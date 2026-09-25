"""
Motor de reconciliacion: aplica las reglas de reconciliation.yaml sobre el
resultado de extract.process_folder().

Lo que un documento por si solo NO puede resolver:
  1. A que embarque pertenece cada archivo (ningun campo unico esta presente
     en TODOS los tipos de documento -- ver agrupar_embarques()).
  2. Si dos valores "son el mismo dato" depende del tipo de campo: IDs se
     comparan exactos, listas de contenedores por superposicion, montos y
     pesos con tolerancia (ver _comparar()).
  3. Cuando NO corresponde comparar: campo no encontrado, confianza baja, o
     directamente sin extractor todavia (ver _resolver_pata()).

Devuelve, por embarque y por regla, uno de tres veredictos -- los mismos que
ya estaban escritos en el tab "Reconciliación" de gui.py:
  VERIFIED         -> al menos 2 patas confiables y todas coinciden
  DISCREPANCY      -> al menos 2 patas confiables y no coinciden
  NOT_VERIFIABLE   -> menos de 2 patas confiables (falta doc, NOT_FOUND,
                     confianza baja, o el campo tiene precision perdida)
"""

import re
import unicodedata
from pathlib import Path

import yaml

HERE = Path(__file__).resolve().parent

CONF_THRESHOLD = 0.85          # mismo umbral que CONF_THRESHOLD en gui.py
NUMERIC_ABS_TOL = 0.01
NUMERIC_REL_TOL = 0.001


# --- Claves de correlacion ---------------------------------------------------
# Ningun campo esta en TODOS los tipos de documento, pero customs_declaration
# tiene SO y contenedor a la vez, asi que hace de puente entre los dos grupos.
# Cada entrada es (doc_type, campo, "so"|"container").
JOIN_KEYS = [
    ("customs_declaration", "f17_contract_no", "so"),
    ("export_invoice", "f22_contract_no", "so"),
    ("sales_contract", "f33_contract_no", "so"),
    ("cargo_manifest", "f46_contract_no", "so"),
    ("tax_filing_directory", "so_no_ref", "so"),
    ("shipping_order", "so_no_ref", "so"),
    ("customs_declaration", "f16_container_nos", "container"),
    ("bill_of_lading", "f43_container_nos", "container"),
    ("cargo_manifest_detail", "f45_container_nos", "container"),
    ("shipping_order", "f58_container_no", "container"),
    # el refund record del fisco NO trae SO ni contenedor: se ata por N° de
    # CDF, que tambien esta en el propio CDF, en la factura (renglon de
    # notas) y en el 3.1
    ("refund_record", "cdf_no", "cdf"),
    ("customs_declaration", "f11_customs_no", "cdf"),
    ("export_invoice", "f28_cdf_no", "cdf"),
    ("tax_filing_directory", "cdf_no_ref", "cdf"),
]

# el "starting point" del equipo: lo que esta en este registro es lo que se
# reclamo al fisco (ver rules.yaml)
CLAIM_DOC_TYPE = "refund_record"

# shipping_order y tax_filing_directory son LOGS (per_row en rules.yaml):
# un solo archivo con filas de muchos embarques. process_xlsx() ya los
# entrega como un record por fila, asi que cada fila se une solo al
# embarque de su SO/contenedor. Las filas de embarques que no tienen
# ningun otro documento en la carpeta quedan en grupos propios, que run()
# descarta (ver _only_logs()).


def load_rules(path=None):
    path = Path(path) if path else HERE / "reconciliation.yaml"
    return yaml.safe_load(open(path, encoding="utf-8"))["rules"]


def load_doc_types(path=None):
    path = Path(path) if path else HERE / "rules.yaml"
    return yaml.safe_load(open(path, encoding="utf-8"))["doc_types"]


def load_checklist(path=None):
    path = Path(path) if path else HERE / "checklist.yaml"
    return yaml.safe_load(open(path, encoding="utf-8"))["documents"]


# --- Normalizacion de valores -------------------------------------------------
def _normalize_id(v):
    """
    IDs que a veces llegan como float (13194443.0) y a veces como texto
    ("13194443") segun el tipo del campo en rules.yaml -- sin esto no
    matchean aunque sean el mismo numero.
    """
    if v is None:
        return None
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = unicodedata.normalize("NFKC", str(v)).strip()
    return s or None


def _as_list(v):
    if v is None:
        return []
    return v if isinstance(v, list) else [v]


# --- Agrupar documentos en embarques (union-find) -----------------------------
class _DSU:
    def __init__(self):
        self.parent = {}

    def find(self, x):
        self.parent.setdefault(x, x)
        while self.parent[x] != x:
            self.parent[x] = self.parent.get(self.parent[x], self.parent[x])
            x = self.parent[x]
        return x

    def union(self, a, b):
        ra, rb = self.find(a), self.find(b)
        if ra != rb:
            self.parent[ra] = rb


def build_records(results):
    """
    Aplana el resultado de extract.process_folder() a una lista de
    "documentos": un dict por pagina/hoja ya clasificada, con su tipo,
    archivo y campos. Paginas SKIPPED/UNRESOLVED/ERROR no aportan nada,
    se descartan aca.
    """
    records = []
    for r in results:
        if not r.get("ok"):
            continue
        for p in r["pages"]:
            doc_type = p["classification"].get("type")
            if not doc_type:
                continue
            records.append({"doc_type": doc_type, "file": r["path"],
                            "page": p["page"], "fields": p["fields"]})
    return records


def cluster_records(records):
    """
    Agrupa records que comparten un valor de SO o de contenedor (union-find).

    Un record SIN ninguna clave (ej. el aviso de liberacion de la pag. 1 del
    CDF, el packing list dentro del contrato) sigue al resto de su archivo:
    si las paginas con clave de ese archivo cayeron todas en un mismo grupo,
    es de ese embarque. Si cayeron en grupos distintos (un log per_row con
    filas de muchos embarques) no hay forma de saber cual, y queda solo --
    run() lo reporta como "sin asignar".
    """
    dsu = _DSU()
    first_holder = {}     # valor normalizado -> indice del primer record que lo tuvo
    keyed = set()

    for idx, rec in enumerate(records):
        for doc_type, field, _kind in JOIN_KEYS:
            if rec["doc_type"] != doc_type or field not in rec["fields"]:
                continue
            raw = rec["fields"][field].get("value")
            for v in _as_list(raw):
                key = _normalize_id(v)
                if not key:
                    continue
                keyed.add(idx)
                if key in first_holder:
                    dsu.union(idx, first_holder[key])
                else:
                    first_holder[key] = idx

    roots_by_file = {}
    for idx in keyed:
        roots_by_file.setdefault(records[idx]["file"], set()).add(dsu.find(idx))
    for idx, rec in enumerate(records):
        roots = roots_by_file.get(rec["file"], set())
        if idx not in keyed and len(roots) == 1:
            dsu.union(idx, next(iter(roots)))

    groups = {}
    for idx in range(len(records)):
        root = dsu.find(idx)
        groups.setdefault(root, []).append(records[idx])
    return list(groups.values())


def _cluster_label(cluster):
    for rec in cluster:
        for doc_type, field, kind in JOIN_KEYS:
            if kind != "so" or rec["doc_type"] != doc_type:
                continue
            v = rec["fields"].get(field, {}).get("value")
            v = _normalize_id(v)
            if v:
                return v
    # sin SO (ej. una exportacion que solo esta en el refund record): el CDF
    cdfs = sorted({v for rec in cluster for v in _key_values(rec, "cdf")})
    if cdfs:
        return f"CDF {cdfs[0]}"
    return " + ".join(sorted({rec["file"] for rec in cluster}))


# --- Resolucion de una "pata" (un ref de una regla) ---------------------------
def _field_spec(doc_types, ref):
    """ref es 'doc_type.campo' o 'doc_type.campo.subcampo' (para f20_price_block)."""
    parts = ref.split(".")
    doc_type, field = parts[0], parts[1]
    spec = doc_types.get(doc_type, {}).get("fields", {}).get(field, {})
    sub = parts[2] if len(parts) > 2 else None
    return spec, sub


def _field_kind(spec, sub):
    if sub:
        for s in spec.get("split_lines", []):
            if s["name"] == sub:
                return "number" if s.get("type") == "number" else "text"
        return "text"
    if spec.get("multi"):
        return "list"
    if spec.get("type") == "number":
        return "number"
    return "text"


def _resolver_pata(cluster, doc_types, item):
    """
    Busca el valor de un leg de regla dentro del cluster. Devuelve un dict
    con status "OK" | "MISSING" | "LOW_CONF" | "LOSSY" y el valor (si hay).
    LOSSY es para campos marcados con nota de precision perdida en
    reconciliation.yaml -- se muestran pero no deciden el veredicto.
    """
    if item.get("ref") is None:
        return {"status": "MISSING", "value": None, "item": item.get("item"),
                "ref": None, "note": item.get("note")}

    doc_type, field = item["ref"].split(".")[0], item["ref"].split(".")[1]
    sub = item["ref"].split(".")[2] if item["ref"].count(".") > 1 else None

    recs = [r for r in cluster if r["doc_type"] == doc_type and field in r["fields"]]
    if not recs:
        return {"status": "MISSING", "value": None, "item": item.get("item"),
                "ref": item["ref"], "note": "documento no presente en este embarque"}

    f = recs[0]["fields"][field]
    if len(recs) > 1 and _field_spec(doc_types, item["ref"])[0].get("multi"):
        # varias filas de un log per_row (ej. un camion por contenedor en
        # shipping_order): la lista del embarque es la union de todas
        f = _merge_list_fields([r["fields"][field] for r in recs])
    val = f["value"]
    if sub:
        # una correccion manual (Revision tab) guarda lo que el analista
        # escribio como TEXTO PLANO, aunque el campo original fuera un dict
        # (ej. f20_price_block) -- ahi val ya no tiene subcampos que sacar.
        if not isinstance(val, dict):
            return {"status": "MISSING", "value": None, "item": item.get("item"),
                    "ref": item["ref"],
                    "note": "sin subcampos (¿se corrigio a mano como texto plano?)"}
        val = val.get(sub)

    base = {"item": item.get("item"), "ref": item["ref"], "unit": item.get("unit"),
            "kind": _field_kind(*_field_spec(doc_types, item["ref"]))}

    if val is None or f["status"] in ("NOT_FOUND", "PARTIAL"):
        return {**base, "status": "MISSING", "value": val, "note": f"campo {f['status']}"}
    if f["conf"] < CONF_THRESHOLD and f["status"] not in ("OK_CORROBORATED", "CORRECTED"):
        return {**base, "status": "LOW_CONF", "value": val,
                "note": f"confianza {f['conf']:.2f}"}
    if item.get("note") and "precision perdida" in item["note"]:
        return {**base, "status": "LOSSY", "value": val, "note": item["note"]}
    return {**base, "status": "OK", "value": val}


def _merge_list_fields(fields):
    """Une los valores de lista; el campo resultante es tan confiable como
    el peor de los que lo forman."""
    found = [f for f in fields if f["value"] is not None]
    if not found:
        return fields[0]
    values = list(dict.fromkeys(v for f in found for v in _as_list(f["value"])))
    worst = min(found, key=lambda f: f["conf"])
    return {**worst, "value": values}


# --- Comparacion ---------------------------------------------------------------
# factor para llevar cada unidad a kilogramos -- unica conversion que hace
# falta hoy (regla 9: neto en kg contra 'cantidad' de la factura en tn).
UNIT_TO_KG = {"kg": 1.0, "ton": 1000.0}


def _convert(value, unit):
    if not unit:
        return value
    try:
        return float(value) * UNIT_TO_KG[unit]
    except (TypeError, ValueError, KeyError):
        return value


def _numbers_close(a, b):
    try:
        a, b = float(a), float(b)
    except (TypeError, ValueError):
        return a == b
    return abs(a - b) <= max(NUMERIC_ABS_TOL, NUMERIC_REL_TOL * max(abs(a), abs(b)))


def _values_equal(a, b, kind):
    if kind == "list":
        return bool(set(map(str, _as_list(a))) & set(map(str, _as_list(b))))
    if kind == "number":
        return _numbers_close(a, b)
    return _normalize_id(a) == _normalize_id(b)


def evaluate_rule(cluster, rule, doc_types):
    legs = [_resolver_pata(cluster, doc_types, item) for item in rule["fields"]]

    # algunas reglas del machete agrupan campos que jamas van a ser el mismo
    # texto aunque ambos esten bien leidos (ej. codigo de modalidad de
    # transporte vs. nombre del buque) -- verificado a mano, ver el 'desc'
    # de la regla. Forzarlas a comparar solo daria falsas discrepancias.
    if rule.get("verifiable") is False:
        return {"id": rule["id"], "desc": rule["desc"].strip(),
                "verdict": "NOT_VERIFIABLE", "legs": legs}

    verdict = _verdict(legs)
    return {"id": rule["id"], "desc": rule["desc"].strip(), "verdict": verdict,
            "legs": legs}


def _agree(legs):
    """True si TODAS las patas dadas tienen el mismo valor (>=2 patas)."""
    kind = legs[0].get("kind", "text")
    base = _convert(legs[0]["value"], legs[0].get("unit"))
    return all(_values_equal(base, _convert(leg["value"], leg.get("unit")), kind)
              for leg in legs[1:])


def _verdict(legs):
    """
    Un valor de baja confianza (LOW_CONF/LOSSY) que IGUAL coincide con otro
    es evidencia a favor, no en contra -- dos lecturas independientes de
    dudosa calidad rara vez coinciden por accidente. Por eso participa en
    el chequeo de coincidencia. Lo que NO hace es acusar: si lo que no
    coincide es una pata de baja confianza, el resultado es NOT_VERIFIABLE
    (podria ser un error de lectura), no DISCREPANCY. Solo se declara
    DISCREPANCY cuando las patas que SI se confian entre si (status OK)
    no coinciden.
    """
    comparable = [leg for leg in legs if leg["status"] in ("OK", "LOW_CONF", "LOSSY")]
    if len(comparable) < 2:
        return "NOT_VERIFIABLE"

    if _agree(comparable):
        return "VERIFIED"

    ok_legs = [leg for leg in comparable if leg["status"] == "OK"]
    if len(ok_legs) < 2:
        return "NOT_VERIFIABLE"          # no hay confianza suficiente para acusar
    return "VERIFIED" if _agree(ok_legs) else "DISCREPANCY"


def _only_logs(cluster, doc_types):
    """
    True si todo el grupo son filas de logs per_row (shipping_order,
    tax_filing_directory): esos logs traen embarques que no son de esta
    carpeta, y dos filas del mismo SO no hacen un embarque a reconciliar.
    """
    return all(doc_types.get(rec["doc_type"], {}).get("per_row") for rec in cluster)


def _claimed(cluster):
    return any(rec["doc_type"] == CLAIM_DOC_TYPE for rec in cluster)


def _key_values(rec, kind=None):
    """Valores de clave de correlacion (SO/contenedor) que trae un record."""
    out = []
    for doc_type, field, k in JOIN_KEYS:
        if rec["doc_type"] != doc_type or (kind and k != kind):
            continue
        raw = rec["fields"].get(field, {}).get("value")
        out += [key for key in map(_normalize_id, _as_list(raw)) if key]
    return out


# --- Checklist de documentacion -----------------------------------------------
PRESENT, MISSING, MANUAL = "PRESENT", "MISSING", "MANUAL"


def evaluate_checklist(cluster, documents):
    """
    Un renglon por documento de checklist.yaml:
      PRESENT -> hay al menos una pagina/fila de alguno de sus doc_types
      MISSING -> tiene extractor y no aparece en el embarque
      MANUAL  -> sin extractor: no se puede afirmar ni que esta ni que falta
    """
    out = []
    for d in documents:
        types = d.get("doc_types") or []
        hits = [{"file": rec["file"], "page": rec["page"]}
                for rec in cluster if rec["doc_type"] in types]
        status = MANUAL if not types else (PRESENT if hits else MISSING)
        out.append({"code": d["code"], "name": d["name"], "status": status,
                    "sources": hits, "doc_types": types})
    return out


def folder_warnings(records, doc_types, documents=()):
    """
    Avisos de la carpeta entera (no de un embarque): mas de un archivo de
    un tipo single_file. Se avisa una vez aca en vez de repetirlo en cada
    embarque -- el problema es de la carpeta, no de cada exportacion.
    """
    out = []
    for dt, cfg in doc_types.items():
        if not cfg.get("single_file"):
            continue
        files = sorted({r["file"] for r in records if r["doc_type"] == dt})
        if len(files) > 1:
            name = next((f"{d['code']} {d['name']}" for d in documents
                         if dt in (d.get("doc_types") or [])), dt)
            out.append(f"{len(files)} files of type '{name}' in the folder "
                       f"({', '.join(files)}) -- keep only one: when they "
                       "disagree, the first one read is the one used")
    return out


def _warnings(cluster, checklist, doc_types, claims_loaded):
    """Senales de que el grupo puede estar mal armado (union-find es
    transitivo: una sola clave mal leida pega dos embarques), o de que no
    cierra contra lo reclamado al fisco."""
    out = []
    if _only_logs(cluster, doc_types) and _claimed(cluster):
        out.append("claimed in the refund record but NO documents found in the folder")
    elif claims_loaded and not _claimed(cluster):
        out.append("has documents but is NOT in the refund record (not claimed?)")
    sos = sorted({v for rec in cluster for v in _key_values(rec, "so")})
    if len(sos) > 1:
        out.append(f"mixes {len(sos)} different SOs ({', '.join(sos)}) -- "
                   "check that these documents belong to the same shipment")
    single = {dt for dt, cfg in doc_types.items() if cfg.get("single_file")}
    for item in checklist:
        if set(item.get("doc_types", [])) & single:
            continue              # ya lo cubre folder_warnings(), una sola vez
        files = sorted({s["file"] for s in item["sources"]})
        if len(files) > 1:
            out.append(f"{item['code']} {item['name']} comes from "
                       f"{len(files)} different files: {', '.join(files)}")
    return out


def run(results, rules=None, doc_types=None, documents=None):
    """
    Devuelve {"shipments": [...], "unassigned": [...], "folder_warnings": [...]}:
      shipments       -> un embarque por grupo con alguna clave (SO/contenedor/
                         CDF): reglas evaluadas, checklist de documentos y avisos
      unassigned      -> paginas sin ninguna clave que tampoco se pudieron
                         atar al embarque de su archivo (ver cluster_records)
      folder_warnings -> problemas de la carpeta entera (ver folder_warnings)
    """
    rules = rules if rules is not None else load_rules()
    doc_types = doc_types if doc_types is not None else load_doc_types()
    documents = documents if documents is not None else load_checklist()

    records = build_records(results)
    clusters = cluster_records(records)
    # "no esta reclamado" solo tiene sentido si el refund record esta en la
    # carpeta; sin el, todas las exportaciones darian ese aviso
    claims_loaded = any(rec["doc_type"] == CLAIM_DOC_TYPE for rec in records)

    shipments, unassigned = [], []
    for cluster in clusters:
        if not any(_key_values(rec) for rec in cluster):
            unassigned += [{"file": rec["file"], "page": rec["page"],
                            "doc_type": rec["doc_type"]} for rec in cluster]
            continue
        if _only_logs(cluster, doc_types) and not _claimed(cluster):
            continue          # embarque que solo figura en los logs, sin docs en la carpeta
        # si esta en el refund record, se reporta AUNQUE no tenga documentos:
        # "se reclamo el reintegro y no hay legajo" es lo mas grave de todo
        # un embarque con un solo documento se reporta igual: sus reglas
        # salen NOT_VERIFIABLE y el checklist muestra lo que falta, que es
        # justamente lo que hay que saber de el
        checklist = evaluate_checklist(cluster, documents)
        shipments.append({
            "shipment": _cluster_label(cluster),
            "doc_types": sorted({rec["doc_type"] for rec in cluster}),
            "files": sorted({rec["file"] for rec in cluster}),
            # que paginas/filas componen el embarque -- lo usa organize.py
            # para copiar documentos y armar el reporte por embarque
            "records": [{"file": rec["file"], "page": rec["page"],
                         "doc_type": rec["doc_type"]} for rec in cluster],
            "rules": [evaluate_rule(cluster, r, doc_types) for r in rules],
            "checklist": checklist,
            "warnings": _warnings(cluster, checklist, doc_types, claims_loaded),
        })
    shipments.sort(key=lambda s: s["shipment"])
    return {"shipments": shipments, "unassigned": unassigned,
            "folder_warnings": folder_warnings(records, doc_types, documents)}


if __name__ == "__main__":
    import json
    import sys

    import extract

    target = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "samples"
    rules_yaml = yaml.safe_load(open(HERE / "rules.yaml", encoding="utf-8"))
    dups = extract.duplicate_single_files(target, rules_yaml)
    if dups:
        sys.exit("Archivos duplicados de un tipo que tiene que ser unico -- dejar uno solo:\n"
                 + "\n".join(f"  {dt}: {', '.join(files)}" for dt, files in dups.items()))
    results = extract.process_folder(target, rules_yaml)
    out = run(results)

    for shipment in out["shipments"]:
        print(f"\n{'=' * 78}\nEmbarque: {shipment['shipment']}  "
              f"(docs: {', '.join(shipment['doc_types'])})")
        for w in shipment["warnings"]:
            print(f"  ! {w}")
        for item in shipment["checklist"]:
            src = ", ".join(f"{s['file']} p{s['page']}" for s in item["sources"])
            print(f"  [{item['status']:7}] {item['code']:8} {item['name'][:45]:45} {src}")
        for r in shipment["rules"]:
            print(f"  [{r['verdict']:14}] regla {r['id']:2}: {r['desc'][:70]}")
            for leg in r["legs"]:
                print(f"      item {leg.get('item')!s:>4} {leg['status']:9} "
                      f"{leg.get('ref') or '(sin extractor)':40} = {leg['value']!r}")
    for w in out["folder_warnings"]:
        print(f"\n!! {w}")
    for u in out["unassigned"]:
        print(f"\nSin asignar: {u['file']} p{u['page']} ({u['doc_type']})")

    dest = HERE / "output" / "reconciliation.json"
    dest.parent.mkdir(exist_ok=True)
    json.dump(out, open(dest, "w", encoding="utf-8"), ensure_ascii=False, indent=2)
    print(f"\nJSON completo: {dest}")
