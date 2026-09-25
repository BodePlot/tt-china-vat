"""
Clasificacion de pagina + extraccion de campos guiada por rules.yaml.

Idea central: NO se usan coordenadas absolutas ni el nombre del archivo.
Cada valor se localiza por su RELACION con la etiqueta impresa, que es
lo unico que no cambia entre escaneos, versiones del formulario o layouts.
"""

import hashlib
import json
import re
import time
import unicodedata
from pathlib import Path

import cv2
import openpyxl
import pymupdf
import yaml

OCR_DPI = 300

# Triage: saltear paginas sin datos (terminos y condiciones, anexos legales).
# Ver triage_page() para por que el criterio es el tamano de las cajas.
TRIAGE_ENABLED = True
TRIAGE_MAX_BOX_HEIGHT = 250      # px a OCR_DPI
# la deteccion del triage corre sobre la imagen achicada a esta escala: no
# hace falta resolucion para medir el tamano de las cajas, y a 0.4 la
# decision de saltear fue identica a la de resolucion completa en samples/
TRIAGE_SCALE = 0.4

# Paginas de un archivo puntual que se saben sin datos y el triage automatico
# NO detecta -- confirmado a mano, no adivinado. Pagina 3 de 3.7.2 Cargo
# Manifest.pdf es texto legal de la naviera (responsabilidad por devolucion
# de contenedor, sanciones) en pocas lineas cortas, no un parrafo grande, asi
# que triage_page() no la marca como boilerplate (ver su umbral por altura de
# caja) y termina corriendo OCR completo para terminar en UNRESOLVED sin
# ningun campo (ver la nota en rules.yaml sobre 3.7.2). Se saltea directo, sin
# gastar ni el detect-only del triage.
KNOWN_BOILERPLATE_PAGES = {
    "3.7.2 Cargo Manifest.pdf": {3},
}

# Cache de OCR por pagina: el OCR tarda 10-25s/pagina y no cambia mientras
# el PDF no cambie. Sin esto, iterar sobre rules.yaml significa repetir
# OCR completo en cada corrida aunque los documentos sean siempre los mismos.
CACHE_ENABLED = True
CACHE_DIR = Path(__file__).resolve().parent / "output" / "ocr_cache"
CACHE_VERSION = 1

_ocr_engine = None


def _cache_file(path):
    st = Path(path).stat()
    # la clave incluye mtime/size (detecta re-escaneos) y los parametros que
    # afectan el resultado del OCR: si cambian, la cache vieja queda huerfana
    # (nunca matchea) y se regenera sola, sin que haya que borrarla a mano.
    raw = (f"{Path(path).resolve()}|{st.st_mtime_ns}|{st.st_size}|"
           f"{OCR_DPI}|{TRIAGE_ENABLED}|{TRIAGE_MAX_BOX_HEIGHT}|v{CACHE_VERSION}")
    digest = hashlib.sha1(raw.encode("utf-8")).hexdigest()
    return CACHE_DIR / f"{digest}.json"


def _load_cache(path):
    if not CACHE_ENABLED:
        return {}
    f = _cache_file(path)
    if f.is_file():
        try:
            return json.loads(f.read_text(encoding="utf-8"))
        except Exception:
            return {}
    return {}


def _save_cache(path, cache):
    if not CACHE_ENABLED:
        return
    CACHE_DIR.mkdir(parents=True, exist_ok=True)
    _cache_file(path).write_text(json.dumps(cache, ensure_ascii=False),
                                  encoding="utf-8")


# --- OCR --------------------------------------------------------------------
def get_ocr():
    global _ocr_engine
    if _ocr_engine is None:
        from rapidocr_onnxruntime import RapidOCR
        _ocr_engine = RapidOCR()
    return _ocr_engine


def normalize(s: str) -> str:
    """
    NFKC pasa los caracteres de ancho completo a normales:
    '（千克）' -> '(千克)',  '１２３' -> '123'.
    Sin esto, los match de etiquetas fallan de forma silenciosa.
    """
    if not s:
        return ""
    s = unicodedata.normalize("NFKC", s)
    return re.sub(r"[\s　]+", "", s)


def page_items(page) -> list[dict]:
    """Cajas de texto con posicion. Usa texto nativo si el PDF lo trae."""
    native = page.get_text("dict")
    boxes = []
    for blk in native.get("blocks", []):
        for line in blk.get("lines", []):
            txt = "".join(sp["text"] for sp in line.get("spans", []))
            if txt.strip():
                x0, y0, x1, y1 = line["bbox"]
                boxes.append({"text": txt, "conf": 1.0, "source": "native",
                              "x": x0, "y": y0, "w": x1 - x0, "h": y1 - y0})
    if len(boxes) >= 5:                      # el PDF tenia texto real
        boxes.sort(key=lambda b: (round(b["y"] / 12), b["x"]))
        # el texto nativo es gratis: no hay nada que ahorrar salteandolo
        return boxes, {"skip": False, "median_box_h": None, "n_boxes": len(boxes)}

    pix = page.get_pixmap(dpi=OCR_DPI)
    png = pix.tobytes("png")

    triage = triage_page(png)
    if triage["skip"]:
        return [], triage                    # no se paga el reconocimiento

    # Si la version instalada ignoro use_rec=False, el triage ya reconocio el
    # texto: se reusa en vez de correr el OCR de nuevo. Sin esto el triage
    # DUPLICA el costo en toda pagina que no se saltea.
    result = triage.pop("_recognized", None)
    det_boxes = triage.pop("_boxes", None)

    # Un escaneo de costado: el reconocimiento sale bien (el modelo endereza
    # cada linea sola) pero las COORDENADAS de las cajas quedan en el marco
    # rotado, y eso rompe right/below/table_col aunque page_regex no lo note.
    # Se detecta por el aspecto de las cajas (altas en vez de anchas); cual
    # de los dos sentidos (+90/-90) es el correcto NO se puede saber por
    # aspecto (los dos dan cajas anchas) -- lo resuelve _try_rotation_fix
    # con el clasificador de angulo, en un solo OCR.
    #
    # Con las cajas del triage el aspecto se ve SIN reconocer nada: una
    # pagina de costado va directo al OCR rotado, sin pagar antes un OCR
    # completo sin rotar que se tiraria (~6s por pagina en 3.5).
    rotation = 0
    if result is None and det_boxes and _looks_rotated([[b, "", 1] for b in det_boxes]):
        angle, fixed = _try_rotation_fix(page)
        if fixed is not None:
            result, rotation = fixed, angle
    if result is None:
        result, _ = get_ocr()(png)
        if _looks_rotated(result):           # sin triage (o sin cajas sueltas)
            angle, fixed = _try_rotation_fix(page)
            if fixed is not None:
                result, rotation = fixed, angle
    triage["rotation"] = rotation

    for box, txt, conf in (result or []):
        xs = [p[0] for p in box]
        ys = [p[1] for p in box]
        boxes.append({"text": txt, "conf": round(float(conf), 3), "source": "ocr",
                      "x": min(xs), "y": min(ys),
                      "w": max(xs) - min(xs), "h": max(ys) - min(ys)})
    boxes.sort(key=lambda b: (round(b["y"] / 12), b["x"]))
    return boxes, triage


def _cell_text(v):
    """
    str(float) usa notacion cientifica para numeros grandes ('5.3382e+17'),
    lo que rompe cualquier regex/comparacion de digitos. Los IDs largos
    (numero de CDF, de factura) llegan como float si Excel los guardo como
    numero en vez de texto -- ya perdieron precision ahi (float64 solo
    tiene ~15-17 digitos significativos), pero al menos se preserva lo
    que SI se puede leer en vez de mostrar notacion cientifica encima.
    """
    if isinstance(v, float) and v.is_integer():
        return f"{v:.0f}"
    return str(v)


def sheet_items(ws) -> list[dict]:
    """
    Celdas de una hoja de Excel, en el mismo formato que page_items(): cada
    celda es una 'caja' con texto y posicion. x/y son numero de columna/fila
    (no pixeles) -- alcanza para que _find_label() y loc_xlsx_col() ubiquen
    el header y bajen por la columna, no hace falta layout real.
    """
    boxes = []
    for row in ws.iter_rows():
        for cell in row:
            if cell.value is None:
                continue
            boxes.append({"text": _cell_text(cell.value), "conf": 1.0,
                          "source": "xlsx", "x": cell.column, "y": cell.row,
                          "w": 1, "h": 1})
    return boxes


# --- Triage de pagina -------------------------------------------------------
def triage_page(png_bytes):
    """
    Decide si vale la pena reconocer el texto de esta pagina, corriendo
    SOLO la etapa de deteccion (~1s, ver _detect_only) en vez del OCR
    completo (4-25s).

    Que distingue una pagina de terminos y condiciones: no es que tenga
    muchas cajas -- tiene MENOS que una pagina de datos. Lo que la delata
    es el TAMANO: el detector agrupa parrafos enteros en cajas gigantes,
    y reconocer una caja gigante cuesta mucho mas que varias chicas.

    Medido sobre 3.5_contract.pdf (6 paginas):
        paginas con datos     -> altura mediana 135-156 px, 4s de OCR
        paginas de terminos   -> altura mediana 440-2880 px, 18-25s de OCR
    No se superponen, por eso el umbral en 250 tiene margen de sobra.

    OJO: calibrar contra documentos reales antes de subir esto a produccion.
    Saltear una pagina que SI tenia un dato es un error silencioso.
    """
    if not TRIAGE_ENABLED:
        return {"skip": False, "median_box_h": None, "n_boxes": None}

    # TODO el calculo va adentro del try: cualquier formato inesperado de la
    # libreria degrada a "no saltear" en vez de tirar abajo el documento.
    try:
        items = _detect_only(png_bytes)
        heights = [h for h in (_box_height(b) for b in items) if h and h > 0]
        if not heights:
            return {"skip": False, "median_box_h": None, "n_boxes": len(items),
                    "reason": "couldn't measure the boxes"}
        heights.sort()
        median = float(heights[len(heights) // 2])
    except Exception as e:
        return {"skip": False, "median_box_h": None, "n_boxes": None,
                "reason": f"triage unavailable ({type(e).__name__}: {e})"}

    skip = median > TRIAGE_MAX_BOX_HEIGHT
    out = {"skip": skip, "median_box_h": round(median, 1),
           "n_boxes": len(items),
           "reason": "boilerplate (paragraph-sized boxes)" if skip else None}

    # si la libreria devolvio tambien el texto, se guarda para no reconocerlo
    # dos veces; page_items() lo consume y lo saca del registro
    if not skip and _has_text(items):
        out["_recognized"] = items
        out["recognized_in_triage"] = True
    elif not skip:
        # cajas sueltas: page_items() las usa para ver si la pagina esta de
        # costado ANTES de pagar el OCR sin rotar (y tambien las saca)
        out["_boxes"] = items
    return out


def _has_text(items):
    """True si los elementos traen texto, o sea que use_rec=False fue ignorado."""
    try:
        return bool(items) and isinstance(items[0][1], str)
    except (TypeError, IndexError, KeyError):
        return False


def _detect_only(png_bytes):
    """
    Solo la deteccion de cajas, sin reconocer el texto, sobre la imagen
    achicada a TRIAGE_SCALE. Devuelve cajas sueltas en pixeles del render
    ORIGINAL (a OCR_DPI), asi TRIAGE_MAX_BOX_HEIGHT no depende de la escala.

    Se llama al detector directo porque rapidocr 1.2.3 IGNORA use_rec=False
    en __call__: el camino "oficial" hacia el OCR completo igual, incluso
    en las paginas de terminos que despues se saltean -- justo el costo que
    el triage existe para evitar. Medido sobre los 3 PDF escaneados de
    samples/: el triage paso de costar un OCR completo por pagina a ~1s.

    Si la version instalada no expone text_detector, cae al camino viejo
    (resultado completo, que _box_height() tambien sabe leer).
    """
    eng = get_ocr()
    if hasattr(eng, "text_detector") and hasattr(eng, "load_img"):
        img = eng.load_img(png_bytes)
        h, w = img.shape[:2]
        small = cv2.resize(img, (max(1, int(w * TRIAGE_SCALE)),
                                 max(1, int(h * TRIAGE_SCALE))))
        boxes, _ = eng.text_detector(small)
        if boxes is None:
            return []
        return [[[x / TRIAGE_SCALE, y / TRIAGE_SCALE] for x, y in b.tolist()]
                for b in boxes]

    try:
        out = get_ocr()(png_bytes, use_det=True, use_cls=False, use_rec=False)
    except TypeError:
        out = get_ocr()(png_bytes)          # version que no acepta los flags

    if isinstance(out, tuple):              # (resultado, tiempos)
        out = out[0]
    if out is None:
        return []
    if hasattr(out, "boxes"):               # objeto de resultado
        out = out.boxes if out.boxes is not None else []
    return list(out)


def _looks_like_points(v):
    """True si v son puntos [[x,y], [x,y], ...] con numeros adentro."""
    try:
        return len(v) >= 3 and len(v[0]) == 2 and all(
            isinstance(c, (int, float)) for c in v[0])
    except (TypeError, IndexError, KeyError):
        return False


def _box_rect(b):
    """
    (ancho, alto) de una caja. Tolera los tres formatos que devuelven las
    distintas versiones, y ante cualquier otro devuelve (None, None) en vez
    de tirar error:
      [[x,y] x4]                    -> caja suelta
      [[[x,y] x4], texto, conf]     -> resultado completo (use_rec ignorado)
      [x1, y1, x2, y2]              -> caja como rectangulo
    """
    if _looks_like_points(b):
        xs, ys = [p[0] for p in b], [p[1] for p in b]
        return max(xs) - min(xs), max(ys) - min(ys)

    # resultado completo: la caja es el primer elemento
    try:
        if _looks_like_points(b[0]):
            xs, ys = [p[0] for p in b[0]], [p[1] for p in b[0]]
            return max(xs) - min(xs), max(ys) - min(ys)
    except (TypeError, IndexError, KeyError):
        pass

    # rectangulo plano
    try:
        if len(b) == 4 and all(isinstance(c, (int, float)) for c in b):
            return abs(b[2] - b[0]), abs(b[3] - b[1])
    except (TypeError, IndexError):
        pass

    return None, None


def _box_height(b):
    return _box_rect(b)[1]


def _ocr_with_flip_ratio(png_bytes):
    """
    OCR completo que ademas devuelve que fraccion de las lineas el
    clasificador de angulo marco como giradas 180 grados. Es el mismo
    pipeline de rapidocr.__call__ (deteccion -> recorte -> clasificador ->
    reconocimiento), armado a mano solo para no tirar ese dato: __call__ lo
    calcula y lo descarta.

    Sin la API interna esperada, cae al __call__ normal con ratio 0.0 (o
    sea "no esta al reves"), que es lo que se asumia antes.
    """
    eng = get_ocr()
    needed = ("load_img", "text_detector", "sorted_boxes", "get_crop_img_list",
              "text_cls", "text_recognizer", "filter_boxes_rec_by_score")
    if not all(hasattr(eng, a) for a in needed):
        return eng(png_bytes)[0], 0.0

    img = eng.load_img(png_bytes)
    boxes, _ = eng.text_detector(img)
    if boxes is None or len(boxes) == 0:
        return None, 0.0
    boxes = eng.sorted_boxes(boxes)
    crops = eng.get_crop_img_list(img, boxes)
    crops, cls_res, _ = eng.text_cls(crops)
    flipped = sum(1 for label, _ in cls_res if label == "180") / len(cls_res)
    rec_res, _ = eng.text_recognizer(crops)
    boxes, rec_res = eng.filter_boxes_rec_by_score(boxes, rec_res)
    result = [[b.tolist(), r[0], str(r[1])] for b, r in zip(boxes, rec_res)]
    return (result or None), flipped


def _looks_rotated(result):
    """
    Cajas 'altas' en vez de 'anchas' en un resultado YA reconocido indican
    que el escaneo esta de costado. El propio result (no una pasada de
    deteccion aparte) alcanza: no cuesta nada extra en el caso comun.
    """
    dims = [_box_rect(item[0]) for item in (result or []) if item]
    dims = [(w, h) for w, h in dims if w and h]
    if len(dims) < 5:
        return False
    med_w = sorted(w for w, _ in dims)[len(dims) // 2]
    med_h = sorted(h for _, h in dims)[len(dims) // 2]
    return med_h > med_w * 1.3


def _try_rotation_fix(page):
    """
    Un solo OCR con la pagina rotada +90. Si quedo cabeza abajo (el escaneo
    venia girado para el otro lado), NO se vuelve a leer: el clasificador de
    angulo ya enderezo cada linea, asi que el texto salio bien -- solo las
    coordenadas quedaron en el marco invertido, y se dan vuelta con una
    cuenta. El resultado es el mismo que haber leido a -90.

    Por que no se elige por confianza de reconocimiento (lo que se hacia
    antes, con un OCR completo a +90 y otro a -90): medido sobre
    3.5 Sales Contract.pdf, las dos orientaciones dan ~0.78 de confianza
    promedio -- el clasificador endereza cada linea sola, asi que el texto
    al reves "se lee bien" igual. Lo que SI separa los dos sentidos es
    cuantas lineas tuvo que dar vuelta el clasificador: 1-8 de ~55 en el
    sentido correcto, 53-58 en el invertido.
    """
    mat = pymupdf.Matrix(OCR_DPI / 72, OCR_DPI / 72).prerotate(90)
    pix = page.get_pixmap(matrix=mat)
    try:
        result, flipped = _ocr_with_flip_ratio(pix.tobytes("png"))
    except Exception:
        return 0, None
    if result is None or flipped <= 0.5:
        return 90, result
    w, h = pix.width, pix.height
    flipped_result = [[[[w - x, h - y] for x, y in box], txt, conf]
                      for box, txt, conf in result]
    return -90, flipped_result


# --- Clasificacion ----------------------------------------------------------
def classify(items, doc_types, min_score=0.5, min_margin=0.15):
    """
    Puntua la pagina contra la firma de cada tipo.
    No alcanza con ganar: hay que ganar POR DIFERENCIA. Si dos tipos
    quedan parejos, no se adivina -> va a revision manual.
    """
    blob = normalize(" ".join(i["text"] for i in items))
    scores = {}

    for name, cfg in doc_types.items():
        sig = cfg.get("signature", {})
        req = [normalize(k) for k in sig.get("required", [])]
        opt = [normalize(k) for k in sig.get("optional", [])]
        neg = [normalize(k) for k in sig.get("negative", [])]

        if any(k in blob for k in neg):
            scores[name] = 0.0
            continue
        hits_req = sum(k in blob for k in req)
        if req and hits_req < len(req):
            scores[name] = 0.4 * (hits_req / len(req))   # parcial, no clasifica
            continue
        hits_opt = sum(k in blob for k in opt) if opt else 0
        scores[name] = 0.7 + 0.3 * (hits_opt / len(opt) if opt else 1)

    ranked = sorted(scores.items(), key=lambda kv: kv[1], reverse=True)
    best, best_score = ranked[0]
    second = ranked[1][1] if len(ranked) > 1 else 0.0

    if best_score < min_score or (best_score - second) < min_margin:
        return {"type": None, "score": best_score, "scores": scores,
                "status": "UNRESOLVED"}
    return {"type": best, "score": round(best_score, 3), "scores": scores,
            "status": "OK"}


# --- Localizadores ----------------------------------------------------------
def _find_label(items, label):
    """Primera caja que contiene la etiqueta (comparando normalizado)."""
    lab_n = normalize(label)
    for it in items:
        if lab_n in normalize(it["text"]):
            return it
    return None


def loc_inline(items, spec):
    """Etiqueta y valor en la MISMA caja: '海关编号：533820250380000144'."""
    lab = _find_label(items, spec["label"])
    if not lab:
        return None
    txt = unicodedata.normalize("NFKC", lab["text"])
    sep = unicodedata.normalize("NFKC", spec.get("sep", ":"))
    tail = txt.split(sep, 1)[1] if sep in txt else txt
    if spec.get("regex"):
        m = re.search(spec["regex"], tail)
        return {**lab, "value": m.group(0)} if m else None
    return {**lab, "value": tail.strip()}


def loc_below(items, spec):
    """
    Valor en la caja de abajo, dentro de la misma columna.
    El filtro de solape horizontal es lo que evita agarrar el valor
    de la celda de al lado (ej. tomar 净重 cuando se pidio 毛重).
    """
    lab = _find_label(items, spec["label"])
    if not lab:
        return None
    lx1, lx2 = lab["x"], lab["x"] + lab["w"]
    max_dy = 2.5 * lab["h"]                     # escala sola con el DPI
    rx = re.compile(spec["regex"]) if spec.get("regex") else None

    cands = []
    for it in items:
        dy = it["y"] - lab["y"]
        if not (0 < dy <= max_dy):
            continue
        if it["x"] >= lx2 or it["x"] + it["w"] <= lx1:   # otra columna
            continue
        val = unicodedata.normalize("NFKC", it["text"]).strip()
        if rx and not rx.search(val):
            continue
        cands.append((dy, {**it, "value": val}))
    if not cands:
        return None
    return min(cands, key=lambda c: c[0])[1]


def loc_right(items, spec):
    """Valor a la derecha, en la misma fila."""
    lab = _find_label(items, spec["label"])
    if not lab:
        return None
    ly1, ly2 = lab["y"], lab["y"] + lab["h"]
    cands = []
    for it in items:
        dx = it["x"] - (lab["x"] + lab["w"])
        # calibrado contra 3.5 Sales Contract.pdf ("合同号：" -> valor a
        # ~8.8x la altura de la etiqueta): 6x se quedaba corto y descartaba
        # el unico valor real en la fila.
        if dx < 0 or dx > 12 * lab["h"]:
            continue
        if it["y"] >= ly2 or it["y"] + it["h"] <= ly1:   # otra fila
            continue
        cands.append((dx, {**it, "value": it["text"].strip()}))
    if not cands:
        return None
    return min(cands, key=lambda c: c[0])[1]


def loc_page_regex(items, spec):
    """Patron sobre todo el texto de la pagina (ej. numeros de contenedor)."""
    blob = unicodedata.normalize("NFKC", " ".join(i["text"] for i in items))
    found = re.findall(spec["regex"], blob)
    if not found:
        return None
    found = [f if isinstance(f, str) else f[0] for f in found]
    if spec.get("unique"):
        found = list(dict.fromkeys(found))
    if spec.get("multi"):
        return {"value": found, "conf": min(i["conf"] for i in items)}
    return {"value": found[0], "conf": 1.0}


def loc_xlsx_col(items, spec):
    """
    Columna de un Excel identificada por el texto de su header (fila 1).
    A diferencia de table_col (una sola fila de datos, cortada al primer
    salto), esto es para logs con muchas filas: junta el valor de TODAS
    las filas debajo del header, porque cada fila es un registro propio
    (ej. un despacho de camion), no una linea partida de la misma celda.
    """
    lab = _find_label(items, spec["header"])
    if not lab:
        return None
    col = lab["x"]
    header_row = lab["y"]
    rx = re.compile(spec["regex"]) if spec.get("regex") else None

    found = []
    for it in sorted(items, key=lambda i: i["y"]):
        if it["x"] != col or it["y"] <= header_row:
            continue
        val = it["text"].strip()
        if rx:
            m = rx.search(val)
            val = m.group(0) if m else None
        if val:
            found.append(val)
    if not found:
        return None
    if spec.get("unique"):
        found = list(dict.fromkeys(found))
    if spec.get("multi"):
        return {"value": found, "conf": 1.0}
    return {"value": found[0], "conf": 1.0}


def loc_table_col(items, spec):
    """
    Columna de tabla delimitada por dos headers.
    Devuelve las lineas apiladas en la celda, ordenadas de arriba a abajo.
    """
    t = spec["table"]
    h_from = _find_label(items, t["header_from"])
    if not h_from:
        return None
    h_to = _find_label(items, t["header_to"]) if t.get("header_to") else None

    x1 = h_from["x"]
    x2 = h_to["x"] if h_to else x1 + h_from["w"] * 4
    header_y = h_from["y"] + h_from["h"]

    rows = []
    for it in items:
        if it["y"] < header_y:
            continue
        cx = it["x"] + it["w"] / 2               # centro de la caja
        if x1 <= cx <= x2:
            rows.append(it)
    if not rows:
        return None
    rows.sort(key=lambda r: r["y"])

    # corta al terminar la primera fila de items (salto vertical grande)
    cut, gap = [rows[0]], 3.0 * rows[0]["h"]
    for prev, cur in zip(rows, rows[1:]):
        if cur["y"] - prev["y"] > gap:
            break
        cut.append(cur)
    return {"lines": cut, "conf": min(r["conf"] for r in cut)}


LOCATORS = {"inline": loc_inline, "below": loc_below, "right": loc_right,
            "page_regex": loc_page_regex, "table_col": loc_table_col,
            "xlsx_col": loc_xlsx_col}


# --- Post-proceso -----------------------------------------------------------
def to_number(s):
    s = re.sub(r"[^\d.\-]", "", unicodedata.normalize("NFKC", str(s)))
    try:
        return float(s)
    except ValueError:
        return None


def resolve_field(items, spec):
    """Aplica la estrategia y arma el resultado con estado y confianza."""
    loc = LOCATORS[spec["strategy"]](items, spec)
    if loc is None:
        return {"value": None, "conf": 0.0, "status": "NOT_FOUND"}

    # --- tabla: lineas apiladas
    if spec["strategy"] == "table_col":
        lines = [unicodedata.normalize("NFKC", l["text"]).strip()
                 for l in loc["lines"]]

        if spec.get("split_lines"):
            out, ok = {}, True
            for i, sub in enumerate(spec["split_lines"]):
                raw = lines[i] if i < len(lines) else None
                if raw is None:
                    ok = False
                    out[sub["name"]] = None
                elif sub.get("type") == "number":
                    out[sub["name"]] = to_number(raw)
                else:
                    out[sub["name"]] = raw
            return {"value": out, "conf": loc["conf"],
                    "status": "OK" if ok else "PARTIAL",
                    "origin": _origin(loc["lines"])}

        sel = spec.get("line", 1)
        val = lines[0] if sel == 1 else spec.get("join", "").join(lines[1:])
        if spec.get("strip_prefix"):
            val = re.sub(spec["strip_prefix"], "", val).strip()
        if spec.get("regex"):
            m = re.search(spec["regex"], val)
            val = m.group(1) if m else val
        if spec.get("type") == "number":
            val = to_number(val)
        return {"value": val or None, "conf": loc["conf"],
                "status": "OK" if val else "NOT_FOUND",
                "origin": _origin(loc["lines"])}

    # --- campo simple
    val = loc["value"]
    if spec.get("type") == "number":
        val = to_number(val)
    res = {"value": val, "conf": loc.get("conf", 1.0), "status": "OK"}
    if "x" in loc:
        res["origin"] = _origin([loc])

    if spec.get("code_from_label"):              # ej. 运输方式(2) -> code 2
        lab = _find_label(items, spec["label"])
        m = re.search(spec["code_from_label"],
                      unicodedata.normalize("NFKC", lab["text"])) if lab else None
        res["code"] = m.group(1) if m else None
    return res


def _origin(boxes):
    """Trazabilidad: de que coordenadas salio el valor."""
    return [{"x": round(b["x"]), "y": round(b["y"]),
             "w": round(b["w"]), "h": round(b["h"]),
             "conf": b.get("conf", 1.0)} for b in boxes]


def corroborate(items, field, spec):
    """Mismo dato leido desde otro lugar de la pagina -> confirma o marca."""
    if not spec.get("corroborate") or field["value"] is None:
        return field
    others = []
    for alt in spec["corroborate"]:
        r = LOCATORS[alt["strategy"]](items, alt)
        if r and r.get("value"):
            v = r["value"]
            others.append(v[0] if isinstance(v, list) else v)
    field["corroborations"] = others
    if others and all(o == field["value"] for o in others):
        field["status"] = "OK_CORROBORATED"
        field["conf"] = max(field["conf"], 0.99)
    elif others:
        field["status"] = "CONFLICT"
    return field


# --- Entrypoint -------------------------------------------------------------
def process_xlsx(path, rules):
    """
    Contraparte de process_pdf() para fuentes tipo log (ej. Shipping Order):
    no hay OCR ni paginas, la hoja completa se clasifica de una. Si el tipo
    tiene per_row (rules.yaml), se extrae un record por fila de datos; si
    no, un unico record para toda la hoja.
    Sin cache: leer un xlsx no cuesta nada comparado con OCR.
    """
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    items = sheet_items(wb.active)
    wb.close()

    cls = classify(items, rules["doc_types"])
    cfg = rules["doc_types"][cls["type"]] if cls["type"] else {}
    header_row = _xlsx_header_row(items, cfg) if cfg.get("per_row") else None
    if header_row is None:
        return {"file": Path(path).name, "pages": [_xlsx_record(items, 1, cls, cfg)]}

    # per_row: un log con registros de muchos embarques -> un record por
    # fila. Cada fila ve los headers (para que xlsx_col encuentre su
    # columna) y solo su propia fila de datos. 'page' es el numero de fila
    # de Excel: identifica el registro igual que una pagina en un PDF.
    header = [it for it in items if it["y"] <= header_row]
    data_rows = sorted({it["y"] for it in items if it["y"] > header_row})
    pages = [_xlsx_record(header + [it for it in items if it["y"] == r], r, cls, cfg)
             for r in data_rows]
    return {"file": Path(path).name, "pages": pages}


def _xlsx_header_row(items, cfg):
    """Fila de headers = la mas baja donde aparece algun header de xlsx_col."""
    rows = [lab["y"] for spec in cfg.get("fields", {}).values()
            if spec.get("strategy") == "xlsx_col"
            for lab in [_find_label(items, spec["header"])] if lab]
    return max(rows) if rows else None


def _xlsx_record(items, row, cls, cfg):
    rec = {"page": row, "source": "xlsx", "render_scale": 1.0,
           "triage": {"skip": False, "median_box_h": None, "n_boxes": len(items)},
           "classification": cls, "fields": {}}
    for fname, spec in cfg.get("fields", {}).items():
        f = resolve_field(items, spec)
        f = corroborate(items, f, spec)
        f["desc_zh"] = spec.get("desc_zh", "")
        rec["fields"][fname] = f
    return rec


def process_pdf(path, rules):
    doc = pymupdf.open(path)
    cache = _load_cache(path)
    cache_dirty = False
    pages = []
    boilerplate_pages = KNOWN_BOILERPLATE_PAGES.get(Path(path).name, set())
    for i, page in enumerate(doc, start=1):
        if i in boilerplate_pages:
            pages.append({
                "page": i, "source": "skipped", "render_scale": 1.0,
                "rotation": 0, "fields": {},
                "triage": {"skip": True,
                          "reason": "known boilerplate, see KNOWN_BOILERPLATE_PAGES"},
                "classification": {"type": None, "score": 0.0, "scores": {},
                                   "status": "SKIPPED"},
            })
            continue
        cached = cache.get(str(i))
        if cached is not None:
            items, triage = cached["items"], cached["triage"]
        else:
            items, triage = page_items(page)
            cache[str(i)] = {"items": items, "triage": triage}
            cache_dirty = True
        src = items[0]["source"] if items else ("skipped" if triage["skip"]
                                               else "empty")
        rec = {"page": i, "source": src,
               # factor para llevar las coords de 'origin' a un render a OCR_DPI:
               # el texto nativo viene en puntos PDF (72dpi), el OCR ya en pixeles
               "render_scale": (OCR_DPI / 72) if src == "native" else 1.0,
               # +90/-90 si la pagina vino escaneada de costado y se corrigio
               # (ver _try_rotation_fix) -- el preview de la GUI necesita
               # reproducir el mismo giro para recortar el lugar correcto.
               "rotation": triage.get("rotation", 0),
               "triage": triage, "fields": {}}

        if triage["skip"]:
            # se registra explicitamente: una pagina salteada en silencio es
            # una pagina que nadie puede auditar despues
            rec["classification"] = {"type": None, "score": 0.0, "scores": {},
                                     "status": "SKIPPED"}
            pages.append(rec)
            continue

        cls = classify(items, rules["doc_types"])
        rec["classification"] = cls
        if cls["type"]:
            for fname, spec in rules["doc_types"][cls["type"]]["fields"].items():
                f = resolve_field(items, spec)
                f = corroborate(items, f, spec)
                f["desc_zh"] = spec.get("desc_zh", "")
                rec["fields"][fname] = f
        pages.append(rec)
    doc.close()
    if cache_dirty:
        _save_cache(path, cache)
    return {"file": Path(path).name, "pages": pages}


PROCESSORS = {".pdf": process_pdf, ".xlsx": process_xlsx}


def duplicate_single_files(folder, rules):
    """
    {doc_type: [archivos]} para los tipos single_file (rules.yaml) que
    aparecen en MAS de un archivo de la carpeta. Se corre ANTES de procesar:
    esos tipos son xlsx, clasificarlos es instantaneo, y si hay duplicados
    no tiene sentido pagar el OCR de los PDFs para despues no poder decidir
    cual de los dos vale.
    """
    folder = Path(folder)
    single = {dt for dt, cfg in rules["doc_types"].items() if cfg.get("single_file")}
    found = {}
    for p in sorted(folder.rglob("*.xlsx")):
        if p.name.startswith("~$"):
            continue
        try:
            wb = openpyxl.load_workbook(p, data_only=True, read_only=True)
            items = sheet_items(wb.active)
            wb.close()
        except Exception:
            continue                          # un xlsx roto lo reporta process_folder
        dt = classify(items, rules["doc_types"])["type"]
        if dt in single:
            found.setdefault(dt, []).append(str(p.relative_to(folder)))
    return {dt: files for dt, files in found.items() if len(files) > 1}


def process_folder(folder, rules, on_progress=None):
    """
    Procesa todos los PDFs y xlsx de la carpeta, incluidas subcarpetas
    (una por numero de orden). Un archivo roto no corta la corrida.

    on_progress(hechos, total, nombre, resultado_o_None) permite que la GUI
    muestre avance sin congelarse: se llama antes y despues de cada archivo.
    """
    folder = Path(folder)
    # "~$archivo.xlsx" es el lock file que crea Excel al tener el archivo
    # abierto -- mismo directorio, misma extension, pero no es un xlsx valido.
    files = sorted(p for p in folder.rglob("*")
                   if p.suffix.lower() in PROCESSORS
                   and not p.name.startswith("~$"))
    results = []

    for path in files:
        rel = str(path.relative_to(folder))
        if on_progress:
            on_progress(len(results), len(files), rel, None)

        t0 = time.perf_counter()
        try:
            r = PROCESSORS[path.suffix.lower()](path, rules)
            r["ok"] = True
        except Exception as e:
            import traceback
            r = {"file": path.name, "ok": False,
                 "error": f"{type(e).__name__}: {e}",
                 "traceback": traceback.format_exc(), "pages": []}
        r["path"] = rel
        r["abs_path"] = str(path)
        r["secs"] = round(time.perf_counter() - t0, 2)
        results.append(r)

        if on_progress:
            on_progress(len(results), len(files), rel, r)
        else:
            print(f"[{len(results)}/{len(files)}] {'OK ' if r['ok'] else 'ERR'} "
                  f"{rel}  ({r['secs']}s)", flush=True)
    return results


def print_report(results):
    for r in results:
        print(f"\n{'=' * 78}\n{r['path']}  ({r['secs']}s)")
        if not r["ok"]:
            print(f"  ERROR: {r['error']}")
            continue
        for p in r["pages"]:
            c = p["classification"]
            tipo = c["type"] or "SIN IDENTIFICAR"
            print(f"\n  -- pag {p['page']} [{p['source']}] -> {tipo} "
                  f"(score {c['score']:.2f}, {c['status']})")
            for k, v in p["fields"].items():
                val = v["value"]
                if isinstance(val, dict):
                    val = " | ".join(f"{a}={b}" for a, b in val.items())
                elif isinstance(val, list):
                    val = " ; ".join(map(str, val))
                extra = f" [code={v['code']}]" if v.get("code") else ""
                print(f"     {k:22} {str(val)[:58]:58} "
                      f"conf={v['conf']:.2f} {v['status']}{extra}")


if __name__ == "__main__":
    import json
    import sys

    HERE = Path(__file__).resolve().parent
    rules = yaml.safe_load(open(HERE / "rules.yaml", encoding="utf-8"))

    # por defecto: ../samples respecto de src/;  o el path que le pases
    target = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE.parent / "samples"

    if target.is_file():
        if target.suffix.lower() not in PROCESSORS:
            sys.exit(f"Extension no soportada: {target.suffix}")
        out = [PROCESSORS[target.suffix.lower()](target, rules)]
        out[0].update(ok=True, path=target.name, secs=0)
    elif target.is_dir():
        out = process_folder(target, rules)
        if not out:
            sys.exit(f"No hay PDFs ni xlsx en: {target}")
    else:
        sys.exit(f"No existe: {target}")

    print_report(out)
    (HERE.parent / "output").mkdir(exist_ok=True)
    dest = HERE.parent / "output" / "extraction.json"
    json.dump(out, open(dest, "w", encoding="utf-8"),
              ensure_ascii=False, indent=2)
    print(f"\nJSON completo (con coordenadas de origen): {dest}")