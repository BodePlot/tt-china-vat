"""
Salida separada por embarque: una carpeta por exportacion con copia de sus
documentos y un reporte Excel (resumen, checklist, reconciliacion y datos
extraidos).

    output/exports/<fecha_hora>/consolidated_report.xlsx   (todas las exportaciones)
    output/exports/<fecha_hora>/<SO>/documents/...
    output/exports/<fecha_hora>/<SO>/<SO>_report.xlsx
    output/exports/<fecha_hora>/_no_export/...       (si hay paginas sin exportacion)

Las palabras (Found, Missing, Match...) salen de words.py: son las mismas que
muestra la pantalla.

Cada corrida va a su propia carpeta con fecha y hora: nunca se pisa ni se
borra una salida anterior. Los originales no se tocan (se copian).

Los logs compartidos (per_row en rules.yaml: Shipping Order, Tax Filing
Directory) NO se copian a cada carpeta -- traen filas de todos los
embarques. Sus filas de ESTE embarque aparecen en la hoja "Extracted data".

Correr solo:  python organize.py [carpeta]      (por defecto samples/)
"""

import re
import shutil
from datetime import datetime
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill

import words
from words import BAD, GOOD, GRAY, WARN

HERE = Path(__file__).resolve().parent
NO_EXPORT_DIR = "_no_export"

# palabra -> (fondo, texto), los mismos colores que la pantalla. Las palabras
# sin color propio (Not required, Info only...) van en gris.
_words = [w for w, _ in (*words.DOC_STATUS.values(), *words.RULE_RESULT.values())]
# estado de cada campo leido (hoja "Extracted data"): ahi van los codigos tal cual
_field_status = {"OK": GOOD, "OK_CORROBORATED": GOOD, "CORRECTED": ("#EFF6FF", "#1D4ED8"),
                 "NOT_FOUND": BAD, "CONFLICT": BAD, "PARTIAL": WARN, "MULTI": WARN}
STATUS_COLORS = {k: (bg.lstrip("#"), fg.lstrip("#")) for k, (bg, fg) in
                 {**dict.fromkeys(_words, GRAY), **words.WORD_COLORS, **_field_status}.items()}
CONSOLIDATED_NAME = "consolidated_report.xlsx"


def _safe_name(text, limit=60):
    """Nombre de carpeta valido en Windows a partir de la etiqueta del embarque."""
    name = re.sub(r'[<>:"/\\|?*\s]+', "_", str(text)).strip("._")
    return (name or "shipment")[:limit]


def _fmt(v):
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, list):
        return ", ".join(map(str, v))
    if isinstance(v, dict):
        return ", ".join(f"{k}={x}" for k, x in v.items())
    return "" if v is None else str(v)


def _sheet(wb, title, head, rows, widths, status_col=None, first=False):
    ws = wb.active if first else wb.create_sheet()
    ws.title = title
    ws.append(head)
    for c in ws[1]:
        c.font = Font(bold=True)
    for r in rows:
        ws.append(r)
    ws.freeze_panes = "A2"
    for i, w in enumerate(widths):
        ws.column_dimensions[chr(ord("A") + i)].width = w
    for row in ws.iter_rows(min_row=2):
        colors = STATUS_COLORS.get(str(row[status_col].value)) if status_col is not None else None
        for cell in row:
            cell.alignment = Alignment(vertical="top", wrap_text=True)
            if colors:
                cell.fill = PatternFill("solid", fgColor=colors[0])
                cell.font = Font(color=colors[1])
    return ws


def _doc_word(item):
    return words.DOC_STATUS[item["status"]][0]


def _rule_word(rule):
    return words.RULE_RESULT[rule["verdict"]][0]


def _rule_values(rule):
    return "\n".join(words.rule_lines(rule))


def _field_rows(records, pages_by_key):
    rows = []
    for rec in sorted(records, key=lambda r: (r["file"], r["page"])):
        page = pages_by_key.get((rec["file"], rec["page"]))
        if not page:
            continue
        if not page["fields"]:
            rows.append([rec["file"], rec["page"], words.doc_label(rec["doc_type"]),
                         "", "", "", "OK"])
        for fname, f in page["fields"].items():
            rows.append([rec["file"], rec["page"], words.doc_label(rec["doc_type"]), fname,
                         _fmt(f["value"]), round(f["conf"], 2), f["status"]])
    return rows


def write_report(path, shipment, source_folder, pages_by_key, log_files,
                 folder_warnings=()):
    ck = shipment["checklist"]
    count = lambda items, key, val: sum(1 for i in items if i[key] == val)
    wb = Workbook()
    rules, kind = shipment["rules"], shipment.get("kind")
    summary = [
        ["Export", shipment["shipment"]],
        ["Export type", words.KIND_TEXT.get(kind, "type unknown")],
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Source folder", str(source_folder)],
        ["Documents found", count(ck, "status", "PRESENT")],
        ["Documents missing", count(ck, "status", "MISSING")],
        ["Documents to check by hand", count(ck, "status", "MANUAL")],
        ["Rules that match", count(rules, "verdict", "VERIFIED")],
        ["Rules that do not match", count(rules, "verdict", "DISCREPANCY")],
        ["Rules that could not be checked", count(rules, "verdict", "NOT_VERIFIABLE")],
        ["Rules shown as info only", count(rules, "verdict", "INFO")],
        ["Rules not applicable", count(rules, "verdict", "NA")],
        ["Notes", "\n".join(shipment["warnings"]) or "none"],
        ["Notes about the folder", "\n".join(folder_warnings) or "none"],
        ["Shared files (not copied)",
         "\n".join(log_files) + "\nTheir rows for this export are in 'Extracted data'."
         if log_files else "none"],
    ]
    _sheet(wb, "Summary", ["Item", "Value"], summary, (30, 90), first=True)
    # mismas columnas y mismas palabras que la pantalla
    _sheet(wb, "Documents", ["Code", "Document", "Status", "Found in"],
           [[i["code"], i["name"], _doc_word(i), words.sources_text(i["sources"], sep="\n")]
            for i in ck],
           (10, 60, 16, 70), status_col=2)
    _sheet(wb, "Rules", ["Rule", "What is checked", "Result", "Values"],
           [[r["id"], words.rule_title(r, kind), _rule_word(r), _rule_values(r)]
            for r in rules],
           (6, 60, 16, 80), status_col=2)
    _sheet(wb, "Extracted data",
           ["File", "Page / row", "Document", "Field", "Value", "Confidence", "Status"],
           _field_rows(shipment["records"], pages_by_key),
           (34, 10, 30, 24, 46, 11, 18), status_col=6)
    wb.save(path)


def _claim_values(shipment, pages_by_key, field):
    """Valores de un campo del refund record en las filas de este embarque."""
    vals = []
    for rec in shipment["records"]:
        if rec["doc_type"] != "refund_record":
            continue
        page = pages_by_key.get((rec["file"], rec["page"]))
        v = page and page["fields"].get(field, {}).get("value")
        if isinstance(v, float):
            v = round(v, 2)                    # 70231.04999999999 -> 70231.05
        if v is not None and v not in vals:
            vals.append(v)
    # un solo valor va como numero (Excel lo puede sumar); varios, como texto
    return vals[0] if len(vals) == 1 else _fmt(vals)


def _no_documents(shipment):
    """Exportacion que solo esta en el refund record, sin ningun documento."""
    return all(rec["doc_type"] == "refund_record" for rec in shipment["records"])


def _other_rows(results):
    """[archivo, paginas] de lo que no es un documento de rules.xlsx."""
    rows = []
    for r in results:
        pages = [p["page"] for p in r.get("pages", [])
                 if p["classification"]["status"] == "OTHER"]
        if pages:
            rows.append([r["path"], ", ".join(map(str, pages))])
    return rows


def write_consolidated(path, shipments, source_folder, pages_by_key,
                       unassigned=(), folder_warnings=(), other=(), unsupported=()):
    """
    Todas las exportaciones en un solo Excel (pedido del equipo de China,
    2026-09-29), para no tener que abrir el reporte de cada una:
      Overview                -> una fila por exportacion: lo reclamado, cuantos
                                 documentos faltan, cuantas reglas no coinciden
      Rules that do not match -> solo esas reglas, con sus valores
      Missing documents       -> solo los documentos que faltan
      All rules               -> todas las reglas de todas las exportaciones
      Other                   -> paginas que no son documentos de rules.xlsx
      Unknown format          -> documentos de rules.xlsx en un formato que no se lee
    El detalle de cada una sigue en su carpeta (<SO>_report.xlsx).
    """
    count = lambda items, key, val: sum(1 for i in items if i[key] == val)
    overview = []
    for s in shipments:
        ck, rules = s["checklist"], s["rules"]
        overview.append([
            s["shipment"],
            words.KIND_TEXT.get(s.get("kind"), ""),
            _claim_values(s, pages_by_key, "cdf_no"),
            _claim_values(s, pages_by_key, "refund_period"),
            _claim_values(s, pages_by_key, "amount_usd"),
            _claim_values(s, pages_by_key, "vat_refunded"),
            count(ck, "status", "PRESENT"), count(ck, "status", "MISSING"),
            ", ".join(i["code"] for i in ck if i["status"] == "MISSING"),
            count(rules, "verdict", "VERIFIED"), count(rules, "verdict", "DISCREPANCY"),
            ", ".join(str(r["id"]) for r in rules if r["verdict"] == "DISCREPANCY"),
            count(rules, "verdict", "NOT_VERIFIABLE"),
            "\n".join(s["warnings"]),
        ])
    wb = Workbook()
    ws = _sheet(wb, "Overview",
                ["Export", "Type", "CDF (refund record)", "Refund period",
                 "USD (refund record)", "VAT refunded CNY", "Documents found",
                 "Documents missing", "Missing codes", "Rules that match",
                 "Rules that do not match", "Rule numbers",
                 "Rules that could not be checked", "Notes"],
                overview, (22, 16, 22, 12, 16, 16, 11, 11, 24, 10, 12, 14, 14, 70), first=True)
    # una exportacion con reglas que no coinciden o documentos faltantes, en rojo
    for row in ws.iter_rows(min_row=2):
        if row[10].value or row[7].value:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=STATUS_COLORS["Missing"][0])
                cell.font = Font(color=STATUS_COLORS["Missing"][1])
    _sheet(wb, "Rules that do not match", ["Export", "Rule", "What is checked", "Values"],
           [[s["shipment"], r["id"], words.rule_title(r, s.get("kind")), _rule_values(r)]
            for s in shipments for r in s["rules"] if r["verdict"] == "DISCREPANCY"],
           (22, 6, 60, 80))
    # una exportacion sin NINGUN documento va en una sola fila, no 16
    missing = []
    for s in shipments:
        if _no_documents(s):
            missing.append([s["shipment"], "all", "in the refund record, but none of "
                                                  "its documents are in the folder"])
            continue
        missing += [[s["shipment"], i["code"], i["name"]]
                    for i in s["checklist"] if i["status"] == "MISSING"]
    _sheet(wb, "Missing documents", ["Export", "Code", "Document"], missing, (22, 10, 70))
    _sheet(wb, "All rules", ["Export", "Rule", "What is checked", "Result", "Values"],
           [[s["shipment"], r["id"], words.rule_title(r, s.get("kind")), _rule_word(r),
             _rule_values(r)]
            for s in shipments for r in s["rules"]],
           (22, 6, 60, 16, 80), status_col=3)
    _sheet(wb, "Other", ["File (not a document of rules.xlsx)", "Pages"], list(other), (80, 40))
    _sheet(wb, words.UNKNOWN_FORMAT, ["File", "Page", "Looks like"],
           [[u["file"], u["page"], words.doc_label(u["doc_type"])] for u in unsupported],
           (80, 8, 40))
    _sheet(wb, "Run info", ["Item", "Value"], [
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Source folder", str(source_folder)],
        ["Exports", len(shipments)],
        ["Pages that belong to no export", len(unassigned)],
        ["Files with Other pages", len(other)],
        ["Pages in an unknown format", len(unsupported)],
        ["Notes about the folder", "\n".join(folder_warnings) or "none"],
    ], (30, 90))
    wb.save(path)


def export_by_shipment(results, recon, doc_types, source_folder, dest_root=None):
    """
    Arma la salida por embarque. Devuelve la carpeta de la corrida.
    results = extract.process_folder(), recon = reconciliation.run().
    """
    dest_root = Path(dest_root) if dest_root else HERE / "output" / "exports"
    run_dir = dest_root / datetime.now().strftime("%Y-%m-%d_%H%M%S")
    run_dir.mkdir(parents=True, exist_ok=False)

    abs_of = {r["path"]: r.get("abs_path") for r in results}
    pages_by_key = {(r["path"], p["page"]): p
                    for r in results if r.get("ok") for p in r["pages"]}
    is_log = lambda dt: bool(doc_types.get(dt, {}).get("per_row"))

    def copy(files, docs_dir):
        docs_dir.mkdir(parents=True, exist_ok=True)
        for f in files:
            src = abs_of.get(f)
            if src and Path(src).is_file():
                # la estructura de subcarpetas de origen se aplana; si dos
                # archivos se llaman igual, el segundo lleva el prefijo de su
                # carpeta para no pisar al primero
                dst = docs_dir / Path(f).name
                if dst.exists():
                    dst = docs_dir / _safe_name(str(Path(f).parent) + "_" + Path(f).name, 150)
                shutil.copy2(src, dst)

    for s in recon["shipments"]:
        d = run_dir / _safe_name(s["shipment"])
        own = sorted({r["file"] for r in s["records"] if not is_log(r["doc_type"])})
        logs = sorted({r["file"] for r in s["records"] if is_log(r["doc_type"])})
        copy(own, d / "documents")
        write_report(d / f"{_safe_name(s['shipment'])}_report.xlsx", s,
                     source_folder, pages_by_key, logs,
                     recon.get("folder_warnings", ()))
    write_consolidated(run_dir / CONSOLIDATED_NAME, recon["shipments"], source_folder,
                       pages_by_key, recon["unassigned"], recon.get("folder_warnings", ()),
                       _other_rows(results), recon.get("unsupported", ()))

    if recon["unassigned"]:
        d = run_dir / NO_EXPORT_DIR
        copy(sorted({u["file"] for u in recon["unassigned"]}), d / "documents")
        wb = Workbook()
        _sheet(wb, words.NO_EXPORT, ["File", "Page / row", "Document", "Why"],
               [[u["file"], u["page"], words.doc_label(u["doc_type"]), u.get("reason", "")]
                for u in recon["unassigned"]],
               (50, 10, 34, 60), first=True)
        wb.save(d / "pages_with_no_export.xlsx")
    return run_dir


if __name__ == "__main__":
    import sys

    import yaml

    import extract
    import reconciliation

    folder = Path(sys.argv[1]) if len(sys.argv) > 1 else HERE / "samples"
    rules_yaml = yaml.safe_load(open(HERE / "rules.yaml", encoding="utf-8"))
    dups = extract.duplicate_single_files(folder, rules_yaml)
    if dups:
        sys.exit("Archivos duplicados de un tipo que tiene que ser unico -- dejar uno solo:\n"
                 + "\n".join(f"  {dt}: {', '.join(files)}" for dt, files in dups.items()))
    results = extract.process_folder(folder, rules_yaml)
    recon = reconciliation.run(results, doc_types=rules_yaml["doc_types"])
    out = export_by_shipment(results, recon, rules_yaml["doc_types"], folder)
    print(f"\nSalida por embarque: {out}")
    for p in sorted(out.iterdir()):
        if p.is_dir():
            print(f"  {p.name}/  ({len(list((p / 'documents').glob('*')))} documentos)")
        else:
            print(f"  {p.name}")
