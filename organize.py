"""
Salida separada por embarque: una carpeta por exportacion con copia de sus
documentos y un reporte Excel (resumen, checklist, reconciliacion y datos
extraidos).

    output/exports/<fecha_hora>/consolidated_report.xlsx   (todas las exportaciones)
    output/exports/<fecha_hora>/<SO>/documents/...
    output/exports/<fecha_hora>/<SO>/<SO>_report.xlsx
    output/exports/<fecha_hora>/_unassigned/...      (si hay paginas sin asignar)

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

HERE = Path(__file__).resolve().parent
UNASSIGNED_DIR = "_unassigned"

# (fondo, texto) -- mismos colores que la GUI
STATUS_COLORS = {
    "PRESENT": ("F0FDF4", "166534"), "VERIFIED": ("F0FDF4", "166534"),
    "OK": ("F0FDF4", "166534"), "OK_CORROBORATED": ("F0FDF4", "166534"),
    "CORRECTED": ("EFF6FF", "1D4ED8"),
    "MISSING": ("FEF2F2", "991B1B"), "DISCREPANCY": ("FEF2F2", "991B1B"),
    "NOT_FOUND": ("FEF2F2", "991B1B"), "CONFLICT": ("FEF2F2", "991B1B"),
    "MANUAL": ("FFEDD5", "C2410C"), "NOT_VERIFIABLE": ("FFEDD5", "C2410C"),
    "PARTIAL": ("FFEDD5", "C2410C"), "MULTI": ("FFEDD5", "C2410C"),
    "INFO": ("F1F5F9", "475569"), "OPTIONAL": ("F1F5F9", "475569"),
    "NA": ("F1F5F9", "475569"),
}
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


def _sources(sources):
    by_file = {}
    for s in sources:
        by_file.setdefault(s["file"], []).append(s["page"])
    return " | ".join(f"{f} ({', '.join(map(str, sorted(p)))})"
                      for f, p in by_file.items())


def _leg(leg):
    ref = leg.get("ref") or f"item {leg.get('item')} (no extractor)"
    if leg["status"] == "MISSING":
        return f"{ref}: not available"
    return f"{ref}: {_fmt(leg['value'])} [{leg['status']}]"


def _field_rows(records, pages_by_key):
    rows = []
    for rec in sorted(records, key=lambda r: (r["file"], r["page"])):
        page = pages_by_key.get((rec["file"], rec["page"]))
        if not page:
            continue
        if not page["fields"]:
            rows.append([rec["file"], rec["page"], rec["doc_type"], "", "", "", "OK"])
        for fname, f in page["fields"].items():
            rows.append([rec["file"], rec["page"], rec["doc_type"], fname,
                         _fmt(f["value"]), round(f["conf"], 2), f["status"]])
    return rows


def write_report(path, shipment, source_folder, pages_by_key, log_files,
                 folder_warnings=()):
    ck = shipment["checklist"]
    count = lambda items, key, val: sum(1 for i in items if i[key] == val)
    wb = Workbook()
    summary = [
        ["Shipment", shipment["shipment"]],
        ["Export type", shipment.get("kind") or "unknown"],
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Source folder", str(source_folder)],
        ["Documents present", count(ck, "status", "PRESENT")],
        ["Documents missing", count(ck, "status", "MISSING")],
        ["Documents to check by hand", count(ck, "status", "MANUAL")],
        ["Rules verified", count(shipment["rules"], "verdict", "VERIFIED")],
        ["Rules with discrepancy", count(shipment["rules"], "verdict", "DISCREPANCY")],
        ["Rules not verifiable", count(shipment["rules"], "verdict", "NOT_VERIFIABLE")],
        ["Rules shown for information", count(shipment["rules"], "verdict", "INFO")],
        ["Rules not applicable", count(shipment["rules"], "verdict", "NA")],
        ["Warnings", "\n".join(shipment["warnings"]) or "none"],
        ["Folder warnings", "\n".join(folder_warnings) or "none"],
        ["Shared logs (not copied)",
         "\n".join(log_files) + "\nTheir rows for this shipment are in 'Extracted data'."
         if log_files else "none"],
    ]
    _sheet(wb, "Summary", ["Item", "Value"], summary, (30, 90), first=True)
    _sheet(wb, "Checklist", ["Code", "Document", "Status", "Found in"],
           [[i["code"], i["name"], i["status"], _sources(i["sources"])] for i in ck],
           (10, 60, 12, 70), status_col=2)
    _sheet(wb, "Reconciliation", ["Rule", "Description", "Verdict", "Detail"],
           [[r["id"], r["desc"], r["verdict"], "\n".join(_leg(l) for l in r["legs"])]
            for r in shipment["rules"]],
           (6, 60, 16, 80), status_col=2)
    _sheet(wb, "Extracted data",
           ["File", "Page / row", "Document type", "Field", "Value", "Confidence", "Status"],
           _field_rows(shipment["records"], pages_by_key),
           (34, 10, 24, 24, 46, 11, 18), status_col=6)
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
    """[archivo, paginas] de lo que no esta en rules.xlsx y no se analizo."""
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
      Overview        -> una fila por exportacion: lo reclamado, cuantos
                         documentos faltan, cuantas reglas fallan, avisos
      Discrepancies   -> solo las reglas con DISCREPANCY, con el detalle
      Missing docs    -> solo los documentos MISSING
      All rules       -> todas las reglas de todas las exportaciones
      Other           -> paginas que no son documentos de rules.xlsx
      Unknown format  -> documentos de rules.xlsx en un formato que no se lee
    El detalle de cada una sigue en su carpeta (<SO>_report.xlsx).
    """
    count = lambda items, key, val: sum(1 for i in items if i[key] == val)
    overview = []
    for s in shipments:
        ck, rules = s["checklist"], s["rules"]
        overview.append([
            s["shipment"],
            s.get("kind") or "",
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
                ["Shipment", "Type", "CDF (refund record)", "Refund period", "USD (refund record)",
                 "VAT refunded CNY", "Docs present", "Docs missing", "Missing codes",
                 "Rules verified", "Discrepancies", "Discrepancy rules",
                 "Not verifiable", "Warnings"],
                overview, (22, 9, 22, 12, 16, 16, 10, 10, 24, 10, 12, 14, 12, 70), first=True)
    # una exportacion con discrepancias o documentos faltantes en rojo
    for row in ws.iter_rows(min_row=2):
        if row[10].value or row[7].value:
            for cell in row:
                cell.fill = PatternFill("solid", fgColor=STATUS_COLORS["DISCREPANCY"][0])
                cell.font = Font(color=STATUS_COLORS["DISCREPANCY"][1])
    _sheet(wb, "Discrepancies", ["Shipment", "Rule", "Description", "Detail"],
           [[s["shipment"], r["id"], r["desc"], "\n".join(_leg(l) for l in r["legs"])]
            for s in shipments for r in s["rules"] if r["verdict"] == "DISCREPANCY"],
           (22, 6, 60, 80))
    # una exportacion sin NINGUN documento va en una sola fila, no 16
    missing = []
    for s in shipments:
        if _no_documents(s):
            missing.append([s["shipment"], "ALL", "claimed in the refund record, "
                                                  "no documents found in the folder"])
            continue
        missing += [[s["shipment"], i["code"], i["name"]]
                    for i in s["checklist"] if i["status"] == "MISSING"]
    _sheet(wb, "Missing docs", ["Shipment", "Code", "Document"], missing, (22, 10, 70))
    _sheet(wb, "All rules", ["Shipment", "Rule", "Description", "Verdict", "Detail"],
           [[s["shipment"], r["id"], r["desc"], r["verdict"],
             "\n".join(_leg(l) for l in r["legs"])]
            for s in shipments for r in s["rules"]],
           (22, 6, 60, 16, 80), status_col=3)
    _sheet(wb, "Other", ["File", "Pages not analyzed"], list(other), (80, 40))
    _sheet(wb, "Unknown format", ["File", "Page", "Looks like"],
           [[u["file"], u["page"], u["doc_type"]] for u in unsupported], (80, 8, 30))
    _sheet(wb, "Run info", ["Item", "Value"], [
        ["Generated", datetime.now().strftime("%Y-%m-%d %H:%M")],
        ["Source folder", str(source_folder)],
        ["Shipments", len(shipments)],
        ["Unassigned pages", len(unassigned)],
        ["Files with Other pages", len(other)],
        ["Pages in an unknown format", len(unsupported)],
        ["Folder warnings", "\n".join(folder_warnings) or "none"],
    ], (24, 90))
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
        d = run_dir / UNASSIGNED_DIR
        copy(sorted({u["file"] for u in recon["unassigned"]}), d / "documents")
        wb = Workbook()
        _sheet(wb, "Unassigned pages", ["File", "Page / row", "Document type"],
               [[u["file"], u["page"], u["doc_type"]] for u in recon["unassigned"]],
               (50, 10, 30), first=True)
        wb.save(d / "unassigned_pages.xlsx")
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
