"""
Las palabras que ve el usuario. Un solo lugar para la pantalla (gui.py) y el
Excel (organize.py): asi dicen siempre lo mismo.

Los modulos trabajan con codigos (PRESENT, NOT_VERIFIABLE...). Al usuario se
le muestra una palabra simple y una frase que la explica.
"""

from functools import lru_cache

import reconciliation

# codigo -> (palabra, frase que la explica)
DOC_STATUS = {
    "PRESENT": ("Found", "the document is in the folder."),
    "MISSING": ("Missing", "the document was not found in the folder."),
    "MANUAL": ("Check by hand", "the tool cannot read this document. "
                                "Open it and check it by hand."),
    "OPTIONAL": ("Not required", "not found, and not required for this type of export."),
    "NA": ("Not applicable", "does not apply to this type of export."),
}
RULE_RESULT = {
    "VERIFIED": ("Match", "the values are the same."),
    "DISCREPANCY": ("Does not match", "the values are different."),
    "NOT_VERIFIABLE": ("Could not check", "a document is missing or a value "
                                          "could not be read."),
    "INFO": ("Info only", "shown side by side, not compared (team's decision)."),
    "NA": ("Not applicable", "this rule does not apply to this type of export."),
}
LEG_NOTES = {"LOSSY": "digits lost in the Excel file", "LOW_CONF": "unclear reading"}
KIND_TEXT = {"liquid": "liquid", "solid": "solid (containers)",
             "unknown": "type unknown"}

NO_EXPORT, NOT_READ, UNKNOWN_FORMAT = "No export", "Not read", "Unknown format"
# "Not read" solo, no dice por que: donde hay lugar (lista de la pantalla, hoja
# del Excel) va con la causa. Es un archivo que no abrio; un documento que
# abrio pero viene en otro formato es UNKNOWN_FORMAT.
NOT_READ_FULL = f"{NOT_READ} (broken file)"

# Por que no abrio un archivo, en palabras simples: (como empieza el error de
# la libreria que lo abre -- PyMuPDF, openpyxl, xlrd --, frase). El error
# tecnico completo queda en la hoja "Not read (broken file)" del Excel.
NOT_READ_WHY = [
    ("FileDataError", "the file is damaged or is not a real PDF"),
    ("EmptyFileError", "the file is empty (0 bytes)"),
    ("ValueError: document closed or encrypted", "the PDF is password protected"),
    ("BadZipFile", "the file is damaged, password protected or not a real Excel file"),
    ("XLRDError", "the file is damaged, password protected or not a real Excel file"),
    ("PermissionError", "the file is open in another program, or you have no "
                        "permission to read it"),
]


def not_read_why(error):
    """Frase simple para el error de un archivo que no abrio."""
    return next((why for start, why in NOT_READ_WHY if (error or "").startswith(start)),
                "the tool failed while reading this file")

# (fondo, texto)
GOOD, BAD = ("#F0FDF4", "#166534"), ("#FEF2F2", "#991B1B")
WARN, GRAY = ("#FFEDD5", "#C2410C"), ("#F1F5F9", "#475569")
WORD_COLORS = {"Found": GOOD, "Match": GOOD,
               "Missing": BAD, "Does not match": BAD, NOT_READ: BAD,
               "Check by hand": WARN, "Could not check": WARN, NO_EXPORT: WARN,
               UNKNOWN_FORMAT: WARN}

# nombre corto de cada documento, con su codigo de rules.xlsx adelante
DOC_LABELS = {
    "refund_record": "Refund record",
    "tax_filing_directory": "3.1 Filing Directory",
    "customs_declaration": "3.2 CDF",
    "export_invoice": "3.3 VAT invoice",
    "bank_receipt": "3.4 Bank Receipt",
    "sales_contract": "3.5 Sales Contract",
    "customs_clearance_expenses": "3.6 Customs Clearance Expenses",
    "bill_of_lading": "3.7.1 Bill of Lading",
    "bill_of_lading_port": "3.7.1 Bill of Lading",
    "bill_of_lading_tanker": "3.7.1 Bill of Lading",
    "cargo_manifest": "3.7.2 Cargo Manifest",
    "cargo_manifest_detail": "3.7.2 Cargo Manifest",
    "cargo_manifest_tanker": "3.7.2 Cargo Manifest",
    "transport_invoice": "3.7.4.1 Transportation Invoice",
    "transport_statement": "3.7.4.2 Transportation Statement",
    "shipping_order": "3.7.4.3 Shipping Order",
    "delivery_note": "3.7.4.4 Delivery Note",
    "coordination_fee_invoice": "3.7.5.1 Coordination Fee Invoice",
    "coordination_fee_statement": "3.7.5.2 Coordination Fee Statement",
    "customs_fee_invoice": "3.8.2.1 Customs Declaration Fee Invoice",
    "customs_fee_statement": "3.8.2.2 Customs Declaration Fee Statement",
}


def doc_label(doc_type):
    return DOC_LABELS.get(doc_type) or str(doc_type).replace("_", " ").capitalize()


def plural(n, one, many):
    return f"{n} {one if n == 1 else many}"


@lru_cache(maxsize=1)
def _rule_titles():
    """Nombre corto de cada regla (reconciliation.yaml). 'title_liquid' es el de
    las exportaciones de liquidos cuando la regla usa otro documento."""
    titles = {}
    for r in reconciliation.load_rules():
        if r.get("title"):
            titles[(r["id"], None)] = r["title"]
        if r.get("title_liquid"):
            titles[(r["id"], "liquid")] = r["title_liquid"]
    return titles


def rule_title(rule, kind=None):
    """rule = una regla ya evaluada de reconciliation.run()."""
    titles = _rule_titles()
    return titles.get((rule["id"], kind)) or titles.get((rule["id"], None)) or rule["desc"]


def sources_text(sources, sep="   |   "):
    """[{file, page}] -> '3.7.4.3 Shipping Order.xlsx (rows 3, 4)'. En un
    xlsx 'page' es el numero de fila (ver extract.process_xlsx)."""
    by_file = {}
    for s in sources:
        by_file.setdefault(s["file"], []).append(s["page"])
    parts = []
    for f, pages in by_file.items():
        excel = f.lower().endswith((".xlsx", ".xls"))
        unit = ("row" if len(pages) == 1 else "rows") if excel else "p."
        parts.append(f"{f} ({unit} {', '.join(map(str, sorted(pages)))})")
    return sep.join(parts)


def leg_value(leg):
    """El valor de una pata, sin '.0' en enteros y con la unidad pegada
    (51000.0 en kg se lee '51000 kg')."""
    v = leg["value"]
    if isinstance(v, list):
        text = ", ".join(str(x) for x in v)
    elif isinstance(v, dict):
        text = ", ".join(f"{k}={x}" for k, x in v.items())
    else:
        if isinstance(v, float) and v.is_integer():
            v = int(v)
        text = "" if v is None else str(v)
    unit = leg.get("unit")
    return f"{text} {unit}" if unit else text


def leg_lines(legs, with_item=False, only_values=False):
    """
    Una linea por documento de la regla: '3.2 CDF: 51000 kg'.
    Un documento con varios formatos (hay 3 de Bill of Lading) tiene una pata
    por formato: se muestra solo la que trajo valor, no las otras vacias.
    with_item agrega el item de rules.xlsx (1-84), para ir directo a esa fila.
    only_values deja solo los valores: sin los documentos que no trajeron
    ninguno y sin la nota de lectura dudosa.
    """
    groups = {}
    for leg in legs:
        label = doc_label(leg["ref"].split(".")[0]) if leg.get("ref") else "(no reader)"
        groups.setdefault(label, []).append(leg)
    lines = []
    for label, group in groups.items():
        found = [l for l in group if l["status"] != "MISSING"]
        for leg in found or ([] if only_values else group[:1]):
            name = label
            if with_item and leg.get("item") is not None:
                name = f"{label}, item {leg['item']}"
            if leg["status"] == "MISSING":
                lines.append(f"{name}: no value")
                continue
            note = None if only_values else LEG_NOTES.get(leg["status"])
            lines.append(f"{name}: {leg_value(leg)}" + (f" ({note})" if note else ""))
    return lines


def rule_lines(rule):
    """Los valores de una regla, uno por documento. Una regla que no aplica no
    muestra nada."""
    return [] if rule["verdict"] == "NA" else leg_lines(rule["legs"], with_item=True)
