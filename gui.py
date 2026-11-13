"""
China VAT Refund - chequeo de los documentos de exportacion.

Correr:  python gui.py

Dos pestañas:
  1. Process -> elegir la carpeta y procesar
  2. Exports -> a la izquierda las exportaciones; a la derecha, de la elegida,
                sus documentos (cuales faltan) y sus reglas (cuales no coinciden)

El detalle campo por campo no se muestra aca: esta en el Excel que se exporta.
"""

import os
import queue
import threading
import time
import tkinter.font
import traceback
from pathlib import Path

import FreeSimpleGUI as sg
import yaml

import extract
import organize
import reconciliation
from words import (BAD, DOC_STATUS, GOOD, GRAY, KIND_TEXT, NO_EXPORT, NOT_READ,
                   NOT_READ_FULL, RULE_RESULT, UNKNOWN_FORMAT, WARN, WORD_COLORS, doc_label,
                   leg_lines, not_read_why, plural, rule_lines, rule_title,
                   sources_text)

HERE = Path(__file__).resolve().parent
ROOT = HERE          # el proyecto es plano: samples/ y output/ viven al lado de gui.py
DEFAULT_SAMPLES = ROOT / "samples"

C = {"primary": "#2563EB", "ok": "#15803D", "warn": "#D97706",
     "bad": "#DC2626", "bg": "#F4F6F8", "surface": "#FFFFFF",
     "text": "#1F2937", "muted": "#6B7280", "sel": "#DBEAFE"}

sg.theme_add_new("China", {
    "BACKGROUND": C["bg"], "TEXT": C["text"], "INPUT": C["surface"],
    "TEXT_INPUT": C["text"], "SCROLL": "#CBD5E1",
    "BUTTON": ("#FFFFFF", C["primary"]), "PROGRESS": (C["primary"], "#E5E7EB"),
    "BORDER": 1, "SLIDER_DEPTH": 0, "PROGRESS_DEPTH": 0})
sg.theme("China")

FONT = ("Segoe UI", 10)
BOLD = ("Segoe UI", 10, "bold")


# =============================================================================
# Datos -> filas
# =============================================================================
# Las palabras (Found, Missing, Match...) estan en words.py: las mismas que
# usa el Excel.
def _export_entry(s):
    """Una exportacion -> lo que se muestra de ella (fila de la lista + panel)."""
    count = lambda items, key, val: sum(1 for i in items if i[key] == val)
    missing = count(s["checklist"], "status", "MISSING")
    by_hand = count(s["checklist"], "status", "MANUAL")
    mismatch = count(s["rules"], "verdict", "DISCREPANCY")
    unchecked = count(s["rules"], "verdict", "NOT_VERIFIABLE")
    empty = organize._no_documents(s)

    if empty:
        summary = ("This export is in the refund record, but none of its "
                   "documents are in the folder.")
    else:
        parts = [missing and plural(missing, "document missing", "documents missing"),
                 mismatch and plural(mismatch, "rule does not match",
                                     "rules do not match"),
                 by_hand and plural(by_hand, "document to check by hand",
                                    "documents to check by hand"),
                 unchecked and plural(unchecked, "rule could not be checked",
                                      "rules could not be checked")]
        summary = ("   ·   ".join(p for p in parts if p)
                   or "All good: no documents missing and all rules match.")

    docs, doc_details = [], []
    for i in s["checklist"]:
        word, why = DOC_STATUS[i["status"]]
        docs.append([i["code"], i["name"], word, sources_text(i["sources"])])
        found = sources_text(i["sources"], sep="\n   ")
        doc_details.append(f"{i['code']}  {i['name']}\n{word}: {why}"
                           + (f"\nFound in:\n   {found}" if found else ""))

    rules, rule_details = [], []
    for r in s["rules"]:
        word, why = RULE_RESULT[r["verdict"]]
        title = rule_title(r, s.get("kind"))
        # en la tabla, una regla que coincide muestra lo que se comparo; los
        # documentos que no trajeron valor se ven al hacer click en la fila
        shown = ([] if r["verdict"] == "NA" else
                 leg_lines(r["legs"], only_values=r["verdict"] == "VERIFIED"))
        full = rule_lines(r)
        rules.append([r["id"], title, word, "   |   ".join(shown)])
        rule_details.append("\n   ".join([f"Rule {r['id']}  —  {title}\n{word}: {why}"]
                                         + full))

    # el titulo dice que es cada numero: Shipping Order (3.7.4.3) y CDF (3.2).
    # Una exportacion sin Shipping Order ya se llama "CDF ..." (reconciliation)
    kind = KIND_TEXT.get(s.get("kind"))
    if s.get("sos"):
        title = [f"Shipping Order {s['shipment']}",
                 s["cdfs"] and "CDF " + ", ".join(s["cdfs"])]
    else:
        title = [str(s["shipment"])]
    title = "   ·   ".join(p for p in title + [kind] if p)
    return {
        # la ultima columna explica el naranja: todo lo que una persona tiene que
        # mirar (documentos "Check by hand" + reglas "Could not check")
        "row": [s["shipment"], "all" if empty else missing, "" if empty else mismatch,
                "" if empty else by_hand + unchecked],
        "color": BAD if missing or mismatch or empty
        else WARN if by_hand or unchecked else GOOD,
        "title": title,
        "summary": summary, "empty": empty,
        "docs": docs, "doc_details": doc_details,
        "rules": rules, "rule_details": rule_details,
        "notes": s["warnings"],
    }


def _special_entry(label, color, title, summary, rows):
    """Entrada de la lista que no es una exportacion (paginas sueltas, archivos
    que no se pudieron leer). rows = [documento, estado, donde]."""
    return {"row": [f"{label}: {len(rows)}", "", "", ""], "color": color,
            "title": title, "summary": summary, "empty": False,
            "docs": [["", doc, word, where] for doc, word, where in rows],
            "doc_details": [f"{doc}\n{word}\n   {where}" for doc, word, where in rows],
            "rules": [], "rule_details": [], "notes": []}


def build_entries(recon, results):
    """
    Todo lo que va en la lista de la izquierda, en orden:
      las exportaciones, las paginas que no son de ninguna exportacion, los
      documentos en un formato desconocido y los archivos que no abrieron.
      Sin esas tres ultimas, un archivo roto solo se veria como un documento
      "Missing" y nadie sabria por que.
    """
    entries = [_export_entry(s) for s in recon["shipments"]]

    # paginas sin exportacion: una fila por (documento, archivo, motivo)
    loose = {}
    for u in recon["unassigned"]:
        key = (doc_label(u["doc_type"]), u["file"], u.get("reason", ""))
        loose.setdefault(key, []).append(u)
    if loose:
        rows = [(doc, NO_EXPORT, f"{sources_text(pages)}  —  {reason}")
                for (doc, _, reason), pages in loose.items()]
        entries.append(_special_entry(
            NO_EXPORT, WARN, "Pages that belong to no export",
            "The tool could not link these pages to any export. Check them by hand.",
            rows))

    # documento reconocido, pero en un formato que no se sabe leer: el archivo
    # esta bien
    unknown = {}
    for u in recon.get("unsupported", ()):
        unknown.setdefault((doc_label(u["doc_type"]), u["file"]), []).append(u)
    if unknown:
        entries.append(_special_entry(
            UNKNOWN_FORMAT, WARN, "Documents in a format the tool cannot read yet",
            "The files are fine and the tool knows which documents they are, but "
            "not this layout. Check them by hand.",
            [(doc, UNKNOWN_FORMAT, sources_text(pages))
             for (doc, _), pages in unknown.items()]))

    # el archivo no abrio: no se sabe que tiene adentro
    broken = [("(whole file)", NOT_READ,
               f"{r['path']}  —  {not_read_why(r.get('error'))}")
              for r in results if not r.get("ok")]
    if broken:
        entries.append(_special_entry(
            NOT_READ_FULL, BAD, "Files that could not be opened",
            "Broken or password-protected files: nothing was read from them. "
            "Get the files again, or check them by hand.",
            broken))
    return entries


# =============================================================================
# Tablas: recorte de texto
# =============================================================================
LIST_WIDTHS = [20, 11, 10, 12]
DOCS_WIDTHS = [8, 48, 14, 34]
RULES_WIDTHS = [5, 40, 14, 45]


def _truncate(text, px, font):
    """Recorta texto para que entre en una celda de 'px' pixeles. El Treeview
    no hace wrap ni pone '…' solo; el texto completo se ve abajo al hacer click.
    Se mide en pixeles porque la fuente es proporcional."""
    text = str(text)
    if font.measure(text) <= px:
        return text
    lo, hi = 0, len(text)
    while lo < hi:                    # el prefijo mas largo que entra con '…'
        mid = (lo + hi + 1) // 2
        if font.measure(text[:mid] + "…") <= px:
            lo = mid
        else:
            hi = mid - 1
    return text[:lo].rstrip() + "…"


def _fit_rows(rows, widths):
    """Recorta cada celda al ancho de su columna."""
    font = tkinter.font.Font(font=FONT)
    # mismo ancho que le da FreeSimpleGUI a la columna (col_width * ancho
    # de 'A' + 10), menos el padding interno de la celda
    px = [w * font.measure("A") + 10 - 12 for w in widths]
    return [[_truncate(v, px[i], font) for i, v in enumerate(row)] for row in rows]


def _row_colors(rows, col):
    """FreeSimpleGUI espera (fila, texto, fondo) -- OJO, no (fila, fondo, texto)."""
    return [(i, *reversed(WORD_COLORS.get(r[col], GRAY))) for i, r in enumerate(rows)]


# =============================================================================
# Layout
# =============================================================================
def tab_process():
    return [
        [sg.Text("Choose the folder with the documents and press Process.",
                 font=("Segoe UI", 11), pad=(0, (16, 10)))],
        [sg.Text("Folder"),
         sg.Input(str(DEFAULT_SAMPLES), key="-FOLDER-", expand_x=True),
         sg.FolderBrowse("Browse", target="-FOLDER-"),
         sg.Button("Process", key="-PROCESS-", size=(16, 1), font=BOLD)],
        [sg.Checkbox("Read every page (slower)", key="-READ_ALL-",
                     text_color=C["muted"],
                     tooltip="Use it only if a document you expected was not found.")],
        [sg.Text("●", key="-ICON-", text_color=C["muted"], font=("Segoe UI", 14)),
         sg.Text("Ready", key="-STATUS-", text_color=C["muted"], font=BOLD,
                 size=(90, 1)),
         sg.Push(),
         sg.Text("—", key="-TOTAL_TIME-", text_color=C["muted"], size=(8, 1)),
         sg.Text("0%", key="-PCT-", text_color=C["muted"])],
        [sg.ProgressBar(100, "h", size=(55, 18), key="-BAR-", expand_x=True)],
        [sg.Text("Log", text_color=C["muted"], pad=(0, (14, 0)))],
        [sg.Multiline(size=(100, 13), key="-LOG-", autoscroll=True, disabled=True,
                      expand_x=True, expand_y=True, background_color="#111827",
                      text_color="#E5E7EB", font=("Consolas", 9),
                      pad=(0, (3, 16)))],
    ]


def tab_exports():
    table = dict(auto_size_columns=False, justification="left", font=FONT,
                 header_font=BOLD, selected_row_colors=(C["text"], C["sel"]),
                 vertical_scroll_only=True, hide_vertical_scroll=False,
                 enable_events=True, select_mode=sg.TABLE_SELECT_MODE_BROWSE,
                 expand_x=False)

    # el detalle va a la izquierda, debajo de la lista: a la derecha las dos
    # tablas ya ocupan todo el alto de una pantalla de notebook
    left = sg.Column([
        [sg.Table(values=[], key="-LIST-",
                  headings=["Shipping Order", "Missing docs", "Mismatches",
                            "Check by hand"],
                  col_widths=LIST_WIDTHS, num_rows=19, **table)],
        [sg.Checkbox("Show exports with no documents", key="-SHOW_EMPTY-",
                     enable_events=True, visible=False)],
        [sg.Text("DETAIL", font=BOLD, text_color=C["muted"], pad=(5, (8, 2))),
         sg.Text("click a row on the right", text_color=C["muted"])],
        [sg.Multiline("", key="-DETAIL-", size=(44, 10), disabled=True,
                      expand_x=True, expand_y=True, font=FONT)],
        [sg.Button("Export all to Excel", key="-ORGANIZE-", disabled=True,
                   size=(22, 1), font=BOLD, button_color=("#FFFFFF", C["ok"]),
                   pad=(5, (8, 0)),
                   tooltip="One folder per export, with its documents and its "
                           "Excel report, plus one Excel with all the exports.")],
    ], vertical_alignment="top", expand_y=True, pad=(0, (8, 8)))

    right = sg.Column([
        [sg.Text("Process a folder to see its exports.", key="-TITLE-",
                 font=("Segoe UI", 13, "bold"), size=(70, 1), pad=(5, 0),
                 expand_x=True)],
        [sg.Text("", key="-SUMMARY-", font=("Segoe UI", 11), size=(90, 1),
                 pad=(5, (0, 4)), expand_x=True)],
        [sg.Text("DOCUMENTS", font=BOLD, text_color=C["muted"], pad=(5, (4, 2)))],
        [sg.Table(values=[], key="-DOCS-",
                  headings=["Code", "Document", "Status", "Found in"],
                  col_widths=DOCS_WIDTHS, num_rows=17, **table)],
        [sg.Text("RULES", font=BOLD, text_color=C["muted"], pad=(5, (8, 2)))],
        [sg.Table(values=[], key="-RULES-",
                  headings=["Rule", "What is checked", "Result", "Values"],
                  col_widths=RULES_WIDTHS, num_rows=14, **table)],
    ], vertical_alignment="top", expand_x=True, expand_y=True, pad=((12, 0), (8, 8)))

    return [[left, right]]


# =============================================================================
# Worker
# =============================================================================
def run_extraction(folder, rules, out_q):
    """Corre en un hilo aparte: si corriera en el principal, la ventana
    quedaria congelada durante todo el OCR (12s por documento escaneado)."""
    try:
        def progress(done, total, name, result):
            out_q.put(("progress", (done, total, name, result)))
        results = extract.process_folder(folder, rules, on_progress=progress)
        out_q.put(("done", results))
    except Exception:
        out_q.put(("error", traceback.format_exc()))


# =============================================================================
# Main
# =============================================================================
def main():
    rules = yaml.safe_load(open(HERE / "rules.yaml", encoding="utf-8"))
    recon_rules = reconciliation.load_rules(HERE / "reconciliation.yaml")
    checklist_docs = reconciliation.load_checklist(HERE / "checklist.yaml")

    layout = [
        [sg.TabGroup([[sg.Tab("1. Process", tab_process(), key="-TAB_PROCESS-"),
                       sg.Tab("2. Exports", tab_exports(), key="-TAB_EXPORTS-")]],
                     expand_x=True, expand_y=True, pad=(10, 8))],
    ]

    # sin tamaño fijo: la ventana toma el de su contenido, que entra en una
    # pantalla de notebook (el nombre de la herramienta va en la barra de titulo)
    win = sg.Window("China VAT Refund — check of the export documents", layout,
                    resizable=True, finalize=True, margins=(0, 0))

    results, out_q = [], queue.Queue()
    recon, processed_folder = None, None
    all_entries, entries, shown = [], [], None   # shown = la entrada en pantalla
    t_start = None                               # inicio de la corrida en curso

    def fmt_elapsed(secs):
        m, s = divmod(int(round(secs)), 60)
        return f"{m}m {s:02d}s" if m else f"{s}s"

    def log(msg):
        win["-LOG-"].print(msg)

    def status(text, color, pct=None):
        win["-ICON-"].update(text_color=color)
        win["-STATUS-"].update(text, text_color=color)
        if pct is not None:
            win["-PCT-"].update(f"{pct}%")
            win["-BAR-"].update(current_count=pct, max=100)

    def show_entry(e):
        nonlocal shown
        shown = e
        win["-TITLE-"].update(e["title"])
        win["-SUMMARY-"].update(e["summary"], text_color=e["color"][1])
        win["-DOCS-"].update(values=_fit_rows(e["docs"], DOCS_WIDTHS),
                             row_colors=_row_colors(e["docs"], 2))
        win["-RULES-"].update(values=_fit_rows(e["rules"], RULES_WIDTHS),
                              row_colors=_row_colors(e["rules"], 2))
        win["-DETAIL-"].update("Notes:\n" + "\n".join(f"  •  {n}" for n in e["notes"])
                               if e["notes"] else "")

    def refresh_list():
        """Arma la lista de la izquierda y muestra la primera."""
        nonlocal entries, shown
        show_empty = win["-SHOW_EMPTY-"].get()
        entries = [e for e in all_entries if show_empty or not e["empty"]]
        rows = [e["row"] for e in entries]
        win["-LIST-"].update(
            values=_fit_rows(rows, LIST_WIDTHS),
            row_colors=[(i, e["color"][1], e["color"][0]) for i, e in enumerate(entries)])
        if entries:
            win["-LIST-"].update(select_rows=[0])
            show_entry(entries[0])
        else:
            shown = None
            win["-TITLE-"].update("No exports with documents in this folder.")
            win["-SUMMARY-"].update("")
            for k in ("-DOCS-", "-RULES-"):
                win[k].update(values=[])
            win["-DETAIL-"].update("")

    while True:
        ev, val = win.read(timeout=120)
        if ev == sg.WIN_CLOSED:
            break

        # reloj en vivo mientras corre el worker (el read() ya despierta cada 120ms)
        if t_start is not None:
            win["-TOTAL_TIME-"].update(fmt_elapsed(time.perf_counter() - t_start))

        # --- cola del worker
        while not out_q.empty():
            kind, payload = out_q.get()
            if kind == "progress":
                done, total, name, res = payload
                pct = int(done / max(total, 1) * 100)
                if res is None:
                    status(f"Processing {name}…", C["primary"], pct)
                else:
                    mark = "OK " if res["ok"] else "ERR"
                    log(f"[{done}/{total}] {mark} {name}  ({res['secs']}s)")
                    if not res["ok"]:
                        # sin esto el log dice 'ERR' y nada mas, que no
                        # alcanza para saber que pasó
                        log(f"        {res.get('error', 'no detail')}")
                        if res.get("traceback"):
                            log(res["traceback"])
                    status(f"{done}/{total} processed", C["primary"], pct)
            elif kind == "done":
                results = payload
                recon = reconciliation.run(results, recon_rules,
                                           rules["doc_types"], checklist_docs)
                all_entries = build_entries(recon, results)
                n_empty = sum(1 for e in all_entries if e["empty"])
                win["-SHOW_EMPTY-"].update(
                    value=False, visible=bool(n_empty),
                    text=f"Show the {n_empty} exports with no documents")
                refresh_list()
                win["-ORGANIZE-"].update(disabled=not recon["shipments"])
                win["-PROCESS-"].update(disabled=False, text="Process")

                with_docs = [s for s in recon["shipments"]
                             if not organize._no_documents(s)]
                n_missing = sum(1 for s in with_docs if any(
                    i["status"] == "MISSING" for i in s["checklist"]))
                n_mismatch = sum(1 for s in with_docs if any(
                    r["verdict"] == "DISCREPANCY" for r in s["rules"]))
                status(f"Done — {plural(len(with_docs), 'export', 'exports')}: "
                       f"{n_missing} with missing documents, "
                       f"{n_mismatch} with rules that do not match", C["ok"], 100)
                elapsed = fmt_elapsed(time.perf_counter() - t_start)
                win["-TOTAL_TIME-"].update(elapsed)
                t_start = None
                log(f"Done. Total time: {elapsed}")
                # problemas de la carpeta entera (ej. dos 3.1): una sola vez,
                # bien visible, en vez de repetirlo en cada exportacion
                for w in recon["folder_warnings"]:
                    log(f"WARNING: {w}")
                if recon["folder_warnings"]:
                    sg.popup_ok("\n\n".join(recon["folder_warnings"]),
                                title="Check the folder")
                win["-TAB_EXPORTS-"].select()        # el resultado, a la vista
            elif kind == "error":
                log(payload)
                win["-PROCESS-"].update(disabled=False, text="Process")
                status("Error during processing", C["bad"], 0)
                t_start = None               # congela el reloj en lo que llego a tardar

        # --- eventos
        if ev == "-PROCESS-":
            folder = Path(val["-FOLDER-"])
            if not folder.is_dir():
                sg.popup_error("The folder doesn't exist:", str(folder))
                continue
            if not any(p.suffix.lower() in extract.PROCESSORS
                      for p in folder.rglob("*")):
                sg.popup_error("No PDFs or xlsx files in:", str(folder))
                continue
            # el 3.1 y el refund record son el ancla del periodo: con dos no
            # se puede saber cual vale, asi que no se procesa (ver rules.yaml)
            dups = extract.duplicate_single_files(folder, rules)
            if dups:
                sg.popup_error(
                    "This folder has more than one file of a type that must be "
                    "unique. Keep only one of each and process again:\n\n"
                    + "\n\n".join(
                        next((f"{d['code']} {d['name']}" for d in checklist_docs
                              if dt in (d.get("doc_types") or [])), dt)
                        + ":\n  " + "\n  ".join(files)
                        for dt, files in dups.items()),
                    title="Duplicate files")
                continue
            win["-LOG-"].update("")
            win["-PROCESS-"].update(disabled=True, text="Processing…")
            status("Starting…", C["primary"], 0)
            t_start = time.perf_counter()
            processed_folder = folder
            extract.QUICK_LOOK_ENABLED = not val["-READ_ALL-"]
            win["-ORGANIZE-"].update(disabled=True)
            threading.Thread(target=run_extraction, args=(folder, rules, out_q),
                             daemon=True).start()

        elif ev == "-SHOW_EMPTY-":
            refresh_list()

        elif ev == "-LIST-" and val["-LIST-"]:
            e = entries[val["-LIST-"][0]]
            if e is not shown:
                show_entry(e)

        elif ev == "-DOCS-" and val["-DOCS-"] and shown:
            win["-DETAIL-"].update(shown["doc_details"][val["-DOCS-"][0]])

        elif ev == "-RULES-" and val["-RULES-"] and shown:
            win["-DETAIL-"].update(shown["rule_details"][val["-RULES-"][0]])

        elif ev == "-ORGANIZE-" and recon:
            try:
                out = organize.export_by_shipment(results, recon, rules["doc_types"],
                                                  processed_folder)
            except Exception as e:
                sg.popup_error("Couldn't write the Excel reports:", str(e))
                continue
            log(f"Excel reports ({len(recon['shipments'])} exports): {out}")
            status(f"Excel reports saved in {out}", C["ok"])
            try:
                os.startfile(out)            # abre la carpeta en el Explorador
            except OSError:
                sg.popup("Excel reports saved in:", str(out))

    win.close()


if __name__ == "__main__":
    main()
