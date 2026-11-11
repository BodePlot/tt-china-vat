"""
Test de regresion: corre la herramienta sobre las carpetas de muestra y compara
el resultado con una foto guardada de cuando se sabia que estaba bien.

    python regression.py --save              guarda la foto
    python regression.py                     compara contra la foto
    python regression.py --no-cache          compara rehaciendo el OCR

--no-cache es para cuando se toca la lectura (OCR, triage, rotacion): con la
cache activa, el OCR viejo sale del disco y un cambio ahi pasaria sin que se
note. Para cambios en rules.yaml o en las reglas alcanza la corrida normal.

La foto queda en output/ (fuera de git) porque tiene datos de los documentos
reales: este test solo corre en una PC que tenga samples/ y samples2/.

Sale con codigo 1 si hay diferencias, asi se puede encadenar con otra cosa.
"""

import argparse
import json
import sys
from pathlib import Path

import yaml

import extract
import reconciliation

HERE = Path(__file__).resolve().parent
FOLDERS = ("samples", "samples2")
SNAPSHOT = HERE / "output" / "regression" / "snapshot.json"
MAX_SHOWN = 60                       # diferencias por seccion antes de cortar


def run_folder(folder, rules_yaml):
    results = extract.process_folder(folder, rules_yaml)
    return results, reconciliation.run(results)


def flatten_extraction(results):
    """{clave: valor} con lo que se leyo de cada pagina, listo para comparar."""
    out = {}
    for r in results:
        key = r["path"].replace("\\", "/")
        if not r["ok"]:
            out[f"{key} | error"] = r["error"]
        for p in r["pages"]:
            cls = p["classification"]
            where = f"{key} | p{p['page']}"
            out[f"{where} | type"] = [cls["type"], cls["status"], p["source"]]
            for name, f in p["fields"].items():
                out[f"{where} | {name}"] = [f.get("status"), f.get("value"),
                                            f.get("code"), f.get("conf")]
    return out


def flatten_reconciliation(out):
    flat = {}
    for s in out["shipments"]:
        name = s["shipment"]
        flat[f"{name} | files"] = s["files"]
        flat[f"{name} | kind"] = s["kind"]
        for w in s["warnings"]:
            flat[f"{name} | warning | {w}"] = True
        for r in s["rules"]:
            flat[f"{name} | rule {r['id']}"] = r["verdict"]
        for item in s["checklist"]:
            flat[f"{name} | checklist {item['code']}"] = item["status"]
    for u in out["unassigned"]:
        flat[f"unassigned | {u['file']} p{u['page']}"] = u["doc_type"]
    for w in out["folder_warnings"]:
        flat[f"folder warning | {w}"] = True
    return flat


def take_snapshot():
    rules_yaml = yaml.safe_load(open(HERE / "rules.yaml", encoding="utf-8"))
    snap, secs = {}, {}
    for name in FOLDERS:
        folder = HERE / name
        if not folder.is_dir():
            sys.exit(f"Falta la carpeta {folder}: este test necesita las muestras")
        results, out = run_folder(folder, rules_yaml)
        snap[name] = {"extraction": flatten_extraction(results),
                      "reconciliation": flatten_reconciliation(out)}
        secs[name] = {r["path"].replace("\\", "/"): r["secs"] for r in results}
    return snap, secs


def differences(old, new):
    """(perdidas, nuevas, cambiadas) entre dos dicts planos."""
    lost = sorted(old.keys() - new.keys())
    added = sorted(new.keys() - old.keys())
    changed = sorted(k for k in old.keys() & new.keys() if old[k] != new[k])
    return lost, added, changed


def _print_block(title, lines):
    if not lines:
        return
    print(f"  {title} ({len(lines)}):")
    for line in lines[:MAX_SHOWN]:
        print(f"    {line}")
    if len(lines) > MAX_SHOWN:
        print(f"    ... y {len(lines) - MAX_SHOWN} mas")


def report(saved, current):
    total = 0
    for name in FOLDERS:
        for section in ("extraction", "reconciliation"):
            old, new = saved[name][section], current[name][section]
            lost, added, changed = differences(old, new)
            n = len(lost) + len(added) + len(changed)
            total += n
            print(f"\n{name} / {section}: {len(old)} datos, {n} diferencias")
            _print_block("Se perdieron", [f"{k} = {old[k]}" for k in lost])
            _print_block("Aparecieron", [f"{k} = {new[k]}" for k in added])
            _print_block("Cambiaron", [f"{k}: {old[k]} -> {new[k]}" for k in changed])
    return total


def main():
    ap = argparse.ArgumentParser(description="Test de regresion sobre las muestras")
    ap.add_argument("--save", action="store_true", help="guardar la foto actual")
    ap.add_argument("--no-cache", action="store_true", help="rehacer el OCR")
    args = ap.parse_args()

    if args.no_cache:
        extract.CACHE_ENABLED = False
    # para que los archivos con caracteres chinos no rompan la consola de Windows
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

    current, secs = take_snapshot()
    if args.save:
        SNAPSHOT.parent.mkdir(parents=True, exist_ok=True)
        SNAPSHOT.write_text(json.dumps(current, ensure_ascii=False, indent=1,
                                       sort_keys=True), encoding="utf-8")
        (SNAPSHOT.parent / "timings.json").write_text(
            json.dumps(secs, ensure_ascii=False, indent=1), encoding="utf-8")
        n = sum(len(s) for f in current.values() for s in f.values())
        print(f"\nFoto guardada con {n} datos: {SNAPSHOT}")
        return

    if not SNAPSHOT.is_file():
        sys.exit("No hay foto todavia: correr primero con --save")
    saved = json.loads(SNAPSHOT.read_text(encoding="utf-8"))
    # la foto paso por JSON: se normaliza igual para no ver falsas diferencias
    current = json.loads(json.dumps(current, ensure_ascii=False))
    total = report(saved, current)
    print(f"\n{'OK: sin diferencias' if total == 0 else f'HAY {total} DIFERENCIAS'}")
    sys.exit(1 if total else 0)


if __name__ == "__main__":
    main()
