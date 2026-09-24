"""
Persistencia de las correcciones manuales.

Por que en un archivo aparte y no dentro de la GUI: una correccion no es
un evento de interfaz, es un dato del negocio. El analista corrigio un
valor que el OCR leyo mal, y eso tiene que sobrevivir al reprocesamiento
(si no, cada corrida le pide corregir lo mismo de nuevo).

Ademas las correcciones son informacion valiosa por si mismas: si el
mismo campo del mismo tipo de documento se corrige siempre, la regla de
extraccion esta mal y hay que arreglarla, no seguir corrigiendo a mano.
"""

import json
from datetime import datetime
from pathlib import Path


class ReviewStore:
    def __init__(self, path):
        self.path = Path(path)
        self.data = {}
        self.load()

    # --- clave -------------------------------------------------------------
    @staticmethod
    def key(doc_path, page, field):
        """Identifica un campo puntual: archivo + pagina + nombre de campo."""
        return f"{doc_path}|{page}|{field}"

    # --- io ----------------------------------------------------------------
    def load(self):
        if self.path.is_file():
            try:
                self.data = json.loads(self.path.read_text(encoding="utf-8"))
            except (json.JSONDecodeError, OSError):
                self.data = {}          # archivo corrupto: no tira la corrida

    def save(self):
        self.path.parent.mkdir(parents=True, exist_ok=True)
        self.path.write_text(
            json.dumps(self.data, ensure_ascii=False, indent=2),
            encoding="utf-8")

    # --- api ---------------------------------------------------------------
    def set(self, doc_path, page, field, value, ocr_value, reason=""):
        """Guarda una correccion, conservando lo que habia leido el OCR."""
        self.data[self.key(doc_path, page, field)] = {
            "corrected": value,
            "ocr_value": ocr_value,       # para despues medir que falla seguido
            "reason": reason,
            "at": datetime.now().isoformat(timespec="seconds"),
        }
        self.save()

    def get(self, doc_path, page, field):
        return self.data.get(self.key(doc_path, page, field))

    def apply(self, results):
        """
        Reaplica las correcciones sobre un set de resultados recien extraido.
        El valor corregido reemplaza al del OCR y el estado pasa a CORRECTED.
        """
        n = 0
        for r in results:
            for p in r.get("pages", []):
                for fname, f in p.get("fields", {}).items():
                    c = self.get(r["path"], p["page"], fname)
                    if not c:
                        continue
                    f["ocr_value"] = f.get("value")
                    f["value"] = c["corrected"]
                    f["status"] = "CORRECTED"
                    f["conf"] = 1.0
                    n += 1
        return n

    def stats_by_field(self):
        """
        Cuantas veces se corrigio cada campo. Si un campo aparece seguido,
        el problema esta en la regla de extraccion, no en el documento.
        """
        counts = {}
        for k in self.data:
            field = k.rsplit("|", 1)[-1]
            counts[field] = counts.get(field, 0) + 1
        return dict(sorted(counts.items(), key=lambda kv: -kv[1]))
