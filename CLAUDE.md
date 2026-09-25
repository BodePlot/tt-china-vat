# China VAT Refund — herramienta de verificación de legajos de exportación

## Qué es y para qué

China devuelve el IVA de lo que se exporta (出口退税). Por cada exportación, la empresa tiene que guardar un legajo de
documentos coherentes entre sí: aduana, factura, contrato, naviera, logística, banco. Esta herramienta lee esos
documentos (PDFs escaneados o nativos, Excels), extrae datos clave, agrupa por exportación, verifica que los datos
crucen entre documentos y que el legajo esté completo.

Pedido del equipo de impuestos. El "machete" es **rules.xlsx**:
- **Fila 1**: bloques de documentos (3.2, 3.3, … 3.8.2.2). Es el mismo número que el prefijo de los archivos de
  `samples/`, y de ahí sale `checklist.yaml`.
- **Fila 3 "Item Code"**: numera 84 datos. Los campos de `rules.yaml` usan ese número (`f11_…`, `f45_…`).
- **Fila 4 "Validation"**: los 13 cruces reales, traducidos 1:1 en `reconciliation.yaml`. Los otros ~70 ítems no se
  cruzan en ninguna regla (pendiente preguntar si son solo para registrar).
- **Tabla china, filas 5-14, columnas B-K**: el equipo dijo que es su **"starting point"**. Es el export
  "Refund record from TAX BUREAU": una fila por exportación reclamada, con CDF, factura, período, tasas, cantidad y
  montos USD/CNY. NO es lo mismo que el 3.1. Ver "Pendientes".

`fields/*.pdf|jpg|webp` son capturas anotadas a mano con el número de ítem sobre el documento. Si contradicen a
rules.xlsx, **gana rules.xlsx**. Ejemplo: el ítem 45 es Container No. de la pág. 2 del manifiesto, aunque la imagen
marque otra cosa.

## Archivos

| Archivo | Rol |
|---|---|
| `extract.py` | Lee PDF/xlsx → clasifica cada página por su **contenido** (firmas) → extrae campos por su **relación con la etiqueta impresa**, nunca por coordenadas absolutas ni por nombre de archivo. Tiene caché de OCR. |
| `rules.yaml` | Por doc_type: firma (`required` / `optional` / `negative`) y campos con estrategia (`inline`, `below`, `right`, `in_label`, `page_regex`, `table_col`, `xlsx_col`). `per_row: true` = log con muchas exportaciones: cada fila es un documento aparte. |
| `reconciliation.py` | Agrupa por exportación (union-find), evalúa reglas, arma checklist y avisos. `run()` → `{"shipments": [...], "unassigned": [...]}`. |
| `reconciliation.yaml` | Las reglas de cruce. `verifiable: false` = la regla compara datos que nunca van a ser iguales como texto. `unit:` convierte unidades (kg/ton). |
| `checklist.yaml` | Los 16 documentos del legajo y qué doc_types cuentan como cada uno. Vacío = sin extractor → estado MANUAL. |
| `organize.py` | Salida por exportación: `output/exports/<fecha_hora>/<SO>/documents/` (copias) + `<SO>_report.xlsx`. Los logs compartidos no se copian; sus filas van al reporte. |
| `gui.py` | FreeSimpleGUI. Pestañas: Process, Results, Review (corrección manual con recorte del documento), Checklist (+ botón "Export by shipment"), Reconciliation. |
| `review_store.py` | Correcciones manuales persistidas en `output/corrections.json`. |

Correr: `python gui.py` · CLI: `python extract.py samples` / `python reconciliation.py samples` / `python organize.py samples`.
Siempre con `venv/Scripts/python.exe`. `samples/`, `output/` y `*.xlsx` están en .gitignore: **samples no tiene backup**.

## Lógica del proceso

1. **Clasificar**: cada página se puntúa contra las firmas de cada tipo; tiene que ganar POR DIFERENCIA, si no queda
   UNRESOLVED. Un PDF puede tener varios documentos (ej. la pág. 1 del CDF es el aviso de liberación).
2. **Extraer**: el valor se ubica relativo a su etiqueta. Las tolerancias son relativas a la altura de la etiqueta,
   así que no dependen de los DPI. `corroborate` relee el mismo dato desde otro lugar para confirmarlo.
3. **Agrupar por exportación** (`cluster_records`): se unen las páginas que comparten SO o N° de contenedor
   (`JOIN_KEYS`). Una página sin claves sigue al resto de su archivo, si ese archivo cae en un solo grupo; si no, queda
   "unassigned". Se descartan los grupos formados solo por filas de logs (exportaciones que no tienen documentos en la
   carpeta).
4. **Reglas**: VERIFIED (≥2 patas confiables y coinciden) / DISCREPANCY (las patas confiables no coinciden) /
   NOT_VERIFIABLE (faltan patas, confianza baja o `verifiable: false`). Una pata de baja confianza puede confirmar,
   pero nunca acusa. Los IDs se comparan exactos, las listas por superposición y los números con tolerancia.
5. **Checklist**: PRESENT / MISSING (tiene extractor y no está) / MANUAL (sin extractor: no se afirma nada).
   Avisos: "mixes N different SOs" y "same document from N different files". El union-find es transitivo, así que
   un contenedor mal leído puede pegar dos exportaciones.

## OCR y rendimiento (CPU only: la PC corporativa no tiene GPU)

- RapidOCR **1.2.3**: su `__call__` **ignora** `use_det` / `use_cls` / `use_rec`. Para detectar sin reconocer hay que
  llamar a `eng.text_detector` directo (ver `_detect_only`).
- **Triage**: detecta cajas a `TRIAGE_SCALE=0.4`. Si la mediana de altura de caja es > 250 px, es una página de
  términos y condiciones y se saltea sin leerla. `KNOWN_BOILERPLATE_PAGES` saltea por nombre de archivo (3.7.2 pág. 3);
  es frágil y está pendiente pasarlo a una detección por contenido.
- **Rotación**: las cajas del triage detectan una página de costado. Se hace UN OCR a +90 y, si el clasificador de
  ángulo marcó >50% de las líneas como 180°, se dan vuelta las coordenadas (resultado = -90) sin releer. La
  confianza de reconocimiento NO sirve para elegir el sentido: da ~0.78 en ambos.
- Medido en los 3 PDF escaneados: 117 s → 55 s, con los 33 campos idénticos. Probado y **descartado**: detección a
  ≤2000 px (pierde un contenedor), 200 DPI, sin clasificador, lotes de reconocimiento más grandes, procesos en
  paralelo (más lento).
- Caché: `output/ocr_cache/`. La clave es ruta + mtime + tamaño + parámetros: mover o renombrar un archivo = re-OCR.

## Samples

- Exportación 1 (**real**): SO 13194443, archivos `3.x …`.
- Exportaciones 2-4 (**sintéticas**, copias de la 1 con otras claves): SO 13198819 / 13196085 / 13196173. Las SO y
  los contenedores salen del Shipping Order real. Los nombres de archivo varían a propósito (con SO, en chino) para
  probar que el nombre no importa.
- Exportación 5 (**sintética con errores a propósito**): SO 13194294, archivos `cdf/invoice/contract/bl/manifest
  13194294.pdf`. La factura trae el N° de CDF mal tipeado (…517 en vez de …511) → regla 13 DISCREPANCY, y USD 49680
  en vez de 46920 → la regla 11 da NOT_VERIFIABLE, no DISCREPANCY (ver Trampas: montos contra un CDF escaneado).
- `3.1 Export Tax Refund Filing Directory.xlsx`: filas 1-5, una por exportación.
- `3.7.4.3 Shipping Order.xlsx`: log de camiones de muchas exportaciones (una fila = un contenedor).
- Resultado esperado para las 4: 7 docs PRESENT, 9 MANUAL, 10 reglas VERIFIED, 4 NOT_VERIFIABLE (3, 4 y 6 por
  diseño; la 12 por falta de fuente), 0 discrepancias.

## Trampas conocidas

- **Montos y pesos nunca dan DISCREPANCY hoy**: se cruzan contra el CDF escaneado, que sale LOW_CONF, y una pata
  dudosa no acusa. Una diferencia real queda como NOT_VERIFIABLE. Se arregla con el Refund Tool (fuente confiable
  de montos) o mejorando la confianza del CDF. Ver "Próximos pasos".
- Escaneos de costado: si los campos con posición salen vacíos pero los `page_regex` funcionan, sospechar rotación.
- `tax_filing_directory`: CDF y factura podrían venir como float en Excel, con precisión perdida → marcados LOSSY.
- FreeSimpleGUI `row_colors` = `(fila, texto, fondo)`, no `(fila, fondo, texto)`.
- `sg.Combo.update(values=…)` borra la selección: pasar `value="All"`.
- Antes de decir "no hay sample" de algo: listar `samples/` completo (hay .xlsx además de .pdf) y comparar campo por
  campo, no el esquema entero.
- Antes de culpar a la extracción por una "discrepancia": mirar la página real. Varias veces el problema era la regla
  (unidades kg/ton en la regla 9; datos no equivalentes en las reglas 3, 4 y 6).

## Próximos pasos (propuestos, sin implementar)

Objetivo: que un error real de monto o peso (ej. la exportación 5, regla 11) no quede escondido como NOT_VERIFIABLE.
En este orden:

1. **Estado `CHECK`** en `reconciliation.py`: las patas no coinciden, pero alguna es LOW_CONF. Sigue sin acusar,
   pero se distingue de "faltan patas". Es barato y sirve para todas las reglas de montos y pesos.
2. **Subir la confianza del CDF con aritmética**: si el bloque de precio trae cantidad × precio unitario = total
   (sin verificar), que las tres cifras cierren confirma la lectura del OCR. Así las reglas 9, 10 y 11 pueden dar
   DISCREPANCY de verdad.
3. **Refund record sintético** sin esperar el real: resuelve la regla 12, le da a la 11 una fuente confiable
   (Excel, sin OCR) y arranca el pendiente #1. Cuando llegue el real, se ajustan las columnas.

## Pendientes / preguntas al equipo

1. **Refund record (la tabla china) como punto de partida**: nuevo doc_type `per_row`, el CDF como JOIN_KEY, alertas
   "reclamada sin documentos" y "documentos sin reclamo", y las reglas 1, 2, 11 y 12 contra esa fuente. Hace falta un
   export real (o armar uno sintético). No tiene número 3.x.
2. ¿Qué documentos son obligatorios para un legajo completo? (todos cuentan igual por ahora)
3. ¿Los ~70 ítems sin regla son solo para registrar, o se cruzan? (ej. LSR No. factura ↔ estado de cuenta, SO)
4. Equivalencias para las reglas 3 y 6; qué quieren cruzar realmente en la 4. Tolerancias de pesos y montos.
5. En el refund record: ¿qué significan las filas con montos en 0 y las varias facturas por CDF?
6. Samples de los documentos sin extractor (3.4, 3.6, 3.7.4.1/2/4, 3.7.5.x, 3.8.2.x). Hay referencias visuales de
   algunos en `fields/`.

