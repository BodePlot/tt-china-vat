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
  cruzan, pero el equipo **quiere que se extraigan igual** (para registrarlos).
- **Tabla china, filas 5-14, columnas B-K**: el equipo dijo que es su **"starting point"**. Es el export
  "Refund record from TAX BUREAU": una fila por exportación reclamada, con CDF, factura, período, tasas, cantidad y
  montos USD/CNY. NO es lo mismo que el 3.1 (confirmado por las tablas de proceso del equipo, ver abajo).

**Tablas de proceso del equipo** (las pasó el usuario el 2026-09-28; el detalle está en la memoria
`project_team_process_tables`):
- Paso 1: se arranca del Refund record. Cada documento se busca por una clave: el CDF → 3.2; la factura → 3.3; la SO
  → 3.5, 3.7.2, 3.7.4.3 y 3.7.4.4; el B/L → 3.7.1; el **LSR No. (transportation)** → 3.7.4.1 y 3.7.4.2; el
  **LSR No. (storage)** → 3.7.5.1.
- El **3.1** es del equipo de Tax, y ahí el equipo TSTP carga **los números LSR de cada CDF**. Nuestro 3.1 de ejemplo
  no tiene esas columnas. El LSR va a hacer falta como JOIN_KEY cuando haya muestras de 3.7.4.x, 3.7.5.x y 3.8.2.x.
- Hay documentos que no están en rules.xlsx: 3.7.3 (contrato con Sinotrans), 3.8.1.1-4 (contratos y poder con
  Yili), y 3.7.4.5, 3.7.5.3 y 3.8.2.3 (pagos a proveedores). Los contratos parecen documentos marco (no uno por
  exportación). Pendiente decidir si entran al checklist.

`fields/*.pdf|jpg|webp` son capturas anotadas a mano con el número de ítem sobre el documento. Si contradicen a
rules.xlsx, **gana rules.xlsx**. Ejemplo: el ítem 45 es Container No. de la pág. 2 del manifiesto, aunque la imagen
marque otra cosa.

## Archivos

| Archivo | Rol |
|---|---|
| `extract.py` | Lee PDF/xlsx → clasifica cada página por su **contenido** (firmas) → extrae campos por su **relación con la etiqueta impresa**, nunca por coordenadas absolutas ni por nombre de archivo. Tiene caché de OCR. |
| `rules.yaml` | Por doc_type: firma (`required` / `optional` / `negative`) y campos con estrategia (`inline`, `below`, `right`, `in_label`, `page_regex`, `table_col`, `xlsx_col`). `per_row: true` = log con muchas exportaciones: cada fila es un documento aparte. |
| `reconciliation.py` | Agrupa por exportación (union-find), evalúa reglas, arma checklist y avisos. `run()` → `{"shipments": [...], "unassigned": [...]}`. |
| `reconciliation.yaml` | Las reglas de cruce. `check: false` = solo se muestran los datos (veredicto INFO, reglas 3 y 4). `compare: alnum` = ignora espacios y "/" (regla 6). `unit:` convierte unidades (kg/ton). |
| `checklist.yaml` | Los 17 documentos del legajo (DS + 3.1 … 3.8.2.2) y qué doc_types cuentan como cada uno. Hoy todos tienen extractor; vacío = sin extractor → MANUAL. |
| `organize.py` | Salida por exportación: `output/exports/<fecha_hora>/<SO>/documents/` (copias) + `<SO>_report.xlsx`. Los logs compartidos no se copian; sus filas van al reporte. Además `consolidated_report.xlsx` con todas las exportaciones (pedido del equipo de China): Overview, Discrepancies, Missing docs, All rules. |
| `gui.py` | FreeSimpleGUI. Pestañas: Process, Results, Review (corrección manual con recorte del documento), Checklist (+ botón "Export by shipment"), Reconciliation. |
| `review_store.py` | Correcciones manuales persistidas en `output/corrections.json`. |

Correr: `python gui.py` · CLI: `python extract.py samples` / `python reconciliation.py samples` / `python organize.py samples`.
Siempre con `venv/Scripts/python.exe`. `xlrd` (en el venv) lee los `.xls` viejos. `output/` y `samples2/` están en
.gitignore; `samples/` sí está commiteado.

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
   NOT_VERIFIABLE (faltan patas o confianza baja) / INFO (`check: false`, solo se muestra). Una pata de baja
   confianza puede confirmar, pero nunca acusa. Los IDs se comparan exactos, las listas por superposición y los
   números **exactos, sin tolerancia** (pedido del equipo: las diferencias las revisan caso por caso).
   Varios documentos del mismo tipo en una exportación (`_combine_legs`): si coinciden, es un valor. Si no, los IDs
   se juntan en una lista (alcanza con que uno coincida) y los números quedan **MULTI**: no se suman ni se elige
   uno, porque pueden ser partes de un total (varias facturas) o versiones del mismo documento (CDF corregido).
5. **Checklist**: PRESENT / MISSING (tiene extractor y no está) / MANUAL (sin extractor: no se afirma nada).
   Avisos: "mixes N different SOs" y "same document from N different files". El union-find es transitivo, así que
   un contenedor mal leído puede pegar dos exportaciones. Si las SO comparten un único CDF, el aviso dice "one CDF
   covers N SOs" (caso válido, confirmado por el equipo).
6. **Un solo archivo ancla por carpeta**: los tipos con `single_file: true` (el 3.1 y el refund record) no pueden
   estar duplicados. `extract.duplicate_single_files()` los busca ANTES de procesar (son xlsx: es instantáneo) y la
   GUI y la CLI **no procesan** si hay más de uno: piden dejar uno solo. `reconciliation.folder_warnings()` queda
   como respaldo si se llama a `run()` directo. Pendiente consultar con el equipo si están de acuerdo con la regla.
7. **Documentos que se adjuntan en vez de unir** (`attach_by` en rules.yaml, ver `_attach()`): se sacan del
   union-find y se agregan después, porque unirían exportaciones que no tienen nada que ver.
   - `lsr`: las facturas mensuales de gastos (3.7.4.1, 3.7.5.1, 3.8.2.1) no traen SO, solo el N° de LSR. Van a
     **cada** exportación cuyas filas de estado de cuenta (`*_lsr_no`) tengan ese LSR, así que una misma factura
     queda en varias exportaciones, y está bien.
   - `amount_usd`: el 3.4 (aviso de crédito del banco) no trae SO ni CDF. Se adjunta solo si su monto coincide con
     el USD de **una** exportación (del refund record o de la factura). Si coincide con varias, queda "unassigned"
     con el motivo. En los samples coincide con las 5, porque las sintéticas copian los montos.
   - Los **estados de cuenta LSR** (3.7.4.2, 3.7.5.2, 3.8.2.2) son PDFs mensuales con filas de todas las SO del mes:
     `rows:` en rules.yaml los lee fila por fila (`extract._pdf_row_records`, page = pág*100+fila) y cada fila se une
     por su SO, como un log `per_row`.
   - `multiple: true` (Delivery Note, uno por camión; factura de exportación, porque un CDF con varias facturas es
     frecuente): no dispara el aviso "same document from N files".

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
- `Refund record from TAX BUREAU.xlsx` (**sintético**, doc_type `refund_record`): la tabla china de rules.xlsx tal
  cual (9 filas reales de 2026, sin documentos → prueban "claimed but NO documents") + 5 filas nuestras. En la 5
  trae el USD correcto (46920), así la factura errónea (49680) da DISCREPANCY en la regla 11.
- `3.7.4.3 Shipping Order.xlsx`: log de camiones de muchas exportaciones (una fila = un contenedor).
- Gastos y cobro (**reales**, de la exportación 13194443, llegaron el 2026-09-28): `3.4 Bank Receipt.png` (captura:
  pasa por OCR), `3.6 … .xlsx` (export de SAP), facturas `3.7.4.1`, `3.7.5.1`, `3.8.2.1`, estados de cuenta LSR
  `3.7.4.2`, `3.7.5.2`, `3.8.2.2` (traen filas de las 5 SO de nuestras exportaciones, que son SO reales) y
  `3.7.4.4 Delivery Note 1/2.pdf` (uno por contenedor).
- Resultado esperado: exportaciones 1-4 → 12 VERIFIED, 2 INFO (reglas 3 y 4), 0 discrepancias.
  Exportación 5 → igual, pero con DISCREPANCY en las reglas 11 y 13. Más 9 "exportaciones" `CDF 5304…`: reclamadas
  sin documentos.
  Checklist: la 13194443 tiene 16/17 PRESENT (falta el 3.4, que queda unassigned por monto ambiguo); las 2-5 tienen
  14/17 (les faltan 3.4, 3.6 y 3.7.4.4, que solo existen para la 13194443).

### samples2/ (reales, llegaron el 2026-10-05 desde el OneDrive del equipo; todavía NO se procesó entero)

- `List of Exports.xls`: el **refund record real** de julio 2026 (export de eTAX, 109 filas). Ya lo lee el
  doc_type `refund_record` (ver el comentario en rules.yaml). Ninguna fila coincide con las exportaciones de
  `samples/`; sí están las 2 de líquidos de abajo.
- `SO 13349718 PP` / `SO 13368987 PP`: polipropileno de enero (no están en el archivo de julio). B/L con **otro
  formato** (impresión del sistema, `IMOCODE:` + `航次号:`) que hoy no se clasifica.
- `SO 6755067 B1-1-丁烯` / `SO 6759315 BD-丁二烯`: líquidos. Subcarpetas con OTRA numeración (3.6.1 B/L, 3.6.2
  manifiesto, 3.7.1.x contrato Sinotrans, 3.7.2.x despachante). Varias versiones del CDF (`已修改价格` / `已修改数量`)
  → sus números van a dar MULTI. Fapiao y B/L escaneados sin texto.
- `3.7.4.3 26年1月出运单.xls` (Shipping Order de enero, se clasifica bien) y `370041183 03092026 9167479.70.pdf`
  (parece un cobro bancario, escaneado).

## Trampas conocidas

- **Montos contra el CDF escaneado no acusan**: el CDF sale LOW_CONF y una pata dudosa no acusa. Las reglas 11 y 12
  sí pueden dar DISCREPANCY gracias al refund record (Excel, confiable). Las 9 y 10 (pesos: CDF vs factura/B/L)
  todavía no. Ver "Próximos pasos".
- Escaneos de costado: si los campos con posición salen vacíos pero los `page_regex` funcionan, sospechar rotación.
- `tax_filing_directory`: CDF y factura podrían venir como float en Excel, con precisión perdida → marcados LOSSY.
- FreeSimpleGUI `row_colors` = `(fila, texto, fondo)`, no `(fila, fondo, texto)`.
- `sg.Combo.update(values=…)` borra la selección: pasar `value="All"`.
- Antes de decir "no hay sample" de algo: listar `samples/` completo (hay .xlsx además de .pdf) y comparar campo por
  campo, no el esquema entero.
- Antes de culpar a la extracción por una "discrepancia": mirar la página real. Varias veces el problema era la regla
  (unidades kg/ton en la regla 9; datos no equivalentes en las reglas 3, 4 y 6).
- Los archivos que manda el usuario suelen llegar como `.txt` en **base64** (Office = empieza con `UEsDB`; `.xls`
  viejo = `0M8R4KGx`). A veces vienen partidos en varios `.txt` que se concatenan en orden, y un pedazo puede venir
  cortado: comparar los largos.

## Próximos pasos (propuestos, sin implementar)

Objetivo: que un error real de peso (reglas 9 y 10) no quede escondido como NOT_VERIFIABLE. Los montos (11 y 12) ya
se resolvieron con el refund record.
En este orden:

1. **Estado `CHECK`** en `reconciliation.py`: las patas no coinciden, pero alguna es LOW_CONF. Sigue sin acusar,
   pero se distingue de "faltan patas". Es barato y sirve para todas las reglas de montos y pesos.
2. **Subir la confianza del CDF con aritmética**: si el bloque de precio trae cantidad × precio unitario = total
   (sin verificar), que las tres cifras cierren confirma la lectura del OCR. Así las reglas 9, 10 y 11 pueden dar
   DISCREPANCY de verdad.
3. ~~Refund record sintético~~: hecho (ver Samples).

## Pendientes / preguntas al equipo

Respuestas del equipo del 2026-09-29: ver la memoria `project_team_answers_2026_09_29`. Ya implementadas: reglas 3/4
como INFO, regla 6 = buque + viaje, ítem 26 completo, sin tolerancia, varias facturas por CDF, CDF con varias SO,
reporte consolidado y refund record real.

1. **Refund record**: el equipo va a mandar UN archivo consolidado (no uno por mes). Falta ver su formato, y si un
   CDF declarado en varios períodos se suma o vale la última fila.
2. ¿Qué documentos son obligatorios para un legajo completo? (todos cuentan igual por ahora)
3. ¿Los ~70 ítems sin regla son solo para registrar, o se cruzan? (ej. LSR No. factura ↔ estado de cuenta, SO)
4. Falta el sample "CDF with multiple SO" (CDF 533820260000000206), que no vino en samples2.
5. samples2: B/L en formato nuevo, CDF en varias versiones (¿cuál vale?) y la numeración distinta de las carpetas
   de líquidos (3.6.1, 3.7.1.x, 3.7.2.x) contra el checklist.
6. ~~Samples de los documentos sin extractor~~: llegaron el 2026-09-28 (reales, de la exportación 13194443) y ya
   tienen extractor. Falta: el ítem 37 (LSR No. custom declaration de 3.6), que no es una columna del export de SAP.
7. Cruce natural que no está en rules.xlsx: el total de cada factura de gastos debería igualar el total de su
   estado de cuenta LSR (lo dice el propio estado de cuenta: "invoice total amount matches the LSR total"). Ej.:
   transporte 16 437,20 CNY en los dos.

