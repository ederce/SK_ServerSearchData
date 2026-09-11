---
name: reportes-excel-izipay
description: >-
  Aplica la identidad visual de marca Izipay (rojo #A6342E atenuado, texto
  brand en headers, fondos brand-soft, 5 niveles de severidad
  good/warning/serious/critical/error con el mismo color exacto que la skill
  HTML hermana) a reportes Excel (.xlsx) armados con openpyxl. No crea hojas
  ni estructuras de datos por su cuenta -- expone funciones de estilo
  (titulo, subtitulo, fila de encabezado, celda de estado coloreada, ancho de
  columnas, congelar encabezado) para que cada script arme su propio
  Workbook/Worksheet con openpyxl y solo llame a estas funciones para que se
  vea igual que los demas reportes Excel de Izipay. Usar esta skill cada vez
  que se genere un .xlsx para el equipo de Ecommerce Izipay (monitoreos,
  investigaciones, resultados de busquedas de transacciones) en lugar de
  definir colores/fuentes sueltos cada vez.
---

# Reportes Excel con identidad Izipay

Pareja en Excel de la skill **`reportes-html-izipay`** (HTML): misma paleta
de marca, misma semantica de severidad de 5 niveles, para que un mismo
reporte se vea consistente en sus dos formatos de salida (Excel y HTML).

## Cuando usar esta skill

Frases tipicas: "arma el reporte en Excel", "que el Excel tenga el mismo
formato de siempre", "genera el .xlsx con los colores de Izipay". En
general, cualquier vez que otro script/skill ya tiene los datos y arma un
`.xlsx` con `openpyxl`, para que los colores/fuentes salgan consistentes con
el resto de reportes en vez de definirlos sueltos en cada script.

> Esta skill (igual que `reportes-html-izipay`) es solo de **estilo**: no
> busca datos ni decide la estructura de hojas/columnas -- eso lo define
> cada script segun su propio reporte (los reportes Excel reales varian
> bastante entre si: `busqueda-transacciones-sftp` usa Resumen+Detalle_Logs,
> el monitoreo agregado usa Resumen+Errores_por_URL, etc.).

## Uso basico

```python
import sys
sys.path.insert(0, r"<ruta-a-esta-skill>")
from izipay_excel import (title_font, subtitle_font, write_title_block,
                           style_header_row, style_status_cell,
                           set_column_widths, freeze_header)
import openpyxl

wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Resumen"

row = write_title_block(ws, "Monitoreo diario notification.dispatcher1",
                         "Fecha: 2026-09-11 | Categoria: PUBLICO", merge_cols=6)
# row ya apunta a la siguiente fila libre (3 en este caso: titulo+subtitulo)

row = style_header_row(ws, row, ["Servidor", "Intentos", "Ultima escritura"])
ws.cell(row, 1, "SVPRDLW364"); ws.cell(row, 2, 1567); ws.cell(row, 3, "2026-09-11 10:31")

ws2 = wb.create_sheet("Errores_por_URL")
row2 = style_header_row(ws2, 1, ["Comercio(s)", "URL", "Categoria"])
ws2.cell(row2, 1, "4078314"); ws2.cell(row2, 2, "https://eppoweb.pe:8051/...")
style_status_cell(ws2.cell(row2, 3), "critical", label="Proxy Izipay (conocido)")
set_column_widths(ws2, [24, 55, 22])
freeze_header(ws2, row=2)

wb.save("reporte.xlsx")
```

### Funciones disponibles (`izipay_excel.py`)

- `title_font(size=14)` / `subtitle_font(size=10)` — fuentes para el
  titulo/subtitulo del bloque inicial de cada hoja.
- `write_title_block(ws, title, subtitle=None, start_row=1, start_col=1, merge_cols=6)`
  — escribe titulo (+ subtitulo opcional) mergeados, devuelve la fila
  siguiente libre.
- `header_font()` / `header_fill()` / `style_header_row(ws, row, headers, start_col=1)`
  — fila de encabezados de tabla (fondo `brand-soft`, texto `brand`,
  negrita). Devuelve la fila siguiente.
- `status_fill(key)` / `status_font(key)` / `style_status_cell(cell, key, label=None)`
  — colorea una celda segun severidad. `key` en
  `good|warning|serious|critical|error|none`, **mismos hex exactos** que los
  chips de `reportes-html-izipay` (good=`#17845A`/soft `#E2F5EC`,
  warning=`#A8710A`/`#FAEDD2`, serious=`#B3541E`/`#FBE3D4`,
  critical=`#B3261E`/`#FBE7E5`, error=`#7A1015`/`#F6DCDC`).
- `set_column_widths(ws, widths, start_col=1)` — ancho de columnas en una
  sola llamada (lista alineada con `start_col`).
- `freeze_header(ws, row=2)` — congela las filas de encabezado.

## Paleta y principios (no reinventar)

Misma paleta y mismos principios que `reportes-html-izipay` (ver su
SKILL.md): rojo `#A6342E` atenuado (nunca el `#FF4240` real de marca en
fondos grandes), 5 niveles de severidad fijos, sobriedad de cabecera. Ver
tambien memoria `reference_izipay_brand_colors` y
`feedback_reporte_tono_sobrio`.

## Validado

Los colores/fuentes de esta libreria son los mismos que ya se usaron y
verificaron manualmente (abriendo el .xlsx generado) en
`Monitoreo_Dispatcher1_Global_2026-09-11.xlsx` (hojas Resumen +
Errores_por_URL con categorias coloreadas) el 2026-09-11. La libreria en si
(funciones sueltas) se probo con un caso de ejemplo equivalente al de "Uso
basico" de arriba antes de darla por lista.

## Requisitos

Solo `openpyxl` (misma dependencia que ya usan
`busqueda-transacciones-sftp`/`trazabilidad-transacciones`/
`reporte-ipn-fallidas` para generar Excel). No requiere `paramiko` ni
conexion de ningun tipo -- esta skill nunca busca datos, solo aplica estilo
sobre un Workbook que ya existe.
