---
name: reportes-html-izipay
description: >-
  Genera reportes HTML autocontenidos con la identidad visual de marca de
  Izipay (rojo #A6342E atenuado + turquesa #3DD2CE, brandbar, tarjetas
  redondeadas, tipografia Segoe UI, modo claro/oscuro con boton persistido en
  localStorage). No hace ninguna busqueda, descarga ni analisis de datos por
  su cuenta: es una libreria/plantilla de maquetado que recibe contenido ya
  calculado (titulo, estadisticas, tablas, callouts, tarjetas por transaccion)
  desde otro script o skill, y devuelve el HTML final como string listo para
  guardar en disco. Usar esta skill cada vez que se necesite producir un
  reporte HTML para el equipo de Ecommerce Izipay (monitoreos, investigaciones
  de incidentes, resultados de busquedas de transacciones, seguimientos), en
  vez de crear un estilo nuevo desde cero, para que todos los reportes se vean
  consistentes entre si. Reutilizable: solo cambian el contenido (kicker,
  titulo, estadisticas, tablas, callouts) que le pasa el script que la usa.
---

# Reportes HTML con identidad Izipay

Skill de **maquetado unicamente**: expone la libreria Python
`izipay_report.py` con la hoja de estilos, el JS (toggle claro/oscuro,
busqueda y orden de tablas, expandir/contraer tarjetas) y funciones
constructoras para armar un reporte HTML con la misma identidad visual en
todos los casos, sin reescribir CSS cada vez.

## Cuando usar esta skill

Frases tipicas: "arma el reporte en HTML", "quiero un reporte con el diseño
de Izipay", "hazlo con el mismo formato de siempre", "genera el HTML de este
monitoreo/investigacion". En general, cualquier vez que otro script o skill
ya tiene los datos analizados y necesita presentarlos como un documento HTML
para compartir con el equipo.

> Esta skill NO reemplaza la logica de negocio de otras skills
> (`busqueda-transacciones-sftp`, `trazabilidad-transacciones`,
> `reporte-ipn-fallidas`): esas siguen encargandose de buscar/cruzar logs y
> decidir que mostrar. Esta skill solo aporta el "como se ve".

## Origen del diseño

Nace de la Opcion A ("Izipay corporativo") elegida por el usuario el
2026-09-10 entre 3 propuestas de diseño, y fue el criterio que ya se aplico
manualmente en:
- El `build_html()` de `busqueda-transacciones-sftp` (reportes por
  TransactionId, con tarjeta-acordeon + linea de tiempo por hito).
- El reporte de monitoreo agregado `Monitoreo_Dispatcher1_Global_*.html`
  (estadisticas + tablas + callouts, sin transacciones puntuales).

Esta skill consolida esos dos estilos en una sola libreria para no tener que
copiar/pegar CSS cada vez que se arma un reporte nuevo.

## Uso basico

```python
import sys
sys.path.insert(0, r"<ruta-a-esta-skill>")
from izipay_report import ReportBuilder, chip, esc, timestamp_now

rb = ReportBuilder(
    kicker="Monitoreo diario",
    h1="Actividad de notificaciones IPN",
    dek="Todos los comercios, 4 servidores PUBLICO, notification.dispatcher1.",
    meta=[f"Generado: {timestamp_now()}", "Categoria: PUBLICO"],
    theme_controls=True,      # boton modo claro/oscuro (default True)
    search_box=False,         # buscador global de tarjetas .txn (default False)
    expand_controls=False,    # botones expandir/contraer todas las tarjetas .txn
)

rb.add_stats([
    ("Intentos totales", "6,513"),
    ("Servidores activos", "4 / 4"),
    ("URLs con error", "23", "critical"),   # 3er elemento opcional: colorea con --status-<key>
])

rb.add_section_title("Actividad por servidor")
rb.add_table(
    headers=["Servidor", "Intentos", "Ultima escritura"],
    rows=[["SVPRDLW364", "1,567", "2026-09-11 10:32"]],
    searchable=False, sortable=True,
)

rb.add_callout(
    f"<b>A vigilar:</b> detalle del hallazgo... {chip('Alerta', 'warning')}",
    variant="warning",   # brand (default) | good | warning | serious | critical
)

html = rb.render()
open("reporte.html", "w", encoding="utf-8").write(html)
```

### Componentes disponibles (`izipay_report.py`)

- `ReportBuilder(kicker, h1, dek, meta, ...)` — arma el documento completo.
  Metodos: `add_stats()`, `add_section_title()`, `add_callout()`,
  `add_table()`, `add_html()` (escape hatch para markup custom) y `render()`.
- `chip(label, key)` — badge de estado con icono, `key` en
  `good|warning|serious|critical|error|none` (mismos 5 niveles de severidad
  que usan `trazabilidad-transacciones` y `busqueda-transacciones-sftp`:
  good=OK/Aprobado, warning=Alerta, serious=Reversado, critical=Rechazado,
  error=nunca notificado tras reintentos).
- `esc(s)` — escapa HTML (usar siempre sobre texto que no sea markup de
  confianza antes de insertarlo en celdas/callouts).
- `CSS` / `JS` — strings crudos, por si un reporte necesita maquetar algo muy
  particular (ej. tarjetas `.txn`/`.timeline`/`.step` tipo acordeon con
  hitos, como hace `busqueda-transacciones-sftp`) sin pasar por
  `ReportBuilder`. Ver las clases ya definidas en `CSS` antes de inventar
  nuevas.
- `timestamp_now()` — fecha/hora local formateada `DD-MM-YYYY HH:MM`, para el
  meta-row del hero.

### Tablas: `searchable` vs `sortable`

- `searchable=True` agrega un input de busqueda arriba de la tabla que
  filtra filas por texto (case-insensitive, contra todo el texto de la fila).
- `sortable=True` hace clickeables los headers para ordenar (numerico si
  todas las celdas de esa columna parsean como numero, alfabetico si no).
- Se pueden combinar. Si se generan **varias tablas** en el mismo reporte,
  cada una lleva su propio buscador/orden independientes (el JS los detecta
  por `table.data[data-searchable]` / `[data-sortable]`, no hace falta IDs
  unicos salvo que se quiera enlazar con anclas `#`).

### Tarjetas tipo acordeon (para reportes por transaccion)

Si el reporte necesita una tarjeta expandible por transaccion con linea de
tiempo de hitos (como `busqueda-transacciones-sftp`), usar las clases ya
definidas en `CSS` (`.txn`, `.txn > summary`, `.txn-meta`, `.timeline`,
`.step`) armando el HTML a mano con `add_html()` y pasando
`expand_controls=True` al `ReportBuilder` para que aparezcan los botones
"Expandir/Contraer todo" (funcionan automaticamente sobre cualquier
`<details class="txn">` presente en el documento, sin configuracion
adicional). Ver `buscar_transacciones_sftp.py` (`build_html()`, en la skill
`busqueda-transacciones-sftp`) como referencia de implementacion completa de
este patron (esa skill mantiene su propio CSS/JS inline por ahora, no importa
todavia de esta libreria; son visualmente equivalentes).

## Paleta y principios (no reinventar)

- Marca: rojo `#A6342E` (atenuado desde el `#FF4240` real de Izipay, se
  sintio "muy fuerte" en superficies grandes) + turquesa `#3DD2CE` (acento,
  el punto junto al wordmark). Ver detalle en memoria
  `reference_izipay_brand_colors`.
- Severidad: 5 niveles fijos `good/warning/serious/critical/error`, siempre
  **icono + etiqueta** (nunca color solo) porque algunos tonos no llegan a
  3:1 de contraste en superficie clara.
- Cabecera del reporte (`brandbar`) siempre en el rojo de marca; el
  **contenido narrativo** (títulos, callouts de alerta) nunca usa el rojo
  vivo real de marca como fondo grande — ver memoria
  `feedback_reporte_tono_sobrio` (principio de sobriedad, aplica incluso al
  color de marca real de la empresa).
- Modo oscuro: toggle manual (boton) siempre gana sobre la preferencia del
  SO; persistido en `localStorage` bajo la key `izipay-report-theme`.

## Validado

Generado y verificado visualmente (Chrome headless, claro y oscuro) contra
datos reales el 2026-09-11: reporte de monitoreo agregado
`Monitoreo_Dispatcher1_Global_2026-09-11.html/.xlsx` (estadisticas + tabla de
actividad por servidor + callouts de hallazgos + tabla de errores por URL
buscable). Ver script de referencia usado para generarlo (equivalente al
ejemplo de uso basico de arriba) en el historial de esta sesion / memoria del
proyecto.

## Requisitos

Ninguno externo: solo Python 3.8+ (libreria estandar). No requiere
`paramiko`/`openpyxl`/conexion SFTP — esta skill nunca busca datos, solo
maqueta lo que ya le pasan.
