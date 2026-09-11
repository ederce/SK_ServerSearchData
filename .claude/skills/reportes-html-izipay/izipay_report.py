#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Libreria reutilizable para generar reportes HTML autocontenidos con la
identidad de marca Izipay (rojo #A6342E atenuado + turquesa #3DD2CE, brandbar,
tarjetas redondeadas, tipografia Segoe UI, modo claro/oscuro). No hace ninguna
busqueda ni analisis de datos: solo recibe contenido ya calculado (titulo,
estadisticas, tablas, callouts) y devuelve el HTML final como string.

Uso tipico:
    from izipay_report import ReportBuilder, esc

    rb = ReportBuilder(
        kicker="Monitoreo diario",
        h1="Actividad de notificaciones IPN",
        dek="Descripcion breve del alcance del reporte.",
        meta=["Fecha: 2026-09-11", "Categoria: PUBLICO"],
    )
    rb.add_stats([("Intentos totales", "6,513"), ("Servidores activos", "4 / 4")])
    rb.add_callout("Texto de alerta o hallazgo destacado.", variant="warning")
    rb.add_section_title("Actividad por servidor")
    rb.add_table(["Servidor", "Intentos"], [["SVPRDLW364", "1,567"]], searchable=False)
    html = rb.render()
    open("reporte.html", "w", encoding="utf-8").write(html)

Componentes disponibles ademas de los de arriba (ver docstrings de cada
funcion/metodo): chip() de estado (good/warning/serious/critical/error),
accordion_item() + timeline_step() para reportes tipo "una tarjeta por
transaccion" (igual patron que usa la skill busqueda-transacciones-sftp).
"""
from datetime import datetime

STATUS_KEYS = ("good", "warning", "serious", "critical", "error", "none")

STATUS_ICONS = {
    "good": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
    "warning": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4M12 17h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>',
    "serious": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7M3 4v5h5"/></svg>',
    "critical": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m15 9-6 6M9 9l6 6"/></svg>',
    "error": '<svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor" stroke="none"><path d="M12 2 1 21h22L12 2Zm0 6.2a1.3 1.3 0 0 1 1.3 1.3v5a1.3 1.3 0 0 1-2.6 0v-5A1.3 1.3 0 0 1 12 8.2Zm0 9.4a1.45 1.45 0 1 1 0 2.9 1.45 1.45 0 0 1 0-2.9Z"/></svg>',
    "none": '<svg viewBox="0 0 24 24" width="13" height="13" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg>',
}

CSS = """
:root { color-scheme: light; }
.report {
  --paper: #ffffff; --page: #f7f5f2;
  --ink: #201a1a; --ink-soft: #56504f; --ink-mute: #948c8b;
  --line: #ece6e2; --line-soft: #f4f0ed; --border: rgba(32,26,26,0.10);
  --brand: #A6342E; --brand-text: #A6342E; --brand-ink: #ffffff; --brand-soft: #f7e6e4; --brand-2: #3DD2CE;
  --status-good: #17845a; --status-good-soft: #e2f5ec;
  --status-warning: #a8710a; --status-warning-soft: #faedd2;
  --status-serious: #b3541e; --status-serious-soft: #fbe3d4;
  --status-critical: #b3261e; --status-critical-soft: #fbe7e5;
  --status-error: #7a1015; --status-error-soft: #f6dcdc;
}
@media (prefers-color-scheme: dark) {
  :root:where(:not([data-theme="light"])) .report {
    color-scheme: dark;
    --paper: #241c1c; --page: #191313;
    --ink: #f7f2f1; --ink-soft: #d1c6c5; --ink-mute: #948c8b;
    --line: #3d3231; --line-soft: #2c2423; --border: rgba(255,255,255,0.10);
    --brand: #e0827c; --brand-text: #e0827c; --brand-ink: #191313; --brand-soft: #3a2322; --brand-2: #6be3e0;
    --status-good: #55c894; --status-good-soft: #17301f;
    --status-warning: #e0b357; --status-warning-soft: #362a11;
    --status-serious: #e58a54; --status-serious-soft: #3a2314;
    --status-critical: #ef7b70; --status-critical-soft: #3a1d1a;
    --status-error: #ff8c8c; --status-error-soft: #3a1414;
  }
}
:root[data-theme="dark"] .report {
  color-scheme: dark;
  --paper: #241c1c; --page: #191313;
  --ink: #f7f2f1; --ink-soft: #d1c6c5; --ink-mute: #948c8b;
  --line: #3d3231; --line-soft: #2c2423; --border: rgba(255,255,255,0.10);
  --brand: #ff6b69; --brand-text: #ff9d9b; --brand-ink: #191313; --brand-soft: #3a2322; --brand-2: #6be3e0;
  --status-good: #55c894; --status-good-soft: #17301f;
  --status-warning: #e0b357; --status-warning-soft: #362a11;
  --status-serious: #e58a54; --status-serious-soft: #3a2314;
  --status-critical: #ef7b70; --status-critical-soft: #3a1d1a;
  --status-error: #ff8c8c; --status-error-soft: #3a1414;
}
* { box-sizing: border-box; }
body { margin: 0; background: var(--page); color: var(--ink); font: 14px/1.55 "Segoe UI", system-ui, -apple-system, sans-serif; }
.brandbar { background: var(--brand); color: var(--brand-ink); position: sticky; top: 0; z-index: 20; }
.brandbar .inner { max-width: 1180px; margin: 0 auto; padding: 14px 22px; display: flex;
  align-items: center; justify-content: space-between; flex-wrap: wrap; gap: 8px; }
.wordmark { font-size: 22px; font-weight: 800; letter-spacing: -.01em; display: flex; align-items: center; gap: 10px; }
.wordmark .dot { width: 11px; height: 11px; border-radius: 50%; background: var(--brand-2); display: inline-block; }
.brandbar .tag { font-size: 12px; opacity: .85; }
.controls-row { background: var(--paper); border-bottom: 1px solid var(--border); position: sticky; top: 53px; z-index: 19; }
.controls { max-width: 1180px; margin: 0 auto; padding: 10px 22px; display: flex; flex-wrap: wrap;
  gap: 10px; align-items: center; justify-content: flex-end; }
.controls input[type=search] { padding: 7px 10px; border-radius: 8px; border: 1px solid var(--border);
  background: var(--page); color: var(--ink); font-size: 13px; min-width: 260px; }
.btn { padding: 7px 12px; border-radius: 8px; border: 1px solid var(--border); background: var(--page);
  color: var(--ink); font-size: 12.5px; cursor: pointer; }
.btn:hover { border-color: var(--brand); }
main { max-width: 1180px; margin: 0 auto; padding: 22px 22px 60px; }
.hero { padding: 6px 0 6px; }
.hero .kicker { font-size: 12px; font-weight: 700; letter-spacing: .08em; text-transform: uppercase;
  color: var(--brand-text); margin-bottom: 8px; }
.hero h1 { font-size: 22px; margin: 0 0 8px; font-weight: 800; line-height: 1.32; }
.hero .dek { font-size: 13.5px; color: var(--ink-soft); max-width: 760px; }
.meta-row { font-size: 12px; color: var(--ink-mute); margin-top: 10px; display: flex; gap: 16px; flex-wrap: wrap; }
.stats { display: grid; grid-template-columns: repeat(auto-fit, minmax(150px, 1fr)); gap: 12px; margin: 22px 0 6px; }
.stat { background: var(--paper); border: 1px solid var(--line); border-radius: 14px; padding: 14px 16px;
  border-top: 3px solid var(--stat-c, var(--brand)); }
.stat .label { font-size: 11px; text-transform: uppercase; letter-spacing: .04em; color: var(--ink-mute);
  margin-bottom: 6px; font-weight: 700; }
.stat .value { font-size: 22px; font-weight: 800; color: var(--stat-c, var(--brand)); }
.stat .sub { font-size: 12px; color: var(--ink-soft); margin-top: 4px; }
h2.sec { font-size: 12.5px; text-transform: uppercase; letter-spacing: .06em; color: var(--brand-text); font-weight: 800;
  margin: 32px 0 12px; padding-bottom: 8px; border-bottom: 2px solid var(--line); }
.empty { background: var(--paper); border: 1px dashed var(--line); border-radius: 12px;
  padding: 26px; text-align: center; color: var(--ink-soft); }
.callout { background: var(--brand-soft); border-radius: 12px; padding: 16px 20px; font-size: 13.5px; margin: 10px 0; }
.callout.good { background: var(--status-good-soft); }
.callout.warning { background: var(--status-warning-soft); }
.callout.serious { background: var(--status-serious-soft); }
.callout.critical { background: var(--status-critical-soft); }
.callout ul { margin: 8px 0 0; padding-left: 20px; }
.callout li { margin-bottom: 4px; }
.mono { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 12.5px; }
span.mono { background: var(--line-soft); border-radius: 5px; padding: 1px 6px; }
.duo { display: grid; grid-template-columns: 1fr 1fr; gap: 16px; margin: 6px 0 10px; }
@media (max-width: 640px) { .duo { grid-template-columns: 1fr; } }
.card { border-radius: 14px; padding: 18px 20px; background: var(--paper); border: 1px solid var(--line); }
.card.good { box-shadow: inset 4px 0 0 var(--status-good); }
.card.warning { box-shadow: inset 4px 0 0 var(--status-warning); }
.card.serious { box-shadow: inset 4px 0 0 var(--status-serious); }
.card.critical { box-shadow: inset 4px 0 0 var(--status-critical); }
.card p { font-size: 13px; color: var(--ink-soft); margin: 4px 0; }
.table-toolbar { display: flex; gap: 10px; align-items: center; margin: 8px 0; flex-wrap: wrap; }
.table-toolbar input[type=search] { padding: 7px 10px; border-radius: 8px; border: 1px solid var(--border);
  background: var(--page); color: var(--ink); font-size: 13px; min-width: 260px; }
.table-toolbar .count { color: var(--ink-mute); font-size: 12.5px; }
.table-wrap { border: 1px solid var(--line); border-radius: 12px; overflow: auto; max-height: 560px; }
table.data { border-collapse: collapse; width: 100%; font-size: 12.5px; }
table.data th { position: sticky; top: 0; background: var(--brand-soft); color: var(--brand-text);
  text-align: left; font-weight: 700; padding: 9px 10px; border-bottom: 1px solid var(--line); white-space: nowrap; }
table.data th.sortable { cursor: pointer; }
table.data th.sorted::after { content: " \\2195"; color: var(--brand-text); }
table.data td { padding: 8px 10px; border-bottom: 1px solid var(--line-soft); vertical-align: top; }
table.data tbody tr:hover { background: var(--brand-soft); }
.hidden-row { display: none !important; }
.ellipsis { display: inline-block; max-width: 420px; white-space: nowrap; overflow: hidden;
  text-overflow: ellipsis; vertical-align: bottom; }
td.nowrap { white-space: nowrap; }
.chip { display: inline-flex; align-items: center; gap: 5px; font-size: 11px; font-weight: 800; letter-spacing: .02em;
  text-transform: uppercase; padding: 3px 10px; border-radius: 999px; white-space: nowrap; }
.chip-good { background: var(--status-good-soft); color: var(--status-good); }
.chip-warning { background: var(--status-warning-soft); color: var(--status-warning); }
.chip-serious { background: var(--status-serious-soft); color: var(--status-serious); }
.chip-critical { background: var(--status-critical-soft); color: var(--status-critical); }
.chip-error { background: var(--status-error-soft); color: var(--status-error); }
.chip-none { background: var(--line-soft); color: var(--ink-mute); }
.txn { background: var(--paper); border: 1px solid var(--line); border-radius: 14px;
  margin-bottom: 12px; overflow: hidden; box-shadow: inset 4px 0 0 var(--status-good); }
.txn[data-worst="warning"] { box-shadow: inset 4px 0 0 var(--status-warning); }
.txn[data-worst="serious"] { box-shadow: inset 4px 0 0 var(--status-serious); }
.txn[data-worst="critical"] { box-shadow: inset 4px 0 0 var(--status-critical); }
.txn[data-worst="error"] { box-shadow: inset 4px 0 0 var(--status-error); }
.txn[data-worst="none"] { box-shadow: inset 4px 0 0 var(--ink-mute); }
.txn > summary { list-style: none; cursor: pointer; padding: 14px 16px; display: flex;
  align-items: center; gap: 10px; flex-wrap: wrap; font-weight: 700; }
.txn > summary::-webkit-details-marker { display: none; }
.txn > summary::before { content: "\\25B8"; display: inline-block; color: var(--ink-mute);
  transition: transform .15s; font-weight: 400; }
.txn[open] > summary::before { transform: rotate(90deg); }
.txn-meta { padding: 0 16px 14px; display: flex; flex-wrap: wrap; gap: 8px 22px; color: var(--ink-soft); font-size: 12.5px; }
.txn-meta b { color: var(--ink); font-weight: 700; }
.timeline { list-style: none; margin: 0; padding: 4px 16px 16px; }
.step { position: relative; padding: 8px 0 8px 22px; border-left: 2px solid var(--line); margin-left: 4px; }
.step:last-child { border-left-color: transparent; }
.step::before { content: ""; position: absolute; left: -6px; top: 12px; width: 10px; height: 10px;
  border-radius: 50%; background: var(--chip-c, var(--ink-mute)); }
.step-head { display: flex; align-items: center; gap: 10px; flex-wrap: wrap; }
.step-evento { font-weight: 700; }
.step-hora { color: var(--ink-mute); font-size: 12px; font-variant-numeric: tabular-nums; }
.step-resultado { color: var(--ink-soft); font-size: 12.5px; margin-top: 3px; }
.step-evidencia { color: var(--ink-mute); font-size: 11.5px; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; margin-top: 2px; }
.idx-link { color: var(--brand-text); font-weight: 700; text-decoration: none; }
.idx-link:hover { text-decoration: underline; }
.mono, .txn-id, .step-evidencia { font-variant-numeric: tabular-nums; }
.detalle-txt { font-family: ui-monospace, SFMono-Regular, Menlo, monospace; font-size: 11.5px;
  white-space: pre-wrap; word-break: break-word; }
footer { text-align: center; color: var(--ink-mute); font-size: 11.5px; padding: 26px 0 6px;
  border-top: 1px solid var(--line); margin-top: 36px; }
@media print { .controls-row, .brandbar { display: none; } body { background: #fff; } }
"""

JS = """
(function(){
  var root = document.documentElement;
  var stored = localStorage.getItem('izipay-report-theme');
  if (stored) root.setAttribute('data-theme', stored);
  var themeBtn = document.getElementById('toggleTheme');
  if (themeBtn) {
    themeBtn.addEventListener('click', function(){
      var cur = root.getAttribute('data-theme') ||
        (window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
      var next = cur === 'dark' ? 'light' : 'dark';
      root.setAttribute('data-theme', next);
      localStorage.setItem('izipay-report-theme', next);
    });
  }

  var expandBtn = document.getElementById('expandAll');
  var collapseBtn = document.getElementById('collapseAll');
  if (expandBtn) expandBtn.addEventListener('click', function(){
    document.querySelectorAll('details.txn').forEach(function(d){ d.open = true; });
  });
  if (collapseBtn) collapseBtn.addEventListener('click', function(){
    document.querySelectorAll('details.txn').forEach(function(d){ d.open = false; });
  });

  function norm(s){ return (s || '').toLowerCase(); }

  var globalSearch = document.getElementById('globalSearch');
  if (globalSearch) {
    globalSearch.addEventListener('input', function(){
      var q = norm(globalSearch.value);
      document.querySelectorAll('details.txn').forEach(function(d){
        d.classList.toggle('hidden-row', q && norm(d.textContent).indexOf(q) === -1);
      });
    });
  }

  document.querySelectorAll('table.data[data-searchable]').forEach(function(table){
    var wrap = table.closest('.table-wrap') || table.parentElement;
    var toolbar = wrap.previousElementSibling;
    var input = toolbar ? toolbar.querySelector('input[type=search]') : null;
    var countEl = toolbar ? toolbar.querySelector('.count') : null;
    var rows = Array.prototype.slice.call(table.querySelectorAll('tbody tr'));
    function filter(){
      var q = input ? norm(input.value) : '';
      var visible = 0;
      rows.forEach(function(tr){
        var match = !q || norm(tr.textContent).indexOf(q) !== -1;
        tr.classList.toggle('hidden-row', !match);
        if (match) visible++;
      });
      if (countEl) countEl.textContent = visible + ' de ' + rows.length;
    }
    if (input) input.addEventListener('input', filter);
    filter();
  });

  document.querySelectorAll('table.data[data-sortable] thead th').forEach(function(th, idx){
    th.classList.add('sortable');
    th.addEventListener('click', function(){
      var table = th.closest('table');
      var tbody = table.querySelector('tbody');
      var asc = th.getAttribute('data-asc') !== 'true';
      table.querySelectorAll('thead th').forEach(function(t){ t.classList.remove('sorted'); t.removeAttribute('data-asc'); });
      th.classList.add('sorted'); th.setAttribute('data-asc', asc);
      var rows = Array.prototype.slice.call(tbody.querySelectorAll('tr'));
      rows.sort(function(a, b){
        var av = a.children[idx].textContent.trim(), bv = b.children[idx].textContent.trim();
        var an = parseFloat(av.replace(/,/g, '')), bn = parseFloat(bv.replace(/,/g, ''));
        var cmp = (!isNaN(an) && !isNaN(bn)) ? an - bn : av.localeCompare(bv);
        return asc ? cmp : -cmp;
      });
      rows.forEach(function(r){ tbody.appendChild(r); });
    });
  });
})();
"""


def esc(s):
    return (str(s).replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
            .replace('"', "&quot;"))


def chip(label, key="none"):
    """Badge de estado (good/warning/serious/critical/error/none), icono + etiqueta."""
    key = key if key in STATUS_KEYS else "none"
    return f'<span class="chip chip-{key}">{STATUS_ICONS[key]}<span>{esc(label)}</span></span>'


class ReportBuilder:
    """Arma un reporte HTML Izipay de forma incremental: crear, ir agregando
    secciones/estadisticas/tablas/callouts en orden, y llamar a render() al
    final para obtener el HTML completo como string."""

    def __init__(self, kicker, h1, dek="", meta=None, tag="Reporte tecnico &middot; Ecommerce",
                 title=None, theme_controls=True, search_box=False, expand_controls=False):
        self.kicker = kicker
        self.h1 = h1
        self.dek = dek
        self.meta = meta or []
        self.tag = tag
        self.title = title or h1
        self.theme_controls = theme_controls
        self.search_box = search_box
        self.expand_controls = expand_controls
        self.body = []

    def add_stats(self, items):
        """items: lista de (label, value) o (label, value, status_key_para_color)."""
        parts = ['<div class="stats">']
        for it in items:
            label, value = it[0], it[1]
            color = f' style="--stat-c: var(--status-{it[2]})"' if len(it) > 2 and it[2] else ""
            parts.append(
                f'<div class="stat"{color}><div class="label">{esc(label)}</div>'
                f'<div class="value">{esc(value)}</div></div>'
            )
        parts.append('</div>')
        self.body.append("".join(parts))

    def add_section_title(self, text):
        self.body.append(f'<h2 class="sec">{esc(text)}</h2>')

    def add_callout(self, html, variant="brand"):
        """html puede incluir markup simple (b, ul/li, span.mono, etc.) ya escapado por el caller."""
        cls = "" if variant == "brand" else f" {variant}"
        self.body.append(f'<div class="callout{cls}">{html}</div>')

    def add_html(self, html):
        """Escape hatch para markup custom que no cubren los helpers de arriba."""
        self.body.append(html)

    def add_table(self, headers, rows, table_id=None, searchable=False, sortable=False,
                  search_placeholder="Buscar..."):
        """headers: lista de strings. rows: lista de listas de celdas ya en HTML
        (usar esc() vos mismo si el contenido no es HTML de confianza)."""
        tid = f' id="{esc(table_id)}"' if table_id else ""
        attrs = ""
        if searchable:
            attrs += ' data-searchable="1"'
        if sortable:
            attrs += ' data-sortable="1"'
        parts = []
        if searchable:
            parts.append(
                '<div class="table-toolbar">'
                f'<input type="search" placeholder="{esc(search_placeholder)}">'
                '<span class="count"></span></div>'
            )
        parts.append(f'<div class="table-wrap"><table class="data"{tid}{attrs}><thead><tr>')
        for h in headers:
            parts.append(f'<th>{esc(h)}</th>')
        parts.append('</tr></thead><tbody>')
        for row in rows:
            parts.append('<tr>' + "".join(f'<td>{cell}</td>' for cell in row) + '</tr>')
        parts.append('</tbody></table></div>')
        self.body.append("".join(parts))

    def render(self):
        controls = []
        if self.search_box:
            controls.append('<input type="search" id="globalSearch" placeholder="Buscar...">')
        if self.expand_controls:
            controls.append('<button class="btn" id="expandAll" type="button">Expandir todo</button>')
            controls.append('<button class="btn" id="collapseAll" type="button">Contraer todo</button>')
        if self.theme_controls:
            controls.append('<button class="btn" id="toggleTheme" type="button">Modo claro/oscuro</button>')
        controls_html = "".join(controls)

        meta_html = "".join(f"<span>{m}</span>" for m in self.meta)

        return f"""<!DOCTYPE html>
<html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>{esc(self.title)}</title>
<style>{CSS}</style>
</head><body class="report">
<div class="brandbar"><div class="inner">
  <div class="wordmark"><span class="dot"></span>izipay</div>
  <div class="tag">{self.tag}</div>
</div></div>
<div class="controls-row"><div class="controls">{controls_html}</div></div>
<main>
  <div class="hero">
    <div class="kicker">{esc(self.kicker)}</div>
    <h1>{esc(self.h1)}</h1>
    {'<p class="dek">' + esc(self.dek) + '</p>' if self.dek else ''}
    <div class="meta-row">{meta_html}</div>
  </div>
{''.join(self.body)}
<footer>Generado con la skill reportes-html-izipay &middot; Equipo Ecommerce Izipay</footer>
</main>
<script>{JS}</script>
</body></html>"""


def timestamp_now():
    return datetime.now().strftime("%d-%m-%Y %H:%M")
