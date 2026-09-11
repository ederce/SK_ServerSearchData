#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Monitoreo diario de notificaciones IPN en los 4 servidores PUBLICO
(SVPRDLW364, SVPRDMW363, SVPRDLW1021, SVPRDMW1020): descubre TODAS las
peticiones de notificacion que hicieron TODOS los comercios (correctas e
incorrectas), agrupa por URL de destino, y clasifica cada falla en una
CAUSA especifica (proxy bloqueado por ACL, proxy sin poder conectar al
destino, certificado invalido/vencido del comercio, certificado "a vigilar"
del lado de Izipay, timeout, HTTP propio del comercio, 2xx no-200), para
poder identificar rapido que atacar primero.

Reutiliza la conexion/descarga SFTP de la skill hermana
`busqueda-transacciones-local-sftp` (mismas credenciales guardadas,
misma topologia `remote_paths.txt`) y el diseño de marca Izipay de las
skills `reportes-html-izipay` / `reportes-excel-izipay`.

Uso tipico (por defecto: hoy, categoria PUBLICO, api notification.dispatcher1):
    python monitorear_ipn_publico.py

Otro dia, u otras APIs de notificacion:
    python monitorear_ipn_publico.py --fecha 2026-09-10
    python monitorear_ipn_publico.py --api notification.dispatcher1,notification.dispatcher2
"""
import argparse, gzip, os, re, sys
from collections import defaultdict
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
SIBLING_SKILL_DIR = os.path.join(SCRIPT_DIR, "..", "busqueda-transacciones-local-sftp")
sys.path.insert(0, os.path.abspath(SIBLING_SKILL_DIR))
try:
    import buscar_transacciones_sftp as sftp_lib
except ImportError:
    sys.exit("No se encontro la skill hermana 'busqueda-transacciones-local-sftp' "
              "(se reutiliza su logica de conexion SFTP). Verifica que ambas skills "
              "esten en la misma carpeta padre '.claude/skills/'.")

sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR, "..", "reportes-html-izipay")))
sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR, "..", "reportes-excel-izipay")))
from izipay_report import ReportBuilder, chip, esc, timestamp_now
from izipay_excel import (write_title_block, style_header_row, style_status_cell,
                           set_column_widths, freeze_header)
import openpyxl
from openpyxl.styles import Font

SERVERS = ["SVPRDLW364", "SVPRDMW363", "SVPRDLW1021", "SVPRDMW1020"]
FIXED_SERVERS = {"SVPRDLW364", "SVPRDMW363"}  # servidores con historial de trust store incompleto
HEALTHY_SERVERS = {"SVPRDLW1021", "SVPRDMW1020"}

DEFAULT_REMOTE_PATHS = os.path.join(SIBLING_SKILL_DIR, "remote_paths.txt")
DEFAULT_CREDENCIALES = os.path.join(SIBLING_SKILL_DIR, "crypto", "sftp_config.json")

TARGET_RE = re.compile(r"SendNotificationToReceptor Target URL:\s*(\S+)")
RESULT_RE = re.compile(r"Finaliza proceso envio notificaci.n externa\.\s*La respuesta obtenida fue:\s*(\{.*\})", re.I)
MERCHANT_RE = re.compile(r'"merchantCode"\s*:\s*"(\d+)"')
SERVER_RE = re.compile(r"SVPRD\w+")
PORT_RE = re.compile(r"WebSiteApi_(\d+)")
STATUSCODE_RE = re.compile(r'"statusCode":(\d+)')
REASON_RE = re.compile(r"respondi.\s+(\w+)")
PROXY_STATUS_RE = re.compile(r"failed with status code '(\d+)'")


def cause_of(body):
    """Clasifica UNA linea de error puntual en una causa especifica.
    Devuelve (cause_key, label, severity)."""
    if "proxy tunnel request" in body:
        m = PROXY_STATUS_RE.search(body)
        proxy_code = m.group(1) if m else "?"
        if proxy_code == "403":
            return "proxy_bloqueo", "Proxy Izipay: bloqueo ACL (revisar con Redes)", "critical"
        return (f"proxy_error_{proxy_code}",
                f"Comercio: proxy no pudo conectar (HTTP {proxy_code} del proxy; URL probablemente invalida/no configurada)",
                "serious")
    if "certificate chain" in body:
        return "certificado", "Comercio: certificado invalido/vencido", "serious"
    if "TaskCanceledException" in body or "Timeout" in body:
        return "timeout", "Comercio: timeout de conexion (100s)", "serious"
    m = STATUSCODE_RE.search(body)
    code = m.group(1) if m else "?"
    mr = REASON_RE.search(body)
    reason = mr.group(1) if mr else ""
    if code.startswith("2"):
        return f"http_2xx_{code}", f"Comercio: respondio {code} {reason}".strip(), "good"
    return f"http_{code}", f"Comercio: HTTP {code} {reason}".strip(), "serious"


def analizar(local_files):
    """Recibe la lista [(remote_path, local_path), ...] ya descargada (misma
    forma que produce download_candidates() de la skill hermana) y devuelve
    (server_activity, server_files, urls) -- misma estructura que se uso en
    el analisis manual del 2026-09-11, en una sola pasada."""
    server_activity = defaultdict(int)
    server_files = defaultdict(set)
    url_ok = defaultdict(int)
    url_fail = defaultdict(int)
    url_merchants_all = defaultdict(set)
    url_merchants_fail = defaultdict(set)
    url_per_server = defaultdict(lambda: defaultdict(lambda: {"ok": 0, "fail": 0}))
    url_causes = defaultdict(dict)

    for remote_path, local_path in local_files:
        srv_m = SERVER_RE.search(remote_path)
        srv = srv_m.group(0) if srv_m else "DESCONOCIDO"
        server_files[srv].add(remote_path)
        opener = gzip.open if local_path.lower().endswith(".gz") else open
        with opener(sftp_lib.long_path(local_path), "rt", encoding="utf-8", errors="replace") as fh:
            pending_url = None
            pending_merchant = None
            for line in fh:
                m = TARGET_RE.search(line)
                if m:
                    server_activity[srv] += 1
                    pending_url = m.group(1).rstrip(".,")
                    pending_merchant = None
                    continue
                mm = MERCHANT_RE.search(line)
                if mm and pending_url:
                    pending_merchant = mm.group(1)
                    url_merchants_all[pending_url].add(pending_merchant)
                    continue
                m = RESULT_RE.search(line)
                if m and pending_url:
                    body = m.group(1)
                    ok = bool(re.search(r'"statuscode"\s*:\s*"?200"?', body, re.I))
                    url_per_server[pending_url][srv]["ok" if ok else "fail"] += 1
                    if ok:
                        url_ok[pending_url] += 1
                    else:
                        url_fail[pending_url] += 1
                        if pending_merchant:
                            url_merchants_fail[pending_url].add(pending_merchant)
                        key, label, severity = cause_of(body)
                        c = url_causes[pending_url].setdefault(
                            key, {"label": label, "severity": severity, "count": 0,
                                  "servers": defaultdict(int), "sample": body[:600]}
                        )
                        c["count"] += 1
                        c["servers"][srv] += 1
                    pending_url = None
                    pending_merchant = None
                    continue

    # "A vigilar": la causa "certificado" de una URL ocurre SOLO en los
    # servidores con historial de trust store incompleto, con exito en un
    # servidor sano -- posible CA/intermedia faltante del lado de Izipay,
    # distinto de un certificado vencido/invalido del comercio (que fallaria
    # tambien en los servidores sanos).
    for url, causes in url_causes.items():
        cert = causes.get("certificado")
        if cert:
            cert_servers = set(cert["servers"].keys())
            healthy_ok = any(url_per_server[url].get(s, {}).get("ok", 0) > 0 for s in HEALTHY_SERVERS)
            if cert_servers and cert_servers <= FIXED_SERVERS and (url_ok[url] > 0 or healthy_ok):
                cert["label"] = "A vigilar (Izipay): error de certificado solo en SVPRDLW364/SVPRDMW363"
                cert["severity"] = "warning"

    urls_out = {}
    for url in set(list(url_ok.keys()) + list(url_fail.keys())):
        causes_out = {}
        for key, c in url_causes.get(url, {}).items():
            causes_out[key] = {"label": c["label"], "severity": c["severity"], "count": c["count"],
                                "servers": dict(c["servers"]), "sample": c["sample"]}
        urls_out[url] = {
            "ok": url_ok.get(url, 0), "fail": url_fail.get(url, 0),
            "merchants_all": sorted(url_merchants_all.get(url, [])),
            "merchants_fail": sorted(url_merchants_fail.get(url, [])),
            "per_server": {s: dict(c) for s, c in url_per_server.get(url, {}).items()},
            "causes": causes_out,
        }
    return dict(server_activity), {k: len(v) for k, v in server_files.items()}, urls_out


def filtrar_candidatos_y_ultima_escritura(all_found, fecha):
    """A partir de TODOS los archivos listados (sin descargar todavia),
    separa los candidatos que cumplen el filtro de fecha y, de paso (mismos
    atributos ya obtenidos, sin una segunda pasada de listado), calcula la
    ultima hora de escritura por servidor/puerto."""
    last_write = defaultdict(dict)
    candidates = []
    for path, attr in all_found:
        if not sftp_lib.date_in_range(os.path.basename(path), fecha, None, None):
            continue
        candidates.append((path, attr))
        srv_m = SERVER_RE.search(path)
        port_m = PORT_RE.search(path)
        if srv_m and port_m:
            srv, port = srv_m.group(0), port_m.group(1)
            cur = last_write[srv].get(port)
            if cur is None or attr.st_mtime > cur:
                last_write[srv][port] = attr.st_mtime
    last_write_fmt = {srv: {p: datetime.fromtimestamp(t).strftime("%Y-%m-%d %H:%M:%S") for p, t in ports.items()}
                      for srv, ports in last_write.items()}
    return candidates, last_write_fmt


# --------------------------------------------------------------------------
# Reporte (Excel + HTML, diseno Izipay via las skills reportes-*-izipay)
# --------------------------------------------------------------------------
def _armar_filas(urls):
    """Una fila por (URL, causa) -- una URL puede tener varias causas
    mezcladas -- ordenadas de mayor a menor cantidad de fallas."""
    rows = []
    for url, info in urls.items():
        if info["fail"] == 0:
            continue
        merchants = info.get("merchants_fail") or info.get("merchants_all") or []
        per_server = info.get("per_server", {})
        per_server_txt = " | ".join(f"{srv}: {c['ok']}/{c['fail']}" for srv, c in sorted(per_server.items()))
        for cause_key, cause in info.get("causes", {}).items():
            servers_txt = ", ".join(f"{s}: {n}" for s, n in sorted(cause["servers"].items(), key=lambda kv: -kv[1]))
            rows.append({
                "url": url, "url_ok": info["ok"], "url_fail": info["fail"],
                "cause_key": cause_key, "label": cause["label"], "severity": cause["severity"],
                "count": cause["count"], "servers_txt": servers_txt, "sample": cause["sample"],
                "merchants": merchants, "per_server_txt": per_server_txt,
            })
    rows.sort(key=lambda r: -r["count"])
    return rows


def construir_excel(rows, server_activity, server_files, last_write, fecha, total_intentos,
                     total_urls, total_urls_error, n_vigilar, n_proxy, xlsx_path):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Resumen"
    r = write_title_block(ws, "Monitoreo diario IPN - 4 servidores PUBLICO",
                           f"Fecha: {fecha}  |  Todos los comercios  |  Categoria PUBLICO", merge_cols=6)
    r += 1
    r = style_header_row(ws, r, ["Servidor", "Intentos de notificacion", "Archivos escaneados",
                                  "Ultima escritura puerto 443", "Ultima escritura puerto 444"])
    for srv in SERVERS:
        ws.cell(r, 1, srv)
        ws.cell(r, 2, server_activity.get(srv, 0))
        ws.cell(r, 3, server_files.get(srv, 0))
        ws.cell(r, 4, last_write.get(srv, {}).get("443", "N/D"))
        ws.cell(r, 5, last_write.get(srv, {}).get("444", "N/D"))
        r += 1
    r += 1
    for label, value in [("Total intentos", total_intentos), ("URLs de notificacion distintas", total_urls),
                          ("URLs con al menos 1 error", total_urls_error),
                          ("Causas de falla distintas", len(rows)),
                          ("Causas 'a vigilar' (patron asimetrico 364/363)", n_vigilar),
                          ("Causas de bloqueo de proxy (ACL)", n_proxy)]:
        ws.cell(r, 1, label).font = Font(bold=True)
        ws.cell(r, 2, value)
        r += 1
    set_column_widths(ws, [46, 22, 18, 22, 22])

    ws2 = wb.create_sheet("Causas_de_falla")
    ws2["A1"] = ("Una fila por CAUSA de falla dentro de cada URL (una URL puede tener varias causas mezcladas). "
                 "Ordenado de mayor a menor cantidad de fallas -- la primera fila es la causa que mas conviene "
                 "atacar. 'OK/FALLO total URL' es el total de esa URL sumando TODAS las causas.")
    ws2["A1"].font = Font(italic=True, color="948C8B", size=9)
    ws2.merge_cells("A1:G1")
    next_row = style_header_row(ws2, 2, ["Comercio(s)", "URL de notificacion", "OK (total)", "FALLO (total)",
                                         "Detalle por servidor (ok/fallo)", "Categoria", "Ejemplo de error"])
    for i, row in enumerate(rows, start=next_row):
        ws2.cell(i, 1, ", ".join(row["merchants"]) or "N/D")
        ws2.cell(i, 2, row["url"])
        ws2.cell(i, 3, row["url_ok"])
        ws2.cell(i, 4, row["url_fail"])
        ws2.cell(i, 5, row["per_server_txt"])
        style_status_cell(ws2.cell(i, 6), row["severity"], label=row["label"])
        ws2.cell(i, 7, row["sample"][:500])
    set_column_widths(ws2, [24, 55, 10, 10, 46, 30, 70])
    freeze_header(ws2, row=next_row)
    wb.save(xlsx_path)


def construir_html(rows, server_activity, server_files, last_write, fecha, total_intentos,
                    total_urls, total_urls_error, n_vigilar, n_proxy, apis_desc, html_path):
    rb = ReportBuilder(
        kicker="Monitoreo diario · IPN",
        h1="Actividad y errores de notificacion IPN en los 4 servidores PUBLICO",
        dek="Todos los comercios, sin filtrar por transaccion puntual. Detecta actividad correcta e "
            "incorrecta, con foco en clasificar cada causa de falla para saber que atacar primero.",
        meta=[f"Fecha: {esc(fecha)}", f"Generado: {timestamp_now()}", "Categoria: PUBLICO", f"API: {esc(apis_desc)}"],
        search_box=False, expand_controls=False, theme_controls=True,
    )
    stats = [
        ("Intentos totales", f"{total_intentos:,}"),
        ("Servidores activos", f"{sum(1 for s in SERVERS if server_activity.get(s, 0) > 0)} / 4"),
        ("URLs distintas", str(total_urls)),
        ("URLs con error", str(total_urls_error), "serious"),
        ("Causas de falla distintas", str(len(rows)), "serious"),
        ("A vigilar (Izipay)", str(n_vigilar), "warning"),
        ("Proxy bloqueado (ACL)", str(n_proxy), "critical"),
    ]
    rb.add_stats(stats)

    rb.add_section_title("Actividad por servidor")
    act_rows = [[esc(s), f"{server_activity.get(s, 0):,}", str(server_files.get(s, 0)),
                 f'<span class="mono">{esc(last_write.get(s, {}).get("443", "N/D"))}</span>',
                 f'<span class="mono">{esc(last_write.get(s, {}).get("444", "N/D"))}</span>']
                for s in SERVERS]
    rb.add_table(["Servidor", "Intentos de notificacion", "Archivos escaneados",
                  "Ultima escritura puerto 443", "Ultima escritura puerto 444"], act_rows)

    vigilar_rows = [r for r in rows if r["severity"] == "warning"]
    if vigilar_rows:
        items = "".join(
            f'<li><span class="mono">{esc(r["url"])}</span> &mdash; {r["count"]} fallo(s) de esta causa, '
            f'solo en {esc(r["servers_txt"])} (URL completa: OK={r["url_ok"]}/FALLO={r["url_fail"]})</li>'
            for r in vigilar_rows
        )
        rb.add_callout(
            f'<b>A vigilar:</b> causas de certificado concentradas solo en SVPRDLW364/SVPRDMW363 '
            f'(servidores con historial de trust store incompleto), con exito en los servidores sanos. '
            f'Revisar si se repite en dias siguientes.<ul>{items}</ul>',
            variant="warning",
        )

    proxy_rows = [r for r in rows if r["cause_key"] == "proxy_bloqueo"]
    if proxy_rows:
        items = "".join(
            f'<li><span class="mono">{esc(r["url"])}</span> &mdash; {r["count"]} fallo(s), servidor(es): {esc(r["servers_txt"])}</li>'
            for r in proxy_rows
        )
        rb.add_callout(f'<b>Bloqueo de proxy (ACL):</b> el proxy corporativo rechaza la conexion antes de '
                        f'llegar al destino -- revisar con Redes la categorizacion del dominio.<ul>{items}</ul>',
                        variant="critical")

    rb.add_section_title("Causas de falla (una fila por causa, ordenadas por impacto)")
    rb.add_html(
        '<p class="dek" style="margin:0 0 10px;font-size:12.5px;">Una URL puede tener varias causas de falla '
        'mezcladas (ej. timeout Y certificado vencido al mismo tiempo); aqui cada causa es su propia fila, '
        'ordenada de mayor a menor cantidad de fallas. <b>OK/FALLO total URL</b> es el total de esa URL '
        'sumando TODAS sus causas (contexto).</p>'
    )
    causa_rows_html = []
    for r in rows:
        causa_rows_html.append([
            f'<span class="mono ellipsis" title="{esc(", ".join(r["merchants"]) or "N/D")}">{esc(", ".join(r["merchants"]) or "N/D")}</span>',
            chip(r["label"], r["severity"]),
            str(r["count"]),
            f'<span class="mono ellipsis" title="{esc(r["url"])}">{esc(r["url"])}</span>',
            f'<span class="mono ellipsis" title="{esc(r["servers_txt"])}">{esc(r["servers_txt"])}</span>',
            f"{r['url_ok']}/{r['url_fail']}",
            f'<span class="ellipsis" title="{esc(r["sample"][:500])}">{esc(r["sample"][:180])}</span>',
        ])
    rb.add_table(
        ["Comercio(s)", "Causa", "Fallas", "URL de notificacion", "Servidores afectados",
         "OK/FALLO total URL", "Ejemplo de error"],
        causa_rows_html, table_id="causasTable", searchable=True, sortable=True,
        search_placeholder="Buscar URL, comercio o servidor...",
    )

    with open(html_path, "w", encoding="utf-8") as fh:
        fh.write(rb.render())


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(
        description="Monitoreo diario de notificaciones IPN en los 4 servidores PUBLICO (todos los comercios)."
    )
    ap.add_argument("--fecha", default=datetime.now().strftime("%Y-%m-%d"),
                     help="Fecha a analizar (YYYY-MM-DD). Default: hoy.")
    ap.add_argument("--api", default="notification.dispatcher1",
                     help="API(s) de notificacion a analizar, coma-separado. Default 'notification.dispatcher1' "
                          "(el unico validado contra logs reales). 'notification.dispatcher2'/'notification.producer' "
                          "deberian compartir el mismo formato pero no se han validado todavia.")
    ap.add_argument("--categoria", default="PUBLICO",
                     help="Categoria de servidores (default 'PUBLICO', el alcance validado de esta skill).")
    ap.add_argument("--remote-paths", default=DEFAULT_REMOTE_PATHS,
                     help="Archivo de topologia remota (default: el de la skill hermana busqueda-transacciones-local-sftp).")
    ap.add_argument("--max-depth", type=int, default=3, help="Profundidad de recorrido recursivo remoto (default 3).")

    ap.add_argument("--host", help="Host SFTP (si se omite, se pide por consola o se usa el guardado).")
    ap.add_argument("--port", help="Puerto SFTP.")
    ap.add_argument("--usuario", help="Usuario SFTP.")
    ap.add_argument("--auth", choices=["password", "key"], help="Metodo de autenticacion.")
    ap.add_argument("--credenciales", default=DEFAULT_CREDENCIALES,
                     help="Archivo JSON de credenciales SFTP guardadas (default: el de la skill hermana).")
    ap.add_argument("--guardar-credenciales", action="store_true")
    ap.add_argument("--olvidar-credenciales", action="store_true")

    ap.add_argument("--cache-dir", default=os.path.join(SCRIPT_DIR, ".cache"),
                     help="Carpeta local de cache de descargas (propia de esta skill, se reutiliza entre corridas).")
    ap.add_argument("--cache-max-days", type=int, default=15,
                     help="Borra automaticamente de --cache-dir los archivos con mas de N dias de antiguedad, "
                          "antes de cada corrida (default 15; usa 0 para desactivar esta limpieza automatica).")
    ap.add_argument("--paralelo-livianos", type=int, default=5)
    ap.add_argument("--paralelo-pesados", type=int, default=2)
    ap.add_argument("--umbral-pesado-mb", type=int, default=300)

    ap.add_argument("--output-xlsx", default=None,
                     help="Nombre del Excel de salida (default 'Monitoreo_IPN_Publico_<fecha>.xlsx').")
    ap.add_argument("--output-html", default=None,
                     help="Nombre del HTML de salida (default 'Monitoreo_IPN_Publico_<fecha>.html').")
    args = ap.parse_args()

    apis = [a.strip() for a in args.api.split(",") if a.strip()]
    reports_dir = os.path.join(SCRIPT_DIR, "_reports", args.fecha)
    os.makedirs(reports_dir, exist_ok=True)
    os.makedirs(args.cache_dir, exist_ok=True)
    removed, freed = sftp_lib.cleanup_old_cache(args.cache_dir, args.cache_max_days)
    if removed:
        print(f"Cache: borrados {removed} archivo(s) con mas de {args.cache_max_days} dias "
              f"({freed / (1024 * 1024):.1f} MB liberados).")
    xlsx_path = os.path.join(reports_dir, args.output_xlsx or f"Monitoreo_IPN_Publico_{args.fecha}.xlsx")
    html_path = os.path.join(reports_dir, args.output_html or f"Monitoreo_IPN_Publico_{args.fecha}.html")

    if args.olvidar_credenciales and os.path.exists(args.credenciales):
        os.remove(args.credenciales)
        print(f"Credenciales guardadas borradas: {args.credenciales}")

    remote_dirs = sftp_lib.load_remote_paths(args.remote_paths, args.categoria, args.api, None)
    print(f"[1/5] Directorios remotos a recorrer (categoria={args.categoria}, api={args.api}): {len(remote_dirs)}")

    print("[2/5] Conectando por SFTP...")
    client, sftp = sftp_lib.connect_sftp(args)
    try:
        print("[3/5] Listando archivos remotos (fecha + ultima escritura)...")
        all_found = []
        for cat, base in remote_dirs:
            found = sftp_lib.walk_remote(sftp, base, max_depth=args.max_depth)
            all_found.extend(found)
            print(f"  {base}: {len(found)} archivo(s)")
        candidates, last_write = filtrar_candidatos_y_ultima_escritura(all_found, args.fecha)
        if not candidates:
            sys.exit(f"No se encontro ningun archivo remoto para la fecha {args.fecha}.")

        print(f"[4/5] Descargando {len(candidates)} archivo(s) a cache local ({args.cache_dir})...")
        local_files = sftp_lib.download_candidates(
            client, candidates, args.cache_dir,
            small_workers=args.paralelo_livianos, large_workers=args.paralelo_pesados,
            large_threshold_mb=args.umbral_pesado_mb,
        )
    finally:
        sftp.close()
        client.close()

    print("[5/5] Analizando y generando reporte...")
    server_activity, server_files, urls = analizar(local_files)
    rows = _armar_filas(urls)
    total_intentos = sum(server_activity.values())
    total_urls = len(urls)
    total_urls_error = sum(1 for u in urls.values() if u["fail"] > 0)
    n_vigilar = sum(1 for r in rows if r["severity"] == "warning")
    n_proxy = sum(1 for r in rows if r["cause_key"] == "proxy_bloqueo")

    construir_excel(rows, server_activity, server_files, last_write, args.fecha, total_intentos,
                     total_urls, total_urls_error, n_vigilar, n_proxy, xlsx_path)
    construir_html(rows, server_activity, server_files, last_write, args.fecha, total_intentos,
                    total_urls, total_urls_error, n_vigilar, n_proxy, args.api, html_path)

    print(f"\nGenerado: {xlsx_path}")
    print(f"Generado: {html_path}")
    print(f"Intentos totales: {total_intentos} | URLs distintas: {total_urls} | "
          f"URLs con error: {total_urls_error} | Causas de falla: {len(rows)} "
          f"(a vigilar: {n_vigilar}, proxy ACL: {n_proxy})")


if __name__ == "__main__":
    main()
