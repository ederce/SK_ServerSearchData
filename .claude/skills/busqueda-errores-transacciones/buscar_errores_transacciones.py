#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Busca transacciones que hayan resultado en ERROR segun los logs NLog de la
plataforma ecommerce. Similar a busqueda-transacciones-local-sftp pero en
lugar de rastrear una transaccion conocida por txid/orden/comercio, busca
TODAS las lineas que cumplen criterios de ERROR:
  - status code (HTTP-like) != 200
  - code (ISO 8583) != "00" y != "RA2"
  - mensaje contiene "Error" o "exception" (case-insensitive)

Dos modos: (A) SFTP remoto, (B) local --logs-dir.
Ruta debe ser una FECHA (obligatorio, a diferencia del otro skill).
Comercio es OPCIONAL: si se da, filtra ese comercio; si no, muestra TODOS.
Rango horario es OPCIONAL: --hora-desde/hasta O --ultimas-N-horas (no ambos).

Requisitos: Python 3.8+, paramiko, openpyxl
    pip install paramiko openpyxl
"""
import argparse, fnmatch, getpass, glob, gzip, io, json, os, queue, re, secrets, subprocess, sys, threading, time, stat as statmod
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime, timedelta

try:
    import paramiko
except ImportError:
    sys.exit("Falta la libreria 'paramiko'. Instala con: pip install paramiko")

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

SCRIPT_DIR_FOR_IMPORTS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR_FOR_IMPORTS, "..", "reportes-html-izipay")))
sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR_FOR_IMPORTS, "..", "reportes-excel-izipay")))
try:
    from izipay_report import ReportBuilder, esc, timestamp_now
    from izipay_excel import (BRAND, header_font, header_fill, write_title_block, style_header_row,
                               style_status_cell, status_font, set_column_widths, freeze_header)
except ImportError as e:
    print(f"Advertencia: no se pudieron importar modulos de reporte: {e}")

MAX_CELL_LEN = 30000
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
# Topologia unica: se comparte con la skill hermana (ahi se edita remote_paths.txt)
DEFAULT_REMOTE_PATHS = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "busqueda-transacciones-local-sftp", "remote_paths.txt"))

# --------------------------------------------------------------------------
# 1) Utilidades de fecha/hora y busqueda de errores
# --------------------------------------------------------------------------
RE_FILE_DATE = re.compile(r"(\d{4}-\d{2}-\d{2})")
RE_FULL_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$")
TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+)\|(?:(\d+)\|)?(\w+)\|(.*)$")
RE_MERCHANT = re.compile(r'"?[Mm]erchant[Cc]ode"?\s*:\s*"?(\d+)"?')
RE_STATUS_CODE = re.compile(r'"statuscode"\s*:\s*"?(\d{3})"?', re.I)
RE_CODE = re.compile(r'"code"\s*:\s*"?(\d{2})"?(?!\d)', re.I)

def parse_args():
    p = argparse.ArgumentParser(
        description="Busca transacciones con ERROR en logs NLog (SFTP o local)")

    # Fecha (OBLIGATORIO)
    p.add_argument("--fecha", help="Dia exacto a buscar (YYYY-MM-DD), ej. 2026-09-25")
    p.add_argument("--fecha-desde", help="Inicio de rango de fechas (YYYY-MM-DD)")
    p.add_argument("--fecha-hasta", help="Fin de rango de fechas (YYYY-MM-DD)")

    # Horas (OPCIONAL, uno u otro, no ambos)
    p.add_argument("--hora-desde", help="Hora inicio (HH:MM)")
    p.add_argument("--hora-hasta", help="Hora fin (HH:MM)")
    p.add_argument("--ultimas-N-horas", type=int, help="Ultimas N horas desde ahora (entero)")

    # Comercio (OPCIONAL)
    p.add_argument("--comercio", help="Codigo de comercio (OPCIONAL; si se omite, muestra TODOS)")

    # Modo local vs SFTP
    p.add_argument("--logs-dir", help="Modo local: carpeta con logs ya descargados")

    # SFTP
    p.add_argument("--categoria", default="todos",
                   help="PUBLICO,CONTROLLER,BUSINESS,INTERNO o todos (default)")
    p.add_argument("--api", help="Nombre de API (coma-separado), ej. authorization,cancel")
    p.add_argument("--ruta", help="Patron glob sobre ruta remota, ej. */SVPRDLW364/*")
    p.add_argument("--remote-paths", default=DEFAULT_REMOTE_PATHS,
                   help=f"Archivo de topologia remota (default {DEFAULT_REMOTE_PATHS})")

    # Credenciales SFTP
    p.add_argument("--host", help="Host SFTP")
    p.add_argument("--port", type=int, help="Puerto SFTP (default 22)")
    p.add_argument("--usuario", help="Usuario SFTP")
    p.add_argument("--auth", choices=["password", "key"],
                   help="Autenticacion: password o key")
    p.add_argument("--credenciales", default=os.path.join(SCRIPT_DIR, "crypto", "sftp_config.json"),
                   help="Archivo JSON de credenciales guardadas (SFTP)")
    p.add_argument("--guardar-credenciales", action="store_true",
                   help="Tras conectar, guarda credenciales cifradas (SFTP)")
    p.add_argument("--olvidar-credenciales", action="store_true",
                   help="Borra archivo de credenciales guardadas (SFTP)")

    # Cache
    p.add_argument("--cache-dir", default=os.path.join(SCRIPT_DIR, ".cache"),
                   help="Carpeta cache local (default .cache de esta skill)")
    p.add_argument("--cache-max-days", type=int, default=15,
                   help="Borra cache >N dias (default 15, 0 para desactivar)")
    p.add_argument("--limpiar-cache", action="store_true", help="Borra cache antes de empezar")

    # Descarga paralela
    p.add_argument("--paralelo-livianos", type=int, default=5,
                   help="Descargas simultaneas archivos livianos (default 5)")
    p.add_argument("--paralelo-pesados", type=int, default=2,
                   help="Descargas simultaneas archivos pesados (default 2)")
    p.add_argument("--umbral-pesado-mb", type=int, default=300,
                   help="Tamano MB umbral pesado (default 300)")
    p.add_argument("--max-depth", type=int, default=3,
                   help="Profundidad recorrido remoto (default 3)")

    # Salida
    p.add_argument("--output-xlsx", help="Nombre archivo Excel (default Errores_Transacciones.xlsx)")
    p.add_argument("--output-html", help="Nombre archivo HTML (default Errores_Transacciones.html)")

    # Encoding
    p.add_argument("--encoding", help="Encoding logs (default utf-8 SFTP, cp1252 local)")

    return p.parse_args()

def validate_args(args):
    """Valida que fecha sea obligatoria y que hora-desde/hasta + ultimas-N-horas no se combinen."""
    if not (args.fecha or (args.fecha_desde and args.fecha_hasta)):
        sys.exit("ERROR: Debes dar --fecha (dia exacto) O --fecha-desde Y --fecha-hasta (OBLIGATORIO)")

    if args.hora_desde and args.hora_hasta and args.ultimas_N_horas:
        sys.exit("ERROR: No puedes combinar --hora-desde/hasta con --ultimas-N-horas")

    if bool(args.hora_desde) != bool(args.hora_hasta):
        sys.exit("ERROR: Si das --hora-desde, tienes que dar --hora-hasta (y viceversa)")

def calcular_rango_horario(ultimas_horas):
    """Calcula hora_desde y hora_hasta para las ultimas N horas desde ahora."""
    if not ultimas_horas:
        return None, None
    ahora = datetime.now()
    hace_n = ahora - timedelta(hours=ultimas_horas)
    hora_desde = hace_n.strftime("%Y-%m-%d %H:%M")
    hora_hasta = ahora.strftime("%Y-%m-%d %H:%M")
    return hora_desde, hora_hasta

def date_in_range(fname, fecha, fecha_desde, fecha_hasta):
    """True si archivo pasa filtro de fecha."""
    if not (fecha or fecha_desde or fecha_hasta):
        return True
    m = RE_FILE_DATE.search(fname)
    if not m:
        return True
    d = m.group(1)
    if fecha:
        return d == fecha
    if fecha_desde and d < fecha_desde:
        return False
    if fecha_hasta and d > fecha_hasta:
        return False
    return True

def is_full_datetime(s):
    """True si el string es 'YYYY-MM-DD HH:MM' (rango que cruza medianoche) vs solo 'HH:MM'."""
    return bool(s) and bool(RE_FULL_DATETIME.match(s))

def time_in_range(ts, hora_desde, hora_hasta):
    """True si timestamp cae en rango horario obligatorio.
    Si hora_desde/hasta son 'HH:MM', compara solo hora del dia.
    Si son 'YYYY-MM-DD HH:MM', compara fecha+hora (permite cruzar medianoche)."""
    if not (hora_desde or hora_hasta):
        return True
    full = is_full_datetime(hora_desde) or is_full_datetime(hora_hasta)
    key = ts[:16] if full else ts[11:16]
    if hora_desde and key < hora_desde:
        return False
    if hora_hasta and key > hora_hasta:
        return False
    return True

# --------------------------------------------------------------------------
# 2) SFTP (reutilizado del skill existente)
# --------------------------------------------------------------------------
def find_crypto_wrapper():
    pattern = os.path.join(SCRIPT_DIR, "crypto", "CryptoWrapper", "bin", "*", "net*", "CryptoWrapper.dll")
    matches = glob.glob(pattern)
    if not matches:
        sys.exit("No se encontro CryptoWrapper.dll compilado. Compilalo con:\n"
                  f"  dotnet build \"{os.path.join(SCRIPT_DIR, 'crypto', 'CryptoWrapper')}\"")
    return matches[0]

def _run_crypto(mode, key, text):
    dll = find_crypto_wrapper()
    proc = subprocess.run(["dotnet", dll, mode], input=f"{key}\n{text}\n",
                           capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        accion = "cifrar" if mode == "encrypt" else "descifrar"
        sys.exit(f"Fallo al {accion} via CryptoWrapper: {proc.stderr.strip()}")
    return proc.stdout.strip()

def crypto_encrypt(key, plaintext):
    return _run_crypto("encrypt", key, plaintext)

def crypto_decrypt(key, payload):
    return _run_crypto("decrypt", key, payload)

def load_saved_credentials(path):
    if not os.path.exists(path):
        return None
    with open(path, encoding="utf-8") as fh:
        return json.load(fh)

def save_credentials(path, host, port, usuario, auth, crypto_key, password_enc):
    os.makedirs(os.path.dirname(path), exist_ok=True)
    data = {"host": host, "port": port, "usuario": usuario, "auth": auth,
            "crypto_key": crypto_key, "password_enc": password_enc}
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(data, fh, indent=2, ensure_ascii=False)

def connect_sftp(args):
    creds = load_saved_credentials(args.credenciales) if not args.olvidar_credenciales else None
    host = args.host or (creds and creds.get("host")) or input("Host SFTP: ").strip()
    port_raw = args.port or (creds and str(creds.get("port"))) or input("Puerto SFTP [22]: ").strip()
    port = int(port_raw) if port_raw else 22
    user = args.usuario or (creds and creds.get("usuario")) or input("Usuario SFTP: ").strip()
    metodo = (args.auth or (creds and creds.get("auth"))
              or input("Autenticacion, 'password' o 'key' [password]: ").strip().lower() or "password")

    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())

    password = None
    if metodo == "key":
        key_path = input("Ruta de la llave privada: ").strip()
        key_pass = getpass.getpass("Passphrase de la llave (Enter si no tiene): ") or None
        pkey = None
        for loader in (paramiko.RSAKey, paramiko.Ed25519Key, paramiko.ECDSAKey):
            try:
                pkey = loader.from_private_key_file(key_path, password=key_pass)
                break
            except Exception:
                continue
        if pkey is None:
            sys.exit("No se pudo cargar la llave privada (formato no reconocido o passphrase incorrecta).")
        client.connect(hostname=host, port=port, username=user, pkey=pkey, timeout=30)
    else:
        if creds and creds.get("password_enc") and creds.get("crypto_key"):
            password = crypto_decrypt(creds["crypto_key"], creds["password_enc"])
            print(f"      Password leida (descifrada) desde {args.credenciales}")
        else:
            password = getpass.getpass("Password SFTP: ")
        client.connect(hostname=host, port=port, username=user, password=password, timeout=30)

    if args.guardar_credenciales and metodo == "password":
        crypto_key = (creds and creds.get("crypto_key")) or secrets.token_urlsafe(32)
        password_enc = crypto_encrypt(crypto_key, password)
        save_credentials(args.credenciales, host, port, user, metodo, crypto_key, password_enc)
        print(f"      Credenciales guardadas (password cifrada AES-256-GCM) en {args.credenciales}")

    return client, client.open_sftp()

def load_remote_paths(path, categoria, api=None, ruta=None):
    if not os.path.exists(path):
        sys.exit(f"No existe el archivo de topologia remota: {path}")
    wanted = None if (not categoria or categoria.lower() == "todos") else {
        c.strip().upper() for c in categoria.split(",")
    }
    api_wanted = {a.strip().lower() for a in api.split(",")} if api else None
    out = []
    with open(path, encoding="utf-8") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            parts = line.split("\t")
            if len(parts) != 2:
                continue
            cat, remote = parts[0].strip(), parts[1].strip()
            if wanted and cat.upper() not in wanted:
                continue
            if api_wanted and remote.rstrip("/").rsplit("/", 1)[-1].lower() not in api_wanted:
                continue
            if ruta and not fnmatch.fnmatch(remote, ruta):
                continue
            out.append((cat, remote))
    if not out:
        sys.exit(f"Los filtros dados no dejaron ninguna ruta remota. Revisa remote_paths.txt.")
    return out

def walk_remote(sftp, base_path, max_depth=3):
    out = []
    stack = [(base_path, 0)]
    while stack:
        path, depth = stack.pop()
        try:
            entries = sftp.listdir_attr(path)
        except (FileNotFoundError, IOError, OSError):
            print(f"  [AVISO] No se pudo listar: {path}")
            continue
        for entry in entries:
            full = path.rstrip("/") + "/" + entry.filename
            if statmod.S_ISDIR(entry.st_mode):
                if depth < max_depth:
                    stack.append((full, depth + 1))
            elif (fnmatch.fnmatch(entry.filename, "*.log")
                  or fnmatch.fnmatch(entry.filename, "*.log.gz")
                  or RE_FILE_DATE.search(entry.filename)):
                out.append((full, entry))
    return out

def long_path(p):
    if os.name != "nt":
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p

def cleanup_old_cache(cache_dir, max_days):
    if not max_days or max_days <= 0 or not os.path.isdir(cache_dir):
        return 0, 0
    cutoff = time.time() - max_days * 86400
    removed, freed = 0, 0
    for root, _dirs, files in os.walk(cache_dir, topdown=False):
        for name in files:
            path = os.path.join(root, name)
            try:
                st = os.stat(long_path(path))
            except OSError:
                continue
            if st.st_mtime < cutoff:
                try:
                    os.remove(long_path(path))
                    removed += 1
                    freed += st.st_size
                except OSError:
                    pass
        if root != cache_dir:
            try:
                if not os.listdir(root):
                    os.rmdir(root)
            except OSError:
                pass
    return removed, freed

def download_candidates(client, candidates, cache_dir, small_workers=5, large_workers=2, large_threshold_mb=300):
    large_threshold = large_threshold_mb * 1024 * 1024
    small = [(p, a) for p, a in candidates if a.st_size < large_threshold]
    large = [(p, a) for p, a in candidates if a.st_size >= large_threshold]

    transport = client.get_transport()
    pool_size = max(1, small_workers) + max(1, large_workers)
    sftp_pool = queue.Queue()
    for _ in range(pool_size):
        sftp_pool.put(paramiko.SFTPClient.from_transport(transport))

    print_lock = threading.Lock()
    total = len(candidates)
    done = {"n": 0}

    def fmt_mb(n):
        return f"{n / (1024 * 1024):.1f} MB"

    def worker(remote_path, attr):
        rel = remote_path.lstrip("/")
        local_path = long_path(os.path.join(cache_dir, rel.replace("/", os.sep)))
        os.makedirs(os.path.dirname(local_path), exist_ok=True)
        sftp = sftp_pool.get()
        try:
            if os.path.exists(local_path) and os.path.getsize(local_path) == attr.st_size:
                with print_lock:
                    print(f"  en cache:    {remote_path}")
            else:
                with print_lock:
                    print(f"  descargando: {remote_path} ({fmt_mb(attr.st_size)})")
                last_pct = {"v": -10}

                def progress(transferred, total_bytes):
                    pct = int(transferred * 100 / total_bytes) if total_bytes else 100
                    if pct - last_pct["v"] >= 10 or transferred >= total_bytes:
                        last_pct["v"] = pct
                        with print_lock:
                            print(f"    [{pct:3d}%] ({fmt_mb(transferred)}/{fmt_mb(total_bytes)})")

                sftp.get(remote_path, local_path, callback=progress)
        finally:
            sftp_pool.put(sftp)
        with print_lock:
            done["n"] += 1
            print(f"  [{done['n']}/{total}] listo: {remote_path}")
        return (remote_path, local_path)

    local_files = []
    with ThreadPoolExecutor(max_workers=max(1, small_workers)) as ex_small, \
         ThreadPoolExecutor(max_workers=max(1, large_workers)) as ex_large:
        futures = [ex_small.submit(worker, p, a) for p, a in small]
        futures += [ex_large.submit(worker, p, a) for p, a in large]
        for fut in as_completed(futures):
            local_files.append(fut.result())

    while not sftp_pool.empty():
        sftp_pool.get().close()

    return local_files

def resolve_report_path(name_arg, default_name):
    reports_dir = os.path.join(SCRIPT_DIR, "_reports", datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(reports_dir, exist_ok=True)
    filename = os.path.basename(name_arg) if name_arg else default_name
    return os.path.join(reports_dir, filename)

# --------------------------------------------------------------------------
# 3) Local mode
# --------------------------------------------------------------------------
def open_local(path, encoding):
    opener = gzip.open if path.endswith(".gz") else open
    return opener(long_path(path), "rt", encoding=encoding, errors="replace")

def iter_log_files(logs_dir):
    plain = sorted(glob.glob(os.path.join(logs_dir, "**", "*.log"), recursive=True))
    gz = sorted(glob.glob(os.path.join(logs_dir, "**", "*.log.gz"), recursive=True))
    plain_set = set(plain)
    gz = [g for g in gz if g[:-3] not in plain_set]
    files = plain + gz
    seen = set()
    return [x for x in files if not (x in seen or seen.add(x))]

def local_server_label(path, logs_dir):
    return os.path.relpath(path, logs_dir).replace("\\", "/")

# --------------------------------------------------------------------------
# 4) SCAN FOR ERRORS (la logica nueva)
# --------------------------------------------------------------------------
def is_error_line(text):
    """True si la linea contiene un error segun los criterios:
    - statuscode != 200
    - code != "00" y != "RA2"
    - mensaje contiene "Error" o "exception"
    """
    low = text.lower()

    # Criterio 1: statuscode != 200
    m_status = RE_STATUS_CODE.search(text)
    if m_status:
        status = m_status.group(1)
        if status != "200":
            return True

    # Criterio 2: code != "00" y != "RA2"
    m_code = RE_CODE.search(text)
    if m_code:
        code = m_code.group(1)
        if code != "00" and code != "RA2":
            return True

    # Criterio 3: "Error" o "exception" en el mensaje
    if "error" in low or "exception" in low:
        return True

    return False

def extract_merchant(msg):
    m = RE_MERCHANT.search(msg)
    return m.group(1) if m else None

def extract_error_code(msg):
    """Extrae el código de error de la línea (statuscode, code, o 'Error'/'exception')."""
    # Intenta extraer statuscode primero
    m_status = RE_STATUS_CODE.search(msg)
    if m_status:
        return f"statuscode={m_status.group(1)}"

    # Intenta extraer code ISO 8583
    m_code = RE_CODE.search(msg)
    if m_code:
        return f"code={m_code.group(1)}"

    # Si contiene Error o exception, retorna eso
    low = msg.lower()
    if "exception" in low:
        return "exception"
    if "error" in low:
        return "error"

    return "unknown"

def scan_for_errors(local_files, encoding, fecha=None, fecha_desde=None, fecha_hasta=None,
                    hora_desde=None, hora_hasta=None, comercio_filter=None):
    """Escanea todos los logs buscando lineas con error. Retorna tupla:
    (errors_by_txid, last_timestamp_by_file) donde:
    - errors_by_txid: {
        "txid1": [
            ("server_label", lineno, "timestamp", "error_code", "linea_completa"),
            ...
        ],
        "txid2": [...],
    }
    - last_timestamp_by_file: {"server_label1": "YYYY-MM-DD HH:MM:SS.fff", ...}
    """
    errors_by_txid = {}
    last_timestamp_by_file = {}

    for remote_label, local_path in local_files:
        max_ts = None
        with open_local(local_path, encoding) as fh:
            for lineno, line in enumerate(fh, 1):
                # Filtro de fecha
                m_ts = TS_RE.match(line)
                if m_ts:
                    ts = m_ts.group(1)
                    if max_ts is None or ts > max_ts:
                        max_ts = ts
                    if not time_in_range(ts, hora_desde, hora_hasta):
                        continue
                else:
                    continue

                # Detectar error
                if not is_error_line(line):
                    continue

                # Extraer txid
                parts = line.split("|", 3)
                if len(parts) < 4:
                    continue
                txid = parts[3].split("|", 1)[0].strip()
                if not txid or " " in txid or len(txid) > 60:
                    continue

                # Filtro de comercio (si se dio)
                if comercio_filter:
                    merchant = extract_merchant(line)
                    if not merchant or merchant != comercio_filter:
                        continue

                # Extraer código de error
                error_code = extract_error_code(line)

                if txid not in errors_by_txid:
                    errors_by_txid[txid] = []
                errors_by_txid[txid].append((remote_label, lineno, ts, error_code, line.rstrip("\n")))

        if max_ts:
            last_timestamp_by_file[remote_label] = max_ts
        print(f"  escaneado: {remote_label}")

    return errors_by_txid, last_timestamp_by_file

# --------------------------------------------------------------------------
# 5) Reporte basico (Excel + HTML)
# --------------------------------------------------------------------------
def generate_report(errors_by_txid, output_xlsx, output_html, last_timestamp_by_file=None):
    """Genera reportes en Excel y HTML con codigo de error y ultima hora de log."""
    if not errors_by_txid:
        print("Sin errores encontrados.")
        return

    if last_timestamp_by_file is None:
        last_timestamp_by_file = {}

    # Excel basico
    wb = Workbook()
    ws = wb.active
    ws.title = "Errores"

    headers = ["TransactionId", "Servidor", "Linea", "Timestamp", "Codigo de Error", "Ultima Hora Log", "Mensaje"]
    ws.append(headers)

    row_num = 2
    for txid in sorted(errors_by_txid.keys()):
        for server, lineno, ts, error_code, line_text in errors_by_txid[txid]:
            msg = line_text[:MAX_CELL_LEN] if len(line_text) > MAX_CELL_LEN else line_text
            last_ts = last_timestamp_by_file.get(server, "N/A")
            ws.append([txid, server, lineno, ts, error_code, last_ts, msg])
            row_num += 1

    wb.save(output_xlsx)
    print(f"Reporte Excel guardado: {output_xlsx}")

    # HTML basico
    html_lines = [
        "<!DOCTYPE html><html><head>",
        "<meta charset='utf-8'>",
        "<title>Errores de Transacciones</title>",
        "<style>",
        "body { font-family: Arial, sans-serif; margin: 20px; }",
        "h1 { color: #333; }",
        "table { border-collapse: collapse; width: 100%; margin-top: 20px; }",
        "th, td { border: 1px solid #ddd; padding: 8px; text-align: left; font-size: 0.9em; }",
        "th { background-color: #f2f2f2; font-weight: bold; }",
        "tr:nth-child(even) { background-color: #f9f9f9; }",
        ".timestamp { color: #666; font-size: 0.85em; }",
        ".txid { font-weight: bold; color: #333; }",
        ".error-code { color: #d9534f; font-weight: bold; }",
        ".last-ts { color: #5cb85c; font-weight: bold; }",
        "</style></head><body>",
        f"<h1>Errores de Transacciones - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</h1>",
        f"<p>Total de TransactionIds con error: {len(errors_by_txid)}</p>",
        "<table><tr><th>TransactionId</th><th>Servidor</th><th>Linea</th><th>Timestamp</th><th>Codigo de Error</th><th>Ultima Hora Log</th><th>Mensaje</th></tr>",
    ]

    for txid in sorted(errors_by_txid.keys()):
        for server, lineno, ts, error_code, line_text in errors_by_txid[txid]:
            msg = esc(line_text[:MAX_CELL_LEN]) if len(line_text) > MAX_CELL_LEN else esc(line_text)
            last_ts = last_timestamp_by_file.get(server, "N/A")
            html_lines.append(f"<tr><td class='txid'>{esc(txid)}</td><td>{esc(server)}</td>"
                              f"<td>{lineno}</td><td class='timestamp'>{esc(ts)}</td>"
                              f"<td class='error-code'>{esc(error_code)}</td>"
                              f"<td class='last-ts'>{esc(last_ts)}</td><td>{msg}</td></tr>")

    html_lines += ["</table></body></html>"]
    html_content = "\n".join(html_lines)
    with open(output_html, "w", encoding="utf-8") as fh:
        fh.write(html_content)
    print(f"Reporte HTML guardado: {output_html}")

def esc(s):
    """Escapa caracteres HTML."""
    return s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;") \
            .replace('"', "&quot;").replace("'", "&#39;")

# --------------------------------------------------------------------------
# MAIN
# --------------------------------------------------------------------------
def main():
    args = parse_args()
    validate_args(args)

    # Calcular rango horario si se da --ultimas-N-horas
    if args.ultimas_N_horas:
        args.hora_desde, args.hora_hasta = calcular_rango_horario(args.ultimas_N_horas)
        print(f"Buscando ultimas {args.ultimas_N_horas} horas: {args.hora_desde} a {args.hora_hasta}")

    default_encoding = "cp1252" if args.logs_dir else "utf-8"
    encoding = args.encoding or default_encoding

    print(f"\n=== BUSQUEDA DE ERRORES EN TRANSACCIONES ===")
    print(f"Fecha: {args.fecha or f'{args.fecha_desde} a {args.fecha_hasta}'}")
    print(f"Horas: {args.hora_desde or '--'} a {args.hora_hasta or '--'}")
    print(f"Comercio: {args.comercio or 'TODOS'}")
    print(f"Modo: {'LOCAL' if args.logs_dir else 'SFTP'}")
    print(f"Encoding: {encoding}\n")

    # Modo LOCAL
    if args.logs_dir:
        if not os.path.isdir(args.logs_dir):
            sys.exit(f"No existe carpeta local: {args.logs_dir}")
        print(f"Explorando carpeta local: {args.logs_dir}")
        local_files = [(local_server_label(f, args.logs_dir), f) for f in iter_log_files(args.logs_dir)]
        if not local_files:
            sys.exit(f"No se encontraron logs en {args.logs_dir}")
    else:
        # Modo SFTP
        print("Conectando a SFTP...")
        client, sftp = connect_sftp(args)

        print("\nCargando topologia remota...")
        remote_paths = load_remote_paths(args.remote_paths, args.categoria, args.api, args.ruta)
        print(f"  {len(remote_paths)} rutas remotas cargadas")

        print("\nLimpiando cache antigua...")
        removed, freed = cleanup_old_cache(args.cache_dir, args.cache_max_days)
        if removed:
            freed_mb = freed / (1024 * 1024)
            print(f"  Borrados {removed} archivos, liberados {freed_mb:.1f} MB")

        if args.limpiar_cache:
            import shutil
            if os.path.exists(args.cache_dir):
                shutil.rmtree(args.cache_dir)
            print("  Cache limpiada completamente")
            os.makedirs(args.cache_dir, exist_ok=True)

        print("\nBuscando archivos candidatos en SFTP...")
        all_files = []
        for cat, remote_base in remote_paths:
            print(f"  Recorriendo {cat} {remote_base}...")
            files = walk_remote(sftp, remote_base, args.max_depth)
            all_files.extend(files)

        candidates = [(p, a) for p, a in all_files if date_in_range(os.path.basename(p), args.fecha, args.fecha_desde, args.fecha_hasta)]
        if not candidates:
            print("\nSin archivos candidatos que descarg ar en el rango de fecha dado.")
            sftp.close()
            client.close()
            return

        print(f"\n  {len(candidates)} archivos candidatos encontrados")
        print("\nDescargando archivos...")
        local_files = download_candidates(client, candidates, args.cache_dir,
                                          args.paralelo_livianos, args.paralelo_pesados,
                                          args.umbral_pesado_mb)
        sftp.close()
        client.close()

    print("\n=== ESCANEANDO LOGS EN BUSCA DE ERRORES ===")
    errors_by_txid, last_timestamp_by_file = scan_for_errors(local_files, encoding,
                                     args.fecha, args.fecha_desde, args.fecha_hasta,
                                     args.hora_desde, args.hora_hasta, args.comercio)

    if not errors_by_txid:
        print("\nSin errores encontrados en los logs analiz ados.")
        return

    print(f"\n=== RESULTADOS ===")
    print(f"TransactionIds con error: {len(errors_by_txid)}")
    for txid in sorted(errors_by_txid.keys()):
        print(f"  {txid}: {len(errors_by_txid[txid])} lineas de error")

    # Generar reportes
    output_xlsx = resolve_report_path(args.output_xlsx, "Errores_Transacciones.xlsx")
    output_html = resolve_report_path(args.output_html, "Errores_Transacciones.html")

    print("\n=== GENERANDO REPORTES ===")
    generate_report(errors_by_txid, output_xlsx, output_html, last_timestamp_by_file)

    print("\n[OK] Listo!")

if __name__ == "__main__":
    main()
