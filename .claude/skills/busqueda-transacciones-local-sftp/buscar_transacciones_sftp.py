#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Busca una o mas transacciones (por TransactionId, Comercio, Nro Orden, Codigo
de Autorizacion y/o Monto) en los logs NLog de la plataforma ecommerce, en
DOS modos de origen de datos: (A) MODO SFTP -- se conecta ella misma a un
servidor remoto, recorre la topologia de carpetas remotas conocida (ver
remote_paths.txt) y descarga solo los archivos candidatos a una cache local
reutilizable; (B) MODO LOCAL (--logs-dir) -- analiza directo una carpeta con
logs ya descargados a mano, sin conectar a nada. Ambos modos comparten el
mismo motor de busqueda/clasificacion y generan el mismo reporte de evidencia
en Excel y HTML. (Fusiona lo que antes era la skill separada
"trazabilidad-transacciones", retirada el 2026-09-11 al unificarse aqui.)

Uso tipico:
    python buscar_transacciones_sftp.py --input input.txt --fecha 2026-07-14
    python buscar_transacciones_sftp.py --txid LP260714-195608076-KF5JXR --fecha-desde 2026-07-10 --fecha-hasta 2026-07-14
    python buscar_transacciones_sftp.py --orden 195608076 --comercio 4080783 --autorizacion 660568 --monto 370.00 --categoria PUBLICO --fecha 2026-07-14 --hora-desde 09:00 --hora-hasta 09:30
    python buscar_transacciones_sftp.py --logs-dir "C:/ruta/a/logs" --txid 174544565028023 --comercio 4078371

Credenciales SFTP: por defecto se piden de forma interactiva por consola
(host, puerto, usuario y password u llave privada), con entrada oculta
(getpass) para la password. Opcionalmente, con --guardar-credenciales se
guarda host/usuario/password en un archivo JSON local (--credenciales), con
la password cifrada AES-256-GCM (key derivada por SHA-256) via el wrapper
.NET en crypto/CryptoWrapper. La key de cifrado tambien se guarda en ese
mismo archivo: esto es proteccion contra un vistazo casual del archivo, NO
un control de acceso real (quien tenga el archivo puede descifrar la
password). Ver seccion "Credenciales guardadas (cifradas)" en SKILL.md.

Requisitos: Python 3.8+, paramiko, openpyxl, y el SDK/runtime de .NET 10
(para el wrapper de cifrado, solo si se usa --guardar-credenciales)
    pip install paramiko openpyxl
"""
import argparse, fnmatch, getpass, glob, gzip, io, json, os, queue, re, secrets, subprocess, sys, threading, stat as statmod
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

try:
    import paramiko
except ImportError:
    sys.exit("Falta la libreria 'paramiko'. Instala con: pip install paramiko")

from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, Border, Side

SCRIPT_DIR_FOR_IMPORTS = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR_FOR_IMPORTS, "..", "reportes-html-izipay")))
sys.path.insert(0, os.path.abspath(os.path.join(SCRIPT_DIR_FOR_IMPORTS, "..", "reportes-excel-izipay")))
from izipay_report import ReportBuilder, esc, timestamp_now
from izipay_excel import (BRAND, header_font, header_fill, write_title_block, style_header_row,
                           style_status_cell, status_font, set_column_widths, freeze_header)

MAX_CELL_LEN = 30000
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DEFAULT_REMOTE_PATHS = os.path.join(SCRIPT_DIR, "remote_paths.txt")

# --------------------------------------------------------------------------
# 1) Criterios de busqueda: input.txt (OPCIONAL) + flags sueltos
#    (mismo esquema/sinonimos que trazabilidad-transacciones, para que un
#    input.txt ya escrito sirva igual en ambas skills)
# --------------------------------------------------------------------------
LABELS = {
    "comercio": "merchants", "codigo de comercio": "merchants", "codigo": "merchants",
    "cod comercio": "merchants", "merchant": "merchants", "merchantcode": "merchants",
    "orden": "orders", "nro orden": "orders", "numero de orden": "orders",
    "order": "orders", "ordernumber": "orders",
    "txid": "txids", "transactionid": "txids", "tx": "txids", "transaction": "txids",
    "autorizacion": "authcodes", "codigo de autorizacion": "authcodes",
    "cod autorizacion": "authcodes", "authcode": "authcodes", "codauth": "authcodes",
    "codigoautorizacion": "authcodes",
    "monto": "amounts", "amount": "amounts", "importe": "amounts",
}
CATS = ("merchants", "orders", "authcodes", "amounts", "txids")


def empty_filters():
    return {c: set() for c in CATS} | {"loose": set()}


def parse_input(path, encoding="utf-8"):
    f = empty_filters()
    if not path or not os.path.exists(path):
        return f
    with open(path, encoding=encoding, errors="replace") as fh:
        for raw in fh:
            line = raw.strip()
            if not line or line.startswith("#"):
                continue
            m = re.match(r"([\wáéíóúñ ]+?)\s*[:=]\s*(.+)$", line)
            key = re.sub(r"\s+", " ", m.group(1).strip().lower()) if m else None
            if m and key in LABELS:
                f[LABELS[key]].add(m.group(2).strip())
            else:
                f["loose"].add(line)
    return f


def add_cli_filters(f, args):
    if args.txid:
        f["txids"] |= set(args.txid)
    if args.comercio:
        f["merchants"] |= set(args.comercio)
    if args.orden:
        f["orders"] |= set(args.orden)
    if args.autorizacion:
        f["authcodes"] |= set(args.autorizacion)
    if args.monto:
        f["amounts"] |= set(args.monto)
    return f


def amount_variants(v):
    v = v.strip()
    variants = {v}
    try:
        num = float(v.replace(",", "."))
        variants.add(f"{num:.2f}")
        if num == int(num):
            variants.add(str(int(num)))
    except ValueError:
        pass
    return variants


def build_needle_groups(f):
    groups = []
    for cat in ("merchants", "orders", "authcodes"):
        if f[cat]:
            groups.append(set(f[cat]))
    if f["amounts"]:
        vs = set()
        for a in f["amounts"]:
            vs |= amount_variants(a)
        groups.append(vs)
    if f["loose"]:
        groups.append(set(f["loose"]))
    return groups or None


def filters_desc(f):
    parts = []
    for cat, label in (("merchants", "comercio"), ("orders", "orden"),
                        ("authcodes", "cod. autorizacion"), ("amounts", "monto"),
                        ("txids", "txid"), ("loose", "libre")):
        if f[cat]:
            parts.append(f"{label}={sorted(f[cat])}")
    return " ; ".join(parts) if parts else "(sin criterios)"


# --------------------------------------------------------------------------
# 2) Topologia remota (remote_paths.txt) + filtro por fecha/categoria
# --------------------------------------------------------------------------
RE_FILE_DATE = re.compile(r"(\d{4}-\d{2}-\d{2})")


def load_remote_paths(path, categoria, api=None, ruta=None):
    """Lee remote_paths.txt -> lista de (categoria, ruta_remota) filtrada por
    categoria (grupo de servidores), api (nombre de la carpeta final, ej.
    'authorization', 'cancel', 'notification.dispatcher1') y/o ruta (patron
    glob sobre la ruta remota completa, ej. '*/SVPRDLW364/*')."""
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
        sys.exit(f"Los filtros dados (categoria='{categoria}', api='{api}', ruta='{ruta}') "
                  f"no dejaron ninguna ruta remota. Revisa remote_paths.txt.")
    return out


def date_in_range(fname, fecha, fecha_desde, fecha_hasta):
    """True si el archivo pasa el filtro de fecha (o si no se pudo determinar
    la fecha del nombre, en cuyo caso se conserva y se avisa por consola)."""
    if not (fecha or fecha_desde or fecha_hasta):
        return True
    m = RE_FILE_DATE.search(fname)
    if not m:
        return True  # sin fecha en el nombre: se conserva, mejor prevenir que descartar
    d = m.group(1)
    if fecha:
        return d == fecha
    if fecha_desde and d < fecha_desde:
        return False
    if fecha_hasta and d > fecha_hasta:
        return False
    return True


RE_FULL_DATETIME = re.compile(r"^\d{4}-\d{2}-\d{2} \d{2}:\d{2}$")


def is_full_datetime(s):
    """True si el string es 'YYYY-MM-DD HH:MM' (rango horario que cruza
    medianoche) en vez de solo 'HH:MM' (rango dentro de un mismo dia)."""
    return bool(s) and bool(RE_FULL_DATETIME.match(s))


def time_in_range(ts, hora_desde, hora_hasta):
    """True si el timestamp de la linea cae dentro del rango horario
    obligatorio dado. ts tiene formato 'YYYY-MM-DD HH:MM:SS.fff'.

    Si hora_desde/hora_hasta vienen en formato 'HH:MM', compara solo la hora
    del dia (mismo rango aplicado a cada dia del filtro de fecha). Si vienen
    en formato completo 'YYYY-MM-DD HH:MM', compara fecha+hora juntas, lo que
    permite un rango continuo que cruce medianoche (ej. 19:00 de un dia hasta
    08:00 del dia siguiente) sin traer horas de fuera de ese rango."""
    if not (hora_desde or hora_hasta):
        return True
    full = is_full_datetime(hora_desde) or is_full_datetime(hora_hasta)
    key = ts[:16] if full else ts[11:16]
    if hora_desde and key < hora_desde:
        return False
    if hora_hasta and key > hora_hasta:
        return False
    return True


def find_crypto_wrapper():
    """Ubica CryptoWrapper.dll ya compilado (dotnet build en crypto/CryptoWrapper)."""
    pattern = os.path.join(SCRIPT_DIR, "crypto", "CryptoWrapper", "bin", "*", "net*", "CryptoWrapper.dll")
    matches = glob.glob(pattern)
    if not matches:
        sys.exit("No se encontro CryptoWrapper.dll compilado. Compilalo una vez con:\n"
                  f"  dotnet build \"{os.path.join(SCRIPT_DIR, 'crypto', 'CryptoWrapper')}\"")
    return matches[0]


def _run_crypto(mode, key, text):
    """Invoca CryptoWrapper.dll (AES-256-GCM, key derivada SHA-256) por stdin;
    key/texto nunca se pasan por argumentos de linea de comandos."""
    dll = find_crypto_wrapper()
    proc = subprocess.run(["dotnet", dll, mode], input=f"{key}\n{text}\n",
                           capture_output=True, text=True, encoding="utf-8")
    if proc.returncode != 0:
        accion = "cifrar" if mode == "encrypt" else "descifrar"
        sys.exit(f"Fallo al {accion} via CryptoWrapper: {proc.stderr.strip()}")
    return proc.stdout


def crypto_encrypt(key, plaintext):
    return _run_crypto("encrypt", key, plaintext)


def crypto_decrypt(key, payload):
    return _run_crypto("decrypt", key, payload)


def resolve_report_path(name_arg, default_name):
    """Fuerza que los reportes finales queden en _reports/<fecha-de-hoy>/ de
    esta skill, usando el nombre dado (o el default) como nombre de archivo."""
    reports_dir = os.path.join(SCRIPT_DIR, "_reports", datetime.now().strftime("%Y-%m-%d"))
    os.makedirs(reports_dir, exist_ok=True)
    filename = os.path.basename(name_arg) if name_arg else default_name
    return os.path.join(reports_dir, filename)


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
    """Resuelve credenciales (archivo guardado > flags > consola) y abre la
    sesion SFTP. Sin --guardar-credenciales, el comportamiento es identico al
    original: todo se pide por consola y nada se guarda en disco."""
    creds = load_saved_credentials(args.credenciales) if not args.olvidar_credenciales else None

    host = args.host or (creds and creds.get("host")) or input("Host SFTP: ").strip()
    port_raw = args.port or (creds and str(creds.get("port"))) or input("Puerto SFTP [22]: ").strip()
    port = int(port_raw) if port_raw else 22
    user = args.usuario or (creds and creds.get("usuario")) or input("Usuario SFTP: ").strip()
    metodo = (args.auth or (creds and creds.get("auth"))
              or input("Autenticacion, 'password' o 'key' [password]: ").strip().lower() or "password")

    client = paramiko.SSHClient()
    client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
    # Nota de seguridad: AutoAddPolicy no valida la huella (fingerprint) del
    # host contra un known_hosts previo. Aceptable en una red interna de
    # confianza; si se requiere verificacion estricta, reemplazar por
    # paramiko.RejectPolicy() + known_hosts cargado de antemano.

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
        print("      ADVERTENCIA: la key de cifrado vive en el mismo archivo -> es proteccion "
              "contra un vistazo casual, no control de acceso real. Protege el archivo (permisos, no compartir).")

    return client, client.open_sftp()


def walk_remote(sftp, base_path, max_depth=3):
    """Recorre recursivamente (hasta max_depth) un directorio remoto y
    devuelve [(ruta_remota_completa, attr)] de archivos *.log / *.log.gz, o
    sin extension pero con fecha embebida en el nombre (logs aun no
    rotados/comprimidos por NLog, ej. 'nlog-ApiPublicCancel-all-2026-08-04')."""
    out = []
    stack = [(base_path, 0)]
    while stack:
        path, depth = stack.pop()
        try:
            entries = sftp.listdir_attr(path)
        except (FileNotFoundError, IOError, OSError):
            print(f"  [AVISO] No se pudo listar (no existe o sin permiso): {path}")
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
    """En Windows, antepone el prefijo \\\\?\\ para saltar el limite clasico
    de MAX_PATH (260 caracteres). La topologia remota real puede anidar
    carpetas "espejo" con el mismo nombre que el archivo (ej. servidores que
    duplican el log del dia dentro de una subcarpeta homonima), lo que empuja
    la ruta local de cache justo al borde o mas alla de ese limite y provoca
    FileNotFoundError al escribir, aunque el directorio si se haya creado."""
    if os.name != "nt":
        return p
    p = os.path.abspath(p)
    if p.startswith("\\\\?\\"):
        return p
    if p.startswith("\\\\"):
        return "\\\\?\\UNC\\" + p[2:]
    return "\\\\?\\" + p


def download_candidates(client, candidates, cache_dir, small_workers=5, large_workers=2, large_threshold_mb=300):
    """Descarga a cache_dir en paralelo (si no esta ya cacheado con el mismo
    tamano) y devuelve [(remote_path, local_path), ...]. Concurrencia
    adaptativa por peso: los archivos livianos (< large_threshold_mb) usan
    hasta `small_workers` descargas simultaneas; los pesados (>=) usan hasta
    `large_workers` (menos, para no saturar el ancho de banda/memoria con
    varios archivos de varios GB a la vez). Cada hilo abre su propio canal
    SFTP sobre el mismo Transport ya autenticado (paramiko.SFTPClient no es
    thread-safe entre hilos, pero el Transport si admite multiples canales).
    Imprime progreso (cada ~10%) para que la espera en archivos grandes no
    parezca colgada."""
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
                            print(f"    [{pct:3d}%] {remote_path} ({fmt_mb(transferred)}/{fmt_mb(total_bytes)})")

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


# --------------------------------------------------------------------------
# 3) Parseo de lineas NLog (identico formato al de trazabilidad-transacciones)
# --------------------------------------------------------------------------
TS_RE = re.compile(r"^(\d{4}-\d{2}-\d{2} \d{2}:\d{2}:\d{2}\.\d+)\|(?:(\d+)\|)?(\w+)\|(.*)$")

ROW_TXID_RE = re.compile(r"\|(\d{8,20})\|")


def extract_row_txid(msg, txid):
    """Algunas lineas (ej. flujo Push/PLIN con el switch bancario) traen un
    identificador numerico distinto al TransactionId principal -- lo
    devuelve como evidencia extra si aparece, o el propio txid si no."""
    m = ROW_TXID_RE.search(msg)
    if m and m.group(1) != txid:
        return m.group(1)
    return txid


def open_local(path, encoding):
    # long_path() evita FileNotFoundError por el limite MAX_PATH (260) de
    # Windows: aplica tanto a la cache SFTP como al modo --logs-dir, cuyas
    # carpetas "espejo" (ver long_path()) pueden anidar rutas igual de largas.
    opener = gzip.open if path.endswith(".gz") else open
    return opener(long_path(path), "rt", encoding=encoding, errors="replace")


def iter_log_files(logs_dir):
    """Modo --logs-dir (sin SFTP, logs ya descargados a mano): *.log (sin
    comprimir) y *.log.gz (comprimidos), recursivo. Si por algun motivo
    existen AMBOS para el mismo archivo (ej. copia descomprimida al lado
    del .gz original), se usa solo el .log para no contar el contenido dos
    veces. Portado de la skill trazabilidad-transacciones (2026-09-11)."""
    plain = sorted(glob.glob(os.path.join(logs_dir, "**", "*.log"), recursive=True))
    gz = sorted(glob.glob(os.path.join(logs_dir, "**", "*.log.gz"), recursive=True))
    plain_set = set(plain)
    gz = [g for g in gz if g[:-3] not in plain_set]
    files = plain + gz
    seen = set()
    return [x for x in files if not (x in seen or seen.add(x))]


def local_server_label(path, logs_dir):
    return os.path.relpath(path, logs_dir).replace("\\", "/")


def discover_txids(local_files, groups, encoding, hora_desde=None, hora_hasta=None):
    found = {}
    for remote_label, local_path in local_files:
        with open_local(local_path, encoding) as fh:
            for lineno, line in enumerate(fh, 1):
                if not all(any(v in line for v in g) for g in groups):
                    continue
                m = TS_RE.match(line)
                if m and not time_in_range(m.group(1), hora_desde, hora_hasta):
                    continue
                parts = line.split("|", 3)
                if len(parts) < 4:
                    continue
                txid = parts[3].split("|", 1)[0].strip()
                if not txid or " " in txid or len(txid) > 60:
                    continue
                found.setdefault(txid, (remote_label, lineno))
        print(f"  escaneado: {remote_label}")
    return found


def scan_by_txid(local_files, txids, encoding, hora_desde=None, hora_hasta=None):
    collected = {tx: [] for tx in txids}
    for remote_label, local_path in local_files:
        with open_local(local_path, encoding) as fh:
            for lineno, line in enumerate(fh, 1):
                for tx in txids:
                    if tx in line:
                        m = TS_RE.match(line)
                        if m and not time_in_range(m.group(1), hora_desde, hora_hasta):
                            continue
                        collected[tx].append((remote_label, lineno, line.rstrip("\n")))
        print(f"  escaneado: {remote_label}")
    return collected


# --------------------------------------------------------------------------
# 4) Extraccion de contexto + clasificacion (mismos patrones validados en
#    trazabilidad-transacciones, porque es el mismo formato NLog/plataforma)
# --------------------------------------------------------------------------
RE_MERCHANT = re.compile(r'"?[Mm]erchant[Cc]ode"?\s*:\s*"?(\d+)"?')
RE_ORDER = re.compile(r'"?[Oo]rder[Nn]umber"?\s*:\s*"([^"]+)"')
RE_AUTHCODE = re.compile(r'"(?:AuthorizationCode|CodeAuth)"\s*:\s*"([^"]+)"')
RE_AMOUNT = re.compile(r'"(?:Amount|AmountValidate)"\s*:\s*"([^"]*)"')
RE_REF = re.compile(r'"ReferenceNumber"\s*:\s*"([^"]+)"')


def extract_context(lines):
    ctx = {"merchant": None, "order": None, "authcode": None, "amount": None, "ref": None}
    for _, _, text in lines:
        if ctx["merchant"] is None:
            m = RE_MERCHANT.search(text)
            if m: ctx["merchant"] = m.group(1)
        if ctx["order"] is None:
            m = RE_ORDER.search(text)
            if m: ctx["order"] = m.group(1)
        if ctx["authcode"] is None:
            m = RE_AUTHCODE.search(text)
            if m and m.group(1): ctx["authcode"] = m.group(1)
        if ctx["amount"] is None:
            m = RE_AMOUNT.search(text)
            if m and m.group(1): ctx["amount"] = m.group(1)
        if ctx["ref"] is None:
            m = RE_REF.search(text)
            if m: ctx["ref"] = m.group(1)
        if all(ctx.values()):
            break
    return ctx


# Perfiles de hitos YA VALIDADOS contra logs reales de un flujo especifico
# (portado de la skill trazabilidad-transacciones el 2026-09-11, ver
# ALL_KNOWN_MILESTONES mas abajo). Para la topologia propia de esta skill
# (authorization/cancel/confirmation/orderinfo/qrnotification/security/
# notification.*/capture/payments/push/qr/yape/accountvalidate/autoreverse)
# todavia no hay un perfil dedicado, asi que build_milestones() cae siempre
# al modo generico por componente para esas -- pero para logs del flujo
# Push/PLIN (tipicamente investigados con --logs-dir sobre logs ya locales)
# se usa este perfil.
PUSH_MILESTONES = [
    (re.compile(r"Validation\s*(?:Application)?\|.*End\|?\s*response", re.I), "Validacion del pago", False),
    (re.compile(r"idTransactionInterbank", re.I), "Validacion con Interbank (switch)", False),
    (re.compile(r"\bSearch\|End response", re.I), "Consulta de estado por el comercio (Search)", True),
    (re.compile(r"\bPayment\|Begin\|request", re.I), "Recepcion del pago (callback Push)", False),
    (re.compile(r"\bGet Info\|End response", re.I), "Consulta de informacion de la transaccion", False),
    (re.compile(r"InsertAuthorization\|authorization:", re.I), "Registro de autorizacion (pendiente)", False),
    (re.compile(r"AuthoritationDomain\|Authorize Response", re.I), "Autorizacion procesada con el banco emisor", False),
    (re.compile(r"InsertNotification\|response:", re.I), "Registro de notificacion interna", False),
    (re.compile(r"UpdateAuthorization\|response:", re.I), "Actualizacion de estado de autorizacion", False),
    (re.compile(r"SendToApiNotification\|urlIpn:", re.I), "Envio de IPN al comercio", False),
    (re.compile(r"ConfirmationDomain\|Confirmation request:", re.I), "Confirmacion de la operacion", False),
    (re.compile(r"Background\s*(?:Job)?\s*Autoreversa\|Begin", re.I), "Programacion de job de Autoreversa", False),
    (re.compile(r"CancelAutoreversa\|Begin", re.I), "Evaluacion de Autoreversa (verifica confirmacion del comercio)", False),
    (re.compile(r"Autoreversa\|(?:Cancel request|Process\|End response)", re.I), "Ejecucion de reverso automatico (Autoreversa)", False),
    (re.compile(r"Autoreversa\|Update\|Begin request", re.I), "Actualizacion de estado final de la orden", False),
    (re.compile(r"\bCancel\|.*response", re.I), "Cancelacion / anulacion de la transaccion", False),
]

# Lista de perfiles conocidos a intentar, en orden, antes de caer al modo
# generico. Agrega aqui el de una API nueva cuando la valides contra logs
# reales de esa API (mismo formato: patron sobre el mensaje, etiqueta
# legible, colapsable).
ALL_KNOWN_MILESTONES = [PUSH_MILESTONES]

# Patrones genericos: fallback API-agnostico cuando ningun perfil conocido
# de ALL_KNOWN_MILESTONES hace match en las lineas de un TransactionId.
GENERIC_MILESTONES = [
    (re.compile(r"\bBegin\b.*[Rr]equest", re.I), "Inicio de solicitud", False),
    (re.compile(r"\bEnd\b.*response", re.I), "Fin de respuesta", False),
    (re.compile(r"\b(?:Error|Exception)\b", re.I), "Error / excepcion", False),
    (re.compile(r'"statuscode"\s*:\s*[45]\d\d', re.I), "Respuesta HTTP con error (4xx/5xx)", False),
    (re.compile(r"\b(?:Notification|Notificaci[oó]n|IPN)\b", re.I), "Notificacion", False),
    (re.compile(r"\bConfirmation\b", re.I), "Confirmacion", False),
    (re.compile(r"\b(?:Cancel|Reverso|Autoreversa)\b", re.I), "Cancelacion / reverso", False),
]


def classify(rest):
    low = rest.lower()
    if re.search(r"urlipn:\s*$", rest.rstrip(), re.I):
        return "ALERTA", "urlIpn vacio -> el comercio no fue notificado por IPN"
    if "reversad" in low:
        return "REVERSADO", snippet(rest)
    if "rechazad" in low or "denegad" in low or "declined" in low:
        return "RECHAZADO", snippet(rest)
    if "aprobado" in low or "exitosa" in low or "successful" in low:
        return "APROBADO", snippet(rest)
    # "message" explicito de exito (ej. {"code":"200","message":"OK",...} de un
    # dispatcher/notificador) no debe marcarse como rechazo aunque el "code"
    # no sea "00": ese "code" es un status HTTP-like (200), no un codigo de
    # autorizacion bancaria.
    msg_ok = re.search(r'"message"\s*:\s*"\s*(ok|exitosa|successful|aprobad[oa]?)\b', low)
    if re.search(r'"statuscode"\s*:\s*[45]\d\d', low) or "exception" in low or re.search(r'\berror\b', low):
        return "ALERTA", snippet(rest)
    # Solo tratamos "code" como codigo de autorizacion bancaria (formato ISO
    # 8583, 2 digitos: "00"=aprobado, "05"/"51"/etc=rechazo) cuando tiene
    # EXACTAMENTE 2 digitos, para no confundirlo con codigos HTTP-like de 3
    # digitos (200/400/500) que usan otros endpoints (ej. dispatcher/notificador).
    m = re.search(r'"code"\s*:\s*"?(\d{2})"?(?!\d)', low)
    if m and not msg_ok and m.group(1) != "00":
        return "RECHAZADO", snippet(rest)
    return "OK", snippet(rest)


def snippet(rest, n=280):
    return rest if len(rest) <= n else rest[:n] + "..."


def generic_component_label(msg):
    comp = msg.split("|", 1)[0].strip()
    return comp if comp else "Evento"


def _scan_with_patterns(lines_sorted, txid, patterns):
    """Escanea SOLO con los patrones de un perfil conocido (ALL_KNOWN_MILESTONES):
    lineas que no matchean ningun patron del perfil se ignoran (a diferencia
    del modo generico, que no descarta nada). Si el perfil no aplica a este
    log, devuelve una lista vacia y build_milestones() cae al modo generico."""
    rows = []
    prefix = txid + "|"
    for server, lineno, text in lines_sorted:
        m = TS_RE.match(text)
        if not m:
            continue
        ts, _thread, _level, rest = m.groups()
        msg = rest[len(prefix):] if rest.startswith(prefix) else rest
        hit = None
        for pattern, label, collapsible in patterns:
            if pattern.search(msg):
                hit = (label, collapsible)
                break
        if not hit:
            continue
        label, collapsible = hit
        estado, resultado = classify(msg)
        row_txid = extract_row_txid(msg, txid)
        if collapsible and rows and rows[-1]["label"] == label:
            rows[-1]["count"] += 1
            rows[-1]["ts_end"] = ts
            rows[-1]["estado"] = estado
            rows[-1]["resultado"] = resultado
            rows[-1]["evidencia"] = f"{rows[-1]['evidencia'].split(' ... ')[0]} ... {server} L{lineno}"
        else:
            rows.append({"label": label, "ts_start": ts, "ts_end": ts, "count": 1,
                         "estado": estado, "resultado": resultado, "row_txid": row_txid,
                         "evidencia": f"{server} L{lineno}"})
    return rows


def generic_milestones(lines_sorted, txid):
    """Modo generico (cualquier API con formato NLog): agrupa por el
    componente que emitio cada linea, colapsando repeticiones consecutivas.
    No requiere conocer el dominio/negocio de esa API."""
    rows = []
    prefix = txid + "|"
    for server, lineno, text in lines_sorted:
        m = TS_RE.match(text)
        if not m:
            continue
        ts, _thread, _level, rest = m.groups()
        msg = rest[len(prefix):] if rest.startswith(prefix) else rest
        if msg.lower().startswith("header"):
            continue
        label = None
        for pattern, generic_label, _ in GENERIC_MILESTONES:
            if pattern.search(msg):
                label = generic_label
                break
        if label is None:
            label = generic_component_label(msg)
        estado, resultado = classify(msg)
        row_txid = extract_row_txid(msg, txid)
        if rows and rows[-1]["label"] == label:
            rows[-1]["count"] += 1
            rows[-1]["ts_end"] = ts
            rows[-1]["estado"] = estado
            rows[-1]["resultado"] = resultado
            rows[-1]["evidencia"] = f"{rows[-1]['evidencia'].split(' ... ')[0]} ... {server} L{lineno}"
        else:
            rows.append({"label": label, "ts_start": ts, "ts_end": ts, "count": 1,
                         "estado": estado, "resultado": resultado, "row_txid": row_txid,
                         "evidencia": f"{server} L{lineno}"})
    return rows


def build_milestones(lines_sorted, txid):
    """Intenta cada perfil conocido (ALL_KNOWN_MILESTONES); si ninguno
    reconoce nada en estas lineas (log de una API distinta, no mapeada), cae
    al modo generico por componente, para que ninguna API se quede sin
    trazabilidad util."""
    for patterns in ALL_KNOWN_MILESTONES:
        rows = _scan_with_patterns(lines_sorted, txid, patterns)
        if rows:
            return rows
    return generic_milestones(lines_sorted, txid)


# --------------------------------------------------------------------------
# 5) Salida: Excel
# --------------------------------------------------------------------------
def build_excel(results, out_path, groups_desc, files_scanned, conn_desc):
    wb = Workbook()

    thin = Side(style="thin", color="BFBFBF")
    BORDER = Border(left=thin, right=thin, top=thin, bottom=thin)
    WRAP = Alignment(wrap_text=True, vertical="top")
    CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)

    def estado_key(estado):
        """Mapea un estado textual (OK/APROBADO/ALERTA/REVERSADO/RECHAZADO/ERROR)
        a la clave de severidad de 5 niveles (good/warning/serious/critical/error)
        que usa la paleta de marca compartida (izipay_excel)."""
        if estado == "ERROR":
            return "error"
        return STATUS_MAP.get(estado, ("warning", estado))[0]

    def info_band(ws, row, value, merge_to_col=7):
        """Banda informativa (titulo de bloque por transaccion, o del cuadro-indice):
        mismo estilo que un encabezado de tabla (fondo brand-soft, texto brand)."""
        ws.merge_cells(start_row=row, start_column=1, end_row=row, end_column=merge_to_col)
        cell = ws.cell(row=row, column=1, value=value)
        cell.font = header_font()
        cell.fill = header_fill()
        cell.alignment = WRAP
        return cell

    ws = wb.active; ws.title = "Resumen"
    r = write_title_block(
        ws, "Busqueda de Transaccion(es) via SFTP",
        f"Criterios: {groups_desc}  |  Archivos descargados/escaneados: {files_scanned}  |  {conn_desc}",
        merge_cols=7,
    )
    ws.row_dimensions[1].height = 26
    ws.row_dimensions[2].height = 22
    r += 1  # fila en blanco (titulo=1, subtitulo=2, blanco=3, headers desde 4)

    headers = ["Paso", "Hora", "Evento", "Resultado", "Estado", "TransactionId (linea)", "Evidencia (ruta remota/linea)"]
    if not results:
        info_band(ws, r, "No se encontraron transacciones que cumplan los criterios dados.")
        ws.cell(row=r, column=1).font = status_font("critical")
        wb.save(out_path)
        return

    freeze_row = r + 1
    if len(results) > 1:
        # Cuadro-indice con el resultado final de cada transaccion (solo si hay
        # mas de una), para identificar rapido cuales tuvieron error y saltar
        # directo a su detalle mas abajo en esta misma hoja.
        block_starts = {}
        row_cursor = r + (2 + len(results) + 1)  # titulo + encabezado + N filas + fila en blanco
        for item in results:
            block_starts[item["txid"]] = row_cursor
            n_effective = len(item["rows"]) if item["rows"] else 1
            row_cursor += n_effective + 3  # titulo + encabezado + N hitos (o msg) + fila en blanco

        info_band(ws, r, f"Resultado final por transaccion ({len(results)} encontradas) - click en el TransactionId para ir al detalle")
        ws.row_dimensions[r].height = 20
        r += 1

        idx_headers = ["TransactionId", "Comercio", "Orden", "Monto", "Resultado Final"]
        r = style_header_row(ws, r, idx_headers)
        freeze_row = r

        for item in results:
            ctx = item["ctx"]; txid = item["txid"]
            worst = transaction_final_status(item["rows"], item["lines"])
            estado_final = WORST_TO_ESTADO[worst]
            vals = [txid, ctx["merchant"] or "N/D", ctx["order"] or "N/D", ctx["amount"] or "N/D", estado_final]
            for c_i, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c_i, value=v)
                cell.border = BORDER
                cell.alignment = CENTER
                if c_i == 5:
                    style_status_cell(cell, worst, label=estado_final)
                if c_i == 1:
                    cell.hyperlink = f"#'Resumen'!A{block_starts[txid]}"
                    cell.font = Font(underline="single", color=BRAND)
            ws.row_dimensions[r].height = 20
            r += 1
        r += 1

    for item in results:
        ctx = item["ctx"]; txid = item["txid"]; rows = item["rows"]
        info_band(ws, r, (f"TransactionId: {txid}   |   Comercio: {ctx['merchant'] or 'N/D'}   |   "
                          f"Orden: {ctx['order'] or 'N/D'}   |   Monto: {ctx['amount'] or 'N/D'}   |   "
                          f"Cod. Autorizacion: {ctx['authcode'] or 'N/D'}   |   Referencia: {ctx['ref'] or 'N/D'}"))
        ws.row_dimensions[r].height = 30
        r += 1

        r = style_header_row(ws, r, headers)

        if not rows:
            ws.merge_cells(f"A{r}:G{r}")
            ws.cell(row=r, column=1, value="(Sin hitos reconocidos en las lineas encontradas; ver hoja Detalle_Logs)").font = status_font("warning", bold=False)
            r += 2
            continue

        for i, row in enumerate(rows, start=1):
            hora = row["ts_start"] if row["ts_start"] == row["ts_end"] else f"{row['ts_start']} -> {row['ts_end']}"
            evento = row["label"] if row["count"] == 1 else f"{row['label']} (x{row['count']})"
            row_txid = row.get("row_txid", item["txid"])
            vals = [i, hora, evento, row["resultado"], row["estado"], row_txid, row["evidencia"]]
            for c_i, v in enumerate(vals, start=1):
                cell = ws.cell(row=r, column=c_i, value=v)
                cell.border = BORDER
                cell.alignment = CENTER if c_i in (1, 2, 5, 6) else WRAP
                if c_i == 5:
                    style_status_cell(cell, estado_key(v), label=v)
            ws.row_dimensions[r].height = 30
            r += 1
        r += 1

    set_column_widths(ws, [6, 20, 38, 42, 16, 24, 46])
    freeze_header(ws, row=freeze_row)

    ws2 = wb.create_sheet("Detalle_Logs")
    d_headers = ["Fecha / Hora", "TransactionId", "Ruta remota origen", "Linea", "Nivel", "Detalle del log"]
    style_header_row(ws2, 1, d_headers)

    all_rows = []
    for item in results:
        for remote_label, lineno, text in item["lines"]:
            if "Header Authorization" in text:
                continue
            m = TS_RE.match(text)
            if not m:
                continue
            ts, _thread, level, rest = m.groups()
            if len(rest) > MAX_CELL_LEN:
                rest = rest[:MAX_CELL_LEN] + " ...[TRUNCADO - ver archivo original]"
            all_rows.append([ts, item["txid"], remote_label, lineno, level, rest])
    all_rows.sort(key=lambda x: (x[0], x[1], x[2], x[3]))

    for r_i, row in enumerate(all_rows, start=2):
        for c_i, v in enumerate(row, start=1):
            cell = ws2.cell(row=r_i, column=c_i, value=v)
            cell.border = BORDER
            cell.font = Font(size=9)
            cell.alignment = CENTER if c_i in (1, 2, 4, 5) else Alignment(wrap_text=True, vertical="top")

    set_column_widths(ws2, [22, 26, 60, 9, 8, 120])
    freeze_header(ws2, row=2)
    ws2.auto_filter.ref = f"A1:F{len(all_rows)+1}"

    wb.save(out_path)


# --------------------------------------------------------------------------
# 6) Salida: HTML (mismo contenido que el Excel, en un reporte autocontenido
#    e interactivo: acordeon por transaccion, filtro/orden en la tabla de
#    detalle, modo oscuro. Paleta y reglas de color siguen la skill dataviz
#    (paleta de estado fija good/warning/serious/critical, icono+etiqueta
#    porque warning/serious no llegan a 3:1 de contraste en superficie clara).
# --------------------------------------------------------------------------
STATUS_MAP = {
    "OK":        ("good",     "OK"),
    "APROBADO":  ("good",     "Aprobado"),
    "ALERTA":    ("warning",  "Alerta"),
    "REVERSADO": ("serious",  "Reversado"),
    "RECHAZADO": ("critical", "Rechazado"),
}
STATUS_RANK = {"good": 0, "warning": 1, "serious": 2, "critical": 3, "error": 4}
KEY_LABEL_ES = {"good": "OK", "warning": "Alerta", "serious": "Reversado", "critical": "Rechazado", "error": "Error", "none": "Sin datos"}
WORST_TO_ESTADO = {"good": "OK", "warning": "ALERTA", "serious": "REVERSADO", "critical": "RECHAZADO", "error": "ERROR", "none": "SIN DATOS"}

# Reintentos de envio de la notificacion IPN externa (dispatcher): cada linea
# "Finaliza proceso envio notificacion externa..." es un intento independiente
# con su propio resultado. Si hubo al menos un intento fallido:
#   - si algun intento tuvo exito (statusCode 200) -> se recupero, queda ALERTA
#   - si TODOS los intentos fallaron -> ERROR (el comercio nunca fue notificado)
NOTIF_ATTEMPT_RE = re.compile(r"Finaliza proceso envio notificaci.n externa\.\s*La respuesta obtenida fue:\s*(\{.*\})", re.I)


def notification_attempts(lines_sorted):
    """Lista de intentos de entrega de la notificacion IPN (True=exitoso, False=fallido)."""
    attempts = []
    for _server, _lineno, text in lines_sorted:
        m = TS_RE.match(text)
        if not m:
            continue
        rest = m.group(4)
        am = NOTIF_ATTEMPT_RE.search(rest)
        if am:
            attempts.append(bool(re.search(r'"statuscode"\s*:\s*"?200"?', am.group(1), re.I)))
    return attempts


def transaction_worst(rows):
    """Peor estado (mayor severidad) entre todos los hitos de una transaccion."""
    worst = None
    for row in rows:
        key, _ = STATUS_MAP.get(row["estado"], ("warning", row["estado"]))
        if worst is None or STATUS_RANK[key] > STATUS_RANK[worst]:
            worst = key
    return worst or "none"


def transaction_final_status(rows, lines_sorted):
    """Resultado final a mostrar en el cuadro-indice / resumen: si hubo
    reintentos de notificacion IPN con alguna falla, prioriza ese resultado
    (ALERTA si se recupero, ERROR si nunca se notifico); si no, usa el peor
    estado generico (transaction_worst)."""
    attempts = notification_attempts(lines_sorted)
    if attempts and not all(attempts):
        return "warning" if any(attempts) else "error"
    return transaction_worst(rows)
STATUS_ICONS = {
    "good": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M20 6 9 17l-5-5"/></svg>',
    "warning": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M12 9v4M12 17h.01M10.3 3.9 1.8 18a2 2 0 0 0 1.7 3h17a2 2 0 0 0 1.7-3L13.7 3.9a2 2 0 0 0-3.4 0Z"/></svg>',
    "serious": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><path d="M3 12a9 9 0 1 0 3-6.7M3 4v5h5"/></svg>',
    "critical": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="m15 9-6 6M9 9l6 6"/></svg>',
    "error": '<svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor" stroke="none"><path d="M12 2 1 21h22L12 2Zm0 6.2a1.3 1.3 0 0 1 1.3 1.3v5a1.3 1.3 0 0 1-2.6 0v-5A1.3 1.3 0 0 1 12 8.2Zm0 9.4a1.45 1.45 0 1 1 0 2.9 1.45 1.45 0 0 1 0-2.9Z"/></svg>',
    "none": '<svg viewBox="0 0 24 24" width="14" height="14" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round"><circle cx="12" cy="12" r="9"/><path d="M12 8v4M12 16h.01"/></svg>',
}


def status_chip(estado):
    key, label = STATUS_MAP.get(estado, ("warning", estado))
    return (f'<span class="chip chip-{key}">{STATUS_ICONS[key]}<span>{esc(label)}</span></span>', key)




def build_html(results, out_path, groups_desc, files_scanned, conn_desc):
    rb = ReportBuilder(
        kicker="Busqueda de transacciones via SFTP",
        h1="Evidencia de transaccion(es) segun criterios de busqueda",
        dek=f"Criterios: {groups_desc}",
        meta=[f"Generado: {esc(timestamp_now())}", f"Archivos escaneados: {files_scanned}", esc(conn_desc)],
        title="Busqueda de Transaccion(es) via SFTP",
        search_box=True, expand_controls=True, theme_controls=True,
    )

    if not results:
        rb.add_html(
            '<div class="empty">'
            f'{STATUS_ICONS["warning"]}'
            '<p><b>No se encontraron transacciones</b> que cumplan los criterios dados.</p>'
            '<p>Revisa el rango de fechas/horas, la categoria de servidores o amplia los criterios.</p>'
            '</div>'
        )
    else:
        items_with_worst = [(item, transaction_final_status(item["rows"], item["lines"])) for item in results]

        if len(results) > 1:
            counts = {}
            for _item, worst in items_with_worst:
                counts[worst] = counts.get(worst, 0) + 1
            stats = [("Transacciones encontradas", str(len(results)))]
            for key in ("good", "warning", "serious", "critical", "error", "none"):
                if counts.get(key):
                    stats.append((KEY_LABEL_ES[key], str(counts[key]), None if key == "none" else key))
            rb.add_stats(stats)

            # Cuadro-indice con el resultado final de cada transaccion, ordenado
            # de peor a mejor, para identificar rapido cuales tuvieron error e
            # ir directo a su detalle (ancla #txn-<id> mas abajo).
            rb.add_section_title(f"Resultado final por transaccion ({len(results)} encontradas)")
            ordered = sorted(items_with_worst, key=lambda p: -STATUS_RANK.get(p[1], -1))
            idx_rows = []
            for item, worst in ordered:
                ctx = item["ctx"]; txid = item["txid"]
                chip_html = f'<span class="chip chip-{worst}">{STATUS_ICONS[worst]}<span>{esc(KEY_LABEL_ES[worst])}</span></span>'
                idx_rows.append([
                    f'<a class="idx-link" href="#txn-{esc(txid)}">{esc(txid)}</a>',
                    esc(ctx["merchant"] or "N/D"), esc(ctx["order"] or "N/D"),
                    esc(ctx["amount"] or "N/D"), chip_html,
                ])
            rb.add_table(["TransactionId", "Comercio", "Orden", "Monto", "Resultado Final"], idx_rows)

        rb.add_section_title("Detalle por transaccion")
        txn_parts = ['<section id="resumen">']
        for item, worst in items_with_worst:
            ctx = item["ctx"]; txid = item["txid"]; rows = item["rows"]
            summary_chip = f'<span class="chip chip-{worst}">{STATUS_ICONS[worst]}<span>{esc(KEY_LABEL_ES[worst])}</span></span>'
            txn_parts.append(f'<details class="txn" id="txn-{esc(txid)}" open data-worst="{worst}">')
            txn_parts.append(
                f'<summary><span class="txn-id">{esc(txid)}</span>{summary_chip}'
                f'<span style="color:var(--ink-mute);font-weight:400;">'
                f'Comercio {esc(ctx["merchant"] or "N/D")} &middot; Orden {esc(ctx["order"] or "N/D")}</span></summary>'
            )
            txn_parts.append(
                '<div class="txn-meta">'
                f'<span><b>Monto</b> {esc(ctx["amount"] or "N/D")}</span>'
                f'<span><b>Cod. Autorizacion</b> {esc(ctx["authcode"] or "N/D")}</span>'
                f'<span><b>Referencia</b> {esc(ctx["ref"] or "N/D")}</span>'
                '</div>'
            )
            if not rows:
                txn_parts.append('<div class="empty">Sin hitos reconocidos en las lineas encontradas; ver detalle de logs mas abajo.</div>')
            else:
                txn_parts.append('<ol class="timeline">')
                for row in rows:
                    hora = row["ts_start"] if row["ts_start"] == row["ts_end"] else f"{row['ts_start']} &rarr; {row['ts_end']}"
                    evento = row["label"] if row["count"] == 1 else f"{row['label']} (x{row['count']})"
                    chip_step, key = status_chip(row["estado"])
                    row_txid = row.get("row_txid", txid)
                    row_txid_html = (f'<div class="step-evidencia">TransactionId de la linea: {esc(row_txid)}</div>'
                                      if row_txid != txid else "")
                    txn_parts.append(
                        f'<li class="step" style="--chip-c: var(--status-{key})">'
                        f'<div class="step-head"><span class="step-evento">{esc(evento)}</span>{chip_step}'
                        f'<span class="step-hora">{esc(hora)}</span></div>'
                        f'<div class="step-resultado">{esc(row["resultado"])}</div>'
                        f'{row_txid_html}'
                        f'<div class="step-evidencia">{esc(row["evidencia"])}</div>'
                        '</li>'
                    )
                txn_parts.append('</ol>')
            txn_parts.append('</details>')
        txn_parts.append('</section>')
        rb.add_html("".join(txn_parts))

    rb.add_section_title("Detalle de logs (evidencia cruda)")
    all_rows = []
    for item in results:
        for remote_label, lineno, text in item["lines"]:
            if "Header Authorization" in text:
                continue
            m = TS_RE.match(text)
            if not m:
                continue
            ts, _thread, level, rest = m.groups()
            if len(rest) > MAX_CELL_LEN:
                rest = rest[:MAX_CELL_LEN] + " ...[TRUNCADO - ver archivo original]"
            all_rows.append([ts, item["txid"], remote_label, lineno, level, rest])
    all_rows.sort(key=lambda x: (x[0], x[1], x[2], x[3]))

    detail_rows = []
    for ts, txid, remote_label, lineno, level, rest in all_rows:
        detail_rows.append([
            f'<span class="mono">{esc(ts)}</span>', esc(txid), esc(remote_label),
            f'<span class="mono">{lineno}</span>', esc(level),
            f'<span class="detalle-txt">{esc(rest)}</span>',
        ])
    rb.add_table(
        ["Fecha/Hora", "TransactionId", "Ruta remota", "Linea", "Nivel", "Detalle"],
        detail_rows, table_id="detailTable", searchable=True, sortable=True,
        search_placeholder="Buscar TransactionId, ruta o detalle...",
    )

    with open(out_path, "w", encoding="utf-8") as fh:
        fh.write(rb.render())


# --------------------------------------------------------------------------
# main
# --------------------------------------------------------------------------
def main():
    ap = argparse.ArgumentParser(description="Busca transacciones en logs NLog remotos via SFTP y genera evidencia en Excel/HTML.")
    ap.add_argument("--input", default="input.txt", help="input.txt con criterios flexibles (OPCIONAL si se usan flags).")
    ap.add_argument("--txid", action="append", help="TransactionId exacto (repetible).")
    ap.add_argument("--comercio", action="append", help="Codigo de comercio (repetible). OBLIGATORIO.")
    ap.add_argument("--orden", action="append", help="Nro de orden (repetible).")
    ap.add_argument("--autorizacion", action="append", help="Codigo de autorizacion (repetible).")
    ap.add_argument("--monto", action="append", help="Monto, ej. 370.00 (repetible).")

    ap.add_argument("--fecha", help="Filtra archivos remotos por fecha exacta embebida en el nombre (YYYY-MM-DD). OBLIGATORIO dar --fecha o --fecha-desde/--fecha-hasta.")
    ap.add_argument("--fecha-desde", help="Filtra desde esta fecha (YYYY-MM-DD), inclusive.")
    ap.add_argument("--fecha-hasta", help="Filtra hasta esta fecha (YYYY-MM-DD), inclusive.")
    ap.add_argument("--hora-desde", help="Filtra desde esta hora, inclusive, sobre la hora real de cada linea. Formato 'HH:MM' (se aplica cada dia dentro del rango de fecha) o 'YYYY-MM-DD HH:MM' (rango continuo que puede cruzar medianoche; en ese caso deriva --fecha-desde solo si no se dio). OBLIGATORIO junto con --hora-hasta.")
    ap.add_argument("--hora-hasta", help="Filtra hasta esta hora, inclusive. Mismo formato que --hora-desde ('HH:MM' o 'YYYY-MM-DD HH:MM'); si es formato completo deriva --fecha-hasta solo si no se dio. OBLIGATORIO junto con --hora-desde.")
    ap.add_argument("--categoria", default="todos",
                     help="PUBLICO, CONTROLLER, BUSINESS, INTERNO "
                          "(coma-separado para varias) o 'todos' (default).")
    ap.add_argument("--api", help="Filtra por nombre de la carpeta API final (coma-separado), ej. "
                                   "'authorization', 'cancel,confirmation', 'notification.dispatcher1'. "
                                   "Ver nombres validos en remote_paths.txt.")
    ap.add_argument("--ruta", help="Filtra por patron glob sobre la ruta remota completa, ej. "
                                    "'*/SVPRDLW364/*' o '*/WebSiteApiController_443/*'.")
    ap.add_argument("--remote-paths", default=DEFAULT_REMOTE_PATHS, help="Archivo con la topologia de rutas remotas.")

    ap.add_argument("--logs-dir",
                     help="Modo LOCAL (sin SFTP): carpeta con los logs YA DESCARGADOS a mano (busqueda "
                          "recursiva de *.log/*.log.gz). Si se da, se omiten credenciales/conexion/descarga "
                          "por completo y se analiza directo esa carpeta. Mutuamente excluyente en la practica "
                          "con --host/--categoria/--api/--ruta/--remote-paths (se ignoran si se da --logs-dir).")

    ap.add_argument("--host", help="Host SFTP (si se omite, se pide por consola).")
    ap.add_argument("--port", help="Puerto SFTP (default 22, se pide por consola si se omite).")
    ap.add_argument("--usuario", help="Usuario SFTP (si se omite, se pide por consola).")
    ap.add_argument("--auth", choices=["password", "key"], help="Metodo de autenticacion (si se omite, se pide por consola).")
    ap.add_argument("--credenciales", default=os.path.join(SCRIPT_DIR, "crypto", "sftp_config.json"),
                     help="Archivo JSON de credenciales SFTP guardadas (password cifrada). Default: crypto/sftp_config.json de esta skill.")
    ap.add_argument("--guardar-credenciales", action="store_true",
                     help="Tras conectar (solo auth password), guarda host/usuario/password (cifrada) en --credenciales para no volver a pedirla.")
    ap.add_argument("--olvidar-credenciales", action="store_true",
                     help="Borra el archivo --credenciales antes de continuar (fuerza a pedir todo por consola de nuevo).")

    ap.add_argument("--cache-dir", default=os.path.join(SCRIPT_DIR, ".cache"), help="Carpeta local para los logs descargados (se reutiliza entre corridas).")
    ap.add_argument("--keep-cache", action="store_true", help="No hace nada especial: la cache SIEMPRE se conserva para reutilizarla; usa --limpiar-cache para borrarla.")
    ap.add_argument("--limpiar-cache", action="store_true", help="Borra la cache local ANTES de descargar (fuerza re-descarga completa).")
    ap.add_argument("--max-depth", type=int, default=3, help="Profundidad maxima al recorrer subcarpetas remotas (default 3).")
    ap.add_argument("--paralelo-livianos", type=int, default=5,
                     help="Descargas simultaneas para archivos livianos, por debajo de --umbral-pesado-mb (default 5).")
    ap.add_argument("--paralelo-pesados", type=int, default=2,
                     help="Descargas simultaneas para archivos pesados, de --umbral-pesado-mb en adelante (default 2, para no saturar la red/memoria con varios GB a la vez).")
    ap.add_argument("--umbral-pesado-mb", type=int, default=300,
                     help="Tamano en MB desde el cual un archivo se considera 'pesado' y usa el limite de --paralelo-pesados en vez de --paralelo-livianos (default 300).")

    ap.add_argument("--output-xlsx", default=None,
                     help="Nombre del Excel de salida (default 'Busqueda_Transacciones_SFTP.xlsx'). "
                          "Siempre se guarda dentro de _reports/<fecha-de-hoy>/ de esta skill.")
    ap.add_argument("--output-html", default=None,
                     help="Nombre del HTML de salida (default 'Busqueda_Transacciones_SFTP.html'). "
                          "Siempre se guarda dentro de _reports/<fecha-de-hoy>/ de esta skill.")
    ap.add_argument("--encoding", default=None,
                     help="Encoding de los logs (utf-8|cp1252|latin-1). Default 'utf-8' en modo SFTP; "
                          "en modo --logs-dir default 'cp1252' (los logs de Push-controller observados "
                          "vienen en ese encoding -- las tildes se corrompen si se leen como utf-8).")
    args = ap.parse_args()
    if args.encoding is None:
        args.encoding = "cp1252" if args.logs_dir else "utf-8"

    f = parse_input(args.input, args.encoding)
    f = add_cli_filters(f, args)
    print(f"[1/6] Criterios: {filters_desc(f)}")
    if not any(f.values()):
        sys.exit("Debes dar al menos un criterio: --txid, --comercio, --orden, --autorizacion, --monto (o input.txt).")
    if not f["merchants"]:
        sys.exit("Debes dar --comercio (obligatorio).")
    if not (f["txids"] or f["orders"]):
        print("      [AVISO] No se dio --txid ni --orden: se hara un descubrimiento global de TODAS "
              "las transacciones del comercio que cumplan el resto de criterios (api/categoria/fecha/hora). "
              "Puede haber muchos resultados y tardar mas; agrega --txid u --orden para acotar si buscas una puntual.")
    if bool(args.hora_desde) != bool(args.hora_hasta):
        sys.exit("Si das --hora-desde o --hora-hasta, debes dar los dos (formato HH:MM, o 'YYYY-MM-DD HH:MM' para un rango continuo que cruce medianoche).")
    if not (args.hora_desde and args.hora_hasta):
        print("      [AVISO] Sin --hora-desde/--hora-hasta: se buscara en el/los dia(s) completo(s), sin acotar por hora.")
    if is_full_datetime(args.hora_desde) or is_full_datetime(args.hora_hasta):
        if not (is_full_datetime(args.hora_desde) and is_full_datetime(args.hora_hasta)):
            sys.exit("Si usas formato completo 'YYYY-MM-DD HH:MM' en --hora-desde/--hora-hasta, debes darlo en ambos.")
        if not args.fecha_desde:
            args.fecha_desde = args.hora_desde[:10]
        if not args.fecha_hasta:
            args.fecha_hasta = args.hora_hasta[:10]
    if not (args.fecha or args.fecha_desde or args.fecha_hasta):
        if args.logs_dir:
            print("      [AVISO] Sin filtro de fecha: se analizaran TODOS los logs de la carpeta --logs-dir.")
        else:
            print("      [AVISO] Sin filtro de fecha: se recorreran TODAS las fechas disponibles en el servidor "
                  "remoto -- puede ser muy lento. Considera usar --fecha / --fecha-desde / --fecha-hasta.")

    args.output_xlsx = resolve_report_path(args.output_xlsx, "Busqueda_Transacciones_SFTP.xlsx")
    args.output_html = resolve_report_path(args.output_html, "Busqueda_Transacciones_SFTP.html")

    if args.fecha:
        fecha_desc = args.fecha
    elif args.fecha_desde or args.fecha_hasta:
        fecha_desc = f"{args.fecha_desde or '*'}..{args.fecha_hasta or '*'}"
    else:
        fecha_desc = "(sin filtro)"
    hora_desc = f"{args.hora_desde}-{args.hora_hasta}" if (args.hora_desde or args.hora_hasta) else "(sin filtro)"

    if args.logs_dir:
        # ------------------------------------------------------------------
        # Modo LOCAL: logs ya descargados a mano, sin SFTP (portado de la
        # skill trazabilidad-transacciones, 2026-09-11). Se salta credenciales,
        # conexion y descarga por completo.
        # ------------------------------------------------------------------
        conn_desc = f"Logs locales: {args.logs_dir} | fecha={fecha_desc} | hora={hora_desc}"
        print(f"[2-5/6] Escaneando carpeta local (sin SFTP): {args.logs_dir}")
        all_files = iter_log_files(args.logs_dir)
        if not all_files:
            sys.exit(f"No se encontraron *.log ni *.log.gz en {args.logs_dir}")
        if args.fecha or args.fecha_desde or args.fecha_hasta:
            kept = [p for p in all_files if date_in_range(os.path.basename(p), args.fecha, args.fecha_desde, args.fecha_hasta)]
        else:
            kept = all_files
        print(f"  {args.logs_dir}: {len(all_files)} archivo(s), {len(kept)} tras filtro de fecha")
        if not kept:
            sys.exit("Ningun archivo local cumple el filtro de fecha dado.")
        local_files = [(local_server_label(p, args.logs_dir), p) for p in kept]
    else:
        # ------------------------------------------------------------------
        # Modo SFTP: conecta al servidor remoto y descarga solo lo necesario.
        # ------------------------------------------------------------------
        if args.limpiar_cache and os.path.isdir(args.cache_dir):
            import shutil
            shutil.rmtree(args.cache_dir)
            print(f"      Cache local borrada: {args.cache_dir}")
        os.makedirs(args.cache_dir, exist_ok=True)

        if args.olvidar_credenciales and os.path.exists(args.credenciales):
            os.remove(args.credenciales)
            print(f"      Credenciales guardadas borradas: {args.credenciales}")

        remote_dirs = load_remote_paths(args.remote_paths, args.categoria, args.api, args.ruta)
        print(f"[2/6] Directorios remotos a recorrer (categoria={args.categoria}, api={args.api or '(todas)'}, "
              f"ruta={args.ruta or '(sin filtro)'}): {len(remote_dirs)}")

        print("[3/6] Conectando por SFTP (se pediran credenciales)...")
        client, sftp = connect_sftp(args)
        conn_desc = f"SFTP: {args.host or '(interactivo)'} | categoria={args.categoria} | fecha={fecha_desc} | hora={hora_desc}"
        try:
            print("[4/6] Listando archivos remotos y aplicando filtro de fecha...")
            candidates = []
            for cat, base in remote_dirs:
                found = walk_remote(sftp, base, max_depth=args.max_depth)
                kept = [(p, a) for p, a in found if date_in_range(os.path.basename(p), args.fecha, args.fecha_desde, args.fecha_hasta)]
                candidates.extend(kept)
                print(f"  {base}: {len(found)} archivo(s), {len(kept)} tras filtro de fecha")

            if not candidates:
                sys.exit("No se encontro ningun archivo de log remoto que cumpla el filtro de fecha/categoria dado.")
            if not (args.fecha or args.fecha_desde or args.fecha_hasta) and len(candidates) > 200:
                print(f"  [AVISO] {len(candidates)} archivos candidatos sin filtro de fecha: la descarga puede ser lenta. "
                      f"Considera usar --fecha / --fecha-desde / --fecha-hasta.")

            print(f"[5/6] Descargando {len(candidates)} archivo(s) candidato(s) a cache local ({args.cache_dir}) "
                  f"[paralelo: {args.paralelo_livianos} livianos / {args.paralelo_pesados} pesados >= {args.umbral_pesado_mb} MB]...")
            local_files = download_candidates(client, candidates, args.cache_dir,
                                               small_workers=args.paralelo_livianos,
                                               large_workers=args.paralelo_pesados,
                                               large_threshold_mb=args.umbral_pesado_mb)
        finally:
            sftp.close()
            client.close()

    print("[6/6] Buscando en los logs descargados...")
    if f["txids"]:
        txids = set(f["txids"])
        print(f"      Busqueda DIRECTA por TransactionId: {sorted(txids)}")
    else:
        groups = build_needle_groups(f)
        found = discover_txids(local_files, groups, args.encoding, args.hora_desde, args.hora_hasta)
        txids = set(found.keys())
        print(f"      TransactionId candidatos encontrados: {len(txids)}")
        if len(txids) > 15:
            print("      [AVISO] Muchas coincidencias; agrega mas criterios (orden, monto, autorizacion) para precisar.")
        if not txids:
            sys.exit("No se encontro ningun TransactionId con esos criterios en los logs descargados.")

    collected = scan_by_txid(local_files, txids, args.encoding, args.hora_desde, args.hora_hasta)

    results = []
    for txid in sorted(txids):
        lines = collected.get(txid, [])
        lines_sorted = sorted(lines, key=lambda x: (TS_RE.match(x[2]).group(1) if TS_RE.match(x[2]) else "", x[0], x[1]))
        ctx = extract_context(lines_sorted)
        rows = build_milestones(lines_sorted, txid)
        results.append({"txid": txid, "ctx": ctx, "rows": rows, "lines": lines_sorted})

    build_excel(results, args.output_xlsx, filters_desc(f), len(local_files), conn_desc)
    build_html(results, args.output_html, filters_desc(f), len(local_files), conn_desc)

    print(f"\nGenerado: {args.output_xlsx}")
    print(f"Generado: {args.output_html}")
    for item in results:
        print(f"  {item['txid']:<30} hitos={len(item['rows']):<3} lineas={len(item['lines']):<5} comercio={item['ctx']['merchant']}")


if __name__ == "__main__":
    main()
