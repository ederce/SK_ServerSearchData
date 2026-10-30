#!/usr/bin/env python3
"""
descargar_logs_sftp.py -- SOLO descarga logs NLog del servidor SFTP a una carpeta
local, sin buscar nada. Es el PRIMER paso del flujo de las skills
busqueda-transacciones-local-sftp y busqueda-errores-transacciones: primero se
baja el dia, y recien despues (solo si el usuario lo pide) se busca con
--logs-dir sobre esa carpeta -- o se revisa a mano.

Estructura local (espejo exacto de la topologia remota, por servidor, sitio =
categoria + puerto, y API):

    <destino>/PuntoWeb_2/<SERVIDOR>/<Sitio>_<puerto>/<api>/<archivo>.log[.gz]
    ej. ...\\LOG_PW2\\PuntoWeb_2\\SVPRDLW366\\WebSiteApiBusiness_444\\qr\\nlog-Business-all-2026-10-29.log

Reutiliza la conexion, credenciales cifradas, topologia (remote_paths.txt) y
filtros de buscar_transacciones_sftp.py. Robustez ante los cortes intermitentes
del servidor: descarga a '<archivo>.part', reanuda desde el offset, reintenta y
reconecta solo; un archivo ya al dia (tamano local >= remoto) se omite. El log
del dia sigue creciendo: volver a correr el script refresca solo lo que cambio.

Uso:
    python descargar_logs_sftp.py                                   # hoy, todas las rutas
    python descargar_logs_sftp.py --categoria BUSINESS --api qr     # solo QR de BUSINESS
    python descargar_logs_sftp.py --fecha 2026-10-28 --categoria CONTROLLER
    python descargar_logs_sftp.py --ruta "*/SVPRDLW366/*"
"""
import argparse, os, sys, threading, time
from concurrent.futures import ThreadPoolExecutor, as_completed
from datetime import datetime

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, SCRIPT_DIR)
import paramiko
from buscar_transacciones_sftp import (connect_sftp, load_remote_paths, walk_remote,
                                       date_in_range, long_path, DEFAULT_REMOTE_PATHS)

# <raiz del proyecto>/LOG_PW2 (la skill vive en <raiz>/.claude/skills/<skill>/)
DEFAULT_DESTINO = os.path.abspath(os.path.join(SCRIPT_DIR, "..", "..", "..", "LOG_PW2"))
MAX_INTENTOS = 5
CANAL_TIMEOUT_S = 60
CHUNK = 1024 * 1024


class Conexion:
    """Una sola sesion SSH compartida; cada hilo abre su propio canal SFTP. Si
    el servidor corta, un unico hilo reconecta (el resto reutiliza)."""

    def __init__(self, args):
        self.args = args
        self.lock = threading.Lock()
        self.gen = 0
        self.client = None
        self._abrir()

    def _abrir(self):
        client, sftp = connect_sftp(self.args)
        sftp.close()
        self.client = client
        self.gen += 1

    def canal(self):
        with self.lock:
            tr = self.client.get_transport()
            if tr is None or not tr.is_active():
                self._abrir()
            sftp = paramiko.SFTPClient.from_transport(self.client.get_transport())
            # Sin timeout, una lectura que el servidor deja colgada bloquea para siempre y los
            # reintentos nunca se activan; con timeout salta socket.timeout y se reintenta/reconecta.
            sftp.get_channel().settimeout(CANAL_TIMEOUT_S)
            return sftp, self.gen

    def reconectar(self, gen):
        with self.lock:
            if gen != self.gen:
                return
            try:
                self.client.close()
            except Exception:
                pass
            for espera in (2, 5, 10, 20):
                try:
                    self._abrir()
                    return
                except Exception as e:
                    print(f"  [reconexion] fallo ({e}); reintento en {espera}s")
                    time.sleep(espera)
            self._abrir()

    def cerrar(self):
        try:
            self.client.close()
        except Exception:
            pass


def descargar_uno(conn, remote, attr, destino):
    """Devuelve 'bajado' | 'al dia'. Lanza excepcion si agota los reintentos."""
    local = long_path(os.path.join(destino, remote.lstrip("/").replace("/", os.sep)))
    part = local + ".part"
    os.makedirs(os.path.dirname(local), exist_ok=True)
    if os.path.exists(local) and os.path.getsize(local) >= attr.st_size:
        return "al dia"
    ultimo = None
    for intento in range(1, MAX_INTENTOS + 1):
        sftp, gen = None, None
        try:
            sftp, gen = conn.canal()
            offset = os.path.getsize(part) if os.path.exists(part) else 0
            if offset > attr.st_size:
                offset = 0
            if offset == 0:
                sftp.get(remote, part)
            else:
                with sftp.open(remote, "rb") as rf, open(part, "ab") as lf:
                    rf.seek(offset)
                    while True:
                        data = rf.read(CHUNK)
                        if not data:
                            break
                        lf.write(data)
            if os.path.getsize(part) < attr.st_size:
                raise IOError(f"descarga incompleta {os.path.getsize(part)}/{attr.st_size} bytes")
            os.replace(part, local)
            return "bajado"
        except Exception as e:
            ultimo = e
            print(f"  [reintento {intento}/{MAX_INTENTOS}] {remote}: {e}")
            if gen is not None:
                conn.reconectar(gen)
            time.sleep(min(2 * intento, 10))
        finally:
            if sftp is not None:
                try:
                    sftp.close()
                except Exception:
                    pass
    raise RuntimeError(f"no se pudo descargar {remote}: {ultimo}")


def main():
    ap = argparse.ArgumentParser(description="Descarga logs NLog del SFTP a una carpeta local (sin buscar).")
    ap.add_argument("--destino", default=DEFAULT_DESTINO,
                    help=f"Carpeta raiz local; los logs quedan en <destino>/PuntoWeb_2/... (default {DEFAULT_DESTINO}).")
    ap.add_argument("--fecha", help="Dia exacto YYYY-MM-DD (default: HOY si no se da ninguna fecha).")
    ap.add_argument("--fecha-desde", help="Rango: desde esta fecha, inclusive.")
    ap.add_argument("--fecha-hasta", help="Rango: hasta esta fecha, inclusive.")
    ap.add_argument("--categoria", default="todos", help="PUBLICO, CONTROLLER, BUSINESS, INTERNO (coma-separado) o 'todos' (default).")
    ap.add_argument("--api", help="Carpeta API final (coma-separado), ej. qr,authorization.")
    ap.add_argument("--ruta", help="Patron glob sobre la ruta remota completa, ej. '*/SVPRDLW366/*'.")
    ap.add_argument("--remote-paths", default=DEFAULT_REMOTE_PATHS)
    ap.add_argument("--max-depth", type=int, default=3)
    ap.add_argument("--paralelo", type=int, default=3, help="Descargas simultaneas (default 3).")
    ap.add_argument("--host"); ap.add_argument("--port"); ap.add_argument("--usuario")
    ap.add_argument("--auth", choices=["password", "key"])
    ap.add_argument("--credenciales", default=os.path.join(SCRIPT_DIR, "crypto", "sftp_config.json"))
    ap.add_argument("--guardar-credenciales", action="store_true")
    ap.add_argument("--olvidar-credenciales", action="store_true")
    args = ap.parse_args()

    if not (args.fecha or args.fecha_desde or args.fecha_hasta):
        args.fecha = datetime.now().strftime("%Y-%m-%d")
        print(f"[AVISO] Sin fecha: se descarga HOY ({args.fecha}).")

    if args.olvidar_credenciales and os.path.exists(args.credenciales):
        os.remove(args.credenciales)

    remote_dirs = load_remote_paths(args.remote_paths, args.categoria, args.api, args.ruta)
    print(f"[1/3] Directorios remotos: {len(remote_dirs)} (categoria={args.categoria}, api={args.api or '(todas)'}, "
          f"ruta={args.ruta or '(sin filtro)'})")

    print("[2/3] Conectando por SFTP y listando archivos...")
    conn = Conexion(args)
    try:
        sftp, _ = conn.canal()
        candidatos = []
        for _cat, base in remote_dirs:
            found = walk_remote(sftp, base, max_depth=args.max_depth)
            kept = [(p, a) for p, a in found if date_in_range(os.path.basename(p), args.fecha, args.fecha_desde, args.fecha_hasta)]
            candidatos.extend(kept)
            print(f"  {base}: {len(found)} archivo(s), {len(kept)} tras filtro de fecha")
        sftp.close()
        if not candidatos:
            sys.exit("No hay logs remotos que cumplan el filtro de fecha/categoria/api dado.")

        total_mb = sum(a.st_size for _, a in candidatos) / (1024 * 1024)
        print(f"[3/3] Descargando {len(candidatos)} archivo(s) ({total_mb:.1f} MB) a {args.destino}\\PuntoWeb_2 ...")
        resultados = {"bajado": 0, "al dia": 0}
        fallidos = []
        hecho = {"n": 0}
        lock = threading.Lock()

        def tarea(p, a):
            estado = descargar_uno(conn, p, a, args.destino)
            with lock:
                hecho["n"] += 1
                print(f"  [{hecho['n']}/{len(candidatos)}] {estado}: {p}")
            return estado

        with ThreadPoolExecutor(max_workers=max(1, args.paralelo)) as ex:
            futs = {ex.submit(tarea, p, a): p for p, a in candidatos}
            for fut in as_completed(futs):
                try:
                    resultados[fut.result()] += 1
                except Exception as e:
                    fallidos.append((futs[fut], str(e)))
    finally:
        conn.cerrar()

    print(f"\nListo: {resultados['bajado']} bajado(s), {resultados['al dia']} ya al dia, {len(fallidos)} fallido(s).")
    print(f"Carpeta: {os.path.join(args.destino, 'PuntoWeb_2')}")
    print("Nota: el log del dia sigue creciendo; vuelve a correr este script para refrescarlo antes de buscar.")
    for p, e in fallidos:
        print(f"  FALLO: {p} -> {e}")
    sys.exit(1 if fallidos else 0)


if __name__ == "__main__":
    main()
