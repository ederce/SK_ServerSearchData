#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Configura UNA SOLA VEZ host/puerto/usuario/password (cifrada) en el archivo
de credenciales (default: crypto/sftp_config.json, mismo que usa
buscar_transacciones_sftp.py). Despues de correr esto, las busquedas ya no
piden nada por consola: leen host/usuario/password directo de ese archivo.

La password SIEMPRE se pide oculta por consola (getpass); nunca se escribe
en texto plano en ningun archivo ni se pasa por argumentos de linea de
comandos.

Requiere que el archivo de credenciales ya tenga una "crypto_key" (la key
fija de cifrado). Si no existe, crea el archivo primero con:
    { "crypto_key": "<tu-key-fija>" }

Uso:
    python configurar_credenciales.py --host MC0305 --port 22 --usuario jchavez
    (host/puerto/usuario tambien se piden por consola si se omiten)
"""
import argparse, getpass, os, sys

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.dirname(SCRIPT_DIR))
import buscar_transacciones_sftp as core  # noqa: E402


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("--host")
    ap.add_argument("--port")
    ap.add_argument("--usuario")
    ap.add_argument("--credenciales", default=os.path.join(SCRIPT_DIR, "sftp_config.json"))
    args = ap.parse_args()

    creds = core.load_saved_credentials(args.credenciales) or {}
    crypto_key = creds.get("crypto_key")
    if not crypto_key:
        sys.exit(f"No hay 'crypto_key' en {args.credenciales}. "
                  f'Crea el archivo primero con: {{"crypto_key": "<tu-key-fija>"}}')

    host = args.host or creds.get("host") or input("Host SFTP: ").strip()
    port = int(args.port or creds.get("port") or input("Puerto SFTP [22]: ").strip() or 22)
    usuario = args.usuario or creds.get("usuario") or input("Usuario SFTP: ").strip()
    password = getpass.getpass("Password SFTP (oculta, no se guarda en texto plano): ")

    password_enc = core.crypto_encrypt(crypto_key, password)
    core.save_credentials(args.credenciales, host, port, usuario, "password", crypto_key, password_enc)
    print(f"Listo: {args.credenciales} actualizado (password cifrada con la crypto_key existente).")
    print("Las siguientes corridas de buscar_transacciones_sftp.py ya no pediran credenciales por consola.")


if __name__ == "__main__":
    main()
