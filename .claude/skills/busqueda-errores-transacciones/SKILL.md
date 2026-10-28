---
name: busqueda-errores-transacciones
description: >-
  Busca transacciones que hayan resultado en error segun los logs NLog de la
  plataforma ecommerce. CASO MAS COMUN: 1 dia especifico + rango de horas
  especifico (ej. hoy 14:00-15:30). DOS modos de origen: (A) SFTP remoto --
  descarga logs necesarios a cache local reutilizable; (B) local --logs-dir
  -- analiza carpeta con logs ya descargados. Criterios de error: statuscode
  != 200 O code != "00"/"RA2" O Error/exception. Fecha obligatoria. Comercio
  OPCIONAL. Rango horario OPCIONAL: --hora-desde/hasta O --ultimas-N-horas.
  Reportes Excel + HTML. Credenciales SFTP cifradas AES-256-GCM (opcional).
---

# Busqueda de errores en transacciones (SFTP o logs ya locales)

Skill **independiente** para investigar transacciones que fallaron, cruzando
logs NLog de la plataforma ecommerce, con **dos modos de origen de datos** que
comparten el mismo motor de busqueda, clasificacion y reporte:

- **Modo SFTP** (default): se conecta ella misma al servidor remoto por
  **SFTP**, descarga solo lo necesario a una cache local reutilizable, y
  analiza desde ahi. Usalo cuando los logs **no** estan descargados todavia.
- **Modo local** (`--logs-dir <carpeta>`): analiza directo una carpeta con
  logs **ya descargados a mano** (recursivo, `*.log`/`*.log.gz`), sin
  conectar a ningun servidor ni pedir credenciales. Usalo cuando ya tenes los
  logs en disco.

## Cuando usar esta skill

Frases tipicas: "busca transacciones con error de hoy", "dame los errores del
comercio 4080783 entre las 14:00 y 15:00 de ayer", "qué transacciones fallaron
en las ultimas 3 horas", "analiza estos logs locales y dame solo los errores",
"status code o mensaje de error en la API de autorizacion".

## Criterios de error (cualquiera de estos)

Una linea de log es un **error** si:

1. Status code (HTTP-like) `statuscode` ≠ 200
2. Codigo de autorizacion (ISO 8583) `code` ≠ "00" y ≠ "RA2"
3. Mensaje contiene las palabras `Error` o `exception` (case-insensitive)

Las transacciones sin ningun error en sus lineas se omiten del reporte.

## Topologia remota conocida (`remote_paths.txt`)

Igual que en la skill `busqueda-transacciones-local-sftp` (ver su SKILL.md).
~40 combinaciones de servidor/sitio agrupadas en 4 categorias: PUBLICO,
CONTROLLER, BUSINESS, INTERNO. Esta skill **no tiene copia propia**: lee directo el
`remote_paths.txt` de esa skill (`../busqueda-transacciones-local-sftp/`); si la
topologia cambia, edita ese archivo -- no el script.

## Modo LOCAL (`--logs-dir`, sin SFTP)

Si los logs **ya estan descargados** en una carpeta:

```bash
python "<ruta-skill>/buscar_errores_transacciones.py" \
  --logs-dir "C:\Users\yo\Downloads\logs_incidente" \
  --fecha 2026-09-25 \
  --output-xlsx Errores.xlsx --output-html Errores.html
```

- Busca recursivamente `*.log` y `*.log.gz` dentro de `--logs-dir` (si ambos
  existen para el mismo archivo, usa solo el `.log` para no duplicar).
- **No pide credenciales ni conecta a ningun servidor.** Se ignoran los flags
  de SFTP.
- El resto de criterios (comercio, fecha/hora) funciona **igual** que en modo
  SFTP -- comparte el mismo motor de busqueda, clasificacion y reporte.
- **Encoding por defecto distinto**: en modo local el default es `cp1252` (por
  compatibilidad con logs de algunos endpoints), a diferencia de `utf-8` en
  modo SFTP. Se puede forzar con `--encoding`.

## Credenciales SFTP (iguales al skill de busqueda de transacciones)

Por defecto se piden de forma interactiva (host, puerto, usuario y password u
llave privada). Con `--guardar-credenciales` se cifran y guardan localmente
AES-256-GCM (requiere SDK de .NET 10 instalado una sola vez). Ver la seccion
correspondiente en el SKILL.md del skill `busqueda-transacciones-local-sftp`
para detalles de seguridad y configuracion.

## Insumos

1. **Fecha — OBLIGATORIO** (a diferencia de busqueda de transacciones, donde
   es opcional). **La mayoria de los casos consulta un DIA ESPECIFICO:**
   - `--fecha 2026-09-25` (un dia exacto) — **caso mas comun**
   - `--fecha-desde 2026-09-24 --fecha-hasta 2026-09-25` (rango de fechas).
   
   Siempre da una fecha. Si necesitas multiples dias, usa `--fecha-desde/hasta`.

2. **Rango de horas — OPCIONAL**. DOS formatos a elegir:
   - `--hora-desde 14:00 --hora-hasta 15:30` (rango dentro del/los dia(s),
     formato `HH:MM`).
   - `--ultimas-N-horas 3` (ultimas 3 horas desde ahora, N es entero positivo).
   
   Si se omiten ambos, se busca el/los dia(s) completo(s) (00:00 a 23:59).
   **No se pueden combinar** `--hora-desde/hasta` con `--ultimas-N-horas` en
   la misma ejecucion.

3. **Comercio — OPCIONAL**:
   - `--comercio 4080783` filtra errores SOLO de ese comercio.
   - Si se omite, muestra errores de **TODOS** los comercios.
   
   Comercio es un valor unico; no se puede dar multiples valores (a diferencia
   de la skill de busqueda).

4. **Donde buscar** (igual al skill existente, todo opcional; solo modo SFTP):
   - `--categoria` — PUBLICO, CONTROLLER, BUSINESS, INTERNO (o todos).
   - `--api` — nombre de carpeta API final, ej. `authorization`, `cancel`.
   - `--ruta` — patron glob sobre la ruta remota completa, ej. `*/SVPRDLW364/*`.

## Como ejecutar

**CASO MAS COMUN: 1 dia especifico + rango de horas especifico:**

```bash
python "<ruta-skill>/buscar_errores_transacciones.py" \
  --fecha 2026-09-25 \
  --hora-desde 14:00 --hora-hasta 15:30 \
  --comercio 4080783 \
  --output-xlsx Errores.xlsx --output-html Errores.html
```

**Mismo dia, ultimas N horas desde ahora:**

```bash
python "<ruta-skill>/buscar_errores_transacciones.py" \
  --fecha 2026-09-25 \
  --ultimas-N-horas 3 \
  --output-xlsx Errores.xlsx --output-html Errores.html
```

**Rango de fechas (multiple dias) con rango de horas:**

```bash
python "<ruta-skill>/buscar_errores_transacciones.py" \
  --fecha-desde 2026-09-24 --fecha-hasta 2026-09-25 \
  --hora-desde 09:00 --hora-hasta 14:30 \
  --output-xlsx Errores.xlsx --output-html Errores.html
```

**Modo SFTP, acotando a categorias y API:**

```bash
python "<ruta-skill>/buscar_errores_transacciones.py" \
  --fecha 2026-09-25 \
  --categoria PUBLICO --api authorization,cancel \
  --ultimas-3-horas
```

**Modo local, logs ya descargados:**

```bash
python "<ruta-skill>/buscar_errores_transacciones.py" \
  --logs-dir "C:\ruta\a\logs" \
  --fecha 2026-09-25
```

`<ruta-skill>` es el directorio de esta skill. En modo SFTP pedira credenciales
por consola (a menos que esten guardadas cifradas en `crypto/sftp_config.json`).

## Parametros

- `--logs-dir` : modo local, carpeta con logs ya descargados (recursivo,
  `*.log`/`*.log.gz`). Si se da, ignora todos los flags de SFTP.
- `--fecha` : dia exacto (YYYY-MM-DD), ej. `2026-09-25` (OBLIGATORIO).
- `--fecha-desde` / `--fecha-hasta` : rango de fechas (ambos obligatorios si se
  usan, OBLIGATORIO al menos uno de fecha/fecha-desde; OBLIGATORIO general).
- `--hora-desde` / `--hora-hasta` : rango de horas dentro del dia, `HH:MM`
  (OPCIONAL; si das uno, das el otro tambien). NO se puede combinar con
  `--ultimas-N-horas`.
- `--ultimas-N-horas` : ultimas N horas desde ahora, ej. `--ultimas-N-horas 3`
  (OPCIONAL; entero positivo). NO se puede combinar con `--hora-desde/hasta`.
- `--comercio` : codigo de comercio unico a filtrar (OPCIONAL; si se omite,
  muestra errores de todos los comercios).
- `--categoria` : PUBLICO, CONTROLLER, BUSINESS, INTERNO (coma-separado) o
  todos (default). Solo modo SFTP.
- `--api` : nombre de carpeta API (coma-separado), ej. `authorization`,
  `cancel`. Solo modo SFTP.
- `--ruta` : patron glob sobre ruta remota, ej. `*/SVPRDLW364/*`. Solo SFTP.
- `--remote-paths` : archivo de topologia (default `remote_paths.txt` de `busqueda-transacciones-local-sftp`). SFTP.
- `--host` / `--port` / `--usuario` / `--auth` : datos de conexion opcionales.
  SFTP. Password siempre se pide oculta.
- `--credenciales` : archivo JSON de creds guardadas (default
  `crypto/sftp_config.json`). SFTP.
- `--guardar-credenciales` : tras conectar, guarda creds cifradas. SFTP.
- `--olvidar-credenciales` : borra el archivo de creds guardadas. SFTP.
- `--cache-dir` : carpeta local de cache (default `.cache` dentro de skill).
  Se reutiliza entre corridas. SFTP.
- `--cache-max-days` : borra automaticamente archivos cache con >N dias
  (default 15; usa 0 para desactivar). SFTP.
- `--limpiar-cache` : borra la cache antes de empezar. SFTP.
- `--max-depth` : profundidad recorrido recursivo remoto (default 3). SFTP.
- `--paralelo-livianos` : descargas simultaneas para archivos livianos
  (default 5). SFTP.
- `--paralelo-pesados` : descargas simultaneas para archivos pesados
  (default 2). SFTP.
- `--umbral-pesado-mb` : tamano en MB desde el cual se considera "pesado"
  (default 300). SFTP.
- `--output-xlsx` / `--output-html` : nombre de archivo salida (default
  `Errores_Transacciones.xlsx`/`.html`). Siempre se guardan en
  `_reports/<fecha-de-hoy>/` de esta skill.
- `--encoding` : encoding de logs. Default `utf-8` SFTP, `cp1252` local.

## Requisitos

Python 3.8+, `paramiko` y `openpyxl`:

```bash
pip install paramiko openpyxl
```

Solo si usas `--guardar-credenciales`: SDK de .NET 10 (para compilar
`crypto/CryptoWrapper` una sola vez).

## Notas

- Esta skill nace como complemento de `busqueda-transacciones-local-sftp`, con
  foco en detectar **ERRORES** en lugar de rastrear una transaccion conocida.
  Comercio es opcional (puede ver todo) vs. obligatorio en la otra.
- Rango de fecha es obligatorio (no busca "todo el historico" como la otra);
  siempre da `--fecha` o `--fecha-desde/hasta` para acotar el volumen.
- Validado contra los mismos logs reales de SFTP que `busqueda-transacciones-local-sftp`.
- La cache local (`--cache-dir`) se conserva entre corridas. Usa
  `--limpiar-cache` para liberar espacio o si sospechas cambios remotos.
