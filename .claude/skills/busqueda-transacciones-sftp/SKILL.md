---
name: busqueda-transacciones-sftp
description: >-
  Busca una o mas transacciones en los logs NLog de la plataforma ecommerce
  que viven en un servidor REMOTO, conectandose directamente por SFTP (no
  requiere que el usuario descargue los logs a mano de antemano, a diferencia
  de la skill trazabilidad-transacciones que asume una carpeta local). Pide
  las credenciales SFTP de forma interactiva por defecto (host, puerto,
  usuario y password o llave privada); opcionalmente se pueden guardar
  cifradas (AES-256-GCM via un wrapper .NET local) con --guardar-credenciales
  para no volver a pedirlas. Requiere SIEMPRE Codigo de Comercio, un rango de
  fechas y un rango de horas (obligatorios). TransactionId o Nro de Orden son
  opcionales: si se dan, acotan la busqueda a una transaccion puntual (mas
  rapido); si no se dan, se hace un descubrimiento global de TODAS las
  transacciones del comercio que cumplan el resto de criterios (api/categoria/
  fecha/hora), util para listar toda la actividad de un comercio en un rango
  sin conocer transacciones puntuales de antemano. Recorre una topologia
  conocida de ~40 servidores/carpetas remotas (publicos, controller/business,
  internos; ver remote_paths.txt) filtrando por esos rangos y opcionalmente
  por categoria de servidor y/o API, descarga solo esos archivos a una cache
  local reutilizable, y busca por los criterios dados (TransactionId exacto,
  o descubrimiento por combinacion de Nro de Orden/Codigo de
  Comercio/Autorizacion/Monto, o solo Codigo de Comercio). Entrega el reporte
  de evidencia en DOS formatos: Excel (.xlsx, hoja Resumen + hoja
  Detalle_Logs) y HTML autocontenido, con el mismo contenido. Usar cuando los
  logs NO estan disponibles localmente y hay que ir a buscarlos al servidor
  por SFTP, ya sea para investigar una transaccion puntual o para listar toda
  la actividad de un comercio en un rango. Reutilizable: solo cambian el
  txid/orden/comercio, el rango de fechas/horas y/o la categoria de
  servidores a recorrer.
---

# Busqueda de transacciones via SFTP

Skill **independiente** para investigar transacciones cuando los logs NLog
todavia no estan en una carpeta local: esta skill se conecta ella misma al
servidor remoto por **SFTP**, descarga solo lo necesario, busca por los
criterios dados, y entrega evidencia en **Excel y HTML**.

> Si los logs YA estan descargados en una carpeta local, usa mejor la skill
> **`trazabilidad-transacciones`** (mas rica: perfiles de hitos validados,
> cruce entre multiples APIs ya presentes en disco). Esta skill
> (`busqueda-transacciones-sftp`) es para cuando hay que **ir a buscarlos**
> al servidor primero.

## Cuando usar esta skill

Frases tipicas: "conectate por SFTP y busca esta transaccion", "necesito
bajar los logs del servidor para esta transaccion", "no tengo los logs
localmente, hay que sacarlos del servidor", "busca en el servidor de
autorizacion/cancel/confirmation esta orden".

## Topologia remota conocida (`remote_paths.txt`)

El servidor SFTP expone la carpeta base `/PuntoWeb_2/<SERVIDOR>/<Sitio_Puerto>/<api>/`
para ~40 combinaciones de servidor/sitio, agrupadas en 4 categorias (ver
comentarios al inicio de `remote_paths.txt`). Cada categoria ya unifica los
servidores "originales" y los "nuevos" (no hay distincion `_NUEVO` separada):

| Categoria | Servidores | Contenido |
|---|---|---|
| `PUBLICO` | SVPRDLW364, SVPRDMW363, SVPRDLW1021, SVPRDMW1020 | WebSiteApi\_443/444 (authorization, cancel, confirmation, orderinfo, qrnotification, security, notification.dispatcher1, capture) + WebSiteForm\_443/444/api |
| `CONTROLLER` | SVPRDMW365, SVPRDLW366, SVPRDMW946, SVPRDLW947 | WebSiteApiController\_443/444 (los servidores 946/947 suman push/qr/yape) |
| `BUSINESS` | SVPRDMW365, SVPRDLW366, SVPRDMW946, SVPRDLW947 | WebSiteApiBusiness\_443/444 (los servidores 946/947 suman authentication/push/qr/yape) |
| `INTERNO` | SVPRDLW368, SVPRDMW367, SVPRDLW653, SVPRDMW652 | WebSiteInterno\_443/444 (authorization, cancel, confirmation, accountvalidate, autoreverse) |

El script **no tiene estas rutas hardcodeadas en el codigo**: las lee de
`remote_paths.txt` (formato `CATEGORIA<TAB>ruta_remota`). Si la topologia
cambia (nuevo servidor, nueva API), edita ese archivo — no el script.

`--categoria` filtra por una o varias de estas etiquetas (coma-separado) o
`todos` (default) para recorrerlas todas.

> **Supuesto a validar contra el servidor real:** se asume que cada carpeta
> `.../<api>/` contiene directamente los `*.log`/`*.log.gz` del dia (mismo
> patron de nombre con fecha embebida que las demas skills, ej.
> `nlog-Authorization-2026-07-14.log.gz`). El recorrido remoto es
> **recursivo** (`--max-depth`, default 3) por si hay subcarpetas adicionales
> (ej. por fecha). Si el patron de nombre real difiere, ajustar `RE_FILE_DATE`
> en el script.

## Descargas en paralelo (concurrencia adaptativa por peso)

Algunos logs reales pesan **varios GB** (se han visto casos de 5+ GB en un
solo archivo del dia). SFTP no permite "grep remoto": para buscar en un
archivo hay que descargarlo entero primero. Para que eso no sea excesivamente
lento, `download_candidates()` descarga **en paralelo**, con dos límites de
concurrencia independientes segun el peso de cada archivo (basado en el
tamaño real reportado por el servidor, no en el nombre):

- **Livianos** (`< --umbral-pesado-mb`, default 300 MB): hasta
  `--paralelo-livianos` descargas simultaneas (default **5**) — varios a la
  vez, porque pesan poco.
- **Pesados** (`>= --umbral-pesado-mb`): hasta `--paralelo-pesados` descargas
  simultaneas (default **2**) — pocos a la vez, para no saturar la red o la
  memoria con varios GB en paralelo.

Ambos grupos se descargan **al mismo tiempo** (no es una fase y luego la
otra): mientras 2 archivos pesados avanzan, los livianos siguen bajando en
paralelo con su propio limite. Cada hilo abre su propio canal SFTP sobre la
misma conexion ya autenticada (`paramiko.SFTPClient.from_transport()`), ya
que un solo `SFTPClient` no es seguro para usar desde varios hilos a la vez.

**Progreso informado**: por cada archivo se imprime su inicio, avances cada
~10% (`[ 50%] archivo (2.5 MB/5.0 MB)`) y su finalizacion (`[3/8] listo:
...`), para que una descarga larga de varios GB no parezca colgada. La
concurrencia (impresion desde varios hilos) esta protegida con un lock para
que las lineas no se entremezclen en la consola.

Ejemplo para forzar mas concurrencia en livianos y menos umbral de "pesado":

```bash
--paralelo-livianos 8 --paralelo-pesados 1 --umbral-pesado-mb 200
```

> Limitacion que sigue existiendo (no resuelta por el paralelismo): el total
> de bytes transferidos por la red es el mismo (no hay filtrado remoto), y la
> cache local (`--cache-dir`) sigue acumulando el peso completo de cada
> archivo descargado. Si el disco se llena, usar `--limpiar-cache` o revisar
> `--cache-dir` manualmente.

## Credenciales SFTP (siempre interactivas)

Por defecto el script **pide por consola** en cada ejecucion: host, puerto
(default 22), usuario, y metodo de autenticacion (`password` con entrada
oculta via `getpass`, o `key` con ruta a llave privada + passphrase). Nada se
guarda en disco a menos que uses `--guardar-credenciales` (ver mas abajo).
Opcionalmente `--host`/`--port`/`--usuario`/`--auth` pre-llenan esos campos
sin exponer la password (que siempre se pide oculta si no hay credenciales
guardadas).

> Nota de seguridad: la conexion usa `paramiko.AutoAddPolicy()`, que **no
> valida la huella (fingerprint)** del servidor contra un `known_hosts`
> previo — aceptable en una red interna de confianza. Si se requiere
> verificacion estricta, cambiar a `RejectPolicy()` con `known_hosts` cargado
> de antemano (ver `connect_sftp()` en el script).

### Credenciales guardadas (cifradas), opcional

Para no volver a escribir usuario/password en cada corrida, se puede guardar
host/usuario/password (auth `password`, no aplica a `key`) en un archivo JSON
local, con la password **cifrada AES-256-GCM** (key derivada por SHA-256, tag
de 128 bits) usando la libreria interna `SecurityUtils` (paquete NuGet en
`crypto/SecurityUtils.1.5.0.nupkg`), invocada desde Python via un pequeno
wrapper .NET (`crypto/CryptoWrapper`).

- **Guardar (opcion A, recomendada para configurar todo de una vez)**: correr
  `python crypto/configurar_credenciales.py` una sola vez. Pide host/puerto/
  usuario/password por consola (password siempre oculta) y escribe
  directamente `host`, `port`, `usuario`, `auth` y `password_enc` en
  `crypto/sftp_config.json`, reusando la `crypto_key` que ya exista en ese
  archivo (si el archivo no existe, crealo primero con
  `{"crypto_key": "<tu-key-fija>"}`). Despues de esto, **ninguna corrida de
  `buscar_transacciones_sftp.py` necesita `--guardar-credenciales` ni pide
  nada por consola**: lee host/usuario/password directo del archivo.
- **Guardar (opcion B, al vuelo)**: agrega `--guardar-credenciales` a
  cualquier corrida de `buscar_transacciones_sftp.py` con `--auth password`.
  Tras conectar exitosamente, escribe (o actualiza) `--credenciales` (default
  `crypto/sftp_config.json`) con `host`, `port`, `usuario`, `auth`,
  `crypto_key` (la key de cifrado, generada al azar si no habia una) y
  `password_enc`. Solo hace falta una vez; corridas siguientes ya no
  necesitan el flag.
- **Usar lo guardado**: en corridas siguientes, si ese archivo existe, el
  script **no vuelve a pedir nada por consola** (ni host/usuario/password):
  descifra la password automaticamente y conecta directo.
- **Olvidar**: `--olvidar-credenciales` borra el archivo antes de continuar
  (la siguiente corrida vuelve a pedir todo por consola, o guarda de nuevo si
  se combina con `--guardar-credenciales` en la misma corrida).
- **Compilar el wrapper una vez** (requiere SDK de .NET 10 instalado):
  ```bash
  dotnet build "<ruta-skill>/crypto/CryptoWrapper"
  ```

### Invocacion directa por prompt (una vez configuradas las credenciales)

Con `crypto/sftp_config.json` completo (host, port, usuario, auth, crypto_key
y password_enc), el script queda **100% no interactivo**: no llama a
`input()` ni `getpass` en ningun punto de la conexion. Esto significa que
**Claude puede ejecutar el script directamente** (via su herramienta de
terminal) a partir de un pedido en lenguaje natural, sin que el usuario tenga
que abrir su propia terminal ni escribir el comando -- a diferencia del flujo
sin credenciales guardadas, donde el prompt oculto de password obliga a que
el usuario corra el comando el mismo (ver seccion de Credenciales SFTP mas
arriba, y usar `!` delante del comando si se corre dentro de Claude Code).

Ejemplo de pedido que ya alcanza para correr una busqueda completa:

> "Busca el txid 5000573822824085769115307690505601604457, comercio 4035823,
> en la API cancel de la categoria PUBLICO, hoy de 00:00 a 19:00"

Claude traduce esto al comando (`--txid`, `--comercio`, `--api cancel`,
`--categoria PUBLICO`, `--fecha`, `--hora-desde`/`--hora-hasta`) y lo corre
el mismo, sin pedir nada mas. Si `sftp_config.json` NO tiene la password
guardada, Claude no debe intentar correrlo el mismo (el prompt de password
quedaria colgado en una ejecucion no interactiva): en ese caso hay que pedirle
al usuario que lo corra el mismo o que configure las credenciales primero
(`crypto/configurar_credenciales.py`).

> ⚠️ **Advertencia de seguridad (importante):** la `crypto_key` que descifra
> la password se guarda **en el mismo archivo** que la password cifrada. Esto
> es **ofuscacion, no control de acceso real**: protege de un vistazo casual
> del archivo (un `cat` accidental, un pantallazo, un grep), pero **cualquiera
> con acceso de lectura a ese archivo puede descifrar la password** (el
> wrapper hace exactamente eso con las 2 lineas del archivo). No subas este
> archivo a ningun repositorio ni lo compartas; trata su carpeta con los
> mismos cuidados que le darias a la password en texto plano. Si se necesita
> control de acceso real, la alternativa es un almacen de secretos del
> sistema operativo (ej. Windows Credential Manager via `keyring`), donde la
> key de descifrado nunca queda accesible desde el propio archivo.

## Insumos

1. **Codigo de Comercio — OBLIGATORIO siempre**, via `input.txt` y/o flags
   (mismo esquema y sinonimos que `trazabilidad-transacciones`):
   - `comercio` / `--comercio` — Codigo de Comercio (siempre obligatorio).
   - `txid` / `--txid` — TransactionId exacto (el mas preciso; **opcional**).
   - `orden` / `--orden` — Nro de Orden (**opcional**).

   Los siguientes tambien son **opcionales**, solo para confirmar o acotar la
   transaccion encontrada:
   - `autorizacion` / `--autorizacion` — Codigo de Autorizacion.
   - `monto` / `--monto` — Monto (se prueban variantes `370`, `370.00`, `370,00`).

   Si se da `txid`, la busqueda es directa y exacta. Si se da `orden` (con o
   sin `autorizacion`/`monto`), se descubre por combinacion de esos criterios.
   Si **no** se da ni `txid` ni `orden`, se hace un **descubrimiento global**:
   todas las transacciones del comercio (combinado con `--api`/`--categoria`/
   `--ruta` y el rango de fecha/hora) que cumplan el resto de criterios dados.
   Sin txid/orden puede haber muchos resultados y tardar mas — usalos cuando
   los tengas para acotar mas rapido.

   Plantilla: **`input.ejemplo.txt`** (cópiala como `input.txt`). Si falta
   comercio, el script se detiene con un error explicito.

2. **Rango de fechas — OBLIGATORIO**:
   - `--fecha 2026-07-14` (un dia exacto), o
   - `--fecha-desde 2026-07-10 --fecha-hasta 2026-07-14` (rango).
   - Si se omiten ambos, el script se detiene con un error explicito (para
     no recorrer/descargar todo el historico por accidente).

3. **Rango de horas — OBLIGATORIO** (`--hora-desde 09:00 --hora-hasta 09:30`,
   formato `HH:MM`): filtra por la hora real de cada linea de log (no el
   nombre del archivo), asi que acota la busqueda incluso dentro de un mismo
   dia con mucho volumen. Si se omite cualquiera de los dos, el script se
   detiene con un error explicito.

4. **Donde buscar** (todo opcional; sin nada de esto se recorren las ~40
   rutas de `remote_paths.txt`). Se pueden combinar (AND) para acotar:
   - `--categoria` — grupo de servidores: `PUBLICO`, `CONTROLLER`, `BUSINESS`,
     `INTERNO` (coma-separado) o `todos` (default). Ej. `--categoria PUBLICO`
     o `--categoria CONTROLLER,BUSINESS`.
   - `--api` — nombre exacto de la carpeta/API final (coma-separado), tal
     como aparece en `remote_paths.txt`: `authorization`, `cancel`,
     `confirmation`, `orderinfo`, `qrnotification`, `security`,
     `notification.dispatcher1`, `notification.dispatcher2`,
     `notification.producer`, `capture`, `payments`, `push`, `qr`, `yape`,
     `authentication`, `accountvalidate`, `autoreverse`. Ej.
     `--api authorization` o `--api cancel,confirmation`.
   - `--ruta` — patron glob sobre la ruta remota completa, para apuntar a un
     servidor o sitio puntual sin depender de categoria/api. Ej.
     `--ruta "*/SVPRDLW364/*"` (solo ese servidor) o
     `--ruta "*/WebSiteApiController_443/*"` (solo ese sitio/puerto).

   Ejemplo apuntando a un servidor y API especificos:
   `--categoria PUBLICO --api authorization`.

## Como ejecutar

```bash
python "<ruta-skill>/buscar_transacciones_sftp.py" --input input.txt \
  --fecha 2026-07-14 --hora-desde 09:00 --hora-hasta 09:30 \
  --output-xlsx Resultado.xlsx --output-html Resultado.html
```

Por TransactionId exacto + comercio, acotando a los servidores publicos:

```bash
python "<ruta-skill>/buscar_transacciones_sftp.py" \
  --txid LP260714-195608076-KF5JXR --comercio 4080783 --fecha 2026-07-14 \
  --hora-desde 09:00 --hora-hasta 09:30 --categoria PUBLICO
```

Por Nro de Orden, confirmando con comercio + autorizacion + monto, en un rango de fechas:

```bash
python "<ruta-skill>/buscar_transacciones_sftp.py" \
  --orden 195608076 --comercio 4080783 --autorizacion 660568 --monto 370.00 \
  --fecha-desde 2026-07-10 --fecha-hasta 2026-07-14 \
  --hora-desde 00:00 --hora-hasta 23:59
```

`<ruta-skill>` es el directorio de esta skill (donde esta este SKILL.md y el
script). Al ejecutar, el script pedira por consola: `Host SFTP`, `Puerto SFTP`,
`Usuario SFTP`, `Autenticacion (password/key)` y la password (oculta) o la
ruta de la llave privada.

Parametros principales:
- `--input` : ruta del `input.txt` (default `input.txt`).
- `--txid` / `--comercio` / `--orden` / `--autorizacion` / `--monto` : criterios sueltos (repetibles).
- `--fecha` / `--fecha-desde` / `--fecha-hasta` : filtro de fecha sobre el nombre del archivo remoto (OBLIGATORIO dar uno).
- `--hora-desde` / `--hora-hasta` : filtro de hora (`HH:MM`) sobre la hora real de cada linea de log (OBLIGATORIOS ambos).
- `--categoria` : `PUBLICO`, `CONTROLLER`, `BUSINESS`, `INTERNO` (coma-separado) o `todos` (default).
- `--api` : nombre de carpeta API final (coma-separado), ej. `authorization`, `cancel,confirmation`.
- `--ruta` : patron glob sobre la ruta remota completa, ej. `*/SVPRDLW364/*`.
- `--remote-paths` : archivo de topologia remota (default `remote_paths.txt` de esta skill).
- `--host` / `--port` / `--usuario` / `--auth` : datos de conexion opcionales (la password/passphrase SIEMPRE se pide oculta, nunca por flag).
- `--credenciales` : archivo JSON de credenciales guardadas (default `crypto/sftp_config.json` de esta skill).
- `--guardar-credenciales` : tras conectar, guarda host/usuario/password (cifrada) en `--credenciales`.
- `--olvidar-credenciales` : borra `--credenciales` antes de continuar.
- `--cache-dir` : carpeta local de cache de descargas (default `.cache` dentro de la skill; se **reutiliza** entre corridas para no re-descargar).
- `--limpiar-cache` : borra la cache local antes de empezar (fuerza re-descarga completa).
- `--max-depth` : profundidad de recorrido recursivo remoto (default 3).
- `--paralelo-livianos` : descargas simultaneas para archivos livianos (default 5).
- `--paralelo-pesados` : descargas simultaneas para archivos pesados (default 2).
- `--umbral-pesado-mb` : tamano en MB desde el cual un archivo se considera "pesado" (default 300).
- `--output-xlsx` / `--output-html` : nombre de archivo de salida (default `Busqueda_Transacciones_SFTP.xlsx`/`.html`). **Siempre** se guardan dentro de `_reports/<fecha-de-hoy>/` de esta skill (ej. `_reports/2026-09-04/Resultado.xlsx`), sin importar la ruta que se pase -- solo se usa el nombre de archivo.
- `--encoding` : encoding de los logs (default `utf-8`; probar `cp1252`/`latin-1` si los acentos salen corruptos).

## Que hace, paso a paso

1. Lee criterios (`input.txt` + flags).
2. Lee `remote_paths.txt`, filtra por `--categoria`.
3. Pide credenciales SFTP por consola y conecta.
4. Recorre recursivamente cada carpeta remota (hasta `--max-depth`), filtrando
   los `*.log`/`*.log.gz` encontrados por el rango de fecha dado.
5. Descarga a `--cache-dir` solo los archivos candidatos, **en paralelo con
   concurrencia adaptativa por peso** (reutiliza los que ya esten cacheados
   con el mismo tamano — util para refinar criterios sin volver a bajar
   todo). Ver "Descargas en paralelo" mas abajo.
6. Cierra la conexion SFTP (todo el resto del procesamiento es local).
7. Si se dio `txid`, busca directo; si no, **descubre** el/los TransactionId
   cuyas lineas contienen todos los demas criterios (texto plano, AND entre
   categorias, igual que `trazabilidad-transacciones`).
8. Arma la linea de tiempo de cada TransactionId con un **modo generico** de
   hitos (agrupa por componente logueado, clasifica
   OK/APROBADO/ALERTA/RECHAZADO/REVERSADO) — ver nota abajo.
9. Genera `--output-xlsx` (hojas Resumen + Detalle_Logs) y `--output-html`
   (mismo contenido, reporte autocontenido) con el mismo contenido y colores,
   guardados siempre en `_reports/<fecha-de-hoy>/` de esta skill (ej.
   `_reports/2026-09-04/Busqueda_Transacciones_SFTP.xlsx`).

## Modo generico de hitos (sin perfil de flujo validado aun)

A diferencia de `trazabilidad-transacciones` (que ya tiene un perfil
`PUSH_MILESTONES` validado contra logs reales de Push-controller/PLIN), para
esta topologia (`authorization`, `cancel`, `confirmation`, `orderinfo`,
`qrnotification`, `security`, `notification.*`, `capture`, `payments`, `push`,
`qr`, `yape`, `accountvalidate`, `autoreverse`) todavia **no hay un perfil de
negocio mapeado**, asi que siempre se usa el modo generico: agrupa lineas
consecutivas por el nombre del componente que las genero y clasifica el
resultado por palabras clave (Aprobado/Rechazado/Reversado/Error/Excepcion/
HTTP 4xx-5xx). Esto ya da evidencia completa y ordenada; si se valida un
flujo especifico contra logs reales de estas APIs, se puede agregar un perfil
dedicado (`GENERIC_MILESTONES` -> agregar lista propia, mismo patron que en
`trazabilidad-transacciones`).

## Salida: estructura del reporte (Excel y HTML, mismo contenido)

- **Resumen**: un bloque por TransactionId encontrado (Comercio, Orden,
  Monto, Cod. Autorizacion, Referencia) + su linea de tiempo (Paso, Hora,
  Evento, Resultado, Estado coloreado, Evidencia con ruta remota y numero de
  linea).
- **Detalle_Logs**: todas las lineas crudas encontradas, ordenadas
  cronologicamente, con TransactionId, ruta remota de origen, numero de
  linea, nivel y detalle completo. Se excluyen lineas `Header Authorization`
  y se truncan campos > 30,000 caracteres.

## Reutilizacion

Para una nueva busqueda **solo cambian los criterios, el rango de fechas y/o
la categoria**:

1. Da los criterios (uno o varios) por `input.txt` o flags.
2. Da el rango de fechas a explorar en el servidor (`--fecha` o
   `--fecha-desde`/`--fecha-hasta`).
3. Opcionalmente acota `--categoria` si ya sabes por donde paso la
   transaccion (mas rapido).
4. Ejecuta el script; ingresa las credenciales SFTP cuando se pidan.

## Requisitos

Python 3.8+, `paramiko` y `openpyxl`:

```bash
pip install paramiko openpyxl
```

Solo si vas a usar `--guardar-credenciales`/credenciales guardadas: SDK de
.NET 10 (para compilar `crypto/CryptoWrapper` una vez, ver seccion de
Credenciales guardadas mas arriba).

## Notas y supuestos

- El TransactionId es el dato mas preciso; usalo siempre que lo tengas.
- La busqueda sin `txid` es por texto plano cruzado (AND entre categorias)
  sobre los archivos descargados, igual semantica que
  `trazabilidad-transacciones`.
- La cache local (`--cache-dir`) se conserva entre corridas a proposito: si
  vas a **refinar** los criterios de una misma investigacion (mismo rango de
  fechas), las corridas siguientes no vuelven a descargar lo ya traido. Usa
  `--limpiar-cache` solo si sospechas que los archivos remotos cambiaron
  (rotacion de logs) o si quieres liberar espacio.
- **Sin validar aun contra el servidor SFTP real**: la estructura de
  `remote_paths.txt` viene de la topologia informada por el usuario, pero el
  patron exacto de nombre de archivo dentro de cada carpeta (`RE_FILE_DATE`),
  el encoding real de estos logs, y si hay subcarpetas adicionales bajo cada
  `<api>/` (que `--max-depth` deberia cubrir) quedan por confirmar en la
  primera corrida real. Si algo no calza, ajustar `RE_FILE_DATE`,
  `--encoding` o `--max-depth` segun lo observado.
- Si el formato NLog cambia, o se quiere mapear un flujo especifico con
  etiquetas de negocio (en vez del modo generico), ajustar
  `GENERIC_MILESTONES`, `RE_MERCHANT`/`RE_ORDER`/`RE_AUTHCODE`/`RE_AMOUNT`/
  `RE_REF` y `classify()` en `buscar_transacciones_sftp.py`.
