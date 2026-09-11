---
name: busqueda-transacciones-local-sftp
description: >-
  Busca una o mas transacciones en los logs NLog de la plataforma ecommerce,
  en DOS modos de origen de datos: (A) MODO SFTP -- se conecta ella misma a un
  servidor REMOTO por SFTP (no requiere que el usuario descargue los logs a
  mano de antemano), recorre una topologia conocida de ~40 servidores/carpetas
  remotas y descarga solo lo necesario a una cache local reutilizable; (B)
  MODO LOCAL (--logs-dir) -- analiza directo una carpeta con logs YA
  DESCARGADOS a mano, sin conectar a ningun servidor ni pedir credenciales.
  Ambos modos comparten el mismo motor de busqueda/clasificacion/reporte.
  Pide las credenciales SFTP de forma interactiva por defecto (host, puerto,
  usuario y password o llave privada) solo en modo SFTP; opcionalmente se
  pueden guardar cifradas (AES-256-GCM via un wrapper .NET local) con
  --guardar-credenciales para no volver a pedirlas. Requiere Codigo de
  Comercio siempre; el rango de fechas y de horas son OPCIONALES en ambos
  modos (si se omiten, se busca en todo lo disponible -- uno o varios dias
  completos). TransactionId o Nro de Orden son opcionales: si se dan, acotan
  la busqueda a una transaccion puntual (mas rapido); si no se dan, se hace un
  descubrimiento global de TODAS las transacciones del comercio que cumplan
  el resto de criterios, util para listar toda la actividad de un comercio
  sin conocer transacciones puntuales de antemano. Tiene un perfil de hitos
  de negocio validado (PUSH_MILESTONES) para el flujo Push/PLIN; para
  cualquier otra API sin perfil conocido, cae automaticamente a un modo
  generico que agrupa por componente logueado. Entrega el reporte de
  evidencia en DOS formatos: Excel (.xlsx, hoja Resumen + hoja Detalle_Logs) y
  HTML autocontenido, con el mismo contenido. Usar cuando un comercio o el
  equipo de operaciones pide investigar/trazar una transaccion (por
  TransactionId, comercio, orden, autorizacion o monto) sin importar si los
  logs ya estan descargados localmente o hay que ir a buscarlos al servidor
  por SFTP. Reutilizable: solo cambian los criterios, el rango de
  fechas/horas, la categoria de servidores (modo SFTP) o la carpeta de logs
  (modo local).
---

# Busqueda de transacciones (SFTP o logs ya locales)

Skill **independiente** para investigar transacciones cruzando logs NLog de
la plataforma ecommerce, con **dos modos de origen de datos** que comparten
el mismo motor de busqueda, clasificacion y reporte:

- **Modo SFTP** (default): se conecta ella misma al servidor remoto por
  **SFTP**, descarga solo lo necesario a una cache local reutilizable, y
  analiza desde ahi. Usalo cuando los logs **no** estan descargados todavia.
- **Modo local** (`--logs-dir <carpeta>`): analiza directo una carpeta con
  logs **ya descargados a mano** (recursivo, `*.log`/`*.log.gz`), sin
  conectar a ningun servidor ni pedir credenciales. Usalo cuando ya tenes los
  logs en disco (ej. te los paso Infraestructura, o los bajaste vos mismo).

> Esta skill nace de fusionar la antigua skill global `trazabilidad-transacciones`
> (que solo tenia modo local) dentro de esta (que solo tenia modo SFTP),
> homologadas el 2026-09-11 para no mantener dos motores de busqueda
> duplicados. `trazabilidad-transacciones` ya no existe como skill separada.

## Cuando usar esta skill

Frases tipicas: "conectate por SFTP y busca esta transaccion", "necesito
bajar los logs del servidor para esta transaccion", "no tengo los logs
localmente, hay que sacarlos del servidor", "busca en el servidor de
autorizacion/cancel/confirmation esta orden", "traza esta transaccion con los
logs que ya tengo en esta carpeta", "quiero la trazabilidad completa de este
comercio/orden".

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

> Cada carpeta `.../<api>/` contiene directamente los `*.log`/`*.log.gz` del
> dia (patron de nombre con fecha embebida, ej.
> `nlog-ApiPublicCancel-all-2026-09-04.log` o
> `nlog-ApiNotication-Dispatcher1-PUBLIC-2026-09-10.log.gz`; `RE_FILE_DATE`
> solo busca `YYYY-MM-DD` en cualquier parte del nombre, asi que tolera
> variantes). El recorrido remoto es **recursivo** (`--max-depth`, default 3)
> porque algunos servidores tienen una subcarpeta "espejo" homonima al lado
> del archivo con una copia duplicada del mismo contenido (ver aviso de
> `long_path()` mas abajo). Validado contra el servidor real multiples veces
> desde 2026-09-04.

## Modo LOCAL (`--logs-dir`, sin SFTP)

Si los logs **ya estan descargados** en una carpeta (a mano, o porque
Infraestructura los paso), usa `--logs-dir <carpeta>` en vez de conectar por
SFTP:

```bash
python "<ruta-skill>/buscar_transacciones_sftp.py" \
  --logs-dir "C:\Users\yo\Downloads\logs_incidente" \
  --txid 174544565028023 --comercio 4078371 \
  --output-xlsx Resultado.xlsx --output-html Resultado.html
```

- Busca recursivamente `*.log` y `*.log.gz` dentro de `--logs-dir` (si existen
  ambos para el mismo archivo -- ej. copia descomprimida al lado del `.gz`
  original -- usa solo el `.log`, para no contar el contenido dos veces).
- **No pide credenciales ni conecta a ningun servidor.** Se ignoran
  `--host`/`--categoria`/`--api`/`--ruta`/`--remote-paths`/`--cache-dir` y
  demas flags especificos de SFTP.
- El resto de criterios (comercio, txid, orden, autorizacion, monto,
  fecha/hora) funciona **exactamente igual** que en modo SFTP -- comparte el
  mismo motor de busqueda, clasificacion y reporte.
- **Encoding por defecto distinto**: en modo local el default es `cp1252` (los
  logs de Push-controller observados vienen en ese encoding; las tildes se
  corrompen si se leen como utf-8), a diferencia del `utf-8` por defecto en
  modo SFTP. Se puede forzar con `--encoding` en cualquiera de los dos modos.
- Usa el perfil de hitos **`PUSH_MILESTONES`** (validado contra logs reales
  de Push-controller/PLIN) si reconoce ese flujo en las lineas encontradas;
  si no, cae al modo generico (ver mas abajo) -- funciona igual aunque la
  transaccion pase por una API sin perfil dedicado todavia.

## Descargas en paralelo (concurrencia adaptativa por peso, solo modo SFTP)

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

1. **Codigo de Comercio — OBLIGATORIO siempre**, via `input.txt` y/o flags:
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

2. **Rango de fechas — OPCIONAL**:
   - `--fecha 2026-07-14` (un dia exacto), o
   - `--fecha-desde 2026-07-10 --fecha-hasta 2026-07-14` (rango).
   - Si se omiten ambos, se busca en **todas** las fechas disponibles (todo
     el historico remoto en modo SFTP -- puede ser lento, se imprime un
     aviso -- o toda la carpeta en modo local). Da un rango cuando lo
     conozcas para acotar y acelerar la busqueda.

3. **Rango de horas — OPCIONAL** (`--hora-desde 09:00 --hora-hasta 09:30`,
   formato `HH:MM`, o `"YYYY-MM-DD HH:MM"` para un rango continuo que cruce
   medianoche): filtra por la hora real de cada linea de log (no el nombre
   del archivo), asi que acota la busqueda incluso dentro de un mismo dia con
   mucho volumen. Si se omiten ambos, se busca el/los dia(s) completo(s) sin
   acotar por hora (util cuando queres uno o varios dias enteros). Si das
   uno de los dos, tenes que dar el otro tambien.

4. **Donde buscar** (todo opcional; solo aplica en modo SFTP -- en modo
   `--logs-dir` se ignoran. Sin nada de esto se recorren las ~40 rutas de
   `remote_paths.txt`). Se pueden combinar (AND) para acotar:
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

Modo local (logs ya descargados), sin fecha/hora (busca en toda la carpeta):

```bash
python "<ruta-skill>/buscar_transacciones_sftp.py" \
  --logs-dir "C:\Users\yo\Downloads\logs_incidente" \
  --txid 174544565028023 --comercio 4078371
```

`<ruta-skill>` es el directorio de esta skill (donde esta este SKILL.md y el
script). Al ejecutar, el script pedira por consola: `Host SFTP`, `Puerto SFTP`,
`Usuario SFTP`, `Autenticacion (password/key)` y la password (oculta) o la
ruta de la llave privada.

Parametros principales:
- `--input` : ruta del `input.txt` (default `input.txt`).
- `--txid` / `--comercio` / `--orden` / `--autorizacion` / `--monto` : criterios sueltos (repetibles).
- `--logs-dir` : **modo local**, carpeta con logs ya descargados (ver seccion "Modo LOCAL" arriba). Si se da, ignora todos los flags de SFTP de abajo.
- `--fecha` / `--fecha-desde` / `--fecha-hasta` : filtro de fecha (OPCIONAL; sin esto se busca en todas las fechas disponibles).
- `--hora-desde` / `--hora-hasta` : filtro de hora, `HH:MM` o `"YYYY-MM-DD HH:MM"` (OPCIONAL; si das uno, das el otro tambien).
- `--categoria` : `PUBLICO`, `CONTROLLER`, `BUSINESS`, `INTERNO` (coma-separado) o `todos` (default). Solo modo SFTP.
- `--api` : nombre de carpeta API final (coma-separado), ej. `authorization`, `cancel,confirmation`. Solo modo SFTP.
- `--ruta` : patron glob sobre la ruta remota completa, ej. `*/SVPRDLW364/*`. Solo modo SFTP.
- `--remote-paths` : archivo de topologia remota (default `remote_paths.txt` de esta skill). Solo modo SFTP.
- `--host` / `--port` / `--usuario` / `--auth` : datos de conexion opcionales (la password/passphrase SIEMPRE se pide oculta, nunca por flag). Solo modo SFTP.
- `--credenciales` : archivo JSON de credenciales guardadas (default `crypto/sftp_config.json` de esta skill). Solo modo SFTP.
- `--guardar-credenciales` : tras conectar, guarda host/usuario/password (cifrada) en `--credenciales`. Solo modo SFTP.
- `--olvidar-credenciales` : borra `--credenciales` antes de continuar. Solo modo SFTP.
- `--cache-dir` : carpeta local de cache de descargas (default `.cache` dentro de la skill; se **reutiliza** entre corridas para no re-descargar). Solo modo SFTP.
- `--limpiar-cache` : borra la cache local antes de empezar (fuerza re-descarga completa). Solo modo SFTP.
- `--max-depth` : profundidad de recorrido recursivo remoto (default 3). Solo modo SFTP.
- `--paralelo-livianos` : descargas simultaneas para archivos livianos (default 5). Solo modo SFTP.
- `--paralelo-pesados` : descargas simultaneas para archivos pesados (default 2). Solo modo SFTP.
- `--umbral-pesado-mb` : tamano en MB desde el cual un archivo se considera "pesado" (default 300). Solo modo SFTP.
- `--output-xlsx` / `--output-html` : nombre de archivo de salida (default `Busqueda_Transacciones_SFTP.xlsx`/`.html`). **Siempre** se guardan dentro de `_reports/<fecha-de-hoy>/` de esta skill (ej. `_reports/2026-09-04/Resultado.xlsx`), sin importar la ruta que se pase -- solo se usa el nombre de archivo.
- `--encoding` : encoding de los logs. Default `utf-8` en modo SFTP, `cp1252` en modo `--logs-dir` (ver "Modo LOCAL" arriba); probar `latin-1` si los acentos salen corruptos.

## Que hace, paso a paso

**Modo SFTP** (default):
1. Lee criterios (`input.txt` + flags).
2. Lee `remote_paths.txt`, filtra por `--categoria`.
3. Pide credenciales SFTP por consola y conecta.
4. Recorre recursivamente cada carpeta remota (hasta `--max-depth`), filtrando
   los `*.log`/`*.log.gz` encontrados por el rango de fecha dado (si se dio).
5. Descarga a `--cache-dir` solo los archivos candidatos, **en paralelo con
   concurrencia adaptativa por peso** (reutiliza los que ya esten cacheados
   con el mismo tamano — util para refinar criterios sin volver a bajar
   todo). Ver "Descargas en paralelo" mas arriba.
6. Cierra la conexion SFTP (todo el resto del procesamiento es local).

**Modo LOCAL** (`--logs-dir`): salta los pasos 2-6 de arriba -- busca
recursivamente `*.log`/`*.log.gz` directo en la carpeta dada, sin conectar a
nada.

**Comun a ambos modos, de ahi en adelante:**

7. Si se dio `txid`, busca directo; si no, **descubre** el/los TransactionId
   cuyas lineas contienen todos los demas criterios (texto plano, AND entre
   categorias).
8. Arma la linea de tiempo de cada TransactionId: intenta primero el perfil
   **`PUSH_MILESTONES`** (validado para Push/PLIN); si ninguna linea matchea
   ese perfil, cae al **modo generico** (agrupa por componente logueado) --
   ver nota abajo. Clasifica cada hito en
   OK/APROBADO/ALERTA/RECHAZADO/REVERSADO/ERROR.
9. Genera `--output-xlsx` (hojas Resumen + Detalle_Logs) y `--output-html`
   (mismo contenido, reporte autocontenido) con el mismo contenido y colores,
   guardados siempre en `_reports/<fecha-de-hoy>/` de esta skill (ej.
   `_reports/2026-09-04/Busqueda_Transacciones_SFTP.xlsx`).

## Perfiles de hitos (PUSH_MILESTONES + modo generico como fallback)

`build_milestones()` intenta primero cada perfil de negocio **ya validado**
contra logs reales (`ALL_KNOWN_MILESTONES`):

- **`PUSH_MILESTONES`**: flujo Push/PLIN (validacion de pago, Interbank
  switch, autorizacion, notificacion IPN, confirmacion, autoreversa,
  cancelacion). Portado de la antigua skill `trazabilidad-transacciones`
  (2026-09-11).

Si ningun perfil conocido reconoce nada en las lineas de un TransactionId
(porque pertenece a una API/flujo distinto, ej. `authorization`, `cancel`,
`confirmation`, `orderinfo`, `qrnotification`, `security`, `notification.*`,
`capture`, `payments`, `qr`, `yape`, `accountvalidate`, `autoreverse`), cae
automaticamente al **modo generico**: agrupa lineas consecutivas por el
nombre del componente que las genero y clasifica el resultado por palabras
clave (Aprobado/Rechazado/Reversado/Error/Excepcion/HTTP 4xx-5xx). Esto ya da
evidencia completa y ordenada; si se valida un flujo especifico de otra API
contra logs reales, se puede agregar un perfil dedicado (lista propia con el
mismo formato que `PUSH_MILESTONES`, sumada a `ALL_KNOWN_MILESTONES`).

## Salida: estructura del reporte (Excel y HTML, mismo contenido)

- **Resumen**: un bloque por TransactionId encontrado (Comercio, Orden,
  Monto, Cod. Autorizacion, Referencia) + su linea de tiempo (Paso, Hora,
  Evento, Resultado, Estado coloreado, TransactionId de la linea -- solo
  aparece si difiere del TransactionId principal, ej. un ID del switch
  bancario en flujo Push --, Evidencia con ruta/archivo y numero de linea).
- **Detalle_Logs**: todas las lineas crudas encontradas, ordenadas
  cronologicamente, con TransactionId, ruta remota de origen, numero de
  linea, nivel y detalle completo. Se excluyen lineas `Header Authorization`
  y se truncan campos > 30,000 caracteres.

## Reutilizacion

Para una nueva busqueda **solo cambian los criterios y, segun el modo, el
rango de fechas/categoria (SFTP) o la carpeta (local)**:

**Modo SFTP:**
1. Da los criterios (uno o varios) por `input.txt` o flags.
2. Opcionalmente da un rango de fechas (`--fecha` o `--fecha-desde`/`--fecha-hasta`)
   para acotar y acelerar -- sin esto se recorre todo el historico remoto.
3. Opcionalmente acota `--categoria` si ya sabes por donde paso la
   transaccion (mas rapido).
4. Ejecuta el script; ingresa las credenciales SFTP cuando se pidan (o usa
   las guardadas).

**Modo local:**
1. Da los criterios (uno o varios) por `input.txt` o flags.
2. Da `--logs-dir <carpeta>` apuntando a donde estan los logs.
3. Ejecuta el script -- no pide nada mas.

## Requisitos

Python 3.8+, `paramiko` y `openpyxl` (paramiko solo hace falta para modo
SFTP, pero se importa siempre):

```bash
pip install paramiko openpyxl
```

Solo si vas a usar `--guardar-credenciales`/credenciales guardadas (modo
SFTP): SDK de .NET 10 (para compilar `crypto/CryptoWrapper` una vez, ver
seccion de Credenciales guardadas mas arriba).

## Notas y supuestos

- El TransactionId es el dato mas preciso; usalo siempre que lo tengas.
- La busqueda sin `txid` es por texto plano cruzado (AND entre categorias)
  sobre los archivos encontrados (descargados o locales, segun el modo).
- La cache local (`--cache-dir`, solo modo SFTP) se conserva entre corridas a
  proposito: si vas a **refinar** los criterios de una misma investigacion
  (mismo rango de fechas), las corridas siguientes no vuelven a descargar lo
  ya traido. Usa `--limpiar-cache` solo si sospechas que los archivos
  remotos cambiaron (rotacion de logs) o si quieres liberar espacio.
- **Validado contra el servidor SFTP real** multiples veces desde
  2026-09-04 (categorias PUBLICO/CONTROLLER/BUSINESS, varias APIs). El
  patron `RE_FILE_DATE`, el encoding `utf-8` por defecto y el recorrido
  recursivo (`--max-depth`) funcionan correctamente contra logs reales.
  Algunos servidores tienen una subcarpeta "espejo" homonima al lado del
  archivo del dia con una copia duplicada -- `long_path()` evita el
  `FileNotFoundError` por limite MAX_PATH (260 caracteres) de Windows que
  eso puede provocar en rutas locales largas (aplica tanto a la cache SFTP
  como a `--logs-dir`).
- Si el formato NLog cambia, o se quiere mapear un flujo especifico nuevo
  con etiquetas de negocio (en vez de caer al modo generico), agregar una
  lista `<NOMBRE>_MILESTONES` (mismo formato que `PUSH_MILESTONES`: patron,
  etiqueta, colapsable) y sumarla a `ALL_KNOWN_MILESTONES` en
  `buscar_transacciones_sftp.py`. Tambien se puede ajustar
  `GENERIC_MILESTONES`, `RE_MERCHANT`/`RE_ORDER`/`RE_AUTHCODE`/`RE_AMOUNT`/
  `RE_REF` y `classify()` si el formato de esos campos cambia.
