---
name: monitoreo-ipn-publico
description: >-
  Monitoreo diario de notificaciones IPN (webhooks de pago) en los 4
  servidores PUBLICO de produccion (SVPRDLW364, SVPRDMW363, SVPRDLW1021,
  SVPRDMW1020): se conecta por SFTP y descubre TODAS las peticiones de
  notificacion que hicieron TODOS los comercios en una fecha dada (correctas
  e incorrectas), sin necesitar un TransactionId/comercio puntual de
  antemano. Agrupa por URL de notificacion del comercio y clasifica cada
  FALLA en una causa especifica y accionable: proxy bloqueado por ACL
  (accionable con Redes), proxy sin poder conectar al destino (URL del
  comercio invalida/no configurada), certificado invalido o vencido del
  comercio, certificado "a vigilar" (posible CA/intermedia faltante del lado
  de Izipay, cuando el error se concentra solo en los servidores con
  historial de trust store incompleto), timeout de conexion, rechazo HTTP
  propio del comercio, o respuesta 2xx distinta de 200. Tambien reporta la
  ultima hora de escritura de log de cada servidor/puerto (para confirmar
  que los 4 siguen activos) y cuantos intentos de notificacion tuvo cada uno.
  Usar al iniciar el dia (o cuando se quiera un chequeo de salud de las
  notificaciones IPN) para revisar de un vistazo que comercios tuvieron
  fallas hoy y priorizar cual atacar primero, sin tener que investigar
  transaccion por transaccion. Entrega el reporte en Excel y HTML (mismo
  diseño de marca Izipay que el resto de reportes del equipo). Alcance
  actual: notificaciones IPN (`notification.dispatcher1`, API validada
  contra logs reales) en categoria PUBLICO -- no cubre otras APIs de
  transacciones (authorization/cancel/confirmation/etc.), que tienen un
  formato de log y una nocion de "error" distintos.
---

# Monitoreo diario de notificaciones IPN (servidores PUBLICO)

Skill de **descubrimiento agregado**, no de investigacion puntual: en vez de
buscar una transaccion especifica (eso ya lo cubre
`busqueda-transacciones-local-sftp`), esta skill escanea **todo** lo que paso
en un dia dado para **todos los comercios**, y arma un reporte de salud de
las notificaciones IPN con las fallas ya clasificadas por causa.

Nace de consolidar el analisis manual hecho el 2026-09-11 (3 scripts sueltos
de scratchpad: descubrimiento por servidor/URL, ultima escritura, y
construccion del reporte) en una sola skill reutilizable dia a dia.

> Reutiliza la conexion/credenciales SFTP y la topologia (`remote_paths.txt`)
> de la skill hermana **`busqueda-transacciones-local-sftp`** (deben estar
> en la misma carpeta `.claude/skills/`), y el diseño de marca Izipay de
> **`reportes-html-izipay`** / **`reportes-excel-izipay`**. No duplica ese
> codigo, lo importa directo.

## Cuando usar esta skill

Frases tipicas: "revisa las notificaciones IPN de hoy", "quiero ver que
comercios tuvieron error de notificacion hoy", "chequeo diario de
notificaciones", "cuantos servidores estan activos y cual es su ultima
escritura de log", "dame el monitoreo de siempre para hoy/ayer".

> Si en cambio se necesita investigar UNA transaccion/comercio puntual
> (con o sin TransactionId), usar `busqueda-transacciones-local-sftp`.

## Que hace, paso a paso

1. Conecta una sola vez por SFTP (credenciales guardadas de la skill
   hermana, no pide nada si ya estan configuradas).
2. Lista (sin descargar) todos los archivos de `notification.dispatcher1`
   en los 4 servidores PUBLICO, calcula la **ultima hora de escritura** por
   servidor/puerto, y filtra los candidatos de la fecha pedida.
3. Descarga esos candidatos a una cache propia de esta skill (`.cache/`,
   con la misma descarga paralela por peso que la skill hermana).
4. En una sola pasada, por cada linea de notificacion encontrada:
   - Agrupa por **URL de destino** (la del comercio).
   - Extrae el **Codigo de Comercio** (`merchantCode`) de la peticion.
   - Clasifica cada intento en exitoso (200) o fallido, y cada fallido en
     una **causa** especifica (ver tabla abajo).
5. Genera el reporte (Excel + HTML) con: actividad por servidor + ultima
   escritura, estadisticas generales, callouts para causas "a vigilar" y
   "proxy bloqueado", y una tabla de **causas de falla** (una fila por
   causa dentro de cada URL, ordenada de mayor a menor impacto).

## Causas de falla detectadas

| Causa | Severidad | Significado | Quien debe actuar |
|---|---|---|---|
| Proxy Izipay: bloqueo ACL | critica | El proxy corporativo responde 403 al intento de tunel, antes de llegar al destino | Redes (revisar categorizacion del dominio) |
| A vigilar (Izipay) | advertencia | Error de certificado SOLO en `SVPRDLW364`/`SVPRDMW363` (exito en servidor sano) -- posible CA/intermedia faltante | Infra (verificar almacen de certificados) |
| Comercio: certificado invalido/vencido | seria | Error de certificado que ocurre tambien en servidores sanos | El comercio (renovar su certificado) |
| Comercio: proxy no pudo conectar | seria | El proxy no logra conectar con el destino (codigo distinto de 403, ej. 500) -- tipicamente URL invalida/no configurada | El comercio (corregir su URL de notificacion) |
| Comercio: timeout de conexion | seria | El servidor del comercio no responde en 100s | El comercio (su servidor esta lento/caido) |
| Comercio: HTTP `<codigo>` | seria | El endpoint del comercio devuelve un error propio (400/403/404/500/etc.) | El comercio |
| Comercio: respondio `2xx` no-200 | informativa | El comercio respondio bien pero con un codigo 2xx distinto de 200 (el dispatcher exige 200 exacto) | Informativo, baja prioridad |

La distincion "A vigilar" vs "Comercio: certificado" se decide automaticamente
comparando en que servidores ocurrio la falla: si es SOLO en
`SVPRDLW364`/`SVPRDMW363` (con exito en un servidor sano para la misma URL),
es probable un problema de Izipay; si ocurre tambien en los servidores sanos,
es el certificado del comercio.

## Como ejecutar

Uso mas comun (hoy, notification.dispatcher1, categoria PUBLICO -- todo por
defecto):

```bash
python "<ruta-skill>/monitorear_ipn_publico.py"
```

Otro dia:

```bash
python "<ruta-skill>/monitorear_ipn_publico.py" --fecha 2026-09-10
```

Otras APIs de notificacion (sin validar todavia contra logs reales, pero
deberian compartir el mismo formato):

```bash
python "<ruta-skill>/monitorear_ipn_publico.py" --api notification.dispatcher1,notification.dispatcher2
```

Parametros principales:
- `--fecha` : fecha a analizar, `YYYY-MM-DD` (default: hoy).
- `--api` : API(s) de notificacion, coma-separado (default `notification.dispatcher1`).
- `--categoria` : categoria de servidores (default `PUBLICO`, el alcance validado).
- `--credenciales` / `--host` / `--port` / `--usuario` / `--auth` /
  `--guardar-credenciales` / `--olvidar-credenciales` : mismo esquema que
  `busqueda-transacciones-local-sftp` (default: reusa las credenciales ya
  guardadas de esa skill).
- `--cache-dir` : cache de descargas propia de esta skill (default `.cache`
  dentro de esta carpeta; no comparte cache con la skill hermana).
- `--paralelo-livianos` / `--paralelo-pesados` / `--umbral-pesado-mb` :
  mismos defaults que la skill hermana (5/2/300).
- `--output-xlsx` / `--output-html` : nombre de salida (default
  `Monitoreo_IPN_Publico_<fecha>.xlsx`/`.html`). Siempre se guardan en
  `_reports/<fecha>/` de esta skill.

## Salida del reporte

- **Excel** (`Resumen` + `Causas_de_falla`): igual estructura que el
  analisis manual del 2026-09-11 -- actividad por servidor con ultima
  escritura, y una tabla con Comercio(s)/URL/OK/FALLO/Detalle por
  servidor/Categoria coloreada/Ejemplo de error.
- **HTML**: mismo contenido con el diseño de marca Izipay (`reportes-html-izipay`):
  tiles de estadisticas, callouts para "a vigilar"/"proxy bloqueado", tabla
  de causas buscable y ordenable por columna.

## Alcance y limitaciones

- **Solo notificaciones IPN, solo categoria PUBLICO.** No cubre
  `authorization`/`cancel`/`confirmation`/`capture`/`payments`/`push`/`qr`/`yape`/etc.
  -- esas APIs no llaman a una "URL del comercio" de la misma forma
  (procesan contra el switch bancario), tienen otro formato de log y otra
  nocion de "error"; agregarlas requeriria un modulo de analisis nuevo, no
  solo cambiar `--api`.
- **Solo `notification.dispatcher1` esta validado** contra logs reales.
  `notification.dispatcher2`/`notification.producer` deberian compartir el
  mismo formato (misma familia de dispatcher), pero no se ha confirmado
  todavia -- probar y ajustar `TARGET_RE`/`RESULT_RE` en
  `monitorear_ipn_publico.py` si el formato real difiere.
- Requiere que la skill hermana `busqueda-transacciones-local-sftp` este
  presente (mismo directorio padre `.claude/skills/`) y con credenciales
  SFTP ya configuradas (`crypto/sftp_config.json` de esa skill).

## Requisitos

Los mismos que las skills que reutiliza: Python 3.8+, `paramiko`,
`openpyxl`. No requiere instalar nada adicional si
`busqueda-transacciones-local-sftp` ya esta funcionando.
