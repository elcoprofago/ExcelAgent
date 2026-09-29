# ExcelAgent

Asesor de Excel para quien sabe poco de Excel. Se le pide algo con palabras comunes ("poneme el total abajo",
"que los importes se vean como plata", "ordename por cliente") y un modelo de lenguaje interpreta el pedido,
mira el libro abierto y lo resuelve con herramientas concretas sobre Excel, explicando lo que hizo.

El modelo puede ser:

- DeepSeek por API key (pago por token). Es el recomendado.
- Un modelo local .gguf servido con llama-server, si la PC lo soporta. Funciona sin internet, pero los modelos
  chicos (medido con Qwen3-4B y Qwen3.5-9B) se equivocan seguido al usar herramientas: ver "Modelos locales".

## Entorno

- Python 3.13 con un entorno virtual en la carpeta padre: `..\.venv` (fuera del repo, en `.gitignore`).
- Dependencias: `requirements.txt` (pywin32, openpyxl, pytesseract, pillow, pytest). Los modulos de la API y de
  modelos locales solo usan la biblioteca estandar.
- Excel de escritorio para Windows (probado con Microsoft 365, version 16.0.20326). Outlook solo hace falta
  para traer contactos.

Crear el entorno desde cero, parado en la carpeta del repo:

    python -m venv ..\.venv
    ..\.venv\Scripts\python.exe -m pip install -r requirements.txt

Si el entorno se copio desde otra PC, `..\.venv\pyvenv.cfg` apunta al Python de aquella (lineas `home` y
`executable`). Corregir esas dos lineas para que apunten al Python 3.13 de esta PC alcanza; no hace falta rehacerlo.

## Lanzadores

- `ExcelAgent.vbs`: doble clic. Arranca sin consola. Es el recomendado.
- `ExcelAgent.bat`: hace lo mismo, delegando en el .vbs.
- Los dos buscan el entorno en `..\.venv` relativo a su propia carpeta: sirven igual en `D:\REPOS\EXCEL\ExcelAgent`
  que en `F:\source\repos\EXCEL\ExcelAgent`. Si el entorno falta, lo dicen en vez de fallar callados.

## Configuracion

Boton **Configuracion** (se guarda solo al cerrar la ventana):

- **API key**: la de DeepSeek (platform.deepseek.com, saldo prepago). Nunca se guarda en claro: cifrada con el
  usuario de Windows o con una contrasena, a eleccion.
- **Modelos locales**: carpetas donde buscar `.gguf`, ruta de `llama-server.exe`, contexto con que se arranca.
- **Avanzado**: permisos, maximo de tokens por respuesta y presupuesto de tokens de la jornada (el 100% de la
  barra de consumo).
- **Permisos**: "Preguntar antes de pisar, borrar o guardar" (predeterminado), "Pisar datos sin preguntar
  (borrar y guardar, si)" o "No preguntar nada".

La configuracion vive en `%APPDATA%\ExcelAgent\config.json` (o en la carpeta que indique la variable
`EXCELAGENT_CONFIG_DIR`).

## Interfaz

- **Abrir Excel**: elige el libro. Si ya estaba abierto en tu Excel, se trabaja sobre esa misma ventana.
- **+**: adjunta un archivo (texto, csv, contactos .vcf, imagen) para que el asistente lo lea o lo vuelque en la hoja.
- **Modelo** y **Esfuerzo**: selector del modelo (DeepSeek o locales encontrados) y del nivel de razonamiento.
- **Medidor**: que esta haciendo (pensando, razonando, trabajando en Excel, esperando permiso), una barra con los
  tokens del dia contra el presupuesto, el detalle de entrada (con cache), salida (con razonamiento) y la
  velocidad en tok/s. El boton ⟳ consulta el saldo de la cuenta de DeepSeek.
- **Nueva charla**: olvida la conversacion (el libro queda como esta).
- **Enviar / Detener**: Detener corta el trabajo en curso, incluido un recorrido largo de Outlook.
- **Reconectar**: vuelve a buscar el Excel abierto (ver "Si Excel no responde").
- A la derecha, el registro de eventos y avance: cada herramienta que usa el asistente y su resultado.

## Seguridad de tus archivos

- Antes de tocar nada, el libro se copia a la subcarpeta `_excelagent_respaldos` junto al libro.
- Segun la aprobacion elegida, el asistente pregunta antes de pisar datos, borrar o guardar.
- El libro se guarda solo cuando se lo pedis.
- Si el asistente repite la misma accion sin avanzar, se corta solo en vez de gastar tokens.
- Al cerrar la interfaz con un Excel conectado se pregunta si cerrarlo tambien.

## Que sabe hacer

Leer y buscar; escribir valores y formulas; rellenar formulas; formatos de numero (moneda, porcentaje, miles,
fechas), fuentes y colores; autoajustar columnas; ordenar, filtrar, tablas, quitar duplicados, reemplazar,
borrar; insertar y eliminar filas o columnas; hojas; graficos; formato condicional; listas desplegables;
inmovilizar paneles; guardar, guardar copia, exportar a CSV; manejar la ventana de Excel; leer y pegar adjuntos;
traer contactos de Outlook o de la agenda del telefono.

Las formulas se escriben en ingles y Excel las muestra en castellano. Los formatos de numero tambien viajan en
ingles; esto se corrigio el 29/09/2026: antes, en un Excel en castellano, "moneda" se veia `$ 1234,5000`.

## Contactos de Outlook

- Pedido: "traeme los contactos de Outlook", o "armame una tabla con los contactos de mis correos".
- Necesita un perfil de Outlook configurado. Un Outlook instalado sin cuenta se queda en el asistente de
  bienvenida: Detener corta la espera.
- Si la carpeta de Contactos esta vacia, junta remitentes y destinatarios de la Bandeja de entrada y de Elementos
  enviados: por defecto los 1200 correos mas recientes de cada carpeta (unos 2 minutos). Pidiendo "todos",
  recorre el buzon completo; medido: unos 16 minutos con 20.000 correos.

## Contactos del telefono (y de WhatsApp)

- WhatsApp no tiene agenda propia: usa la del telefono. ExcelAgent no lee WhatsApp Web (automatizarlo va contra las
  condiciones de uso de WhatsApp y arriesga la cuenta); la via es exportar la agenda y adjuntarla con **+**.
- Formatos: vCard (`.vcf`, lo que exporta la app Contactos de Android o icloud.com) y CSV de contactos (Google
  Contactos, Outlook). Se reducen a Nombre, Telefono, Otros telefonos y Correo (`contactos.py`).
- Los telefonos se escriben como texto: sin eso Excel les sacaba el `+`, los mostraba como `5,49114E+12` o tomaba
  `+54 9 11 ...` por una formula. Vale para toda columna cuyo encabezado diga telefono, celular, movil, phone o
  WhatsApp, y para todo valor con forma de telefono.
- Para pocos contactos alcanza una captura de la lista adjuntada con **+** (OCR: conviene revisar los numeros).

## Lectura de texto en imagenes

- Motor: Tesseract 5.5.3 en `..\tesseract` (fuera del repo), con los idiomas eng y spa. Tambien se busca en
  `C:\Program Files\Tesseract-OCR` o donde indique la variable `TESSERACT_CMD`.
- Envoltura de Python: pytesseract y pillow en el entorno virtual.

## Si Excel no responde

Antes de conectarse, el asistente consulta si hay un Excel abierto y le da seis segundos para contestar. Si no
contesta (un cuadro de dialogo abierto, una celda en edicion), no se le cuelga encima: avisa en el registro y
trabaja con una instancia nueva. Para volver a la original: destrabala en Excel y apreta **Reconectar**.

Solo se usa un Excel abierto que tenga ventana. Uno sin ventana es de otro programa que automatiza Excel (o de
las pruebas): no se toca y se trabaja con una instancia propia. Antes se adoptaba y podia fallar con "Excel no
pudo abrir el archivo" (medido el 29/09/2026). Si el libro igual no queda abierto, se reintenta una vez en una
instancia propia.

## Modelos locales

Se arrancan con llama-server al primer pedido y se apagan al salir. ExcelAgent no depende de ningun otro
programa: `llama-server.exe` y sus DLL van en `..\bin` (fuera del repo, como `..\tesseract`) o en `bin\` junto al
programa, o se elige la ruta en Configuracion. Los modelos se buscan en `Models\` junto al programa, en su carpeta
padre y en la raiz de la unidad; otras carpetas se agregan en Configuracion.

En esta PC, `..\bin` tiene llama-server b9775 con CUDA (678 MB). Pruebas reales del
29/09/2026 con el mismo pedido coloquial (ordenar, formatear como moneda, poner un total):

- DeepSeek Flash: todo correcto, 8 herramientas, unos 24.000 tokens de entrada.
- Qwen3-4B-Instruct: ordeno bien, formato equivocado, no puso el total.
- Qwen3.5-9B: ordeno y formateo bien, despues repitio una escritura vacia hasta el tope de pasos. Eso dejo dos
  correcciones: la escritura vacia ahora es un error y la repeticion identica corta el bucle.

## Archivo .log

Un archivo por dia en `logs\excelagent_AAAAMMDD.log` (en `.gitignore`). La variable `EXCELAGENT_LOG_DIR` lo manda a
otra carpeta: las pruebas la usan para no mezclarse con el registro real. La respuesta del asistente va en una sola
linea (antes, una linea por cada pieza del texto).

## Pruebas

    ..\.venv\Scripts\python.exe -m pytest tests -q

- Abren Excel sin ventana y trabajan con libros de juguete en una carpeta temporal. Tardan unos dos minutos.
- `tests\modulos\`: chequeos de la API, la configuracion, el medidor y los modelos locales, que se corren como
  scripts aparte (`test_modulos.py` los lanza). El de modelos locales usa archivos .gguf falsos dispersos (no
  ocupan disco) y los borra aunque la prueba muera a mitad de camino.
- No usan la API key real ni gastan tokens: el modelo se reemplaza por un guion.
- La prueba de graficos fallo una vez de cada varias corridas sin causa encontrada; repetida, pasa.
- La de modelos locales en la GPU se omite (y lo dice) si otros programas ocupan la VRAM y el modelo de prueba no
  entra.
- Mientras corren, conviene no usar **Abrir Excel** en un ExcelAgent de una version anterior al 29/09/2026: esa
  version adoptaba el Excel oculto de las pruebas.

## Estructura

    gui.py            Interfaz tkinter
    dialogos.py       Configuracion, API key, permisos
    asistente.py      Une la GUI con el modelo elegido y con Excel; cuenta tokens
    agente.py         Bucle modelo -> herramientas -> modelo, con freno ante repeticiones
    prompts.py        Instrucciones del asesor
    herramientas.py   Las herramientas que el modelo puede usar sobre Excel
    excel.py          Sesion COM con Excel, respaldos, adjuntos, OCR, Outlook
    dsapi.py          API de DeepSeek y configuracion
    localmodels.py    Modelos .gguf con llama-server
    meter.py          Medidor de tokens
    contactos.py      Contactos exportados del telefono (.vcf) o en CSV; telefonos como texto
    secret.py         Cifrado de la key con contrasena
    limpiar_temporales.ps1   Borra los restos de desarrollo de la version anterior (probar con -WhatIf)
    tests\            Pruebas
