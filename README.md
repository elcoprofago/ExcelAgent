# ExcelAgent

Asistente local que trabaja sobre libros de Excel reales, usando el Excel instalado en esta PC
mediante COM con pywin32. Tiene interfaz grafica con chat, boton para designar el archivo,
registro de eventos y avance en porcentaje.

## Entorno ya instalado

- Python 3.13.15 con winget, alcance usuario
- Entorno virtual: D:\REPOS\EXCEL\.venv
- Librerias: pywin32 312, openpyxl 3.1.5, pytest 9.1.1
- Excel 16.0 build 19127, automatizable por COM

## Como se usa

1. Doble clic en ExcelAgent.bat
2. Boton Abrir Excel: elegi el libro (.xlsx, .xlsm, .xls, .xlsb). Se abre en Excel a la vista.
3. Escribi la consigna en el cuadro de chat y apreta Enviar o Enter.
4. El panel derecho muestra el registro de eventos y el avance en porcentaje.

## Consignas que entiende

    hojas | resumen | leer A1:C10
    escribir 1500 en C2
    formula =B2*C2 en D2
    negrita en A1:C1 | sin negrita en A1:C1
    color rojo en A1:A9 | fondo amarillo en A1:C1
    formato moneda en C2:C50 (o numero, entero, porcentaje, fecha, texto)
    sumar columna C | promedio columna C | maximo columna C | contar columna C
    crear hoja Ventas | activar hoja Ventas | renombrar hoja A a B
    ordenar por columna B descendente
    reemplazar IVA por IVA21 | buscar IVA
    quitar duplicados | autofiltro | quitar filtro | autoajustar
    exportar csv D:/salida.csv
    ejecutar macro MiMacro
    guardar | guardar como D:/copia.xlsx

## Seguridad de tus archivos

- Antes de tocar nada, el programa copia el libro a la subcarpeta _excelagent_respaldos.
- Si el libro ya estaba abierto en tu Excel, se trabaja sobre esa misma ventana.
- Al cerrar la interfaz se te pregunta si queres cerrar Excel tambien.
- El libro se guarda cuando se lo pides con la consigna guardar.

## Pruebas

    D:\REPOS\EXCEL\.venv\Scripts\python.exe -m pytest D:\REPOS\EXCEL\ExcelAgent\tests -q

Los scripts tests\probar_excel.py y tests\probar_excel2.py abren Excel de verdad
sobre libros temporales y verifican el resultado con openpyxl.

## Estructura

    .venv                      entorno virtual con pywin32, openpyxl y pytest
    ExcelAgent\app.py          nucleo: eventos, sesion COM, consignas y agente
    ExcelAgent\gui.py          interfaz grafica
    ExcelAgent\ExcelAgent.bat  lanzador
    ExcelAgent\tests\          pruebas


## A definir

- Licencia: a definir
- Autoria y copyright: a definir


## Interfaz

- Fondo general azul claro, panel de chat celeste claro y panel de registro en azul oscuro.
- El registro usa texto enriquecido: verde para lo hecho, celeste para el avance, amarillo para avisos y rojo para errores.
- El boton + adjunta archivos de texto o imagenes al chat.
- El libro designado se abre minimizado; con la consigna mostrar excel se trae al frente.

## Contactos de Outlook

- Consigna: contactos de Outlook, o bien crea una tabla con los contactos de Outlook.
- Necesita un perfil de Outlook configurado en esta PC. Hoy no hay ninguno, por eso responde
  Outlook no devolvio contactos en la carpeta predeterminada.
- Alternativa ya probada: guarda los contactos en un .txt, .csv o una imagen, adjuntala con el
  boton + y despues usa la consigna volcar adjunto en A1.

## Lectura de texto en imagenes

- Motor: Tesseract 5.5.3 en D:\REPOS\EXCEL\tesseract, instalado sin permisos de administrador
- Consignas: adjuntos, analizar adjunto, volcar adjunto en A1
- Envoltura de Python: pytesseract y pillow en el entorno virtual
- Idiomas cargados ademas de eng: spa, descargado aparte

## Si Excel no responde

Antes de conectarse, el asistente consulta si hay un Excel abierto y le da seis segundos
para contestar. Si no contesta (por ejemplo, hay un cuadro de dialogo abierto o una celda en
edicion), no se le cuelga encima: avisa en el registro y trabaja con una instancia nueva.

Cuando eso pasa, el libro designado se abre en la instancia nueva, no en la ventana trabada.
Para volver a trabajar sobre la original: destrabala en Excel (cerra el dialogo o sali de la
edicion) y apreta el boton Reconectar.

## Cuadro de consignas

- Tiene tres lineas de alto como minimo y crece solo hasta diez mientras escribis.
- Enter envia; Shift+Enter baja de linea.
- Se pueden pegar varias consignas, una por linea.
- Se ejecutan en orden y el avance se reparte entre ellas; todo queda en un mensaje.

## Pendiente

- Empaquetar como .exe para que no aparezca ninguna consola. Se hara cuando la app este
  terminada; por ahora el lanzador usa pythonw.exe, que no abre consola.

## Archivo .log

Cada suceso tambien queda escrito en el proyecto, uno por dia.
Esta en ExcelAgent\logs\excelagent_AAAAMMDD.log. Sirve para revisar despues que ocurrio.

## Lanzadores

- ExcelAgent.vbs: doble clic. Arranca la aplicacion sin ninguna consola. Es el recomendado.
- ExcelAgent.bat: hace lo mismo, delegando en el .vbs; la consola se cierra al instante.
- Si preferis el acceso directo: apuntalo a pythonw.exe con gui.py como argumento, o al .vbs.

## Contactos desde los correos

- Consigna: recopila los contactos de los correos enviados y recibidos, o cualquier frase
  parecida que mencione contactos y correos. Tambien entiende listas numeradas (1-, 2-, 3.).
- Recorre la Bandeja de entrada y los Elementos enviados, y arma dos columnas:
  Nombre en la columna A y correo en la columna B, con el encabezado en negrita.
- Por defecto mira los 1200 correos mas recientes de cada carpeta, unos 2 minutos.
- Con la palabra todos recorre el buzon completo; medido: unos 16 minutos con 20.000 correos.
- Si la carpeta de Contactos de Outlook esta vacia, usa los correos automaticamente.

## Tercera columna con numero de orden (agregado 27-09)

Si la consigna pide 'numero de orden', 'numeracion' o 'columna de orden', se agrega
una tercera columna C con el orden de cada contacto.

## Aclaraciones sobre 'todos'

- 'traer todos los contactos' significa la lista completa de contactos: NO recorre todo el buzon.
- Para recorrer el buzon entero hay que decirlo asi: 'todo el buzon', 'buzon completo',
  'todos los correos', 'todos los mensajes' o 'sin limite'.


## Guardar una copia, renombrar y liberar el libro (agregado 27-09)

- Consigna: salvala como contacto-clean en el escritorio; guardala en mis documentos; libera el proceso.
- Se pueden pedir juntas en un mismo mensaje.
- Guardar como: genera una copia con el nombre indicado y deja intacto el original.
- Si ya existe un archivo con ese nombre, avisa y no lo sobrescribe.
- Liberar: deja libre el libro; si el Excel lo habia abierto el asistente, tambien finaliza esa instancia.
- Si el libro tenia cambios pendientes, los graba antes de dejarlo.

