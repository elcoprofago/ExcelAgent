# Capa de Excel: la conexion COM con el libro y las operaciones que usa el asistente.

from __future__ import annotations

import os
import re
import threading
import time

from dataclasses import dataclass

NL = chr(10)
XL_MINIMIZADO = -4140    # xlMinimized
XL_NORMAL = -4143        # xlNormal


def hora_actual():
    t = time.localtime()
    hh = str(t.tm_hour).zfill(2)
    mm = str(t.tm_min).zfill(2)
    ss = str(t.tm_sec).zfill(2)
    return hh + ':' + mm + ':' + ss

@dataclass
class Evento:
    'Evento emitido por el agente.'
    tipo: str
    texto: str
    nivel: str = 'info'
    pct: float = 0.0
    hora: str = str()


class Bus:
    'Recolecta eventos y los reenvia a la interfaz.'

    def __init__(self, sink=None):
        self.sink = sink
        self.historial = []


    def emite(self, tipo, texto, nivel='info', pct=0.0):
        ev = Evento(tipo, texto, nivel, pct, hora_actual())
        self.historial.append(ev)
        if self.sink:
            self.sink(ev)

    def log(self, texto, nivel='info'):
        self.emite('log', texto, nivel)

    def chat(self, texto, nivel='chat'):
        self.emite('chat', texto, nivel)

    def progreso(self, pct, texto):
        self.emite('progreso', texto, 'avance', pct)


class ExcelError(RuntimeError):
    'Error controlado de la sesion con Excel.'


def con_limite(funcion, segundos, que='Excel'):
    'Ejecuta una llamada COM en un hilo aparte y falla si no responde a tiempo.'
    caja = {}
    def correr():
        import pythoncom
        pythoncom.CoInitialize()
        try:
            caja['valor'] = funcion()
        except Exception as exc:
            caja['error'] = exc
    hilo = threading.Thread(target=correr, daemon=True)
    hilo.start()
    hilo.join(segundos)
    if hilo.is_alive():
        raise ExcelError(que + ' no respondio en ' + str(int(segundos)) + ' segundos. Parece trabado.')
    if 'error' in caja:
        raise caja['error']
    if 'valor' not in caja:
        raise ExcelError(que + ' se corto de forma inesperada al consultarlo.')
    return caja['valor']


class Sesion:
    'Conexion COM con el Excel instalado y el libro abierto.'

    def __init__(self, bus):
        self.bus = bus
        self.app = None
        self.wb = None
        self.ruta = None
        self.respaldo = None
        self.propia = False
        self.abierto_por_nosotros = False
        self.adjuntos = []      # adjuntos cargados desde la interfaz
        self.detener = threading.Event()


    def conectar(self, visible=True):
        import pythoncom
        import win32com.client as win32
        pythoncom.CoInitialize()
        if self.app is None:
            viva, aviso = self._instancia_que_responde(win32)
            if aviso:
                self.bus.log(aviso, 'warn')
            if viva is not None:
                self.app = viva
                origen = 'una instancia de Excel ya abierta'
                try:
                    self.app.DisplayAlerts = False
                    self.app.Visible = bool(visible)
                except Exception as exc:
                    self.bus.log('La instancia adoptada no respondia: armo una nueva. ' + str(exc), 'warn')
                    self.app = None
            if self.app is None:
                self._instancia_nueva(win32, visible)
                origen = 'una instancia nueva de Excel'
        else:
            origen = 'la conexion ya abierta'
            self.app.Visible = bool(visible)
        self.bus.log('Conexion lista con ' + origen, 'ok')

    def _instancia_que_responde(self, win32):
        # Nunca se adhiere a un Excel trabado. La sonda corre en un hilo aparte,
        # pero el objeto se toma en este hilo: cruzar objetos COM entre hilos los invalida.
        # Solo se adopta un Excel con ventana: uno sin ventana es de otro programa que lo automatiza
        # (o de las pruebas) y puede cerrarse en cualquier momento. Medido el 29/09/2026: adoptar el
        # Excel oculto de las pruebas hizo fallar 'Excel no pudo abrir el archivo'. La sonda solo lee.
        def sonda():
            obj = win32.GetActiveObject('Excel.Application')
            return bool(obj.Visible), int(obj.Workbooks.Count)
        try:
            visible, _ = con_limite(sonda, 6)
        except ExcelError as exc:
            self.bus.log('No pude usar el Excel abierto: ' + str(exc), 'warn')
            return None, 'El Excel abierto no se puede usar: trabajo con una instancia nueva. Destrabalo y apreta Reconectar.'
        except Exception:
            return None, str()
        if not visible:
            self.bus.log('Hay un Excel sin ventana (de otro programa): no lo toco y trabajo con una instancia nueva.', 'info')
            return None, str()
        try:
            return win32.GetActiveObject('Excel.Application'), str()
        except Exception:
            return None, str()

    def abrir(self, ruta, visible=True, respaldo=True, minimizar=True):
        ruta = os.path.abspath(ruta)
        if not os.path.isfile(ruta):
            raise ExcelError('No existe el archivo: ' + ruta)
        self.conectar(visible)
        if minimizar:
            # Medido: en una instancia nueva, minimizar recien despues de Open no tiene efecto
            # (WindowState sigue en normal) salvo que la ventana se haya minimizado antes de abrir.
            try:
                self.app.WindowState = XL_MINIMIZADO
            except Exception:
                pass
        self.wb = self._buscar_abierto(ruta)
        if self.wb is None:
            self.wb = self._abrir_libro(ruta, visible)
            self.abierto_por_nosotros = True
        if self.wb is None:
            self.wb = self._buscar_abierto(ruta)
        if self.wb is None and not self.propia:
            # Open en un Excel ajeno puede no fallar y aun asi no dejar el libro (por ejemplo, si esa
            # instancia se estaba cerrando): se reintenta una vez en una instancia propia.
            self.bus.log('El Excel en uso no dejo el libro abierto. Pruebo con una instancia nueva.', 'warn')
            import win32com.client as win32
            self._instancia_nueva(win32, visible)
            self.wb = self._abrir_libro(ruta, visible) or self._buscar_abierto(ruta)
        if self.wb is None:
            raise ExcelError('Excel no pudo abrir el archivo: ' + ruta)
        self.ruta = ruta
        self.bus.log('Libro abierto: ' + ruta, 'ok')
        if minimizar:
            self.minimizar()
        if respaldo:
            self.respaldo = self.copia_respaldo()
        return self.hojas()


    def _buscar_abierto(self, ruta):
        # Las llamadas COM van siempre en este mismo hilo: cruzar objetos entre hilos
        # los invalida y Windows termina el proceso (RPC_E_WRONG_THREAD).
        try:
            return self.buscar_abierto(ruta)
        except Exception as exc:
            self.bus.log('No pude revisar los libros abiertos: ' + str(exc)[:70], 'warn')
            return None

    def _abrir_libro(self, ruta, visible=True):
        # Si el Excel en uso no puede abrir el libro, se arma una instancia nueva.
        try:
            return self.app.Workbooks.Open(ruta)
        except Exception as exc:
            self.bus.log('El Excel en uso no pudo abrir el libro: ' + str(exc)[:70] + '. Pruebo con una instancia nueva.', 'warn')
            import win32com.client as win32
            self._instancia_nueva(win32, visible)
            self.bus.log('Conexion lista con una instancia nueva de Excel', 'ok')
            return self.app.Workbooks.Open(ruta)

    def _instancia_nueva(self, win32, visible):
        self.app = win32.DispatchEx('Excel.Application')
        self.propia = True
        self.app.DisplayAlerts = False
        self.app.Visible = bool(visible)


    def minimizar(self):
        'Deja la ventana de Excel en la barra de tareas.'
        if self.app is None:
            return
        if not self.propia:
            self.bus.log('Excel ya estaba abierto: tambien minimizo esa ventana', 'warn')
        self.app.WindowState = XL_MINIMIZADO
        self.bus.log('Ventana de Excel minimizada', 'ok')

    def mostrar(self):
        'Trae la ventana de Excel al frente.'
        self.app.WindowState = XL_NORMAL
        self.app.Visible = True
        self.bus.log('Ventana de Excel a la vista', 'ok')

    def buscar_abierto(self, ruta):
        objetivo = os.path.abspath(ruta).lower()
        for wb in self.app.Workbooks:
            try:
                if os.path.abspath(str(wb.FullName)).lower() == objetivo:
                    return wb
            except Exception:
                continue
        return None

    def copia_respaldo(self):
        carpeta = os.path.join(os.path.dirname(self.ruta), '_excelagent_respaldos')
        os.makedirs(carpeta, exist_ok=True)
        t = time.localtime()
        sello = str(t.tm_year) + str(t.tm_mon).zfill(2) + str(t.tm_mday).zfill(2)
        hora = str(t.tm_hour).zfill(2) + str(t.tm_min).zfill(2) + str(t.tm_sec).zfill(2)
        base = os.path.splitext(os.path.basename(self.ruta))[0]
        ext = os.path.splitext(self.ruta)[1]
        destino = os.path.join(carpeta, base + '_' + sello + '_' + hora + ext)
        with open(self.ruta, 'rb') as f1, open(destino, 'wb') as f2:
            f2.write(f1.read())
        self.bus.log('Copia de respaldo: ' + destino, 'ok')
        return destino

    def hojas(self):
        return [ws.Name for ws in self.wb.Worksheets]

    def hoja(self, nombre=None):
        if self.wb is None:
            raise ExcelError('No hay ningun libro abierto')
        if not nombre:
            return self.wb.ActiveSheet
        nombres = self.hojas()
        if nombre not in nombres:
            raise ExcelError('No existe la hoja ' + nombre + '. Hojas: ' + ', '.join(nombres))
        return self.wb.Worksheets(nombre)

    def cerrar(self, guardar=False):
        if self.wb is not None and self.abierto_por_nosotros:
            self.wb.Close(guardar)
            self.abierto_por_nosotros = False
        self.wb = None

    def cambios_pendientes(self):
        'Dice si el libro que abrio el asistente tiene cambios sin guardar.'
        if self.wb is None or not self.abierto_por_nosotros:
            return False
        try:
            return not bool(self.wb.Saved)
        except Exception:
            return True     # Si no se puede saber, se trata como pendiente: nunca se descarta a ciegas.

    def entregar_al_usuario(self):
        # Deja el libro abierto y a la vista, con las alertas de Excel activas, para que
        # el usuario decida. UserControl evita que Excel se cierre al soltar la conexion.
        try:
            nombre = str(self.wb.Name)
        except Exception:
            nombre = os.path.basename(str(self.ruta or 'abierto'))
        try:
            self.app.DisplayAlerts = True
            self.app.Visible = True
            self.app.WindowState = XL_NORMAL
            self.app.UserControl = True
        except Exception as exc:
            self.bus.log('No pude dejar Excel a la vista: ' + str(exc), 'warn')
        self.bus.log('El libro ' + nombre + ' tiene cambios sin guardar: queda abierto en Excel para que decidas si guardarlos.', 'warn')
        self.wb = None
        self.app = None
        self.abierto_por_nosotros = False

    def salir(self, cerrar_excel=False, descartar=False):
        if self.cambios_pendientes() and not descartar:
            self.entregar_al_usuario()
            self.bus.log('Sesion finalizada', 'ok')
            return
        self.cerrar(False)
        if cerrar_excel and self.app is not None:
            if self.propia:
                self.app.Quit()
            self.app = None
        self.bus.log('Sesion finalizada', 'ok')

# Nombres de colores en castellano (herramientas.COLOR_NAMES les suma los ingleses). Los formatos de numero ya no
# estan aca: ver herramientas.NUMBER_FORMATS y set_number_format.
COLORES = {
      'rojo': '#FF0000'
    , 'verde': '#00B050'
    , 'amarillo': '#FFFF00'
    , 'azul': '#0070C0'
    , 'naranja': '#FFC000'
    , 'violeta': '#7030A0'
    , 'gris': '#D9D9D9'
    , 'celeste': '#00B0F0'
    , 'rosa': '#FF99CC'
    , 'negro': '#000000'
}


def color_bgr(color):
    h = color.lstrip('#')
    r = int(h[0:2], 16)
    g = int(h[2:4], 16)
    b = int(h[4:6], 16)
    return r + (g << 8) + (b << 16)


def col_num(letra):
    n = 0
    for ch in letra.upper():
        n = n * 26 + (ord(ch) - 64)
    return n


def valor(texto):
    t = texto.strip()
    if len(t) > 1 and t[0] == t[-1] and t[0] in (chr(34), chr(39)):
        return t[1:-1]
    if re.fullmatch('[+-]?[0-9]+', t):
        return int(t)
    if re.fullmatch('[+-]?[0-9]*[.,][0-9]+', t):
        return float(t.replace(',', '.'))
    return t


def col_letter(num):
    'Convierte un numero en letra de columna: 1 -> A, 27 -> AA.'
    s = str()
    while num > 0:
        num, resto = divmod(num - 1, 26)
        s = chr(65 + resto) + s
    return s


def rango_desde(ref, filas, columnas):
    'Arma el rango explicito A1:C4. Range.Resize no acepta escritura en bloque.'
    m = re.match('([A-Za-z]{1,3})([0-9]{1,7})', ref)
    if not m:
        raise ExcelError('Referencia de celda invalida: ' + ref)
    c1 = col_num(m.group(1))
    f1 = int(m.group(2))
    c2 = c1 + max(columnas, 1) - 1
    f2 = f1 + max(filas, 1) - 1
    return col_letter(c1) + str(f1) + ':' + col_letter(c2) + str(f2)


PR_SMTP = 'http://schemas.microsoft.com/mapi/proptags/0x39FE001E'
LIMITE_CORREOS = 1200
CARPETAS_CORREO = (('Bandeja de entrada', 6), ('Elementos enviados', 5))


def limpiar_nombre(nombre):
    # Outlook suele devolver los nombres de Exchange con comillas o dos puntos pegados.
    n = str(nombre or str()).strip()
    comilla = chr(39)
    doble = chr(34)
    n = n.strip(comilla).strip(doble).strip()
    while n.startswith(':') or n.startswith(comilla) or n.startswith(doble):
        n = n[1:].strip()
    n = n.strip(comilla).strip(doble).strip()
    return n


def _smtp_de_entrada(entrada, cache):
    # Devuelve el par nombre y correo de un remitente o destinatario, con cache.
    try:
        clave = str(entrada.Address)
    except Exception:
        return (str(), str())
    if clave in cache:
        return cache[clave]
    par = (str(), str())
    try:
        ae = entrada.AddressEntry
        if ae.Type == 'EX':
            nombre = str(ae.Name)
            correo = str()
            try:
                correo = str(ae.PropertyAccessor.GetProperty(PR_SMTP))
            except Exception:
                correo = str()
            if not correo:
                ex = ae.GetExchangeUser()
                if ex is not None:
                    correo = str(ex.PrimarySmtpAddress)
            par = (nombre, correo)
        else:
            par = (str(ae.Name), str(ae.Address))
    except Exception:
        par = (str(), str())
    par = (limpiar_nombre(par[0]), par[1].strip())
    cache[clave] = par
    return par


def correo_valido(direccion):
    # Descarta direcciones vacias o mal formadas, como undisclosed-recipients.
    d = str(direccion or str()).strip()
    if d.count('@') != 1:
        return str()
    usuario, dominio = d.split('@')
    if not usuario or not dominio or ' ' in d:
        return str()
    return d


def op_contactos_correo(s, todos=False, con_orden=False):
    # Recorre los correos recibidos y enviados y arma el listado de contactos.
    import win32com.client as win32
    if s.detener.is_set():
        # antes de abrir Outlook: sin cuenta configurada, queda en el asistente de bienvenida y no vuelve
        raise ExcelError('Trabajo detenido por el usuario: se revisaron 0 correos.')
    try:
        ol = win32.GetActiveObject('Outlook.Application')
    except Exception:
        ol = win32.Dispatch('Outlook.Application')
    ns = ol.GetNamespace('MAPI')
    limite = 0 if todos else LIMITE_CORREOS
    cache = {}
    encontrados = {}
    propias = set()
    try:
        ae = ns.CurrentUser.AddressEntry
        propias.add(str(ae.Address).strip().lower())
        ex = ae.GetExchangeUser()
        if ex is not None:
            propias.add(str(ex.PrimarySmtpAddress).strip().lower())
    except Exception:
        pass
    propias.discard(str())
    revisados = 0
    for numero, (nombre, idx) in enumerate(CARPETAS_CORREO):
        try:
            carpeta = ns.GetDefaultFolder(idx)
        except Exception as exc:
            s.bus.log('No pude abrir ' + nombre + ': ' + str(exc), 'warn')
            continue
        items = carpeta.Items
        en_carpeta = 0
        item = items.GetFirst()
        while item is not None and (todos or en_carpeta < limite):
            en_carpeta += 1
            revisados += 1
            if s.detener.is_set():
                raise ExcelError('Trabajo detenido por el usuario: se revisaron ' + str(revisados) + ' correos.')
            if en_carpeta in (200, 400, 600, 800, 1000, 1200):
                total_carpeta = limite or items.Count
                s.bus.log('Revisando ' + nombre + ': ' + str(en_carpeta) + ' de ' + str(total_carpeta) + ' correos')
                avance = (numero + en_carpeta / float(total_carpeta)) * 45.0
                s.bus.progreso(min(avance, 95.0), 'Revisando ' + nombre)
            try:
                if item.Class == 43:
                    n, d = _smtp_de_entrada(item.Sender, cache)
                    d = correo_valido(d)
                    if numero == 1 and d:
                        propias.add(d.lower())
                    if d:
                        clave = d.lower()
                        if clave not in encontrados or not encontrados[clave]:
                            encontrados[clave] = n
                    for r in item.Recipients:
                        nr, dr = _smtp_de_entrada(r, cache)
                        dr = correo_valido(dr)
                        if dr:
                            clav = dr.lower()
                            if clav not in encontrados or not encontrados[clav]:
                                encontrados[clav] = nr
            except Exception:
                pass
            item = items.GetNext()
    propias.discard(str())
    for correo_propio in propias:
        encontrados.pop(correo_propio, None)
    if not encontrados:
        raise ExcelError('No pude extraer contactos de los correos revisados')
    encabezado = ['Nombre', 'Correo']
    if con_orden:
        encabezado.append('Orden')
    filas = [encabezado]
    orden = 0
    for correo in sorted(encontrados, key=lambda c: (encontrados[c] or c).lower()):
        orden += 1
        fila = [encontrados[correo] or correo, correo]
        if con_orden:
            fila.append(orden)
        filas.append(fila)
    ancho = len(encabezado)
    h = s.hoja()
    destino = rango_desde('A1', len(filas), ancho)
    h.Range(destino).Value2 = tuple(tuple(f) for f in filas)
    h.Range(rango_desde('A1', 1, ancho)).Font.Bold = True
    h.Range(destino).Columns.AutoFit()
    detalle = 'Contactos volcados en la hoja ' + h.Name + ': ' + str(len(filas) - 1) + ' contactos'
    if con_orden:
        detalle = detalle + NL + 'Nombre en la columna A, correo en la columna B y numero de orden en la columna C (revisados ' + str(revisados) + ' correos).'
    else:
        detalle = detalle + NL + 'Nombre en la columna A y correo en la columna B (revisados ' + str(revisados) + ' correos).'
    if not todos:
        detalle = detalle + NL + 'Se miraron los ' + str(LIMITE_CORREOS) + ' correos mas recientes de cada carpeta.'
        detalle = detalle + NL + 'Para revisar todo el buzon hay que pedirlo (tarda varios minutos; medido: unos 16 con 20.000 correos).'
    return detalle

def carpeta_conocida(texto):
    # Resuelve escritorio o documentos a una ruta real de esta PC.
    casa = os.path.expanduser('~')
    bajo = texto.lower()
    if 'escritorio' in bajo or 'desktop' in bajo:
        ruta = os.path.join(casa, 'Desktop')
        if os.path.isdir(ruta):
            return ruta
    if 'documento' in bajo:
        ruta = os.path.join(casa, 'Documents')
        if os.path.isdir(ruta):
            return ruta
    return None

EXTENSIONES_EXCEL = ('.xlsx', '.xlsm', '.xls', '.xlsb')


def destino_copia(s, nombre, carpeta=None):
    # Arma la ruta de la copia: sin extension, lleva la del original; sin carpeta, va junto al original.
    if os.path.splitext(nombre)[1].lower() not in EXTENSIONES_EXCEL:
        nombre = nombre + os.path.splitext(s.ruta)[1]
    if os.path.isabs(nombre):
        return nombre
    return os.path.join(carpeta or os.path.dirname(s.ruta), nombre)


def op_guardar_como(s, destino):
    # Guarda una copia con el nombre nuevo, sin tocar el archivo original.
    destino = os.path.abspath(destino)
    ext_original = os.path.splitext(s.ruta)[1].lower()
    if os.path.splitext(destino)[1].lower() != ext_original:
        # SaveCopyAs no convierte de formato: con otra extension, Excel no podria abrir la copia.
        raise ExcelError('La copia tiene que llevar la misma extension que el original (' + ext_original + ')')
    carpeta = os.path.dirname(destino)
    if carpeta and not os.path.isdir(carpeta):
        os.makedirs(carpeta, exist_ok=True)
    if os.path.exists(destino):
        raise ExcelError('Ya existe un archivo con ese nombre: ' + destino)
    s.wb.SaveCopyAs(destino)
    return 'Copia guardada como ' + destino


def _filas_con_datos(h):
    'Cuenta las filas que tienen algun dato, leyendo el rango de una sola vez.'
    datos = h.UsedRange.Value2
    if not isinstance(datos, tuple):
        datos = ((datos,),)
    n = 0
    for fila in datos:
        if not isinstance(fila, tuple):
            fila = (fila,)
        if any(c is not None for c in fila):
            n += 1
    return n


def partir_filas(texto):
    'Divide el texto en filas y columnas por tabulaciones, punto y coma o comas.'
    filas = []
    for linea in texto.splitlines():
        if not linea.strip():
            continue
        if chr(9) in linea:
            partes = linea.split(chr(9))
        elif ';' in linea:
            partes = linea.split(';')
        elif ',' in linea:
            partes = linea.split(',')
        else:
            partes = [linea]
        filas.append([p.strip() for p in partes])
    return filas


def valor_celda(t):
    'Convierte a numero solo lo que es claramente numerico, respetando ceros a la izquierda.'
    s = t.strip()
    if not re.fullmatch('[+-]?[0-9]+([.,][0-9]+)?', s):
        return t
    if len(s.lstrip('+-')) > 1 and s.lstrip('+-')[0] == '0':
        return t
    return valor(s)


def leer_texto(ruta):
    'Lee un archivo de texto probando varias codificaciones.'
    for codif in ('utf-8-sig', 'utf-8', 'cp1252', 'latin-1'):
        try:
            with open(ruta, 'r', encoding=codif) as f:
                return f.read()
        except UnicodeDecodeError:
            continue
    with open(ruta, 'rb') as f:
        return f.read().decode('latin-1', 'replace')


def ruta_tesseract():
    'Busca el motor de OCR: variable de entorno, proyecto y rutas habituales.'
    candidatos = [os.environ.get('TESSERACT_CMD', '')]
    candidatos.append(os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'tesseract', 'tesseract.exe'))
    candidatos.append('C:/Program Files/Tesseract-OCR/tesseract.exe')
    candidatos.append('C:/Program Files (x86)/Tesseract-OCR/tesseract.exe')
    for c in candidatos:
        if c and os.path.isfile(c):
            return os.path.abspath(c)
    for carpeta in os.environ.get('PATH', '').split(os.pathsep):
        p = os.path.join(carpeta, 'tesseract.exe')
        if os.path.isfile(p):
            return p
    return None


def ocr_imagen(ruta):
    'Reconoce el texto de una imagen con pytesseract.'
    try:
        import pytesseract
        from PIL import Image
    except ImportError:
        raise ExcelError('Falta pytesseract o Pillow en el entorno virtual')
    exe = ruta_tesseract()
    if not exe:
        raise ExcelError('No encontre tesseract.exe. Instalalo o defini la variable TESSERACT_CMD')
    pytesseract.pytesseract.tesseract_cmd = exe
    idioma = 'eng'
    try:
        if 'spa' in pytesseract.get_languages(config=''):
            idioma = 'spa+eng'
    except Exception:
        pass
    with Image.open(ruta) as img:
        return pytesseract.image_to_string(img, lang=idioma)


def procesar_adjunto(ruta):
    """Lee un adjunto y devuelve su contenido utilizable. Los formatos estan en adjuntos.py; 'filas' es la primera
    parte con datos (la primera hoja, la primera tabla) y 'partes' tiene todas."""
    if not os.path.isfile(ruta):
        raise ExcelError('No existe el adjunto: ' + ruta)
    import adjuntos
    leido = adjuntos.leer(ruta)
    partes = leido['partes']
    for p in partes:
        p['columnas'] = max([len(f) for f in p['filas']] + [0])
    primera = next((p for p in partes if p['filas']), partes[0])
    texto = leido['texto']
    if texto.strip():
        resumen = 'Primeras lineas:' + NL + NL.join(texto.splitlines()[:8])
    else:
        resumen = 'No se reconocio texto en el adjunto.'
    return {'ruta': ruta, 'nombre': os.path.basename(ruta), 'clase': leido['clase'], 'texto': texto,
            'filas': primera['filas'], 'columnas': primera['columnas'], 'partes': partes, 'resumen': resumen}


def op_crear_tabla(s, ref=None):
    'Convierte un rango, o el area usada, en una tabla de Excel.'
    h = s.hoja()
    rango = h.Range(ref) if ref else h.UsedRange
    # Si pidieron una sola celda, la tabla abarca el bloque de datos con su encabezado.
    if rango.Rows.Count < 2:
        s.bus.log('En ' + str(rango.Address) + ' no hay encabezado ni datos: tomo el area usada ' + str(h.UsedRange.Address), 'warn')
        rango = h.UsedRange
    if rango.Rows.Count < 2:
        raise ExcelError('Hacen falta al menos dos filas: encabezado y datos')
    if h.ListObjects.Count > 0:
        raise ExcelError('La hoja ya tiene una tabla. Quitala antes de crear otra')
    tabla = h.ListObjects.Add(1, rango, True, 1)
    tabla.TableStyle = 'TableStyleMedium2'
    return 'Tabla creada en ' + str(rango.Address) + ' con ' + str(tabla.ListRows.Count) + ' filas de datos'


def op_minimizar_excel(s):
    s.minimizar()
    return 'Excel minimizado en la barra de tareas'


def op_mostrar_excel(s):
    s.mostrar()
    return 'Excel a la vista'


CONTACTOS_CAMPOS = ('FullName', 'Email1Address', 'BusinessTelephoneNumber', 'CompanyName', 'JobTitle')
CONTACTOS_ENCABEZADOS = ('Nombre', 'Correo', 'Telefono', 'Empresa', 'Cargo')


def op_contactos_outlook(s, con_tabla=False, todos=False):
    # trae los contactos de la carpeta predeterminada de Outlook a la hoja activa
    import win32com.client as win32
    try:
        outlook = win32.GetActiveObject('Outlook.Application')
    except Exception:
        try:
            outlook = win32.Dispatch('Outlook.Application')
        except Exception as exc:
            raise ExcelError('No pude iniciar Outlook: ' + str(exc))
    try:
        espacio = outlook.GetNamespace('MAPI')
        carpeta = espacio.GetDefaultFolder(10)
        items = carpeta.Items
    except Exception as exc:
        raise ExcelError('Outlook no tiene perfil configurado o no responde: ' + str(exc))
    filas = [list(CONTACTOS_ENCABEZADOS)]
    vistos = set()
    for item in items:
        try:
            if item.Class != 40:
                continue
            datos = []
            for campo in CONTACTOS_CAMPOS:
                try:
                    datos.append(str(getattr(item, campo) or ''))
                except Exception:
                    datos.append('')
            clave = datos[0] + '|' + datos[1]
            if clave in vistos:
                continue
            vistos.add(clave)
            filas.append(datos)
        except Exception:
            continue
    if len(filas) == 1:
        s.bus.log('La carpeta de Contactos esta vacia: busco en los correos.', 'warn')
        # Antes se pasaba con_tabla en el lugar de todos: pedir tabla recorria el buzon completo y no armaba tabla.
        mensaje = op_contactos_correo(s, todos)
        if con_tabla:
            mensaje = mensaje + NL + op_crear_tabla(s, 'A1')
        return mensaje
    h = s.hoja()
    ancho = len(CONTACTOS_ENCABEZADOS)
    destino = rango_desde('A1', len(filas), ancho)
    h.Range(destino).Value2 = tuple(tuple(f) for f in filas)
    mensaje = 'Contactos de Outlook volcados en ' + h.Name + ': ' + str(len(filas) - 1) + ' contactos'
    if con_tabla:
        mensaje = mensaje + NL + op_crear_tabla(s, 'A1')
    return mensaje
