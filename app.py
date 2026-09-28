# Asistente local para trabajar sobre libros de Excel abiertos.

from __future__ import annotations

import difflib
import os
import re
import threading
import time
import unicodedata

from dataclasses import dataclass

NL = chr(10)
PERCENT = chr(37)
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
            else:
                self.app = win32.DispatchEx('Excel.Application')
                self.propia = True
                origen = 'una instancia nueva de Excel'
            try:
                self.app.DisplayAlerts = False
                self.app.Visible = bool(visible)
            except Exception as exc:
                self.bus.log('La instancia adoptada no respondia: armo una nueva. ' + str(exc), 'warn')
                self.app = win32.DispatchEx('Excel.Application')
                self.propia = True
                self.app.DisplayAlerts = False
                self.app.Visible = bool(visible)
                origen = 'una instancia nueva de Excel'
        else:
            origen = 'la conexion ya abierta'
            self.app.Visible = bool(visible)
        self.bus.log('Conexion lista con ' + origen, 'ok')

    def _instancia_que_responde(self, win32):
        # Nunca se adhiere a un Excel trabado. La sonda corre en un hilo aparte,
        # pero el objeto se toma en este hilo: cruzar objetos COM entre hilos los invalida.
        def sonda():
            obj = win32.GetActiveObject('Excel.Application')
            obj.DisplayAlerts = False
            return int(obj.Workbooks.Count)
        try:
            con_limite(sonda, 6)
        except ExcelError as exc:
            self.bus.log('No pude usar el Excel abierto: ' + str(exc), 'warn')
            return None, 'El Excel abierto no se puede usar: trabajo con una instancia nueva. Destrabalo y apreta Reconectar.'
        except Exception:
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
        self.wb = self._buscar_abierto(ruta)
        if self.wb is None:
            self.wb = self._abrir_libro(ruta, visible)
            self.abierto_por_nosotros = True
        if self.wb is None:
            self.wb = self._buscar_abierto(ruta)
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
            self.app = win32.DispatchEx('Excel.Application')
            self.app.DisplayAlerts = False
            self.app.Visible = bool(visible)
            self.propia = True
            self.bus.log('Conexion lista con una instancia nueva de Excel', 'ok')
            return self.app.Workbooks.Open(ruta)


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

    def salir(self, cerrar_excel=False):
        self.cerrar(False)
        if cerrar_excel and self.app is not None:
            if self.propia:
                self.app.Quit()
            self.app = None
        self.bus.log('Sesion finalizada', 'ok')



# Este Excel esta en espanol: los formatos se aplican con NumberFormatLocal
# usando punto para los miles y coma para los decimales.
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

FORMATOS = {
      'moneda': '$#.##0,00'
    , 'numero': '#.##0,00'
    , 'entero': '#.##0'
    , 'porcentaje': '0,00' + PERCENT
    , 'fecha': 'dd/mm/aaaa'
    , 'texto': '@'
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
        detalle = detalle + NL + 'Para revisar todo el buzon, agrega la palabra todos a la consigna (tarda varios minutos).'
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

def detectar_guardado(t):
    # Reconoce pedidos de guardar una copia, renombrar o cerrar el proceso.
    bajo = sin_acentos(t).lower()
    raices = ('guardar', 'guarda', 'salvar', 'salva', 'renombrar', 'renombra')
    sufijos = ('la', 'lo', 'las', 'los', str())
    quiere_guardar = bool(re.search('|'.join(r + s for r in raices for s in sufijos), bajo))
    quiere_cerrar = bool(re.search('cerrar|cierra|cerra|cerralo|cerrala|liberar|libera|liberalo|soltar|suelta|soltalo|soltala', bajo))
    if not (quiere_guardar or quiere_cerrar):
        return None
    pasos = []
    partes = []
    destino = None
    m = re.search('como (.+?)(?: en el | en la | en |$)', bajo)
    if m:
        nombre = m.group(1).strip().strip(chr(34)).strip(chr(39)).strip()
        if nombre:
            if not nombre.endswith(('.xlsx', '.xlsm', '.xls', '.xlsb')):
                nombre = nombre + '.xlsx'
            carpeta = carpeta_conocida(bajo)
            if carpeta:
                destino = os.path.join(carpeta, nombre)
    if quiere_guardar:
        if destino:
            pasos.append(lambda s: op_guardar_como(s, destino))
            partes.append('Guardar una copia como ' + destino)
        else:
            pasos.append(op_guardar)
            partes.append('Guardar el libro')
    if quiere_cerrar:
        pasos.append(op_cerrar_excel)
        partes.append('Cerrar el proceso de Excel')
    if not pasos:
        return None
    return ' y '.join(partes), pasos


def op_guardar_como(s, destino):
    # Guarda una copia con el nombre nuevo, sin tocar el archivo original.
    destino = os.path.abspath(destino)
    if not destino.lower().endswith(('.xlsx', '.xlsm', '.xls', '.xlsb')):
        destino = destino + '.xlsx'
    carpeta = os.path.dirname(destino)
    if carpeta and not os.path.isdir(carpeta):
        os.makedirs(carpeta, exist_ok=True)
    if os.path.exists(destino):
        raise ExcelError('Ya existe un archivo con ese nombre: ' + destino)
    s.wb.SaveCopyAs(destino)
    return 'Copia guardada como ' + destino


def op_cerrar_excel(s):
    # Deja el libro y el Excel como pide la consigna.
    if s.wb is None:
        return 'No hay ningun libro abierto'
    nombre = str(s.wb.Name)
    partes = []
    if s.abierto_por_nosotros:
        try:
            if not s.wb.Saved:
                s.wb.Save()
                partes.append('Guardados los cambios de ' + nombre)
        except Exception as exc:
            partes.append('no pude guardar antes de cerrar: ' + str(exc))
        partes.append('El asistente solto el libro ' + nombre)
    else:
        partes.append('El libro ' + nombre + ' ya estaba abierto antes: lo dejo como estaba')
    s.cerrar(False)
    if s.app is not None and s.propia:
        try:
            s.app.Quit()
            s.app = None
            partes.append('Excel quedo terminado')
        except Exception as exc:
            partes.append('no pude terminar Excel: ' + str(exc))
    return '. '.join(partes) + '. Para seguir, elegi un libro con Abrir Excel.'


def sin_acentos(texto):
    # Deja las letras sin tilde, para comparar palabras sin depender de los acentos.
    base = unicodedata.normalize('NFKD', str(texto or str()))
    return str().join(c for c in base if not unicodedata.combining(c))


def detectar_contactos(t):
    # Reconoce pedidos de contactos escritos como se habla, sin forma de comando.
    bajo = sin_acentos(t).lower()
    if not re.search('contacto|agenda|directorio', bajo):
        return None    # Si piden guardar, renombrar o cerrar, no es una extraccion nueva.
    raices = ('guardar', 'guarda', 'salvar', 'salva', 'renombrar', 'renombra', 'cerrar', 'cierra')
    sufijos = ('la', 'lo', 'las', 'los', str())
    patron_guardar = '|'.join(r + s for r in raices for s in sufijos)
    pide_guardar = bool(re.search(patron_guardar, bajo))
    marca_extraccion = 'tabla|listado|elabora|recopila|extrae|agrega'
    con_armado = bool(re.search('listado|tabla|planilla|elabora|confecciona|prepara|arma|crea|genera|trae|hace|agrega', bajo))
    if pide_guardar and not re.search(marca_extraccion, bajo):
        return None
    if not con_armado:
        return None
    todos = bool(re.search('todo el buzon|buzon completo|todos los correos|todos los mensajes|sin limite', bajo))
    con_tabla = bool(re.search('tabla', bajo))
    con_orden = bool(re.search('numero de orden|numeracion|orden numerico|columna de orden|nro de orden|con orden', bajo))
    pasos = [lambda s: op_contactos_correo(s, todos, con_orden)]
    if con_tabla:
        pasos.append(lambda s: op_crear_tabla(s))
    titulo = 'Traer los contactos desde los correos de Outlook'
    if todos:
        titulo = titulo + ' (buzon completo)'
    if con_orden:
        titulo = titulo + ' con numero de orden'
    if con_tabla:
        titulo = titulo + ' como tabla'
    return titulo, pasos


def ayuda():
    return NL.join([
          'Consignas que entiendo:'
        , '  hojas | resumen | leer A1:C10'
        , '  escribir 1500 en C2'
        , '  formula =B2*C2 en D2'
        , '  negrita en A1:C1 | sin negrita en A1:C1'
        , '  color rojo en A1:A9 | fondo amarillo en A1:C1'
        , '  formato moneda en C2:C50 (o numero, entero, porcentaje, fecha, texto)'
        , '  sumar columna C | promedio columna C | maximo columna C | contar columna C'
        , '  crear hoja Ventas | activar hoja Ventas | renombrar hoja A a B'
        , '  ordenar por columna B descendente'
        , '  reemplazar IVA por IVA21 | buscar IVA'
        , '  quitar duplicados | autofiltro | quitar filtro | autoajustar'
        , '  crear tabla | crear tabla en A1:C20'
        , '  contactos desde los correos (bandeja de entrada y enviados)'
        , '  ... con la palabra todos recorre el buzon completo'

        , '  adjuntos | analizar adjunto | volcar adjunto en A1'
        , '  minimizar excel | mostrar excel'
        , '  exportar csv D:/salida.csv'
        , '  ejecutar macro MiMacro'
        , '  guardar | guardar como D:/copia.xlsx'
        , ''
        , 'Podes escribir la frase como te salga: saco preambulos y entiendo los verbos en imperativo.'
    ])



def op_escribir(s, ref, dato):
    celda = s.hoja().Range(ref)
    celda.Value2 = dato
    return ref + ' = ' + str(celda.Value2)


def op_formula(s, ref, expr):
    if not expr.startswith('='):
        expr = '=' + expr
    celda = s.hoja().Range(ref)
    celda.Formula = expr
    return ref + ' = ' + expr + '  ->  ' + str(celda.Value2)


def op_formato(s, ref, negrita=None, color=None, fondo=None, num=None):
    r = s.hoja().Range(ref)
    if negrita is not None:
        r.Font.Bold = negrita
    if color:
        r.Font.Color = color_bgr(color)
    if fondo:
        r.Interior.Color = color_bgr(fondo)
    if num:
        r.NumberFormatLocal = num
    return 'Formato aplicado en ' + ref



def op_resumen(s, col, func='SUM', etiqueta='TOTAL'):
    h = s.hoja()
    n = col_num(col)
    ultima = h.Cells.Item(h.Rows.Count, n).End(-4162).Row
    if ultima < 2:
        ultima = 1
    fila = ultima + 1
    celda = h.Cells.Item(fila, n)
    celda.Formula = '=' + func + '(' + col + '2:' + col + str(ultima) + ')'
    if n != 1:
        h.Cells.Item(fila, 1).Value2 = etiqueta
    return etiqueta + ' en ' + col + str(fila) + ' = ' + str(celda.Value2)


def op_hoja(s, accion, nombre, nuevo=None):
    if accion == 'nueva':
        if nombre in s.hojas():
            s.wb.Worksheets(nombre).Activate()
            return 'La hoja ' + nombre + ' ya existia: se activo'
        ws = s.wb.Worksheets.Add(After=s.wb.Worksheets(s.wb.Worksheets.Count))
        ws.Name = nombre
        return 'Hoja creada: ' + nombre
    if accion == 'activar':
        s.hoja(nombre).Activate()
        return 'Hoja activa: ' + nombre
    s.wb.Worksheets(nombre).Name = nuevo
    return 'La hoja ' + nombre + ' ahora se llama ' + nuevo



def op_ordenar(s, col, asc=True):
    h = s.hoja()
    usado = h.UsedRange
    if usado.Rows.Count < 2:
        raise ExcelError('No hay filas suficientes para ordenar')
    clave = h.Cells.Item(usado.Row + 1, col_num(col))
    h.Sort.SortFields.Clear()
    h.Sort.SortFields.Add(Key=clave, SortOn=0, Order=1 if asc else 2)
    h.Sort.SetRange(usado)
    h.Sort.Header = 1
    h.Sort.Apply()
    h.Sort.SortFields.Clear()
    return 'Ordenado por la columna ' + col


def op_reemplazar(s, busca, pon):
    rango = s.hoja().UsedRange
    total = 0
    celda = rango.Find(busca, LookIn=-4163, LookAt=2)
    if celda is not None:
        primera = celda.Address
        while celda is not None and total < 5000:
            total += 1
            celda = rango.FindNext(celda)
            if celda is not None and celda.Address == primera:
                break
    if total == 0:
        return 'No se encontro ' + busca
    rango.Replace(What=busca, Replacement=pon, LookAt=2)
    return str(total) + ' coincidencia(s) reemplazadas'



def op_buscar(s, texto):
    r = s.hoja().UsedRange
    c = r.Find(texto, LookIn=-4163, LookAt=2)
    if c is None:
        return 'No se encontro ' + texto
    sitios = []
    primera = c.Address
    while c is not None and len(sitios) < 30:
        sitios.append(c.Address)
        c = r.FindNext(c)
        if c is not None and c.Address == primera:
            break
    return texto + ' aparece en: ' + ', '.join(sitios)


def op_leer(s, ref):
    datos = s.hoja().Range(ref).Value2
    if not isinstance(datos, tuple):
        datos = ((datos,),)
    lineas = []
    for fila in datos:
        if not isinstance(fila, tuple):
            fila = (fila,)
        lineas.append(' | '.join('' if v is None else str(v) for v in fila))
    return ref + ':' + NL + NL.join(lineas)



def op_describir(s):
    h = s.hoja()
    u = h.UsedRange
    cols = min(u.Columns.Count, 12)
    enc = [str(h.Cells.Item(u.Row, c).Value2) for c in range(1, cols + 1)]
    cab = 'Hoja ' + h.Name + ': ' + u.Address
    cab = cab + '  (' + str(u.Rows.Count) + ' filas x ' + str(u.Columns.Count) + ' columnas)'
    return cab + NL + 'Encabezados: ' + ' | '.join(enc)


def op_hojas(s):
    return 'Hojas: ' + ', '.join(s.hojas())


def op_macro(s, nombre):
    return 'Macro ' + nombre + ' devolvio: ' + str(s.app.Run(nombre))



def op_csv(s, destino=None):
    h = s.hoja()
    if not destino:
        destino = os.path.splitext(s.ruta)[0] + '_' + h.Name + '.csv'
    destino = os.path.abspath(destino)
    import csv
    valores = h.UsedRange.Value2
    if not isinstance(valores, tuple):
        valores = ((valores,),)
    with open(destino, 'w', encoding='utf-8-sig', newline='') as fh:
        esc = csv.writer(fh, delimiter=';')
        for fila in valores:
            if not isinstance(fila, tuple):
                fila = (fila,)
            esc.writerow(['' if v is None else v for v in fila])
    return 'Exportado a ' + destino


def op_guardar(s, destino=None):
    if destino:
        s.ruta = s.wb.SaveAs(os.path.abspath(destino))
        return 'Guardado como ' + s.ruta
    s.wb.Save()
    return 'Guardado: ' + s.ruta

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

def op_quitar_duplicados(s):
    h = s.hoja()
    antes = _filas_con_datos(h)
    cols = h.UsedRange.Columns.Count
    # RemoveDuplicates falla si la ventana de Excel esta minimizada:
    # se oculta y se pasa a normal un instante, sin que se vea en pantalla.
    estado = s.app.WindowState
    visible = s.app.Visible
    try:
        if estado == XL_MINIMIZADO:
            s.app.Visible = False
            s.app.WindowState = XL_NORMAL
        h.UsedRange.RemoveDuplicates(tuple(range(1, cols + 1)), 1)
    finally:
        if estado == XL_MINIMIZADO:
            s.app.WindowState = estado
            s.app.Visible = visible
    despues = _filas_con_datos(h)
    return 'Duplicados quitados: filas con datos ' + str(antes) + ' -> ' + str(despues)


def op_filtro(s, activar=True):
    h = s.hoja()
    if activar:
        h.UsedRange.AutoFilter()
        return 'Autofiltro activado'
    h.AutoFilterMode = False
    return 'Autofiltro quitado'


def op_ancho(s):
    s.hoja().Columns.AutoFit()
    return 'Ancho de columnas ajustado'


RE_RANGO = '([A-Za-z]{1,3}[0-9]{1,7}(:[A-Za-z]{1,3}[0-9]{1,7})?)'
RE_RANGO_LIBRE = '[A-Za-z]{1,3}[0-9]{1,7}(?::[A-Za-z]{1,3}[0-9]{1,7})?'

VERBOS = {
      'escribi': 'escribir'
    , 'escribe': 'escribir'
    , 'anota': 'anotar'
    , 'pon': 'poner'
    , 'pone': 'poner'
    , 'crea': 'crear'
    , 'arma': 'crear'
    , 'genera': 'crear'
    , 'hace': 'crear'
    , 'agrega': 'agregar'
    , 'busca': 'buscar'
    , 'reemplaza': 'reemplazar'
    , 'ordena': 'ordenar'
    , 'guarda': 'guardar'
    , 'exporta': 'exportar'
    , 'suma': 'sumar'
    , 'sumame': 'sumar'
    , 'sumale': 'sumar'
    , 'totaliza': 'sumar'
    , 'minimiza': 'minimizar'
    , 'muestra': 'mostrar'
    , 'mostra': 'mostrar'
    , 'restaura': 'mostrar'
    , 'activa': 'activar'
    , 'renombra': 'renombrar'
    , 'volca': 'volcar'
    , 'analiza': 'analizar'
    , 'quita': 'quitar'
    , 'saca': 'quitar'
    , 'elimina': 'quitar'
    , 'lista': 'listar'
    , 'aplica': 'aplicar'
}


def quitar_marcadores(texto):
    # Saca los numeros y guiones de las listas: uno guion, dos parentesis, guion suelto.
    patron = re.compile(r'^\s*(?:\(?\d{1,2}\)?[.\-:)]|[-*])\s+')
    return chr(10).join(patron.sub(str(), linea) for linea in texto.splitlines())


def normalizar(texto):
    # Limpia marcadores y preambulos, y unifica los verbos en imperativo.
    t = quitar_marcadores(texto).strip()
    bajas = t.lower()
    preambulos = ('quiero que ', 'necesito que ', 'podes ', 'podrias ', 'en la hoja designada ', 'en la hoja activa ', 'en la hoja ', 'en el libro ', 'en la planilla ', 'por favor ', 'porfa ', 'che ', 'dale ')
    cambio = True
    while cambio:
        cambio = False
        for pre in preambulos:
            if bajas.startswith(pre):
                t = t[len(pre):].strip()
                bajas = t.lower()
                cambio = True
    for forma, canonico in VERBOS.items():
        t = re.sub('(?<![A-Za-z])' + forma + '(?![A-Za-z])', canonico, t, count=1, flags=re.IGNORECASE)
    return t


def extras_del_texto(texto):
    # Junta los pedidos de guardar o terminar que vengan junto a otra tarea.
    partes = []
    pasos = []
    for linea in str(texto or str()).splitlines():
        linea = linea.strip()
        if not linea:
            continue
        pedido = detectar_guardado(linea)
        if pedido is None:
            continue
        resumen, trabajos = pedido
        partes.append(resumen)
        pasos.extend(trabajos)
    if not pasos:
        return None
    return ' y '.join(partes), pasos


def interpretar(texto):
    t = normalizar(texto)
    if re.match('^(ayuda|help|comandos)$', t, re.IGNORECASE):
        return ayuda(), []
    if re.match('^(hojas|listar hojas|ver hojas)$', t, re.IGNORECASE):
        return 'Listar las hojas', [op_hojas]
    if re.match('^(resumen|describir|analizar|analizar hoja)$', t, re.IGNORECASE):
        return 'Resumen de la hoja activa', [op_describir]
    detector = detectar_contactos(t)
    if detector is not None:
        extra = extras_del_texto(texto)
        if extra is None:
            return detector
        resumen, pasos = detector
        resumen_extra, pasos_extra = extra
        return resumen + ' y ' + resumen_extra, pasos + pasos_extra

    m = re.match('^leer ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        ref = m.group(1).upper()
        return 'Leer ' + ref, [lambda s: op_leer(s, ref)]
    m = re.match('^(escribir|poner|anotar|cargar|colocar) (.+?) en ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        ref = m.group(3).upper()
        dato = valor(m.group(2))
        return 'Escribir en ' + ref, [lambda s: op_escribir(s, ref, dato)]
    m = re.match('^(formula|poner formula|poner la formula) (.+?) en ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        expr = m.group(2)
        ref = m.group(3).upper()
        return 'Formula en ' + ref, [lambda s: op_formula(s, ref, expr)]

    m = re.match('^(aplicar |poner |pon )?(negrita|sin negrita) en ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        on = m.group(2).lower() == 'negrita'
        ref = m.group(3).upper()
        return 'Negrita en ' + ref, [lambda s: op_formato(s, ref, negrita=on)]
    m = re.match('^(aplicar |poner |pon )?color ([A-Za-z]+) en ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        c = COLORES.get(m.group(2).lower())
        ref = m.group(3).upper()
        if not c:
            return 'No conozco el color ' + m.group(2), []
        return 'Color de letra en ' + ref, [lambda s: op_formato(s, ref, color=c)]
    m = re.match('^(aplicar |poner |pon )?fondo ([A-Za-z]+) en ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        c = COLORES.get(m.group(2).lower())
        ref = m.group(3).upper()
        if not c:
            return 'No conozco el color ' + m.group(2), []
        return 'Fondo en ' + ref, [lambda s: op_formato(s, ref, fondo=c)]

    m = re.match('^(aplicar |poner |pon )?formato ([A-Za-z]+) en ' + RE_RANGO + '$', t, re.IGNORECASE)
    if m:
        f = FORMATOS.get(m.group(2).lower())
        ref = m.group(3).upper()
        if not f:
            return 'No conozco el formato ' + m.group(2), []
        return 'Formato en ' + ref, [lambda s: op_formato(s, ref, num=f)]
    m = re.match('^(sumar|totalizar|total) (la )?columna ([A-Za-z]{1,3})$', t, re.IGNORECASE)
    if m:
        col = m.group(3).upper()
        return 'Total de la columna ' + col, [lambda s: op_resumen(s, col, 'SUM', 'TOTAL')]
    m = re.match('^(promedio|maximo|max|minimo|min|contar) (la )?columna ([A-Za-z]{1,3})$', t, re.IGNORECASE)
    if m:
        pares = {'promedio': 'AVERAGE', 'maximo': 'MAX', 'max': 'MAX', 'minimo': 'MIN', 'min': 'MIN', 'contar': 'COUNT'}
        cual = m.group(1).lower()
        col = m.group(3).upper()
        func = pares[cual]
        return cual.upper() + ' de la columna ' + col, [lambda s: op_resumen(s, col, func, cual.upper())]

    m = re.match('^(crear|agregar|nueva|nuevo|anadir) (una )?hoja (.+)$', t, re.IGNORECASE)
    if m:
        nombre = m.group(3).strip()
        return 'Crear la hoja ' + nombre, [lambda s: op_hoja(s, 'nueva', nombre)]
    m = re.match('^(activar|ir a|seleccionar) (la )?hoja (.+)$', t, re.IGNORECASE)
    if m:
        nombre = m.group(3).strip()
        return 'Activar la hoja ' + nombre, [lambda s: op_hoja(s, 'activar', nombre)]
    m = re.match('^renombrar (la )?hoja (.+?) a (.+)$', t, re.IGNORECASE)
    if m:
        nombre = m.group(2).strip()
        nuevo = m.group(3).strip()
        return 'Renombrar la hoja ' + nombre, [lambda s: op_hoja(s, 'renombrar', nombre, nuevo)]
    m = re.match('^ordenar por (la )?columna ([A-Za-z]{1,3})( (ascendente|descendente|asc|desc))?$', t, re.IGNORECASE)
    if m:
        col = m.group(2).upper()
        sentido = (m.group(4) or '').lower()
        asc = not sentido.startswith('desc')
        return 'Ordenar por la columna ' + col, [lambda s: op_ordenar(s, col, asc)]
    m = re.match('^reemplazar (.+?) por (.+)$', t, re.IGNORECASE)
    if m:
        busca = m.group(1)
        pon = m.group(2)
        return 'Reemplazar ' + busca, [lambda s: op_reemplazar(s, busca, pon)]
    m = re.match('^(buscar|encontrar) (.+)$', t, re.IGNORECASE)
    if m:
        busca = m.group(2)
        return 'Buscar ' + busca, [lambda s: op_buscar(s, busca)]

    if re.match('^(adjuntos|ver adjuntos|listar adjuntos)$', t, re.IGNORECASE):
        return 'Listar los adjuntos cargados', [op_listar_adjuntos]
    if re.match('^(analizar|resumir) (el )?adjunto$', t, re.IGNORECASE):
        return 'Analizar el ultimo adjunto', [op_analizar_adjunto]
    m = re.match('^(volcar|pegar|insertar|escribir) (el )?adjunto( en ([A-Za-z]{1,3}[0-9]{1,7}))?$', t, re.IGNORECASE)
    if m:
        ref = (m.group(4) or 'A1').upper()
        return 'Volcar el adjunto en ' + ref, [lambda s: op_volcar_adjunto(s, ref)]
    if re.search('outlook', t, re.IGNORECASE) and re.search('contacto|agenda|directorio', t, re.IGNORECASE):
        con_tabla = bool(re.search('tabla', t, re.IGNORECASE))
        return 'Traer los contactos de Outlook' + (' como tabla' if con_tabla else ''), [lambda s: op_contactos_outlook(s, con_tabla)]

    if re.match('^(quitar|sacar|eliminar) (los |las )?duplicados$', t, re.IGNORECASE):
        return 'Quitar filas duplicadas', [op_quitar_duplicados]
    if re.match('^(autofiltro|poner filtro|activar filtro)$', t, re.IGNORECASE):
        return 'Activar el autofiltro', [lambda s: op_filtro(s, True)]
    if re.match('^(quitar filtro|sacar filtro)$', t, re.IGNORECASE):
        return 'Quitar el autofiltro', [lambda s: op_filtro(s, False)]
    if re.match('^(autoajustar|ajustar ancho|acomodar columnas)$', t, re.IGNORECASE):
        return 'Ajustar el ancho de columnas', [op_ancho]
    m = re.match('^exportar csv( (.+))?$', t, re.IGNORECASE)
    if m:
        destino = (m.group(2) or '').strip() or None
        return 'Exportar la hoja a CSV', [lambda s: op_csv(s, destino)]

    if re.search('(crear|armar|generar|hacer) (una )?tabla', t, re.IGNORECASE) and not re.search('contacto', t, re.IGNORECASE):
        refs = re.findall(RE_RANGO_LIBRE, t)
        ref = refs[0].upper() if refs else ''
        return 'Crear una tabla' + (' en ' + ref if ref else ''), [lambda s: op_crear_tabla(s, ref)]
    if re.match('^(minimizar|ocultar) excel$', t, re.IGNORECASE):
        return 'Minimizar la ventana de Excel', [op_minimizar_excel]
    if re.match('^(mostrar|restaurar|ver) excel$', t, re.IGNORECASE):
        return 'Mostrar la ventana de Excel', [op_mostrar_excel]
    m = re.match('^ejecutar macro ([A-Za-z0-9_.]+)$', t, re.IGNORECASE)
    if m:
        nombre = m.group(1)
        return 'Ejecutar la macro ' + nombre, [lambda s: op_macro(s, nombre)]
    guardado = detectar_guardado(t)
    if guardado is not None:
        return guardado
    return no_entendido(t), []



class Agente:
    'Orquesta: interpreta la consigna, ejecuta y responde.'

    def __init__(self, bus):
        self.bus = bus
        self.sesion = Sesion(bus)

    def abrir(self, ruta, visible=True, respaldo=True, minimizar=True):
        self.bus.progreso(10, 'Abriendo el libro')
        hojas = self.sesion.abrir(ruta, visible, respaldo, minimizar)
        self.bus.progreso(100, 'Libro listo')
        return hojas

    def adjuntar(self, ruta):
        self.bus.log('Adjuntando: ' + ruta)
        self.bus.progreso(25, 'Leyendo el adjunto')
        info = procesar_adjunto(ruta)
        self.sesion.adjuntos.append(info)
        self.bus.progreso(100, 'Adjunto listo')
        detalle = info['nombre'] + ' (' + info['clase'] + ', ' + str(len(info['texto'])) + ' caracteres, ' + str(len(info['filas'])) + ' filas)'
        self.bus.log('Adjunto leido: ' + detalle, 'ok')
        self.bus.chat('Adjunto ' + detalle + NL + info['resumen'])
        return info


    def pedir(self, texto):
        self.sesion.detener.clear()
        texto = (texto or str()).strip()
        if not texto:
            return self.responde('No escribiste ninguna consigna.')
        if self.sesion.wb is None:
            return self.responde('Primero elegi un archivo con el boton Abrir Excel')
        self.bus.chat(texto, 'usuario')
        if self.sesion.adjuntos:
            self.bus.log('Adjuntos cargados: ' + str(len(self.sesion.adjuntos)))
        lineas = [l.strip() for l in texto.splitlines() if l.strip()]
        if len(lineas) > 1:
            resumen, trabajos = interpretar(texto)
            if not trabajos:
                return self._pedir_varias(lineas)
            self.bus.log("El mensaje se lee como una sola tarea de " + str(len(trabajos)) + " paso/s")
            self.bus.chat("Entendi: " + resumen)
        else:
            resumen, trabajos = interpretar(texto)
            if trabajos:
                self.bus.chat("Entendi: " + resumen)
        self.bus.progreso(3, 'Interpretando la consigna')
        respuesta = self._ejecutar(texto, 5.0, 90.0)
        self.bus.progreso(100, 'Trabajo terminado')
        return self.responde(respuesta)


    def _pedir_varias(self, lineas):
        'Ejecuta una consigna por linea, en orden.'
        self.bus.log('Consignas en el mensaje: ' + str(len(lineas)))
        salida = []
        total = float(len(lineas))
        for i, linea in enumerate(lineas, start=1):
            if self.sesion.detener.is_set():
                salida.append('Se corto el trabajo a pedido.'),
                break
            self.bus.log('Consigna ' + str(i) + ' de ' + str(len(lineas)) + ': ' + linea)
            base = 100.0 * (i - 1) / total
            salida.append(self._ejecutar(linea, base, 100.0 / total))
        self.bus.progreso(100, 'Trabajo terminado')
        return self.responde(NL.join(salida))

    def _ejecutar(self, texto, base=0.0, ancho=100.0):
        'Interpreta y ejecuta una consigna. Nunca lanza: siempre devuelve texto.'
        try:
            resumen, trabajos = interpretar(texto)
        except Exception as exc:
            self.bus.log('No pude interpretar la consigna: ' + str(exc), 'error')
            return 'No pude interpretar esa consigna: ' + str(exc)
        if not trabajos:
            return resumen
        self.bus.log('Plan: ' + resumen + ' (' + str(len(trabajos)) + ' paso/s)')
        salida = []
        for i, trabajo in enumerate(trabajos, start=1):
            if self.sesion.detener.is_set():
                self.bus.log('Se corto el trabajo a pedido.', 'warn'),
                salida.append('Se corto el trabajo a pedido.'),
                break
            self.bus.log('Paso ' + str(i) + ': ' + resumen)
            self.bus.progreso(base + ancho * i / len(trabajos), resumen)
            try:
                salida.append(str(trabajo(self.sesion)))
            except Exception as exc:
                self.bus.log('Fallo el paso ' + str(i) + ': ' + str(exc), 'error')
                salida.append('ERROR: ' + str(exc))
        return NL.join(salida)

    def responde(self, texto):
        self.bus.chat(texto)
        return texto

    def detener(self):
        # El boton Detener deja anotado el pedido; se aplica en el proximo control.
        self.sesion.detener.set()
        self.bus.log('Pedido de corte anotado.', 'warn')


    def salir(self, cerrar_excel=False):
        self.sesion.salir(cerrar_excel)



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
    'Lee un adjunto (texto o imagen) y devuelve su contenido utilizable.'
    if not os.path.isfile(ruta):
        raise ExcelError('No existe el adjunto: ' + ruta)
    ext = os.path.splitext(ruta)[1].lower()
    imagenes = ('.png', '.jpg', '.jpeg', '.bmp', '.tif', '.tiff', '.webp', '.gif')
    clase = 'imagen' if ext in imagenes else 'texto'
    texto = ocr_imagen(ruta) if clase == 'imagen' else leer_texto(ruta)
    filas = partir_filas(texto)
    columnas = 0
    for f in filas:
        if len(f) > columnas:
            columnas = len(f)
    if texto.strip():
        resumen = 'Primeras lineas:' + NL + NL.join(texto.splitlines()[:8])
    else:
        resumen = 'No se reconocio texto en el adjunto.'
    return {'ruta': ruta, 'nombre': os.path.basename(ruta), 'clase': clase, 'texto': texto, 'filas': filas, 'columnas': columnas, 'resumen': resumen}



def op_listar_adjuntos(s):
    if not s.adjuntos:
        return 'No hay adjuntos cargados. Usa el boton + para agregar un archivo de texto o una imagen.'
    partes = []
    for i, a in enumerate(s.adjuntos, start=1):
        partes.append(str(i) + '. ' + a['nombre'] + ' (' + a['clase'] + ', ' + str(len(a['texto'])) + ' caracteres, ' + str(len(a['filas'])) + ' filas x ' + str(a['columnas']) + ' columnas)')
    return 'Adjuntos cargados:' + NL + NL.join(partes)


def op_analizar_adjunto(s):
    if not s.adjuntos:
        raise ExcelError('No hay adjuntos cargados: usa el boton + para agregar uno')
    a = s.adjuntos[-1]
    lineas = len(a['texto'].splitlines())
    cab = a['nombre'] + ' (' + a['clase'] + ')'
    detalle = 'Caracteres: ' + str(len(a['texto'])) + ' | Lineas: ' + str(lineas)
    detalle = detalle + ' | Filas detectadas: ' + str(len(a['filas'])) + ' | Columnas: ' + str(a['columnas'])
    return cab + NL + detalle + NL + a['resumen']



def op_volcar_adjunto(s, ref='A1'):
    # vuelca el adjunto cargado en la hoja activa
    if not s.adjuntos:
        raise ExcelError('No hay adjuntos cargados: usa el boton + para agregar uno')
    a = s.adjuntos[-1]
    if not a['filas']:
        raise ExcelError('El adjunto no tiene texto utilizable')
    ancho = a['columnas']
    datos = []
    for fila in a['filas']:
        completa = [valor_celda(v) for v in fila]
        completa = completa + [''] * (ancho - len(completa))
        datos.append(tuple(completa))
    destino = rango_desde(ref, len(datos), ancho)
    s.hoja().Range(destino).Value2 = tuple(datos)
    return 'Adjunto ' + a['nombre'] + ' volcado en ' + ref + ': ' + str(len(datos)) + ' filas x ' + str(ancho) + ' columnas'


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


def op_contactos_outlook(s, con_tabla=False):
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
        return op_contactos_correo(s, con_tabla)
    h = s.hoja()
    ancho = len(CONTACTOS_ENCABEZADOS)
    destino = rango_desde('A1', len(filas), ancho)
    h.Range(destino).Value2 = tuple(tuple(f) for f in filas)
    mensaje = 'Contactos de Outlook volcados en ' + h.Name + ': ' + str(len(filas) - 1) + ' contactos'
    if con_tabla:
        mensaje = mensaje + NL + op_crear_tabla(s, 'A1')
    return mensaje


CLAVES_AYUDA = ('hojas', 'resumen', 'leer A1', 'escribir', 'formula', 'negrita', 'color', 'fondo', 'formato', 'sumar columna', 'promedio columna', 'crear hoja', 'activar hoja', 'renombrar hoja', 'ordenar por columna', 'reemplazar', 'buscar', 'quitar duplicados', 'autofiltro', 'autoajustar', 'crear tabla', 'contactos de outlook', 'adjuntos', 'volcar adjunto', 'minimizar excel', 'exportar csv', 'ejecutar macro', 'guardar')


def no_entendido(texto):
    'Mensaje de fallback, con sugerencias parecidas a lo que se escribio.'
    palabras = [p for p in re.findall('[a-z]+', texto.lower()) if len(p) > 2]
    parecidos = []
    for clave in CLAVES_AYUDA:
        if any(p in clave for p in palabras):
            parecidos.append(clave)
    if not parecidos:
        parecidos = difflib.get_close_matches(texto.lower(), list(CLAVES_AYUDA), n=3, cutoff=0.35)
    mensaje = 'No entendi esa consigna.'
    if parecidos:
        mensaje = mensaje + ' Quizas quisiste decir: ' + ' | '.join(parecidos[:4]) + '.'
    return mensaje + ' Escribi ayuda para ver la lista completa.'

