# Contactos que el usuario trae de afuera: la agenda del telefono exportada como vCard (.vcf) o como CSV (Google
# Contactos, Outlook). Es la via para "los contactos de WhatsApp": WhatsApp no tiene agenda propia, usa la del
# telefono, y leer WhatsApp Web automatizando el navegador va contra sus condiciones de uso (riesgo de que suspendan
# la cuenta). Sin COM: se prueba sin Excel.
import csv
import io
import quopri
import re

ENCABEZADOS = ('Nombre', 'Telefono', 'Otros telefonos', 'Correo')

# Encabezados de columna que indican telefonos: sus valores se escriben en Excel como texto.
_COLUMNA_TELEFONO = re.compile(r'\b(tel[eé]fonos?|tel|phone|celular|cel|m[oó]vil|mobile|whats ?app|fono)\b', re.I)


def es_columna_telefono(encabezado):
    return bool(_COLUMNA_TELEFONO.search(str(encabezado or '')))


def parece_telefono(v):
    """Un telefono escrito como numero pierde el '+', los ceros de adelante o se ve como 5,49116E+12; si empieza con
    '+' y tiene espacios, Excel lo toma por formula. Parece telefono: 7 o mas digitos con '+' adelante, o con
    separadores de telefono (espacio, guion, parentesis) entre los digitos."""
    s = str(v).strip()
    if not re.fullmatch(r'\+?[0-9()][0-9 ()./-]*[0-9]', s):
        return False
    digitos = len(re.sub(r'[^0-9]', '', s))
    if digitos < 7:
        return False
    return s.startswith('+') or bool(re.search(r'[0-9][ ()-]+[0-9]', s))


# ------------------------------------------------------------ vCard

def _desplegar(texto):
    'Une las lineas plegadas del vCard (las que siguen empiezan con espacio) y las de quoted-printable (terminan en =).'
    lineas = []
    for linea in texto.splitlines():
        if lineas and linea[:1] in (' ', chr(9)):
            lineas[-1] += linea[1:]
        elif lineas and lineas[-1].endswith('=') and 'QUOTED-PRINTABLE' in lineas[-1].upper():
            lineas[-1] = lineas[-1][:-1] + linea
        else:
            lineas.append(linea)
    return lineas


def _valor(parametros, valor):
    if 'QUOTED-PRINTABLE' in parametros.upper():
        m = re.search(r'CHARSET=([^;:]+)', parametros, re.I)
        crudo = quopri.decodestring(valor.encode('latin-1', 'replace'))
        try:
            valor = crudo.decode(m.group(1) if m else 'utf-8')
        except (LookupError, UnicodeDecodeError):
            valor = crudo.decode('latin-1')
    return valor.replace('\\,', ',').replace('\\;', ';').replace('\\n', ' ').replace('\\N', ' ').strip()


def filas_de_vcard(texto):
    'Una fila por contacto: Nombre, Telefono, Otros telefonos, Correo. Con encabezado.'
    filas = [list(ENCABEZADOS)]
    actual = None
    for linea in _desplegar(texto):
        if ':' not in linea:
            continue
        clave, valor = linea.split(':', 1)
        nombre = clave.split(';')[0].split('.')[-1].upper()     # 'item1.TEL' -> 'TEL'
        if nombre == 'BEGIN' and valor.strip().upper() == 'VCARD':
            actual = {'fn': '', 'n': '', 'tel': [], 'email': []}
        elif actual is None:
            continue
        elif nombre == 'END':
            nom = actual['fn'] or actual['n']
            tels = [t for i, t in enumerate(actual['tel']) if t not in actual['tel'][:i]]
            if nom or tels or actual['email']:
                filas.append([nom, tels[0] if tels else '', ' / '.join(tels[1:]), actual['email'][0] if actual['email'] else ''])
            actual = None
        elif nombre == 'FN':
            actual['fn'] = _valor(clave, valor)
        elif nombre == 'N':
            # N: apellido;nombre;segundo;prefijo;sufijo. Solo se usa si falta FN.
            partes = [p.strip() for p in _valor(clave, valor.replace('\\;', chr(0))).split(';')]
            partes = [p.replace(chr(0), ';') for p in partes] + [''] * 5
            actual['n'] = ' '.join(p for p in (partes[3], partes[1], partes[2], partes[0], partes[4]) if p)
        elif nombre == 'TEL':
            t = _valor(clave, valor)
            if t.lower().startswith('tel:'):
                t = t[4:]
            if t:
                actual['tel'].append(t)
        elif nombre == 'EMAIL':
            e = _valor(clave, valor)
            if e:
                actual['email'].append(e)
    return filas


# ------------------------------------------------------------ CSV

def filas_de_csv(texto):
    'Parte un CSV respetando las comillas (Google exporta nombres con comas entre comillas).'
    muestra = texto[:4096]
    cuentas = {sep: muestra.count(sep) for sep in (chr(9), ';', ',')}
    sep = max(cuentas, key=cuentas.get) if any(cuentas.values()) else ','
    return [[c.strip() for c in f] for f in csv.reader(io.StringIO(texto), delimiter=sep) if any(c.strip() for c in f)]


def _indices(encabezado, patron):
    return [i for i, h in enumerate(encabezado) if re.search(patron, h, re.I)]


def reducir_contactos(filas):
    """Si el CSV es una exportacion de contactos (Google tiene mas de 30 columnas), lo reduce a Nombre, Telefono,
    Otros telefonos y Correo. Si no lo parece, devuelve None y se usa tal cual."""
    if len(filas) < 2:
        return None
    enc = filas[0]
    tel = [i for i in _indices(enc, r'phone|tel[eé]fono|m[oó]vil|mobile|celular') if not re.search(r'type|tipo|label|etiqueta', enc[i], re.I)]
    completo = _indices(enc, r'^(name|display name|full name|nombre completo|nombre para mostrar)$')
    partes = [i for p in (r'^(first name|given name|nombre de pila|nombre)$', r'^(middle name|additional name|segundo nombre)$',
                          r'^(last name|family name|apellidos?)$') for i in _indices(enc, p)[:1]]
    correo = [i for i in _indices(enc, r'e-?mail|correo') if not re.search(r'type|tipo|label|etiqueta|display|mostrar', enc[i], re.I)]
    if not tel or not (completo or partes):
        return None
    salida = [list(ENCABEZADOS)]
    for f in filas[1:]:
        f = f + [''] * (len(enc) - len(f))
        nombre = f[completo[0]] if completo and f[completo[0]] else ' '.join(f[i] for i in partes if f[i])
        tels = []
        for i in tel:
            for t in re.split(r'\s*:::\s*', f[i]):        # Google junta varios numeros con ' ::: '
                if t and t not in tels:
                    tels.append(t)
        mails = [m for i in correo for m in re.split(r'\s*:::\s*', f[i]) if m]
        if nombre or tels:
            salida.append([nombre, tels[0] if tels else '', ' / '.join(tels[1:]), mails[0] if mails else ''])
    return salida
