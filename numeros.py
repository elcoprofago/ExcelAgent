# Numeros y fechas que llegan como texto (un CSV, una tabla de Word o de un PDF, lo que escribe el modelo): se
# interpretan como lo que son antes de escribirlos en la hoja. Sin esto, '1.500' quedaba 1,5, '$ 1.500' y '15%' como
# texto, y '01/09/2026' Excel lo tomaba como 9 de enero (por COM lee los textos con la configuracion de EE.UU.).
#
# Lo ambiguo se decide por la columna: '1.500' es mil quinientos si en la misma columna hay '2.300,50' (coma decimal),
# y uno coma cinco si hay '3.25'. Sin pistas en la columna, vale la coma decimal del Excel del usuario (Argentina).
# Quedan como texto los codigos con ceros a la izquierda ('007') y los numeros de mas de 15 cifras (un CBU): Excel
# los guardaria redondeados.
import datetime
import re

_ESPACIOS = str.maketrans({' ': ' ', ' ': ' '})
_MONEDA = re.compile(r'^(?P<signo1>[+-]?)\s*(?P<moneda>US\$|U\$S|\$|€)?\s*(?P<signo2>[+-]?)\s*(?P<num>[0-9][0-9., ]*)\s*'
                     r'(?P<pct>%)?$')
# dd/mm/aaaa, dd-mm-aa, dd.mm.aaaa (con punto solo con el año entero: '1.10.20' es mas probable una version).
_FECHA = re.compile(r'^(\d{1,2})([/-])(\d{1,2})\2(\d{4}|\d{2})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$|'
                    r'^(\d{1,2})(\.)(\d{1,2})\.(\d{4})(?:\s+(\d{1,2}):(\d{2})(?::(\d{2}))?)?$')
_FECHA_ISO = re.compile(r'^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?)?$')
_HORA = re.compile(r'^(\d{1,2}):(\d{2})(?::(\d{2}))?$')
_FORMATO_MONEDA = {'$': '$ #,##0.00', 'US$': '"US$ "#,##0.00', 'U$S': '"US$ "#,##0.00', '€': '"€ "#,##0.00'}
MAX_CIFRAS = 15


class Texto(str):
    'Un texto que tiene que quedar como texto aunque parezca numero (se escribe con apostrofo).'


def _partes(num):
    """Las lecturas posibles de '1.500,50' y parecidos: lista de (entero, fraccion, separador_decimal) y si el caso
    es ambiguo. Vacia si no es un numero bien escrito."""
    if ' ' in num:
        m = re.fullmatch(r'(\d{1,3}(?: \d{3})+)(?:([.,])(\d+))?', num)
        return ([(m.group(1).replace(' ', ''), m.group(3) or '', m.group(2))], False) if m else ([], False)
    seps = [c for c in num if c in '.,']
    if not seps:
        return [(num, '', None)], False
    if len(set(seps)) == 2:
        dec = seps[-1]
        mil = ',' if dec == '.' else '.'
        entero, _, frac = num.rpartition(dec)
        if seps.count(dec) == 1 and re.fullmatch(r'\d{1,3}(?:' + re.escape(mil) + r'\d{3})+', entero) and frac.isdigit():
            return [(entero.replace(mil, ''), frac, dec)], False
        return [], False
    sep = seps[0]
    if len(seps) > 1:
        if re.fullmatch(r'\d{1,3}(?:' + re.escape(sep) + r'\d{3})+', num):
            return [(num.replace(sep, ''), '', None)], False
        return [], False
    entero, frac = num.split(sep)
    if not (entero.isdigit() and frac.isdigit()):
        return [], False
    if len(frac) == 3 and entero != '0' and len(entero) <= 3 and entero[0] != '0':
        return [(entero, frac, sep), (entero + frac, '', None)], True
    return [(entero, frac, sep)], False


def es_fecha(texto):
    s = str(texto).strip()
    return bool(_FECHA.match(s) or _FECHA_ISO.match(s))


def pista_decimal(texto):
    'El separador decimal que un texto deja claro (\'2.300,50\' -> \',\'; \'3.25\' -> \'.\'), o None.'
    m = _MONEDA.match(str(texto).strip().translate(_ESPACIOS))
    if not m:
        return None
    num = m.group('num').strip()
    seps = [c for c in num if c in '.,']
    if not seps or ' ' in num:
        return None
    if len(set(seps)) == 2:
        return seps[-1]
    if len(seps) > 1:
        return ',' if seps[0] == '.' else '.'
    lecturas, ambiguo = _partes(num)
    return lecturas[0][2] if lecturas and not ambiguo else None


def _grupos_fecha(m):
    'dia-o-mes, mes-o-dia, año, hora, minutos, segundos de un _FECHA, sea cual sea la alternativa que coincidio.'
    g = m.groups()
    g = g[:7] if g[0] is not None else g[7:]
    return int(g[0]), int(g[2]), int(g[3]), g[4], g[5], g[6]


def pista_fecha(texto):
    "'dmy' si el primer numero pasa de 12 ('25/09/2026'), 'mdy' si el segundo ('09/25/2026'); None si no dice nada."
    m = _FECHA.match(str(texto).strip())
    if not m:
        return None
    a, b = _grupos_fecha(m)[:2]
    return 'dmy' if a > 12 >= b else 'mdy' if b > 12 >= a else None


def convencion(textos, decimal=','):
    """Separador decimal y orden de fechas de una columna, por mayoria de las pistas que dan sus valores; lo que no
    tenga pistas queda con los del usuario (decimal, dia/mes/año)."""
    votos, fechas = {}, {}
    for t in textos:
        if isinstance(t, str) and t.strip():
            d = pista_decimal(t)
            if d:
                votos[d] = votos.get(d, 0) + 1
            f = pista_fecha(t)
            if f:
                fechas[f] = fechas.get(f, 0) + 1
    dec = decimal
    if votos and len(set(votos.values())) == len(votos):
        dec = max(votos, key=votos.get)
    orden = 'dmy'
    if fechas.get('mdy', 0) > fechas.get('dmy', 0):
        orden = 'mdy'
    return dec, orden


def _fecha(anio, mes, dia, h, mi, s):
    if anio < 100:
        anio += 2000 if anio < 70 else 1900
    try:
        if h is None:
            return datetime.date(anio, mes, dia)
        return datetime.datetime(anio, mes, dia, h, mi, s or 0)
    except ValueError:
        return None


def interpretar(texto, decimal=',', orden='dmy'):
    """(valor, formato): el numero o la fecha que el texto escribe y el formato para que se vea igual ('$ #,##0.00',
    '0%', 'dd/mm/yyyy'), o None si no hace falta. Un codigo que tiene que seguir siendo texto vuelve como Texto; lo
    que no es numero ni fecha vuelve tal cual."""
    if not isinstance(texto, str):
        return texto, None
    s = texto.strip().translate(_ESPACIOS)
    if not s:
        return texto, None
    m = _FECHA.match(s)
    if m:
        a, b, anio, h, mi, se = _grupos_fecha(m)
        dia, mes = (a, b) if orden == 'dmy' else (b, a)
        if a > 12 >= b:
            dia, mes = a, b
        elif b > 12 >= a:
            dia, mes = b, a
        h = int(h) if h else None
        v = _fecha(anio, mes, dia, h, int(mi or 0), int(se or 0))
        if v:
            return v, 'dd/mm/yyyy hh:mm' if h is not None else 'dd/mm/yyyy'
        return texto, None
    m = _FECHA_ISO.match(s)
    if m:
        h = int(m.group(4)) if m.group(4) else None
        v = _fecha(int(m.group(1)), int(m.group(2)), int(m.group(3)), h, int(m.group(5) or 0), int(m.group(6) or 0))
        if v:
            return v, 'dd/mm/yyyy hh:mm' if h is not None else 'dd/mm/yyyy'
        return texto, None
    m = _HORA.match(s)
    if m:
        h, mi, se = int(m.group(1)), int(m.group(2)), int(m.group(3) or 0)
        if h < 24 and mi < 60 and se < 60:
            return datetime.time(h, mi, se), 'hh:mm:ss' if m.group(3) else 'hh:mm'
        return texto, None
    m = _MONEDA.match(s)
    if not m or (m.group('signo1') and m.group('signo2')):
        return texto, None
    num = m.group('num').strip()
    if re.fullmatch(r'0\d+(?:[.,]\d+)?', num):
        return Texto(s), None
    if s.startswith('+') and len(re.sub(r'\D', '', num)) >= 7 and not m.group('pct'):
        return Texto(s), None      # un telefono con codigo de pais: como numero perderia el '+'
    lecturas, ambiguo = _partes(num)
    if not lecturas:
        return texto, None
    entero, frac, sep = lecturas[0]
    if ambiguo and sep != decimal:
        entero, frac, sep = lecturas[1]
    if len(entero.lstrip('0')) + len(frac) > MAX_CIFRAS:
        return Texto(s), None
    negativo = '-' in (m.group('signo1') + m.group('signo2'))
    v = float(entero + '.' + frac) if frac else int(entero)
    if negativo:
        v = -v
    formato = None
    if m.group('pct'):
        v = v / 100
        formato = '0.' + '0' * len(frac) + '%' if frac else '0%'
    elif m.group('moneda'):
        formato = _FORMATO_MONEDA[m.group('moneda')]
    elif ' ' in num or (not frac and len(entero) > 3 and num != entero) or (frac and re.search(r'\d[.,]\d{3}[.,]', num)):
        formato = '#,##0' + ('.' + '0' * len(frac) if frac else '')
    return v, formato
