import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import app


def test_ayuda_lista_comandos():
    assert 'escribir 1500 en C2' in app.ayuda()


def test_interpretar_escribir():
    resumen, trabajos = app.interpretar('escribir 1500 en c2')
    assert len(trabajos) == 1
    assert 'C2' in resumen


def test_interpretar_formula():
    resumen, trabajos = app.interpretar('formula =B2*C2 en D2')
    assert len(trabajos) == 1
    assert 'D2' in resumen


def test_valor_numeros_y_texto():
    assert app.valor('42') == 42
    assert app.valor('1500,50') == 1500.5
    assert app.valor('texto libre') == 'texto libre'



def test_interpretar_desconocido():
    resumen, trabajos = app.interpretar('haceme un cafe')
    assert trabajos == []
    assert 'No entendi' in resumen


def test_resumen_de_hoja_y_hojas():
    _, t1 = app.interpretar('resumen')
    _, t2 = app.interpretar('hojas')
    assert len(t1) == 1
    assert len(t2) == 1


def test_color_bgr():
    assert app.color_bgr('#FF0000') == 255
    assert app.color_bgr('#00FF00') == 65280


def test_col_num():
    assert app.col_num('A') == 1
    assert app.col_num('AA') == 27


def test_bus_guarda_eventos():
    bus = app.Bus()
    bus.log('hola', 'ok')
    bus.progreso(50, 'mitad')
    assert len(bus.historial) == 2
    assert bus.historial[1].pct == 50



def test_entiende_la_consigna_del_usuario():
    # Ahora la frase pide dos cosas: traer contactos de los correos y armar una tabla.
    frase = 'En la hoja designada crea una tabla que contenga un listado de contactos que surjan de Outlook'
    resumen, trabajos = app.interpretar(frase)
    assert len(trabajos) == 2
    assert 'Outlook' in resumen
    assert 'tabla' in resumen

def test_verbos_flexibles():
    frases = ('escribi 99 en e1', 'sumame la columna C', 'sumar la columna C', 'quita los duplicados', 'mostra excel', 'minimiza excel', 'guarda')
    for frase in frases:
        resumen, trabajos = app.interpretar(frase)
        assert trabajos, 'no entendio: ' + frase


def test_no_entendido_sugiere_algo():
    resumen, trabajos = app.interpretar('haceme un cafe')
    assert trabajos == []
    assert 'No entendi' in resumen



def test_partir_filas_y_valor_celda():
    assert app.partir_filas('a,b' + chr(10) + 'c,d') == [['a', 'b'], ['c', 'd']]
    assert app.partir_filas('x' + chr(9) + 'y') == [['x', 'y']]
    assert app.valor_celda('1500') == 1500
    assert app.valor_celda('011') == '011'
    assert app.valor_celda('ACME') == 'ACME'



def test_procesar_adjunto_texto(tmp_path):
    ruta = tmp_path / 'contactos.csv'
    ruta.write_text('Nombre;Correo' + chr(10) + 'Ana;ana@x.com' + chr(10) + 'Beto;beto@x.com' + chr(10), encoding='utf-8')
    info = app.procesar_adjunto(str(ruta))
    assert info['clase'] == 'texto'
    assert info['columnas'] == 2
    assert len(info['filas']) == 3
    assert 'Ana' in info['texto']



def test_ocr_lee_una_imagen(tmp_path):
    from PIL import Image, ImageDraw, ImageFont
    ruta = tmp_path / 'contactos.png'
    img = Image.new('RGB', (760, 220), 'white')
    d = ImageDraw.Draw(img)
    fuente = ImageFont.truetype('C:/Windows/Fonts/segoeui.ttf', 32)
    d.text((20, 20), 'CONTACTOS', fill='black', font=fuente)
    d.text((20, 80), 'Ana Perez  ana@x.com', fill='black', font=fuente)
    d.text((20, 140), 'Beto Diaz  beto@x.com', fill='black', font=fuente)
    img.save(ruta)
    info = app.procesar_adjunto(str(ruta))
    assert info['clase'] == 'imagen'
    assert 'CONTACTOS' in info['texto'].upper()
    assert 'ana@x.com' in info['texto'].lower()



def test_con_limite_devuelve_el_valor():
    assert app.con_limite(lambda: 42, 5) == 42


def test_con_limite_falla_si_no_responde():
    import time as reloj
    try:
        app.con_limite(lambda: reloj.sleep(4), 1)
        raise AssertionError('deberia haber fallado por tiempo')
    except app.ExcelError as exc:
        assert 'no respondio' in str(exc)


def test_con_limite_propaga_el_error_real():
    def rompe():
        raise ValueError('averia de prueba')
    try:
        app.con_limite(rompe, 5)
        raise AssertionError('deberia haber propagado el error')
    except ValueError as exc:
        assert 'averia de prueba' in str(exc)


def test_marcadores_de_lista():
    assert app.normalizar('1- escribir 5 en A1') == 'escribir 5 en A1'
    assert app.normalizar('2) sumar la columna B') == 'sumar la columna B'
    assert app.normalizar('- negrita en A1') == 'negrita en A1'
    assert app.normalizar('3. hojas') == 'hojas'


def test_contactos_desde_los_correos():
    m1 = '1- Quiero que desde el Outlook instalado, recopiles los contactos registrados en los correos enviados y recibidos.'
    m1 = m1 + chr(10) + '2- Con esos contactos, elabora en el libro designado, un listado de contactos.'
    resumen, trabajos = app.interpretar(m1)
    assert len(trabajos) >= 1
    assert 'correos' in resumen


def test_limpiar_nombre():
    comilla = chr(39)
    dos = chr(34)
    assert app.limpiar_nombre('Mauro Blanco' + comilla) == 'Mauro Blanco'
    assert app.limpiar_nombre(dos + ': Ana Perez' + comilla + dos) == 'Ana Perez'
    assert app.limpiar_nombre('  Beto  ') == 'Beto'


def test_consignas_simples_quedan_separadas():
    # Un mensaje con varias ordenes distintas no debe leerse como una sola tarea.
    mensaje = '1- escribir 5 en A1' + chr(10) + '2- negrita en A1:B1'
    resumen, trabajos = app.interpretar(mensaje)
    assert trabajos == []



def test_mensaje_de_guardar_y_cerrar():
    # Regresion: este mensaje se leia como una extraccion nueva y repetia todo el trabajo.
    frase = 'Por ahora la planilla solicitada esta bien hecha. Salvala y cierra el proceso, renombrandola como ' + chr(34) + 'contactos-clean' + chr(34) + ' en el escritorio.'
    resumen, trabajos = app.interpretar(frase)
    assert len(trabajos) == 2
    assert 'Guardar una copia' in resumen
    assert 'Cerrar el proceso' in resumen
    assert 'contactos-clean' in resumen
    assert 'correos' not in resumen

def test_contactos_con_numero_de_orden():
    frase = 'Desde el correo de Outlook traer los contactos y elaborar una tabla con numero de orden'
    resumen, trabajos = app.interpretar(frase)
    assert 'numero de orden' in resumen
    assert len(trabajos) == 2


def test_contactos_de_todo_el_buzon():
    resumen, trabajos = app.interpretar('recorre todo el buzon y trae los contactos')
    assert 'buzon completo' in resumen
    assert len(trabajos) == 1

def test_correo_valido_descarta_basura():
    # Regresion: undisclosed-recipients y otros valores rotos llegaban a la planilla.
    assert app.correo_valido(str()) == str()
    assert app.correo_valido('undisclosed-recipients:') == str()
    assert app.correo_valido('ana perez@x.com') == str()
    assert app.correo_valido('ana@@x.com') == str()
    assert app.correo_valido('ana@x.com') == 'ana@x.com'



def test_todos_los_contactos_no_es_todo_el_buzon():
    # Regresion: "traer todos los contactos" disparaba un recorrido de 20.000 correos.
    resumen, trabajos = app.interpretar('Desde el correo de Outlook traer todos los contactos (bandeja de entrada y de salida).')
    assert 'buzon completo' not in resumen
    assert len(trabajos) >= 1


def test_todo_el_buzon_si_es_completo():
    resumen, trabajos = app.interpretar('recorre todo el buzon y trae los contactos')
    assert 'buzon completo' in resumen



def test_orden_con_acento():
    # Regresion: la palabra con acento no disparaba la tercera columna.
    frase = "traer los contactos de los correos y elaborar una tabla con n" + chr(250) + "mero de orden"
    resumen, trabajos = app.interpretar(frase)
    assert "numero de orden" in resumen
    assert len(trabajos) == 2



def test_mensaje_de_tres_lineas_con_guardar_y_cerrar():
    # Regresion: el mensaje real de tres lineas se leia como dos pasos y repetia la extraccion.
    lineas = ["Desde el correo de Outlook traer los contactos y elaborar una tabla con numero de orden", "guardala como copia-de-prueba en el escritorio", "cerra el proceso"]
    resumen, trabajos = app.interpretar(chr(10).join(lineas))
    assert len(trabajos) == 4
    assert "Guardar una copia" in resumen
    assert "Cerrar el proceso" in resumen



def test_detener_corta_el_recorrido_de_correos():
    # El boton Detener debe cortar el trabajo en el proximo control.
    s = app.Sesion(app.Bus())
    import pythoncom
    pythoncom.CoInitialize()
    s.detener.set()
    try:
        app.op_contactos_correo(s, False, False)
        raise AssertionError("deberia haber cortado el recorrido")
    except app.ExcelError as exc:
        assert "detenido" in str(exc)

