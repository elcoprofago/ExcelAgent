import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import gui


def test_interfaz_texto_y_avance():
    v = gui.Ventana()
    try:
        v.bus.log('marca uno', 'ok')
        v.bus.progreso(42, 'marca dos')
        v.bus.chat('pregunta', 'usuario')
        v._procesar_eventos()
        panel = v.log.get('1.0', 'end')
        assert 'marca uno' in panel
        assert int(v.pct.get()) == 42
        assert 'pregunta' in v.chat.get('1.0', 'end')
    finally:
        v.raiz.destroy()

def brillo(color):
    c = color.lstrip('#')
    r, g, b = int(c[0:2], 16), int(c[2:4], 16), int(c[4:6], 16)
    return (r * 299 + g * 587 + b * 114) / 1000


def test_colores_pedidos():
    v = gui.Ventana()
    try:
        assert v.log.cget('bg') == gui.FONDO_LOG
        assert v.chat.cget('bg') == gui.FONDO_CHAT
        assert v.raiz.cget('bg') == gui.FONDO_APP
        assert brillo(gui.FONDO_LOG) < 90
        assert brillo(gui.FONDO_CHAT) > 220
        assert brillo(gui.FONDO_APP) > 200
        assert v.log.tag_cget('error', 'foreground') == gui.COLOR_ERROR
        assert v.log.tag_cget('ok', 'foreground') == gui.COLOR_OK
    finally:
        v.raiz.destroy()



def test_texto_enriquecido_en_el_registro():
    v = gui.Ventana()
    try:
        v.bus.log('algo salio mal', 'error')
        v._procesar_eventos()
        rangos = []
        idx = '1.0'
        while True:
            idx = v.log.tag_nextrange('error', idx)
            if not idx:
                break
            rangos.append(idx)
            idx = idx[1]
        assert rangos
        texto = v.log.get(rangos[0][0], rangos[0][1])
        assert 'algo salio mal' in texto
    finally:
        v.raiz.destroy()



def test_boton_adjuntar_presente():
    v = gui.Ventana()
    try:
        encontrados = []
        def recorrer(w):
            for h in w.winfo_children():
                try:
                    if h.cget('text') == '+':
                        encontrados.append(h)
                except Exception:
                    pass
                recorrer(h)
        recorrer(v.raiz)
        assert encontrados
    finally:
        v.raiz.destroy()



def test_cuadro_de_consignas_tiene_tres_lineas():
    v = gui.Ventana()
    try:
        assert gui.ALTO_ENTRADA == 3
        assert int(v.entrada.cget('height')) == 3
    finally:
        v.raiz.destroy()


def test_cuadro_de_consignas_crece_con_el_texto():
    v = gui.Ventana()
    try:
        cinco = 'uno' + chr(10) + 'dos' + chr(10) + 'tres' + chr(10) + 'cuatro' + chr(10) + 'cinco'
        v.entrada.insert('1.0', cinco)
        v._ajustar_alto()
        assert int(v.entrada.cget('height')) == 5
        veinte = ('linea' + chr(10)) * 20
        v.entrada.delete('1.0', 'end')
        v.entrada.insert('1.0', veinte)
        v._ajustar_alto()
        assert int(v.entrada.cget('height')) == gui.ALTO_ENTRADA_MAX
    finally:
        v.raiz.destroy()


def test_el_registro_sobrevive_a_un_fallo():
    # Regresion: un fallo al mostrar un evento mataba el bucle y la ventana quedaba muda.
    v = gui.Ventana()
    try:
        original = v._escribir_log
        estado = {'activo': True}
        def averiado(ev):
            if estado['activo']:
                estado['activo'] = False
                raise RuntimeError('averia simulada')
            return original(ev)
        v._escribir_log = averiado
        v.bus.log('evento que dispara la averia')
        v.bus.log('MARCA_POSTERIOR')
        v._procesar_eventos()
        texto = v.log.get('1.0', 'end')
        assert 'MARCA_POSTERIOR' in texto
        assert ('No pude mostrar' in texto) or ('Fallo el bucle' in texto)
    finally:
        v.raiz.destroy()


def test_no_avisa_si_hay_actividad_reciente():
    # Regresion: avisaba que Excel no respondia mientras leia Outlook.
    v = gui.Ventana()
    try:
        v.ocupado_desde = gui.time.time() - 60
        v.ultimo_evento = gui.time.time()
        v._vigilar_espera()
        v._procesar_eventos()
        texto = v.log.get('1.0', 'end')
        assert 'Sin novedades' not in texto
    finally:
        v.raiz.destroy()

def test_avisa_tras_silencio_prolongado():
    v = gui.Ventana()
    try:
        ahora = gui.time.time()
        v.ocupado_desde = ahora - 60
        v.ultimo_evento = ahora - 60
        v._vigilar_espera()
        v._procesar_eventos()
        texto = v.log.get('1.0', 'end')
        assert 'Sin novedades' in texto
    finally:
        v.raiz.destroy()

def test_boton_visible_con_la_ventana_comprimida():
    # Regresion: con la ventana chica el boton Enviar quedaba fuera de la vista.
    v = gui.Ventana()
    try:
        v.raiz.geometry("720x420")
        for k in range(12):
            v.raiz.update()
            gui.time.sleep(0.05)
        b = v.boton_principal
        alto_ventana = v.raiz.winfo_height()
        y = b.winfo_rooty() - v.raiz.winfo_rooty()
        assert b.winfo_ismapped()
        assert b.winfo_width() > 10
        assert b.winfo_height() > 10
        assert (y + b.winfo_height()) <= alto_ventana
    finally:
        v.raiz.destroy()



def test_boton_alterna_entre_enviar_y_detener():
    v = gui.Ventana()
    try:
        for k in range(8):
            v.raiz.update()
            gui.time.sleep(0.05)
        imagen_libre = str(v.boton_principal.cget("image"))
        v.ocupado_desde = gui.time.time()
        v._actualizar_botones()
        imagen_ocupado = str(v.boton_principal.cget("image"))
        assert imagen_libre != imagen_ocupado
        v.boton_principal.invoke()
        fin = gui.time.time() + 3
        while gui.time.time() < fin and not v.agente.sesion.detener.is_set():
            v.raiz.update()
            gui.time.sleep(0.05)
        assert v.agente.sesion.detener.is_set()
    finally:
        v.raiz.destroy()

