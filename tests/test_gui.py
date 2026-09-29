import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import tempfile

# La configuracion real (%APPDATA%\ExcelAgent) y el registro real (logs\) no se tocan: cada corrida usa una
# carpeta propia. Regresion del 29/09/2026: el registro del usuario quedaba mezclado con las marcas de las pruebas.
os.environ['EXCELAGENT_CONFIG_DIR'] = tempfile.mkdtemp(prefix='excelagent_cfg_')
os.environ['EXCELAGENT_LOG_DIR'] = os.path.join(os.environ['EXCELAGENT_CONFIG_DIR'], 'logs')

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



def test_clip_arriba_de_enviar():
    # Pedido del usuario (29/09/2026): el '+' de la barra de arriba costaba verlo; ahora es un clip sobre Enviar.
    v = gui.Ventana()
    try:
        for k in range(8):
            v.raiz.update()
            gui.time.sleep(0.05)
        c, b = v.boton_adjuntar, v.boton_principal
        assert c.winfo_ismapped() and str(c.cget('image'))          # con el icono, no con texto
        assert c.winfo_rooty() + c.winfo_height() <= b.winfo_rooty()  # arriba de Enviar
        assert abs(c.winfo_rootx() - b.winfo_rootx()) <= 4          # en la misma columna
        # Control: ya no queda un '+' en la barra de arriba.
        def textos(w):
            for h in w.winfo_children():
                try:
                    yield h.cget('text')
                except Exception:
                    pass
                yield from textos(h)
        assert '+' not in list(textos(v.raiz))
        c.configure(command=lambda: v.bus.log('marca clip', 'ok'))
        c.invoke()
        v._procesar_eventos()
        assert 'marca clip' in v.log.get('1.0', 'end')
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
        alto_ventana = v.raiz.winfo_height()
        for b in (v.boton_principal, v.boton_adjuntar):
            y = b.winfo_rooty() - v.raiz.winfo_rooty()
            assert b.winfo_ismapped()
            assert b.winfo_width() > 10
            assert b.winfo_height() > 10
            assert y >= 0 and (y + b.winfo_height()) <= alto_ventana
        # Regresion: a 760 de ancho Nueva charla quedaba afuera y el saldo, cortado.
        v.saldo_var.set('Saldo: US$ 5.91 - sesion: -US$ 0.02 - 19:00')
        v.raiz.update()
        ancho = v.raiz.winfo_width()
        def todos(w):
            for h in w.winfo_children():
                yield h
                yield from todos(h)
        for w in todos(v.raiz):
            try:
                t = w.cget('text') or str(w.cget('textvariable'))
            except Exception:
                continue
            if t in ('Nueva charla', 'Configuracion', '⟳', 'Actualizar', str(v.saldo_var)):
                x = w.winfo_rootx() - v.raiz.winfo_rootx()
                assert w.winfo_ismapped() and x >= 0 and x + w.winfo_width() <= ancho, t
                assert w.winfo_width() >= w.winfo_reqwidth(), t      # entero, no recortado
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



class Stream:
    'Como el real: cancel() corta la respuesta (el real cierra la conexion).'
    def __init__(self, eventos):
        self.eventos = eventos
        self.cortado = False

    def cancel(self):
        self.cortado = True

    def __iter__(self):
        for e in self.eventos:
            if self.cortado:
                return
            if callable(e):
                e()
            else:
                yield e


def esperar(v, condicion, segundos=5):
    fin = gui.time.time() + segundos
    while gui.time.time() < fin and not condicion():
        v.raiz.update()
        v._procesar_eventos()
        gui.time.sleep(0.02)
    return condicion()


def test_pedido_completo_con_modelo_falso():
    uso = ('usage', {'prompt_tokens': 1200, 'completion_tokens': 30, 'prompt_cache_hit_tokens': 1000})
    v = gui.Ventana(stream_factory=lambda entry, msgs, tools: Stream(
        [('content', 'Usa '), ('content', '=SUMA(A1:A3).'), uso, ('finish', 'stop')]))
    try:
        v.entrada.insert('1.0', 'como sumo?')
        v._enviar()
        assert v.en_pedido
        assert esperar(v, lambda: not v.en_pedido)
        chat = v.chat.get('1.0', 'end')
        assert 'VOS: como sumo?' in chat
        assert 'ASISTENTE: Usa =SUMA(A1:A3).' in chat
        assert v.estado_var.get() == 'Listo'
        tokens = v.tokens_var.get()
        assert '1.230' in tokens and 'cache 1.000' in tokens
        assert v.agente.totals['in'] == 1200
        # La respuesta va al registro en una sola linea, no una por pieza del streaming.
        with open(v.archivo_log, encoding='utf-8') as f:
            lineas = [l for l in f if '[asistente]' in l]
        assert len(lineas) == 1 and lineas[0].rstrip().endswith('[asistente] Usa =SUMA(A1:A3).')
    finally:
        v.raiz.destroy()


def test_sin_key_no_envia_y_conserva_el_texto():
    v = gui.Ventana()
    try:
        assert v.agente.necesita_key()
        v.entrada.insert('1.0', 'hola')
        v._enviar()
        v._procesar_eventos()
        assert not v.en_pedido
        assert v._texto_consigna() == 'hola'
        assert 'API key' in v.chat.get('1.0', 'end')
    finally:
        v.raiz.destroy()


def test_detener_corta_un_pedido_en_curso():
    v = gui.Ventana()
    try:
        v.agente._stream_factory = v._stream_factory = lambda entry, msgs, tools: Stream(
            [('content', 'empiezo')] + [lambda: gui.time.sleep(0.05)] * 200 + [('finish', 'stop')])
        v.entrada.insert('1.0', 'algo largo')
        v._enviar()
        assert esperar(v, lambda: 'empiezo' in v.chat.get('1.0', 'end'))
        v.boton_principal.invoke()
        assert esperar(v, lambda: not v.en_pedido)
        assert 'Interrumpido' in v.chat.get('1.0', 'end')
    finally:
        v.raiz.destroy()


def test_selector_de_modelo_y_esfuerzo():
    v = gui.Ventana()
    try:
        opciones = list(v.modelo_cb.cget('values'))
        assert 'deepseek-flash' in ' '.join(opciones) and 'deepseek-v4-pro' in ' '.join(opciones)
        pro = [o for o in opciones if 'deepseek-v4-pro' in o][0]
        v.modelo_var.set(pro)
        v._modelo_cambiado()
        assert v.agente.model_id == 'deepseek-v4-pro' and v.cfg['model'] == 'deepseek-v4-pro'
        assert list(v.esfuerzo_cb.cget('values')) == [gui.SIN_EFFORT, 'low', 'high', 'max']
        v.esfuerzo_var.set('max')
        v._esfuerzo_cambiado()
        assert v.agente.effort == 'max' and v.cfg['effort'] == 'max'
        # un modelo local sin razonamiento configurable deja el esfuerzo deshabilitado
        v.agente.model_id = gui.localmodels.MODEL_ID_PREFIX + 'Z:/no/existe.gguf'
        v._refrescar_esfuerzo()
        assert str(v.esfuerzo_cb.cget('state')) == 'disabled'
    finally:
        v.raiz.destroy()


def test_formato_de_miles():
    assert gui.miles(1234567) == '1.234.567'
    assert gui.miles(0) == '0'


def test_saldo_muestra_la_hora_y_lo_gastado_en_la_sesion(monkeypatch):
    # El saldo llega con dos decimales: con pedidos chicos el numero no cambia y el boton parecia no andar.
    saldos = iter(['5.93', '5.93', '5.91'])
    def falso(key):
        t = next(saldos)
        return {'available': True, 'text': 'US$ ' + t, 'amounts': {'USD': float(t)}}
    monkeypatch.setattr(gui.dsapi, 'get_balance', falso)
    v = gui.Ventana()
    monkeypatch.setattr(type(v.cfg), 'api_key', property(lambda self: 'k-falsa'))
    try:
        v._refrescar_saldo()
        assert esperar(v, lambda: v.saldo_var.get() != 'Saldo: ...')
        assert v.saldo_var.get().startswith('Saldo: US$ 5.93 - ') and 'sesion' not in v.saldo_var.get()
        v._refrescar_saldo(manual=True)
        assert esperar(v, lambda: v.saldo_var.get() != 'Saldo: ...')
        assert 'sesion: -US$ 0.00' in v.saldo_var.get()
        assert 'dos decimales' in v.log.get('1.0', 'end')           # el boton deja constancia de que consulto
        v._refrescar_saldo()
        assert esperar(v, lambda: v.saldo_var.get() != 'Saldo: ...')
        assert 'sesion: -US$ 0.02' in v.saldo_var.get()
    finally:
        v.raiz.destroy()


def test_titulo_con_la_version():
    v = gui.Ventana()
    try:
        assert v.raiz.title() == 'ExcelAgent ' + gui.version.VERSION + ' - asesor de Excel'
    finally:
        v.raiz.destroy()


def test_boton_actualizar(monkeypatch):
    # Sin red: GitHub, la instalacion y los cuadros de dialogo son falsos.
    dialogos = []
    monkeypatch.setattr(gui.messagebox, 'showinfo', lambda t, m: dialogos.append(('info', m)))
    monkeypatch.setattr(gui.messagebox, 'showerror', lambda t, m: dialogos.append(('error', m)))
    respuestas = iter([True, False])        # si a instalar, no a reiniciar ahora
    monkeypatch.setattr(gui.messagebox, 'askyesno', lambda t, m: dialogos.append(('pregunta', m)) or next(respuestas))
    instaladas = []
    monkeypatch.setattr(gui.actualizar, 'actualizar',
                        lambda rel, avance=None: (avance(50, 100), instaladas.append(rel['version']), 'C:/respaldo')[-1])
    monkeypatch.setattr(gui.actualizar, 'relanzar', lambda: instaladas.append('relanzo'))
    rel = {'tag': 'v' + gui.version.VERSION, 'version': gui.version.VERSION, 'nombre': 'misma', 'notas': ''}
    monkeypatch.setattr(gui.actualizar, 'ultima_release', lambda: dict(rel))
    v = gui.Ventana()
    try:
        # 1) ya esta la ultima
        v._actualizar()
        assert esperar(v, lambda: not v.actualizando)
        assert dialogos == [('info', 'Ya tenes la ultima version (' + gui.version.VERSION + ').')]
        assert v.boton_actualizar.cget('state') == 'normal'
        # 2) hay una nueva: la ofrece con sus notas, la instala y no reinicia porque se dijo que no
        dialogos.clear()
        rel.update(tag='v99.0.0', version='99.0.0', nombre='Version 99', notas='Cosas nuevas')
        v._actualizar()
        assert v.boton_actualizar.cget('state') == 'disabled'
        assert esperar(v, lambda: not v.actualizando and len(dialogos) == 2)
        assert 'Cosas nuevas' in dialogos[0][1] and 'Version 99' in dialogos[0][1]
        assert 'instalada la version 99.0.0' in dialogos[1][1]
        assert instaladas == ['99.0.0']
        assert int(v.pct.get()) == 100 and 'C:/respaldo' in v.log.get('1.0', 'end')
        # 3) un error se muestra y el boton vuelve a andar
        dialogos.clear()
        monkeypatch.setattr(gui.actualizar, 'ultima_release',
                            lambda: (_ for _ in ()).throw(gui.actualizar.ErrorActualizacion('sin internet')))
        v._actualizar()
        assert esperar(v, lambda: not v.actualizando)
        assert dialogos == [('error', 'sin internet')] and v.boton_actualizar.cget('state') == 'normal'
        # 4) con un pedido en curso no arranca
        dialogos.clear()
        v.en_pedido = True
        v._actualizar()
        assert not v.actualizando and dialogos[0][0] == 'info' and 'trabajo en curso' in dialogos[0][1]
        v.en_pedido = False
    finally:
        v.raiz.destroy()
