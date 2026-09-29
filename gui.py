# Interfaz grafica: chat con el asesor, registro de eventos, selector de modelo y medidor de tokens.
from __future__ import annotations

import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import adjuntos
import asistente
import dialogos
import dsapi
import excel
import localmodels
import meter

FONDO_APP = '#CCE2F5'
FONDO_CHAT = '#E8F4FF'
FONDO_LOG = '#0B2545'
FONDO_BOTON = '#A9CDEB'
FONDO_ENTRADA = '#FFFFFF'
COLOR_TITULO = '#0B2545'
COLOR_HORA = '#8FB8DE'
COLOR_INFO = '#D6E4F0'
COLOR_OK = '#7CFC9A'
COLOR_AVISO = '#FFD166'
COLOR_ERROR = '#FF6B6B'
COLOR_AVANCE = '#7FD1FF'
COLOR_USUARIO = '#0B4A8B'
COLOR_AGENTE = '#14453A'
COLOR_AVISO_CHAT = '#8A4B08'
COLOR_APAGADO = '#5B7FA6'
CARPETA_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
IMG_ENVIAR = os.path.join(CARPETA_ASSETS, 'b-env.png')
IMG_DETENER = os.path.join(CARPETA_ASSETS, 'b-stop.png')
PISTA_TEXTO = 'Enter envia - Shift+Enter baja de linea - contale con tus palabras que queres hacer'
ALTO_ENTRADA = 3
ALTO_ENTRADA_MAX = 10
ESPERA_MAXIMA = 25
REAVISO = 30
# EXCELAGENT_LOG_DIR es para las pruebas: antes escribian en este mismo registro y se mezclaban con el uso real.
RUTA_LOGS = os.environ.get('EXCELAGENT_LOG_DIR') or os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
MARCA = '------------------------------------------------------------'
SIN_EFFORT = '(por defecto)'
MAX_LOG_HERRAMIENTA = 300
BUSY_NOTE = 'Hay un pedido en curso. Espera a que termine o apreta Detener.'

AYUDA = """Soy tu asesor de Excel. Pedime las cosas como te salgan, por ejemplo:
- "ponele lindo al encabezado"
- "sacame los repetidos"
- "quiero ver cuanto vendio cada cliente"
- "marcame en rojo lo que pase de 1000"
- "armame un grafico de las ventas por mes"
- "como hago para que la primera fila quede fija?"

Para trabajar sobre un archivo, elegilo con Abrir Excel (hago una copia de respaldo al abrirlo). Con + me pasas otra planilla
(Excel, OpenOffice, csv), un documento (Word, PDF, OpenOffice, .md), una imagen o los contactos exportados del telefono
(.vcf) para leerlos o volcarlos en la planilla. Antes de pisar datos, borrar o guardar te pido permiso. No guardo el archivo si no me lo pedis.

Arriba elegis el modelo: los de DeepSeek necesitan la API key (Configuracion); los [local] corren en esta PC, sin internet,
pero son mas lentos. El esfuerzo es cuanto piensa antes de contestar: mas esfuerzo, mejores respuestas y mas tokens.
La barra de consumo muestra los tokens usados desde que abriste el programa."""


def miles(n):
    return f'{int(n):,}'.replace(',', '.')


def recortar(texto, n=MAX_LOG_HERRAMIENTA):
    t = ' '.join(str(texto).split())
    return t if len(t) <= n else t[:n] + '...'


class Ventana:
    'Interfaz principal.'

    def __init__(self, cfg=None, stream_factory=None):
        self.raiz = tk.Tk()
        self.raiz.title('ExcelAgent - asesor de Excel')
        self.raiz.geometry('1120x760')
        self.raiz.configure(bg=FONDO_APP)
        self.cfg = cfg or dsapi.Config()
        self.peticiones = queue.Queue()
        self.eventos = queue.Queue()
        self.bus = excel.Bus(sink=self.eventos.put)
        self._stream_factory = stream_factory
        self.agente = asistente.Asistente(self.cfg, self.bus, stream_factory=stream_factory)
        self.medidor = meter.TokenMeter()
        self.ruta = None
        self.ocupado_desde = None
        self.en_pedido = False
        self.cancel_ev = None
        self._dialogo = None
        self._paso_con_texto = False
        self._texto_paso = []       # la respuesta en curso: va al registro entera, no pieza por pieza
        self.ultimo_aviso = 0.0
        self.ultimo_evento = time.time()
        self.archivo_log = None
        self.pct = tk.DoubleVar(value=0.0)
        self.texto_pct = tk.StringVar(value='0 ' + chr(37))
        self.ruta_var = tk.StringVar(value='(sin archivo designado)')
        self.modelo_var = tk.StringVar()
        self.esfuerzo_var = tk.StringVar()
        self.saldo_var = tk.StringVar(value='Saldo: -')
        self.estado_var = tk.StringVar(value='Listo')
        self.tokens_var = tk.StringVar(value='')
        self._display_a_id = {}
        self.img_enviar = None
        self.img_detener = None
        self.trabajando_antes = False
        self.cerrando = False
        self.tarea_after = None
        self._cargar_iconos()
        self.raiz.minsize(760, 460)
        self._preparar_log()
        self._construir()
        self._instalar_manejadores()
        self._arrancar_trabajador()
        self.raiz.bind("<Destroy>", self._al_soltar)
        self._agendar_eventos()
        self.raiz.protocol('WM_DELETE_WINDOW', self._al_cerrar)
        self._refrescar_modelos_widgets()
        self.bus.log('Interfaz lista. Podes preguntar algo o elegir un libro con Abrir Excel.', 'ok')
        if self.cfg.load_warning:
            self.bus.log(self.cfg.load_warning, 'warn')
        self._al_arrancar()

    # ------------------------------------------------------------ registro en disco y errores

    def _preparar_log(self):
        try:
            if not os.path.isdir(RUTA_LOGS):
                os.makedirs(RUTA_LOGS, exist_ok=True)
            t = time.localtime()
            nombre = 'excelagent_' + str(t.tm_year) + str(t.tm_mon).zfill(2) + str(t.tm_mday).zfill(2) + '.log'
            self.archivo_log = os.path.join(RUTA_LOGS, nombre)
            self._anotar(MARCA)
            self._anotar('Sesion iniciada')
        except Exception:
            self.archivo_log = None

    def _anotar(self, texto):
        'Deja constancia en disco. Nunca debe romper la interfaz.'
        if not self.archivo_log:
            return
        try:
            with open(self.archivo_log, 'a', encoding='utf-8') as f:
                f.write(excel.hora_actual() + ' ' + str(texto) + chr(10))
        except Exception:
            pass

    def _instalar_manejadores(self):
        'Cualquier error queda a la vista y en disco, aunque no haya consola.'
        self.raiz.report_callback_exception = self._error_tk
        sys.excepthook = self._error_sys
        try:
            threading.excepthook = self._error_hilo
        except Exception:
            pass

    def _error_tk(self, tipo, valor, rastro):
        self._error_interno('Error en la interfaz: ' + str(valor))

    def _error_sys(self, tipo, valor, rastro):
        self._error_interno('Error no controlado: ' + str(valor))

    def _error_hilo(self, datos):
        self._error_interno('Error en un hilo de trabajo: ' + str(datos.exc_value))

    def _error_interno(self, texto, nivel='error'):
        'Muestra un problema por pantalla y lo deja en el archivo de registro.'
        self._anotar(texto)
        try:
            self.log.configure(state='normal')
            self.log.insert('end', excel.hora_actual() + ' ', 'hora')
            self.log.insert('end', '[' + nivel + '] ', nivel)
            self.log.insert('end', texto + chr(10), nivel)
            self.log.see('end')
            self.log.configure(state='disabled')
        except Exception:
            pass

    # ------------------------------------------------------------ armado

    def _construir(self):
        barra = tk.Frame(self.raiz, bg=FONDO_APP, padx=10, pady=8)
        barra.pack(fill='x')
        tk.Button(barra, text='Abrir Excel', width=13, bg=FONDO_BOTON, command=self._elegir).pack(side='left')
        caja = tk.Entry(barra, textvariable=self.ruta_var, state='readonly')
        caja.pack(side='left', fill='x', expand=True, padx=(10, 0))
        tk.Button(barra, text='+', width=3, bg=FONDO_BOTON, command=self._adjuntar).pack(side='left', padx=(10, 0))
        tk.Button(barra, text='Ayuda', width=8, bg=FONDO_BOTON, command=self._ayuda).pack(side='left', padx=(8, 0))
        tk.Button(barra, text='Reconectar', width=11, bg=FONDO_BOTON, command=self._reiniciar_conexion).pack(side='left', padx=(8, 0))

        # segunda barra: con que modelo se trabaja y cuanto piensa
        modelos = tk.Frame(self.raiz, bg=FONDO_APP, padx=10)
        modelos.pack(fill='x')
        tk.Label(modelos, text='Modelo', bg=FONDO_APP, fg=COLOR_TITULO).pack(side='left')
        self.modelo_cb = ttk.Combobox(modelos, textvariable=self.modelo_var, state='readonly', width=42)
        self.modelo_cb.pack(side='left', padx=(6, 12))
        self.modelo_cb.bind('<<ComboboxSelected>>', lambda e: self._modelo_cambiado())
        tk.Label(modelos, text='Esfuerzo', bg=FONDO_APP, fg=COLOR_TITULO).pack(side='left')
        self.esfuerzo_cb = ttk.Combobox(modelos, textvariable=self.esfuerzo_var, state='readonly', width=14)
        self.esfuerzo_cb.pack(side='left', padx=(6, 12))
        self.esfuerzo_cb.bind('<<ComboboxSelected>>', lambda e: self._esfuerzo_cambiado())
        tk.Button(modelos, text='Configuracion', width=13, bg=FONDO_BOTON, command=self._abrir_configuracion).pack(side='left')
        tk.Button(modelos, text='Nueva charla', width=12, bg=FONDO_BOTON, command=self._nueva_charla).pack(side='left', padx=(8, 0))
        tk.Button(modelos, text='⟳', width=3, bg=FONDO_BOTON, command=self._refrescar_saldo).pack(side='right')
        tk.Label(modelos, textvariable=self.saldo_var, bg=FONDO_APP, fg=COLOR_TITULO).pack(side='right', padx=(0, 6))

        # medidor: estado del pedido y consumo de tokens
        medidor = tk.Frame(self.raiz, bg=FONDO_APP, padx=10, pady=4)
        medidor.pack(fill='x')
        tk.Label(medidor, textvariable=self.estado_var, bg=FONDO_APP, fg=COLOR_TITULO, width=34, anchor='w',
                 font=('Segoe UI', 9, 'bold')).pack(side='left')
        self.consumo = ttk.Progressbar(medidor, maximum=100.0, length=160)
        self.consumo.pack(side='left', padx=(0, 8))
        self.tokens_lbl = tk.Label(medidor, textvariable=self.tokens_var, bg=FONDO_APP, fg=COLOR_APAGADO, anchor='w')
        self.tokens_lbl.pack(side='left', fill='x', expand=True)

        cuerpo = tk.PanedWindow(self.raiz, orient='horizontal', sashwidth=6, bg=FONDO_APP)
        cuerpo.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        izq = tk.LabelFrame(cuerpo, text='Chat con el asistente', bg=FONDO_APP, fg=COLOR_TITULO, padx=6, pady=6)
        self.chat = ScrolledText(izq, wrap='word', state='disabled', height=20, bg=FONDO_CHAT, fg=COLOR_AGENTE, relief='flat')
        self.chat.tag_config('hora', foreground=COLOR_APAGADO)
        self.chat.tag_config('usuario', foreground=COLOR_USUARIO, font=('Segoe UI', 10, 'bold'))
        self.chat.tag_config('agente', foreground=COLOR_AGENTE, font=('Segoe UI', 10))
        self.chat.tag_config('aviso', foreground=COLOR_AVISO_CHAT, font=('Segoe UI', 10, 'bold'))

        envio = tk.Frame(izq, bg=FONDO_APP)
        marco = tk.Frame(envio, bg=FONDO_APP)
        marco.pack(side='left', fill='both', expand=True)
        self.entrada = tk.Text(marco, height=ALTO_ENTRADA, wrap='word', bg=FONDO_ENTRADA, fg=COLOR_AGENTE, relief='solid', borderwidth=1, undo=True, font=('Segoe UI', 10))
        self.entrada.pack(side='left', fill='both', expand=True)
        self.barra_entrada = tk.Scrollbar(marco, command=self.entrada.yview)
        self.barra_entrada.pack(side='right', fill='y')
        self.entrada.configure(yscrollcommand=self.barra_entrada.set)
        self.entrada.bind('<Return>', self._tecla_enter)
        self.entrada.bind('<Shift-Return>', self._tecla_shift)
        self.entrada.bind('<<Modified>>', self._ajustar_alto)
        self.boton_principal = tk.Button(envio, bd=0, relief='flat', highlightthickness=0, cursor='hand2', bg=FONDO_APP, activebackground=FONDO_APP, command=self._enviar_o_detener)
        if self.img_enviar is not None:
            self.boton_principal.configure(image=self.img_enviar, text=str())
        else:
            self.boton_principal.configure(text='Enviar', width=10, bg=FONDO_BOTON)
        # El boton va primero en el empaquetado: asi conserva su ancho y no queda recortado.
        self.boton_principal.pack(side='right', padx=(8, 0), anchor='s')
        marco.pack_forget()
        marco.pack(side='left', fill='both', expand=True)
        self.pista = tk.Label(izq, text=PISTA_TEXTO, bg=FONDO_APP, fg=COLOR_TITULO, anchor='w')
        # El pie se empaqueta primero y de abajo hacia arriba: el boton queda siempre a la vista.
        self.pista.pack(side='bottom', fill='x', pady=(4, 0))
        envio.pack(side='bottom', fill='x', pady=(6, 0))
        self.chat.pack(fill='both', expand=True)
        cuerpo.add(izq, stretch='always')

        der = tk.LabelFrame(cuerpo, text='Registro de eventos y avance', bg=FONDO_APP, fg=COLOR_TITULO, padx=6, pady=6)
        self.log = ScrolledText(der, wrap='word', state='disabled', width=52, bg=FONDO_LOG, fg=COLOR_INFO, relief='flat', insertbackground=COLOR_INFO)
        self.log.pack(fill='both', expand=True)
        self.log.tag_config('hora', foreground=COLOR_HORA)
        self.log.tag_config('info', foreground=COLOR_INFO)
        self.log.tag_config('avance', foreground=COLOR_AVANCE)
        self.log.tag_config('ok', foreground=COLOR_OK)
        self.log.tag_config('warn', foreground=COLOR_AVISO)
        self.log.tag_config('error', foreground=COLOR_ERROR)

        avance = tk.Frame(der, bg=FONDO_APP)
        avance.pack(fill='x', pady=(6, 0))
        self.barra = ttk.Progressbar(avance, maximum=100.0, variable=self.pct)
        self.barra.pack(side='left', fill='x', expand=True)
        self.etiqueta_pct = tk.Label(avance, textvariable=self.texto_pct, width=8, bg=FONDO_APP, fg=COLOR_TITULO)
        self.etiqueta_pct.pack(side='left', padx=(6, 0))
        cuerpo.add(der)

    # ------------------------------------------------------------ lo que necesitan los dialogos

    def post(self, fn):
        'Corre fn en el hilo de la ventana (desde cualquier hilo).'
        self.eventos.put(fn)

    def clave_cambiada(self):
        self._refrescar_modelos_remotos()
        self._refrescar_saldo()

    def reescanear_modelos(self):
        self.agente.local_models = localmodels.scan_models(localmodels.default_model_dirs(self.cfg))
        self._refrescar_modelos_widgets()
        return self.agente.local_models

    # ------------------------------------------------------------ modelos, esfuerzo y saldo

    def _al_arrancar(self):
        threading.Thread(target=self._buscar_locales, daemon=True).start()
        if self.cfg.api_key:
            self.clave_cambiada()
        elif self.cfg.needs_unlock:
            self.bus.log('La API key esta guardada con contraseña: te la pido para desbloquearla.', 'warn')
            self.raiz.after(300, lambda: dialogos.UnlockDialog(self))
        else:
            self.bus.log('Sin API key de DeepSeek: cargala en Configuracion, '
                         'o elegi un modelo [local].', 'warn')

    def _buscar_locales(self):
        try:
            hallados = localmodels.scan_models(localmodels.default_model_dirs(self.cfg))
        except Exception as exc:        # noqa: BLE001 - una carpeta ilegible no debe romper el arranque
            self.bus.log('No pude buscar modelos locales: ' + str(exc), 'warn')
            return

        def aplicar():
            self.agente.local_models = hallados
            self._refrescar_modelos_widgets()
            if hallados:
                self.bus.log(f'Modelos locales encontrados: {len(hallados)}', 'info')
        self.post(aplicar)

    def _refrescar_modelos_remotos(self):
        key = self.cfg.api_key
        if not key:
            return

        def trabajo():
            try:
                lista = dsapi.list_models(key)
            except dsapi.ApiError as exc:
                self.bus.log('No pude pedir la lista de modelos a DeepSeek: ' + str(exc), 'warn')
                return

            def aplicar():
                self.agente.remote_models = lista
                self._refrescar_modelos_widgets()
            self.post(aplicar)
        threading.Thread(target=trabajo, daemon=True).start()

    def _refrescar_saldo(self):
        key = self.cfg.api_key
        if not key:
            self.saldo_var.set('Saldo: sin key')
            return
        self.saldo_var.set('Saldo: ...')

        def trabajo():
            try:
                b = dsapi.get_balance(key)
                texto = 'Saldo: ' + b['text'] + ('' if b['available'] else ' (no disponible)')
            except dsapi.ApiError as exc:
                texto = 'Saldo: ?'
                self.bus.log('No pude consultar el saldo: ' + str(exc), 'warn')
            self.post(lambda: self.saldo_var.set(texto))
        threading.Thread(target=trabajo, daemon=True).start()

    def _refrescar_modelos_widgets(self):
        a = self.agente
        entradas = a.all_entries()
        actual = a.entry()
        if actual['id'] not in [e['id'] for e in entradas]:
            entradas.insert(0, actual)      # el elegido se conserva aunque ya no este en las listas
        self._display_a_id = {a.display(e): e['id'] for e in entradas}
        self.modelo_cb.configure(values=list(self._display_a_id))
        self.modelo_var.set(a.display(actual))
        self._refrescar_esfuerzo()

    def _refrescar_esfuerzo(self):
        entrada = self.agente.entry()
        esfuerzos = entrada.get('efforts') or []
        if entrada['kind'] == 'local' and not esfuerzos:
            self.esfuerzo_cb.configure(values=['-'], state='disabled')
            self.esfuerzo_var.set('-')
            return
        self.esfuerzo_cb.configure(values=[SIN_EFFORT] + esfuerzos, state='readonly')
        self.esfuerzo_var.set(self.agente.effort if self.agente.effort in esfuerzos else SIN_EFFORT)

    def _modelo_cambiado(self):
        mid = self._display_a_id.get(self.modelo_var.get())
        if self.en_pedido:
            self.bus.chat(BUSY_NOTE, 'aviso')
            self._refrescar_modelos_widgets()
            return
        if not mid or mid == self.agente.model_id:
            return
        self.agente.model_id = mid
        self.cfg['model'] = mid
        self._guardar_cfg()
        self._refrescar_esfuerzo()
        entrada = self.agente.entry()
        self.bus.log('Modelo: ' + self.agente.display(entrada), 'info')
        if entrada['kind'] == 'remote' and not self.cfg.api_key:
            self.bus.chat('Para usar ' + entrada['id'] + ' falta la API key de DeepSeek (boton Configuracion).', 'aviso')

    def _esfuerzo_cambiado(self):
        v = self.esfuerzo_var.get()
        self.agente.effort = '' if v in (SIN_EFFORT, '-') else v
        self.cfg['effort'] = self.agente.effort
        self._guardar_cfg()

    def _guardar_cfg(self):
        try:
            self.cfg.save()
        except OSError as exc:
            self.bus.log('No pude guardar la configuracion: ' + str(exc), 'warn')

    def _abrir_configuracion(self):
        dialogos.SettingsDialog(self)

    def _nueva_charla(self):
        if self.en_pedido:
            self.bus.chat(BUSY_NOTE, 'aviso')
            return
        self.agente.nueva_charla()
        self.chat.configure(state='normal')
        self.chat.delete('1.0', 'end')
        self.chat.configure(state='disabled')
        self.bus.log('Charla nueva: el asistente ya no recuerda lo anterior (el libro sigue abierto).', 'ok')

    # ------------------------------------------------------------ libro y adjuntos

    def _elegir(self):
        tipos = [('Libros de Excel', '*.xlsx *.xlsm *.xls *.xlsb'), ('Todos los archivos', '*.*')]
        ruta = filedialog.askopenfilename(title='Elegi el libro de Excel', filetypes=tipos)
        if not ruta:
            self.bus.log('No se eligio ningun archivo.', 'warn')
            return
        self.ruta = os.path.abspath(ruta)
        self.ruta_var.set(self.ruta)
        self.bus.log('Archivo designado: ' + self.ruta, 'ok')
        self.bus.progreso(0, 'Pendiente de abrir')
        self.peticiones.put(('abrir', self.ruta))

    def _adjuntar(self):
        rutas = filedialog.askopenfilenames(title='Adjuntar planillas, documentos, textos, contactos o imagenes',
                                            filetypes=adjuntos.TIPOS_DIALOGO)
        if not rutas:
            self.bus.log('No se adjunto ningun archivo.', 'warn')
            return
        for r in rutas:
            self.peticiones.put(('adjuntar', os.path.abspath(r)))

    # ------------------------------------------------------------ entrada

    def _texto_consigna(self):
        return self.entrada.get('1.0', 'end-1c')

    def _tecla_enter(self, evento=None):
        self._enviar()
        return 'break'

    def _tecla_shift(self, evento=None):
        self.entrada.insert('insert', chr(10))
        return 'break'

    def _ajustar_alto(self, evento=None):
        contenido = self._texto_consigna()
        lineas = contenido.count(chr(10)) + 1
        alto = max(ALTO_ENTRADA, min(lineas, ALTO_ENTRADA_MAX))
        actual = int(self.entrada.cget('height'))
        if actual != alto:
            self.entrada.configure(height=alto)
        self.entrada.edit_modified(False)

    def _enviar(self):
        if self.en_pedido:
            return
        texto = self._texto_consigna().strip()
        if not texto:
            self.bus.log('No escribiste ninguna consigna.', 'warn')
            return
        if self.agente.necesita_key() and self._stream_factory is None:
            if self.cfg.needs_unlock:
                dialogos.UnlockDialog(self)
            else:
                self.bus.chat('Falta la API key de DeepSeek: cargala con el boton Configuracion, '
                              'o elegi un modelo [local] arriba.', 'aviso')
            return
        self.entrada.delete('1.0', 'end')
        self.entrada.configure(height=ALTO_ENTRADA)
        self.bus.chat(texto, 'usuario')
        self.cancel_ev = threading.Event()
        self.en_pedido = True
        self.estado_var.set('Pensando...')
        self.peticiones.put(('pedir', (texto, self.cancel_ev)))

    def _ayuda(self):
        self.bus.chat(AYUDA)

    def _reiniciar_conexion(self):
        # Sirve cuando Excel quedo sin responder: sesion nueva en un hilo nuevo. La charla y los modelos se conservan.
        viejo = self.agente
        viejo.detener()
        nuevo = asistente.Asistente(self.cfg, self.bus, stream_factory=self._stream_factory)
        for campo in ('messages', 'totals', 'remote_models', 'local_models', 'model_id', 'effort', 'ediciones_permitidas'):
            setattr(nuevo, campo, getattr(viejo, campo))
        nuevo._server, nuevo._server_key, nuevo._cpt = viejo._server, viejo._server_key, viejo._cpt
        self.bus.log('Armando una conexion nueva con Excel.', 'warn')
        self.agente = nuevo
        self.peticiones = queue.Queue()
        self.ocupado_desde = None
        self.en_pedido = False
        self.ruta = None
        self.ruta_var.set('(sin archivo designado)')
        self._arrancar_trabajador()
        self.bus.log('Conexion nueva lista. Volve a elegir el libro con Abrir Excel.', 'ok')

    # ------------------------------------------------------------ hilo de trabajo (todo lo que toca Excel)

    def _arrancar_trabajador(self):
        cola = self.peticiones
        agente = self.agente
        self.trabajador = threading.Thread(target=self._bucle, args=(cola, agente), daemon=True)
        self.trabajador.start()

    def _bucle(self, cola, agente):
        # Este hilo no debe morir nunca: si muere, la interfaz deja de responder sin avisar.
        while True:
            try:
                tarea = cola.get()
                if tarea is None:
                    return
                self.ocupado_desde = time.time()
                self.ultimo_aviso = 0.0
                try:
                    clase, dato = tarea
                except Exception:
                    self.bus.log('Tarea con formato inesperado: ' + repr(tarea), 'error')
                    continue
                try:
                    if clase == 'abrir':
                        agente.abrir(dato, visible=True, respaldo=True, minimizar=True)
                    elif clase == 'adjuntar':
                        agente.adjuntar(dato)
                    elif clase == 'pedir':
                        texto, cancel = dato
                        r = {'outcome': None, 'error': None, 'restore': None}
                        try:
                            r = agente.pedir(texto, cancel=cancel, emit=self._emitir, confirm=self._confirmar)
                        finally:
                            self.post(lambda r=r: self._pedido_terminado(r))
                    elif clase == 'salir':
                        agente.salir(cerrar_excel=bool(dato))
                except Exception as exc:
                    self.bus.log('Error: ' + str(exc), 'error')
                    self.bus.chat('Hubo un problema: ' + str(exc), 'error')
            except Exception as exc:
                self.bus.log('Fallo inesperado: ' + str(exc), 'error')
            finally:
                self.ocupado_desde = None

    def _emitir(self, kind, *a):
        'Eventos del agente (hilo de trabajo) -> ventana.'
        self.post(lambda: self._evento_agente(kind, *a))

    def _confirmar(self, kind, title, detail):
        """Corre en el hilo de trabajo: pide permiso en la ventana y espera sin bloquearla."""
        ev, caja = threading.Event(), {}
        cancel = self.cancel_ev

        def listo(r):
            if 'r' not in caja:
                caja['r'] = r
                ev.set()

        def mostrar():
            try:
                self._dialogo = dialogos.ConfirmDialog(self.raiz, kind, title, detail, listo)
                self.estado_var.set('Esperando tu permiso...')
            except tk.TclError:
                listo('deny')
        self.post(mostrar)
        self.bus.log('Pide permiso: ' + title, 'warn')
        while not ev.wait(0.1):
            if self.cerrando or (cancel is not None and cancel.is_set()):
                self.post(self._cerrar_dialogo)
                return False
        self._dialogo = None
        if caja['r'] == 'allow_session':
            self.agente.permitir_ediciones()
            self.bus.log('Reemplazos permitidos para el resto de esta charla.', 'ok')
        self.bus.log('Permiso ' + ('concedido' if caja['r'] != 'deny' else 'denegado') + '.', 'ok' if caja['r'] != 'deny' else 'warn')
        return caja['r'] != 'deny'

    def _cerrar_dialogo(self):
        d, self._dialogo = self._dialogo, None
        if d is not None:
            try:
                d.dismiss()
            except tk.TclError:
                pass

    # ------------------------------------------------------------ eventos del agente (hilo de la ventana)

    def _evento_agente(self, kind, *a):
        self.ultimo_evento = time.time()
        if kind == 'step_begin':
            self.medidor.step_begin()
            self._paso_con_texto = False
            self._texto_paso = []
            self.estado_var.set('Pensando...')
        elif kind == 'reasoning':
            self.medidor.piece(a[0])
            self.estado_var.set('Razonando...')
        elif kind == 'content':
            self.medidor.piece(a[0])
            self.estado_var.set('Escribiendo...')
            self._escribir_en_vivo(a[0])
        elif kind == 'usage':
            self.medidor.usage(a[0])
        elif kind == 'step_end':
            self.medidor.step_end()
            p = a[0]
            if self._paso_con_texto:
                self._escribir_en_vivo(chr(10))
            self._anotar_respuesta()
            if p.get('finish') == 'length':
                self.bus.chat('La respuesta se corto por el limite de tokens (se ajusta en Configuracion, Avanzado).', 'aviso')
        elif kind == 'tool_start':
            self.estado_var.set('Trabajando en Excel: ' + str(a[1]) + '...')
            self.bus.log('-> ' + str(a[1]) + ' ' + recortar(a[2], 160), 'avance')
        elif kind == 'tool_result':
            r = str(a[2])
            nivel = 'warn' if r.startswith(('ERROR', 'DENIED')) else 'ok'
            self.bus.log(str(a[1]) + ': ' + recortar(r), nivel)
            self.estado_var.set('Pensando...')
        elif kind == 'notice':
            self.bus.chat(a[0], 'aviso')
        elif kind == 'status':
            self.estado_var.set(a[0])
            self.bus.log(a[0], 'info')

    def _escribir_en_vivo(self, texto):
        self.chat.configure(state='normal')
        if not self._paso_con_texto:
            self._paso_con_texto = True
            self.chat.insert('end', '[' + excel.hora_actual() + '] ', 'hora')
            self.chat.insert('end', 'ASISTENTE: ', 'agente')
        self.chat.insert('end', texto, 'agente')
        self.chat.see('end')
        self.chat.configure(state='disabled')
        self._texto_paso.append(texto)

    def _anotar_respuesta(self):
        # Antes se anotaba cada pieza del streaming como una linea '[asistente]' aparte (una por palabra).
        texto = ''.join(self._texto_paso).strip()
        self._texto_paso = []
        if texto:
            self._anotar('[asistente] ' + texto)

    def _pedido_terminado(self, r):
        self.medidor.step_end()          # un paso cortado no se pierde del acumulado
        self._anotar_respuesta()         # ni su texto en el registro
        self._cerrar_dialogo()
        self.en_pedido = False
        self.cancel_ev = None
        if r.get('error'):
            self.bus.chat(r['error'], 'error')
            self.bus.log(r['error'], 'error')
        elif r.get('outcome') == 'cancelled':
            self.bus.chat('Interrumpido.', 'aviso')
        if r.get('restore') and not self._texto_consigna().strip():
            self.entrada.insert('1.0', r['restore'])
        self.estado_var.set('Listo')
        if self.agente.entry()['kind'] == 'remote' and self.cfg.api_key:
            self._refrescar_saldo()

    def _actualizar_medidor(self):
        t = self.agente.totals
        total = self.medidor.total()
        presupuesto = max(1, int(self.cfg['token_budget']))
        self.consumo.configure(value=min(100.0, 100.0 * total / presupuesto))
        partes = ['Tokens de hoy: ' + ('~' if self.medidor.estimated() else '') + miles(total)
                  + f' ({100.0 * total / presupuesto:.1f}' + chr(37) + ' del presupuesto)']
        if t['in'] or t['out']:
            partes.append('entrada ' + miles(t['in']) + (' (cache ' + miles(t['hit']) + ')' if t['hit'] else ''))
            partes.append('salida ' + miles(t['out']) + (' (razonando ' + miles(t['reason']) + ')' if t['reason'] else ''))
        v = self.medidor.rate()
        if v is not None:
            partes.append(('~' if v[1] else '') + f'{v[0]:.0f} tok/s')
        self.tokens_var.set(' - '.join(partes))

    # ------------------------------------------------------------ bucle de la ventana

    def _procesar_eventos(self):
        # Si un evento falla, se informa y el bucle sigue: la ventana nunca queda muda.
        try:
            while True:
                try:
                    ev = self.eventos.get_nowait()
                except queue.Empty:
                    break
                try:
                    if callable(ev):
                        ev()
                    else:
                        self._mostrar_evento(ev)
                except Exception as exc:
                    self._error_interno('No pude mostrar un evento: ' + str(exc))
            self._actualizar_medidor()
        except Exception as exc:
            self._error_interno('Fallo el bucle de eventos: ' + str(exc))
        finally:
            self._actualizar_botones()
            self._agendar_eventos()

    def _agendar_eventos(self):
        # Reprograma el repaso de eventos; se detiene si la ventana ya se solto.
        if self.cerrando:
            return
        try:
            if self.tarea_after is not None:
                self.raiz.after_cancel(self.tarea_after)     # uno solo pendiente, aunque se llame a mano
            self.tarea_after = self.raiz.after(120, self._procesar_eventos)
        except Exception:
            self.cerrando = True

    def _al_soltar(self, evento=None):
        # Al soltarse la ventana se cancela el repaso pendiente.
        if evento is not None and str(evento.widget) != str(self.raiz):
            return
        self.cerrando = True
        if self.tarea_after is not None:
            try:
                self.raiz.after_cancel(self.tarea_after)
            except Exception:
                pass
            self.tarea_after = None

    def _vigilar_espera(self):
        # Avisa solo si no hay novedades: eso es lo que indica que algo quedo trabado.
        if self.ocupado_desde is None:
            return
        ahora = time.time()
        silencio = ahora - self.ultimo_evento
        if silencio < ESPERA_MAXIMA:
            return
        if ahora - self.ultimo_aviso < REAVISO:
            return
        self.ultimo_aviso = ahora
        self.bus.log('Sin novedades hace ' + str(int(silencio)) + ' segundos. Puede que el modelo este pensando o '
                     'cargandose; si sigue asi, revisa que Excel no tenga un cuadro de dialogo abierto o apreta Reconectar.', 'warn')

    def _mostrar_evento(self, ev):
        self.ultimo_evento = time.time()
        if ev.tipo == 'chat':
            self._escribir_chat(ev)
        else:
            self._escribir_log(ev)
        if ev.tipo == 'progreso':
            self.pct.set(ev.pct)
            self.texto_pct.set(str(int(ev.pct)) + ' ' + chr(37))
        self._anotar('[' + ev.nivel + '] ' + ev.texto)
        self._vigilar_espera()

    def _escribir_chat(self, ev):
        if ev.nivel == 'usuario':
            quien = 'VOS'
            etiqueta = 'usuario'
        elif ev.nivel in ('error', 'warn', 'aviso'):
            quien = 'ASISTENTE'
            etiqueta = 'aviso'
        else:
            quien = 'ASISTENTE'
            etiqueta = 'agente'
        self.chat.configure(state='normal')
        self.chat.insert('end', '[' + ev.hora + '] ', 'hora')
        self.chat.insert('end', quien + ': ', etiqueta)
        self.chat.insert('end', ev.texto + chr(10), etiqueta)
        self.chat.see('end')
        self.chat.configure(state='disabled')

    def _escribir_log(self, ev):
        nivel = ev.nivel if ev.nivel in ('info', 'avance', 'ok', 'warn', 'error') else 'info'
        self.log.configure(state='normal')
        self.log.insert('end', ev.hora + ' ', 'hora')
        self.log.insert('end', '[' + nivel + '] ', nivel)
        self.log.insert('end', ev.texto + chr(10), nivel)
        self.log.see('end')
        self.log.configure(state='disabled')

    def _cargar_iconos(self):
        # Iconos de enviar y detener, en la carpeta assets.
        try:
            if os.path.isfile(IMG_ENVIAR):
                self.img_enviar = tk.PhotoImage(file=IMG_ENVIAR)
            if os.path.isfile(IMG_DETENER):
                self.img_detener = tk.PhotoImage(file=IMG_DETENER)
        except Exception as exc:
            self.img_enviar = None
            self.img_detener = None
            self._error_interno('No pude cargar los iconos: ' + str(exc), 'warn')

    def _enviar_o_detener(self):
        # El mismo boton hace las dos cosas.
        if self.ocupado_desde is not None or self.en_pedido:
            self._detener()
            return
        self._enviar()

    def _detener(self):
        # Directo desde la ventana, sin pasar por la cola: la cola esta ocupada con el trabajo que hay que cortar.
        if self.cancel_ev is not None:
            self.cancel_ev.set()
        self.agente.detener()
        self._cerrar_dialogo()
        self.estado_var.set('Deteniendo...')
        self.bus.log('Pedido de detener enviado.', 'warn')

    def _actualizar_botones(self):
        # Cambia el icono y el texto segun haya trabajo en curso o no.
        trabajando = self.ocupado_desde is not None or self.en_pedido
        if trabajando == self.trabajando_antes:
            return
        self.trabajando_antes = trabajando
        if trabajando:
            imagen = self.img_detener
            texto = 'Detener'
        else:
            imagen = self.img_enviar
            texto = 'Enviar'
        if imagen is not None:
            self.boton_principal.configure(image=imagen, text=str())
        else:
            self.boton_principal.configure(text=texto)
        self.pista.configure(text='Trabajando. Apreta Detener para cortar el trabajo.' if trabajando else PISTA_TEXTO)

    def _al_cerrar(self):
        if self.ocupado_desde is not None or self.en_pedido:
            self._detener()
        # solo se pregunta si hay un Excel conectado; sin el, la pregunta no tiene sentido (y antes aparecia igual)
        cerrar = self.agente.sesion.app is not None and messagebox.askyesno('Salir', 'Cierro tambien la ventana de Excel?')
        self.peticiones.put(('salir', bool(cerrar)))
        # espera a que el hilo de trabajo termine de salir (cierra el libro y apaga el modelo local)
        fin = time.time() + 8
        while (not self.peticiones.empty() or self.ocupado_desde is not None) and time.time() < fin:
            self.raiz.update()
            time.sleep(0.05)
        self.agente.detener_servidor()      # por si el hilo de trabajo quedo trabado: el modelo local no queda vivo
        self._anotar('Sesion cerrada')
        self.raiz.destroy()


def main():
    ventana = Ventana()
    ventana.raiz.mainloop()


if __name__ == '__main__':
    main()
