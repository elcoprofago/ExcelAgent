# Interfaz grafica con cuadro de texto, boton de apertura, registro y avance.
from __future__ import annotations

import os
import queue
import sys
import threading
import time
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText

import app
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
CARPETA_ASSETS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'assets')
IMG_ENVIAR = os.path.join(CARPETA_ASSETS, 'b-env.png')
IMG_DETENER = os.path.join(CARPETA_ASSETS, 'b-stop.png')
PISTA_TEXTO = 'Enter envia - Shift+Enter baja de linea - podes pegar varias consignas, una por linea'
ALTO_ENTRADA = 3
ALTO_ENTRADA_MAX = 10
ESPERA_MAXIMA = 25
REAVISO = 30
RUTA_LOGS = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'logs')
MARCA = '------------------------------------------------------------'


class Ventana:
    'Interfaz principal.'

    def __init__(self):
        self.raiz = tk.Tk()
        self.raiz.title('ExcelAgent - libros de Excel')
        self.raiz.geometry('1040x720')
        self.raiz.configure(bg=FONDO_APP)
        self.peticiones = queue.Queue()
        self.eventos = queue.Queue()
        self.bus = app.Bus(sink=self.eventos.put)
        self.agente = app.Agente(self.bus)
        self.ruta = None
        self.ocupado_desde = None
        self.ultimo_aviso = 0.0
        self.ultimo_evento = time.time()
        self.archivo_log = None
        self.pct = tk.DoubleVar(value=0.0)
        self.texto_pct = tk.StringVar(value='0 ' + chr(37))
        self.ruta_var = tk.StringVar(value='(sin archivo designado)')
        self.img_enviar = None
        self.img_detener = None
        self.trabajando_antes = False
        self._cargar_iconos()
        self.raiz.minsize(720, 420)
        self._preparar_log()
        self._construir()
        self._instalar_manejadores()
        self._arrancar_trabajador()
        self.raiz.after(120, self._procesar_eventos)
        self.raiz.protocol('WM_DELETE_WINDOW', self._al_cerrar)
        self.bus.log('Interfaz lista. Elegi un libro con el boton Abrir Excel.', 'ok')

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
                f.write(app.hora_actual() + ' ' + str(texto) + chr(10))
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
            self.log.insert('end', app.hora_actual() + ' ', 'hora')
            self.log.insert('end', '[' + nivel + '] ', nivel)
            self.log.insert('end', texto + chr(10), nivel)
            self.log.see('end')
            self.log.configure(state='disabled')
        except Exception:
            pass

    def _construir(self):
        barra = tk.Frame(self.raiz, bg=FONDO_APP, padx=10, pady=8)
        barra.pack(fill='x')
        tk.Button(barra, text='Abrir Excel', width=13, bg=FONDO_BOTON, command=self._elegir).pack(side='left')
        caja = tk.Entry(barra, textvariable=self.ruta_var, state='readonly')
        caja.pack(side='left', fill='x', expand=True, padx=(10, 0))
        tk.Button(barra, text='+', width=3, bg=FONDO_BOTON, command=self._adjuntar).pack(side='left', padx=(10, 0))
        tk.Button(barra, text='Ayuda', width=8, bg=FONDO_BOTON, command=self._ayuda).pack(side='left', padx=(8, 0))
        tk.Button(barra, text='Reconectar', width=11, bg=FONDO_BOTON, command=self._reiniciar_conexion).pack(side='left', padx=(8, 0))

        cuerpo = tk.PanedWindow(self.raiz, orient='horizontal', sashwidth=6, bg=FONDO_APP)
        cuerpo.pack(fill='both', expand=True, padx=10, pady=(0, 10))

        izq = tk.LabelFrame(cuerpo, text='Chat con el asistente', bg=FONDO_APP, fg=COLOR_TITULO, padx=6, pady=6)
        self.chat = ScrolledText(izq, wrap='word', state='disabled', height=20, bg=FONDO_CHAT, fg=COLOR_AGENTE, relief='flat')
        self.chat.tag_config('hora', foreground='#5B7FA6')
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
        tipos = [('Textos', '*.txt *.csv *.tsv *.log *.md'), ('Imagenes', '*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.webp'), ('Todos los archivos', '*.*')]
        rutas = filedialog.askopenfilenames(title='Adjuntar textos o imagenes', filetypes=tipos)
        if not rutas:
            self.bus.log('No se adjunto ningun archivo.', 'warn')
            return
        for r in rutas:
            self.peticiones.put(('adjuntar', os.path.abspath(r)))

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
        texto = self._texto_consigna().strip()
        self.entrada.delete('1.0', 'end')
        self.entrada.configure(height=ALTO_ENTRADA)
        if not texto:
            self.bus.log('No escribiste ninguna consigna.', 'warn')
            return
        if self.ruta is None:
            self.bus.chat('Primero elegi un archivo con el boton Abrir Excel.', 'warn')
            return
        self.peticiones.put(('pedir', texto))

    def _ayuda(self):
        self.bus.chat(app.ayuda())

    def _reiniciar_conexion(self):
        # Sirve cuando Excel quedo sin responder: se arma una conexion nueva en un hilo nuevo.
        def trabajo():
            self.bus.log('Armando una conexion nueva con Excel.', 'warn')
            self.agente = app.Agente(self.bus)
            self.peticiones = queue.Queue()
            self.ocupado_desde = None
            self.raiz.after(0, self._arrancar_trabajador)
            self.bus.log('Conexion nueva lista. Volve a elegir el libro con Abrir Excel.', 'ok')
        threading.Thread(target=trabajo, daemon=True).start()

    def _arrancar_trabajador(self):
        cola = self.peticiones
        self.trabajador = threading.Thread(target=self._bucle, args=(cola,), daemon=True)
        self.trabajador.start()

    def _bucle(self, cola):
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
                        self.agente.abrir(dato, visible=True, respaldo=True, minimizar=True)
                    elif clase == 'adjuntar':
                        self.agente.adjuntar(dato)
                    elif clase == 'pedir':
                        self.agente.pedir(dato)
                    elif clase == 'detener':
                        self.agente.detener()
                    elif clase == 'salir':
                        self.agente.salir(cerrar_excel=bool(dato))
                except Exception as exc:
                    self.bus.log('Error: ' + str(exc), 'error')
                    self.bus.chat('Hubo un problema: ' + str(exc), 'error')
            except Exception as exc:
                self.bus.log('Fallo inesperado: ' + str(exc), 'error')
            finally:
                self.ocupado_desde = None

    def _procesar_eventos(self):
        # Si un evento falla, se informa y el bucle sigue: la ventana nunca queda muda.
        try:
            while True:
                try:
                    ev = self.eventos.get_nowait()
                except queue.Empty:
                    break
                try:
                    self._mostrar_evento(ev)
                except Exception as exc:
                    self._error_interno('No pude mostrar un evento: ' + str(exc))
        except Exception as exc:
            self._error_interno('Fallo el bucle de eventos: ' + str(exc))
        finally:
            self._actualizar_botones()
            self.raiz.after(120, self._procesar_eventos)

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
        self.bus.log('Sin novedades hace ' + str(int(silencio)) + ' segundos. Si sigue asi, revisa que Excel no tenga un cuadro de dialogo abierto o apreta Reconectar.', 'warn')

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
        # Iconos de enviar y detener, traidos de DeepSeekChat a la carpeta assets.
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
        # El mismo boton hace las dos cosas, como en DeepSeekChat.
        if self.ocupado_desde is not None:
            self.peticiones.put(('detener', None))
            return
        self._enviar()

    def _actualizar_botones(self):
        # Cambia el icono y el texto segun haya trabajo en curso o no.
        trabajando = self.ocupado_desde is not None
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
        cerrar = messagebox.askyesno('Salir', 'Cierro tambien la ventana de Excel?')
        self.peticiones.put(('salir', bool(cerrar)))
        fin = time.time() + 5
        while not self.peticiones.empty() and time.time() < fin:
            self.raiz.update()
            time.sleep(0.05)
        self._anotar('Sesion cerrada')
        self.raiz.destroy()


def main():
    ventana = Ventana()
    ventana.raiz.mainloop()


if __name__ == '__main__':
    main()
