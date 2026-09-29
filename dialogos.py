# Ventanas auxiliares: pedir permiso para un cambio, desbloquear la key y la configuracion.
# Siguen el modelo de DeepSeekChat (dialogs.py) con los colores de ExcelAgent.
#
# Ninguna es modal a la fuerza mientras el asistente trabaja: el agente espera la respuesta en otro hilo y el usuario
# tiene que poder apretar Detener en la ventana principal.
#
# La ventana principal que las abre tiene que ofrecer: raiz (tk.Tk), cfg (dsapi.Config), post(fn) para correr algo en
# el hilo de la ventana, clave_cambiada() y reescanear_modelos() -> lista de modelos locales.
from __future__ import annotations

import json
import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import dsapi
import localmodels

TITULO = 'ExcelAgent'
FONDO = '#CCE2F5'
PANEL = '#FFFFFF'
TEXTO = '#0B2545'
APAGADO = '#5B7FA6'
ERROR = '#B00020'
BIEN = '#14653A'
BOTON = '#A9CDEB'
AGREGA = '#14653A'
QUITA = '#B00020'

APROBACIONES = {'ask': 'Preguntar antes de pisar, borrar o guardar',
                'edits': 'Pisar datos sin preguntar (borrar y guardar, si)',
                'all': 'No preguntar nada'}


def _ubicar(win, raiz, dx=80, dy=60):
    win.update_idletasks()
    win.geometry(f'+{raiz.winfo_rootx() + dx}+{raiz.winfo_rooty() + dy}')


def _etiqueta(padre, texto, negrita=False, color=TEXTO, **kw):
    return tk.Label(padre, text=texto, bg=FONDO, fg=color, justify='left', anchor='w',
                    font=('Segoe UI', 10, 'bold') if negrita else ('Segoe UI', 9), **kw)


def _boton(padre, texto, orden, principal=False, **kw):
    return tk.Button(padre, text=texto, command=orden, bg=TEXTO if principal else BOTON,
                     fg='white' if principal else TEXTO, activebackground=BOTON, relief='flat', padx=10, pady=3, **kw)


def _marco(padre, **kw):
    return tk.Frame(padre, bg=FONDO, **kw)


# ---------------------------------------------------------------- pedir permiso

class ConfirmDialog:
    """Muestra que quiere hacer el asistente. on_result recibe 'allow', 'allow_session' o 'deny' (una sola vez)."""

    CABECERAS = {'edit': 'El asistente quiere reemplazar datos que ya estaban',
                 'delete': 'El asistente quiere borrar algo',
                 'save': 'El asistente quiere guardar el archivo'}

    def __init__(self, raiz, kind, title, detail, on_result):
        self.kind, self.on_result, self._hecho = kind, on_result, False
        self.win = w = tk.Toplevel(raiz)
        w.title('El asistente pide permiso')
        w.configure(bg=FONDO)
        w.transient(raiz)
        f = _marco(w, padx=14, pady=14)
        f.pack(fill='both', expand=True)
        _etiqueta(f, self.CABECERAS.get(kind, 'El asistente pide permiso'), negrita=True).pack(anchor='w')
        _etiqueta(f, title, color=APAGADO, wraplength=620).pack(anchor='w', pady=(2, 8))
        vista = tk.Text(f, wrap='word', width=80, height=8, bg=PANEL, fg=TEXTO, relief='flat', padx=8, pady=6,
                        font=('Segoe UI', 10))
        vista.insert('1.0', detail or '')
        vista.configure(state='disabled')
        vista.pack(fill='both', expand=True)
        bf = _marco(f)
        bf.pack(fill='x', pady=(12, 0))
        self.btn_si = _boton(bf, 'Permitir', lambda: self._fin('allow'), principal=True)
        self.btn_si.pack(side='left')
        if kind == 'edit':
            _boton(bf, 'Permitir todos los reemplazos de esta charla', lambda: self._fin('allow_session')).pack(
                side='left', padx=8)
        self.btn_no = _boton(bf, 'No permitir', lambda: self._fin('deny'))
        self.btn_no.pack(side='right')
        w.protocol('WM_DELETE_WINDOW', lambda: self._fin('deny'))
        w.bind('<Escape>', lambda e: self._fin('deny'))
        _ubicar(w, raiz)
        w.lift()
        # borrar o guardar: un Enter por accidente no debe ser un si
        (self.btn_si if kind == 'edit' else self.btn_no).focus_set()

    def _fin(self, resultado):
        if self._hecho:
            return
        self._hecho = True
        try:
            self.win.destroy()
        except tk.TclError:
            pass
        self.on_result(resultado)

    def dismiss(self):
        'Cierra negando (lo usa la ventana principal al detener).'
        self._fin('deny')


# ---------------------------------------------------------------- desbloquear la key

class UnlockDialog:
    def __init__(self, app):
        self.app = app
        self.win = w = tk.Toplevel(app.raiz)
        w.title('Desbloquear API key')
        w.configure(bg=FONDO)
        w.transient(app.raiz)
        w.resizable(False, False)
        f = _marco(w, padx=16, pady=16)
        f.pack()
        _etiqueta(f, 'La API key esta guardada con contraseña.', negrita=True).pack(anchor='w')
        _etiqueta(f, 'Sin ella solo funcionan los modelos locales.', color=APAGADO).pack(anchor='w', pady=(0, 8))
        self.pw = tk.StringVar()
        self.entrada = tk.Entry(f, textvariable=self.pw, show='•', width=34)
        self.entrada.pack(fill='x')
        self.msg = _etiqueta(f, '', color=ERROR)
        self.msg.pack(anchor='w', pady=(6, 0))
        bf = _marco(f)
        bf.pack(fill='x', pady=(10, 0))
        _boton(bf, 'Desbloquear', self.enviar, principal=True).pack(side='left')
        _boton(bf, 'Ahora no', w.destroy).pack(side='left', padx=8)
        _boton(bf, 'Olvide la contraseña...', self.olvide).pack(side='right')
        self.entrada.bind('<Return>', lambda e: self.enviar())
        w.bind('<Escape>', lambda e: w.destroy())
        _ubicar(w, app.raiz, 140, 120)
        w.grab_set()
        self.entrada.focus_set()

    def enviar(self):
        if self.app.cfg.unlock(self.pw.get()):
            self.win.destroy()
            self.app.clave_cambiada()
        else:
            self.msg.configure(text='Contraseña incorrecta.')
            self.pw.set('')

    def olvide(self):
        if messagebox.askyesno(TITULO, 'La key guardada no se puede recuperar sin la contraseña.\n\n'
                               'Borrarla para poder cargar una nueva? (La key sigue existiendo en tu cuenta de DeepSeek; '
                               'solo se borra la copia de este programa.)', parent=self.win):
            self.app.cfg.set_api_key('')
            self.win.destroy()
            self.app.clave_cambiada()


# ---------------------------------------------------------------- importar de DeepSeekChat

def config_deepseekchat():
    'Donde guarda DeepSeekChat su configuracion cuando esta instalado (no portable).'
    return os.path.join(os.environ.get('APPDATA') or os.path.expanduser('~'), 'DeepSeekChat', 'config.json')


def importar_de_deepseekchat(cfg, ruta=None):
    """Copia de la configuracion de DeepSeekChat la API key y los ajustes de modelos locales. La key nunca pasa por
    la pantalla: si alla esta cifrada con el usuario de Windows se descifra y se vuelve a cifrar aca; si esta con
    contraseña, se copia cifrada tal cual (se desbloquea con la misma contraseña). Devuelve un texto con lo hecho;
    lanza ValueError si no hay nada que importar."""
    ruta = ruta or config_deepseekchat()
    try:
        with open(ruta, 'r', encoding='utf-8') as f:
            otra = json.load(f)
    except FileNotFoundError:
        raise ValueError(f'No encontre la configuracion de DeepSeekChat en {ruta}.')
    except (OSError, ValueError) as e:
        raise ValueError(f'No pude leer {ruta}: {e}')
    if not isinstance(otra, dict):
        raise ValueError(f'{ruta} no tiene el formato esperado.')
    hecho = []
    if otra.get('api_key_pw'):
        cfg['api_key_enc'] = ''
        cfg['api_key_pw'] = otra['api_key_pw']
        cfg.set_session_key('')
        hecho.append('la API key (cifrada con contraseña: te la va a pedir para desbloquearla)')
    elif otra.get('api_key_enc'):
        try:
            key = dsapi.unprotect(otra['api_key_enc'])
        except (OSError, ValueError) as e:
            raise ValueError(f'La key de DeepSeekChat no se pudo descifrar con este usuario de Windows ({e}).')
        cfg.set_api_key(key)
        del key
        hecho.append('la API key (cifrada con tu usuario de Windows)')
    if otra.get('llama_server_path'):
        cfg['llama_server_path'] = str(otra['llama_server_path'])
        hecho.append('la ruta de llama-server')
    dirs = [d for d in (otra.get('model_dirs') or []) if isinstance(d, str) and d]
    nuevas = [d for d in dirs if d not in cfg['model_dirs']]
    if nuevas:
        cfg['model_dirs'] = list(cfg['model_dirs']) + nuevas
        hecho.append('las carpetas de modelos (' + ', '.join(nuevas) + ')')
    if isinstance(otra.get('local_ctx'), int) and otra['local_ctx'] >= 2048:
        cfg['local_ctx'] = otra['local_ctx']
        hecho.append(f"el contexto local ({otra['local_ctx']})")
    if not hecho:
        raise ValueError('DeepSeekChat no tiene ni key ni modelos locales configurados.')
    cfg.save()
    return 'Importe ' + '; '.join(hecho) + '.'


# ---------------------------------------------------------------- configuracion

class SettingsDialog:
    LOCAL = ('model_dirs', 'llama_server_path', 'local_ctx')

    def __init__(self, app):
        self.app = app
        self.win = w = tk.Toplevel(app.raiz)
        w.title('Configuracion')
        w.configure(bg=FONDO)
        w.transient(app.raiz)
        w.resizable(False, False)
        estilo = ttk.Style(w)
        estilo.configure('EA.TNotebook', background=FONDO)
        estilo.configure('EA.TFrame', background=FONDO)
        fuera = _marco(w, padx=12, pady=12)
        fuera.pack(fill='both', expand=True)
        nb = ttk.Notebook(fuera, style='EA.TNotebook')
        nb.pack(fill='both', expand=True)
        self.tab_key = _marco(nb, padx=14, pady=14)
        self.tab_local = _marco(nb, padx=14, pady=14)
        self.tab_mas = _marco(nb, padx=14, pady=14)
        nb.add(self.tab_key, text='API key')
        nb.add(self.tab_local, text='Modelos locales')
        nb.add(self.tab_mas, text='Avanzado')
        self._armar_key()
        self._armar_local()
        self._armar_mas()
        bf = _marco(fuera)
        bf.pack(fill='x', pady=(10, 0))
        _boton(bf, 'Abrir carpeta de datos', self.abrir_datos).pack(side='left')
        _boton(bf, 'Cerrar', self.cerrar, principal=True).pack(side='right')
        self.opt_msg = _etiqueta(fuera, 'Los cambios se guardan solos al cerrar esta ventana.', color=APAGADO,
                                 wraplength=560)
        self.opt_msg.pack(fill='x', pady=(6, 0))
        w.protocol('WM_DELETE_WINDOW', self.cerrar)
        self._foto = {k: self.app.cfg[k] for k in self.LOCAL}
        _ubicar(w, app.raiz)
        w.grab_set()
        self.key_entry.focus_set()

    def cerrar(self):
        'Cerrar guarda lo que se haya tocado.'
        if self.guardar(silencioso=True):
            if any(self.app.cfg[k] != self._foto[k] for k in self.LOCAL):
                self.app.reescanear_modelos()
        self.win.destroy()

    # ------------------------------------------------------------ API key

    def _armar_key(self):
        f, cfg = self.tab_key, self.app.cfg
        _etiqueta(f, 'API key de DeepSeek', negrita=True).grid(row=0, column=0, columnspan=3, sticky='w')
        _etiqueta(f, 'Se consigue en platform.deepseek.com, seccion API keys. El uso se paga con saldo prepago.',
                  color=APAGADO, wraplength=520).grid(row=1, column=0, columnspan=3, sticky='w')
        self.key_estado = _etiqueta(f, '', wraplength=520)
        self.key_estado.grid(row=2, column=0, columnspan=3, sticky='w', pady=(4, 6))
        self.key_var = tk.StringVar()
        self.key_entry = tk.Entry(f, textvariable=self.key_var, show='•', width=60)
        self.key_entry.grid(row=3, column=0, columnspan=3, sticky='we')
        self.ver_var = tk.BooleanVar(value=False)
        tk.Checkbutton(f, text='Mostrar', variable=self.ver_var, bg=FONDO, activebackground=FONDO,
                       command=lambda: self.key_entry.configure(show='' if self.ver_var.get() else '•')).grid(
            row=4, column=0, sticky='w', pady=4)

        _etiqueta(f, 'Donde guardarla', negrita=True).grid(row=5, column=0, columnspan=3, sticky='w', pady=(8, 2))
        self.modo = tk.StringVar(value='password' if cfg.portable or cfg.key_mode == 'password' else 'dpapi')
        opciones = (('dpapi', 'Cifrada con mi usuario de Windows (solo sirve en esta PC)'),
                    ('password', 'Cifrada con una contraseña (sirve en cualquier PC)'),
                    ('session', 'No guardarla: vale solo mientras el programa este abierto'))
        for i, (valor, texto) in enumerate(opciones):
            rb = tk.Radiobutton(f, text=texto, value=valor, variable=self.modo, bg=FONDO, activebackground=FONDO,
                                anchor='w', command=self._modo_cambiado)
            rb.grid(row=6 + i, column=0, columnspan=3, sticky='w')
            if valor == 'dpapi' and cfg.portable:
                rb.configure(state='disabled')
        self.pw1, self.pw2 = tk.StringVar(), tk.StringVar()
        _etiqueta(f, 'Contraseña').grid(row=9, column=0, sticky='w', pady=(6, 0))
        self.pw1_e = tk.Entry(f, textvariable=self.pw1, show='•', width=28)
        self.pw1_e.grid(row=9, column=1, sticky='w', pady=(6, 0))
        _etiqueta(f, 'Repetirla').grid(row=10, column=0, sticky='w')
        self.pw2_e = tk.Entry(f, textvariable=self.pw2, show='•', width=28)
        self.pw2_e.grid(row=10, column=1, sticky='w')

        bf = _marco(f)
        bf.grid(row=11, column=0, columnspan=3, sticky='we', pady=(10, 0))
        _boton(bf, 'Importar desde DeepSeekChat', self.importar).pack(side='left')
        self.btn_probar = _boton(bf, 'Probar y guardar', self.probar_y_guardar, principal=True)
        self.btn_probar.pack(side='right')
        _boton(bf, 'Borrar key', self.borrar_key).pack(side='right', padx=6)
        self.key_msg = _etiqueta(f, '', wraplength=520)
        self.key_msg.grid(row=12, column=0, columnspan=3, sticky='w', pady=(6, 0))
        self._modo_cambiado()
        self._estado_key()

    def _modo_cambiado(self):
        on = self.modo.get() == 'password'
        for e in (self.pw1_e, self.pw2_e):
            e.configure(state='normal' if on else 'disabled')

    def _estado_key(self):
        cfg = self.app.cfg
        if cfg.api_key:
            como = {'password': 'guardada con contraseña', 'dpapi': 'guardada, cifrada con tu usuario de Windows',
                    'none': 'solo en memoria (no esta guardada)'}[cfg.key_mode]
            texto, color = f'En uso: {dsapi.Config.mask(cfg.api_key)} - {como}.', TEXTO
        elif cfg.needs_unlock:
            texto, color = 'Hay una key guardada con contraseña, todavia bloqueada.', TEXTO
        elif cfg.has_key():
            texto, color = 'Hay una key guardada pero no se pudo descifrar; carga una nueva.', ERROR
        else:
            texto, color = 'No hay key cargada.', TEXTO
        self.key_estado.configure(text=texto, fg=color)

    def _msg(self, texto, error=False):
        self.key_msg.configure(text=texto, fg=ERROR if error else BIEN)

    def probar_y_guardar(self):
        key = self.key_var.get().strip()
        if not key:
            self._msg('Pega una key primero.', error=True)
            return
        modo, pw = self.modo.get(), self.pw1.get()
        if modo == 'password':
            if len(pw) < 6:
                self._msg('La contraseña necesita al menos 6 caracteres.', error=True)
                return
            if pw != self.pw2.get():
                self._msg('Las dos contraseñas no coinciden.', error=True)
                return
        self.btn_probar.configure(state='disabled')
        self._msg('Probando contra DeepSeek...')

        def trabajo():
            try:
                res = (dsapi.get_balance(key), None)
            except dsapi.ApiError as e:
                res = (None, e)
            self.app.post(lambda: self._probado(key, modo, pw, *res))
        threading.Thread(target=trabajo, daemon=True).start()

    def _probado(self, key, modo, pw, saldo, err):
        try:
            self.btn_probar.configure(state='normal')
        except tk.TclError:
            return      # la ventana se cerro mientras se probaba
        if err is not None:
            self._msg(f'No se guardo. {err}', error=True)
            return
        if not saldo['available']:
            self._msg('La key es valida pero DeepSeek informa la cuenta como no disponible. No se guardo.', error=True)
            return
        cfg = self.app.cfg
        try:
            if modo == 'session':
                cfg.set_session_key(key)
            else:
                cfg.set_api_key(key, pw if modo == 'password' else None)
        except (OSError, ValueError) as e:
            self._msg(f'No se pudo guardar: {e}', error=True)
            return
        self.key_var.set('')
        self.pw1.set('')
        self.pw2.set('')
        self._estado_key()
        msg = f"Key valida y {'cargada para esta sesion' if modo == 'session' else 'guardada'}. Saldo: {saldo['text']}"
        if modo == 'session' and cfg.has_key():
            msg += '\nOjo: la copia guardada antes sigue en disco; "Borrar key" la elimina.'
        self._msg(msg)
        self.app.clave_cambiada()

    def borrar_key(self):
        if messagebox.askyesno(TITULO, 'Borrar la API key guardada?', parent=self.win):
            self.app.cfg.set_api_key('')
            self._estado_key()
            self._msg('Key borrada.')
            self.app.clave_cambiada()

    def importar(self):
        cfg = self.app.cfg
        if cfg.has_key() and not messagebox.askyesno(
                TITULO, 'Ya hay una key guardada en ExcelAgent. Reemplazarla por la de DeepSeekChat?', parent=self.win):
            return
        try:
            texto = importar_de_deepseekchat(cfg)
        except (ValueError, OSError) as e:
            self._msg(str(e), error=True)
            return
        self._cargar_local()
        self._estado_key()
        self._msg(texto)
        self.app.clave_cambiada()
        if cfg.needs_unlock:
            UnlockDialog(self.app)

    # ------------------------------------------------------------ modelos locales

    def _armar_local(self):
        f = self.tab_local
        _etiqueta(f, 'Modelos locales (.gguf)', negrita=True).grid(row=0, column=0, columnspan=3, sticky='w')
        _etiqueta(f, 'Corren en esta PC con llama-server, sin internet y sin costo. Necesitan memoria: un modelo de '
                  '4 GB pide unos 6 GB de RAM libres. Son mas lentos y se equivocan mas que DeepSeek.',
                  color=APAGADO, wraplength=540).grid(row=1, column=0, columnspan=3, sticky='w', pady=(2, 6))
        _etiqueta(f, 'Se buscan siempre en:\n' + '\n'.join(localmodels.default_model_dirs()), color=APAGADO,
                  wraplength=540).grid(row=2, column=0, columnspan=3, sticky='w', pady=(0, 8))
        _etiqueta(f, 'Carpetas adicionales').grid(row=3, column=0, columnspan=3, sticky='w')
        self.dirs = tk.Listbox(f, height=4, width=72, bg=PANEL, fg=TEXTO, relief='flat')
        self.dirs.grid(row=4, column=0, columnspan=3, sticky='we')
        bf = _marco(f)
        bf.grid(row=5, column=0, columnspan=3, sticky='w', pady=4)
        _boton(bf, 'Agregar carpeta...', self.agregar_dir).pack(side='left')
        _boton(bf, 'Quitar', self.quitar_dir).pack(side='left', padx=6)
        _etiqueta(f, 'llama-server.exe (vacio = buscar junto al programa)').grid(
            row=6, column=0, columnspan=3, sticky='w', pady=(10, 0))
        self.srv_var = tk.StringVar()
        srv = tk.Entry(f, textvariable=self.srv_var, width=60)
        srv.grid(row=7, column=0, columnspan=2, sticky='we')
        for ev in ('<FocusOut>', '<Return>'):
            srv.bind(ev, lambda e: self._estado_srv())
        _boton(f, 'Examinar...', self.elegir_srv).grid(row=7, column=2, padx=(6, 0))
        self.srv_estado = _etiqueta(f, '', wraplength=540)
        self.srv_estado.grid(row=8, column=0, columnspan=3, sticky='w', pady=(2, 0))
        _etiqueta(f, 'Contexto del modelo local (tokens)').grid(row=9, column=0, sticky='w', pady=(10, 0))
        self.ctx_var = tk.IntVar()
        tk.Spinbox(f, from_=2048, to=131072, increment=2048, textvariable=self.ctx_var, width=9).grid(
            row=9, column=1, sticky='w', pady=(10, 0))
        _etiqueta(f, 'Mas contexto = mas memoria. 16384 es un punto medio razonable.', color=APAGADO).grid(
            row=10, column=0, columnspan=3, sticky='w')
        _boton(f, 'Buscar modelos ahora', self.buscar_ya).grid(row=11, column=0, sticky='w', pady=(10, 0))
        self.buscar_lbl = _etiqueta(f, '', wraplength=420)
        self.buscar_lbl.grid(row=11, column=1, columnspan=2, sticky='w', pady=(10, 0), padx=8)
        self._cargar_local()

    def _cargar_local(self):
        c = self.app.cfg
        self.dirs.delete(0, 'end')
        for d in c['model_dirs']:
            self.dirs.insert('end', d)
        self.srv_var.set(c['llama_server_path'])
        self.ctx_var.set(int(c['local_ctx']))
        self._estado_srv()

    def _estado_srv(self):
        p = localmodels.find_llama_server(self.srv_var.get().strip())
        self.srv_estado.configure(
            text=f'Encontrado: {p}' if p else 'No se encontro llama-server.exe: los modelos locales no van a poder arrancar.',
            fg=APAGADO if p else ERROR)

    def agregar_dir(self):
        d = filedialog.askdirectory(parent=self.win, title='Carpeta con modelos .gguf')
        if d and os.path.normpath(d) not in self.dirs.get(0, 'end'):
            self.dirs.insert('end', os.path.normpath(d))

    def quitar_dir(self):
        for i in reversed(self.dirs.curselection()):
            self.dirs.delete(i)

    def elegir_srv(self):
        p = filedialog.askopenfilename(parent=self.win, title='llama-server.exe',
                                       filetypes=[('llama-server', 'llama-server.exe'), ('Todos', '*.*')])
        if p:
            self.srv_var.set(os.path.normpath(p))
            self._estado_srv()

    def buscar_ya(self):
        if not self.guardar(silencioso=True):
            return
        self._foto = {k: self.app.cfg[k] for k in self.LOCAL}
        hallados = self.app.reescanear_modelos()
        self.buscar_lbl.configure(text=f'{len(hallados)} modelo(s) encontrado(s).' if hallados
                                  else 'No se encontro ningun modelo .gguf.')
        self._estado_srv()

    # ------------------------------------------------------------ avanzado

    def _armar_mas(self):
        f, cfg = self.tab_mas, self.app.cfg
        _etiqueta(f, 'Permisos', negrita=True).grid(row=0, column=0, columnspan=2, sticky='w')
        self.aprob_var = tk.StringVar(value=APROBACIONES.get(cfg['approval'], APROBACIONES['ask']))
        ttk.Combobox(f, textvariable=self.aprob_var, values=list(APROBACIONES.values()), state='readonly',
                     width=48).grid(row=1, column=0, columnspan=2, sticky='w', pady=(2, 10))
        _etiqueta(f, 'Maximo de tokens por respuesta').grid(row=2, column=0, sticky='w')
        self.max_var = tk.IntVar(value=int(cfg['max_tokens']))
        tk.Spinbox(f, from_=1024, to=393216, increment=1024, textvariable=self.max_var, width=9).grid(
            row=2, column=1, sticky='w', padx=6)
        _etiqueta(f, 'Presupuesto de tokens de la jornada (100% de la barra de consumo)').grid(
            row=3, column=0, sticky='w', pady=(10, 0))
        self.presu_var = tk.IntVar(value=int(cfg['token_budget']))
        tk.Spinbox(f, from_=10000, to=100000000, increment=100000, textvariable=self.presu_var, width=11).grid(
            row=3, column=1, sticky='w', padx=6, pady=(10, 0))
        _etiqueta(f, f'Datos (configuracion y registros del modelo local): {cfg.dir}'
                  + ('   [modo portable]' if cfg.portable else ''), color=APAGADO, wraplength=540).grid(
            row=4, column=0, columnspan=2, sticky='w', pady=(14, 0))

    # ------------------------------------------------------------ guardar

    def guardar(self, silencioso=False):
        c = self.app.cfg
        try:
            c['max_tokens'] = max(256, min(393216, int(self.max_var.get())))
            c['token_budget'] = max(10000, int(self.presu_var.get()))
            ctx = max(2048, min(131072, int(self.ctx_var.get())))
        except (tk.TclError, ValueError):
            self.opt_msg.configure(text='Hay un valor numerico invalido: no se guardo.', fg=ERROR)
            return False
        c['local_ctx'] = ctx
        c['model_dirs'] = list(self.dirs.get(0, 'end'))
        c['llama_server_path'] = self.srv_var.get().strip()
        c['approval'] = next((k for k, v in APROBACIONES.items() if v == self.aprob_var.get()), 'ask')
        try:
            c.save()
        except OSError as e:
            self.opt_msg.configure(text=f'No se pudo guardar: {e}', fg=ERROR)
            return False
        if not silencioso:
            self.opt_msg.configure(text='Opciones guardadas.', fg=BIEN)
        return True

    def abrir_datos(self):
        os.makedirs(self.app.cfg.dir, exist_ok=True)
        os.startfile(self.app.cfg.dir)
