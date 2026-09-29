# El asesor sin ventana: une la sesion de Excel, las herramientas y el modelo (DeepSeek o uno local).
# La ventana (gui.py) lo usa desde su hilo de trabajo; los tests, directamente, con un modelo falso.
#
# Todo lo que toca Excel (abrir, adjuntar, pedir, salir) tiene que correr en UN mismo hilo: los objetos COM no se
# pueden cruzar entre hilos. Lo unico que se llama desde otro hilo es detener(), que no toca Excel.
from __future__ import annotations

import json
import os
import threading

import agente as ag
import dsapi
import excel
import herramientas
import localmodels
import prompts

REMOTE_BUDGET = 600000      # caracteres de historial que se mandan a DeepSeek (su ventana sobra)
LOCAL_TIMEOUT = 900
MAX_STATE_SHEETS = 30


def nuevos_totales():
    return {'in': 0, 'out': 0, 'reason': 0, 'hit': 0}


class Asistente:
    def __init__(self, cfg, bus, stream_factory=None, max_steps=ag.DEFAULT_MAX_STEPS):
        self.cfg = cfg
        self.bus = bus
        self.sesion = excel.Sesion(bus)
        self.messages = []
        self.totals = nuevos_totales()
        self.remote_models = list(dsapi.FALLBACK_MODELS)
        self.local_models = []
        self.model_id = cfg['model']
        self.effort = cfg['effort']
        self.max_steps = max_steps
        self.agent = None
        self.toolbox = None
        self.cancel_ev = None
        self.ediciones_permitidas = False
        self._stream_factory = stream_factory    # (entry, msgs, tools) -> stream; solo para tests
        self._server = None
        self._server_key = None
        self._cpt = {}                           # caracteres por token medidos, por modelo local

    # ------------------------------------------------------------ libro y adjuntos (hilo de Excel)

    @property
    def ruta(self):
        return self.sesion.ruta

    def abrir(self, ruta, visible=True, respaldo=True, minimizar=True):
        if self.sesion.wb is not None and os.path.abspath(ruta).lower() != str(self.sesion.ruta or '').lower():
            # otro libro: el anterior se cierra solo si no tiene cambios; si los tiene, queda en Excel para el usuario
            if self.sesion.cambios_pendientes():
                self.sesion.entregar_al_usuario()
            else:
                self.sesion.cerrar(False)
        hojas = self.sesion.abrir(ruta, visible=visible, respaldo=respaldo, minimizar=minimizar)
        self.bus.log('Hojas: ' + ', '.join(hojas), 'info')
        self.bus.chat('Listo, abri ' + os.path.basename(ruta) + '. Contame que queres hacer con el.')
        return hojas

    def adjuntar(self, ruta):
        info = excel.procesar_adjunto(ruta)
        self.sesion.adjuntos.append(info)
        partes = info.get('partes') or []
        otras = f", en {len(partes)} partes: " + ', '.join(p['nombre'] for p in partes) if len(partes) > 1 else ''
        self.bus.log(f"Adjunto {len(self.sesion.adjuntos)}: {info['nombre']} ({info['clase']}, "
                     f"{len(info['filas'])} fila(s){otras})", 'ok')
        self.bus.chat(f"Recibi el adjunto {info['nombre']}. Decime que hago con el (por ejemplo, pasarlo a una hoja).")
        return info

    def estado_libro(self):
        """Lo que el modelo necesita saber del libro en cada pedido, en pocas lineas. Vacio si no hay libro."""
        s = self.sesion
        if s.wb is None:
            extra = ''
            if s.adjuntos:
                extra = 'Attachments loaded by the user: ' + ', '.join(
                    f"{i}. {a['nombre']}" for i, a in enumerate(s.adjuntos, 1))
            return extra
        try:
            lines = [f'File: {s.ruta}']
            activa = str(s.wb.ActiveSheet.Name)
            hojas = list(s.wb.Worksheets)
            for ws in hojas[:MAX_STATE_SHEETS]:
                ur = ws.UsedRange
                usado = str(ur.Address).replace('$', '') if int(ws.Application.WorksheetFunction.CountA(ur)) else 'empty'
                lines.append(f"- sheet '{ws.Name}': used {usado}" + (' (active)' if ws.Name == activa else ''))
            if len(hojas) > MAX_STATE_SHEETS:
                lines.append(f'- ... and {len(hojas) - MAX_STATE_SHEETS} more sheets (see workbook_info)')
            if not s.wb.Saved:
                lines.append('Unsaved changes: yes')
            if s.adjuntos:
                lines.append('Attachments loaded by the user: ' + ', '.join(
                    f"{i}. {a['nombre']}" for i, a in enumerate(s.adjuntos, 1)))
            return '\n'.join(lines)
        except Exception as e:      # noqa: BLE001 - el pedido sigue; el modelo puede llamar a workbook_info
            return f'File: {s.ruta}\n(Could not read the workbook state: {herramientas.com_text(e)})'

    # ------------------------------------------------------------ modelos

    def entry_for(self, mid):
        """El modelo (remoto o local) de un id guardado. Si ya no esta en las listas se fabrica uno."""
        for m in self.remote_models:
            if m['id'] == mid:
                return {**m, 'kind': 'remote'}
        for m in self.local_models:
            if m['id'] == mid:
                return self.local_entry(m)
        if mid.startswith(localmodels.MODEL_ID_PREFIX):
            path = mid[len(localmodels.MODEL_ID_PREFIX):]
            base = os.path.basename(path)
            if not os.path.isfile(path):
                same = [m for m in self.local_models if os.path.basename(m['path']).lower() == base.lower()]
                if len(same) == 1:
                    return self.local_entry(same[0])
            return {'id': mid, 'kind': 'local', 'name': base[:-5] if base.lower().endswith('.gguf') else base,
                    'path': path, 'size': 0, 'efforts': [], 'context': int(self.cfg['local_ctx']),
                    'missing': not os.path.isfile(path)}
        return {'id': mid, 'kind': 'remote', 'name': mid, 'efforts': [], 'context': 0}

    def local_entry(self, m):
        efforts = list(localmodels.REASONING_CHOICES) if localmodels.reasoning_for(m['path']) is not None else []
        return {**m, 'kind': 'local', 'efforts': efforts, 'context': int(self.cfg['local_ctx'])}

    @staticmethod
    def display(entry):
        if entry['kind'] == 'remote':
            return entry['id']
        if entry.get('missing'):
            return f"[local] {entry['name']} (no encontrado)"
        return f"[local] {entry['name']}" + (f" ({localmodels.format_size(entry['size'])})" if entry.get('size') else '')

    def all_entries(self):
        return [{**m, 'kind': 'remote'} for m in self.remote_models] + [self.local_entry(m) for m in self.local_models]

    def entry(self):
        return self.entry_for(self.model_id)

    def necesita_key(self):
        return self.entry()['kind'] == 'remote' and not self.cfg.api_key

    # ------------------------------------------------------------ modelo en marcha

    def local_max_tokens(self):
        """Tope de respuesta de un modelo local; es tambien lo que se le reserva dentro de su ventana."""
        return max(512, min(int(self.cfg['max_tokens']), int(self.cfg['local_ctx']) // 3))

    def make_stream(self, entry, base, api_msgs, tools):
        msgs = [{'role': 'system', 'content': prompts.system_prompt()}] + api_msgs
        if self._stream_factory is not None:
            return self._stream_factory(entry, msgs, tools)
        if entry['kind'] == 'local':
            return dsapi.ChatStream(None, entry['name'], msgs, None, self.local_max_tokens(), base=base, tools=tools,
                                    timeout=LOCAL_TIMEOUT)
        effort = self.effort if self.effort in (entry['efforts'] or []) else None
        return dsapi.ChatStream(self.cfg.api_key, entry['id'], msgs, effort, int(self.cfg['max_tokens']), tools=tools)

    def agent_budget(self, entry, toolbox):
        if entry['kind'] != 'local':
            return {'budget_chars': REMOTE_BUDGET}
        overhead = len(prompts.system_prompt()) + len(json.dumps(toolbox.specs, ensure_ascii=False))
        return {'ctx_tokens': int(self.cfg['local_ctx']), 'reserve_tokens': self.local_max_tokens(),
                'overhead_chars': overhead,
                'chars_per_token': self._cpt.get(entry.get('path') or entry['name'], ag.DEFAULT_CHARS_PER_TOKEN)}

    def local_server(self):
        exe = localmodels.find_llama_server(self.cfg['llama_server_path'])
        if not exe:
            raise localmodels.LocalError('No se encontro llama-server.exe (Configuracion -> Modelos locales).')
        key = (exe, int(self.cfg['local_ctx']))
        if self._server is None or self._server_key != key:
            if self._server is not None:
                self._server.stop()
            self._server = localmodels.LocalServer(exe, os.path.join(self.cfg.dir, 'logs'), key[1])
            self._server_key = key
        return self._server

    # ------------------------------------------------------------ un pedido

    def pedir(self, texto, cancel=None, emit=None, confirm=None):
        """Corre el agente hasta la respuesta final. emit(kind, *args) recibe los eventos del agente (en este hilo).
        Devuelve {'outcome', 'error', 'restore'}; restore es el texto a devolver al campo de entrada cuando el pedido
        no llego a tener respuesta (error o detenido antes de empezar)."""
        cancel = cancel or threading.Event()
        self.cancel_ev = cancel
        emit = emit or (lambda *a: None)
        entry = self.entry()
        texto = texto.strip()
        n0 = len(self.messages)
        outcome, error, agent = None, None, None

        def relay(kind, *a):
            if kind == 'usage':
                u = a[0]
                self.totals['in'] += u.get('prompt_tokens', 0)
                self.totals['out'] += u.get('completion_tokens', 0)
                self.totals['reason'] += (u.get('completion_tokens_details') or {}).get('reasoning_tokens', 0)
                self.totals['hit'] += u.get('prompt_cache_hit_tokens', 0)
            emit(kind, *a)

        try:
            if entry['kind'] == 'remote' and not self.cfg.api_key and self._stream_factory is None:
                raise dsapi.ApiError('Falta la API key de DeepSeek: cargala en Configuracion (o elegi un modelo local).')
            self.toolbox = herramientas.ToolBox(self.sesion, confirm=confirm, approval=self.aprobacion())
            self.messages.append({'role': 'user', 'content': prompts.with_state(texto, self.estado_libro())})
            base = None
            if entry['kind'] == 'local' and self._stream_factory is None:
                emit('status', 'Cargando el modelo local (puede tardar un rato)...')
                if entry.get('missing'):
                    raise localmodels.LocalError(f"No se encuentra el modelo {entry['path']}.")
                base = self.local_server().ensure(entry['path'], int(self.cfg['local_ctx']), cancel,
                                                  reasoning=localmodels.reasoning_for(entry['path'], self.effort))
            tb = self.toolbox
            agent = ag.Agent(lambda msgs, tools: self.make_stream(entry, base, msgs, tools), tb,
                             max_steps=self.max_steps, label=self.display(entry), **self.agent_budget(entry, tb))
            self.agent = agent
            if cancel.is_set():
                outcome = 'cancelled'
            else:
                outcome = agent.run(self.messages, relay, cancel)
        except Exception as e:      # noqa: BLE001 - cualquier falla se muestra, ninguna se traga
            if cancel.is_set():
                outcome = 'cancelled'
            elif isinstance(e, (dsapi.ApiError, localmodels.LocalError)):
                error = str(e)
            else:
                error = f'Error inesperado: {e!r}'
        finally:
            self.agent = None
            if agent is not None and agent.ctx_tokens:
                self._cpt[entry.get('path') or entry['name']] = agent.chars_per_token
            ag.repair(self.messages)
        restore = None
        if (error or outcome == 'cancelled') and len(self.messages) <= n0 + 1:
            del self.messages[n0:]      # no hubo respuesta: el texto vuelve al campo para reintentar
            restore = texto
        return {'outcome': outcome, 'error': error, 'restore': restore}

    def detener(self):
        'Se puede llamar desde cualquier hilo: corta la espera del modelo y lo que este haciendo la herramienta.'
        ev = self.cancel_ev
        if ev is not None:
            ev.set()
        self.sesion.detener.set()
        a = self.agent
        if a is not None:
            a.cancel_stream()

    def ultima_respuesta(self):
        last = next((m for m in reversed(self.messages) if m.get('role') == 'assistant' and m.get('content')), {})
        return str(last.get('content') or '').strip()

    def aprobacion(self):
        'La de Configuracion, salvo que en esta charla se hayan permitido todos los reemplazos.'
        return 'edits' if self.cfg['approval'] == 'ask' and self.ediciones_permitidas else self.cfg['approval']

    def permitir_ediciones(self):
        'El usuario eligio "Permitir todos los reemplazos de esta charla" (borrar y guardar se siguen preguntando).'
        self.ediciones_permitidas = True
        if self.toolbox is not None:
            self.toolbox.approval = self.aprobacion()

    def nueva_charla(self):
        self.messages = []
        self.ediciones_permitidas = False

    # ------------------------------------------------------------ cierre

    def detener_servidor(self):
        if self._server is not None:
            self._server.stop()
            self._server = None
            self._server_key = None

    def salir(self, cerrar_excel=False):
        self.detener_servidor()
        self.sesion.salir(cerrar_excel=cerrar_excel)
