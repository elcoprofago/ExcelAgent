import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
from openpyxl import Workbook

import asistente
import dsapi
import excel
import localmodels
import prompts


class Guion:
    'Modelo falso: cada paso es una lista de eventos (o una excepcion). Registra lo que recibio.'
    def __init__(self, *pasos):
        self.pasos = list(pasos)
        self.i = 0
        self.vistos = []

    def __call__(self, entry, msgs, tools):
        self.vistos.append((entry, msgs, tools))
        paso = self.pasos[min(self.i, len(self.pasos) - 1)]
        self.i += 1
        return Stream(paso)


class Stream:
    def __init__(self, eventos):
        self.eventos = eventos

    def cancel(self):
        pass

    def __iter__(self):
        if isinstance(self.eventos, Exception):
            raise self.eventos
        for e in self.eventos:
            if callable(e):
                e()
            else:
                yield e


def uso(i=100, o=20):
    return ('usage', {'prompt_tokens': i, 'completion_tokens': o, 'prompt_cache_hit_tokens': 60,
                      'completion_tokens_details': {'reasoning_tokens': 5}})


def llamada(id_, nombre, args):
    return {'id': id_, 'type': 'function', 'function': {'name': nombre, 'arguments': json.dumps(args)}}


def respuesta(texto):
    return [('content', texto), uso(), ('finish', 'stop')]


@pytest.fixture(scope='module')
def excel_vivo():
    s = excel.Sesion(excel.Bus())
    yield s
    s.salir(cerrar_excel=True, descartar=True)


@pytest.fixture
def nuevo(tmp_path, excel_vivo):
    def armar(*pasos):
        cfg = dsapi.Config(str(tmp_path / 'cfg'))
        guion = Guion(*pasos)
        a = asistente.Asistente(cfg, excel.Bus(), stream_factory=guion)
        a.sesion = excel_vivo      # una sola instancia de Excel para todo el modulo
        return a, guion
    yield armar
    excel_vivo.cerrar(False)


def libro_de_juguete(tmp_path):
    ruta = tmp_path / 'ventas.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Datos'
    for f in (('Cliente', 'Importe'), ('ACME', 1000), ('Beta', 250), ('Zenit', 700)):
        ws.append(list(f))
    wb.create_sheet('Control')['A1'] = 'NO TOCAR'
    wb.save(str(ruta))
    return str(ruta)


def test_pregunta_general_sin_libro(nuevo):
    a, guion = nuevo(respuesta('Se usa =SUMA(A1:A3).'))
    eventos = []
    r = a.pedir('como sumo una columna?', emit=lambda k, *x: eventos.append(k))
    assert r == {'outcome': 'done', 'error': None, 'restore': None}
    _, msgs, tools = guion.vistos[0]
    assert msgs[0] == {'role': 'system', 'content': prompts.system_prompt()}
    assert msgs[1]['content'].startswith('como sumo una columna?')
    assert prompts.NO_WORKBOOK in msgs[1]['content']
    assert {t['function']['name'] for t in tools} >= {'read_range', 'write_range'}
    assert a.ultima_respuesta() == 'Se usa =SUMA(A1:A3).'
    assert a.totals == {'in': 100, 'out': 20, 'reason': 5, 'hit': 60}
    assert 'content' in eventos and 'usage' in eventos


def test_pedido_vago_ejecuta_herramientas_en_el_libro(nuevo, tmp_path):
    a, guion = nuevo(
        [('tool_calls', [llamada('c1', 'write_range', {'start_cell': 'D1', 'values': [['Total'], ['=SUM(B2:B4)']]})]),
         uso(), ('finish', 'tool_calls')],
        respuesta('Listo: puse el total en D2.'))
    a.abrir(libro_de_juguete(tmp_path), visible=False, respaldo=False, minimizar=False)
    r = a.pedir('poneme el total')
    assert r['outcome'] == 'done' and r['error'] is None
    wb = a.sesion.wb
    assert wb.Worksheets('Datos').Range('D2').Value2 == 1950
    assert wb.Worksheets('Control').Range('A1').Value2 == 'NO TOCAR'
    estado = guion.vistos[0][1][1]['content']
    assert "sheet 'Datos': used A1:B4 (active)" in estado
    resultado = [m for m in a.messages if m['role'] == 'tool'][0]['content']
    assert resultado.startswith('Wrote 2 x 1 cells')
    # el segundo paso ve el resultado de la herramienta
    assert guion.vistos[1][1][-1]['role'] == 'tool'


def test_la_misma_escritura_repetida_corta_el_bucle(nuevo, tmp_path):
    # Regresion (Qwen3.5-9B local): una escritura 'exitosa' repetida reiniciaba el contador de repeticiones cada vez y
    # el agente llegaba al tope de 60 pasos. El guion repite su unico paso para siempre.
    a, guion = nuevo(
        [('tool_calls', [llamada('c1', 'write_range', {'start_cell': 'D1', 'values': [['Total']]})]),
         uso(), ('finish', 'tool_calls')])
    a.cfg['approval'] = 'all'
    a.abrir(libro_de_juguete(tmp_path), visible=False, respaldo=False, minimizar=False)
    r = a.pedir('poneme el titulo')
    assert r['outcome'] == 'loop'
    assert guion.i <= 5
    assert a.sesion.wb.Worksheets('Datos').Range('D1').Value2 == 'Total'


def test_error_de_la_api_devuelve_el_texto(nuevo):
    a, _ = nuevo(dsapi.ApiError('Saldo insuficiente'))
    r = a.pedir('hola')
    assert r == {'outcome': None, 'error': 'Saldo insuficiente', 'restore': 'hola'}
    assert a.messages == []


def test_error_inesperado_no_se_traga(nuevo):
    a, _ = nuevo(RuntimeError('averia'))
    r = a.pedir('hola')
    assert 'averia' in r['error'] and r['restore'] == 'hola'


def test_detener_durante_la_respuesta(nuevo):
    a, _ = nuevo(None)
    a._stream_factory.pasos = [[('content', 'empiezo'), lambda: a.detener(), ('content', ' y sigo'), ('finish', 'stop')]]
    r = a.pedir('hace algo largo')
    assert r['outcome'] == 'cancelled'
    assert a.sesion.detener.is_set()


def test_sin_key_no_sale_a_la_red(tmp_path):
    cfg = dsapi.Config(str(tmp_path / 'cfg'))
    a = asistente.Asistente(cfg, excel.Bus())
    assert a.entry()['kind'] == 'remote' and a.necesita_key()
    r = a.pedir('hola')
    assert 'API key' in r['error'] and r['restore'] == 'hola'
    assert a.messages == []


def test_modelo_local_inexistente(tmp_path):
    cfg = dsapi.Config(str(tmp_path / 'cfg'))
    a = asistente.Asistente(cfg, excel.Bus())
    a.model_id = localmodels.MODEL_ID_PREFIX + str(tmp_path / 'no-existe.gguf')
    e = a.entry()
    assert e['kind'] == 'local' and e['missing']
    assert a.display(e) == '[local] no-existe (no encontrado)'
    assert not a.necesita_key()
    r = a.pedir('hola')
    assert r['error'] and r['restore'] == 'hola'


def test_presupuesto_local_descuenta_instrucciones_y_herramientas(tmp_path):
    cfg = dsapi.Config(str(tmp_path / 'cfg'))
    a = asistente.Asistente(cfg, excel.Bus())
    import herramientas
    tb = herramientas.ToolBox(a.sesion)
    b = a.agent_budget({'kind': 'local', 'name': 'x', 'path': 'x.gguf'}, tb)
    assert b['ctx_tokens'] == cfg['local_ctx']
    assert b['reserve_tokens'] == a.local_max_tokens() == max(512, min(cfg['max_tokens'], cfg['local_ctx'] // 3))
    assert b['overhead_chars'] > len(prompts.system_prompt())
    assert a.agent_budget({'kind': 'remote', 'id': 'deepseek-flash'}, tb) == {'budget_chars': asistente.REMOTE_BUDGET}
