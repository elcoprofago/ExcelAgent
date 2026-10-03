# Boton 'Actualizar' (actualizar.py) contra una instalacion de juguete y un GitHub falso: sin red.
import io
import json
import os
import stat
import sys
import urllib.error
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

import actualizar
from actualizar import ErrorActualizacion

API = 'https://api.github.com/repos/elcoprofago/ExcelAgent/releases/latest'
ZIP = 'https://api.github.com/repos/elcoprofago/ExcelAgent/zipball/v1.1.0'
ARBOL = 'https://api.github.com/repos/elcoprofago/ExcelAgent/git/trees/v1.0.0?recursive=1'


def escribir(carpeta, archivos):
    for ruta, texto in archivos.items():
        p = os.path.join(carpeta, ruta)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, 'w', encoding='utf-8') as f:
            f.write(texto)


def leer(carpeta, ruta):
    with open(os.path.join(carpeta, ruta), encoding='utf-8') as f:
        return f.read()


def existe(carpeta, ruta):
    return os.path.lexists(os.path.join(carpeta, ruta))


# La version 1.0.0 tal como esta publicada (lo que devuelve el arbol de su etiqueta).
PROGRAMA_VIEJO = {'gui.py': 'viejo', 'version.py': "VERSION = '1.0.0'\n", 'ExcelAgent.vbs': 'vbs',
                  'requirements.txt': 'a==1\n', 'assets/b-env.png': 'img vieja', 'obsoleto.py': 'ya no va',
                  'viejos/solo.py': 'ya no va'}
# Personalizaciones y archivos del usuario: controles, tienen que quedar tal cual.
DEL_USUARIO = {'agregado_por_el_usuario.txt': 'mio', 'planillas/cuentas.xlsx': 'datos',
               'logs/excelagent_hoy.log': 'registro', 'data/config.json': '{"model": "x"}', 'portable.flag': '',
               'Models/chico.gguf': 'pesos', 'bin/llama-server.exe': 'exe'}
# Lo que no es ni del usuario ni de la version nueva: se borra.
SOBRANTE = {'.git/HEAD': 'ref: refs/heads/main', '.git/objects/ab/cdef': 'objeto',
            '_versiones_anteriores/1.0.0_20260101/gui.py': 'muy viejo', '__pycache__/gui.cpython-313.pyc': 'pyc',
            'assets/__pycache__/x.pyc': 'pyc'}
VIEJA = {**PROGRAMA_VIEJO, **DEL_USUARIO, **SOBRANTE}
NUEVA = {'gui.py': 'nuevo', 'version.py': "VERSION = '1.1.0'\n", 'ExcelAgent.vbs': 'vbs', 'requirements.txt': 'a==1\n',
         'assets/b-env.png': 'img nueva', 'modulo_nuevo.py': 'nuevo'}


def zip_de(archivos, raiz='elcoprofago-ExcelAgent-abc1234/'):
    b = io.BytesIO()
    with zipfile.ZipFile(b, 'w') as z:
        for ruta, texto in archivos.items():
            z.writestr(raiz + ruta, texto)
    return b.getvalue()


class Respuesta(io.BytesIO):
    def __init__(self, datos):
        super().__init__(datos)
        self.headers = {'Content-Length': str(len(datos))}


def github(zip_bytes, tag='v1.1.0', arbol=PROGRAMA_VIEJO):
    "Un urlopen falso: la API de la ultima release, el zip y el arbol de la version instalada (arbol=None: 404)."
    def abrir(pedido, timeout=None):
        assert pedido.get_header('User-agent', '').startswith('ExcelAgent/')
        if pedido.full_url == API:
            return Respuesta(json.dumps({'tag_name': tag, 'name': 'Version ' + tag, 'body': 'Novedades',
                                         'zipball_url': ZIP, 'html_url': 'https://github.com/x'}).encode())
        if pedido.full_url == ZIP:
            return Respuesta(zip_bytes)
        if pedido.full_url == ARBOL:
            if arbol is None:
                raise urllib.error.HTTPError(pedido.full_url, 404, 'Not Found', {}, None)
            entradas = [{'path': p, 'type': 'blob'} for p in arbol] + [{'path': 'assets', 'type': 'tree'}]
            return Respuesta(json.dumps({'tree': entradas, 'truncated': False}).encode())
        raise AssertionError(pedido.full_url)
    return abrir


@pytest.fixture
def instalada(tmp_path):
    p = str(tmp_path / 'ExcelAgent')
    escribir(p, VIEJA)
    os.chmod(os.path.join(p, '.git/objects/ab/cdef'), stat.S_IREAD)     # como los objetos de git: solo lectura
    return p


def sin_pip(programa):
    raise AssertionError('no deberia instalar bibliotecas: requirements.txt no cambio')


def test_instala_la_nueva_conserva_lo_del_usuario_y_borra_el_resto(instalada):
    abrir = github(zip_de(NUEVA))
    rel = actualizar.ultima_release(abrir=abrir)
    assert rel['version'] == '1.1.0' and actualizar.hay_nueva(rel, '1.0.0')
    avisos = actualizar.actualizar(rel, instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert avisos == []
    for ruta, texto in NUEVA.items():
        assert leer(instalada, ruta) == texto
    # controles: las personalizaciones y los archivos del usuario quedan tal cual
    for ruta, texto in DEL_USUARIO.items():
        assert leer(instalada, ruta) == texto, ruta
    # lo de la version anterior que la nueva no trae, .git, los respaldos viejos y __pycache__: borrados
    for ruta in ('obsoleto.py', 'viejos', '.git', '_versiones_anteriores', '__pycache__', 'assets/__pycache__',
                 actualizar.EN_CURSO):
        assert not existe(instalada, ruta), ruta
    # y nada mas que eso: lo que hay es exactamente la version nueva mas lo del usuario
    hay = {os.path.relpath(os.path.join(c, n), instalada) for c, _, ns in os.walk(instalada) for n in ns}
    assert hay == {os.path.normpath(r) for r in {**NUEVA, **DEL_USUARIO}}


def test_sin_la_lista_de_la_version_anterior_no_borra_archivos_dudosos(instalada):
    abrir = github(zip_de(NUEVA), arbol=None)       # p. ej. una version que no esta publicada
    avisos = actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert leer(instalada, 'gui.py') == 'nuevo'
    assert len(avisos) == 1 and 'version anterior' in avisos[0]
    assert leer(instalada, 'obsoleto.py') == 'ya no va'           # en la duda, no se borra
    assert not existe(instalada, '.git') and not existe(instalada, '_versiones_anteriores')
    for ruta, texto in DEL_USUARIO.items():
        assert leer(instalada, ruta) == texto, ruta


def test_si_falla_a_mitad_vuelve_exactamente_a_la_anterior(instalada):
    nueva = dict(NUEVA, **{'requirements.txt': 'a==2\n'})
    llamadas = []

    def pip_que_falla(programa):
        llamadas.append(leer(programa, 'requirements.txt'))
        if len(llamadas) == 1:
            raise ErrorActualizacion('sin internet')
    abrir = github(zip_de(nueva))
    rel = actualizar.ultima_release(abrir=abrir)
    with pytest.raises(ErrorActualizacion, match='Volvi a dejar la version 1.0.0'):
        actualizar.actualizar(rel, instalada, '1.0.0', abrir=abrir, pip=pip_que_falla)
    for ruta, texto in VIEJA.items():                 # todo como estaba, .git incluido
        assert leer(instalada, ruta) == texto, ruta
    assert not existe(instalada, 'modulo_nuevo.py') and not existe(instalada, actualizar.EN_CURSO)
    # pip corrio con las nuevas y, al volver, con las de la anterior
    assert llamadas == ['a==2\n', 'a==1\n']


def test_falla_de_copia_tambien_vuelve_atras(instalada, monkeypatch):
    original = actualizar._copiar
    cuenta = {'n': 0}
    pisados = sum(1 for r in NUEVA if r in PROGRAMA_VIEJO)

    def copiar(o, d):
        cuenta['n'] += 1
        if cuenta['n'] == pisados + 3:          # despues de apartar los pisados, a mitad de la copia nueva
            raise PermissionError('archivo en uso')
        original(o, d)
    monkeypatch.setattr(actualizar, '_copiar', copiar)
    abrir = github(zip_de(NUEVA))
    with pytest.raises(ErrorActualizacion, match='archivo en uso'):
        actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    for ruta, texto in VIEJA.items():
        assert leer(instalada, ruta) == texto, ruta
    assert not existe(instalada, 'modulo_nuevo.py') and not existe(instalada, actualizar.EN_CURSO)


def test_lo_que_no_se_puede_borrar_se_avisa_y_no_frena(instalada, monkeypatch):
    original = actualizar._borrar

    def borrar(ruta):
        if os.path.basename(ruta) == 'obsoleto.py':
            raise PermissionError(13, 'El archivo esta en uso')
        original(ruta)
    monkeypatch.setattr(actualizar, '_borrar', borrar)
    abrir = github(zip_de(NUEVA))
    avisos = actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert leer(instalada, 'gui.py') == 'nuevo'
    assert avisos == ['No pude borrar obsoleto.py (El archivo esta en uso).']
    assert not existe(instalada, '.git') and not existe(instalada, 'viejos')


def test_resto_de_una_actualizacion_cortada_no_estorba(instalada):
    escribir(instalada, {actualizar.EN_CURSO + '/gui.py': 'de un corte de luz'})
    abrir = github(zip_de(NUEVA))
    assert actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir,
                                 pip=sin_pip) == []
    assert leer(instalada, 'gui.py') == 'nuevo' and not existe(instalada, actualizar.EN_CURSO)


def test_un_cambio_solo_de_mayusculas_no_borra_el_archivo_nuevo(tmp_path):
    p = str(tmp_path)
    escribir(p, {'Modulo.py': 'nuevo'})
    assert actualizar.limpiar(p, ['Modulo.py'], {'modulo.py'}) == []
    assert leer(p, 'Modulo.py') == 'nuevo'


def test_version_de_adentro_distinta_de_la_etiqueta_no_instala(instalada):
    abrir = github(zip_de(dict(NUEVA, **{'version.py': "VERSION = '1.0.9'\n"})))
    with pytest.raises(ErrorActualizacion, match='no coinciden'):
        actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    for ruta, texto in VIEJA.items():
        assert leer(instalada, ruta) == texto, ruta
    assert not existe(instalada, actualizar.EN_CURSO)


def test_paquete_con_rutas_afuera_se_rechaza(instalada, tmp_path):
    abrir = github(zip_de({'../../fuera.py': 'x', **NUEVA}, raiz=''))
    with pytest.raises(ErrorActualizacion, match='fuera de su carpeta'):
        actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert leer(instalada, 'gui.py') == 'viejo'


def test_comparacion_de_versiones_y_errores_de_github():
    assert actualizar.hay_nueva({'tag': 'v1.10.0'}, '1.9.3')
    assert not actualizar.hay_nueva({'tag': 'v1.0.0'}, '1.0.0')
    with pytest.raises(ErrorActualizacion, match='no entiendo'):
        actualizar.numero('ultima')

    def no_hay(pedido, timeout=None):
        raise urllib.error.HTTPError(pedido.full_url, 404, 'Not Found', {}, None)

    def sin_red(pedido, timeout=None):
        raise urllib.error.URLError('getaddrinfo failed')
    with pytest.raises(ErrorActualizacion, match='ninguna version publicada'):
        actualizar.ultima_release(abrir=no_hay)
    with pytest.raises(ErrorActualizacion, match='conexion a internet'):
        actualizar.ultima_release(abrir=sin_red)
    assert actualizar.archivos_de_version('v1.0.0', abrir=sin_red) is None
