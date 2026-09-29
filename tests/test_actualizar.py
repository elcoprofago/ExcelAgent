# Boton 'Actualizar' (actualizar.py) contra una instalacion de juguete y un GitHub falso: sin red.
import io
import json
import os
import sys
import urllib.error
import zipfile

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

import actualizar
from actualizar import ErrorActualizacion

API = 'https://api.github.com/repos/elcoprofago/ExcelAgent/releases/latest'
ZIP = 'https://api.github.com/repos/elcoprofago/ExcelAgent/zipball/v1.1.0'


def escribir(carpeta, archivos):
    for ruta, texto in archivos.items():
        p = os.path.join(carpeta, ruta)
        os.makedirs(os.path.dirname(p), exist_ok=True)
        with open(p, 'w', encoding='utf-8') as f:
            f.write(texto)


def leer(carpeta, ruta):
    with open(os.path.join(carpeta, ruta), encoding='utf-8') as f:
        return f.read()


VIEJA = {'gui.py': 'viejo', 'version.py': "VERSION = '1.0.0'\n", 'ExcelAgent.vbs': 'vbs', 'requirements.txt': 'a==1\n',
         'assets/b-env.png': 'img vieja', 'agregado_por_el_usuario.txt': 'mio',     # control: debe quedar
         'logs/excelagent_hoy.log': 'registro'}                                      # control: ni se respalda ni se toca
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


def github(zip_bytes, tag='v1.1.0'):
    'Un urlopen falso: la API de la ultima release y el zip.'
    def abrir(pedido, timeout=None):
        assert pedido.get_header('User-agent', '').startswith('ExcelAgent/')
        if pedido.full_url == API:
            return Respuesta(json.dumps({'tag_name': tag, 'name': 'Version ' + tag, 'body': 'Novedades',
                                         'zipball_url': ZIP, 'html_url': 'https://github.com/x'}).encode())
        if pedido.full_url == ZIP:
            return Respuesta(zip_bytes)
        raise AssertionError(pedido.full_url)
    return abrir


@pytest.fixture
def instalada(tmp_path):
    p = str(tmp_path / 'ExcelAgent')
    escribir(p, VIEJA)
    return p


def sin_pip(programa):
    raise AssertionError('no deberia instalar bibliotecas: requirements.txt no cambio')


def test_instala_la_nueva_y_respalda_la_anterior_entera(instalada):
    abrir = github(zip_de(NUEVA))
    rel = actualizar.ultima_release(abrir=abrir)
    assert rel['version'] == '1.1.0' and actualizar.hay_nueva(rel, '1.0.0')
    respaldo = actualizar.actualizar(rel, instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert leer(instalada, 'gui.py') == 'nuevo' and leer(instalada, 'assets/b-env.png') == 'img nueva'
    assert leer(instalada, 'modulo_nuevo.py') == 'nuevo'
    assert actualizar.version_de(instalada) == '1.1.0'
    # controles
    assert leer(instalada, 'agregado_por_el_usuario.txt') == 'mio'
    assert leer(instalada, 'logs/excelagent_hoy.log') == 'registro'
    # el respaldo es la version anterior completa, sin el registro
    assert os.path.basename(respaldo).startswith('1.0.0_')
    assert sorted(actualizar.archivos(respaldo)) == sorted(os.path.normpath(r) for r in VIEJA if not r.startswith('logs'))
    assert leer(respaldo, 'gui.py') == 'viejo'
    assert not any(n.endswith('.nuevo') for n in actualizar.archivos(instalada))


def test_si_falla_a_mitad_vuelve_exactamente_a_la_anterior(instalada, monkeypatch):
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
    assert leer(instalada, 'gui.py') == 'viejo' and leer(instalada, 'assets/b-env.png') == 'img vieja'
    assert not os.path.exists(os.path.join(instalada, 'modulo_nuevo.py'))      # lo agregado se saca
    assert leer(instalada, 'agregado_por_el_usuario.txt') == 'mio'
    # pip corrio con las nuevas y, al volver, con las de la anterior
    assert llamadas == ['a==2\n', 'a==1\n']


def test_falla_de_copia_tambien_vuelve_atras(instalada, monkeypatch):
    original = actualizar._copiar
    cuenta = {'n': 0}
    antes = len(list(actualizar.archivos(instalada)))

    def copiar(o, d):
        cuenta['n'] += 1
        if cuenta['n'] == antes + 3:             # despues del respaldo, a mitad de la copia nueva
            raise PermissionError('archivo en uso')
        original(o, d)
    monkeypatch.setattr(actualizar, '_copiar', copiar)
    abrir = github(zip_de(NUEVA))
    with pytest.raises(ErrorActualizacion, match='archivo en uso'):
        actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    for ruta, texto in VIEJA.items():
        assert leer(instalada, ruta) == texto
    assert not os.path.exists(os.path.join(instalada, 'modulo_nuevo.py'))


def test_version_de_adentro_distinta_de_la_etiqueta_no_instala(instalada):
    abrir = github(zip_de(dict(NUEVA, **{'version.py': "VERSION = '1.0.9'\n"})))
    with pytest.raises(ErrorActualizacion, match='no coinciden'):
        actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert leer(instalada, 'gui.py') == 'viejo'
    assert not os.path.exists(os.path.join(instalada, actualizar.CARPETA_RESPALDOS))


def test_paquete_con_rutas_afuera_se_rechaza(instalada, tmp_path):
    abrir = github(zip_de({'../../fuera.py': 'x', **NUEVA}, raiz=''))
    with pytest.raises(ErrorActualizacion, match='fuera de su carpeta'):
        actualizar.actualizar(actualizar.ultima_release(abrir=abrir), instalada, '1.0.0', abrir=abrir, pip=sin_pip)
    assert leer(instalada, 'gui.py') == 'viejo'


def test_clon_de_git_con_cambios_no_se_pisa(instalada):
    os.makedirs(os.path.join(instalada, '.git'))

    class R:
        returncode, stdout, stderr = 0, ' M gui.py\n', ''
    with pytest.raises(ErrorActualizacion, match='sin commitear'):
        actualizar.aplicar(instalada, instalada, '1.0.0', pip=sin_pip, correr=lambda *a, **k: R())
    assert not os.path.exists(os.path.join(instalada, actualizar.CARPETA_RESPALDOS))


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
