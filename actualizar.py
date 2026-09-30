# Boton 'Actualizar': trae la ultima release publicada en GitHub y la instala sobre la carpeta del programa.
#
# Orden, y por que:
#   1. se baja el zip de la release a una carpeta temporal y se valida (que sea ExcelAgent y que la version de adentro
#      coincida con la etiqueta);
#   2. se respalda la version instalada ENTERA en _versiones_anteriores\<version>_<fecha>, no solo lo que se va a pisar:
#      es el estado al que se vuelve;
#   3. recien ahi se copian los archivos nuevos. Si algo falla a mitad de camino, se restaura el respaldo y se borra lo
#      que la copia habia agregado: queda exactamente la version anterior.
# No se tocan: logs\, la configuracion (en %APPDATA%), el Python (..\python del instalador o ..\.venv) ni
# ..\tesseract. Tampoco se borra ningun archivo que la version nueva ya no traiga: puede haberlo puesto el usuario, y un
# archivo de mas no rompe nada.
import json
import os
import re
import shutil
import subprocess
import sys
import tempfile
import time
import urllib.error
import urllib.request
import zipfile

import version

API = 'https://api.github.com/repos/{repo}/releases/latest'
CARPETA_RESPALDOS = '_versiones_anteriores'
# Carpetas de primer nivel que no son del programa: ni se respaldan ni se pisan.
NO_TOCAR = {'logs', '__pycache__', '.git', CARPETA_RESPALDOS, '_excelagent_respaldos'}
PROGRAMA = os.path.dirname(os.path.abspath(__file__))
SIN_VENTANA = 0x08000000        # CREATE_NO_WINDOW: pip y git sin consola que parpadee


class ErrorActualizacion(Exception):
    pass


def numero(etiqueta):
    'v1.2.3 o 1.2.3 -> (1, 2, 3), para comparar versiones.'
    m = re.fullmatch(r'v?(\d+)\.(\d+)\.(\d+)', (etiqueta or '').strip())
    if not m:
        raise ErrorActualizacion(f'La version publicada tiene una etiqueta que no entiendo: {etiqueta!r}')
    return tuple(int(x) for x in m.groups())


def _pedir(url, timeout, abrir):
    pedido = urllib.request.Request(url, headers={'User-Agent': 'ExcelAgent/' + version.VERSION,
                                                  'Accept': 'application/vnd.github+json'})
    return abrir(pedido, timeout=timeout)


def ultima_release(repo=version.REPO, abrir=urllib.request.urlopen):
    'La ultima release publicada: etiqueta, version, notas y el zip del codigo.'
    try:
        with _pedir(API.format(repo=repo), 15, abrir) as r:
            d = json.load(r)
    except urllib.error.HTTPError as exc:
        if exc.code == 404:
            raise ErrorActualizacion('Todavia no hay ninguna version publicada para descargar.') from None
        raise ErrorActualizacion(f'GitHub respondio con el error {exc.code} al buscar la ultima version.') from None
    except (urllib.error.URLError, OSError, ValueError) as exc:
        raise ErrorActualizacion('No pude consultar GitHub (' + str(getattr(exc, 'reason', exc)) +
                                 '). Revisa la conexion a internet y proba de nuevo.') from None
    etiqueta = d.get('tag_name') or ''
    return {'tag': etiqueta, 'version': '.'.join(map(str, numero(etiqueta))), 'nombre': d.get('name') or etiqueta,
            'notas': (d.get('body') or '').strip(), 'zip': d.get('zipball_url'), 'pagina': d.get('html_url')}


def hay_nueva(rel, actual=version.VERSION):
    return numero(rel['tag']) > numero(actual)


def descargar(url, destino, abrir=urllib.request.urlopen, avance=None):
    try:
        with _pedir(url, 60, abrir) as r, open(destino, 'wb') as f:
            total = int(r.headers.get('Content-Length') or 0)
            leido = 0
            while True:
                trozo = r.read(65536)
                if not trozo:
                    break
                f.write(trozo)
                leido += len(trozo)
                if avance:
                    avance(leido, total)
    except (urllib.error.URLError, OSError) as exc:
        raise ErrorActualizacion('Se corto la descarga (' + str(getattr(exc, 'reason', exc)) +
                                 '). No cambie nada: proba de nuevo.') from None
    return destino


def extraer(zip_path, carpeta):
    'El zip de GitHub trae todo dentro de una carpeta dueño-repo-commit: devuelve esa carpeta.'
    base = os.path.realpath(carpeta)
    try:
        with zipfile.ZipFile(zip_path) as z:
            for nombre in z.namelist():
                ruta = os.path.realpath(os.path.join(base, nombre))
                if ruta != base and not ruta.startswith(base + os.sep):
                    raise ErrorActualizacion('El paquete descargado trae rutas fuera de su carpeta: no lo uso.')
            z.extractall(base)
    except zipfile.BadZipFile:
        raise ErrorActualizacion('El paquete descargado esta dañado. No cambie nada: proba de nuevo.') from None
    hijos = os.listdir(base)
    if len(hijos) == 1 and os.path.isdir(os.path.join(base, hijos[0])):
        return os.path.join(base, hijos[0])
    return base


def version_de(carpeta):
    try:
        with open(os.path.join(carpeta, 'version.py'), encoding='utf-8') as f:
            m = re.search(r"^VERSION\s*=\s*'([^']+)'", f.read(), re.M)
    except OSError:
        return None
    return m.group(1) if m else None


def validar(raiz, rel):
    for f in ('gui.py', 'version.py', 'ExcelAgent.vbs', 'requirements.txt'):
        if not os.path.isfile(os.path.join(raiz, f)):
            raise ErrorActualizacion(f'El paquete descargado no parece ExcelAgent (falta {f}). No cambie nada.')
    adentro = version_de(raiz)
    if adentro != rel['version']:
        raise ErrorActualizacion(f'La version publicada {rel["tag"]} trae adentro la version {adentro}: no coinciden y '
                                 'no la instalo. Hay que corregir la publicacion.')


def archivos(raiz):
    'Rutas relativas de los archivos del programa (sin logs, respaldos, .git ni __pycache__).'
    for carpeta, subcarpetas, nombres in os.walk(raiz):
        rel = os.path.relpath(carpeta, raiz)
        if rel == '.':
            subcarpetas[:] = [d for d in subcarpetas if d not in NO_TOCAR]
        else:
            subcarpetas[:] = [d for d in subcarpetas if d != '__pycache__']
        for n in nombres:
            yield n if rel == '.' else os.path.join(rel, n)


def _copiar(origen, destino):
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    temporal = destino + '.nuevo'
    shutil.copy2(origen, temporal)
    os.replace(temporal, destino)       # el archivo viejo se reemplaza entero o no se toca


def respaldar(programa, actual):
    'Copia de toda la version instalada; se verifica apenas se crea (mismos archivos, mismos tamaños).'
    destino = os.path.join(programa, CARPETA_RESPALDOS, actual + '_' + time.strftime('%Y%m%d_%H%M%S'))
    lista = sorted(archivos(programa))
    for r in lista:
        _copiar(os.path.join(programa, r), os.path.join(destino, r))
    copiados = sorted(archivos(destino))
    if copiados != lista or any(os.path.getsize(os.path.join(programa, r)) != os.path.getsize(os.path.join(destino, r))
                                for r in lista):
        raise ErrorActualizacion('El respaldo de la version actual no quedo completo: no actualizo. Respaldo: ' + destino)
    return destino


def restaurar(respaldo, programa, agregados):
    'Vuelve al estado del respaldo: repone todos sus archivos y borra los que agrego la actualizacion.'
    for r in archivos(respaldo):
        _copiar(os.path.join(respaldo, r), os.path.join(programa, r))
    for r in agregados:
        try:
            os.remove(os.path.join(programa, r))
        except FileNotFoundError:
            pass


def _chequear_git(programa, correr=subprocess.run):
    'Una carpeta clonada con git y con cambios sin commitear no se pisa: esos cambios no estan en ningun otro lado.'
    if not os.path.isdir(os.path.join(programa, '.git')):
        return
    try:
        r = correr(['git', '-C', programa, 'status', '--porcelain'], capture_output=True, text=True, timeout=30,
                   creationflags=SIN_VENTANA)
    except OSError:
        raise ErrorActualizacion('Esta carpeta es un repositorio de git y no encuentro git para ver si tiene cambios '
                                 'sin guardar. Actualizala con git (git pull).') from None
    if r.returncode != 0:
        raise ErrorActualizacion('Esta carpeta es un repositorio de git y git no pudo revisarla: ' + r.stderr.strip())
    if r.stdout.strip():
        raise ErrorActualizacion('Esta carpeta es un repositorio de git con cambios sin commitear: no la piso. '
                                 'Commitealos o actualiza con git pull.')


def instalar_dependencias(programa, python=None, correr=subprocess.run):
    python = python or sys.executable.replace('pythonw.exe', 'python.exe')
    # -E -s, igual que el lanzador: ni variables PYTHON* ni las bibliotecas del usuario (%APPDATA%\Python). Si no, pip
    # puede dar por instalada una biblioteca que esta ahi y no en el Python del programa.
    r = correr([python, '-E', '-s', '-m', 'pip', 'install', '-r', os.path.join(programa, 'requirements.txt')],
               capture_output=True, text=True, timeout=900, creationflags=SIN_VENTANA)
    if r.returncode != 0:
        ultimas = '\n'.join((r.stderr or r.stdout or '').strip().splitlines()[-3:])
        raise ErrorActualizacion('No pude instalar las bibliotecas nuevas:\n' + ultimas)


def aplicar(raiz, programa, actual, pip=instalar_dependencias, correr=subprocess.run):
    """Instala la version de 'raiz' sobre 'programa'. Devuelve la carpeta del respaldo. Ante cualquier falla deja la
    version anterior tal cual estaba y lo dice."""
    _chequear_git(programa, correr)
    respaldo = respaldar(programa, actual)
    agregados = []
    pip_corrio = False
    try:
        for r in sorted(archivos(raiz)):
            destino = os.path.join(programa, r)
            if not os.path.exists(destino):
                agregados.append(r)
            _copiar(os.path.join(raiz, r), destino)
        with open(os.path.join(respaldo, 'requirements.txt'), 'rb') as a, \
                open(os.path.join(programa, 'requirements.txt'), 'rb') as b:
            # sin contar los fines de linea: una copia sacada con git en Windows tiene CRLF y el zip de GitHub LF
            cambiaron = a.read().replace(b'\r\n', b'\n') != b.read().replace(b'\r\n', b'\n')
        if cambiaron:
            pip_corrio = True
            pip(programa)
    except Exception as exc:
        motivo = str(exc)
        try:
            restaurar(respaldo, programa, agregados)
        except Exception as exc2:
            raise ErrorActualizacion(f'Fallo la actualizacion ({motivo}) y tampoco pude volver atras ({exc2}). '
                                     f'La version anterior entera esta en {respaldo}.') from None
        extra = ''
        if pip_corrio:
            # pip pudo haber cambiado algunas bibliotecas antes de fallar: se vuelven a las versiones de la anterior.
            try:
                pip(programa)
            except Exception as exc3:
                extra = ('\nOjo: las bibliotecas pueden haber quedado a medio cambiar. Para arreglarlo: '
                         '"' + sys.executable.replace('pythonw.exe', 'python.exe') +
                         '" -E -s -m pip install -r requirements.txt (' + str(exc3)[:200] + ')')
        raise ErrorActualizacion(f'Fallo la actualizacion: {motivo}\nVolvi a dejar la version {actual} como estaba '
                                 f'(respaldo en {respaldo}).' + extra) from None
    return respaldo


def actualizar(rel, programa=PROGRAMA, actual=version.VERSION, avance=None, abrir=urllib.request.urlopen,
               pip=instalar_dependencias):
    'Todo el recorrido: bajar, validar, respaldar e instalar. Devuelve la carpeta del respaldo.'
    temporal = tempfile.mkdtemp(prefix='excelagent_actualizacion_')
    try:
        zip_path = descargar(rel['zip'], os.path.join(temporal, 'release.zip'), abrir, avance)
        raiz = extraer(zip_path, os.path.join(temporal, 'contenido'))
        validar(raiz, rel)
        return aplicar(raiz, programa, actual, pip)
    finally:
        shutil.rmtree(temporal, ignore_errors=True)


def relanzar(programa=PROGRAMA):
    'Arranca la version recien instalada con su propio lanzador (sin consola).'
    subprocess.Popen(['wscript.exe', '//NoLogo', os.path.join(programa, 'ExcelAgent.vbs')], creationflags=SIN_VENTANA)
