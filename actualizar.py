# Boton 'Actualizar': trae la ultima release publicada en GitHub y la instala sobre la carpeta del programa.
#
# Lo unico que sobrevive de una version a otra son las personalizaciones del usuario: la configuracion (en %APPDATA%,
# fuera de esta carpeta), y en esta carpeta CONSERVAR (config portable, modelos locales, llama-server, registro) y
# cualquier archivo que no sea del programa (puesto por el usuario). Todo lo demas se reemplaza o se borra, incluido
# .git: una copia sacada con git se actualiza igual, tenga o no cambios.
#
# Orden, y por que:
#   1. se baja el zip de la release a una carpeta temporal y se valida (que sea ExcelAgent y que la version de adentro
#      coincida con la etiqueta);
#   2. se aparta una copia de los archivos que se van a pisar (en _actualizacion_en_curso), para volver atras;
#   3. se copian los archivos nuevos (y pip, si cambio requirements.txt). Si algo falla, se reponen los pisados y se
#      borra lo agregado: queda exactamente la version anterior;
#   4. instalada la nueva, se borra: la copia apartada (no es un respaldo, se borra apenas la nueva queda bien), .git,
#      _versiones_anteriores (de las versiones que si guardaban respaldo), los __pycache__ y los archivos de la version
#      anterior que la nueva ya no trae. Esa lista sale de GitHub (el arbol de la etiqueta de la version instalada): sin
#      ella no se sabria si un archivo es del programa o del usuario, y en la duda no se borra.
# El Python (..\python del instalador o ..\.venv) y ..\tesseract estan fuera de esta carpeta: no se tocan.
import json
import os
import re
import shutil
import stat
import subprocess
import sys
import tempfile
import urllib.error
import urllib.request
import zipfile

import version

API = 'https://api.github.com/repos/{repo}/releases/latest'
ARBOL = 'https://api.github.com/repos/{repo}/git/trees/{ref}?recursive=1'
EN_CURSO = '_actualizacion_en_curso'
# Primer nivel: personalizaciones del usuario. Nunca se pisan ni se borran.
CONSERVAR = {'logs', 'data', 'portable.flag', 'Models', 'bin'}
# Primer nivel: lo que no es del usuario ni de la version nueva. Se borra al terminar.
BORRAR = {'.git', '_versiones_anteriores'}
PROGRAMA = os.path.dirname(os.path.abspath(__file__))
SIN_VENTANA = 0x08000000        # CREATE_NO_WINDOW: pip sin consola que parpadee


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
    'Rutas relativas de los archivos de una carpeta, sin CONSERVAR, BORRAR, la copia apartada ni __pycache__.'
    afuera = CONSERVAR | BORRAR | {EN_CURSO}
    for carpeta, subcarpetas, nombres in os.walk(raiz):
        rel = os.path.relpath(carpeta, raiz)
        if rel == '.':
            subcarpetas[:] = [d for d in subcarpetas if d not in afuera and d != '__pycache__']
            nombres = [n for n in nombres if n not in afuera]
        else:
            subcarpetas[:] = [d for d in subcarpetas if d != '__pycache__']
        for n in nombres:
            yield n if rel == '.' else os.path.join(rel, n)


def archivos_de_version(etiqueta, repo=version.REPO, abrir=urllib.request.urlopen):
    """Los archivos que trae la version 'etiqueta' segun GitHub (rutas con el separador de Windows), o None si no se
    pudo saber: sin internet, una version sin publicar (una copia de desarrollo) o una lista truncada."""
    try:
        with _pedir(ARBOL.format(repo=repo, ref=etiqueta), 15, abrir) as r:
            d = json.load(r)
    except (urllib.error.URLError, OSError, ValueError):
        return None
    if d.get('truncated') or not isinstance(d.get('tree'), list):
        return None
    return {os.path.normpath(e['path']) for e in d['tree'] if e.get('type') == 'blob' and e.get('path')}


def _copiar(origen, destino):
    os.makedirs(os.path.dirname(destino), exist_ok=True)
    temporal = destino + '.nuevo'
    shutil.copy2(origen, temporal)
    os.replace(temporal, destino)       # el archivo viejo se reemplaza entero o no se toca


def apartar(programa, pisados):
    'Copia de los archivos que se van a pisar; se verifica apenas se crea (mismos archivos, mismos tamaños).'
    destino = os.path.join(programa, EN_CURSO)
    if os.path.exists(destino):
        # resto de una actualizacion cortada (un corte de luz): la carpeta quedo a medio camino entre dos versiones y
        # esta actualizacion pisa todos los archivos del programa igual, asi que ese resto ya no sirve para volver.
        _borrar(destino)
    for r in pisados:
        _copiar(os.path.join(programa, r), os.path.join(destino, r))
    for r in pisados:
        if os.path.getsize(os.path.join(programa, r)) != os.path.getsize(os.path.join(destino, r)):
            raise ErrorActualizacion(f'La copia de {r} no quedo igual: no actualizo.')
    return destino


def restaurar(apartados, programa, agregados):
    'Vuelve a la version anterior: repone los archivos pisados y borra los que agrego la actualizacion.'
    for r in archivos(apartados):
        _copiar(os.path.join(apartados, r), os.path.join(programa, r))
    for r in agregados:
        try:
            os.remove(os.path.join(programa, r))
        except FileNotFoundError:
            pass


def _borrar(ruta):
    'Borra un archivo o una carpeta entera. Los objetos de .git son de solo lectura: se les saca esa marca y se reintenta.'
    def solo_lectura(funcion, camino, _):
        os.chmod(camino, stat.S_IWRITE)
        funcion(camino)
    if os.path.isdir(ruta) and not os.path.islink(ruta):
        shutil.rmtree(ruta, onerror=solo_lectura)
    elif os.path.lexists(ruta):
        try:
            os.remove(ruta)
        except PermissionError:
            os.chmod(ruta, stat.S_IWRITE)
            os.remove(ruta)


def limpiar(programa, nuevos, viejos):
    """Con la version nueva ya instalada, borra lo que no es de ella ni del usuario. Devuelve avisos: lo que no se pudo
    borrar (un archivo abierto, por ejemplo) y si no se supo cuales eran los archivos de la version anterior."""
    avisos = []
    borrar = [os.path.join(programa, n) for n in sorted(BORRAR | {EN_CURSO})]
    for carpeta, subcarpetas, _ in os.walk(programa):
        if os.path.relpath(carpeta, programa) == '.':
            subcarpetas[:] = [d for d in subcarpetas if d not in CONSERVAR | BORRAR | {EN_CURSO}]
        if '__pycache__' in subcarpetas:
            subcarpetas.remove('__pycache__')
            borrar.append(os.path.join(carpeta, '__pycache__'))
    if viejos is None:
        avisos.append('No pude saber cuales eran los archivos de la version anterior: los que la nueva ya no trae '
                      'quedaron en la carpeta (no molestan).')
    else:
        # sin distinguir mayusculas: en Windows 'Foo.py' y 'foo.py' son el mismo archivo
        nuevos_ = {os.path.normcase(r) for r in nuevos}
        for r in sorted(viejos):
            if os.path.normcase(r) in nuevos_ or r.split(os.sep)[0] in CONSERVAR | BORRAR:
                continue
            borrar.append(os.path.join(programa, r))
    carpetas = set()
    for ruta in borrar:
        if not os.path.lexists(ruta):
            continue
        try:
            _borrar(ruta)
            carpetas.add(os.path.dirname(ruta))
        except OSError as exc:
            avisos.append(f'No pude borrar {os.path.relpath(ruta, programa)} ({exc.strerror or exc}).')
    # carpetas que quedaron vacias al sacar archivos viejos (de la mas profunda a la de arriba)
    for c in sorted(carpetas, key=len, reverse=True):
        while os.path.normcase(c) != os.path.normcase(programa) and os.path.isdir(c) and not os.listdir(c):
            os.rmdir(c)
            c = os.path.dirname(c)
    return avisos


def instalar_dependencias(programa, python=None, correr=subprocess.run):
    python = python or sys.executable.replace('pythonw.exe', 'python.exe')
    # -E -s, igual que el lanzador: ni variables PYTHON* ni las bibliotecas del usuario (%APPDATA%\Python). Si no, pip
    # puede dar por instalada una biblioteca que esta ahi y no en el Python del programa.
    r = correr([python, '-E', '-s', '-m', 'pip', 'install', '-r', os.path.join(programa, 'requirements.txt')],
               capture_output=True, text=True, timeout=900, creationflags=SIN_VENTANA)
    if r.returncode != 0:
        ultimas = '\n'.join((r.stderr or r.stdout or '').strip().splitlines()[-3:])
        raise ErrorActualizacion('No pude instalar las bibliotecas nuevas:\n' + ultimas)


def _leer_requisitos(ruta):
    try:
        with open(ruta, 'rb') as f:
            # sin contar los fines de linea: una copia sacada con git en Windows tiene CRLF y el zip de GitHub LF
            return f.read().replace(b'\r\n', b'\n')
    except FileNotFoundError:
        return None


def aplicar(raiz, programa, actual, pip=instalar_dependencias, viejos=None):
    """Instala la version de 'raiz' sobre 'programa' y despues borra lo que no es de ella ni del usuario ('viejos': los
    archivos de la version anterior, o None si no se sabe). Devuelve avisos. Ante una falla antes de quedar instalada,
    deja la version anterior tal cual estaba y lo dice."""
    nuevos = sorted(archivos(raiz))
    pisados = [r for r in nuevos if os.path.exists(os.path.join(programa, r))]
    agregados = [r for r in nuevos if not os.path.exists(os.path.join(programa, r))]
    req_antes = _leer_requisitos(os.path.join(programa, 'requirements.txt'))
    apartados = os.path.join(programa, EN_CURSO)
    pip_corrio = False
    try:
        apartar(programa, pisados)
        for r in nuevos:
            _copiar(os.path.join(raiz, r), os.path.join(programa, r))
        if _leer_requisitos(os.path.join(programa, 'requirements.txt')) != req_antes:
            pip_corrio = True
            pip(programa)
    except Exception as exc:
        motivo = str(exc)
        try:
            restaurar(apartados, programa, agregados)
        except Exception as exc2:
            raise ErrorActualizacion(f'Fallo la actualizacion ({motivo}) y tampoco pude volver atras ({exc2}). '
                                     f'Los archivos de la version anterior estan en {apartados}.') from None
        try:
            _borrar(apartados)
        except OSError:
            pass
        extra = ''
        if pip_corrio:
            # pip pudo haber cambiado algunas bibliotecas antes de fallar: se vuelven a las versiones de la anterior.
            try:
                pip(programa)
            except Exception as exc3:
                extra = ('\nOjo: las bibliotecas pueden haber quedado a medio cambiar. Para arreglarlo: '
                         '"' + sys.executable.replace('pythonw.exe', 'python.exe') +
                         '" -E -s -m pip install -r requirements.txt (' + str(exc3)[:200] + ')')
        raise ErrorActualizacion(f'Fallo la actualizacion: {motivo}\nVolvi a dejar la version {actual} como '
                                 'estaba.' + extra) from None
    return limpiar(programa, nuevos, viejos)


def actualizar(rel, programa=PROGRAMA, actual=version.VERSION, avance=None, abrir=urllib.request.urlopen,
               pip=instalar_dependencias):
    'Todo el recorrido: bajar, validar, instalar y limpiar. Devuelve avisos (lista vacia si no hubo ninguno).'
    temporal = tempfile.mkdtemp(prefix='excelagent_actualizacion_')
    try:
        zip_path = descargar(rel['zip'], os.path.join(temporal, 'release.zip'), abrir, avance)
        raiz = extraer(zip_path, os.path.join(temporal, 'contenido'))
        validar(raiz, rel)
        viejos = archivos_de_version('v' + actual, abrir=abrir)
        return aplicar(raiz, programa, actual, pip, viejos)
    finally:
        shutil.rmtree(temporal, ignore_errors=True)


def relanzar(programa=PROGRAMA):
    'Arranca la version recien instalada con su propio lanzador (sin consola).'
    subprocess.Popen(['wscript.exe', '//NoLogo', os.path.join(programa, 'ExcelAgent.vbs')], creationflags=SIN_VENTANA)
