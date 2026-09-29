import os
import subprocess
import sys

import pytest

# Las pruebas de los modulos copiados de DeepSeekChat (dsapi, secret, meter, localmodels) son scripts que
# informan por codigo de salida: se corren aparte, cada uno con su propia carpeta de configuracion.
CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'de_deepseekchat')


@pytest.mark.parametrize('nombre', ['meter', 'config', 'dsapi', 'retry', 'local'])
def test_script_de_deepseekchat(nombre, tmp_path):
    env = dict(os.environ, EXCELAGENT_CONFIG_DIR=str(tmp_path / 'cfg'), PYTHONDONTWRITEBYTECODE='1',
               PYTHONPATH=os.pathsep.join(p for p in sys.path if p))
    r = subprocess.run([sys.executable, os.path.join(CARPETA, 'chk_' + nombre + '.py')], capture_output=True,
                       text=True, encoding='utf-8', errors='replace', env=env, timeout=600, cwd=str(tmp_path))
    fallas = [l for l in r.stdout.splitlines() if l.startswith('FALLA')]
    assert r.returncode == 0, '\n'.join(fallas) or (r.stdout[-1500:] + r.stderr[-1500:])
