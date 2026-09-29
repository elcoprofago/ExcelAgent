import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import dsapi

# Con una key inventada: la real nunca se lee.
KEY_FALSA = 'sk-prueba-0123456789abcdef'


def guardado(cfg):
    with open(os.path.join(cfg.dir, 'config.json'), encoding='utf-8') as f:
        return f.read()


def test_key_cifrada_con_el_usuario_de_windows(tmp_path):
    cfg = dsapi.Config(str(tmp_path / 'cfg'))
    cfg.set_api_key(KEY_FALSA)
    assert cfg.key_mode == 'dpapi'
    assert KEY_FALSA not in guardado(cfg)
    assert dsapi.Config(cfg.dir).api_key == KEY_FALSA


def test_key_con_contrasena(tmp_path):
    cfg = dsapi.Config(str(tmp_path / 'cfg'))
    cfg.set_api_key(KEY_FALSA, 'clave-de-prueba')
    assert KEY_FALSA not in guardado(cfg)
    otra = dsapi.Config(cfg.dir)
    assert otra.needs_unlock and not otra.api_key
    otra.unlock('clave-de-prueba')
    assert otra.api_key == KEY_FALSA
