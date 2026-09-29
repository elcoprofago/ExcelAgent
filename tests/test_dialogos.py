import json
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

import dialogos
import dsapi

# Todo con una configuracion de DeepSeekChat falsa y una key inventada: la real nunca se lee.
KEY_FALSA = 'sk-prueba-0123456789abcdef'


def config_de_deepseekchat(tmp_path, **campos):
    ruta = tmp_path / 'dsc' / 'config.json'
    ruta.parent.mkdir()
    ruta.write_text(json.dumps(campos), encoding='utf-8')
    return str(ruta)


def test_importa_la_key_cifrada_con_el_usuario_de_windows(tmp_path):
    ruta = config_de_deepseekchat(tmp_path, api_key_enc=dsapi.protect(KEY_FALSA), model_dirs=['E:\\Models'],
                                  llama_server_path='C:\\no\\existe\\llama-server.exe', local_ctx=20480)
    cfg = dsapi.Config(str(tmp_path / 'ea'))
    texto = dialogos.importar_de_deepseekchat(cfg, ruta)
    assert KEY_FALSA not in texto
    assert cfg.api_key == KEY_FALSA and cfg.key_mode == 'dpapi'
    assert cfg['model_dirs'][-1] == 'E:\\Models' and cfg['local_ctx'] == 20480
    # queda guardada: una configuracion nueva sobre la misma carpeta la recupera, y no en claro
    guardado = open(os.path.join(cfg.dir, 'config.json'), encoding='utf-8').read()
    assert KEY_FALSA not in guardado
    assert dsapi.Config(cfg.dir).api_key == KEY_FALSA


def test_importa_la_key_con_contrasena_sin_descifrarla(tmp_path):
    origen = dsapi.Config(str(tmp_path / 'dsc_cfg'))
    origen.set_api_key(KEY_FALSA, 'clave-de-prueba')
    ruta = config_de_deepseekchat(tmp_path, api_key_pw=origen['api_key_pw'])
    cfg = dsapi.Config(str(tmp_path / 'ea'))
    dialogos.importar_de_deepseekchat(cfg, ruta)
    assert not cfg.api_key and cfg.needs_unlock
    otra = dsapi.Config(cfg.dir)
    assert otra.needs_unlock
    otra.unlock('clave-de-prueba')
    assert otra.api_key == KEY_FALSA


def test_sin_nada_que_importar(tmp_path):
    cfg = dsapi.Config(str(tmp_path / 'ea'))
    with pytest.raises(ValueError, match='ni key ni modelos'):
        dialogos.importar_de_deepseekchat(cfg, config_de_deepseekchat(tmp_path, model='deepseek-flash'))
    with pytest.raises(ValueError, match='No encontre'):
        dialogos.importar_de_deepseekchat(cfg, str(tmp_path / 'no-existe.json'))
