# A que Excel se conecta la sesion. Con Excel falsos: no abre ninguno de verdad.
# Regresion del 29/09/2026: la sesion adopto el Excel oculto que usaban las pruebas (otro programa que
# automatiza Excel haria lo mismo), le cambio DisplayAlerts y Visible, y al abrir el libro fallo con
# 'Excel no pudo abrir el archivo'.
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest
import win32com.client

import excel


class Libro:
    def __init__(self, ruta):
        self.FullName = ruta
        self.Worksheets = [type('Hoja', (), {'Name': 'Hoja1'})()]


class Libros(list):
    def __init__(self, app):
        super().__init__()
        self.app = app

    @property
    def Count(self):
        return len(self)

    def Open(self, ruta):
        if self.app.open_no_deja_nada:
            return None             # como un Excel que se esta cerrando: no falla y no abre
        wb = Libro(ruta)
        self.append(wb)
        return wb


class App:
    def __init__(self, visible, open_no_deja_nada=False):
        self.Visible = visible
        self.DisplayAlerts = True
        self.WindowState = -4143
        self.open_no_deja_nada = open_no_deja_nada
        self.Workbooks = Libros(self)


@pytest.fixture
def entorno(monkeypatch, tmp_path):
    ruta = tmp_path / 'libro.xlsx'
    ruta.write_bytes(b'x')
    nuevas = []

    def nueva(_):
        app = App(visible=False)
        nuevas.append(app)
        return app

    monkeypatch.setattr(win32com.client, 'DispatchEx', nueva)

    def usar_activa(app):
        monkeypatch.setattr(win32com.client, 'GetActiveObject', lambda _: app)
    return str(ruta), nuevas, usar_activa


def test_no_adopta_ni_toca_un_excel_sin_ventana(entorno):
    ruta, nuevas, usar_activa = entorno
    ajeno = App(visible=False)
    usar_activa(ajeno)
    s = excel.Sesion(excel.Bus())
    s.abrir(ruta, visible=False, respaldo=False, minimizar=False)
    assert s.propia and s.app is nuevas[0]
    assert ajeno.DisplayAlerts is True and ajeno.Visible is False and ajeno.Workbooks.Count == 0


def test_adopta_el_excel_con_ventana_del_usuario(entorno):
    # Control: el Excel del usuario, con ventana, se sigue usando.
    ruta, nuevas, usar_activa = entorno
    del_usuario = App(visible=True)
    usar_activa(del_usuario)
    s = excel.Sesion(excel.Bus())
    s.abrir(ruta, visible=True, respaldo=False, minimizar=False)
    assert not s.propia and s.app is del_usuario and not nuevas
    assert del_usuario.Workbooks.Count == 1


def test_si_el_excel_adoptado_no_deja_el_libro_reintenta_en_uno_propio(entorno):
    ruta, nuevas, usar_activa = entorno
    usar_activa(App(visible=True, open_no_deja_nada=True))
    s = excel.Sesion(excel.Bus())
    s.abrir(ruta, visible=True, respaldo=False, minimizar=False)
    assert s.propia and s.app is nuevas[0] and s.wb is not None
