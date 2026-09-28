import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook

import app



def libro_de_prueba(tmp_path, filas):
    ruta = tmp_path / 'mini.xlsx'
    wb = Workbook()
    ws = wb.active
    ws.title = 'Datos'
    for fila in filas:
        ws.append(list(fila))
    wb.save(str(ruta))
    return str(ruta)


FILAS = (('Cliente', 'Importe', 'Zona'),
         ('ACME', 1000, 'Norte'),
         ('ACME', 250, 'Sur'),
         ('Zenit', 700, 'Norte'),
         ('ACME', 1000, 'Norte'))


def test_abre_minimizado_y_deja_la_ventana_minimizada(tmp_path):
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = app.Sesion(app.Bus())
    try:
        s.abrir(ruta, visible=True, respaldo=False, minimizar=True)
        assert s.app.WindowState == app.XL_MINIMIZADO
        assert s.app.Visible is True
    finally:
        s.salir(cerrar_excel=True)


def test_quitar_duplicados_con_excel_minimizado(tmp_path):
    # regresion: RemoveDuplicates falla con la ventana minimizada (error COM).
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = app.Sesion(app.Bus())
    try:
        s.abrir(ruta, visible=True, respaldo=False, minimizar=True)
        resultado = app.op_quitar_duplicados(s)
        assert '5 -> 4' in resultado
        assert s.app.WindowState == app.XL_MINIMIZADO
        assert s.app.Visible is True
    finally:
        s.salir(cerrar_excel=True)


def test_operaciones_varias_con_excel_minimizado(tmp_path):
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = app.Sesion(app.Bus())
    try:
        s.abrir(ruta, visible=True, respaldo=False, minimizar=True)
        app.op_ordenar(s, 'B', False)
        app.op_filtro(s, True)
        app.op_formato(s, 'B2:B5', num=app.FORMATOS['moneda'])
        app.op_reemplazar(s, 'Sur', 'Centro')
        app.op_resumen(s, 'B', 'SUM', 'TOTAL')
        app.op_crear_tabla(s, 'A1:C5')
        assert s.app.WindowState == app.XL_MINIMIZADO
    finally:
        s.salir(cerrar_excel=True)


def test_tabla_desde_una_sola_celda(tmp_path):
    # Regresion: la consigna pedia la tabla en A1 y fallaba con 'Hacen falta al menos dos filas'.
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = app.Sesion(app.Bus())
    try:
        s.abrir(ruta, visible=False, respaldo=False)
        texto = app.op_crear_tabla(s, 'A1')
        assert 'Tabla' in texto
        h = s.hoja()
        assert h.ListObjects.Count == 1
        direccion = str(h.ListObjects.Item(1).Range.Address).replace(chr(36), str())
        assert direccion == 'A1:C5'
    finally:
        s.salir(cerrar_excel=True)