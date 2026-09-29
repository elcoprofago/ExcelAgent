import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from openpyxl import Workbook

import excel
import herramientas



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
    s = excel.Sesion(excel.Bus())
    try:
        s.abrir(ruta, visible=True, respaldo=False, minimizar=True)
        assert s.app.WindowState == excel.XL_MINIMIZADO
        assert s.app.Visible is True
    finally:
        s.salir(cerrar_excel=True, descartar=True)


def test_quitar_duplicados_con_excel_minimizado(tmp_path):
    # regresion: RemoveDuplicates falla con la ventana minimizada (error COM).
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = excel.Sesion(excel.Bus())
    try:
        s.abrir(ruta, visible=True, respaldo=False, minimizar=True)
        resultado = herramientas.ToolBox(s, approval='all').tool_remove_duplicates()
        assert '5 -> 4' in resultado
        assert s.app.WindowState == excel.XL_MINIMIZADO
        assert s.app.Visible is True
    finally:
        s.salir(cerrar_excel=True, descartar=True)


def test_operaciones_varias_con_excel_minimizado(tmp_path):
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = excel.Sesion(excel.Bus())
    try:
        s.abrir(ruta, visible=True, respaldo=False, minimizar=True)
        tb = herramientas.ToolBox(s, approval='all')
        for nombre, args in (('sort_range', {'key_column': 'B', 'descending': True}),
                             ('autofilter', {'on': True}),
                             ('format_range', {'range': 'B2:B5', 'number_format': 'currency'}),
                             ('replace', {'find': 'Sur', 'replace_with': 'Centro'}),
                             ('write_range', {'start_cell': 'B6', 'values': [['=SUM(B2:B5)']]}),
                             ('create_table', {'range': 'A1:C5'})):
            r = tb.execute(nombre, args)
            assert not r.startswith(('ERROR', 'DENIED')), nombre + ': ' + r
        assert s.app.WindowState == excel.XL_MINIMIZADO
    finally:
        s.salir(cerrar_excel=True, descartar=True)


def test_tabla_desde_una_sola_celda(tmp_path):
    # Regresion: la consigna pedia la tabla en A1 y fallaba con 'Hacen falta al menos dos filas'.
    ruta = libro_de_prueba(tmp_path, FILAS)
    s = excel.Sesion(excel.Bus())
    try:
        s.abrir(ruta, visible=False, respaldo=False)
        texto = excel.op_crear_tabla(s, 'A1')
        assert 'Tabla' in texto
        h = s.hoja()
        assert h.ListObjects.Count == 1
        direccion = str(h.ListObjects.Item(1).Range.Address).replace(chr(36), str())
        assert direccion == 'A1:C5'
    finally:
        s.salir(cerrar_excel=True, descartar=True)
