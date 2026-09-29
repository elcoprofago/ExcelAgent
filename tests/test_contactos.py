# Contactos traidos de afuera (agenda del telefono: es la que usa WhatsApp). Sin Excel.
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import contactos

# Como los exporta Android: quoted-printable con salto suave, varios TEL, item1.EMAIL, una linea plegada.
VCARD = '\r\n'.join([
    'BEGIN:VCARD', 'VERSION:2.1',
    'N;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:G=C3=B3mez;Mar=C3=ADa;;;',
    'FN;CHARSET=UTF-8;ENCODING=QUOTED-PRINTABLE:Mar=C3=ADa G=C3=B3mez =',
    '(trabajo)',
    'TEL;CELL:+54 9 11 5555-1234',
    'TEL;HOME:011 4555-1234',
    'TEL;CELL:+54 9 11 5555-1234',
    'item1.EMAIL;TYPE=INTERNET:maria@ejemplo.com',
    'END:VCARD',
    'BEGIN:VCARD', 'VERSION:3.0',
    'N:Perez;Juan;;;',
    'TEL;TYPE=CELL:+5491144443333',
    'NOTE:una nota larga que sigue',
    '  en la linea de abajo',
    'END:VCARD',
    'BEGIN:VCARD', 'VERSION:3.0', 'FN:Sin datos', 'END:VCARD',
    ''])

# Google Contactos (Exportar > CSV de Google): muchas columnas, nombres con coma entre comillas, varios
# numeros juntos con ' ::: '.
GOOGLE = ('First Name,Middle Name,Last Name,Nickname,E-mail 1 - Type,E-mail 1 - Value,Phone 1 - Type,Phone 1 - Value,'
          'Phone 2 - Type,Phone 2 - Value\n'
          'Ana,,"Lopez, hija",,* Home,ana@x.com,Mobile,+54 9 351 555-0001 ::: +54 9 351 555-0002,Home,0351 422-1111\n'
          'Beto,,,,,,Mobile,+5493515550003,,\n')


def test_vcard_una_fila_por_contacto():
    filas = contactos.filas_de_vcard(VCARD)
    assert filas[0] == ['Nombre', 'Telefono', 'Otros telefonos', 'Correo']
    assert filas[1] == ['María Gómez (trabajo)', '+54 9 11 5555-1234', '011 4555-1234', 'maria@ejemplo.com']
    assert filas[2] == ['Juan Perez', '+5491144443333', '', '']     # sin FN: arma el nombre con N
    assert filas[3] == ['Sin datos', '', '', '']
    assert len(filas) == 4


def test_csv_de_google_se_reduce_a_las_columnas_utiles():
    filas = contactos.reducir_contactos(contactos.filas_de_csv(GOOGLE))
    assert filas == [['Nombre', 'Telefono', 'Otros telefonos', 'Correo'],
                     ['Ana Lopez, hija', '+54 9 351 555-0001', '+54 9 351 555-0002 / 0351 422-1111', 'ana@x.com'],
                     ['Beto', '+5493515550003', '', '']]


def test_csv_que_no_es_de_contactos_queda_tal_cual():
    # Control: una planilla comun no se toca.
    filas = contactos.filas_de_csv('Producto;Importe\nYerba;1500\n')
    assert filas == [['Producto', 'Importe'], ['Yerba', '1500']]
    assert contactos.reducir_contactos(filas) is None


def test_que_parece_telefono():
    for v in ('+5491155551234', '+54 9 11 5555-1234', '011 4555-1234', '(011) 4555 1234'):
        assert contactos.parece_telefono(v), v
    for v in ('1500000', '3.5', '12/05/2024', '12-3', 'Ana', ''):
        assert not contactos.parece_telefono(v), v
    for h in ('Telefono', 'Otros telefonos', 'Phone 1 - Value', 'Mobile Phone', 'Teléfono móvil', 'WhatsApp'):
        assert contactos.es_columna_telefono(h), h
    for h in ('Celeste', 'Hotel', 'Importe', 'Nombre'):
        assert not contactos.es_columna_telefono(h), h
