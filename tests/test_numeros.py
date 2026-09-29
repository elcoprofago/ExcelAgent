# Numeros y fechas escritos como texto (numeros.py). Sin Excel: solo la interpretacion.
import datetime as dt
import os
import sys

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

import pytest

import numeros
from numeros import Texto

CASOS = [
    # texto, separador decimal de la columna, orden de fecha, valor, formato
    ('1500', ',', 'dmy', 1500, None),
    ('1.500', ',', 'dmy', 1500, '#,##0'),
    ('1.500', '.', 'dmy', 1.5, None),
    ('1,500', ',', 'dmy', 1.5, None),
    ('1,500', '.', 'dmy', 1500, '#,##0'),
    ('1.500,50', ',', 'dmy', 1500.5, '#,##0.00'),
    ('1,500.50', ',', 'dmy', 1500.5, '#,##0.00'),       # sin ambiguedad: el ultimo separador es el decimal
    ('1500,5', ',', 'dmy', 1500.5, None),
    ('3.25', ',', 'dmy', 3.25, None),                   # dos decimales: no puede ser separador de miles
    ('0,5', ',', 'dmy', 0.5, None),
    ('0.500', ',', 'dmy', 0.5, None),
    ('0', ',', 'dmy', 0, None),
    ('-12', ',', 'dmy', -12, None),
    ('1.234.567', ',', 'dmy', 1234567, '#,##0'),
    ('1 500 000', ',', 'dmy', 1500000, '#,##0'),
    ('$ 1.500', ',', 'dmy', 1500, '$ #,##0.00'),
    ('-$1.500,25', ',', 'dmy', -1500.25, '$ #,##0.00'),
    ('US$ 20', ',', 'dmy', 20, '"US$ "#,##0.00'),
    ('15%', ',', 'dmy', 0.15, '0%'),
    ('12,5 %', ',', 'dmy', 0.125, '0.0%'),
    ('01/09/2026', ',', 'dmy', dt.date(2026, 9, 1), 'dd/mm/yyyy'),
    ('01/09/2026', ',', 'mdy', dt.date(2026, 1, 9), 'dd/mm/yyyy'),
    ('25/09/2026', ',', 'mdy', dt.date(2026, 9, 25), 'dd/mm/yyyy'),   # 25 no es mes: manda el dato
    ('1-9-26', ',', 'dmy', dt.date(2026, 9, 1), 'dd/mm/yyyy'),
    ('01.09.2026', ',', 'dmy', dt.date(2026, 9, 1), 'dd/mm/yyyy'),
    ('2026-09-01', ',', 'dmy', dt.date(2026, 9, 1), 'dd/mm/yyyy'),
    ('01/09/2026 14:30', ',', 'dmy', dt.datetime(2026, 9, 1, 14, 30), 'dd/mm/yyyy hh:mm'),
    ('14:30', ',', 'dmy', dt.time(14, 30), 'hh:mm'),
    # Controles: siguen siendo texto
    ('007', ',', 'dmy', Texto('007'), None),
    ('+5491144443333', ',', 'dmy', Texto('+5491144443333'), None),
    ('2850590940090418135201', ',', 'dmy', Texto('2850590940090418135201'), None),   # CBU: 22 cifras
    ('1.10.20', ',', 'dmy', '1.10.20', None),
    ('192.168.0.1', ',', 'dmy', '192.168.0.1', None),
    ('31/02/2026', ',', 'dmy', '31/02/2026', None),
    ('Ana', ',', 'dmy', 'Ana', None),
    ('1.5.3', ',', 'dmy', '1.5.3', None),
    ('12,34,5', ',', 'dmy', '12,34,5', None),
    ('25:99', ',', 'dmy', '25:99', None),
    ('A-12', ',', 'dmy', 'A-12', None),
    ('', ',', 'dmy', '', None),
]


@pytest.mark.parametrize('texto, decimal, orden, valor, formato', CASOS)
def test_interpretar(texto, decimal, orden, valor, formato):
    v, f = numeros.interpretar(texto, decimal, orden)
    assert (v, type(v), f) == (valor, type(valor), formato)


def test_la_columna_decide_lo_ambiguo():
    assert numeros.convencion(['1.500', '2.300,50', '980']) == (',', 'dmy')
    assert numeros.convencion(['1.500', '3.25']) == ('.', 'dmy')
    assert numeros.convencion(['1.500']) == (',', 'dmy')                 # sin pistas: la del usuario
    assert numeros.convencion(['1.500'], decimal='.') == ('.', 'dmy')
    assert numeros.convencion(['09/25/2026', '01/09/2026']) == (',', 'mdy')
    assert numeros.convencion(['1,5', '2.5']) == (',', 'dmy')            # empate: la del usuario
