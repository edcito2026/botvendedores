#!/usr/bin/env python3
# Small test script to validate phone matching logic
import os
import sys

# Ensure the repo root is on path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)) + '/../')

from BOT_RESPONDER_VENDEDORES import cargar_vendedores, obtener_vendedor_de_excel

TEST_NUMBERS = [
    '+51970507377',
    '51970507377',
    '970507377',
    '970-507-377',
    '970 507 377'
]

if __name__ == '__main__':
    print('Cargando vendedores desde:', os.environ.get('EXCEL_VENDEDORES', 'vendedores.xlsx'))
    v = cargar_vendedores()
    print('Vendedores cargados:', len(v))
    for n in TEST_NUMBERS:
        nombre = obtener_vendedor_de_excel(n)
        print(f'{n} -> {nombre}')
