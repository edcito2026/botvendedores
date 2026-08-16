#!/usr/bin/env python3
"""
🚀 EXTRACTOR OPTIMIZADO DE BD ARCOR - V2
✅ Solo período actual (mes en curso)
✅ Columnas optimizadas (sin redundancias)
✅ Tamaño reducido: 67MB → ~5-8MB
✅ Mantiene tabla CLIENTES para futuro
"""

import sqlite3
import shutil
from datetime import datetime
import os
import sys
from calendar import monthrange

BASE_PATH = os.path.dirname(os.path.abspath(__file__))
PARENT_PATH = os.path.dirname(BASE_PATH)
REPO_PATH = os.path.dirname(os.path.dirname(os.path.dirname(BASE_PATH)))

SOURCE_DB = os.path.join(REPO_PATH, 'EXTRACTOR.db')
TARGET_DB = os.path.join(BASE_PATH, 'ventas.db')
BACKUP_DB = os.path.join(BASE_PATH, 'ventas_backup.db')
LOG_FILE = os.path.join(BASE_PATH, 'logs', 'extract.log')

os.makedirs(os.path.dirname(LOG_FILE), exist_ok=True)

def log(msg, level='INFO'):
    """Guardar mensaje en log"""
    timestamp = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    log_msg = f'[{timestamp}] {level}: {msg}'
    print(log_msg)
    try:
        with open(LOG_FILE, 'a', encoding='utf-8') as f:
            f.write(log_msg + '\n')
    except:
        pass

def conectar_db(db_path):
    """Conectar a BD"""
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn
    except Exception as e:
        log(f'Error conectando a {db_path}: {str(e)}', 'ERROR')
        raise

def obtener_periodo_actual():
    """Obtener período actual (YYYYMM)"""
    ahora = datetime.now()
    periodo = f"{ahora.year}{ahora.month:02d}"
    return periodo, ahora.year, ahora.month

def crear_tabla_ventas_optimizada(cursor_target):
    """Crear tabla VENTAS2026 OPTIMIZADA - Solo columnas necesarias"""
    log('Creando tabla VENTAS2026 optimizada...')

    cursor_target.execute('''
        CREATE TABLE IF NOT EXISTS VENTAS2026 (
            Vendedor TEXT,
            Imp_Total REAL,
            Cod_Clie TEXT,
            Documento TEXT,
            Periodo TEXT,
            Proveedor TEXT,
            Calif TEXT,
            lin_neg TEXT,
            PRIMARY KEY (Periodo, Documento, Cod_Clie, Vendedor)
        )
    ''')

def crear_tabla_clientes(cursor_target):
    """Crear tabla clientes COMPLETA (para futuro)"""
    log('Creando tabla clientes...')

    cursor_target.execute('''
        CREATE TABLE IF NOT EXISTS clientes (
            Cod_Clie TEXT PRIMARY KEY,
            Raz_Social TEXT,
            Direccion TEXT,
            Provincia TEXT,
            Distrito TEXT,
            Zona TEXT,
            Canal TEXT,
            Calif TEXT,
            Giro TEXT,
            Distrito1 TEXT,
            Cdg_Ubigeo TEXT,
            Latitud TEXT,
            Longitud TEXT,
            Estado TEXT,
            Cdg_Vend TEXT,
            Vendedor TEXT,
            Nom_Ruta TEXT,
            DV TEXT,
            Lim_Cred TEXT,
            Cat_Clie TEXT,
            Tlfno TEXT
        )
    ''')

def crear_tabla_cuotas_optimizada(cursor_target):
    """Crear tabla cuotas OPTIMIZADA"""
    log('Creando tabla cuotas optimizada...')

    cursor_target.execute('''
        CREATE TABLE IF NOT EXISTS cuotas (
            Vendedor TEXT,
            NRO_MES INTEGER,
            AÑO INTEGER,
            Cuota_Soles REAL,
            Cuota_Cobertura REAL,
            Proveedor TEXT,
            PRIMARY KEY (Vendedor, AÑO, NRO_MES, Proveedor)
        )
    ''')

def extraer_ventas_optimizado(conn_source, cursor_target, periodo, año):
    """Extrae VENTAS2026 - SOLO período actual y columnas necesarias"""
    log(f'📅 Extrayendo VENTAS2026 (Período: {periodo}, Proveedor: ARCOR)...')

    query = '''
        SELECT
            Vendedor,
            CAST(Imp_Total AS REAL) as Imp_Total,
            Cod_Clie,
            Documento,
            Periodo,
            Proveedor,
            Calif,
            COALESCE(lin_neg, 'ARCOR') as lin_neg
        FROM VENTAS2026
        WHERE Proveedor = 'ARCOR'
        AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
    '''

    try:
        cursor_source = conn_source.cursor()
        cursor_source.execute(query, (periodo,))
        rows = cursor_source.fetchall()

        log(f'Total de registros a copiar: {len(rows):,}')

        if len(rows) > 0:
            cols = ['Vendedor', 'Imp_Total', 'Cod_Clie', 'Documento', 'Periodo', 'Proveedor', 'Calif', 'lin_neg']
            placeholders = ','.join(['?' for _ in cols])
            insert_sql = f'INSERT OR IGNORE INTO VENTAS2026 ({",".join(cols)}) VALUES ({placeholders})'

            cursor_target.executemany(insert_sql, rows)
            log(f'✓ {len(rows):,} registros de VENTAS2026 copiados (período {periodo})')

        return len(rows)

    except Exception as e:
        log(f'Error extrayendo VENTAS2026: {str(e)}', 'ERROR')
        raise

def extraer_clientes_completo(conn_source, cursor_target):
    """Extrae tabla clientes COMPLETA"""
    log('Extrayendo tabla clientes (completa para futuro)...')

    try:
        cursor_source = conn_source.cursor()
        cursor_source.execute('SELECT * FROM clientes')
        rows = cursor_source.fetchall()

        log(f'Total de clientes: {len(rows):,}')

        if len(rows) > 0:
            cols = [description[0] for description in cursor_source.description]
            placeholders = ','.join(['?' for _ in cols])
            insert_sql = f'INSERT OR IGNORE INTO clientes ({",".join(cols)}) VALUES ({placeholders})'

            cursor_target.executemany(insert_sql, rows)
            log(f'✓ {len(rows):,} registros de clientes copiados')

        return len(rows)

    except Exception as e:
        log(f'Error extrayendo clientes: {str(e)}', 'ERROR')
        raise

def extraer_cuotas_optimizado(conn_source, cursor_target, periodo, año, mes):
    """Extrae cuotas - SOLO mes actual y columnas necesarias"""
    log(f'Extrayendo cuotas ARCOR (Mes: {mes}, Año: {año})...')

    try:
        cursor_source = conn_source.cursor()
        cursor_source.execute('''
            SELECT
                Vendedor,
                NRO_MES,
                AÑO,
                CAST(Cuota_Soles AS REAL) as Cuota_Soles,
                CAST(Cuota_Cobertura AS REAL) as Cuota_Cobertura,
                Proveedor
            FROM CuotasDistribuidas
            WHERE Proveedor = 'ARCOR'
            AND NRO_MES = ?
            AND AÑO = ?
        ''', (mes, año))

        rows = cursor_source.fetchall()

        log(f'Total de cuotas a copiar: {len(rows):,}')

        if len(rows) > 0:
            cols = ['Vendedor', 'NRO_MES', 'AÑO', 'Cuota_Soles', 'Cuota_Cobertura', 'Proveedor']
            placeholders = ','.join(['?' for _ in cols])
            insert_sql = f'INSERT OR IGNORE INTO cuotas ({",".join(cols)}) VALUES ({placeholders})'

            cursor_target.executemany(insert_sql, rows)
            log(f'✓ {len(rows):,} registros de cuotas ARCOR copiados (mes {mes}/{año})')

        return len(rows)

    except Exception as e:
        log(f'Error extrayendo cuotas: {str(e)}', 'ERROR')
        raise

def crear_indices(cursor_target):
    """Crear índices para optimización"""
    log('Creando índices de optimización...')

    indices = [
        ('idx_ventas_vendedor', 'VENTAS2026', 'Vendedor'),
        ('idx_ventas_cliente', 'VENTAS2026', 'Cod_Clie'),
        ('idx_ventas_periodo', 'VENTAS2026', 'Periodo'),
        ('idx_clientes_vendedor', 'clientes', 'Vendedor'),
        ('idx_cuotas_vendedor', 'cuotas', 'Vendedor'),
    ]

    for idx_name, tabla, columna in indices:
        try:
            cursor_target.execute(f'CREATE INDEX IF NOT EXISTS {idx_name} ON {tabla}({columna})')
        except Exception as e:
            log(f'Advertencia creando índice {idx_name}: {str(e)}', 'WARN')

def obtener_tamaño_archivo(ruta):
    """Obtener tamaño de archivo en MB"""
    try:
        size_bytes = os.path.getsize(ruta)
        size_mb = size_bytes / (1024 * 1024)
        return f'{size_mb:.2f} MB'
    except:
        return 'N/A'

def main():
    """Proceso principal"""
    print('\n' + '='*70)
    print('  🚀 EXTRACTOR OPTIMIZADO DE BD ARCOR - V2')
    print('='*70 + '\n')

    inicio = datetime.now()
    log('Iniciando extracción optimizada...')

    # Obtener período actual
    periodo_actual, año, mes = obtener_periodo_actual()
    print(f'📅 Período actual: {periodo_actual} ({mes}/{año})')
    print(f'✅ Extrayendo SOLO datos del mes en curso\n')

    if not os.path.exists(SOURCE_DB):
        log(f'❌ BD fuente no encontrada: {SOURCE_DB}', 'ERROR')
        return False

    log(f'BD fuente: {SOURCE_DB} ({obtener_tamaño_archivo(SOURCE_DB)})')

    try:
        conn_source = conectar_db(SOURCE_DB)

        if os.path.exists(TARGET_DB):
            log('Respaldando BD anterior...')
            shutil.copy(TARGET_DB, BACKUP_DB)
            os.remove(TARGET_DB)

        conn_target = conectar_db(TARGET_DB)
        cursor_target = conn_target.cursor()

        # Crear tablas optimizadas
        crear_tabla_ventas_optimizada(cursor_target)
        crear_tabla_clientes(cursor_target)
        crear_tabla_cuotas_optimizada(cursor_target)

        # Extraer datos
        ventas_count = extraer_ventas_optimizado(conn_source, cursor_target, int(periodo_actual), año)
        clientes_count = extraer_clientes_completo(conn_source, cursor_target)
        cuotas_count = extraer_cuotas_optimizado(conn_source, cursor_target, periodo_actual, año, mes)

        # Crear índices
        crear_indices(cursor_target)

        # Confirmar cambios
        conn_target.commit()
        conn_source.close()
        conn_target.close()

        duracion = datetime.now() - inicio
        tamaño_destino = obtener_tamaño_archivo(TARGET_DB)
        tamaño_fuente = obtener_tamaño_archivo(SOURCE_DB)

        print('\n' + '='*70)
        print('  ✅ EXTRACCIÓN COMPLETADA EXITOSAMENTE')
        print('='*70)
        print(f'\n📊 RESUMEN:')
        print(f'  BD Fuente:        {tamaño_fuente}')
        print(f'  BD Destino:       {tamaño_destino}')
        print(f'\n📈 DATOS EXTRAÍDOS:')
        print(f'  ✓ VENTAS2026 (período {periodo_actual}): {ventas_count:,} registros')
        print(f'  ✓ clientes (completa):          {clientes_count:,} registros')
        print(f'  ✓ cuotas (mes {mes}/{año}):             {cuotas_count:,} registros')
        print(f'\n🔍 OPTIMIZACIONES:')
        print(f'  • Período extraído: {periodo_actual}')
        print(f'  • Columnas reducidas: VENTAS2026 (26→8 con Calif, lin_neg), CUOTAS (10→5)')
        print(f'  • Tabla CLIENTES: Completa (preservada para futuro)')
        print(f'  • Conversión de Periodo: CAST(CAST(Periodo AS FLOAT) AS INTEGER)')
        print(f'\n⏱️  TIEMPO: {duracion.total_seconds():.2f} segundos')
        print('\n' + '='*70 + '\n')

        log(f'✅ Extracción optimizada completada en {duracion.total_seconds():.2f} seg')
        return True

    except Exception as e:
        log(f'❌ Error: {str(e)}', 'ERROR')
        print(f'\n❌ ERROR: {str(e)}\n')
        return False

if __name__ == '__main__':
    exito = main()
    if len(sys.argv) == 1:
        input('Presiona Enter para cerrar...')
