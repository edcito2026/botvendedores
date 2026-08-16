#!/usr/bin/env python3
"""
🚀 EXTRACTOR OPTIMIZADO DE BD ARCOR
Extrae tablas de EXTRACTOR.db y crea ventas.db optimizado
Compatible con Windows y Linux (GitHub Actions)
"""

import sqlite3
import shutil
from datetime import datetime
import os
import sys

# ====================================
# CONFIGURACIÓN (rutas relativas)
# ====================================

# Detectar ruta base según OS
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
PARENT_PATH = os.path.dirname(BASE_PATH)  # Automatizacion\WhatsApp
REPO_PATH = os.path.dirname(os.path.dirname(PARENT_PATH))  # E:\PRUEBASEXTRACTOR

# Rutas relativas que funcionan en Windows y Linux
SOURCE_DB = os.path.join(REPO_PATH, 'EXTRACTOR.db')
TARGET_DB = os.path.join(BASE_PATH, 'ventas.db')
BACKUP_DB = os.path.join(BASE_PATH, 'ventas_backup.db')
LOG_FILE = os.path.join(BASE_PATH, 'logs', 'extract.log')

# Crear carpeta de logs si no existe
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
        pass  # Si falla logging, continúa

def conectar_db(db_path):
    """Conectar a BD"""
    try:
        conn = sqlite3.connect(db_path)
        conn.row_factory = sqlite3.Row
        return conn
    except Exception as e:
        log(f'Error conectando a {db_path}: {str(e)}', 'ERROR')
        raise

def crear_tabla_ventas(cursor_target):
    """Crear tabla VENTAS2026 optimizada en BD destino"""
    log('Creando tabla VENTAS2026...')

    cursor_target.execute('''
        CREATE TABLE IF NOT EXISTS VENTAS2026 (
            Producto TEXT,
            Proveedor TEXT,
            Lin_Neg TEXT,
            Categoria TEXT,
            Promo TEXT,
            CodProd TEXT,
            Cliente TEXT,
            Zona TEXT,
            Canal TEXT,
            Departamento TEXT,
            Distrito TEXT,
            Calif TEXT,
            Cod_Clie TEXT,
            Giro TEXT,
            Periodo TEXT,
            Dia_Sem TEXT,
            Documento TEXT,
            Forma_Pago TEXT,
            F_Emis TEXT,
            Cdg_Vend TEXT,
            Vendedor TEXT,
            Bonif TEXT,
            Cant_Fct_Vta TEXT,
            Imp_Total TEXT,
            ID TEXT PRIMARY KEY,
            TF_Gratuita TEXT
        )
    ''')

def crear_tabla_clientes(cursor_target):
    """Crear tabla clientes optimizada en BD destino"""
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

def extraer_ventas(conn_source, cursor_target):
    """Extraer VENTAS2026 con filtros ARCOR"""
    log('Extrayendo VENTAS2026 (Proveedor = ARCOR)...')

    query = '''
        SELECT *
        FROM VENTAS2026
        WHERE Proveedor = 'ARCOR'
    '''

    try:
        cursor_source = conn_source.cursor()
        cursor_source.execute(query)
        rows = cursor_source.fetchall()

        log(f'Total de registros a copiar: {len(rows)}')

        if len(rows) > 0:
            cols = [description[0] for description in cursor_source.description]
            periodo_idx = cols.index('Periodo') if 'Periodo' in cols else -1

            rows_procesadas = []
            for row in rows:
                row_list = list(row)
                if periodo_idx >= 0 and row_list[periodo_idx]:
                    try:
                        periodo_str = str(int(float(row_list[periodo_idx])))
                        row_list[periodo_idx] = periodo_str
                    except:
                        pass
                rows_procesadas.append(tuple(row_list))

            placeholders = ','.join(['?' for _ in cols])
            insert_sql = f'INSERT INTO VENTAS2026 ({",".join(cols)}) VALUES ({placeholders})'

            cursor_target.executemany(insert_sql, rows_procesadas)
            log(f'✓ {len(rows_procesadas)} registros de VENTAS2026 copiados (Periodo normalizado)')

        return len(rows)

    except Exception as e:
        log(f'Error extrayendo VENTAS2026: {str(e)}', 'ERROR')
        raise

def extraer_clientes(conn_source, cursor_target):
    """Extraer tabla clientes completa"""
    log('Extrayendo tabla clientes...')

    try:
        cursor_source = conn_source.cursor()
        cursor_source.execute('SELECT * FROM clientes')
        rows = cursor_source.fetchall()

        log(f'Total de clientes: {len(rows)}')

        if len(rows) > 0:
            cols = [description[0] for description in cursor_source.description]
            placeholders = ','.join(['?' for _ in cols])
            insert_sql = f'INSERT INTO clientes ({",".join(cols)}) VALUES ({placeholders})'

            cursor_target.executemany(insert_sql, rows)
            log(f'✓ {len(rows)} registros de clientes copiados')

        return len(rows)

    except Exception as e:
        log(f'Error extrayendo clientes: {str(e)}', 'ERROR')
        raise

def crear_tabla_cuotas(cursor_target):
    """Crear tabla cuotas optimizada en BD destino"""
    log('Creando tabla cuotas...')

    cursor_target.execute('''
        CREATE TABLE IF NOT EXISTS cuotas (
            Vendedor TEXT,
            Codigo INTEGER,
            Canal TEXT,
            Proveedor TEXT,
            Mes TEXT,
            NRO_MES INTEGER,
            AÑO INTEGER,
            Periodo DATE,
            Cuota_Soles REAL,
            Cuota_Cobertura REAL,
            PRIMARY KEY (Vendedor, AÑO, NRO_MES, Proveedor)
        )
    ''')

def extraer_cuotas(conn_source, cursor_target):
    """Extraer cuotas ARCOR desde CuotasDistribuidas"""
    log('Extrayendo cuotas ARCOR...')

    try:
        cursor_source = conn_source.cursor()
        cursor_source.execute('''
            SELECT *
            FROM CuotasDistribuidas
            WHERE Proveedor = 'ARCOR'
        ''')
        rows = cursor_source.fetchall()

        log(f'Total de cuotas a copiar: {len(rows)}')

        if len(rows) > 0:
            cols = [description[0] for description in cursor_source.description]
            placeholders = ','.join(['?' for _ in cols])
            insert_sql = f'INSERT INTO cuotas ({",".join(cols)}) VALUES ({placeholders})'

            cursor_target.executemany(insert_sql, rows)
            log(f'✓ {len(rows)} registros de cuotas ARCOR copiados')

        return len(rows)

    except Exception as e:
        log(f'Error extrayendo cuotas: {str(e)}', 'ERROR')
        raise

def crear_indices(cursor_target):
    """Crear índices para optimizar búsquedas"""
    log('Creando índices para optimización...')

    indices = [
        ('idx_ventas_cliente', 'VENTAS2026', 'Cliente'),
        ('idx_ventas_proveedor', 'VENTAS2026', 'Proveedor'),
        ('idx_ventas_fecha', 'VENTAS2026', 'F_Emis'),
        ('idx_ventas_codprod', 'VENTAS2026', 'CodProd'),
        ('idx_ventas_vendedor', 'VENTAS2026', 'Vendedor'),
        ('idx_clientes_razonsocial', 'clientes', 'Raz_Social'),
        ('idx_clientes_zona', 'clientes', 'Zona'),
        ('idx_clientes_vendedor', 'clientes', 'Vendedor'),
        ('idx_cuotas_vendedor', 'cuotas', 'Vendedor'),
        ('idx_cuotas_año_mes', 'cuotas', 'AÑO'),
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
    print('\n' + '='*60)
    print('  🚀 EXTRACTOR OPTIMIZADO DE BD ARCOR')
    print('='*60 + '\n')

    inicio = datetime.now()
    log('Iniciando extracción...')

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

        crear_tabla_ventas(cursor_target)
        crear_tabla_clientes(cursor_target)
        crear_tabla_cuotas(cursor_target)

        ventas_count = extraer_ventas(conn_source, cursor_target)
        clientes_count = extraer_clientes(conn_source, cursor_target)
        cuotas_count = extraer_cuotas(conn_source, cursor_target)

        crear_indices(cursor_target)

        conn_target.commit()
        conn_source.close()
        conn_target.close()

        duracion = datetime.now() - inicio
        tamaño_destino = obtener_tamaño_archivo(TARGET_DB)
        tamaño_fuente = obtener_tamaño_archivo(SOURCE_DB)

        print('\n' + '='*60)
        print('  ✅ EXTRACCIÓN COMPLETADA EXITOSAMENTE')
        print('='*60)
        print(f'\n📊 RESUMEN:')
        print(f'  BD Fuente: {SOURCE_DB}')
        print(f'  Tamaño: {tamaño_fuente}')
        print(f'  BD Destino: {TARGET_DB}')
        print(f'  Tamaño: {tamaño_destino}')
        print(f'\n📈 DATOS EXTRAÍDOS:')
        print(f'  ✓ VENTAS2026: {ventas_count:,}')
        print(f'  ✓ clientes: {clientes_count:,}')
        print(f'  ✓ cuotas: {cuotas_count:,}')
        print(f'\n⏱️  TIEMPO: {duracion.total_seconds():.2f} segundos')
        print(f'\n📋 FILTROS: Proveedor = ARCOR')
        print('\n' + '='*60 + '\n')

        log(f'✅ Completada en {duracion.total_seconds():.2f} seg')
        return True

    except Exception as e:
        log(f'❌ Error: {str(e)}', 'ERROR')
        print(f'\n❌ ERROR: {str(e)}\n')
        return False

if __name__ == '__main__':
    exito = main()
    if len(sys.argv) == 1:  # Solo pide input si se ejecuta manualmente
        input('Presiona Enter para cerrar...')
