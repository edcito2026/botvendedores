#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🤖 BOT RESPONDER A VENDEDORES
Recibe "resumen" → Valida → Obtiene datos → Responde
Mejoras: uso de variables de entorno, manejo robusto de webhook, cache simple para Excel,
mejor logging (logger.exception), uso de response.ok, quitar debug=True en app.run.
"""

from flask import Flask, request, jsonify
import sqlite3
import requests
import logging
import os
import openpyxl
from datetime import datetime
from threading import Lock

# ===== CONFIG / CREDENCIALES (leer desde variables de entorno) =====
ACCESS_TOKEN = os.environ.get('WHATSAPP_ACCESS_TOKEN')
PHONE_NUMBER_ID = os.environ.get('WHATSAPP_PHONE_NUMBER_ID')
VERIFY_TOKEN = os.environ.get('WHATSAPP_VERIFY_TOKEN')
API_VERSION = os.environ.get('WHATSAPP_API_VERSION', 'v18.0')

# Rutas / recursos (también configurables)
BD_PATH = os.environ.get('BD_PATH', 'ventas.db')
EXCEL_VENDEDORES = os.environ.get('EXCEL_VENDEDORES', 'vendedores.xlsx')

if PHONE_NUMBER_ID:
    API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"
else:
    API_URL = None

# ===== LOGGING =====
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/bot_responder.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ===== Excel cache (simple) =====
_EXCEL_CACHE = {'mtime': None, 'data': None}
_EXCEL_LOCK = Lock()


def cargar_vendedores():
    """Carga y cachea los vendedores desde el Excel para evitar abrir el archivo en cada request.
    Normaliza teléfonos guardando sólo los últimos 9 dígitos (para Perú).
    Asume: Nombre en columna A, Teléfono en columna B. Si tu archivo es distinto, ajusta EXCEL_VENDEDORES
    o pide que haga el cambio."""
    try:
        if not os.path.exists(EXCEL_VENDEDORES):
            logger.warning(f'Archivo Excel no encontrado: {EXCEL_VENDEDORES}')
            return []

        mtime = os.path.getmtime(EXCEL_VENDEDORES)
        with _EXCEL_LOCK:
            if _EXCEL_CACHE['data'] is not None and _EXCEL_CACHE['mtime'] == mtime:
                return _EXCEL_CACHE['data']

            wb = openpyxl.load_workbook(EXCEL_VENDEDORES, read_only=True, data_only=True)
            ws = wb.active
            vendedores = []

            # Ajustado: min_row=1 porque en tu archivo los datos comienzan desde la fila 1 (sin encabezado)
            for row in ws.iter_rows(min_row=1, values_only=True):
                if not row:
                    continue
                # Asumimos nombre en columna A (row[0]) y teléfono en columna B (row[1])
                nombre = row[0] if len(row) > 0 else None
                tel = row[1] if len(row) > 1 else None
                if nombre and tel:
                    tel_clean = ''.join(filter(str.isdigit, str(tel)))
                    # Guardar sólo últimos 9 dígitos (móvil peruano)
                    tel_last9 = tel_clean[-9:] if len(tel_clean) >= 9 else tel_clean
                    vendedores.append({'nombre': str(nombre).strip(), 'telefono': tel_last9})

            _EXCEL_CACHE['mtime'] = mtime
            _EXCEL_CACHE['data'] = vendedores
            logger.debug(f'Vendedores cargados: {len(vendedores)}')
            return vendedores

    except Exception:
        logger.exception('Error cargando Excel')
        return []


# ===== FUNCIONES =====


def obtener_vendedor_de_excel(numero_telefono):
    """Busca vendedor en Excel por teléfono: compara por últimos 9 dígitos."""
    try:
        numero_limpio = ''.join(filter(str.isdigit, str(numero_telefono)))
        numero_last9 = numero_limpio[-9:] if len(numero_limpio) >= 9 else numero_limpio

        vendedores = cargar_vendedores()
        logger.debug(f'Buscando vendedor para número entrante: {numero_telefono} -> last9={numero_last9}')
        logger.debug(f'Lista de teléfonos (last9) en Excel: {[v.get("telefono") for v in vendedores]}')

        for v in vendedores:
            tel = v.get('telefono', '')
            if not tel:
                continue
            if tel == numero_last9 or tel.endswith(numero_last9):
                logger.info(f'Vendedor encontrado: {v["nombre"]} para número {numero_telefono}')
                return v.get('nombre')

        logger.warning(f'No se encontró vendedor para número {numero_telefono}')
        return None

    except Exception:
        logger.exception('Error obteniendo vendedor desde Excel')
        return None


def obtener_datos_vendedor(nombre_vendedor):
    """Obtiene datos de ventas del vendedor"""
    try:
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Ventas ARCOR
        cursor.execute('''
            SELECT
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total_ventas,
                COUNT(DISTINCT Cod_Clie) as clientes,
                COUNT(DISTINCT Documento) as documentos
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR"
        ''', (nombre_vendedor,))

        venta_data = cursor.fetchone()
        if not venta_data:
            conn.close()
            return None

        ventas = {
            'total': venta_data['total_ventas'] or 0,
            'clientes': venta_data['clientes'] or 0,
            'documentos': venta_data['documentos'] or 0
        }

        # Ticket promedio
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR"
        ''', (nombre_vendedor,))

        ticket = cursor.fetchone()
        ventas['ticket'] = ticket['ticket'] or 0

        # Cuota
        cursor.execute('''
            SELECT Cuota_Soles
            FROM cuotas
            WHERE Vendedor = ? AND AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"
            LIMIT 1
        ''', (nombre_vendedor,))

        cuota_data = cursor.fetchone()
        ventas['cuota'] = cuota_data['Cuota_Soles'] if cuota_data and 'Cuota_Soles' in cuota_data.keys() else 0

        # Cumplimiento
        ventas['cumplimiento'] = (ventas['total'] / ventas['cuota'] * 100) if ventas['cuota'] > 0 else 0

        # Días restantes (se mantiene lógica original, revisar si debe excluir sábados)
        cursor.execute('''
            WITH RECURSIVE dates AS (
                SELECT DATE('2026-08-01') as fecha
                UNION ALL
                SELECT DATE(fecha, '+1 day') FROM dates
                WHERE fecha <= DATE('now')
            )
            SELECT COUNT(*) as dias FROM dates
            WHERE CAST(strftime('%w', fecha) AS INTEGER) IN (1,2,3,4,5,6)
        ''')
        dias_data = cursor.fetchone()
        dias_transcurridos = dias_data['dias'] or 1
        total_dias_habiles = 26
        dias_restantes = total_dias_habiles - dias_transcurridos

        ventas['dias_restantes'] = dias_restantes
        ventas['proyeccion'] = round(ventas['total'] + ((ventas['total'] / dias_transcurridos) * dias_restantes), 2) if dias_transcurridos > 0 else ventas['total']

        conn.close()
        return ventas

    except Exception:
        logger.exception('Error consultando BD para vendedor')
        return None


def generar_respuesta(nombre, ventas):
    """Genera mensaje de respuesta (personal o general)"""

    # Si es reporte general para Jefe/Supervisor
    if ventas.get('es_general'):
        kpis = obtener_kpis_completos()
        if not kpis:
            return "❌ Error obteniendo datos"

        # Construir categorías con emojis
        categorias_emojis = {
            'GOLOSINAS': '🍬',
            'CHOCOLATES': '🍫',
            'CHICLES': '💫',
            'GALLETAS': '🍪',
            'ALIMENTOS': '🥫'
        }

        categorias_texto = ""
        for idx, cat in enumerate(kpis.get('categorias', []), 1):
            emoji = categorias_emojis.get(cat.get('Lin_Neg', ''), '')
            if emoji:
                categorias_texto += f"\n{emoji} {cat['Lin_Neg']} - S/. {cat['venta_linea']:,.2f} ({cat['pct_participacion']:.1f}%)"
            else:
                categorias_texto += f"\n{cat['Lin_Neg']} - S/. {cat['venta_linea']:,.2f} ({cat['pct_participacion']:.1f}%)"

        return f"""📊 REPORTE ARCOR - AGOSTO
{datetime.now().strftime("%d/%m/%Y %H:%M")}

RESULTADOS ACTUALES:
Ventas: S/. {kpis['kpi_ventas']:,.2f}
Cobertura: {kpis['cobertura_actual']} clientes
Ticket Promedio: S/. {kpis['ticket_promedio']:,.2f}
TROYA: S/. {kpis['troya_venta']:,.2f} ({kpis['troya_pct']:.1f}%)

LÍNEAS DE NEGOCIO:{categorias_texto}

PROYECCIÓN AL 31/AGOSTO:
Ventas: S/. {kpis['proyeccion_ventas']:,.2f}
Cobertura: {kpis['cobertura_proyectada']} clientes

CUMPLIMIENTO: {kpis['pct_cumplimiento']:.1f}%

PENDIENTE: {kpis['dias_restantes']} días hábiles

Sistema Automatizado N&J"""

    # Si es reporte personal
    return f"""📊 REPORTE PERSONAL - AGOSTO
{datetime.now().strftime("%d/%m/%Y %H:%M")}

👤 {nombre}

TU DESEMPEÑO:
Ventas: S/. {ventas['total']:,.2f}
Cuota: S/. {ventas['cuota']:,.2f}
Cumplimiento: {ventas['cumplimiento']:.1f}%

ESTADÍSTICAS:
Clientes: {ventas['clientes']}
Documentos: {ventas['documentos']}
Ticket: S/. {ventas['ticket']:,.2f}

PROYECCIÓN:
Venta Final: S/. {ventas['proyeccion']:,.2f}
Cumpl. Final: {(ventas['proyeccion']/ventas['cuota']*100):.1f}%

Días Pendientes: {ventas['dias_restantes']}

Sistema N&J"""

def obtener_kpis_completos():
    """Obtiene todos los KPIs completos para Jefe/Supervisor"""
    try:
        logger.info('🔄 Obteniendo KPIs completos ARCOR...')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 1. KPI VENTAS TOTAL ARCOR
        cursor.execute('SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total FROM VENTAS2026 WHERE Periodo = "202608" AND Proveedor = "ARCOR"')
        kpi_ventas = cursor.fetchone()['total'] or 0

        # 2. COBERTURA ACTUAL ARCOR
        cursor.execute('SELECT COUNT(DISTINCT Cod_Clie) as total FROM VENTAS2026 WHERE Periodo = "202608" AND Proveedor = "ARCOR"')
        cobertura_actual = cursor.fetchone()['total'] or 0

        # 3. TICKET PROMEDIO ARCOR
        cursor.execute('SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket FROM VENTAS2026 WHERE Periodo = "202608" AND Proveedor = "ARCOR"')
        ticket_promedio = cursor.fetchone()['ticket'] or 0

        # 4. TROYA (Calif=D) ARCOR
        cursor.execute('SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as troya_venta FROM VENTAS2026 WHERE Periodo = "202608" AND Calif = "D" AND Proveedor = "ARCOR"')
        troya_venta = cursor.fetchone()['troya_venta'] or 0
        troya_pct = (troya_venta / kpi_ventas * 100) if kpi_ventas > 0 else 0

        # 5. CATEGORÍAS POR LÍN_NEG ARCOR
        cursor.execute('''
            SELECT Lin_Neg, ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as venta_linea
            FROM VENTAS2026
            WHERE Periodo = "202608" AND Lin_Neg != "MATERIAL POP" AND Proveedor = "ARCOR"
            GROUP BY Lin_Neg
            ORDER BY venta_linea DESC
        ''')
        categorias = []
        total_lineas = 0
        for row in cursor.fetchall():
            categorias.append(dict(row))
            total_lineas += row['venta_linea']

        for cat in categorias:
            cat['pct_participacion'] = (cat['venta_linea'] / total_lineas * 100) if total_lineas > 0 else 0

        # 6. DÍAS HÁBILES TRANSCURRIDOS
        cursor.execute('''
            WITH RECURSIVE dates AS (
                SELECT DATE('2026-08-01') as fecha
                UNION ALL
                SELECT DATE(fecha, '+1 day') FROM dates
                WHERE fecha <= DATE('now')
            )
            SELECT COUNT(*) as dias FROM dates
            WHERE CAST(strftime('%w', fecha) AS INTEGER) IN (1,2,3,4,5,6)
        ''')
        dias_transcurridos = cursor.fetchone()['dias'] or 1

        # 7. PROYECCIONES
        total_dias_habiles = 26
        dias_restantes = total_dias_habiles - dias_transcurridos
        proyeccion_cobertura = round((cobertura_actual / dias_transcurridos) * total_dias_habiles) if dias_transcurridos > 0 else 0
        proyeccion_ventas = round(kpi_ventas + ((kpi_ventas / dias_transcurridos) * dias_restantes), 2) if dias_transcurridos > 0 else 0

        # 8. % CUMPLIMIENTO ARCOR
        cursor.execute('SELECT ROUND(SUM(Cuota_Soles), 2) as cuota_total FROM cuotas WHERE AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"')
        cuota_total = cursor.fetchone()['cuota_total'] or 1
        pct_cumplimiento = (proyeccion_ventas / cuota_total * 100) if cuota_total > 0 else 0

        conn.close()

        return {
            'kpi_ventas': kpi_ventas,
            'cobertura_actual': cobertura_actual,
            'cobertura_proyectada': proyeccion_cobertura,
            'ticket_promedio': ticket_promedio,
            'troya_venta': troya_venta,
            'troya_pct': troya_pct,
            'categorias': categorias,
            'proyeccion_ventas': proyeccion_ventas,
            'pct_cumplimiento': pct_cumplimiento,
            'dias_restantes': dias_restantes
        }

    except Exception:
        logger.exception('❌ Error obteniendo KPIs')
        return None


def obtener_datos_generales():
    """Obtiene datos consolidados de TODOS los vendedores ARCOR"""
    try:
        logger.info('🔄 Consultando datos generales consolidados...')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # Total de ventas ARCOR
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total
            FROM VENTAS2026
            WHERE Periodo = "202608" AND Proveedor = "ARCOR"
        ''')
        total_ventas = cursor.fetchone()['total'] or 0

        # Cobertura (clientes únicos)
        cursor.execute('''
            SELECT COUNT(DISTINCT Cod_Clie) as total
            FROM VENTAS2026
            WHERE Periodo = "202608" AND Proveedor = "ARCOR"
        ''')
        cobertura = cursor.fetchone()['total'] or 0

        # Ticket promedio
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket
            FROM VENTAS2026
            WHERE Periodo = "202608" AND Proveedor = "ARCOR"
        ''')
        ticket = cursor.fetchone()['ticket'] or 0

        # Cuota total ARCOR
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Cuota_Soles AS REAL)), 2) as cuota
            FROM cuotas
            WHERE AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"
        ''')
        cuota_total = cursor.fetchone()['cuota'] or 0

        conn.close()

        cumplimiento = (total_ventas / cuota_total * 100) if cuota_total > 0 else 0

        logger.info('✅ Datos generales obtenidos')
        return {
            'total': total_ventas,
            'clientes': cobertura,
            'ticket': ticket,
            'cuota': cuota_total,
            'cumplimiento': cumplimiento,
            'es_general': True
        }

    except Exception:
        logger.exception('❌ Error obteniendo datos generales')
        return None


def enviar_respuesta(numero_destino, mensaje):
    """Envía mensaje por WhatsApp API"""
    try:
        if not ACCESS_TOKEN or not API_URL:
            logger.error('ACCESS_TOKEN o API_URL no configurados. Setea las variables de entorno correspondientes.')
            return False

        numero = str(numero_destino).replace('+', '').replace(' ', '')
        if not numero.startswith('51'):
            numero = f"51{numero}"

        payload = {
            "messaging_product": "whatsapp",
            "to": numero,
            "type": "text",
            "text": {"body": mensaje}
        }

        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        response = requests.post(API_URL, json=payload, headers=headers, timeout=10)

        if response.ok:
            logger.info(f'✅ Respuesta enviada a {numero}')
            return True
        else:
            logger.error(f'Error enviando ({response.status_code}): {response.text}')
            return False

    except Exception:
        logger.exception('Error enviando respuesta')
        return False


# ===== WEBHOOK =====

@app.route('/webhook', methods=['GET'])
def verificar():
    """Verifica webhook con Meta"""
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if VERIFY_TOKEN and verify_token == VERIFY_TOKEN:
        logger.info('✅ Webhook verificado')
        return challenge
    logger.warning('Webhook verificación fallida')
    return "Unauthorized", 403


@app.route('/webhook', methods=['POST'])
def recibir():
    """Recibe y responde a mensajes"""
    try:
        data = request.get_json(force=True, silent=True)
        logger.info('Mensaje recibido')

        if not data:
            logger.warning('Payload vacío o inválido')
            return jsonify({"status": "ok"}), 200

        entry = data.get('entry') or []
        if not entry:
            return jsonify({"status": "ok"}), 200

        cambios = entry[0].get('changes') or []
        if not cambios:
            return jsonify({"status": "ok"}), 200

        cambio = cambios[0].get('value') or {}
        messages = cambio.get('messages') or []
        if not messages:
            return jsonify({"status": "ok"}), 200

        mensaje_obj = messages[0]
        numero_remitente = mensaje_obj.get('from')
        texto_mensaje = (mensaje_obj.get('text') or {}).get('body', '')

        if not numero_remitente or not texto_mensaje:
            logger.info('Mensaje sin texto o sin remitente')
            return jsonify({"status": "ok"}), 200

        texto_mensaje = texto_mensaje.strip()
        logger.info(f'De: {numero_remitente} | Mensaje: "{texto_mensaje}"')

        # Verificar palabra clave
        if "resumen" not in texto_mensaje.lower():
            logger.info('Palabra clave no encontrada')
            return jsonify({"status": "ok"}), 200

        logger.info('✅ Palabra clave detectada')

        # Validar vendedor
        nombre = obtener_vendedor_de_excel(numero_remitente)

        if not nombre:
            logger.warning(f'Número no autorizado: {numero_remitente}')
            enviar_respuesta(numero_remitente, "❌ Número no autorizado")
            return jsonify({"status": "unauthorized"}), 403

        logger.info(f'✅ Vendedor válido: {nombre}')

        # Detectar si es Jefe de Ventas o Supervisor ARCOR
        es_jefe = str(nombre).upper() in ['JEFE DE VENTAS', 'SUPERVISOR ARCOR']

        if es_jefe:
            logger.info('📊 Resumen GENERAL solicitado por Jefe/Supervisor')
            ventas = obtener_datos_generales()
        else:
            ventas = obtener_datos_vendedor(nombre)

        if not ventas:
            logger.error('Sin datos de ventas')
            enviar_respuesta(numero_remitente, "⚠️ Sin datos disponibles")
            return jsonify({"status": "no_data"}), 404

        logger.info('✅ Datos obtenidos')

        # Generar y enviar respuesta
        mensaje = generar_respuesta(nombre, ventas)

        if enviar_respuesta(numero_remitente, mensaje):
            logger.info(f'✅ Reporte enviado a {nombre}')
            return jsonify({"status": "success"}), 200
        else:
            logger.error('Error enviando respuesta')
            return jsonify({"status": "send_error"}), 500

    except Exception:
        logger.exception('Error procesando webhook')
        return jsonify({"status": "error"}), 500


if __name__ == '__main__':
    logger.info('='*60)
    logger.info('🤖 BOT RESPONDER VENDEDORES INICIANDO')
    logger.info('='*60)
    logger.info(f'Escuchando: http://localhost:5000/webhook')
    logger.info(f'Palabra clave: "resumen"')
    logger.info('='*60 + '\n')

    # En producción, ejecutar con Gunicorn/uvicorn. Aquí solo para desarrollo local.
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
