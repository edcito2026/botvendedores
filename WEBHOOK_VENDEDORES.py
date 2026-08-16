#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🔔 WEBHOOK RECEPCIÓN DE MENSAJES WHATSAPP
Recibe mensajes de vendedores, valida su número, extrae ventas personalizadas
y envía reporte dinámico por WhatsApp API
FIXED: Índices Excel, palabra clave "resumen", min_row=1 (sin encabezados)
"""

from flask import Flask, request, jsonify
import sqlite3
import requests
import logging
from datetime import datetime
import json
import openpyxl
import os
from threading import Lock

# ===== CONFIGURACIÓN (variables de entorno) =====
ACCESS_TOKEN = os.environ.get('WHATSAPP_ACCESS_TOKEN')
PHONE_NUMBER_ID = os.environ.get('WHATSAPP_PHONE_NUMBER_ID')
VERIFY_TOKEN = os.environ.get('WHATSAPP_VERIFY_TOKEN')
API_VERSION = os.environ.get('WHATSAPP_API_VERSION', 'v18.0')

# Palabra clave para solicitar reporte
PALABRA_CLAVE = "resumen"

# Rutas (relativas + fallback)
BD_PATH = os.environ.get('BD_PATH', 'ventas.db')
EXCEL_VENDEDORES = os.environ.get('EXCEL_VENDEDORES', 'vendedores.xlsx')

if PHONE_NUMBER_ID:
    API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"
else:
    API_URL = None

# Logging
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/webhook_vendedores.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ===== Excel cache (thread-safe) =====
_EXCEL_CACHE = {'mtime': None, 'data': None}
_EXCEL_LOCK = Lock()


# ===== 1. VALIDAR VENDEDOR =====
def obtener_vendedores_autorizados():
    """Lee lista de vendedores autorizados desde Excel (con cache thread-safe)"""
    try:
        if not os.path.exists(EXCEL_VENDEDORES):
            logger.warning(f'Archivo Excel no encontrado: {EXCEL_VENDEDORES}')
            return {}

        mtime = os.path.getmtime(EXCEL_VENDEDORES)
        with _EXCEL_LOCK:
            if _EXCEL_CACHE['data'] is not None and _EXCEL_CACHE['mtime'] == mtime:
                return _EXCEL_CACHE['data']

            logger.info('📋 Leyendo vendedores autorizados desde Excel...')
            vendedores = {}

            wb = openpyxl.load_workbook(EXCEL_VENDEDORES, read_only=True, data_only=True)
            ws = wb.active

            # Estructura correcta del Excel (sin encabezados):
            # Fila 1+: Datos (ABDEL MARTIN... | 970507377 | 19)
            # Columna A: Nombre | B: Teléfono | C: Clientes
            for row in ws.iter_rows(min_row=1, values_only=True):
                if not row or len(row) < 3:
                    continue

                if row[1]:  # Teléfono en columna B (index 1)
                    telefono = str(row[1]).strip()
                    # Normalizar teléfono: quitar espacios y caracteres especiales
                    telefono_limpio = ''.join(filter(str.isdigit, telefono))
                    # Usar últimos 9 dígitos (número móvil peruano)
                    telefono_last9 = telefono_limpio[-9:] if len(telefono_limpio) >= 9 else telefono_limpio

                    vendedores[telefono_last9] = {
                        'nombre': str(row[0]).strip() if row[0] else None,  # Nombre en columna A (index 0)
                        'codigo': None,  # No hay código en este Excel
                        'telefono': telefono_last9
                    }

            _EXCEL_CACHE['mtime'] = mtime
            _EXCEL_CACHE['data'] = vendedores
            logger.info(f'✅ {len(vendedores)} vendedores autorizados cargados')
            return vendedores

    except Exception as e:
        logger.exception(f'❌ Error leyendo Excel: {e}')
        return {}


def validar_vendedor(numero_telefonico):
    """Valida si el número de teléfono está autorizado (compara últimos 9 dígitos)"""
    # Normalizar número
    numero_limpio = ''.join(filter(str.isdigit, numero_telefonico))
    # Usar últimos 9 dígitos (número móvil peruano)
    numero_last9 = numero_limpio[-9:] if len(numero_limpio) >= 9 else numero_limpio

    logger.debug(f'Buscando vendedor para: {numero_telefonico} -> last9={numero_last9}')

    vendedores = obtener_vendedores_autorizados()

    if numero_last9 in vendedores:
        logger.info(f'✅ Vendedor autorizado: {vendedores[numero_last9]["nombre"]}')
        return True, vendedores[numero_last9]
    else:
        logger.warning(f'⚠️  Número no autorizado: {numero_telefonico}')
        logger.debug(f'Números disponibles: {list(vendedores.keys())}')
        return False, None


# ===== 2. EXTRAER DATOS DE VENTAS DEL VENDEDOR =====
def obtener_ventas_vendedor(nombre_vendedor):
    """Obtiene datos de ventas del vendedor específico desde BD"""
    try:
        logger.info(f'🔄 Consultando ventas de {nombre_vendedor}...')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 1. Ventas totales
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

        # 2. Ticket promedio
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR"
        ''', (nombre_vendedor,))

        ticket = cursor.fetchone()
        ventas['ticket'] = ticket['ticket'] or 0

        # 3. Cuota
        cursor.execute('''
            SELECT Cuota_Soles
            FROM cuotas
            WHERE Vendedor = ? AND AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"
            LIMIT 1
        ''', (nombre_vendedor,))

        cuota_data = cursor.fetchone()
        ventas['cuota'] = cuota_data['Cuota_Soles'] if cuota_data else 0

        # 4. Cumplimiento
        ventas['cumplimiento'] = (ventas['total'] / ventas['cuota'] * 100) if ventas['cuota'] > 0 else 0

        # 5. Proyección (usando fórmula: Venta + (Venta/Días) * Días_Restantes)
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
        dias_restantes = 26 - dias_transcurridos

        ventas['dias_restantes'] = dias_restantes
        ventas['proyeccion'] = round(ventas['total'] + ((ventas['total'] / dias_transcurridos) * dias_restantes), 2)

        conn.close()
        logger.info(f'✅ Datos obtenidos para {nombre_vendedor}')
        return ventas

    except Exception as e:
        logger.error(f'❌ Error obteniendo ventas: {e}')
        return None


# ===== 3. GENERAR MENSAJE PERSONALIZADO =====
def generar_mensaje_vendedor(nombre, ventas):
    """Genera mensaje personalizado con datos del vendedor"""

    if not ventas:
        return None

    cumplimiento_emoji = "🟢" if ventas['cumplimiento'] >= 75 else "🟡" if ventas['cumplimiento'] >= 50 else "🔴"

    mensaje = f"""📊 REPORTE PERSONAL - AGOSTO
{datetime.now().strftime("%d/%m/%Y %H:%M")}

👤 Vendedor: {nombre}

TU DESEMPEÑO:
Ventas Actuales: S/. {ventas['total']:,.2f}
Cuota: S/. {ventas['cuota']:,.2f}
Cumplimiento: {ventas['cumplimiento']:.1f}% {cumplimiento_emoji}

ESTADÍSTICAS:
Clientes Visitados: {ventas['clientes']}
Documentos: {ventas['documentos']}
Ticket Promedio: S/. {ventas['ticket']:,.2f}

PROYECCIÓN AL 31/AGOSTO:
Venta Proyectada: S/. {ventas['proyeccion']:,.2f}
Cumplimiento Proyectado: {((ventas['proyeccion']/ventas['cuota']*100) if ventas['cuota'] > 0 else 0):.1f}%

PENDIENTE: {ventas['dias_restantes']} días hábiles

¡Sigue adelante! 💪

Sistema Automatizado N&J"""

    return mensaje


# ===== 4. ENVIAR MENSAJE POR WHATSAPP =====
def enviar_mensaje_whatsapp(numero_destino, mensaje):
    """Envía mensaje por WhatsApp API"""
    try:
        logger.info(f'📤 Enviando mensaje a {numero_destino}...')

        # Asegurar formato correcto del número
        numero = numero_destino.replace('+', '').replace(' ', '')
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
            logger.info(f'✅ Mensaje enviado correctamente a {numero}')
            return True
        else:
            logger.error(f'❌ Error enviando mensaje ({response.status_code}): {response.text}')
            return False

    except Exception:
        logger.exception('❌ Error enviando mensaje')
        return False


# ===== WEBHOOK ENDPOINTS =====

@app.route('/webhook', methods=['GET'])
def verificar_webhook():
    """Verifica que Meta pueda conectarse al webhook"""
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if VERIFY_TOKEN and verify_token == VERIFY_TOKEN:
        logger.info('✅ Webhook verificado por Meta')
        return challenge
    else:
        logger.warning('⚠️  Token de verificación incorrecto o no configurado')
        return "Unauthorized", 403


@app.route('/webhook', methods=['POST'])
def recibir_mensaje():
    """Recibe mensajes de WhatsApp"""
    try:
        data = request.get_json(force=True, silent=True)
        logger.info('📨 Mensaje recibido')

        if not data:
            logger.warning('Payload vacío o inválido')
            return jsonify({"status": "ok"}), 200

        # Extraer información del mensaje
        if 'entry' not in data or not data['entry']:
            return jsonify({"status": "ok"}), 200

        entrada = data['entry'][0]
        if 'changes' not in entrada:
            return jsonify({"status": "ok"}), 200

        cambio = entrada['changes'][0]['value']

        # Verificar si hay mensajes
        if 'messages' not in cambio:
            return jsonify({"status": "ok"}), 200

        mensaje_obj = cambio['messages'][0]
        numero_remitente = mensaje_obj['from']

        # Extraer texto del mensaje
        if mensaje_obj['type'] != 'text':
            logger.info(f'Mensaje tipo {mensaje_obj["type"]} ignorado')
            return jsonify({"status": "ok"}), 200

        texto_mensaje = mensaje_obj['text']['body'].strip()
        logger.info(f'📱 Mensaje de {numero_remitente}: "{texto_mensaje}"')

        # TAREA 1: Filtrar por palabra clave
        if PALABRA_CLAVE.lower() not in texto_mensaje.lower():
            logger.info(f'Palabra clave no encontrada. Ignorando.')
            return jsonify({"status": "ok"}), 200

        logger.info(f'✅ Palabra clave detectada: "{PALABRA_CLAVE}"')

        # TAREA 2: Validar vendedor
        es_valido, datos_vendedor = validar_vendedor(numero_remitente)

        if not es_valido:
            mensaje_respuesta = "❌ Número no autorizado. Por favor contacta a tu supervisor."
            logger.warning(f'Enviando alerta de número no autorizado a {numero_remitente}')
            enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            return jsonify({"status": "unauthorized"}), 403

        # TAREA 3: Extraer datos dinámicos
        ventas = obtener_ventas_vendedor(datos_vendedor['nombre'])

        if not ventas:
            mensaje_respuesta = "⚠️ No se encontraron datos de ventas. Intenta más tarde."
            enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            return jsonify({"status": "no_data"}), 404

        # TAREA 4: Generar y enviar reporte
        mensaje_personalizado = generar_mensaje_vendedor(datos_vendedor['nombre'], ventas)

        if mensaje_personalizado:
            exito = enviar_mensaje_whatsapp(numero_remitente, mensaje_personalizado)

            if exito:
                logger.info(f'✅ Reporte enviado exitosamente a {datos_vendedor["nombre"]}')
                return jsonify({"status": "success"}), 200
            else:
                logger.error(f'❌ Error al enviar reporte')
                return jsonify({"status": "send_error"}), 500
        else:
            logger.error('Error generando mensaje')
            return jsonify({"status": "generate_error"}), 500

    except Exception as e:
        logger.error(f'❌ Error procesando webhook: {e}')
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/status', methods=['GET'])
def status():
    """Endpoint de prueba para verificar que el servidor está activo"""
    return jsonify({
        "status": "active",
        "timestamp": datetime.now().isoformat(),
        "webhook": "/webhook",
        "palabra_clave": PALABRA_CLAVE
    }), 200


# ===== EJECUTAR SERVIDOR =====
if __name__ == '__main__':
    logger.info('='*70)
    logger.info('🚀 SERVIDOR WEBHOOK WHATSAPP INICIANDO')
    logger.info('='*70)
    logger.info(f'Escuchando en: http://localhost:5000/webhook')
    logger.info(f'Palabra clave: "{PALABRA_CLAVE}"')
    logger.info(f'BD: {BD_PATH}')
    logger.info(f'Excel vendedores: {EXCEL_VENDEDORES}')
    logger.info(f'Puerto: {os.environ.get("PORT", 5000)}')
    logger.info(f'Normalization: Últimos 9 dígitos (Perú)')
    logger.info(f'Cache: Thread-safe + mtime check')
    logger.info('='*70 + '\n')

    # En producción, ejecutar con Gunicorn/uvicorn. Aquí solo para desarrollo local.
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
