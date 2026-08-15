#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🤖 BOT RESPONDER A VENDEDORES - VERSIÓN MEJORADA CON DEBUGGING
Recibe "resumen" → Valida → Obtiene datos → Responde
"""

from flask import Flask, request, jsonify
import sqlite3
import requests
import logging
import os
import openpyxl
from datetime import datetime
import json

# ===== CREDENCIALES =====
ACCESS_TOKEN = "EAAO1HSTvFqoBSND9HEaEJi4lKRKBhdU4YhAeiBSH2bu67zxZCvRPqTONojFdRjp112QBxObzZCE8Q2LaLhGV8aJY3kixWsS4fZAxrepU0lFinc7i3iOFCUTTc1GRPGKN8z7w8rC0lqvMZBsQZAodBSTsOCqZAHjjVlQnaI9pT7H9tDEnGFUJOBj5K3iU6aZBDFKzgZDZD"
PHONE_NUMBER_ID = "1202656292939375"
VERIFY_TOKEN = "tu_token_verificacion_seguro"
API_VERSION = "v18.0"
API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"

# ===== RUTAS =====
BD_PATH = 'ventas.db'
EXCEL_VENDEDORES = 'vendedores.xlsx'

# ===== LOGGING DETALLADO =====
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
    level=logging.DEBUG,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/bot_debug.log', encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)


def obtener_vendedor_de_excel(numero_telefono):
    """Busca vendedor en Excel por teléfono"""
    try:
        numero_limpio = ''.join(filter(str.isdigit, numero_telefono))
        if numero_limpio.startswith('51'):
            numero_limpio = numero_limpio[2:]

        logger.info(f'🔍 Buscando vendedor con teléfono: {numero_limpio}')

        wb = openpyxl.load_workbook(EXCEL_VENDEDORES)
        ws = wb.active

        for row in ws.iter_rows(min_row=2, values_only=True):
            if row[2]:
                tel = ''.join(filter(str.isdigit, str(row[2])))
                logger.debug(f'  Comparando: {tel} == {numero_limpio}')
                if tel == numero_limpio or tel.endswith(numero_limpio):
                    logger.info(f'✅ Vendedor encontrado: {row[1]}')
                    return row[1]

        logger.warning(f'⚠️ Vendedor NO encontrado con teléfono: {numero_limpio}')
        return None

    except Exception as e:
        logger.error(f'❌ Error leyendo Excel: {e}', exc_info=True)
        return None


def obtener_datos_vendedor(nombre_vendedor):
    """Obtiene datos de ventas del vendedor"""
    try:
        logger.info(f'📊 Obteniendo datos para: {nombre_vendedor}')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

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
            logger.warning(f'⚠️ Sin datos de ventas para: {nombre_vendedor}')
            conn.close()
            return None

        ventas = {
            'total': venta_data['total_ventas'] or 0,
            'clientes': venta_data['clientes'] or 0,
            'documentos': venta_data['documentos'] or 0
        }

        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR"
        ''', (nombre_vendedor,))

        ticket = cursor.fetchone()
        ventas['ticket'] = ticket['ticket'] or 0

        cursor.execute('''
            SELECT Cuota_Soles
            FROM cuotas
            WHERE Vendedor = ? AND AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"
            LIMIT 1
        ''', (nombre_vendedor,))

        cuota_data = cursor.fetchone()
        ventas['cuota'] = cuota_data['Cuota_Soles'] if cuota_data else 0

        ventas['cumplimiento'] = (ventas['total'] / ventas['cuota'] * 100) if ventas['cuota'] > 0 else 0

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

        logger.info(f'✅ Datos obtenidos: S/. {ventas["total"]}')
        conn.close()
        return ventas

    except Exception as e:
        logger.error(f'❌ Error BD: {e}', exc_info=True)
        return None


def generar_respuesta(nombre, ventas):
    """Genera mensaje de respuesta"""
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


def enviar_respuesta(numero_destino, mensaje):
    """Envía mensaje por WhatsApp API"""
    try:
        numero = numero_destino.replace('+', '').replace(' ', '')
        if not numero.startswith('51'):
            numero = f"51{numero}"

        logger.info(f'📤 Enviando a: {numero}')

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

        if response.status_code == 200:
            logger.info(f'✅ Mensaje ENVIADO a {numero}')
            return True
        else:
            logger.error(f'❌ Error WhatsApp {response.status_code}: {response.text}')
            return False

    except Exception as e:
        logger.error(f'❌ Error enviando: {e}', exc_info=True)
        return False


# ===== WEBHOOK =====

@app.route('/webhook', methods=['GET'])
def verificar():
    """Verifica webhook con Meta"""
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    logger.info(f'🔐 Verificación webhook:')
    logger.info(f'  Token recibido: {verify_token}')
    logger.info(f'  Token esperado: {VERIFY_TOKEN}')

    if verify_token == VERIFY_TOKEN:
        logger.info('✅ Webhook VERIFICADO')
        return challenge

    logger.error('❌ Token incorrecto')
    return "Unauthorized", 403


@app.route('/webhook', methods=['POST'])
def recibir():
    """Recibe y responde a mensajes"""
    try:
        data = request.get_json()
        logger.info('='*70)
        logger.info('📨 NUEVO EVENTO WEBHOOK')
        logger.info(f'Payload: {json.dumps(data, indent=2)}')
        logger.info('='*70)

        if 'entry' not in data or not data['entry']:
            logger.warning('⚠️ No hay entry en payload')
            return jsonify({"status": "ok"}), 200

        entrada = data['entry'][0]
        if 'changes' not in entrada:
            logger.warning('⚠️ No hay changes')
            return jsonify({"status": "ok"}), 200

        cambio = entrada['changes'][0]['value']
        logger.info(f'Changes: {cambio}')

        if 'messages' not in cambio:
            logger.info('ℹ️ Sin mensajes (status de entrega)')
            return jsonify({"status": "ok"}), 200

        mensaje_obj = cambio['messages'][0]
        logger.info(f'Mensaje: {mensaje_obj}')

        numero_remitente = mensaje_obj.get('from', 'desconocido')

        # Validar tipo de mensaje
        if 'text' not in mensaje_obj:
            tipo = mensaje_obj.get('type', 'desconocido')
            logger.warning(f'⚠️ NO es texto. Tipo: {tipo}')
            return jsonify({"status": "ok"}), 200

        texto_mensaje = mensaje_obj['text'].get('body', '').strip()

        logger.info(f'📱 Remitente: {numero_remitente}')
        logger.info(f'💬 Texto: "{texto_mensaje}"')

        # Verificar palabra clave
        if "resumen" not in texto_mensaje.lower():
            logger.info(f'❌ Palabra "resumen" NO encontrada')
            logger.info(f'   Mensaje: "{texto_mensaje}"')
            return jsonify({"status": "ok"}), 200

        logger.info('✅ Palabra "resumen" DETECTADA')

        # Validar vendedor
        nombre = obtener_vendedor_de_excel(numero_remitente)

        if not nombre:
            logger.warning(f'❌ Número NO autorizado: {numero_remitente}')
            enviar_respuesta(numero_remitente, "❌ Número no autorizado")
            return jsonify({"status": "unauthorized"}), 403

        logger.info(f'✅ Vendedor válido: {nombre}')

        # Obtener datos
        ventas = obtener_datos_vendedor(nombre)

        if not ventas:
            logger.error(f'❌ Sin datos de ventas para: {nombre}')
            enviar_respuesta(numero_remitente, "⚠️ Sin datos disponibles")
            return jsonify({"status": "no_data"}), 404

        logger.info('✅ Datos obtenidos')

        # Generar y enviar respuesta
        mensaje = generar_respuesta(nombre, ventas)

        if enviar_respuesta(numero_remitente, mensaje):
            logger.info(f'✅✅ REPORTE ENVIADO EXITOSAMENTE a {nombre}')
            return jsonify({"status": "success"}), 200
        else:
            logger.error('❌ Error enviando respuesta')
            return jsonify({"status": "send_error"}), 500

    except Exception as e:
        logger.error(f'❌ ERROR: {e}', exc_info=True)
        return jsonify({"status": "error"}), 500


if __name__ == '__main__':
    logger.info('='*70)
    logger.info('🤖 BOT RESPONDER VENDEDORES INICIANDO')
    logger.info('='*70)
    logger.info(f'URL: http://localhost:5000/webhook')
    logger.info(f'Palabra clave: "resumen"')
    logger.info(f'VERIFY_TOKEN actual: {VERIFY_TOKEN}')
    logger.info('='*70)

    app.run(host='0.0.0.0', port=5000, debug=True)
