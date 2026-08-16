#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🔔 WEBHOOK RECEPCIÓN DE MENSAJES WHATSAPP - VERSIÓN MEJORADA
Recibe mensajes de vendedores y jefes, valida, y envía reportes dinámicos personalizados
✅ Reportes complejos para vendedor (KPIs + proyecciones)
✅ Reportes agregados para jefe/supervisor (líneas de negocio + TROYA)
✅ Credenciales en variables de entorno (seguro)
✅ Palabra clave: "resumen"
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

# ===== CONFIGURACIÓN CON VARIABLES DE ENTORNO =====
ACCESS_TOKEN = os.environ.get('WHATSAPP_ACCESS_TOKEN', 'EAAO1HSTvFqoBSND9HEaEJi4lKRKBhdU4YhAeiBSH2bu67zxZCvRPqTONojFdRjp112QBxObzZCE8Q2LaLhGV8aJY3kixWsS4fZAxrepU0lFinc7i3iOFCUTTc1GRPGKN8z7w8rC0lqvMZBsQZAodBSTsOCqZAHjjVlQnaI9pT7H9tDEnGFUJOBj5K3iU6aZBDFKzgZDZD')
PHONE_NUMBER_ID = os.environ.get('WHATSAPP_PHONE_NUMBER_ID', '1202656292939375')
VERIFY_TOKEN = os.environ.get('WHATSAPP_VERIFY_TOKEN', 'tu_token_verificacion_seguro')
API_VERSION = os.environ.get('WHATSAPP_API_VERSION', 'v25.0')
API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"

# Palabra clave para solicitar reporte
PALABRA_CLAVE = "resumen"

# Rutas con fallback
BASE_PATH = os.path.dirname(os.path.abspath(__file__))
BD_PATH = os.environ.get('BD_PATH', os.path.join(BASE_PATH, 'ventas.db'))
EXCEL_VENDEDORES = os.environ.get('EXCEL_VENDEDORES', os.path.join(BASE_PATH, 'vendedores.xlsx'))

# Cache para Excel (thread-safe)
_EXCEL_CACHE = {}
_EXCEL_LOCK = Lock()

# Logging
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler(os.path.join(BASE_PATH, 'logs', 'webhook_mejorado.log'), encoding='utf-8'),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)


# ===== 1. VALIDAR VENDEDOR/JEFE =====
def obtener_vendedores_autorizados():
    """Lee lista de vendedores autorizados desde Excel (con cache)"""
    try:
        # Verificar cache
        with _EXCEL_LOCK:
            if _EXCEL_CACHE and os.path.getmtime(EXCEL_VENDEDORES) == _EXCEL_CACHE.get('mtime'):
                return _EXCEL_CACHE.get('data', {})

        logger.info('📋 Leyendo vendedores autorizados desde Excel...')
        vendedores = {}

        wb = openpyxl.load_workbook(EXCEL_VENDEDORES)
        ws = wb.active

        # Excel: Columna A = Nombre, Columna B = Teléfono, Columna C = Clientes
        for row in ws.iter_rows(min_row=1, values_only=True):
            if row[1] and row[0]:  # Teléfono y Nombre
                telefono = str(row[1]).strip()
                telefono_limpio = ''.join(filter(str.isdigit, telefono))
                telefono_last9 = telefono_limpio[-9:] if len(telefono_limpio) >= 9 else telefono_limpio

                vendedores[telefono_last9] = {
                    'nombre': str(row[0]).strip(),
                    'telefonooriginal': telefono_limpio,
                    'clientes': row[2] if len(row) > 2 else None
                }

        # Guardar en cache
        with _EXCEL_LOCK:
            _EXCEL_CACHE['data'] = vendedores
            _EXCEL_CACHE['mtime'] = os.path.getmtime(EXCEL_VENDEDORES)

        logger.info(f'✅ {len(vendedores)} vendedores autorizados cargados')
        return vendedores

    except Exception as e:
        logger.error(f'❌ Error leyendo Excel: {e}')
        return {}


def validar_vendedor(numero_telefonico):
    """Valida si el número de teléfono está autorizado"""
    numero_limpio = ''.join(filter(str.isdigit, numero_telefonico))
    numero_last9 = numero_limpio[-9:] if len(numero_limpio) >= 9 else numero_limpio

    vendedores = obtener_vendedores_autorizados()

    if numero_last9 in vendedores:
        logger.info(f'✅ Usuario autorizado: {vendedores[numero_last9]["nombre"]}')
        return True, vendedores[numero_last9]
    else:
        logger.warning(f'⚠️  Número no autorizado: {numero_telefonico}')
        return False, None


# ===== 2. EXTRAER DATOS DE VENTAS PARA VENDEDOR =====
def obtener_datos_vendedor(nombre_vendedor):
    """Extrae datos complejos de vendedor: ventas, cuota, KPIs, TROYA"""
    try:
        logger.info(f'🔄 Consultando datos de {nombre_vendedor}...')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 1. Ventas totales y clientes
        cursor.execute('''
            SELECT
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total_ventas,
                COUNT(DISTINCT Cod_Clie) as clientes
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR"
        ''', (nombre_vendedor,))

        venta_data = cursor.fetchone()
        if not venta_data:
            conn.close()
            return None

        datos = {
            'vendedor': nombre_vendedor,
            'total_ventas': venta_data['total_ventas'] or 0,
            'clientes': venta_data['clientes'] or 0
        }

        # 2. Ticket promedio general
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR"
        ''', (nombre_vendedor,))

        ticket = cursor.fetchone()
        datos['ticket_promedio'] = ticket['ticket'] or 0

        # 3. Cuota y Cuota Cobertura
        cursor.execute('''
            SELECT Cuota_Soles, Cuota_Cobertura
            FROM cuotas
            WHERE Vendedor = ? AND AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"
            LIMIT 1
        ''', (nombre_vendedor,))

        cuota_data = cursor.fetchone()
        datos['cuota'] = cuota_data['Cuota_Soles'] if cuota_data else 0
        datos['cuota_cobertura'] = int(cuota_data['Cuota_Cobertura']) if cuota_data and cuota_data['Cuota_Cobertura'] else 0

        # 4. Cumplimiento actual (ventas)
        datos['cumplimiento'] = (datos['total_ventas'] / datos['cuota'] * 100) if datos['cuota'] > 0 else 0

        # 5. TROYA - Ventas en clientes con Calif = 'D'
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as troya
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR" AND Calif = "D"
        ''', (nombre_vendedor,))

        troya_data = cursor.fetchone()
        datos['ventas_troya'] = troya_data['troya'] or 0

        # 6. Clientes TROYA (Calif=D) que compraron vs no compraron
        # Clientes que compraron (tienen al menos 1 transacción con Calif=D)
        cursor.execute('''
            SELECT COUNT(DISTINCT Cod_Clie) as clientes_troya_compraron
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR" AND Calif = "D"
        ''', (nombre_vendedor,))

        troya_comp = cursor.fetchone()
        datos['clientes_troya_compraron'] = troya_comp['clientes_troya_compraron'] or 0

        # Clientes que NO compraron (tienen Calif=D pero 0 ventas)
        # Se calcula: total clientes con Calif=D en base - clientes que sí compraron
        cursor.execute('''
            SELECT COUNT(DISTINCT Cod_Clie) as total_clientes_d
            FROM clientes
            WHERE Vendedor = ? AND Calif = "D"
        ''', (nombre_vendedor,))

        total_d = cursor.fetchone()
        total_d_clientes = total_d['total_clientes_d'] or 0
        datos['clientes_troya_no_compraron'] = max(0, total_d_clientes - datos['clientes_troya_compraron'])

        # 7. Ticket promedio solo de clientes TROYA (Calif=D)
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket_troya
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = "202608" AND Proveedor = "ARCOR" AND Calif = "D"
        ''', (nombre_vendedor,))

        ticket_troya = cursor.fetchone()
        datos['ticket_troya'] = ticket_troya['ticket_troya'] or 0

        # 8. Proyección (Venta + (Venta/Días) * Días_Restantes)
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

        datos['dias_restantes'] = dias_restantes
        datos['proyeccion_ventas'] = round(datos['total_ventas'] + ((datos['total_ventas'] / dias_transcurridos) * dias_restantes), 2)
        datos['cumplimiento_proyectado'] = (datos['proyeccion_ventas'] / datos['cuota'] * 100) if datos['cuota'] > 0 else 0

        # 9. Proyección de cobertura (clientes)
        # Fórmula similar: clientes_actuales + (clientes_actuales/días) * días_restantes
        datos['proyeccion_cobertura'] = int(round(datos['clientes'] + ((datos['clientes'] / dias_transcurridos) * dias_restantes)))
        datos['cumplimiento_cobertura_proyectado'] = (datos['proyeccion_cobertura'] / datos['cuota_cobertura'] * 100) if datos['cuota_cobertura'] > 0 else 0

        conn.close()
        logger.info(f'✅ Datos obtenidos para {nombre_vendedor}')
        return datos

    except Exception as e:
        logger.error(f'❌ Error obteniendo datos vendedor: {e}')
        return None


# ===== 3. EXTRAER DATOS AGREGADOS PARA JEFE =====
def obtener_datos_generales():
    """Extrae datos consolidados ARCOR para jefe/supervisor"""
    try:
        logger.info('🔄 Consultando datos generales ARCOR...')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        datos = {}

        # 1. Ventas totales y cobertura ARCOR
        cursor.execute('''
            SELECT
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total_ventas,
                COUNT(DISTINCT Cod_Clie) as clientes_totales
            FROM VENTAS2026
            WHERE Proveedor = "ARCOR"
        ''')

        venta_data = cursor.fetchone()
        datos['total_ventas'] = venta_data['total_ventas'] or 0
        datos['cobertura'] = venta_data['clientes_totales'] or 0

        # 2. Ticket promedio general
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket
            FROM VENTAS2026
            WHERE Proveedor = "ARCOR"
        ''')

        ticket = cursor.fetchone()
        datos['ticket_promedio'] = ticket['ticket'] or 0

        # 3. Cuota total ARCOR (Ventas y Cobertura)
        cursor.execute('''
            SELECT
                ROUND(SUM(Cuota_Soles), 2) as cuota_ventas,
                ROUND(SUM(Cuota_Cobertura), 2) as cuota_cobertura
            FROM cuotas
            WHERE AÑO = 2026 AND NRO_MES = 8 AND Proveedor = "ARCOR"
        ''')

        cuota_data = cursor.fetchone()
        datos['cuota_ventas'] = cuota_data['cuota_ventas'] if cuota_data else 0
        datos['cuota_cobertura'] = int(cuota_data['cuota_cobertura']) if cuota_data and cuota_data['cuota_cobertura'] else 0
        datos['cumplimiento'] = (datos['total_ventas'] / datos['cuota_ventas'] * 100) if datos['cuota_ventas'] > 0 else 0

        # 4. Ventas TROYA (Calif = 'D')
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as ventas_troya
            FROM VENTAS2026
            WHERE Proveedor = "ARCOR" AND Calif = "D"
        ''')

        troya_venta = cursor.fetchone()
        datos['ventas_troya'] = troya_venta['ventas_troya'] or 0

        # 5. Líneas de negocio (agrupado por lin_neg)
        cursor.execute('''
            SELECT
                lin_neg,
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as ventas_linea
            FROM VENTAS2026
            WHERE Proveedor = "ARCOR"
            GROUP BY lin_neg
            ORDER BY ventas_linea DESC
        ''')

        lineas = cursor.fetchall()
        datos['lineas_negocio'] = {row['lin_neg']: row['ventas_linea'] for row in lineas}

        # 6. Clientes TROYA (Calif=D) que compraron vs no compraron
        cursor.execute('''
            SELECT COUNT(DISTINCT Cod_Clie) as clientes_troya_compraron
            FROM VENTAS2026
            WHERE Proveedor = "ARCOR" AND Calif = "D"
        ''')

        troya_comp = cursor.fetchone()
        datos['clientes_troya_compraron'] = troya_comp['clientes_troya_compraron'] or 0

        # Clientes con Calif=D que NO compraron
        cursor.execute('''
            SELECT COUNT(DISTINCT Cod_Clie) as total_clientes_d
            FROM clientes
            WHERE Calif = "D"
        ''')

        total_d = cursor.fetchone()
        total_d_clientes = total_d['total_clientes_d'] or 0
        datos['clientes_troya_no_compraron'] = max(0, total_d_clientes - datos['clientes_troya_compraron'])

        # 7. Ticket promedio solo de clientes TROYA (Calif=D)
        cursor.execute('''
            SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket_troya
            FROM VENTAS2026
            WHERE Proveedor = "ARCOR" AND Calif = "D"
        ''')

        ticket_troya = cursor.fetchone()
        datos['ticket_troya'] = ticket_troya['ticket_troya'] or 0

        # 8. Proyección general (Ventas y Cobertura)
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

        datos['dias_restantes'] = dias_restantes
        datos['proyeccion_ventas'] = round(datos['total_ventas'] + ((datos['total_ventas'] / dias_transcurridos) * dias_restantes), 2)
        datos['cumplimiento_ventas_proyectado'] = (datos['proyeccion_ventas'] / datos['cuota_ventas'] * 100) if datos['cuota_ventas'] > 0 else 0

        datos['proyeccion_cobertura'] = int(round(datos['cobertura'] + ((datos['cobertura'] / dias_transcurridos) * dias_restantes)))
        datos['cumplimiento_cobertura_proyectado'] = (datos['proyeccion_cobertura'] / datos['cuota_cobertura'] * 100) if datos['cuota_cobertura'] > 0 else 0

        conn.close()
        logger.info('✅ Datos generales obtenidos')
        return datos

    except Exception as e:
        logger.error(f'❌ Error obteniendo datos generales: {e}')
        return None


# ===== 4. DETECTAR SI ES JEFE O VENDEDOR =====
def es_jefe(nombre_usuario):
    """Detecta si el usuario es jefe/supervisor"""
    jefes_keywords = ['jefe', 'supervisor', 'gerente', 'coordinador', 'arcor']
    nombre_lower = nombre_usuario.lower()
    return any(keyword in nombre_lower for keyword in jefes_keywords)


# ===== 5. GENERAR MENSAJES PERSONALIZADOS =====
def generar_mensaje_vendedor(datos):
    """Genera reporte detallado para vendedor con estructura nueva"""
    if not datos:
        return None

    cumpl = datos['cumplimiento']
    cumpl_cobertura_proy = datos['cumplimiento_cobertura_proyectado']

    cumplimiento_emoji = "🟢" if cumpl >= 90 else "🟡" if cumpl >= 75 else "🔴"
    proyectado_emoji = "🟢" if datos['cumplimiento_proyectado'] >= 90 else "🟡" if datos['cumplimiento_proyectado'] >= 75 else "🔴"
    cobertura_emoji = "🟢" if cumpl_cobertura_proy >= 90 else "🟡" if cumpl_cobertura_proy >= 75 else "🔴"

    mensaje = f"""📊 REPORTE PERSONAL - AGOSTO
{datetime.now().strftime('%d/%m/%Y %H:%M')}

👤 {datos['vendedor'].upper()}

🎯 OBJETIVOS MES:
├─ Cuota Ventas: S/. {datos['cuota']:,.2f}
└─ Cuota Cobertura: {datos['cuota_cobertura']} clientes

💼 DESEMPEÑO:
├─ Ventas Actuales: S/. {datos['total_ventas']:,.2f}
├─ Cobertura: {datos['clientes']} clientes
├─ Ticket Promedio: S/. {datos['ticket_promedio']:,.2f}
└─ Ventas TROYA (Calif=D): S/. {datos['ventas_troya']:,.2f}

⚠️ TROYA (Clientes Críticos - Calif=D):
├─ Compraron: {datos['clientes_troya_compraron']} clientes
├─ No Compraron: {datos['clientes_troya_no_compraron']} clientes
└─ Ticket Promedio TROYA: S/. {datos['ticket_troya']:,.2f}

🚀 PROYECCIÓN AL 31/AGOSTO:
├─ Ventas Proyectadas: S/. {datos['proyeccion_ventas']:,.2f} ({datos['cumplimiento_proyectado']:.1f}%) {proyectado_emoji}
└─ Cobertura Proyectada: {datos['proyeccion_cobertura']} clientes ({cumpl_cobertura_proy:.1f}%) {cobertura_emoji}

💪 ¡Sigue adelante!
Sistema N&J"""

    return mensaje


def generar_mensaje_jefe(datos):
    """Genera reporte consolidado para jefe/supervisor con nueva estructura"""
    if not datos:
        return None

    cumpl = datos['cumplimiento']
    cumpl_cobertura_proy = datos['cumplimiento_cobertura_proyectado']

    cumplimiento_emoji = "🟢" if cumpl >= 90 else "🟡" if cumpl >= 75 else "🔴"
    proyectado_ventas_emoji = "🟢" if datos['cumplimiento_ventas_proyectado'] >= 90 else "🟡" if datos['cumplimiento_ventas_proyectado'] >= 75 else "🔴"
    cobertura_emoji = "🟢" if cumpl_cobertura_proy >= 90 else "🟡" if cumpl_cobertura_proy >= 75 else "🔴"

    lineas_txt = ""
    for linea, ventas in sorted(datos['lineas_negocio'].items(), key=lambda x: x[1], reverse=True):
        pct = (ventas / datos['total_ventas'] * 100) if datos['total_ventas'] > 0 else 0
        lineas_txt += f"  • {linea}: S/. {ventas:,.0f} ({pct:.1f}%)\n"

    mensaje = f"""📊 REPORTE ARCOR - AGOSTO
{datetime.now().strftime('%d/%m/%Y %H:%M')}

🎯 OBJETIVOS MES:
├─ Cuota Ventas: S/. {datos['cuota_ventas']:,.2f}
└─ Cuota Cobertura: {datos['cuota_cobertura']} clientes

💼 DESEMPEÑO:
├─ Ventas Actuales: S/. {datos['total_ventas']:,.2f}
├─ Cobertura: {datos['cobertura']} clientes
├─ Ticket Promedio: S/. {datos['ticket_promedio']:,.2f}
└─ Ventas TROYA (Calif=D): S/. {datos['ventas_troya']:,.2f}

📋 LÍNEAS DE NEGOCIO:
{lineas_txt}
⚠️ TROYA (Clientes Críticos - Calif=D):
├─ Compraron: {datos['clientes_troya_compraron']} clientes
├─ No Compraron: {datos['clientes_troya_no_compraron']} clientes
└─ Ticket Promedio TROYA: S/. {datos['ticket_troya']:,.2f}

🚀 PROYECCIÓN AL 31/AGOSTO:
├─ Ventas Proyectadas: S/. {datos['proyeccion_ventas']:,.2f} ({datos['cumplimiento_ventas_proyectado']:.1f}%) {proyectado_ventas_emoji}
└─ Cobertura Proyectada: {datos['proyeccion_cobertura']} clientes ({cumpl_cobertura_proy:.1f}%) {cobertura_emoji}

📞 Equipo de Ventas N&J"""

    return mensaje


# ===== 6. ENVIAR MENSAJE POR WHATSAPP =====
def enviar_mensaje_whatsapp(numero_destino, mensaje):
    """Envía mensaje por WhatsApp API"""
    try:
        logger.info(f'📤 Enviando mensaje a {numero_destino}...')

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

    except Exception as e:
        logger.error(f'❌ Error: {e}')
        return False


# ===== WEBHOOK ENDPOINTS =====

@app.route('/webhook', methods=['GET'])
def verificar_webhook():
    """Verifica que Meta pueda conectarse al webhook"""
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if verify_token == VERIFY_TOKEN:
        logger.info('✅ Webhook verificado por Meta')
        return challenge
    else:
        logger.warning('⚠️  Token de verificación incorrecto')
        return "Unauthorized", 403


@app.route('/webhook', methods=['POST'])
def recibir_mensaje():
    """Recibe mensajes de WhatsApp"""
    try:
        data = request.get_json()
        logger.info(f'📨 Mensaje recibido: {json.dumps(data, indent=2)}')

        if 'entry' not in data or not data['entry']:
            return jsonify({"status": "ok"}), 200

        entrada = data['entry'][0]
        if 'changes' not in entrada:
            return jsonify({"status": "ok"}), 200

        cambio = entrada['changes'][0]['value']

        if 'messages' not in cambio:
            return jsonify({"status": "ok"}), 200

        mensaje_obj = cambio['messages'][0]
        numero_remitente = mensaje_obj['from']

        if mensaje_obj['type'] != 'text':
            logger.info(f'Mensaje tipo {mensaje_obj["type"]} ignorado')
            return jsonify({"status": "ok"}), 200

        texto_mensaje = mensaje_obj['text']['body'].strip()
        logger.info(f'📱 Mensaje de {numero_remitente}: "{texto_mensaje}"')

        # VALIDAR PALABRA CLAVE
        if PALABRA_CLAVE.lower() not in texto_mensaje.lower():
            logger.info(f'Palabra clave "{PALABRA_CLAVE}" no encontrada. Ignorando.')
            return jsonify({"status": "ok"}), 200

        logger.info(f'✅ Palabra clave detectada: "{PALABRA_CLAVE}"')

        # VALIDAR USUARIO
        es_valido, datos_usuario = validar_vendedor(numero_remitente)

        if not es_valido:
            mensaje_respuesta = "❌ Número no autorizado. Por favor contacta a tu supervisor."
            logger.warning(f'Enviando alerta de número no autorizado a {numero_remitente}')
            enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            return jsonify({"status": "unauthorized"}), 403

        # DETERMINAR TIPO DE USUARIO Y GENERAR REPORTE
        nombre_usuario = datos_usuario['nombre']
        es_jefe_supervisor = es_jefe(nombre_usuario)

        if es_jefe_supervisor:
            logger.info(f'👔 Detectado: JEFE/SUPERVISOR ({nombre_usuario})')
            datos = obtener_datos_generales()
            mensaje = generar_mensaje_jefe(datos)
        else:
            logger.info(f'👤 Detectado: VENDEDOR ({nombre_usuario})')
            datos = obtener_datos_vendedor(nombre_usuario)
            mensaje = generar_mensaje_vendedor(datos)

        if not datos or not mensaje:
            mensaje_respuesta = "⚠️ No se encontraron datos. Intenta más tarde."
            enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            return jsonify({"status": "no_data"}), 404

        # ENVIAR REPORTE
        exito = enviar_mensaje_whatsapp(numero_remitente, mensaje)

        if exito:
            logger.info(f'✅ Reporte enviado a {nombre_usuario}')
            return jsonify({"status": "success"}), 200
        else:
            logger.error(f'❌ Error al enviar reporte')
            return jsonify({"status": "send_error"}), 500

    except Exception as e:
        logger.error(f'❌ Error procesando webhook: {e}')
        return jsonify({"status": "error", "message": str(e)}), 500


@app.route('/status', methods=['GET'])
def status():
    """Endpoint de prueba"""
    return jsonify({
        "status": "active",
        "timestamp": datetime.now().isoformat(),
        "webhook": "/webhook",
        "palabra_clave": PALABRA_CLAVE,
        "versión": "MEJORADA"
    }), 200


# ===== EJECUTAR SERVIDOR =====
if __name__ == '__main__':
    logger.info('='*70)
    logger.info('🚀 SERVIDOR WEBHOOK WHATSAPP - VERSIÓN MEJORADA')
    logger.info('='*70)
    logger.info(f'Escuchando en: http://localhost:5000/webhook')
    logger.info(f'Palabra clave: "{PALABRA_CLAVE}"')
    logger.info(f'BD: {BD_PATH}')
    logger.info(f'Excel vendedores: {EXCEL_VENDEDORES}')
    logger.info('='*70 + '\n')

    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
