#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
🔔 WEBHOOK RECEPCIÓN DE MENSAJES WHATSAPP (refactorizado)
- Configuración vía variables de entorno
- Cache simple para cargar vendedores desde Excel (misma estructura que BOT_RESPONDER_VENDEDORES.py)
- Normalización por últimos 9 dígitos (Perú)
- Verificación HMAC (X-Hub-Signature-256)
- Consultas SQL parametrizadas
- Manejo robusto del JSON entrante
- No ejecutar en debug en producción (usar gunicorn)
"""

from flask import Flask, request, jsonify
import sqlite3
import requests
import logging
import os
import openpyxl
from datetime import datetime
from threading import Lock
import hmac
import hashlib

# ===== CONFIG / CREDENCIALES (leer desde variables de entorno) =====
ACCESS_TOKEN = os.environ.get('WHATSAPP_ACCESS_TOKEN')
PHONE_NUMBER_ID = os.environ.get('WHATSAPP_PHONE_NUMBER_ID')
VERIFY_TOKEN = os.environ.get('WHATSAPP_VERIFY_TOKEN')
API_VERSION = os.environ.get('WHATSAPP_API_VERSION', 'v18.0')
APP_SECRET = os.environ.get('WHATSAPP_APP_SECRET')  # para verificar X-Hub-Signature-256

# Rutas / recursos (configurables)
BD_PATH = os.environ.get('BD_PATH', 'ventas.db')
EXCEL_VENDEDORES = os.environ.get('EXCEL_VENDEDORES', 'vendedores.xlsx')

# Column configuration for vendedores.xlsx (0-based indices)
# Mantener misma estructura que BOT_RESPONDER_VENDEDORES.py: nombre=col A (0), telefono=col B (1), start row = 1
VEN_COL_NOMBRE = int(os.environ.get('VEN_COL_NOMBRE', '0'))
VEN_COL_TELEFONO = int(os.environ.get('VEN_COL_TELEFONO', '1'))
VEN_START_ROW = int(os.environ.get('VEN_START_ROW', '1'))

# Logging level configurable via env (DEBUG/INFO/WARNING/ERROR)
LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO').upper()
try:
    log_level = getattr(logging, LOG_LEVEL)
except Exception:
    log_level = logging.INFO

if PHONE_NUMBER_ID:
    API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"
else:
    API_URL = None

# ===== LOGGING =====
os.makedirs('logs', exist_ok=True)
logging.basicConfig(
    level=log_level,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('logs/webhook_vendedores.log', encoding='utf-8'),
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
    Usa VEN_COL_NOMBRE y VEN_COL_TELEFONO (0-based) y VEN_START_ROW (1-based).
    """
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

            # openpyxl iter_rows min_row expects 1-based indexing
            for row in ws.iter_rows(min_row=VEN_START_ROW, values_only=True):
                if not row:
                    continue
                nombre = row[VEN_COL_NOMBRE] if len(row) > VEN_COL_NOMBRE else None
                tel = row[VEN_COL_TELEFONO] if len(row) > VEN_COL_TELEFONO else None
                if nombre and tel:
                    tel_clean = ''.join(filter(str.isdigit, str(tel)))
                    tel_last9 = tel_clean[-9:] if len(tel_clean) >= 9 else tel_clean
                    vendedores.append({'nombre': str(nombre).strip(), 'telefono': tel_last9})

            _EXCEL_CACHE['mtime'] = mtime
            _EXCEL_CACHE['data'] = vendedores
            logger.debug(f'Vendedores cargados: {len(vendedores)}')
            return vendedores

    except Exception:
        logger.exception('Error cargando Excel')
        return []


# ===== Helpers =====

def verify_signature(request):
    """Verifica X-Hub-Signature-256 usando APP_SECRET. Devuelve True si coincide.
    Si APP_SECRET no está configurado, devuelve True (opcional, pero se recomienda configurar en producción).
    """
    if not APP_SECRET:
        logger.warning('APP_SECRET no configurado: saltando verificación de firma')
        return True

    signature = request.headers.get('X-Hub-Signature-256')
    if not signature:
        logger.warning('No X-Hub-Signature-256 en headers')
        return False

    try:
        sha_name, sig = signature.split('=', 1)
    except Exception:
        logger.warning('Formato de X-Hub-Signature-256 inválido')
        return False

    if sha_name != 'sha256':
        logger.warning('Algoritmo de firma no soportado: %s', sha_name)
        return False

    expected = hmac.new(APP_SECRET.encode('utf-8'), request.get_data(), hashlib.sha256).hexdigest()
    return hmac.compare_digest(sig, expected)


def obtener_vendedor_de_excel(numero_telefono):
    """Busca vendedor en Excel por teléfono: compara por últimos 9 dígitos."""
    try:
        numero_limpio = ''.join(filter(str.isdigit, str(numero_telefono)))
        numero_last9 = numero_limpio[-9:] if len(numero_limpio) >= 9 else numero_limpio

        vendedores = cargar_vendedores()
        logger.debug(f'Buscando vendedor para número entrante: {numero_telefono} -> last9={numero_last9}')

        for v in vendedores:
            tel = v.get('telefono', '')
            if not tel:
                continue
            if tel == numero_last9 or tel.endswith(numero_last9):
                logger.info(f"Vendedor encontrado: {v['nombre']} para número {numero_telefono}")
                return v.get('nombre')

        logger.warning(f'No se encontró vendedor para número {numero_telefono}')
        return None

    except Exception:
        logger.exception('Error obteniendo vendedor desde Excel')
        return None


# ===== DB / Consultas =====

def obtener_datos_vendedor(nombre_vendedor):
    """Obtiene datos de ventas del vendedor (parámetros fijos para Periodo/Proveedor como en el proyecto)."""
    try:
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        # 1. Ventas totales (parametrizada)
        periodo = os.environ.get('PERIODO', '202608')
        proveedor = os.environ.get('PROVEEDOR', 'ARCOR')

        cursor.execute('''
            SELECT
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total_ventas,
                COUNT(DISTINCT Cod_Clie) as clientes,
                COUNT(DISTINCT Documento) as documentos
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = ?
        ''', (nombre_vendedor, periodo, proveedor))

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
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = ?
        ''', (nombre_vendedor, periodo, proveedor))

        ticket = cursor.fetchone()
        ventas['ticket'] = ticket['ticket'] or 0

        # 3. Cuota
        anio = int(os.environ.get('ANIO', '2026'))
        nro_mes = int(os.environ.get('NRO_MES', '8'))
        cursor.execute('''
            SELECT Cuota_Soles
            FROM cuotas
            WHERE Vendedor = ? AND AÑO = ? AND NRO_MES = ? AND Proveedor = ?
            LIMIT 1
        ''', (nombre_vendedor, anio, nro_mes, proveedor))

        cuota_data = cursor.fetchone()
        ventas['cuota'] = cuota_data['Cuota_Soles'] if cuota_data and 'Cuota_Soles' in cuota_data.keys() else 0

        # 4. Cumplimiento
        ventas['cumplimiento'] = (ventas['total'] / ventas['cuota'] * 100) if ventas['cuota'] > 0 else 0

        # 5. Días transcurridos y proyección: mantiene comportamiento original usando SQL CTE,
        # pero con periodo parametrizado. Si prefieres, se puede calcular en Python.
        cursor.execute('''
            WITH RECURSIVE dates AS (
              SELECT DATE(? || '-01') as fecha
              UNION ALL
              SELECT DATE(fecha, '+1 day') FROM dates
              WHERE fecha <= DATE('now')
            )
            SELECT COUNT(*) as dias FROM dates
            WHERE CAST(strftime('%w', fecha) AS INTEGER) IN (1,2,3,4,5)
        ''', (f"{periodo[:4]}-{periodo[4:6]}",))

        dias_data = cursor.fetchone()
        dias_transcurridos = dias_data['dias'] or 1
        total_dias_habiles = int(os.environ.get('TOTAL_DIAS_HABILES', '26'))
        dias_restantes = total_dias_habiles - dias_transcurridos

        ventas['dias_restantes'] = dias_restantes
        ventas['proyeccion'] = round(ventas['total'] + ((ventas['total'] / dias_transcurridos) * dias_restantes), 2) if dias_transcurridos > 0 else ventas['total']

        conn.close()
        logger.info(f'✅ Datos obtenidos para {nombre_vendedor}')
        return ventas

    except Exception:
        logger.exception('Error obteniendo ventas')
        return None


def obtener_datos_generales():
    """Obtiene datos consolidados (KPIs) para Jefes/Supervisores."""
    try:
        logger.info('🔄 Obteniendo KPIs completos...')
        conn = sqlite3.connect(BD_PATH)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        periodo = os.environ.get('PERIODO', '202608')
        proveedor = os.environ.get('PROVEEDOR', 'ARCOR')
        anio = int(os.environ.get('ANIO', '2026'))
        nro_mes = int(os.environ.get('NRO_MES', '8'))
        total_dias_habiles = int(os.environ.get('TOTAL_DIAS_HABILES', '26'))

        cursor.execute('SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as total FROM VENTAS2026 WHERE Periodo = ? AND Proveedor = ?', (periodo, proveedor))
        kpi_ventas = cursor.fetchone()['total'] or 0

        cursor.execute('SELECT COUNT(DISTINCT Cod_Clie) as total FROM VENTAS2026 WHERE Periodo = ? AND Proveedor = ?', (periodo, proveedor))
        cobertura_actual = cursor.fetchone()['total'] or 0

        cursor.execute('SELECT ROUND(SUM(CAST(Imp_Total AS REAL)) / COUNT(DISTINCT Documento), 2) as ticket FROM VENTAS2026 WHERE Periodo = ? AND Proveedor = ?', (periodo, proveedor))
        ticket_promedio = cursor.fetchone()['ticket'] or 0

        cursor.execute('SELECT ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as troya_venta FROM VENTAS2026 WHERE Periodo = ? AND Calif = ? AND Proveedor = ?', (periodo, 'D', proveedor))
        troya_venta = cursor.fetchone()['troya_venta'] or 0
        troya_pct = (troya_venta / kpi_ventas * 100) if kpi_ventas > 0 else 0

        cursor.execute('''
            SELECT Lin_Neg, ROUND(SUM(CAST(Imp_Total AS REAL)), 2) as venta_linea
            FROM VENTAS2026
            WHERE Periodo = ? AND Lin_Neg != ? AND Proveedor = ?
            GROUP BY Lin_Neg
            ORDER BY venta_linea DESC
        ''', (periodo, 'MATERIAL POP', proveedor))

        categorias = []
        total_lineas = 0
        for row in cursor.fetchall():
            categorias.append(dict(row))
            total_lineas += row['venta_linea']

        for cat in categorias:
            cat['pct_participacion'] = (cat['venta_linea'] / total_lineas * 100) if total_lineas > 0 else 0

        # Días transcurridos (CTE similar al anterior)
        cursor.execute('''
            WITH RECURSIVE dates AS (
                SELECT DATE(? || '-01') as fecha
                UNION ALL
                SELECT DATE(fecha, '+1 day') FROM dates
                WHERE fecha <= DATE('now')
            )
            SELECT COUNT(*) as dias FROM dates
            WHERE CAST(strftime('%w', fecha) AS INTEGER) IN (1,2,3,4,5)
        ''', (f"{periodo[:4]}-{periodo[4:6]}",))

        dias_transcurridos = cursor.fetchone()['dias'] or 1
        dias_restantes = total_dias_habiles - dias_transcurridos

        proyeccion_ventas = round(kpi_ventas + ((kpi_ventas / dias_transcurridos) * dias_restantes), 2) if dias_transcurridos > 0 else kpi_ventas

        cursor.execute('SELECT ROUND(SUM(Cuota_Soles), 2) as cuota_total FROM cuotas WHERE AÑO = ? AND NRO_MES = ? AND Proveedor = ?', (anio, nro_mes, proveedor))
        cuota_total = cursor.fetchone()['cuota_total'] or 1
        pct_cumplimiento = (proyeccion_ventas / cuota_total * 100) if cuota_total > 0 else 0

        conn.close()

        return {
            'kpi_ventas': kpi_ventas,
            'cobertura_actual': cobertura_actual,
            'cobertura_proyectada': round((cobertura_actual / dias_transcurridos) * total_dias_habiles) if dias_transcurridos > 0 else cobertura_actual,
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


# ===== Mensajes =====

def generar_mensaje_personal(nombre, ventas):
    cumplimiento_emoji = "🟢" if ventas.get('cumplimiento', 0) >= 75 else "🟡" if ventas.get('cumplimiento', 0) >= 50 else "🔴"

    return f"""📊 REPORTE PERSONAL - AGOSTO
{datetime.now().strftime('%d/%m/%Y %H:%M')}

👤 Vendedor: {nombre}

TU DESEMPEÑO:
Ventas Actuales: S/. {ventas.get('total',0):,.2f}
Cuota: S/. {ventas.get('cuota',0):,.2f}
Cumplimiento: {ventas.get('cumplimiento',0):.1f}% {cumplimiento_emoji}

ESTADÍSTICAS:
Clientes Visitados: {ventas.get('clientes',0)}
Documentos: {ventas.get('documentos',0)}
Ticket Promedio: S/. {ventas.get('ticket',0):,.2f}

PROYECCIÓN AL 31/AGOSTO:
Venta Proyectada: S/. {ventas.get('proyeccion',0):,.2f}
Cumplimiento Proyectado: {(ventas.get('proyeccion',0)/ventas.get('cuota',1)*100):.1f}%

PENDIENTE: {ventas.get('dias_restantes',0)} días hábiles

¡Sigue adelante! 💪

Sistema Automatizado N&J"""


def generar_mensaje_general(kpis):
    categorias_emojis = {
        'GOLOSINAS': '🍬',
        'CHOCOLATES': '🍫',
        'CHICLES': '💫',
        'GALLETAS': '🍪',
        'ALIMENTOS': '🥫'
    }

    categorias_texto = ""
    for cat in kpis.get('categorias', []):
        emoji = categorias_emojis.get(cat.get('Lin_Neg', ''), '')
        if emoji:
            categorias_texto += f"\n{emoji} {cat['Lin_Neg']} - S/. {cat['venta_linea']:,.2f} ({cat['pct_participacion']:.1f}%)"
        else:
            categorias_texto += f"\n{cat['Lin_Neg']} - S/. {cat['venta_linea']:,.2f} ({cat['pct_participacion']:.1f}%)"

    return f"""📊 REPORTE ARCOR - AGOSTO
{datetime.now().strftime('%d/%m/%Y %H:%M')}

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


# ===== Envío WhatsApp =====

def enviar_mensaje_whatsapp(numero_destino, mensaje):
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
    """Verifica que Meta pueda conectarse al webhook (challenge)"""
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if VERIFY_TOKEN and verify_token == VERIFY_TOKEN:
        logger.info('✅ Webhook verificado por Meta')
        return challenge
    else:
        logger.warning('⚠️  Token de verificación incorrecto')
        return "Unauthorized", 403


@app.route('/webhook', methods=['POST'])
def recibir_mensaje():
    """Recibe mensajes de WhatsApp"""
    try:
        # Verificar firma HMAC antes de procesar
        if not verify_signature(request):
            logger.warning('Firma HMAC inválida o ausente')
            return jsonify({"status": "forbidden"}), 403

        data = request.get_json(force=True, silent=True)
        logger.info('📨 Mensaje recibido')
        logger.debug(f'Payload: {data}')

        if not data:
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
        logger.info(f'📱 Mensaje de {numero_remitente}: "{texto_mensaje}"')

        # Filtrar por palabra clave (resumen)
        if 'resumen' not in texto_mensaje.lower():
            logger.info('Palabra clave no encontrada. Ignorando.')
            return jsonify({"status": "ok"}), 200

        logger.info('✅ Palabra clave detectada')

        # Validar vendedor usando Excel cache
        nombre = obtener_vendedor_de_excel(numero_remitente)
        if not nombre:
            mensaje_respuesta = "❌ Número no autorizado. Por favor contacta a tu supervisor."
            logger.warning(f'Número no autorizado: {numero_remitente}')
            enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            return jsonify({"status": "unauthorized"}), 403

        logger.info(f'✅ Vendedor autorizado: {nombre}')

        # Determinar si es jefe/supervisor (nombres especiales)
        es_jefe = str(nombre).upper() in ['JEFE DE VENTAS', 'SUPERVISOR ARCOR']

        if es_jefe:
            ventas = obtener_datos_generales()
            if not ventas:
                enviar_mensaje_whatsapp(numero_remitente, "❌ Error obteniendo datos generales")
                return jsonify({"status": "no_data"}), 404

            mensaje = generar_mensaje_general(ventas)
        else:
            ventas = obtener_datos_vendedor(nombre)
            if not ventas:
                enviar_mensaje_whatsapp(numero_remitente, "⚠️ No se encontraron datos de ventas. Intenta más tarde.")
                return jsonify({"status": "no_data"}), 404

            mensaje = generar_mensaje_personal(nombre, ventas)

        if mensaje:
            exito = enviar_mensaje_whatsapp(numero_remitente, mensaje)
            if exito:
                logger.info(f'✅ Reporte enviado exitosamente a {nombre}')
                return jsonify({"status": "success"}), 200
            else:
                logger.error('❌ Error al enviar reporte')
                return jsonify({"status": "send_error"}), 500
        else:
            logger.error('Error generando mensaje')
            return jsonify({"status": "generate_error"}), 500

    except Exception:
        logger.exception('❌ Error procesando webhook')
        return jsonify({"status": "error"}), 500


@app.route('/status', methods=['GET'])
def status():
    """Endpoint de prueba para verificar que el servidor está activo"""
    return jsonify({
        "status": "active",
        "timestamp": datetime.now().isoformat(),
        "webhook": "/webhook",
        "palabra_clave": "resumen"
    }), 200


if __name__ == '__main__':
    logger.info('='*70)
    logger.info('🚀 WEBHOOK VENDEDORES (refactor) INICIANDO')
    logger.info('='*70)
    logger.info(f'Escuchando en: http://localhost:5000/webhook')
    logger.info('Palabra clave: "resumen"')
    logger.info(f'BD: {BD_PATH}')
    logger.info(f'Excel vendedores: {EXCEL_VENDEDORES}')
    logger.info('='*70 + '\n')

    # En producción, ejecutar con Gunicorn/uvicorn. Aquí solo para desarrollo local.
    app.run(host='0.0.0.0', port=int(os.environ.get('PORT', 5000)), debug=False)
