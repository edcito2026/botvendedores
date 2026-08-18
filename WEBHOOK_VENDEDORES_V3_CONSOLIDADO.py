#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WEBHOOK VENDEDORES V4 - CORRECTO
Reportes TROYA (Calif='D') con JOIN correcto por Cod_Clie y Cdg_Vend (CAST AS REAL)
Proveedor: ARCOR
"""

from flask import Flask, request, jsonify
import sqlite3
import requests
import logging
import os
from datetime import datetime
from zoneinfo import ZoneInfo
import openpyxl
from threading import Lock

# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

BASE_PATH = os.path.dirname(os.path.abspath(__file__))

ACCESS_TOKEN = os.environ.get("WHATSAPP_ACCESS_TOKEN")
PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")
VERIFY_TOKEN = os.environ.get("WHATSAPP_VERIFY_TOKEN")
API_VERSION = os.environ.get("WHATSAPP_API_VERSION", "v25.0")

if not ACCESS_TOKEN or not PHONE_NUMBER_ID or not VERIFY_TOKEN:
    raise RuntimeError("Faltan variables de entorno WhatsApp")

API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"

BD_PATH = os.environ.get("BD_PATH", os.path.join(BASE_PATH, "ventas.db"))
EXCEL_VENDEDORES = os.environ.get("EXCEL_VENDEDORES", os.path.join(BASE_PATH, "vendedores.xlsx"))

try:
    TZ_LIMA = ZoneInfo("America/Lima")
except:
    TZ_LIMA = None

# ============================================================
# 2. LOGGING
# ============================================================

LOG_DIR = os.path.join(BASE_PATH, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(os.path.join(LOG_DIR, "webhook_v4.log"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

_excel_cache = {}
_excel_cache_lock = Lock()
_excel_mtime = 0

# ============================================================
# 3. UTILIDADES
# ============================================================

def ahora_local():
    """Fecha/hora actual en Perú"""
    if TZ_LIMA:
        return datetime.now(TZ_LIMA)
    return datetime.now()

def normalizar_telefono(numero):
    """Normaliza teléfono a formato 51XXXXXXXXX"""
    if not numero:
        return None
    limpio = ''.join(filter(str.isdigit, numero))
    if limpio.startswith('51'):
        return limpio
    if limpio.startswith('0'):
        return f"51{limpio[1:]}"
    if len(limpio) == 9:
        return f"51{limpio}"
    return limpio

def obtener_ultimo_numero(numero):
    """Últimos 9 dígitos"""
    limpio = ''.join(filter(str.isdigit, numero))
    return limpio[-9:] if len(limpio) >= 9 else limpio

def cargar_vendedores_excel():
    """Carga vendedores del Excel con caché"""
    global _excel_cache, _excel_mtime

    try:
        if not os.path.exists(EXCEL_VENDEDORES):
            logger.error(f"❌ Excel no encontrado: {EXCEL_VENDEDORES}")
            return {}

        mtime = os.path.getmtime(EXCEL_VENDEDORES)

        with _excel_cache_lock:
            if _excel_cache and mtime == _excel_mtime:
                return _excel_cache

            vendedores = {}
            wb = openpyxl.load_workbook(EXCEL_VENDEDORES)
            ws = wb['Vendedores'] if 'Vendedores' in wb.sheetnames else wb.active

            for row in ws.iter_rows(min_row=4, values_only=True):
                if row[0] is None:
                    break

                codigo = row[0]
                nombre = row[1]
                telefono = row[2]

                if codigo and nombre and telefono:
                    tel_norm = normalizar_telefono(str(telefono))
                    tel_9dig = obtener_ultimo_numero(str(telefono))

                    vendedores[tel_norm] = {
                        'codigo': codigo,
                        'nombre': str(nombre).strip(),
                        'telefono': tel_norm
                    }

                    vendedores[tel_9dig] = vendedores[tel_norm]

            _excel_cache = vendedores
            _excel_mtime = mtime
            logger.info(f"✅ {len(set(v['codigo'] for v in vendedores.values()))} vendedores cargados")
            return vendedores

    except Exception as e:
        logger.error(f"❌ Error leyendo Excel: {e}")
        return {}

def obtener_clientes_troya_vendedor(cod_vendedor):
    """
    Obtiene clientes TROYA (Calif='D') del vendedor CON VENTAS ARCOR

    Query correcta:
    - JOIN por Cod_Clie (CAST AS REAL)
    - JOIN por Cdg_Vend (CAST AS REAL)
    - Filtra por Proveedor='ARCOR'
    - Período actual (202608) vs anterior (202607)
    """
    try:
        conn = sqlite3.connect(BD_PATH, timeout=15)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        ahora = ahora_local()
        periodo_actual = f"{ahora.year}{ahora.month:02d}"
        periodo_anterior = f"{ahora.year}{ahora.month-1:02d}" if ahora.month > 1 else f"{ahora.year-1}12"

        # Query con JOINs correctos
        query = """
        SELECT
          c.Cod_Clie,
          c.Raz_Social,
          c.Vendedor,
          c.DV,
          c.Giro,
          COALESCE(ROUND(SUM(CASE WHEN v_actual.Periodo = ? THEN v_actual.Imp_Total ELSE 0 END), 2), 0) as venta_actual,
          COALESCE(ROUND(SUM(CASE WHEN v_anterior.Periodo = ? THEN v_anterior.Imp_Total ELSE 0 END), 2), 0) as venta_anterior
        FROM clientes c
        LEFT JOIN VENTAS2026 v_actual ON CAST(c.Cod_Clie AS REAL) = CAST(v_actual.Cod_Clie AS REAL)
          AND CAST(c.Cdg_Vend AS REAL) = CAST(v_actual.Cdg_Vend AS REAL)
          AND v_actual.Proveedor = 'ARCOR'
        LEFT JOIN VENTAS2026 v_anterior ON CAST(c.Cod_Clie AS REAL) = CAST(v_anterior.Cod_Clie AS REAL)
          AND CAST(c.Cdg_Vend AS REAL) = CAST(v_anterior.Cdg_Vend AS REAL)
          AND v_anterior.Proveedor = 'ARCOR'
        WHERE c.Calif = 'D'
          AND CAST(c.Cdg_Vend AS REAL) = ?
        GROUP BY c.Cod_Clie, c.Raz_Social, c.Vendedor, c.DV, c.Giro
        ORDER BY venta_actual DESC, c.Raz_Social
        """

        cursor.execute(query, (periodo_actual, periodo_anterior, float(cod_vendedor)))
        clientes = [dict(row) for row in cursor.fetchall()]

        conn.close()
        logger.info(f"✅ {len(clientes)} clientes TROYA obtenidos para vendedor {cod_vendedor}")
        return clientes

    except Exception as e:
        logger.error(f"❌ Error consultando clientes TROYA: {e}")
        return []

def generar_mensaje_troya(vendedor, clientes):
    """Genera mensaje TROYA formateado"""
    if not clientes:
        return f"""Hola {vendedor['nombre']}, 👋

No tienes clientes TROYA (Calif=D) con datos.

🤖 Bot N&J"""

    nombre = vendedor['nombre'].strip().title()
    ahora = ahora_local()
    mes_actual = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                  "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"][ahora.month - 1]
    mes_anterior = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
                    "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"][(ahora.month-2) % 12]

    con_compra = [c for c in clientes if c['venta_actual'] > 0]
    sin_compra = [c for c in clientes if c['venta_actual'] == 0]

    total_venta_actual = sum(c['venta_actual'] for c in clientes)
    total_venta_anterior = sum(c['venta_anterior'] for c in clientes)
    diferencia = total_venta_actual - total_venta_anterior

    mensaje = f"""👤 {nombre.upper()}
📊 CLIENTES TROYA - {mes_actual} {ahora.year}

📈 RESUMEN:
Total: {len(clientes)} | ✅ {len(con_compra)} compran | ❌ {len(sin_compra)} sin compra
Ventas: ${total_venta_actual:,.2f} | Anterior: ${total_venta_anterior:,.2f} | Δ ${diferencia:+,.2f}

"""

    if con_compra:
        mensaje += f"✅ CON COMPRA ({len(con_compra)}):\n"
        for c in con_compra[:10]:  # Máximo 10
            diferencia_cliente = c['venta_actual'] - c['venta_anterior']
            mensaje += f"• {c['Raz_Social'][:30]}: ${c['venta_actual']:,.0f} (Δ ${diferencia_cliente:+,.0f})\n"
        if len(con_compra) > 10:
            mensaje += f"... +{len(con_compra)-10} más\n"

    if sin_compra:
        mensaje += f"\n❌ SIN COMPRA ({len(sin_compra)}):\n"
        for c in sin_compra[:5]:  # Máximo 5
            mensaje += f"• {c['Raz_Social'][:30]} (Día {c['DV']})\n"
        if len(sin_compra) > 5:
            mensaje += f"... +{len(sin_compra)-5} más\n"

    mensaje += "\n🤖 Bot N&J Distribuciones"

    return mensaje

def enviar_respuesta_api(numero_destino, mensaje_texto):
    """Envía respuesta por WhatsApp Business API"""
    try:
        payload = {
            "messaging_product": "whatsapp",
            "to": numero_destino,
            "type": "text",
            "text": {"body": mensaje_texto}
        }

        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        response = requests.post(API_URL, json=payload, headers=headers, timeout=10)

        if response.status_code == 200:
            logger.info(f"✅ Mensaje enviado a {numero_destino}")
            return True
        else:
            logger.error(f"❌ Error API ({response.status_code}): {response.text}")
            return False

    except Exception as e:
        logger.error(f"❌ Error enviando: {e}")
        return False

# ============================================================
# 4. WEBHOOK ROUTES
# ============================================================

@app.route("/webhook", methods=["GET"])
def webhook_verify():
    """Verifica webhook con Meta"""
    mode = request.args.get("hub.mode")
    token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if mode == "subscribe" and token == VERIFY_TOKEN:
        logger.info("✅ Webhook verificado")
        return challenge, 200
    else:
        logger.warning("❌ Verificación fallida")
        return jsonify({"status": "error"}), 403

@app.route("/webhook", methods=["POST"])
def webhook_receive():
    """Recibe mensajes entrantes"""
    try:
        data = request.get_json()
        if not data:
            return jsonify({"status": "ok"}), 200

        messages = data.get("entry", [{}])[0].get("changes", [{}])[0].get("value", {}).get("messages", [])
        if not messages:
            return jsonify({"status": "ok"}), 200

        mensaje_obj = messages[0]
        numero_origen = mensaje_obj.get("from")
        texto_mensaje = mensaje_obj.get("text", {}).get("body", "").strip().upper()

        logger.info(f"📩 Mensaje de {numero_origen}: {texto_mensaje[:50]}")

        # Detectar palabra clave TROYA
        if "TROYA" not in texto_mensaje:
            logger.info("⊘ No contiene TROYA")
            return jsonify({"status": "ok"}), 200

        # Cargar vendedores
        vendedores = cargar_vendedores_excel()
        if not vendedores:
            logger.error("❌ No hay vendedores")
            return jsonify({"status": "error"}), 500

        # Buscar vendedor
        vendedor = vendedores.get(numero_origen) or vendedores.get(obtener_ultimo_numero(numero_origen))
        if not vendedor:
            logger.warning(f"⚠️ Vendedor no encontrado: {numero_origen}")
            respuesta = "Hola 👋\n\nNo estás registrado en N&J.\nContacta al equipo.\n\n🤖 Bot N&J"
            enviar_respuesta_api(numero_origen, respuesta)
            return jsonify({"status": "ok"}), 200

        logger.info(f"👤 Vendedor: {vendedor['nombre']} (Código: {vendedor['codigo']})")

        # Obtener clientes TROYA
        clientes_troya = obtener_clientes_troya_vendedor(vendedor['codigo'])

        # Generar mensaje
        mensaje_respuesta = generar_mensaje_troya(vendedor, clientes_troya)

        # Enviar
        logger.info(f"📤 Enviando a {vendedor['nombre']}")
        if enviar_respuesta_api(numero_origen, mensaje_respuesta):
            logger.info(f"✅ Respuesta exitosa")
        else:
            logger.error(f"❌ Error en respuesta")

        return jsonify({"status": "ok"}), 200

    except Exception as e:
        logger.error(f"❌ Error procesando webhook: {e}")
        return jsonify({"status": "error"}), 500

@app.route("/health", methods=["GET"])
def health():
    """Health check"""
    return jsonify({"status": "ok", "timestamp": ahora_local().isoformat()}), 200

# ============================================================
# 5. MAIN
# ============================================================

if __name__ == "__main__":
    logger.info("="*70)
    logger.info("WEBHOOK VENDEDORES V4 - CORRECTO")
    logger.info("="*70)
    logger.info(f"📊 BD: {BD_PATH}")
    logger.info(f"📋 Excel: {EXCEL_VENDEDORES}")
    logger.info(f"🔑 Palabra clave: TROYA")
    logger.info(f"🏷️ Proveedor: ARCOR")
    logger.info(f"🎯 Filtro: Calif='D'")
    logger.info("="*70)

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
