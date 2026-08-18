#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
📢 ENVIAR MENSAJE A TODOS LOS VENDEDORES
Plantilla Meta: VENDEDORES
Uso: python ENVIAR_MENSAJE_TODOS.py
"""

import os
import sqlite3
import requests
import logging
from datetime import datetime
import openpyxl
import time

# ============================================================
# CONFIGURACIÓN
# ============================================================

BASE_PATH = os.path.dirname(os.path.abspath(__file__))

# Credenciales WhatsApp
ACCESS_TOKEN = os.environ.get("WHATSAPP_ACCESS_TOKEN")
PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")
API_VERSION = os.environ.get("WHATSAPP_API_VERSION", "v25.0")

if not ACCESS_TOKEN or not PHONE_NUMBER_ID:
    raise RuntimeError("Faltan variables de entorno WHATSAPP_ACCESS_TOKEN o WHATSAPP_PHONE_NUMBER_ID")

API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"

EXCEL_VENDEDORES = os.path.join(BASE_PATH, "vendedores.xlsx")

# Plantilla Meta
PLANTILLA_NOMBRE = "vendedor"

# Mensaje directo (si no usas plantilla)
MENSAJE_DIRECTO = """Hola! Soy el Bot de Ventas N&J 🤖
Envía 'resumen' y te mando tu reporte al instante.
¡Pruébalo ahora!"""

# Logging
LOG_DIR = os.path.join(BASE_PATH, "logs")
os.makedirs(LOG_DIR, exist_ok=True)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[
        logging.FileHandler(
            os.path.join(LOG_DIR, "enviar_mensaje_todos.log"),
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

# ============================================================
# FUNCIONES
# ============================================================

def normalizar_telefono(numero):
    """Normaliza un teléfono dejando solo dígitos."""
    return "".join(filter(str.isdigit, str(numero)))


def obtener_vendedores():
    """Lee lista de vendedores desde Excel."""
    try:
        logger.info("📋 Leyendo vendedores desde Excel...")
        vendedores = []

        wb = openpyxl.load_workbook(
            EXCEL_VENDEDORES,
            read_only=True,
            data_only=True
        )
        ws = wb.active

        # A=Nombre, B=Teléfono, C=Clientes, D=Rol
        for row in ws.iter_rows(min_row=1, values_only=True):
            if len(row) < 2 or not row[0] or not row[1]:
                continue

            nombre = str(row[0]).strip()
            telefono = normalizar_telefono(row[1])
            rol = str(row[3]).strip().upper() if len(row) > 3 and row[3] else ""

            if not telefono:
                continue

            vendedores.append({
                "nombre": nombre,
                "telefono": telefono,
                "telefono_last9": telefono[-9:] if len(telefono) >= 9 else telefono,
                "rol": rol
            })

        wb.close()

        logger.info(f"✅ {len(vendedores)} vendedores cargados")
        return vendedores

    except Exception as e:
        logger.error(f"❌ Error leyendo Excel: {e}")
        return []


def enviar_mensaje_whatsapp(numero_destino, nombre_vendedor, reintentos=3):
    """Envía mensaje por WhatsApp API con reintentos."""
    try:
        numero = normalizar_telefono(numero_destino)
        if not numero.startswith("51"):
            numero = f"51{numero}"

        payload = {
            "messaging_product": "whatsapp",
            "to": numero,
            "type": "template",
            "template": {
                "name": PLANTILLA_NOMBRE,
                "language": {"code": "en"}
            }
        }

        headers = {
            "Authorization": f"Bearer {ACCESS_TOKEN}",
            "Content-Type": "application/json"
        }

        logger.info(f"📤 Enviando a {nombre_vendedor} ({numero[-4:]})...")

        for intento in range(1, reintentos + 1):
            try:
                response = requests.post(
                    API_URL,
                    json=payload,
                    headers=headers,
                    timeout=15
                )

                if response.ok:
                    logger.info(f"✅ Enviado a {nombre_vendedor} ({numero[-4:]})")
                    return True

                if intento < reintentos:
                    logger.warning(f"⚠️ Intento {intento}/{reintentos} falló. Reintentando en 2s...")
                    time.sleep(2)
                else:
                    logger.error(
                        f"❌ Error enviando a {nombre_vendedor} ({response.status_code}): {response.text}"
                    )
                    return False

            except requests.RequestException as e:
                if intento < reintentos:
                    logger.warning(f"⚠️ Error de conexión intento {intento}/{reintentos}: {e}")
                    time.sleep(2)
                else:
                    logger.error(f"❌ Error de conexión a {nombre_vendedor}: {e}")
                    return False

    except Exception as e:
        logger.error(f"❌ Error enviando a {nombre_vendedor}: {e}")
        return False


def main():
    """Función principal."""
    print("=" * 70)
    print("  📢 ENVIAR MENSAJE A TODOS LOS VENDEDORES")
    print("=" * 70 + "\n")

    logger.info("Iniciando envío de mensajes...")

    # Obtener vendedores
    vendedores = obtener_vendedores()

    if not vendedores:
        logger.error("❌ No hay vendedores para enviar")
        return False

    # Estadísticas
    exitosos = 0
    fallidos = 0
    inicio = datetime.now()

    # Enviar a cada vendedor
    for vendedor in vendedores:
        exito = enviar_mensaje_whatsapp(
            vendedor["telefono"],
            vendedor["nombre"],
            reintentos=3
        )

        if exito:
            exitosos += 1
        else:
            fallidos += 1

        # Pequeña pausa entre envíos para evitar rate limit
        time.sleep(1)

    # Resumen
    duracion = datetime.now() - inicio

    print("\n" + "=" * 70)
    print("  ✅ ENVÍO COMPLETADO")
    print("=" * 70)
    print(f"\n📊 RESUMEN:")
    print(f"  ✅ Exitosos: {exitosos}")
    print(f"  ❌ Fallidos: {fallidos}")
    print(f"  📈 Total: {len(vendedores)}")
    print(f"  ⏱️  Tiempo: {duracion.total_seconds():.1f} segundos")
    print("\n" + "=" * 70 + "\n")

    logger.info(f"Envío completado: {exitosos} exitosos, {fallidos} fallidos")

    return fallidos == 0


if __name__ == "__main__":
    exito = main()
    exit(0 if exito else 1)
