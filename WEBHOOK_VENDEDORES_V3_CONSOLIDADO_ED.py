#!/usr/bin/env python3
# -*- coding: utf-8 -*-

from flask import Flask, request, jsonify
import sqlite3
import requests
import logging
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
import openpyxl
import os
import re
from threading import Lock
from concurrent.futures import ThreadPoolExecutor
import time

# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

BASE_PATH = os.path.dirname(os.path.abspath(__file__))

ACCESS_TOKEN = os.environ.get("WHATSAPP_ACCESS_TOKEN")
PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")
VERIFY_TOKEN = os.environ.get("WHATSAPP_VERIFY_TOKEN")
API_VERSION = os.environ.get("WHATSAPP_API_VERSION", "v25.0")

if not ACCESS_TOKEN:
    raise RuntimeError("Falta WHATSAPP_ACCESS_TOKEN")
if not PHONE_NUMBER_ID:
    raise RuntimeError("Falta WHATSAPP_PHONE_NUMBER_ID")
if not VERIFY_TOKEN:
    raise RuntimeError("Falta WHATSAPP_VERIFY_TOKEN")

API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"

BD_PATH = os.environ.get("BD_PATH", os.path.join(BASE_PATH, "ventas.db"))
EXCEL_VENDEDORES = os.environ.get("EXCEL_VENDEDORES", os.path.join(BASE_PATH, "vendedores.xlsx"))

def cargar_dias_laborables():
    valor = os.environ.get("DIAS_LABORABLES", "0,1,2,3,4,5")
    try:
        dias = {int(x.strip()) for x in valor.split(",") if x.strip() != ""}
        if not dias.issubset(set(range(7))):
            raise ValueError
        return dias
    except ValueError:
        raise RuntimeError("DIAS_LABORABLES inválido. Ejemplo: 0,1,2,3,4,5")

DIAS_LABORABLES = cargar_dias_laborables()

try:
    TZ_LIMA = ZoneInfo("America/Lima")
except Exception:
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
        logging.FileHandler(os.path.join(LOG_DIR, "webhook_v3_troya.log"), encoding="utf-8"),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# ============================================================
# 3. UTILIDADES
# ============================================================

MESES_ES = ["ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
            "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"]

def ahora_local():
    if TZ_LIMA:
        return datetime.now(TZ_LIMA)
    return datetime.now()

def obtener_contexto_periodo():
    ahora = ahora_local()
    hoy = ahora.date()
    anio = hoy.year
    mes = hoy.month
    primer_dia = date(anio, mes, 1)
    primer_dia_siguiente = date(anio + 1, 1, 1) if mes == 12 else date(anio, mes + 1, 1)
    ultimo_dia = primer_dia_siguiente - timedelta(days=1)
    return {
        "hoy": hoy,
        "ahora": ahora,
        "anio": anio,
        "mes": mes,
        "periodo": f"{anio}{mes:02d}",
        "nombre_mes": MESES_ES[mes - 1],
        "primer_dia": primer_dia,
        "ultimo_dia": ultimo_dia,
    }

def calcular_dias_laborables_periodo():
    ctx = obtener_contexto_periodo()
    hoy = ctx["hoy"]
    primer_dia = ctx["primer_dia"]
    ultimo_dia = ctx["ultimo_dia"]
    fecha_corte = min(max(hoy, primer_dia), ultimo_dia)
    transcurridos = restantes = total = 0
    fecha = primer_dia
    while fecha <= ultimo_dia:
        if fecha.weekday() in DIAS_LABORABLES:
            total += 1
            if fecha <= fecha_corte:
                transcurridos += 1
            elif fecha > hoy:
                restantes += 1
        fecha += timedelta(days=1)
    if hoy > ultimo_dia:
        restantes = 0
        transcurridos = total
    if hoy < primer_dia:
        transcurridos = 0
        restantes = total
    return {"dias_laborables_total": total, "dias_transcurridos": transcurridos, "dias_restantes": restantes}

def normalizar_telefono(numero):
    return "".join(filter(str.isdigit, str(numero)))

# ============================================================
# 4. DATABASE & CACHE
# ============================================================

_DB_INIT_LOCK = Lock()
_DB_INITIALIZED = False
_EXCEL_CACHE = {}
_EXCEL_LOCK = Lock()

def get_db_connection():
    conn = sqlite3.connect(BD_PATH, timeout=30)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA busy_timeout = 30000")
    conn.execute("PRAGMA journal_mode = WAL")
    conn.execute("PRAGMA synchronous = NORMAL")
    return conn

def inicializar_bd():
    global _DB_INITIALIZED
    if _DB_INITIALIZED:
        return
    with _DB_INIT_LOCK:
        if _DB_INITIALIZED:
            return
        conn = None
        try:
            conn = get_db_connection()
            conn.execute("""
                CREATE TABLE IF NOT EXISTS webhook_procesados (
                    message_id TEXT PRIMARY KEY,
                    telefono TEXT,
                    recibido_en TEXT NOT NULL
                )
            """)
            conn.commit()
            _DB_INITIALIZED = True
            logger.info("✅ BD inicializada")
        except Exception as e:
            logger.error(f"❌ Error inicializando BD: {e}")
            raise
        finally:
            if conn:
                conn.close()

# Un único punto de escritura para la deduplicación.
_MESSAGE_CLAIM_LOCK = Lock()


def reclamar_message_id(message_id, telefono):
    inicializar_bd()
    if not message_id:
        return True
    conn = None
    try:
        with _MESSAGE_CLAIM_LOCK:
            conn = get_db_connection()
            cursor = conn.execute(
                """INSERT OR IGNORE INTO webhook_procesados (message_id, telefono, recibido_en) VALUES (?, ?, ?)""",
                (message_id, telefono, ahora_local().isoformat())
            )
            conn.commit()
        return cursor.rowcount == 1
    except Exception as e:
        logger.error(f"❌ No se pudo reclamar message_id {message_id}: {e}")
        return None
    finally:
        if conn:
            conn.close()


def liberar_message_id(message_id):
    if not message_id:
        return
    conn = None
    try:
        conn = get_db_connection()
        conn.execute("DELETE FROM webhook_procesados WHERE message_id = ?", (message_id,))
        conn.commit()
    except Exception as e:
        logger.error(f"❌ Error liberando message_id: {e}")
    finally:
        if conn:
            conn.close()


# ============================================================
# 5. VENDEDORES
# ============================================================

def obtener_vendedores_autorizados():
    try:
        mtime = os.path.getmtime(EXCEL_VENDEDORES)
        with _EXCEL_LOCK:
            if _EXCEL_CACHE and mtime == _EXCEL_CACHE.get("mtime"):
                return _EXCEL_CACHE.get("data", {})

        vendedores = {}
        wb = openpyxl.load_workbook(EXCEL_VENDEDORES, read_only=True, data_only=True)
        ws = wb.active

        for row in ws.iter_rows(min_row=1, values_only=True):
            if len(row) < 2 or not row[0] or not row[1]:
                continue
            nombre = str(row[0]).strip()
            telefono = normalizar_telefono(row[1])
            if not telefono:
                continue
            rol = str(row[3]).strip().upper() if len(row) > 3 and row[3] else ""
            vendedores[telefono] = {
                "nombre": nombre,
                "codigo": row[2] if len(row) > 2 else nombre,
                "telefono": telefono,
                "telefono_last9": telefono[-9:] if len(telefono) >= 9 else telefono,
                "rol": rol,
            }
        wb.close()

        with _EXCEL_LOCK:
            _EXCEL_CACHE["data"] = vendedores
            _EXCEL_CACHE["mtime"] = mtime

        logger.info(f"✅ {len(vendedores)} vendedores cargados")
        return vendedores
    except Exception as e:
        logger.error(f"❌ Error leyendo Excel: {e}")
        return {}

def validar_vendedor(numero_telefonico):
    numero_limpio = normalizar_telefono(numero_telefonico)
    vendedores = obtener_vendedores_autorizados()

    if numero_limpio in vendedores:
        usuario = vendedores[numero_limpio]
        logger.info(f"✅ Vendedor: {usuario['nombre']}")
        return True, usuario

    last9 = numero_limpio[-9:] if len(numero_limpio) >= 9 else numero_limpio
    for usuario in vendedores.values():
        if usuario.get("telefono_last9") == last9:
            logger.info(f"✅ Vendedor (por últimos 9): {usuario['nombre']}")
            return True, usuario

    logger.warning(f"⚠️ Número no autorizado: {numero_limpio[-4:] if numero_limpio else 'vacío'}")
    return False, None

def es_jefe(rol=""):
    rol = (rol or "").strip().upper()
    if rol:
        palabras_clave = {"JEFE", "SUPERVISOR", "GERENTE", "COORDINADOR"}
        return any(palabra in rol for palabra in palabras_clave)
    return False

# ============================================================
# 6. OBTENER CLIENTES TROYA - QUERIES CORREGIDAS
# ============================================================

def obtener_clientes_troya(nombre_vendedor):
    """Obtiene clientes TROYA del vendedor sin duplicar importes.

    Reglas del reporte vendedor:
      - Cliente debe ser Calif='D' en la tabla clientes.
      - Vendedor se determina por clientes.Vendedor.
      - Ventas deben ser ARCOR.
      - Se consideran únicamente período actual y período anterior.
      - El cruce conserva Cod_Clie + Cdg_Vend para respetar la cartera
        asignada al vendedor.
      - Las ventas se agregan antes del JOIN para evitar multiplicaciones
        entre movimientos del período actual y anterior.
    """
    try:
        logger.info(f"🔍 Buscando TROYA para: '{nombre_vendedor}'")

        conn = sqlite3.connect(BD_PATH, timeout=30)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        ahora = ahora_local()
        periodo_actual = f"{ahora.year}{ahora.month:02d}"
        mes_anterior = ahora.month - 1 if ahora.month > 1 else 12
        anio_anterior = ahora.year if ahora.month > 1 else ahora.year - 1
        periodo_anterior = f"{anio_anterior}{mes_anterior:02d}"

        # Primero agregamos las ventas por cliente + vendedor + período.
        # Esto evita el producto cruzado que inflaba los importes cuando
        # existían varias ventas en ambos períodos.
        query = """
        WITH ventas_agregadas AS (
            SELECT
                CAST(Cod_Clie AS REAL) AS Cod_Clie_real,
                CAST(Cdg_Vend AS REAL) AS Cdg_Vend_real,
                Periodo,
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) AS venta
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND Periodo IN (?, ?)
            GROUP BY
                CAST(Cod_Clie AS REAL),
                CAST(Cdg_Vend AS REAL),
                Periodo
        )
        SELECT
            c.Cod_Clie,
            c.Raz_Social,
            c.Vendedor,
            c.DV,
            c.Giro,
            COALESCE(MAX(CASE
                WHEN va.Periodo = ? THEN va.venta
                ELSE 0
            END), 0) AS venta_actual,
            COALESCE(MAX(CASE
                WHEN va.Periodo = ? THEN va.venta
                ELSE 0
            END), 0) AS venta_anterior
        FROM clientes c
        LEFT JOIN ventas_agregadas va
          ON CAST(c.Cod_Clie AS REAL) = va.Cod_Clie_real
         AND CAST(c.Cdg_Vend AS REAL) = va.Cdg_Vend_real
        WHERE c.Calif = 'D'
          AND c.Vendedor LIKE ?
        GROUP BY
            c.Cod_Clie,
            c.Raz_Social,
            c.Vendedor,
            c.DV,
            c.Giro
        ORDER BY
            venta_actual DESC,
            c.Raz_Social
        """

        cursor.execute(
            query,
            (periodo_actual, periodo_anterior, periodo_actual,
             periodo_anterior, f"%{nombre_vendedor}%")
        )
        clientes = [dict(row) for row in cursor.fetchall()]

        for cliente in clientes:
            cliente["tiene_compras"] = float(cliente.get("venta_actual") or 0) > 0

        total_actual = sum(float(c.get("venta_actual") or 0) for c in clientes)
        total_anterior = sum(float(c.get("venta_anterior") or 0) for c in clientes)

        logger.info(
            f"📊 TROYA vendedor {nombre_vendedor}: "
            f"actual S/. {total_actual:,.2f} | "
            f"anterior S/. {total_anterior:,.2f} | "
            f"clientes {len(clientes)}"
        )

        conn.close()
        return clientes

    except Exception as e:
        logger.error(f"❌ Error obteniendo clientes TROYA: {e}")
        return []


def obtener_clientes_troya_generales():
    """Obtiene clientes TROYA consolidados para jefe/supervisor.

    La venta TROYA se toma de VENTAS2026 como fuente de verdad:
      - Proveedor = ARCOR
      - Calif = D
      - período actual/anterior

    Las ventas se agregan primero por Cod_Clie + período y luego se
    relacionan con clientes SOLO por Cod_Clie.

    No se utiliza Cdg_Vend para decidir si una venta entra al
    consolidado. La conciliación mostró diferencias de Cdg_Vend entre
    VENTAS2026 y clientes que excluían ventas válidas.
    """
    try:
        logger.info("🔍 Buscando TROYA GENERAL para todos los vendedores")

        conn = sqlite3.connect(BD_PATH, timeout=30)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        ahora = ahora_local()
        periodo_actual = f"{ahora.year}{ahora.month:02d}"
        mes_anterior = ahora.month - 1 if ahora.month > 1 else 12
        anio_anterior = ahora.year if ahora.month > 1 else ahora.year - 1
        periodo_anterior = f"{anio_anterior}{mes_anterior:02d}"

        # Fuente de verdad para importes:
        # VENTAS2026 filtrada por ARCOR + Calif='D'.
        # Se consolida antes del JOIN para evitar multiplicación de filas.
        query = """
        WITH ventas_agregadas AS (
            SELECT
                CAST(Cod_Clie AS REAL) AS Cod_Clie_real,
                Periodo,
                ROUND(SUM(CAST(Imp_Total AS REAL)), 2) AS venta
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND Calif = 'D'
              AND Periodo IN (?, ?)
            GROUP BY
                CAST(Cod_Clie AS REAL),
                Periodo
        )
        SELECT
            c.Cod_Clie,
            c.Raz_Social,
            c.Vendedor,
            c.DV,
            COALESCE(MAX(CASE
                WHEN va.Periodo = ? THEN va.venta
                ELSE 0
            END), 0) AS venta_actual,
            COALESCE(MAX(CASE
                WHEN va.Periodo = ? THEN va.venta
                ELSE 0
            END), 0) AS venta_anterior
        FROM clientes c
        LEFT JOIN ventas_agregadas va
          ON CAST(c.Cod_Clie AS REAL) = va.Cod_Clie_real
        WHERE c.Calif = 'D'
        GROUP BY
            c.Cod_Clie,
            c.Raz_Social,
            c.Vendedor,
            c.DV
        ORDER BY
            c.Vendedor,
            venta_actual DESC,
            c.Raz_Social
        """

        cursor.execute(
            query,
            (periodo_actual, periodo_anterior, periodo_actual, periodo_anterior)
        )
        clientes = [dict(row) for row in cursor.fetchall()]

        for cliente in clientes:
            cliente["tiene_compras"] = float(cliente.get("venta_actual") or 0) > 0

        total_actual = sum(float(c.get("venta_actual") or 0) for c in clientes)
        total_anterior = sum(float(c.get("venta_anterior") or 0) for c in clientes)

        logger.info(
            f"📊 Encontrados: {len(clientes)} clientes TROYA TOTAL | "
            f"Actual: S/. {total_actual:,.2f} | "
            f"Anterior: S/. {total_anterior:,.2f}"
        )

        conn.close()
        return clientes

    except Exception as e:
        logger.error(f"❌ Error obteniendo clientes TROYA generales: {e}")
        return []

# ============================================================
# 7. GENERADOR DE MENSAJES TROYA
# ============================================================

def obtener_datos_concurso_chenobyl(nombre_vendedor):
    """Obtiene CHERNOBYL y TABLETAS para el concurso del mes actual."""
    PRODUCTOS_CHERNOBYL = [
        "MOGUL EXTREME ROCKS 12*10*45GR",
        "MOGUL INDIVIDUAL GUSANO EXTREME 20*90GR",
        "ROLLO MOGUL ACIDO EXTREME 12X12X35GR",
        "MOGUL INDIVIDUAL OSO EXTREME 24*80GR PERU",
        "MOGUL INDIVIDUAL SANDIA EXTREME 24*80GR PERU",
        "MOGUL OSO EXTREME 12X10X55G",
        "OSITO EXTREME 12X12X25GR",
        "MOGUL TUBITO TUTTIFRUTTI EXTREME 30*70GR",
        "MOGUL JELLY BEAN EXTREME 12X10X50GR",
        "MOGUL CONFITADO EXTREME 12*16*30GR",
        "MOGUL CONFITADO INDIVID EXTREME 40*70GR",
        "MOGUL TUBITO EXTREME FRUTILLA 30*70GR",
        "ROLLO MOGUL ACIDO EXT 12X12X35GR",
        "ROLLO MOGUL FRUTALES 12*12*35GR PERU",
        "ROLLO MOGUL RED BERRIES 12X12X35GR",
        "MOGUL LENGUA MAXX EXTREME TUTI 6*24*15",
        "MOGUL OSITO EXTREME 25GR - GRANEL",
        "MOGUL SANDIA EXTREME 12*10*50GR",
    ]
    PRODUCTOS_TABLETAS = [
        "NIKOLO CAFE **NUEVO*** 10DSP*12UU* 29GR",
        "NIKOLO PEQUEÑO 10DSP*300GR",
        "NIKOLO XL  CHOCOLATE 12D*516GR*12UU",
        "NIKOLO PEQUEÑO MENTA 10DSP*276GR",
        "NIKOLO PEQUEÑO FRESA 10DSP*12UU* 29GR",
        "TABLETA GALLETA BOB 6DSP*41GR*18UU",
        "PRIVILEGIO TABLETA MANJAR 10*348GR",
    ]
    PRODUCTOS_GALLETAS = [
        "GALLETA BOB VAINILLA 14X216GR",
        "GALLETA BOB CHOCOLATE 14X216GR",
        "GALLETA SAPITO FRESA 14X216GR",
        "GALLETA SAPITO VAINILLA 14X216GR",
        "GALLETA BOB PACK X 6UU CHOCO 20*24GR",
        "GALLETA TORTINI BON O BON 65*90GR",
        "GALLETA BOB PACK X 6UU VAINILLA 20*24GR",
    ]

    conn = None
    try:
        ctx = obtener_contexto_periodo()
        periodo = ctx["periodo"]
        conn = get_db_connection()
        cursor = conn.cursor()

        def grupo(productos):
            ph = ",".join("?" for _ in productos)
            params = [f"%{nombre_vendedor}%", periodo] + [p.upper().strip() for p in productos]
            cursor.execute(f"""
                SELECT TRIM(Producto) producto,
                       COALESCE(SUM(CAST(Imp_Total AS REAL)),0) venta
                FROM VENTAS2026
                WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor='ARCOR'
                  AND UPPER(TRIM(COALESCE(Producto,''))) IN ({ph})
                  AND CAST(Imp_Total AS REAL) > 0
                GROUP BY UPPER(TRIM(Producto))
            """, params)
            ventas = {str(r["producto"]).strip().upper(): float(r["venta"] or 0)
                      for r in cursor.fetchall()}

            cursor.execute(f"""
                SELECT COALESCE(SUM(CAST(Imp_Total AS REAL)),0) venta_total,
                       COUNT(DISTINCT Cod_Clie) clientes_unicos
                FROM VENTAS2026
                WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor='ARCOR'
                  AND UPPER(TRIM(COALESCE(Producto,''))) IN ({ph})
                  AND CAST(Imp_Total AS REAL) > 0
            """, params)
            r = cursor.fetchone()

            detalle = [{"producto": p, "venta": ventas.get(p.upper().strip(), 0.0)}
                       for p in productos]
            sin_venta = [x for x in detalle if x["venta"] <= 0]
            menor = sorted([x for x in detalle if x["venta"] > 0],
                           key=lambda x: x["venta"])[:3]
            return {
                "venta": float(r["venta_total"] or 0),
                "clientes": int(r["clientes_unicos"] or 0),
                "sin_venta": sin_venta,
                "menor": menor,
            }

        ch = grupo(PRODUCTOS_CHERNOBYL)
        tb = grupo(PRODUCTOS_TABLETAS)
        ga = grupo(PRODUCTOS_GALLETAS)

        return {
            "periodo": periodo,
            "nombre_mes": ctx["nombre_mes"],
            "venta_total": ch["venta"],
            "clientes_unicos": ch["clientes"],
            "productos_sin_venta": ch["sin_venta"],
            "productos_menor_venta": ch["menor"],
            "venta_total_tabletas": tb["venta"],
            "clientes_unicos_tabletas": tb["clientes"],
            "tabletas_sin_venta": tb["sin_venta"],
            "tabletas_menor_venta": tb["menor"],
            "venta_total_galletas": ga["venta"],
            "clientes_unicos_galletas": ga["clientes"],
            "galletas_sin_venta": ga["sin_venta"],
            "galletas_menor_venta": ga["menor"],
        }
    except Exception as e:
        logger.exception(f"❌ Error obteniendo CONCURSO para {nombre_vendedor}: {e}")
        return None
    finally:
        if conn:
            conn.close()


def generar_mensaje_concurso_chenobyl(datos):
    if not datos:
        return "⚠️ No se pudo obtener la información del concurso."

    mensaje = f"""🏆 CONCURSO {datos['nombre_mes'].upper()} {datos['periodo'][:4]}

━━━━━━━━━━━━━━━━━━━━
☢️ CHERNOBYL
━━━━━━━━━━━━━━━━━━━━

💰 VENTAS: S/. {datos['venta_total']:,.2f}
👥 COBERTURA: {datos['clientes_unicos']} clientes únicos

"""

    sin_venta = datos.get("productos_sin_venta", [])
    menor = datos.get("productos_menor_venta", [])
    if sin_venta:
        mensaje += "❌ SIN VENTA:\n" + "".join(f"• {x['producto']}\n" for x in sin_venta)
    if menor:
        mensaje += "\n⚠️ 3 CON MENOR VENTA:\n" + "".join(
            f"• {x['producto']}\n" for x in menor
        )

    mensaje += f"""
━━━━━━━━━━━━━━━━━━━━
🍫 TABLETAS
━━━━━━━━━━━━━━━━━━━━

💰 VENTAS: S/. {datos['venta_total_tabletas']:,.2f}
👥 COBERTURA: {datos['clientes_unicos_tabletas']} clientes únicos

"""
    sin_venta = datos.get("tabletas_sin_venta", [])
    menor = datos.get("tabletas_menor_venta", [])
    if sin_venta:
        mensaje += "❌ SIN VENTA:\n" + "".join(f"• {x['producto']}\n" for x in sin_venta)
    if menor:
        mensaje += "\n⚠️ 3 CON MENOR VENTA:\n" + "".join(
            f"• {x['producto']}\n" for x in menor
        )

    mensaje += f"""
━━━━━━━━━━━━━━━━━━━━
🍪 GALLETAS
━━━━━━━━━━━━━━━━━━━━

💰 VENTAS: S/. {datos['venta_total_galletas']:,.2f}
👥 COBERTURA: {datos['clientes_unicos_galletas']} clientes únicos

"""
    sin_venta = datos.get("galletas_sin_venta", [])
    menor = datos.get("galletas_menor_venta", [])

    if sin_venta:
        mensaje += "❌ SIN VENTA:\n" + "".join(
            f"• {x['producto']}\n" for x in sin_venta
        )

    if menor:
        mensaje += "\n⚠️ 3 CON MENOR VENTA:\n" + "".join(
            f"• {x['producto']}\n" for x in menor
        )

    mensaje += "\n━━━━━━━━━━━━━━━━━━━━\n🚀 ¡A seguir impulsando el concurso! 💪🔥\n\n🤖 Bot N&J"
    return mensaje

def generar_mensaje_troya(vendedor, clientes):
    """Genera reporte personalizado de clientes TROYA (formato compacto)"""

    if not clientes:
        return f"""Hola {vendedor.get('nombre', 'Vendedor')}, 👋

No tienes clientes TROYA registrados actualmente.

🤖 Bot N&J"""

    nombre = vendedor.get('nombre', 'Vendedor').strip().title()
    cantidad = len(clientes)
    con_compra = sum(1 for c in clientes if c.get('tiene_compras'))
    sin_compra = cantidad - con_compra

    ahora = ahora_local()
    mes_nombre = MESES_ES[ahora.month - 1]
    mes_anterior = ahora.month - 1 if ahora.month > 1 else 12
    mes_anterior_nombre = MESES_ES[mes_anterior - 1]

    # Totales
    total_actual = sum(c.get('venta_actual', 0) for c in clientes)

    mensaje = f"""👤 {nombre}
📅 {mes_nombre} 2026 | {cantidad} TROYA (✅{con_compra} ❌{sin_compra})

✅ CON COMPRA
"""

    # Clientes CON COMPRA
    con_compra_list = [c for c in clientes if c.get('tiene_compras')]
    if con_compra_list:
        for cliente in con_compra_list:
            cliente_nombre = cliente['Raz_Social'].strip().title()[:24]
            venta_act = cliente.get('venta_actual', 0)
            mensaje += f"S/. {venta_act:<7.0f} {cliente_nombre}\n"

    # Clientes SIN COMPRA
    sin_compra_list = [c for c in clientes if not c.get('tiene_compras')]
    if sin_compra_list:
        mensaje += f"\n❌ SIN COMPRA\n"
        for cliente in sin_compra_list:
            cliente_nombre = cliente['Raz_Social'].strip().title()[:20]
            dia = cliente.get('DV', 'N/A')
            mensaje += f"{cliente_nombre} {dia}\n"

    # Resumen y comparativo contra el mes anterior.
    total_anterior = sum(float(c.get("venta_anterior") or 0) for c in clientes)
    diferencia = total_actual - total_anterior
    if total_anterior > 0:
        variacion_pct = (diferencia / total_anterior) * 100
    else:
        variacion_pct = 100.0 if total_actual > 0 else 0.0

    mensaje += (
        f"\n📊 Total {mes_nombre}: S/. {total_actual:,.0f}"
        f"\n📈 Comparativo TROYA"
        f"\n├─ {mes_anterior_nombre}: S/. {total_anterior:,.0f}"
        f"\n├─ {mes_nombre}: S/. {total_actual:,.0f}"
        f"\n└─ Variación: S/. {diferencia:+,.0f} ({variacion_pct:+.1f}%)"
    )

    if total_actual > total_anterior:
        mensaje += (
            "\n\n🏆 ¡FELICITACIONES! 🎉"
            "\nSuperaste tus ventas a clientes TROYA del mes anterior."
            "\n¡Sigue así y vamos por más! 💪🚀"
        )
    elif total_actual < total_anterior:
        mensaje += (
            "\n\n🎯 ¡A ENFOCARSE EN TROYA!"
            "\nEste mes estás por debajo de tu venta del mes anterior."
            "\nVisita y trabaja especialmente a tus clientes TROYA "
            "para recuperar la brecha. 💪🔥"
        )
    else:
        mensaje += (
            "\n\n💪 ¡VAMOS POR MÁS TROYA!"
            "\nMantienes el mismo nivel de ventas del mes anterior."
            "\nCada visita y cada compra cuenta. 🚀"
        )

    mensaje += "\n\n🤖 Bot N&J"

    return mensaje


def generar_mensaje_troya_jefe(clientes):
    """Genera reporte consolidado de TROYA para jefe/supervisor"""

    if not clientes:
        return """👨‍💼 CLIENTES TROYA CONSOLIDADOS

No hay clientes TROYA registrados.

🤖 Bot N&J"""

    ahora = ahora_local()
    mes_nombre = MESES_ES[ahora.month - 1]
    mes_anterior = ahora.month - 1 if ahora.month > 1 else 12
    mes_anterior_nombre = MESES_ES[mes_anterior - 1]

    # Totales por período
    total_actual = sum(c.get('venta_actual', 0) for c in clientes)
    total_anterior = sum(c.get('venta_anterior', 0) for c in clientes)
    diferencia = total_actual - total_anterior
    pct = (diferencia / total_anterior * 100) if total_anterior > 0 else 0

    # Desglose TROYA: ARCOR neto de SAYON + SAYON.
    # SAYON se identifica por nombre de producto dentro del universo TROYA.
    periodo_actual = f"{ahora.year}{ahora.month:02d}"
    conn = None
    try:
        conn = get_db_connection()
        row = conn.execute("""
            SELECT COALESCE(SUM(CAST(Imp_Total AS REAL)), 0) AS sayon
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND Calif = 'D'
              AND Periodo = ?
              AND UPPER(COALESCE(Producto, '')) LIKE '%SAYON%'
        """, (periodo_actual,)).fetchone()
        sayon_actual = float(row["sayon"] or 0)
    finally:
        if conn:
            conn.close()

    arcor_neto = total_actual - sayon_actual

    # Agrupar por vendedor
    por_vendedor = {}
    for cliente in clientes:
        vendedor = cliente.get('Vendedor', 'SIN VENDEDOR')
        if vendedor not in por_vendedor:
            por_vendedor[vendedor] = {'sin_compra': 0}
        if not cliente.get('tiene_compras'):
            por_vendedor[vendedor]['sin_compra'] += 1

    # Contar totales para resumen
    total_clientes = len(clientes)
    con_compra_total = sum(1 for c in clientes if c.get('tiene_compras'))
    sin_compra_total = total_clientes - con_compra_total
    ticket_promedio_general = (
        sum(float(c.get('venta_actual') or 0) for c in clientes if float(c.get('venta_actual') or 0) > 0)
        / max(con_compra_total, 1)
    )

    mensaje = f"""👨‍💼 CLIENTES TROYA CONSOLIDADOS
{mes_nombre} 2026 | {ahora_local().strftime('%d/%m/%Y %H:%M')}

Total {mes_nombre}: S/. {total_actual:,.0f}
Total {mes_anterior_nombre}: S/. {total_anterior:,.0f}
Diferencia: S/. {diferencia:+,.0f} ({pct:+.1f}%)

💼 VENTAS {mes_nombre}:
├─ ARCOR: S/. {arcor_neto:,.0f}
└─ SAYON: S/. {sayon_actual:,.0f}

📊 TOTAL: {total_clientes} | ✅ {con_compra_total} | ❌ {sin_compra_total} | 💵 S/. {ticket_promedio_general:.0f}

```
VENDEDOR                 SIN COMPRA
───────────────────────────────────
"""

    for vendedor in sorted(por_vendedor.keys()):
        datos = por_vendedor[vendedor]
        nombre_corto = vendedor.split()[0] if vendedor else vendedor
        mensaje += f"{nombre_corto:<24}{datos['sin_compra']:>10}\n"

    mensaje += "```\n"

    mensaje += "\n🤖 Bot N&J Distribuciones"

    return mensaje

# ============================================================
# 8. REPORTE VENDEDOR (EXISTENTE)
# ============================================================

def obtener_datos_vendedor(nombre_vendedor):
    """Extrae KPIs del vendedor para el período actual"""
    conn = None
    try:
        ctx = obtener_contexto_periodo()
        periodo = ctx["periodo"]
        anio = ctx["anio"]
        mes = ctx["mes"]
        dias = calcular_dias_laborables_periodo()

        conn = get_db_connection()
        cursor = conn.cursor()

        cursor.execute("""
            SELECT
                ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS total_ventas,
                COUNT(DISTINCT Cod_Clie) AS clientes
            FROM VENTAS2026
            WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor = 'ARCOR'
        """, (f"%{nombre_vendedor}%", periodo))

        venta_data = cursor.fetchone()

        datos = {
            "vendedor": nombre_vendedor,
            "total_ventas": venta_data["total_ventas"] or 0,
            "clientes": venta_data["clientes"] or 0,
            "periodo": periodo,
            "nombre_mes": ctx["nombre_mes"],
            "dias_transcurridos": dias["dias_transcurridos"],
            "dias_restantes": dias["dias_restantes"],
            "dias_laborables_total": dias["dias_laborables_total"],
        }

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Cod_Clie), 0), 2), 0) AS ticket
            FROM VENTAS2026
            WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor = 'ARCOR'
              AND CAST(Imp_Total AS REAL) > 0
        """, (f"%{nombre_vendedor}%", periodo))
        ticket = cursor.fetchone()
        datos["ticket_promedio"] = ticket["ticket"] or 0

        cursor.execute("""
            SELECT Cuota_Soles, Cuota_Cobertura
            FROM cuotas
            WHERE Vendedor = ? AND AÑO = ? AND NRO_MES = ? AND Proveedor = 'ARCOR'
            LIMIT 1
        """, (nombre_vendedor, anio, mes))

        cuota_data = cursor.fetchone()
        datos["cuota"] = float(cuota_data["Cuota_Soles"] or 0) if cuota_data else 0
        datos["cuota_cobertura"] = int(cuota_data["Cuota_Cobertura"] or 0) if cuota_data else 0
        datos["cumplimiento"] = (datos["total_ventas"] / datos["cuota"] * 100 if datos["cuota"] > 0 else 0)

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)), 2), 0) AS troya
            FROM VENTAS2026
            WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor = 'ARCOR' AND Calif = 'D'
        """, (f"%{nombre_vendedor}%", periodo))
        troya_data = cursor.fetchone()
        datos["ventas_troya"] = troya_data["troya"] or 0

        # Desglose TROYA individual: productos cuyo nombre contiene SAYON.
        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)), 2), 0) AS sayon
            FROM VENTAS2026
            WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor = 'ARCOR'
              AND Calif = 'D'
              AND UPPER(COALESCE(Producto, '')) LIKE '%SAYON%'
        """, (f"%{nombre_vendedor}%", periodo))
        sayon_data = cursor.fetchone()
        datos["ventas_sayon"] = sayon_data["sayon"] or 0
        datos["ventas_arcor_neto"] = datos["ventas_troya"] - datos["ventas_sayon"]

        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS clientes_troya
            FROM VENTAS2026
            WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor = 'ARCOR' AND Calif = 'D'
        """, (f"%{nombre_vendedor}%", periodo))
        troya_comp = cursor.fetchone()
        datos["clientes_troya_compraron"] = troya_comp["clientes_troya"] or 0

        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS total_clientes_d
            FROM clientes
            WHERE Vendedor LIKE ? AND Calif = 'D'
        """, (f"%{nombre_vendedor}%",))
        total_d = cursor.fetchone()
        total_d_clientes = total_d["total_clientes_d"] or 0
        datos["clientes_troya_no_compraron"] = max(0, total_d_clientes - datos["clientes_troya_compraron"])

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Cod_Clie), 0), 2), 0) AS ticket_troya
            FROM VENTAS2026
            WHERE Vendedor LIKE ? AND Periodo = ? AND Proveedor = 'ARCOR' AND Calif = 'D'
              AND CAST(Imp_Total AS REAL) > 0
        """, (f"%{nombre_vendedor}%", periodo))
        ticket_troya = cursor.fetchone()
        datos["ticket_troya"] = ticket_troya["ticket_troya"] or 0

        dias_transcurridos = max(dias["dias_transcurridos"], 1)
        dias_restantes = max(dias["dias_restantes"], 0)
        promedio_diario = datos["total_ventas"] / dias_transcurridos
        datos["venta_promedio_diaria"] = round(promedio_diario, 2)
        datos["proyeccion_ventas"] = round(datos["total_ventas"] + promedio_diario * dias_restantes, 2)
        datos["cumplimiento_proyectado"] = (datos["proyeccion_ventas"] / datos["cuota"] * 100 if datos["cuota"] > 0 else 0)

        promedio_cobertura = datos["clientes"] / dias_transcurridos
        datos["proyeccion_cobertura"] = int(round(datos["clientes"] + promedio_cobertura * dias_restantes))
        datos["cumplimiento_cobertura_proyectado"] = (datos["proyeccion_cobertura"] / datos["cuota_cobertura"] * 100 if datos["cuota_cobertura"] > 0 else 0)

        logger.info(f"✅ Datos obtenidos para {nombre_vendedor}")
        return datos
    except Exception as e:
        logger.exception(f"❌ Error obteniendo datos: {e}")
        return None
    finally:
        if conn:
            conn.close()

def generar_mensaje_vendedor(datos):
    """Genera reporte detallado para vendedor"""
    if not datos:
        return None

    cumpl = datos["cumplimiento"]
    cumpl_cobertura_proy = datos["cumplimiento_cobertura_proyectado"]

    proyectado_emoji = ("🟢" if datos["cumplimiento_proyectado"] >= 90 else "🟡" if datos["cumplimiento_proyectado"] >= 75 else "🔴")
    cobertura_emoji = ("🟢" if cumpl_cobertura_proy >= 90 else "🟡" if cumpl_cobertura_proy >= 75 else "🔴")
    cumplimiento_emoji = ("🟢" if cumpl >= 90 else "🟡" if cumpl >= 75 else "🔴")

    return f"""📊 REPORTE PERSONAL - {datos['nombre_mes']}
{ahora_local().strftime('%d/%m/%Y %H:%M')}

👤 {datos['vendedor'].upper()}

🎯 OBJETIVOS MES:
├─ Ventas: S/. {datos['cuota']:,.2f}
└─ Cobertura: {datos['cuota_cobertura']} clientes

💼 DESEMPEÑO:
├─ Ventas: S/. {datos['total_ventas']:,.2f}
├─ Avance: {cumpl:.1f}% {cumplimiento_emoji}
├─ Cobertura: {datos['clientes']} clientes
├─ Ticket : S/. {datos['ticket_promedio']:,.2f}
├─ ARCOR: S/. {datos['ventas_arcor_neto']:,.2f}\n├─ SAYON: S/. {datos['ventas_sayon']:,.2f}\n└─ TROYA: S/. {datos['ventas_troya']:,.2f}

⚠️ TROYA RESUMEN:
├─ Compraron: {datos['clientes_troya_compraron']} clientes
├─ No Compraron: {datos['clientes_troya_no_compraron']} clientes
└─ Ticket: S/. {datos['ticket_troya']:,.2f}

📅 RITMO DEL MES:
├─ Días transcurridos: {datos['dias_transcurridos']}
├─ Días restantes: {datos['dias_restantes']}
└─ Venta x día: S/. {datos['venta_promedio_diaria']:,.2f}

🚀 PROYECCIÓN AL CIERRE:
├─ Ventas: S/. {datos['proyeccion_ventas']:,.2f} ({datos['cumplimiento_proyectado']:.1f}%) {proyectado_emoji}
└─ Cobertura: {datos['proyeccion_cobertura']} clientes ({cumpl_cobertura_proy:.1f}%) {cobertura_emoji}

💪 ¡Sigue adelante!
🤖 Bot N&J Distribuciones"""

# ============================================================
# 9. REPORTE JEFE (EXISTENTE)
# ============================================================

def obtener_datos_generales():
    """Extrae datos consolidados ARCOR para el período actual"""
    conn = None
    try:
        ctx = obtener_contexto_periodo()
        periodo = ctx["periodo"]
        anio = ctx["anio"]
        mes = ctx["mes"]
        dias = calcular_dias_laborables_periodo()

        conn = get_db_connection()
        cursor = conn.cursor()
        datos = {
            "periodo": periodo,
            "nombre_mes": ctx["nombre_mes"],
            "dias_transcurridos": dias["dias_transcurridos"],
            "dias_restantes": dias["dias_restantes"],
            "dias_laborables_total": dias["dias_laborables_total"],
        }

        cursor.execute("""
            SELECT ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS total_ventas,
                   COUNT(DISTINCT Cod_Clie) AS clientes_totales
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ?
        """, (periodo,))
        venta_data = cursor.fetchone()
        datos["total_ventas"] = venta_data["total_ventas"] or 0
        datos["cobertura"] = venta_data["clientes_totales"] or 0

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Cod_Clie), 0), 2), 0) AS ticket
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ?
              AND CAST(Imp_Total AS REAL) > 0
        """, (periodo,))
        ticket = cursor.fetchone()
        datos["ticket_promedio"] = ticket["ticket"] or 0

        cursor.execute("""
            SELECT ROUND(COALESCE(SUM(Cuota_Soles), 0), 2) AS cuota_ventas,
                   ROUND(COALESCE(SUM(Cuota_Cobertura), 0), 0) AS cuota_cobertura
            FROM cuotas
            WHERE AÑO = ? AND NRO_MES = ? AND Proveedor = 'ARCOR'
        """, (anio, mes))
        cuota_data = cursor.fetchone()
        datos["cuota_ventas"] = float(cuota_data["cuota_ventas"] or 0) if cuota_data else 0
        datos["cuota_cobertura"] = int(cuota_data["cuota_cobertura"] or 0) if cuota_data else 0
        datos["cumplimiento"] = (datos["total_ventas"] / datos["cuota_ventas"] * 100 if datos["cuota_ventas"] > 0 else 0)

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)), 2), 0) AS ventas_troya
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ? AND Calif = 'D'
        """, (periodo,))
        troya_venta = cursor.fetchone()
        datos["ventas_troya"] = troya_venta["ventas_troya"] or 0

        # Líneas de negocio EXCLUYENDO SAYON de cada línea a la que pertenece.
        # SAYON se muestra como una línea independiente para no mezclar su venta
        # con la línea de negocio original del producto.
        cursor.execute("""
            SELECT lin_neg,
                   ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS ventas_linea
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ?
            GROUP BY lin_neg
            ORDER BY ventas_linea DESC
        """, (periodo,))
        lineas = cursor.fetchall()

        cursor.execute("""
            SELECT lin_neg,
                   ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS ventas_sayon_linea
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND Periodo = ?
              AND UPPER(COALESCE(Producto, '')) LIKE '%SAYON%'
            GROUP BY lin_neg
        """, (periodo,))
        sayon_por_linea = {
            (row["lin_neg"] if row["lin_neg"] else "SIN LÍNEA"): float(row["ventas_sayon_linea"] or 0)
            for row in cursor.fetchall()
        }

        datos["lineas_negocio"] = {}
        for row in lineas:
            linea = row["lin_neg"] if row["lin_neg"] else "SIN LÍNEA"
            venta_linea = float(row["ventas_linea"] or 0)
            venta_sayon_linea = sayon_por_linea.get(linea, 0)
            datos["lineas_negocio"][linea] = round(max(0, venta_linea - venta_sayon_linea), 2)

        # Venta SAYON independiente: todos los productos cuyo nombre contiene SAYON.
        cursor.execute("""
            SELECT ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS ventas_sayon
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND Periodo = ?
              AND UPPER(COALESCE(Producto, '')) LIKE '%SAYON%'
        """, (periodo,))
        sayon_data = cursor.fetchone()
        datos["ventas_sayon"] = float(sayon_data["ventas_sayon"] or 0)

        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS clientes_troya
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ? AND Calif = 'D'
        """, (periodo,))
        troya_comp = cursor.fetchone()
        datos["clientes_troya_compraron"] = troya_comp["clientes_troya"] or 0

        cursor.execute("SELECT COUNT(DISTINCT Cod_Clie) AS total_clientes_d FROM clientes WHERE Calif = 'D'")
        total_d = cursor.fetchone()
        total_d_clientes = total_d["total_clientes_d"] or 0
        datos["clientes_troya_no_compraron"] = max(0, total_d_clientes - datos["clientes_troya_compraron"])

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Cod_Clie), 0), 2), 0) AS ticket_troya
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ? AND Calif = 'D'
              AND CAST(Imp_Total AS REAL) > 0
        """, (periodo,))
        ticket_troya = cursor.fetchone()
        datos["ticket_troya"] = ticket_troya["ticket_troya"] or 0

        dias_transcurridos = max(dias["dias_transcurridos"], 1)
        dias_restantes = max(dias["dias_restantes"], 0)
        promedio_diario = datos["total_ventas"] / dias_transcurridos
        datos["venta_promedio_diaria"] = round(promedio_diario, 2)

        # Proyección SAYON independiente, usando el mismo ritmo diario del mes.
        promedio_diario_sayon = datos["ventas_sayon"] / dias_transcurridos
        datos["proyeccion_sayon"] = round(
            datos["ventas_sayon"] + promedio_diario_sayon * dias_restantes, 2
        )

        # Proyección ARCOR neta de SAYON.
        datos["proyeccion_ventas"] = round(
            datos["total_ventas"] + promedio_diario * dias_restantes, 2
        )
        datos["proyeccion_arcor_sin_sayon"] = round(
            max(0, datos["proyeccion_ventas"] - datos["proyeccion_sayon"]), 2
        )
        datos["cumplimiento_ventas_proyectado"] = (
            datos["proyeccion_arcor_sin_sayon"] / datos["cuota_ventas"] * 100
            if datos["cuota_ventas"] > 0 else 0
        )

        promedio_cobertura = datos["cobertura"] / dias_transcurridos
        datos["proyeccion_cobertura"] = int(round(datos["cobertura"] + promedio_cobertura * dias_restantes))
        datos["cumplimiento_cobertura_proyectado"] = (datos["proyeccion_cobertura"] / datos["cuota_cobertura"] * 100 if datos["cuota_cobertura"] > 0 else 0)

        logger.info("✅ Datos generales obtenidos")
        return datos
    except Exception as e:
        logger.exception(f"❌ Error obteniendo datos generales: {e}")
        return None
    finally:
        if conn:
            conn.close()

def generar_mensaje_jefe(datos):
    """Genera reporte consolidado para jefe/supervisor"""
    if not datos:
        return None

    cumpl = datos["cumplimiento"]
    cumpl_cobertura_proy = datos["cumplimiento_cobertura_proyectado"]
    proyectado_ventas_emoji = ("🟢" if datos["cumplimiento_ventas_proyectado"] >= 90 else "🟡" if datos["cumplimiento_ventas_proyectado"] >= 75 else "🔴")
    cobertura_emoji = ("🟢" if cumpl_cobertura_proy >= 90 else "🟡" if cumpl_cobertura_proy >= 75 else "🔴")
    cumplimiento_emoji = ("🟢" if cumpl >= 90 else "🟡" if cumpl >= 75 else "🔴")

    lineas_txt = ""
    for linea, ventas in sorted(datos["lineas_negocio"].items(), key=lambda x: x[1], reverse=True):
        # MATERIAL POP deja de mostrarse en el reporte; SAYON ocupa su lugar
        # como categoría independiente.
        if linea.strip().upper() == "MATERIAL POP":
            continue
        pct = ventas / datos["total_ventas"] * 100 if datos["total_ventas"] > 0 else 0
        lineas_txt += f"  • {linea}: S/. {ventas:,.0f} ({pct:.1f}%)\n"

    # SAYON se presenta separado de la línea de negocio original.
    pct_sayon = datos["ventas_sayon"] / datos["total_ventas"] * 100 if datos["total_ventas"] > 0 else 0
    lineas_txt += f"  • SAYON: S/. {datos['ventas_sayon']:,.0f} ({pct_sayon:.1f}%)\n"

    return f"""📊 REPORTE ARCOR - {datos['nombre_mes']}
{ahora_local().strftime('%d/%m/%Y %H:%M')}

🎯 OBJETIVOS MES:
├─ Ventas: S/. {datos['cuota_ventas']:,.2f}
└─ Cobertura: {datos['cuota_cobertura']} clientes

💼 DESEMPEÑO:
├─ Ventas: S/. {datos['total_ventas']:,.2f}
├─ Avance: {cumpl:.1f}% {cumplimiento_emoji}
├─ Cobertura: {datos['cobertura']} clientes
├─ Ticket: S/. {datos['ticket_promedio']:,.2f}
├─ ARCOR: S/. {datos['ventas_arcor_neto']:,.2f}\n├─ SAYON: S/. {datos['ventas_sayon']:,.2f}\n└─ TROYA: S/. {datos['ventas_troya']:,.2f}

📋 LÍNEAS DE NEGOCIO:
{lineas_txt if lineas_txt else '  • Sin ventas registradas'}
⚠️ RESUMEN TROYA:
├─ Compraron: {datos['clientes_troya_compraron']} clientes
├─ No Compraron: {datos['clientes_troya_no_compraron']} clientes
└─ Ticket: S/. {datos['ticket_troya']:,.2f}

📅 RITMO DEL MES:
├─ Días transcurridos: {datos['dias_transcurridos']}
├─ Días restantes: {datos['dias_restantes']}
└─ Venta x día: S/. {datos['venta_promedio_diaria']:,.2f}

🚀 PROYECCIÓN AL CIERRE:
├─ ARCOR: S/. {datos['proyeccion_arcor_sin_sayon']:,.2f} ({datos['cumplimiento_ventas_proyectado']:.1f}%) {proyectado_ventas_emoji}
├─ SAYON: S/. {datos['proyeccion_sayon']:,.2f}
└─ Cobertura: {datos['proyeccion_cobertura']} clientes ({cumpl_cobertura_proy:.1f}%) {cobertura_emoji}

🤖 Bot N&J Distribuciones"""

# ============================================================
# 10. WHATSAPP
# ============================================================

def enviar_mensaje_whatsapp(numero_destino, mensaje):
    """Envía un mensaje de texto por WhatsApp Cloud API"""
    try:
        numero = normalizar_telefono(numero_destino)
        if not numero.startswith("51"):
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

        logger.info(f"📤 Enviando mensaje a *{numero[-4:]}")

        response = requests.post(API_URL, json=payload, headers=headers, timeout=15)

        if response.ok:
            logger.info(f"✅ Mensaje enviado a *{numero[-4:]}")
            return True

        logger.error(f"❌ Error API ({response.status_code})")
        return False
    except Exception as e:
        logger.error(f"❌ Error enviando WhatsApp: {e}")
        return False

# ============================================================
# 11. WEBHOOK
# ============================================================

@app.route("/webhook", methods=["GET"])
def verificar_webhook():
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")
    if verify_token == VERIFY_TOKEN:
        logger.info("✅ Webhook verificado")
        return challenge or "", 200
    logger.warning("⚠️ Token incorrecto")
    return "Unauthorized", 403

@app.route("/webhook", methods=["POST"])
def recibir_mensaje():
    """Recibe, deduplica y encola el mensaje; responde 200 inmediatamente."""
    try:
        data = request.get_json(silent=True) or {}
        if "entry" not in data or not data["entry"]:
            return jsonify({"status": "ok"}), 200
        entrada = data["entry"][0]
        cambios = entrada.get("changes") or []
        if not cambios:
            return jsonify({"status": "ok"}), 200
        cambio = cambios[0].get("value") or {}
        mensajes = cambio.get("messages") or []
        if not mensajes:
            return jsonify({"status": "ok"}), 200
        mensaje_obj = mensajes[0]
        message_id = mensaje_obj.get("id")
        numero_remitente = mensaje_obj.get("from", "")
        if mensaje_obj.get("type") != "text":
            return jsonify({"status": "ok"}), 200
        texto_mensaje = (mensaje_obj.get("text", {}).get("body", "").strip()).upper()
        logger.info(f"📨 Mensaje de *{numero_remitente[-4:]}: {texto_mensaje[:50]}")
        palabra_detectada = (
            "CONCURSO" if "CONCURSO" in texto_mensaje
            else ("RESUMEN" if "RESUMEN" in texto_mensaje
                  else ("TROYA" if "TROYA" in texto_mensaje else None))
        )
        if not palabra_detectada:
            logger.info("⊘ Sin palabra clave")
            return jsonify({"status": "ok"}), 200
        reclamado = reclamar_message_id(message_id, numero_remitente)
        if reclamado is False:
            logger.info(f"♻️ Webhook duplicado ignorado: {message_id}")
            return jsonify({"status": "duplicate"}), 200
        if reclamado is None:
            logger.error(f"⛔ No se procesa {message_id}: no pudo registrarse en SQLite")
            return jsonify({"status": "db_busy"}), 200
        WEBHOOK_EXECUTOR.submit(procesar_mensaje_en_segundo_plano, {"message_id": message_id, "numero_remitente": numero_remitente, "palabra_detectada": palabra_detectada})
        return jsonify({"status": "accepted", "message_id": message_id, "queued": True}), 200
    except Exception as e:
        logger.exception(f"❌ Error recibiendo webhook: {e}")
        return jsonify({"status": "error"}), 200


def procesar_mensaje_en_segundo_plano(job):
    message_id = job.get("message_id")
    numero_remitente = job.get("numero_remitente", "")
    palabra_detectada = job.get("palabra_detectada")
    try:
        logger.info(f"⚙️ Procesando en segundo plano message_id={message_id} palabra={palabra_detectada}")
        es_valido, datos_usuario = validar_vendedor(numero_remitente)
        if not es_valido:
            enviar_mensaje_whatsapp(numero_remitente, "❌ No autorizado. Contacta a tu supervisor.")
            return
        nombre_usuario = datos_usuario["nombre"]
        rol = datos_usuario.get("rol", "")
        if palabra_detectada == "CONCURSO":
            logger.info(f"☢️ Vendedor solicita CHERNOBYL: {nombre_usuario}")
            datos_concurso = obtener_datos_concurso_chenobyl(nombre_usuario)
            mensaje = generar_mensaje_concurso_chenobyl(datos_concurso)
            if enviar_mensaje_whatsapp(numero_remitente, mensaje):
                logger.info(f"✅ Reporte CHERNOBYL enviado a {nombre_usuario} (message_id={message_id})")
            else:
                logger.error(f"❌ Error enviando CHERNOBYL a {nombre_usuario}; message_id conservado: {message_id}")
            return
        if palabra_detectada == "TROYA":
            if es_jefe(rol):
                logger.info(f"👔 Jefe/Supervisor solicita TROYA: {nombre_usuario}")
                clientes_troya = obtener_clientes_troya_generales()
                mensaje = generar_mensaje_troya_jefe(clientes_troya)
            else:
                logger.info(f"👤 Vendedor solicita TROYA: {nombre_usuario}")
                clientes_troya = obtener_clientes_troya(nombre_usuario)
                mensaje = generar_mensaje_troya(datos_usuario, clientes_troya)
            if not mensaje:
                mensaje = "⚠️ Error generando reporte TROYA"
            if enviar_mensaje_whatsapp(numero_remitente, mensaje):
                logger.info(f"✅ Reporte TROYA enviado a {nombre_usuario} (message_id={message_id})")
            else:
                logger.error(f"❌ Error enviando TROYA a {nombre_usuario}; message_id conservado: {message_id}")
            return
        if es_jefe(rol):
            logger.info(f"👔 Jefe/Supervisor: {nombre_usuario}")
            datos = obtener_datos_generales()
            mensaje = generar_mensaje_jefe(datos)
        else:
            logger.info(f"👤 Vendedor: {nombre_usuario}")
            datos = obtener_datos_vendedor(nombre_usuario)
            mensaje = generar_mensaje_vendedor(datos)
        if not datos or not mensaje:
            mensaje = "⚠️ No hay datos disponibles"
        if enviar_mensaje_whatsapp(numero_remitente, mensaje):
            logger.info(f"✅ Reporte enviado a {nombre_usuario} (message_id={message_id})")
        else:
            logger.error(f"❌ Error enviando RESUMEN a {nombre_usuario}; message_id conservado: {message_id}")
    except Exception as e:
        logger.exception(f"❌ Error procesando message_id={message_id}: {e}")


WEBHOOK_EXECUTOR = ThreadPoolExecutor(max_workers=1, thread_name_prefix="whatsapp-worker")

@app.route("/status", methods=["GET"])
def status():
    ctx = obtener_contexto_periodo()
    dias = calcular_dias_laborables_periodo()
    return jsonify({
        "status": "active",
        "timestamp": ctx["ahora"].isoformat(),
        "webhook": "/webhook",
        "version": "V4 ANTI-DUPLICADOS - WORKER",
        "periodo": ctx["periodo"],
        "palabras_clave": ["RESUMEN", "TROYA", "CONCURSO"],
    }), 200

# ============================================================
# 12. MAIN
# ============================================================

if __name__ == "__main__":
    logger.info("=" * 70)
    logger.info("🚀 WEBHOOK V4 - ANTI-DUPLICADOS + WORKER")
    logger.info("=" * 70)
    ctx = obtener_contexto_periodo()
    dias = calcular_dias_laborables_periodo()
    logger.info(f"Período: {ctx['periodo']} - {ctx['nombre_mes']}")
    logger.info(f"Días laborables: {dias}")
    logger.info(f"BD: {BD_PATH}")
    logger.info(f"Excel: {EXCEL_VENDEDORES}")
    logger.info("Palabras clave: RESUMEN | TROYA | CONCURSO")
    logger.info("QUERIES TROYA: CAST AS REAL | Cdg_Vend | ARCOR | WORKER ÚNICO")
    logger.info("=" * 70)

    inicializar_bd()

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
