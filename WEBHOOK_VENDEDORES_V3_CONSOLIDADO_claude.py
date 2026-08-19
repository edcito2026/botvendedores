#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WEBHOOK RECEPCIÓN DE MENSAJES WHATSAPP - V3 CONSOLIDADO
✨ QUERIES TROYA CORREGIDAS (CAST AS REAL, Cdg_Vend, ARCOR)

✨ NUEVAS CARACTERÍSTICAS:
- Palabras clave: "RESUMEN" = reporte general, "TROYA" = reporte clientes TROYA
- Reportes personalizados según rol: VENDEDOR vs JEFE/SUPERVISOR
- Reporte TROYA: lista clientes Calif=D con split CON/SIN COMPRA
- Gestión de créditos optimizada: UN SOLO servicio Render
- QUERIES TROYA CORRECTAS: JOIN por Cod_Clie y Cdg_Vend (CAST AS REAL)

Variables de entorno:
WHATSAPP_ACCESS_TOKEN
WHATSAPP_PHONE_NUMBER_ID
WHATSAPP_VERIFY_TOKEN
WHATSAPP_API_VERSION (default v25.0)
BD_PATH (default ventas.db)
EXCEL_VENDEDORES (default vendedores.xlsx)
DIAS_LABORABLES (default 0,1,2,3,4,5)
"""

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
    conn = sqlite3.connect(BD_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
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

def reclamar_message_id(message_id, telefono):
    inicializar_bd()
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.execute("""
            INSERT OR IGNORE INTO webhook_procesados (message_id, telefono, recibido_en)
            VALUES (?, ?, ?)
        """, (message_id, telefono, ahora_local().isoformat()))
        conn.commit()
        return cursor.rowcount == 1
    except Exception as e:
        logger.error(f"❌ Error registrando message_id: {e}")
        return True
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
    """Obtiene clientes TROYA (Calif='D') con JOIN CORRECTO: CAST AS REAL"""
    try:
        logger.info(f"🔍 Buscando TROYA para: '{nombre_vendedor}'")

        conn = sqlite3.connect(BD_PATH, timeout=15)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        ahora = ahora_local()
        periodo_actual = f"{ahora.year}{ahora.month:02d}"
        mes_anterior = ahora.month - 1 if ahora.month > 1 else 12
        anio_anterior = ahora.year if ahora.month > 1 else ahora.year - 1
        periodo_anterior = f"{anio_anterior}{mes_anterior:02d}"

        # Query CORRECTA: JOIN por Cod_Clie y Cdg_Vend (CAST AS REAL)
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
          AND c.Vendedor = ?
        GROUP BY c.Cod_Clie, c.Raz_Social, c.Vendedor, c.DV, c.Giro
        ORDER BY venta_actual DESC, c.Raz_Social
        """

        cursor.execute(query, (periodo_actual, periodo_anterior, nombre_vendedor))
        clientes = [dict(row) for row in cursor.fetchall()]

        # Agregar flag tiene_compras
        for cliente in clientes:
            cliente['tiene_compras'] = cliente['venta_actual'] > 0

        logger.info(f"📊 Encontrados: {len(clientes)} clientes TROYA")
        conn.close()
        return clientes

    except Exception as e:
        logger.error(f"❌ Error obteniendo clientes TROYA: {e}")
        return []


def obtener_clientes_troya_generales():
    """Obtiene clientes TROYA de TODOS los vendedores - QUERY CORRECTA"""
    try:
        logger.info("🔍 Buscando TROYA GENERAL para todos los vendedores")

        conn = sqlite3.connect(BD_PATH, timeout=15)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()

        ahora = ahora_local()
        periodo_actual = f"{ahora.year}{ahora.month:02d}"

        # Query CORRECTA: JOIN por Cod_Clie y Cdg_Vend (CAST AS REAL)
        query = """
        SELECT
          c.Cod_Clie,
          c.Raz_Social,
          c.Vendedor,
          c.DV,
          COALESCE(ROUND(SUM(v.Imp_Total), 2), 0) as venta_actual
        FROM clientes c
        LEFT JOIN VENTAS2026 v ON CAST(c.Cod_Clie AS REAL) = CAST(v.Cod_Clie AS REAL)
          AND CAST(c.Cdg_Vend AS REAL) = CAST(v.Cdg_Vend AS REAL)
          AND v.Periodo = ?
          AND v.Proveedor = 'ARCOR'
        WHERE c.Calif = 'D'
        GROUP BY c.Cod_Clie, c.Raz_Social, c.Vendedor, c.DV
        ORDER BY c.Vendedor, venta_actual DESC, c.Raz_Social
        """

        cursor.execute(query, (periodo_actual,))
        clientes = [dict(row) for row in cursor.fetchall()]

        # Agregar flag tiene_compras
        for cliente in clientes:
            cliente['tiene_compras'] = cliente['venta_actual'] > 0

        logger.info(f"📊 Encontrados: {len(clientes)} clientes TROYA TOTAL")
        conn.close()
        return clientes

    except Exception as e:
        logger.error(f"❌ Error obteniendo clientes TROYA generales: {e}")
        return []

# ============================================================
# 7. GENERADOR DE MENSAJES TROYA
# ============================================================

def generar_mensaje_troya(vendedor, clientes):
    """Genera reporte personalizado de clientes TROYA (formato compacto)"""

    if not clientes:
        return f"""Hola {vendedor.get('nombre', 'Vendedor')}, 👋

No tienes clientes TROYA (Calif=D) registrados actualmente.

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
    total_anterior = sum(c.get('venta_anterior', 0) for c in clientes)

    mensaje = f"""👤 {nombre}
📅 {mes_nombre} 2026 | {cantidad} TROYA (✅{con_compra} ❌{sin_compra})

✅ CON COMPRA
"""

    # Clientes CON COMPRA con headers
    con_compra_list = [c for c in clientes if c.get('tiene_compras')]
    if con_compra_list:
        mensaje += f"{mes_anterior_nombre:<8} {mes_nombre:<8}\n"
        for cliente in con_compra_list:
            cliente_nombre = cliente['Raz_Social'].strip().title()[:20]
            venta_ant = cliente.get('venta_anterior', 0)
            venta_act = cliente.get('venta_actual', 0)
            mensaje += f"S/.{venta_ant:<4.0f} S/.{venta_act:<4.0f} {cliente_nombre}\n"

    # Clientes SIN COMPRA
    sin_compra_list = [c for c in clientes if not c.get('tiene_compras')]
    if sin_compra_list:
        mensaje += f"\n❌ SIN COMPRA\n"
        for cliente in sin_compra_list:
            cliente_nombre = cliente['Raz_Social'].strip().title()[:16]
            dia = cliente.get('DV', 'N/A')
            venta_ant = cliente.get('venta_anterior', 0)
            if venta_ant > 0:
                mensaje += f"{cliente_nombre} {dia} S/.{venta_ant:.0f}\n"
            else:
                mensaje += f"{cliente_nombre} {dia}\n"

    # Resumen
    diferencia = total_actual - total_anterior
    pct = (diferencia / total_anterior * 100) if total_anterior > 0 else 0

    mensaje += f"\n📊 Totales: S/.{total_anterior:.0f} → S/.{total_actual:.0f}"
    if pct != 0:
        emoji = "📈" if pct > 0 else "📉"
        mensaje += f" {emoji} {pct:+.0f}%"

    mensaje += "\n\n🤖 Bot N&J"

    return mensaje


def generar_mensaje_troya_jefe(clientes):
    """Genera reporte consolidado de TROYA para jefe/supervisor"""

    if not clientes:
        return """👨‍💼 REPORTE TROYA CONSOLIDADO

No hay clientes TROYA (Calif=D) registrados.

🤖 Bot N&J"""

    ahora = ahora_local()
    mes_nombre = MESES_ES[ahora.month - 1]

    total_clientes = len(clientes)
    con_compra_total = sum(1 for c in clientes if c.get('tiene_compras'))
    sin_compra_total = total_clientes - con_compra_total

    # Agrupar por vendedor
    por_vendedor = {}
    for cliente in clientes:
        vendedor = cliente.get('Vendedor', 'SIN VENDEDOR')
        if vendedor not in por_vendedor:
            por_vendedor[vendedor] = {'total': 0, 'con_compra': 0, 'sin_compra': 0}
        por_vendedor[vendedor]['total'] += 1
        if cliente.get('tiene_compras'):
            por_vendedor[vendedor]['con_compra'] += 1
        else:
            por_vendedor[vendedor]['sin_compra'] += 1

    pct_con_compra = (con_compra_total / total_clientes * 100) if total_clientes > 0 else 0

    mensaje = f"""👨‍💼 REPORTE TROYA CONSOLIDADO - {mes_nombre}
{ahora_local().strftime('%d/%m/%Y %H:%M')}

📊 RESUMEN GENERAL:
├─ Total Clientes TROYA: {total_clientes}
├─ ✅ Con Compra: {con_compra_total} ({pct_con_compra:.1f}%)
└─ ❌ Sin Compra: {sin_compra_total}

📈 POR VENDEDOR:
"""

    for vendedor in sorted(por_vendedor.keys()):
        datos = por_vendedor[vendedor]
        pct = (datos['con_compra'] / datos['total'] * 100) if datos['total'] > 0 else 0
        mensaje += f"\n{vendedor}\n"
        mensaje += f"  Total: {datos['total']} | ✅ {datos['con_compra']} ({pct:.0f}%) | ❌ {datos['sin_compra']}"

    mensaje += """

⚡ ACCIONES REQUERIDAS:
• Todos deben reactivar clientes SIN COMPRA
• Mantener activos los clientes CON COMPRA
• Enviar promociones especiales

🤖 Bot N&J"""

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
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = 'ARCOR'
        """, (nombre_vendedor, periodo))

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
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Documento), 0), 2), 0) AS ticket
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = 'ARCOR'
        """, (nombre_vendedor, periodo))
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
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = 'ARCOR' AND Calif = 'D'
        """, (nombre_vendedor, periodo))
        troya_data = cursor.fetchone()
        datos["ventas_troya"] = troya_data["troya"] or 0

        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS clientes_troya
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = 'ARCOR' AND Calif = 'D'
        """, (nombre_vendedor, periodo))
        troya_comp = cursor.fetchone()
        datos["clientes_troya_compraron"] = troya_comp["clientes_troya"] or 0

        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS total_clientes_d
            FROM clientes
            WHERE Vendedor = ? AND Calif = 'D'
        """, (nombre_vendedor,))
        total_d = cursor.fetchone()
        total_d_clientes = total_d["total_clientes_d"] or 0
        datos["clientes_troya_no_compraron"] = max(0, total_d_clientes - datos["clientes_troya_compraron"])

        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Documento), 0), 2), 0) AS ticket_troya
            FROM VENTAS2026
            WHERE Vendedor = ? AND Periodo = ? AND Proveedor = 'ARCOR' AND Calif = 'D'
        """, (nombre_vendedor, periodo))
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
├─ Cuota Ventas: S/. {datos['cuota']:,.2f}
└─ Cuota Cobertura: {datos['cuota_cobertura']} clientes

💼 DESEMPEÑO:
├─ Ventas Actuales: S/. {datos['total_ventas']:,.2f}
├─ Cumplimiento: {cumpl:.1f}% {cumplimiento_emoji}
├─ Cobertura: {datos['clientes']} clientes
├─ Ticket Promedio: S/. {datos['ticket_promedio']:,.2f}
└─ Ventas TROYA: S/. {datos['ventas_troya']:,.2f}

⚠️ TROYA RESUMEN:
├─ Compraron: {datos['clientes_troya_compraron']} clientes
├─ No Compraron: {datos['clientes_troya_no_compraron']} clientes
└─ Ticket TROYA: S/. {datos['ticket_troya']:,.2f}

📅 RITMO DEL MES:
├─ Días laborables transcurridos: {datos['dias_transcurridos']}
├─ Días laborables restantes: {datos['dias_restantes']}
└─ Venta promedio/día: S/. {datos['venta_promedio_diaria']:,.2f}

🚀 PROYECCIÓN AL CIERRE:
├─ Ventas: S/. {datos['proyeccion_ventas']:,.2f} ({datos['cumplimiento_proyectado']:.1f}%) {proyectado_emoji}
└─ Cobertura: {datos['proyeccion_cobertura']} clientes ({cumpl_cobertura_proy:.1f}%) {cobertura_emoji}

💪 ¡Sigue adelante!
Sistema Automatizado N&J"""

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
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Documento), 0), 2), 0) AS ticket
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ?
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

        cursor.execute("""
            SELECT lin_neg, ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS ventas_linea
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ?
            GROUP BY lin_neg
            ORDER BY ventas_linea DESC
        """, (periodo,))
        lineas = cursor.fetchall()
        datos["lineas_negocio"] = {(row["lin_neg"] if row["lin_neg"] else "SIN LÍNEA"): row["ventas_linea"] for row in lineas}

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
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)) / NULLIF(COUNT(DISTINCT Documento), 0), 2), 0) AS ticket_troya
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR' AND Periodo = ? AND Calif = 'D'
        """, (periodo,))
        ticket_troya = cursor.fetchone()
        datos["ticket_troya"] = ticket_troya["ticket_troya"] or 0

        dias_transcurridos = max(dias["dias_transcurridos"], 1)
        dias_restantes = max(dias["dias_restantes"], 0)
        promedio_diario = datos["total_ventas"] / dias_transcurridos
        datos["venta_promedio_diaria"] = round(promedio_diario, 2)
        datos["proyeccion_ventas"] = round(datos["total_ventas"] + promedio_diario * dias_restantes, 2)
        datos["cumplimiento_ventas_proyectado"] = (datos["proyeccion_ventas"] / datos["cuota_ventas"] * 100 if datos["cuota_ventas"] > 0 else 0)

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
        pct = ventas / datos["total_ventas"] * 100 if datos["total_ventas"] > 0 else 0
        lineas_txt += f"  • {linea}: S/. {ventas:,.0f} ({pct:.1f}%)\n"

    return f"""📊 REPORTE ARCOR - {datos['nombre_mes']}
{ahora_local().strftime('%d/%m/%Y %H:%M')}

🎯 OBJETIVOS MES:
├─ Cuota Ventas: S/. {datos['cuota_ventas']:,.2f}
└─ Cuota Cobertura: {datos['cuota_cobertura']} clientes

💼 DESEMPEÑO:
├─ Ventas Actuales: S/. {datos['total_ventas']:,.2f}
├─ Cumplimiento: {cumpl:.1f}% {cumplimiento_emoji}
├─ Cobertura: {datos['cobertura']} clientes
├─ Ticket Promedio: S/. {datos['ticket_promedio']:,.2f}
└─ Ventas TROYA: S/. {datos['ventas_troya']:,.2f}

📋 LÍNEAS DE NEGOCIO:
{lineas_txt if lineas_txt else '  • Sin ventas registradas'}
⚠️ RESUMEN TROYA:
├─ Compraron: {datos['clientes_troya_compraron']} clientes
├─ No Compraron: {datos['clientes_troya_no_compraron']} clientes
└─ Ticket Promedio: S/. {datos['ticket_troya']:,.2f}

📅 RITMO DEL MES:
├─ Días laborables transcurridos: {datos['dias_transcurridos']}
├─ Días laborables restantes: {datos['dias_restantes']}
└─ Venta promedio/día: S/. {datos['venta_promedio_diaria']:,.2f}

🚀 PROYECCIÓN AL CIERRE:
├─ Ventas: S/. {datos['proyeccion_ventas']:,.2f} ({datos['cumplimiento_ventas_proyectado']:.1f}%) {proyectado_ventas_emoji}
└─ Cobertura: {datos['proyeccion_cobertura']} clientes ({cumpl_cobertura_proy:.1f}%) {cobertura_emoji}

📞 Equipo de Ventas N&J"""

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
    """Recibe mensajes y responde según palabra clave"""
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
        tipo = mensaje_obj.get("type")

        if tipo != "text":
            return jsonify({"status": "ok"}), 200

        texto_mensaje = (mensaje_obj.get("text", {}).get("body", "").strip()).upper()
        logger.info(f"📨 Mensaje de *{numero_remitente[-4:]}: {texto_mensaje[:50]}")

        # Detectar palabra clave: RESUMEN o TROYA
        palabra_detectada = None
        if "RESUMEN" in texto_mensaje:
            palabra_detectada = "RESUMEN"
        elif "TROYA" in texto_mensaje:
            palabra_detectada = "TROYA"

        if not palabra_detectada:
            logger.info("⊘ Sin palabra clave")
            return jsonify({"status": "ok"}), 200

        # Evitar duplicados
        if message_id and not reclamar_message_id(message_id, numero_remitente):
            logger.info(f"♻️ Webhook duplicado")
            return jsonify({"status": "duplicate"}), 200

        logger.info(f"✅ Palabra clave: {palabra_detectada}")

        # Validar vendedor
        es_valido, datos_usuario = validar_vendedor(numero_remitente)

        if not es_valido:
            enviar_mensaje_whatsapp(numero_remitente, "❌ No autorizado. Contacta a tu supervisor.")
            return jsonify({"status": "unauthorized"}), 200

        nombre_usuario = datos_usuario["nombre"]
        codigo_usuario = datos_usuario.get("codigo", nombre_usuario)
        rol = datos_usuario.get("rol", "")

        # ========== PALABRA CLAVE: TROYA ==========
        if palabra_detectada == "TROYA":
            es_jefe_supervisor = es_jefe(rol)

            if es_jefe_supervisor:
                logger.info(f"👔 Jefe/Supervisor solicita TROYA: {nombre_usuario}")
                clientes_troya = obtener_clientes_troya_generales()
                mensaje = generar_mensaje_troya_jefe(clientes_troya)
            else:
                logger.info(f"👤 Vendedor solicita TROYA: {nombre_usuario}")
                clientes_troya = obtener_clientes_troya(nombre_usuario)
                mensaje = generar_mensaje_troya(datos_usuario, clientes_troya)

            if not mensaje:
                mensaje = "⚠️ Error generando reporte TROYA"

            exito = enviar_mensaje_whatsapp(numero_remitente, mensaje)

            if exito:
                logger.info(f"✅ Reporte TROYA enviado a {nombre_usuario}")
                return jsonify({"status": "success"}), 200

            liberar_message_id(message_id)
            return jsonify({"status": "send_error"}), 500

        # ========== PALABRA CLAVE: RESUMEN ==========
        else:
            es_jefe_supervisor = es_jefe(rol)

            if es_jefe_supervisor:
                logger.info(f"👔 Jefe/Supervisor: {nombre_usuario}")
                datos = obtener_datos_generales()
                mensaje = generar_mensaje_jefe(datos)
            else:
                logger.info(f"👤 Vendedor: {nombre_usuario}")
                datos = obtener_datos_vendedor(nombre_usuario)
                mensaje = generar_mensaje_vendedor(datos)

            if not datos or not mensaje:
                mensaje = "⚠️ No hay datos disponibles"
                exito = enviar_mensaje_whatsapp(numero_remitente, mensaje)
                if not exito:
                    liberar_message_id(message_id)
                return jsonify({"status": "no_data"}), 200

            exito = enviar_mensaje_whatsapp(numero_remitente, mensaje)

            if exito:
                logger.info(f"✅ Reporte enviado a {nombre_usuario}")
                return jsonify({"status": "success"}), 200

            liberar_message_id(message_id)
            return jsonify({"status": "send_error"}), 500

    except Exception as e:
        logger.exception(f"❌ Error procesando webhook: {e}")
        return jsonify({"status": "error"}), 500

@app.route("/status", methods=["GET"])
def status():
    ctx = obtener_contexto_periodo()
    dias = calcular_dias_laborables_periodo()
    return jsonify({
        "status": "active",
        "timestamp": ctx["ahora"].isoformat(),
        "webhook": "/webhook",
        "version": "V3 TROYA Correcto",
        "periodo": ctx["periodo"],
        "palabras_clave": ["RESUMEN", "TROYA"],
    }), 200

# ============================================================
# 12. MAIN
# ============================================================

if __name__ == "__main__":
    logger.info("=" * 70)
    logger.info("🚀 WEBHOOK V3 CONSOLIDADO - QUERIES TROYA CORREGIDAS")
    logger.info("=" * 70)
    ctx = obtener_contexto_periodo()
    dias = calcular_dias_laborables_periodo()
    logger.info(f"Período: {ctx['periodo']} - {ctx['nombre_mes']}")
    logger.info(f"Días laborables: {dias}")
    logger.info(f"BD: {BD_PATH}")
    logger.info(f"Excel: {EXCEL_VENDEDORES}")
    logger.info("Palabras clave: RESUMEN | TROYA")
    logger.info("QUERIES TROYA: CAST AS REAL | Cdg_Vend | ARCOR")
    logger.info("=" * 70)

    inicializar_bd()

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
