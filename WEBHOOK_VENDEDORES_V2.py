#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
WEBHOOK RECEPCIÓN DE MENSAJES WHATSAPP - V2 CORREGIDO

Cambios en esta versión:
✅ Normalización de período: CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
✅ Maneja períodos con decimal ('202608.0' → 202608)
- Credenciales por variables de entorno (sin secretos)
- Período y nombre de mes dinámicos
- Proyección basada en días laborables reales
- Protección contra webhooks duplicados
- Conexiones SQLite seguras
- Zona horaria America/Lima
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


# ============================================================
# 1. CONFIGURACIÓN
# ============================================================

BASE_PATH = os.path.dirname(os.path.abspath(__file__))

ACCESS_TOKEN = os.environ.get("WHATSAPP_ACCESS_TOKEN")
PHONE_NUMBER_ID = os.environ.get("WHATSAPP_PHONE_NUMBER_ID")
VERIFY_TOKEN = os.environ.get("WHATSAPP_VERIFY_TOKEN")
API_VERSION = os.environ.get("WHATSAPP_API_VERSION", "v25.0")

if not ACCESS_TOKEN:
    raise RuntimeError("Falta la variable de entorno WHATSAPP_ACCESS_TOKEN")
if not PHONE_NUMBER_ID:
    raise RuntimeError("Falta la variable de entorno WHATSAPP_PHONE_NUMBER_ID")
if not VERIFY_TOKEN:
    raise RuntimeError("Falta la variable de entorno WHATSAPP_VERIFY_TOKEN")

API_URL = f"https://graph.facebook.com/{API_VERSION}/{PHONE_NUMBER_ID}/messages"

PALABRA_CLAVE = os.environ.get("PALABRA_CLAVE", "resumen").strip().lower()
BD_PATH = os.environ.get("BD_PATH", os.path.join(BASE_PATH, "ventas.db"))
EXCEL_VENDEDORES = os.environ.get(
    "EXCEL_VENDEDORES",
    os.path.join(BASE_PATH, "vendedores.xlsx")
)

# 0=lunes ... 6=domingo. Por defecto: lunes-sábado; domingo no laborable.
def cargar_dias_laborables():
    valor = os.environ.get("DIAS_LABORABLES", "0,1,2,3,4,5")
    try:
        dias = {int(x.strip()) for x in valor.split(",") if x.strip() != ""}
        if not dias.issubset(set(range(7))):
            raise ValueError
        return dias
    except ValueError:
        raise RuntimeError(
            "DIAS_LABORABLES debe contener números del 0 al 6. "
            "Ejemplo: 0,1,2,3,4,5"
        )

DIAS_LABORABLES = cargar_dias_laborables()

# Zona horaria comercial de Perú.
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
        logging.FileHandler(
            os.path.join(LOG_DIR, "webhook_v2.log"),
            encoding="utf-8"
        ),
        logging.StreamHandler()
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)


# ============================================================
# 3. UTILIDADES DE FECHA / PERIODO
# ============================================================

MESES_ES = [
    "ENERO", "FEBRERO", "MARZO", "ABRIL", "MAYO", "JUNIO",
    "JULIO", "AGOSTO", "SETIEMBRE", "OCTUBRE", "NOVIEMBRE", "DICIEMBRE"
]


def ahora_local():
    """Fecha/hora actual en Perú."""
    if TZ_LIMA:
        return datetime.now(TZ_LIMA)
    return datetime.now()


def obtener_contexto_periodo():
    """Obtiene periodo, año, mes y fechas del mes actual."""
    ahora = ahora_local()
    hoy = ahora.date()
    anio = hoy.year
    mes = hoy.month
    primer_dia = date(anio, mes, 1)

    if mes == 12:
        primer_dia_siguiente = date(anio + 1, 1, 1)
    else:
        primer_dia_siguiente = date(anio, mes + 1, 1)

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
    """
    Calcula días laborables transcurridos y restantes del mes actual.
    No usa un número fijo como 26.
    Domingo queda excluido por defecto.
    """
    ctx = obtener_contexto_periodo()
    hoy = ctx["hoy"]
    primer_dia = ctx["primer_dia"]
    ultimo_dia = ctx["ultimo_dia"]

    # Si estamos fuera del mes por alguna razón, protegemos los límites.
    fecha_corte = min(max(hoy, primer_dia), ultimo_dia)

    transcurridos = 0
    restantes = 0
    total = 0

    fecha = primer_dia
    while fecha <= ultimo_dia:
        if fecha.weekday() in DIAS_LABORABLES:
            total += 1
            if fecha <= fecha_corte:
                transcurridos += 1
            elif fecha > hoy:
                restantes += 1
        fecha += timedelta(days=1)

    # Si hoy es posterior al mes, no deben quedar días.
    if hoy > ultimo_dia:
        restantes = 0
        transcurridos = total

    # Si hoy es anterior al mes, no deben aparecer días transcurridos.
    if hoy < primer_dia:
        transcurridos = 0
        restantes = total

    return {
        "dias_laborables_total": total,
        "dias_transcurridos": transcurridos,
        "dias_restantes": restantes,
    }


def periodo_sql():
    """Retorna el periodo actual como string YYYYMM."""
    return obtener_contexto_periodo()["periodo"]


# ============================================================
# 4. SQLITE
# ============================================================

_DB_INIT_LOCK = Lock()
_DB_INITIALIZED = False


def get_db_connection():
    """Abre una conexión SQLite con configuración segura para lecturas concurrentes."""
    conn = sqlite3.connect(BD_PATH, timeout=15)
    conn.row_factory = sqlite3.Row
    return conn


def inicializar_bd():
    """Crea la tabla de control de webhooks procesados si no existe."""
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
            logger.info("✅ Control de webhooks duplicados inicializado")
        except Exception as e:
            logger.error(f"❌ Error inicializando BD: {e}")
            raise
        finally:
            if conn:
                conn.close()


def reclamar_message_id(message_id, telefono):
    """Reserva un message_id. False = ya fue procesado/reservado."""
    inicializar_bd()
    conn = None
    try:
        conn = get_db_connection()
        cursor = conn.execute(
            """
            INSERT OR IGNORE INTO webhook_procesados
                (message_id, telefono, recibido_en)
            VALUES (?, ?, ?)
            """,
            (message_id, telefono, ahora_local().isoformat())
        )
        conn.commit()
        return cursor.rowcount == 1
    except Exception as e:
        logger.error(f"❌ Error registrando message_id: {e}")
        # Si falla el control de duplicados, no bloqueamos el webhook.
        return True
    finally:
        if conn:
            conn.close()


def liberar_message_id(message_id):
    """Permite reintentar si el envío falló."""
    if not message_id:
        return
    conn = None
    try:
        conn = get_db_connection()
        conn.execute(
            "DELETE FROM webhook_procesados WHERE message_id = ?",
            (message_id,)
        )
        conn.commit()
    except Exception as e:
        logger.error(f"❌ Error liberando message_id: {e}")
    finally:
        if conn:
            conn.close()


# ============================================================
# 5. VALIDAR VENDEDOR / JEFE
# ============================================================

_EXCEL_CACHE = {}
_EXCEL_LOCK = Lock()


def normalizar_telefono(numero):
    """Normaliza un teléfono dejando solo dígitos."""
    return "".join(filter(str.isdigit, str(numero)))


def obtener_vendedores_autorizados():
    """Lee vendedores autorizados desde Excel y los mantiene en cache."""
    try:
        mtime = os.path.getmtime(EXCEL_VENDEDORES)

        with _EXCEL_LOCK:
            if _EXCEL_CACHE and mtime == _EXCEL_CACHE.get("mtime"):
                return _EXCEL_CACHE.get("data", {})

        logger.info("📋 Leyendo vendedores autorizados desde Excel...")
        vendedores = {}

        wb = openpyxl.load_workbook(
            EXCEL_VENDEDORES,
            read_only=True,
            data_only=True
        )
        ws = wb.active

        # A=Nombre, B=Teléfono, C=Clientes, D=Rol (opcional)
        for row in ws.iter_rows(min_row=1, values_only=True):
            if len(row) < 2 or not row[0] or not row[1]:
                continue

            nombre = str(row[0]).strip()
            telefono = normalizar_telefono(row[1])

            if not telefono:
                continue

            # Preferimos el número completo; conservamos fallback de últimos 9.
            rol = str(row[3]).strip().upper() if len(row) > 3 and row[3] else ""

            vendedores[telefono] = {
                "nombre": nombre,
                "telefonooriginal": telefono,
                "telefono_last9": telefono[-9:] if len(telefono) >= 9 else telefono,
                "clientes": row[2] if len(row) > 2 else None,
                "rol": rol,
            }

        wb.close()

        with _EXCEL_LOCK:
            _EXCEL_CACHE["data"] = vendedores
            _EXCEL_CACHE["mtime"] = mtime

        logger.info(f"✅ {len(vendedores)} usuarios autorizados cargados")
        return vendedores

    except Exception as e:
        logger.error(f"❌ Error leyendo Excel: {e}")
        return {}


def validar_vendedor(numero_telefonico):
    """Valida el teléfono autorizado. Mantiene compatibilidad con el formato anterior."""
    numero_limpio = normalizar_telefono(numero_telefonico)
    vendedores = obtener_vendedores_autorizados()

    if numero_limpio in vendedores:
        usuario = vendedores[numero_limpio]
        logger.info(f"✅ Usuario autorizado: {usuario['nombre']}")
        return True, usuario

    # Compatibilidad con Excel antiguo que pudiera no incluir código de país.
    last9 = numero_limpio[-9:] if len(numero_limpio) >= 9 else numero_limpio
    for usuario in vendedores.values():
        if usuario.get("telefono_last9") == last9:
            logger.info(f"✅ Usuario autorizado por últimos 9 dígitos: {usuario['nombre']}")
            return True, usuario

    logger.warning(f"⚠️ Número no autorizado: {numero_limpio[-4:] if numero_limpio else 'vacío'}")
    return False, None


def es_jefe(nombre_usuario, rol=""):
    """
    Determina el rol.
    Prioridad: columna D del Excel.
    Fallback: mantiene compatibilidad con la lógica anterior.
    """
    rol = (rol or "").strip().upper()

    if rol:
        # Busca palabras clave en el rol (no coincidencia exacta)
        jefe_keywords = ["JEFE", "SUPERVISOR", "GERENTE", "COORDINADOR"]
        return any(keyword in rol for keyword in jefe_keywords)

    # Compatibilidad temporal con Excel antiguo.
    nombre_lower = nombre_usuario.lower()
    jefes_keywords = ["jefe", "supervisor", "gerente", "coordinador", "arcor"]
    return any(keyword in nombre_lower for keyword in jefes_keywords)


# ============================================================
# 6. DATOS DEL VENDEDOR
# ============================================================


def obtener_datos_vendedor(nombre_vendedor):
    """Extrae KPIs del vendedor únicamente para el periodo actual."""
    conn = None
    try:
        ctx = obtener_contexto_periodo()
        periodo = ctx["periodo"]
        periodo_int = int(float(periodo))  # Normaliza '202608' → 202608
        anio = ctx["anio"]
        mes = ctx["mes"]
        dias = calcular_dias_laborables_periodo()

        logger.info(f"🔄 Consultando datos de {nombre_vendedor} - periodo {periodo}...")

        conn = get_db_connection()
        cursor = conn.cursor()

        # 1. Ventas y clientes - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS total_ventas,
                COUNT(DISTINCT Cod_Clie) AS clientes
            FROM VENTAS2026
            WHERE Vendedor = ?
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Proveedor = 'ARCOR'
        """, (nombre_vendedor, periodo_int))

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

        # 2. Ticket promedio general - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                COALESCE(
                    ROUND(
                        SUM(CAST(Imp_Total AS REAL)) /
                        NULLIF(COUNT(DISTINCT Documento), 0),
                        2
                    ), 0
                ) AS ticket
            FROM VENTAS2026
            WHERE Vendedor = ?
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Proveedor = 'ARCOR'
        """, (nombre_vendedor, periodo_int))
        ticket = cursor.fetchone()
        datos["ticket_promedio"] = ticket["ticket"] or 0

        # 3. Cuota
        cursor.execute("""
            SELECT Cuota_Soles, Cuota_Cobertura
            FROM cuotas
            WHERE Vendedor = ?
              AND AÑO = ?
              AND NRO_MES = ?
              AND Proveedor = 'ARCOR'
            LIMIT 1
        """, (nombre_vendedor, anio, mes))

        cuota_data = cursor.fetchone()
        datos["cuota"] = float(cuota_data["Cuota_Soles"] or 0) if cuota_data else 0
        datos["cuota_cobertura"] = int(cuota_data["Cuota_Cobertura"] or 0) if cuota_data else 0

        # 4. Cumplimiento
        datos["cumplimiento"] = (
            datos["total_ventas"] / datos["cuota"] * 100
            if datos["cuota"] > 0 else 0
        )

        # 5. TROYA - ventas en Calif=D - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)), 2), 0) AS troya
            FROM VENTAS2026
            WHERE Vendedor = ?
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Proveedor = 'ARCOR'
              AND Calif = 'D'
        """, (nombre_vendedor, periodo_int))
        troya_data = cursor.fetchone()
        datos["ventas_troya"] = troya_data["troya"] or 0

        # 6. Clientes TROYA que compraron - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS clientes_troya_compraron
            FROM VENTAS2026
            WHERE Vendedor = ?
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Proveedor = 'ARCOR'
              AND Calif = 'D'
        """, (nombre_vendedor, periodo_int))
        troya_comp = cursor.fetchone()
        datos["clientes_troya_compraron"] = troya_comp["clientes_troya_compraron"] or 0

        # 7. Base TROYA del vendedor
        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS total_clientes_d
            FROM clientes
            WHERE Vendedor = ?
              AND Calif = 'D'
        """, (nombre_vendedor,))
        total_d = cursor.fetchone()
        total_d_clientes = total_d["total_clientes_d"] or 0
        datos["clientes_troya_no_compraron"] = max(
            0,
            total_d_clientes - datos["clientes_troya_compraron"]
        )

        # 8. Ticket TROYA - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                COALESCE(
                    ROUND(
                        SUM(CAST(Imp_Total AS REAL)) /
                        NULLIF(COUNT(DISTINCT Documento), 0),
                        2
                    ), 0
                ) AS ticket_troya
            FROM VENTAS2026
            WHERE Vendedor = ?
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Proveedor = 'ARCOR'
              AND Calif = 'D'
        """, (nombre_vendedor, periodo_int))
        ticket_troya = cursor.fetchone()
        datos["ticket_troya"] = ticket_troya["ticket_troya"] or 0

        # 9. Proyección dinámica por días laborables
        dias_transcurridos = max(dias["dias_transcurridos"], 1)
        dias_restantes = max(dias["dias_restantes"], 0)

        promedio_diario = datos["total_ventas"] / dias_transcurridos
        datos["venta_promedio_diaria"] = round(promedio_diario, 2)
        datos["proyeccion_ventas"] = round(
            datos["total_ventas"] + promedio_diario * dias_restantes,
            2
        )
        datos["cumplimiento_proyectado"] = (
            datos["proyeccion_ventas"] / datos["cuota"] * 100
            if datos["cuota"] > 0 else 0
        )

        # 10. Proyección cobertura
        promedio_cobertura = datos["clientes"] / dias_transcurridos
        datos["proyeccion_cobertura"] = int(round(
            datos["clientes"] + promedio_cobertura * dias_restantes
        ))
        datos["cumplimiento_cobertura_proyectado"] = (
            datos["proyeccion_cobertura"] / datos["cuota_cobertura"] * 100
            if datos["cuota_cobertura"] > 0 else 0
        )

        logger.info(f"✅ Datos obtenidos para {nombre_vendedor}")
        return datos

    except Exception as e:
        logger.exception(f"❌ Error obteniendo datos vendedor: {e}")
        return None
    finally:
        if conn:
            conn.close()


# ============================================================
# 7. DATOS GENERALES PARA JEFE / SUPERVISOR
# ============================================================


def obtener_datos_generales():
    """Extrae datos consolidados ARCOR únicamente para el periodo actual."""
    conn = None
    try:
        ctx = obtener_contexto_periodo()
        periodo = ctx["periodo"]
        periodo_int = int(float(periodo))  # Normaliza '202608' → 202608
        anio = ctx["anio"]
        mes = ctx["mes"]
        dias = calcular_dias_laborables_periodo()

        logger.info(f"🔄 Consultando datos generales ARCOR - periodo {periodo}...")

        conn = get_db_connection()
        cursor = conn.cursor()
        datos = {
            "periodo": periodo,
            "nombre_mes": ctx["nombre_mes"],
            "dias_transcurridos": dias["dias_transcurridos"],
            "dias_restantes": dias["dias_restantes"],
            "dias_laborables_total": dias["dias_laborables_total"],
        }

        # 1. Ventas y cobertura ARCOR del mes actual - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS total_ventas,
                COUNT(DISTINCT Cod_Clie) AS clientes_totales
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
        """, (periodo_int,))
        venta_data = cursor.fetchone()
        datos["total_ventas"] = venta_data["total_ventas"] or 0
        datos["cobertura"] = venta_data["clientes_totales"] or 0

        # 2. Ticket promedio general del mes - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                COALESCE(
                    ROUND(
                        SUM(CAST(Imp_Total AS REAL)) /
                        NULLIF(COUNT(DISTINCT Documento), 0),
                        2
                    ), 0
                ) AS ticket
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
        """, (periodo_int,))
        ticket = cursor.fetchone()
        datos["ticket_promedio"] = ticket["ticket"] or 0

        # 3. Cuota total ARCOR del mes
        cursor.execute("""
            SELECT
                ROUND(COALESCE(SUM(Cuota_Soles), 0), 2) AS cuota_ventas,
                ROUND(COALESCE(SUM(Cuota_Cobertura), 0), 0) AS cuota_cobertura
            FROM cuotas
            WHERE AÑO = ?
              AND NRO_MES = ?
              AND Proveedor = 'ARCOR'
        """, (anio, mes))
        cuota_data = cursor.fetchone()
        datos["cuota_ventas"] = float(cuota_data["cuota_ventas"] or 0) if cuota_data else 0
        datos["cuota_cobertura"] = int(cuota_data["cuota_cobertura"] or 0) if cuota_data else 0
        datos["cumplimiento"] = (
            datos["total_ventas"] / datos["cuota_ventas"] * 100
            if datos["cuota_ventas"] > 0 else 0
        )

        # 4. Ventas TROYA del mes - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT COALESCE(ROUND(SUM(CAST(Imp_Total AS REAL)), 2), 0) AS ventas_troya
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Calif = 'D'
        """, (periodo_int,))
        troya_venta = cursor.fetchone()
        datos["ventas_troya"] = troya_venta["ventas_troya"] or 0

        # 5. Líneas de negocio del mes - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                lin_neg,
                ROUND(COALESCE(SUM(CAST(Imp_Total AS REAL)), 0), 2) AS ventas_linea
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
            GROUP BY lin_neg
            ORDER BY ventas_linea DESC
        """, (periodo_int,))
        lineas = cursor.fetchall()
        datos["lineas_negocio"] = {
            row["lin_neg"] if row["lin_neg"] else "SIN LÍNEA": row["ventas_linea"]
            for row in lineas
        }

        # 6. Clientes TROYA que compraron - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS clientes_troya_compraron
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Calif = 'D'
        """, (periodo_int,))
        troya_comp = cursor.fetchone()
        datos["clientes_troya_compraron"] = troya_comp["clientes_troya_compraron"] or 0

        # Base TROYA total
        cursor.execute("""
            SELECT COUNT(DISTINCT Cod_Clie) AS total_clientes_d
            FROM clientes
            WHERE Calif = 'D'
        """)
        total_d = cursor.fetchone()
        total_d_clientes = total_d["total_clientes_d"] or 0
        datos["clientes_troya_no_compraron"] = max(
            0,
            total_d_clientes - datos["clientes_troya_compraron"]
        )

        # 7. Ticket TROYA del mes - NORMALIZAR PERÍODO
        cursor.execute("""
            SELECT
                COALESCE(
                    ROUND(
                        SUM(CAST(Imp_Total AS REAL)) /
                        NULLIF(COUNT(DISTINCT Documento), 0),
                        2
                    ), 0
                ) AS ticket_troya
            FROM VENTAS2026
            WHERE Proveedor = 'ARCOR'
              AND CAST(CAST(Periodo AS FLOAT) AS INTEGER) = ?
              AND Calif = 'D'
        """, (periodo_int,))
        ticket_troya = cursor.fetchone()
        datos["ticket_troya"] = ticket_troya["ticket_troya"] or 0

        # 8. Proyección general dinámica
        dias_transcurridos = max(dias["dias_transcurridos"], 1)
        dias_restantes = max(dias["dias_restantes"], 0)

        promedio_diario = datos["total_ventas"] / dias_transcurridos
        datos["venta_promedio_diaria"] = round(promedio_diario, 2)
        datos["proyeccion_ventas"] = round(
            datos["total_ventas"] + promedio_diario * dias_restantes,
            2
        )
        datos["cumplimiento_ventas_proyectado"] = (
            datos["proyeccion_ventas"] / datos["cuota_ventas"] * 100
            if datos["cuota_ventas"] > 0 else 0
        )

        promedio_cobertura = datos["cobertura"] / dias_transcurridos
        datos["proyeccion_cobertura"] = int(round(
            datos["cobertura"] + promedio_cobertura * dias_restantes
        ))
        datos["cumplimiento_cobertura_proyectado"] = (
            datos["proyeccion_cobertura"] / datos["cuota_cobertura"] * 100
            if datos["cuota_cobertura"] > 0 else 0
        )

        logger.info("✅ Datos generales obtenidos")
        return datos

    except Exception as e:
        logger.exception(f"❌ Error obteniendo datos generales: {e}")
        return None
    finally:
        if conn:
            conn.close()


# ============================================================
# 8. GENERAR MENSAJES
# ============================================================


def generar_mensaje_vendedor(datos):
    """Genera reporte detallado para vendedor."""
    if not datos:
        return None

    cumpl = datos["cumplimiento"]
    cumpl_cobertura_proy = datos["cumplimiento_cobertura_proyectado"]

    proyectado_emoji = (
        "🟢" if datos["cumplimiento_proyectado"] >= 90
        else "🟡" if datos["cumplimiento_proyectado"] >= 75
        else "🔴"
    )
    cobertura_emoji = (
        "🟢" if cumpl_cobertura_proy >= 90
        else "🟡" if cumpl_cobertura_proy >= 75
        else "🔴"
    )
    cumplimiento_emoji = (
        "🟢" if cumpl >= 90
        else "🟡" if cumpl >= 75
        else "🔴"
    )

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


def generar_mensaje_jefe(datos):
    """Genera reporte consolidado para jefe/supervisor."""
    if not datos:
        return None

    cumpl = datos["cumplimiento"]
    cumpl_cobertura_proy = datos["cumplimiento_cobertura_proyectado"]

    proyectado_ventas_emoji = (
        "🟢" if datos["cumplimiento_ventas_proyectado"] >= 90
        else "🟡" if datos["cumplimiento_ventas_proyectado"] >= 75
        else "🔴"
    )
    cobertura_emoji = (
        "🟢" if cumpl_cobertura_proy >= 90
        else "🟡" if cumpl_cobertura_proy >= 75
        else "🔴"
    )
    cumplimiento_emoji = (
        "🟢" if cumpl >= 90
        else "🟡" if cumpl >= 75
        else "🔴"
    )

    lineas_txt = ""
    for linea, ventas in sorted(
        datos["lineas_negocio"].items(),
        key=lambda x: x[1],
        reverse=True
    ):
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
# 9. WHATSAPP
# ============================================================


def enviar_mensaje_whatsapp(numero_destino, mensaje):
    """Envía un mensaje de texto por WhatsApp Cloud API."""
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

        logger.info(f"📤 Enviando reporte a teléfono terminado en {numero[-4:]}")

        response = requests.post(
            API_URL,
            json=payload,
            headers=headers,
            timeout=15
        )

        if response.ok:
            logger.info(f"✅ Mensaje enviado correctamente a {numero[-4:]}")
            return True

        logger.error(
            f"❌ Error enviando mensaje ({response.status_code})"
        )
        return False

    except requests.RequestException as e:
        logger.error(f"❌ Error HTTP enviando WhatsApp: {e}")
        return False
    except Exception as e:
        logger.exception(f"❌ Error enviando WhatsApp: {e}")
        return False


# ============================================================
# 10. WEBHOOK
# ============================================================


@app.route("/webhook", methods=["GET"])
def verificar_webhook():
    """Verifica que Meta pueda conectarse al webhook."""
    verify_token = request.args.get("hub.verify_token")
    challenge = request.args.get("hub.challenge")

    if verify_token == VERIFY_TOKEN:
        logger.info("✅ Webhook verificado por Meta")
        return challenge or "", 200

    logger.warning("⚠️ Token de verificación incorrecto")
    return "Unauthorized", 403


@app.route("/webhook", methods=["POST"])
def recibir_mensaje():
    """Recibe mensajes de WhatsApp, valida y responde."""
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
            # También llegan eventos de estado sin messages.
            return jsonify({"status": "ok"}), 200

        mensaje_obj = mensajes[0]
        message_id = mensaje_obj.get("id")
        numero_remitente = mensaje_obj.get("from", "")
        tipo = mensaje_obj.get("type")

        if tipo != "text":
            logger.info(
                f"Mensaje no textual ignorado. tipo={tipo}, telefono=*{numero_remitente[-4:]}"
            )
            return jsonify({"status": "ok"}), 200

        texto_mensaje = (
            mensaje_obj.get("text", {}).get("body", "").strip()
        )
        logger.info(
            f"📨 Mensaje recibido de *{numero_remitente[-4:]}: tipo=text"
        )

        # Validar palabra clave antes de reservar el mensaje.
        if PALABRA_CLAVE not in texto_mensaje.lower():
            return jsonify({"status": "ok"}), 200

        # Evitar reportes duplicados por redelivery del webhook.
        if message_id and not reclamar_message_id(message_id, numero_remitente):
            logger.info(f"♻️ Webhook duplicado ignorado: {message_id}")
            return jsonify({"status": "duplicate"}), 200

        logger.info(f"✅ Palabra clave detectada: {PALABRA_CLAVE}")

        # Validar usuario.
        es_valido, datos_usuario = validar_vendedor(numero_remitente)

        if not es_valido:
            mensaje_respuesta = (
                "❌ Número no autorizado. Por favor contacta a tu supervisor."
            )
            exito = enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            if not exito:
                liberar_message_id(message_id)
                return jsonify({"status": "send_error"}), 500
            return jsonify({"status": "unauthorized"}), 200

        nombre_usuario = datos_usuario["nombre"]
        rol = datos_usuario.get("rol", "")
        es_jefe_supervisor = es_jefe(nombre_usuario, rol)

        if es_jefe_supervisor:
            logger.info(f"👔 Detectado: {rol or 'JEFE/SUPERVISOR'} ({nombre_usuario})")
            datos = obtener_datos_generales()
            mensaje = generar_mensaje_jefe(datos)
        else:
            logger.info(f"👤 Detectado: VENDEDOR ({nombre_usuario})")
            datos = obtener_datos_vendedor(nombre_usuario)
            mensaje = generar_mensaje_vendedor(datos)

        if not datos or not mensaje:
            mensaje_respuesta = "⚠️ No se encontraron datos. Intenta más tarde."
            exito = enviar_mensaje_whatsapp(numero_remitente, mensaje_respuesta)
            if not exito:
                liberar_message_id(message_id)
                return jsonify({"status": "send_error"}), 500
            return jsonify({"status": "no_data"}), 200

        exito = enviar_mensaje_whatsapp(numero_remitente, mensaje)

        if exito:
            logger.info(f"✅ Reporte enviado a {nombre_usuario}")
            return jsonify({"status": "success"}), 200

        # Si falló el envío, liberamos el ID para permitir un nuevo intento.
        liberar_message_id(message_id)
        logger.error("❌ Error al enviar reporte")
        return jsonify({"status": "send_error"}), 500

    except Exception as e:
        logger.exception(f"❌ Error procesando webhook: {e}")
        return jsonify({"status": "error"}), 500


# ============================================================
# 11. STATUS
# ============================================================


@app.route("/status", methods=["GET"])
def status():
    """Endpoint de estado del servicio."""
    ctx = obtener_contexto_periodo()
    dias = calcular_dias_laborables_periodo()

    return jsonify({
        "status": "active",
        "timestamp": ctx["ahora"].isoformat(),
        "webhook": "/webhook",
        "palabra_clave": PALABRA_CLAVE,
        "version": "V2 CORREGIDO",
        "periodo": ctx["periodo"],
        "mes": ctx["nombre_mes"],
        "dias_laborables": dias,
    }), 200


# ============================================================
# 12. ARRANQUE
# ============================================================


if __name__ == "__main__":
    logger.info("=" * 70)
    logger.info("🚀 SERVIDOR WEBHOOK WHATSAPP - V2 CORREGIDO")
    logger.info("=" * 70)
    ctx = obtener_contexto_periodo()
    dias = calcular_dias_laborables_periodo()
    logger.info(f"Periodo actual: {ctx['periodo']} - {ctx['nombre_mes']}")
    logger.info(f"Días laborables: {dias}")
    logger.info(f"BD: {BD_PATH}")
    logger.info(f"Excel vendedores: {EXCEL_VENDEDORES}")
    logger.info("Credenciales: variables de entorno")
    logger.info("=" * 70)

    inicializar_bd()

    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
