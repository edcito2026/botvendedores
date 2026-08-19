#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""COACH COMERCIAL - Worker independiente para Render.
Usa las mismas variables de entorno de WhatsApp/BD/Excel del Webhook V7.
"""
import os, time, sqlite3, logging
from datetime import datetime, date, timedelta
from zoneinfo import ZoneInfo
from collections import defaultdict
import requests, openpyxl

BASE=os.path.dirname(os.path.abspath(__file__))
TOKEN=os.environ.get('WHATSAPP_ACCESS_TOKEN')
PHONE_ID=os.environ.get('WHATSAPP_PHONE_NUMBER_ID')
API_VER=os.environ.get('WHATSAPP_API_VERSION','v25.0')
if not TOKEN or not PHONE_ID: raise RuntimeError('Faltan credenciales WhatsApp del Webhook V7')
API_URL=f'https://graph.facebook.com/{API_VER}/{PHONE_ID}/messages'
BD=os.environ.get('BD_PATH',os.path.join(BASE,'ventas.db'))
VENDEDORES_XLSX=os.environ.get('EXCEL_VENDEDORES',os.path.join(BASE,'vendedores.xlsx'))
FERIADOS_XLSX=os.environ.get('FERIADOS_XLSX',os.path.join(BASE,'FERIADOS.xlsx'))
HORA=int(os.environ.get('COACH_HORA','7')); MINUTO=int(os.environ.get('COACH_MINUTO','0'))
TZ=ZoneInfo('America/Lima'); LAB={0,1,2,3,4,5}
DIAS=['LUNES','MARTES','MIÉRCOLES','JUEVES','VIERNES','SÁBADO','DOMINGO']
MESES=['ENERO','FEBRERO','MARZO','ABRIL','MAYO','JUNIO','JULIO','AGOSTO','SETIEMBRE','OCTUBRE','NOVIEMBRE','DICIEMBRE']
LOG=os.path.join(BASE,'logs'); os.makedirs(LOG,exist_ok=True)
logging.basicConfig(level=logging.INFO,format='%(asctime)s - %(levelname)s - %(message)s',handlers=[logging.FileHandler(os.path.join(LOG,'coach_comercial.log'),encoding='utf-8'),logging.StreamHandler()])
logger=logging.getLogger('coach')

def now(): return datetime.now(TZ)
def tel(x): return ''.join(filter(str.isdigit,str(x or '')))
def db():
    c=sqlite3.connect(BD,timeout=30); c.row_factory=sqlite3.Row; c.execute('PRAGMA busy_timeout=30000'); c.execute('PRAGMA journal_mode=WAL'); return c

def feriados():
    s=set()
    try:
        if not os.path.exists(FERIADOS_XLSX): return s
        w=openpyxl.load_workbook(FERIADOS_XLSX,read_only=True,data_only=True)
        for ws in w.worksheets:
            for row in ws.iter_rows(values_only=True):
                for v in row:
                    if isinstance(v,datetime): s.add(v.date())
                    elif isinstance(v,date): s.add(v)
                    elif isinstance(v,str):
                        for f in ('%d/%m/%Y','%Y-%m-%d','%d-%m-%Y'):
                            try: s.add(datetime.strptime(v.strip(),f).date()); break
                            except ValueError: pass
        w.close()
    except Exception as e: logger.error('Feriados: %s',e)
    return s
FERIADOS=feriados()
def laborable(d): return d.weekday() in LAB and d not in FERIADOS
def contexto():
    h=now(); y,m=h.year,h.month; p=date(y,m,1); q=date(y+1,1,1) if m==12 else date(y,m+1,1); u=q-timedelta(days=1)
    return {'hoy':h.date(),'anio':y,'mes':m,'periodo':f'{y}{m:02d}','mes_nombre':MESES[m-1],'primer':p,'ultimo':u}
def dias_mes(c):
    fs=[]; d=c['primer']
    while d<=c['ultimo']:
        if laborable(d): fs.append(d)
        d+=timedelta(days=1)
    return fs,[x for x in fs if x<=c['hoy']],[x for x in fs if x>c['hoy']]
def periodo_ant(c): return f"{c['anio']-1}12" if c['mes']==1 else f"{c['anio']}{c['mes']-1:02d}"

def vendedores():
    out={}
    w=openpyxl.load_workbook(VENDEDORES_XLSX,read_only=True,data_only=True); ws=w.active
    for r in ws.iter_rows(values_only=True):
        if len(r)>=2 and r[0] and r[1]:
            n=str(r[0]).strip(); t=tel(r[1]); rol=str(r[3]).upper().strip() if len(r)>3 and r[3] else ''
            if t: out[t]={'nombre':n,'telefono':t,'rol':rol}
    w.close(); return out

def init_control():
    c=db(); c.execute('CREATE TABLE IF NOT EXISTS coach_envios(fecha TEXT,telefono TEXT,vendedor TEXT,enviado_en TEXT,PRIMARY KEY(fecha,telefono))'); c.commit(); c.close()
def reclamar(fecha,t,n):
    c=db()
    try:
        x=c.execute('INSERT OR IGNORE INTO coach_envios VALUES(?,?,?,?)',(fecha,t,n,now().isoformat())); c.commit(); return x.rowcount==1
    finally: c.close()

def cuota(cur,n,c):
    r=cur.execute("SELECT COALESCE(SUM(Cuota_Soles),0) cuota,COALESCE(SUM(Cuota_Cobertura),0) cob FROM cuotas WHERE Vendedor=? AND AÑO=? AND NRO_MES=? AND Proveedor='ARCOR'",(n,c['anio'],c['mes'])).fetchone(); return float(r['cuota'] or 0),int(r['cob'] or 0)
def base(cur,n,p):
    r=cur.execute("SELECT COALESCE(SUM(CAST(Imp_Total AS REAL)),0) v,COUNT(DISTINCT Cod_Clie) cli FROM VENTAS2026 WHERE Vendedor=? AND Periodo=? AND Proveedor='ARCOR' AND CAST(Imp_Total AS REAL)>0",(n,p)).fetchone(); v=float(r['v'] or 0); cli=int(r['cli'] or 0); return v,cli,(v/cli if cli else 0)
def dia_idx(v):
    s=str(v or '').upper(); names={'LUNES':0,'MARTES':1,'MIERCOLES':2,'MIÉRCOLES':2,'JUEVES':3,'VIERNES':4,'SABADO':5,'SÁBADO':5}
    try:
        if s[:1].isdigit(): return int(s.split()[0])-1
    except: pass
    for k,x in names.items():
        if k in s: return x
    return None

def pesos(cur,n,c):
    ini=c['hoy']-timedelta(days=120); rec=c['hoy']-timedelta(days=35); total=defaultdict(float); reciente=defaultdict(float)
    for row in cur.execute("SELECT Dia_Sem,SUM(CAST(Imp_Total AS REAL)) v FROM VENTAS2026 WHERE Vendedor=? AND Proveedor='ARCOR' AND CAST(Imp_Total AS REAL)>0 AND date(F_Emis)>=date(?) AND date(F_Emis)<date(?) GROUP BY Dia_Sem",(n,ini.isoformat(),c['primer'].isoformat())):
        d=dia_idx(row['Dia_Sem']);
        if d is not None and d<6: total[d]+=float(row['v'] or 0)
    for row in cur.execute("SELECT Dia_Sem,SUM(CAST(Imp_Total AS REAL)) v FROM VENTAS2026 WHERE Vendedor=? AND Proveedor='ARCOR' AND CAST(Imp_Total AS REAL)>0 AND date(F_Emis)>=date(?) AND date(F_Emis)<date(?) GROUP BY Dia_Sem",(n,rec.isoformat(),c['primer'].isoformat())):
        d=dia_idx(row['Dia_Sem']);
        if d is not None and d<6: reciente[d]+=float(row['v'] or 0)
    def norm(x):
        z=sum(x.values()); return {d:v/z for d,v in x.items()} if z else {}
    a,b=norm(total),norm(reciente); p={d:(.6*b.get(d,0)+.4*a.get(d,0) if a and b else (b.get(d,0) if b else a.get(d,0))) for d in range(6)}; z=sum(p.values()); return {d:(p[d]/z if z else 1/6) for d in range(6)}
def objetivo(saldo,fs,p):
    if saldo<=0 or not fs:return 0
    vals=[max(p.get(x.weekday(),0),1e-9) for x in fs]; return saldo*vals[0]/sum(vals)

def troya(cur,n,p,didx):
    # Dia_Sem no existe en clientes; se infiere el día habitual desde VENTAS2026.
    nd=DIAS[didx]; q="""
    WITH h AS (SELECT CAST(Cod_Clie AS REAL) id,Dia_Sem,COUNT(*) f,SUM(CAST(Imp_Total AS REAL)) v FROM VENTAS2026 WHERE Vendedor=? AND Proveedor='ARCOR' AND CAST(Imp_Total AS REAL)>0 GROUP BY CAST(Cod_Clie AS REAL),Dia_Sem),
    dh AS (SELECT id,Dia_Sem FROM (SELECT h.*,ROW_NUMBER() OVER(PARTITION BY id ORDER BY f DESC,v DESC) rn FROM h) WHERE rn=1),
    cp AS (SELECT DISTINCT CAST(Cod_Clie AS REAL) id FROM VENTAS2026 WHERE Vendedor=? AND Proveedor='ARCOR' AND Periodo=? AND CAST(Imp_Total AS REAL)>0)
    SELECT c.Raz_Social,MAX(h.v) potencial FROM clientes c JOIN dh ON CAST(c.Cod_Clie AS REAL)=dh.id LEFT JOIN h ON h.id=dh.id AND h.Dia_Sem=dh.Dia_Sem LEFT JOIN cp ON cp.id=CAST(c.Cod_Clie AS REAL) WHERE c.Calif='D' AND c.Vendedor=? AND dh.Dia_Sem LIKE ? AND cp.id IS NULL GROUP BY c.Cod_Clie,c.Raz_Social ORDER BY potencial DESC LIMIT 5"""
    return [dict(r) for r in cur.execute(q,(n,n,p,n,f'%{nd[:3]}%')).fetchall()]
def promociones_3_menor_venta(cur, n, p, pa):
    """3 promociones con menor venta del vendedor en el período actual."""
    q = """
    SELECT TRIM(Promo) x,
           SUM(CAST(Imp_Total AS REAL)) actual
    FROM VENTAS2026
    WHERE Vendedor=?
      AND Proveedor='ARCOR'
      AND Periodo=?
      AND TRIM(COALESCE(Promo,''))<>''
    GROUP BY TRIM(Promo)
    ORDER BY actual ASC
    LIMIT 3
    """
    return [dict(r) for r in cur.execute(q, (n, p)).fetchall()]


def productos_5_oportunidad(cur, n, p):
    """
    1) TOP 30 productos más vendidos de ARCOR en general en el período.
    2) Dentro de esos 30, ordenar por venta del vendedor.
    3) Mostrar los 5 con menor venta del vendedor.
    """
    q = """
    WITH top30 AS (
        SELECT TRIM(Producto) producto
        FROM VENTAS2026
        WHERE Proveedor='ARCOR'
          AND Periodo=?
          AND TRIM(COALESCE(Producto,''))<>''
          AND CAST(Imp_Total AS REAL)>0
        GROUP BY TRIM(Producto)
        ORDER BY SUM(CAST(Imp_Total AS REAL)) DESC
        LIMIT 20
    ),
    vendedor AS (
        SELECT TRIM(Producto) producto,
               SUM(CAST(Imp_Total AS REAL)) venta
        FROM VENTAS2026
        WHERE Vendedor=?
          AND Proveedor='ARCOR'
          AND Periodo=?
          AND TRIM(COALESCE(Producto,''))<>''
          AND CAST(Imp_Total AS REAL)>0
        GROUP BY TRIM(Producto)
    )
    SELECT t.producto,
           COALESCE(v.venta,0) actual
    FROM top30 t
    LEFT JOIN vendedor v ON v.producto=t.producto
    ORDER BY actual ASC, t.producto
    LIMIT 5
    """
    return [
        dict(r)
        for r in cur.execute(q, (p, n, p)).fetchall()
    ]


def analizar(cur,n,c):
    q,cob=cuota(cur,n,c)
    v,cli,ticket=base(cur,n,c['periodo'])
    saldo=max(q-v,0)
    fs,trans,rest=dias_mes(c)
    plan=([c['hoy']] if laborable(c['hoy']) else [])+rest
    p=pesos(cur,n,c)
    obj=objetivo(saldo,plan,p)
    cn=int((obj/ticket)+.999999) if obj>0 and ticket>0 else 0

    cobpend=max(cob-cli,0)
    cobh=0
    if plan and cobpend:
        pesos_plan=[max(p.get(x.weekday(),0),1e-9) for x in plan]
        cobh=int(round(cobpend*pesos_plan[0]/sum(pesos_plan)))

    return {
        'vendedor':n,
        'cuota':q,
        'cob_cuota':cob,
        'venta':v,
        'clientes':cli,
        'ticket':ticket,
        'saldo':saldo,
        'objetivo':obj,
        'clientes_necesarios':max(cn,cobh),
        'peso_hoy':p.get(c['hoy'].weekday(),0),
        'troya':troya(cur,n,c['periodo'],c['hoy'].weekday()) if laborable(c['hoy']) else [],
        'promos':promociones_3_menor_venta(cur,n,c['periodo'],periodo_ant(c)),
        'productos':productos_5_oportunidad(cur,n,c['periodo'])
    }


def estado_comercial(a,c):
    total,trans,rest=dias_mes(c)
    if not total or a["cuota"]<=0:
        return "ENFOQUE",0.0
    avance_real=a["venta"]/a["cuota"]
    avance_esperado=len(trans)/len(total)
    brecha_pp=(avance_real-avance_esperado)*100
    if len(rest)<=2 and brecha_pp < -5: return "URGENTE",brecha_pp
    if brecha_pp < -10: return "RECUPERACION",brecha_pp
    if brecha_pp < -5: return "REACCION",brecha_pp
    if brecha_pp <= 5: return "ENFOQUE",brecha_pp
    return "BUEN_RITMO",brecha_pp

SALUDOS={
"BUEN_RITMO":[
"☀️ ¡Buenos días, {nombre}! 😊\n¡Vienes haciendo un buen trabajo! Hoy vamos a mantener el ritmo y aprovechar nuevas oportunidades. 💪",
"🌟 ¡Excelente día, {nombre}!\nTu avance viene bien encaminado. ¡Vamos a seguir construyendo un gran cierre! 🔥",
"🚀 ¡Buenos días, {nombre}!\n¡Seguimos avanzando! Hoy vamos por otro buen día de ventas. 💪"],
"ENFOQUE":[
"☀️ ¡Buenos días, {nombre}! 😊\nHoy tenemos una nueva oportunidad para acercarnos a la cuota. ¡Vamos con foco! 🎯",
"💪 ¡Buenos días, {nombre}!\nEl objetivo sigue al alcance. Hoy concentremos el esfuerzo donde tenemos mejores oportunidades.",
"🎯 ¡Buen día, {nombre}!\nCada día cuenta. Hoy salgamos con un objetivo claro y aprovechemos cada oportunidad."],
"REACCION":[
"🔥 ¡Buenos días, {nombre}!\nHoy necesitamos reaccionar. Estamos por debajo del ritmo necesario y cada venta cuenta. ¡Vamos a recuperar terreno!",
"⚡ ¡Buenos días, {nombre}!\nTenemos una brecha que debemos comenzar a cerrar hoy. Enfoquémonos en las oportunidades que pueden generar venta.",
"🎯 ¡Vamos, {nombre}!\nHoy toca recuperar ritmo. Tenemos claro cuánto necesitamos vender y dónde buscar las oportunidades."],
"RECUPERACION":[
"🚨 ¡Buenos días, {nombre}!\nNecesitamos recuperar ritmo. La brecha frente al objetivo ya es importante. Hoy debemos aprovechar cada oportunidad.",
"🔥 ¡Buenos días, {nombre}!\nHoy es un día clave para reaccionar. Vamos a buscar el objetivo desde el primer cliente.",
"⚠️ ¡Buenos días, {nombre}!\nEstamos por debajo del ritmo necesario. Hoy necesitamos ejecución, foco y aprovechar cada oportunidad."],
"URGENTE":[
"🚨 ¡Buenos días, {nombre}!\nEntramos en una etapa decisiva. Quedan muy pocos días y todavía existe una brecha importante. Hoy necesitamos máxima concentración.",
"🔥 ¡Buenos días, {nombre}!\nEstamos en la recta final. Cada cliente y cada venta cuentan. Hoy tenemos que salir a buscar el objetivo.",
"🚨 ¡Buenos días, {nombre}!\nQuedan pocos días para cerrar el mes. Hoy necesitamos un día fuerte y enfocado para reducir la brecha."]
}
def saludo_adaptativo(a,c):
    estado,brecha=estado_comercial(a,c)
    opciones=SALUDOS[estado]
    clave=f"{a['vendedor']}|{c['hoy'].isoformat()}|{estado}"
    i=sum(ord(x) for x in clave)%len(opciones)
    return opciones[i].format(nombre=a["vendedor"].title()),estado,brecha


def msg(a,c):
    d=DIAS[c['hoy'].weekday()]
    saludo, estado, brecha_pp = saludo_adaptativo(a,c)

    if a['saldo']<=0:
        return f"""{saludo}

📊 Cuota: S/. {a['cuota']:,.2f}
💰 Vendido: S/. {a['venta']:,.2f}
🔻 Falta: S/. 0.00

🏆 ¡CUOTA CUMPLIDA!

💵 Ticket: S/. {a['ticket']:,.2f}

🎯 Hoy: mantén el ritmo, protege el ticket y aprovecha las oportunidades TROYA.

🤖 Coach Comercial N&J"""

    s=f"""{saludo}

📊 Cuota: S/. {a['cuota']:,.2f}
💰 Vendido: S/. {a['venta']:,.2f}
🔻 Falta: S/. {a['saldo']:,.2f}

🎯 VENTA OBJETIVO HOY
S/. {a['objetivo']:,.2f}

📅 Este objetivo considera tu comportamiento histórico
por día de semana y los días laborables que quedan.

👥 COBERTURA
Busca aproximadamente {a['clientes_necesarios']} clientes con compra,
usando tu ticket actual de S/. {a['ticket']:,.2f}."""

    if a['troya']:
        s+=f"""

🎯 TROYA — {d}
Clientes TROYA de tu día de visita que aún NO compraron:
"""
        for r in a['troya']:
            s+=f"• {r['Raz_Social'][:34]}\n"
        s+="👉 Prioridad: recuperar estas compras hoy."

    if a['promos']:
        s+="\n\n🏷️ PROMOCIONES A IMPULSAR\n"
        for i,r in enumerate(a['promos'],1):
            s+=f"{i}️⃣ {r['x'][:45]}\n"

    if a['productos']:
        s+="""\n📦 PRODUCTOS A IMPULSAR
Dentro de los 20 productos más vendidos de ARCOR,
estos son tus 5 con menor venta:
"""
        for i,r in enumerate(a['productos'],1):
            s+=f"{i}️⃣ {r['producto'][:45]}\n"

    s+=f"""

💵 TICKET
Actual: S/. {a['ticket']:,.2f}
👉 Busca aumentar el valor de cada pedido.

🔥 PRIORIDAD DE HOY
1️⃣ Alcanzar S/. {a['objetivo']:,.2f}
2️⃣ Recuperar TROYA del {d.lower()}
3️⃣ Impulsar promociones y productos con menor venta
4️⃣ Aumentar ticket por cliente

🤖 Coach Comercial N&J"""
    return s


def enviar(t,m):
    t=tel(t); t=t if t.startswith('51') else '51'+t
    try:
        r=requests.post(API_URL,json={'messaging_product':'whatsapp','to':t,'type':'text','text':{'body':m}},headers={'Authorization':f'Bearer {TOKEN}','Content-Type':'application/json'},timeout=20)
        if r.ok: logger.info('✅ Coach enviado a *%s',t[-4:]); return True
        logger.error('❌ WhatsApp %s %s',r.status_code,r.text[:300]); return False
    except Exception as e: logger.exception('❌ Envío Coach: %s',e); return False

def ejecutar():
    c=contexto()
    if not laborable(c['hoy']): logger.info('⏭️ Día no laborable: %s',c['hoy']); return
    init_control(); vs=vendedores(); conn=db()
    try:
        for t,u in vs.items():
            if any(x in u['rol'] for x in ('JEFE','SUPERVISOR','GERENTE','COORDINADOR')): continue
            if not reclamar(c['hoy'].isoformat(),t,u['nombre']): continue
            try:
                a=analizar(conn.cursor(),u['nombre'],c)
                estado, brecha_pp = estado_comercial(a,c)
                logger.info('🧠 %s | estado=%s | brecha=%+.1f pp',u['nombre'],estado,brecha_pp)
                if a['cuota']<=0: logger.warning('⚠️ Sin cuota ARCOR: %s',u['nombre']); continue
                enviar(t,msg(a,c))
            except Exception: logger.exception('❌ Coach %s',u['nombre'])
    finally: conn.close()

def segundos():
    a=now(); x=a.replace(hour=HORA,minute=MINUTO,second=0,microsecond=0)
    if a>=x:x+=timedelta(days=1)
    return max((x-a).total_seconds(),1)

if __name__=='__main__':
    logger.info('🧠 COACH COMERCIAL WORKER V1 | %02d:%02d America/Lima',HORA,MINUTO)
    logger.info('BD=%s | VENDEDORES=%s | FERIADOS=%s',BD,VENDEDORES_XLSX,FERIADOS_XLSX)
    init_control(); ultima=None
    while True:
        a=now(); h=a.date()
        if a.hour==HORA and a.minute==MINUTO and ultima!=h:
            try: ejecutar()
            except Exception: logger.exception('❌ Error general Coach')
            ultima=h
        time.sleep(20)
