import io, csv, smtplib, datetime, threading, os, sqlite3
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from flask import Flask, request, redirect, session, render_template_string, Response, send_file, g
from fpdf import FPDF

app = Flask(__name__)
app.secret_key = "liga-mendocina-demo-123"
import os
DATA_DIR = "/data" if os.path.exists("/data") else "."
DB = os.path.join(DATA_DIR, "liga.db")
# Migrar base vieja al disco la primera vez
if DATA_DIR == "/data" and not os.path.exists(DB) and os.path.exists("liga.db"):
    import shutil
    shutil.copy("liga.db", DB)
    print("MIGRADO liga.db a /data")

EMAIL_REMITENTE = "fabiand@didiosistemas.com.ar"
EMAIL_CLAVE_APP = "iutu skcd dmut zhij"





# --- MIGRACION BAJA V2: asegurar columna estado ---
try:
    import sqlite3 as _sq
    _mdb = _sq.connect("liga.db")
    _cols = [r[1] for r in _mdb.execute("PRAGMA table_info(jugadores)").fetchall()]
    if "estado" not in _cols:
        _mdb.execute("ALTER TABLE jugadores ADD COLUMN estado TEXT DEFAULT 'ACTIVO'")
        _mdb.execute("UPDATE jugadores SET estado='ACTIVO' WHERE estado IS NULL OR estado=''")
        _mdb.commit()
    _mdb.close()
except Exception as e:
    print("Migracion estado V2:", e)

@app.route("/logo.png")
def logo():
    if os.path.exists("logo.png"):
        return send_file("logo.png", mimetype="image/png")
    return "", 404

ADMIN_USER = "admin@liga.com"
ADMIN_PASS = "admin123"

JUGADORES = {
    "50035043": {"nombre": "CARLOS GABRIEL CHAVEZ", "club": "LUJAN SPORT CLUB", "categoria": "U-16", "torneo": "COPA PLATA 7MA"},
    "50812892": {"nombre": "BENJAMIN CASTRO CAMPOS", "club": "LUJAN SPORT CLUB", "categoria": "U-15", "torneo": "COPA PLATA 8VA"},
    "51469824": {"nombre": "MATIAS SEBASTIAN GONZALEZ", "club": "LUJAN SPORT CLUB", "categoria": "U-15", "torneo": "COPA PLATA 8VA"},
    "50953809": {"nombre": "LAUTARO SALCEDO", "club": "LUJAN SPORT CLUB", "categoria": "U-15", "torneo": "COPA PLATA 8VA"},
    "51090610": {"nombre": "MARIANO ANDRES MARTINEZ", "club": "LUJAN SPORT CLUB", "categoria": "U-15", "torneo": "COPA PLATA 8VA"},
    "52516346": {"nombre": "VALENTINO SALCEDO", "club": "LUJAN SPORT CLUB", "categoria": "U-13", "torneo": "COPA PLATA 10MA"},
    "53158209": {"nombre": "AGOSTINO VIEL EDCORIAZA", "club": "LUJAN SPORT CLUB", "categoria": "U-13", "torneo": "COPA PLATA 10MA"},
    "52939620": {"nombre": "JOAQUIN ARIEL PALEO DAVILA", "club": "LUJAN SPORT CLUB", "categoria": "U-13", "torneo": "COPA PLATA 10MA"},
    "51016535": {"nombre": "JUAN CRUZ OCTAVIO GIUNTA VEGA", "club": "DEPORTIVO MAIPU", "categoria": "U-15", "torneo": "COPA ORO 8VA"},
}
CARGOS = [
    {"id": 1, "nombre": "Inscripcion Anual Liga 2026/2027", "monto": 40000},
    {"id": 2, "nombre": "Seguro Deportivo Obligatorio", "monto": 8000},
    {"id": 3, "nombre": "Carnet / Ficha del Jugador", "monto": 5000},
]

def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB)
        g.db.row_factory = sqlite3.Row
    return g.db

@app.teardown_appcontext
def close_db(e=None):
    db = g.pop("db", None)
    if db: db.close()

def init_db():
    db = sqlite3.connect(DB)
    db.execute("CREATE TABLE IF NOT EXISTS usuarios (email TEXT PRIMARY KEY, dni_jugador TEXT, password TEXT)")
    db.execute("CREATE TABLE IF NOT EXISTS pagos (email TEXT, cargo_id INTEGER, PRIMARY KEY(email, cargo_id))")
    db.execute("CREATE TABLE IF NOT EXISTS jugadores (dni TEXT PRIMARY KEY, nombre TEXT, club TEXT, categoria TEXT, torneo TEXT)")
    db.execute("CREATE TABLE IF NOT EXISTS clubes (nombre TEXT PRIMARY KEY)")
    db.execute("CREATE TABLE IF NOT EXISTS clubes_usuarios (club TEXT PRIMARY KEY, usuario TEXT UNIQUE, password TEXT)")
    db.execute("CREATE TABLE IF NOT EXISTS categorias (nombre TEXT PRIMARY KEY)")
    db.execute("CREATE TABLE IF NOT EXISTS cargos (nombre TEXT PRIMARY KEY, monto INTEGER)")
    db.execute("CREATE TABLE IF NOT EXISTS temporadas (nombre TEXT PRIMARY KEY, activa INTEGER DEFAULT 1)")
    db.execute("CREATE TABLE IF NOT EXISTS tipos_competicion (nombre TEXT PRIMARY KEY)")
    db.execute("CREATE TABLE IF NOT EXISTS niveles (nombre TEXT PRIMARY KEY)")
    db.execute("CREATE TABLE IF NOT EXISTS torneos (nombre TEXT PRIMARY KEY, temporada TEXT, tipo TEXT)")
    # Datos iniciales desde tu Excel real
    db.execute("INSERT OR IGNORE INTO temporadas VALUES ('2026/2027', 1)")
    db.execute("INSERT OR IGNORE INTO temporadas VALUES ('2025/2026', 0)")
    db.execute("INSERT OR IGNORE INTO tipos_competicion VALUES ('TORNEO DIVISIONES INFERIORES COPA PLATA 7MA')")
    db.execute("INSERT OR IGNORE INTO tipos_competicion VALUES ('TORNEO DIVISIONES INFERIORES COPA PLATA 8VA')")
    db.execute("INSERT OR IGNORE INTO tipos_competicion VALUES ('TORNEO DIVISIONES INFERIORES COPA PLATA 10MA')")
    db.execute("INSERT OR IGNORE INTO tipos_competicion VALUES ('TORNEO DIVISIONES INFERIORES COPA ORO 8VA')")
    db.execute("INSERT OR IGNORE INTO niveles VALUES ('Aficionado')")
    db.execute("INSERT OR IGNORE INTO niveles VALUES ('Federado')")
    db.execute("INSERT OR IGNORE INTO niveles VALUES ('Profesional')")
    db.execute("INSERT OR IGNORE INTO torneos VALUES ('TORNEO CLAUSURA DIVISIONES INFERIORES COPA PLATA 7MA 26/27','2026/2027','TORNEO DIVISIONES INFERIORES COPA PLATA 7MA')")
    db.execute("INSERT OR IGNORE INTO torneos VALUES ('TORNEO CLAUSURA DIVISIONES INFERIORES COPA PLATA 8VA 26/27','2026/2027','TORNEO DIVISIONES INFERIORES COPA PLATA 8VA')")
    db.execute("INSERT OR IGNORE INTO torneos VALUES ('TORNEO CLAUSURA DIVISIONES INFERIORES COPA PLATA 10MA 26/27','2026/2027','TORNEO DIVISIONES INFERIORES COPA PLATA 10MA')")
    db.execute("INSERT OR IGNORE INTO torneos VALUES ('TORNEO CLAUSURA DIVISIONES INFERIORES COPA ORO 8VA 26/27','2026/2027','TORNEO DIVISIONES INFERIORES COPA ORO 8VA')")
    # Ampliar jugadores con los campos nuevos (si no existen)
    try: db.execute("ALTER TABLE jugadores ADD COLUMN temporada TEXT")
    except: pass
    try: db.execute("ALTER TABLE jugadores ADD COLUMN tipo TEXT")
    except: pass
    try: db.execute("ALTER TABLE jugadores ADD COLUMN nivel TEXT DEFAULT 'Aficionado'")
    except: pass
    try: db.execute("ALTER TABLE jugadores ADD COLUMN fecha_nac TEXT")
    except: pass
    cur4 = db.execute("SELECT COUNT(*) FROM categorias").fetchone()[0]
    if cur4 == 0:
        cats = set()
        for j0 in JUGADORES.values():
            cats.add(j0["categoria"])
        for c0 in cats:
            db.execute("INSERT OR IGNORE INTO categorias VALUES (?)", (c0,))
    cur5 = db.execute("SELECT COUNT(*) FROM cargos").fetchone()[0]
    if cur5 == 0:
        for n0, m0 in [("Inscripcion Anual Liga 2026/2027", 40000), ("Seguro Deportivo Obligatorio", 8000), ("Carnet / Ficha del Jugador", 5000)]:
            db.execute("INSERT OR IGNORE INTO cargos VALUES (?, ?)", (n0, m0))
    cur3 = db.execute("SELECT COUNT(*) FROM clubes").fetchone()[0]
    if cur3 == 0:
        clubs_existentes = set()
        for j0 in JUGADORES.values():
            clubs_existentes.add(j0["club"])
        for c0 in clubs_existentes:
            db.execute("INSERT OR IGNORE INTO clubes VALUES (?)", (c0,))
    # Si la tabla esta vacia, cargar los jugadores que ya teniamos en el codigo
    cur2 = db.execute("SELECT COUNT(*) FROM jugadores").fetchone()[0]
    if cur2 == 0:
        for dni0, j0 in JUGADORES.items():
            db.execute("INSERT OR IGNORE INTO jugadores VALUES (?,?,?,?,?)", (dni0, j0["nombre"], j0["club"], j0["categoria"], j0["torneo"]))
    # Cargar jugadores de la DB al diccionario en memoria
    for r in db.execute("SELECT * FROM jugadores").fetchall():
        JUGADORES[r[0]] = {"nombre": r[1], "club": r[2], "categoria": r[3], "torneo": r[4]}
    pass # no crear demo automatica
    db.commit(); db.close()

def get_usuario(email):
    db = get_db()
    row = db.execute("SELECT * FROM usuarios WHERE email=?", (email,)).fetchone()
    if not row: return None
    pagos = [r[0] for r in db.execute("SELECT cargo_id FROM pagos WHERE email=?", (email,)).fetchall()]
    return {"dni_jugador": row["dni_jugador"], "password": row["password"], "pagos": pagos}

def crear_usuario(email, dni):
    db = get_db()
    # CONTROL DNI UNICO: un DNI no puede estar en 2 usuarios
    db.execute("INSERT INTO usuarios VALUES (?,?,?)", (email, dni, dni))
    db.commit()

def agregar_pago(email, cid):
    db = get_db()
    db.execute("INSERT OR IGNORE INTO pagos VALUES (?,?)", (email, cid))
    db.commit()

def pagar_todo_db(email):
    db = get_db()
    for c in get_cargos_pago():
        db.execute("INSERT OR IGNORE INTO pagos VALUES (?,?)", (email, c["id"]))
    db.commit()

def todos_usuarios():
    db = get_db()
    rows = db.execute("SELECT * FROM usuarios").fetchall()
    out = {}
    for r in rows:
        pagos = [x[0] for x in db.execute("SELECT cargo_id FROM pagos WHERE email=?", (r["email"],)).fetchall()]
        out[r["email"]] = {"dni_jugador": r["dni_jugador"], "password": r["password"], "pagos": pagos}
    return out

def agregar_jugador_db(dni, nombre, club, categoria, torneo):
    db = get_db()
    db.execute("INSERT INTO jugadores (dni, nombre, club, categoria, torneo, estado) VALUES (?,?,?,?,?, 'ACTIVO') ON CONFLICT(dni) DO UPDATE SET nombre=excluded.nombre, club=excluded.club, categoria=excluded.categoria, torneo=excluded.torneo, estado='ACTIVO'", (dni, nombre, club, categoria, torneo))
    db.commit()
    JUGADORES[dni] = {"nombre": nombre, "club": club, "categoria": categoria, "torneo": torneo, "estado": "ACTIVO"}

def borrar_jugador_db(dni):
    db = get_db()
    db.execute("UPDATE jugadores SET estado='BAJA' WHERE dni=?", (dni,))
    db.commit()
    if dni in JUGADORES:
        JUGADORES[dni]['estado'] = 'BAJA'

def reactivar_jugador_db(dni):
    db = get_db()
    db.execute("UPDATE jugadores SET estado='ACTIVO' WHERE dni=?", (dni,))
    db.commit()
    if dni in JUGADORES:
        JUGADORES[dni]['estado'] = 'ACTIVO'

def get_clubes():
    db = get_db()
    return [r[0] for r in db.execute("SELECT nombre FROM clubes ORDER BY nombre").fetchall()]


# --- PARCHE CAT: USUARIOS CLUBES ---
def slug_club(nombre):
    n = str(nombre).lower().strip()
    for a,b in [("á","a"),("é","e"),("í","i"),("ó","o"),("ú","u"),("ñ","n"),("ü","u")]:
        n = n.replace(a,b)
    out = ""
    for ch in n:
        if ch.isalnum():
            out += ch
        elif ch in [" ","-","_","(",")",".","/"]:
            if not out.endswith("_"):
                out += "_"
    return out.strip("_") or "club"

def get_usuario_club(club):
    db = get_db()
    try:
        return db.execute("SELECT * FROM clubes_usuarios WHERE club=?", (club,)).fetchone()
    except Exception:
        return None

def crear_usuario_club(club):
    db = get_db()
    db.execute("CREATE TABLE IF NOT EXISTS clubes_usuarios (club TEXT PRIMARY KEY, usuario TEXT UNIQUE, password TEXT)")
    base = slug_club(club)
    usuario = base
    i = 2
    while db.execute("SELECT 1 FROM clubes_usuarios WHERE usuario=?", (usuario,)).fetchone():
        usuario = f"{base}_{i}"
        i += 1
    db.execute("INSERT OR REPLACE INTO clubes_usuarios VALUES (?,?,?)", (club, usuario, "1234"))
    db.commit()
    return usuario

def agregar_club_db(nombre):
    db = get_db()
    db.execute("INSERT OR IGNORE INTO clubes VALUES (?)", (nombre,))
    db.commit()

def borrar_club_db(nombre):
    db = get_db()
    db.execute("DELETE FROM clubes WHERE nombre=?", (nombre,))
    db.commit()

def get_categorias():
    db = get_db()
    return [r[0] for r in db.execute("SELECT nombre FROM categorias ORDER BY nombre").fetchall()]

def agregar_categoria_db(nombre):
    db = get_db()
    db.execute("INSERT OR IGNORE INTO categorias VALUES (?)", (nombre,))
    db.commit()

def borrar_categoria_db(nombre):
    db = get_db()
    db.execute("DELETE FROM categorias WHERE nombre=?", (nombre,))
    db.commit()

def migrar_cargos_viejos():
    try:
        db = get_db()
        rows = [r[0] for r in db.execute("SELECT nombre FROM cargos").fetchall()]
        if "INSCRIPCION" in rows or "SEGURO" in rows or "CARNET" in rows:
            db.execute("DELETE FROM cargos")
            for n0, m0 in [("Inscripcion Anual Liga 2026/2027", 40000), ("Seguro Deportivo Obligatorio", 8000), ("Carnet / Ficha del Jugador", 5000)]:
                db.execute("INSERT OR IGNORE INTO cargos VALUES (?, ?)", (n0, m0))
            db.commit()
    except Exception as e:
        print("migrar cargos:", e)

def get_cargos_pago():
    migrar_cargos_viejos()
    db = get_db()
    rows = db.execute("SELECT nombre, monto FROM cargos ORDER BY rowid").fetchall()
    return [{"id": i+1, "nombre": n, "monto": m} for i, (n, m) in enumerate(rows)]

def get_cargos():
    db = get_db()
    return db.execute("SELECT nombre, monto FROM cargos ORDER BY rowid").fetchall()

def agregar_cargo_db(nombre, monto):
    db = get_db()
    db.execute("INSERT OR REPLACE INTO cargos VALUES (?, ?)", (nombre, monto))
    db.commit()

def borrar_cargo_db(nombre):
    db = get_db()
    db.execute("DELETE FROM cargos WHERE nombre=?", (nombre,))
    db.commit()

def limpiar_datos_prueba():
    db = get_db()
    db.execute("DELETE FROM pagos")
    db.execute("DELETE FROM usuarios")
    db.commit()

init_db()

def _enviar_email_real(destinatario, asunto, texto, pdf_bytes=None, pdf_nombre="recibo.pdf"):
    try:
        msg = MIMEMultipart()
        msg["From"] = EMAIL_REMITENTE
        msg["To"] = destinatario
        msg["Subject"] = asunto
        msg.attach(MIMEText(texto, "plain"))
        if pdf_bytes:
            adj = MIMEApplication(pdf_bytes, _subtype="pdf")
            adj.add_header("Content-Disposition", "attachment", filename=pdf_nombre)
            msg.attach(adj)
        server = smtplib.SMTP("smtp.gmail.com", 587, timeout=20)
        server.starttls()
        server.login(EMAIL_REMITENTE, EMAIL_CLAVE_APP)
        server.send_message(msg)
        server.quit()
        print("Email enviado a " + destinatario)
        return True
    except Exception as e:
        print("Error enviando email: " + str(e))
        return False

def enviar_email(destinatario, asunto, texto, pdf_bytes=None, pdf_nombre="recibo.pdf"):
    threading.Thread(target=_enviar_email_real, args=(destinatario, asunto, texto, pdf_bytes, pdf_nombre), daemon=True).start()
    return True

def generar_pdf_recibo(nombre_jugador, dni, club, total):
    pdf = FPDF()
    pdf.add_page()
    # Logo arriba centrado
    try:
        import os
        if os.path.exists("logo.png"):
            pdf.image("logo.png", x=80, y=6, w=50)
            pdf.ln(52)
    except:
        pass
    pdf.set_font("Arial", "B", 18)
    pdf.cell(0, 15, "LIGA MENDOCINA DE FUTBOL", ln=True, align="C")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "Recibo de Pago - Autogestion", ln=True, align="C")
    pdf.ln(10)
    pdf.cell(0, 10, "Jugador: " + nombre_jugador, ln=True)
    pdf.cell(0, 10, "DNI: " + dni + " - Club: " + club, ln=True)
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 10, "Detalle:", ln=True)
    pdf.set_font("Arial", "", 12)
    try:
        lista = get_cargos_pago()
    except:
        lista = []
    if lista:
        for c in lista:
            pdf.cell(0, 10, " - " + c["nombre"] + ": $" + str(c["monto"]), ln=True)
    else:
        pdf.cell(0, 10, " - Concepto: Inscripcion + Seguro + Carnet", ln=True)
    pdf.cell(0, 10, "Total abonado: $" + str(total), ln=True)
    pdf.cell(0, 10, "Medio: Mercado Pago (DEMO)", ln=True)
    pdf.cell(0, 10, "Estado: HABILITADO", ln=True)
    pdf.cell(0, 10, "Fecha: " + datetime.datetime.now().strftime("%d/%m/%Y %H:%M"), ln=True)
    return bytes(pdf.output())

ESTILO = """
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<style>
body{font-family:Arial,sans-serif;background:#eef2f0;padding:76px 16px 16px;margin:0}
.card{background:white;max-width:780px;margin:12px auto;padding:24px;border-radius:16px;box-shadow:0 4px 20px rgba(0,0,0,.08)}
.logo{width:140px;display:block;margin:0 auto 8px;border-radius:12px}
.logo-mini{width:70px;display:block;margin:0 auto 6px;border-radius:10px}
h1{color:#0a4a2b;text-align:center;font-size:26px} h2{color:#0a4a2b;font-size:22px}
input{width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px}
button{background:#00a650;color:white;border:none;padding:15px;width:100%;border-radius:10px;font-size:16px;cursor:pointer;font-weight:700;margin-top:6px}
.btn-mp{background:#009ee3} .btn-admin{background:#222} .btn-excel{background:#1d6f42} .btn-sec{background:#eee;color:#222} .btn-danger{background:#dc3545}
.cargo{display:flex;justify-content:space-between;padding:14px 4px;border-bottom:1px solid #eee}
.pagado{color:#0a7a2b;font-weight:700}
.badge{padding:8px 14px;border-radius:20px;font-weight:700;display:inline-block;font-size:14px}
.habilitado{background:#d4edda;color:#155724} .pendiente-bg{background:#f8d7da;color:#721c24}
a{color:#009ee3;text-decoration:none} .center{text-align:center} .small{font-size:13px;color:#666}
table{width:100%;border-collapse:collapse;margin-top:15px}
th,td{padding:10px;border-bottom:1px solid #eee;text-align:left;font-size:14px}
th{background:#0a4a2b;color:white}
.kpi{display:flex;gap:10px;margin:15px 0} .kpi div{flex:1;background:#f8f9fa;padding:15px;border-radius:12px;text-align:center} .kpi b{font-size:22px;display:block}
.topbar{display:flex;justify-content:space-between;align-items:center}

/* --- NAVBAR LIGA --- */
.navbar{position:fixed;top:0;left:0;right:0;z-index:1000;background:#0a4a2b;box-shadow:0 2px 12px rgba(0,0,0,.25);padding:0 16px}
.navbar-inner{max-width:900px;margin:0 auto;display:flex;align-items:center;gap:14px;height:60px}
.navbar-brand{display:flex;align-items:center;gap:10px;color:white;text-decoration:none}
.navbar-brand img{width:38px;height:38px;object-fit:cover;border-radius:8px;background:white}
.navbar-brand b{font-size:15px;line-height:1.1;display:block}
.navbar-brand span{font-size:11px;opacity:.85;font-weight:400;display:block}
.navbar-links{display:flex;gap:4px;margin-left:auto;flex-wrap:wrap}
.navbar-links a{color:white;padding:8px 12px;border-radius:8px;font-size:14px;font-weight:700;white-space:nowrap}
.navbar-links a:hover{background:rgba(255,255,255,.15)}
.navbar-links a.btn-nav-salir{background:#dc3545}
.navbar-links a.btn-nav-ingresar{background:#00a650}
@media(max-width:640px){.navbar-brand span{display:none}.navbar-brand b{font-size:13px}.navbar-links a{padding:7px 8px;font-size:12px}body{padding:68px 10px 10px}}

/* PARCHE CAT - alinear botones */
.nav-links, .top-nav { display: flex !important; align-items: center !important; gap: 8px !important; }
.nav-links a, .top-nav a { display: inline-flex !important; align-items: center !important; justify-content: center !important; line-height: 1 !important; padding: 8px 14px !important; margin: 0 !important; box-sizing: border-box; }
.export-row, .botones-exportar { display: flex !important; gap: 10px !important; width: 100% !important; }
.export-row a, .botones-exportar a { flex: 0 0 auto !important; display: flex !important; align-items: center !important; justify-content: center !important; text-align: center !important; padding: 12px 8px !important; margin: 0 !important; box-sizing: border-box; white-space: nowrap; }


/* PARCHE CAT CENTRAR - 3 botones EXPORTAR al centro */
.export-row, .botones-exportar { display: flex !important; justify-content: center !important; align-items: center !important; gap: 10px !important; width: 100% !important; flex-wrap: wrap !important; }
.export-row a, .botones-exportar a { flex: 0 0 auto !important; display: inline-flex !important; align-items: center !important; justify-content: center !important; }

</style>
<nav class="navbar"><div class="navbar-inner">
<a class="navbar-brand" href="/"><img src="/logo.png"><span><b>LIGA MENDOCINA</b><span>Autogestion de Jugadores</span></span></a>
<div class="navbar-links">
<a href="/">Inicio</a>
<a href="/registro">Crear Cuenta</a>
<a href="/login" class="btn-nav-ingresar">Ingresar</a>
<a href="/admin">Admin</a>
<a href="/logout" class="btn-nav-salir">Salir</a>
</div></div></nav>

"""

@app.route("/")
def index():
    return render_template_string(ESTILO + """
    <div class="card center">
        <img src="/logo.png" class="logo">
        <h1>LIGA MENDOCINA DE FUTBOL</h1>
        <p><b>Autogestion de Jugadores</b><br>Paga la inscripcion sin hacer cola</p>
        <a href="/registro"><button>SOY PADRE / MADRE - CREAR MI CUENTA</button></a><br><br>
        <a href="/login"><button style="background:#222">YA TENGO CUENTA - INGRESAR</button></a><br><br>
        <a href="/admin">Acceso Liga / Admin</a>
        <p class="small">Demo - Para probar usa DNI: 50035043 (CHAVEZ)</p>
    </div>
    """)

@app.route("/registro", methods=["GET","POST"])
def registro():
    if request.method == "POST":
        email = request.form["email"].lower().strip()
        dni = request.form["dni"].strip()
        if dni not in JUGADORES:
            return render_template_string(ESTILO + "<div class='card center'><h2>DNI no encontrado</h2><p>El DNI " + dni + " no esta en la base.</p><a href='/registro'><button class='btn-sec'>Volver</button></a></div>")
        # CONTROL DNI UNICO: un DNI = un solo usuario
        db_tmp = get_db()
        existe = db_tmp.execute("SELECT email FROM usuarios WHERE dni_jugador=?", (dni,)).fetchone()
        if existe:
            return render_template_string(ESTILO + "<div class='card center'><h2>⚠️ Ese jugador ya está asociado</h2><p>El DNI <b>" + dni + "</b> ya está asociado al usuario <b>" + existe[0] + "</b>.<br>Si sos el papá/mamá, pedile al Admin que libere ese jugador.<br>Si te equivocaste de DNI, volvé e intentá con otro.</p><a href='/registro'><button class='btn-sec'>Volver a Registro</button></a></div>")
        if get_usuario(email):
            return redirect("/login")
        crear_usuario(email, dni)
        jug = JUGADORES[dni]
        enviar_email(email, "Tu cuenta Liga Mendocina - Bienvenido", "Hola! Tu cuenta fue creada.\nUsuario: " + email + "\nClave: " + dni + "\nJugador: " + jug["nombre"])
        return render_template_string(ESTILO + "<div class='card center'><h2>Cuenta creada!</h2><p>Te enviamos un correo a <b>" + email + "</b></p><p><b>Clave:</b> " + dni + "<br><b>Jugador:</b> " + jug["nombre"] + "</p><a href='/login'><button>INGRESAR AHORA</button></a></div>")
    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini"><h2 class="center">Crear cuenta</h2>
    <form method="post"><label>Tu Email</label><input name="email" type="email" required placeholder="papa@gmail.com">
    <label>DNI del Jugador</label><input name="dni" required placeholder="Ej: 50035043" inputmode="numeric">
    <button>VALIDAR Y CREAR CUENTA</button></form><a href="/">Volver</a></div>
    """)

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].lower().strip()
        clave = request.form["password"].strip()
        # PARCHE CAT: LOGIN CLUB
        try:
            row_club = get_db().execute("SELECT * FROM clubes_usuarios WHERE usuario=?", (email,)).fetchone()
        except Exception:
            row_club = None
        if row_club and row_club["password"] == clave:
            session["club_usuario"] = email
            session["club"] = row_club["club"]
            return redirect("/club/panel")
        u = get_usuario(email)
        if u and u["password"] == clave:
            session["email"] = email
            return redirect("/panel")
        return render_template_string(ESTILO + "<div class='card center'><h2>Datos incorrectos</h2><a href='/login'><button class='btn-sec'>Reintentar</button></a></div>")
    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini"><h2 class="center">Ingresar</h2>
    <form method="post"><label>Usuario (tu email o usuario de club)</label><input name="email" type="text" required>
    <label>Clave</label><input name="password" type="password" required placeholder="Tu DNI">
    <button>INGRESAR</button></form><p style='text-align:center;margin-top:12px'><a href="/olvidado" style='color:#e91e63;font-weight:bold'>¿Olvidaste tu DNI / Clave? Recuperar acá</a></p><a href="/">Volver</a></div>
    """)

@app.route("/panel")
def panel():
    email = session.get("email")
    if not email: return redirect("/login")
    u = get_usuario(email)
    if not u: return redirect("/login")
    jug = JUGADORES[u["dni_jugador"]]
    CARGOS_DYN = get_cargos_pago()
    total = sum(c["monto"] for c in CARGOS_DYN)
    pagado_total = sum(c["monto"] for c in CARGOS_DYN if c["id"] in u["pagos"])
    pendiente = total - pagado_total
    # HISTORIAL DE PAGOS
    historial_html = "<h3>Historial de pagos</h3>"
    if not u['pagos']:
        historial_html += "<p class='small'>Todavia no hay pagos registrados.</p>"
    else:
        historial_html += "<div style='background:#f8f9fa;border-radius:10px;padding:10px;'>"
        for c in get_cargos_pago():
            if c['id'] in u['pagos']:
                historial_html += "<div class='cargo'><span>✅ " + c['nombre'] + " - PAGADO</span><span><b>$" + str(c['monto']) + "</b> <a href='/descargar_recibo/" + str(c['id']) + "'><button style='width:auto;padding:6px 12px;font-size:13px;background:#1d6f42;'>DESCARGAR PDF</button></a></span></div>"
        historial_html += "<p style='text-align:right;'><b>Total pagado: $" + str(pagado_total) + "</b></p>"
        if pendiente == 0:
            historial_html += "<a href='/descargar_recibo_total'><button class='btn-sec'>DESCARGAR COMPROBANTE TOTAL EN PDF</button></a>"
        historial_html += "</div>"
    habilitado = pendiente == 0
    cargos_html = ""
    for c in get_cargos_pago():
        if c["id"] in u["pagos"]:
            estado = '<span class="pagado">PAGADO</span>'
        else:
            estado = '<a href="/pagar/' + str(c["id"]) + '"><button style="width:auto;padding:9px 18px">PAGAR</button></a>'
        cargos_html += '<div class="cargo"><span>' + c["nombre"] + '</span><span><b>$' + str(c["monto"]) + '</b> ' + estado + '</span></div>'
    if habilitado:
        badge = '<span class="badge habilitado">HABILITADO PARA JUGAR</span>'
        btn_todo = ""
    else:
        badge = '<span class="badge pendiente-bg">PENDIENTE DE PAGO</span>'
        btn_todo = '<a href="/pagar_todo"><button class="btn-mp">PAGAR TODO $' + str(total) + ' CON MERCADO PAGO - DEMO</button></a>'
    return render_template_string(ESTILO + "<div class='card'><div class='topbar'><b>Hola, " + email + "</b> <a href='/logout'>Salir</a></div><div class='center'><img src='/logo.png' class='logo-mini'></div><h2 class='center'>" + jug["nombre"] + "</h2><div style='background:#f0f7f2;border:2px solid #0a4a2b;border-radius:14px;padding:16px;margin:14px 0;'><div style='text-align:center;font-size:13px;color:#666;font-weight:700;'>FICHA DEL JUGADOR</div><div style='margin-top:10px;font-size:15px;'><b>JUGADOR:</b> " + jug["nombre"] + "<br><b>DNI:</b> " + u["dni_jugador"] + "<br><b>CLUB:</b> " + jug["club"] + "<br><b>CATEGORIA:</b> " + jug["categoria"] + "<br><b>TORNEO:</b> " + jug["torneo"] + "</div></div><p class='center'>Estado: " + badge + "</p><h3>Cargos a pagar</h3>" + cargos_html + historial_html + "<p style='text-align:right'><b>Total pendiente: $" + str(pendiente) + "</b></p>" + btn_todo + "</div>")

@app.route("/pagar/<int:cid>")
def pagar(cid):
    email = session.get("email")
    if not email: return redirect("/login")
    agregar_pago(email, cid)
    u = get_usuario(email)
    if len(u["pagos"]) == len(get_cargos_pago()):
        jug = JUGADORES[u["dni_jugador"]]
        monto_pago = sum(c["monto"] for c in get_cargos_pago() if c["id"] == cid)
        pdf_bytes = generar_pdf_recibo(jug["nombre"], u["dni_jugador"], jug["club"], monto_pago)
        enviar_email(email, "Recibo de Pago - Liga Mendocina", "Gracias! Pagaste $" + str(monto_pago) + " por " + jug["nombre"] + ". Adjunto tu recibo.", pdf_bytes=pdf_bytes)
    return redirect("/panel")

@app.route("/pagar_todo")
def pagar_todo():
    email = session.get("email")
    if not email: return redirect("/login")
    pagar_todo_db(email)
    u = get_usuario(email)
    jug = JUGADORES[u["dni_jugador"]]
    monto_pago = sum(c["monto"] for c in get_cargos_pago())
    pdf_bytes = generar_pdf_recibo(jug["nombre"], u["dni_jugador"], jug["club"], monto_pago)
    enviar_email(email, "Recibo de Pago - Liga Mendocina", "Gracias! Pagaste $" + str(monto_pago) + " por " + jug["nombre"] + ". Adjunto tu recibo. Ya estas HABILITADO.", pdf_bytes=pdf_bytes)
    return redirect("/panel")




@app.route("/descargar_recibo/<int:cid>")
def descargar_recibo(cid):
    email = session.get("email")
    if not email: return redirect("/login")
    u = get_usuario(email)
    if not u or cid not in u["pagos"]: return redirect("/panel")
    jug = JUGADORES[u["dni_jugador"]]
    cargo = [c for c in get_cargos_pago() if c["id"] == cid][0]
    pdf_bytes = generar_pdf_recibo(jug["nombre"], u["dni_jugador"], jug["club"], cargo["monto"])
    return (pdf_bytes, 200, {"Content-Type": "application/pdf", "Content-Disposition": "attachment; filename=recibo_" + str(cid) + ".pdf"})

@app.route("/descargar_recibo_total")
def descargar_recibo_total():
    email = session.get("email")
    if not email: return redirect("/login")
    u = get_usuario(email)
    jug = JUGADORES[u["dni_jugador"]]
    total_pagado = sum(c["monto"] for c in get_cargos_pago() if c["id"] in u["pagos"])
    pdf_bytes = generar_pdf_recibo(jug["nombre"], u["dni_jugador"], jug["club"], total_pagado)
    return (pdf_bytes, 200, {"Content-Type": "application/pdf", "Content-Disposition": "attachment; filename=recibo_total.pdf"})

@app.route("/admin/usuarios")
def admin_usuarios():
    email = session.get("email")
    if not email: return redirect("/login")
    db = get_db()
    rows = db.execute("SELECT email, dni_jugador, password FROM usuarios ORDER BY email").fetchall()
    html = ESTILO + "<div class=\"card\" style=\"max-width:1050px\"><h2>ABM Usuarios - Papas</h2><p>Total: <b>" + str(len(rows)) + "</b> usuarios</p><input id=\"buscador\" onkeyup=\"filtrar()\" placeholder=\"Buscar por email, DNI o jugador...\" style=\"width:100%;padding:10px;border:1px solid #ddd;border-radius:8px;margin-bottom:12px;font-size:15px\"><p><a href='/admin'>Volver a Admin</a></p><div style=\"overflow-x:auto\">"
    html += "<table id=\"tabla\" style=\"width:100%;border-collapse:collapse;font-size:14px\"><tr style=\"background:#0a4a2b;color:white\"><th style=\"padding:10px;text-align:left\">Email</th><th style=\"padding:10px;text-align:left\">DNI Jugador</th><th style=\"padding:10px;text-align:left\">Jugador</th><th style=\"padding:10px;text-align:left\">Club</th><th style=\"padding:10px\">Pagos</th><th style=\"padding:10px;text-align:left\">Acciones</th></tr>"
    for r in rows:
        em, dni_j, pw = r[0], r[1], r[2]
        jug = JUGADORES.get(dni_j, {})
        nombre = jug.get("nombre","- SIN JUGADOR -")
        club = jug.get("club","-")
        n_pagos = db.execute("SELECT COUNT(*) FROM pagos WHERE email=?", (em,)).fetchone()[0]
        html += f"<tr><td>{em}</td><td>{dni_j}</td><td>{nombre}</td><td>{club}</td><td>{n_pagos}</td>"
        html += f"<td><a href='/admin/usuarios/cambiar/{em}'><button style='background:#2196f3;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer'>Cambiar Jugador</button></a> "
        html += f"<a href='/admin/usuarios/reset/{em}'><button style='background:#ff9800;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer'>Reset 1234</button></a> "
        html += f"<a href='/admin/usuarios/borrar/{em}' onclick=\"return confirm('Dar de Baja {em}? Se borra usuario y sus pagos. Podra registrarse de nuevo.')\"><button style='background:#dc3545;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer'>DAR DE BAJA</button></a></div></td></tr>"
    html += "</table></div><p style=\"font-size:13px;color:#666;margin-top:12px\">Reset: la clave queda en <b>1234</b><br>Dar de Baja: borra usuario + pagos, se puede registrar de nuevo.</p></div><script>function filtrar(){var q=document.getElementById('buscador').value.toLowerCase();var filas=document.querySelectorAll('#tabla tr');for(var i=1;i<filas.length;i++){filas[i].style.display=filas[i].innerText.toLowerCase().includes(q)?'':'none';}}</script>"
    return html

@app.route("/admin/usuarios/borrar/<em>")
def borrar_usuario(em):
    if not session.get("email"): return redirect("/login")
    db = get_db()
    db.execute("DELETE FROM pagos WHERE email=?", (em,))
    db.execute("DELETE FROM usuarios WHERE email=?", (em,))
    db.commit()
    return redirect("/admin/usuarios")

@app.route("/admin/usuarios/reset/<em>")
def reset_usuario(em):
    if not session.get("email"): return redirect("/login")
    db = get_db()
    db.execute("UPDATE usuarios SET password='1234' WHERE email=?", (em,))
    db.commit()
    return redirect("/admin/usuarios")


@app.route("/admin/usuarios/cambiar/<em>", methods=["GET","POST"])
def cambiar_jugador(em):
    if not session.get("email"): return redirect("/login")
    db = get_db()
    if request.method == "POST":
        nuevo_dni = request.form.get("dni_jugador")
        # CONTROL DNI UNICO en Cambiar Jugador
        otro = db.execute("SELECT email FROM usuarios WHERE dni_jugador=? AND email!=?", (nuevo_dni, em)).fetchone()
        if otro:
            return ESTILO + f"<div class='card'><h2>⚠️ DNI ya ocupado</h2><p>El DNI <b>{nuevo_dni}</b> ya está asociado a <b>{otro[0]}</b>.<br>Elegí otro jugador o liberá primero ese usuario.</p><p><a href='/admin/usuarios'>← Volver</a></p></div>"
        db.execute("UPDATE usuarios SET dni_jugador=? WHERE email=?", (nuevo_dni, em))
        db.execute("DELETE FROM pagos WHERE email=?", (em,))
        db.commit()
        return redirect("/admin/usuarios")
    opts = ""
    for dni_j, j in JUGADORES.items():
        opts += f"<option value='{dni_j}'>{j['nombre']} - DNI {dni_j} - {j['club']}</option>"
    return f"<h1>Cambiar Jugador de {em}</h1><p>Ojo: al cambiar de jugador se borran sus pagos, porque son de otro chico.</p><form method='POST'><select name='dni_jugador'>{opts}</select><br><br><button>Guardar</button></form><br><a href='/admin/usuarios'>Volver</a>"

@app.route("/olvidado", methods=["GET","POST"])
def olvidado():
    if request.method == "POST":
        em = request.form.get("email","").strip().lower()
        db = get_db()
        row = db.execute("SELECT dni_jugador FROM usuarios WHERE email=?", (em,)).fetchone()
        if not row:
            return "Ese email no esta registrado. <a href='/login'>Volver</a>"
        db.execute("UPDATE usuarios SET password='1234' WHERE email=?", (em,))
        db.commit()
        jug = JUGADORES.get(row[0], {})
        return ESTILO + f"<div class=\"card\"><h2>✅ Recuperado!</h2><p>Email: {em}<br>DNI asociado: <b>{row[0]}</b><br>Jugador: <b>{jug.get('nombre','-')}</b><br>Tu clave nueva es: <b>1234</b></p><p><a href='/login'>← Ir a Login</a></p></div>"
    msg = ""
    if request.method == "POST":
        email = request.form.get("email","").strip().lower()
        conn = sqlite3.connect(DB)
        cur = conn.cursor()
        cur.execute("SELECT dni, nombre FROM usuarios WHERE email=?", (email,))
        r = cur.fetchone()
        conn.close()
        if r:
            msg = f"Tu DNI es: <b>{r[0]}</b> - Jugador: {r[1]}. Pedile al admin que te resetee la clave."
        else:
            msg = "No encontramos ese email. Verificá que sea el mismo con el que te registraste."
    return f"""{ESTILO}
    <div class="card">
    <h2>🔑 Olvidé mi DNI / Clave</h2>
    <p style="color:#666;font-size:14px">Ingresá el email con el que te registraste y te mostramos tu DNI.</p>
    <form method="post">
    <label>Email:</label><br>
    <input name="email" type="email" required placeholder="ejemplo@gmail.com"><br><br>
    <button class="btn" type="submit">Recuperar</button>
    </form>
    <div style="margin-top:14px">{msg}</div>
    <p style="margin-top:16px"><a href="/login">← Volver a Login</a></p>
    </div>
    """

@app.route("/admin", methods=["GET","POST"])
def admin_login():
    if request.method == "POST":
        if request.form["email"] == ADMIN_USER and request.form["password"] == ADMIN_PASS:
            session["admin"] = True
            return redirect("/admin/panel")
        return render_template_string(ESTILO + "<div class='card center'><h2>Datos incorrectos</h2><a href='/admin'><button class='btn-sec'>Reintentar</button></a></div>")
    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini"><h2 class="center">Panel Liga - Admin</h2>
    <form method="post"><label>Usuario Admin</label><input name="email" required placeholder="admin@liga.com">
    <label>Clave</label><input name="password" type="password" required><button class="btn-admin">ENTRAR AL PANEL</button></form>
    <p class="small">Demo: admin@liga.com / admin123</p><a href="/">Volver</a></div>
    """)


@app.route("/admin/jugadores/exportar")
def exportar_jugadores():
    if not session.get("admin"): return redirect("/admin")
    output = io.StringIO()
    w = csv.writer(output)
    w.writerow(["DNI", "NOMBRE", "CLUB", "CATEGORIA", "TORNEO", "ESTADO_PAGO", "PENDIENTE", "TOTAL_PAGADO"])
    q2 = request.args.get("q", "").strip().upper()
    filtro2 = request.args.get("filtro", "todos")
    _hm = _map_dni_a_habilitado()
    for dni, j in JUGADORES.items():
        if q2 and not (q2 in str(dni) or q2 in j["nombre"].upper() or q2 in j["club"].upper()):
            continue
        _h, _pe, _pa = _hm.get(str(dni),(False,0,0))
        if filtro2 == "habilitados" and not _h: continue
        if filtro2 == "pendientes" and _h: continue
        w.writerow([dni, j["nombre"], j["club"], j["categoria"], j["torneo"], "HABILITADO" if _h else "PENDIENTE", _pe, _pa])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=jugadores_liga.csv"})

@app.route("/admin/panel")
def admin_panel():
    if not session.get("admin"): return redirect("/admin")
    USUARIOS = todos_usuarios()
    total_recaudado = 0
    cant_pagados = 0
    cant_pendientes = 0
    filas = ""
    # --- DASHBOARD ---
    CARGOS_LIST = get_cargos_pago()
    total_por_jugador = sum(cc["monto"] for cc in CARGOS_LIST)
    total_posible = total_por_jugador * len(USUARIOS)
    # por cargo: {id: [cobrado, cantidad]}
    por_cargo_stats = {cc["id"]: {"nombre": cc["nombre"], "monto": cc["monto"], "cobrado": 0, "cant": 0} for cc in CARGOS_LIST}
    for email, u in USUARIOS.items():
        jug = JUGADORES.get(u["dni_jugador"], {"nombre":"?","club":"?","categoria":"?"})
        pagado = sum(c["monto"] for c in get_cargos_pago() if c["id"] in u["pagos"])
        for cc in CARGOS_LIST:
            if cc["id"] in u["pagos"]:
                por_cargo_stats[cc["id"]]["cobrado"] += cc["monto"]
                por_cargo_stats[cc["id"]]["cant"] += 1
        total_recaudado += pagado
        habilitado = pagado == sum(c["monto"] for c in get_cargos_pago())
        if habilitado: cant_pagados += 1
        else: cant_pendientes += 1
        if habilitado:
            badge = '<span class="badge habilitado">HABILITADO</span>'
        else:
            badge = '<span class="badge pendiente-bg">PENDIENTE</span>'
        filas += "<tr><td><b>" + jug["nombre"] + "</b><br><span class='small'>" + email + "</span></td><td>" + jug["club"] + "<br>" + jug["categoria"] + "</td><td>$" + str(pagado) + "</td><td>" + badge + "</td></tr>"
    html = "<div class='card' style='max-width:900px'><div class='center'><img src='/logo.png' class='logo-mini'></div>"
    html += "<div class='topbar'><h2>Panel Liga Mendocina</h2><a href='/admin/salir'>Salir</a></div>"
    total_pendiente = total_posible - total_recaudado
    pct = round(total_recaudado * 100 / total_posible) if total_posible > 0 else 0
    html += "<div class='kpi'><div><span>RECAUDADO</span><b>$" + str(total_recaudado) + "</b></div><div><span>TOTAL PENDIENTE</span><b style='color:#c0392b'>$" + str(total_pendiente) + "</b></div><div><span>HABILITADOS</span><b>" + str(cant_pagados) + "</b></div><div><span>PENDIENTES</span><b>" + str(cant_pendientes) + "</b></div></div>"
    html += "<div style='background:#eee;border-radius:10px;height:22px;margin:10px 0;overflow:hidden'><div style='background:#2e7d32;height:22px;border-radius:8px;width:" + str(pct) + "%'></div></div><p class='center' style='margin:2px 0;font-weight:bold'>" + str(pct) + "% cobrado</p><p class='small center'>Total posible: $" + str(total_posible) + " | $" + str(total_por_jugador) + " por jugador x " + str(len(USUARIOS)) + " usuarios</p>"
    # Tabla por cargo
    filas_cargo = ""
    for cid, d in por_cargo_stats.items():
        pendiente_cargo = d["monto"] * len(USUARIOS) - d["cobrado"]
        filas_cargo += "<tr><td>" + d["nombre"] + "</td><td>$" + str(d["monto"]) + "</td><td style='color:green'><b>$" + str(d["cobrado"]) + "</b></td><td style='color:#c0392b'><b>$" + str(pendiente_cargo) + "</b></td><td>" + str(d["cant"]) + " / " + str(len(USUARIOS)) + "</td></tr>"
    # --- GRAFICO TORTA ---
    colores = ["#2e7d32", "#ff7f00", "#007bff", "#6a1b9a", "#e91e63", "#00bcd4"]
    total_para_torta = total_recaudado if total_recaudado > 0 else total_pendiente
    titulo_torta = "Recaudado por Cargo" if total_recaudado > 0 else "Pendiente por Cargo (todavia no hay cobros)"
    # Construir conic-gradient
    grad_parts = []
    leyenda = ""
    acumulado = 0
    idx_color = 0
    for cid, d in por_cargo_stats.items():
        valor = d["cobrado"] if total_recaudado > 0 else (d["monto"] * len(USUARIOS) - d["cobrado"])
        if total_para_torta == 0:
            pct_cargo = 0
        else:
            pct_cargo = round(valor * 100 / total_para_torta)
        color = colores[idx_color % len(colores)]
        idx_color += 1
        inicio = acumulado
        acumulado += pct_cargo
        # Evitar que el ultimo no cierre a 100 por redondeo
        if list(por_cargo_stats.keys())[-1] == cid:
            acumulado = 100
        grad_parts.append(f"{color} {inicio}% {acumulado}%")
        leyenda += f"<div style='display:flex;align-items:center;gap:8px;margin:4px 0'><span style='width:14px;height:14px;background:{color};border-radius:3px;display:inline-block'></span><b>{d['nombre']}</b>: ${valor} ({pct_cargo}%)</div>"
    grad_str = ", ".join(grad_parts) if grad_parts else "#eee 0% 100%"
    html += f"<h3 style='margin-top:20px'>🍩 {titulo_torta}</h3><div style='display:flex;gap:20px;align-items:center;flex-wrap:wrap;justify-content:center;background:#fafafa;border-radius:12px;padding:15px'><div style='position:relative;width:170px;height:170px;flex-shrink:0'><div style='width:170px;height:170px;border-radius:50%;background:conic-gradient({grad_str});box-shadow:0 2px 10px rgba(0,0,0,0.15)'></div><div style='position:absolute;top:50%;left:50%;transform:translate(-50%,-50%);width:95px;height:95px;background:white;border-radius:50%;display:flex;flex-direction:column;align-items:center;justify-content:center;box-shadow:inset 0 0 8px rgba(0,0,0,0.08)'><span style='font-size:11px;color:#888'>TOTAL</span><b style='font-size:17px'>${total_para_torta}</b><span style='font-size:10px;color:#888'>{len(USUARIOS)} jugador(s)</span></div></div><div>{leyenda}</div></div>"
    html += "<h3 style='margin-top:15px'>Por Cargo</h3><table><tr><th>Cargo</th><th>Monto</th><th>Cobrado</th><th>Pendiente</th><th>Pagaron</th></tr>" + filas_cargo + "</table><br>" 
    html += "<div style='display:flex;gap:8px;flex-wrap:wrap;justify-content:center;align-items:center;width:100%'><a href='/admin/exportar?filtro=todos'><button class='btn-excel'>EXPORTAR TODOS A EXCEL</button></a><a href='/admin/exportar?filtro=habilitados'><button style='background:#16a34a'>EXPORTAR SOLO HABILITADOS</button></a><a href='/admin/exportar?filtro=pendientes'><button style='background:#dc2626'>EXPORTAR SOLO CON DEUDA</button></a></div>"
    html += "<a href='/admin/jugadores'><button style='background:#0a4a2b'>+ CARGAR NUEVO JUGADOR</button></a>"
    html += "<a href='/admin/clubes'><button style='background:#6a1b9a'>CLUBES - ALTA / BAJA / MODIFICAR</button></a>"
    html += "<a href='/admin/categorias'><button style='background:#007bff'>CATEGORIAS - ALTA / BAJA / MODIFICAR</button></a>"
    html += "<a href='/admin/cargos'><button style='background:#ff7f00'>CARGOS - INSCRIPCION / SEGURO / CARNET</button></a>"
    html += "<a href='/admin/torneos'><button style='background:#0a2452'>TORNEOS</button></a>"
    html += "<a href='/admin/temporadas'><button style='background:#009688'>TEMPORADAS</button></a>"
    html += "<a href='/admin/tipos'><button style='background:#795548'>TIPO COMPETICION</button></a>"
    html += "<a href='/admin/niveles'><button style='background:#607d8b'>NIVEL</button></a>"
    html += "<a href='/admin/importar'><button style='background:#ff5722'>IMPORTAR EXCEL 9000</button></a>"
    html += "<a href='/admin/tutores'><button style='background:#e91e63'>TUTORES - PAPAS / VER JUGADORES</button></a>"
    html += "<a href='/admin/limpiar'><button class='btn-danger'>LIMPIAR DATOS DE PRUEBA</button></a>"
    html += "<p class='small center'>Limpiar borra todas las cuentas y deja el panel en cero. Guardado en SQLite.</p>"
    html += "<table><tr><th>Jugador / Padre</th><th>Club</th><th>Pagado</th><th>Estado</th></tr>" + filas + "</table></div>"
    return render_template_string(ESTILO + html)

@app.route("/admin/limpiar", methods=["GET","POST"])
def admin_limpiar():
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        limpiar_datos_prueba()
        return redirect("/admin/panel")
    return render_template_string(ESTILO + """
    <div class="card center"><img src="/logo.png" class="logo-mini">
    <div style="font-size:48px">ADVERTENCIA</div>
    <h2>Limpiar datos de prueba?</h2>
    <p>Se van a borrar <b>TODAS</b> las cuentas creadas y todos los pagos.<br>El panel queda totalmente vacio, en cero.</p>
    <form method="post"><button class="btn-danger">SI, LIMPIAR TODO</button></form>
    <a href="/admin/panel"><button class="btn-sec">Cancelar - Volver al panel</button></a></div>
    """)

@app.route("/admin/exportar")
def admin_exportar():
    import csv, io
    from flask import Response
    filtro2 = request.args.get("filtro", "todos")
    q2 = request.args.get("q", "").strip().upper()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Jugador","DNI","Club","Categoria","Email Padre","Monto Pagado","Estado"])
    hab_map = _map_dni_a_habilitado()
    # email por dni_jugador
    email_por_dni = {}
    try:
        for email, u in todos_usuarios().items():
            email_por_dni[str(u.get("dni_jugador",""))] = email
    except: pass
    for dni, jug in JUGADORES.items():
        hab, pagado, total = hab_map.get(str(dni), (False, 0, 0))
        estado = "HABILITADO" if hab else "PENDIENTE"
        if filtro2 == "habilitados" and estado != "HABILITADO": continue
        if filtro2 == "pendientes" and estado != "PENDIENTE": continue
        if q2 and not (q2 in str(dni) or q2 in str(jug.get("nombre","")).upper() or q2 in str(jug.get("club","")).upper()): continue
        writer.writerow([jug.get("nombre",""), dni, jug.get("club",""), jug.get("categoria",""), email_por_dni.get(str(dni),""), pagado, estado])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=pagos.csv"})


def _map_dni_a_habilitado():
    try:
        cargos = get_cargos_pago()
    except:
        cargos = []
    total = sum(c["monto"] for c in cargos)
    mapa = {}
    # primero todos los JUGADORES en 0
    for dni in JUGADORES.keys():
        mapa[str(dni)] = (False, 0, total)
    # despues pisar con lo que realmente pago cada USUARIO
    try:
        for u in todos_usuarios().values():
            dni_j = str(u.get("dni_jugador",""))
            if not dni_j: continue
            ids_pagos = u.get("pagos", []) or []
            pagado = sum(c["monto"] for c in cargos if c["id"] in ids_pagos)
            hab = (pagado >= total and total > 0)
            mapa[dni_j] = (hab, pagado, total)
    except Exception as e:
        print("map habilitado error", e)
    return mapa


@app.route("/admin/jugadores", methods=["GET","POST"])
def admin_jugadores():

    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        dni = request.form["dni"].strip()
        nombre = request.form["nombre"].strip().upper()
        club = request.form["club"].strip().upper()
        categoria = request.form["categoria"].strip().upper()
        torneo = request.form["torneo"].strip().upper()
        if dni and nombre:
            agregar_jugador_db(dni, nombre, club, categoria, torneo)
            return redirect("/admin/jugadores")

    # --- BUSCADOR + PAGINADOR ---
    q = request.args.get("q", "").strip().upper()
    filtro = request.args.get("filtro", "todos")
    page = request.args.get("page", 1, type=int)
    per_page = 10

    # Filtrar por DNI, Nombre o Club
    todos = list(JUGADORES.items())
    # ORDEN BAJA: ACTIVOS primero por nombre, DE BAJA al final
    todos.sort(key=lambda x: (1 if x[1].get("estado") == "BAJA" else 0, x[1].get("nombre", "")))
    if q:
        todos = [(dni, j) for dni, j in todos if q in dni or q in j["nombre"].upper() or q in j["club"].upper()]
    _hab_map = _map_dni_a_habilitado()
    if filtro == "habilitados":
        todos = [(dni, j) for dni, j in todos if _hab_map.get(str(dni),(False,0,0))[0]]
    elif filtro == "pendientes":
        todos = [(dni, j) for dni, j in todos if not _hab_map.get(str(dni),(False,0,0))[0]]

    total = len(todos)
    total_pages = (total + per_page - 1) // per_page
    if page < 1: page = 1
    if page > total_pages and total_pages > 0: page = total_pages
    inicio = (page - 1) * per_page
    pagina_actual = todos[inicio:inicio + per_page]

    clubes = get_clubes()
    cats_list = get_categorias()
    opciones_categorias = ""
    for cc in cats_list:
        opciones_categorias += "<option value=\"" + cc + "\">" + cc + "</option>"

    opciones_clubes = ""
    for c in clubes:
        opciones_clubes += "<option value=\"" + c + "\">" + c + "</option>"

    # Torneos para desplegable
    try:
        _db = sqlite3.connect("liga.db")
        _lista = [r[0] for r in _db.execute("SELECT nombre FROM torneos ORDER BY nombre")]
        _db.close()
    except Exception:
        _lista = []
    for _tt in sorted(set([j.get("torneo","") for j in JUGADORES.values() if j.get("torneo")])):
        _lista.append(_tt) if _tt not in _lista else None
    opciones_torneos = ""
    for _tt in _lista:
        opciones_torneos += "<option value=\"" + _tt + "\">" + _tt + "</option>"
    filas_j = ""
    for dni, j in pagina_actual:
        _es_baja = (j.get("estado") == "BAJA")
        _badge = " <span style='background:#dc3545;color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:900;white-space:nowrap;display:inline-block'>DE BAJA</span>" if _es_baja else " <span style='background:#28a745;color:white;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:900;white-space:nowrap;display:inline-block'>ACTIVO</span>"
        _estilo = " style='background:#ffe0e0'" if _es_baja else ""
        _boton_estado = "<a href='/admin/jugadores/reactivar/" + dni + "'><button style='width:auto;padding:7px 12px;background:#28a745'>✅ REACTIVAR</button></a>" if _es_baja else "<a href='/admin/jugadores/borrar/" + dni + "'><button style='width:auto;padding:7px 12px;background:#dc3545'>DAR DE BAJA</button></a>"
        _hab, _pen, _pag = _hab_map.get(str(dni),(False,0,0))
        _badge_pago = "✅ HABILITADO" if _hab else "❌ PENDIENTE $" + str(_pen)
        filas_j += "<tr" + _estilo + "><td><b>" + j["nombre"] + "</b>" + _badge + "<br><span class='small'>DNI: " + dni + "</span></td><td>" + j["club"] + "</td><td>" + j["categoria"] + "</td><td>" + j["torneo"] + "</td><td>" + _badge_pago + "</td><td><a href='/admin/jugadores/editar/" + dni + "'><button style='width:auto;padding:7px 12px;background:#f0ad4e'>EDITAR</button></a> " + _boton_estado + "</td></tr>"

    # Links paginador
    paginador = ""
    if total_pages > 1:
        paginador += "<div style='display:flex;gap:8px;justify-content:center;margin:20px 0;align-items:center;flex-wrap:wrap;'>"
        if page > 1:
            paginador += "<a href='/admin/jugadores?q=" + q + "&filtro=" + filtro + "&page=" + str(page-1) + "'><button style='width:auto;padding:8px 14px;background:#eee;color:#333'>⬅ Anterior</button></a>"
        paginador += "<span><b>Página " + str(page) + " de " + str(total_pages) + "</b> - " + str(total) + " encontrados</span>"
        if page < total_pages:
            paginador += "<a href='/admin/jugadores?q=" + q + "&filtro=" + filtro + "&page=" + str(page+1) + "'><button style='width:auto;padding:8px 14px;background:#eee;color:#333'>Siguiente ➡</button></a>"
        paginador += "</div>"

    return render_template_string(ESTILO + """
    <div class="card" style="max-width:900px"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2><div style="display:flex;justify-content:space-between;align-items:center;">
  <h3 style="margin:0;">Cargar Nuevo Jugador</h3>
  <button type="button" id="btn_toggle_form" onclick="toggleForm()" style="background:#0f5132;color:white;border:none;padding:8px 14px;border-radius:8px;cursor:pointer;font-weight:bold;">+ Cargar Nuevo</button>
</div>
</h2><a href="/admin/panel">Volver al panel</a></div>
    <div id="form_carga" style="display:none;"><p class="small">El admin carga el jugador, despues el padre se registra con ese DNI.</p>
    <form method="post">
      <label>DNI del Jugador</label><input name="dni" required placeholder="Ej: 51234567" inputmode="numeric">
      <label>Nombre y Apellido</label><input name="nombre" required placeholder="Ej: JUAN PEREZ">
      <label>Club</label><select name="club" required style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px">""" + opciones_clubes + """</select><p class="small">Si el club no esta en la lista, cargalo primero en <a href="/admin/clubes"><b>CLUBES</b></a>.</p>
      <label>Categoria</label><select name="categoria" required style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px">""" + opciones_categorias + """</select>
      <label>Torneo</label><input list="lista_torneos" name="torneo" required placeholder="Escribi o elegi torneo..." value="COPA PLATA 8VA" style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px"><datalist id="lista_torneos">""" + opciones_torneos + """</datalist><p class="small">Si el torneo no esta, cargalo en <a href="/admin/torneos"><b>TORNEOS</b></a>.</p>
      <button>GUARDAR JUGADOR</button>
    </form>
    </div>
    <script>
    function toggleForm(){
      var f = document.getElementById("form_carga");
      var b = document.getElementById("btn_toggle_form");
      if(f.style.display === "none"){
        f.style.display = "block";
        b.textContent = "- Ocultar Formulario";
      } else {
        f.style.display = "none";
        b.textContent = "+ Cargar Nuevo";
      }
    }
    </script>
    <h3>Jugadores cargados (""" + str(total) + """)</h3>
    <div style="text-align:center; margin:15px 0;">
      <a href="/admin/jugadores/exportar?q=""" + q + """&filtro=""" + filtro + """" style="display:block; background:#28a745; color:white; padding:18px; border-radius:12px; font-weight:900; text-decoration:none; font-size:18px;">⬇ EXPORTAR A EXCEL - BAJAR JUGADORES FILTRADOS</a>
    </div>
    <div style="display:flex;gap:8px;flex-wrap:wrap;margin:10px 0;"><a href="/admin/jugadores?filtro=todos"><button style="width:auto;padding:10px 14px;background:#0a4a2b">Todos</button></a><a href="/admin/jugadores?filtro=habilitados"><button style="width:auto;padding:10px 14px;background:#16a34a">✅ Solo Habilitados</button></a><a href="/admin/jugadores?filtro=pendientes"><button style="width:auto;padding:10px 14px;background:#dc2626">❌ Con Deuda</button></a></div><form method="get" style="display:flex;gap:10px;margin:15px 0;"><input type="hidden" name="filtro" value="" + filtro + "">
      <input name="q" value=\"""" + q + """\" placeholder="Buscar por DNI, Nombre o Club..." style="flex:1;">
      <button style="width:auto;padding:12px 20px;">🔍 BUSCAR</button>
    </form>
    """ + paginador + """
<script>
document.querySelectorAll("table tr").forEach(function(tr){
  if(tr.textContent.includes("DE BAJA")){
    tr.style.background = "#ffe0e0";
    tr.style.opacity = "0.85";
  }
});
</script>

    <table><tr><th>Jugador / DNI</th><th>Club</th><th>Categoria</th><th>Torneo</th><th>Estado Pago</th><th>Acciones</th></tr>""" + filas_j + """</table>
    """ + paginador + """
    </div>
    """)


@app.route("/admin/jugadores/editar/<dni>", methods=["GET","POST"])
def admin_editar_jugador(dni):
    if not session.get("admin"): return redirect("/admin")
    j = JUGADORES.get(dni)
    if not j: return redirect("/admin/jugadores")
    clubes_e = get_clubes()
    cats_e = get_categorias()
    opciones_categorias = ""
    for cc in cats_e:
        selc = " selected" if cc == j["categoria"] else ""
        opciones_categorias += "<option value=\"" + cc + "\"" + selc + ">" + cc + "</option>"
    opciones_e = ""
    for c in clubes_e:
        sel = " selected" if c == j["club"] else ""
        opciones_e += "<option value=\"" + c + "\"" + sel + ">" + c + "</option>"
    # Torneos editar
    try:
        _db2 = sqlite3.connect("liga.db")
        _lista_e = [r[0] for r in _db2.execute("SELECT nombre FROM torneos ORDER BY nombre")]
        _db2.close()
    except Exception:
        _lista_e = []
    for _tt in sorted(set([jj.get("torneo","") for jj in JUGADORES.values() if jj.get("torneo")])):
        _lista_e.append(_tt) if _tt not in _lista_e else None
    if j.get("torneo") and j["torneo"] not in _lista_e:
        _lista_e.append(j["torneo"])
    opciones_torneos_e = ""
    for _tt in _lista_e:
        _selt = " selected" if _tt == j["torneo"] else ""
        opciones_torneos_e += "<option value=\"" + _tt + "\"" + _selt + ">" + _tt + "</option>"
    if request.method == "POST":
        nombre = request.form["nombre"].strip().upper()
        club = request.form["club"].strip().upper()
        categoria = request.form["categoria"].strip().upper()
        torneo = request.form["torneo"].strip().upper()
        agregar_jugador_db(dni, nombre, club, categoria, torneo)
        return redirect("/admin/jugadores")
    # --- PAGOS DE ESTE JUGADOR (directo a DB) ---
    pagos_html = ""
    try:
        _dbp = sqlite3.connect("liga.db")
        _dbp.row_factory = sqlite3.Row
        # 1. Tutores de este DNI
        try:
            _tutores = [r["email"] for r in _dbp.execute("SELECT email FROM usuarios WHERE dni_jugador=?", (str(dni),)).fetchall()]
        except Exception:
            _tutores = [r[0] for r in _dbp.execute("SELECT email FROM usuarios WHERE dni_jugador=?", (str(dni),)).fetchall()]
        # 2. Cargos / conceptos de pago - MISMO que usa el panel del tutor
        # OJO: id = posicion+1, no existe en la tabla. Por eso hay que usar get_cargos_pago()
        _cargos = get_cargos_pago()
        if not _tutores:
            pagos_html = "<div style='background:#fff7ed;border:1px solid #fed7aa;border-radius:10px;padding:14px;margin-top:16px'><h3 style='margin:0 0 8px'>Pagos de este jugador</h3><p class='small' style='margin:0'>Este jugador todavia no tiene tutor asignado, por eso no hay pagos.</p></div>"
        else:
            for em_tutor in _tutores:
                _ids = set([r[0] for r in _dbp.execute("SELECT cargo_id FROM pagos WHERE email=?", (em_tutor,)).fetchall()])
                _filas = ""
                _tot_pag = 0
                _tot_pen = 0
                for c in _cargos:
                    _cid = c.get("id")
                    _monto = c.get("monto", 0) or 0
                    _nombre = c.get("nombre") or c.get("descripcion") or c.get("concepto") or f"Cargo {_cid}"
                    _pagado = _cid in _ids
                    if _pagado: _tot_pag += _monto
                    else: _tot_pen += _monto
                    _badge = "<span style='background:#dcfce7;color:#166534;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700'>PAGADO</span>" if _pagado else "<span style='background:#fee2e2;color:#991b1b;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700'>PENDIENTE</span>"
                    _filas += f"<tr><td style='padding:8px'>{_nombre}</td><td style='padding:8px'>${_monto}</td><td style='padding:8px'>{_badge}</td></tr>"
                if not _filas:
                    _filas = "<tr><td colspan='3' style='padding:8px'>No hay cargos definidos.</td></tr>"
                pagos_html += f"<div style='background:#f8fafc;border:1px solid #e2e8f0;border-radius:10px;padding:14px;margin-top:16px'><h3 style='margin:0 0 8px'>Pagos de este jugador</h3><p class='small' style='margin:0 0 8px'>Tutor: <b>{em_tutor}</b><br>Total pagado: <b style='color:#16a34a'>${_tot_pag}</b> &nbsp;|&nbsp; Pendiente: <b style='color:#dc2626'>${_tot_pen}</b></p><table style='width:100%;border-collapse:collapse'><tr style='background:#eef2f7'><th style='text-align:left;padding:8px'>Concepto</th><th style='text-align:left;padding:8px'>Monto</th><th style='text-align:left;padding:8px'>Estado</th></tr>{_filas}</table></div>"
        _dbp.close()
    except Exception as e:
        pagos_html = f"<div style='background:#fef2f2;border:1px solid #fecaca;border-radius:10px;padding:12px;margin-top:16px'><p class='small'>No se pudieron cargar los pagos: {e}</p></div>"

    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>Editar Jugador</h2><a href="/admin/jugadores">Volver</a></div>
    <p>DNI: <b>""" + dni + """</b> (no se puede cambiar)</p>
    <form method="post">
      <label>Nombre y Apellido</label><input name="nombre" required value='""" + j["nombre"] + """'>
      <label>Club</label><select name="club" required style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px">""" + opciones_e + """</select>
      <label>Categoria</label><select name="categoria" required style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px">""" + opciones_categorias + """</select>
      <label>Torneo</label><input list="lista_torneos_e" name="torneo" required value='""" + j["torneo"] + """' style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px;font-size:16px"><datalist id="lista_torneos_e">""" + opciones_torneos_e + """</datalist>
      <button>GUARDAR CAMBIOS</button>
    </form>""" + pagos_html + """
    </div>
    """)

@app.route("/admin/jugadores/borrar/<dni>", methods=["GET","POST"])
def admin_borrar_jugador(dni):
    if not session.get("admin"): return redirect("/admin")
    j = JUGADORES.get(dni)
    if not j: return redirect("/admin/jugadores")
    if request.method == "POST":
        borrar_jugador_db(dni)
        return redirect("/admin/jugadores")
    return render_template_string(ESTILO + """
    <div class="card center"><img src="/logo.png" class="logo-mini">
    <div style="font-size:48px">ADVERTENCIA</div>
    <h2>Dar de Baja jugador?</h2>
    <p><b>""" + j["nombre"] + """</b><br>DNI: """ + dni + """<br>""" + j["club"] + """ - """ + j["categoria"] + """</p>
    <p class="small" style="color:#c00"><b>NO se borra nada.</b><br>El jugador pasa a estado DE BAJA, se va al final de la lista en rojo.<br>Sus pagos y su historial quedan guardados. Se puede reactivar cuando quieras.</p>
    <form method="post"><button class="btn-danger">SI, DAR DE BAJA</button></form>
    <a href="/admin/jugadores"><button class="btn-sec">Cancelar</button></a></div>
    """)

@app.route("/admin/jugadores/reactivar/<dni>", methods=["GET","POST"])
def admin_reactivar_jugador(dni):
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        reactivar_jugador_db(dni)
        return redirect("/admin/jugadores")
    j = JUGADORES.get(dni)
    if not j: return redirect("/admin/jugadores")
    return render_template_string(ESTILO + """
    <div class="card center"><img src="/logo.png" class="logo-mini">
    <div style="font-size:48px">✅</div>
    <h2>¿Reactivar jugador?</h2>
    <p><b>""" + j["nombre"] + """</b><br>DNI: """ + dni + """<br>""" + j["club"] + """ - """ + j["categoria"] + """</p>
    <p class="small">Vuelve a aparecer como ACTIVO en la lista. Su historial y pagos NO se tocan.</p>
    <form method="post"><button style="background:#28a745;color:white;border:none;padding:12px 22px;border-radius:8px;cursor:pointer;font-weight:900">SI, REACTIVAR</button></form>
    <a href="/admin/jugadores"><button class="btn-sec">Cancelar</button></a></div>
    """)

@app.route("/admin/clubes", methods=["GET","POST"])
def admin_clubes():
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        nombre = request.form["nombre"].strip().upper()
        if nombre:
            agregar_club_db(nombre)
            return redirect("/admin/clubes")
    clubes = get_clubes()
    cats_list = get_categorias()
    opciones_categorias = ""
    for cc in cats_list:
        opciones_categorias += "<option value=\"" + cc + "\">" + cc + "</option>"
    # alias para editar, mismo nombre

    filas_c = ""
    for c in clubes:
        cant = len([j for j in JUGADORES.values() if j["club"] == c])
        uc = get_usuario_club(c)
        if uc:
            btn_club = "<a href='/admin/clubes/usuario/" + c + "'><button style='width:auto;padding:7px 12px;background:#6a1b9a'>USUARIO: " + uc["usuario"] + "</button></a>"
        else:
            btn_club = "<a href='/admin/clubes/usuario/" + c + "'><button style='width:auto;padding:7px 12px;background:#6a1b9a'>GENERAR USUARIO</button></a>"
        filas_c += "<tr><td><b>" + c + "</b></td><td>" + str(cant) + " jugadores</td><td><a href='/admin/clubes/editar/" + c + "'><button style='width:auto;padding:7px 12px;background:#f0ad4e'>EDITAR</button></a> <a href='/admin/clubes/borrar/" + c + "'><button style='width:auto;padding:7px 12px;background:#dc3545'>DAR DE BAJA</button></a> " + btn_club + "</td></tr>"
    return render_template_string(ESTILO + """
    <div class="card" style="max-width:800px"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>ABM Clubes</h2><a href="/admin/panel">Volver al panel</a></div>
    <p class="small">El club se carga una sola vez aca. Despues al cargar jugadores solo se elige de la lista.</p>
    <form method="post"><label>Nombre del Club</label><input name="nombre" required placeholder="Ej: GODOY CRUZ ANTONIO TOMBA">
    <button style="background:#6a1b9a">+ AGREGAR CLUB</button></form>
    <h3>Clubes cargados (""" + str(len(clubes)) + """)</h3>
    <table><tr><th>Club</th><th>Jugadores</th><th>Acciones</th></tr>""" + filas_c + """</table>
    <a href="/admin/jugadores"><button>CARGAR JUGADORES CON ESTOS CLUBES</button></a></div>
    """)

@app.route("/admin/clubes/editar/<path:nombre>", methods=["GET","POST"])
def admin_editar_club(nombre):
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        nuevo = request.form["nombre"].strip().upper()
        if nuevo and nuevo != nombre:
            db = get_db()
            db.execute("INSERT OR IGNORE INTO clubes VALUES (?)", (nuevo,))
            db.execute("UPDATE jugadores SET club=? WHERE club=?", (nuevo, nombre))
            db.execute("DELETE FROM clubes WHERE nombre=?", (nombre,))
            db.commit()
            for dni0, j0 in JUGADORES.items():
                if j0["club"] == nombre:
                    j0["club"] = nuevo
        return redirect("/admin/clubes")
    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>Editar Club</h2><a href="/admin/clubes">Volver</a></div>
    <p class="small">Si cambias el nombre, se actualiza en todos los jugadores de ese club.</p>
    <form method="post"><label>Nombre del Club</label><input name="nombre" required value='""" + nombre + """'>
    <button style="background:#6a1b9a">GUARDAR CAMBIOS</button></form></div>
    """)

@app.route("/admin/clubes/borrar/<nombre>", methods=["GET","POST"])
def admin_borrar_club(nombre):
    if not session.get("admin"): return redirect("/admin")
    cant = len([j for j in JUGADORES.values() if j["club"] == nombre])
    if request.method == "POST":
        borrar_club_db(nombre)
        return redirect("/admin/clubes")
    return render_template_string(ESTILO + """
    <div class="card center"><img src="/logo.png" class="logo-mini">
    <div style="font-size:48px">ADVERTENCIA</div>
    <h2>Dar de Baja club?</h2><p><b>""" + nombre + """</b><br>""" + str(cant) + """ jugadores usan este club.</p>
    <p class="small">El club se borra de la lista, los jugadores quedan con el nombre viejo pero no podran elegirlo de nuevo.</p>
    <form method="post"><button class="btn-danger">SI, DAR DE BAJA</button></form>
    <a href="/admin/clubes"><button class="btn-sec">Cancelar</button></a></div>
    """)

@app.route("/admin/categorias", methods=["GET","POST"])
def admin_categorias():
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        nombre = request.form["nombre"].strip().upper()
        if nombre:
            agregar_categoria_db(nombre)
            return redirect("/admin/categorias")
    cats = get_categorias()
    filas = ""
    for c in cats:
        cant = len([j for j in JUGADORES.values() if j["categoria"] == c])
        filas += "<tr><td><b>" + c + "</b></td><td>" + str(cant) + " jugadores</td><td><a href='/admin/categorias/editar/" + c + "'><button style='width:auto;padding:7px 12px;background:#f0ad4e'>EDITAR</button></a> <a href='/admin/categorias/borrar/" + c + "'><button style='width:auto;padding:7px 12px;background:#dc3545'>DAR DE BAJA</button></a></div></td></tr>"
    return render_template_string(ESTILO + """
    <div class="card" style="max-width:800px"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>ABM Categorias</h2><a href="/admin/panel">Volver al panel</a></div>
    <p class="small">Ej: U-11, U-13, U-15. Se carga una vez aca, despues se elige de la lista.</p>
    <form method="post"><label>Nombre de Categoria</label><input name="nombre" required placeholder="Ej: U-17">
    <button style="background:#007bff">+ AGREGAR CATEGORIA</button></form>
    <h3>Categorias (""" + str(len(cats)) + """)</h3>
    <table><tr><th>Categoria</th><th>Jugadores</th><th>Acciones</th></tr>""" + filas + """</table></div>
    """)

@app.route("/admin/categorias/editar/<path:nombre>", methods=["GET","POST"])
def admin_editar_categoria(nombre):
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        nuevo = request.form["nombre"].strip().upper()
        if nuevo and nuevo != nombre:
            db = get_db()
            db.execute("INSERT OR IGNORE INTO categorias VALUES (?)", (nuevo,))
            db.execute("UPDATE jugadores SET categoria=? WHERE categoria=?", (nuevo, nombre))
            db.execute("DELETE FROM categorias WHERE nombre=?", (nombre,))
            db.commit()
            for dni0, j0 in JUGADORES.items():
                if j0["categoria"] == nombre:
                    j0["categoria"] = nuevo
        return redirect("/admin/categorias")
    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>Editar Categoria</h2><a href="/admin/categorias">Volver</a></div>
    <form method="post"><label>Nombre</label><input name="nombre" required value='""" + nombre + """'>
    <button style="background:#007bff">GUARDAR</button></form></div>
    """)

@app.route("/admin/categorias/borrar/<nombre>", methods=["GET","POST"])
def admin_borrar_categoria(nombre):
    if not session.get("admin"): return redirect("/admin")
    cant = len([j for j in JUGADORES.values() if j["categoria"] == nombre])
    if request.method == "POST":
        borrar_categoria_db(nombre)
        return redirect("/admin/categorias")
    return render_template_string(ESTILO + """
    <div class="card center"><img src="/logo.png" class="logo-mini">
    <div style="font-size:48px">ADVERTENCIA</div><h2>Dar de Baja categoria?</h2><p><b>""" + nombre + """</b><br>""" + str(cant) + """ jugadores.</p>
    <form method="post"><button class="btn-danger">SI, DAR DE BAJA</button></form>
    <a href="/admin/categorias"><button class="btn-sec">Cancelar</button></a></div>
    """)

@app.route("/admin/cargos", methods=["GET","POST"])
def admin_cargos():
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        nombre = request.form["nombre"].strip().upper()
        try: monto = int(request.form["monto"])
        except: monto = 0
        if nombre:
            agregar_cargo_db(nombre, monto)
            return redirect("/admin/cargos")
    cargos = get_cargos()
    filas = ""
    for n, m in cargos:
        filas += "<tr><td><b>" + n + "</b></td><td>$" + f"{m:,}".replace(",", ".") + "</td><td><a href='/admin/cargos/editar/" + n + "'><button style='width:auto;padding:7px 12px;background:#f0ad4e'>EDITAR</button></a> <a href='/admin/cargos/borrar/" + n + "'><button style='width:auto;padding:7px 12px;background:#dc3545'>DAR DE BAJA</button></a></div></td></tr>"
    return render_template_string(ESTILO + """
    <div class="card" style="max-width:800px"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>ABM Cargos</h2><a href="/admin/panel">Volver al panel</a></div>
    <p class="small">Aca definis que se cobra: Inscripcion, Seguro, Carnet, y cuanto sale cada uno.</p>
    <form method="post"><label>Nombre del Cargo</label><input name="nombre" required placeholder="Ej: Seguro Deportivo Obligatorio">
    <label>Monto $</label><input name="monto" type="number" required placeholder="Ej: 30000">
    <button style="background:#ff7f00">+ AGREGAR / ACTUALIZAR CARGO</button></form>
    <h3>Cargos (""" + str(len(cargos)) + """)</h3>
    <table><tr><th>Cargo</th><th>Monto</th><th>Acciones</th></tr>""" + filas + """</table></div>
    """)

@app.route("/admin/cargos/editar/<path:nombre>", methods=["GET","POST"])
def admin_editar_cargo(nombre):
    if not session.get("admin"): return redirect("/admin")
    db = get_db()
    row = db.execute("SELECT monto FROM cargos WHERE nombre=?", (nombre,)).fetchone()
    monto_actual = row[0] if row else 0
    if request.method == "POST":
        nuevo = request.form["nombre"].strip().upper()
        try: monto = int(request.form["monto"])
        except: monto = monto_actual
        if nuevo:
            db.execute("DELETE FROM cargos WHERE nombre=?", (nombre,))
            db.execute("INSERT OR REPLACE INTO cargos VALUES (?, ?)", (nuevo, monto))
            db.commit()
        return redirect("/admin/cargos")
    return render_template_string(ESTILO + """
    <div class="card"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>Editar Cargo</h2><a href="/admin/cargos">Volver</a></div>
    <form method="post"><label>Nombre</label><input name="nombre" required value='""" + nombre + """'>
    <label>Monto $</label><input name="monto" type="number" required value='""" + str(monto_actual) + """'>
    <button style="background:#ff7f00">GUARDAR</button></form></div>
    """)

@app.route("/admin/cargos/borrar/<nombre>", methods=["GET","POST"])
def admin_borrar_cargo(nombre):
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        borrar_cargo_db(nombre)
        return redirect("/admin/cargos")
    return render_template_string(ESTILO + """
    <div class="card center"><img src="/logo.png" class="logo-mini">
    <div style="font-size:48px">ADVERTENCIA</div><h2>Dar de Baja cargo?</h2><p><b>""" + nombre + """</b></p>
    <form method="post"><button class="btn-danger">SI, DAR DE BAJA</button></form>
    <a href="/admin/cargos"><button class="btn-sec">Cancelar</button></a></div>
    """)

@app.route("/admin/salir")




def admin_salir():
    session.pop("admin", None); return redirect("/")


# --- PARCHE CAT: RUTAS CLUB ---
@app.route("/admin/clubes/usuario/<path:club>", methods=["GET","POST"])
def admin_generar_usuario_club(club):
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        usuario = crear_usuario_club(club)
        return render_template_string(ESTILO + f"<div class='card center'><h2>Usuario creado</h2><p>Club: <b>{club}</b><br>Usuario: <b>{usuario}</b><br>Clave: <b>1234</b></p><p class='small'>Pasale esto al club. Entra en /login con ese usuario.</p><a href='/admin/clubes'><button>Volver a Clubes</button></a></div>")
    uc = get_usuario_club(club)
    if uc:
        return render_template_string(ESTILO + f"<div class='card center'><h2>Usuario del Club</h2><p>Club: <b>{club}</b><br>Usuario: <b>{uc['usuario']}</b><br>Clave actual: <b>{uc['password']}</b></p><form method='post'><button>RESETEAR A 1234</button></form><br><a href='/admin/clubes'>Volver</a></div>")
    return render_template_string(ESTILO + f"<div class='card center'><h2>Generar Usuario</h2><p>Club: <b>{club}</b><br>Usuario sugerido: <b>{slug_club(club)}</b><br>Clave: <b>1234</b></p><form method='post'><button>GENERAR USUARIO</button></form><br><a href='/admin/clubes'>Cancelar</a></div>")

@app.route("/club/panel")
def club_panel():
    usuario = session.get("club_usuario")
    club = session.get("club")
    if not usuario: return redirect("/login")
    hab_map = _map_dni_a_habilitado()
    filas = ""
    total_hab = 0
    for dni, jug in JUGADORES.items():
        if jug.get("club") != club: continue
        hab, pagado, total = hab_map.get(str(dni), (False, 0, 0))
        if not hab: continue
        total_hab += 1
        filas += f"<tr><td><b>{jug.get('nombre','')}</b></td><td>{dni}</td><td>{jug.get('categoria','')}</td><td>HABILITADO</td></tr>"
    return render_template_string(ESTILO + f"<div class='card' style='max-width:900px'><h2>{club}</h2><p>Usuario: <b>{usuario}</b> | Jugadores HABILITADOS: <b>{total_hab}</b></p><div style='text-align:center;margin:15px 0'><a href='/club/exportar' style='display:block;background:#28a745;color:white;padding:18px;border-radius:12px;font-weight:900;text-decoration:none;font-size:18px'>DESCARGAR PLANILLA HABILITADOS</a></div><table><tr><th>Jugador</th><th>DNI</th><th>Categoria</th><th>Estado</th></tr>{filas}</table><br><a href='/logout'><button class='btn-sec'>Salir</button></a></div>")

@app.route("/club/exportar")
def club_exportar():
    usuario = session.get("club_usuario")
    club = session.get("club")
    if not usuario: return redirect("/login")
    hab_map = _map_dni_a_habilitado()
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Jugador","DNI","Club","Categoria","Estado"])
    for dni, jug in JUGADORES.items():
        if jug.get("club") != club: continue
        hab, pagado, total = hab_map.get(str(dni), (False, 0, 0))
        if not hab: continue
        writer.writerow([jug.get("nombre",""), dni, club, jug.get("categoria",""), "HABILITADO"])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": f"attachment; filename=habilitados_{slug_club(club)}.csv"})

@app.route("/logout")
def logout():
    session.clear(); return redirect("/")



# --- V2.9 - ABM Torneos, Temporadas, Tipos, Niveles + Importar Excel ---
def _lista(tabla):
    db = get_db()
    return [r[0] for r in db.execute(f"SELECT nombre FROM {tabla} ORDER BY nombre")]

@app.route("/admin/torneos", methods=["GET","POST"])
def admin_torneos():
    if not session.get("admin"): return redirect("/admin")
    db = get_db()
    if request.method == "POST":
        nombre = request.form["nombre"].strip().upper()
        temporada = request.form.get("temporada","2026/2027")
        tipo = request.form.get("tipo","")
        if nombre: db.execute("INSERT OR IGNORE INTO torneos VALUES (?,?,?)", (nombre, temporada, tipo)); db.commit()
        return redirect("/admin/torneos")
    filas = ""
    for r in db.execute("SELECT * FROM torneos ORDER BY nombre"):
        filas += f"<tr><td><b>{r['nombre']}</b></td><td>{r['temporada']}</td><td>{r['tipo']}</td><td><a href='/admin/torneos/borrar/{r['nombre']}'><button class='btn-danger'>Dar de Baja</button></a></td></tr>"
    temps = "".join([f"<option>{t}</option>" for t in _lista("temporadas")])
    tipos = "".join([f"<option>{t}</option>" for t in _lista("tipos_competicion")])
    return render_template_string(ESTILO + f"""
    <div class="card"><img src="/logo.png" class="logo-mini">
    <div class="topbar"><h2>🏆 Torneos</h2><a href="/admin/panel">Volver</a></div>
    <form method="post"><label>Nombre del torneo</label><input name="nombre" required placeholder="TORNEO CLAUSURA...">
    <label>Temporada</label><select name="temporada" style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px">{temps}</select>
    <label>Tipo de competición</label><select name="tipo" style="width:100%;padding:14px;margin:8px 0 14px;border:1.5px solid #ddd;border-radius:10px">{tipos}</select>
    <button>+ CREAR TORNEO</button></form>
    <table><tr><th>Torneo</th><th>Temporada</th><th>Tipo</th><th></th></tr>{filas}</table></div>""")

@app.route("/admin/torneos/borrar/<path:nombre>", methods=["GET","POST"])
def admin_borrar_torneo(nombre):
    if not session.get("admin"): return redirect("/admin")
    db = get_db()
    if request.method == "POST":
        # No deja borrar si tiene jugadores
        cnt = db.execute("SELECT COUNT(*) FROM jugadores WHERE torneo=?", (nombre,)).fetchone()[0]
        if cnt > 0:
            return f"<p>⛔ No podés borrar {nombre}, tiene {cnt} jugadores. <a href='/admin/torneos'>Volver</a></p>"
        db.execute("DELETE FROM torneos WHERE nombre=?", (nombre,)); db.commit()
        return redirect("/admin/torneos")
    return render_template_string(ESTILO + f"""<div class="card center"><h2>¿Dar de Baja torneo?</h2><p><b>{nombre}</b></p><form method="post"><button class="btn-danger">SI, DAR DE BAJA</button></form><a href="/admin/torneos"><button class="btn-sec">Cancelar</button></a></div>""")

@app.route("/admin/temporadas", methods=["GET","POST"])
def admin_temporadas():
    if not session.get("admin"): return redirect("/admin")
    db = get_db()
    if request.method == "POST":
        n = request.form["nombre"].strip()
        if n: db.execute("INSERT OR IGNORE INTO temporadas VALUES (?,1)", (n,)); db.commit()
        return redirect("/admin/temporadas")
    filas = "".join([f"<tr><td><b>{r[0]}</b></td><td>{'✅ Activa' if r[1] else '⏸️ Cerrada'}</td></tr>" for r in db.execute("SELECT * FROM temporadas")])
    return render_template_string(ESTILO + f"""<div class="card"><div class="topbar"><h2>📅 Temporadas</h2><a href="/admin/panel">Volver</a></div><form method="post"><input name="nombre" required placeholder="2026/2027"><button>+ CREAR</button></form><table><tr><th>Temporada</th><th>Estado</th></tr>{filas}</table></div>""")

@app.route("/admin/tipos", methods=["GET","POST"])
def admin_tipos():
    if not session.get("admin"): return redirect("/admin")
    db = get_db()
    if request.method == "POST":
        n = request.form["nombre"].strip().upper()
        if n: db.execute("INSERT OR IGNORE INTO tipos_competicion VALUES (?)", (n,)); db.commit()
        return redirect("/admin/tipos")
    filas = "".join([f"<tr><td><b>{r[0]}</b></td></tr>" for r in db.execute("SELECT * FROM tipos_competicion")])
    return render_template_string(ESTILO + f"""<div class="card"><div class="topbar"><h2>⚽ Tipo de Competición</h2><a href="/admin/panel">Volver</a></div><form method="post"><input name="nombre" required placeholder="TORNEO DIVISIONES..."><button>+ CREAR</button></form><table><tr><th>Tipo</th></tr>{filas}</table></div>""")

@app.route("/admin/niveles", methods=["GET","POST"])
def admin_niveles():
    if not session.get("admin"): return redirect("/admin")
    db = get_db()
    if request.method == "POST":
        n = request.form["nombre"].strip()
        if n: db.execute("INSERT OR IGNORE INTO niveles VALUES (?)", (n,)); db.commit()
        return redirect("/admin/niveles")
    filas = "".join([f"<tr><td><b>{r[0]}</b></td></tr>" for r in db.execute("SELECT * FROM niveles")])
    return render_template_string(ESTILO + f"""<div class="card"><div class="topbar"><h2>📊 Nivel</h2><a href="/admin/panel">Volver</a></div><form method="post"><input name="nombre" required placeholder="Aficionado"><button>+ CREAR</button></form><table><tr><th>Nivel</th></tr>{filas}</table></div>""")

@app.route("/admin/importar", methods=["GET","POST"])
def admin_importar():
    if not session.get("admin"): return redirect("/admin")
    if request.method == "POST":
        import pandas as pd
        f = request.files.get("file")
        if not f: return "Subí el Excel"
        f.save("/tmp/import.xlsx")
        df = pd.read_excel("/tmp/import.xlsx", sheet_name="Data", header=0)
        db = get_db(); nuevos=0; dup=0
        for _, row in df.iterrows():
            try:
                dni = str(int(float(row.get("DNI Nº")))) if pd.notna(row.get("DNI Nº")) else ""
                if not dni: continue
                if db.execute("SELECT 1 FROM jugadores WHERE dni=?", (dni,)).fetchone(): dup+=1; continue
                nombre = str(row.get("Nombre de persona")) if pd.notna(row.get("Nombre de persona")) else ""
                club = str(row.get("Club")) if pd.notna(row.get("Club")) else ""
                torneo = str(row.get("Nombre")) if pd.notna(row.get("Nombre")) else ""
                temporada = str(row.get("Temporada")) if pd.notna(row.get("Temporada")) else "2026/2027"
                tipo = str(row.get("Tipo de competición")) if pd.notna(row.get("Tipo de competición")) else ""
                cat = str(row.get("Categoría de edad")) if pd.notna(row.get("Categoría de edad")) else ""
                nivel = str(row.get("Nivel")) if pd.notna(row.get("Nivel")) else "Aficionado"
                fnac = str(row.get("Fecha de nacimiento"))[:10] if pd.notna(row.get("Fecha de nacimiento")) else ""
                db.execute("INSERT OR IGNORE INTO clubes VALUES (?)", (club,))
                db.execute("INSERT OR IGNORE INTO categorias VALUES (?)", (cat,))
                db.execute("INSERT OR IGNORE INTO temporadas VALUES (?,1)", (temporada,))
                db.execute("INSERT OR IGNORE INTO tipos_competicion VALUES (?)", (tipo,))
                db.execute("INSERT OR IGNORE INTO niveles VALUES (?)", (nivel,))
                db.execute("INSERT OR IGNORE INTO torneos VALUES (?,?,?)", (torneo, temporada, tipo))
                db.execute("INSERT INTO jugadores(dni,nombre,club,categoria,torneo,temporada,tipo,nivel,fecha_nac) VALUES (?,?,?,?,?,?,?,?,?)", (dni, nombre, club, cat, torneo, temporada, tipo, nivel, fnac))
                nuevos+=1
            except Exception as e: print("Error fila", e); continue
        db.commit()
        return f"<div class='card'><h2>✅ Importación lista</h2><p>Nuevos: <b>{nuevos}</b><br>Duplicados (ya existían): <b>{dup}</b></p><a href='/admin/panel'><button>Volver al panel</button></a></div>"
    return render_template_string(ESTILO + """<div class="card"><div class="topbar"><h2>📥 Importar Excel de 9000</h2><a href="/admin/panel">Volver</a></div><p>Subí el Excel que te dio la Liga (hoja Data). Usa el DNI para no duplicar.</p><form method="post" enctype="multipart/form-data"><input type="file" name="file" accept=".xlsx" required><button>IMPORTAR</button></form></div>""")



# --- ADMIN TUTORES ---
@app.route("/admin/tutores")

def admin_tutores():
    if not session.get("admin"): return redirect("/admin")
    q = request.args.get("q", "").strip().upper()
    filtro = request.args.get("filtro", "todos")

    # Tutores = usuarios (email + dni_jugador) + datos del jugador en JUGADORES
    db = get_db()
    try:
        usuarios = db.execute("SELECT email, dni_jugador FROM usuarios ORDER BY email").fetchall()
    except Exception:
        usuarios = []
    try:
        db.close()
    except: pass

    tutores = []
    for u in usuarios:
        email = u[0] if isinstance(u, tuple) else u["email"]
        dni_j = u[1] if isinstance(u, tuple) else u["dni_jugador"]
        jug = JUGADORES.get(str(dni_j), {})
        nombre_jug = jug.get("nombre","") if jug else "(jugador no encontrado)"
        tutores.append({
            "email": email,
            "dni": str(dni_j),
            "jugador": nombre_jug,
            "club": jug.get("club","") if jug else "",
            "categoria": jug.get("categoria","") if jug else "",
            "estado": jug.get("estado","ACTIVO") if jug else "SIN DATOS",
        })

    if q:
        tutores = [x for x in tutores if q in x["email"].upper() or q in x["jugador"].upper() or q in x["dni"]]

    tutores.sort(key=lambda x: x["email"].lower())

    filas = ""
    for tu in tutores:
        color = "#16a34a" if tu["estado"] == "ACTIVO" else "#dc2626"
        filas += f"<tr><td style='word-break:break-all'>{tu['email']}</td><td><b>{tu['jugador']}</b><br><span style='font-size:12px;color:#666'>DNI {tu['dni']} - {tu['club']} - {tu['categoria']}</span></td><td><span style='background:#dcfce7;color:#166534;padding:4px 10px;border-radius:20px;font-size:12px;font-weight:700'>{tu['estado']}</span></td><td><div class='acciones'><a href='/admin/jugadores?q={tu['dni']}'><button style='background:#2563eb;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer;margin:2px'>Ver jugador</button></a> <a href='/admin/usuarios/cambiar/{tu['email']}'><button style='background:#f59e0b;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer;margin:2px'>Cambiar Jugador</button></a> <a href='/admin/usuarios/reset/{tu['email']}'><button style='background:#6b7280;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer;margin:2px'>Reset 1234</button></a> <a href='/admin/usuarios/borrar/{tu['email']}' onclick=\"return confirm('Dar de Baja {tu['email']}? Se borra usuario y sus pagos.')\"><button style='background:#dc3545;color:white;border:none;padding:6px 10px;border-radius:6px;cursor:pointer;margin:2px'>DAR DE BAJA</button></a></div></td></tr>"
    if not filas:
        filas = "<tr><td colspan='4' style='text-align:center;padding:20px'>No hay tutores para mostrar.</td></tr>"

    _style = """
    <style>
      .card input[name="q"]{flex:1 !important;width:auto !important;min-width:0 !important}
      .card form{flex-wrap:wrap !important}
      .card form button{width:auto !important;flex:none !important;white-space:nowrap}
      .card table button{width:auto !important;display:inline-block !important;padding:6px 10px !important;font-size:12px !important;white-space:nowrap}
      .card table td:last-child{white-space:normal !important;min-width:280px}
      .acciones{display:flex;flex-wrap:wrap;gap:4px}
      .acciones a{flex:1 1 110px}
      .acciones button{width:100% !important}
    </style>
    """
    return ESTILO + _style + f"""
    <div class='card' style='max-width:1100px;margin:20px auto'>
      <a href="/admin">← Volver al Panel</a>
      <h2>👨‍👩‍👧 Tutores</h2>
      <p style="color:#555">Listado de tutores y jugador a su cargo. El tutor restablece su clave solo desde "Olvidé mi clave" en el ingreso.</p>
      <div style="display:flex;gap:8px;flex-wrap:wrap;margin:10px 0;"><a href="/admin/jugadores?filtro=todos"><button style="width:auto;padding:10px 14px;background:#0a4a2b">Todos</button></a><a href="/admin/jugadores?filtro=habilitados"><button style="width:auto;padding:10px 14px;background:#16a34a">✅ Solo Habilitados</button></a><a href="/admin/jugadores?filtro=pendientes"><button style="width:auto;padding:10px 14px;background:#dc2626">❌ Con Deuda</button></a></div><form method="get" style="margin:12px 0;display:flex;gap:8px">
        <input name="q" value="{q}" placeholder="Buscar por email, jugador o DNI..." style="flex:1;padding:10px;border:1px solid #ccc;border-radius:8px">
        <button style="padding:10px 18px;background:#111827;color:#fff;border:0;border-radius:8px;cursor:pointer">Buscar</button>
      </form>
      <div style="overflow-x:auto">
      <table style="width:100%;border-collapse:collapse">
        <tr style="background:#f3f4f6"><th style="text-align:left;padding:12px">Tutor (email)</th><th style="text-align:left;padding:12px">Jugador asignado</th><th style="text-align:left;padding:12px">Estado</th><th style="text-align:left;padding:12px">Acciones</th></tr>
        {filas}
      </table>
      </div>
      <p style="font-size:12px;color:#888;margin-top:10px">Total tutores: {len(tutores)}</p>
    </div>
    """

if __name__ == "__main__":
    import os
    port = int(os.environ.get("PORT", 5000))
    app.run(host="0.0.0.0", port=port, debug=False)
