from flask import Flask, request, redirect, url_for, session, render_template_string, Response
import datetime, csv, io
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from fpdf import FPDF

app = Flask(__name__)
app.secret_key = "liga-mendocina-demo-123"

# ============================================================
# 👉👉 ACÁ SOLO CAMBIÁS ESTAS 2 LÍNEAS 👈👈
# ============================================================
EMAIL_REMITENTE = "fabiand@didiosistemas.com.ar"
EMAIL_CLAVE_APP = "iutu skcd dmut zhij"
# ============================================================

# --- DATOS ADMIN PARA LA REUNIÓN ---
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
    {"id": 1, "nombre": "Inscripción Anual Liga 2026/2027", "monto": 40000},
    {"id": 2, "nombre": "Seguro Deportivo Obligatorio", "monto": 8000},
    {"id": 3, "nombre": "Carnet / Ficha del Jugador", "monto": 5000},
]
USUARIOS = {
    "demo@gmail.com": {"dni_jugador": "50035043", "password": "50035043", "pagos": []},
}

def enviar_email(destinatario, asunto, texto, pdf_bytes=None, pdf_nombre="recibo.pdf"):
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
        server = smtplib.SMTP("smtp.gmail.com", 587)
        server.starttls()
        server.login(EMAIL_REMITENTE, EMAIL_CLAVE_APP)
        server.send_message(msg)
        server.quit()
        print(f"✅ Email enviado a {destinatario}")
        return True
    except Exception as e:
        print(f"❌ Error enviando email: {e}")
        return False

def generar_pdf_recibo(nombre_jugador, dni, club, total):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", "B", 18)
    pdf.cell(0, 15, "LIGA MENDOCINA DE FUTBOL", ln=True, align="C")
    pdf.set_font("Arial", "", 12)
    pdf.cell(0, 10, "Recibo de Pago - Autogestion", ln=True, align="C")
    pdf.ln(10)
    pdf.cell(0, 10, f"Jugador: {nombre_jugador}", ln=True)
    pdf.cell(0, 10, f"DNI: {dni} - Club: {club}", ln=True)
    pdf.cell(0, 10, "Concepto: Inscripcion + Seguro + Carnet", ln=True)
    pdf.cell(0, 10, f"Total abonado: ${total:,}", ln=True)
    pdf.cell(0, 10, "Medio: Mercado Pago (DEMO)", ln=True)
    pdf.cell(0, 10, "Estado: HABILITADO", ln=True)
    pdf.cell(0, 10, f"Fecha: {datetime.datetime.now().strftime('%d/%m/%Y %H:%M')}", ln=True)
    return bytes(pdf.output())

ESTILO = """
<style>
body{font-family:Arial,sans-serif;background:#f2f2f2;padding:20px}
.card{background:white;max-width:750px;margin:auto;padding:25px;border-radius:12px;box-shadow:0 2px 10px #ccc}
h1{color:#0a4a2b;text-align:center} h2{color:#0a4a2b}
input{width:100%;padding:12px;margin:8px 0;border:1px solid #ccc;border-radius:8px;box-sizing:border-box}
button{background:#00a650;color:white;border:none;padding:14px;width:100%;border-radius:8px;font-size:16px;cursor:pointer;font-weight:bold}
button:hover{background:#008a43}
.btn-mp{background:#009ee3} .btn-admin{background:#333} .btn-excel{background:#1d6f42}
.cargo{display:flex;justify-content:space-between;padding:12px;border-bottom:1px solid #eee}
.pagado{color:green;font-weight:bold}
.badge{padding:8px 14px;border-radius:20px;font-weight:bold;display:inline-block}
.habilitado{background:#d4edda;color:#155724} .pendiente-bg{background:#f8d7da;color:#721c24}
a{color:#009ee3} .center{text-align:center} .small{font-size:13px;color:#666}
table{width:100%;border-collapse:collapse;margin-top:15px}
th,td{padding:10px;border-bottom:1px solid #eee;text-align:left;font-size:14px}
th{background:#0a4a2b;color:white}
.kpi{display:flex;gap:10px;margin:15px 0}
.kpi div{flex:1;background:#f8f9fa;padding:15px;border-radius:10px;text-align:center}
.kpi b{font-size:22px;display:block}
</style>
"""

@app.route("/")
def index():
    return render_template_string(ESTILO + """
    <div class="card center">
        <h1>⚽ LIGA MENDOCINA DE FÚTBOL</h1>
        <p><b>Autogestión de Jugadores</b><br>Pagá la inscripción sin hacer cola</p>
        <a href="/registro"><button>SOY PADRE / MADRE - CREAR MI CUENTA</button></a><br><br>
        <a href="/login"><button style="background:#333">YA TENGO CUENTA - INGRESAR</button></a><br><br>
        <a href="/admin">Acceso Liga / Admin →</a>
        <p class="small">Demo - Para probar usá DNI: 50035043 (CHAVEZ)</p>
    </div>
    """)

@app.route("/registro", methods=["GET","POST"])
def registro():
    if request.method == "POST":
        email = request.form["email"].lower().strip()
        dni = request.form["dni"].strip()
        if dni not in JUGADORES:
            return render_template_string(ESTILO + f"""<div class="card center"><h2>❌ DNI no encontrado</h2><p>El DNI <b>{dni}</b> no está en la base.</p><a href="/registro">Volver</a></div>""")
        if email in USUARIOS:
            return redirect("/login")
        USUARIOS[email] = {"dni_jugador": dni, "password": dni, "pagos": []}
        jug = JUGADORES[dni]
        enviar_email(email, "Tu cuenta Liga Mendocina - Bienvenido", f"Hola!\n\nTu cuenta fue creada con exito.\n\nUsuario: {email}\nClave inicial: {dni} (DNI del jugador)\nJugador vinculado: {jug['nombre']} - {jug['club']} {jug['categoria']}\n\nLiga Mendocina de Futbol")
        return render_template_string(ESTILO + f"""<div class="card center"><h2>✅ ¡Cuenta creada!</h2><p>Te enviamos un correo a <b>{email}</b></p><p><b>Usuario:</b> {email}<br><b>Clave:</b> {dni}<br><b>Jugador:</b> {jug['nombre']}</p><a href="/login"><button>INGRESAR AHORA</button></a></div>""")
    return render_template_string(ESTILO + """<div class="card"><h2>Crear cuenta</h2><form method="post"><label>Tu Email</label><input name="email" type="email" required placeholder="papa@gmail.com"><label>DNI del Jugador</label><input name="dni" required placeholder="Ej: 50035043"><button>VALIDAR Y CREAR CUENTA</button></form><a href="/">Volver</a></div>""")

@app.route("/login", methods=["GET","POST"])
def login():
    if request.method == "POST":
        email = request.form["email"].lower().strip()
        clave = request.form["password"].strip()
        u = USUARIOS.get(email)
        if u and u["password"] == clave:
            session["email"] = email
            return redirect("/panel")
        return render_template_string(ESTILO + '<div class="card center"><h2>❌ Datos incorrectos</h2><a href="/login">Reintentar</a></div>')
    return render_template_string(ESTILO + """<div class="card"><h2>Ingresar</h2><form method="post"><label>Usuario (tu email)</label><input name="email" type="email" required><label>Clave</label><input name="password" type="password" required><button>INGRESAR</button></form><a href="/">Volver</a></div>""")

@app.route("/panel")
def panel():
    email = session.get("email")
    if not email: return redirect("/login")
    u = USUARIOS[email]
    jug = JUGADORES[u["dni_jugador"]]
    total = sum(c["monto"] for c in CARGOS)
    pagado_total = sum(c["monto"] for c in CARGOS if c["id"] in u["pagos"])
    pendiente = total - pagado_total
    habilitado = pendiente == 0
    cargos_html = ""
    for c in CARGOS:
        estado = '<span class="pagado">✅ PAGADO</span>' if c["id"] in u["pagos"] else f'<a href="/pagar/{c["id"]}"><button style="width:auto;padding:8px 16px">PAGAR</button></a>'
        cargos_html += f'<div class="cargo"><span>{c["nombre"]}</span><span><b>${c["monto"]:,}</b> {estado}</span></div>'
    badge = '<span class="badge habilitado">🟢 HABILITADO PARA JUGAR</span>' if habilitado else '<span class="badge pendiente-bg">🔴 PENDIENTE DE PAGO</span>'
    btn_todo = "" if habilitado else f'<a href="/pagar_todo"><button class="btn-mp">PAGAR TODO ${total:,} CON MERCADO PAGO - DEMO</button></a>'
    return render_template_string(ESTILO + f"""<div class="card"><div style="display:flex;justify-content:space-between"><b>Hola, {email}</b> <a href="/logout">Salir</a></div><h2>{jug['nombre']}</h2><p>Club: <b>{jug['club']}</b> | Categoría: <b>{jug['categoria']}</b><br>Torneo: {jug['torneo']} | DNI: {u['dni_jugador']}</p><p>Estado: {badge}</p><h3>Cargos a pagar</h3>{cargos_html}<p style="text-align:right"><b>Total pendiente: ${pendiente:,}</b></p>{btn_todo}</div>""")

@app.route("/pagar/<int:cid>")
def pagar(cid):
    email = session.get("email")
    if not email: return redirect("/login")
    if cid not in USUARIOS[email]["pagos"]:
        USUARIOS[email]["pagos"].append(cid)
    if len(USUARIOS[email]["pagos"]) == len(CARGOS):
        jug = JUGADORES[USUARIOS[email]["dni_jugador"]]
        pdf_bytes = generar_pdf_recibo(jug["nombre"], USUARIOS[email]["dni_jugador"], jug["club"], 53000)
        enviar_email(email, "Recibo de Pago - Liga Mendocina", f"Gracias! Pagaste $53.000 por {jug['nombre']}.\nAdjunto tu recibo en PDF. Ya estas HABILITADO.\n\nLiga Mendocina", pdf_bytes=pdf_bytes)
    return redirect("/panel")

@app.route("/pagar_todo")
def pagar_todo():
    email = session.get("email")
    if not email: return redirect("/login")
    USUARIOS[email]["pagos"] = [c["id"] for c in CARGOS]
    jug = JUGADORES[USUARIOS[email]["dni_jugador"]]
    pdf_bytes = generar_pdf_recibo(jug["nombre"], USUARIOS[email]["dni_jugador"], jug["club"], 53000)
    enviar_email(email, "Recibo de Pago - Liga Mendocina", f"Gracias! Pagaste $53.000 por {jug['nombre']}.\nAdjunto tu recibo en PDF. Ya estas HABILITADO.\n\nLiga Mendocina", pdf_bytes=pdf_bytes)
    return redirect("/panel")

# ====== PANEL ADMIN ======
@app.route("/admin", methods=["GET","POST"])
def admin_login():
    if request.method == "POST":
        if request.form["email"] == ADMIN_USER and request.form["password"] == ADMIN_PASS:
            session["admin"] = True
            return redirect("/admin/panel")
        return render_template_string(ESTILO + '<div class="card center"><h2>❌ Datos incorrectos</h2><a href="/admin">Reintentar</a></div>')
    return render_template_string(ESTILO + """<div class="card"><h2>🔐 Panel Liga - Admin</h2><p class="small">Acceso solo para personal de la Liga</p><form method="post"><label>Usuario Admin</label><input name="email" required placeholder="admin@liga.com"><label>Clave</label><input name="password" type="password" required><button class="btn-admin">ENTRAR AL PANEL</button></form><p class="small">Demo: admin@liga.com / admin123</p><a href="/">Volver</a></div>""")

@app.route("/admin/panel")
def admin_panel():
    if not session.get("admin"): return redirect("/admin")
    total_recaudado = 0
    cant_pagados = 0
    cant_pendientes = 0
    filas = ""
    for email, u in USUARIOS.items():
        jug = JUGADORES.get(u["dni_jugador"], {"nombre":"?","club":"?","categoria":"?"})
        pagado = sum(c["monto"] for c in CARGOS if c["id"] in u["pagos"])
        total_recaudado += pagado
        habilitado = pagado == 53000
        if habilitado: cant_pagados += 1
        else: cant_pendientes += 1
        badge = '<span class="badge habilitado">HABILITADO</span>' if habilitado else '<span class="badge pendiente-bg">PENDIENTE</span>'
        filas += f"<tr><td><b>{jug['nombre']}</b><br><span class='small'>{email}</span></td><td>{jug['club']}<br>{jug['categoria']}</td><td>${pagado:,}</td><td>{badge}</td></tr>"
    return render_template_string(ESTILO + f"""
    <div class="card" style="max-width:900px">
        <div style="display:flex;justify-content:space-between"><h2>📊 Panel Liga Mendocina</h2><a href="/admin/salir">Salir</a></div>
        <div class="kpi">
            <div><span>💰 RECAUDADO HOY</span><b>${total_recaudado:,}</b></div>
            <div><span>✅ HABILITADOS</span><b>{cant_pagados}</b></div>
            <div><span>🔴 PENDIENTES</span><b>{cant_pendientes}</b></div>
        </div>
        <a href="/admin/exportar"><button class="btn-excel">⬇ EXPORTAR A EXCEL / CSV</button></a>
        <table><tr><th>Jugador / Padre</th><th>Club</th><th>Pagado</th><th>Estado</th></tr>{filas}</table>
        <p class="small center">En versión final: filtro por club, por categoría, y reporte para ARCA</p>
    </div>""")

@app.route("/admin/exportar")
def admin_exportar():
    if not session.get("admin"): return redirect("/admin")
    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["Jugador","DNI","Club","Categoria","Email Padre","Monto Pagado","Estado"])
    for email, u in USUARIOS.items():
        jug = JUGADORES.get(u["dni_jugador"], {"nombre":"?","club":"?","categoria":"?"})
        pagado = sum(c["monto"] for c in CARGOS if c["id"] in u["pagos"])
        estado = "HABILITADO" if pagado == 53000 else "PENDIENTE"
        writer.writerow([jug["nombre"], u["dni_jugador"], jug["club"], jug["categoria"], email, pagado, estado])
    return Response(output.getvalue(), mimetype="text/csv", headers={"Content-Disposition": "attachment; filename=liga_mendocina.csv"})

@app.route("/admin/salir")
def admin_salir():
    session.pop("admin", None); return redirect("/")

@app.route("/logout")
def logout():
    session.clear(); return redirect("/")

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
