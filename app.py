from flask import Flask, render_template, request, send_file
from fpdf import FPDF
import os
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

app = Flask(__name__)

# Clase personalizada para generar el PDF
class PDF(FPDF):
    def header(self):
        self.set_font("Arial", "B", 12)
        self.cell(0, 10, "Formulario de Visita al Cliente", align="L")
        logo_path = "static/images/logoCTR.png"
        if os.path.exists(logo_path):
            self.image(logo_path, 170, 8, 30)
        self.ln(20)

    def footer(self):
        self.set_y(-15)
        self.set_font("Arial", "I", 8)
        self.cell(0, 10, f"Página {self.page_no()}", 0, 0, "C")

# Función para generar el PDF
def generar_pdf(data):
    pdf = PDF()
    pdf.add_page()

    pdf.set_font("Arial", "B", 12)
    for key, value in data.items():
        pdf.cell(80, 10, f"{key}:", 0, 0, "L")
        pdf.set_font("Arial", "", 12)
        pdf.cell(0, 10, f"{value if value else 'N/A'}", 0, 1, "L")
        pdf.set_font("Arial", "B", 12)

    directory = "pdfs"
    if not os.path.exists(directory):
        os.makedirs(directory)

    nombre_cliente = data.get("Nombre del Cliente", "Cliente")
    nombre_cliente = nombre_cliente.replace(" ", "_")
    filename = os.path.join(directory, f"ReporteVisitaCliente_{nombre_cliente}.pdf")
    pdf.output(filename)
    return filename

# Función para enviar el PDF por correo
def enviar_email(asunto, mensaje, archivo_pdf):
    smtp_server = "smtp.office365.com"
    smtp_port = 587
    email_remitente = "francisco.galvez@ctr.com.mx"
    email_password = "65228Esg"
    email_destinatario = "francisco.galvez@ctr.com.mx"

    try:
        msg = MIMEMultipart()
        msg["From"] = email_remitente
        msg["To"] = email_destinatario
        msg["Subject"] = asunto
        msg.attach(MIMEText(mensaje, "plain"))

        with open(archivo_pdf, "rb") as attachment:
            part = MIMEBase("application", "octet-stream")
            part.set_payload(attachment.read())
        encoders.encode_base64(part)
        part.add_header("Content-Disposition", f"attachment; filename={os.path.basename(archivo_pdf)}")
        msg.attach(part)

        server = smtplib.SMTP(smtp_server, smtp_port)
        server.starttls()
        server.login(email_remitente, email_password)
        server.send_message(msg)
        server.quit()

        return "Correo enviado exitosamente"
    except Exception as e:
        return f"Error al enviar el correo: {str(e)}"

# Ruta para mostrar el formulario HTML
@app.route('/')
def index():
    return render_template('index.html')

# Ruta para recibir datos, generar y enviar el PDF, y descargarlo
@app.route('/generar_pdf', methods=['POST'])
def generar_pdf_endpoint():
    data = {
        "Nombre del Cliente": request.form.get("nombre"),
        "Fecha de la Visita": request.form.get("fecha"),
        "Departamento": request.form.get("departamento"),
        "Usuario": request.form.get("usuario"),
        "Sector": request.form.get("sector"),
        "Vendedor": request.form.get("vendedor"),
        "Especialista de Producto": request.form.get("especialista"),
        "Equipos o Productos Presentados": request.form.get("productos"),
        "Nivel de Interés": request.form.get("interes"),
        "Comentarios/Preguntas": request.form.get("comentarios"),
        "Equipos de la Competencia": request.form.get("competencia"),
        "Monto Estimado": request.form.get("monto"),
        "Oportunidades para Otros Productos": request.form.get("oportunidades"),
        "Requiere Seguimiento": "Sí" if request.form.get("seguimiento") else "No",
    }

    try:
        # Generar el PDF
        filename = generar_pdf(data)

        # Enviar el PDF por correo
        asunto = "Reporte de Visita al Cliente"
        mensaje = "Adjunto encontrarás el reporte de visita al cliente generado desde el formulario."
        enviar_email(asunto, mensaje, filename)

        # Descargar el PDF automáticamente
        return send_file(filename, as_attachment=True, download_name=os.path.basename(filename), mimetype='application/pdf')
    except Exception as e:
        return f"Error: {str(e)}", 500

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    app.run(host='0.0.0.0', port=port, debug=True)
