import os
import time
import smtplib
import pandas as pd
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr, make_msgid # Added make_msgid
from openpyxl import Workbook
import logging
from .config import SMTP_SERVER, SMTP_PORT, EMAIL_USER, EMAIL_PASSWORD, load_email_templates, LOTE_SIZE # Import LOTE_SIZE
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger

logger = logging.getLogger(__name__)

class EmailSender:
    def __init__(self):
        self.subject_template, self.body_template = load_email_templates()

    def is_valid_email(self, email):
        import re
        pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        return re.match(pattern, email) is not None

    def is_valid_dni(self, dni):
        return dni.isdigit() and len(dni) == 8

    def send_batch(self, recipients, mes, path_boletas, progress_callback=None):
        # LOTE_SIZE = 30 # Removed local definition, now imported
        successful_sends = [] # List to store details of successful sends
        errores = []
        total_recipients = len(recipients)

        # Siempre recarga la plantilla antes de enviar
        def get_current_templates():
            from .config import load_email_templates
            return load_email_templates()

        def enviar_lote(lote, start_index):
            subject_template, body_template = get_current_templates()  # <-- recarga aquí
            try:
                if progress_callback:
                    progress_callback(f"🔗 Conectando al servidor SMTP...")
                server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
                server.login(EMAIL_USER, EMAIL_PASSWORD)
                logger.info("Conexión SMTP establecida para el lote.")
            except Exception as e:
                logger.error(f"Error con servidor SMTP al conectar lote: {str(e)}")
                for offset, recipient in enumerate(lote):
                    i = start_index + offset
                    nombre = recipient["nombre"]
                    email = str(recipient["email"]).strip()
                    dni = str(recipient["dni"]).strip()
                    errores.append((i, nombre, email, dni, f"Error SMTP: {str(e)} (No se pudo conectar)"))
                return
            for offset, recipient in enumerate(lote):
                i = start_index + offset
                if progress_callback:
                    progress_callback(f"📧 Enviando correo {i-1} de {total_recipients}...")
                nombre = recipient["nombre"]
                email = str(recipient["email"]).strip()
                dni = str(recipient["dni"]).strip()
                if not self.is_valid_dni(dni):
                    errores.append((i, nombre, email, dni, "DNI inválido"))
                    continue
                if not self.is_valid_email(email):
                    errores.append((i, nombre, email, dni, "Email inválido"))
                    continue
                dni_formatted = dni.zfill(8)
                pdf_path = os.path.join(path_boletas, mes, f"{dni_formatted}.pdf")
                if not os.path.exists(pdf_path):
                    errores.append((i, nombre, email, dni, "PDF no encontrado"))
                    continue
                msg = MIMEMultipart("alternative")
                msg["From"] = formataddr(("Clínica Santa Rosa", EMAIL_USER))
                msg["To"] = formataddr((nombre, email))
                msg["Subject"] = self.subject_template.replace("{MES}", mes.capitalize())
                msg.add_header('Disposition-Notification-To', EMAIL_USER)
                msg_id = make_msgid()
                msg['Message-ID'] = msg_id
                html_content = self.body_template.replace("{NOMBRE}", nombre).replace("{MES}", mes.capitalize())
                # Agregar Message-ID al final del cuerpo
                html_content += f"<br><br><small><b>Identificador de envío (Message-ID):</b> {msg_id}</small>"
                msg.attach(MIMEText(html_content, "html"))
                try:
                    with open(pdf_path, "rb") as f:
                        part = MIMEApplication(f.read(), _subtype="pdf")
                        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
                        msg.attach(part)
                    server.send_message(msg)
                    # Collect details of successfully sent email
                    successful_sends.append({
                        'nombre': nombre,
                        'email': email,
                        'dni': dni,
                        'pdf_path': pdf_path,
                        'message_id': msg_id,
                        'asunto': msg["Subject"],
                        'cuerpo_html': html_content
                    })
                    time.sleep(1.5) # Keep the delay to avoid overwhelming the server
                except Exception as e:
                    errores.append((i, nombre, email, dni, f"Error SMTP: {str(e)}"))
            server.quit()
            time.sleep(2)
        for batch_start in range(0, total_recipients, LOTE_SIZE):
            batch = recipients[batch_start:batch_start + LOTE_SIZE]
            enviar_lote(batch, batch_start + 2)
        return successful_sends, errores

    def generar_constancia_envio(self, remitente, destinatario, asunto, cuerpo, fecha_envio, adjunto_path, message_id=None):
        """
        Genera un PDF de constancia de envío exitoso y lo combina con el PDF adjunto enviado.
        :param remitente: (nombre, email)
        :param destinatario: (nombre, email)
        :param asunto: str
        :param cuerpo: str (puede ser HTML)
        :param fecha_envio: str (YYYY-MM-DD HH:MM:SS)
        :param adjunto_path: ruta al PDF adjunto enviado
        :param message_id: Message-ID del correo (opcional)
        :return: ruta al PDF combinado generado
        """
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader
        from PyPDF2 import PdfMerger
        import tempfile
        import shutil
        # Crear PDF temporal de constancia
        constancia_pdf = tempfile.NamedTemporaryFile(delete=False, suffix="_constancia.pdf")
        c = canvas.Canvas(constancia_pdf.name, pagesize=letter)
        width, height = letter
        # Añadir logo
        logo_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'img', 'logocsr.png')
        if os.path.exists(logo_path):
            logo_img = ImageReader(logo_path)
            logo_width = 60
            logo_height = 60
            c.drawImage(logo_img, width/2-logo_width/2, height-80, width=logo_width, height=logo_height, mask='auto')
            y = height - 100
        else:
            y = height - 40
        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width/2, y, "CONSTANCIA DE ENVÍO DE CORREO ELECTRÓNICO")
        y -= 30
        c.setFont("Helvetica", 10)
        c.drawString(40, y, f"Fecha/hora de envío: {fecha_envio}")
        y -= 18
        if message_id:
            c.setFont("Helvetica", 9)
            c.drawString(40, y, f"Message-ID: {message_id}")
            y -= 18
        c.setFont("Helvetica", 10)
        c.line(40, y, width-40, y)
        y -= 25
        # Mostrar correctamente el remitente (nombre y correo)
        c.drawString(40, y, f"De: {remitente[0]} <{remitente[1] if remitente[1] else EMAIL_USER}>")
        y -= 18
        c.drawString(40, y, f"Para: {destinatario[0]} <{destinatario[1]}>")
        y -= 18
        c.drawString(40, y, f"Asunto: {asunto}")
        y -= 18
        c.drawString(40, y, f"Adjunto: {os.path.basename(adjunto_path)}")
        y -= 25
        c.setFont("Helvetica-Bold", 11)
        c.drawString(40, y, "Mensaje:")
        y -= 18
        c.setFont("Helvetica", 10)
        # Imprimir cuerpo del mensaje respetando saltos de párrafo del HTML
        import re
        import textwrap
        # Reemplazar etiquetas <br> y <p> por saltos de línea reales
        cuerpo_html = cuerpo.replace('<br>', '\n').replace('<br/>', '\n').replace('<br />', '\n')
        cuerpo_html = re.sub(r'<p[^>]*>', '\n', cuerpo_html, flags=re.IGNORECASE)
        cuerpo_html = cuerpo_html.replace('</p>', '\n')
        # Eliminar el resto de etiquetas HTML
        cuerpo_texto = re.sub('<[^<]+?>', '', cuerpo_html)
        max_width = 90  # caracteres por línea
        for line in cuerpo_texto.splitlines():
            if not line.strip():
                y -= 10  # Espacio extra para líneas vacías (párrafos)
                continue
            wrapped_lines = textwrap.wrap(line.strip(), width=max_width, break_long_words=False, replace_whitespace=False)
            for wrapped in wrapped_lines:
                if y < 60:
                    c.showPage(); y = height - 40; c.setFont("Helvetica", 10)
                c.drawString(50, y, wrapped)
                y -= 15
            if wrapped_lines:
                y -= 5  # Espacio extra entre párrafos
        y -= 10
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(40, y, "Este documento certifica que el correo fue enviado exitosamente desde la aplicación.")
        c.save()
        # Combinar constancia y PDF adjunto
        output_dir = os.path.join(os.path.dirname(adjunto_path), "constancias")
        os.makedirs(output_dir, exist_ok=True)
        output_pdf_path = os.path.join(output_dir, f"constancia_{os.path.splitext(os.path.basename(adjunto_path))[0]}.pdf")
        merger = PdfMerger()
        merger.append(constancia_pdf.name)
        merger.append(adjunto_path)
        merger.write(output_pdf_path)
        merger.close()
        constancia_pdf.close()
        try:
            os.unlink(constancia_pdf.name)
        except Exception:
            pass
        return output_pdf_path

    def generar_log_errores(self, errores):
        if not errores:
            return None
        wb = Workbook()
        ws = wb.active
        ws.title = "Errores"
        ws.append(["Fila", "Nombre", "Email", "DNI", "Motivo"])
        for fila in errores:
            ws.append(fila)
        base_filename = "logError"
        extension = ".xlsx"
        attempt = 0
        error_file_saved = None
        from datetime import datetime
        while attempt < 5:
            try:
                if attempt == 0:
                    filename = f"{base_filename}{extension}"
                else:
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    filename = f"{base_filename}_{timestamp}{extension}"
                wb.save(filename)
                error_file_saved = filename
                break
            except PermissionError:
                attempt += 1
                continue
            except Exception:
                break
        return error_file_saved
