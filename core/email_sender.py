import os
import time
import smtplib
import pandas as pd
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr, make_msgid
import openpyxl
from openpyxl import Workbook
import logging
# Updated import: LOTE_SIZE is still needed, load_email_templates is used in constructor
from .config import SMTP_SERVER, SMTP_PORT, EMAIL_USER, EMAIL_PASSWORD, load_email_templates, LOTE_SIZE
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger

logger = logging.getLogger(__name__)

class EmailSender:
    def __init__(self, template_name="default"): # Added template_name parameter
        # Load the specified template upon instantiation
        try:
            self.subject_template, self.body_template = load_email_templates(template_name)
            logger.info(f"EmailSender initialized with template: '{template_name}'")
        except Exception as e:
            logger.error(f"Failed to load template '{template_name}' during EmailSender init: {e}. Using fallback defaults.")
            # Fallback to hardcoded defaults if load_email_templates fails critically
            self.subject_template = "Boleta del mes de {MES}"
            self.body_template = ("<html><body><h2>Hola {NOMBRE},</h2>"
                                      "<p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p>"
                                      "<p>Saludos cordiales,<br>Clínica Santa Rosa</p>"
                                       "<p>Favor de confirmar la recepciòn de este correo.</p>"
                                       "</mody></html>")

    def is_valid_email(self, email):
        import re
        # Pattern correctly defined as a raw string for regex
        pattern = r"^[\\w\\.\\-]+@[\\w\\.\\-]+\\.\\w+$"
        return re.match(pattern, email) is not None

    def is_valid_dni(self, dni):
        return dni.isdigit() and len(dni) == 8

    def send_batch(self, recipients, mes, path_boletas, progress_callback=None):
        successful_sends = []
        errores = []
        total_recipients = len(recipients)

        def enviar_lote(lote, start_index):
            try:
                if progress_callback:
                    progress_callback(f"◉ Conectando al servidor SMTP...")
                server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
                server.login(EMAIL_USER, EMAIL_PASSWORD)
                logger.info("Conexión SMTP establecida para el lote.")
            except Exception as e:
                logger.error(f'Error con servidor SMTP al conectar lote: {str(e)}')
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
                    progress_callback(f"↳ Enviando correo {i+1} de {total_recipients}...")

                nombre = recipient["nombre"]
                email = str(recipient["email"]).strip()
                dni = str(recipient["dni"]).strip()

                if not self.is_valid_dni(dni):
                    errores.append((i, nombre, email, dni, "DNI Inválido"))
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
                # Use placeholders directly in replace, no f-string needed here for self.subject_template
                msg["Subject"] = self.subject_template.replace("{MES}", mes.capitalize()).replace("{NOMBRE}", nombre)
                msg.add_header('Disposition-Notification-To', EMAIL_USER)
                msg_id = make_msgid()
                msg['Message-ID'] = msg_id

                # Same for self.body_template
                html_content = self.body_template.replace("{NOMBRE}", nombre).replace("{MES}", mes.capitalize())
                html_content += f"<br><br><small><b>Identificador de envío (Message-ID): {msg_id}</b></small>"
                msg.attach(MIMEText(html_content, "html"))

                try:
                    with open(pdf_path, "rb") as f:
                        part = MIMEApplication(f.read(), _subtype="pdf")
                        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
                        msg.attach(part)

                    server.send_message(msg)
                    successful_sends.append(({
                        'nombre': nombre,
                        'email': email,
                        'dni': dni,
                        'pdf_path': pdf_path,
                        'message_id': msg_id,
                        'asunto': msg["Subject"],
                        'cuerpo_html': html_content
                    }))
                    time.sleep(1.5)
                except Exception as e:
                    errores.append((i, nombre, email, dni, f'Error SMTP al enviar: {str(e)}'))

            try:
                server.quit()
                logger.info("Conexión SMTP cerrada para el lote.")
            except Exception as e:
                logger.warning(f"Error al cerrar conexión SMTP: {e}")

            time.sleep(2)

        for batch_start in range(0, total_recipients, LOTE_SIZE):
            batch = recipients[batch_start:batch_start + LOTE_SIZE]
            enviar_lote(batch, batch_start + 0) # Fixed: start_index should be batch_start

        return successful_sends, errores

    def generar_constancia_envio(self, remitente, destinatario, asunto, cuerpo, fecha_envio, adjunto_path, message_id=None):
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader # Ensure this is imported
        from PyPDF2 import PdfMerger # Ensure this is imported
        import tempfile
        import shutil # Keep for good measure, though os.unlink is used
        constancia_pdf = tempfile.NamedTemporaryFile(delete=False, suffix="_constancia.pdf")
        c = canvas.Canvas(constancia_pdf.name, pagesize=letter)
        width, height = letter
        # Correct path assuming 'img' is at the root, and this file is in 'core'
        logo_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'img', 'logocsr.png')
        if os.path.exists(logo_path):
            try:
                logo_img = ImageReader(logo_path)
                logo_width = 60
                logo_height = 60
                c.drawImage(logo_img, width/2-logo_width/2, height-80, width=logo_width, height=logo_height, mask='auto')
                y = height - 100
            except Exception as img_e: # Catch error if image is problematic
                logger.error(f'Error loading logo for constancia: {img_e}')
                y = height - 40 # Default position if logo fails
        else:
            logger.warning(f'Logo not found at {logo_path} for constancia.')
            y = height - 40

        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width/2, y, "CONSTANCIA DE ENVIO DE CORREO ELECTRÓNICO")
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
        import re # Ensure re is imported
        import textwrap # Ensure textwrap is imported
        # Use literal \n for splitlines, after replacing HTML breaks
        cuerpo_html_newlines = cuerpo.replace('<br>', '\\n').replace('<br/>', '\\n').replace('<br />', '\\n') # Python literal \\n
        cuerpo_html_newlines = re.sub('<p>', '\\n', cuerpo_html_newlines, flags=re.IGNORECASE)
        cuerpo_html_newlines = cuerpo_html_newlines.replace('</p>', '\\n')
        cuerpo_texto = re.sub('<[^<]+?>', '', cuerpo_html_newlines) # Remove HTML tags

        max_width = 90
        for line in cuerpo_texto.split('\\n'): # Split by literal '\n'
            if not line.strip():
                y -= 10
                continue
            wrapped_lines = textwrap.wrap(line.strip(), width=max_width, break_long_words=False, replace_whitespace=False)
            for wrapped in wrapped_lines:
                if y < 60: # Check if new page is needed
                    c.showPage()
                    y = height - 40 # Reset y position
                    c.setFont("Helvetica", 10) # Reset font
                c.drawString(50, y, wrapped)
                y -= 15
            if wrapped_lines:
                y -= 5
        y -= 10
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(40, y, "Este documento certifica que el correo fue enviado exitosamente desde la aplicación.")
        c.save()

        output_dir = os.path.join(os.path.dirname(adjunto_path), "constancias")
        os.makedirs(output_dir, exist_ok=True)
        output_pdf_path = os.path.join(output_dir, f'constancia_{os.path.splitext(os.path.basename(adjunto_path))[0]}.pdf')

        merger = PdfMerger()
        merger.append(constancia_pdf.name)
        merger.append(adjunto_path)
        merger.write(output_pdf_path)
        merger.close()

        constancia_pdf.close()
        try:
            os.unlink(constancia_pdf.name)
        except Exception as e:
            logger.warning(f"Could not delete temporary constancia PDF {constancia_pdf.name}: {e}")
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
        base_filename = "logerror"
        extension = ".xlsx"
        attempt = 0
        error_file_saved = None
        # datetime is already imported at the top of the file
        while attempt < 5:
            try:
                if attempt ==.py
import os
import time
import smtplib
import pandas as pd
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr, make_msgid
import openpyxl
from openpyxl import Workbook
import logging
# Updated import: LOTE_SIZE is still needed, load_email_templates is used in constructor
from .config import SMTP_SERVER, SMTP_PORT, EMAIL_USER, EMAIL_PASSWORD, load_email_templates, LOTE_SIZE
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from PyPDF2 import PdfMerger

logger = logging.getLogger(__name__)

class EmailSender:
    def __init__(self, template_name="default"): # Added template_name parameter
        # Load the specified template upon instantiation
        try:
            self.subject_template, self.body_template = load_email_templates(template_name)
            logger.info(f"EmailSender initialized with template: '{template_name}'")
        except Exception as e:
            logger.error(f"Failed to load template '{template_name}' during EmailSender init: {e}. Using fallback defaults.")
            # Fallback to hardcoded defaults if load_email_templates fails critically
            self.subject_template = "Boleta del mes de {MES}"
            self.body_template = ("<html><body><h2>Hola {NOMBRE},</h2>"
                                      "<p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p>"
                                      "<p>Saludos cordiales,<br>Clínica Santa Rosa</p>"
                                       "<p>Favor de confirmar la recepciòn de este correo.</p>"
                                       "</mody></html>")

    def is_valid_email(self, email):
        import re
        # Pattern correctly defined as a raw string for regex
        pattern = r"^[\\w\\.\\-]+@[\\w\\.\\-]+\\.\\w+$"
        return re.match(pattern, email) is not None

    def is_valid_dni(self, dni):
        return dni.isdigit() and len(dni) == 8

    def send_batch(self, recipients, mes, path_boletas, progress_callback=None):
        successful_sends = []
        errores = []
        total_recipients = len(recipients)

        def enviar_lote(lote, start_index):
            try:
                if progress_callback:
                    progress_callback(f"◉ Conectando al servidor SMTP...")
                server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
                server.login(EMAIL_USER, EMAIL_PASSWORD)
                logger.info("Conexión SMTP establecida para el lote.")
            except Exception as e:
                logger.error(f'Error con servidor SMTP al conectar lote: {str(e)}')
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
                    progress_callback(f"↳ Enviando correo {i+1} de {total_recipients}...")

                nombre = recipient["nombre"]
                email = str(recipient["email"]).strip()
                dni = str(recipient["dni"]).strip()

                if not self.is_valid_dni(dni):
                    errores.append((i, nombre, email, dni, "DNI Inválido"))
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
                # Use placeholders directly in replace, no f-string needed here for self.subject_template
                msg["Subject"] = self.subject_template.replace("{MES}", mes.capitalize()).replace("{NOMBRE}", nombre)
                msg.add_header('Disposition-Notification-To', EMAIL_USER)
                msg_id = make_msgid()
                msg['Message-ID'] = msg_id

                # Same for self.body_template
                html_content = self.body_template.replace("{NOMBRE}", nombre).replace("{MES}", mes.capitalize())
                html_content += f"<br><br><small><b>Identificador de envío (Message-ID): {msg_id}</b></small>"
                msg.attach(MIMEText(html_content, "html"))

                try:
                    with open(pdf_path, "rb") as f:
                        part = MIMEApplication(f.read(), _subtype="pdf")
                        part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
                        msg.attach(part)

                    server.send_message(msg)
                    successful_sends.append(({
                        'nombre': nombre,
                        'email': email,
                        'dni': dni,
                        'pdf_path': pdf_path,
                        'message_id': msg_id,
                        'asunto': msg["Subject"],
                        'cuerpo_html': html_content
                    }))
                    time.sleep(1.5)
                except Exception as e:
                    errores.append((i, nombre, email, dni, f'Error SMTP al enviar: {str(e)}'))

            try:
                server.quit()
                logger.info("Conexión SMTP cerrada para el lote.")
            except Exception as e:
                logger.warning(f"Error al cerrar conexión SMTP: {e}")

            time.sleep(2)

        for batch_start in range(0, total_recipients, LOTE_SIZE):
            batch = recipients[batch_start:batch_start + LOTE_SIZE]
            enviar_lote(batch, batch_start + 0) # Fixed: start_index should be batch_start

        return successful_sends, errores

    def generar_constancia_envio(self, remitente, destinatario, asunto, cuerpo, fecha_envio, adjunto_path, message_id=None):
        from reportlab.lib.pagesizes import letter
        from reportlab.pdfgen import canvas
        from reportlab.lib.utils import ImageReader # Ensure this is imported
        from PyPDF2 import PdfMerger # Ensure this is imported
        import tempfile
        import shutil # Keep for good measure, though os.unlink is used
        constancia_pdf = tempfile.NamedTemporaryFile(delete=False, suffix="_constancia.pdf")
        c = canvas.Canvas(constancia_pdf.name, pagesize=letter)
        width, height = letter
        # Correct path assuming 'img' is at the root, and this file is in 'core'
        logo_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'img', 'logocsr.png')
        if os.path.exists(logo_path):
            try:
                logo_img = ImageReader(logo_path)
                logo_width = 60
                logo_height = 60
                c.drawImage(logo_img, width/2-logo_width/2, height-80, width=logo_width, height=logo_height, mask='auto')
                y = height - 100
            except Exception as img_e: # Catch error if image is problematic
                logger.error(f'Error loading logo for constancia: {img_e}')
                y = height - 40 # Default position if logo fails
        else:
            logger.warning(f'Logo not found at {logo_path} for constancia.')
            y = height - 40

        c.setFont("Helvetica-Bold", 16)
        c.drawCentredString(width/2, y, "CONSTANCIA DE ENVIO DE CORREO ELECTRÓNICO")
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
        import re # Ensure re is imported
        import textwrap # Ensure textwrap is imported
        # Use literal \n for splitlines, after replacing HTML breaks
        cuerpo_html_newlines = cuerpo.replace('<br>', '\\n').replace('<br/>', '\\n').replace('<br />', '\\n') # Python literal \\n
        cuerpo_html_newlines = re.sub('<p>', '\\n', cuerpo_html_newlines, flags=re.IGNORECASE)
        cuerpo_html_newlines = cuerpo_html_newlines.replace('</p>', '\\n')
        cuerpo_texto = re.sub('<[^<]+?>', '', cuerpo_html_newlines) # Remove HTML tags

        max_width = 90
        for line in cuerpo_texto.split('\\n'): # Split by literal '\n'
            if not line.strip():
                y -= 10
                continue
            wrapped_lines = textwrap.wrap(line.strip(), width=max_width, break_long_words=False, replace_whitespace=False)
            for wrapped in wrapped_lines:
                if y < 60: # Check if new page is needed
                    c.showPage()
                    y = height - 40 # Reset y position
                    c.setFont("Helvetica", 10) # Reset font
                c.drawString(50, y, wrapped)
                y -= 15
            if wrapped_lines:
                y -= 5
        y -= 10
        c.setFont("Helvetica-Oblique", 9)
        c.drawString(40, y, "Este documento certifica que el correo fue enviado exitosamente desde la aplicación.")
        c.save()

        output_dir = os.path.join(os.path.dirname(adjunto_path), "constancias")
        os.makedirs(output_dir, exist_ok=True)
        output_pdf_path = os.path.join(output_dir, f'constancia_{os.path.splitext(os.path.basename(adjunto_path))[0]}.pdf')

        merger = PdfMerger()
        merger.append(constancia_pdf.name)
        merger.append(adjunto_path)
        merger.write(output_pdf_path)
        merger.close()

        constancia_pdf.close()
        try:
            os.unlink(constancia_pdf.name)
        except Exception as e:
            logger.warning(f"Could not delete temporary constancia PDF {constancia_pdf.name}: {e}")
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
        base_filename = "logerror"
        extension = ".xlsx"
        attempt = 0
        error_file_saved = None
        # datetime is already imported at the top of the file
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
                time.sleep(0.1) # Brief pause before retry
                continue
            except Exception as e:
                logger.error(f"Failed to save error log '{filename}': {e}")
                break
        return error_file_saved
