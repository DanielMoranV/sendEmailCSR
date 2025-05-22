import smtplib
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
import os
import time
from datetime import datetime
import logging
from dotenv import load_dotenv
import argparse

# Cargar variables de entorno
load_dotenv()

class BulkEmailSender:
    def __init__(self, smtp_server, smtp_port, email, password):
        self.smtp_server = smtp_server
        self.smtp_port = smtp_port
        self.email = email
        self.password = password
        self.server = None

        # Logging
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler('email_log.txt'),
                logging.StreamHandler()
            ]
        )
        self.logger = logging.getLogger(__name__)
    
    def connect(self):
        try:
            self.server = smtplib.SMTP_SSL(self.smtp_server, self.smtp_port)
            self.server.login(self.email, self.password)
            self.logger.info("Conexión SMTP establecida exitosamente")
            return True
        except Exception as e:
            self.logger.error(f"Error al conectar con SMTP: {str(e)}")
            return False
    
    def disconnect(self):
        if self.server:
            self.server.quit()
            self.logger.info("Conexión SMTP cerrada")
    
    def create_message(self, recipient_email, recipient_name, subject, body_html, body_text, pdf_path=None):
        msg = MIMEMultipart('alternative')
        msg['From'] = formataddr(('Soporte CSR', self.email))
        msg['To'] = formataddr((recipient_name, recipient_email))
        msg['Subject'] = subject

        body_html_personalized = body_html.replace('{nombre}', recipient_name)
        body_text_personalized = body_text.replace('{nombre}', recipient_name)

        msg.attach(MIMEText(body_text_personalized, 'plain', 'utf-8'))
        msg.attach(MIMEText(body_html_personalized, 'html', 'utf-8'))

        if pdf_path and os.path.exists(pdf_path):
            try:
                with open(pdf_path, 'rb') as pdf_file:
                    pdf_attachment = MIMEApplication(pdf_file.read(), _subtype='pdf')
                    pdf_attachment.add_header('Content-Disposition', 'attachment', filename=os.path.basename(pdf_path))
                    msg.attach(pdf_attachment)
                self.logger.info(f"PDF adjuntado: {pdf_path}")
            except Exception as e:
                self.logger.error(f"Error al adjuntar PDF: {str(e)}")
        else:
            self.logger.warning(f"Archivo PDF no encontrado: {pdf_path}")

        return msg
    
    def send_bulk_emails(self, recipients_file, subject, body_html, body_text, mes, delay=1):
        try:
            df = pd.read_csv(recipients_file)
            recipients = df.to_dict('records')
            self.logger.info(f"Cargados {len(recipients)} destinatarios")
        except Exception as e:
            self.logger.error(f"Error al leer archivo de destinatarios: {str(e)}")
            return

        if not self.connect():
            return

        successful_sends = 0
        failed_sends = 0

        try:
            for recipient in recipients:
                dni = str(recipient['dni'])
                pdf_path = os.path.join(mes, f"{dni}.pdf")

                try:
                    msg = self.create_message(
                        recipient['email'],
                        recipient['nombre'],
                        subject,
                        body_html,
                        body_text,
                        pdf_path
                    )

                    self.server.send_message(msg)
                    successful_sends += 1
                    self.logger.info(f"Correo enviado a: {recipient['email']}")
                    time.sleep(delay)

                except Exception as e:
                    failed_sends += 1
                    self.logger.error(f"Error enviando a {recipient['email']}: {str(e)}")
                    continue

        finally:
            self.disconnect()

        self.logger.info(f"Envío completado - Exitosos: {successful_sends}, Fallidos: {failed_sends}")

# Main con argumento mes
def main():
    parser = argparse.ArgumentParser(description="Envía correos masivos con PDF por mes.")
    parser.add_argument('--mes', required=True, help='Mes de los PDFs (ej: enero, febrero)')
    args = parser.parse_args()

    smtp_server = os.getenv('SMTP_SERVER')
    smtp_port = os.getenv('SMTP_PORT')

    email = os.getenv('EMAIL_USER')
    password = os.getenv('EMAIL_PASSWORD')

    if not email or not password:
        print("Error: No se encontraron las credenciales en el archivo .env")
        return

    sender = BulkEmailSender(smtp_server, smtp_port, email, password)

    subject = "Asunto del correo masivo"
    body_html = """
    <html><body>
        <h2>Hola {nombre},</h2>
        <p>Adjuntamos el documento correspondiente al mes solicitado.</p>
        <p>Saludos cordiales,<br>Tu equipo</p>
    </body></html>
    """
    body_text = """
    Hola {nombre},

    Adjuntamos el documento correspondiente al mes solicitado.

    Saludos cordiales,
    Tu equipo
    """

    sender.send_bulk_emails(
        recipients_file="destinatarios.csv",
        subject=subject,
        body_html=body_html,
        body_text=body_text,
        mes=args.mes,
        delay=2
    )

if __name__ == "__main__":
    main()
