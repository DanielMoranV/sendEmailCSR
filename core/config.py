import os
from dotenv import load_dotenv
import sys # Import sys module

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")
DEFAULT_PATH = "C:/BoletasCSR"
EMAIL_CONFIG_FILE = "email_config.txt"
LOTE_SIZE = 30 # Default batch size for sending emails


def load_email_templates():
    subject = "Boleta del mes de {MES}"
    body = """
        <html><body>
        <h2>Hola {NOMBRE},</h2>
        <p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p>
        <p>Saludos cordiales,<br>Clínica Santa Rosa</p>
        <p>Favor de confirmar la recepción de este correo.</p>
        </body></html>"""
    if os.path.exists(EMAIL_CONFIG_FILE):
        try:
            with open(EMAIL_CONFIG_FILE, 'r', encoding='utf-8') as f:
                for line in f:
                    line = line.strip()
                    if line.startswith("SUBJECT="):
                        subject = line.split("=", 1)[1]
                    elif line.startswith("BODY="):
                        body = line.split("=", 1)[1]
        except Exception as e:
            print(f"Error loading email templates from {EMAIL_CONFIG_FILE}: {e}. Using default templates.", file=sys.stderr)
            # Still return default subject and body if an error occurs
    return subject, body
