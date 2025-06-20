import os
import dotenv
from dotenv import load_dotenv
import sys

load_dotenv()

SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWOJD")
DEFAULT_PATH = os.path.join(os.path.expanduser("~"), "BoletasCSR")
EMAIL_CONFIG_FILE = "email_config.txt"  # Old config file, for fallback
TEMPLATES_DIR = "templates"
LOTE_SIZE = 30

if not os.path.exists(TEMPLATES_DIR):
    try:
        os.makedirs(TEMPLATES_DIR)
        print(f"Created missing templates directory: {TEMPLATES_DIR}", file=sys.stderr)
    except OSError as e:
        print(f"Error creating {TEMPLATES_DIR} on load: {e}", file=sys.stderr)

def get_template_names():
    global TEMPLATES_DIR    # i added this line
    if not os.path.isdir(TEMPLATES_DIR):
        print(f"Templates directory '{TEMPLATES_DIR}' not found.", file=sys.stderr)
        return []
    try:
        files = [f for f in os.listdir(TEMPLATES_DIR) if os.path.isfile(os.path.join(TEMPLATES_DIR, f)) and f.endswith(".txt")]
        return sorted([os.path.splitext(f)[0] for f in files])
    except Exception as e:
        print(f"Error listing templates in '{TEMPLATES_DIR}': {e}", file=sys.stderr)
        return []

def load_email_templates(template_name="default"):
    default_subject = "Boleta del mes de {MES}"
    default_body = ("<html><body><h2>Hola {NOMBRE},</h2>"
                    "<p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p>"
                    "<p>Saludos cordiales,<br>Clínica Santa Rosa</p>"
                    "<p>Favor de confirmar la recepción de este correo.</p>"
                    "</body></html>")
    subject, body = default_subject, default_body
    if not template_name:
        print("Warning: No template name. Using defaults.", file=sys.stderr)
        return subject, body

    template_path = os.path.join(TEMPLATES_DIR, f"{template_name}.txt")
    source_desc = f"template '{template_path}'"

    file_to_load = None
    if os.path.exists(template_path):
        file_to_load = template_path
    elif template_name == "default" and os.path.exists(EMAIL_CONFIG_FILE):
        print(f'Warning: \'{template_path}\' not found. Falling back to \'{EMAIL_CONFIG_FILE}\'.', file=sys.stderr)
        file_to_load = EMAIL_CONFIG_FILE
        source_desc = f"fallback '{EMAIL_CONFIG_FILE}'"

    if file_to_load:
        try:
            with open(file_to_load, "r", encoding="utf-8") as f:
                content = f.read()

            parsed_subject, parsed_body = None, None
            subject_marker, body_marker = "SUBJECT=", "BODY="

            s_idx, b_idx = content.find(subject_marker), content.find(body_marker)

            if s_idx != -1:
                val_start = s_idx + len(subject_marker)
                end_idx = b_idx if (b_idx > s_idx) else len(content)
                # Subject is first line after marker, or up to BODY= if on same line.
                first_newline = content.find('\n', val_start)
                if first_newline != -1 and first_newline < end_idx : # Subject is a single line
                     parsed_subject = content[val_start:first_newline].strip()
                else: # Subject until body_marker or end of content
                     parsed_subject = content[val_start:end_idx].strip()

            if b_idx != -1:
                parsed_body = content[b_idx + len(body_marker):].strip()

            if parsed_subject is not None: subject = parsed_subject
            if parsed_body is not None: body = parsed_body

            if parsed_subject is not None or parsed_body is not None:
                print(f"Loaded from {source_desc}.", file=sys.stderr)
            else: # File existed but markers not found.
                print(f'Warning: Markers not found in {source_desc}. Using defaults.', file=sys.stderr)

        except Exception as e:
            print(f"Error reading {source_desc}: {e}. Using defaults.", file=sys.stderr)
    else:
        print(f"Warning: No template file found for '{template_name}'. Using defaults.", file=sys.stderr)
    return subject, body

def save_email_template(template_name, subject, body):
    if not (template_name and isinstance(template_name, str) and template_name.strip()):
        raise ValueError("Template name invalid.")
    if any(c in template_name for c in "/\\.."):
        raise ValueError("Template name has invalid chars.")
    if not os.path.isdir(TEMPLATES_DIR): os.makedirs(TEMPLATES_DIR)

    path = os.path.join(TEMPLATES_DIR, f"{template_name}.txt")
    try:
        with open(path, 'w', encoding="utf-8") as f:
            f.write(f"SUBJECT={subject}\nBODY={body}") # \n to separate, body can be multi-line
        print(f'Saved \'{template_name}\' to {path}.', file=sys.stderr)
    except Exception as e:
        raise IOError(f"Error saving '{template_name}': {e}")

def delete_email_template(template_name):
    if not (template_name and isinstance(template_name, str) and template_name.strip()):
        raise ValueError("Template name invalid.")
    if any(c in template_name for c in "/\\.."):
        raise ValueError("Template name has invalid chars.")

    path = os.path.join(TEMPLATES_DIR, f"{template_name}.txt")
    try:
        if os.path.exists(path):
            os.remove(path)
            print(f'Deleted \'{template_name}\' from {path}.', file=sys.stderr)
        else:
            print(f"Warning: Delete failed, '{path}' not found.", file=sys.stderr)
    except Exception as e:
        raise IOError(f"Error deleting '{template_name}': {e}")
