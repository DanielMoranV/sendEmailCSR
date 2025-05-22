import os
import time
import smtplib
import pandas as pd
from datetime import datetime
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from email.utils import formataddr
from tkinter import filedialog, StringVar, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from dotenv import load_dotenv
import logging
import re
from openpyxl import Workbook
import threading

# Configurar logs
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s - %(levelname)s - %(message)s",
    handlers=[logging.FileHandler("email_log.txt"), logging.StreamHandler()],
)
logger = logging.getLogger(__name__)

# Cargar variables del entorno
load_dotenv()
SMTP_SERVER = os.getenv("SMTP_SERVER")
SMTP_PORT = int(os.getenv("SMTP_PORT", 465))
EMAIL_USER = os.getenv("EMAIL_USER")
EMAIL_PASSWORD = os.getenv("EMAIL_PASSWORD")

DEFAULT_PATH = "C:/BoletasCSR"
EMAIL_CONFIG_FILE = "email_config.txt" # Define el nombre del archivo de configuración

class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Envío de Boletas")
        self.root.geometry("320x450")

        self.current_month = datetime.now().strftime("%m")
        self.months = [
            ("01", "enero"), ("02", "febrero"), ("03", "marzo"), ("04", "abril"),
            ("05", "mayo"), ("06", "junio"), ("07", "julio"), ("08", "agosto"),
            ("09", "septiembre"), ("10", "octubre"), ("11", "noviembre"), ("12", "diciembre")
        ]

        self.mes_var = StringVar(value=next(name for num, name in self.months if num == self.current_month))
        self.path_var = StringVar(value=DEFAULT_PATH)
        self.excel_path_var = StringVar(value="")
        self.status_var = StringVar(value="")
        self.progress_var = StringVar(value="")

        # Variable para controlar el estado de procesamiento
        self.is_processing = False

        # Cargar configuración de email
        self.email_subject_template = f"Boleta del mes de {{MES}}" # Valor por defecto
        self.email_body_template = """
            <html><body>
            <h2>Hola {NOMBRE},</h2>
            <p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p>
            <p>Saludos cordiales,<br>Clínica Santa Rosa</p>
            <p>Favor de confirmar la recepción de este correo.</p>
            </body></html>""" # Valor por defecto
        self.load_email_config() # Llama a la nueva función para cargar la configuración

        self.build_gui()

    def load_email_config(self):
        """Carga la configuración del asunto y cuerpo del email desde un archivo."""
        config_path = EMAIL_CONFIG_FILE
        if os.path.exists(config_path):
            try:
                with open(config_path, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("SUBJECT="):
                            self.email_subject_template = line.split("=", 1)[1]
                            logger.info(f"Asunto de email cargado: {self.email_subject_template}")
                        elif line.startswith("BODY="):
                            self.email_body_template = line.split("=", 1)[1]
                            logger.info(f"Cuerpo de email cargado.")
            except Exception as e:
                logger.error(f"Error al leer el archivo de configuración de email '{config_path}': {e}")
                self.update_status(f"⚠️ Error al leer configuración de email. Usando valores por defecto.", "warning")
        else:
            logger.info(f"Archivo de configuración de email '{config_path}' no encontrado. Usando valores por defecto.")


    def build_gui(self):
        container = tb.Frame(self.root, padding=10)
        container.pack(fill=BOTH, expand=True)

        tb.Label(container, text="Seleccionar Mes:", font=("Segoe UI", 11)).pack(anchor=W)
        tb.Combobox(container, values=[m[1] for m in self.months], textvariable=self.mes_var, font=("Segoe UI", 10), width=40).pack()

        tb.Label(container, text="Ruta de Boletas:", font=("Segoe UI", 11)).pack(anchor=W, pady=(10, 0))
        tb.Entry(container, textvariable=self.path_var, width=55).pack()

        tb.Label(container, text="Archivo Excel de Empleados:", font=("Segoe UI", 11)).pack(anchor=W, pady=(10, 0))
        self.select_excel_btn = tb.Button(container, text="Seleccionar Archivo Excel", command=self.select_excel, bootstyle=INFO, width=180)
        self.select_excel_btn.pack(pady=5)
        tb.Entry(container, textvariable=self.excel_path_var, width=55).pack()

        # Botón de envío (guardamos referencia para habilitarlo/deshabilitarlo)
        self.send_btn = tb.Button(container, text="Enviar Boletas", command=self.start_email_process, bootstyle=SUCCESS, width=180)
        self.send_btn.pack(pady=20)

        # Etiqueta de progreso (para mostrar el progreso durante el envío)
        self.progress_label = tb.Label(container, textvariable=self.progress_var, font=("Segoe UI", 10))
        self.progress_label.pack()

        # Barra de progreso
        self.progress_bar = tb.Progressbar(container, mode='indeterminate', bootstyle="success-striped")
        self.progress_bar.pack(fill=X, pady=5)
        self.progress_bar.pack_forget()  # Ocultar inicialmente

        self.status_label = tb.Label(container, textvariable=self.status_var, font=("Segoe UI", 10))
        self.status_label.pack()

    def select_excel(self):
        if self.is_processing:
            return  # No permitir seleccionar archivo durante el procesamiento

        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_path_var.set(file_path)
            logger.info(f"Archivo Excel seleccionado: {file_path}")

    def start_email_process(self):
        """Inicia el proceso de envío en un hilo separado"""
        if self.is_processing:
            return

        # Deshabilitar controles
        self.set_processing_state(True)

        # Ejecutar en hilo separado para no bloquear la interfaz
        thread = threading.Thread(target=self.send_emails_thread)
        thread.daemon = True
        thread.start()

    def set_processing_state(self, processing):
        """Controla el estado de la interfaz durante el procesamiento"""
        self.is_processing = processing

        if processing:
            # Deshabilitar botones
            self.send_btn.configure(state='disabled', text="Procesando...")
            self.select_excel_btn.configure(state='disabled')

            # Mostrar barra de progreso
            self.progress_bar.pack(fill=X, pady=5)
            self.progress_bar.start(10)  # Velocidad de animación

            # Limpiar estado anterior
            self.status_var.set("")
            self.progress_var.set("Iniciando envío de correos...")

        else:
            # Habilitar botones
            self.send_btn.configure(state='normal', text="Enviar Boletas")
            self.select_excel_btn.configure(state='normal')

            # Ocultar barra de progreso
            self.progress_bar.stop()
            self.progress_bar.pack_forget()

            # Limpiar mensaje de progreso
            self.progress_var.set("")

    def update_progress(self, message):
        """Actualiza el mensaje de progreso de forma segura desde hilos"""
        self.root.after(0, lambda: self.progress_var.set(message))

    def update_status(self, message, style="secondary"):
        """Actualiza el estado final de forma segura desde hilos"""
        def update():
            self.status_var.set(message)
            self.status_label.configure(bootstyle=style)
        self.root.after(0, update)

    def send_emails_thread(self):
        """Método que ejecuta el envío de emails en un hilo separado"""
        try:
            self.send_emails()
        except Exception as e:
            logger.error(f"Error durante el procesamiento: {str(e)}")
            self.update_status(f"❌ Error inesperado: {str(e)}", "danger")
        finally:
            # Siempre restaurar el estado de la interfaz
            self.root.after(0, lambda: self.set_processing_state(False))

    def is_valid_email(self, email):
        pattern = r"^[\w\.-]+@[\w\.-]+\.\w+$"
        return re.match(pattern, email) is not None

    def is_valid_dni(self, dni):
        # Verificar que solo contenga dígitos y tenga exactamente 8 caracteres
        return dni.isdigit() and len(dni) == 8

    def send_emails(self):
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            self.update_status("❌ Error: Debes seleccionar un archivo Excel válido.", "danger")
            return

        self.update_progress("📖 Leyendo archivo Excel...")

        try:
            df = pd.read_excel(excel_path, dtype={"dni": str})  # tratar DNI como texto
        except Exception as e:
            logger.error(f"Error al leer el archivo Excel: {str(e)}")
            self.update_status(f"❌ Error al leer Excel: {str(e)}", "danger")
            return

        required_columns = {"nombre", "email", "dni"}
        df.columns = df.columns.str.lower()
        if not required_columns.issubset(df.columns):
            self.update_status("❌ Error: El Excel debe contener: nombre, email y dni.", "danger")
            return

        df = df[["nombre", "email", "dni"]].dropna(subset=["nombre", "email", "dni"])
        recipients = df.to_dict("records")

        if not recipients:
            self.update_status("❌ Sin registros válidos en el archivo Excel.", "danger")
            return

        self.update_progress("🔗 Conectando al servidor SMTP...")

        try:
            server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
            server.login(EMAIL_USER, EMAIL_PASSWORD)
            logger.info("Conexión SMTP establecida.")
        except Exception as e:
            logger.error(f"Error con servidor SMTP: {str(e)}")
            self.update_status(f"❌ Error SMTP: {str(e)}", "danger")
            return

        enviados = 0
        errores = []
        total_recipients = len(recipients)

        for i, recipient in enumerate(recipients, start=2):  # Fila inicia en 2 (por encabezado)
            # Actualizar progreso
            self.update_progress(f"📧 Enviando correo {i-1} de {total_recipients}...")

            nombre = recipient["nombre"]
            email = str(recipient["email"]).strip()
            dni = str(recipient["dni"]).strip()

            # Validaciones
            if not self.is_valid_dni(dni):
                error = f"Fila {i} - DNI inválido: {dni}"
                errores.append((i, nombre, email, dni, "DNI inválido"))
                logger.warning(error)
                continue
            if not self.is_valid_email(email):
                error = f"Fila {i} - Email inválido: {email}"
                errores.append((i, nombre, email, dni, "Email inválido"))
                logger.warning(error)
                continue

            # Solo después de validar, formatear el DNI para buscar el archivo
            dni_formatted = dni.zfill(8)
            pdf_path = os.path.join(self.path_var.get(), self.mes_var.get(), f"{dni_formatted}.pdf")
            if not os.path.exists(pdf_path):
                errores.append((i, nombre, email, dni, "PDF no encontrado"))
                logger.warning(f"PDF no encontrado para DNI {dni} en fila {i}")
                continue

            msg = MIMEMultipart("alternative")
            msg["From"] = formataddr(("Clínica Santa Rosa", EMAIL_USER))
            msg["To"] = formataddr((nombre, email))

            # Usar las plantillas cargadas para el asunto y el cuerpo del correo
            # y reemplazar los marcadores de posición
            mes_capitalizado = self.mes_var.get().capitalize()
            msg["Subject"] = self.email_subject_template.replace("{MES}", mes_capitalizado)

            html_content = self.email_body_template.replace("{NOMBRE}", nombre).replace("{MES}", mes_capitalizado)
            msg.attach(MIMEText(html_content, "html"))

            try:
                with open(pdf_path, "rb") as f:
                    part = MIMEApplication(f.read(), _subtype="pdf")
                    part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
                    msg.attach(part)

                server.send_message(msg)
                logger.info(f"Correo enviado a {email}")
                enviados += 1
                time.sleep(1.5)
            except Exception as e:
                logger.error(f"Error enviando a {email}: {str(e)}")
                errores.append((i, nombre, email, dni, f"Error SMTP: {str(e)}"))

        server.quit()

        self.update_progress("📝 Generando reporte de errores...")

        # Guardar errores en logError.xlsx
        error_file_saved = None
        if errores:
            wb = Workbook()
            ws = wb.active
            ws.title = "Errores"
            ws.append(["Fila", "Nombre", "Email", "DNI", "Motivo"])
            for fila in errores:
                ws.append(fila)

            # Intentar guardar el archivo, si falla crear uno con timestamp
            base_filename = "logError"
            extension = ".xlsx"
            attempt = 0

            while attempt < 5:  # Máximo 5 intentos
                try:
                    if attempt == 0:
                        filename = f"{base_filename}{extension}"
                    else:
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        filename = f"{base_filename}_{timestamp}{extension}"

                    wb.save(filename)
                    error_file_saved = filename
                    logger.info(f"Errores registrados en {filename}")
                    break

                except PermissionError:
                    if attempt == 0:
                        logger.warning("logError.xlsx está abierto. Intentando crear archivo alternativo...")
                        # Usar self.root.after para mostrar mensaje desde el hilo principal
                        self.root.after(0, lambda: messagebox.showwarning("Archivo en uso",
                                                                         "El archivo logError.xlsx está abierto.\n"
                                                                         "Se creará un archivo alternativo con timestamp."))
                    attempt += 1
                    continue
                except Exception as e:
                    logger.error(f"Error al guardar archivo de errores: {str(e)}")
                    self.root.after(0, lambda: messagebox.showerror("Error", f"No se pudo guardar el archivo de errores:\n{str(e)}"))
                    break

            # Intentar abrir el archivo creado
            if error_file_saved:
                try:
                    os.startfile(error_file_saved)
                except Exception as e:
                    logger.error(f"No se pudo abrir {error_file_saved} automáticamente: {str(e)}")
            else:
                self.root.after(0, lambda: messagebox.showerror("Error", "No se pudo crear el archivo de errores después de varios intentos."))

        # Mostrar resultados finales
        total_procesados = len(recipients)
        total_enviados = enviados
        total_errores = len(errores)

        logger.info(f"Resumen: Total procesados: {total_procesados}, Enviados: {total_enviados}, Errores: {total_errores}")

        # Construir mensaje de estado en múltiples líneas
        mensaje_lineas = []

        # Línea 1: Total procesados
        mensaje_lineas.append(f"📊 Total procesados: {total_procesados}")

        # Línea 2: Enviados exitosos
        if total_enviados > 0:
            mensaje_lineas.append(f"✅ Enviados correctamente: {total_enviados}")

        # Línea 3: Errores
        if total_errores > 0:
            mensaje_lineas.append(f"❌ Errores encontrados: {total_errores}")

        # Línea 4: Archivo de log (si existe)
        if error_file_saved:
            mensaje_lineas.append(f"📄 Ver detalles en: {error_file_saved}")

        # Unir las líneas con saltos de línea
        mensaje_final = "\n".join(mensaje_lineas)

        # Configurar color según el resultado
        if total_enviados > 0 and total_errores > 0:
            style = "warning"
        elif total_enviados > 0 and total_errores == 0:
            style = "success"
        elif total_enviados == 0 and total_errores > 0:
            style = "danger"
        else:
            style = "secondary"
            mensaje_final = "No se procesaron correos"

        self.update_status(mensaje_final, style)

if __name__ == "__main__":
    app = tb.Window(themename="superhero")
    gui = EmailSenderGUI(app)
    app.mainloop()