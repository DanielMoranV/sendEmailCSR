import os
import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import filedialog, StringVar, messagebox
import threading
import pandas as pd
from core.email_sender import EmailSender
from core.config import DEFAULT_PATH

class EmailSenderGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Envío de Boletas")
        self.root.geometry("320x450")

        self.current_month = pd.Timestamp.now().strftime("%m")
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
        self.is_processing = False

        self.sender = EmailSender()
        self.build_gui()

    def build_gui(self):
        container = tb.Frame(self.root, padding=10)
        container.pack(fill="both", expand=True)
        tb.Label(container, text="Seleccionar Mes:", font=("Segoe UI", 11)).pack(anchor="w")
        tb.Combobox(container, values=[m[1] for m in self.months], textvariable=self.mes_var, font=("Segoe UI", 10), width=40).pack()
        tb.Label(container, text="Ruta de Boletas:", font=("Segoe UI", 11)).pack(anchor="w", pady=(10, 0))
        tb.Entry(container, textvariable=self.path_var, width=55).pack()
        tb.Label(container, text="Archivo Excel de Empleados:", font=("Segoe UI", 11)).pack(anchor="w", pady=(10, 0))
        self.select_excel_btn = tb.Button(container, text="Seleccionar Archivo Excel", command=self.select_excel, bootstyle=INFO, width=180)
        self.select_excel_btn.pack(pady=5)
        tb.Entry(container, textvariable=self.excel_path_var, width=55).pack()
        self.send_btn = tb.Button(container, text="Enviar Boletas", command=self.start_email_process, bootstyle=SUCCESS, width=180)
        self.send_btn.pack(pady=20)
        self.progress_label = tb.Label(container, textvariable=self.progress_var, font=("Segoe UI", 10))
        self.progress_label.pack()
        self.progress_bar = tb.Progressbar(container, mode='indeterminate', bootstyle="success-striped")
        self.progress_bar.pack(fill="x", pady=5)
        self.progress_bar.pack_forget()
        self.status_label = tb.Label(container, textvariable=self.status_var, font=("Segoe UI", 10))
        self.status_label.pack()

    def select_excel(self):
        if self.is_processing:
            return
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        if file_path:
            self.excel_path_var.set(file_path)

    def start_email_process(self):
        if self.is_processing:
            return
        self.set_processing_state(True)
        thread = threading.Thread(target=self.send_emails_thread)
        thread.daemon = True
        thread.start()

    def set_processing_state(self, processing):
        self.is_processing = processing
        if processing:
            self.send_btn.configure(state='disabled', text="Procesando...")
            self.select_excel_btn.configure(state='disabled')
            self.progress_bar.pack(fill="x", pady=5)
            self.progress_bar.start(10)
            self.status_var.set("")
            self.progress_var.set("Iniciando envío de correos...")
        else:
            self.send_btn.configure(state='normal', text="Enviar Boletas")
            self.select_excel_btn.configure(state='normal')
            self.progress_bar.stop()
            self.progress_bar.pack_forget()
            self.progress_var.set("")

    def update_progress(self, message):
        self.root.after(0, lambda: self.progress_var.set(message))

    def update_status(self, message, style="secondary"):
        def update():
            self.status_var.set(message)
            self.status_label.configure(bootstyle=style)
        self.root.after(0, update)

    def send_emails_thread(self):
        try:
            self.send_emails()
        except Exception as e:
            self.update_status(f"❌ Error inesperado: {str(e)}", "danger")
        finally:
            self.root.after(0, lambda: self.set_processing_state(False))

    def send_emails(self):
        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            self.update_status("❌ Error: Debes seleccionar un archivo Excel válido.", "danger")
            return
        self.update_progress("📖 Leyendo archivo Excel...")
        try:
            df = pd.read_excel(excel_path, dtype={"dni": str})
        except Exception as e:
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
        mes = self.mes_var.get()
        path_boletas = self.path_var.get()
        # Enviar correos y recolectar constancias generadas
        constancias_generadas = []
        original_send_batch = self.sender.send_batch
        def progress_and_constancia(msg):
            self.update_progress(msg)
        # Monkeypatch para capturar constancias (usamos un wrapper temporal)
        def send_batch_with_constancia(recipients, mes, path_boletas, progress_callback=None):
            LOTE_SIZE = 30
            enviados = 0
            errores = []
            total_recipients = len(recipients)
            from core.email_sender import EmailSender
            sender = self.sender
            def enviar_lote(lote, start_index):
                nonlocal enviados
                import os
                import time
                from datetime import datetime
                try:
                    if progress_callback:
                        progress_callback(f"🔗 Conectando al servidor SMTP...")
                    import smtplib
                    from core.config import SMTP_SERVER, SMTP_PORT, EMAIL_USER, EMAIL_PASSWORD
                    server = smtplib.SMTP_SSL(SMTP_SERVER, SMTP_PORT)
                    server.login(EMAIL_USER, EMAIL_PASSWORD)
                except Exception as e:
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
                    if not sender.is_valid_dni(dni):
                        errores.append((i, nombre, email, dni, "DNI inválido"))
                        continue
                    if not sender.is_valid_email(email):
                        errores.append((i, nombre, email, dni, "Email inválido"))
                        continue
                    dni_formatted = dni.zfill(8)
                    pdf_path = os.path.join(path_boletas, mes, f"{dni_formatted}.pdf")
                    if not os.path.exists(pdf_path):
                        errores.append((i, nombre, email, dni, "PDF no encontrado"))
                        continue
                    # Preparar datos para constancia
                    msg_asunto = sender.subject_template.replace("{MES}", mes.capitalize())
                    html_content = sender.body_template.replace("{NOMBRE}", nombre).replace("{MES}", mes.capitalize())
                    try:
                        # Envío real
                        from email.mime.multipart import MIMEMultipart
                        from email.mime.text import MIMEText
                        from email.mime.application import MIMEApplication
                        from email.utils import formataddr
                        msg = MIMEMultipart("alternative")
                        msg["From"] = formataddr(("Clínica Santa Rosa", sender.EMAIL_USER if hasattr(sender, 'EMAIL_USER') else ""))
                        msg["To"] = formataddr((nombre, email))
                        msg["Subject"] = msg_asunto
                        msg.add_header('Disposition-Notification-To', sender.EMAIL_USER if hasattr(sender, 'EMAIL_USER') else "")
                        msg.attach(MIMEText(html_content, "html"))
                        with open(pdf_path, "rb") as f:
                            part = MIMEApplication(f.read(), _subtype="pdf")
                            part.add_header("Content-Disposition", "attachment", filename=os.path.basename(pdf_path))
                            msg.attach(part)
                        server.send_message(msg)
                        enviados += 1
                        # Generar constancia PDF
                        fecha_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        constancia_path = sender.generar_constancia_envio(
                            remitente=("Clínica Santa Rosa", sender.EMAIL_USER if hasattr(sender, 'EMAIL_USER') else ""),
                            destinatario=(nombre, email),
                            asunto=msg_asunto,
                            cuerpo=html_content,
                            fecha_envio=fecha_envio,
                            adjunto_path=pdf_path
                        )
                        constancias_generadas.append(constancia_path)
                        # Eliminado messagebox para no molestar al usuario
                        time.sleep(1.5)
                    except Exception as e:
                        errores.append((i, nombre, email, dni, f"Error SMTP: {str(e)}"))
                server.quit()
                time.sleep(2)
            for batch_start in range(0, total_recipients, LOTE_SIZE):
                batch = recipients[batch_start:batch_start + LOTE_SIZE]
                enviar_lote(batch, batch_start + 2)
            return enviados, errores
        self.sender.send_batch = send_batch_with_constancia
        enviados, errores = self.sender.send_batch(
            recipients, mes, path_boletas, progress_callback=self.update_progress
        )
        # Guardar log de errores si los hay
        error_file_saved = self.sender.generar_log_errores(errores) if errores else None
        total_procesados = len(recipients)
        total_enviados = enviados
        total_errores = len(errores)
        mensaje_lineas = []
        mensaje_lineas.append(f"📊 Total procesados: {total_procesados}")
        if total_enviados > 0:
            mensaje_lineas.append(f"✅ Enviados correctamente: {total_enviados}")
        if total_errores > 0:
            mensaje_lineas.append(f"❌ Errores encontrados: {total_errores}")
        if error_file_saved:
            mensaje_lineas.append(f"📄 Ver detalles en: {error_file_saved}")
        mensaje_final = "\n".join(mensaje_lineas)
        # Determinar estilo
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
        # Mostrar resumen de constancias generadas solo una vez
        if constancias_generadas:
            carpeta_constancias = os.path.dirname(constancias_generadas[0])
            messagebox.showinfo(
                "Constancias generadas",
                f"Se generaron {len(constancias_generadas)} constancias PDF en la carpeta:\n{carpeta_constancias}"
            )
        # Si hubo errores, abrir el archivo de errores automáticamente
        if error_file_saved:
            try:
                os.startfile(error_file_saved)
            except Exception:
                pass
