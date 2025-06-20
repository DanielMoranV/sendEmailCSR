import os
import ttkbootstrap as tb
import webbrowser
from ttkbootstrap.constants import *
from tkinter import filedialog, StringVar, messagebox
from gui.email_template_modal import open_template_editor_modal
import threading
import pandas as pd
from core.email_sender import EmailSender
from core.config import DEFAULT_PATH, EMAIL_USER
from datetime import datetime

class EmailSenderGUI:
    def open_template_editor_modal(self):
        open_template_editor_modal(self.root)

    def __init__(self, root):
        self.root = root
        self.root.title("📧 Sistema de Envío de Boletas - Clínica Santa Rosa")
        self.root.geometry("950x650")  # Reducido significativamente
        self.root.resizable(True, True)
        
        # Configurar el tema
        self.style = tb.Style("flatly")

        self.current_month = pd.Timestamp.now().strftime("%m")
        self.months = [
            ("01", "Enero"), ("02", "Febrero"), ("03", "Marzo"), ("04", "Abril"),
            ("05", "Mayo"), ("06", "Junio"), ("07", "Julio"), ("08", "Agosto"),
            ("09", "Septiembre"), ("10", "Octubre"), ("11", "Noviembre"), ("12", "Diciembre")
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
        # Frame principal
        main_frame = tb.Frame(self.root, bootstyle="light")
        main_frame.pack(fill="both", expand=True, padx=8, pady=5)

        # Header compacto
        self.create_compact_header(main_frame)
        
        # Layout de tres columnas compactas
        self.create_compact_layout(main_frame)
        
        # Footer minimalista
        self.create_compact_footer(main_frame)

    def create_compact_header(self, parent):
        """Header más compacto"""
        header_frame = tb.Frame(parent, bootstyle="info", padding=8)
        header_frame.pack(fill="x", pady=(0, 8))
        
        title_label = tb.Label(
            header_frame, 
            text="📧 Sistema de Envío de Emails - Clínica Santa Rosa",
            font=("Segoe UI", 14, "bold"),
            bootstyle="inverse-info"
        )
        title_label.pack()

    def create_compact_layout(self, parent):
        """Layout de tres columnas compactas"""
        # Frame para las tres columnas
        columns_frame = tb.Frame(parent)
        columns_frame.pack(fill="both", expand=True)
        
        # Columna 1 - Configuración (35%)
        config_column = tb.Frame(columns_frame, padding=5)
        config_column.pack(side="left", fill="both", expand=True, padx=(0, 5))
        
        # Columna 2 - Acciones (30%)
        actions_column = tb.Frame(columns_frame, padding=5)
        actions_column.pack(side="left", fill="both", expand=True, padx=2)
        
        # Columna 3 - Estado (35%)
        status_column = tb.Frame(columns_frame, padding=5)
        status_column.pack(side="left", fill="both", expand=True, padx=(5, 0))
        
        # Llenar columnas
        self.create_config_column(config_column)
        self.create_actions_column(actions_column)
        self.create_status_column(status_column)

    def create_config_column(self, parent):
        """Columna de configuración compacta"""
        # Sección de mes
        month_frame = tb.LabelFrame(parent, text="📅 Mes", bootstyle="info", padding=8)
        month_frame.pack(fill="x", pady=(0, 8))
        
        self.month_combo = tb.Combobox(
            month_frame,
            values=[m[1] for m in self.months],
            textvariable=self.mes_var,
            font=("Segoe UI", 10),
            bootstyle="info",
            state="readonly"
        )
        self.month_combo.pack(fill="x")

        # Sección de archivos
        files_frame = tb.LabelFrame(parent, text="📁 Archivos", bootstyle="warning", padding=8)
        files_frame.pack(fill="x", pady=(0, 8))
        
        # Directorio de boletas
        tb.Label(files_frame, text="Directorio boletas:", font=("Segoe UI", 9)).pack(anchor="w")
        self.path_entry = tb.Entry(files_frame, textvariable=self.path_var, font=("Segoe UI", 9))
        self.path_entry.pack(fill="x", pady=(2, 8))
        
        # Excel
        self.select_excel_btn = tb.Button(
            files_frame,
            text="📂 Seleccionar Excel",
            command=self.select_excel,
            bootstyle="warning-outline"
        )
        self.select_excel_btn.pack(fill="x", pady=(0, 5))
        
        self.excel_entry = tb.Entry(
            files_frame,
            textvariable=self.excel_path_var,
            font=("Segoe UI", 8),
            state="readonly"
        )
        self.excel_entry.pack(fill="x")

        # Información compacta
        info_frame = tb.LabelFrame(parent, text="💡 Info", bootstyle="secondary", padding=8)
        info_frame.pack(fill="x", pady=(0, 8))
        
        info_text = "📋 Excel: nombre, email, dni\n📁 PDFs: {dni}.pdf"
        tb.Label(info_frame, text=info_text, font=("Segoe UI", 8), justify="left").pack(anchor="w")

        # Botón editar plantilla
        edit_btn = tb.Button(
            parent,
            text="✏️ Editar Plantilla",
            command=self.open_template_editor_modal,
            bootstyle="info-outline"
        )
        edit_btn.pack(fill="x")

    def create_actions_column(self, parent):
        """Columna de acciones compacta"""
        actions_frame = tb.LabelFrame(parent, text="🎯 Acciones", bootstyle="success", padding=10)
        actions_frame.pack(fill="x", pady=(0, 8))
        
        # Botón principal
        self.send_btn = tb.Button(
            actions_frame,
            text="📨 ENVIAR EMAILS",
            command=self.start_email_process,
            bootstyle="success"
        )
        self.send_btn.pack(fill="x", pady=(0, 10))
        
        # Botones secundarios
        verify_btn = tb.Button(
            actions_frame,
            text="🔍 Verificar",
            command=self.verify_configuration,
            bootstyle="info-outline"
        )
        verify_btn.pack(fill="x", pady=(0, 5))
        
        clear_btn = tb.Button(
            actions_frame,
            text="🗑️ Limpiar",
            command=self.clear_data,
            bootstyle="secondary-outline"
        )
        clear_btn.pack(fill="x")

        # Estadísticas compactas
        stats_frame = tb.LabelFrame(parent, text="📈 Estadísticas", bootstyle="secondary", padding=8)
        stats_frame.pack(fill="x")
        
        # Crear grid de estadísticas
        stats_grid = tb.Frame(stats_frame)
        stats_grid.pack(fill="x")
        
        # Inicializar variables de estadísticas
        self.stat_vars = getattr(self, 'stat_vars', {})
        for key in ["procesados", "enviados", "errores", "tiempo"]:
            if key not in self.stat_vars:
                self.stat_vars[key] = StringVar(value="0" if key != "tiempo" else "0s")
        
        # Crear labels de estadísticas en grid 2x2
        self.create_compact_stat_item(stats_grid, "📧 Procesados", self.stat_vars["procesados"], 0, 0)
        self.create_compact_stat_item(stats_grid, "✅ Enviados", self.stat_vars["enviados"], 0, 1)
        self.create_compact_stat_item(stats_grid, "❌ Errores", self.stat_vars["errores"], 1, 0)
        self.create_compact_stat_item(stats_grid, "⏱️ Tiempo", self.stat_vars["tiempo"], 1, 1)

    def create_compact_stat_item(self, parent, emoji, var, row, col):
        """Crear estadística compacta"""
        frame = tb.Frame(parent)
        frame.grid(row=row, column=col, padx=2, pady=1, sticky="ew")
        parent.grid_columnconfigure(col, weight=1)
        
        tb.Label(frame, text=emoji, font=("Segoe UI", 10)).pack(side="left")
        tb.Label(frame, textvariable=var, font=("Segoe UI", 9, "bold")).pack(side="right")

    def create_status_column(self, parent):
        """Columna de estado compacta"""
        # Progreso
        progress_frame = tb.LabelFrame(parent, text="📊 Progreso", bootstyle="primary", padding=8)
        progress_frame.pack(fill="x", pady=(0, 8))
        
        self.progress_label = tb.Label(
            progress_frame,
            textvariable=self.progress_var,
            font=("Segoe UI", 9),
            wraplength=200
        )
        self.progress_label.pack(anchor="w", pady=(0, 5))
        
        self.progress_bar = tb.Progressbar(
            progress_frame,
            mode='indeterminate',
            bootstyle="success-striped"
        )
        self.progress_bar.pack(fill="x")

        # Estado
        status_frame = tb.LabelFrame(parent, text="📋 Estado", bootstyle="primary", padding=8)
        status_frame.pack(fill="both", expand=True)
        
        # Frame con scroll para el estado
        status_container = tb.Frame(status_frame)
        status_container.pack(fill="both", expand=True)
        
        self.status_label = tb.Label(
            status_container,
            textvariable=self.status_var,
            font=("Segoe UI", 9),
            wraplength=250,
            justify="left",
            anchor="nw"
        )
        self.status_label.pack(fill="both", expand=True, anchor="nw")

    def create_compact_footer(self, parent):
        """Footer compacto"""
        footer_frame = tb.Frame(parent, bootstyle="dark", padding=5)
        footer_frame.pack(fill="x", side="bottom")
        
        footer_label = tb.Label(
            footer_frame,
            text="🏥 Clínica Santa Rosa | v2.0 | soporteti@csr.pe.com",
            font=("Segoe UI", 8),
            bootstyle="inverse-dark"
        )
        footer_label.pack()

    def update_stats(self, procesados, enviados, errores, tiempo):
        """Actualizar estadísticas"""
        self.stat_vars["procesados"].set(str(procesados))
        self.stat_vars["enviados"].set(str(enviados))
        self.stat_vars["errores"].set(str(errores))
        self.stat_vars["tiempo"].set(str(tiempo))

    def verify_configuration(self):
        """Verificar configuración"""
        issues = []
        
        if not self.excel_path_var.get():
            issues.append("• No hay archivo Excel")
        elif not os.path.exists(self.excel_path_var.get()):
            issues.append("• Excel no existe")
            
        if not os.path.exists(self.path_var.get()):
            issues.append("• Directorio boletas no existe")
            
        if issues:
            messagebox.showwarning("⚠️ Problemas", "\n".join(issues))
        else:
            messagebox.showinfo("✅ OK", "Configuración correcta")

    def clear_data(self):
        """Limpiar datos"""
        if self.is_processing:
            return
        self.excel_path_var.set("")
        self.status_var.set("")
        self.progress_var.set("")
        self.update_stats(0, 0, 0, "0s")
        self.update_status("🔄 Limpiado", "info")

    def select_excel(self):
        if self.is_processing:
            return
        file_path = filedialog.askopenfilename(
            title="Seleccionar Excel",
            filetypes=[("Excel", "*.xlsx"), ("Todos", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            filename = os.path.basename(file_path)
            self.update_status(f"✅ {filename}", "success")

    def start_email_process(self):
        if self.is_processing:
            return
        
        if not self.excel_path_var.get():
            messagebox.showwarning("Excel requerido", "Selecciona un archivo Excel")
            return
        
        self.set_processing_state(True)
        thread = threading.Thread(target=self.send_emails_thread)
        thread.daemon = True
        thread.start()

    def set_processing_state(self, processing):
        self.is_processing = processing
        if processing:
            self.send_btn.configure(
                state='disabled',
                text="⏳ PROCESANDO...",
                bootstyle="warning"
            )
            self.select_excel_btn.configure(state='disabled')
            self.month_combo.configure(state='disabled')
            self.progress_bar.start(10)
            self.status_var.set("")
            self.progress_var.set("🔄 Iniciando...")
        else:
            self.send_btn.configure(
                state='normal',
                text="📨 ENVIAR BOLETAS",
                bootstyle="success"
            )
            self.select_excel_btn.configure(state='normal')
            self.month_combo.configure(state='readonly')
            self.progress_bar.stop()
            self.progress_var.set("")

    def update_progress(self, message):
        # Acortar mensajes de progreso
        short_message = message.replace("Enviando correo a ", "📧 ").replace("Procesando archivo Excel...", "📖 Excel...")
        if len(short_message) > 40:
            short_message = short_message[:37] + "..."
        self.root.after(0, lambda: self.progress_var.set(short_message))

    def update_status(self, message, style="secondary"):
        def update():
            self.status_var.set(message)
            self.status_label.configure(bootstyle=style)
        self.root.after(0, update)

    def send_emails_thread(self):
        try:
            self.send_emails()
        except Exception as e:
            self.update_status(f"❌ Error: {str(e)}", "danger")
        finally:
            self.root.after(0, lambda: self.set_processing_state(False))

    def send_emails(self):
        # Refresca la plantilla de email antes de enviar
        try:
            from core.config import load_email_templates
            subject, body = load_email_templates()
            self.sender.subject_template = subject
            self.sender.body_template = body
        except Exception as e:
            self.update_status(f"❌ Error plantilla: {str(e)}", "danger")
            return

        excel_path = self.excel_path_var.get()
        if not excel_path or not os.path.exists(excel_path):
            self.update_status("❌ Excel inválido", "danger")
            return
            
        self.update_progress("📖 Leyendo Excel...")
        try:
            df = pd.read_excel(excel_path, dtype={"dni": str})
        except Exception as e:
            self.update_status(f"❌ Error Excel: {str(e)}", "danger")
            return
            
        required_columns = {"nombre", "email", "dni"}
        df.columns = df.columns.str.lower()
        if not required_columns.issubset(df.columns):
            self.update_status("❌ Faltan columnas: nombre, email, dni", "danger")
            return
            
        df = df[["nombre", "email", "dni"]].dropna(subset=["nombre", "email", "dni"])
        recipients = df.to_dict("records")
        
        if not recipients:
            self.update_status("❌ Sin registros válidos", "danger")
            return
            
        mes = self.mes_var.get()
        path_boletas = self.path_var.get()
        
        import time
        start_time = time.time()

        successful_sends, errores = self.sender.send_batch(
            recipients, mes, path_boletas, progress_callback=self.update_progress
        )

        # Generar constancias
        constancias_generadas = []
        if successful_sends:
            self.update_progress(f"📨 Generando {len(successful_sends)} constancias...")
            for item in successful_sends:
                try:
                    fecha_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    constancia_path = self.sender.generar_constancia_envio(
                        remitente=("Clínica Santa Rosa", EMAIL_USER),
                        destinatario=(item['nombre'], item['email']),
                        asunto=item['asunto'],
                        cuerpo=item['cuerpo_html'],
                        fecha_envio=fecha_envio,
                        adjunto_path=item['pdf_path'],
                        message_id=item['message_id']
                    )
                    constancias_generadas.append(constancia_path)
                except Exception as e:
                    self.update_status(f"⚠️ Error constancia {item['email']}: {str(e)}", "warning")
                    errores.append(('-', item['nombre'], item['email'], item['dni'], f"Error constancia: {str(e)}"))

        elapsed = int(time.time() - start_time)
        elapsed_str = f"{elapsed}s" if elapsed < 60 else f"{elapsed//60}m {elapsed%60}s"

        error_file_saved = self.sender.generar_log_errores(errores) if errores else None
        total_procesados = len(recipients)
        total_enviados = len(successful_sends)
        total_errores = len(errores)

        # Actualizar estadísticas
        self.update_stats(total_procesados, total_enviados, total_errores, elapsed_str)

        # Mensaje final compacto
        mensaje_lineas = [f"📊 {total_procesados} procesados"]
        if total_enviados > 0:
            mensaje_lineas.append(f"✅ {total_enviados} enviados")
        if total_errores > 0:
            mensaje_lineas.append(f"❌ {total_errores} errores")
        if error_file_saved:
            mensaje_lineas.append(f"📄 Log: {os.path.basename(error_file_saved)}")

        mensaje_final = "\n".join(mensaje_lineas)

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

        if constancias_generadas:
            carpeta_constancias = os.path.dirname(constancias_generadas[0])
            messagebox.showinfo(
                "✅ Completado",
                f"{len(constancias_generadas)} constancias PDF en:\n{carpeta_constancias}"
            )

        if error_file_saved:
            try:
                webbrowser.open(os.path.realpath(error_file_saved))
            except Exception:
                pass