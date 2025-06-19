import os
import ttkbootstrap as tb
import webbrowser # Import webbrowser
from ttkbootstrap.constants import *
from tkinter import filedialog, StringVar, messagebox
from gui.email_template_modal import open_template_editor_modal
import threading
import pandas as pd
from core.email_sender import EmailSender
from core.config import DEFAULT_PATH, EMAIL_USER # Import EMAIL_USER
from datetime import datetime # Import datetime

class EmailSenderGUI:
    # El método ahora solo llama a la función modularizada
    def open_template_editor_modal(self):
        open_template_editor_modal(self.root)

    def __init__(self, root):
        self.root = root
        self.root.title("📧 Sistema de Envío de Boletas - Clínica Santa Rosa")
        self.root.geometry("1000x900")  # Aumentamos el ancho para dos columnas
        self.root.resizable(True, True)
        
        # Configurar el tema
        self.style = tb.Style("flatly")
        
        # Configurar colores personalizados
        self.colors = {
            'primary': '#2E86AB',
            'secondary': '#A23B72',
            'success': '#4CAF50',
            'warning': '#FF9800',
            'danger': '#F44336',
            'info': '#2196F3',
            'light': '#F8F9FA',
            'dark': '#343A40'
        }

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
        # Frame principal con fondo degradado simulado
        main_frame = tb.Frame(self.root, bootstyle="light")
        main_frame.pack(fill="both", expand=True)

        # Header con título y logo
        self.create_header(main_frame)
        
        # Contenedor principal con padding
        container = tb.Frame(main_frame, padding=20, bootstyle="light")
        container.pack(fill="both", expand=True)
        
        # Crear layout de dos columnas
        self.create_two_column_layout(container)
        
        # Footer con información (span completo)
        self.create_footer(main_frame)

    def create_two_column_layout(self, parent):
        """Crear el layout de dos columnas"""
        # Frame principal para las dos columnas
        columns_frame = tb.Frame(parent)
        columns_frame.pack(fill="both", expand=True)
        
        # Columna izquierda - Configuración
        left_column = tb.Frame(columns_frame, padding=10)
        left_column.pack(side="left", fill="both", expand=True, padx=(0, 10))
        
        # Columna derecha - Acciones y Estado
        right_column = tb.Frame(columns_frame, padding=10)
        right_column.pack(side="right", fill="both", expand=True, padx=(10, 0))
        
        # Llenar columna izquierda
        self.create_left_column_content(left_column)
        
        # Llenar columna derecha
        self.create_right_column_content(right_column)

    def create_left_column_content(self, parent):
        """Crear contenido de la columna izquierda - Configuración"""
        # Título de la columna
        title_label = tb.Label(
            parent,
            text="⚙️ CONFIGURACIÓN",
            font=("Segoe UI", 16, "bold"),
            bootstyle="primary"
        )
        title_label.pack(pady=(0, 20))
        
        # Sección de selección de mes
        self.create_month_section(parent)
        
        # Separador
        self.create_separator(parent)
        
        # Sección de rutas
        self.create_paths_section(parent)
        
        # Información adicional
        self.create_info_section(parent)

        # Botón para editar plantilla de email
        edit_template_btn = tb.Button(
            parent,
            text="✏️ Editar plantilla de email",
            command=self.open_template_editor_modal,
            bootstyle="info-outline",
            width=35
        )
        edit_template_btn.pack(pady=(15, 0))

    def create_right_column_content(self, parent):
        """Crear contenido de la columna derecha - Acciones y Estado"""
        # Título de la columna
        title_label = tb.Label(
            parent,
            text="🚀 ACCIONES Y ESTADO",
            font=("Segoe UI", 16, "bold"),
            bootstyle="success"
        )
        title_label.pack(pady=(0, 20))
        
        # Sección de acciones
        self.create_actions_section(parent)
        
        # Separador
        self.create_separator(parent)
        
        # Sección de progreso
        self.create_progress_section(parent)
        
        # Estadísticas rápidas
        self.create_stats_section(parent)

    def create_header(self, parent):
        """Crear header con título y decoraciones"""
        header_frame = tb.Frame(parent, bootstyle="info", padding=20)
        header_frame.pack(fill="x")
        
        # Contenedor para centrar el contenido
        header_content = tb.Frame(header_frame, bootstyle="info")
        header_content.pack(expand=True)
        
        # Título principal
        title_label = tb.Label(
            header_content, 
            text="📧 Sistema de Envío de Boletas",
            font=("Segoe UI", 22, "bold"),
            bootstyle="inverse-info"
        )
        title_label.pack()
        
        # Subtítulo
        subtitle_label = tb.Label(
            header_content,
            text="Clínica Santa Rosa - Gestión Automatizada de Envíos",
            font=("Segoe UI", 12),
            bootstyle="inverse-info"
        )
        subtitle_label.pack(pady=(5, 0))

    def create_month_section(self, parent):
        """Crear sección de selección de mes"""
        month_frame = tb.LabelFrame(
            parent,
            text="📅 Período de Envío",
            bootstyle="info",
            padding=20
        )
        month_frame.pack(fill="x", pady=(0, 15))
        
        # Label con icono
        month_label = tb.Label(
            month_frame,
            text="🗓️ Mes a procesar:",
            font=("Segoe UI", 11, "bold"),
            bootstyle="info"
        )
        month_label.pack(anchor="w", pady=(0, 8))
        
        # Combobox estilizado
        self.month_combo = tb.Combobox(
            month_frame,
            values=[m[1] for m in self.months],
            textvariable=self.mes_var,
            font=("Segoe UI", 11),
            bootstyle="info",
            state="readonly",
            width=25
        )
        self.month_combo.pack(anchor="w", fill="x")

    def create_paths_section(self, parent):
        """Crear sección de rutas y archivos"""
        paths_frame = tb.LabelFrame(
            parent,
            text="📁 Archivos y Rutas",
            bootstyle="warning",
            padding=20
        )
        paths_frame.pack(fill="x", pady=(0, 15))
        
        # Ruta de boletas
        boletas_label = tb.Label(
            paths_frame,
            text="📄 Directorio de boletas:",
            font=("Segoe UI", 11, "bold"),
            bootstyle="warning"
        )
        boletas_label.pack(anchor="w", pady=(0, 5))
        
        self.path_entry = tb.Entry(
            paths_frame,
            textvariable=self.path_var,
            font=("Segoe UI", 10),
            bootstyle="warning"
        )
        self.path_entry.pack(fill="x", pady=(0, 15))
        
        # Archivo Excel
        excel_label = tb.Label(
            paths_frame,
            text="📊 Base de datos (Excel):",
            font=("Segoe UI", 11, "bold"),
            bootstyle="warning"
        )
        excel_label.pack(anchor="w", pady=(0, 5))
        
        # Botón para seleccionar Excel
        self.select_excel_btn = tb.Button(
            paths_frame,
            text="📂 Seleccionar Archivo Excel",
            command=self.select_excel,
            bootstyle="warning-outline",
            width=30
        )
        self.select_excel_btn.pack(fill="x", pady=(0, 8))
        
        # Entry para mostrar archivo seleccionado
        self.excel_entry = tb.Entry(
            paths_frame,
            textvariable=self.excel_path_var,
            font=("Segoe UI", 9),
            bootstyle="warning",
            state="readonly"
        )
        self.excel_entry.pack(fill="x")

    def create_info_section(self, parent):
        """Crear sección de información adicional"""
        info_frame = tb.LabelFrame(
            parent,
            text="💡 Información",
            bootstyle="secondary",
            padding=15
        )
        info_frame.pack(fill="x", pady=(15, 0))
        
        info_text = """📋 Requisitos del archivo Excel:
• Columnas: nombre, email, dni
• Formato: .xlsx
• Sin filas vacías en datos principales

📁 Estructura de archivos:
• Boletas en formato: {dni}.pdf
• Organizadas por mes"""
        
        info_label = tb.Label(
            info_frame,
            text=info_text,
            font=("Segoe UI", 9),
            bootstyle="secondary",
            justify="left"
        )
        info_label.pack(anchor="w")

    def create_actions_section(self, parent):
        """Crear sección de acciones principales"""
        actions_frame = tb.LabelFrame(
            parent,
            text="🎯 Panel de Control",
            bootstyle="success",
            padding=25
        )
        actions_frame.pack(fill="x", pady=(0, 15))
        
        # Botón principal de envío
        self.send_btn = tb.Button(
            actions_frame,
            text="📨 INICIAR ENVÍO DE BOLETAS",
            command=self.start_email_process,
            bootstyle="success",
            width=35
        )
        self.send_btn.pack(pady=(0, 15))
        
        # Botones secundarios en una fila
        secondary_frame = tb.Frame(actions_frame)
        secondary_frame.pack(fill="x", pady=(0, 10))
        
        # Botón para verificar configuración
        verify_btn = tb.Button(
            secondary_frame,
            text="🔍 Verificar Config",
            command=self.verify_configuration,
            bootstyle="info-outline",
            width=17
        )
        verify_btn.pack(side="left", padx=(0, 5))
        
        # Botón para limpiar datos
        clear_btn = tb.Button(
            secondary_frame,
            text="🗑️ Limpiar",
            command=self.clear_data,
            bootstyle="secondary-outline",
            width=17
        )
        clear_btn.pack(side="right", padx=(5, 0))

    def create_progress_section(self, parent):
        """Crear sección de progreso y estado"""
        progress_frame = tb.LabelFrame(
            parent,
            text="📊 Estado del Proceso",
            bootstyle="primary",
            padding=20
        )
        progress_frame.pack(fill="x", pady=(0, 15))
        
        # Label de progreso
        self.progress_label = tb.Label(
            progress_frame,
            textvariable=self.progress_var,
            font=("Segoe UI", 10, "bold"),
            bootstyle="primary"
        )
        self.progress_label.pack(anchor="w", pady=(0, 10))
        
        # Barra de progreso
        self.progress_bar = tb.Progressbar(
            progress_frame,
            mode='indeterminate',
            bootstyle="success-striped",
            length=300
        )
        self.progress_bar.pack(fill="x", pady=(0, 15))
        
        # Label de estado
        self.status_label = tb.Label(
            progress_frame,
            textvariable=self.status_var,
            font=("Segoe UI", 10),
            bootstyle="primary",
            wraplength=350,
            justify="left"
        )
        self.status_label.pack(anchor="w")

    def create_stats_section(self, parent):
        """Crear sección de estadísticas rápidas con variables persistentes"""
        stats_frame = tb.LabelFrame(
            parent,
            text="📈 Estadísticas",
            bootstyle="secondary",
            padding=15
        )
        stats_frame.pack(fill="x", pady=(15, 0))

        stats_grid = tb.Frame(stats_frame)
        stats_grid.pack(fill="x")

        # Crear StringVars para cada estadística si no existen
        self.stat_vars = getattr(self, 'stat_vars', {})
        for key in ["procesados", "enviados", "errores", "tiempo"]:
            if key not in self.stat_vars:
                self.stat_vars[key] = StringVar(value="0" if key != "tiempo" else "0s")

        self.stat_labels = getattr(self, 'stat_labels', {})
        self.stat_labels["procesados"] = self.create_stat_item(stats_grid, "📧 Procesados:", self.stat_vars["procesados"], 0, 0)
        self.stat_labels["enviados"] = self.create_stat_item(stats_grid, "✅ Enviados:", self.stat_vars["enviados"], 0, 1)
        self.stat_labels["errores"] = self.create_stat_item(stats_grid, "❌ Errores:", self.stat_vars["errores"], 1, 0)
        self.stat_labels["tiempo"] = self.create_stat_item(stats_grid, "⏱️ Tiempo:", self.stat_vars["tiempo"], 1, 1)

    def create_stat_item(self, parent, label, var, row, col):
        """Crear un elemento de estadística usando StringVar"""
        frame = tb.Frame(parent)
        frame.grid(row=row, column=col, padx=5, pady=2, sticky="ew")
        parent.grid_columnconfigure(col, weight=1)
        label_widget = tb.Label(frame, text=label, font=("Segoe UI", 9))
        label_widget.pack(side="left")
        value_widget = tb.Label(frame, textvariable=var, font=("Segoe UI", 9, "bold"), bootstyle="secondary")
        value_widget.pack(side="right")
        return value_widget

    def update_stats(self, procesados, enviados, errores, tiempo):
        """Actualizar valores de estadísticas en la GUI"""
        self.stat_vars["procesados"].set(str(procesados))
        self.stat_vars["enviados"].set(str(enviados))
        self.stat_vars["errores"].set(str(errores))
        self.stat_vars["tiempo"].set(str(tiempo))



    def create_footer(self, parent):
        """Crear footer con información adicional"""
        footer_frame = tb.Frame(parent, bootstyle="dark", padding=15)
        footer_frame.pack(fill="x", side="bottom")
        
        footer_content = tb.Frame(footer_frame, bootstyle="dark")
        footer_content.pack(expand=True)
        
        footer_label = tb.Label(
            footer_content,
            text="🏥 Clínica Santa Rosa | Sistema de Gestión de Boletas v2.0 | Soporte: soporteti@csr.pe.com",
            font=("Segoe UI", 9),
            bootstyle="inverse-dark"
        )
        footer_label.pack()

    def create_separator(self, parent):
        """Crear separador visual"""
        separator = tb.Separator(parent, orient="horizontal")
        separator.pack(fill="x", pady=8)

    def verify_configuration(self):
        """Verificar la configuración actual"""
        issues = []
        
        if not self.excel_path_var.get():
            issues.append("• No se ha seleccionado archivo Excel")
        elif not os.path.exists(self.excel_path_var.get()):
            issues.append("• El archivo Excel no existe")
            
        if not os.path.exists(self.path_var.get()):
            issues.append("• El directorio de boletas no existe")
            
        if issues:
            messagebox.showwarning("⚠️ Problemas de Configuración", "\n".join(issues))
        else:
            messagebox.showinfo("✅ Configuración Correcta", "Todos los requisitos están listos para el envío.")

    def clear_data(self):
        """Limpiar datos de la interfaz"""
        if self.is_processing:
            return
        self.excel_path_var.set("")
        self.status_var.set("")
        self.progress_var.set("")
        # Reiniciar las estadísticas
        if hasattr(self, 'update_stats'):
            self.update_stats(0, 0, 0, "0s")
        self.update_status("🔄 Datos limpiados", "info")


    def select_excel(self):
        if self.is_processing:
            return
        file_path = filedialog.askopenfilename(
            title="Seleccionar archivo Excel",
            filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")]
        )
        if file_path:
            self.excel_path_var.set(file_path)
            filename = os.path.basename(file_path)
            self.update_status(f"✅ Archivo seleccionado: {filename}", "success")

    def start_email_process(self):
        if self.is_processing:
            return
        
        if not self.excel_path_var.get():
            messagebox.showwarning(
                "Archivo requerido",
                "Por favor selecciona un archivo Excel antes de continuar."
            )
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
            self.progress_var.set("🔄 Iniciando proceso de envío...")
        else:
            self.send_btn.configure(
                state='normal',
                text="📨 INICIAR ENVÍO DE BOLETAS",
                bootstyle="success"
            )
            self.select_excel_btn.configure(state='normal')
            self.month_combo.configure(state='readonly')
            self.progress_bar.stop()
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
        # Refresca la plantilla de email antes de enviar
        try:
            from core.config import load_email_templates
            subject, body = load_email_templates()
            self.sender.subject_template = subject
            self.sender.body_template = body
        except Exception as e:
            self.update_status(f"❌ Error al recargar plantilla: {str(e)}", "danger")
            return

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
        
        import time
        start_time = time.time()

        # Call send_batch directly
        successful_sends, errores = self.sender.send_batch(
            recipients, mes, path_boletas, progress_callback=self.update_progress
        )

        # Generate constancias for successful sends
        constancias_generadas = []
        if successful_sends:
            self.update_progress(f"📨 Generando constancias para {len(successful_sends)} envíos exitosos...")
            for item in successful_sends:
                try:
                    fecha_envio = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    constancia_path = self.sender.generar_constancia_envio(
                        remitente=("Clínica Santa Rosa", EMAIL_USER), # Use imported EMAIL_USER
                        destinatario=(item['nombre'], item['email']),
                        asunto=item['asunto'],
                        cuerpo=item['cuerpo_html'],
                        fecha_envio=fecha_envio,
                        adjunto_path=item['pdf_path'],
                        message_id=item['message_id']
                    )
                    constancias_generadas.append(constancia_path)
                except Exception as e:
                    self.update_status(f"⚠️ Error generando constancia para {item['email']}: {str(e)}", "warning")
                    # Log this error as well, perhaps to a separate log or add to 'errores' list with a specific marker
                    errores.append(('-', item['nombre'], item['email'], item['dni'], f"Error al generar constancia: {str(e)}"))

        elapsed = int(time.time() - start_time)
        elapsed_str = f"{elapsed}s" if elapsed < 60 else f"{elapsed//60}m {elapsed%60}s"

        error_file_saved = self.sender.generar_log_errores(errores) if errores else None
        total_procesados = len(recipients)
        total_enviados = len(successful_sends) # Use length of successful_sends
        total_errores = len(errores)

        # Actualizar estadísticas en la GUI
        self.update_stats(total_procesados, total_enviados, total_errores, elapsed_str)

        mensaje_lineas = []
        mensaje_lineas.append(f"📊 Total procesados: {total_procesados}")
        if total_enviados > 0:
            mensaje_lineas.append(f"✅ Enviados correctamente: {total_enviados}")
        if total_errores > 0:
            mensaje_lineas.append(f"❌ Errores encontrados: {total_errores}")
        if error_file_saved:
            mensaje_lineas.append(f"📄 Ver detalles en: {os.path.basename(error_file_saved)}")

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
                "✅ Proceso Completado",
                f"Se generaron {len(constancias_generadas)} constancias PDF en la carpeta:\n{carpeta_constancias}"
            )

        if error_file_saved:
            try:
                webbrowser.open(os.path.realpath(error_file_saved)) # Changed to webbrowser.open
            except Exception:
                pass