import ttkbootstrap as tb
from ttkbootstrap.constants import *
from tkinter import messagebox, StringVar, Text, BooleanVar
from core.template_manager import TemplateManager
import threading

class TemplateManagerGUI:
    """Interfaz gráfica para gestionar plantillas de email"""
    
    def __init__(self, parent_window):
        self.parent = parent_window
        self.template_manager = TemplateManager()
        self.selected_template_id = None
        
        # Crear ventana modal
        self.window = tb.Toplevel(parent_window)
        self.window.title("📝 Gestor de Plantillas de Email")
        self.window.geometry("900x650")
        self.window.resizable(True, True)
        self.window.transient(parent_window)
        self.window.grab_set()
        
        # Variables
        self.template_var = StringVar()
        self.name_var = StringVar()
        self.subject_var = StringVar()
        self.is_default_var = BooleanVar()
        
        self.build_gui()
        self.refresh_template_list()
        
        # Centrar ventana
        self.center_window()
    
    def center_window(self):
        """Centrar ventana en la pantalla"""
        self.window.update_idletasks()
        x = (self.window.winfo_screenwidth() // 2) - (self.window.winfo_width() // 2)
        y = (self.window.winfo_screenheight() // 2) - (self.window.winfo_height() // 2)
        self.window.geometry(f"+{x}+{y}")
    
    def build_gui(self):
        """Construir interfaz gráfica"""
        # Frame principal
        main_frame = tb.Frame(self.window, padding=10)
        main_frame.pack(fill="both", expand=True)
        
        # Header
        header_frame = tb.Frame(main_frame, bootstyle="info", padding=10)
        header_frame.pack(fill="x", pady=(0, 10))
        
        title_label = tb.Label(
            header_frame,
            text="📝 Gestor de Plantillas de Email",
            font=("Segoe UI", 16, "bold"),
            bootstyle="inverse-info"
        )
        title_label.pack()
        
        # Layout principal - dos columnas
        content_frame = tb.Frame(main_frame)
        content_frame.pack(fill="both", expand=True)
        
        # Columna izquierda - Lista de plantillas (30%)
        left_frame = tb.Frame(content_frame, padding=5)
        left_frame.pack(side="left", fill="both", expand=False, padx=(0, 5))
        
        # Columna derecha - Editor (70%)
        right_frame = tb.Frame(content_frame, padding=5)
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))
        
        self.create_template_list(left_frame)
        self.create_template_editor(right_frame)
        self.create_bottom_buttons(main_frame)
        
        # Variables disponibles panel
        self.create_variables_panel(right_frame)
    
    def create_template_list(self, parent):
        """Crear lista de plantillas"""
        list_frame = tb.LabelFrame(parent, text="📋 Plantillas Disponibles", bootstyle="primary", padding=10)
        list_frame.pack(fill="both", expand=True)
        
        # Botones de acción de lista
        buttons_frame = tb.Frame(list_frame)
        buttons_frame.pack(fill="x", pady=(0, 10))
        
        self.new_btn = tb.Button(
            buttons_frame,
            text="➕ Nueva",
            command=self.new_template,
            bootstyle="success-outline",
            width=12
        )
        self.new_btn.pack(side="left", padx=(0, 5))
        
        self.duplicate_btn = tb.Button(
            buttons_frame,
            text="📋 Duplicar",
            command=self.duplicate_template,
            bootstyle="info-outline",
            width=12,
            state="disabled"
        )
        self.duplicate_btn.pack(side="left", padx=(0, 5))
        
        self.delete_btn = tb.Button(
            buttons_frame,
            text="🗑️ Eliminar",
            command=self.delete_template,
            bootstyle="danger-outline",
            width=12,
            state="disabled"
        )
        self.delete_btn.pack(side="left")
        
        # Lista de plantillas
        list_container = tb.Frame(list_frame)
        list_container.pack(fill="both", expand=True)
        
        # Scrollbar para la lista
        scrollbar = tb.Scrollbar(list_container)
        scrollbar.pack(side="right", fill="y")
        
        self.template_listbox = tb.Treeview(
            list_container,
            columns=("default",),
            show="tree headings",
            yscrollcommand=scrollbar.set,
            height=15
        )
        self.template_listbox.heading("#0", text="Plantilla")
        self.template_listbox.heading("default", text="Predeterminada")
        self.template_listbox.column("#0", width=200)
        self.template_listbox.column("default", width=80)
        
        self.template_listbox.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.template_listbox.yview)
        
        # Bind eventos
        self.template_listbox.bind("<<TreeviewSelect>>", self.on_template_select)
        self.template_listbox.bind("<Double-1>", self.on_template_double_click)
    
    def create_template_editor(self, parent):
        """Crear editor de plantillas"""
        editor_frame = tb.LabelFrame(parent, text="✏️ Editor de Plantilla", bootstyle="warning", padding=10)
        editor_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        # Nombre de plantilla
        name_frame = tb.Frame(editor_frame)
        name_frame.pack(fill="x", pady=(0, 10))
        
        tb.Label(name_frame, text="Nombre:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.name_entry = tb.Entry(
            name_frame,
            textvariable=self.name_var,
            font=("Segoe UI", 10),
            state="disabled"
        )
        self.name_entry.pack(fill="x", pady=(2, 0))
        
        # Asunto
        subject_frame = tb.Frame(editor_frame)
        subject_frame.pack(fill="x", pady=(0, 10))
        
        tb.Label(subject_frame, text="Asunto:", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        self.subject_entry = tb.Entry(
            subject_frame,
            textvariable=self.subject_var,
            font=("Segoe UI", 10),
            state="disabled"
        )
        self.subject_entry.pack(fill="x", pady=(2, 0))
        
        # Cuerpo del mensaje
        body_frame = tb.Frame(editor_frame)
        body_frame.pack(fill="both", expand=True, pady=(0, 10))
        
        tb.Label(body_frame, text="Cuerpo del mensaje (HTML):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
        
        # Frame para texto con scrollbar
        text_container = tb.Frame(body_frame)
        text_container.pack(fill="both", expand=True, pady=(2, 0))
        
        self.body_text = Text(
            text_container,
            wrap="word",
            font=("Consolas", 9),
            state="disabled",
            height=12
        )
        text_scrollbar = tb.Scrollbar(text_container, command=self.body_text.yview)
        self.body_text.config(yscrollcommand=text_scrollbar.set)
        
        self.body_text.pack(side="left", fill="both", expand=True)
        text_scrollbar.pack(side="right", fill="y")
        
        # Checkbox predeterminada
        default_frame = tb.Frame(editor_frame)
        default_frame.pack(fill="x", pady=(0, 10))
        
        self.default_check = tb.Checkbutton(
            default_frame,
            text="Usar como plantilla predeterminada",
            variable=self.is_default_var,
            bootstyle="success-round-toggle",
            state="disabled"
        )
        self.default_check.pack(anchor="w")
        
        # Botones de editor
        editor_buttons_frame = tb.Frame(editor_frame)
        editor_buttons_frame.pack(fill="x")
        
        self.save_btn = tb.Button(
            editor_buttons_frame,
            text="💾 Guardar",
            command=self.save_template,
            bootstyle="success",
            state="disabled"
        )
        self.save_btn.pack(side="left", padx=(0, 5))
        
        self.cancel_btn = tb.Button(
            editor_buttons_frame,
            text="❌ Cancelar",
            command=self.cancel_edit,
            bootstyle="secondary",
            state="disabled"
        )
        self.cancel_btn.pack(side="left", padx=(0, 5))
        
        self.preview_btn = tb.Button(
            editor_buttons_frame,
            text="👁️ Vista Previa",
            command=self.preview_template,
            bootstyle="info-outline",
            state="disabled"
        )
        self.preview_btn.pack(side="right")
    
    def create_variables_panel(self, parent):
        """Crear panel de variables disponibles"""
        variables_frame = tb.LabelFrame(parent, text="📌 Variables Disponibles", bootstyle="secondary", padding=8)
        variables_frame.pack(fill="x")
        
        variables = self.template_manager.get_available_variables()
        
        for i, var in enumerate(variables):
            var_frame = tb.Frame(variables_frame)
            var_frame.pack(fill="x", pady=2)
            
            var_label = tb.Label(
                var_frame,
                text=var["variable"],
                font=("Consolas", 9, "bold"),
                bootstyle="info"
            )
            var_label.pack(side="left")
            
            desc_label = tb.Label(
                var_frame,
                text=f" - {var['description']}",
                font=("Segoe UI", 8)
            )
            desc_label.pack(side="left")
    
    def create_bottom_buttons(self, parent):
        """Crear botones del fondo"""
        bottom_frame = tb.Frame(parent, padding=(0, 10, 0, 0))
        bottom_frame.pack(fill="x", side="bottom")
        
        close_btn = tb.Button(
            bottom_frame,
            text="🚪 Cerrar",
            command=self.close_window,
            bootstyle="secondary"
        )
        close_btn.pack(side="right")
        
        refresh_btn = tb.Button(
            bottom_frame,
            text="🔄 Actualizar",
            command=self.refresh_template_list,
            bootstyle="info-outline"
        )
        refresh_btn.pack(side="right", padx=(0, 10))
    
    def refresh_template_list(self):
        """Actualizar lista de plantillas"""
        # Limpiar lista actual
        for item in self.template_listbox.get_children():
            self.template_listbox.delete(item)
        
        # Cargar plantillas
        templates = self.template_manager.get_all_templates()
        templates.sort(key=lambda x: (not x.get('is_default', False), x.get('name', '')))
        
        for template in templates:
            default_text = "✓" if template.get('is_default', False) else ""
            self.template_listbox.insert(
                "",
                "end",
                iid=template['id'],
                text=template['name'],
                values=(default_text,)
            )
    
    def on_template_select(self, event):
        """Manejar selección de plantilla"""
        selection = self.template_listbox.selection()
        if not selection:
            self.selected_template_id = None
            self.disable_editor()
            return
        
        template_id = selection[0]
        self.selected_template_id = template_id
        
        # Habilitar botones de lista
        self.duplicate_btn.config(state="normal")
        self.delete_btn.config(state="normal")
        
        # Mostrar plantilla en editor (solo lectura)
        self.load_template_to_editor(template_id)
    
    def on_template_double_click(self, event):
        """Manejar doble click en plantilla"""
        if self.selected_template_id:
            self.edit_template()
    
    def load_template_to_editor(self, template_id):
        """Cargar plantilla en el editor"""
        template = self.template_manager.get_template(template_id)
        if not template:
            return
        
        self.name_var.set(template['name'])
        self.subject_var.set(template['subject'])
        self.is_default_var.set(template.get('is_default', False))
        
        self.body_text.config(state="normal")
        self.body_text.delete(1.0, "end")
        self.body_text.insert(1.0, template['body'])
        self.body_text.config(state="disabled")
    
    def disable_editor(self):
        """Deshabilitar editor"""
        self.name_var.set("")
        self.subject_var.set("")
        self.is_default_var.set(False)
        
        self.body_text.config(state="normal")
        self.body_text.delete(1.0, "end")
        self.body_text.config(state="disabled")
        
        # Deshabilitar botones
        self.duplicate_btn.config(state="disabled")
        self.delete_btn.config(state="disabled")
        self.save_btn.config(state="disabled")
        self.cancel_btn.config(state="disabled")
        self.preview_btn.config(state="disabled")
    
    def enable_editor(self):
        """Habilitar editor para edición"""
        self.name_entry.config(state="normal")
        self.subject_entry.config(state="normal")
        self.body_text.config(state="normal")
        self.default_check.config(state="normal")
        
        self.save_btn.config(state="normal")
        self.cancel_btn.config(state="normal")
        self.preview_btn.config(state="normal")
        
        # Deshabilitar botones de lista mientras se edita (no el Treeview)
        self.new_btn.config(state="disabled")
        self.duplicate_btn.config(state="disabled")
        self.delete_btn.config(state="disabled")
    
    def new_template(self):
        """Crear nueva plantilla"""
        self.selected_template_id = None
        
        # Limpiar campos
        self.name_var.set("Nueva Plantilla")
        self.subject_var.set("Boleta del mes de {MES}")
        self.is_default_var.set(False)
        
        self.body_text.config(state="normal")
        self.body_text.delete(1.0, "end")
        self.body_text.insert(1.0, "<html><body><h2>Hola {NOMBRE},</h2><p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p><p>Saludos cordiales,<br>Clínica Santa Rosa</p></body></html>")
        
        self.enable_editor()
        self.name_entry.focus()
        self.name_entry.select_range(0, "end")
    
    def edit_template(self):
        """Editar plantilla seleccionada"""
        if not self.selected_template_id:
            return
        
        self.enable_editor()
        self.name_entry.focus()
    
    def save_template(self):
        """Guardar plantilla"""
        name = self.name_var.get().strip()
        subject = self.subject_var.get().strip()
        body = self.body_text.get(1.0, "end-1c").strip()
        is_default = self.is_default_var.get()
        
        if not name:
            messagebox.showerror("Error", "El nombre de la plantilla es requerido")
            self.name_entry.focus()
            return
        
        if not subject:
            messagebox.showerror("Error", "El asunto es requerido")
            self.subject_entry.focus()
            return
        
        if not body:
            messagebox.showerror("Error", "El cuerpo del mensaje es requerido")
            self.body_text.focus()
            return
        
        # Validar variables
        missing_vars = self.template_manager.validate_template_variables(subject, body)
        if missing_vars:
            result = messagebox.askywarning(
                "Variables faltantes",
                f"Las siguientes variables requeridas no están presentes:\n{', '.join(missing_vars)}\n\n¿Desea continuar de todos modos?",
                type=messagebox.YESNO
            )
            if result != messagebox.YES:
                return
        
        # Guardar plantilla
        if self.selected_template_id:
            # Actualizar existente
            success, message = self.template_manager.update_template(
                self.selected_template_id,
                name=name,
                subject=subject,
                body=body,
                set_as_default=is_default
            )
        else:
            # Crear nueva
            success, message = self.template_manager.create_template(
                name=name,
                subject=subject,
                body=body,
                set_as_default=is_default
            )
            if success:
                self.selected_template_id = message  # message contiene el ID
        
        if success:
            messagebox.showinfo("Éxito", "Plantilla guardada exitosamente")
            self.refresh_template_list()
            self.disable_editor_after_save()
            if self.selected_template_id:
                self.template_listbox.selection_set(self.selected_template_id)
                self.template_listbox.focus(self.selected_template_id)
        else:
            messagebox.showerror("Error", f"Error guardando plantilla: {message}")
    
    def cancel_edit(self):
        """Cancelar edición"""
        self.disable_editor_after_save()
        if self.selected_template_id:
            self.load_template_to_editor(self.selected_template_id)
        else:
            self.disable_editor()
    
    def disable_editor_after_save(self):
        """Deshabilitar editor después de guardar"""
        self.name_entry.config(state="disabled")
        self.subject_entry.config(state="disabled")
        self.body_text.config(state="disabled")
        self.default_check.config(state="disabled")
        
        self.save_btn.config(state="disabled")
        self.cancel_btn.config(state="disabled")
        self.preview_btn.config(state="disabled")
        
        # Re-habilitar botones de lista
        self.new_btn.config(state="normal")
        if self.selected_template_id:
            self.duplicate_btn.config(state="normal")
            self.delete_btn.config(state="normal")
    
    def duplicate_template(self):
        """Duplicar plantilla seleccionada"""
        if not self.selected_template_id:
            return
        
        template = self.template_manager.get_template(self.selected_template_id)
        if not template:
            return
        
        new_name = f"Copia de {template['name']}"
        success, message = self.template_manager.duplicate_template(
            self.selected_template_id,
            new_name
        )
        
        if success:
            messagebox.showinfo("Éxito", "Plantilla duplicada exitosamente")
            self.refresh_template_list()
            # Seleccionar la nueva plantilla
            new_id = message
            self.template_listbox.selection_set(new_id)
            self.template_listbox.focus(new_id)
            self.selected_template_id = new_id
            self.load_template_to_editor(new_id)
        else:
            messagebox.showerror("Error", f"Error duplicando plantilla: {message}")
    
    def delete_template(self):
        """Eliminar plantilla seleccionada"""
        if not self.selected_template_id:
            return
        
        template = self.template_manager.get_template(self.selected_template_id)
        if not template:
            return
        
        result = messagebox.askyesno(
            "Confirmar eliminación",
            f"¿Está seguro de que desea eliminar la plantilla '{template['name']}'?\n\nEsta acción no se puede deshacer."
        )
        
        if result:
            success, message = self.template_manager.delete_template(self.selected_template_id)
            
            if success:
                messagebox.showinfo("Éxito", "Plantilla eliminada exitosamente")
                self.refresh_template_list()
                self.selected_template_id = None
                self.disable_editor()
            else:
                messagebox.showerror("Error", f"Error eliminando plantilla: {message}")
    
    def preview_template(self):
        """Vista previa de la plantilla"""
        subject = self.subject_var.get()
        body = self.body_text.get(1.0, "end-1c")
        
        # Reemplazar variables con datos de ejemplo
        preview_subject = subject.replace("{NOMBRE}", "Juan Pérez").replace("{MES}", "Enero")
        preview_body = body.replace("{NOMBRE}", "Juan Pérez").replace("{MES}", "Enero")
        
        # Crear ventana de vista previa
        preview_window = tb.Toplevel(self.window)
        preview_window.title("👁️ Vista Previa de Plantilla")
        preview_window.geometry("600x500")
        preview_window.transient(self.window)
        preview_window.grab_set()
        
        # Contenido de vista previa
        main_frame = tb.Frame(preview_window, padding=15)
        main_frame.pack(fill="both", expand=True)
        
        # Asunto
        tb.Label(main_frame, text="Asunto:", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        subject_frame = tb.Frame(main_frame, bootstyle="info", padding=10)
        subject_frame.pack(fill="x", pady=(5, 15))
        
        tb.Label(subject_frame, text=preview_subject, font=("Segoe UI", 11)).pack(anchor="w")
        
        # Cuerpo
        tb.Label(main_frame, text="Cuerpo del mensaje:", font=("Segoe UI", 12, "bold")).pack(anchor="w")
        body_frame = tb.Frame(main_frame, bootstyle="light", padding=10)
        body_frame.pack(fill="both", expand=True, pady=(5, 15))
        
        # Mostrar HTML renderizado (simplificado)
        import html
        preview_text = Text(body_frame, wrap="word", font=("Segoe UI", 10), state="normal")
        preview_text.pack(fill="both", expand=True)
        
        # Insertar contenido (simplificar HTML para mostrar)
        display_text = preview_body
        display_text = display_text.replace("<html>", "").replace("</html>", "")
        display_text = display_text.replace("<body>", "").replace("</body>", "")
        display_text = display_text.replace("<h2>", "").replace("</h2>", "\n")
        display_text = display_text.replace("<p>", "").replace("</p>", "\n\n")
        display_text = display_text.replace("<br>", "\n").replace("<br/>", "\n")
        display_text = html.unescape(display_text.strip())
        
        preview_text.insert(1.0, display_text)
        preview_text.config(state="disabled")
        
        # Botón cerrar
        close_btn = tb.Button(
            main_frame,
            text="Cerrar",
            command=preview_window.destroy,
            bootstyle="secondary"
        )
        close_btn.pack(pady=(0, 5))
    
    def close_window(self):
        """Cerrar ventana"""
        self.window.destroy()

def open_template_manager(parent_window):
    """Función para abrir el gestor de plantillas"""
    TemplateManagerGUI(parent_window)