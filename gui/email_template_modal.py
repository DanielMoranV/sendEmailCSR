import os
from tkinter import Toplevel, Entry, Text, END, BOTH, LEFT, RIGHT, Y, X, Scrollbar, VERTICAL, INSERT, messagebox
import ttkbootstrap as tb
from tkhtmlview import HTMLLabel
from core.config import EMAIL_CONFIG_FILE, load_email_templates

def open_template_editor_modal(root):
    modal = Toplevel(root)
    modal.title("Editar plantilla de email")
    modal.geometry("800x650")
    modal.transient(root)
    modal.grab_set()
    modal.resizable(True, True)

    # Cargar plantilla actual
    subject, body = load_email_templates()

    # --- Ayuda visual ---
    help_frame = tb.Frame(modal, padding=8)
    help_frame.pack(fill=X)
    help_text = (
        "Variables disponibles: "
        "{NOMBRE} (nombre del destinatario), "
        "{MES} (mes seleccionado).\n"
        "Puedes usar HTML: <b>negrita</b>, <u>subrayado</u>, <i>cursiva</i>, <br>, etc.\n"
        "Ejemplo: <b>Hola {NOMBRE}</b> - Boleta de {MES}."
    )
    tb.Label(help_frame, text=help_text, font=("Segoe UI", 9), justify="left", bootstyle="info").pack(anchor="w")

    # --- Contenedor principal de dos columnas ---
    columns_frame = tb.Frame(modal, padding=5)
    columns_frame.pack(fill=BOTH, expand=True)
    columns_frame.columnconfigure(0, weight=2, minsize=350)
    columns_frame.columnconfigure(1, weight=3, minsize=350)
    columns_frame.rowconfigure(0, weight=1)

    # --- Columna izquierda: Edición ---
    edit_frame = tb.Frame(columns_frame)
    edit_frame.grid(row=0, column=0, sticky="nsew", padx=(0, 10))

    # Campo asunto
    tb.Label(edit_frame, text="Asunto:", font=("Segoe UI", 10, "bold")).pack(anchor="w", pady=(0,2))
    subject_entry = Entry(edit_frame, font=("Segoe UI", 10))
    subject_entry.pack(fill=X, pady=(0,8))
    subject_entry.insert(0, subject)

    # Campo cuerpo
    tb.Label(edit_frame, text="Cuerpo (HTML):", font=("Segoe UI", 10, "bold")).pack(anchor="w")
    text_body = Text(edit_frame, wrap="word", font=("Segoe UI", 10), height=13)
    text_body.pack(fill=BOTH, expand=True, pady=(0,2))
    text_body.insert(1.0, body)
    # Scrollbar
    sb = Scrollbar(edit_frame, orient=VERTICAL, command=text_body.yview)
    sb.pack(side=RIGHT, fill=Y)
    text_body.config(yscrollcommand=sb.set)

    # Botones de formato rápido
    def insert_tag(tag_open, tag_close):
        try:
            start = text_body.index("sel.first")
            end = text_body.index("sel.last")
            selected = text_body.get(start, end)
            text_body.delete(start, end)
            text_body.insert(start, f"{tag_open}{selected}{tag_close}")
        except:
            text_body.insert(INSERT, f"{tag_open}{tag_close}")

    btns_frame = tb.Frame(edit_frame, padding=3)
    btns_frame.pack(fill=X, pady=(3,0))
    tb.Button(btns_frame, text="<b>Negrita</b>", command=lambda: insert_tag("<b>", "</b>"), width=12).pack(side=LEFT, padx=2)
    tb.Button(btns_frame, text="<u>Subrayado</u>", command=lambda: insert_tag("<u>", "</u>"), width=12).pack(side=LEFT, padx=2)
    tb.Button(btns_frame, text="<i>Cursiva</i>", command=lambda: insert_tag("<i>", "</i>"), width=12).pack(side=LEFT, padx=2)
    tb.Button(btns_frame, text="<br> Salto línea", command=lambda: insert_tag("<br>", ""), width=14).pack(side=LEFT, padx=2)

    # --- Columna derecha: Vista previa ---
    preview_frame = tb.LabelFrame(columns_frame, text="Vista previa", bootstyle="primary", padding=8)
    preview_frame.grid(row=0, column=1, sticky="nsew")
    preview_label = HTMLLabel(preview_frame, html="", background="#f8f9fa")
    preview_label.pack(fill=BOTH, expand=True)

    def update_preview(*_):
        # Simula reemplazo de variables para preview
        subj = subject_entry.get().replace("{MES}", "Junio").replace("{NOMBRE}", "Juan Pérez")
        html = text_body.get(1.0, END)
        html = html.replace("{MES}", "Junio").replace("{NOMBRE}", "Juan Pérez")
        preview_label.set_html(f"<b>Asunto:</b> {subj}<hr>" + html)
    subject_entry.bind("<KeyRelease>", update_preview)
    text_body.bind("<KeyRelease>", update_preview)
    update_preview()

    # --- Botón guardar ---
    def save_template():
        new_subject = subject_entry.get()
        new_body = text_body.get(1.0, END).strip()
        try:
            # Guardar en archivo
            lines = []
            lines.append(f"SUBJECT={new_subject}\n")
            lines.append(f"BODY={new_body}\n")
            with open(EMAIL_CONFIG_FILE, "w", encoding="utf-8") as f:
                f.writelines(lines)
            messagebox.showinfo("Plantilla guardada", "La plantilla de email se guardó correctamente.")
            modal.destroy()
        except Exception as e:
            messagebox.showerror("Error", f"No se pudo guardar la plantilla: {str(e)}")

    btn_save = tb.Button(modal, text="💾 Guardar plantilla", bootstyle="success", command=save_template, width=24)
    btn_save.pack(pady=10)

    # Botón cerrar
    btn_close = tb.Button(modal, text="Cerrar", bootstyle="secondary-outline", command=modal.destroy, width=12)
    btn_close.pack()
