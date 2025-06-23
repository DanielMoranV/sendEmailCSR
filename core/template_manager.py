import os
import json
from datetime import datetime
from typing import Dict, List, Optional, Tuple

class TemplateManager:
    """Gestor de plantillas de email con funcionalidades CRUD"""
    
    def __init__(self, templates_file="email_templates.json"):
        self.templates_file = templates_file
        self.templates = self._load_templates()
        
    def _load_templates(self) -> Dict:
        """Cargar plantillas desde archivo JSON"""
        default_templates = {
            "default": {
                "id": "default",
                "name": "Plantilla Predeterminada",
                "subject": "Boleta del mes de {MES}",
                "body": "<html><body><h2>Hola {NOMBRE},</h2><p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p><p>Saludos cordiales,<br>Clínica Santa Rosa</p><p>Favor de confirmar la recepción de este correo.</p></body></html>",
                "created_at": datetime.now().isoformat(),
                "updated_at": datetime.now().isoformat(),
                "is_default": True
            }
        }
        
        if os.path.exists(self.templates_file):
            try:
                with open(self.templates_file, 'r', encoding='utf-8') as f:
                    templates = json.load(f)
                    # Asegurar que exista una plantilla por defecto
                    if not any(t.get('is_default', False) for t in templates.values()):
                        templates.update(default_templates)
                    return templates
            except (json.JSONDecodeError, Exception):
                # Si hay error, usar plantillas por defecto
                pass
        
        # Migrar desde email_config.txt si existe
        return self._migrate_from_old_config(default_templates)
    
    def _migrate_from_old_config(self, default_templates: Dict) -> Dict:
        """Migrar configuración antigua desde email_config.txt"""
        old_config_file = "email_config.txt"
        
        if os.path.exists(old_config_file):
            try:
                subject = default_templates["default"]["subject"]
                body = default_templates["default"]["body"]
                
                with open(old_config_file, 'r', encoding='utf-8') as f:
                    for line in f:
                        line = line.strip()
                        if line.startswith("SUBJECT="):
                            subject = line.split("=", 1)[1]
                        elif line.startswith("BODY="):
                            body = line.split("=", 1)[1]
                
                # Actualizar plantilla por defecto con datos migrados
                default_templates["default"]["subject"] = subject
                default_templates["default"]["body"] = body
                default_templates["default"]["name"] = "Plantilla Migrada (Predeterminada)"
                
                # Guardar las plantillas migradas
                self._save_templates(default_templates)
                
            except Exception as e:
                print(f"Error migrando desde {old_config_file}: {e}")
        
        return default_templates
    
    def _save_templates(self, templates: Dict = None) -> bool:
        """Guardar plantillas en archivo JSON"""
        try:
            templates_to_save = templates or self.templates
            with open(self.templates_file, 'w', encoding='utf-8') as f:
                json.dump(templates_to_save, f, indent=2, ensure_ascii=False)
            return True
        except Exception as e:
            print(f"Error guardando plantillas: {e}")
            return False
    
    def get_all_templates(self) -> List[Dict]:
        """Obtener todas las plantillas"""
        return list(self.templates.values())
    
    def get_template(self, template_id: str) -> Optional[Dict]:
        """Obtener una plantilla específica por ID"""
        return self.templates.get(template_id)
    
    def get_default_template(self) -> Dict:
        """Obtener la plantilla marcada como predeterminada"""
        for template in self.templates.values():
            if template.get('is_default', False):
                return template
        
        # Si no hay predeterminada, devolver la primera disponible
        if self.templates:
            return list(self.templates.values())[0]
        
        # Crear plantilla por defecto si no existe ninguna
        default = {
            "id": "default",
            "name": "Plantilla Predeterminada",
            "subject": "Boleta del mes de {MES}",
            "body": "<html><body><h2>Hola {NOMBRE},</h2><p>Adjuntamos tu boleta correspondiente al mes de {MES}.</p><p>Saludos cordiales,<br>Clínica Santa Rosa</p><p>Favor de confirmar la recepción de este correo.</p></body></html>",
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "is_default": True
        }
        self.templates["default"] = default
        self._save_templates()
        return default
    
    def create_template(self, name: str, subject: str, body: str, set_as_default: bool = False) -> Tuple[bool, str]:
        """Crear nueva plantilla"""
        if not name.strip():
            return False, "El nombre de la plantilla es requerido"
        
        if not subject.strip():
            return False, "El asunto es requerido"
        
        if not body.strip():
            return False, "El cuerpo del mensaje es requerido"
        
        # Generar ID único
        template_id = f"template_{int(datetime.now().timestamp())}"
        
        # Si se marca como predeterminada, desmarcar otras
        if set_as_default:
            for template in self.templates.values():
                template['is_default'] = False
        
        new_template = {
            "id": template_id,
            "name": name.strip(),
            "subject": subject.strip(),
            "body": body.strip(),
            "created_at": datetime.now().isoformat(),
            "updated_at": datetime.now().isoformat(),
            "is_default": set_as_default
        }
        
        self.templates[template_id] = new_template
        
        if self._save_templates():
            return True, template_id
        else:
            # Revertir cambios si no se pudo guardar
            del self.templates[template_id]
            return False, "Error guardando la plantilla"
    
    def update_template(self, template_id: str, name: str = None, subject: str = None, 
                       body: str = None, set_as_default: bool = None) -> Tuple[bool, str]:
        """Actualizar plantilla existente"""
        if template_id not in self.templates:
            return False, "Plantilla no encontrada"
        
        template = self.templates[template_id]
        
        # Actualizar campos proporcionados
        if name is not None:
            if not name.strip():
                return False, "El nombre no puede estar vacío"
            template["name"] = name.strip()
        
        if subject is not None:
            if not subject.strip():
                return False, "El asunto no puede estar vacío"
            template["subject"] = subject.strip()
        
        if body is not None:
            if not body.strip():
                return False, "El cuerpo no puede estar vacío"
            template["body"] = body.strip()
        
        # Manejar cambio de plantilla predeterminada
        if set_as_default is not None:
            if set_as_default:
                # Desmarcar otras como predeterminadas
                for t in self.templates.values():
                    t['is_default'] = False
                template['is_default'] = True
            else:
                template['is_default'] = False
                # Asegurar que al menos una quede como predeterminada
                if not any(t.get('is_default', False) for t in self.templates.values()):
                    # Si no queda ninguna predeterminada, marcar la primera
                    if self.templates:
                        list(self.templates.values())[0]['is_default'] = True
        
        template["updated_at"] = datetime.now().isoformat()
        
        if self._save_templates():
            return True, "Plantilla actualizada exitosamente"
        else:
            return False, "Error guardando los cambios"
    
    def delete_template(self, template_id: str) -> Tuple[bool, str]:
        """Eliminar plantilla"""
        if template_id not in self.templates:
            return False, "Plantilla no encontrada"
        
        template = self.templates[template_id]
        was_default = template.get('is_default', False)
        
        # No permitir eliminar si es la única plantilla
        if len(self.templates) <= 1:
            return False, "No se puede eliminar la única plantilla disponible"
        
        del self.templates[template_id]
        
        # Si la plantilla eliminada era predeterminada, marcar otra como predeterminada
        if was_default and self.templates:
            list(self.templates.values())[0]['is_default'] = True
        
        if self._save_templates():
            return True, "Plantilla eliminada exitosamente"
        else:
            return False, "Error eliminando la plantilla"
    
    def duplicate_template(self, template_id: str, new_name: str = None) -> Tuple[bool, str]:
        """Duplicar plantilla existente"""
        if template_id not in self.templates:
            return False, "Plantilla no encontrada"
        
        original = self.templates[template_id]
        
        if not new_name:
            new_name = f"Copia de {original['name']}"
        
        return self.create_template(
            name=new_name,
            subject=original['subject'],
            body=original['body'],
            set_as_default=False
        )
    
    def set_default_template(self, template_id: str) -> Tuple[bool, str]:
        """Establecer plantilla como predeterminada"""
        if template_id not in self.templates:
            return False, "Plantilla no encontrada"
        
        # Desmarcar todas como predeterminadas
        for template in self.templates.values():
            template['is_default'] = False
        
        # Marcar la seleccionada como predeterminada
        self.templates[template_id]['is_default'] = True
        self.templates[template_id]['updated_at'] = datetime.now().isoformat()
        
        if self._save_templates():
            return True, "Plantilla establecida como predeterminada"
        else:
            return False, "Error guardando los cambios"
    
    def get_template_for_email(self, template_id: str = None) -> Tuple[str, str]:
        """Obtener asunto y cuerpo de plantilla para usar en envío de email"""
        if template_id and template_id in self.templates:
            template = self.templates[template_id]
        else:
            template = self.get_default_template()
        
        return template['subject'], template['body']
    
    def validate_template_variables(self, subject: str, body: str) -> List[str]:
        """Validar que las variables requeridas estén presentes"""
        required_vars = ['{NOMBRE}', '{MES}']
        missing_vars = []
        
        combined_text = subject + " " + body
        
        for var in required_vars:
            if var not in combined_text:
                missing_vars.append(var)
        
        return missing_vars
    
    def get_available_variables(self) -> List[Dict[str, str]]:
        """Obtener lista de variables disponibles para usar en plantillas"""
        return [
            {
                "variable": "{NOMBRE}",
                "description": "Nombre del destinatario del correo"
            },
            {
                "variable": "{MES}",
                "description": "Mes correspondiente a la boleta"
            }
        ]