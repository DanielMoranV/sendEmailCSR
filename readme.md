# 📧 Aplicación de Envío de Boletas - Clínica Santa Rosa

Una aplicación de escritorio desarrollada en Python para el envío automatizado de boletas de pago por correo electrónico a los empleados de la Clínica Santa Rosa.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![GUI](https://img.shields.io/badge/GUI-ttkbootstrap-green)
![License](https://img.shields.io/badge/License-MIT-yellow)

## 🚀 Características

- **Interfaz gráfica moderna** con tema oscuro usando ttkbootstrap
- **Envío masivo de correos** con archivos PDF adjuntos
- **Validación automática** de emails y DNIs
- **Animación de carga** con progreso en tiempo real
- **Prevención de envíos duplicados** deshabilitando controles durante procesamiento
- **Reporte de errores** automático en formato Excel
- **Logging completo** de todas las operaciones
- **Configuración por variables de entorno** para mayor seguridad

## 📋 Requisitos del Sistema

### Software
- Python 3.8 o superior
- Windows (para abrir archivos automáticamente)

### Dependencias Python
```
pandas
openpyxl
ttkbootstrap
python-dotenv
```

## 🛠️ Instalación

1. **Clona o descarga el proyecto:**
```bash
git clone <url-del-repositorio>
cd email-sender-app
```

2. **Instala las dependencias:**
```bash
pip install pandas openpyxl ttkbootstrap python-dotenv
```

3. **Configura las variables de entorno:**
Crea un archivo `.env` en el directorio raíz con el siguiente contenido:
```env
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=465
EMAIL_USER=tu-email@gmail.com
EMAIL_PASSWORD=tu-contraseña-de-aplicacion
```

> **⚠️ Importante:** Para Gmail, necesitas usar una "Contraseña de Aplicación" en lugar de tu contraseña normal.

## 📁 Estructura de Archivos

```
proyecto/
├── main.py                 # Archivo principal de la aplicación
├── .env                    # Variables de entorno (no incluir en git)
├── README.md              # Este archivo
├── email_log.txt          # Log automático de operaciones
├── logError.xlsx          # Reporte de errores (generado automáticamente)
└── C:/BoletasCSR/         # Directorio por defecto de boletas
    ├── enero/
    │   ├── 12345678.pdf
    │   └── 87654321.pdf
    ├── febrero/
    └── ...
```

## 📊 Formato del Archivo Excel

El archivo Excel debe contener las siguientes columnas (sin importar el orden):

| Columna | Tipo | Descripción | Ejemplo |
|---------|------|-------------|---------|
| `nombre` | Texto | Nombre completo del empleado | Juan Pérez |
| `email` | Texto | Correo electrónico válido | juan.perez@email.com |
| `dni` | Texto/Número | DNI de 8 dígitos | 12345678 |

### Ejemplo de archivo Excel:
```
| nombre      | email                | dni      |
|-------------|---------------------|----------|
| Juan Pérez  | juan.perez@mail.com | 12345678 |
| Ana García  | ana.garcia@mail.com | 87654321 |
| Luis Torres | luis.torres@mail.com| 11223344 |
```

## 🎯 Uso de la Aplicación

### 1. Iniciar la aplicación
```bash
python main.py
```

### 2. Configurar parámetros
- **Mes:** Selecciona el mes correspondiente a las boletas
- **Ruta de Boletas:** Directorio donde están los PDFs (por defecto: `C:/BoletasCSR`)
- **Archivo Excel:** Selecciona el archivo con los datos de empleados

### 3. Iniciar envío
- Haz clic en "Enviar Boletas"
- La aplicación mostrará el progreso en tiempo real
- Los botones se deshabilitarán durante el procesamiento

### 4. Revisar resultados
- Verás un resumen con enviados correctamente y errores
- Si hay errores, se generará automáticamente `logError.xlsx`
- Revisa `email_log.txt` para detalles técnicos

## ✅ Validaciones Automáticas

La aplicación valida automáticamente:

- **DNI:** Debe tener exactamente 8 dígitos numéricos
- **Email:** Formato válido de correo electrónico
- **Archivo PDF:** Debe existir en la ruta especificada con formato `{DNI}.pdf`
- **Conexión SMTP:** Verifica conectividad antes de enviar

## 📝 Plantilla de Correo

Cada correo enviado incluye:

- **Asunto:** "Boleta del mes de {Mes}"
- **Remitente:** Clínica Santa Rosa
- **Contenido HTML** personalizado con el nombre del empleado
- **Archivo adjunto:** PDF de la boleta correspondiente

## 🔧 Configuración SMTP

### Gmail
```env
SMTP_SERVER=smtp.gmail.com
SMTP_PORT=465
EMAIL_USER=tu-email@gmail.com
EMAIL_PASSWORD=contraseña-de-aplicacion
```

### Outlook/Hotmail
```env
SMTP_SERVER=smtp.live.com
SMTP_PORT=587
EMAIL_USER=tu-email@outlook.com
EMAIL_PASSWORD=tu-contraseña
```

### Otros proveedores
Consulta la documentación de tu proveedor de email para obtener los valores correctos.

## 📊 Manejo de Errores

### Tipos de errores registrados:
- **DNI inválido:** DNI que no tiene 8 dígitos
- **Email inválido:** Formato de correo incorrecto
- **PDF no encontrado:** El archivo de boleta no existe
- **Error SMTP:** Problemas al enviar el correo

### Archivo de errores:
Si ocurren errores, se genera `logError.xlsx` con:
- Número de fila en el Excel original
- Nombre del empleado
- Email
- DNI
- Motivo del error

## 🚨 Solución de Problemas

### Error de conexión SMTP
- Verifica las credenciales en el archivo `.env`
- Para Gmail, asegúrate de usar una contraseña de aplicación
- Verifica que la verificación en 2 pasos esté activada (Gmail)

### Archivos PDF no encontrados
- Verifica que la ruta sea correcta
- Los PDFs deben nombrarse con el DNI (ej: `12345678.pdf`)
- Asegúrate de que la carpeta del mes exista

### Problemas con el archivo Excel
- Verifica que contenga las columnas: `nombre`, `email`, `dni`
- Evita celdas vacías en estas columnas
- Guarda el archivo en formato `.xlsx`

## 🔒 Seguridad

- Las credenciales se almacenan en archivo `.env` (no incluir en control de versiones)
- Usa contraseñas de aplicación específicas, no tu contraseña principal
- Los logs no contienen información sensible de autenticación

## 📈 Funciones de Logging

La aplicación registra automáticamente:
- Conexiones SMTP exitosas
- Correos enviados correctamente
- Errores durante el envío
- Archivos Excel procesados
- Validaciones fallidas

## 🤝 Contribución

Para contribuir al proyecto:
1. Fork el repositorio
2. Crea una rama para tu feature (`git checkout -b feature/nueva-funcionalidad`)
3. Commit tus cambios (`git commit -am 'Agrega nueva funcionalidad'`)
4. Push a la rama (`git push origin feature/nueva-funcionalidad`)
5. Crea un Pull Request

## 📄 Licencia

Este proyecto está bajo la Licencia MIT. Ver el archivo `LICENSE` para más detalles.

## 👥 Soporte

Para soporte técnico o reportar bugs:
- Crea un issue en el repositorio
- Incluye los logs relevantes de `email_log.txt`
- Describe los pasos para reproducir el problema

---

**Desarrollado para Clínica Santa Rosa** 🏥