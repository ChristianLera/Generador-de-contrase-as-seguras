# 🔐 Gestor Profesional de Contraseñas Seguras

Aplicación de escritorio con interfaz gráfica para generar, almacenar y gestionar contraseñas seguras. Desarrollada en Python con Tkinter y almacenamiento en Excel.

## ✨ Características

- 🔑 **Generador de contraseñas** con parámetros personalizables:
  - Longitud ajustable (8-64 caracteres)
  - Inclusión/exclusión de mayúsculas, minúsculas, dígitos y símbolos
  - Evitar caracteres ambiguos (il1Lo0O)
- 💾 **Bóveda de contraseñas** con almacenamiento en Excel
- 🔍 **Búsqueda** de contraseñas por descripción
- 📊 **Estadísticas** de seguridad y uso de la bóveda
- 📋 **Copiar al portapapeles** con un clic
- 📤 **Importar/Exportar** a archivos CSV
- 💿 **Respaldo automático** de la base de datos
- 🎨 **Interfaz moderna** con tema oscuro profesional

## 🖥️ Capturas de pantalla

*[Agrega aquí capturas de pantalla de tu aplicación]*

## 📋 Requisitos

- Python 3.8 o superior
- Pip (gestor de paquetes de Python)

## 🚀 Instalación y ejecución

### Windows (archivo batch)

```bash
ejecutar.bat
```

### Windows PowerShell

```powershell
.\ejecutar.ps1
```

### Manual

```bash
pip install -r requirements.txt
python GeneradorDeContraseñasSeguras.py
```

## 📦 Dependencias

| Paquete | Versión | Propósito |
|---------|---------|-----------|
| pandas | >=1.3.0 | Gestión de datos en Excel |
| openpyxl | >=3.0.9 | Motor de lectura/escritura Excel |
| pyperclip | >=1.8.2 | Copiar al portapapeles |

## 🗂️ Estructura del proyecto

```
Gestor-Contrasenas/
├── GeneradorDeContraseñasSeguras.py   # Aplicación principal
├── requirements.txt                   # Dependencias Python
├── ejecutar.bat                       # Lanzador para Windows
├── ejecutar.ps1                       # Lanzador para PowerShell
├── mis_contrasenas.xlsx              # Base de datos (se crea automáticamente)
└── README.md                          # Este archivo
```

## 🛠️ Uso

### Generador de contraseñas
1. Ajusta los parámetros (longitud, tipos de caracteres)
2. Haz clic en "Generar"
3. Copia la contraseña con "Copiar"
4. Guarda la contraseña con descripción y categoría

### Bóveda de contraseñas
- Visualiza todas las contraseñas guardadas
- Busca por descripción
- Copia la contraseña real al portapapeles
- Elimina contraseñas obsoletas
- Importa/Exporta a CSV

### Estadísticas
- Total de contraseñas guardadas
- Distribución por categoría
- Niveles de seguridad
- Últimas contraseñas agregadas

## 🔒 Seguridad

- Las contraseñas se almacenan en un archivo Excel **sin encriptación** (recomendación: guarda el archivo en una unidad cifrada o usa herramientas como VeraCrypt)
- El generador utiliza el módulo `secrets` de Python, considerado criptográficamente seguro
- Las contraseñas se enmascaran en la interfaz (solo se muestra parcialmente)

## 👨‍💻 Autor

**Christian Lera**

## 📄 Licencia

Este proyecto es de uso libre para fines educativos y personales.
