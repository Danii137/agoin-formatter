# 🏢 AGOIN - Formateador Corporativo de Documentos

Aplicación web que convierte cualquier documento Word al formato corporativo de **AGOIN - Arquitectura y Gestión de Operaciones Inmobiliarias S.L.P.**

## 🌟 Características

- ✅ **Extracción automática** de información del proyecto
- 📝 **Preserva** contenido, tablas e imágenes
- 🎨 **Aplica formato corporativo** AGOIN (márgenes, encabezados, pies de página)
- 🔄 **Interfaz intuitiva** y fácil de usar
- ☁️ **100% en línea** - No requiere instalación
- 💰 **Gratis** - Desplegable en Streamlit Cloud

## 🚀 Instalación Local

### Requisitos previos
- Python 3.8 o superior
- pip (gestor de paquetes de Python)

### Pasos de instalación

1. **Clonar el repositorio**
```bash
git clone https://github.com/TU_USUARIO/agoin-document-formatter.git
cd agoin-document-formatter
```

2. **Crear entorno virtual (recomendado)**
```bash
python -m venv venv
source venv/bin/activate  # En Windows: venv\Scripts\activate
```

3. **Instalar dependencias**
```bash
pip install -r requirements.txt
```

4. **Ejecutar la aplicación**
```bash
streamlit run app.py
```

5. **Abrir en el navegador**
La aplicación se abrirá automáticamente en `http://localhost:8501`

## ☁️ Despliegue en Streamlit Cloud (GRATIS)

### Opción 1: Despliegue automático desde GitHub

1. **Sube tu proyecto a GitHub**
   - Crea un nuevo repositorio en GitHub
   - Sube todos los archivos del proyecto

2. **Accede a Streamlit Cloud**
   - Ve a https://share.streamlit.io
   - Inicia sesión con tu cuenta de GitHub

3. **Despliega la aplicación**
   - Haz clic en "New app"
   - Selecciona tu repositorio
   - Selecciona la rama (main/master)
   - Archivo principal: `app.py`
   - Haz clic en "Deploy"

4. **¡Listo!** Tu aplicación estará disponible en una URL pública

### Opción 2: Despliegue directo

```bash
# Instala streamlit si no lo tienes
pip install streamlit

# Despliega directamente
streamlit run app.py --server.port 8501
```

## 📖 Uso

1. **Sube tu documento**
   - Haz clic en "Subir Documento"
   - Selecciona un archivo Word (.docx)

2. **Revisa la información extraída**
   - La aplicación detecta automáticamente:
     * Título del proyecto
     * Ubicación
     * Tipo de sección
   - Puedes editar cualquier campo si es necesario

3. **Convierte el documento**
   - Haz clic en "Convertir al Formato AGOIN"
   - Espera unos segundos

4. **Descarga el resultado**
   - Haz clic en "Descargar Documento Formateado"
   - El documento tendrá el formato corporativo AGOIN completo

## 🎨 Formato Corporativo AGOIN

El documento formateado incluye:

### Márgenes
- Superior: 2.5 cm
- Inferior: 2.5 cm
- Izquierdo: 3.0 cm
- Derecho: 3.0 cm

### Encabezado
- Título del proyecto
- Ubicación del proyecto
- Nombre de la sección con numeración

### Pie de página
- Nombre de la empresa: ARQUITECTURA Y GESTION DE OPERACIONES INMOBILIARIAS S.L.P.
- Dirección: AVDA DE IRLANDA, 24 4ºD 45005 TOLEDO
- Contacto: 925.29.93.00 info@agoin.es

### Contenido
- Preserva todo el texto original
- Mantiene tablas con formato adaptado
- Conserva imágenes incluidas

## 🛠️ Tecnologías Utilizadas

- **Streamlit** - Framework para aplicaciones web en Python
- **python-docx** - Manipulación de documentos Word
- **Pillow** - Procesamiento de imágenes
- **PyPDF2** - Soporte para PDFs (futuras versiones)

## 📝 Estructura del Proyecto

```
document_formatter/
│
├── app.py                 # Aplicación principal Streamlit
├── requirements.txt       # Dependencias del proyecto
├── README.md             # Este archivo
├── assets/               # Recursos (logo, etc.)
├── templates/            # Plantillas de documentos
└── utils/                # Funciones auxiliares
```

## 🤝 Contribuciones

Las contribuciones son bienvenidas. Por favor:

1. Haz fork del proyecto
2. Crea una rama para tu función (`git checkout -b feature/AmazingFeature`)
3. Commit tus cambios (`git commit -m 'Add some AmazingFeature'`)
4. Push a la rama (`git push origin feature/AmazingFeature`)
5. Abre un Pull Request

## 📄 Licencia

Este proyecto es propiedad de **AGOIN - Arquitectura y Gestión de Operaciones Inmobiliarias S.L.P.**

## 📧 Contacto

**AGOIN**
- 📍 AVDA DE IRLANDA, 24 4ºD 45005 TOLEDO
- 📞 925.29.93.00
- ✉️ info@agoin.es

## 🐛 Reporte de Errores

Si encuentras algún error, por favor abre un issue en GitHub con:
- Descripción del problema
- Pasos para reproducirlo
- Capturas de pantalla (si aplica)
- Documento de ejemplo (si es posible)

---

Desarrollado con ❤️ para AGOIN
