import streamlit as st
from docx import Document
from docx.shared import Pt, Cm, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from PIL import Image
import io
import re
from datetime import datetime

# Configuración de la página
st.set_page_config(
    page_title="AGOIN - Formateador de Documentos",
    page_icon="🏢",
    layout="wide"
)

# CSS personalizado con los colores corporativos de AGOIN
st.markdown("""
<style>
    .main-header {
        background: linear-gradient(135deg, #1a5c4d 0%, #2d8b73 100%);
        padding: 2rem;
        border-radius: 10px;
        color: white;
        text-align: center;
        margin-bottom: 2rem;
    }
    .info-box {
        background-color: #f0f8f5;
        border-left: 5px solid #1a5c4d;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .stButton>button {
        background-color: #1a5c4d;
        color: white;
        font-weight: bold;
        border-radius: 5px;
        padding: 0.5rem 2rem;
        border: none;
    }
    .stButton>button:hover {
        background-color: #2d8b73;
    }
</style>
""", unsafe_allow_html=True)

# Encabezado principal
st.markdown("""
<div class="main-header">
    <h1>📄 AGOIN - Formateador Corporativo</h1>
    <p>Convierte cualquier documento al formato corporativo de AGOIN</p>
</div>
""", unsafe_allow_html=True)

def extract_project_info(doc):
    """Extrae información del proyecto del documento original"""
    project_info = {
        'title': '',
        'location': '',
        'section': ''
    }

    # Buscar en los primeros 10 párrafos
    text = ' '.join([p.text for p in doc.paragraphs[:10]])

    # Patrones para detectar información
    # Buscar palabras clave como "PROYECTO", "VIVIENDA", etc.
    project_pattern = re.search(r'PROYECTO[^\n]{0,200}', text, re.IGNORECASE)
    if project_pattern:
        project_info['title'] = project_pattern.group(0).strip()

    # Buscar dirección o ubicación
    location_pattern = re.search(r'(?:CALLE|AVENIDA|AVDA|C/)[^\n]{0,150}', text, re.IGNORECASE)
    if location_pattern:
        project_info['location'] = location_pattern.group(0).strip()

    # Detectar tipo de memoria o sección
    section_keywords = ['MEMORIA DESCRIPTIVA', 'MEMORIA CONSTRUCTIVA', 'MEMORIA JUSTIFICATIVA', 
                       'PLIEGO', 'PRESUPUESTO', 'MEDICIONES']
    for keyword in section_keywords:
        if keyword.lower() in text.lower():
            project_info['section'] = keyword
            break

    return project_info

def apply_agoin_format(input_doc, project_title, project_location, section_name):
    """Aplica el formato corporativo de AGOIN al documento"""

    # Crear nuevo documento con formato AGOIN
    output_doc = Document()

    # Configurar márgenes (2.5cm arriba/abajo, 3cm izq/der)
    for section in output_doc.sections:
        section.top_margin = Cm(2.5)
        section.bottom_margin = Cm(2.5)
        section.left_margin = Cm(3.0)
        section.right_margin = Cm(3.0)
        section.page_height = Cm(29.7)  # A4
        section.page_width = Cm(21.0)

        # Configurar encabezado
        header = section.header
        header.is_linked_to_previous = False

        # Logo y título en encabezado
        header_para1 = header.paragraphs[0]
        header_para1.text = project_title if project_title else "[TÍTULO DEL PROYECTO]"
        header_para1.alignment = WD_ALIGN_PARAGRAPH.LEFT

        header_para2 = header.add_paragraph()
        header_para2.text = project_location if project_location else "[UBICACIÓN DEL PROYECTO]"
        header_para2.alignment = WD_ALIGN_PARAGRAPH.LEFT

        header_para3 = header.add_paragraph()
        section_text = f"{section_name}         " if section_name else "[SECCIÓN]         "
        header_para3.text = section_text
        header_para3.alignment = WD_ALIGN_PARAGRAPH.LEFT

        # Configurar pie de página
        footer = section.footer
        footer.is_linked_to_previous = False

        footer_para1 = footer.paragraphs[0]
        footer_para1.text = "ARQUITECTURA Y GESTION DE OPERACIONES INMOBILIARIAS S.L.P."
        footer_para1.alignment = WD_ALIGN_PARAGRAPH.CENTER

        footer_para2 = footer.add_paragraph()
        footer_para2.text = "AVDA DE IRLANDA, 24 4ºD 45005 TOLEDO"
        footer_para2.alignment = WD_ALIGN_PARAGRAPH.CENTER

        footer_para3 = footer.add_paragraph()
        footer_para3.text = "925.29.93.00 info@agoin.es"
        footer_para3.alignment = WD_ALIGN_PARAGRAPH.CENTER

    # Copiar contenido del documento original
    for element in input_doc.element.body:
        # Copiar párrafos
        if element.tag.endswith('p'):
            for para in input_doc.paragraphs:
                new_para = output_doc.add_paragraph()

                # Copiar formato de texto
                for run in para.runs:
                    new_run = new_para.add_run(run.text)
                    if run.bold:
                        new_run.bold = True
                    if run.italic:
                        new_run.italic = True
                    if run.underline:
                        new_run.underline = True

                # Mantener alineación
                new_para.alignment = para.alignment

        # Copiar tablas
        elif element.tag.endswith('tbl'):
            for table in input_doc.tables:
                # Crear nueva tabla con las mismas dimensiones
                new_table = output_doc.add_table(rows=len(table.rows), cols=len(table.columns))
                new_table.style = 'Table Grid'

                # Copiar contenido de celdas
                for i, row in enumerate(table.rows):
                    for j, cell in enumerate(row.cells):
                        new_table.rows[i].cells[j].text = cell.text

    return output_doc

# Interfaz principal
col1, col2 = st.columns([2, 1])

with col1:
    st.markdown("### 📤 Subir Documento")
    uploaded_file = st.file_uploader(
        "Selecciona un archivo Word (.docx)",
        type=['docx'],
        help="Sube el documento que deseas convertir al formato AGOIN"
    )

with col2:
    st.markdown("""
    <div class="info-box">
        <h4>ℹ️ Información</h4>
        <p><strong>Formato soportado:</strong> Word (.docx)</p>
        <p><strong>Preserva:</strong> Texto, tablas, imágenes</p>
        <p><strong>Aplica:</strong> Formato corporativo AGOIN</p>
    </div>
    """, unsafe_allow_html=True)

if uploaded_file is not None:
    try:
        # Cargar documento
        doc = Document(uploaded_file)

        st.success("✅ Documento cargado correctamente")

        # Extraer información automáticamente
        st.markdown("### 🔍 Información Detectada")
        project_info = extract_project_info(doc)

        col_a, col_b = st.columns(2)

        with col_a:
            project_title = st.text_area(
                "Título del Proyecto",
                value=project_info['title'],
                help="Se extrajo automáticamente. Puedes modificarlo si es necesario.",
                height=100
            )

            section_name = st.text_input(
                "Nombre de la Sección",
                value=project_info['section'],
                help="Ejemplo: MEMORIA DESCRIPTIVA, MEMORIA CONSTRUCTIVA, etc."
            )

        with col_b:
            project_location = st.text_area(
                "Ubicación del Proyecto",
                value=project_info['location'],
                help="Se extrajo automáticamente. Puedes modificarlo si es necesario.",
                height=100
            )

        # Botón para procesar
        st.markdown("### 🔄 Procesar Documento")

        if st.button("🚀 Convertir al Formato AGOIN", use_container_width=True):
            with st.spinner("Procesando documento..."):
                try:
                    # Aplicar formato AGOIN
                    output_doc = apply_agoin_format(
                        doc, 
                        project_title, 
                        project_location, 
                        section_name
                    )

                    # Guardar en memoria
                    output_buffer = io.BytesIO()
                    output_doc.save(output_buffer)
                    output_buffer.seek(0)

                    st.success("✅ ¡Documento formateado exitosamente!")

                    # Botón de descarga
                    st.download_button(
                        label="📥 Descargar Documento Formateado",
                        data=output_buffer,
                        file_name=f"AGOIN_{uploaded_file.name}",
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )

                except Exception as e:
                    st.error(f"❌ Error al procesar: {str(e)}")

    except Exception as e:
        st.error(f"❌ Error al cargar el documento: {str(e)}")

# Pie de página
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Desarrollado para AGOIN - Arquitectura y Gestión de Operaciones Inmobiliarias S.L.P.</p>
    <p>© 2025 - Todos los derechos reservados</p>
</div>
""", unsafe_allow_html=True)
