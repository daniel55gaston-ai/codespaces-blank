import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import pdfplumber
import io

# Configuración de página
st.set_page_config(page_title="Generador MT Valero", layout="wide")

# Memoria de la sesión para los datos del PDF
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}
if 'mi_reporte' not in st.session_state:
    st.session_state.mi_reporte = None

st.title("🚀 Generador de Reportes Final")

with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Plantilla (.pptx)", type=["pptx"], key="u_pptx")
    # SECCIÓN DEL PDF AGREGADA
    archivo_pdf = st.file_uploader("Subir Hoja de Trabajo (PDF)", type=["pdf"], key="u_pdf")
    fotos = st.file_uploader("Subir Fotos", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="u_fotos")

if plantilla:
    if st.session_state.mi_reporte is None:
        st.session_state.mi_reporte = Presentation(io.BytesIO(plantilla.read()))
    
    prs = st.session_state.mi_reporte

    # Lógica para leer el PDF de Valero
    if archivo_pdf and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            # Buscamos datos específicos en el texto del PDF
            for linea in texto.split('\n'):
                if "Cliente:" in linea:
                    st.session_state.datos_pdf["cliente"] = linea.split("Cliente:")[1].strip()
                if "Fecha:" in linea:
                    st.session_state.datos_pdf["fecha"] = linea.split("Fecha:")[1].strip()
        st.success("✅ Datos extraídos: " + st.session_state.datos_pdf["cliente"])

    # Interfaz de edición
    col1, col2 = st.columns(2)
    with col1:
        cliente_final = st.text_input("Nombre del Cliente:", value=st.session_state.datos_pdf["cliente"])
    with col2:
        fecha_final = st.text_input("Fecha:", value=st.session_state.datos_pdf["fecha"])

    st.subheader("✍️ Descripción del Trabajo")
    texto_grande = st.text_area("Contenido (Formato Times New Roman 12):", height=150)

    if st.button("➕ Añadir Diapositiva"):
        # Se añade diapositiva usando el diseño de la plantilla
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        
        for shape in slide.placeholders:
            nombre = shape.name.upper()
            
            # Llenar Cliente y Fecha si existen los espacios
            if "CLIENT" in nombre: shape.text = cliente_final
            elif "DATE" in nombre or "FECHA" in nombre: shape.text = fecha_final
            
            # Formato de texto para el cuerpo
            elif any(x in nombre for x in ["CONTENT", "BODY", "DESCRIPCION"]):
                tf = shape.text_frame
                tf.text = texto_grande
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
                        run.font.color.rgb = RGBColor(0, 0, 0)
        st.success("✅ Diapositiva añadida.")

    # Descarga del archivo final
    st.divider()
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    
    st.download_button(
        label="📥 DESCARGAR REPORTE FINAL (.PPTX)",
        data=output,
        file_name="Reporte_Finalizado.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    