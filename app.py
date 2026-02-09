import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import pdfplumber
import io

# Configuración de página
st.set_page_config(page_title="Generador MT Valero", layout="wide")

if 'mi_reporte' not in st.session_state:
    st.session_state.mi_reporte = None

st.title("🚀 Generador de Reportes Final")

with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Plantilla (.pptx)", type=["pptx"], key="u_pptx")
    fotos = st.file_uploader("Subir Fotos", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="u_fotos")

if plantilla:
    if st.session_state.mi_reporte is None:
        st.session_state.mi_reporte = Presentation(io.BytesIO(plantilla.read()))
    
    prs = st.session_state.mi_reporte
    
    # Editor de texto con formato Times New Roman 12
    st.subheader("✍️ Descripción del Trabajo")
    texto_grande = st.text_area("Contenido:", height=150)

    if st.button("➕ Añadir Diapositiva"):
        # Usar el diseño 1 (título y contenido) por defecto
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        
        for shape in slide.placeholders:
            if "CONTENT" in shape.name.upper() or "BODY" in shape.name.upper():
                tf = shape.text_frame
                tf.text = texto_grande
                # Aplicar formato Times New Roman 12
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
                        run.font.color.rgb = RGBColor(0, 0, 0)
        st.success("✅ Diapositiva añadida con éxito.")

    # Sección de Descarga Segura
    st.divider()
    output = io.BytesIO()
    prs.save(output)
    output.seek(0) # Volver al inicio del archivo para la descarga
    
    st.download_button(
        label="📥 DESCARGAR REPORTE FINAL (.PPTX)",
        data=output,
        file_name="Reporte_MT_Valero.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    