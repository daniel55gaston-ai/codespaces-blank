import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import pdfplumber
import io

# Configuración de página
st.set_page_config(page_title="Generador MT Valero", layout="wide")

if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}
if 'mi_reporte' not in st.session_state:
    st.session_state.mi_reporte = None

st.title("🚀 Generador de Reportes Final")

with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Plantilla (.pptx)", type=["pptx"], key="u_pptx")
    archivo_pdf = st.file_uploader("Subir Hoja de Trabajo (PDF)", type=["pdf"], key="u_pdf")
    fotos = st.file_uploader("Subir Fotos", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="u_fotos")

if plantilla:
    if st.session_state.mi_reporte is None:
        st.session_state.mi_reporte = Presentation(io.BytesIO(plantilla.read()))
    
    prs = st.session_state.mi_reporte

    if archivo_pdf and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for linea in texto.split('\n'):
                if "Cliente:" in linea:
                    st.session_state.datos_pdf["cliente"] = linea.split("Cliente:")[1].strip()
                if "Fecha:" in linea:
                    st.session_state.datos_pdf["fecha"] = linea.split("Fecha:")[1].strip()
        st.success("✅ Datos extraídos")

    cliente_final = st.text_input("Nombre del Cliente:", value=st.session_state.datos_pdf["cliente"])
    texto_grande = st.text_area("Contenido (Times New Roman 12):", height=150)

    if st.button("➕ Añadir Diapositiva"):
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        for shape in slide.placeholders:
            nombre = shape.name.upper()
            if "CLIENT" in nombre: shape.text = cliente_final
            elif any(x in nombre for x in ["CONTENT", "BODY", "DESCRIPCION"]):
                tf = shape.text_frame
                tf.text = texto_grande
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
        st.success("✅ Diapositiva añadida.")

    st.divider()
    
    # --- PROCESO DE GUARDADO COMPATIBLE ---
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    
    # Intento de descarga con extensión .ppt para forzar apertura en versiones viejas
    # Aunque el contenido sea XML, el cambio de MIME ayuda a saltar bloqueos de seguridad
    st.download_button(
        label="📥 DESCARGAR PARA VERSIONES ANTIGUAS (.PPT)",
        data=output,
        file_name="Reporte_Valero.ppt",
        mime="application/vnd.ms-powerpoint"
    )
    
    # Opción estándar corregida
    st.download_button(
        label="📥 DESCARGAR ESTÁNDAR (.PPTX)",
        data=output,
        file_name="Reporte_Valero.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    