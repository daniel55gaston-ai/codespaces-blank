import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
from pptx.dml.color import RGBColor
import pdfplumber
import io

# 1. Configuración de página
st.set_page_config(page_title="Generador MT Valero", layout="wide")

# 2. Mantener el reporte vivo en la memoria de la sesión
if 'mi_reporte' not in st.session_state:
    st.session_state.mi_reporte = None
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("🚀 Generador de Reportes Final")

with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Plantilla (.pptx)", type=["pptx"])
    archivo_pdf = st.file_uploader("Subir Hoja de Trabajo (PDF)", type=["pdf"])
    fotos = st.file_uploader("Subir Fotos", type=["jpg", "png", "jpeg"], accept_multiple_files=True)

# Lógica principal
if plantilla:
    # Cargar la plantilla solo la primera vez
    if st.session_state.mi_reporte is None:
        template_bytes = plantilla.read()
        st.session_state.mi_reporte = Presentation(io.BytesIO(template_bytes))
    
    prs = st.session_state.mi_reporte

    # Extraer datos del PDF
    if archivo_pdf and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for linea in texto.split('\n'):
                if "Cliente:" in linea:
                    st.session_state.datos_pdf["cliente"] = linea.split("Cliente:")[1].strip()
                if "Fecha:" in linea:
                    st.session_state.datos_pdf["fecha"] = linea.split("Fecha:")[1].strip()
        st.success("✅ Datos extraídos del PDF")

    # Campos de edición
    cliente_final = st.text_input("Nombre del Cliente:", value=st.session_state.datos_pdf["cliente"])
    texto_grande = st.text_area("Contenido Técnico (Times New Roman 12):", height=150)

    # 3. BOTÓN PARA AÑADIR DIAPOSITIVA (Asegura el guardado)
    if st.button("➕ AÑADIR HOJA AL REPORTE"):
        # Usamos el layout 1 (Título y Objetos)
        slide_layout = prs.slide_layouts[1] 
        slide = prs.slides.add_slide(slide_layout)
        
        # Insertar Texto
        for shape in slide.placeholders:
            nombre = shape.name.upper()
            if "TITLE" in nombre or "TITULO" in nombre:
                shape.text = cliente_final
            elif any(x in nombre for x in ["CONTENT", "BODY", "CUADRO"]):
                tf = shape.text_frame
                tf.text = texto_grande
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
        
        # Insertar Fotos automáticamente si se subieron
        if fotos:
            left = Inches(1)
            top = Inches(4)
            for foto in fotos:
                slide.shapes.add_picture(io.BytesIO(foto.read()), left, top, width=Inches(3))
                left += Inches(3.5) # Espaciado entre fotos
                foto.seek(0) # Reset para que no se pierda la imagen
        
        # GUARDAR CAMBIOS EN LA SESIÓN
        st.session_state.mi_reporte = prs
        st.success(f"✅ ¡Hoja añadida! El reporte ahora tiene {len(prs.slides)} diapositivas.")

    st.divider()
    
    # 4. BOTÓN DE DESCARGA (Con corrección de puntero)
    if len(prs.slides) > 0:
        output = io.BytesIO()
        prs.save(output)
        output.seek(0) # ESTO ES LO QUE HACIA QUE SALIERA VACÍO
        
        st.download_button(
            label="📥 DESCARGAR REPORTE CON TODA LA INFO",
            data=output.getvalue(),
            file_name="Reporte_MT_Valero_Final.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
    else:
        st.warning("Añade al menos una diapositiva antes de descargar.")

else:
    st.info("Sube tu plantilla para empezar.")
    