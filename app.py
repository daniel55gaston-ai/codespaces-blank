import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import pdfplumber
import io

# Configuración de la interfaz
st.set_page_config(page_title="Generador MT Valero", layout="wide")

# Inicialización de la memoria de sesión
if 'mi_reporte' not in st.session_state:
    st.session_state.mi_reporte = None
if 'datos_auto' not in st.session_state:
    st.session_state.datos_auto = {"Cliente": "", "Ubicacion": "", "Fecha": "", "Serie": ""}

st.title("🚀 Generador de Reportes Final")

# --- BARRA LATERAL: CARGA DE ARCHIVOS ---
with st.sidebar:
    st.header("1. Carga de Archivos")
    # Se usan llaves (keys) únicas para evitar errores de ID duplicado
    plantilla = st.file_uploader("Subir Plantilla (.pptx)", type=None, key="u_pptx")
    archivo_pdf = st.file_uploader("Subir PDF (Hoja de Trabajo)", type="pdf", key="u_pdf")
    fotos = st.file_uploader("Fotos de WhatsApp", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="u_fotos")

# --- LÓGICA DE PROCESAMIENTO ---
if plantilla:
    if st.session_state.mi_reporte is None:
        st.session_state.mi_reporte = Presentation(io.BytesIO(plantilla.read()))
    
    prs = st.session_state.mi_reporte

    # Escaneo del PDF basado en tu formato de Valero
    if archivo_pdf and st.button("🔍 Escanear datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            lineas = texto.split('\n')
            for l in lineas:
                if "Cliente:" in l: st.session_state.datos_auto["Cliente"] = l.split("Cliente:")[1].strip()
                if "Ubicación:" in l: st.session_state.datos_auto["Ubicacion"] = l.split("Ubicación:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_auto["Fecha"] = l.split("Fecha:")[1].strip()
        st.success("Datos extraídos correctamente.")

    col1, col2 = st.columns([1, 1.2])

    with col1:
        st.subheader("🎨 Configuración de Diseño")
        nombres_layouts = [f"{i}: {l.name}" for i, l in enumerate(prs.slide_layouts)]
        sel_layout = st.selectbox("Elegir diseño del patrón:", nombres_layouts)
        idx_layout = int(sel_layout.split(":")[0])
        layout_actual = prs.slide_layouts[idx_layout]

        # Selector de fotos por sección
        ph_fotos = [s for s in layout_actual.placeholders if "Picture" in s.name or "Imagen" in s.name]
        asignaciones = {}
        if ph_fotos:
            st.write("**Asignar fotos a espacios:**")
            for ph in ph_fotos:
                opciones = ["Ninguna"] + [f.name for f in fotos] if fotos else ["Ninguna"]
                asignaciones[ph.name] = st.selectbox(f"Espacio: {ph.name}", opciones, key=f"f_{ph.name}")

    with col2:
        st.subheader("✍️ Información de la Hoja")
        c_nombre = st.text_input("Cliente:", value=st.session_state.datos_auto["Cliente"], key="in_cliente")
        c_fecha = st.text_input("Fecha:", value=st.session_state.datos_auto["Fecha"], key="in_fecha")
        
        st.write("**Descripción Técnica (Times New Roman 12):**")
        texto_grande = st.text_area("Notas del trabajo:", height=180, key="in_area_texto")

    # BOTÓN PARA GENERAR DIAPOSITIVA
    if st.button("➕ Añadir Diapositiva"):
        nueva_slide = prs.slides.add_slide(layout_actual)
        
        for shape in nueva_slide.placeholders:
            nombre_ph = shape.name.upper()
            
            # Llenar datos de texto
            if "CLIENT" in nombre_ph: shape.text = c_nombre
            elif "DATE" in nombre_ph or "FECHA" in nombre_ph: shape.text = c_fecha
            
            # Aplicar formato Times New Roman 12
            elif any(k in nombre_ph for k in ["BODY", "CONTENT", "DESCRIPCION"]):
                tf = shape.text_frame
                tf.text = texto_grande
                for p in tf.paragraphs:
                    for run in p.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
                        run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Insertar fotos según selección
            if shape.name in asignaciones and asignaciones[shape.name] != "Ninguna":
                f_file = next(f for f in fotos if f.name == asignaciones[shape.name])
                img_stream = io.BytesIO(f_file.read())
                shape.insert_picture(img_stream)
                f_file.seek(0)
                
        st.balloons()
        st.success(f"¡Hoja añadida! Total: {len(prs.slides)}")

    # DESCARGA
    st.divider()
    output_ppt = io.BytesIO()
    prs.save(output_ppt)
    st.download_button(
        label="📥 DESCARGAR REPORTE FINAL",
        data=output_ppt.getvalue(),
        file_name="Reporte_Final.pptx",
        mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
    )
    
    if st.button("🗑️ Reiniciar Reporte"):
        st.session_state.mi_reporte = None
        st.rerun()

else:
    st.info("👋 Sube tu plantilla para comenzar.")
    