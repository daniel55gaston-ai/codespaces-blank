import streamlit as st
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import pdfplumber
import io

# Configuración profesional de la página
st.set_page_config(page_title="Generador MT Valero", layout="wide")

# Inicialización de la memoria de sesión
if 'mi_reporte' not in st.session_state:
    st.session_state.mi_reporte = None
if 'datos_auto' not in st.session_state:
    st.session_state.datos_auto = {"Cliente": "", "Ubicacion": "", "Fecha": "", "Serie": ""}

st.title("🚀 Constructor de Reportes MT / Valero")

# --- BARRA LATERAL: CARGA DE ARCHIVOS ---
with st.sidebar:
    st.header("📂 1. Carga de Archivos")
    # Keys únicas para evitar errores de ID duplicado
    plantilla = st.file_uploader("Subir Plantilla (.pptx)", type=None, key="u_pptx")
    archivo_pdf = st.file_uploader("Subir PDF (Hoja de Trabajo Valero)", type="pdf", key="u_pdf")
    fotos = st.file_uploader("Fotos de WhatsApp/Evidencia", type=["jpg", "png", "jpeg"], accept_multiple_files=True, key="u_fotos")

# --- LÓGICA PRINCIPAL ---
if plantilla:
    if st.session_state.mi_reporte is None:
        st.session_state.mi_reporte = Presentation(io.BytesIO(plantilla.read()))
    
    prs = st.session_state.mi_reporte

    # Escaneo automático del PDF de Valero
    if archivo_pdf and st.button("🔍 Escanear datos del PDF automáticamente"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto_full = pdf.pages[0].extract_text()
            lineas = texto_full.split('\n')
            for l in lineas:
                if "Cliente:" in l: st.session_state.datos_auto["Cliente"] = l.split("Cliente:")[1].strip()
                if "Ubicación:" in l: st.session_state.datos_auto["Ubicacion"] = l.split("Ubicación:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_auto["Fecha"] = l.split("Fecha:")[1].strip()
        st.toast("Datos del PDF cargados con éxito", icon="✅")

    # INTERFAZ DE EDICIÓN
    col_diseno, col_datos = st.columns([1, 1.2])

    with col_diseno:
        st.subheader("🎨 Configuración de Hoja")
        nombres_layouts = [f"{i}: {l.name}" for i, l in enumerate(prs.slide_layouts)]
        sel_layout = st.selectbox("Selecciona el Diseño de Patrón:", nombres_layouts)
        idx_layout = int(sel_layout.split(":")[0])
        layout_actual = prs.slide_layouts[idx_layout]

        # Identificar cuadros para fotos
        ph_fotos = [s for s in layout_actual.placeholders if "Picture" in s.name or "Imagen" in s.name]
        asignaciones_fotos = {}
        if ph_fotos and fotos:
            st.write("📸 **Asignar fotos a secciones:**")
            for ph in ph_fotos:
                asignaciones_fotos[ph.name] = st.selectbox(f"Espacio: {ph.name}", ["Ninguna"] + [f.name for f in fotos], key=f"sel_{ph.name}")

    with col_datos:
        st.subheader("✍️ Información de la Diapositiva")
        cliente = st.text_input("Cliente:", value=st.session_state.datos_auto["Cliente"])
        ubicacion = st.text_input("Ubicación:", value=st.session_state.datos_auto["Ubicacion"])
        fecha = st.text_input("Fecha:", value=st.session_state.datos_auto["Fecha"])
        
        st.write("**Descripción Técnica (Times New Roman 12):**")
        texto_grande = st.text_area("Escribe el contenido aquí:", height=200, placeholder="Escribe la descripción del trabajo realizado...")

    # ACCIÓN: AÑADIR DIAPOSITIVA
    if st.button("🚀 Añadir Diapositiva al Reporte"):
        nueva_slide = prs.slides.add_slide(layout_actual)
        
        for shape in nueva_slide.placeholders:
            nombre_ph = shape.name.upper()
            
            # Llenar textos automáticos
            if "CLIENT" in nombre_ph: shape.text = cliente
            elif "LOCATION" in nombre_ph or "UBICACION" in nombre_ph: shape.text = ubicacion
            elif "DATE" in nombre_ph or "FECHA" in nombre_ph: shape.text = fecha
            
            # Formato Times New Roman 12 para el cuadro central
            elif any(k in nombre_ph for k in ["BODY", "CONTENT", "DESCRIPCION"]):
                tf = shape.text_frame
                tf.text = texto_grande
                for paragraph in tf.paragraphs:
                    for run in paragraph.runs:
                        run.font.name = 'Times New Roman'
                        run.font.size = Pt(12)
                        run.font.color.rgb = RGBColor(0, 0, 0)
            
            # Insertar fotos seleccionadas
            if shape.name in asignaciones_fotos:
                foto_nombre = asignaciones_fotos[shape.name]
                if foto_nombre != "Ninguna":
                    archivo_foto = next(f for f in fotos if f.name == foto_nombre)
                    shape.insert_picture(io.BytesIO(archivo_foto.read()))
                    archivo_foto.seek(0)

        st.success(f"✅ Diapositiva añadida. Total: {len(prs.slides)}")

    # BOTONES DE CIERRE
    st.divider()
    output = io.BytesIO()
    prs.save(output)
    st.download_button("📥 DESCARGAR REPORTE PPTX", data=output.getvalue(), file_name="Reporte_Finalizado.pptx")
    
    if st.button("🗑️ Reiniciar Todo"):
        st.session_state.mi_reporte = None
        st.rerun()
else:
    st.info("👋 Por favor, carga tu plantilla .pptx en la barra lateral para comenzar.")
    