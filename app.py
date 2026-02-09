import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
import pdfplumber
import io

# Configuración de página
st.set_page_config(page_title="Editor de Reportes MT", layout="wide")

# Memoria de sesión para gestionar múltiples hojas y fotos
if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("📋 Generador de Reportes Pro (PPTX + PDF)")

# --- BARRA LATERAL: CARGA DE RECURSOS ---
with st.sidebar:
    st.header("1. Recursos")
    plantilla = st.file_uploader("Subir Plantilla Base (.pptx)", type=["pptx"])
    archivo_pdf = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"])
    fotos_subidas = st.file_uploader("Galería de Fotos (WhatsApp)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
    if archivo_pdf and st.button("🔍 Escanear PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for l in texto.split('\n'):
                if "Cliente:" in l: st.session_state.datos_pdf["cliente"] = l.split("Cliente:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_pdf["fecha"] = l.split("Fecha:")[1].strip()
        st.success("Datos de Valero cargados.")

# --- INTERFAZ PRINCIPAL ---
col_form, col_preview = st.columns([1, 1])

with col_form:
    st.subheader("✍️ Configurar Nueva Hoja")
    cliente = st.text_input("Cliente:", value=st.session_state.datos_pdf["cliente"])
    fecha = st.text_input("Fecha:", value=st.session_state.datos_pdf["fecha"])
    contenido = st.text_area("Descripción (Times New Roman 12):", height=150)
    
    # Selector de fotos para ESTA hoja específica
    fotos_seleccionadas = []
    if fotos_subidas:
        st.write("📸 Selecciona fotos para esta hoja:")
        for f in fotos_subidas:
            if st.checkbox(f"Incluir {f.name}", key=f"check_{f.name}_{len(st.session_state.hojas)}"):
                fotos_seleccionadas.append(f)

    if st.button("➕ GUARDAR HOJA AL REPORTE"):
        nueva_hoja = {
            "cliente": cliente,
            "fecha": fecha,
            "contenido": contenido,
            "fotos": fotos_seleccionadas
        }
        st.session_state.hojas.append(nueva_hoja)
        st.success(f"Hoja #{len(st.session_state.hojas)} lista.")

with col_preview:
    st.subheader("👁️ Vista Previa y Control")
    if st.session_state.hojas:
        num_hoja = st.number_input("Ir a la hoja:", min_value=1, max_value=len(st.session_state.hojas), step=1)
        h = st.session_state.hojas[num_hoja-1]
        
        with st.container(border=True):
            st.markdown(f"**Hoja Actual:** {num_hoja}")
            st.markdown(f"**Cliente:** {h['cliente']} | **Fecha:** {h['fecha']}")
            st.markdown(f"<p style='font-family:Times New Roman; font-size:14px;'>{h['contenido']}</p>", unsafe_allow_html=True)
            if h['fotos']:
                st.write(f"🖼️ Fotos asignadas: {len(h['fotos'])}")
        
        if st.button("🗑️ Eliminar Hoja Actual"):
            st.session_state.hojas.pop(num_hoja-1)
            st.rerun()
    else:
        st.info("Aún no has creado hojas.")

# --- GENERACIÓN FINAL ---
st.divider()
if st.session_state.hojas and plantilla:
    if st.button("🚀 GENERAR Y DESCARGAR REPORTE FINAL"):
        prs = Presentation(io.BytesIO(plantilla.read()))
        
        for h in st.session_state.hojas:
            # Añadir slide usando el diseño 1 de tu PPT
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            
            # Rellenar textos con formato
            for shape in slide.placeholders:
                if "TITLE" in shape.name.upper(): shape.text = h['cliente']
                elif any(x in shape.name.upper() for x in ["BODY", "CONTENT", "CUADRO"]):
                    tf = shape.text_frame
                    tf.text = h['contenido']
                    for p in tf.paragraphs:
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)
            
            # Insertar fotos en la hoja correspondiente
            left = Inches(0.5)
            for f in h['fotos']:
                slide.shapes.add_picture(io.BytesIO(f.read()), left, Inches(4.5), width=Inches(2.5))
                left += Inches(2.7)
                f.seek(0) # Reset para reuso

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 DESCARGAR PPTX EDITADO",
            data=output.getvalue(),
            file_name="Reporte_MT_Equipment_Final.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        