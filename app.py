import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
import pdfplumber
import io

# Configuración de página
st.set_page_config(page_title="Editor MT Valero Pro", layout="wide")

# Inicialización de memoria de sesión
if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("📸 Generador de Reportes Inteligente")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Plantilla Base (.pptx)", type=["pptx"])
    archivo_pdf = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"])
    fotos_totales = st.file_uploader("Galería Completa (WhatsApp)", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
    if archivo_pdf and st.button("🔍 Escanear PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for l in texto.split('\n'):
                if "Cliente:" in l: st.session_state.datos_pdf["cliente"] = l.split("Cliente:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_pdf["fecha"] = l.split("Fecha:")[1].strip()
        st.success("Datos extraídos.")

# --- INTERFAZ DE TRABAJO ---
col_edit, col_prev = st.columns([1, 1])

with col_edit:
    st.subheader("✍️ Configurar Hoja")
    
    # Usamos un formulario para poder limpiar todo al enviar
    with st.form("form_hoja", clear_on_submit=True):
        cliente = st.text_input("Cliente:", value=st.session_state.datos_pdf["cliente"])
        fecha = st.text_input("Fecha:", value=st.session_state.datos_pdf["fecha"])
        descripcion = st.text_area("Descripción Técnica (Times New Roman 12):")
        
        st.write("🖼️ **Selecciona fotos para esta hoja:**")
        fotos_seleccionadas = []
        if fotos_totales:
            for f in fotos_totales:
                # Cada checkbox tiene una clave única basada en el total de hojas creadas
                if st.checkbox(f"Incluir {f.name}", key=f"sel_{f.name}_{len(st.session_state.hojas)}"):
                    fotos_seleccionadas.append(f)
        
        enviar = st.form_submit_button("➕ GUARDAR HOJA Y LIMPIAR")
        
        if enviar:
            if descripcion or fotos_seleccionadas:
                nueva_hoja = {
                    "cliente": cliente,
                    "fecha": fecha,
                    "descripcion": descripcion,
                    "fotos": fotos_seleccionadas
                }
                st.session_state.hojas.append(nueva_hoja)
                st.success(f"✅ Hoja #{len(st.session_state.hojas)} guardada. Formulario listo para la siguiente.")
                st.rerun()
            else:
                st.warning("Escribe algo o selecciona una foto antes de guardar.")

with col_prev:
    st.subheader("👁️ Vista Previa del Reporte")
    if st.session_state.hojas:
        idx = st.number_input("Revisar Hoja #:", min_value=1, max_value=len(st.session_state.hojas), step=1) - 1
        h = st.session_state.hojas[idx]
        
        with st.container(border=True):
            st.markdown(f"**Cliente:** {h['cliente']} | **Fecha:** {h['fecha']}")
            st.markdown(f"**Descripción:** {h['descripcion']}")
            st.write(f"🖼️ Fotos: {len(h['fotos'])}")
            for img in h['fotos']:
                st.caption(f"✅ {img.name}")
        
        if st.button("🗑️ Borrar esta hoja"):
            st.session_state.hojas.pop(idx)
            st.rerun()
    else:
        st.info("Las hojas guardadas aparecerán aquí.")

# --- GENERACIÓN FINAL ---
st.divider()
if st.session_state.hojas and plantilla:
    if st.button("🚀 DESCARGAR REPORTE FINAL"):
        prs = Presentation(io.BytesIO(plantilla.read()))
        for h in st.session_state.hojas:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            for shape in slide.placeholders:
                if "TITLE" in shape.name.upper(): shape.text = h['cliente']
                elif any(x in shape.name.upper() for x in ["BODY", "CONTENT"]):
                    tf = shape.text_frame
                    tf.text = h['descripcion']
                    for p in tf.paragraphs:
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)
            
            # Acomodo de fotos
            x_pos = Inches(0.5)
            for f in h['fotos']:
                slide.shapes.add_picture(io.BytesIO(f.read()), x_pos, Inches(4.5), width=Inches(2.4))
                x_pos += Inches(2.6)
                f.seek(0)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        st.download_button("📥 DESCARGAR PPTX", output.getvalue(), "Reporte_MT.pptx")
        