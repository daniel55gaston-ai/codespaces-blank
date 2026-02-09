import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
import pdfplumber
import io

# Configuración de página
st.set_page_config(page_title="Editor MT Valero Pro", layout="wide")

if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("🛠️ Centro de Control de Reportes MT")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Carga de Recursos")
    plantilla = st.file_uploader("Subir Plantilla Base (.pptx)", type=["pptx"])
    archivo_pdf = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"])
    fotos_totales = st.file_uploader("Galería de WhatsApp", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
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
    st.subheader("📝 Configurar Nueva Hoja")
    
    if plantilla:
        # Cargamos la plantilla temporalmente para leer los nombres de sus diapositivas
        prs_temp = Presentation(io.BytesIO(plantilla.getvalue()))
        # Obtenemos los nombres de los diseños que ya tiene tu PowerPoint
        nombres_disenos = [layout.name for layout in prs_temp.slide_layouts]
        
        with st.form("form_hoja", clear_on_submit=True):
            # 1. ELIGES EL NOMBRE QUE YA TIENE ASIGNADO EL PPT
            diseno_elegido = st.selectbox("Selecciona el tipo de diapositiva (Nombre asignado):", nombres_disenos)
            
            cliente = st.text_input("Cliente:", value=st.session_state.datos_pdf["cliente"])
            descripcion = st.text_area("Descripción Técnica (Times New Roman 12):")
            
            st.write("🖼️ **Selecciona fotos para esta hoja:**")
            fotos_seleccionadas = []
            if fotos_totales:
                cols_fotos = st.columns(3)
                for i, f in enumerate(fotos_totales):
                    with cols_fotos[i % 3]:
                        st.image(f, width=100)
                        if st.checkbox("Incluir", key=f"sel_{f.name}_{len(st.session_state.hojas)}"):
                            fotos_seleccionadas.append(f)
            
            enviar = st.form_submit_button("➕ GUARDAR HOJA Y LIMPIAR")
            
            if enviar:
                idx_diseno = nombres_disenos.index(diseno_elegido)
                nueva_hoja = {
                    "titulo": diseno_elegido, # El nombre asignado en el PPT
                    "idx_layout": idx_diseno,
                    "cliente": cliente,
                    "fecha": st.session_state.datos_pdf["fecha"],
                    "descripcion": descripcion,
                    "fotos": fotos_seleccionadas
                }
                st.session_state.hojas.append(nueva_hoja)
                st.success(f"✅ Hoja '{diseno_elegido}' guardada.")
                st.rerun()
    else:
        st.warning("⚠️ Sube una plantilla (.pptx) en la izquierda para ver los nombres de las diapositivas.")

with col_prev:
    st.subheader("👁️ Vista Previa del Reporte")
    if st.session_state.hojas:
        idx = st.number_input("Navegar por Hoja #:", min_value=1, max_value=len(st.session_state.hojas), step=1) - 1
        h = st.session_state.hojas[idx]
        
        with st.container(border=True):
            st.markdown(f"### {idx+1}. {h['titulo']}")
            st.markdown(f"**Cliente:** {h['cliente']} | **Fecha:** {h['fecha']}")
            st.info(f"**Texto:** {h['descripcion']}")
            if h['fotos']:
                st.write(f"🖼️ {len(h['fotos'])} fotos seleccionadas.")
        
        if st.button("🗑️ Eliminar Hoja Actual"):
            st.session_state.hojas.pop(idx)
            st.rerun()

# --- GENERACIÓN FINAL ---
st.divider()
if st.session_state.hojas and plantilla:
    if st.button("🚀 DESCARGAR REPORTE FINAL"):
        prs = Presentation(io.BytesIO(plantilla.read()))
        for h in st.session_state.hojas:
            # USA EL DISEÑO EXACTO QUE ELEGISTE POR NOMBRE
            slide = prs.slides.add_slide(prs.slide_layouts[h['idx_layout']])
            
            for shape in slide.placeholders:
                nombre_shape = shape.name.upper()
                if "TITLE" in nombre_shape or "TITULO" in nombre_shape: 
                    shape.text = h['titulo']
                elif any(x in nombre_shape for x in ["BODY", "CONTENT"]):
                    tf = shape.text_frame
                    tf.text = h['descripcion']
                    for p in tf.paragraphs:
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)
            
            x_pos = Inches(0.5)
            for f in h['fotos']:
                slide.shapes.add_picture(io.BytesIO(f.read()), x_pos, Inches(4.5), width=Inches(2.4))
                x_pos += Inches(2.6)
                f.seek(0)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        st.download_button("📥 DESCARGAR PPTX", output.getvalue(), "Reporte_MT_Final.pptx")
        