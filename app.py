import streamlit as st
from pptx import Presentation
from pptx.util import Pt, Inches
import pdfplumber
import io

# Configuración de la interfaz
st.set_page_config(page_title="Editor MT Valero Pro", layout="wide")

# Inicialización de memoria de sesión
if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("🚀 Generador de Reportes Pro (PPTX + PDF)")

# --- BARRA LATERAL: CARGA DE RECURSOS ---
with st.sidebar:
    st.header("1. Recursos")
    plantilla = st.file_uploader("Subir Plantilla Base (.pptx)", type=["pptx"])
    archivo_pdf = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"]) #
    fotos_subidas = st.file_uploader("Galería de Fotos (WhatsApp)", type=["jpg", "png", "jpeg"], accept_multiple_files=True) #
    
    # LÓGICA DE EXTRACCIÓN DEL PDF
    if archivo_pdf and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto_completo = ""
            for pagina in pdf.pages:
                texto_completo += pagina.extract_text() + "\n"
            
            # Buscamos patrones comunes en reportes de Valero
            for linea in texto_completo.split('\n'):
                if "Cliente:" in linea:
                    st.session_state.datos_pdf["cliente"] = linea.split("Cliente:")[1].strip()
                if "Fecha:" in linea:
                    st.session_state.datos_pdf["fecha"] = linea.split("Fecha:")[1].strip()
        st.success(f"✅ Datos extraídos: {st.session_state.datos_pdf['cliente']}")

# --- INTERFAZ PRINCIPAL ---
col_form, col_preview = st.columns([1, 1])

with col_form:
    st.subheader("✍️ Configurar Nueva Hoja")
    # Estos campos se llenan automáticamente tras la extracción
    cliente = st.text_input("Cliente:", value=st.session_state.datos_pdf["cliente"])
    fecha = st.text_input("Fecha:", value=st.session_state.datos_pdf["fecha"])
    contenido = st.text_area("Descripción (Times New Roman 12):", height=150)
    
    # OPCIÓN PARA ASIGNAR FOTOS A ESTA HOJA
    fotos_seleccionadas = []
    if fotos_subidas:
        st.write("📸 Selecciona fotos para esta hoja específica:")
        for f in fotos_subidas:
            # Creamos una casilla de verificación por cada foto
            if st.checkbox(f"Incluir {f.name}", key=f"foto_{f.name}_{len(st.session_state.hojas)}"):
                fotos_seleccionadas.append(f)

    if st.button("➕ GUARDAR HOJA AL REPORTE"):
        nueva_hoja = {
            "cliente": cliente,
            "fecha": fecha,
            "contenido": contenido,
            "fotos": fotos_seleccionadas
        }
        st.session_state.hojas.append(nueva_hoja)
        st.success(f"Hoja #{len(st.session_state.hojas)} guardada correctamente.")

with col_preview:
    st.subheader("👁️ Vista Previa del Reporte") #
    if st.session_state.hojas:
        num_hoja = st.number_input("Navegar por hojas:", min_value=1, max_value=len(st.session_state.hojas), step=1)
        h = st.session_state.hojas[num_hoja-1]
        
        with st.container(border=True):
            st.markdown(f"**Visualizando Hoja:** {num_hoja}")
            st.markdown(f"**Cliente:** {h['cliente']}")
            st.markdown(f"**Contenido:** {h['contenido']}")
            st.write(f"🖼️ Fotos en esta hoja: {len(h['fotos'])}")
            
        if st.button("🗑️ Eliminar esta hoja"):
            st.session_state.hojas.pop(num_hoja-1)
            st.rerun()
    else:
        st.info("Aún no has añadido ninguna hoja al reporte.") #

# --- GENERACIÓN FINAL DEL PPTX ---
st.divider()
if st.session_state.hojas and plantilla:
    if st.button("🚀 GENERAR REPORTE FINAL"):
        prs = Presentation(io.BytesIO(plantilla.read()))
        
        for h in st.session_state.hojas:
            slide = prs.slides.add_slide(prs.slide_layouts[1])
            
            # Insertar texto con formato
            for shape in slide.placeholders:
                if "TITLE" in shape.name.upper():
                    shape.text = h['cliente']
                elif any(x in shape.name.upper() for x in ["BODY", "CONTENT"]):
                    tf = shape.text_frame
                    tf.text = h['contenido']
                    for p in tf.paragraphs:
                        for run in p.runs:
                            run.font.name = 'Times New Roman'
                            run.font.size = Pt(12)
            
            # Insertar fotos seleccionadas para esta diapositiva
            left = Inches(0.5)
            for f in h['fotos']:
                slide.shapes.add_picture(io.BytesIO(f.read()), left, Inches(4.5), width=Inches(2.5))
                left += Inches(2.7)
                f.seek(0)

        output = io.BytesIO()
        prs.save(output)
        output.seek(0)
        
        st.download_button(
            label="📥 DESCARGAR REPORTE FINAL (.PPTX)",
            data=output.getvalue(),
            file_name="Reporte_Final_MT_Valero.pptx",
            mime="application/vnd.openxmlformats-officedocument.presentationml.presentation"
        )
        