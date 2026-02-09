import streamlit as st
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import io

# Configuración de la interfaz
st.set_page_config(page_title="Generador MT Valero PRO", layout="wide")

# Inicialización de la memoria de sesión (Crucial para no perder datos)
if 'hojas' not in st.session_state:
    st.session_state.hojas = [] # Lista de todas las hojas creadas
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("📋 Generador de Reportes PDF (Editable)")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Carga de Archivos")
    archivo_pdf = st.file_uploader("Subir Hoja de Trabajo (PDF Valero)", type=["pdf"])
    
    if archivo_pdf and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for linea in texto.split('\n'):
                if "Cliente:" in linea: st.session_state.datos_pdf["cliente"] = linea.split("Cliente:")[1].strip()
                if "Fecha:" in linea: st.session_state.datos_pdf["fecha"] = linea.split("Fecha:")[1].strip()
        st.success("Datos cargados correctamente.")

# --- INTERFAZ DE TRABAJO ---
col_edit, col_prev = st.columns([1, 1])

with col_edit:
    st.subheader("✍️ Editar Hoja Actual")
    cliente = st.text_input("Nombre del Cliente:", value=st.session_state.datos_pdf["cliente"])
    fecha = st.text_input("Fecha:", value=st.session_state.datos_pdf["fecha"])
    contenido = st.text_area("Descripción Técnica (Times New Roman 12):", height=200, placeholder="Escribe aquí el resultado del mantenimiento...")
    
    if st.button("➕ Guardar y Añadir otra Hoja"):
        nueva_hoja = {"cliente": cliente, "fecha": fecha, "contenido": contenido}
        st.session_state.hojas.append(nueva_hoja)
        st.balloons()
        st.success(f"Hoja #{len(st.session_state.hojas)} guardada.")

with col_prev:
    st.subheader("👁️ Vista Previa del Reporte")
    if st.session_state.hojas:
        # Navegador entre hojas
        index = st.number_input("Ver Hoja número:", min_value=1, max_value=len(st.session_state.hojas), step=1) - 1
        hoja_actual = st.session_state.hojas[index]
        
        # Simulación de la hoja en pantalla
        st.info(f"Mostrando Hoja {index + 1} de {len(st.session_state.hojas)}")
        with st.container(border=True):
            st.markdown(f"**Cliente:** {hoja_actual['cliente']}")
            st.markdown(f"**Fecha:** {hoja_actual['fecha']}")
            st.divider()
            st.markdown(f"<p style='font-family:Times New Roman; font-size:16px;'>{hoja_actual['contenido']}</p>", unsafe_allow_html=True)
            
        if st.button("🗑️ Borrar esta hoja"):
            st.session_state.hojas.pop(index)
            st.rerun()
    else:
        st.write("Aún no has añadido ninguna hoja al reporte.")

# --- GENERACIÓN DE PDF FINAL ---
st.divider()
if st.session_state.hojas:
    if st.button("🚀 GENERAR PDF FINAL"):
        buffer = io.BytesIO()
        p = canvas.Canvas(buffer, pagesize=letter)
        
        for hoja in st.session_state.hojas:
            # Configuración de la página PDF
            p.setFont("Times-Roman", 16)
            p.drawString(50, 750, f"Cliente: {hoja['cliente']}")
            p.drawString(450, 750, f"Fecha: {hoja['fecha']}")
            p.line(50, 740, 550, 740)
            
            # Contenido en Times New Roman 12
            p.setFont("Times-Roman", 12)
            text_obj = p.beginText(50, 700)
            text_obj.textLines(hoja['contenido'])
            p.drawText(text_obj)
            
            p.showPage() # Nueva hoja
        
        p.save()
        buffer.seek(0)
        
        st.download_button(
            label="📥 DESCARGAR REPORTE EN PDF",
            data=buffer,
            file_name="Reporte_Final_MT_Valero.pdf",
            mime="application/pdf"
        )
        