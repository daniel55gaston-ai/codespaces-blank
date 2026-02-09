import streamlit as st
import pdfplumber
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.utils import ImageReader
import io

# Configuración de página
st.set_page_config(page_title="Generador PDF MT Valero", layout="wide")

# Inicialización de memoria de sesión
if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("📄 Generador de Reportes PDF Editable")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Carga de Recursos")
    archivo_pdf = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"])
    fotos_totales = st.file_uploader("Galería de WhatsApp", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
    if archivo_pdf and st.button("🔍 Escanear PDF de Origen"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for l in texto.split('\n'):
                if "Cliente:" in l: st.session_state.datos_pdf["cliente"] = l.split("Cliente:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_pdf["fecha"] = l.split("Fecha:")[1].strip()
        st.success("Datos de Valero extraídos.")

# --- INTERFAZ DE TRABAJO ---
col_edit, col_prev = st.columns([1, 1])

with col_edit:
    st.subheader("📝 Configurar Nueva Hoja")
    
    with st.form("form_pdf", clear_on_submit=True):
        titulo_hoja = st.text_input("Título de la Hoja:", placeholder="Ej: Inspección de Válvula de Seguridad")
        cliente = st.text_input("Cliente:", value=st.session_state.datos_pdf["cliente"])
        descripcion = st.text_area("Descripción Técnica (Times New Roman 12):", height=150)
        
        st.write("🖼️ **Selecciona las fotos para esta hoja:**")
        fotos_seleccionadas = []
        if fotos_totales:
            cols_img = st.columns(3)
            for i, f in enumerate(fotos_totales):
                with cols_img[i % 3]:
                    st.image(f, width=100)
                    if st.checkbox("Incluir", key=f"foto_{f.name}_{len(st.session_state.hojas)}"):
                        fotos_seleccionadas.append(f)
        
        if st.form_submit_button("➕ GUARDAR HOJA AL PDF"):
            if titulo_hoja or descripcion:
                nueva_hoja = {
                    "titulo": titulo_hoja,
                    "cliente": cliente,
                    "fecha": st.session_state.datos_pdf["fecha"],
                    "descripcion": descripcion,
                    "fotos": fotos_seleccionadas
                }
                st.session_state.hojas.append(nueva_hoja)
                st.rerun()

with col_prev:
    st.subheader("👁️ Vista Previa del Reporte")
    if st.session_state.hojas:
        idx = st.number_input("Ver Hoja #:", min_value=1, max_value=len(st.session_state.hojas), step=1) - 1
        h = st.session_state.hojas[idx]
        
        with st.container(border=True):
            st.markdown(f"### {idx+1}. {h['titulo']}")
            st.markdown(f"**Cliente:** {h['cliente']} | **Fecha:** {h['fecha']}")
            st.write(f"**Texto:** {h['descripcion']}")
            if h['fotos']:
                st.write(f"🖼️ {len(h['fotos'])} fotos en esta página.")
        
        if st.button("🗑️ Eliminar esta hoja"):
            st.session_state.hojas.pop(idx)
            st.rerun()
    else:
        st.info("Añade hojas a la izquierda para ver la previa.")

# --- GENERACIÓN DEL PDF FINAL ---
st.divider()
if st.session_state.hojas:
    if st.button("🚀 GENERAR Y DESCARGAR PDF FINAL"):
        buffer = io.BytesIO()
        c = canvas.Canvas(buffer, pagesize=letter)
        width, height = letter

        for h in st.session_state.hojas:
            # Encabezado
            c.setFont("Times-Bold", 14)
            c.drawString(50, height - 50, f"REPORTE: {h['titulo'].upper()}")
            c.setFont("Times-Roman", 10)
            c.drawString(50, height - 70, f"Cliente: {h['cliente']} | Fecha: {h['fecha']}")
            c.line(50, height - 75, 550, height - 75)

            # Cuerpo de texto (Times New Roman 12)
            c.setFont("Times-Roman", 12)
            text_obj = c.beginText(50, height - 100)
            # Dividir texto por líneas para que quepa
            lines = h['descripcion'].split('\n')
            for line in lines:
                text_obj.textLine(line)
            c.drawText(text_obj)

            # Insertar Fotos (Acomodo automático)
            if h['fotos']:
                x_offset = 50
                y_offset = height - 450
                for f in h['fotos']:
                    img_data = ImageReader(io.BytesIO(f.read()))
                    c.drawImage(img_data, x_offset, y_offset, width=150, height=150, preserveAspectRatio=True)
                    x_offset += 170
                    if x_offset > 400: # Salto de fila de fotos
                        x_offset = 50
                        y_offset -= 160
                    f.seek(0)

            c.showPage() # Siguiente página

        c.save()
        buffer.seek(0)
        st.download_button(
            label="📥 DESCARGAR REPORTE EN PDF",
            data=buffer,
            file_name="Reporte_Final_Valero.pdf",
            mime="application/pdf"
        )