import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import pdfplumber
import io
import subprocess
import os

# Configuración de la aplicación
st.set_page_config(page_title="MT Valero - PDF Editable", layout="wide")

if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("📄 Generador de Reportes PDF Editable")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Reporte_MT_Final.pptx", type=["pptx"])
    archivo_pdf = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"])
    fotos_totales = st.file_uploader("Galería de WhatsApp", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
    if archivo_pdf and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf) as pdf:
            texto = pdf.pages[0].extract_text()
            for l in texto.split('\n'):
                if "Cliente:" in l: st.session_state.datos_pdf["cliente"] = l.split("Cliente:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_pdf["fecha"] = l.split("Fecha:")[1].strip()
        st.success("Datos extraídos.")

# --- CUERPO PRINCIPAL ---
if plantilla:
    prs_base = Presentation(io.BytesIO(plantilla.getvalue()))
    nombres_layouts = [layout.name for layout in prs_base.slide_layouts]

    col_edit, col_prev = st.columns([1, 1])

    with col_edit:
        st.subheader("📝 Configurar Diapositiva")
        with st.form("form_hoja", clear_on_submit=True):
            diseno = st.selectbox("Seleccionar diseño:", nombres_layouts)
            # El texto que irá directo a tu tabla
            contenido_tabla = st.text_area("Contenido Técnico (Times New Roman 12):")
            
            st.write("🖼️ **Fotos para esta página:**")
            fotos_escogidas = []
            if fotos_totales:
                cols = st.columns(3)
                for i, f in enumerate(fotos_totales):
                    with cols[i % 3]:
                        st.image(f, width=100)
                        # Clave única para evitar errores de duplicado
                        if st.checkbox("Incluir", key=f"sel_{f.name}_{len(st.session_state.hojas)}"):
                            fotos_escogidas.append(f)
            
            if st.form_submit_button("➕ GUARDAR HOJA"):
                st.session_state.hojas.append({
                    "layout_idx": nombres_layouts.index(diseno),
                    "texto": contenido_tabla,
                    "fotos": fotos_escogidas
                })
                st.rerun()

    # --- GENERACIÓN Y DESCARGA ---
    st.divider()
    if st.session_state.hojas:
        # Generamos el PPTX primero (es el molde para el PDF)
        prs_final = Presentation(io.BytesIO(plantilla.getvalue()))
        # Limpiar diapos originales
        for i in range(len(prs_final.slides)-1, -1, -1):
            rId = prs_final.slides._sldIdLst[i].rId
            prs_final.part.drop_rel(rId)
            del prs_final.slides._sldIdLst[i]

        for h in st.session_state.hojas:
            slide = prs_final.slides.add_slide(prs_final.slide_layouts[h['layout_idx']])
            img_ptr = 0
            for shape in slide.shapes:
                # Escribir en la tabla con Times New Roman 12
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text == "" or "DESCRIPCIÓN" in cell.text.upper():
                                cell.text = h['texto']
                                for p in cell.text_frame.paragraphs:
                                    for r in p.runs:
                                        r.font.name = 'Times New Roman'
                                        r.font.size = Pt(12)
                # Inyectar fotos en tus cuadros técnicos
                elif shape.is_placeholder and (shape.placeholder_format.type == 18 or "PICTURE" in shape.name.upper()):
                    if img_ptr < len(h['fotos']):
                        shape.insert_picture(io.BytesIO(h['fotos'][img_ptr].read()))
                        h['fotos'][img_ptr].seek(0)
                        img_ptr += 1

        output_pptx = io.BytesIO()
        prs_final.save(output_pptx)
        
        col_pptx, col_pdf = st.columns(2)
        with col_pptx:
            st.download_button("📥 DESCARGAR PPTX", output_pptx.getvalue(), "Reporte.pptx")
        
        with col_pdf:
            if st.button("🚀 PREPARAR PDF EDITABLE"):
                with open("temp.pptx", "wb") as f:
                    f.write(output_pptx.getvalue())
                try:
                    # Conversión real usando LibreOffice instalado en terminal
                    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "temp.pptx"], check=True)
                    with open("temp.pdf", "rb") as f:
                        st.download_button("✅ DESCARGAR PDF FINAL", f.read(), "Reporte_Editable.pdf", "application/pdf")
                    os.remove("temp.pptx")
                    os.remove("temp.pdf")
                except:
                    st.error("Error: Ejecuta primero el comando de instalación en la terminal.")
else:
    st.info("👈 Sube tu plantilla 'Reporte_MT_Final.pptx' para comenzar.")
    