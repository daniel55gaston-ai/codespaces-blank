import streamlit as st
from pptx import Presentation
from pptx.util import Pt
import pdfplumber
import io
import subprocess
import os

# Configuración de la App
st.set_page_config(page_title="MT Valero - Reporte PDF Editable", layout="wide")

if 'hojas' not in st.session_state:
    st.session_state.hojas = []
if 'datos_pdf' not in st.session_state:
    st.session_state.datos_pdf = {"cliente": "", "fecha": ""}

st.title("📄 Generador de Reporte PDF Editable")

# --- BARRA LATERAL ---
with st.sidebar:
    st.header("1. Carga de Archivos")
    plantilla = st.file_uploader("Subir Reporte_MT_Final.pptx", type=["pptx"])
    archivo_pdf_valero = st.file_uploader("Subir Hoja Valero (PDF)", type=["pdf"])
    fotos_totales = st.file_uploader("Galería de WhatsApp", type=["jpg", "png", "jpeg"], accept_multiple_files=True)
    
    if archivo_pdf_valero and st.button("🔍 Extraer Datos del PDF"):
        with pdfplumber.open(archivo_pdf_valero) as pdf:
            texto = pdf.pages[0].extract_text()
            for l in texto.split('\n'):
                if "Cliente:" in l: st.session_state.datos_pdf["cliente"] = l.split("Cliente:")[1].strip()
                if "Fecha:" in l: st.session_state.datos_pdf["fecha"] = l.split("Fecha:")[1].strip()
        st.success("Datos extraídos.")

# --- CONSTRUCCIÓN DEL REPORTE ---
if plantilla:
    prs_base = Presentation(io.BytesIO(plantilla.getvalue()))
    nombres_layouts = [layout.name for layout in prs_base.slide_layouts]

    col_edit, col_prev = st.columns([1, 1])

    with col_edit:
        st.subheader("📝 Configurar Diapositiva")
        with st.form("editor_form", clear_on_submit=True):
            diseno = st.selectbox("Diseño de tu Plantilla:", nombres_layouts)
            texto_tabla = st.text_area("Texto para la Tabla (Times New Roman 12):")
            
            st.write("🖼️ **Selecciona fotos:**")
            fotos_seleccionadas = []
            if fotos_totales:
                c_f = st.columns(3)
                for i, f in enumerate(fotos_totales):
                    with c_f[i % 3]:
                        st.image(f, width=100)
                        if st.checkbox("Incluir", key=f"sel_{f.name}_{len(st.session_state.hojas)}"):
                            fotos_seleccionadas.append(f)
            
            if st.form_submit_button("➕ GUARDAR HOJA"):
                st.session_state.hojas.append({
                    "layout_idx": nombres_layouts.index(diseno),
                    "texto": texto_tabla,
                    "fotos": fotos_seleccionadas
                })
                st.rerun()

    with col_prev:
        st.subheader("👁️ Vista Previa")
        if st.session_state.hojas:
            idx = st.number_input("Ver Hoja #:", min_value=1, max_value=len(st.session_state.hojas)) - 1
            h = st.session_state.hojas[idx]
            st.info(f"Hoja guardada con {len(h['fotos'])} fotos.")
            if st.button("🗑️ Eliminar esta hoja"):
                st.session_state.hojas.pop(idx)
                st.rerun()

    # --- BOTONES DE DESCARGA ---
    st.divider()
    if st.session_state.hojas:
        # Generar el PPTX en memoria primero
        prs_final = Presentation(io.BytesIO(plantilla.getvalue()))
        for i in range(len(prs_final.slides)-1, -1, -1):
            rId = prs_final.slides._sldIdLst[i].rId
            prs_final.part.drop_rel(rId)
            del prs_final.slides._sldIdLst[i]

        for h in st.session_state.hojas:
            slide = prs_final.slides.add_slide(prs_final.slide_layouts[h['layout_idx']])
            img_count = 0
            for shape in slide.shapes:
                # Inyectar en Tabla o Cuadro de Texto
                if shape.has_table:
                    for row in shape.table.rows:
                        for cell in row.cells:
                            if cell.text == "" or "DESCRIPCIÓN" in cell.text.upper():
                                cell.text = h['texto']
                                for p in cell.text_frame.paragraphs:
                                    for run in p.runs:
                                        run.font.name = 'Times New Roman'
                                        run.font.size = Pt(12)
                elif shape.is_placeholder and shape.placeholder_format.type in [2, 7]:
                    shape.text = h['texto']
                    for p in shape.text_frame.paragraphs:
                        for r in p.runs:
                            r.font.name = 'Times New Roman'
                            r.font.size = Pt(12)
                # Inyectar Fotos en Placeholders
                elif shape.is_placeholder and (shape.placeholder_format.type == 18 or "PICTURE" in shape.name.upper()):
                    if img_count < len(h['fotos']):
                        shape.insert_picture(io.BytesIO(h['fotos'][img_count].read()))
                        h['fotos'][img_count].seek(0)
                        img_count += 1

        output_pptx = io.BytesIO()
        prs_final.save(output_pptx)
        
        col1, col2 = st.columns(2)
        
        with col1:
            st.download_button("📥 DESCARGAR PPTX", output_pptx.getvalue(), "Reporte.pptx")
        
        with col2:
            if st.button("🚀 PREPARAR PDF EDITABLE"):
                # Guardamos temporalmente para convertir
                with open("temp.pptx", "wb") as f:
                    f.write(output_pptx.getvalue())
                
                try:
                    # Comando para convertir a PDF manteniendo el diseño
                    subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "temp.pptx"], check=True)
                    
                    with open("temp.pdf", "rb") as f:
                        st.download_button("✅ DESCARGAR PDF FINAL", f.read(), "Reporte_MT_Final.pdf", "application/pdf")
                    
                    # Limpieza
                    os.remove("temp.pptx")
                    os.remove("temp.pdf")
                except Exception as e:
                    st.error("Error al generar PDF. Asegúrate de haber instalado LibreOffice en la terminal.")
                    