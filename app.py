import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from docx import Document
from io import BytesIO

# Manejo de importación flexible para evitar errores de versión
try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI", layout="wide")

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading('Informe de Ajuste Estratégico', 0)
    doc.add_paragraph(f'Empresa Cliente: {empresa}')
    doc.add_paragraph(f'Nicho de Mercado: {nicho}')
    doc.add_paragraph("-" * 30)
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto / Servicio'
    hdr_cells[1].text = 'Índice de Match (1.0 - 5.0)'
    
    for prod, score in productos_data.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(prod)
        row_cells[1].text = f"{score:.2f}"
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.title("🎯 Valoración de Nicho y Capacidades")
    
    # Conexión a Google Sheets
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        # ttl=0 asegura que si cambias el Excel, la app se actualice al refrescar
        df = conn.read(ttl=0)
        
        # Limpieza de nombres de columnas (elimina espacios accidentales)
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=['Empresa', 'Nicho / Segmento'])
    except Exception as e:
        st.error(f"Error de conexión o lectura: {e}")
        st.info("Asegúrate de que el link en Secrets sea correcto y el Sheet sea público.")
        st.stop()

    # Filtros Dinámicos
    col1, col2 = st.columns(2)
    with col1:
        emp_list = df['Empresa'].unique()
        emp_sel = st.selectbox("Seleccione la Empresa", emp_list)
        df_e = df[df['Empresa'] == emp_sel]
    with col2:
        nicho_list = df_e['Nicho / Segmento'].unique()
        nicho_sel = st.selectbox("Seleccione el Nicho", nicho_list)
        df_n = df_e[df_e['Nicho / Segmento'] == nicho_sel]

    st.divider()

    # Selección de Productos
    prods = df_n['Producto / Servicio'].unique()
    sel_prods = st.multiselect("Seleccione Productos/Servicios del Portafolio", prods)

    if sel_prods:
        fig = go.Figure()
        res_scores = {}

        for p in sel_prods:
            d_p = df_n[df_n['Producto / Servicio'] == p]
            
            # Radar
            fig.add_trace(go.Scatterpolar(
                r=d_p['Calificación (1-5)'],
                theta=d_p['Factor de Éxito'],
                fill='toself',
                name=p
            ))
            
            # Cálculo Ponderado
            c = pd.to_numeric(d_p['Calificación (1-5)'], errors='coerce').fillna(0)
            w = pd.to_numeric(d_p['Peso (%)'], errors='coerce').fillna(0)
            res_scores[p] = (c * w / 100).sum()

        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            title=f"Comparativa de Ajuste para {nicho_sel}"
        )
        st.plotly_chart(fig, use_container_width=True)

        # Métricas
        m_cols = st.columns(len(sel_prods))
        for i, (name, val) in enumerate(res_scores.items()):
            m_cols[i].metric(name, f"{val:.2f} / 5.0")

        # Botón de Descarga
        st.divider()
        word_data = generate_docx(emp_sel, nicho_sel, res_scores)
        st.download_button(
            label="📄 Descargar Reporte en Word",
            data=word_data,
            file_name=f"Analisis_{emp_sel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.info("Seleccione productos para generar el análisis visual.")

if __name__ == "__main__":
    main()