import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from docx import Document
from io import BytesIO

# Importación robusta
try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI - Valoración", layout="wide")

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading(f'Análisis Estratégico: {empresa}', 0)
    doc.add_paragraph(f'Nicho Analizado: {nicho}')
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto'
    hdr_cells[1].text = 'Puntaje de Ajuste (Match)'
    
    for prod, score in productos_data.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(prod)
        row_cells[1].text = f"{score:.2f}"
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.title("🎯 Sistema de Valoración de Portafolio")

    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        
        # LIMPIEZA AUTOMÁTICA DE ESPACIOS
        df.columns = df.columns.str.strip()
        
        # Mapeo según tus columnas detectadas y simplificadas
        cols_necesarias = ['Empresa', 'Nicho', 'Producto', 'Factor', 'Peso', 'Calificacion']
        
        # Verificación de existencia
        for col in cols_necesarias:
            if col not in df.columns:
                st.error(f"No se encuentra la columna: '{col}'")
                st.info(f"Columnas detectadas en tu Sheet: {list(df.columns)}")
                st.stop()

        df = df.dropna(subset=['Empresa', 'Nicho'])

    except Exception as e:
        st.error(f"Error crítico de conexión: {e}")
        st.stop()

    # Filtros en Cascada
    c1, c2 = st.columns(2)
    with c1:
        emp_sel = st.selectbox("1. Seleccione Empresa", df['Empresa'].unique())
        df_e = df[df['Empresa'] == emp_sel]
    with c2:
        nicho_sel = st.selectbox("2. Seleccione Nicho", df_e['Nicho'].unique())
        df_n = df_e[df_e['Nicho'] == nicho_sel]

    st.divider()

    # Selección de Productos
    prods_disp = df_n['Producto'].unique()
    prods_sel = st.multiselect("3. Seleccione Productos para comparar", prods_disp)

    if prods_sel:
        fig = go.Figure()
        res_dict = {}

        for p in prods_sel:
            d_p = df_n[df_n['Producto'] == p]
            
            # Gráfico de Radar
            fig.add_trace(go.Scatterpolar(
                r=d_p['Calificacion'],
                theta=d_p['Factor'],
                fill='toself',
                name=p
            ))
            
            # Cálculo de ajuste (Ponderación)
            # Aseguramos que Peso y Calificacion sean números
            c = pd.to_numeric(d_p['Calificacion'], errors='coerce').fillna(0)
            w = pd.to_numeric(d_p['Peso'], errors='coerce').fillna(0)
            
            # El score es la suma de (Calificacion * Peso)
            # Si el peso viene en formato 0.1, se multiplica directo. Si viene como 10, se divide por 100.
            if w.max() > 1:
                score_ponderado = (c * w / 100).sum()
            else:
                score_ponderado = (c * w).sum()
                
            res_dict[p] = score_ponderado

        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            title=f"Ajuste de Portafolio para: {nicho_sel}"
        )
        st.plotly_chart(fig, use_container_width=True)

        # Métricas visuales
        cols = st.columns(len(prods_sel))
        for i, (name, val) in enumerate(res_dict.items()):
            cols[i].metric(name, f"{val:.2f} / 5.0")

        st.divider()
        
        # Botón para Reporte Word
        word_report = generate_docx(emp_sel, nicho_sel, res_dict)
        st.download_button(
            label="📄 Descargar Reporte Word",
            data=word_report,
            file_name=f"Valoracion_{emp_sel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.info("Seleccione productos para visualizar el análisis de capacidades.")

if __name__ == "__main__":
    main()