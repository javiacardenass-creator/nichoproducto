import streamlit as st
from streamlit_gsheets import GSheetsConnection
import pandas as pd
import plotly.graph_objects as go
from docx import Document
from io import BytesIO

# Configuración de página
st.set_page_config(page_title="Consultoría Estratégica BI", layout="wide")

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading(f'Reporte de Valoración Estratégica: {empresa}', 0)
    doc.add_paragraph(f'Nicho Analizado: {nicho}')
    
    for prod, score in productos_data.items():
        doc.add_heading(f'Producto: {prod}', level=1)
        doc.add_paragraph(f'Índice de Ajuste Final: {score:.2f} / 5.0')
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.sidebar.title("Configuración")
    
    # Conexión a Google Sheets
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read()
    except Exception as e:
        st.error("Error al conectar con Google Sheets. Revisa las credenciales.")
        st.stop()

    st.title("🎯 Valoración de Nicho vs. Portafolio")
    st.markdown("---")

    # 1. Filtros de Cascada
    col_a, col_b = st.columns(2)
    
    with col_a:
        empresa_sel = st.selectbox("1. Seleccione la Empresa", df['Empresa'].unique())
        df_empresa = df[df['Empresa'] == empresa_sel]

    with col_b:
        nicho_sel = st.selectbox("2. Seleccione el Nicho", df_empresa['Nicho / Segmento'].unique())
        df_nicho = df_empresa[df_empresa['Nicho / Segmento'] == nicho_sel]

    # 2. Selección de Productos
    productos_disponibles = df_nicho['Producto / Servicio'].unique()
    productos_sel = st.multiselect("3. Compare Productos/Servicios", productos_disponibles)

    if productos_sel:
        fig = go.Figure()
        resultados_reporte = {}

        for prod in productos_sel:
            d_prod = df_nicho[df_nicho['Producto / Servicio'] == prod]
            
            # Gráfico de Radar
            fig.add_trace(go.Scatterpolar(
                r=d_prod['Calificación (1-5)'],
                theta=d_prod['Factor de Éxito'],
                fill='toself',
                name=prod
            ))
            
            # Cálculo de Score Ponderado
            score = (d_prod['Calificación (1-5)'] * d_prod['Peso (%)'] / 100).sum()
            resultados_reporte[prod] = score

        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            height=600
        )
        
        st.plotly_chart(fig, use_container_width=True)

        # 3. Métricas y Reporte
        st.subheader("Análisis de Ajuste (Match)")
        cols_metrics = st.columns(len(productos_sel))
        
        for i, (prod, score) in enumerate(resultados_reporte.items()):
            cols_metrics[i].metric(label=prod, value=f"{score:.2f} / 5.0")

        # Botón para descargar reporte
        docx_file = generate_docx(empresa_sel, nicho_sel, resultados_reporte)
        st.download_button(
            label="📄 Descargar Reporte en Word",
            data=docx_file,
            file_name=f"Reporte_{empresa_sel}_{nicho_sel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.info("Seleccione uno o más productos para visualizar el radar de capacidades.")

if __name__ == "__main__":
    main()