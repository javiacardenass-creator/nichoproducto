import streamlit as st
from st_gsheets_connection import GSheetsConnection
import pandas as pd
import plotly.graph_objects as go
from docx import Document
from io import BytesIO

# Configuración de página
st.set_page_config(page_title="Consultoría BI - Valoración de Nicho", layout="wide")

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading('Reporte de Valoración de Portafolio', 0)
    doc.add_paragraph(f'Empresa: {empresa}')
    doc.add_paragraph(f'Nicho/Segmento: {nicho}')
    doc.add_paragraph("---")
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto / Servicio'
    hdr_cells[1].text = 'Índice de Ajuste (Score)'
    
    for prod, score in productos_data.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(prod)
        row_cells[1].text = f"{score:.2f}"
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

def main():
    st.title("🎯 Matriz de Ajuste: Producto vs. Nicho")
    
    # 1. Conexión y Limpieza de Datos
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        # ttl=0 para que siempre traiga datos frescos del Excel
        df = conn.read(ttl=0)
        
        # Limpieza CRÍTICA: Quitar espacios en los nombres de las columnas
        df.columns = df.columns.str.strip()
        df = df.dropna(subset=['Empresa', 'Nicho / Segmento'])
        
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        st.info("Revisa que el ID del Sheet en los Secrets sea correcto y que el archivo sea público (o compartido con el correo de la cuenta de servicio).")
        st.stop()

    # 2. Selectores en cascada
    col1, col2 = st.columns(2)
    
    with col1:
        # Usamos nombres exactos según tu Excel
        lista_empresas = df['Empresa'].unique()
        empresa_sel = st.selectbox("🏢 Seleccione la Empresa", lista_empresas)
        df_empresa = df[df['Empresa'] == empresa_sel]

    with col2:
        lista_nichos = df_empresa['Nicho / Segmento'].unique()
        nicho_sel = st.selectbox("🎯 Seleccione el Nicho", lista_nichos)
        df_nicho = df_empresa[df_empresa['Nicho / Segmento'] == nicho_sel]

    # 3. Selección de Productos y Visualización
    st.divider()
    productos_disponibles = df_nicho['Producto / Servicio'].unique()
    productos_sel = st.multiselect("📦 Seleccione los Productos para comparar", productos_disponibles)

    if productos_sel:
        fig = go.Figure()
        scores_finales = {}

        for prod in productos_sel:
            d_prod = df_nicho[df_nicho['Producto / Servicio'] == prod]
            
            # Gráfico de Radar
            fig.add_trace(go.Scatterpolar(
                r=d_prod['Calificación (1-5)'],
                theta=d_prod['Factor de Éxito'],
                fill='toself',
                name=prod
            ))
            
            # Cálculo de Score Ponderado: (Calificación * Peso) / 100
            # Aseguramos que los datos sean numéricos
            calif = pd.to_numeric(d_prod['Calificación (1-5)'], errors='coerce').fillna(0)
            peso = pd.to_numeric(d_prod['Peso (%)'], errors='coerce').fillna(0)
            
            score_ponderado = (calif * peso / 100).sum()
            scores_finales[prod] = score_ponderado

        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            showlegend=True,
            title=f"Ajuste de Portafolio para {nicho_sel}"
        )
        
        st.plotly_chart(fig, use_container_width=True)

        # Mostrar métricas
        cols_m = st.columns(len(productos_sel))
        for i, (p, s) in enumerate(scores_finales.items()):
            cols_m[i].metric(p, f"{s:.2f} / 5.0")

        # Generar Reporte
        st.divider()
        reporte_word = generate_docx(empresa_sel, nicho_sel, scores_finales)
        st.download_button(
            label="💾 Descargar Informe Ejecutivo (Word)",
            data=reporte_word,
            file_name=f"Analisis_{empresa_sel}_{nicho_sel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    else:
        st.warning("Seleccione al menos un producto para generar el radar.")

if __name__ == "__main__":
    main()