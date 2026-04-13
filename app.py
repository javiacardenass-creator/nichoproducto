import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import requests
from docx import Document
from docx.shared import Inches
from io import BytesIO
import matplotlib.pyplot as plt

# --- CONFIGURACIÓN ---
APPSCRIPT_URL = st.secrets.get("https://script.google.com/macros/s/AKfycbxI1AeGjMdgYQzT4jZKktxFUIa1xWN1rYfh3EsdkmrM-mmVT0UpzgVnHyQIzTFlz2214w/exec")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI Pro", layout="wide")

# --- FUNCIONES DE APOYO ---
def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
        df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame()

# Gestión de Logos en memoria de sesión
def save_logo(empresa, file):
    if file:
        st.session_state[f"logo_{empresa}"] = file.getvalue()

def get_logo(empresa):
    return st.session_state.get(f"logo_{empresa}", None)

# --- GENERACIÓN DE INFORME WORD PRO ---
def generate_advanced_docx(empresa, nicho, df_n, prods_sel, resumen_df):
    doc = Document()
    
    # 1. Insertar Logo si existe
    logo_data = get_logo(empresa)
    if logo_data:
        image_stream = BytesIO(logo_data)
        doc.add_picture(image_stream, width=Inches(1.5))
    
    doc.add_heading(f'Informe Estratégico de Valoración', 0)
    doc.add_heading(f'Cliente: {empresa}', level=1)
    doc.add_paragraph(f'Nicho de Mercado: {nicho}')

    # 2. Insertar Gráfico de Radar (Generado con Matplotlib para Word)
    doc.add_heading('Análisis Visual de Capacidades (Radar)', level=2)
    
    plt.figure(figsize=(6, 6))
    ax = plt.subplot(111, polar=True)
    
    for p in prods_sel:
        d_p = df_n[df_n['Producto'] == p]
        factors = d_p['Factor'].tolist()
        values = d_p['Calificacion'].tolist()
        values += values[:1] # Cerrar el círculo
        angles = [n / float(len(factors)) * 2 * 3.14159 for n in range(len(factors))]
        angles += angles[:1]
        
        ax.plot(angles, values, linewidth=1, linestyle='solid', label=p)
        ax.fill(angles, values, alpha=0.1)
        plt.xticks(angles[:-1], factors)

    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))
    
    memfile = BytesIO()
    plt.savefig(memfile, format='png', bbox_inches='tight')
    doc.add_picture(memfile, width=Inches(5))
    plt.close()

    # 3. Matriz Resumen Valorada
    doc.add_heading('Matriz Resumen de Valoración por Producto', level=2)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto'
    hdr_cells[1].text = 'Puntaje Final (Ponderado)'
    
    for _, row in resumen_df.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Producto'])
        row_cells[1].text = f"{row['Puntaje Final']:.2f}"

    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- INTERFAZ ---
df = get_data()
if df.empty:
    st.error("Error al conectar con la base de datos.")
    st.stop()

# Sidebar para Logo y Selección
with st.sidebar:
    st.header("🏢 Identidad Corporativa")
    emp_sel = st.selectbox("Empresa Cliente", df['Empresa'].unique())
    
    # Carga de Logo
    logo_file = st.file_uploader(f"Subir logo de {emp_sel}", type=['png', 'jpg', 'jpeg'])
    if logo_file:
        save_logo(emp_sel, logo_file)
    
    current_logo = get_logo(emp_sel)
    if current_logo:
        st.image(current_logo, width=150)
    
    st.divider()
    df_e = df[df['Empresa'] == emp_sel]
    nicho_sel = st.selectbox("Nicho Analizado", df_e['Nicho'].unique())
    df_n = df_e[df_e['Nicho'] == nicho_sel]
    prods_sel = st.multiselect("Productos a comparar", df_n['Producto'].unique(), default=df_n['Producto'].unique())

# Tabs Principales
tab1, tab2, tab3 = st.tabs(["📊 Radar", "📋 Matriz Resumen", "⚙️ Gestión"])

# Cálculo de Matriz Resumen (Global para las pestañas)
ajuste = 100 if df_n['Peso'].max() > 1.1 else 1
df_n['Ponderado'] = df_n['Calificacion'] * (df_n['Peso'] / ajuste)
resumen_df = df_n[df_n['Producto'].isin(prods_sel)].groupby('Producto')['Ponderado'].sum().reset_index()
resumen_df.columns = ['Producto', 'Puntaje Final']

# --- PESTAÑA 1: RADAR ---
with tab1:
    if prods_sel:
        fig = go.Figure()
        for p in prods_sel:
            d_p = df_n[df_n['Producto'] == p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])), title=f"Capacidades en {nicho_sel}")
        st.plotly_chart(fig, use_container_width=True)

# --- PESTAÑA 2: MATRIZ RESUMEN ---
with tab2:
    st.subheader(f"Resumen Ejecutivo: {nicho_sel}")
    
    c_m1, c_m2 = st.columns([2, 1])
    
    with c_m1:
        st.write("**Matriz de Valoración Final**")
        try:
            st.dataframe(resumen_df.style.background_gradient(cmap='YlGn').format({'Puntaje Final': "{:.2f}"}), use_container_width=True)
        except:
            st.dataframe(resumen_df, use_container_width=True)

    with c_m2:
        st.write("**Reporte Profesional**")
        btn_word = st.download_button(
            label="📂 Descargar Informe Word (.docx)",
            data=generate_advanced_docx(emp_sel, nicho_sel, df_n, prods_sel, resumen_df),
            file_name=f"Informe_Estrategico_{emp_sel}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
        if btn_word:
            st.balloons()

    st.divider()
    st.write("**Detalle por Atributos (Factores)**")
    pivot_t = df_n[df_n['Producto'].isin(prods_sel)].pivot_table(index='Factor', columns='Producto', values='Calificacion')
    st.table(pivot_t)

# --- PESTAÑA 3: GESTIÓN (IGUAL A LA ANTERIOR) ---
with tab3:
    st.header("⚙️ Administración de Matriz")
    # ... (Aquí va tu código de formulario de gestión anterior)