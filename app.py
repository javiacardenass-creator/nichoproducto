import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import requests
from docx import Document
from io import BytesIO

# --- CONFIGURACIÓN DE CONEXIÓN ---
# Se recomienda usar st.secrets para la URL de AppScript
APPSCRIPT_URL = st.secrets.get("https://script.google.com/macros/s/AKfycbxI1AeGjMdgYQzT4jZKktxFUIa1xWN1rYfh3EsdkmrM-mmVT0UpzgVnHyQIzTFlz2214w/exec")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    try:
        from streamlit_gsheets import GSheetsConnection
    except ImportError:
        GSheetsConnection = None

st.set_page_config(page_title="Consultoría BI - Estrategia", layout="wide")

# --- FUNCIONES DE APOYO ---
def get_data():
    if GSheetsConnection is None:
        st.error("Librería de conexión no encontrada.")
        return pd.DataFrame()
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        # Limpieza y conversión numérica
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
        df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
        return df
    except Exception as e:
        st.error(f"Error al leer Google Sheets: {e}")
        return pd.DataFrame()

def generate_docx(empresa, nicho, df_resumen):
    doc = Document()
    doc.add_heading(f'Análisis de Valoración: {empresa}', 0)
    doc.add_paragraph(f'Nicho: {nicho}')
    
    doc.add_heading('Resumen de Puntuación Ponderada', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto / Servicio'
    hdr_cells[1].text = 'Puntaje Total (0-5)'
    
    for _, row in df_resumen.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Producto'])
        row_cells[1].text = f"{row['Puntaje Final']:.2f}"
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- LÓGICA PRINCIPAL ---
st.title("🎯 Sistema de Valoración Estratégica")

df = get_data()

tab_val, tab_matriz, tab_admin = st.tabs(["📊 Análisis Radar", "📋 Matriz de Resultados", "⚙️ Gestión de Datos"])

if df.empty:
    st.warning("La base de datos está vacía o no se pudo leer.")
    st.stop()

# Selectores globales para la sesión actual
with st.sidebar:
    st.header("Filtros de Análisis")
    emp_sel = st.selectbox("Empresa Cliente", df['Empresa'].unique())
    df_e = df[df['Empresa'] == emp_sel]
    
    nicho_sel = st.selectbox("Nicho / Segmento", df_e['Nicho'].unique())
    df_n = df_e[df_e['Nicho'] == nicho_sel]
    
    prods_sel = st.multiselect("Productos a Evaluar", df_n['Producto'].unique(), default=df_n['Producto'].unique()[:2] if len(df_n['Producto'].unique()) > 1 else df_n['Producto'].unique())

# --- PESTAÑA 1: RADAR ---
with tab_val:
    if prods_sel:
        fig = go.Figure()
        for p in prods_sel:
            d_p = df_n[df_n['Producto'] == p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        
        fig.update_layout(
            polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
            title=f"Comparativa de Capacidades en {nicho_sel}"
        )
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Selecciona productos en la barra lateral para ver el Radar.")

# --- PESTAÑA 2: MATRIZ VISUAL ---
with tab_matriz:
    if prods_sel:
        st.subheader(f"Análisis Ponderado - {emp_sel} ({nicho_sel})")
        
        # Cálculo de Suma Producto
        # Si el peso máximo es mayor a 1.1 asumimos que está en escala 1-100
        ajuste_peso = 100 if df_n['Peso'].max() > 1.1 else 1
        df_n['Ponderado'] = df_n['Calificacion'] * (df_n['Peso'] / ajuste_peso)
        
        resumen = df_n[df_n['Producto'].isin(prods_sel)].groupby('Producto')['Ponderado'].sum().reset_index()
        resumen.columns = ['Producto', 'Puntaje Final']
        
        c_res1, c_res2 = st.columns(2)
        
        with c_res1:
            st.write("**Puntaje Final Calculado**")
            try:
                st.dataframe(
                    resumen.style.background_gradient(cmap='Greens', subset=['Puntaje Final'])
                    .format({'Puntaje Final': "{:.2f}"}),
                    use_container_width=True
                )
            except:
                st.dataframe(resumen, use_container_width=True)
            
            st.download_button("💾 Exportar Resultados a Word", 
                               generate_docx(emp_sel, nicho_sel, resumen), 
                               f"Reporte_{emp_sel}.docx")

        with c_res2:
            st.write("**Promedio de Fortalezas del Nicho**")
            promedio_f = df_n[df_n['Producto'].isin(prods_sel)].groupby('Factor')['Calificacion'].mean().reset_index()
            fig_bar = px.bar(promedio_f, x='Factor', y='Calificacion', color='Calificacion', color_continuous_scale='Viridis', range_y=[0,5])
            st.plotly_chart(fig_bar, use_container_width=True)

        st.divider()
        st.write("**Matriz Detallada (Factores vs Productos)**")
        pivot = df_n[df_n['Producto'].isin(prods_sel)].pivot_table(index='Factor', columns='Producto', values='Calificacion', aggfunc='mean')
        st.table(pivot.style.format("{:.1f}"))

# --- PESTAÑA 3: GESTIÓN ---
with tab_admin:
    st.header("🛠️ Configuración de la Matriz")
    
    with st.form("form_admin", clear_on_submit=True):
        col_a, col_b = st.columns(2)
        
        with col_a:
            op_e = ["➕ Crear Nueva..."] + list(df['Empresa'].unique())
            s_e = st.selectbox("Seleccionar Empresa", op_e)
            n_e = st.text_input("Nombre de nueva empresa") if s_e == "➕ Crear Nueva..." else s_e
            
            op_n = ["➕ Crear Nuevo..."] + (list(df[df['Empresa'] == s_e]['Nicho'].unique()) if s_e != "➕ Crear Nueva..." else [])
            s_n = st.selectbox("Seleccionar Nicho", op_n)
            n_n = st.text_input("Nombre de nuevo nicho") if s_n == "➕ Crear Nuevo..." else s_n

        with col_b:
            op_p = ["➕ Crear Nuevo..."] + (list(df[(df['Empresa']==s_e)&(df['Nicho']==s_n)]['Producto'].unique()) if s_n != "➕ Crear Nuevo..." else [])
            s_p = st.selectbox("Seleccionar Producto", op_p)
            n_p = st.text_input("Nombre de nuevo producto") if s_p == "➕ Crear Nuevo..." else s_p
            
            f_nom = st.text_input("Nombre del Factor (Atributo)")
            f_pes = st.number_input("Peso (%)", 1, 100, 20)
            f_cal = st.slider("Calificación (1-5)", 1, 5, 3)

        if st.form_submit_button("🚀 Guardar y Actualizar"):
            final_e = n_e if n_e else s_e
            final_n = n_n if n_n else s_n
            final_p = n_p if n_p else s_p
            
            if all([final_e, final_n, final_p, f_nom]) and final_e != "➕ Crear Nueva...":
                payload = {
                    "empresa": final_e, "nicho": final_n, "producto": final_p,
                    "factor": f_nom, "peso": f_pes, "calificacion": f_cal
                }
                try:
                    res = requests.post(APPSCRIPT_URL, json=payload)
                    if res.status_code == 200:
                        st.success("✅ Guardado exitoso.")
                        st.cache_data.clear()
                    else: st.error("Error en AppScript.")
                except Exception as e: st.error(f"Error de red: {e}")
            else:
                st.warning("Completa todos los campos obligatorios.")