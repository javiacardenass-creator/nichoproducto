import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import requests
from docx import Document
from io import BytesIO

# --- CONFIGURACIÓN ---
APPSCRIPT_URL = st.secrets.get("https://script.google.com/macros/s/AKfycbxI1AeGjMdgYQzT4jZKktxFUIa1xWN1rYfh3EsdkmrM-mmVT0UpzgVnHyQIzTFlz2214w/exec")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI - Valoración", layout="wide")

def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        # Asegurar tipos numéricos
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
        df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame(columns=['Empresa', 'Nicho', 'Producto', 'Factor', 'Peso', 'Calificacion'])

def generate_docx(empresa, nicho, df_resumen):
    doc = Document()
    doc.add_heading(f'Reporte de Valoración Estratégica: {empresa}', 0)
    doc.add_paragraph(f'Nicho/Segmento Analizado: {nicho}')
    
    doc.add_heading('Matriz de Resultados Ponderados', level=1)
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto / Servicio'
    hdr_cells[1].text = 'Puntaje Final (0-5)'
    
    for _, row in df_resumen.iterrows():
        row_cells = table.add_row().cells
        row_cells[0].text = str(row['Producto'])
        row_cells[1].text = f"{row['Puntaje Final']:.2f}"
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- LÓGICA DE INTERFAZ ---
st.title("🎯 Sistema de Valoración Estratégica BI")

df = get_data()

tab_val, tab_matriz, tab_admin = st.tabs(["📊 Radar de Análisis", "📋 Matriz de Valoración", "⚙️ Gestión de Datos"])

# --- PESTAÑA 1: RADAR ---
with tab_val:
    if not df.empty:
        c1, c2 = st.columns(2)
        with c1:
            emp_v = st.selectbox("Empresa", df['Empresa'].unique(), key="v_emp")
            df_e = df[df['Empresa'] == emp_v]
        with c2:
            nicho_v = st.selectbox("Nicho", df_e['Nicho'].unique(), key="v_nic")
            df_n = df_e[df_e['Nicho'] == nicho_v]

        prods_sel = st.multiselect("Comparar Productos", df_n['Producto'].unique(), default=df_n['Producto'].unique()[:2])

        if prods_sel:
            fig = go.Figure()
            for p in prods_sel:
                d_p = df_n[df_n['Producto'] == p]
                fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
            
            fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])), showlegend=True)
            st.plotly_chart(fig, use_container_width=True)

# --- PESTAÑA 2: MATRIZ VISUAL (CÁLCULOS) ---
with tab_matriz:
    st.header(f"Resultados Ponderados: {nicho_v}")
    
    if prods_sel:
        # 1. Cálculo de Suma Producto (Puntaje Final)
        # Agrupamos por producto y calculamos (Calificación * Peso)
        df_n['Ponderado'] = df_n['Calificacion'] * (df_n['Peso']/100 if df_n['Peso'].max() > 1 else df_n['Peso'])
        
        resumen = df_n[df_n['Producto'].isin(prods_sel)].groupby('Producto')['Ponderado'].sum().reset_index()
        resumen.columns = ['Producto', 'Puntaje Final']
        
        # 2. Promedio por Factor (Para ver el estado del nicho)
        promedio_factores = df_n[df_n['Producto'].isin(prods_sel)].groupby('Factor')['Calificacion'].mean().reset_index()
        
        col_res1, col_res2 = st.columns([1, 1])
        
        with col_res1:
            st.subheader("Puntaje Final por Producto")
            st.dataframe(resumen.style.background_gradient(cmap='Blues', subset=['Puntaje Final']).format({'Puntaje Final': "{:.2f}"}), use_container_width=True)
            
            st.download_button("💾 Generar Reporte Word", 
                               generate_docx(emp_v, nicho_v, resumen), 
                               file_name=f"Reporte_{emp_v}.docx")

        with col_res2:
            st.subheader("Promedio de Desempeño por Factor")
            fig_bar = px.bar(promedio_factores, x='Factor', y='Calificacion', 
                             title="Promedio General del Nicho",
                             color='Calificacion', color_continuous_scale='RdYlGn', range_y=[0,5])
            st.plotly_chart(fig_bar, use_container_width=True)

        st.divider()
        st.subheader("Vista Detallada de la Matriz")
        # Tabla pivote para ver Factores vs Productos
        pivot_df = df_n[df_n['Producto'].isin(prods_sel)].pivot(index='Factor', columns='Producto', values='Calificacion')
        st.table(pivot_df)

# --- PESTAÑA 3: GESTIÓN (IDEM ANTERIOR) ---
with tab_admin:
    st.header("🛠️ Configuración de Estructura")
    with st.form("form_gestion", clear_on_submit=True):
        c_a, c_b = st.columns(2)
        with c_a:
            op_emp = ["➕ Nueva..."] + list(df['Empresa'].unique())
            s_emp = st.selectbox("Empresa", op_emp)
            n_emp = st.text_input("Nombre Nueva Empresa") if s_emp == "➕ Nueva..." else s_emp
            
            op_nic = ["➕ Nuevo..."] + (list(df[df['Empresa']==s_emp]['Nicho'].unique()) if s_emp != "➕ Nueva..." else [])
            s_nic = st.selectbox("Nicho", op_nic)
            n_nic = st.text_input("Nombre Nuevo Nicho") if s_nic == "➕ Nuevo..." else s_nic
            
        with c_b:
            op_pro = ["➕ Nuevo..."] + (list(df[(df['Empresa']==s_emp)&(df['Nicho']==s_nic)]['Producto'].unique()) if s_nic != "➕ Nuevo..." else [])
            s_pro = st.selectbox("Producto", op_pro)
            n_pro = st.text_input("Nombre Nuevo Producto") if s_pro == "➕ Nuevo..." else s_pro
            
            n_fac = st.text_input("Nombre del Factor")
            f_pes = st.number_input("Peso (%)", 1, 100, 20)
            f_cal = st.slider("Calificación (1-5)", 1, 5, 3)

        if st.form_submit_button("🚀 Guardar"):
            final_e, final_n, final_p = (n_emp if n_emp else s_emp), (n_nic if n_nic else s_nic), (n_pro if n_pro else s_pro)
            if all([final_e, final_n, final_p, n_fac]):
                payload = {"empresa": final_e, "nicho": final_n, "producto": final_p, "factor": n_fac, "peso": f_pes, "calificacion": f_cal}
                res = requests.post(APPSCRIPT_URL, json=payload)
                if res.status_code == 200:
                    st.success("Guardado."); st.cache_data.clear()
                else: st.error("Error AppScript")