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
    try:
        from streamlit_gsheets import GSheetsConnection
    except ImportError:
        GSheetsConnection = None

st.set_page_config(page_title="Consultoría BI Pro", layout="wide")

# --- FUNCIONES DE APOYO ---
def get_data():
    if GSheetsConnection is None: return pd.DataFrame()
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
        df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame()

if 'logos' not in st.session_state:
    st.session_state['logos'] = {}

def generate_report_docx(empresa, nicho, df_n, prods_sel, resumen_df):
    doc = Document()
    if empresa in st.session_state['logos']:
        img_data = BytesIO(st.session_state['logos'][empresa])
        doc.add_picture(img_data, width=Inches(1.2))
    
    doc.add_heading(f'Informe de Valoración: {empresa}', 0)
    doc.add_paragraph(f'Nicho: {nicho}')

    # Radar para Word
    plt.figure(figsize=(6, 6))
    ax = plt.subplot(111, polar=True)
    for p in prods_sel:
        d_p = df_n[df_n['Producto'] == p]
        val = d_p['Calificacion'].tolist()
        val += val[:1]
        angles = [n / float(len(d_p['Factor'])) * 2 * 3.14159 for n in range(len(d_p['Factor']))]
        angles += angles[:1]
        ax.plot(angles, val, linewidth=1, label=p)
        ax.fill(angles, val, alpha=0.1)
        plt.xticks(angles[:-1], d_p['Factor'].tolist())
    plt.legend(loc='upper right', bbox_to_anchor=(0.1, 0.1))
    
    tmp_img = BytesIO()
    plt.savefig(tmp_img, format='png')
    doc.add_picture(tmp_img, width=Inches(5))
    plt.close()

    # Tabla Resumen
    t = doc.add_table(rows=1, cols=2)
    t.style = 'Table Grid'
    t.rows[0].cells[0].text = 'Producto'
    t.rows[0].cells[1].text = 'Puntaje Final'
    for _, r in resumen_df.iterrows():
        row = t.add_row().cells
        row[0].text = str(r['Producto'])
        row[1].text = f"{r['Puntaje Final']:.2f}"

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

# --- LÓGICA PRINCIPAL ---
df = get_data()
if df.empty:
    st.error("Error de conexión. Verifica tus credenciales de Google Sheets.")
    st.stop()

# Sidebar para Logo y Filtros
with st.sidebar:
    st.header("🏢 Filtros y Logo")
    emp_list = sorted(df['Empresa'].unique())
    emp_v = st.selectbox("Empresa Cliente", emp_list)
    
    logo_up = st.file_uploader("Subir Logo (PNG/JPG)", type=['png','jpg'])
    if logo_up: st.session_state['logos'][emp_v] = logo_up.getvalue()
    if emp_v in st.session_state['logos']: st.image(st.session_state['logos'][emp_v], width=120)
    
    st.divider()
    nicho_list = sorted(df[df['Empresa']==emp_v]['Nicho'].unique())
    nic_v = st.selectbox("Nicho Analizado", nicho_list)
    df_n = df[(df['Empresa']==emp_v) & (df['Nicho']==nic_v)].copy()
    
    prod_list = sorted(df_n['Producto'].unique())
    prods_v = st.multiselect("Productos a comparar", prod_list, default=prod_list)

# Nuevo Orden de Pestañas
t_admin, t_radar, t_matriz = st.tabs(["⚙️ Gestión de Datos", "📊 Gráfico Radar", "📋 Matriz Resumen"])

# Cálculos Globales
ajuste = 100 if df_n['Peso'].max() > 1.1 else 1
df_n['Ponderado'] = df_n['Calificacion'] * (df_n['Peso'] / ajuste)
res_df = df_n[df_n['Producto'].isin(prods_v)].groupby('Producto')['Ponderado'].sum().reset_index()
res_df.columns = ['Producto', 'Puntaje Final']

# --- PESTAÑA 1: GESTIÓN DE DATOS ---
with t_admin:
    st.header("🛠️ Registro y Actualización")
    with st.form("form_gestion", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            e_opt = ["➕ Crear Nueva..."] + list(df['Empresa'].unique())
            e_s = st.selectbox("Empresa", e_opt)
            e_f = st.text_input("Nombre Nueva Empresa") if e_s == "➕ Crear Nueva..." else e_s
            
            n_opt = ["➕ Crear Nuevo..."] + (list(df[df['Empresa']==e_s]['Nicho'].unique()) if e_s != "➕ Crear Nueva..." else [])
            n_s = st.selectbox("Nicho", n_opt)
            n_f = st.text_input("Nombre Nuevo Nicho") if n_s == "➕ Crear Nuevo..." else n_s
            
        with c2:
            p_opt = ["➕ Crear Nuevo..."] + (list(df[(df['Empresa']==e_s)&(df['Nicho']==n_s)]['Producto'].unique()) if n_s != "➕ Crear Nuevo..." else [])
            p_s = st.selectbox("Producto", p_opt)
            p_f = st.text_input("Nombre Nuevo Producto") if p_s == "➕ Crear Nuevo..." else p_s
            
            f_n = st.text_input("Factor/Atributo")
            w_n = st.number_input("Peso (%)", 1, 100, 20)
            c_n = st.slider("Calificación (1-5)", 1, 5, 3)
            
        if st.form_submit_button("🚀 Guardar en Google Sheets"):
            if e_f and n_f and p_f and f_n:
                payload = {"empresa": e_f, "nicho": n_f, "producto": p_f, "factor": f_n, "peso": w_n, "calificacion": c_n}
                res = requests.post(APPSCRIPT_URL, json=payload)
                if res.status_code == 200:
                    st.success(f"Dato guardado: {f_n}")
                    st.cache_data.clear()
                else: st.error("Error en conexión con AppScript.")

# --- PESTAÑA 2: GRÁFICO RADAR ---
with t_radar:
    if prods_v:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_n[df_n['Producto']==p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])), title=f"Análisis Radar: {nic_v}")
        st.plotly_chart(fig, use_container_width=True)

# --- PESTAÑA 3: MATRIZ RESUMEN ---
with t_matriz:
    st.header(f"Resultados Ejecutivos: {nic_v}")
    col_mat1, col_mat2 = st.columns([1, 1])
    
    with col_mat1:
        st.subheader("Puntaje Ponderado Final")
        st.dataframe(res_df.style.background_gradient(cmap='Greens').format({'Puntaje Final': "{:.2f}"}), use_container_width=True)
        st.download_button("📂 Descargar Informe Word", 
                           generate_report_docx(emp_v, nic_v, df_n, prods_v, res_df),
                           f"Reporte_{emp_v}.docx")

    with col_mat2:
        st.subheader("Promedio de Desempeño por Factor")
        # Recuperamos el gráfico de barras perdido
        prom_factores = df_n[df_n['Producto'].isin(prods_v)].groupby('Factor')['Calificacion'].mean().reset_index()
        fig_bar = px.bar(prom_factores, x='Factor', y='Calificacion', 
                         color='Calificacion', color_continuous_scale='RdYlGn', range_y=[0,5])
        st.plotly_chart(fig_bar, use_container_width=True)

    st.divider()
    st.subheader("Matriz Comparativa de Calificaciones")
    pivot = df_n[df_n['Producto'].isin(prods_v)].pivot_table(index='Factor', columns='Producto', values='Calificacion')
    st.table(pivot.style.format("{:.1f}"))