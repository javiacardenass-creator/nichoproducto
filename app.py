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
# Asegúrate de tener esta URL en Settings > Secrets de Streamlit
APPSCRIPT_URL = st.secrets.get("https://script.google.com/macros/s/AKfycbxI1AeGjMdgYQzT4jZKktxFUIa1xWN1rYfh3EsdkmrM-mmVT0UpzgVnHyQIzTFlz2214w/exec")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    try:
        from streamlit_gsheets import GSheetsConnection
    except ImportError:
        GSheetsConnection = None

st.set_page_config(page_title="Consultoría BI Pro", layout="wide")

# --- FUNCIONES DE DATOS ---
def get_data():
    if GSheetsConnection is None:
        return pd.DataFrame()
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
        df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
        return df
    except:
        return pd.DataFrame()

# Manejo de Logos
if 'logos' not in st.session_state:
    st.session_state['logos'] = {}

def generate_report_docx(empresa, nicho, df_n, prods_sel, resumen_df):
    doc = Document()
    
    # Logo si existe
    if empresa in st.session_state['logos']:
        img_data = BytesIO(st.session_state['logos'][empresa])
        doc.add_picture(img_data, width=Inches(1.2))
    
    doc.add_heading(f'Informe de Valoración: {empresa}', 0)
    doc.add_paragraph(f'Nicho: {nicho}')

    # Gráfico de Radar para Word
    doc.add_heading('Análisis de Capacidades', level=1)
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

    # Tabla de Resultados
    doc.add_heading('Matriz de Resultados Ponderados', level=1)
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

# --- INTERFAZ ---
df = get_data()
if df.empty:
    st.error("No hay datos o la conexión falló.")
    st.stop()

# Sidebar
with st.sidebar:
    st.header("🏢 Configuración Cliente")
    emp_list = sorted(df['Empresa'].unique())
    emp_v = st.selectbox("Seleccionar Empresa", emp_list)
    
    logo_up = st.file_uploader("Cargar Logo", type=['png','jpg'])
    if logo_up:
        st.session_state['logos'][emp_v] = logo_up.getvalue()
    if emp_v in st.session_state['logos']:
        st.image(st.session_state['logos'][emp_v], width=100)
    
    st.divider()
    nicho_list = sorted(df[df['Empresa']==emp_v]['Nicho'].unique())
    nic_v = st.selectbox("Nicho", nicho_list)
    df_n = df[(df['Empresa']==emp_v) & (df['Nicho']==nic_v)]
    
    prod_list = sorted(df_n['Producto'].unique())
    prods_v = st.multiselect("Productos", prod_list, default=prod_list)

# Tabs
t1, t2, t3 = st.tabs(["📊 Gráfico Radar", "📋 Matriz Resumen", "⚙️ Gestión de Datos"])

# Cálculos
ajuste = 100 if df_n['Peso'].max() > 1.1 else 1
df_n['Ponderado'] = df_n['Calificacion'] * (df_n['Peso'] / ajuste)
res_df = df_n[df_n['Producto'].isin(prods_v)].groupby('Producto')['Ponderado'].sum().reset_index()
res_df.columns = ['Producto', 'Puntaje Final']

with t1:
    if prods_v:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_n[df_n['Producto']==p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        st.plotly_chart(fig, use_container_width=True)

with t2:
    st.header(f"Resultados para {nic_v}")
    c1, c2 = st.columns([2,1])
    with c1:
        st.dataframe(res_df.style.background_gradient(cmap='Greens').format({'Puntaje Final': "{:.2f}"}), use_container_width=True)
    with c2:
        st.write("**Reporte Word**")
        st.download_button("📂 Descargar Informe", 
                           generate_report_docx(emp_v, nic_v, df_n, prods_v, res_df),
                           f"Reporte_{emp_v}.docx")
    
    st.divider()
    st.write("**Matriz Detallada**")
    pivot = df_n[df_n['Producto'].isin(prods_v)].pivot_table(index='Factor', columns='Producto', values='Calificacion')
    st.table(pivot)

with t3:
    st.header("🛠️ Registro de Información")
    with st.form("f_admin", clear_on_submit=True):
        col1, col2 = st.columns(2)
        with col1:
            e_opt = ["➕ Crear Nueva..."] + list(df['Empresa'].unique())
            e_s = st.selectbox("Empresa", e_opt)
            e_final = st.text_input("Nombre Nueva Empresa") if e_s == "➕ Crear Nueva..." else e_s
            
            n_opt = ["➕ Crear Nuevo..."] + (list(df[df['Empresa']==e_s]['Nicho'].unique()) if e_s != "➕ Crear Nueva..." else [])
            n_s = st.selectbox("Nicho", n_opt)
            n_final = st.text_input("Nombre Nuevo Nicho") if n_s == "➕ Crear Nuevo..." else n_s
            
        with col2:
            p_opt = ["➕ Crear Nuevo..."] + (list(df[(df['Empresa']==e_s)&(df['Nicho']==n_s)]['Producto'].unique()) if n_s != "➕ Crear Nuevo..." else [])
            p_s = st.selectbox("Producto", p_opt)
            p_final = st.text_input("Nombre Nuevo Producto") if p_s == "➕ Crear Nuevo..." else p_s
            
            f_new = st.text_input("Factor")
            w_new = st.number_input("Peso (%)", 1, 100, 20)
            c_new = st.slider("Calificación", 1, 5, 3)
            
        if st.form_submit_button("🚀 Guardar"):
            if e_final and n_final and p_final and f_new:
                payload = {"empresa": e_final, "nicho": n_final, "producto": p_final, "factor": f_new, "peso": w_new, "calificacion": c_new}
                res = requests.post(APPSCRIPT_URL, json=payload)
                if res.status_code == 200:
                    st.success("Guardado"); st.cache_data.clear()
                else: st.error("Error AppScript")