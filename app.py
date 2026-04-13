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
APPSCRIPT_URL = st.secrets.get("https://script.google.com/macros/s/AKfycbzVLmOtsHnzL3IJXAcZgqovnjooghk5yG_5H3b7Dwhx9HNlzcpZXuhTq3HRXB9J9Cvfvg/exec", None)

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI Pro", layout="wide")

# --- FUNCIONES DE DATOS ---
def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
        df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
        if 'Accionables' not in df.columns: df['Accionables'] = ""
        return df
    except Exception as e:
        st.error(f"Error al conectar con Sheets: {e}")
        return pd.DataFrame()

if 'logos' not in st.session_state:
    st.session_state['logos'] = {}

# --- GENERADOR DE INFORME WORD ---
def generate_report_docx(empresa, nicho, df_n, prods_sel, resumen_df):
    doc = Document()
    if empresa in st.session_state['logos']:
        img_data = BytesIO(st.session_state['logos'][empresa])
        doc.add_picture(img_data, width=Inches(1.2))
    
    doc.add_heading(f'Informe Estratégico: {empresa}', 0)
    doc.add_paragraph(f'Nicho: {nicho}')

    # Gráfico Radar
    plt.figure(figsize=(6, 6))
    ax = plt.subplot(111, polar=True)
    for p in prods_sel:
        d_p = df_n[df_n['Producto'] == p]
        if not d_p.empty:
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
    doc.add_picture(tmp_img, width=Inches(4))
    plt.close()

    doc.add_heading('Plan de Acción Estratégico', level=1)
    for p in prods_sel:
        score = resumen_df[resumen_df['Producto'] == p]['Puntaje Final'].values[0]
        doc.add_heading(f'Producto: {p} (Score: {score:.2f})', level=2)
        acciones = df_n[df_n['Producto'] == p]['Accionables'].unique()
        for acc in acciones:
            if str(acc).strip() and str(acc) != "nan":
                doc.add_paragraph(f"• {acc}", style='List Bullet')

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

# --- INTERFAZ ---
df = get_data()
if df.empty:
    st.warning("Cargando base de datos...")
    st.stop()

with st.sidebar:
    st.header("🏢 Filtros y Logo")
    emp_v = st.selectbox("Empresa Cliente", sorted(df['Empresa'].unique()))
    logo_up = st.file_uploader("Logo", type=['png','jpg'])
    if logo_up: st.session_state['logos'][emp_v] = logo_up.getvalue()
    if emp_v in st.session_state['logos']: st.image(st.session_state['logos'][emp_v], width=100)
    
    st.divider()
    nicho_opts = sorted(df[df['Empresa']==emp_v]['Nicho'].unique())
    nic_v = st.selectbox("Nicho Analizado", nicho_opts)
    df_n = df[(df['Empresa']==emp_v) & (df['Nicho']==nic_v)].copy()
    prods_v = st.multiselect("Productos", sorted(df_n['Producto'].unique()), default=sorted(df_n['Producto'].unique()))

t_admin, t_radar, t_matriz = st.tabs(["⚙️ Gestión de Datos", "📊 Gráfico Radar", "📋 Matriz & Accionables"])

# Cálculos
ajuste = 100 if df_n['Peso'].max() > 1.1 else 1
df_n['Ponderado'] = df_n['Calificacion'] * (df_n['Peso'] / ajuste)
res_df = df_n[df_n['Producto'].isin(prods_v)].groupby('Producto')['Ponderado'].sum().reset_index()
res_df.columns = ['Producto', 'Puntaje Final']

# --- PESTAÑA 1: GESTIÓN (Factores Dinámicos) ---
with t_admin:
    st.header("🛠️ Registro de Factores y Acciones")
    with st.form("form_gestion", clear_on_submit=True):
        c1, c2 = st.columns(2)
        with c1:
            e_s = st.selectbox("Empresa", ["➕ Crear Nueva..."] + list(df['Empresa'].unique()))
            e_f = st.text_input("Nombre Nueva Empresa") if e_s == "➕ Crear Nueva..." else e_s
            
            n_opts = ["➕ Crear Nuevo..."] + (list(df[df['Empresa']==e_s]['Nicho'].unique()) if e_s != "➕ Crear Nueva..." else [])
            n_s = st.selectbox("Nicho", n_opts)
            n_f = st.text_input("Nombre Nuevo Nicho") if n_s == "➕ Crear Nuevo..." else n_s

        with c2:
            p_opts = ["➕ Crear Nuevo..."] + (list(df[(df['Empresa']==e_s)&(df['Nicho']==n_s)]['Producto'].unique()) if n_s != "➕ Crear Nuevo..." else [])
            p_s = st.selectbox("Producto", p_opts)
            p_f = st.text_input("Nombre Nuevo Producto") if p_s == "➕ Crear Nuevo..." else p_s
            
            # NUEVO: Selector dinámico de Factores
            f_opts = ["➕ Crear Nuevo Factor..."] + (list(df[(df['Empresa']==e_s)&(df['Nicho']==n_s)&(df['Producto']==p_s)]['Factor'].unique()) if p_s != "➕ Crear Nuevo..." else [])
            f_s = st.selectbox("Factor de Éxito", f_opts)
            f_f = st.text_input("Nombre Nuevo Factor") if f_s == "➕ Crear Nuevo Factor..." else f_s
            
            sub_c1, sub_c2 = st.columns(2)
            w_n = sub_c1.number_input("Peso (%)", 1, 100, 20)
            c_n = sub_c2.slider("Calificación", 1, 5, 3)
        
        acc_n = st.text_area("🎯 Accionables / Conclusiones")
        
        if st.form_submit_button("🚀 Guardar Todo"):
            if not APPSCRIPT_URL:
                st.error("URL de AppScript no configurada en Secrets.")
            elif e_f and n_f and p_f and f_f:
                payload = {"empresa": e_f, "nicho": n_f, "producto": p_f, "factor": f_f, "peso": w_n, "calificacion": c_n, "accionables": acc_n}
                try:
                    res = requests.post(APPSCRIPT_URL, json=payload, timeout=10)
                    if res.status_code == 200:
                        st.success(f"✅ Guardado: {f_f} para {p_f}")
                        st.cache_data.clear()
                    else: st.error("Error en AppScript.")
                except Exception as e: st.error(f"Error: {e}")

# --- PESTAÑA 2: RADAR ---
with t_radar:
    if prods_v:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_n[df_n['Producto']==p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        st.plotly_chart(fig, use_container_width=True)

# --- PESTAÑA 3: MATRIZ & ACCIONABLES ---
with t_matriz:
    st.header(f"Resultados Ejecutivos: {nic_v}")
    col_mat1, col_mat2 = st.columns([1, 1])
    
    with col_mat1:
        st.subheader("Análisis Ponderado")
        st.dataframe(res_df.style.background_gradient(cmap='Greens').format({'Puntaje Final': "{:.2f}"}), use_container_width=True)
        st.download_button("📂 Descargar Informe Word", generate_report_docx(emp_v, nic_v, df_n, prods_v, res_df), f"Reporte_{emp_v}.docx")

    with col_mat2:
        st.subheader("Desempeño por Factor")
        prom_f = df_n[df_n['Producto'].isin(prods_v)].groupby('Factor')['Calificacion'].mean().reset_index()
        st.plotly_chart(px.bar(prom_f, x='Factor', y='Calificacion', color='Calificacion', range_y=[0,5]), use_container_width=True)

    st.divider()
    st.subheader("🎯 Plan de Acción por Producto")
    for p in prods_v:
        with st.expander(f"Accionables para {p}"):
            acciones = df_n[df_n['Producto'] == p]['Accionables'].unique()
            for a in acciones:
                if str(a).strip() and str(a) != "nan":
                    st.info(a)