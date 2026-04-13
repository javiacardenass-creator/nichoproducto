import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import requests
from docx import Document
from docx.shared import Inches
from io import BytesIO
import matplotlib.pyplot as plt

# --- 1. CONFIGURACIÓN ---
APPSCRIPT_URL = st.secrets.get("APPSCRIPT_URL")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI Pro", layout="wide")

# --- 2. FUNCIONES DE DATOS ---
def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        # ttl=0 asegura que traiga datos nuevos cada vez que se refresca la app
        df = conn.read(ttl=0)
        if df is not None and not df.empty:
            df.columns = df.columns.str.strip()
            # Validar columnas numéricas
            cols_num = ['Peso', 'Calificacion']
            for col in cols_num:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)
            
            if 'Accionables' not in df.columns:
                df['Accionables'] = ""
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return pd.DataFrame()

# --- 3. PROCESAMIENTO ---
df = get_data()

if df.empty:
    st.warning("⚠️ No se detectaron datos. Revisa la conexión con Google Sheets.")
    st.stop()

# --- SIDEBAR (FILTROS) ---
with st.sidebar:
    st.header("🏢 Filtros Globales")
    # Filtro de Empresa
    lista_empresas = sorted(df['Empresa'].unique())
    emp_v = st.selectbox("Empresa Cliente", lista_empresas)
    
    # Filtro de Nicho (Depende de Empresa)
    df_emp = df[df['Empresa'] == emp_v]
    lista_nichos = sorted(df_emp['Nicho'].unique())
    nic_v = st.selectbox("Nicho Analizado", lista_nichos)
    
    # Filtro de Productos (Depende de Nicho)
    df_nic = df_emp[df_emp['Nicho'] == nic_v].copy()
    lista_prods = sorted(df_nic['Producto'].unique())
    prods_v = st.multiselect("Filtrar Productos", lista_prods, default=lista_prods)
    
    st.divider()
    logo_up = st.file_uploader("Subir Logo Cliente", type=['png','jpg'])
    if logo_up:
        st.image(logo_up, width=150)

# Filtrado final para visualizaciones
df_final = df_nic[df_nic['Producto'].isin(prods_v)].copy()

# Pestañas
t_admin, t_radar, t_matriz = st.tabs(["⚙️ Gestión de Datos", "📊 Gráfico Radar", "📋 Matriz de Comparación"])

# --- PESTAÑA 1: GESTIÓN DE DATOS ---
with t_admin:
    st.subheader("🛠️ Registro y Actualización")
    with st.form("form_registro", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            # Empresa
            e_opt = ["➕ Nueva..."] + list(df['Empresa'].unique())
            e_sel = st.selectbox("Seleccionar Empresa", e_opt)
            e_input = st.text_input("Nombre si es Nueva") if e_sel == "➕ Nueva..." else e_sel
            
            # Nicho
            n_base = df[df['Empresa'] == e_sel]['Nicho'].unique() if e_sel != "➕ Nueva..." else []
            n_opt = ["➕ Nuevo..."] + list(n_base)
            n_sel = st.selectbox("Seleccionar Nicho", n_opt)
            n_input = st.text_input("Nombre si es Nuevo") if n_sel == "➕ Nuevo..." else n_sel

        with col2:
            # Producto
            p_base = df[(df['Empresa'] == e_sel) & (df['Nicho'] == n_sel)]['Producto'].unique() if n_sel != "➕ Nuevo..." else []
            p_opt = ["➕ Nuevo..."] + list(p_base)
            p_sel = st.selectbox("Seleccionar Producto", p_opt)
            p_input = st.text_input("Nombre si es Nuevo") if p_sel == "➕ Nuevo..." else p_sel
            
            # Factor
            f_base = df_nic['Factor'].unique()
            f_opt = ["➕ Nuevo..."] + list(f_base)
            f_sel = st.selectbox("Seleccionar Factor", f_opt)
            f_input = st.text_input("Nombre si es Nuevo Factor") if f_sel == "➕ Nuevo..." else f_sel

        st.divider()
        c_a, c_b, c_c = st.columns(3)
        peso = c_a.number_input("Peso (%)", 1, 100, 20)
        calif = c_b.slider("Calificación", 1, 5, 3)
        
        accion = st.text_area("Accionables / Conclusiones")
        
        if st.form_submit_button("🚀 Guardar en Google Sheets"):
            if not APPSCRIPT_URL:
                st.error("Falta URL de AppScript en Secrets.")
            elif e_input and n_input and p_input and f_input:
                payload = {
                    "empresa": e_input, "nicho": n_input, "producto": p_input,
                    "factor": f_input, "peso": peso, "calificacion": calif,
                    "accionables": accion
                }
                res = requests.post(APPSCRIPT_URL, json=payload)
                if res.status_code == 200:
                    st.success("Guardado exitoso"); st.cache_data.clear()
                    st.rerun()
                else: st.error("Error al guardar.")

# --- PESTAÑA 2: RADAR ---
with t_radar:
    if prods_v:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_final[df_final['Producto'] == p]
            if not d_p.empty:
                fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])), title=f"Radar de Nicho: {nic_v}")
        st.plotly_chart(fig, use_container_width=True)

# --- PESTAÑA 3: MATRIZ DE COMPARACIÓN ---
with t_matriz:
    st.subheader(f"Matriz Comparativa: {nic_v}")
    
    if not df_final.empty:
        # Crear la tabla cruzada: Factores vs Productos
        matriz_pivot = df_final.pivot_table(
            index='Factor', 
            columns='Producto', 
            values='Calificacion', 
            aggfunc='mean'
        ).fillna(0)
        
        st.write("**Calificación por Factor y Producto**")
        st.dataframe(matriz_pivot.style.background_gradient(cmap='YlGn', axis=None).format("{:.1f}"), use_container_width=True)
        
        # Resumen de puntajes ponderados
        df_final['Ponderado'] = df_final['Calificacion'] * (df_final['Peso'] / 100)
        resumen = df_final.groupby('Producto')['Ponderado'].sum().reset_index()
        resumen.columns = ['Producto', 'Puntaje Total Ponderado']
        
        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            st.write("**Ranking Ponderado**")
            st.table(resumen.sort_values(by='Puntaje Total Ponderado', ascending=False))
        with c2:
            st.write("**Accionables Registrados**")
            for p in prods_v:
                accs = df_final[df_final['Producto'] == p]['Accionables'].unique()
                with st.expander(f"Acciones para {p}"):
                    for a in accs:
                        if str(a).strip() and str(a) != 'nan': st.info(a)
    else:
        st.info("Selecciona productos para generar la matriz.")