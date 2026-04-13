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

st.set_page_config(page_title="Consultoría Estratégica BI", layout="wide")

# --- 2. CARGA DE DATOS ---
@st.cache_data(ttl=0)
def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        if df is not None and not df.empty:
            df.columns = df.columns.str.strip()
            df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
            df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
            if 'Accionables' not in df.columns:
                df['Accionables'] = ""
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error de conexión: {e}")
        return pd.DataFrame()

df_raw = get_data()

if df_raw.empty:
    st.warning("⚠️ No hay datos o la conexión falló. Revisa los Secrets.")
    st.stop()

# --- 3. BARRA LATERAL (CLIENTE Y NICHO) ---
with st.sidebar:
    st.header("🏢 Estructura de Proyecto")
    
    # SELECCIÓN O CREACIÓN DE CLIENTE
    e_opt = ["➕ Crear Nuevo..."] + sorted(list(df_raw['Empresa'].unique()))
    e_sel = st.selectbox("Cliente / Empresa", e_opt)
    e_final = st.text_input("Nombre de nueva Empresa", key="k_e_sidebar") if e_sel == "➕ Crear Nuevo..." else e_sel
    
    st.divider()
    
    # SELECCIÓN O CREACIÓN DE NICHO
    n_base = df_raw[df_raw['Empresa'] == e_sel]['Nicho'].unique() if e_sel != "➕ Crear Nuevo..." else []
    n_opt = ["➕ Crear Nuevo..."] + sorted(list(n_base))
    n_sel = st.selectbox("Nicho de Mercado", n_opt)
    n_final = st.text_input("Nombre de nuevo Nicho", key="k_n_sidebar") if n_sel == "➕ Crear Nuevo..." else n_sel

    st.divider()
    
    # FILTRO DE PRODUCTOS PARA ANÁLISIS
    df_contexto = df_raw[(df_raw['Empresa'] == e_final) & (df_raw['Nicho'] == n_final)]
    prods_disponibles = sorted(df_contexto['Producto'].unique())
    prods_v = st.multiselect("Productos a comparar en Reporte", prods_disponibles, default=prods_disponibles)

# --- 4. CUERPO PRINCIPAL ---
st.title(f"Análisis: {n_final}")
st.caption(f"Cliente: {e_final}")

tab_gestion, tab_radar, tab_matriz = st.tabs(["📝 Carga de Evaluación", "📊 Radar Competitivo", "📋 Matriz de Revisión"])

# PESTAÑA 1: GESTIÓN (PRODUCTOS Y FACTORES)
with tab_gestion:
    st.subheader("🛠️ Evaluación de Factores")
    with st.form("form_evaluacion", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            # PRODUCTO
            p_base = df_contexto['Producto'].unique()
            p_opt = ["➕ Nuevo Producto..."] + sorted(list(p_base))
            p_sel = st.selectbox("Seleccionar Producto", p_opt)
            p_final = st.text_input("Nombre del nuevo Producto", key="k_p_main") if p_sel == "➕ Nuevo Producto..." else p_sel

        with col2:
            # FACTOR
            f_base = df_contexto['Factor'].unique()
            f_opt = ["➕ Nuevo Factor..."] + sorted(list(f_base))
            f_sel = st.selectbox("Seleccionar Factor", f_opt)
            f_final = st.text_input("Nombre del nuevo Factor", key="k_f_main") if f_sel == "➕ Nuevo Factor..." else f_sel

        st.divider()
        c_a, c_b = st.columns(2)
        peso = c_a.number_input("Importancia / Peso (%)", 1, 100, 20)
        calif = c_b.slider("Calificación (1-5)", 1, 5, 3)
        
        accionables = st.text_area("🎯 Recomendaciones / Acciones Estratégicas")
        
        if st.form_submit_button("🚀 Guardar Evaluación"):
            if not APPSCRIPT_URL:
                st.error("Error: APPSCRIPT_URL no configurada.")
            elif e_final and n_final and p_final and f_final:
                payload = {
                    "empresa": e_final, "nicho": n_final, "producto": p_final,
                    "factor": f_final, "peso": peso, "calificacion": calif,
                    "accionables": accionables
                }
                res = requests.post(APPSCRIPT_URL, json=payload)
                if res.status_code == 200:
                    st.success("Dato guardado correctamente.")
                    st.cache_data.clear()
                    st.rerun()
                else: st.error("Fallo al conectar con Google Sheets.")

# PESTAÑA 2: RADAR
with tab_radar:
    df_radar = df_contexto[df_contexto['Producto'].isin(prods_v)]
    if not df_radar.empty:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_radar[df_radar['Producto'] == p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Selecciona productos en la barra lateral para generar el Radar.")

# PESTAÑA 3: MATRIZ DE REVISIÓN
with tab_matriz:
    df_matriz = df_contexto[df_contexto['Producto'].isin(prods_v)]
    if not df_matriz.empty:
        # MATRIZ CRUZADA REAL: FACTOR VS PRODUCTO
        pivot = df_matriz.pivot_table(
            index='Factor', 
            columns='Producto', 
            values='Calificacion', 
            aggfunc='mean'
        ).fillna(0)
        
        st.write("### Comparativa de Calificaciones por Factor")
        st.dataframe(pivot.style.background_gradient(cmap='RdYlGn', axis=None).format("{:.1f}"), use_container_width=True)
        
        st.divider()
        # RANKING PONDERADO
        df_matriz['Ponderado'] = df_matriz['Calificacion'] * (df_matriz['Peso'] / 100)
        resumen = df_matriz.groupby('Producto')['Ponderado'].sum().reset_index()
        resumen.columns = ['Producto', 'Puntaje Final']
        
        c1, c2 = st.columns([1, 2])
        with c1:
            st.write("#### Ranking Final")
            st.table(resumen.sort_values(by='Puntaje Final', ascending=False))
        with c2:
            st.write("#### Acciones por Producto")
            for p in prods_v:
                accs = df_matriz[df_matriz['Producto'] == p]['Accionables'].unique()
                with st.expander(f"Plan de acción: {p}"):
                    for a in accs:
                        if str(a).strip() and str(a) != 'nan': st.info(a)
    else:
        st.info("No hay datos suficientes para la matriz.")