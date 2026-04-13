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

# --- 3. BARRA LATERAL (CLIENTE Y NICHO) ---
with st.sidebar:
    st.header("🏢 Estructura de Proyecto")
    
    # Selección de Empresa
    e_opt = ["➕ Crear Nuevo..."] + sorted(list(df_raw['Empresa'].unique())) if not df_raw.empty else ["➕ Crear Nuevo..."]
    e_sel = st.selectbox("Cliente / Empresa", e_opt)
    e_final = st.text_input("Nombre de nueva Empresa", key="k_e_sidebar") if e_sel == "➕ Crear Nuevo..." else e_sel
    
    st.divider()
    
    # Selección de Nicho
    n_base = df_raw[df_raw['Empresa'] == e_sel]['Nicho'].unique() if (not df_raw.empty and e_sel != "➕ Crear Nuevo...") else []
    n_opt = ["➕ Crear Nuevo..."] + sorted(list(n_base))
    n_sel = st.selectbox("Nicho de Mercado", n_opt)
    n_final = st.text_input("Nombre de nuevo Nicho", key="k_n_sidebar") if n_sel == "➕ Crear Nuevo..." else n_sel

    st.divider()
    
    # Filtro de Productos para Análisis
    df_contexto = df_raw[(df_raw['Empresa'] == e_final) & (df_raw['Nicho'] == n_final)] if not df_raw.empty else pd.DataFrame()
    prods_disponibles = sorted(df_contexto['Producto'].unique()) if not df_contexto.empty else []
    prods_v = st.multiselect("Productos a comparar en Reporte", prods_disponibles, default=prods_disponibles)

# --- 4. CUERPO PRINCIPAL ---
st.title(f"Análisis: {n_final}")
st.caption(f"Cliente: {e_final}")

tab_gestion, tab_radar, tab_matriz = st.tabs(["📝 Carga de Evaluación", "📊 Radar Competitivo", "📋 Matriz de Revisión"])

# PESTAÑA 1: GESTIÓN
with tab_gestion:
    st.subheader("🛠️ Evaluación de Factores")
    with st.form("form_evaluacion", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            # PRODUCTO: Siempre incluimos la opción de crear uno nuevo
            p_base = sorted(list(df_contexto['Producto'].unique())) if not df_contexto.empty else []
            p_opt = ["➕ Nuevo Producto..."] + p_base
            p_sel = st.selectbox("Seleccionar Producto", p_opt)
            p_final = st.text_input("Nombre del nuevo Producto", key="k_p_main") if p_sel == "➕ Nuevo Producto..." else p_sel

        with col2:
            # FACTOR: Siempre incluimos la opción de crear uno nuevo
            f_base = sorted(list(df_contexto['Factor'].unique())) if not df_contexto.empty else []
            f_opt = ["➕ Nuevo Factor..."] + f_base
            f_sel = st.selectbox("Seleccionar Factor", f_opt)
            f_final = st.text_input("Nombre del nuevo Factor", key="k_f_main") if f_sel == "➕ Nuevo Factor..." else f_sel

        st.divider()
        c_a, c_b = st.columns(2)
        peso = c_a.number_input("Importancia / Peso (%)", 1, 100, 20)
        calif = c_b.slider("Calificación (1-5)", 1, 5, 3)
        
        accionables = st.text_area("🎯 Recomendaciones / Acciones Estratégicas")
        
        if st.form_submit_button("🚀 Guardar Evaluación"):
            if not APPSCRIPT_URL:
                st.error("Error: APPSCRIPT_URL no configurada en Secrets.")
            elif e_final and n_final and p_final and f_final:
                payload = {
                    "empresa": e_final, "nicho": n_final, "producto": p_final,
                    "factor": f_final, "peso": peso, "calificacion": calif,
                    "accionables": accionables
                }
                try:
                    res = requests.post(APPSCRIPT_URL, json=payload, timeout=10)
                    if res.status_code == 200:
                        st.success(f"¡{p_final} guardado exitosamente!")
                        st.cache_data.clear()
                        st.rerun()
                    else:
                        st.error("Error al guardar en la base de datos.")
                except Exception as e:
                    st.error(f"Error de conexión: {e}")
            else:
                st.warning("Por favor completa todos los campos (Empresa, Nicho, Producto y Factor).")

# PESTAÑA 2: RADAR
with tab_radar:
    if not df_contexto.empty and prods_v:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_contexto[df_contexto['Producto'] == p]
            if not d_p.empty:
                fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos suficientes o productos seleccionados para el Radar.")

# PESTAÑA 3: MATRIZ
with tab_matriz:
    if not df_contexto.empty and prods_v:
        df_matriz = df_contexto[df_contexto['Producto'].isin(prods_v)]
        if not df_matriz.empty:
            pivot = df_matriz.pivot_table(
                index='Factor', 
                columns='Producto', 
                values='Calificacion', 
                aggfunc='mean'
            ).fillna(0)
            
            st.write("### Comparativa de Calificaciones por Factor")
            st.dataframe(pivot.style.background_gradient(cmap='RdYlGn', axis=None).format("{:.1f}"), use_container_width=True)
            
            st.divider()
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
        st.info("Selecciona productos y asegúrate de tener datos cargados para ver la matriz.")