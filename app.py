import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from st_gsheets_connection import GSheetsConnection
from docx import Document
from io import BytesIO

st.set_page_config(page_title="Consultoría Estratégica BI", layout="wide")

# --- CONEXIÓN ---
conn = st.connection("gsheets", type=GSheetsConnection)

def get_data():
    df = conn.read(ttl=0)
    df.columns = df.columns.str.strip()
    return df

# --- INTERFAZ PRINCIPAL ---
st.title("🚀 Plataforma de Estrategia Comercial")

tab1, tab2, tab3 = st.tabs(["🏢 Gestión de Empresas", "⚙️ Configuración de Matriz", "📊 Panel de Valoración"])

# --- TAB 1: CREAR EMPRESA ---
with tab1:
    st.header("Registrar Nueva Empresa")
    with st.form("form_empresa"):
        nueva_empresa = st.text_input("Nombre de la Empresa Cliente")
        nuevo_nicho = st.text_input("Nicho de Mercado inicial")
        submit_emp = st.form_submit_button("Guardar Empresa")
        
        if submit_emp and nueva_empresa and nuevo_nicho:
            # Aquí podrías implementar la lógica para añadir fila al Sheet
            # Por ahora, daremos las instrucciones para el flujo de datos
            st.success(f"Empresa '{nueva_empresa}' lista para configuración.")

# --- TAB 2: CREAR PRODUCTOS Y FACTORES ---
with tab2:
    st.header("Configurar Productos y Factores")
    df = get_data()
    
    empresa_target = st.selectbox("Seleccione Empresa para configurar", df['Empresa'].unique(), key="setup_emp")
    
    col_f1, col_f2 = st.columns(2)
    with col_f1:
        st.subheader("📦 Producto")
        nuevo_prod = st.text_input("Nombre del Producto/Servicio")
    with col_f2:
        st.subheader("💡 Factores de Éxito")
        nuevo_factor = st.text_input("Nombre del Factor (ej. Precio, Soporte)")
        peso_factor = st.number_input("Peso del Factor (%)", min_value=1, max_value=100, value=20)
        calif_factor = st.slider("Calificación de Capacidad", 1, 5, 3)

    if st.button("Añadir a la Matriz"):
        st.info("Nota: Para guardar datos de vuelta a Google Sheets, se requiere configurar una 'Service Account' con permisos de edición.")

# --- TAB 3: VALORACIÓN (Tu código actual optimizado) ---
with tab3:
    df = get_data()
    
    col_v1, col_v2 = st.columns(2)
    with col_v1:
        emp_v = st.selectbox("Empresa", df['Empresa'].unique(), key="v_emp")
        df_e = df[df['Empresa'] == emp_v]
    with col_v2:
        nicho_v = st.selectbox("Nicho", df_e['Nicho'].unique(), key="v_nicho")
        df_n = df_e[df_e['Nicho'] == nicho_v]

    prods_sel = st.multiselect("Seleccione Productos", df_n['Producto'].unique())

    if prods_sel:
        fig = go.Figure()
        res_dict = {}

        for p in prods_sel:
            d_p = df_n[df_n['Producto'] == p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
            
            c = pd.to_numeric(d_p['Calificacion'], errors='coerce').fillna(0)
            w = pd.to_numeric(d_p['Peso'], errors='coerce').fillna(0)
            res_dict[p] = (c * (w/100 if w.max() > 1 else w)).sum()

        st.plotly_chart(fig, use_container_width=True)
        
        # Métricas
        c_met = st.columns(len(prods_sel))
        for i, (name, val) in enumerate(res_dict.items()):
            c_met[i].metric(name, f"{val:.2f} / 5.0")