import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import requests
from docx import Document
from io import BytesIO

# --- CONFIGURACIÓN ---
# PEGA AQUÍ TU URL DE APPSCRIPT
APPSCRIPT_URL = "TU_URL_DE_IMPLEMENTACION_AQUI"

st.set_page_config(page_title="Consultoría BI - Valoración", layout="wide")

# --- FUNCIONES ---
def get_data():
    try:
        # Usamos el conector de secrets para lectura
        from st_gsheets_connection import GSheetsConnection
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer datos: {e}")
        return pd.DataFrame()

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading(f'Análisis de Capacidades: {empresa}', 0)
    doc.add_paragraph(f'Nicho: {nicho}')
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto'
    hdr_cells[1].text = 'Match Ponderado'
    for prod, score in productos_data.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(prod)
        row_cells[1].text = f"{score:.2f}"
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- INTERFAZ ---
st.title("🎯 Sistema Integral de Estrategia")

tab_val, tab_admin = st.tabs(["📊 Valoración y Radar", "⚙️ Gestión (Crear Empresa/Producto)"])

# --- TAB 1: VALORACIÓN ---
with tab_val:
    df = get_data()
    if not df.empty and 'Empresa' in df.columns:
        c1, c2 = st.columns(2)
        with c1:
            emp_v = st.selectbox("Empresa", df['Empresa'].unique())
            df_e = df[df['Empresa'] == emp_v]
        with c2:
            nicho_v = st.selectbox("Nicho", df_e['Nicho'].unique())
            df_n = df_e[df_e['Nicho'] == nicho_v]

        prods_sel = st.multiselect("Comparar Productos", df_n['Producto'].unique())

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
            
            m_cols = st.columns(len(prods_sel))
            for i, (name, val) in enumerate(res_dict.items()):
                m_cols[i].metric(name, f"{val:.2f} / 5.0")
            
            st.divider()
            st.download_button("💾 Informe Word", generate_docx(emp_v, nicho_v, res_dict), f"Analisis_{emp_v}.docx")

# --- TAB 2: ADMINISTRACIÓN (CREACIÓN) ---
with tab_admin:
    st.header("Registrar Nueva Información")
    st.write("Completa los campos para añadir datos directamente al Excel.")
    
    with st.form("form_registro", clear_on_submit=True):
        col_a, col_b = st.columns(2)
        with col_a:
            f_emp = st.text_input("Nombre Empresa")
            f_nic = st.text_input("Nicho / Segmento")
            f_pro = st.text_input("Producto / Servicio")
        with col_b:
            f_fac = st.text_input("Factor de Éxito (ej: Precio, Marca)")
            f_pes = st.number_input("Peso del factor (1-100)", 1, 100, 20)
            f_cal = st.slider("Calificación de Capacidad (1-5)", 1, 5, 3)
        
        if st.form_submit_button("🚀 Guardar en Google Sheets"):
            if f_emp and f_nic and f_pro and f_fac:
                payload = {
                    "empresa": f_emp, "nicho": f_nic, "producto": f_pro,
                    "factor": f_fac, "peso": f_pes, "calificacion": f_cal
                }
                try:
                    response = requests.post(APPSCRIPT_URL, json=payload)
                    if response.status_code == 200:
                        st.success(f"✅ ¡Éxito! {f_pro} guardado correctamente.")
                        st.balloons()
                    else:
                        st.error("Error al conectar con AppScript. Revisa la URL.")
                except Exception as e:
                    st.error(f"Error de red: {e}")
            else:
                st.warning("Por favor, llena todos los campos antes de guardar.")