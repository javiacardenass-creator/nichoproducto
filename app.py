import streamlit as st
import pandas as pd
import plotly.graph_objects as go
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
        return df
    except:
        return pd.DataFrame(columns=['Empresa', 'Nicho', 'Producto', 'Factor', 'Peso', 'Calificacion'])

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
st.title("🎯 Sistema de Valoración Dinámico")

tab_val, tab_admin = st.tabs(["📊 Radar de Valoración", "⚙️ Gestión de Estructura"])

df = get_data()

# --- PESTAÑA 1: VALORACIÓN ---
with tab_val:
    if not df.empty and 'Empresa' in df.columns:
        c1, c2 = st.columns(2)
        with c1:
            emp_v = st.selectbox("Empresa", df['Empresa'].unique(), key="v_emp")
            df_e = df[df['Empresa'] == emp_v]
        with c2:
            nicho_v = st.selectbox("Nicho", df_e['Nicho'].unique(), key="v_nic")
            df_n = df_e[df_e['Nicho'] == nicho_v]

        prods_sel = st.multiselect("Comparar Productos", df_n['Producto'].unique())

        if prods_sel:
            fig = go.Figure()
            res_dict = {}
            for p in prods_sel:
                d_p = df_n[df_n['Producto'] == p]
                fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
                calif = pd.to_numeric(d_p['Calificacion'], errors='coerce').fillna(0)
                peso = pd.to_numeric(d_p['Peso'], errors='coerce').fillna(0)
                res_dict[p] = (calif * (peso/100 if peso.max() > 1 else peso)).sum()

            st.plotly_chart(fig, use_container_width=True)
            m_cols = st.columns(len(prods_sel))
            for i, (name, val) in enumerate(res_dict.items()):
                m_cols[i].metric(name, f"{val:.2f} / 5.0")
            
            st.divider()
            st.download_button("💾 Informe Word", generate_docx(emp_v, nicho_v, res_dict), f"Analisis_{emp_v}.docx")

# --- PESTAÑA 2: GESTIÓN (CARGAR O CREAR) ---
with tab_admin:
    st.header("🛠️ Configuración de Empresa y Portafolio")
    st.write("Selecciona elementos existentes o crea nuevos para expandir la matriz.")

    with st.form("form_gestion", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            # Lógica para Empresa
            opciones_emp = ["➕ Crear nueva empresa..."] + list(df['Empresa'].unique())
            sel_emp = st.selectbox("Empresa", opciones_emp)
            nom_emp = st.text_input("Nombre de nueva empresa") if sel_emp == "➕ Crear nueva empresa..." else sel_emp

            # Lógica para Nicho
            if sel_emp != "➕ Crear nueva empresa...":
                opciones_nic = ["➕ Crear nuevo nicho..."] + list(df[df['Empresa'] == sel_emp]['Nicho'].unique())
            else:
                opciones_nic = ["➕ Crear nuevo nicho..."]
            sel_nic = st.selectbox("Nicho", opciones_nic)
            nom_nic = st.text_input("Nombre de nuevo nicho") if sel_nic == "➕ Crear nuevo nicho..." else sel_nic

        with col2:
            # Lógica para Producto
            if sel_nic != "➕ Crear nuevo nicho..." and sel_emp != "➕ Crear nueva empresa...":
                opciones_pro = ["➕ Crear nuevo producto..."] + list(df[(df['Empresa'] == sel_emp) & (df['Nicho'] == sel_nic)]['Producto'].unique())
            else:
                opciones_pro = ["➕ Crear nuevo producto..."]
            sel_pro = st.selectbox("Producto", opciones_pro)
            nom_pro = st.text_input("Nombre de nuevo producto") if sel_pro == "➕ Crear nuevo producto..." else sel_pro

            # Factor, Peso y Calificación siempre se crean/asignan
            nom_fac = st.text_input("Nuevo Factor de Éxito")
            f_peso = st.number_input("Peso (%)", 1, 100, 20)
            f_cali = st.slider("Calificación (1-5)", 1, 5, 3)

        if st.form_submit_button("🚀 Guardar Configuración"):
            # Validar que tengamos nombres finales
            final_emp = nom_emp if nom_emp else sel_emp
            final_nic = nom_nic if nom_nic else sel_nic
            final_pro = nom_pro if nom_pro else sel_pro

            if final_emp and final_nic and final_pro and nom_fac:
                payload = {
                    "empresa": final_emp, "nicho": final_nic, "producto": final_pro,
                    "factor": nom_fac, "peso": f_peso, "calificacion": f_cali
                }
                try:
                    res = requests.post(APPSCRIPT_URL, json=payload)
                    if res.status_code == 200:
                        st.success(f"✅ Se ha registrado el factor '{nom_fac}' para {final_pro}")
                        st.cache_data.clear() # Limpia caché para recargar datos
                    else:
                        st.error("Error al conectar con AppScript.")
                except Exception as e:
                    st.error(f"Error de red: {e}")
            else:
                st.warning("Completa todos los campos necesarios.")