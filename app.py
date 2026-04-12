import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import requests
from docx import Document
from io import BytesIO

# --- CONFIGURACIÓN DE CONEXIÓN ---
# 1. Pega aquí la URL que obtuviste al "Implementar" en Apps Script
APPSCRIPT_URL = "https://script.google.com/macros/s/AKfycbzmIlzhVjPyCtg16hEzsIK-CribvHSphonlaQfxJnwZ17tOI6AISZxMHvfvNYuu9LclEA/exec"

# 2. Importación segura del conector de Google Sheets
try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    try:
        from streamlit_gsheets import GSheetsConnection
    except ImportError:
        GSheetsConnection = None

st.set_page_config(page_title="Consultoría BI - Valoración", layout="wide")

# --- FUNCIONES DE APOYO ---
def get_data():
    if GSheetsConnection is None:
        st.error("Error: La librería 'st-gsheets-connection' no está instalada.")
        return pd.DataFrame()
    
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        # Limpieza de nombres de columnas
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        st.error(f"Error al leer datos desde Google Sheets: {e}")
        return pd.DataFrame()

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading(f'Reporte Estratégico: {empresa}', 0)
    doc.add_paragraph(f'Nicho de Mercado: {nicho}')
    
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto / Servicio'
    hdr_cells[1].text = 'Índice de Ajuste'
    
    for prod, score in productos_data.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(prod)
        row_cells[1].text = f"{score:.2f}"
    
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- INTERFAZ PRINCIPAL ---
st.title("🎯 Sistema de Gestión y Valoración de Nichos")

tab_val, tab_admin = st.tabs(["📊 Radar de Valoración", "⚙️ Gestión (Crear Datos)"])

# --- PESTAÑA 1: RADAR Y VALORACIÓN ---
with tab_val:
    df = get_data()
    
    if not df.empty and 'Empresa' in df.columns:
        col_v1, col_v2 = st.columns(2)
        
        with col_v1:
            emp_v = st.selectbox("1. Seleccione Empresa", df['Empresa'].unique())
            df_e = df[df['Empresa'] == emp_v]
        
        with col_v2:
            nicho_v = st.selectbox("2. Seleccione Nicho", df_e['Nicho'].unique())
            df_n = df_e[df_e['Nicho'] == nicho_v]

        st.divider()
        
        prods_sel = st.multiselect("3. Seleccione Productos para comparar", df_n['Producto'].unique())

        if prods_sel:
            fig = go.Figure()
            res_dict = {}

            for p in prods_sel:
                d_p = df_n[df_n['Producto'] == p]
                
                # Radar
                fig.add_trace(go.Scatterpolar(
                    r=d_p['Calificacion'],
                    theta=d_p['Factor'],
                    fill='toself',
                    name=p
                ))
                
                # Cálculo de Score Ponderado
                calif = pd.to_numeric(d_p['Calificacion'], errors='coerce').fillna(0)
                peso = pd.to_numeric(d_p['Peso'], errors='coerce').fillna(0)
                
                # Si el peso es > 1 (ej: 20), dividimos por 100
                score = (calif * (peso/100 if peso.max() > 1 else peso)).sum()
                res_dict[p] = score

            fig.update_layout(
                polar=dict(radialaxis=dict(visible=True, range=[0, 5])),
                showlegend=True,
                title=f"Ajuste de Capacidades - Nicho: {nicho_v}"
            )
            
            st.plotly_chart(fig, use_container_width=True)

            # Métricas
            m_cols = st.columns(len(prods_sel))
            for i, (name, val) in enumerate(res_dict.items()):
                m_cols[i].metric(name, f"{val:.2f} / 5.0")

            st.divider()
            
            # Botón de Reporte
            reporte = generate_docx(emp_v, nicho_v, res_dict)
            st.download_button(
                label="📄 Descargar Informe Word",
                data=reporte,
                file_name=f"Analisis_{emp_v}_{nicho_v}.docx",
                mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
            )
    else:
        st.warning("No se detectaron datos. Asegúrate de que el Google Sheet tenga los encabezados correctos.")

# --- PESTAÑA 2: ADMINISTRACIÓN (POST A APPSCRIPT) ---
with tab_admin:
    st.header("➕ Registrar Nueva Información")
    st.info("Desde aquí puedes crear empresas, nichos y productos sin abrir el Excel.")
    
    with st.form("form_registro", clear_on_submit=True):
        f_col1, f_col2 = st.columns(2)
        
        with f_col1:
            f_emp = st.text_input("Nombre de la Empresa")
            f_nic = st.text_input("Nombre del Nicho")
            f_pro = st.text_input("Producto o Servicio")
        
        with f_col2:
            f_fac = st.text_input("Factor de Éxito (Atributo)")
            f_pes = st.number_input("Peso del Factor (1-100)", 1, 100, 20)
            f_cal = st.slider("Calificación de Capacidad (1-5)", 1, 5, 3)
        
        enviar = st.form_submit_button("🚀 Guardar en Google Sheets")

        if enviar:
            if f_emp and f_nic and f_pro and f_fac:
                # Datos para enviar al AppScript
                payload = {
                    "empresa": f_emp,
                    "nicho": f_nic,
                    "producto": f_pro,
                    "factor": f_fac,
                    "peso": f_pes,
                    "calificacion": f_cal
                }
                
                try:
                    res = requests.post(APPSCRIPT_URL, json=payload)
                    if res.status_code == 200:
                        st.success(f"✅ ¡Datos guardados! La empresa '{f_emp}' ha sido actualizada.")
                        st.balloons()
                    else:
                        st.error("Error al conectar con AppScript. Verifica la URL de implementación.")
                except Exception as e:
                    st.error(f"Falla de red al intentar guardar: {e}")
            else:
                st.warning("⚠️ Completa todos los campos para poder registrar la información.")