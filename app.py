import streamlit as st
import pandas as pd
import plotly.graph_objects as go
from docx import Document
from io import BytesIO

# Importación segura del conector
try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Estrategia Comercial BI", layout="wide")

# --- FUNCIONES DE APOYO ---
def get_data(conn):
    try:
        df = conn.read(ttl=0)
        df.columns = df.columns.str.strip()
        # Asegurar que las columnas existen para evitar errores de carga inicial
        cols = ['Empresa', 'Nicho', 'Producto', 'Factor', 'Peso', 'Calificacion']
        for c in cols:
            if c not in df.columns:
                df[c] = None
        return df
    except:
        return pd.DataFrame(columns=['Empresa', 'Nicho', 'Producto', 'Factor', 'Peso', 'Calificacion'])

def generate_docx(empresa, nicho, productos_data):
    doc = Document()
    doc.add_heading(f'Análisis Estratégico: {empresa}', 0)
    doc.add_paragraph(f'Nicho: {nicho}')
    table = doc.add_table(rows=1, cols=2)
    table.style = 'Table Grid'
    hdr_cells = table.rows[0].cells
    hdr_cells[0].text = 'Producto'
    hdr_cells[1].text = 'Puntaje de Ajuste'
    for prod, score in productos_data.items():
        row_cells = table.add_row().cells
        row_cells[0].text = str(prod)
        row_cells[1].text = f"{score:.2f}"
    bio = BytesIO()
    doc.save(bio)
    return bio.getvalue()

# --- LÓGICA PRINCIPAL ---
def main():
    st.title("🎯 Plataforma de Consultoría Estratégica")
    
    conn = st.connection("gsheets", type=GSheetsConnection)
    df = get_data(conn)

    # Navegación por pestañas
    tab_val, tab_admin = st.tabs(["📊 Valoración y Radar", "⚙️ Gestión de Datos"])

    # --- PESTAÑA 1: VALORACIÓN ---
    with tab_val:
        if not df.empty and df['Empresa'].notna().any():
            c1, c2 = st.columns(2)
            with c1:
                emp_sel = st.selectbox("Seleccione Empresa", df['Empresa'].unique())
                df_e = df[df['Empresa'] == emp_sel]
            with c2:
                nicho_sel = st.selectbox("Seleccione Nicho", df_e['Nicho'].unique())
                df_n = df_e[df_e['Nicho'] == nicho_sel]

            prods_sel = st.multiselect("Seleccione Productos para comparar", df_n['Producto'].unique())

            if prods_sel:
                fig = go.Figure()
                res_dict = {}
                for p in prods_sel:
                    d_p = df_n[df_n['Producto'] == p]
                    fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
                    
                    c = pd.to_numeric(d_p['Calificacion'], errors='coerce').fillna(0)
                    w = pd.to_numeric(d_p['Peso'], errors='coerce').fillna(0)
                    res_dict[p] = (c * w / 100).sum() if w.max() > 1 else (c * w).sum()

                st.plotly_chart(fig, use_container_width=True)
                
                # Métricas
                cols_m = st.columns(len(prods_sel))
                for i, (name, val) in enumerate(res_dict.items()):
                    cols_m[i].metric(name, f"{val:.2f} / 5.0")

                st.divider()
                if st.download_button("📄 Descargar Reporte Word", generate_docx(emp_sel, nicho_sel, res_dict), f"Analisis_{emp_sel}.docx"):
                    st.success("Reporte generado.")
        else:
            st.info("No hay datos registrados. Ve a la pestaña de Gestión para crear tu primera empresa.")

    # --- PESTAÑA 2: GESTIÓN (CREACIÓN) ---
    with tab_admin:
        st.header("Administración de Matriz")
        
        with st.expander("➕ Crear Nueva Empresa / Nicho"):
            with st.form("nueva_empresa"):
                n_emp = st.text_input("Nombre de la Empresa")
                n_nic = st.text_input("Nicho")
                if st.form_submit_button("Registrar Base"):
                    st.warning("Para persistir cambios en el Excel, conecta tu Service Account JSON en Secrets.")
                    # Aquí iría: conn.create(data=[[n_emp, n_nic, ...]])

        with st.expander("📋 Agregar Producto y Factores"):
            if not df.empty:
                with st.form("nuevo_producto"):
                    f_emp = st.selectbox("Empresa Destino", df['Empresa'].unique())
                    f_nic = st.selectbox("Nicho Destino", df[df['Empresa']==f_emp]['Nicho'].unique())
                    f_prod = st.text_input("Nombre del Producto")
                    f_fact = st.text_input("Factor de Éxito")
                    f_peso = st.number_input("Peso (%)", 1, 100, 20)
                    f_cali = st.slider("Calificación (1-5)", 1, 5, 3)
                    
                    if st.form_submit_button("Añadir a Matriz"):
                        st.info("Datos listos para enviar. Requiere permisos de edición en el Sheet.")
            else:
                st.write("Crea una empresa primero.")

if __name__ == "__main__":
    main()