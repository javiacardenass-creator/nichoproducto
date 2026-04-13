import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import requests
from docx import Document
from docx.shared import Inches
from io import BytesIO
import matplotlib.pyplot as plt

# --- 1. CONFIGURACIÓN DE SEGURIDAD ---
APPSCRIPT_URL = st.secrets.get("APPSCRIPT_URL")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="Consultoría BI Pro", layout="wide")

# --- 2. CARGA DE DATOS ---
@st.cache_data(ttl=0)
def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        if df is not None and not df.empty:
            df.columns = df.columns.str.strip()
            # Limpieza y validación de tipos
            df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
            df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
            if 'Accionables' not in df.columns:
                df['Accionables'] = ""
            return df
        return pd.DataFrame()
    except Exception as e:
        st.error(f"Error al conectar con la base de datos: {e}")
        return pd.DataFrame()

# --- 3. GENERADOR DE WORD ---
def generate_report_docx(empresa, nicho, df_n, prods_sel, resumen_df):
    doc = Document()
    doc.add_heading(f'Informe Estratégico: {empresa}', 0)
    doc.add_paragraph(f'Análisis del Nicho: {nicho}')

    # Gráfico Radar para el Word
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

    doc.add_heading('Plan de Acción', level=1)
    for p in prods_sel:
        score = resumen_df[resumen_df['Producto'] == p]['Puntaje Final'].values[0]
        doc.add_heading(f'Producto: {p} (Score: {score:.2f})', level=2)
        acciones = df_n[df_n['Producto'] == p]['Accionables'].unique()
        for acc in acciones:
            if str(acc).strip() and str(acc) != "nan":
                doc.add_paragraph(f"• {acc}")

    out = BytesIO()
    doc.save(out)
    return out.getvalue()

# --- 4. LÓGICA DE INTERFAZ ---
df_raw = get_data()

if df_raw.empty:
    st.warning("⚠️ La base de datos está vacía o no se puede conectar. Revisa tus Secrets.")
    st.stop()

# --- SIDEBAR (FILTROS) ---
with st.sidebar:
    st.header("🏢 Filtros de Análisis")
    emp_list = sorted(df_raw['Empresa'].unique())
    emp_v = st.selectbox("Empresa Cliente", emp_list)
    
    nic_list = sorted(df_raw[df_raw['Empresa']==emp_v]['Nicho'].unique())
    nic_v = st.selectbox("Nicho Analizado", nic_list)
    
    # Filtrado intermedio para productos
    df_nic = df_raw[(df_raw['Empresa']==emp_v) & (df_raw['Nicho']==nic_v)].copy()
    
    prods_list = sorted(df_nic['Producto'].unique())
    prods_v = st.multiselect("Productos a Comparar", prods_list, default=prods_list)

# Datos filtrados finales para visualización
df_final = df_nic[df_nic['Producto'].isin(prods_v)].copy()

# Pestañas
t_admin, t_radar, t_matriz = st.tabs(["⚙️ Gestión de Datos", "📊 Gráfico Radar", "📋 Matriz de Revisión"])

# --- PESTAÑA 1: GESTIÓN DE DATOS ---
with t_admin:
    st.subheader("🛠️ Registro Dinámico de Factores")
    with st.form("form_registro", clear_on_submit=True):
        col1, col2 = st.columns(2)
        
        with col1:
            e_opt = ["➕ Nueva..."] + list(df_raw['Empresa'].unique())
            e_sel = st.selectbox("Empresa", e_opt)
            e_final = st.text_input("Nombre Nueva Empresa", key="k_emp") if e_sel == "➕ Nueva..." else e_sel
            
            n_base = df_raw[df_raw['Empresa'] == e_sel]['Nicho'].unique() if e_sel != "➕ Nueva..." else []
            n_opt = ["➕ Nuevo..."] + list(n_base)
            n_sel = st.selectbox("Nicho", n_opt)
            n_final = st.text_input("Nombre Nuevo Nicho", key="k_nic") if n_sel == "➕ Nuevo..." else n_sel

        with col2:
            p_base = df_raw[(df_raw['Empresa'] == e_sel) & (df_raw['Nicho'] == n_sel)]['Producto'].unique() if n_sel != "➕ Nuevo..." else []
            p_opt = ["➕ Nuevo..."] + list(p_base)
            p_sel = st.selectbox("Producto", p_opt)
            p_final = st.text_input("Nombre Nuevo Producto", key="k_prod") if p_sel == "➕ Nuevo..." else p_sel
            
            f_base = df_nic['Factor'].unique()
            f_opt = ["➕ Nuevo Factor..."] + list(f_base)
            f_sel = st.selectbox("Factor de Éxito", f_opt)
            f_final = st.text_input("Nombre Nuevo Factor", key="k_fac") if f_sel == "➕ Nuevo Factor..." else f_sel

        st.divider()
        c_a, c_b = st.columns(2)
        w_n = c_a.number_input("Peso (%)", 1, 100, 20)
        c_n = c_b.slider("Calificación", 1, 5, 3)
        acc_n = st.text_area("🎯 Accionables / Conclusiones")
        
        if st.form_submit_button("🚀 Guardar en Google Sheets"):
            if not APPSCRIPT_URL:
                st.error("Error: APPSCRIPT_URL no configurada.")
            elif e_final and n_final and p_final and f_final:
                payload = {"empresa": e_final, "nicho": n_final, "producto": p_final, "factor": f_final, "peso": w_n, "calificacion": c_n, "accionables": acc_n}
                res = requests.post(APPSCRIPT_URL, json=payload)
                if res.status_code == 200:
                    st.success("Guardado con éxito"); st.cache_data.clear(); st.rerun()
                else: st.error("Error al guardar.")

# --- PESTAÑA 2: GRÁFICO RADAR ---
with t_radar:
    if prods_v:
        fig = go.Figure()
        for p in prods_v:
            d_p = df_final[df_final['Producto'] == p]
            fig.add_trace(go.Scatterpolar(r=d_p['Calificacion'], theta=d_p['Factor'], fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])), title=f"Radar: {nic_v}")
        st.plotly_chart(fig, use_container_width=True)

# --- PESTAÑA 3: MATRIZ DE REVISIÓN ---
with t_matriz:
    st.subheader(f"Matriz Comparativa: {nic_v}")
    
    if not df_final.empty:
        # Matriz Cruzada (Factores vs Productos)
        matriz_pivot = df_final.pivot_table(
            index='Factor', 
            columns='Producto', 
            values='Calificacion', 
            aggfunc='mean'
        ).fillna(0)
        
        st.write("### Calificación Detallada por Factor")
        st.dataframe(matriz_pivot.style.background_gradient(cmap='RdYlGn', axis=None).format("{:.1f}"), use_container_width=True)
        
        st.divider()
        # Cálculo de Puntajes Ponderados
        df_final['Ponderado'] = df_final['Calificacion'] * (df_final['Peso'] / 100)
        resumen = df_final.groupby('Producto')['Ponderado'].sum().reset_index()
        resumen.columns = ['Producto', 'Puntaje Final']
        
        c1, c2 = st.columns([1, 2])
        with c1:
            st.write("#### Ranking Ponderado")
            st.table(resumen.sort_values(by='Puntaje Final', ascending=False))
            st.download_button("📂 Descargar Word", generate_report_docx(emp_v, nic_v, df_nic, prods_v, resumen), f"Reporte_{emp_v}.docx")
        with c2:
            st.write("#### Accionables por Producto")
            for p in prods_v:
                accs = df_final[df_final['Producto'] == p]['Accionables'].unique()
                with st.expander(f"Ver acciones: {p}"):
                    for a in accs:
                        if str(a).strip() and str(a) != 'nan': st.info(a)
    else:
        st.info("Selecciona productos para generar la matriz.")