import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import requests
from io import BytesIO

# --- 1. CONFIGURACIÓN Y CARGA ---
APPSCRIPT_URL = st.secrets.get("APPSCRIPT_URL")

try:
    from st_gsheets_connection import GSheetsConnection
except ImportError:
    from streamlit_gsheets import GSheetsConnection

st.set_page_config(page_title="BI Strategic Matrix", layout="wide")

@st.cache_data(ttl=0)
def get_data():
    try:
        conn = st.connection("gsheets", type=GSheetsConnection)
        df = conn.read(ttl=0)
        if df is not None and not df.empty:
            df.columns = df.columns.str.strip()
            df['Peso'] = pd.to_numeric(df['Peso'], errors='coerce').fillna(0)
            df['Calificacion'] = pd.to_numeric(df['Calificacion'], errors='coerce').fillna(0)
            return df
        return pd.DataFrame(columns=['Empresa', 'Nicho', 'Producto', 'Factor', 'Peso', 'Calificacion', 'Accionables'])
    except:
        return pd.DataFrame()

df_raw = get_data()

# --- 2. BARRA LATERAL (ESTRUCTURA) ---
with st.sidebar:
    st.header("🏢 Configuración de Proyecto")
    
    # Empresa
    e_opt = ["➕ Crear Nuevo..."] + sorted(list(df_raw['Empresa'].unique())) if not df_raw.empty else ["➕ Crear Nuevo..."]
    e_sel = st.selectbox("Cliente", e_opt)
    e_final = st.text_input("Nueva Empresa", key="side_e") if e_sel == "➕ Crear Nuevo..." else e_sel
    
    # Nicho
    n_base = df_raw[df_raw['Empresa'] == e_sel]['Nicho'].unique() if (not df_raw.empty and e_sel != "➕ Crear Nuevo...") else []
    n_opt = ["➕ Crear Nuevo..."] + sorted(list(n_base))
    n_sel = st.selectbox("Nicho", n_opt)
    n_final = st.text_input("Nuevo Nicho", key="side_n") if n_sel == "➕ Crear Nuevo..." else n_sel

    st.divider()
    st.info("💡 Define arriba el Cliente y Nicho para habilitar la matriz de carga.")

# --- 3. CUERPO PRINCIPAL ---
st.title("📊 Matriz de Evaluación Estratégica")

tab_carga, tab_analisis = st.tabs(["📥 Matriz de Entrada", "📈 Visualización y Reporte"])

# --- PESTAÑA: MATRIZ DE ENTRADA (DISEÑO VISUAL MEJORADO) ---
with tab_carga:
    st.subheader("📝 Registro Masivo de Calificaciones")
    
    # 1. Definición de Dimensiones
    col_dim1, col_dim2 = st.columns(2)
    with col_dim1:
        prods_input = st.text_input("Productos (separados por coma)", placeholder="Ej: Producto A, Producto B")
    with col_dim2:
        factors_input = st.text_input("Factores de Éxito (separados por coma)", placeholder="Ej: Precio, Calidad, Innovación")

    if prods_input and factors_input:
        lista_p = [p.strip() for p in prods_input.split(",") if p.strip()]
        lista_f = [f.strip() for f in factors_input.split(",") if f.strip()]
        
        # 2. Creación de la Matriz Visual
        st.write("---")
        st.write("#### Introduzca Calificaciones (1-5) y Pesos (%)")
        
        # Usamos columnas de Streamlit para simular una tabla de entrada
        header_cols = st.columns([2, 1] + [1] * len(lista_p))
        header_cols[0].write("**Factor**")
        header_cols[1].write("**Peso (%)**")
        for i, p in enumerate(lista_p):
            header_cols[i+2].write(f"**{p}**")

        form_data = []
        
        # Generar filas de la matriz
        for f_idx, f_name in enumerate(lista_f):
            row_cols = st.columns([2, 1] + [1] * len(lista_p))
            row_cols[0].markdown(f"**{f_name}**")
            p_val = row_cols[1].number_input(f"W_{f_idx}", 0, 100, 20, label_visibility="collapsed")
            
            calificaciones_fila = []
            for p_idx, p_name in enumerate(lista_p):
                c_val = row_cols[p_idx+2].number_input(f"C_{f_idx}_{p_idx}", 1, 5, 3, label_visibility="collapsed")
                calificaciones_fila.append(c_val)
            
            form_data.append({"factor": f_name, "peso": p_val, "calificaciones": calificaciones_fila})

        # 3. Accionables Globales por Producto
        st.write("---")
        st.write("#### Accionables / Conclusiones")
        acc_cols = st.columns(len(lista_p))
        lista_acc = []
        for i, p in enumerate(lista_p):
            txt = acc_cols[i].text_area(f"Acción para {p}", key=f"acc_{i}")
            lista_acc.append(txt)

        # 4. Botón de Envío Masivo
        if st.button("💾 Guardar Matriz Completa", type="primary"):
            if not APPSCRIPT_URL:
                st.error("URL de AppScript no configurada.")
            else:
                exitos = 0
                with st.spinner("Guardando registros..."):
                    for i, p_name in enumerate(lista_p):
                        for f_item in form_data:
                            payload = {
                                "empresa": e_final,
                                "nicho": n_final,
                                "producto": p_name,
                                "factor": f_item["factor"],
                                "peso": f_item["peso"],
                                "calificacion": f_item["calificaciones"][i],
                                "accionables": lista_acc[i]
                            }
                            res = requests.post(APPSCRIPT_URL, json=payload)
                            if res.status_code == 200: exitos += 1
                
                st.success(f"Se han guardado {exitos} registros exitosamente.")
                st.cache_data.clear()
                st.rerun()
    else:
        st.info("👆 Ingrese los nombres de los productos y factores para desplegar la matriz de carga.")

# --- PESTAÑA: ANÁLISIS ---
with tab_analisis:
    df_v = df_raw[(df_raw['Empresa'] == e_final) & (df_raw['Nicho'] == n_final)]
    
    if not df_v.empty:
        # Matriz de Revisión (Visualización con mapa de calor)
        st.subheader("📋 Matriz de Desempeño Actual")
        pivot = df_v.pivot_table(index='Factor', columns='Producto', values='Calificacion', aggfunc='mean').fillna(0)
        st.dataframe(pivot.style.background_gradient(cmap='RdYlGn', axis=None).format("{:.1f}"), use_container_width=True)
        
        # Gráfico Radar
        st.write("---")
        st.subheader("📊 Radar Comparativo")
        fig = go.Figure()
        for p in pivot.columns:
            fig.add_trace(go.Scatterpolar(r=pivot[p], theta=pivot.index, fill='toself', name=p))
        fig.update_layout(polar=dict(radialaxis=dict(visible=True, range=[0, 5])))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("No hay datos registrados para este Nicho.")