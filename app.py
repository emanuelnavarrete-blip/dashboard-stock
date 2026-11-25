import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="Gestión de Inventario | Stock Real", layout="wide")

# --- ESTILO VISUAL CORPORATIVO Y BLOQUEO DE TRADUCTOR ---
st.markdown("""
    <style>
    .stApp { background-color: #F4F6F9; }
    [data-testid="stSidebar"] { background-color: #0F2C59; }
    [data-testid="stSidebar"] * { color: #FFFFFF !important; }
    [data-testid="stSidebar"] div[data-baseweb="select"] * ,
    [data-testid="stSidebar"] input { color: #1E293B !important; }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border-top: 4px solid #0F2C59;
        padding: 18px 20px;
        border-radius: 8px;
        box-shadow: 0 2px 6px rgba(15, 44, 89, 0.08);
    }
    h1, h2, h3 { color: #0F2C59 !important; }
    </style>
    <script>
        document.documentElement.lang = 'es';
        document.documentElement.setAttribute('translate', 'no');
    </script>
""", unsafe_allow_html=True)

st.title("📊 Control Corporativo de Stock Real")

@st.cache_data
def cargar_datos():
    datos = pd.read_excel("STOCK REAL 23.xlsx", sheet_name="Stock")
    datos.columns = datos.columns.str.strip()
    return datos

df = cargar_datos()

# Detectar columnas dinámicamente
col_tipo = [c for c in df.columns if c.upper().startswith("TIPO DE")][0]
col_producto = [c for c in df.columns if c.upper().startswith("NOM. P")][0]

# --- BARRA LATERAL DE FILTROS ---
st.sidebar.header("Filtros de Consulta")

opciones_tipo = ["Todos"] + sorted(df[col_tipo].dropna().unique().tolist())
tipo_elegido = st.sidebar.selectbox("1. Tipo de Destino (CEDI / Tienda):", opciones_tipo)
if tipo_elegido != "Todos":
    df = df[df[col_tipo] == tipo_elegido]

opciones_centro = ["Todos"] + sorted(df["Nombre Centro Destino"].dropna().unique().tolist())
centro_elegido = st.sidebar.selectbox("2. Nombre Centro Destino:", opciones_centro)
if centro_elegido != "Todos":
    df = df[df["Nombre Centro Destino"] == centro_elegido]

opciones_estado = ["Todos"] + sorted(df["ESTADO"].dropna().unique().tolist())
estado_elegido = st.sidebar.selectbox("3. Estado:", opciones_estado)
if estado_elegido != "Todos":
    df = df[df["ESTADO"] == estado_elegido]

busqueda = st.sidebar.text_input("Buscar por SKU o Producto:")
if busqueda:
    df = df[
        df["COD. PRODUCTO"].astype(str).str.contains(busqueda, case=False, na=False) |
        df[col_producto].astype(str).str.contains(busqueda, case=False, na=False)
    ]

# --- TARJETAS KPI EJECUTIVAS ---
total_real = df["STOCK REAL"].sum()
total_skus = df["COD. PRODUCTO"].nunique()
total_centros = df["Nombre Centro Destino"].nunique()
quiebres = len(df[df["STOCK REAL"] == 0])

col1, col2, col3, col4 = st.columns(4)
col1.metric("Stock Real Total", f"{int(total_real):,} Unid.")
col2.metric("SKUs Únicos", f"{total_skus:,}")
col3.metric("Centros de Destino", f"{total_centros:,}")
col4.metric("Líneas sin Stock (0)", f"{quiebres:,}")

st.divider()

# --- FILA 1 DE GRÁFICOS EJECUTIVOS ---
g1, g2 = st.columns([1, 1.6])

with g1:
    st.subheader("Distribución: CEDI vs. Tiendas")
    dist_tipo = df.groupby(col_tipo, as_index=False)["STOCK REAL"].sum()
    
    dona = alt.Chart(dist_tipo).mark_arc(innerRadius=65, outerRadius=115).encode(
        theta=alt.Theta("STOCK REAL:Q"),
        color=alt.Color(
            f"{col_tipo}:N",
            scale=alt.Scale(range=["#0F2C59", "#3B82F6", "#93C5FD"]),
            legend=alt.Legend(title="Tipo", orient="bottom")
        ),
        tooltip=[col_tipo, alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    ).properties(height=300)
    st.altair_chart(dona, use_container_width=True)

with g2:
    st.subheader("Top 10 Centros de Destino por Stock Real")
    top_centros = (
        df.groupby("Nombre
    
