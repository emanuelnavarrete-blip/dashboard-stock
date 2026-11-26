import os
import streamlit as st
import pandas as pd
import altair as alt

st.set_page_config(page_title="Inventario | Stock Real", layout="wide")

# --- ESTILO SOBRIO Y BLOQUEO DE TRADUCTOR ---
st.markdown("""
    <style>
    .stApp { background-color: #F6F5F0; }
    [data-testid="stSidebar"] {
        background-color: #182E26;
        border-right: 1px solid #2A473C;
    }
    [data-testid="stSidebar"] * { color: #F4F2EC !important; }
    [data-testid="stSidebar"] div[data-baseweb="select"] * ,
    [data-testid="stSidebar"] input { color: #1C1E1D !important; }
    div[data-testid="stMetric"] {
        background-color: #FFFFFF;
        border: 1px solid #E4E1D8;
        padding: 16px 20px;
        border-radius: 6px;
    }
    h1, h2, h3 {
        color: #182E26 !important;
        font-weight: 600 !important;
        letter-spacing: -0.01em;
    }
    .sub-caption {
        color: #5C615E;
        font-size: 0.9rem;
        margin-top: -8px;
        margin-bottom: 14px;
    }
    </style>
    <script>
        document.documentElement.lang = 'es';
        document.documentElement.setAttribute('translate', 'no');
    </script>
""", unsafe_allow_html=True)

st.title("Stock Real de Inventario")
st.markdown('<p class="sub-caption">Consulta consolidada por tipo de destino, centro, estado y SKU</p>', unsafe_allow_html=True)

COLOR_PRINCIPAL = "#1E3A34"
PALETA_DONA = ["#1E3A34", "#8C7A5B", "#C5BFA9"]


@st.cache_data
def cargar_datos():
    archivo = "STOCK REAL 26.xlsx" if os.path.exists("STOCK REAL 26.xlsx") else "STOCK REAL 23.xlsx"
    try:
        datos = pd.read_excel(archivo, sheet_name="Stock")
    except Exception:
        datos = pd.read_excel(archivo)
    datos.columns = datos.columns.astype(str).str.strip()
    return datos


df = cargar_datos()

# Detectar columnas dinámicamente
col_tipo = [c for c in df.columns if c.upper().startswith("TIPO DE")][0]
col_producto = [c for c in df.columns if c.upper().startswith("NOM. P")][0]

# --- BARRA LATERAL DE FILTROS ---
st.sidebar.header("Filtros")

opciones_tipo = ["Todos"] + sorted(df[col_tipo].dropna().unique().tolist())
tipo_elegido = st.sidebar.selectbox("Tipo de destino", opciones_tipo)
if tipo_elegido != "Todos":
    df = df[df[col_tipo] == tipo_elegido]

opciones_centro = ["Todos"] + sorted(df["Nombre Centro Destino"].dropna().unique().tolist())
centro_elegido = st.sidebar.selectbox("Centro de destino", opciones_centro)
if centro_elegido != "Todos":
    df = df[df["Nombre Centro Destino"] == centro_elegido]

opciones_estado = ["Todos"] + sorted(df["ESTADO"].dropna().unique().tolist())
estado_elegido = st.sidebar.selectbox("Estado", opciones_estado)
if estado_elegido != "Todos":
    df = df[df["ESTADO"] == estado_elegido]

busqueda = st.sidebar.text_input("SKU o producto")
if busqueda:
    df = df[
        df["COD. PRODUCTO"].astype(str).str.contains(busqueda, case=False, na=False) |
        df[col_producto].astype(str).str.contains(busqueda, case=False, na=False)
    ]

# --- INDICADORES ---
total_real = df["STOCK REAL"].sum()
total_skus = df["COD. PRODUCTO"].nunique()
total_centros = df["Nombre Centro Destino"].nunique()
quiebres = len(df[df["STOCK REAL"] == 0])

col1, col2, col3, col4 = st.columns(4)
col1.metric("Stock real total", f"{int(total_real):,} unid.")
col2.metric("SKUs únicos", f"{total_skus:,}")
col3.metric("Centros de destino", f"{total_centros:,}")
col4.metric("Líneas en cero", f"{quiebres:,}")

st.divider()

# --- FILA 1 DE GRÁFICOS ---
g1, g2 = st.columns([1, 1.6])

with g1:
    st.subheader("CEDI frente a tiendas")
    dist_tipo = df.groupby(col_tipo, as_index=False)["STOCK REAL"].sum()

    dona = alt.Chart(dist_tipo).mark_arc(innerRadius=62, outerRadius=112, stroke="#FFFFFF", strokeWidth=2).encode(
        theta=alt.Theta("STOCK REAL:Q"),
        color=alt.Color(
            f"{col_tipo}:N",
            scale=alt.Scale(range=PALETA_DONA),
            legend=alt.Legend(title="", orient="bottom")
        ),
        tooltip=[col_tipo, alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    ).properties(height=300)
    st.altair_chart(dona, use_container_width=True)

with g2:
    st.subheader("Top 10 centros por stock real")
    top_centros = (
        df.groupby("Nombre Centro Destino", as_index=False)["STOCK REAL"]
        .sum()
        .sort_values("STOCK REAL", ascending=False)
        .head(10)
    )

    barras_centro = alt.Chart(top_centros).mark_bar(color=COLOR_PRINCIPAL, cornerRadiusEnd=3).encode(
        x=alt.X("STOCK REAL:Q", title="Unidades", axis=alt.Axis(format=",.0f")),
        y=alt.Y("Nombre Centro Destino:N", sort="-x", title=""),
        tooltip=["Nombre Centro Destino", alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    )
    texto_centro = barras_centro.mark_text(align="left", dx=5, color=COLOR_PRINCIPAL, fontWeight="bold").encode(
        text=alt.Text("STOCK REAL:Q", format=",.0f")
    )
    st.altair_chart((barras_centro + texto_centro).properties(height=300), use_container_width=True)

# --- FILA 2 DE GRÁFICOS ---
g3, g4 = st.columns([1.6, 1])

with g3:
    st.subheader("Top 10 productos por stock real")
    top_prod = (
        df.groupby(col_producto, as_index=False)["STOCK REAL"]
        .sum()
        .sort_values("STOCK REAL", ascending=False)
        .head(10)
    )

    barras_prod = alt.Chart(top_prod).mark_bar(color=COLOR_PRINCIPAL, cornerRadiusEnd=3).encode(
        x=alt.X("STOCK REAL:Q", title="Unidades", axis=alt.Axis(format=",.0f")),
        y=alt.Y(f"{col_producto}:N", sort="-x", title=""),
        tooltip=[col_producto, alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    )
    texto_prod = barras_prod.mark_text(align="left", dx=5, color=COLOR_PRINCIPAL, fontWeight="bold").encode(
        text=alt.Text("STOCK REAL:Q", format=",.0f")
    )
    st.altair_chart((barras_prod + texto_prod).properties(height=300), use_container_width=True)

with g4:
    st.subheader("Stock real por estado")
    dist_estado = (
        df.groupby("ESTADO", as_index=False)["STOCK REAL"]
        .sum()
        .sort_values("STOCK REAL", ascending=False)
    )

    barras_estado = alt.Chart(dist_estado).mark_bar(
        color=COLOR_PRINCIPAL,
        cornerRadiusTopLeft=3,
        cornerRadiusTopRight=3
    ).encode(
        x=alt.X("ESTADO:N", sort="-y", title="", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("STOCK REAL:Q", title="Unidades", axis=alt.Axis(format=",.0f")),
        tooltip=["ESTADO", alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    )
    texto_estado = barras_estado.mark_text(dy=-8, color=COLOR_PRINCIPAL, fontWeight="bold").encode(
        text=alt.Text("STOCK REAL:Q", format=",.0f")
    )
    st.altair_chart((barras_estado + texto_estado).properties(height=300), use_container_width=True)

st.divider()

# --- TABLA DETALLADA ---
st.subheader("Detalle por centro y SKU")

detalle_sku = (
    df.groupby(["Nombre Centro Destino", "COD. PRODUCTO", col_producto], as_index=False)["STOCK REAL"]
    .sum()
)

detalle_sku = detalle_sku.rename(columns={
    "Nombre Centro Destino": "Centro / Tienda",
    "COD. PRODUCTO": "SKU",
    col_producto: "Producto",
    "STOCK REAL": "Stock Real"
})

detalle_sku = detalle_sku.sort_values(by="Stock Real", ascending=False)

subtotal_tabla = detalle_sku["Stock Real"].sum()

st.info(f"**Subtotal según filtros:** {subtotal_tabla:,.0f} unidades")

fila_total = pd.DataFrame([{
    "Centro / Tienda": "TOTAL GENERAL",
    "SKU": "",
    "Producto": "",
    "Stock Real": subtotal_tabla
}])

detalle_con_total = pd.concat([detalle_sku, fila_total], ignore_index=True)

st.dataframe(
    detalle_con_total,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Stock Real": st.column_config.NumberColumn("Stock Real", format="%d")
    }
)
