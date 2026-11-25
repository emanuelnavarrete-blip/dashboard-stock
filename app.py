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
        df.groupby("Nombre Centro Destino", as_index=False)["STOCK REAL"]
        .sum()
        .sort_values("STOCK REAL", ascending=False)
        .head(10)
    )
    
    barras_centro = alt.Chart(top_centros).mark_bar(color="#0F2C59", cornerRadiusEnd=4).encode(
        x=alt.X("STOCK REAL:Q", title="Unidades", axis=alt.Axis(format=",.0f")),
        y=alt.Y("Nombre Centro Destino:N", sort="-x", title=""),
        tooltip=["Nombre Centro Destino", alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    )
    texto_centro = barras_centro.mark_text(align="left", dx=5, color="#0F2C59", fontWeight="bold").encode(
        text=alt.Text("STOCK REAL:Q", format=",.0f")
    )
    st.altair_chart((barras_centro + texto_centro).properties(height=300), use_container_width=True)

# --- FILA 2 DE GRÁFICOS EJECUTIVOS ---
g3, g4 = st.columns([1.6, 1])

with g3:
    st.subheader("Top 10 Productos (SKU) con Mayor Stock")
    top_prod = (
        df.groupby(col_producto, as_index=False)["STOCK REAL"]
        .sum()
        .sort_values("STOCK REAL", ascending=False)
        .head(10)
    )
    
    barras_prod = alt.Chart(top_prod).mark_bar(color="#1D4ED8", cornerRadiusEnd=4).encode(
        x=alt.X("STOCK REAL:Q", title="Unidades", axis=alt.Axis(format=",.0f")),
        y=alt.Y(f"{col_producto}:N", sort="-x", title=""),
        tooltip=[col_producto, alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    )
    texto_prod = barras_prod.mark_text(align="left", dx=5, color="#1D4ED8", fontWeight="bold").encode(
        text=alt.Text("STOCK REAL:Q", format=",.0f")
    )
    st.altair_chart((barras_prod + texto_prod).properties(height=300), use_container_width=True)

with g4:
    st.subheader("Stock Real por Estado")
    dist_estado = (
        df.groupby("ESTADO", as_index=False)["STOCK REAL"]
        .sum()
        .sort_values("STOCK REAL", ascending=False)
    )
    
    barras_estado = alt.Chart(dist_estado).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
        x=alt.X("ESTADO:N", sort="-y", title="", axis=alt.Axis(labelAngle=0)),
        y=alt.Y("STOCK REAL:Q", title="Unidades", axis=alt.Axis(format=",.0f")),
        color=alt.Color("ESTADO:N", scale=alt.Scale(range=["#0F2C59", "#0284C7", "#64748B"]), legend=None),
        tooltip=["ESTADO", alt.Tooltip("STOCK REAL:Q", format=",.0f")]
    )
    texto_estado = barras_estado.mark_text(dy=-8, color="#0F2C59", fontWeight="bold").encode(
        text=alt.Text("STOCK REAL:Q", format=",.0f")
    )
    st.altair_chart((barras_estado + texto_estado).properties(height=300), use_container_width=True)

st.divider()

# --- TABLA DETALLADA: TIENDA/CEDI + SKU + PRODUCTO ---
st.subheader("Detalle de Stock Real por Tienda / CEDI y SKU - Producto")

detalle_sku = (
    df.groupby(["Nombre Centro Destino", "COD. PRODUCTO", col_producto], as_index=False)["STOCK REAL"]
    .sum()
)

detalle_sku = detalle_sku.rename(columns={
    "Nombre Centro Destino": "Tienda / CEDI",
    "COD. PRODUCTO": "SKU",
    col_producto: "Producto",
    "STOCK REAL": "Stock Real"
})

detalle_sku = detalle_sku.sort_values(by="Stock Real", ascending=False)

# 1. Calcular el subtotal dinámico según los filtros aplicados
subtotal_tabla = detalle_sku["Stock Real"].sum()

# 2. Mostrar recuadro resumen justo encima de la tabla
st.info(f"**Subtotal Stock Real (según filtros aplicados):** {subtotal_tabla:,.0f} Unidades")

# 3. Agregar fila de TOTAL al final de la tabla
fila_total = pd.DataFrame([{
    "Tienda / CEDI": "TOTAL GENERAL",
    "SKU": "",
    "Producto": "",
    "Stock Real": subtotal_tabla
}])

detalle_con_total = pd.concat([detalle_sku, fila_total], ignore_index=True)

# 4. Mostrar la tabla con formato
st.dataframe(
    detalle_con_total,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Stock Real": st.column_config.NumberColumn("Stock Real", format="%d")
