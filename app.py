from io import BytesIO
import re
import altair as alt
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Dashboard de Stock Real",
    page_icon="🌲",
    layout="wide",
)

# Estética ejecutiva en Verde Bosque
VERDE_OSCURO = "#143D29"
VERDE_BOSQUE = "#1E5631"
VERDE_CLARO = "#E8F2EC"
VERDE_ACENTO = "#2E7D32"
BORDE_SUAVE = "#C8DDD0"

st.markdown(
    f"""
    <style>
        .main {{
            background-color: #F7FAF8;
        }}
        .titulo-ejecutivo {{
            background: linear-gradient(90deg, {VERDE_OSCURO} 0%, {VERDE_BOSQUE} 100%);
            color: #FFFFFF;
            padding: 18px 24px;
            border-radius: 10px;
            margin-bottom: 16px;
        }}
        .titulo-ejecutivo h1 {{
            color: #FFFFFF;
            font-size: 1.75rem;
            margin: 0 0 4px 0;
        }}
        .titulo-ejecutivo p {{
            color: #DDF0E4;
            margin: 0;
            font-size: 0.95rem;
        }}
        div[data-testid="stMetric"] {{
            background-color: #FFFFFF;
            border: 1px solid {BORDE_SUAVE};
            border-left: 5px solid {VERDE_BOSQUE};
            padding: 12px 16px;
            border-radius: 8px;
        }}
    </style>
    """,
    unsafe_allow_html=True,
)


def _limpiar_texto(serie: pd.Series) -> pd.Series:
    return (
        serie.fillna("")
        .astype(str)
        .str.strip()
        .replace({"nan": "", "None": "", "NAN": ""})
    )


def _limpiar_numero(serie: pd.Series) -> pd.Series:
    limpio = (
        serie.fillna(0)
        .astype(str)
        .str.replace(" ", "", regex=False)
        .str.replace(",", "", regex=False)
    )
    return pd.to_numeric(limpio, errors="coerce").fillna(0.0)


@st.cache_data(show_spinner="Cargando datos de stock...")
def cargar_datos(ruta_excel: str = "STOCK REAL 26.xlsx") -> pd.DataFrame:
    xl = pd.ExcelFile(ruta_excel, engine="openpyxl")
    hoja = "Stock" if "Stock" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(ruta_excel, sheet_name=hoja, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Detección dinámica de columna de stock ("STOCK REAL" vs "STOCK TOTAL")
    col_stock = None
    for candidata in ["STOCK REAL", "STOCK TOTAL"]:
        for c in df.columns:
            if c.upper() == candidata:
                col_stock = c
                break
        if col_stock:
            break
    if col_stock is None:
        for c in df.columns:
            if "STOCK" in c.upper():
                col_stock = c
                break

    if col_stock is None:
        raise ValueError("No se encontró una columna de STOCK REAL o STOCK TOTAL en el archivo.")

    df["STOCK_NUM"] = _limpiar_numero(df[col_stock])

    columnas_esperadas = [
        "TIPO DE DESTINO",
        "Nombre Centro Destino",
        "COD. PRODUCTO",
        "NOM. PRODUCTO",
        "ESTADO",
    ]
    for col in columnas_esperadas:
        if col not in df.columns:
            df[col] = "SIN DATO"
        else:
            df[col] = _limpiar_texto(df[col])
            df.loc[df[col] == "", col] = "SIN DATO"

    df["STOCK REAL"] = df["STOCK_NUM"]
    cols_finales = columnas_esperadas + ["STOCK REAL"]
    return df[cols_finales].copy()


def generar_excel_con_tabla(df_export: pd.DataFrame) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Stock Filtrado"

    encabezados = list(df_export.columns)
    ws.append(encabezados)

    fill_header = PatternFill(start_color="1E5631", end_color="1E5631", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True, size=11)

    for col_idx in range(1, len(encabezados) + 1):
        celda = ws.cell(row=1, column=col_idx)
        celda.fill = fill_header
        celda.font = font_header
        celda.alignment = Alignment(horizontal="center", vertical="center")

    col_stock_idx = encabezados.index("STOCK REAL") + 1 if "STOCK REAL" in encabezados else None

    for row in df_export.itertuples(index=False, name=None):
        ws.append(list(row))

    n_filas = len(df_export)
    n_cols = len(encabezados)
    ultima_fila = n_filas + 1

    if col_stock_idx is not None and n_filas > 0:
        col_letra_stock = get_column_letter(col_stock_idx)
        for r in range(2, ultima_fila + 1):
            c_val = ws.cell(row=r, column=col_stock_idx)
            c_val.number_format = "#,##0.00"

        fila_total = ultima_fila + 1
        ws.cell(row=fila_total, column=1, value="TOTAL").font = Font(bold=True, color="143D29")
        celda_total = ws.cell(
            row=fila_total,
            column=col_stock_idx,
            value=f"=SUM({col_letra_stock}2:{col_letra_stock}{ultima_fila})",
        )
        celda_total.font = Font(bold=True, color="143D29")
        celda_total.number_format = "#,##0.00"
        celda_total.fill = PatternFill(start_color="E8F2EC", end_color="E8F2EC", fill_type="solid")

    if n_filas > 0:
        ultima_letra = get_column_letter(n_cols)
        ref_tabla = f"A1:{ultima_letra}{ultima_fila}"
        nombre_limpio = re.sub(r"[^A-Za-z0-9_]", "_", "TablaStock")
        tabla = Table(displayName=nombre_limpio, ref=ref_tabla)
        estilo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False,
        )
        tabla.tableStyleInfo = estilo
        ws.add_table(tabla)

    for col_idx in range(1, n_cols + 1):
        letra = get_column_letter(col_idx)
        max_len = len(str(encabezados[col_idx - 1]))
        for r in range(2, min(ultima_fila + 1, 502)):
            val_str = str(ws.cell(row=r, column=col_idx).value or "")
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[letra].width = min(max(max_len + 4, 14), 45)

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida.getvalue()


st.markdown(
    """
    <div class="titulo-ejecutivo">
        <h1>🌲 Dashboard Ejecutivo de Stock Real</h1>
        <p>Control de inventario por tipo de destino, centro, estado y SKU</p>
    </div>
    """,
    unsafe_allow_html=True,
)

try:
    df = cargar_datos("STOCK REAL 26.xlsx")
except Exception as e:
    st.error(f"No se pudo cargar el archivo de datos: {e}")
    st.stop()

# Filtros en barra lateral
st.sidebar.header("Filtros de consulta")

tipos_sel = st.sidebar.multiselect(
    "Tipo de destino",
    options=sorted(df["TIPO DE DESTINO"].unique()),
    default=[],
)

estados_sel = st.sidebar.multiselect(
    "Estado",
    options=sorted(df["ESTADO"].unique()),
    default=[],
)

df_base_centros = df.copy()
if tipos_sel:
    df_base_centros = df_base_centros[df_base_centros["TIPO DE DESTINO"].isin(tipos_sel)]
if estados_sel:
    df_base_centros = df_base_centros[df_base_centros["ESTADO"].isin(estados_sel)]

centros_sel = st.sidebar.multiselect(
    "Centro destino",
    options=sorted(df_base_centros["Nombre Centro Destino"].unique()),
    default=[],
)

sku_buscar = st.sidebar.text_input(
    "Buscar SKU (código o nombre de producto)",
    placeholder="Ej. 100101001001 o CARBON",
).strip()

solo_con_stock = st.sidebar.checkbox("Mostrar solo registros con stock distinto de 0", value=False)

# Aplicar filtros
df_filtrado = df.copy()
if tipos_sel:
    df_filtrado = df_filtrado[df_filtrado["TIPO DE DESTINO"].isin(tipos_sel)]
if estados_sel:
    df_filtrado = df_filtrado[df_filtrado["ESTADO"].isin(estados_sel)]
if centros_sel:
    df_filtrado = df_filtrado[df_filtrado["Nombre Centro Destino"].isin(centros_sel)]
if sku_buscar:
    mask_sku = df_filtrado["COD. PRODUCTO"].str.contains(sku_buscar, case=False, na=False) | df_filtrado[
        "NOM. PRODUCTO"
    ].str.contains(sku_buscar, case=False, na=False)
    df_filtrado = df_filtrado[mask_sku]
if solo_con_stock:
    df_filtrado = df_filtrado[df_filtrado["STOCK REAL"] != 0]

# KPIs principales
stock_total = float(df_filtrado["STOCK REAL"].sum())
registros_cnt = len(df_filtrado)
skus_unicos = df_filtrado["COD. PRODUCTO"].nunique()
centros_unicos = df_filtrado["Nombre Centro Destino"].nunique()

c1, c2, c3, c4 = st.columns(4)
c1.metric("Stock Real Total", f"{stock_total:,.2f}")
c2.metric("Registros filtrados", f"{registros_cnt:,}")
c3.metric("SKUs únicos", f"{skus_unicos:,}")
c4.metric("Centros destino", f"{centros_unicos:,}")

st.markdown("---")

if df_filtrado.empty:
    st.warning("No hay registros que coincidan con los filtros seleccionados.")
else:
    col_izq, col_der = st.columns(2)

    with col_izq:
        st.subheader("Stock Real por Tipo de Destino")
        df_tipo = (
            df_filtrado.groupby("TIPO DE DESTINO", as_index=False)["STOCK REAL"]
            .sum()
            .sort_values("STOCK REAL", ascending=False)
        )
        chart_tipo = (
            alt.Chart(df_tipo)
            .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
            .encode(
                x=alt.X("STOCK REAL:Q", title="Stock Real"),
                y=alt.Y("TIPO DE DESTINO:N", sort="-x", title="Tipo de Destino"),
                color=alt.value(VERDE_BOSQUE),
                tooltip=[
                    alt.Tooltip("TIPO DE DESTINO:N", title="Tipo"),
                    alt.Tooltip("STOCK REAL:Q", title="Stock Real", format=",.2f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(chart_tipo, use_container_width=True)

    with col_der:
        st.subheader("Top 10 Centros por Stock Real")
        df_centro = (
            df_filtrado.groupby("Nombre Centro Destino", as_index=False)["STOCK REAL"]
            .sum()
            .sort_values("STOCK REAL", ascending=False)
            .head(10)
        )
        chart_centro = (
            alt.Chart(df_centro)
            .mark_bar(cornerRadiusTopRight=6, cornerRadiusBottomRight=6)
            .encode(
                x=alt.X("STOCK REAL:Q", title="Stock Real"),
                y=alt.Y("Nombre Centro Destino:N", sort="-x", title="Centro Destino"),
                color=alt.value(VERDE_ACENTO),
                tooltip=[
                    alt.Tooltip("Nombre Centro Destino:N", title="Centro"),
                    alt.Tooltip("STOCK REAL:Q", title="Stock Real", format=",.2f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(chart_centro, use_container_width=True)

    st.subheader("Detalle de Inventario")
    excel_bytes = generar_excel_con_tabla(df_filtrado)
    st.download_button(
        label="📊 Exportar a Excel (Tabla y Fórmula SUMA)",
        data=excel_bytes,
        file_name="Stock_Real_Filtrado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    st.dataframe(df_filtrado.head(2000), use_container_width=True, hide_index=True)
    if len(df_filtrado) > 2000:
        st.caption(
            f"Mostrando los primeros 2,000 registros en pantalla de {len(df_filtrado):,} filtrados. "
            "La descarga en Excel incluye todos los registros filtrados."
        )
