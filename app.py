from io import BytesIO
import re
import altair as alt
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
import pandas as pd
import streamlit as st

TITULO_REPORTE = "REPORTE STOCK - VENTA - CANTIDADES VENDIDAS AL 25"

st.set_page_config(
    page_title=TITULO_REPORTE,
    page_icon="📊",
    layout="wide",
)

# Paleta corporativa sobria para gerencia (azul marino + pizarra + gris neutro)
NAVY_OSCURO = "#162A3E"
NAVY_MEDIO = "#243B53"
AZUL_ACERO = "#335C67"
GRIS_TEXTO = "#1F2933"
GRIS_SUB = "#52606D"
FONDO_APP = "#F4F6F8"
FONDO_TARJETA = "#FFFFFF"
BORDE_LINEA = "#D9E2EC"

st.markdown(
    f"""
    <style>
        html, body, .stApp {{
            background-color: {FONDO_APP};
            color: {GRIS_TEXTO};
        }}
        .encabezado-gerencia {{
            background-color: {NAVY_OSCURO};
            color: #FFFFFF;
            padding: 20px 26px;
            border-radius: 6px;
            border-bottom: 3px solid {AZUL_ACERO};
            margin-bottom: 18px;
        }}
        .encabezado-gerencia h1 {{
            color: #FFFFFF;
            font-size: 1.55rem;
            font-weight: 600;
            letter-spacing: -0.01em;
            margin: 0;
        }}
        div[data-testid="stMetric"] {{
            background-color: {FONDO_TARJETA};
            border: 1px solid {BORDE_LINEA};
            border-top: 3px solid {NAVY_MEDIO};
            padding: 14px 16px;
            border-radius: 6px;
        }}
        div[data-testid="stMetric"] label {{
            color: {GRIS_SUB} !important;
            font-size: 0.82rem !important;
            font-weight: 500 !important;
        }}
        div[data-testid="stMetric"] div[data-testid="stMetricValue"] {{
            color: {GRIS_TEXTO} !important;
            font-weight: 600 !important;
        }}
        section[data-testid="stSidebar"] {{
            background-color: #FFFFFF;
            border-right: 1px solid {BORDE_LINEA};
        }}
        /* Traducir "Choose options" en multiselect a español */
        div[data-baseweb="select"] span[data-testid="stSearchInput"] input::placeholder {{
            color: transparent !important;
        }}
        div[data-baseweb="select"] > div > div:first-child {{
            position: relative;
        }}
        div[data-baseweb="select"] div:contains("Choose options") {{
            font-size: 0 !important;
        }}
        div[data-baseweb="select"] div:has(> span):not(:has(span[data-baseweb="tag"])) span {{
            visibility: hidden;
            position: relative;
        }}
        div[data-baseweb="select"] div:has(> span):not(:has(span[data-baseweb="tag"])) span::after {{
            content: "Seleccione opciones";
            visibility: visible;
            position: absolute;
            left: 0;
            top: 0;
            color: #616E7C;
            font-size: 0.92rem;
            white-space: nowrap;
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
        .str.replace("S/", "", regex=False)
        .str.replace("$", "", regex=False)
        .str.replace(" ", "", regex=False)
        .str.replace(",", "", regex=False)
    )
    return pd.to_numeric(limpio, errors="coerce").fillna(0.0)


def _buscar_columna(
    columnas: list[str],
    candidatos: list[str],
    palabras_clave: list[str],
    excluir_palabras: list[str] | None = None,
) -> str | None:
    excl = [x.upper() for x in (excluir_palabras or [])]
    mapa_upper = {c.upper().strip(): c for c in columnas}
    for cand in candidatos:
        if cand.upper().strip() in mapa_upper:
            return mapa_upper[cand.upper().strip()]
    for c in columnas:
        c_up = c.upper()
        if any(ex in c_up for ex in excl):
            continue
        if any(pk.upper() in c_up for pk in palabras_clave):
            return c
    return None


@st.cache_data(show_spinner="Cargando información...")
def cargar_datos(ruta_excel: str = "STOCK REAL 26.xlsx") -> tuple[pd.DataFrame, str, str]:
    xl = pd.ExcelFile(ruta_excel, engine="openpyxl")
    hoja = (
        "STOCK REAL 26"
        if "STOCK REAL 26" in xl.sheet_names
        else ("Stock" if "Stock" in xl.sheet_names else xl.sheet_names[0])
    )
    df = pd.read_excel(ruta_excel, sheet_name=hoja, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    col_stock = _buscar_columna(
        list(df.columns),
        ["STOCK REAL", "STOCK TOTAL"],
        ["STOCK"],
    )
    if col_stock is None:
        raise ValueError("No se encontró una columna de STOCK REAL o STOCK TOTAL en el archivo.")

    col_prov = _buscar_columna(
        list(df.columns),
        ["PROVEEDOR", "NOM. PROVEEDOR", "NOMBRE PROVEEDOR"],
        ["PROVEEDOR"],
    )
    col_fact = _buscar_columna(
        list(df.columns),
        [
            "FACTURACION SIN IGV AL 25",
            "FACTURACIÓN SIN IGV AL 25",
            "VENTA",
            "FACTURACION",
            "FACTURACIÓN",
        ],
        ["FACTURACION", "FACTURACIÓN", "VENTA"],
    )
    col_cant_vta = _buscar_columna(
        list(df.columns),
        [
            "CANTIDADES VENDIDAS AL 25",
            "CANTIDAD VENDIDA AL 25",
            "CANTIDAD VENDIDAS AL 25",
            "CANTIDADES VENDIDAS",
            "CANTIDAD VENDIDA",
            "UNIDADES VENDIDAS AL 25",
            "VENTA UNIDADES AL 25",
        ],
        ["CANTIDAD", "CANTIDADES", "UNIDADES"],
        excluir_palabras=["FACTURACION", "FACTURACIÓN", "STOCK"],
    )

    df["STOCK REAL"] = _limpiar_numero(df[col_stock])

    if col_prov is not None:
        df["PROVEEDOR"] = _limpiar_texto(df[col_prov])
        df.loc[df["PROVEEDOR"] == "", "PROVEEDOR"] = "SIN PROVEEDOR"
    else:
        df["PROVEEDOR"] = "SIN PROVEEDOR"

    nombre_fact = col_fact if col_fact else "FACTURACION SIN IGV AL 25"
    df["FACT_NUM"] = _limpiar_numero(df[col_fact]) if col_fact is not None else 0.0

    nombre_cant = col_cant_vta if col_cant_vta else "CANTIDADES VENDIDAS AL 25"
    df["CANT_VTA_NUM"] = _limpiar_numero(df[col_cant_vta]) if col_cant_vta is not None else 0.0

    columnas_texto = [
        "TIPO DE DESTINO",
        "Nombre Centro Destino",
        "COD. PRODUCTO",
        "NOM. PRODUCTO",
        "ESTADO",
    ]
    for col in columnas_texto:
        if col not in df.columns:
            df[col] = "SIN DATO"
        else:
            df[col] = _limpiar_texto(df[col])
            df.loc[df[col] == "", col] = "SIN DATO"

    cols_finales = [
        "TIPO DE DESTINO",
        "Nombre Centro Destino",
        "PROVEEDOR",
        "COD. PRODUCTO",
        "NOM. PRODUCTO",
        "ESTADO",
        "STOCK REAL",
        "FACT_NUM",
        "CANT_VTA_NUM",
    ]
    return df[cols_finales].copy(), nombre_fact, nombre_cant


def consolidar_por_sku(
    df_in: pd.DataFrame,
    nombre_fact: str,
    nombre_cant: str,
) -> pd.DataFrame:
    """
    - Suma STOCK REAL por SKU.
    - NO suma FACT_NUM ni CANT_VTA_NUM repetidas entre filas del mismo SKU:
      toma el primer valor correspondiente al SKU sin duplicar.
    """
    if df_in.empty:
        return pd.DataFrame(
            columns=[
                "COD. PRODUCTO",
                "NOM. PRODUCTO",
                "PROVEEDOR",
                "ESTADO",
                "STOCK REAL",
                nombre_cant,
                nombre_fact,
            ]
        )

    resumen = (
        df_in.groupby("COD. PRODUCTO", as_index=False)
        .agg(
            {
                "NOM. PRODUCTO": "first",
                "PROVEEDOR": lambda s: " / ".join(
                    sorted(set(x for x in s if x and x != "SIN PROVEEDOR"))
                )
                or "SIN PROVEEDOR",
                "ESTADO": lambda s: " / ".join(
                    sorted(set(x for x in s if x and x != "SIN DATO"))
                )
                or "SIN DATO",
                "STOCK REAL": "sum",
                "CANT_VTA_NUM": "first",
                "FACT_NUM": "first",
            }
        )
        .rename(
            columns={
                "CANT_VTA_NUM": nombre_cant,
                "FACT_NUM": nombre_fact,
            }
        )
    )
    cols_orden = [
        "COD. PRODUCTO",
        "NOM. PRODUCTO",
        "PROVEEDOR",
        "ESTADO",
        "STOCK REAL",
        nombre_cant,
        nombre_fact,
    ]
    return resumen[cols_orden].sort_values("STOCK REAL", ascending=False).reset_index(drop=True)


def formatear_para_pantalla(
    df_tabla: pd.DataFrame,
    cols_moneda: list[str],
    cols_numero: list[str],
) -> pd.DataFrame:
    out = df_tabla.copy()
    for c in cols_moneda:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0).map(
                lambda x: f"S/ {x:,.2f}"
            )
    for c in cols_numero:
        if c in out.columns:
            out[c] = pd.to_numeric(out[c], errors="coerce").fillna(0.0).map(
                lambda x: f"{x:,.2f}"
            )
    return out


def generar_excel_con_tabla(
    df_export: pd.DataFrame,
    cols_sumar_excel: list[str],
    cols_moneda: list[str],
    cols_numero: list[str],
) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Reporte Stock"

    encabezados = list(df_export.columns)
    ws.append(encabezados)

    fill_header = PatternFill(start_color="162A3E", end_color="162A3E", fill_type="solid")
    font_header = Font(color="FFFFFF", bold=True, size=10)

    for col_idx in range(1, len(encabezados) + 1):
        celda = ws.cell(row=1, column=col_idx)
        celda.fill = fill_header
        celda.font = font_header
        celda.alignment = Alignment(horizontal="center", vertical="center")

    for row in df_export.itertuples(index=False, name=None):
        ws.append(list(row))

    n_filas = len(df_export)
    n_cols = len(encabezados)
    ultima_fila = n_filas + 1

    if n_filas > 0:
        for c_name in cols_moneda:
            if c_name in encabezados:
                c_idx = encabezados.index(c_name) + 1
                for r in range(2, ultima_fila + 1):
                    ws.cell(row=r, column=c_idx).number_format = '"S/ "#,##0.00'

        for c_name in cols_numero:
            if c_name in encabezados:
                c_idx = encabezados.index(c_name) + 1
                for r in range(2, ultima_fila + 1):
                    ws.cell(row=r, column=c_idx).number_format = "#,##0.00"

        cols_sumar_validas = [c for c in cols_sumar_excel if c in encabezados]
        if cols_sumar_validas:
            fila_total = ultima_fila + 1
            ws.cell(row=fila_total, column=1, value="TOTAL").font = Font(
                bold=True, color="162A3E"
            )
            for c_name in cols_sumar_validas:
                c_idx = encabezados.index(c_name) + 1
                col_letra = get_column_letter(c_idx)
                celda_total = ws.cell(
                    row=fila_total,
                    column=c_idx,
                    value=f"=SUM({col_letra}2:{col_letra}{ultima_fila})",
                )
                celda_total.font = Font(bold=True, color="162A3E")
                celda_total.number_format = (
                    '"S/ "#,##0.00' if c_name in cols_moneda else "#,##0.00"
                )
                celda_total.fill = PatternFill(
                    start_color="E6ECF2", end_color="E6ECF2", fill_type="solid"
                )

        ultima_letra = get_column_letter(n_cols)
        ref_tabla = f"A1:{ultima_letra}{ultima_fila}"
        nombre_limpio = re.sub(r"[^A-Za-z0-9_]", "_", "TablaStock")
        tabla = Table(displayName=nombre_limpio, ref=ref_tabla)
        estilo = TableStyleInfo(
            name="TableStyleMedium2",
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
        ws.column_dimensions[letra].width = min(max(max_len + 4, 15), 45)

    salida = BytesIO()
    wb.save(salida)
    salida.seek(0)
    return salida.getvalue()


st.markdown(
    f"""
    <div class="encabezado-gerencia">
        <h1>{TITULO_REPORTE}</h1>
    </div>
    """,
    unsafe_allow_html=True,
)

try:
    df, nombre_fact, nombre_cant = cargar_datos("STOCK REAL 26.xlsx")
except Exception as e:
    st.error(f"No se pudo cargar el archivo de datos: {e}")
    st.stop()

# Filtros en barra lateral
st.sidebar.header("Filtros")

proveedores_sel = st.sidebar.multiselect(
    "Proveedor",
    options=sorted(df["PROVEEDOR"].unique()),
    default=[],
    placeholder="Seleccione opciones",
)

tipos_sel = st.sidebar.multiselect(
    "Tipo de destino",
    options=sorted(df["TIPO DE DESTINO"].unique()),
    default=[],
    placeholder="Seleccione opciones",
)

estados_sel = st.sidebar.multiselect(
    "Estado",
    options=sorted(df["ESTADO"].unique()),
    default=[],
    placeholder="Seleccione opciones",
)

df_base_centros = df.copy()
if proveedores_sel:
    df_base_centros = df_base_centros[df_base_centros["PROVEEDOR"].isin(proveedores_sel)]
if tipos_sel:
    df_base_centros = df_base_centros[df_base_centros["TIPO DE DESTINO"].isin(tipos_sel)]
if estados_sel:
    df_base_centros = df_base_centros[df_base_centros["ESTADO"].isin(estados_sel)]

centros_sel = st.sidebar.multiselect(
    "Centro destino",
    options=sorted(df_base_centros["Nombre Centro Destino"].unique()),
    default=[],
    placeholder="Seleccione opciones",
)

sku_buscar = st.sidebar.text_input(
    "Buscar SKU (código o descripción)",
    placeholder="Ej. 100101001001 o CARBON",
).strip()

solo_con_stock = st.sidebar.checkbox("Solo registros con stock distinto de 0", value=False)

vista_tabla = st.sidebar.radio(
    "Vista de tabla",
    options=[
        "Consolidado por SKU (Stock sumado / Venta y Facturación sin duplicar)",
        "Detalle por fila (sin sumar venta ni facturación)",
    ],
    index=0,
)

# Aplicar filtros
df_filtrado = df.copy()
if proveedores_sel:
    df_filtrado = df_filtrado[df_filtrado["PROVEEDOR"].isin(proveedores_sel)]
if tipos_sel:
    df_filtrado = df_filtrado[df_filtrado["TIPO DE DESTINO"].isin(tipos_sel)]
if estados_sel:
    df_filtrado = df_filtrado[df_filtrado["ESTADO"].isin(estados_sel)]
if centros_sel:
    df_filtrado = df_filtrado[df_filtrado["Nombre Centro Destino"].isin(centros_sel)]
if sku_buscar:
    mask_sku = (
        df_filtrado["COD. PRODUCTO"].str.contains(sku_buscar, case=False, na=False)
        | df_filtrado["NOM. PRODUCTO"].str.contains(sku_buscar, case=False, na=False)
    )
    df_filtrado = df_filtrado[mask_sku]
if solo_con_stock:
    df_filtrado = df_filtrado[df_filtrado["STOCK REAL"] != 0]

df_sku = consolidar_por_sku(df_filtrado, nombre_fact, nombre_cant)

# KPIs principales
stock_total = float(df_filtrado["STOCK REAL"].sum())
cant_vta_sin_duplicar = float(df_sku[nombre_cant].sum()) if not df_sku.empty else 0.0
fact_sin_duplicar = float(df_sku[nombre_fact].sum()) if not df_sku.empty else 0.0
skus_unicos = len(df_sku)

c1, c2, c3, c4 = st.columns(4)
c1.metric("Stock Real Total", f"{stock_total:,.2f}")
c2.metric(f"{nombre_cant} (por SKU)", f"{cant_vta_sin_duplicar:,.2f}")
c3.metric(f"{nombre_fact} (por SKU)", f"S/ {fact_sin_duplicar:,.2f}")
c4.metric("SKUs únicos", f"{skus_unicos:,}")

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
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                x=alt.X("STOCK REAL:Q", title="Stock Real", axis=alt.Axis(format=",.0f")),
                y=alt.Y("TIPO DE DESTINO:N", sort="-x", title="Tipo de Destino"),
                color=alt.value(NAVY_MEDIO),
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
            .mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4)
            .encode(
                x=alt.X("STOCK REAL:Q", title="Stock Real", axis=alt.Axis(format=",.0f")),
                y=alt.Y("Nombre Centro Destino:N", sort="-x", title="Centro Destino"),
                color=alt.value(AZUL_ACERO),
                tooltip=[
                    alt.Tooltip("Nombre Centro Destino:N", title="Centro"),
                    alt.Tooltip("STOCK REAL:Q", title="Stock Real", format=",.2f"),
                ],
            )
            .properties(height=300)
        )
        st.altair_chart(chart_centro, use_container_width=True)

    st.subheader("Detalle por SKU y Facturación")

    cols_moneda_lista = [nombre_fact]
    cols_numero_lista = ["STOCK REAL", nombre_cant]

    if vista_tabla.startswith("Consolidado"):
        tabla_base = df_sku
        cols_sumar_xls = ["STOCK REAL", nombre_cant, nombre_fact]
        nombre_xls = "REPORTE_STOCK_VENTA_CANTIDADES_25.xlsx"
        st.caption(
            f"Consolidado por SKU: **STOCK REAL** suma todos los registros del SKU; "
            f"**{nombre_cant}** y **{nombre_fact}** se asignan por SKU sin sumar duplicados."
        )
    else:
        tabla_base = df_filtrado.rename(
            columns={
                "FACT_NUM": nombre_fact,
                "CANT_VTA_NUM": nombre_cant,
            }
        )
        cols_sumar_xls = ["STOCK REAL"]
        nombre_xls = "REPORTE_STOCK_DETALLE_FILAS.xlsx"
        st.caption(
            f"Detalle por fila: la tabla en Excel suma **STOCK REAL**, mientras **{nombre_cant}** y "
            f"**{nombre_fact}** no se suman en el total para evitar duplicidades por SKU."
        )

    excel_bytes = generar_excel_con_tabla(
        tabla_base,
        cols_sumar_excel=cols_sumar_xls,
        cols_moneda=cols_moneda_lista,
        cols_numero=cols_numero_lista,
    )
    st.download_button(
        label="Exportar a Excel (Tabla y Fórmula SUMA)",
        data=excel_bytes,
        file_name=nombre_xls,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    tabla_pantalla = formatear_para_pantalla(
        tabla_base.head(2000),
        cols_moneda=cols_moneda_lista,
        cols_numero=cols_numero_lista,
    )
    st.dataframe(tabla_pantalla, use_container_width=True, hide_index=True)

    if len(tabla_base) > 2000:
        st.caption(
            f"Mostrando los primeros 2,000 registros en pantalla de {len(tabla_base):,}. "
            "La descarga en Excel incluye todos los registros filtrados."
        )
