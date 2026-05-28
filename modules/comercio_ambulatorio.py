# modules/comercio_ambulatorio.py

import pandas as pd
import plotly.express as px
import streamlit as st
from pathlib import Path
from utils.google_sheets import get_resoluciones_sheet_or_none, normalize_text

# Paleta de colores por año
YEAR_COLORS = {
    "2023": "#e74c3c",
    "2024": "#3498db",
    "2025": "#2ecc71",
    "2026": "#f39c12",
}

YEAR_ORDER = ["2023", "2024", "2025", "2026"]
COSTO_MENSUAL_ESTIMADO = 30.0
DIA_CORTE_PAGO = 15

MONTH_ORDER = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"
]

MONTH_MAP = {
    1: "Enero",
    2: "Febrero",
    3: "Marzo",
    4: "Abril",
    5: "Mayo",
    6: "Junio",
    7: "Julio",
    8: "Agosto",
    9: "Septiembre",
    10: "Octubre",
    11: "Noviembre",
    12: "Diciembre"
}


def get_spanish_month(month_num):
    return MONTH_MAP.get(month_num, "")


def clean_resolution_dates(series):
    return (
        series.astype(str)
        .str.strip()
        .str.replace("'", "", regex=False)
        .str.replace(r"/+", "/", regex=True)
    )


def extract_resolution_year(series):
    return series.astype(str).str.extract(r"(20\d{2})", expand=False)


def refresh_year_order(df):
    global YEAR_ORDER
    years = sorted(df["AÑO"].dropna().astype(str).unique())
    if years:
        YEAR_ORDER = years
        for idx, year in enumerate(YEAR_ORDER):
            YEAR_COLORS.setdefault(year, px.colors.qualitative.Set2[idx % len(px.colors.qualitative.Set2)])


def load_comercio_ambulatorio_drive_data():
    df_raw = get_resoluciones_sheet_or_none()
    if df_raw is None:
        return None

    required = {"TIPO DE PROCEDIMIENTO", "FECHA RESOLUCION"}
    if not required.issubset(df_raw.columns):
        st.warning("El Sheet no tiene las columnas requeridas para Comercio Ambulatorio.")
        return None

    df = df_raw.copy()
    df["TIPO_NORMALIZADO"] = df["TIPO DE PROCEDIMIENTO"].map(normalize_text)
    df = df[df["TIPO_NORMALIZADO"] == "COMERCIO AMBULATORIO"].copy()
    if df.empty:
        return None

    fecha_limpia = clean_resolution_dates(df["FECHA RESOLUCION"])
    df["FECHA_EMITIDA"] = pd.to_datetime(
        fecha_limpia,
        dayfirst=True,
        errors="coerce",
    )
    df["ANIO_REFERENCIA"] = pd.to_numeric(df.get("PERIODO"), errors="coerce")
    df["ANIO_RESOLUCION"] = pd.to_numeric(
        extract_resolution_year(df.get("RESOLUCION DE SG", "")),
        errors="coerce",
    )
    df["ANIO_CONTEO"] = df["FECHA_EMITIDA"].dt.year.astype("float64")
    df["ANIO_CONTEO"] = df["ANIO_CONTEO"].fillna(df["ANIO_REFERENCIA"]).fillna(df["ANIO_RESOLUCION"])

    mismatch_mask = (
        df["FECHA_EMITIDA"].notna()
        & df["ANIO_RESOLUCION"].notna()
        & (df["FECHA_EMITIDA"].dt.year != df["ANIO_RESOLUCION"])
    )
    df.loc[mismatch_mask, "FECHA_EMITIDA"] = df.loc[mismatch_mask].apply(
        lambda row: row["FECHA_EMITIDA"].replace(year=int(row["ANIO_RESOLUCION"])),
        axis=1,
    )
    df.loc[mismatch_mask, "ANIO_CONTEO"] = df.loc[mismatch_mask, "ANIO_RESOLUCION"].astype("float64")

    df = df.dropna(subset=["ANIO_CONTEO"])
    if df.empty:
        return None

    df["AÑO"] = df["ANIO_CONTEO"].astype(int).astype(str)
    df["MES_NUM"] = df["FECHA_EMITIDA"].dt.month.fillna(13).astype(int)
    df["MES"] = df["MES_NUM"].map(get_spanish_month).fillna("Sin fecha")
    df.loc[df["MES_NUM"] == 13, "MES"] = "Sin fecha"
    df = df.sort_values("FECHA_EMITIDA").reset_index(drop=True)
    df.attrs["source"] = "drive"
    refresh_year_order(df)
    return df


def load_comercio_ambulatorio_data():
    """Carga y procesa los datos de autorizaciones de comercio ambulatorio."""
    drive_df = load_comercio_ambulatorio_drive_data()

    try:
        data_path = Path(__file__).parent.parent / "data" / "comercio_ambulatorio.csv"

        df_raw = pd.read_csv(
            data_path,
            sep=";",
            encoding="utf-8-sig",
            dtype=str
        )

        df_raw.columns = df_raw.columns.str.strip()

        # Caso 1: si el archivo ya viene con FECHA_EMITIDA
        if "FECHA_EMITIDA" in df_raw.columns:
            df = df_raw.copy()

            df["FECHA_EMITIDA"] = pd.to_datetime(
                df["FECHA_EMITIDA"],
                dayfirst=True,
                errors="coerce"
            )

            df = df.dropna(subset=["FECHA_EMITIDA"])
            df["AÑO"] = df["FECHA_EMITIDA"].dt.year.astype(str)

        else:
            # Caso 2: columnas por año: 2023, 2024, 2025, 2026...
            year_cols = [col for col in df_raw.columns if col.strip().isdigit()]

            if not year_cols:
                raise ValueError(
                    "No se encontró la columna 'FECHA_EMITIDA' ni columnas de años como 2023, 2024, 2025, 2026."
                )

            df = df_raw[year_cols].melt(
                var_name="AÑO",
                value_name="FECHA_EMITIDA"
            )

            df["FECHA_EMITIDA"] = df["FECHA_EMITIDA"].astype(str).str.strip()

            df = df[
                df["FECHA_EMITIDA"].notna() &
                (df["FECHA_EMITIDA"] != "") &
                (df["FECHA_EMITIDA"].str.lower() != "nan")
            ]

            df["FECHA_EMITIDA"] = pd.to_datetime(
                df["FECHA_EMITIDA"],
                format="%d/%m/%Y",
                errors="coerce"
            )

            df = df.dropna(subset=["FECHA_EMITIDA"])

        df["AÑO"] = df["AÑO"].astype(str)
        df["MES_NUM"] = df["FECHA_EMITIDA"].dt.month
        df["MES"] = df["MES_NUM"].map(get_spanish_month)

        df = df.sort_values("FECHA_EMITIDA").reset_index(drop=True)

        if drive_df is not None and not drive_df.empty:
            active_year = drive_df["AÑO"].astype(int).max()
            historical_df = df[df["AÑO"].astype(int) < active_year].copy()
            df = pd.concat([historical_df, drive_df], ignore_index=True)
            df = df.sort_values("FECHA_EMITIDA").reset_index(drop=True)
            df.attrs["source"] = "mixed"
        else:
            df.attrs["source"] = "local"

        refresh_year_order(df)

        return df

    except Exception as e:
        st.error(f"🚨 Error al cargar datos: {str(e)}")
        return None


def load_comercio_ambulatorio_recaudacion_data():
    """Carga los datos fijos de recaudación de comercio ambulatorio."""
    data = [
        {"AÑO": "2023", "PERMISOS": 398, "MESES": 12, "COSTO": 30.0, "TOTAL_RECAUDADO": 143280.0},
        {"AÑO": "2024", "PERMISOS": 183, "MESES": 12, "COSTO": 30.0, "TOTAL_RECAUDADO": 65880.0},
        {"AÑO": "2025", "PERMISOS": 125, "MESES": 12, "COSTO": 30.0, "TOTAL_RECAUDADO": 45000.0},
        {"AÑO": "2026", "PERMISOS": 66, "MESES": 4, "COSTO": 30.0, "TOTAL_RECAUDADO": 4350.0},
    ]

    df = pd.DataFrame(data)
    df["AÑO"] = pd.Categorical(df["AÑO"], categories=YEAR_ORDER, ordered=True)
    return df


def calcular_recaudacion_estimada_vigencia(df):
    """Estima recaudacion desde el mes de emision hasta el cierre aplicable."""
    base = df.copy()
    base["FECHA_EMITIDA"] = pd.to_datetime(base["FECHA_EMITIDA"], errors="coerce")
    base["TIENE_FECHA"] = base["FECHA_EMITIDA"].notna()

    base["MES_INICIO_PAGO"] = 0
    con_fecha = base["TIENE_FECHA"]
    base.loc[con_fecha, "MES_INICIO_PAGO"] = base.loc[con_fecha, "FECHA_EMITIDA"].dt.month

    hoy = pd.Timestamp.today()
    anio_actual = hoy.year
    mes_actual = hoy.month
    anio_emision = base["FECHA_EMITIDA"].dt.year

    base["MES_FIN_PAGO"] = 12
    base.loc[anio_emision == anio_actual, "MES_FIN_PAGO"] = mes_actual
    base.loc[anio_emision > anio_actual, "MES_FIN_PAGO"] = 0

    base["MESES_COBRABLES"] = (base["MES_FIN_PAGO"] - base["MES_INICIO_PAGO"] + 1).clip(lower=0, upper=12)
    base.loc[~base["TIENE_FECHA"], "MESES_COBRABLES"] = 0
    base["RECAUDACION_ESTIMADA"] = base["MESES_COBRABLES"] * COSTO_MENSUAL_ESTIMADO
    return base


def grafico_comparativa_meses(df):
    """Gráfico de barras comparativo por meses y años."""
    st.subheader("📊 Comparativa por Meses")

    comparativa = (
        df.groupby(["MES", "MES_NUM", "AÑO"])
        .size()
        .reset_index(name="AUTORIZACIONES")
        .sort_values(["MES_NUM", "AÑO"])
    )

    fig = px.bar(
        comparativa,
        x="MES",
        y="AUTORIZACIONES",
        color="AÑO",
        barmode="group",
        color_discrete_map=YEAR_COLORS,
        category_orders={
            "MES": MONTH_ORDER,
            "AÑO": YEAR_ORDER
        },
        height=450,
        labels={
            "AUTORIZACIONES": "Cantidad de Autorizaciones",
            "MES": "Mes",
            "AÑO": "Año",
        }
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes",
        yaxis_title="Cantidad de Autorizaciones",
        hovermode="x unified",
        legend_title="Año"
    )

    st.plotly_chart(fig, use_container_width=True)


def grafico_crecimiento_mensual(df):
    """Gráfico de líneas de crecimiento mensual por año."""
    st.subheader("📈 Crecimiento Mensual por Año")

    monthly_data = (
        df.groupby(["MES", "MES_NUM", "AÑO"])
        .size()
        .reset_index(name="AUTORIZACIONES")
        .sort_values(["AÑO", "MES_NUM"])
    )

    fig = px.line(
        monthly_data,
        x="MES",
        y="AUTORIZACIONES",
        color="AÑO",
        markers=True,
        line_shape="spline",
        color_discrete_map=YEAR_COLORS,
        category_orders={
            "MES": MONTH_ORDER,
            "AÑO": YEAR_ORDER
        },
        height=450,
        labels={
            "AUTORIZACIONES": "Cantidad de Autorizaciones",
            "MES": "Mes",
            "AÑO": "Año",
        }
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes",
        yaxis_title="Cantidad de Autorizaciones",
        hovermode="x unified",
        legend_title="Año"
    )

    fig.update_traces(
        marker=dict(size=8),
        line=dict(width=3)
    )

    st.plotly_chart(fig, use_container_width=True)


def grafico_comparativa_por_ano(df):
    """Gráfico de barras con totales por año."""
    st.subheader("📅 Total de Autorizaciones por Año")

    anual = (
        df.groupby("AÑO")
        .size()
        .reindex(YEAR_ORDER, fill_value=0)
        .reset_index(name="TOTAL_AUTORIZACIONES")
    )

    fig = px.bar(
        anual,
        x="AÑO",
        y="TOTAL_AUTORIZACIONES",
        color="AÑO",
        text="TOTAL_AUTORIZACIONES",
        color_discrete_map=YEAR_COLORS,
        category_orders={"AÑO": YEAR_ORDER},
        height=350,
        labels={
            "TOTAL_AUTORIZACIONES": "Total de Autorizaciones",
            "AÑO": "Año",
        }
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Total de Autorizaciones",
        showlegend=False
    )

    fig.update_xaxes(type="category")

    fig.update_traces(
        textposition="outside",
        marker_line_color="rgba(0,0,0,0.3)",
        marker_line_width=2
    )

    st.plotly_chart(fig, use_container_width=True)


def grafico_2026_mensual(df):
    """Vista enfocada en el año actual para autorizaciones emitidas por mes."""
    if "2026" not in df["AÑO"].astype(str).unique():
        return

    st.subheader("Autorizaciones 2026 por mes")

    df_2026 = df[df["AÑO"].astype(str) == "2026"].copy()
    mensual = (
        df_2026.groupby(["MES_NUM", "MES"])
        .size()
        .reset_index(name="AUTORIZACIONES")
        .sort_values("MES_NUM")
    )
    mensual["ACUMULADO"] = mensual["AUTORIZACIONES"].cumsum()

    fig = px.bar(
        mensual,
        x="MES",
        y="AUTORIZACIONES",
        text="AUTORIZACIONES",
        color_discrete_sequence=["#f39c12"],
        category_orders={"MES": [*MONTH_ORDER, "Sin fecha"]},
        height=420,
        labels={
            "MES": "Mes",
            "AUTORIZACIONES": "Autorizaciones",
        },
    )
    fig.add_scatter(
        x=mensual["MES"],
        y=mensual["ACUMULADO"],
        mode="lines+markers+text",
        name="Acumulado",
        text=mensual["ACUMULADO"],
        textposition="top center",
        line=dict(color="#0f4c81", width=3),
    )
    fig.update_traces(textposition="outside", selector=dict(type="bar"))
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes",
        yaxis_title="Autorizaciones",
        legend_title="Serie",
    )

    st.plotly_chart(fig, use_container_width=True)


def tabla_resumen(df):
    """Tabla resumen por mes y año."""
    st.subheader("📋 Tabla Resumen: Autorizaciones por Mes y Año")

    resumen = (
        df.groupby(["MES_NUM", "MES", "AÑO"])
        .size()
        .reset_index(name="TOTAL")
    )

    tabla_df = (
        resumen.pivot_table(
            index=["MES_NUM", "MES"],
            columns="AÑO",
            values="TOTAL",
            aggfunc="sum",
            fill_value=0
        )
        .reset_index()
        .sort_values("MES_NUM")
    )

    for year in YEAR_ORDER:
        if year not in tabla_df.columns:
            tabla_df[year] = 0

    tabla_df["Total"] = tabla_df[YEAR_ORDER].sum(axis=1)
    tabla_df = tabla_df[["MES", "2023", "2024", "2025", "2026", "Total"]]
    tabla_df = tabla_df.rename(columns={"MES": "Mes"})

    st.dataframe(
        tabla_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Mes": st.column_config.TextColumn("Mes", width="medium"),
            "2023": st.column_config.NumberColumn("2023", format="%d"),
            "2024": st.column_config.NumberColumn("2024", format="%d"),
            "2025": st.column_config.NumberColumn("2025", format="%d"),
            "2026": st.column_config.NumberColumn("2026", format="%d"),
            "Total": st.column_config.NumberColumn("Total", format="%d"),
        }
    )


def estadisticas_recaudacion_estimada(df):
    estimada = calcular_recaudacion_estimada_vigencia(df)
    total_estimado = float(estimada["RECAUDACION_ESTIMADA"].sum())
    permisos_calculados = int(estimada["TIENE_FECHA"].sum())
    permisos_sin_fecha = int((~estimada["TIENE_FECHA"]).sum())
    promedio_permiso = total_estimado / permisos_calculados if permisos_calculados else 0

    st.subheader("Recaudacion estimada por vigencia")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Estimacion total", f"S/ {total_estimado:,.2f}")
    c2.metric("Permisos calculados", f"{permisos_calculados:,}")
    c3.metric("Promedio por permiso", f"S/ {promedio_permiso:,.2f}")
    c4.metric("Sin fecha", permisos_sin_fecha)

    st.caption(
        "Estimacion: para el año actual se calcula desde el mes de emision hasta el mes actual; "
        "para años cerrados se calcula hasta diciembre. Costo mensual: S/ 30."
    )


def grafico_recaudacion_estimada_por_ano(df):
    estimada = calcular_recaudacion_estimada_vigencia(df)
    anual = (
        estimada.groupby("AÑO")["RECAUDACION_ESTIMADA"]
        .sum()
        .reindex(YEAR_ORDER, fill_value=0)
        .reset_index()
    )

    fig = px.bar(
        anual,
        x="AÑO",
        y="RECAUDACION_ESTIMADA",
        color="AÑO",
        text="RECAUDACION_ESTIMADA",
        color_discrete_map=YEAR_COLORS,
        category_orders={"AÑO": YEAR_ORDER},
        height=420,
        labels={"AÑO": "Año", "RECAUDACION_ESTIMADA": "Recaudacion estimada (S/)"},
    )
    fig.update_traces(texttemplate="S/ %{y:,.2f}", textposition="outside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Recaudacion estimada (S/)",
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True)


def grafico_recaudacion_estimada_mensual(df):
    estimada = calcular_recaudacion_estimada_vigencia(df)
    mensual = (
        estimada[estimada["TIENE_FECHA"]]
        .groupby(["AÑO", "MES_NUM", "MES"])
        .agg(
            PERMISOS=("FECHA_EMITIDA", "size"),
            RECAUDACION_ESTIMADA=("RECAUDACION_ESTIMADA", "sum"),
        )
        .reset_index()
        .sort_values(["AÑO", "MES_NUM"])
    )

    if mensual.empty:
        return

    fig = px.bar(
        mensual,
        x="MES",
        y="RECAUDACION_ESTIMADA",
        color="AÑO",
        barmode="group",
        text="RECAUDACION_ESTIMADA",
        color_discrete_map=YEAR_COLORS,
        category_orders={"MES": MONTH_ORDER, "AÑO": YEAR_ORDER},
        height=450,
        labels={
            "MES": "Mes de emision",
            "AÑO": "Año",
            "RECAUDACION_ESTIMADA": "Recaudacion estimada (S/)",
        },
    )
    fig.update_traces(texttemplate="S/ %{y:,.0f}", textposition="outside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes de emision",
        yaxis_title="Recaudacion estimada (S/)",
        legend_title="Año",
    )
    st.plotly_chart(fig, use_container_width=True)


def tabla_recaudacion_estimada(df):
    estimada = calcular_recaudacion_estimada_vigencia(df)
    resumen = (
        estimada.groupby(["AÑO", "MES_NUM", "MES"])
        .agg(
            Permisos=("FECHA_EMITIDA", "size"),
            Meses_cobrables=("MESES_COBRABLES", "sum"),
            Recaudacion_estimada=("RECAUDACION_ESTIMADA", "sum"),
        )
        .reset_index()
        .sort_values(["AÑO", "MES_NUM"])
        .rename(columns={"AÑO": "Año", "MES": "Mes"})
    )
    resumen = resumen[["Año", "Mes", "Permisos", "Meses_cobrables", "Recaudacion_estimada"]]

    st.dataframe(
        resumen,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año"),
            "Mes": st.column_config.TextColumn("Mes"),
            "Permisos": st.column_config.NumberColumn("Permisos", format="%d"),
            "Meses_cobrables": st.column_config.NumberColumn("Meses cobrables", format="%d"),
            "Recaudacion_estimada": st.column_config.NumberColumn("Recaudacion estimada", format="S/ %.2f"),
        },
    )


def estadisticas_generales(df):
    """Muestra KPIs generales de autorizaciones."""
    st.subheader("📊 Estadísticas Generales de Autorizaciones")

    c1, c2, c3, c4 = st.columns(4)

    total_autorizaciones = len(df)
    total_anios = df["AÑO"].nunique()

    mes_max = (
        df.groupby("MES")
        .size()
        .sort_values(ascending=False)
        .index[0]
    )

    promedio_mes = (
        df.groupby(["AÑO", "MES_NUM"])
        .size()
        .mean()
    )

    c1.metric("📜 Total Autorizaciones", total_autorizaciones)
    c2.metric("📅 Años", total_anios)
    c3.metric("🏆 Mes Más Activo", mes_max)
    c4.metric("📈 Promedio/Mes", f"{promedio_mes:.1f}")


def estadisticas_recaudacion(recaud_df):
    """Muestra KPIs generales de recaudación."""
    st.subheader("💰 Estadísticas Generales de Recaudación")

    c1, c2, c3, c4 = st.columns(4)

    total_recaudado = float(recaud_df["TOTAL_RECAUDADO"].sum())
    anio_max_recaudacion = recaud_df.loc[recaud_df["TOTAL_RECAUDADO"].idxmax(), "AÑO"]
    promedio_anual = recaud_df["TOTAL_RECAUDADO"].mean()
    costo_mensual = recaud_df["COSTO"].iloc[0]

    c1.metric("💵 Total Recaudado", f"S/ {total_recaudado:,.2f}")
    c2.metric("🏆 Año con Mayor Recaudación", str(anio_max_recaudacion))
    c3.metric("📈 Promedio Anual", f"S/ {promedio_anual:,.2f}")
    c4.metric("🧾 Costo Mensual", f"S/ {costo_mensual:,.2f}")


def grafico_recaudacion_por_ano(recaud_df):
    """Gráfico de barras de recaudación por año."""
    st.subheader("💰 Recaudación por Año")

    fig = px.bar(
        recaud_df,
        x="AÑO",
        y="TOTAL_RECAUDADO",
        color="AÑO",
        text="TOTAL_RECAUDADO",
        color_discrete_map=YEAR_COLORS,
        category_orders={"AÑO": YEAR_ORDER},
        height=350,
        labels={
            "AÑO": "Año",
            "TOTAL_RECAUDADO": "Recaudación Total (S/)"
        }
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Recaudación Total (S/)",
        showlegend=False
    )

    fig.update_xaxes(type="category")

    fig.update_traces(
        texttemplate="S/ %{y:,.2f}",
        textposition="outside",
        marker_line_color="rgba(0,0,0,0.3)",
        marker_line_width=2
    )

    st.plotly_chart(fig, use_container_width=True)


def grafico_permisos_vs_recaudacion(recaud_df):
    """Gráfico comparativo entre permisos y recaudación."""
    st.subheader("📈 Permisos vs Recaudación")

    df_chart = recaud_df.copy()
    df_chart["AÑO"] = df_chart["AÑO"].astype(str)

    fig = px.scatter(
        df_chart,
        x="PERMISOS",
        y="TOTAL_RECAUDADO",
        color="AÑO",
        size="PERMISOS",
        text="AÑO",
        color_discrete_map=YEAR_COLORS,
        category_orders={"AÑO": YEAR_ORDER},
        height=420,
        labels={
            "PERMISOS": "Cantidad de Permisos",
            "TOTAL_RECAUDADO": "Recaudación Total (S/)"
        }
    )

    fig.update_traces(
        textposition="top center",
        marker=dict(line=dict(width=1, color="rgba(0,0,0,0.3)"))
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Cantidad de Permisos",
        yaxis_title="Recaudación Total (S/)"
    )

    st.plotly_chart(fig, use_container_width=True)


def tabla_recaudacion(recaud_df):
    """Tabla resumen de recaudación."""
    st.subheader("📋 Tabla Resumen de Recaudación")

    tabla_df = recaud_df.copy().rename(columns={
        "AÑO": "Año",
        "PERMISOS": "Permisos",
        "MESES": "Meses",
        "COSTO": "Costo",
        "TOTAL_RECAUDADO": "Total Recaudado"
    })

    st.dataframe(
        tabla_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año", width="small"),
            "Permisos": st.column_config.NumberColumn("Permisos", format="%d"),
            "Meses": st.column_config.NumberColumn("Meses", format="%d"),
            "Costo": st.column_config.NumberColumn("Costo", format="S/ %.2f"),
            "Total Recaudado": st.column_config.NumberColumn("Total Recaudado", format="S/ %.2f"),
        }
    )


def observaciones(df):
    """Muestra observaciones automáticas del comportamiento anual y mensual."""
    st.subheader("📝 Observaciones de Autorizaciones")

    total_anual = (
        df.groupby("AÑO")
        .size()
        .reindex(YEAR_ORDER, fill_value=0)
    )

    mes_general = (
        df.groupby(["MES_NUM", "MES"])
        .size()
        .reset_index(name="TOTAL")
        .sort_values(["TOTAL", "MES_NUM"], ascending=[False, True])
        .iloc[0]
    )

    pico_por_anio = (
        df.groupby(["AÑO", "MES_NUM", "MES"])
        .size()
        .reset_index(name="TOTAL")
        .sort_values(["AÑO", "TOTAL", "MES_NUM"], ascending=[True, False, True])
        .drop_duplicates(subset=["AÑO"])
        .sort_values("AÑO")
    )

    def obtener_pico(anio):
        fila = pico_por_anio[pico_por_anio["AÑO"] == anio]
        return fila.iloc[0] if not fila.empty else None

    pico_2023 = obtener_pico("2023")
    pico_2024 = obtener_pico("2024")
    pico_2025 = obtener_pico("2025")
    pico_2026 = obtener_pico("2026")

    def variacion_pct(base, actual):
        if base == 0:
            return None
        return ((actual - base) / base) * 100

    var_23_24 = variacion_pct(total_anual.get("2023", 0), total_anual.get("2024", 0))
    var_24_25 = variacion_pct(total_anual.get("2024", 0), total_anual.get("2025", 0))
    var_25_26 = variacion_pct(total_anual.get("2025", 0), total_anual.get("2026", 0))

    nota_2026 = ""
    df_2026 = df[df["AÑO"] == "2026"]
    if not df_2026.empty:
        ultimo_mes_2026 = int(df_2026["MES_NUM"].max())
        ultimo_mes_nombre = MONTH_MAP.get(ultimo_mes_2026, "")
        if ultimo_mes_2026 < 12:
            nota_2026 = (
                f"- El año **2026** presenta información parcial hasta **{ultimo_mes_nombre}**, "
                "por lo que su comparación con años completos debe interpretarse con cautela.\n"
            )

    texto = (
        f"- En el periodo analizado, el año con mayor número de autorizaciones emitidas fue "
        f"**{total_anual.idxmax()}**, con **{int(total_anual.max())}** registros.\n"
        f"- El mes con mayor concentración de autorizaciones en todo el periodo fue "
        f"**{mes_general['MES']}**, con **{int(mes_general['TOTAL'])}** registros acumulados.\n"
    )

    if pico_2023 is not None:
        texto += (
            f"- En **2023**, el mes con mayor número de autorizaciones fue "
            f"**{pico_2023['MES']}**, con **{int(pico_2023['TOTAL'])}** registros.\n"
        )

    if pico_2024 is not None:
        texto += (
            f"- En **2024**, el mes con mayor número de autorizaciones fue "
            f"**{pico_2024['MES']}**, con **{int(pico_2024['TOTAL'])}** registros.\n"
        )

    if pico_2025 is not None:
        texto += (
            f"- En **2025**, el mes con mayor número de autorizaciones fue "
            f"**{pico_2025['MES']}**, con **{int(pico_2025['TOTAL'])}** registros.\n"
        )

    if pico_2026 is not None:
        texto += (
            f"- En **2026**, el mes con mayor número de autorizaciones fue "
            f"**{pico_2026['MES']}**, con **{int(pico_2026['TOTAL'])}** registros.\n"
        )

    if var_23_24 is not None:
        tendencia = "disminución" if var_23_24 < 0 else "incremento"
        texto += (
            f"- Entre **2023 y 2024** se observa una **{tendencia}** de "
            f"**{abs(var_23_24):.1f}%** en el total de autorizaciones emitidas.\n"
        )

    if var_24_25 is not None:
        tendencia = "disminución" if var_24_25 < 0 else "incremento"
        texto += (
            f"- Entre **2024 y 2025** se observa una **{tendencia}** de "
            f"**{abs(var_24_25):.1f}%** en el total de autorizaciones emitidas.\n"
        )

    if var_25_26 is not None:
        tendencia = "disminución" if var_25_26 < 0 else "incremento"
        texto += (
            f"- Entre **2025 y 2026** se observa una **{tendencia}** de "
            f"**{abs(var_25_26):.1f}%** en el total registrado.\n"
        )

    texto += nota_2026

    st.info(texto)


def observaciones_recaudacion(recaud_df):
    """Muestra observaciones automáticas de la recaudación."""
    st.subheader("📝 Observaciones de Recaudación")

    mayor_recaudacion = recaud_df.loc[recaud_df["TOTAL_RECAUDADO"].idxmax()]
    menor_recaudacion = recaud_df.loc[recaud_df["TOTAL_RECAUDADO"].idxmin()]

    def variacion_pct(base, actual):
        if base == 0:
            return None
        return ((actual - base) / base) * 100

    rec_2023 = recaud_df.loc[recaud_df["AÑO"] == "2023", "TOTAL_RECAUDADO"].iloc[0]
    rec_2024 = recaud_df.loc[recaud_df["AÑO"] == "2024", "TOTAL_RECAUDADO"].iloc[0]
    rec_2025 = recaud_df.loc[recaud_df["AÑO"] == "2025", "TOTAL_RECAUDADO"].iloc[0]
    rec_2026 = recaud_df.loc[recaud_df["AÑO"] == "2026", "TOTAL_RECAUDADO"].iloc[0]

    var_23_24 = variacion_pct(rec_2023, rec_2024)
    var_24_25 = variacion_pct(rec_2024, rec_2025)
    var_25_26 = variacion_pct(rec_2025, rec_2026)

    total_recaudado = recaud_df["TOTAL_RECAUDADO"].sum()

    texto = (
        f"- La recaudación total del periodo asciende a **S/ {total_recaudado:,.2f}**.\n"
        f"- El año con mayor recaudación fue **{mayor_recaudacion['AÑO']}**, con **S/ {mayor_recaudacion['TOTAL_RECAUDADO']:,.2f}**.\n"
        f"- El año con menor recaudación fue **{menor_recaudacion['AÑO']}**, con **S/ {menor_recaudacion['TOTAL_RECAUDADO']:,.2f}**.\n"
    )

    if var_23_24 is not None:
        tendencia = "disminución" if var_23_24 < 0 else "incremento"
        texto += (
            f"- Entre **2023 y 2024** se registra una **{tendencia}** de "
            f"**{abs(var_23_24):.1f}%** en la recaudación.\n"
        )

    if var_24_25 is not None:
        tendencia = "disminución" if var_24_25 < 0 else "incremento"
        texto += (
            f"- Entre **2024 y 2025** se registra una **{tendencia}** de "
            f"**{abs(var_24_25):.1f}%** en la recaudación.\n"
        )

    if var_25_26 is not None:
        tendencia = "disminución" if var_25_26 < 0 else "incremento"
        texto += (
            f"- Entre **2025 y 2026** se registra una **{tendencia}** de "
            f"**{abs(var_25_26):.1f}%** en la recaudación acumulada.\n"
        )

    texto += (
        "- El valor consignado para **2026** corresponde únicamente a **3 meses**, por lo que no resulta directamente comparable con años completos.\n"
        "- La recaudación presentada se basa en el cuadro consolidado proporcionado para permisos, meses, costo y total anual."
    )

    st.info(texto)


def show_comercio_ambulatorio_module():
    """Módulo completo de Comercio Ambulatorio."""
    st.header("📍 Módulo de Autorizaciones de Comercio Ambulatorio")
    st.markdown("---")

    with st.spinner("🔍 Cargando datos..."):
        df = load_comercio_ambulatorio_data()

    if df is None or df.empty:
        st.error("No se pudieron cargar los datos.")
        return

    fuente = df.attrs.get("source")

    if fuente == "mixed":
        st.success("Historico local conservado y ano actual actualizado desde Google Drive.")
    elif fuente == "drive":
        st.success("Datos actualizados desde Google Drive: autorizaciones emitidas por fecha de resolucion.")

    estadisticas_generales(df)
    st.markdown("---")

    grafico_comparativa_meses(df)
    st.markdown("---")

    grafico_crecimiento_mensual(df)
    st.markdown("---")

    grafico_2026_mensual(df)
    st.markdown("---")

    grafico_comparativa_por_ano(df)
    st.markdown("---")

    tabla_resumen(df)
    st.markdown("---")

    observaciones(df)
