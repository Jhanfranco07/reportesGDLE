import re
from html import escape
from pathlib import Path

import pandas as pd
import plotly.express as px
import streamlit as st
from utils.google_sheets import (
    extract_google_sheet_id,
    get_resoluciones_sheet_or_none,
    get_secret_value,
    load_resoluciones_sheet,
    normalize_column_name,
    normalize_text,
    open_google_worksheet,
    parse_money_series,
    read_google_worksheet_with_rows,
)

# Colores por período
YEAR_COLORS = {
    "2023": "#e74c3c",
    "2024": "#3498db",
    "2025": "#2ecc71",
    "2026 (Ene-Abr)": "#f39c12",
}

YEAR_ORDER = ["2023", "2024", "2025", "2026 (Ene-Abr)"]

RISK_COLORS = {
    "MEDIO": "#3498db",
    "ALTO": "#f39c12",
    "MUY ALTO": "#e74c3c",
    "ALTOS Y MUY ALTOS": "#e74c3c",
}

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
    12: "Diciembre",
}

DRIVE_TABS = ["RESOLUCIONES 2023", "RESOLUCIONES 2024", "RESOLUCIONES 2025", "RESOLUCIONES 2026"]

PRIMARY_LICENSE_PROCEDURES = {
    "LICENCIA TEMPORAL",
    "LICENCIA INDETERMINADA",
}

TRACKED_PROCEDURES = {
    *PRIMARY_LICENSE_PROCEDURES,
    "LICENCIA DE FUNCIONAMIENTO",
    "TRANSFERENCIA DE LICENCIA DE FUNCIONAMIENTO",
    "DUPLICADO DE LICENCIA DE FUNCIONAMIENTO",
}

PROCEDURE_LABELS = {
    "LICENCIA TEMPORAL": "Licencia temporal",
    "LICENCIA INDETERMINADA": "Licencia indeterminada",
    "LICENCIA DE FUNCIONAMIENTO": "Improcedente con pago",
    "TRANSFERENCIA DE LICENCIA DE FUNCIONAMIENTO": "Transferencia",
    "DUPLICADO DE LICENCIA DE FUNCIONAMIENTO": "Duplicado",
}

PROCEDURE_COLORS = {
    "Licencia temporal": "#3498db",
    "Licencia indeterminada": "#2ecc71",
    "Improcedente con pago": "#e74c3c",
    "Transferencia": "#f39c12",
    "Duplicado": "#7f8c8d",
}


def apply_chart_style(fig, legend_title=None, height=None):
    if height is not None:
        fig.update_layout(height=height)
    fig.update_layout(
        template="plotly_white",
        plot_bgcolor="#ffffff",
        paper_bgcolor="#ffffff",
        font=dict(color="#344054", size=12),
        margin=dict(l=26, r=42, t=42, b=44),
        legend=dict(
            title=legend_title,
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1,
        ),
        hoverlabel=dict(bgcolor="#0f3447", font_color="#ffffff"),
        uniformtext_minsize=10,
        uniformtext_mode="show",
    )
    fig.update_xaxes(
        showgrid=True,
        gridcolor="#e6edf5",
        zeroline=False,
        automargin=True,
        tickfont=dict(size=11),
    )
    fig.update_yaxes(
        showgrid=False,
        zeroline=False,
        automargin=True,
        tickfont=dict(size=11),
    )
    return fig


def color_scale_for_values(values, colorscale="Tealgrn"):
    numeric = pd.to_numeric(pd.Series(values), errors="coerce").fillna(0)
    if numeric.empty:
        return []
    minimum = float(numeric.min())
    maximum = float(numeric.max())
    if maximum == minimum:
        positions = [0.72 for _ in numeric]
    else:
        positions = ((numeric - minimum) / (maximum - minimum) * 0.72 + 0.22).tolist()
    return px.colors.sample_colorscale(colorscale, positions)


def polish_bar_traces(fig, text_color="#344054"):
    fig.update_traces(
        marker_line_color="rgba(15, 52, 71, 0.34)",
        marker_line_width=1.15,
        opacity=0.96,
        textfont=dict(size=12, color=text_color),
        cliponaxis=False,
    )
    return fig

DATE_COLUMNS = ["FECHA RESOLUCION", "FECHA RESOLUC.", "FECHA RESOLUC", "FEC.RESOLUC.", "FEC.RESOLUC"]
ZONE_REFERENCE_PATH = Path("script/data/LIC - 2023-2026 - VIGENTES - EMISION.xlsx")
ZONE_REFERENCE_SHEET = "DATA 2023-2024-2025-2026"
ZONE_LABELS = {
    "1": "ZONA 1: PACHACAMAC HISTORICO",
    "2": "ZONA 2: PAUL POBLET LIND",
    "3": "ZONA 3: CENTROS POBLADOS RURALES",
    "4": "ZONA 4: JOSE GALVEZ BARRENECHEA",
    "5": "ZONA 5: MANCHAY",
}
ZONE_ORDER = list(ZONE_LABELS.values()) + ["SIN ZONA"]
ZONE_COLORS = {
    "ZONA 1: PACHACAMAC HISTORICO": "#1f77b4",
    "ZONA 2: PAUL POBLET LIND": "#8e44ad",
    "ZONA 3: CENTROS POBLADOS RURALES": "#16a085",
    "ZONA 4: JOSE GALVEZ BARRENECHEA": "#f39c12",
    "ZONA 5: MANCHAY": "#2f9e44",
    "SIN ZONA": "#7f8c8d",
    "OTRAS ZONAS": "#9b59b6",
    "MANCHAY": "#2f9e44",
    "PACHACAMAC": "#1f77b4",
    "JOSE GALVEZ": "#f39c12",
}
ADDRESS_UPDATE_CONFIGS = {
    "RESOLUCIONES 2026": [
        Path("script/data/BASE DE DATOS 2026 - REGISTRO.xlsx"),
        Path("script/data/BASE DE DATOS 2025 - REGISTRO.xlsx"),
    ],
    "RESOLUCIONES 2025": [
        Path("script/data/BASE DE DATOS 2025 - REGISTRO.xlsx"),
        Path("script/data/BASE DE DATOS 2024 - REGISTRO.xlsx"),
    ],
    "RESOLUCIONES 2024": [
        Path("script/data/BASE DE DATOS 2024 - REGISTRO.xlsx"),
        Path("script/data/BASE DE DATOS 2023 - REGISTRO.xlsx"),
    ],
    "RESOLUCIONES 2023": [
        Path("script/data/BASE DE DATOS 2023 - REGISTRO.xlsx"),
    ],
}
ADDRESS_UPDATE_TAB = "RESOLUCIONES 2026"
LICENSE_SHEET_ID_STATE_KEY = "licencias_update_sheet_id"

EXPEDIENTE_COLUMNS = [
    "EXPEDIENTE / D.S.",
    "EXPEDIENTE / DS",
    "EDIENTE / D.S.",
    "EDIENTE / DS",
    "EDIENTE",
    "EXPEDIENTE",
    "EXPEDIENTES",
    "EXPEDIENTE N",
    "EXPEDIENTE NRO",
    "EXPEDIENTE NUMERO",
    "DOCUMENTO SIMPLE",
    "N EXPEDIENTE",
    "N DE EXPEDIENTE",
    "N D.S",
    "N DS",
    "N EXP",
    "NRO. EXPEDIENTE",
    "NRO EXPEDIENTE",
    "NRO DE EXPEDIENTE",
    "NO EXPEDIENTE",
    "NO. EXPEDIENTE",
    "NUMERO EXPEDIENTE",
    "NUMERO DE EXPEDIENTE",
    "N° EXPEDIENTE",
    "N EXPEDIENTE",
    "N° DE EXPEDIENTE",
]


def first_existing_column(df, candidates):
    for col in candidates:
        if col in df.columns:
            return col
    return None


def parse_resolution_dates(series):
    cleaned = (
        series.astype(str)
        .str.strip()
        .str.replace("'", "", regex=False)
        .str.replace(r"^(\d{1,2})-(\d{1,2})-(\d{4})$", r"\1/\2/\3", regex=True)
        .str.replace(r"/+", "/", regex=True)
    )
    dates = pd.to_datetime(cleaned, dayfirst=True, errors="coerce")

    numeric = pd.to_numeric(series, errors="coerce")
    serial_mask = dates.isna() & numeric.between(20000, 60000)
    if serial_mask.any():
        dates.loc[serial_mask] = pd.to_datetime(
            numeric.loc[serial_mask],
            unit="D",
            origin="1899-12-30",
            errors="coerce",
        )
    return dates


def procedure_group(procedure):
    if procedure in PRIMARY_LICENSE_PROCEDURES:
        return "Licencias temporales/indeterminadas"
    if procedure == "TRANSFERENCIA DE LICENCIA DE FUNCIONAMIENTO":
        return "Transferencias"
    if procedure == "DUPLICADO DE LICENCIA DE FUNCIONAMIENTO":
        return "Duplicados"
    if procedure == "LICENCIA DE FUNCIONAMIENTO":
        return "Improcedentes con pago"
    return "Otros"


def normalize_license_procedure(value):
    text = normalize_text(value)
    if text == "LICENCIA DE FUNCIONAMIENTO TEMPORAL":
        return "LICENCIA TEMPORAL"
    if text == "LICENCIA DE FUNCIONAMIENTO INDETERMINADA":
        return "LICENCIA INDETERMINADA"
    return text


def get_sheet_procedure_column(df):
    return first_existing_column(df, ["TIPO DE PROCEDIMIENTO", "ASUNTO"])


def get_update_procedures_for_tab(tab_name):
    if str(tab_name).strip().upper() == "RESOLUCIONES 2023":
        return PRIMARY_LICENSE_PROCEDURES
    return TRACKED_PROCEDURES


def refresh_year_order(resumen_df):
    global YEAR_ORDER
    years = sorted(resumen_df["PERIODO"].dropna().astype(str).unique())
    if years:
        YEAR_ORDER = years
        for idx, year in enumerate(YEAR_ORDER):
            YEAR_COLORS.setdefault(year, px.colors.qualitative.Set2[idx % len(px.colors.qualitative.Set2)])


def classify_itse(value):
    text = normalize_text(value)
    if "MUY ALTO" in text:
        return "MUY ALTO", "MUY ALTO"
    if "ALTO" in text:
        return "ALTO", "ALTO"
    if "MEDIO" in text:
        return "MEDIO", "MEDIO"
    return None, None


def normalize_zone(value):
    text = normalize_text(value)
    if not text:
        return "SIN ZONA"
    numeric_match = re.fullmatch(r"([1-5])(?:\.0+)?", text)
    if numeric_match:
        return ZONE_LABELS[numeric_match.group(1)]
    zone_match = re.search(r"ZONA\s*([1-5])", text)
    if zone_match:
        return ZONE_LABELS[zone_match.group(1)]
    if "PAUL" in text and "POBLET" in text:
        return ZONE_LABELS["2"]
    if "CENTRO" in text and "POBLADO" in text:
        return ZONE_LABELS["3"]
    if "RURAL" in text and ("CENTRO" in text or "POBLADO" in text):
        return ZONE_LABELS["3"]
    if "JOSE" in text and ("GALVEZ" in text or "GALVES" in text):
        return ZONE_LABELS["4"]
    if "MANCHAY" in text:
        return ZONE_LABELS["5"]
    if "PACHACAMAC" in text:
        return ZONE_LABELS["1"]
    return text


def normalize_licencias_drive_sheet(df_raw, tab_name):
    fecha_col = first_existing_column(df_raw, DATE_COLUMNS)
    risk_col = first_existing_column(df_raw, ["TIPO DE ITSE", "TIPO ITSE"])
    procedure_col = get_sheet_procedure_column(df_raw)
    if fecha_col is None or procedure_col is None or risk_col is None:
        st.warning(f"La hoja {tab_name} no tiene las columnas requeridas para Licencias de Funcionamiento.")
        return None

    df = df_raw.copy()
    if procedure_col != "TIPO DE PROCEDIMIENTO":
        df["TIPO DE PROCEDIMIENTO"] = df[procedure_col]
    if risk_col != "TIPO DE ITSE":
        df["TIPO DE ITSE"] = df[risk_col]
    if "COSTO" not in df.columns:
        df["COSTO"] = 0
    df["PROCEDIMIENTO_NORMALIZADO"] = df["TIPO DE PROCEDIMIENTO"].map(normalize_license_procedure)
    df = df[df["PROCEDIMIENTO_NORMALIZADO"].isin(TRACKED_PROCEDURES)].copy()
    if df.empty:
        return None

    df["FECHA_RESOLUCION"] = parse_resolution_dates(df[fecha_col])
    df = df.dropna(subset=["FECHA_RESOLUCION"])
    if df.empty:
        return None

    df["COSTO_NUM"] = parse_money_series(df["COSTO"])
    risk_data = df["TIPO DE ITSE"].apply(classify_itse)
    df["RIESGO_DETALLE"] = risk_data.apply(lambda item: item[0])
    df["RIESGO_AGRUPADO"] = risk_data.apply(lambda item: item[1])
    tab_year = re.search(r"(20\d{2})", str(tab_name))
    df["PERIODO"] = tab_year.group(1) if tab_year else df["FECHA_RESOLUCION"].dt.year.astype(str)
    df["MES_NUM"] = df["FECHA_RESOLUCION"].dt.month
    df["MES"] = df["MES_NUM"].map(MONTH_MAP)
    df["TIPO_PROCEDIMIENTO"] = df["PROCEDIMIENTO_NORMALIZADO"].map(PROCEDURE_LABELS)
    df["GRUPO_REPORTE"] = df["PROCEDIMIENTO_NORMALIZADO"].map(procedure_group)
    df["ES_LICENCIA_PRINCIPAL"] = df["PROCEDIMIENTO_NORMALIZADO"].isin(PRIMARY_LICENSE_PROCEDURES)
    if "ZONA" in df.columns:
        df["ZONA_NORMALIZADA"] = df["ZONA"].map(normalize_zone)
    else:
        df["ZONA_NORMALIZADA"] = "SIN ZONA"
    df["HOJA_ORIGEN"] = tab_name
    return df


def load_licencias_drive_records():
    frames = []
    for tab_name in DRIVE_TABS:
        df_raw = get_resoluciones_sheet_or_none(tab_name=tab_name, show_warning=False)
        if df_raw is None:
            continue
        normalized = normalize_licencias_drive_sheet(df_raw, tab_name)
        if normalized is not None and not normalized.empty:
            normalized = normalized.loc[:, ~normalized.columns.duplicated()].copy()
            frames.append(normalized)

    if not frames:
        return None

    all_columns = []
    for frame in frames:
        for column in frame.columns:
            if column not in all_columns:
                all_columns.append(column)
    frames = [frame.reindex(columns=all_columns) for frame in frames]
    df = pd.concat(frames, ignore_index=True)
    df = df.sort_values(["FECHA_RESOLUCION", "TIPO_PROCEDIMIENTO"]).reset_index(drop=True)
    df.attrs["source"] = "drive"
    return df


def build_license_summary_from_records(records_df):
    df = records_df[records_df["ES_LICENCIA_PRINCIPAL"]].copy()
    df = df.dropna(subset=["RIESGO_AGRUPADO"])
    if df.empty:
        return None

    detalle_df = (
        df.groupby(["PERIODO", "MES_NUM", "MES", "RIESGO_DETALLE", "RIESGO_AGRUPADO"], observed=False)
        .agg(
            EXPEDIENTES=("FECHA_RESOLUCION", "size"),
            COSTO=("COSTO_NUM", "mean"),
            TOTAL=("COSTO_NUM", "sum"),
        )
        .reset_index()
        .sort_values(["PERIODO", "MES_NUM", "RIESGO_DETALLE"])
    )

    resumen_df = (
        df.groupby("PERIODO", observed=False)
        .agg(
            EXPEDIENTES=("FECHA_RESOLUCION", "size"),
            RECAUDACION=("COSTO_NUM", "sum"),
        )
        .reset_index()
        .sort_values("PERIODO")
    )

    detalle_df.attrs["source"] = "drive"
    resumen_df.attrs["source"] = "drive"
    return detalle_df, resumen_df


def empty_tramites_df():
    return pd.DataFrame(
        columns=[
            "PERIODO",
            "MES",
            "MES_NUM",
            "TIPO_PROCEDIMIENTO",
            "GRUPO_REPORTE",
            "RIESGO_AGRUPADO",
            "COSTO_NUM",
            "FECHA_RESOLUCION",
            "HOJA_ORIGEN",
            "PROCEDIMIENTO_NORMALIZADO",
            "ES_LICENCIA_PRINCIPAL",
        ]
    )


def load_licencias_funcionamiento_data():
    """Carga los datos fijos de Licencias de Funcionamiento."""
    drive_records = load_licencias_drive_records()
    drive_data = build_license_summary_from_records(drive_records) if drive_records is not None else None

    # Detalle transcrito del cuadro fuente
    detalle_data = [
        {"PERIODO": "2025", "RIESGO_DETALLE": "MEDIO", "RIESGO_AGRUPADO": "MEDIO", "EXPEDIENTES": 600, "COSTO": 200.90, "TOTAL": 120540.00},
        {"PERIODO": "2025", "RIESGO_DETALLE": "ALTOS Y MUY ALTOS", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 350, "COSTO": 678.90, "TOTAL": 237615.00},

        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "MEDIO DEL MES DE ENERO", "RIESGO_AGRUPADO": "MEDIO", "EXPEDIENTES": 67, "COSTO": 200.90, "TOTAL": 13460.30},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "MEDIO DEL MES DE FEBRERO", "RIESGO_AGRUPADO": "MEDIO", "EXPEDIENTES": 67, "COSTO": 193.20, "TOTAL": 12944.00},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "MEDIO DEL MES DE MARZO", "RIESGO_AGRUPADO": "MEDIO", "EXPEDIENTES": 61, "COSTO": 193.20, "TOTAL": 11785.20},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "MEDIO DEL MES DE ABRIL", "RIESGO_AGRUPADO": "MEDIO", "EXPEDIENTES": 29, "COSTO": 193.20, "TOTAL": 5602.80},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "ALTOS Y MUY ALTOS DEL MES DE ENERO", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 5, "COSTO": 678.90, "TOTAL": 3395.00},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "ALTOS Y MUY ALTOS DEL MES DE FEBRERO", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 5, "COSTO": 678.90, "TOTAL": 3395.00},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "ALTO DEL MES DE MARZO", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 3, "COSTO": 356.40, "TOTAL": 1069.20},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "ALTO DEL MES DE ABRIL", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 1, "COSTO": 356.40, "TOTAL": 356.40},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "MUY ALTO DEL MES DE MARZO", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 27, "COSTO": 631.20, "TOTAL": 17042.40},
        {"PERIODO": "2026 (Ene-Abr)", "RIESGO_DETALLE": "MUY ALTO DEL MES DE ABRIL", "RIESGO_AGRUPADO": "ALTOS Y MUY ALTOS", "EXPEDIENTES": 20, "COSTO": 631.20, "TOTAL": 12624.00},
    ]

    # Resumen anual según el total consolidado mostrado en tu cuadro
    resumen_data = [
        {"PERIODO": "2025", "EXPEDIENTES": 950, "RECAUDACION": 358155.00},
        {"PERIODO": "2026 (Ene-Abr)", "EXPEDIENTES": 285, "RECAUDACION": 81673.00},
    ]

    detalle_df = pd.DataFrame(detalle_data)
    resumen_df = pd.DataFrame(resumen_data)

    if drive_data is not None:
        drive_detalle_df, drive_resumen_df = drive_data
        drive_years = set(drive_resumen_df["PERIODO"].astype(str))

        detalle_years = detalle_df["PERIODO"].astype(str).str.extract(r"(\d{4})")[0].astype(int)
        resumen_years = resumen_df["PERIODO"].astype(str).str.extract(r"(\d{4})")[0].astype(int)

        detalle_df = pd.concat(
            [detalle_df[~detalle_years.astype(str).isin(drive_years)], drive_detalle_df],
            ignore_index=True,
        )
        resumen_df = pd.concat(
            [resumen_df[~resumen_years.astype(str).isin(drive_years)], drive_resumen_df],
            ignore_index=True,
        )
        detalle_df.attrs["source"] = "mixed"
        resumen_df.attrs["source"] = "mixed"
        drive_records.attrs["source"] = "drive"
        refresh_year_order(resumen_df)
        return detalle_df, resumen_df, drive_records

    if drive_records is not None:
        detalle_df["PERIODO"] = pd.Categorical(detalle_df["PERIODO"], categories=YEAR_ORDER, ordered=True)
        resumen_df["PERIODO"] = pd.Categorical(resumen_df["PERIODO"], categories=YEAR_ORDER, ordered=True)
        detalle_df.attrs["source"] = "local"
        resumen_df.attrs["source"] = "local"
        drive_records.attrs["source"] = "drive"
        return detalle_df, resumen_df, drive_records

    detalle_df["PERIODO"] = pd.Categorical(detalle_df["PERIODO"], categories=YEAR_ORDER, ordered=True)
    resumen_df["PERIODO"] = pd.Categorical(resumen_df["PERIODO"], categories=YEAR_ORDER, ordered=True)
    detalle_df.attrs["source"] = "local"
    resumen_df.attrs["source"] = "local"
    tramites_df = empty_tramites_df()
    tramites_df.attrs["source"] = "local"

    return detalle_df, resumen_df, tramites_df


def estadisticas_generales(resumen_df, tramites_df=None):
    st.subheader("Estadísticas generales")

    if tramites_df is not None and not tramites_df.empty:
        licencias = tramites_df[tramites_df["ES_LICENCIA_PRINCIPAL"]].copy()
        licencias = licencias[licencias["RIESGO_AGRUPADO"].notna()].copy()
        resumen_base = (
            licencias.groupby("PERIODO", observed=False)
            .agg(EXPEDIENTES=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
            .reset_index()
        )
    else:
        resumen_base = resumen_df.copy()

    total_expedientes = int(resumen_base["EXPEDIENTES"].sum())
    total_recaudado = float(resumen_base["RECAUDACION"].sum())
    periodo_max = resumen_base.loc[resumen_base["RECAUDACION"].idxmax(), "PERIODO"]
    promedio_expedientes = resumen_base["EXPEDIENTES"].mean()
    show_metric_row(
        [
            ("Licencias emitidas", f"{total_expedientes:,}"),
            ("Ingresos temp./indet.", f"S/ {total_recaudado:,.2f}"),
            ("Mayor ingreso", str(periodo_max)),
            ("Promedio de licencias", f"{promedio_expedientes:.1f}"),
        ]
    )
    return

    c1, c2, c3, c4 = st.columns(4)

    total_expedientes = int(resumen_df["EXPEDIENTES"].sum())
    total_recaudado = float(resumen_df["RECAUDACION"].sum())
    periodo_max = resumen_df.loc[resumen_df["RECAUDACION"].idxmax(), "PERIODO"]
    promedio_expedientes = resumen_df["EXPEDIENTES"].mean()

    c1.metric("Total de expedientes", f"{total_expedientes:,}")
    c2.metric("Recaudación total", f"S/ {total_recaudado:,.2f}")
    c3.metric("Mayor recaudación", str(periodo_max))
    c4.metric("Promedio de expedientes", f"{promedio_expedientes:.1f}")


def grafico_expedientes(resumen_df):
    st.subheader("Expedientes por año")

    fig = px.bar(
        resumen_df,
        x="PERIODO",
        y="EXPEDIENTES",
        color="PERIODO",
        text="EXPEDIENTES",
        color_discrete_map=YEAR_COLORS,
        category_orders={"PERIODO": YEAR_ORDER},
        height=420,
        labels={"PERIODO": "Año", "EXPEDIENTES": "Nro. de expedientes"}
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Nro. de expedientes",
        showlegend=False
    )

    fig.update_xaxes(type="category")
    fig.update_traces(
        textposition="outside",
        marker_line_color="rgba(0,0,0,0.3)",
        marker_line_width=2
    )
    polish_bar_traces(fig)
    apply_chart_style(fig, height=420)
    fig.update_layout(showlegend=False)

    st.plotly_chart(fig, use_container_width=True, key="licencias_general_expedientes")


def grafico_recaudacion(resumen_df):
    st.subheader("Recaudación por año")

    fig = px.bar(
        resumen_df,
        x="PERIODO",
        y="RECAUDACION",
        color="PERIODO",
        text="RECAUDACION",
        color_discrete_map=YEAR_COLORS,
        category_orders={"PERIODO": YEAR_ORDER},
        height=420,
        labels={"PERIODO": "Año", "RECAUDACION": "Recaudación (S/)"}
    )

    fig.update_traces(
        texttemplate="S/ %{y:,.2f}",
        textposition="outside",
        marker_line_color="rgba(0,0,0,0.3)",
        marker_line_width=2
    )
    polish_bar_traces(fig)

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Recaudación (S/)",
        showlegend=False
    )

    fig.update_xaxes(type="category")
    apply_chart_style(fig, height=420)
    fig.update_layout(showlegend=False)

    st.plotly_chart(fig, use_container_width=True, key="licencias_general_recaudacion")


def grafico_riesgo_apilado(detalle_df):
    st.subheader("Expedientes por Riesgo")

    riesgo_resumen = (
        detalle_df.groupby(["PERIODO", "RIESGO_AGRUPADO"], observed=False)["EXPEDIENTES"]
        .sum()
        .reset_index()
    )

    fig = px.bar(
        riesgo_resumen,
        x="PERIODO",
        y="EXPEDIENTES",
        color="RIESGO_AGRUPADO",
        barmode="stack",
        category_orders={"PERIODO": YEAR_ORDER},
        color_discrete_map=RISK_COLORS,
        height=450,
        labels={
            "PERIODO": "Año",
            "EXPEDIENTES": "Nro. de expedientes",
            "RIESGO_AGRUPADO": "Riesgo"
        }
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Nro. de expedientes",
        legend_title="Riesgo"
    )

    fig.update_xaxes(type="category")
    apply_chart_style(fig, legend_title="Riesgo", height=450)

    st.plotly_chart(fig, use_container_width=True, key="licencias_general_riesgo_apilado")


def grafico_recaudacion_riesgo(detalle_df):
    st.subheader("Recaudación por riesgo")

    riesgo_recaudacion = (
        detalle_df.groupby(["PERIODO", "RIESGO_AGRUPADO"], observed=False)["TOTAL"]
        .sum()
        .reset_index()
    )

    fig = px.bar(
        riesgo_recaudacion,
        x="PERIODO",
        y="TOTAL",
        color="RIESGO_AGRUPADO",
        barmode="group",
        category_orders={"PERIODO": YEAR_ORDER},
        color_discrete_map=RISK_COLORS,
        height=450,
        labels={
            "PERIODO": "Año",
            "TOTAL": "Recaudación (S/)",
            "RIESGO_AGRUPADO": "Riesgo"
        }
    )

    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Recaudación (S/)",
        legend_title="Riesgo"
    )

    fig.update_xaxes(type="category")
    apply_chart_style(fig, legend_title="Riesgo", height=450)

    st.plotly_chart(fig, use_container_width=True, key="licencias_general_recaudacion_riesgo")


def grafico_mensual_licencias(detalle_df):
    if not {"MES", "MES_NUM"}.issubset(detalle_df.columns):
        return

    st.subheader("Recaudación mensual por licencias")

    mensual = (
        detalle_df.groupby(["PERIODO", "MES", "MES_NUM"], observed=False)
        .agg(EXPEDIENTES=("EXPEDIENTES", "sum"), RECAUDACION=("TOTAL", "sum"))
        .reset_index()
        .sort_values(["PERIODO", "MES_NUM"])
    )

    fig = px.bar(
        mensual,
        x="MES",
        y="RECAUDACION",
        color="PERIODO",
        barmode="group",
        text="RECAUDACION",
        category_orders={"MES": MONTH_ORDER, "PERIODO": YEAR_ORDER},
        color_discrete_map=YEAR_COLORS,
        height=450,
        labels={
            "MES": "Mes",
            "RECAUDACION": "Recaudación (S/)",
            "PERIODO": "Año",
        },
    )
    fig.update_traces(texttemplate="S/ %{y:,.2f}", textposition="outside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes",
        yaxis_title="Recaudación (S/)",
        legend_title="Año",
    )
    st.plotly_chart(fig, use_container_width=True, key="licencias_general_mensual")

    tabla = mensual.pivot_table(
        index=["MES_NUM", "MES"],
        columns="PERIODO",
        values="RECAUDACION",
        aggfunc="sum",
        fill_value=0,
    ).reset_index()

    for year in YEAR_ORDER:
        if year not in tabla.columns:
            tabla[year] = 0

    tabla["Total"] = tabla[YEAR_ORDER].sum(axis=1)
    tabla = tabla[["MES", *YEAR_ORDER, "Total"]].rename(columns={"MES": "Mes"})

    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Mes": st.column_config.TextColumn("Mes"),
            **{
                year: st.column_config.NumberColumn(year, format="S/ %.2f")
                for year in YEAR_ORDER
            },
            "Total": st.column_config.NumberColumn("Total", format="S/ %.2f"),
        },
    )


def grafico_2026_por_mes_y_riesgo(detalle_df):
    if not {"MES", "MES_NUM", "RIESGO_AGRUPADO"}.issubset(detalle_df.columns):
        return

    detalle_2026 = detalle_df[detalle_df["PERIODO"].astype(str).str.startswith("2026")].copy()
    if detalle_2026.empty:
        return

    st.subheader("Licencias 2026 por mes y riesgo")

    mensual_riesgo = (
        detalle_2026.groupby(["MES", "MES_NUM", "RIESGO_AGRUPADO"], observed=False)
        .agg(EXPEDIENTES=("EXPEDIENTES", "sum"), RECAUDACION=("TOTAL", "sum"))
        .reset_index()
        .sort_values(["MES_NUM", "RIESGO_AGRUPADO"])
    )

    fig = px.bar(
        mensual_riesgo,
        x="MES",
        y="EXPEDIENTES",
        color="RIESGO_AGRUPADO",
        barmode="stack",
        text="EXPEDIENTES",
        category_orders={"MES": MONTH_ORDER, "RIESGO_AGRUPADO": ["MEDIO", "ALTO", "MUY ALTO"]},
        color_discrete_map=RISK_COLORS,
        height=450,
        labels={
            "MES": "Mes",
            "EXPEDIENTES": "Expedientes",
            "RIESGO_AGRUPADO": "Riesgo",
        },
    )
    fig.update_traces(textposition="inside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes",
        yaxis_title="Expedientes",
        legend_title="Riesgo",
    )
    st.plotly_chart(fig, use_container_width=True, key="licencias_2026_mensual_riesgo")

    recaudacion_mensual = (
        detalle_2026.groupby(["MES", "MES_NUM"], observed=False)["TOTAL"]
        .sum()
        .reset_index(name="RECAUDACION")
        .sort_values("MES_NUM")
    )

    fig_recaudacion = px.line(
        recaudacion_mensual,
        x="MES",
        y="RECAUDACION",
        markers=True,
        text="RECAUDACION",
        category_orders={"MES": MONTH_ORDER},
        color_discrete_sequence=["#0f4c81"],
        height=380,
        labels={"MES": "Mes", "RECAUDACION": "Recaudación (S/)"},
    )
    fig_recaudacion.update_traces(texttemplate="S/ %{y:,.2f}", textposition="top center", line=dict(width=3))
    fig_recaudacion.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Mes",
        yaxis_title="Recaudación (S/)",
        showlegend=False,
    )
    st.plotly_chart(fig_recaudacion, use_container_width=True, key="licencias_2026_recaudacion_mensual")


def estadisticas_procedimientos(tramites_df):
    if tramites_df is None or tramites_df.empty:
        return

    st.subheader("Ingresos por trámites de licencias")
    total_tramites = int(len(tramites_df))
    total_ingresos = float(tramites_df["COSTO_NUM"].sum())
    principales = tramites_df[tramites_df["ES_LICENCIA_PRINCIPAL"]]
    otros = tramites_df[~tramites_df["ES_LICENCIA_PRINCIPAL"]]

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Total trámites", f"{total_tramites:,}")
    c2.metric("Ingresos totales", f"S/ {total_ingresos:,.2f}")
    c3.metric("Temp. e indet.", f"{len(principales):,}")
    c4.metric("Otros trámites", f"{len(otros):,}")


def grafico_ingresos_por_tipo(tramites_df):
    if tramites_df is None or tramites_df.empty:
        return

    resumen = (
        tramites_df.groupby(["PERIODO", "TIPO_PROCEDIMIENTO"], observed=False)
        .agg(TRAMITES=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
        .reset_index()
        .sort_values(["PERIODO", "TIPO_PROCEDIMIENTO"])
    )

    fig = px.bar(
        resumen,
        x="PERIODO",
        y="RECAUDACION",
        color="TIPO_PROCEDIMIENTO",
        barmode="stack",
        text="RECAUDACION",
        color_discrete_map=PROCEDURE_COLORS,
        category_orders={"PERIODO": YEAR_ORDER},
        height=450,
        labels={
            "PERIODO": "Año",
            "RECAUDACION": "Recaudación (S/)",
            "TIPO_PROCEDIMIENTO": "Tipo de trámite",
        },
    )
    fig.update_traces(texttemplate="S/ %{y:,.2f}", textposition="inside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Recaudación (S/)",
        legend_title="Tipo de trámite",
    )
    st.plotly_chart(fig, use_container_width=True, key="licencias_tramites_ingresos_tipo")

    fig_tramites = px.bar(
        resumen,
        x="PERIODO",
        y="TRAMITES",
        color="TIPO_PROCEDIMIENTO",
        barmode="stack",
        text="TRAMITES",
        color_discrete_map=PROCEDURE_COLORS,
        category_orders={"PERIODO": YEAR_ORDER},
        height=420,
        labels={
            "PERIODO": "Año",
            "TRAMITES": "Trámites",
            "TIPO_PROCEDIMIENTO": "Tipo de trámite",
        },
    )
    fig_tramites.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Año",
        yaxis_title="Trámites",
        legend_title="Tipo de trámite",
    )
    st.plotly_chart(fig_tramites, use_container_width=True, key="licencias_tramites_cantidad_tipo")


def grafico_mensual_procedimientos(tramites_df):
    if tramites_df is None or tramites_df.empty:
        return

    mensual = (
        tramites_df.groupby(["PERIODO", "MES_NUM", "MES", "TIPO_PROCEDIMIENTO"], observed=False)
        .agg(TRAMITES=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
        .reset_index()
        .sort_values(["PERIODO", "MES_NUM", "TIPO_PROCEDIMIENTO"])
    )

    if mensual.empty:
        return

    fig = px.bar(
        mensual,
        x="MES",
        y="RECAUDACION",
        color="TIPO_PROCEDIMIENTO",
        facet_col="PERIODO",
        facet_col_wrap=2,
        barmode="stack",
        text="RECAUDACION",
        category_orders={"MES": MONTH_ORDER, "PERIODO": YEAR_ORDER},
        color_discrete_map=PROCEDURE_COLORS,
        height=520,
        labels={
            "MES": "Mes",
            "RECAUDACION": "Recaudación (S/)",
            "TIPO_PROCEDIMIENTO": "Tipo de trámite",
        },
    )
    fig.update_traces(texttemplate="S/ %{y:,.0f}", textposition="inside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        legend_title="Tipo de trámite",
    )
    fig.for_each_annotation(lambda annotation: annotation.update(text=annotation.text.replace("PERIODO=", "")))
    st.plotly_chart(fig, use_container_width=True, key="licencias_tramites_mensual")


def get_zoned_license_records(tramites_df, year=None):
    if tramites_df is None or tramites_df.empty or "ZONA_NORMALIZADA" not in tramites_df.columns:
        return pd.DataFrame()

    df = tramites_df[tramites_df["PROCEDIMIENTO_NORMALIZADO"].isin(PRIMARY_LICENSE_PROCEDURES)].copy()
    if year is not None:
        df = filter_period(df, year)
    complete_mask = df["RIESGO_AGRUPADO"].notna()
    if "TIPO DE ITSE" in df.columns:
        complete_mask &= df["TIPO DE ITSE"].map(lambda value: not is_blank(value))
    expediente_cols = [
        column
        for column in df.columns
        if any(token in normalize_text(column) for token in ["EXPEDIENTE", "EDIENTE"])
    ]
    if expediente_cols:
        has_expediente = df[expediente_cols].apply(
            lambda row: any(not is_blank(value) for value in row),
            axis=1,
        )
        complete_mask &= has_expediente
    df = df[complete_mask].copy()
    return df


def render_zone_license_report(tramites_df, year=None, key_suffix=None):
    zoned = get_zoned_license_records(tramites_df, year)
    if zoned.empty:
        return

    chart_key = key_suffix or year or "general"
    title_year = f" {year}" if year else ""
    title = f"Licencias por zona{title_year}" if year else "Licencias totales por zona"
    st.subheader(title)

    resumen = (
        zoned.groupby("ZONA_NORMALIZADA", observed=False)
        .agg(LICENCIAS=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
        .reset_index()
    )
    resumen["ZONA_ORDEN"] = resumen["ZONA_NORMALIZADA"].map(
        {zone: index for index, zone in enumerate(ZONE_ORDER)}
    ).fillna(len(ZONE_ORDER)).astype(int)
    resumen = resumen.sort_values(["ZONA_ORDEN", "ZONA_NORMALIZADA"])

    zone_category_order = resumen["ZONA_NORMALIZADA"].tolist()
    resumen_chart = resumen.sort_values(["ZONA_ORDEN", "ZONA_NORMALIZADA"], ascending=False).copy()
    fig = px.bar(
        resumen_chart,
        x="LICENCIAS",
        y="ZONA_NORMALIZADA",
        color="ZONA_NORMALIZADA",
        text="LICENCIAS",
        orientation="h",
        color_discrete_map=ZONE_COLORS,
        height=max(360, len(resumen_chart) * 82),
        labels={"ZONA_NORMALIZADA": "Zona", "LICENCIAS": "Licencias emitidas"},
    )
    polish_bar_traces(fig)
    fig.update_traces(texttemplate="%{x:,.0f}", textposition="outside", cliponaxis=False)
    apply_chart_style(fig, height=max(360, len(resumen_chart) * 82))
    fig.update_layout(showlegend=False, margin=dict(l=220, r=90, t=32, b=40))
    fig.update_xaxes(title="Licencias emitidas")
    fig.update_yaxes(title="Zona", categoryorder="array", categoryarray=zone_category_order)
    st.plotly_chart(fig, use_container_width=True, key=f"licencias_zona_{chart_key}")

    total_licencias = int(resumen["LICENCIAS"].sum())
    if total_licencias:
        resumen_pct = resumen_chart.copy()
        resumen_pct["PORCENTAJE"] = resumen_pct["LICENCIAS"] / total_licencias * 100
        fig_pct = px.bar(
            resumen_pct,
            x="PORCENTAJE",
            y="ZONA_NORMALIZADA",
            color="ZONA_NORMALIZADA",
            text="PORCENTAJE",
            orientation="h",
            color_discrete_map=ZONE_COLORS,
            height=max(360, len(resumen_pct) * 82),
            labels={"ZONA_NORMALIZADA": "Zona", "PORCENTAJE": "% del total"},
        )
        polish_bar_traces(fig_pct)
        fig_pct.update_traces(texttemplate="%{x:.1f}%", textposition="outside", cliponaxis=False)
        apply_chart_style(fig_pct, height=max(360, len(resumen_pct) * 82))
        fig_pct.update_layout(showlegend=False, margin=dict(l=220, r=90, t=32, b=40))
        fig_pct.update_xaxes(title="% del total de licencias emitidas", ticksuffix="%")
        fig_pct.update_yaxes(title="Zona", categoryorder="array", categoryarray=zone_category_order)
        st.plotly_chart(fig_pct, use_container_width=True, key=f"licencias_zona_pct_{chart_key}")

    tipo_zona = (
        zoned.groupby(["ZONA_NORMALIZADA", "TIPO_PROCEDIMIENTO"], observed=False)
        .agg(LICENCIAS=("FECHA_RESOLUCION", "size"))
        .reset_index()
    )
    tipo_zona["ZONA_ORDEN"] = tipo_zona["ZONA_NORMALIZADA"].map(
        {zone: index for index, zone in enumerate(ZONE_ORDER)}
    ).fillna(len(ZONE_ORDER)).astype(int)
    tipo_zona = tipo_zona.sort_values(["ZONA_ORDEN", "TIPO_PROCEDIMIENTO"])

    fig_tipo = px.bar(
        tipo_zona,
        x="ZONA_NORMALIZADA",
        y="LICENCIAS",
        color="TIPO_PROCEDIMIENTO",
        text="LICENCIAS",
        barmode="stack",
        color_discrete_map=PROCEDURE_COLORS,
        height=410,
        labels={
            "ZONA_NORMALIZADA": "Zona",
            "LICENCIAS": "Licencias emitidas",
            "TIPO_PROCEDIMIENTO": "Tipo de tramite",
        },
    )
    polish_bar_traces(fig_tipo, text_color="#ffffff")
    fig_tipo.update_traces(textposition="inside", insidetextanchor="middle")
    apply_chart_style(fig_tipo, legend_title="Tipo", height=410)
    fig_tipo.update_layout(xaxis_title="Zona", yaxis_title="Licencias emitidas")
    fig_tipo.update_xaxes(categoryorder="array", categoryarray=zone_category_order)
    st.plotly_chart(fig_tipo, use_container_width=True, key=f"licencias_zona_tipo_{chart_key}")

    tabla = resumen.drop(columns=["ZONA_ORDEN"], errors="ignore").rename(
        columns={
            "ZONA_NORMALIZADA": "Zona",
            "LICENCIAS": "Licencias",
            "RECAUDACION": "Recaudacion",
        }
    )
    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Zona": st.column_config.TextColumn("Zona"),
            "Licencias": st.column_config.NumberColumn("Licencias", format="%d"),
            "Recaudacion": st.column_config.NumberColumn("Recaudacion", format="S/ %.2f"),
        },
    )

    missing_zone = zoned[zoned["ZONA_NORMALIZADA"] == "SIN ZONA"].copy()
    if missing_zone.empty:
        return

    st.warning(
        f"Hay {len(missing_zone):,} licencias temporales/indeterminadas sin ZONA en la fuente."
    )
    expediente_col = find_expediente_column(missing_zone)
    columns = [
        column
        for column in [
            "RESOLUCION",
            expediente_col,
            "TIPO DE PROCEDIMIENTO",
            "SECTOR",
            "ZONA",
            "FECHA_RESOLUCION",
        ]
        if column and column in missing_zone.columns
    ]
    if columns:
        detail = missing_zone[columns].copy()
        if "FECHA_RESOLUCION" in detail.columns:
            detail["FECHA_RESOLUCION"] = detail["FECHA_RESOLUCION"].dt.strftime("%d/%m/%Y")
        detail = detail.rename(
            columns={
                "RESOLUCION": "Resolucion",
                expediente_col: "Expediente" if expediente_col else expediente_col,
                "TIPO DE PROCEDIMIENTO": "Tipo",
                "SECTOR": "Sector",
                "ZONA": "Zona",
                "FECHA_RESOLUCION": "Fecha resoluc.",
            }
        )
        st.dataframe(detail, use_container_width=True, hide_index=True)


def tabla_resumen_procedimientos(tramites_df):
    if tramites_df is None or tramites_df.empty:
        return

    st.subheader("Resumen por tipo de trámite")
    resumen = (
        tramites_df.groupby(["PERIODO", "TIPO_PROCEDIMIENTO", "GRUPO_REPORTE"], observed=False)
        .agg(
            Trámites=("FECHA_RESOLUCION", "size"),
            Recaudación=("COSTO_NUM", "sum"),
            Costo_promedio=("COSTO_NUM", "mean"),
        )
        .reset_index()
        .sort_values(["PERIODO", "TIPO_PROCEDIMIENTO"])
        .rename(
            columns={
                "PERIODO": "Año",
                "TIPO_PROCEDIMIENTO": "Tipo de trámite",
                "GRUPO_REPORTE": "Grupo",
            }
        )
    )

    st.dataframe(
        resumen,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año"),
            "Tipo de trámite": st.column_config.TextColumn("Tipo de trámite", width="large"),
            "Grupo": st.column_config.TextColumn("Grupo", width="medium"),
            "Trámites": st.column_config.NumberColumn("Trámites", format="%d"),
            "Recaudación": st.column_config.NumberColumn("Recaudación", format="S/ %.2f"),
            "Costo_promedio": st.column_config.NumberColumn("Costo promedio", format="S/ %.2f"),
        },
    )


def tabla_detalle_tramites(tramites_df):
    if tramites_df is None or tramites_df.empty:
        return

    st.subheader("Detalle de trámites importados desde Drive")
    detalle = tramites_df[
        [
            "PERIODO",
            "MES",
            "FECHA_RESOLUCION",
            "TIPO_PROCEDIMIENTO",
            "GRUPO_REPORTE",
            "RIESGO_AGRUPADO",
            "COSTO_NUM",
            "HOJA_ORIGEN",
        ]
    ].copy()
    detalle["FECHA_RESOLUCION"] = detalle["FECHA_RESOLUCION"].dt.strftime("%d/%m/%Y")
    detalle = detalle.rename(
        columns={
            "PERIODO": "Año",
            "MES": "Mes",
            "FECHA_RESOLUCION": "Fecha resoluc.",
            "TIPO_PROCEDIMIENTO": "Tipo de trámite",
            "GRUPO_REPORTE": "Grupo",
            "RIESGO_AGRUPADO": "Riesgo",
            "COSTO_NUM": "Costo",
            "HOJA_ORIGEN": "Hoja",
        }
    )

    st.dataframe(
        detalle,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año"),
            "Mes": st.column_config.TextColumn("Mes"),
            "Fecha resoluc.": st.column_config.TextColumn("Fecha resoluc."),
            "Tipo de trámite": st.column_config.TextColumn("Tipo de trámite", width="large"),
            "Grupo": st.column_config.TextColumn("Grupo", width="medium"),
            "Riesgo": st.column_config.TextColumn("Riesgo"),
            "Costo": st.column_config.NumberColumn("Costo", format="S/ %.2f"),
            "Hoja": st.column_config.TextColumn("Hoja"),
        },
    )


def filter_period(df, year):
    if df is None or df.empty or "PERIODO" not in df.columns:
        return df.iloc[0:0].copy() if df is not None else pd.DataFrame()
    return df[df["PERIODO"].astype(str).str.startswith(str(year))].copy()


def is_blank(value):
    if value is None or pd.isna(value):
        return True
    return str(value).strip() == ""


def normalize_expediente(value):
    text = normalize_text(value)
    text = re.sub(r"\s+", "", text)
    text = text.replace("EXP.", "").replace("EXP", "")
    text = re.sub(r"[^0-9A-Z/-]", "", text)
    match = re.search(r"(\d{1,6})[-/](20\d{2})", text)
    if match:
        return f"{int(match.group(1))}-{match.group(2)}"
    return text.lstrip("0")


def normalize_resolution_key(value):
    text = normalize_text(value)
    text = re.sub(r"[^0-9A-Z]+", "-", text).strip("-")
    parts = text.split("-")
    if len(parts) >= 2 and parts[0].isdigit() and re.fullmatch(r"20\d{2}", parts[1]):
        parts[0] = str(int(parts[0]))
    return "-".join(parts)


def normalize_match_name(value):
    return re.sub(r"[^0-9A-Z]+", "", normalize_text(value))


def normalize_reference_procedure(value):
    text = normalize_text(value)
    if "TEMPORAL" in text:
        return "LICENCIA TEMPORAL"
    if "INDETERMINADA" in text:
        return "LICENCIA INDETERMINADA"
    return normalize_license_procedure(text)


def get_configured_license_sheet_id():
    try:
        return extract_google_sheet_id(get_secret_value("GOOGLE_SHEET_ID"))
    except Exception:
        return ""


def get_service_account_email():
    try:
        return str(st.secrets["gcp_service_account"].get("client_email", "")).strip()
    except Exception:
        return ""


def render_license_sheet_id_input():
    default_sheet_id = st.session_state.get(LICENSE_SHEET_ID_STATE_KEY) or get_configured_license_sheet_id()
    sheet_value = st.text_input(
        "ID o URL del Google Sheet",
        value=default_sheet_id,
        help="Puedes usar el ID guardado en secrets o pegar aqui la URL completa del archivo.",
    )
    sheet_id = extract_google_sheet_id(sheet_value)
    st.session_state[LICENSE_SHEET_ID_STATE_KEY] = sheet_id
    return sheet_id


def find_column(df, candidates):
    return first_existing_column(df, [normalize_column_name(candidate) for candidate in candidates])


def find_expediente_column(df):
    candidate_columns = [
        column
        for column in [normalize_column_name(candidate) for candidate in EXPEDIENTE_COLUMNS]
        if column in df.columns
    ]
    for column in candidate_columns:
        values = df[column].fillna("").astype(str).str.strip()
        if values.ne("").any():
            return column

    for column in df.columns:
        normalized = normalize_text(column)
        if "EXPEDIENT" in normalized or "EDIENTE" in normalized:
            values = df[column].fillna("").astype(str).str.strip()
            if values.ne("").any():
                return column
    if candidate_columns:
        return candidate_columns[0]
    return None


def find_risk_column(df):
    return first_existing_column(df, ["TIPO DE ITSE", "TIPO ITSE", "RIESGO"])


def find_resolution_column(df):
    for column in df.columns:
        text = normalize_text(column)
        if "RESOLUCION" in text and "R.R" not in text:
            return column
    return None


def read_excel_with_detected_headers(path, sheet_name):
    variants = []
    for header_row in [0, 1]:
        try:
            df = pd.read_excel(path, sheet_name=sheet_name, header=header_row)
        except Exception:
            continue
        df.columns = [normalize_column_name(col) for col in df.columns]
        df = df.loc[:, [not str(col).startswith("UNNAMED") and bool(col) for col in df.columns]]
        df = df.loc[:, ~df.columns.duplicated()]
        variants.append(df)
    return variants


@st.cache_data(show_spinner=False)
def load_local_license_database(path_text):
    path = Path(path_text)
    if not path.exists():
        raise FileNotFoundError(f"No se encontro la BD local: {path}")

    frames = []
    excel = pd.ExcelFile(path)
    for sheet_name in excel.sheet_names:
        usable = None
        for df_candidate in read_excel_with_detected_headers(path, sheet_name):
            expediente_col = find_expediente_column(df_candidate)
            direccion_col = find_column(df_candidate, ["DIRECCION", "DIRECION", "DIRRECION"])
            if expediente_col is not None and direccion_col is not None:
                usable = (df_candidate, expediente_col, direccion_col)
                break
        if usable is None:
            continue
        df, expediente_col, direccion_col = usable
        risk_col = find_risk_column(df)
        columns = [expediente_col, direccion_col]
        if risk_col is not None:
            columns.append(risk_col)
        partial = df[columns].copy()
        partial.columns = ["EXPEDIENTE_BD", "DIRECCION_BD"] + (["RIESGO_BD"] if risk_col is not None else [])
        if "RIESGO_BD" not in partial.columns:
            partial["RIESGO_BD"] = ""
        partial["HOJA_BD"] = sheet_name
        partial["ARCHIVO_BD"] = path.name
        frames.append(partial)

    if not frames:
        raise ValueError("La BD local no tiene una hoja con columnas de expediente y direccion.")

    db = pd.concat(frames, ignore_index=True)
    db["EXPEDIENTE_KEY"] = db["EXPEDIENTE_BD"].map(normalize_expediente)
    db["DIRECCION_BD"] = db["DIRECCION_BD"].fillna("").astype(str).str.strip()
    db["RIESGO_BD"] = db["RIESGO_BD"].fillna("").astype(str).str.strip()
    db = db[(db["EXPEDIENTE_KEY"] != "") & ((db["DIRECCION_BD"] != "") | (db["RIESGO_BD"] != ""))]
    db = db.drop_duplicates(subset=["EXPEDIENTE_KEY"], keep="first")
    return db


def load_local_license_databases(paths):
    frames = []
    errors = []
    for path in paths:
        try:
            frames.append(load_local_license_database(str(path)))
        except Exception as exc:
            errors.append(f"{path}: {exc}")

    if not frames:
        raise ValueError("No se pudo leer ninguna BD local. " + " | ".join(errors))

    db = pd.concat(frames, ignore_index=True)
    db = db.drop_duplicates(subset=["EXPEDIENTE_KEY"], keep="first")
    return db


def require_sheet_columns(df, required_columns):
    missing = [column for column in required_columns if column not in df.columns]
    if missing:
        available = ", ".join(str(column) for column in df.columns if column != "__SHEET_ROW")
        raise ValueError(
            "Faltan columnas en la hoja: "
            + ", ".join(missing)
            + f". Columnas detectadas: {available}"
        )


def build_direccion_preview(sheet_id, tab_name=ADDRESS_UPDATE_TAB):
    _, sheet_df, _ = read_google_worksheet_with_rows(sheet_id, tab_name)
    require_sheet_columns(
        sheet_df,
        ["DIRECCION DEL ESTABLECIMIENTO"],
    )
    procedure_col = get_sheet_procedure_column(sheet_df)
    if procedure_col is None:
        raise ValueError("No se encontro columna TIPO DE PROCEDIMIENTO o ASUNTO en la hoja.")
    expediente_col = find_expediente_column(sheet_df)
    if expediente_col is None:
        available = ", ".join(str(column) for column in sheet_df.columns if column != "__SHEET_ROW")
        raise ValueError(
            "No se encontro una columna de expediente en la hoja. "
            f"Columnas detectadas: {available}"
        )

    work_df = sheet_df.copy()
    work_df["TIPO_NORMALIZADO"] = work_df[procedure_col].map(normalize_license_procedure)
    work_df["EXPEDIENTE_KEY"] = work_df[expediente_col].map(normalize_expediente)
    mask = (
        work_df["TIPO_NORMALIZADO"].isin(get_update_procedures_for_tab(tab_name))
        & work_df["DIRECCION DEL ESTABLECIMIENTO"].map(is_blank)
    )

    rows = []
    for _, row in work_df[mask].iterrows():
        rows.append(
            {
                "fila": int(row["__SHEET_ROW"]),
                "resolucion": row.get("RESOLUCION", ""),
                "n resolucion": row.get("N° RESOLUCION", ""),
                "resolucion sg": row.get("RESOLUCION DE SG", ""),
                "expediente": row.get(expediente_col, ""),
                "tipo": row.get(procedure_col, ""),
                "sector": row.get("SECTOR", ""),
                "direccion actual": row.get("DIRECCION DEL ESTABLECIMIENTO", ""),
                "sector vacio": "SI" if is_blank(row.get("SECTOR", "")) else "NO",
            }
        )

    return pd.DataFrame(rows)


def build_direccion_matches_preview(candidates_df, db_paths):
    if candidates_df is None or candidates_df.empty:
        return pd.DataFrame()

    db = load_local_license_databases(db_paths)
    db_by_expediente = db.set_index("EXPEDIENTE_KEY")

    rows = []
    for _, row in candidates_df.iterrows():
        expediente_key = normalize_expediente(row.get("expediente", ""))
        match = db_by_expediente.loc[expediente_key] if expediente_key in db_by_expediente.index else None
        direccion = "" if match is None else str(match["DIRECCION_BD"]).strip()
        rows.append(
            {
                "fila": int(row["fila"]),
                "resolucion": row.get("resolucion", ""),
                "n resolucion": row.get("n resolucion", ""),
                "resolucion sg": row.get("resolucion sg", ""),
                "expediente": row.get("expediente", ""),
                "tipo": row.get("tipo", ""),
                "sector": row.get("sector", ""),
                "direccion encontrada": direccion,
                "archivo bd": "" if match is None else str(match.get("ARCHIVO_BD", "")),
                "hoja bd": "" if match is None else str(match.get("HOJA_BD", "")),
                "estado": "CON COINCIDENCIA" if direccion else "SIN COINCIDENCIA",
            }
        )

    return pd.DataFrame(rows)


def normalize_risk_for_sheet(value):
    text = normalize_text(value)
    if not text:
        return ""
    if "MUY ALTO" in text:
        return "ITSE RIESGO MUY ALTO"
    if "ALTO" in text:
        return "ITSE RIESGO ALTO"
    if "MEDIO" in text:
        return "ITSE RIESGO MEDIO"
    if "BAJO" in text:
        return "ITSE RIESGO BAJO"
    return str(value).strip()


def build_risk_preview(sheet_id, tab_name=ADDRESS_UPDATE_TAB):
    _, sheet_df, _ = read_google_worksheet_with_rows(sheet_id, tab_name)
    procedure_col = get_sheet_procedure_column(sheet_df)
    if procedure_col is None:
        raise ValueError("No se encontro columna TIPO DE PROCEDIMIENTO o ASUNTO en la hoja.")
    risk_col = find_risk_column(sheet_df)
    if risk_col is None:
        raise ValueError("No se encontro columna TIPO DE ITSE / TIPO ITSE en la hoja.")
    expediente_col = find_expediente_column(sheet_df)
    if expediente_col is None:
        available = ", ".join(str(column) for column in sheet_df.columns if column != "__SHEET_ROW")
        raise ValueError(
            "No se encontro una columna de expediente en la hoja. "
            f"Columnas detectadas: {available}"
        )

    work_df = sheet_df.copy()
    work_df["TIPO_NORMALIZADO"] = work_df[procedure_col].map(normalize_license_procedure)
    mask = work_df["TIPO_NORMALIZADO"].isin(get_update_procedures_for_tab(tab_name)) & work_df[risk_col].map(is_blank)

    rows = []
    for _, row in work_df[mask].iterrows():
        rows.append(
            {
                "fila": int(row["__SHEET_ROW"]),
                "resolucion": row.get("RESOLUCION", ""),
                "n resolucion": row.get("N° RESOLUCION", ""),
                "resolucion sg": row.get("RESOLUCION DE SG", ""),
                "expediente": row.get(expediente_col, ""),
                "tipo": row.get(procedure_col, ""),
                "riesgo actual": row.get(risk_col, ""),
                "columna riesgo": risk_col,
            }
        )
    return pd.DataFrame(rows)


def build_risk_matches_preview(candidates_df, db_paths):
    if candidates_df is None or candidates_df.empty:
        return pd.DataFrame()

    db = load_local_license_databases(db_paths)
    db_by_expediente = db.set_index("EXPEDIENTE_KEY")

    rows = []
    for _, row in candidates_df.iterrows():
        expediente_key = normalize_expediente(row.get("expediente", ""))
        match = db_by_expediente.loc[expediente_key] if expediente_key in db_by_expediente.index else None
        riesgo = "" if match is None else normalize_risk_for_sheet(match.get("RIESGO_BD", ""))
        rows.append(
            {
                "fila": int(row["fila"]),
                "resolucion": row.get("resolucion", ""),
                "n resolucion": row.get("n resolucion", ""),
                "resolucion sg": row.get("resolucion sg", ""),
                "expediente": row.get("expediente", ""),
                "tipo": row.get("tipo", ""),
                "riesgo encontrado": riesgo,
                "archivo bd": "" if match is None else str(match.get("ARCHIVO_BD", "")),
                "hoja bd": "" if match is None else str(match.get("HOJA_BD", "")),
                "estado": "CON COINCIDENCIA" if riesgo else "SIN COINCIDENCIA",
                "columna riesgo": row.get("columna riesgo", "TIPO DE ITSE"),
            }
        )
    return pd.DataFrame(rows)


def apply_sheet_updates(sheet_id, preview_df, target_column, value_column, tab_name=ADDRESS_UPDATE_TAB):
    if preview_df is None or preview_df.empty:
        return 0

    try:
        from gspread.cell import Cell
    except ImportError as exc:
        raise RuntimeError("Falta la dependencia gspread para escribir en Google Sheets.") from exc

    worksheet = open_google_worksheet(sheet_id, tab_name)
    headers = [normalize_column_name(header) for header in worksheet.row_values(1)]
    target_header = normalize_column_name(target_column)
    if target_header not in headers:
        raise ValueError(f"No se encontro la columna {target_column} en la hoja.")

    target_col = headers.index(target_header) + 1
    cells = []
    for _, row in preview_df.iterrows():
        value = str(row[value_column]).strip()
        if not value:
            continue
        cells.append(Cell(row=int(row["fila"]), col=target_col, value=value))

    if not cells:
        return 0

    worksheet.update_cells(cells, value_input_option="USER_ENTERED")
    load_resoluciones_sheet.clear()
    return len(cells)


def build_zona_search_preview(sheet_id, search_text, zone_value, tab_name=ADDRESS_UPDATE_TAB):
    _, sheet_df, _ = read_google_worksheet_with_rows(sheet_id, tab_name)
    require_sheet_columns(sheet_df, ["ZONA"])

    tokens = re.findall(r"[0-9A-Z]+", normalize_text(search_text))
    if not tokens:
        return pd.DataFrame()

    expediente_col = find_expediente_column(sheet_df)
    resolution_col = find_resolution_column(sheet_df)
    procedure_col = get_sheet_procedure_column(sheet_df)
    searchable_columns = [
        column
        for column in [
            "DIRECCION DEL ESTABLECIMIENTO",
            "SECTOR",
            "APELLIDOS Y NOMBRES / RAZON SOCIAL",
            expediente_col,
            resolution_col,
            procedure_col,
            "GIRO / ACTIVIDAD COMERCIAL",
            "GIRO DEL NEGOCIO",
        ]
        if column and column in sheet_df.columns
    ]
    if not searchable_columns:
        raise ValueError("No se encontraron columnas para buscar coincidencias de zona.")

    def matches_search(row):
        haystack = " ".join(normalize_text(row.get(column, "")) for column in searchable_columns)
        return all(token in haystack for token in tokens)

    mask = sheet_df["ZONA"].map(is_blank) & sheet_df.apply(matches_search, axis=1)

    rows = []
    for _, row in sheet_df[mask].iterrows():
        rows.append(
            {
                "aplicar": True,
                "fila": int(row["__SHEET_ROW"]),
                "resolucion": row.get(resolution_col, "") if resolution_col else "",
                "expediente": row.get(expediente_col, "") if expediente_col else "",
                "razon social": row.get("APELLIDOS Y NOMBRES / RAZON SOCIAL", ""),
                "tipo": row.get(procedure_col, "") if procedure_col else "",
                "direccion": row.get("DIRECCION DEL ESTABLECIMIENTO", ""),
                "sector": row.get("SECTOR", ""),
                "zona actual": row.get("ZONA", ""),
                "zona propuesta": normalize_zone(zone_value),
            }
        )

    return pd.DataFrame(rows)


@st.cache_data(show_spinner=False)
def load_zone_reference_database(path_text=str(ZONE_REFERENCE_PATH)):
    path = Path(path_text)
    if not path.exists():
        raise FileNotFoundError(f"No se encontro el Excel de zonas: {path}")

    df = pd.read_excel(path, sheet_name=ZONE_REFERENCE_SHEET)
    df.columns = [normalize_column_name(column) for column in df.columns]
    df = df.loc[:, [bool(column) for column in df.columns]]
    df = df.loc[:, ~df.columns.duplicated()].copy()

    required = [
        "PERIODO",
        "EXPEDIENTE",
        "RESOLUCION",
        "APELLIDOS Y NOMBRES / RAZON SOCIAL",
        "ZONA",
    ]
    require_sheet_columns(df, required)

    procedure_col = first_existing_column(df, ["OTROS", "TIPO DE PROCEDIMIENTO", "ASUNTO", "VIGENCIA"])
    if procedure_col is None:
        raise ValueError("El Excel de zonas no tiene una columna de tipo de licencia para validar.")

    out = df.copy()
    out["PERIODO_KEY"] = out["PERIODO"].astype(str).str.extract(r"(20\d{2})")[0]
    out["EXPEDIENTE_KEY"] = out["EXPEDIENTE"].map(normalize_expediente)
    out["RESOLUCION_KEY"] = out["RESOLUCION"].map(normalize_resolution_key)
    out["NOMBRE_KEY"] = out["APELLIDOS Y NOMBRES / RAZON SOCIAL"].map(normalize_match_name)
    out["PROCEDIMIENTO_NORMALIZADO"] = out[procedure_col].map(normalize_reference_procedure)
    out["ZONA_NUMERO"] = out["ZONA"].astype(str).str.extract(r"([1-5])")[0]
    out["ZONA_PROPUESTA"] = out["ZONA_NUMERO"].map(ZONE_LABELS).fillna("")
    out = out[
        (out["PERIODO_KEY"].notna())
        & (out["EXPEDIENTE_KEY"] != "")
        & (out["RESOLUCION_KEY"] != "")
        & (out["ZONA_PROPUESTA"] != "")
        & (out["PROCEDIMIENTO_NORMALIZADO"].isin(PRIMARY_LICENSE_PROCEDURES))
    ].copy()
    out = out.drop_duplicates(
        subset=["PERIODO_KEY", "EXPEDIENTE_KEY", "RESOLUCION_KEY", "NOMBRE_KEY", "PROCEDIMIENTO_NORMALIZADO"],
        keep="first",
    )
    return out


def build_zone_reference_preview(sheet_id, tab_name, only_blank_zone=True):
    _, sheet_df, _ = read_google_worksheet_with_rows(sheet_id, tab_name)
    require_sheet_columns(sheet_df, ["ZONA", "APELLIDOS Y NOMBRES / RAZON SOCIAL"])

    procedure_col = get_sheet_procedure_column(sheet_df)
    expediente_col = find_expediente_column(sheet_df)
    resolution_col = find_resolution_column(sheet_df)
    name_col = "APELLIDOS Y NOMBRES / RAZON SOCIAL"
    if procedure_col is None or expediente_col is None or resolution_col is None:
        raise ValueError("La hoja no tiene columnas suficientes para cruzar zona por tipo, expediente y resolucion.")

    year_match = re.search(r"(20\d{2})", str(tab_name))
    tab_year = year_match.group(1) if year_match else ""
    reference_df = load_zone_reference_database()
    if tab_year:
        reference_df = reference_df[reference_df["PERIODO_KEY"] == tab_year].copy()
    if reference_df.empty:
        return pd.DataFrame()

    sheet_df = sheet_df.copy()
    sheet_df["PROCEDIMIENTO_NORMALIZADO"] = sheet_df[procedure_col].map(normalize_license_procedure)
    sheet_df = sheet_df[sheet_df["PROCEDIMIENTO_NORMALIZADO"].isin(PRIMARY_LICENSE_PROCEDURES)].copy()
    if only_blank_zone:
        sheet_df = sheet_df[sheet_df["ZONA"].map(is_blank)].copy()
    if sheet_df.empty:
        return pd.DataFrame()

    rows = []
    reference_by_exp = {
        key: group.copy()
        for key, group in reference_df.groupby("EXPEDIENTE_KEY", observed=False)
        if key
    }
    for _, row in sheet_df.iterrows():
        expediente_key = normalize_expediente(row.get(expediente_col, ""))
        resolution_key = normalize_resolution_key(row.get(resolution_col, ""))
        name_key = normalize_match_name(row.get(name_col, ""))
        procedure = row.get("PROCEDIMIENTO_NORMALIZADO", "")
        candidates = reference_by_exp.get(expediente_key, pd.DataFrame())
        if candidates.empty:
            continue

        scored = candidates.copy()
        scored["MATCH_RESOLUCION"] = scored["RESOLUCION_KEY"] == resolution_key
        scored["MATCH_RAZON_SOCIAL"] = scored["NOMBRE_KEY"] == name_key
        scored["MATCH_TIPO"] = scored["PROCEDIMIENTO_NORMALIZADO"] == procedure
        scored["SCORE"] = (
            scored["MATCH_RESOLUCION"].astype(int) * 4
            + scored["MATCH_RAZON_SOCIAL"].astype(int) * 2
            + scored["MATCH_TIPO"].astype(int) * 2
        )
        scored = scored.sort_values(["SCORE", "MATCH_RESOLUCION", "MATCH_RAZON_SOCIAL", "MATCH_TIPO"], ascending=False)
        best = scored.iloc[0]
        if not bool(best["MATCH_RESOLUCION"]) and not (bool(best["MATCH_RAZON_SOCIAL"]) and bool(best["MATCH_TIPO"])):
            continue

        if bool(best["MATCH_RESOLUCION"]) and bool(best["MATCH_RAZON_SOCIAL"]) and bool(best["MATCH_TIPO"]):
            estado = "COINCIDENCIA COMPLETA"
            aplicar = True
        elif bool(best["MATCH_RESOLUCION"]) and bool(best["MATCH_TIPO"]):
            estado = "COINCIDENCIA FUERTE"
            aplicar = True
        else:
            estado = "REVISAR COINCIDENCIA"
            aplicar = False

        rows.append(
            {
                "aplicar": aplicar,
                "fila": int(row["__SHEET_ROW"]),
                "estado": estado,
                "zona actual": row.get("ZONA", ""),
                "zona propuesta": best.get("ZONA_PROPUESTA", ""),
                "expediente drive": row.get(expediente_col, ""),
                "expediente excel": best.get("EXPEDIENTE", ""),
                "resolucion drive": row.get(resolution_col, ""),
                "resolucion excel": best.get("RESOLUCION", ""),
                "razon social drive": row.get(name_col, ""),
                "razon social excel": best.get("APELLIDOS Y NOMBRES / RAZON SOCIAL", ""),
                "tipo drive": row.get(procedure_col, ""),
                "tipo excel": best.get("OTROS", best.get("PROCEDIMIENTO_NORMALIZADO", "")),
                "valida resolucion": "SI" if bool(best["MATCH_RESOLUCION"]) else "NO",
                "valida razon social": "SI" if bool(best["MATCH_RAZON_SOCIAL"]) else "NO",
                "valida tipo": "SI" if bool(best["MATCH_TIPO"]) else "NO",
            }
        )

    return pd.DataFrame(rows)


def format_google_write_error(exc):
    message = str(exc)
    if "403" in message or "does not have permission" in message:
        service_account_email = get_service_account_email()
        email_text = f" ({service_account_email})" if service_account_email else ""
        return (
            "Google Sheets rechazo la escritura por permisos. "
            f"Comparte el archivo Drive con la cuenta de servicio{email_text} como Editor y vuelve a intentar."
        )
    return f"No se pudo escribir en Google Sheets: {exc}"


def render_direccion_update_tool(sheet_id, tab_name, db_paths):
    st.subheader("Completar DIRECCION DEL ESTABLECIMIENTO")
    st.caption(
        f"Hoja objetivo: {tab_name}. Paso 1: revisa filas con direccion vacia. "
        "Paso 2: busca coincidencias en las BD locales configuradas."
    )

    key_suffix = normalize_text(tab_name).replace(" ", "_")

    if st.button("1. Ver expedientes con direccion vacia", key=f"preview_direcciones_{key_suffix}", use_container_width=True):
        try:
            st.session_state["direccion_candidates"] = build_direccion_preview(sheet_id, tab_name)
            st.session_state["direccion_sheet_id"] = sheet_id
            st.session_state["direccion_tab_name"] = tab_name
            st.session_state.pop("direccion_matches", None)
        except Exception as exc:
            st.session_state.pop("direccion_candidates", None)
            st.session_state.pop("direccion_matches", None)
            st.error(f"No se pudo preparar la vista previa: {exc}")

    candidates = st.session_state.get("direccion_candidates")
    matches = st.session_state.get("direccion_matches")
    if candidates is not None and (
        st.session_state.get("direccion_sheet_id") != sheet_id
        or st.session_state.get("direccion_tab_name") != tab_name
    ):
        st.session_state.pop("direccion_candidates", None)
        st.session_state.pop("direccion_matches", None)
        candidates = None
        matches = None
    if candidates is None:
        return

    if candidates.empty:
        st.info("No hay filas pendientes para completar direccion.")
        return

    sector_empty = int((candidates["sector vacio"] == "SI").sum()) if "sector vacio" in candidates.columns else 0
    show_metric_row(
        [
            ("Direcciones vacias", f"{len(candidates):,}"),
            ("Sector vacio", f"{sector_empty:,}"),
            ("Con sector", f"{len(candidates) - sector_empty:,}"),
        ]
    )
    st.dataframe(candidates, use_container_width=True, hide_index=True)

    if st.button("2. Buscar coincidencias en BD local", key=f"match_direcciones_{key_suffix}", use_container_width=True):
        try:
            matches = build_direccion_matches_preview(candidates, db_paths)
            st.session_state["direccion_matches"] = matches
        except Exception as exc:
            st.session_state.pop("direccion_matches", None)
            st.error(f"No se pudo buscar en la BD local: {exc}")
            return

    if matches is None:
        return

    if matches.empty:
        st.warning("No se encontraron candidatos para cruzar con la BD local.")
        return

    matches_to_write = matches[matches["estado"] == "CON COINCIDENCIA"].copy()
    st.markdown("#### Coincidencias encontradas en BD")
    show_metric_row(
        [
            ("Revisados en BD", f"{len(matches):,}"),
            ("Con coincidencia", f"{len(matches_to_write):,}"),
            ("Sin coincidencia", f"{len(matches) - len(matches_to_write):,}"),
        ]
    )
    st.dataframe(matches, use_container_width=True, hide_index=True)

    if matches_to_write.empty:
        st.warning("No hay direcciones encontradas para escribir.")
        return

    if st.button("Escribir direcciones encontradas", key=f"apply_direcciones_{key_suffix}", use_container_width=True):
        try:
            updated = apply_sheet_updates(
                sheet_id,
                matches_to_write,
                "DIRECCION DEL ESTABLECIMIENTO",
                "direccion encontrada",
                tab_name=tab_name,
            )
            st.success(f"Se actualizaron {updated:,} filas en Google Sheets.")
            st.session_state.pop("direccion_candidates", None)
            st.session_state.pop("direccion_matches", None)
        except Exception as exc:
            st.error(format_google_write_error(exc))


def render_risk_update_tool(sheet_id, tab_name, db_paths):
    st.subheader("Completar TIPO ITSE")
    st.caption(
        f"Hoja objetivo: {tab_name}. Busca expedientes con TIPO ITSE vacio y cruza con la columna RIESGO de la BD local."
    )
    key_suffix = normalize_text(tab_name).replace(" ", "_")

    if st.button("1. Ver expedientes sin tipo ITSE", key=f"preview_riesgo_{key_suffix}", use_container_width=True):
        try:
            st.session_state["risk_candidates"] = build_risk_preview(sheet_id, tab_name)
            st.session_state["risk_sheet_id"] = sheet_id
            st.session_state["risk_tab_name"] = tab_name
            st.session_state.pop("risk_matches", None)
        except Exception as exc:
            st.session_state.pop("risk_candidates", None)
            st.session_state.pop("risk_matches", None)
            st.error(f"No se pudo preparar la vista previa: {exc}")

    candidates = st.session_state.get("risk_candidates")
    matches = st.session_state.get("risk_matches")
    if candidates is not None and (
        st.session_state.get("risk_sheet_id") != sheet_id
        or st.session_state.get("risk_tab_name") != tab_name
    ):
        st.session_state.pop("risk_candidates", None)
        st.session_state.pop("risk_matches", None)
        candidates = None
        matches = None
    if candidates is None:
        return

    if candidates.empty:
        st.info("No hay filas pendientes para completar TIPO ITSE.")
        return

    show_metric_row([("Tipo ITSE vacio", f"{len(candidates):,}")])
    st.dataframe(candidates.drop(columns=["columna riesgo"], errors="ignore"), use_container_width=True, hide_index=True)

    if st.button("2. Buscar riesgos en BD local", key=f"match_riesgo_{key_suffix}", use_container_width=True):
        try:
            matches = build_risk_matches_preview(candidates, db_paths)
            st.session_state["risk_matches"] = matches
        except Exception as exc:
            st.session_state.pop("risk_matches", None)
            st.error(f"No se pudo buscar en la BD local: {exc}")
            return

    if matches is None:
        return

    matches_to_write = matches[matches["estado"] == "CON COINCIDENCIA"].copy()
    show_metric_row(
        [
            ("Revisados en BD", f"{len(matches):,}"),
            ("Con coincidencia", f"{len(matches_to_write):,}"),
            ("Sin coincidencia", f"{len(matches) - len(matches_to_write):,}"),
        ]
    )
    st.dataframe(matches.drop(columns=["columna riesgo"], errors="ignore"), use_container_width=True, hide_index=True)

    if matches_to_write.empty:
        st.warning("No hay riesgos encontrados para escribir.")
        return

    target_column = str(matches_to_write["columna riesgo"].dropna().iloc[0])
    if st.button("Escribir tipos ITSE encontrados", key=f"apply_riesgo_{key_suffix}", use_container_width=True):
        try:
            updated = apply_sheet_updates(
                sheet_id,
                matches_to_write,
                target_column,
                "riesgo encontrado",
                tab_name=tab_name,
            )
            st.success(f"Se actualizaron {updated:,} filas en Google Sheets.")
            st.session_state.pop("risk_candidates", None)
            st.session_state.pop("risk_matches", None)
        except Exception as exc:
            st.error(format_google_write_error(exc))


def render_zona_search_update_tool(sheet_id, tab_name):
    st.subheader("Completar ZONA por busqueda de palabras clave")
    st.caption(
        f"Hoja objetivo: {tab_name}. Busca texto en direccion, sector, razon social, expediente, resolucion y tipo; "
        "muestra solo filas con ZONA vacia y permite validar una por una antes de escribir."
    )

    key_suffix = normalize_text(tab_name).replace(" ", "_")
    zone_options = list(ZONE_LABELS.values())
    search_text = st.text_input("Buscar palabra clave", value="MANCHAY", key=f"zona_search_text_{key_suffix}")
    zone_value = st.selectbox("Zona a escribir", options=zone_options, index=4, key=f"zona_value_{key_suffix}")

    if st.button("Buscar filas para completar zona", key=f"preview_zona_{key_suffix}", use_container_width=True):
        try:
            st.session_state["zona_preview"] = build_zona_search_preview(
                sheet_id,
                search_text,
                zone_value,
                tab_name=tab_name,
            )
            st.session_state["zona_sheet_id"] = sheet_id
            st.session_state["zona_tab_name"] = tab_name
            st.session_state["zona_search"] = search_text
            st.session_state["zona_value"] = zone_value
        except Exception as exc:
            st.session_state.pop("zona_preview", None)
            st.error(f"No se pudo preparar la vista previa: {exc}")

    preview = st.session_state.get("zona_preview")
    if preview is not None and (
        st.session_state.get("zona_sheet_id") != sheet_id
        or st.session_state.get("zona_tab_name") != tab_name
    ):
        st.session_state.pop("zona_preview", None)
        preview = None
    if preview is None:
        return

    if preview.empty:
        st.info("No hay filas pendientes para la busqueda indicada.")
        return

    edited_preview = st.data_editor(
        preview,
        use_container_width=True,
        hide_index=True,
        column_config={
            "aplicar": st.column_config.CheckboxColumn("Aplicar"),
            "fila": st.column_config.NumberColumn("Fila", format="%d"),
            "resolucion": st.column_config.TextColumn("Resolucion"),
            "expediente": st.column_config.TextColumn("Expediente"),
            "razon social": st.column_config.TextColumn("Razon social"),
            "tipo": st.column_config.TextColumn("Tipo"),
            "direccion": st.column_config.TextColumn("Direccion"),
            "sector": st.column_config.TextColumn("Sector"),
            "zona actual": st.column_config.TextColumn("Zona actual"),
            "zona propuesta": st.column_config.TextColumn("Zona propuesta"),
        },
        disabled=["fila", "resolucion", "expediente", "razon social", "tipo", "direccion", "sector", "zona actual"],
        key=f"zona_editor_{key_suffix}",
    )
    selected_preview = edited_preview[
        edited_preview["aplicar"].fillna(False)
        & edited_preview["zona propuesta"].astype(str).str.strip().ne("")
    ].copy()
    show_metric_row(
        [
            ("Encontradas", f"{len(preview):,}"),
            ("Seleccionadas", f"{len(selected_preview):,}"),
        ]
    )

    if st.button("Escribir zonas seleccionadas", key=f"apply_zona_{key_suffix}", use_container_width=True):
        try:
            updated = apply_sheet_updates(sheet_id, selected_preview, "ZONA", "zona propuesta", tab_name=tab_name)
            st.success(f"Se actualizaron {updated:,} filas en Google Sheets.")
            st.session_state.pop("zona_preview", None)
        except Exception as exc:
            st.error(format_google_write_error(exc))


def render_zone_reference_update_tool(sheet_id, tab_name):
    st.subheader("Completar ZONA desde DATA 2023-2026")
    st.caption(
        f"Hoja objetivo: {tab_name}. Cruza contra {ZONE_REFERENCE_PATH.name} / {ZONE_REFERENCE_SHEET} "
        "usando expediente, resolucion, razon social y tipo de licencia."
    )

    key_suffix = normalize_text(tab_name).replace(" ", "_")
    only_blank_zone = st.checkbox(
        "Solo filas con ZONA vacia",
        value=True,
        key=f"zone_reference_only_blank_{key_suffix}",
    )

    if st.button("Buscar zonas en DATA 2023-2026", key=f"preview_zone_reference_{key_suffix}", use_container_width=True):
        try:
            st.session_state["zone_reference_preview"] = build_zone_reference_preview(
                sheet_id,
                tab_name,
                only_blank_zone=only_blank_zone,
            )
            st.session_state["zone_reference_sheet_id"] = sheet_id
            st.session_state["zone_reference_tab_name"] = tab_name
            st.session_state["zone_reference_only_blank"] = only_blank_zone
        except Exception as exc:
            st.session_state.pop("zone_reference_preview", None)
            st.error(f"No se pudo preparar la vista previa: {exc}")

    preview = st.session_state.get("zone_reference_preview")
    if preview is not None and (
        st.session_state.get("zone_reference_sheet_id") != sheet_id
        or st.session_state.get("zone_reference_tab_name") != tab_name
        or st.session_state.get("zone_reference_only_blank") != only_blank_zone
    ):
        st.session_state.pop("zone_reference_preview", None)
        preview = None
    if preview is None:
        return

    if preview.empty:
        st.info("No se encontraron coincidencias para completar ZONA con la DATA 2023-2026.")
        return

    complete_count = int((preview["estado"] == "COINCIDENCIA COMPLETA").sum())
    strong_count = int((preview["estado"] == "COINCIDENCIA FUERTE").sum())
    review_count = int((preview["estado"] == "REVISAR COINCIDENCIA").sum())
    show_metric_row(
        [
            ("Coincidencia completa", f"{complete_count:,}"),
            ("Coincidencia fuerte", f"{strong_count:,}"),
            ("Por revisar", f"{review_count:,}"),
        ]
    )

    edited_preview = st.data_editor(
        preview,
        use_container_width=True,
        hide_index=True,
        column_config={
            "aplicar": st.column_config.CheckboxColumn("Aplicar"),
            "fila": st.column_config.NumberColumn("Fila", format="%d"),
            "estado": st.column_config.TextColumn("Estado"),
            "zona actual": st.column_config.TextColumn("Zona actual"),
            "zona propuesta": st.column_config.TextColumn("Zona propuesta"),
            "expediente drive": st.column_config.TextColumn("Expediente Drive"),
            "expediente excel": st.column_config.TextColumn("Expediente Excel"),
            "resolucion drive": st.column_config.TextColumn("Resolucion Drive"),
            "resolucion excel": st.column_config.TextColumn("Resolucion Excel"),
            "razon social drive": st.column_config.TextColumn("Razon social Drive"),
            "razon social excel": st.column_config.TextColumn("Razon social Excel"),
            "tipo drive": st.column_config.TextColumn("Tipo Drive"),
            "tipo excel": st.column_config.TextColumn("Tipo Excel"),
        },
        disabled=[
            "fila",
            "estado",
            "zona actual",
            "expediente drive",
            "expediente excel",
            "resolucion drive",
            "resolucion excel",
            "razon social drive",
            "razon social excel",
            "tipo drive",
            "tipo excel",
            "valida resolucion",
            "valida razon social",
            "valida tipo",
        ],
        key=f"zone_reference_editor_{key_suffix}",
    )
    selected_preview = edited_preview[
        edited_preview["aplicar"].fillna(False)
        & edited_preview["zona propuesta"].astype(str).str.strip().ne("")
    ].copy()
    show_metric_row(
        [
            ("Encontradas", f"{len(preview):,}"),
            ("Seleccionadas", f"{len(selected_preview):,}"),
        ]
    )

    if st.button("Escribir zonas seleccionadas desde DATA", key=f"apply_zone_reference_{key_suffix}", use_container_width=True):
        try:
            updated = apply_sheet_updates(sheet_id, selected_preview, "ZONA", "zona propuesta", tab_name=tab_name)
            st.success(f"Se actualizaron {updated:,} filas en Google Sheets.")
            st.session_state.pop("zone_reference_preview", None)
        except Exception as exc:
            st.error(format_google_write_error(exc))


def render_resoluciones_2026_update_tools():
    st.subheader("Actualizaciones de resoluciones")
    st.info(
        "Las vistas previas no escriben cambios. La escritura ocurre solo con el boton de confirmacion de cada bloque."
    )
    service_account_email = get_service_account_email()
    if service_account_email:
        st.caption(f"Cuenta de servicio usada por la app: {service_account_email}")
    sheet_id = render_license_sheet_id_input()
    if not sheet_id:
        st.warning("Ingresa el ID o URL del Google Sheet para continuar.")
        return

    tab_name = st.selectbox(
        "Hoja a actualizar",
        options=list(ADDRESS_UPDATE_CONFIGS.keys()),
        index=0,
        key="address_update_tab_name",
    )
    db_paths = ADDRESS_UPDATE_CONFIGS[tab_name]
    st.caption("BD local usada: " + ", ".join(path.name for path in db_paths))

    render_direccion_update_tool(sheet_id, tab_name, db_paths)
    st.markdown("---")
    render_risk_update_tool(sheet_id, tab_name, db_paths)
    st.markdown("---")
    render_zona_search_update_tool(sheet_id, tab_name)


def render_drive_refresh_button():
    if st.button("Actualizar datos desde Drive", key="refresh_licencias_drive", use_container_width=True):
        load_resoluciones_sheet.clear()
        st.success("Cache limpiada. Recargando datos desde Google Drive.")
        st.rerun()


def render_general_licencias(detalle_df, resumen_df, tramites_df):
    estadisticas_generales(resumen_df, tramites_df)
    st.markdown("---")

    grafico_expedientes(resumen_df)
    st.markdown("---")

    grafico_riesgo_apilado(detalle_df)
    st.markdown("---")

    render_zone_license_report(tramites_df, key_suffix="general_total")
    st.markdown("---")

    grafico_recaudacion(resumen_df)
    st.markdown("---")

    grafico_recaudacion_riesgo(detalle_df)
    st.markdown("---")

    grafico_mensual_licencias(detalle_df)
    st.markdown("---")

    grafico_2026_por_mes_y_riesgo(detalle_df)
    st.markdown("---")

    estadisticas_procedimientos(tramites_df)
    grafico_ingresos_por_tipo(tramites_df)
    grafico_mensual_procedimientos(tramites_df)
    st.markdown("---")
    tabla_resumen_procedimientos(tramites_df)
    tabla_detalle_tramites(tramites_df)
    if tramites_df is not None and not tramites_df.empty:
        st.markdown("---")

    tabla_resumen_anual(resumen_df)
    st.markdown("---")

    tabla_detallada(detalle_df)
    st.markdown("---")

    observaciones(resumen_df, detalle_df)


def show_metric_row(metrics):
    cards = []
    for label, value in metrics:
        cards.append(
            '<div class="gde-metric-card">'
            f'<div class="gde-metric-label">{escape(str(label))}</div>'
            f'<div class="gde-metric-value">{escape(str(value))}</div>'
            "</div>"
        )
    st.markdown(f'<div class="gde-metric-grid">{"".join(cards)}</div>', unsafe_allow_html=True)


def variacion_porcentual(base, actual):
    if base in (0, None) or pd.isna(base):
        return None
    return ((actual - base) / base) * 100


def build_monthly_license_summary(detalle_year):
    if detalle_year.empty or not {"MES", "MES_NUM"}.issubset(detalle_year.columns):
        return pd.DataFrame()
    return (
        detalle_year.groupby(["MES_NUM", "MES"], observed=False)
        .agg(EXPEDIENTES=("EXPEDIENTES", "sum"), RECAUDACION=("TOTAL", "sum"))
        .reset_index()
        .sort_values("MES_NUM")
    )


def render_year_license_monthly_charts(year, detalle_year, tramites_year):
    mensual = build_monthly_license_summary(detalle_year)
    if mensual.empty:
        return

    st.subheader("Licencias emitidas por mes")
    mensual["COLOR"] = color_scale_for_values(mensual["EXPEDIENTES"], "Tealgrn")
    fig_expedientes = px.bar(
        mensual,
        x="MES",
        y="EXPEDIENTES",
        text="EXPEDIENTES",
        category_orders={"MES": MONTH_ORDER},
        color="MES",
        color_discrete_sequence=mensual["COLOR"].tolist(),
        height=430,
        labels={"MES": "Mes", "EXPEDIENTES": "Licencias emitidas"},
    )
    fig_expedientes.update_traces(textposition="outside", texttemplate="%{y:,.0f}")
    polish_bar_traces(fig_expedientes)
    apply_chart_style(fig_expedientes, height=430)
    fig_expedientes.update_layout(
        xaxis_title="Mes",
        yaxis_title="Licencias emitidas",
        showlegend=False,
        bargap=0.18,
        margin=dict(l=28, r=48, t=46, b=56),
    )
    fig_expedientes.update_yaxes(gridcolor="#dfe9f3")
    st.plotly_chart(fig_expedientes, use_container_width=True, key=f"licencias_anual_{year}_emitidas_mensual")
    render_zone_license_report(tramites_year, year=year, key_suffix=f"tab_{year}")

    st.subheader("Recaudación mensual por licencias")
    recaudacion_movil = mensual.sort_values("MES_NUM", ascending=False).copy()
    fig_recaudacion = px.bar(
        recaudacion_movil,
        x="RECAUDACION",
        y="MES",
        orientation="h",
        text="RECAUDACION",
        color_discrete_sequence=["#0f4c81"],
        height=max(420, len(recaudacion_movil) * 38),
        labels={"MES": "Mes", "RECAUDACION": "Recaudación (S/)"},
    )
    fig_recaudacion.update_traces(texttemplate="S/ %{x:,.2f}", textposition="outside", cliponaxis=False)
    fig_recaudacion.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Recaudación (S/)",
        yaxis_title="Mes",
        showlegend=False,
        margin=dict(l=10, r=110, t=20, b=40),
    )
    st.plotly_chart(fig_recaudacion, use_container_width=True, key=f"licencias_anual_{year}_recaudacion_mensual")

    tabla = mensual.rename(
        columns={
            "MES": "Mes",
            "EXPEDIENTES": "Licencias emitidas",
            "RECAUDACION": "Recaudación",
        }
    )[["Mes", "Licencias emitidas", "Recaudación"]]
    st.dataframe(
        tabla,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Mes": st.column_config.TextColumn("Mes"),
            "Licencias emitidas": st.column_config.NumberColumn("Licencias emitidas", format="%d"),
            "Recaudación": st.column_config.NumberColumn("Recaudación", format="S/ %.2f"),
        },
    )


def render_year_observations(year, detalle_year, tramites_year):
    st.subheader("Lectura del año")
    if detalle_year.empty:
        st.info("No hay detalle suficiente para generar tendencias de licencias emitidas.")
        return

    mensual = build_monthly_license_summary(detalle_year)
    riesgo = (
        detalle_year.groupby("RIESGO_AGRUPADO", observed=False)["EXPEDIENTES"]
        .sum()
        .sort_values(ascending=False)
    )

    total_exp = int(detalle_year["EXPEDIENTES"].sum())
    total_rec = float(detalle_year["TOTAL"].sum())
    texto = (
        f"- En **{year}** se registran **{total_exp:,} licencias emitidas** "
        f"por **S/ {total_rec:,.2f}** de recaudación.\n"
    )

    if not mensual.empty:
        pico_exp = mensual.sort_values(["EXPEDIENTES", "MES_NUM"], ascending=[False, True]).iloc[0]
        pico_rec = mensual.sort_values(["RECAUDACION", "MES_NUM"], ascending=[False, True]).iloc[0]
        texto += (
            f"- El mes con mayor número de licencias emitidas fue **{pico_exp['MES']}**, "
            f"con **{int(pico_exp['EXPEDIENTES']):,} registros**.\n"
            f"- La mayor recaudación mensual se registró en **{pico_rec['MES']}**, "
            f"con **S/ {float(pico_rec['RECAUDACION']):,.2f}**.\n"
        )

    if not riesgo.empty:
        detalle_riesgo = ", ".join(f"{idx}: {int(val):,}" for idx, val in riesgo.items())
        texto += f"- La distribución por riesgo fue: **{detalle_riesgo}**.\n"

    if not tramites_year.empty:
        otros = tramites_year[~tramites_year["ES_LICENCIA_PRINCIPAL"]]
        if not otros.empty:
            resumen_otros = (
                otros.groupby("TIPO_PROCEDIMIENTO", observed=False)
                .agg(TRAMITES=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
                .sort_values("RECAUDACION", ascending=False)
            )
            principal = resumen_otros.iloc[0]
            texto += (
                f"- En trámites complementarios, el mayor ingreso corresponde a "
                f"**{resumen_otros.index[0]}**, con **{int(principal['TRAMITES']):,} trámites** "
                f"y **S/ {float(principal['RECAUDACION']):,.2f}**.\n"
            )

    st.info(texto)


def render_year_license_section(year, detalle_year, tramites_year):
    st.subheader("Licencias emitidas: temporales e indeterminadas")

    main_records = tramites_year[tramites_year["ES_LICENCIA_PRINCIPAL"]].copy()
    expedientes = int(detalle_year["EXPEDIENTES"].sum()) if not detalle_year.empty else len(main_records)
    ingresos = float(detalle_year["TOTAL"].sum()) if not detalle_year.empty else float(main_records["COSTO_NUM"].sum())
    riesgo_medio = int(detalle_year.loc[detalle_year["RIESGO_AGRUPADO"] == "MEDIO", "EXPEDIENTES"].sum()) if not detalle_year.empty else 0
    riesgo_alto = int(detalle_year.loc[detalle_year["RIESGO_AGRUPADO"] == "ALTO", "EXPEDIENTES"].sum()) if not detalle_year.empty else 0
    riesgo_muy_alto = int(detalle_year.loc[detalle_year["RIESGO_AGRUPADO"] == "MUY ALTO", "EXPEDIENTES"].sum()) if not detalle_year.empty else 0
    riesgo_alto_consolidado = int(detalle_year.loc[detalle_year["RIESGO_AGRUPADO"] == "ALTOS Y MUY ALTOS", "EXPEDIENTES"].sum()) if not detalle_year.empty else 0

    metricas_riesgo = [
        ("Expedientes", f"{expedientes:,}"),
        ("Ingresos", f"S/ {ingresos:,.2f}"),
        ("Riesgo medio", f"{riesgo_medio:,}"),
        ("Riesgo alto", f"{riesgo_alto:,}"),
        ("Riesgo muy alto", f"{riesgo_muy_alto:,}"),
    ]
    if riesgo_alto_consolidado:
        metricas_riesgo.append(("Alto y muy alto consolidado", f"{riesgo_alto_consolidado:,}"))

    show_metric_row(metricas_riesgo)

    if detalle_year.empty:
        st.info("No hay detalle mensual por riesgo para este año.")
        return

    if not {"MES", "MES_NUM"}.issubset(detalle_year.columns):
        tabla_detallada(detalle_year)
        return

    mensual_riesgo = (
        detalle_year.groupby(["MES", "MES_NUM", "RIESGO_AGRUPADO"], observed=False)
        .agg(EXPEDIENTES=("EXPEDIENTES", "sum"), RECAUDACION=("TOTAL", "sum"))
        .reset_index()
        .sort_values(["MES_NUM", "RIESGO_AGRUPADO"])
    )

    fig = px.bar(
        mensual_riesgo,
        x="MES",
        y="EXPEDIENTES",
        color="RIESGO_AGRUPADO",
        barmode="stack",
        text="EXPEDIENTES",
        category_orders={"MES": MONTH_ORDER},
        color_discrete_map=RISK_COLORS,
        height=410,
        labels={"MES": "Mes", "EXPEDIENTES": "Expedientes", "RIESGO_AGRUPADO": "Riesgo"},
    )
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", legend_title="Riesgo")
    st.plotly_chart(fig, use_container_width=True, key=f"licencias_anual_{year}_riesgo_mensual")

    render_year_license_monthly_charts(year, detalle_year, tramites_year)

    render_year_observations(year, detalle_year, tramites_year)

    tabla_detallada(detalle_year)


def render_year_group_section(year, title, df_group):
    st.subheader(title)
    if df_group.empty:
        st.info("No se encontraron registros para este grupo.")
        return

    total_tramites = len(df_group)
    total_ingresos = float(df_group["COSTO_NUM"].sum())
    promedio = total_ingresos / total_tramites if total_tramites else 0
    tipo_principal = df_group["TIPO_PROCEDIMIENTO"].mode().iloc[0] if not df_group["TIPO_PROCEDIMIENTO"].mode().empty else "-"

    show_metric_row(
        [
            ("Trámites", f"{total_tramites:,}"),
            ("Ingresos", f"S/ {total_ingresos:,.2f}"),
            ("Costo promedio", f"S/ {promedio:,.2f}"),
            ("Tipo principal", tipo_principal),
        ]
    )

    resumen = (
        df_group.groupby(["TIPO_PROCEDIMIENTO", "MES", "MES_NUM"], observed=False)
        .agg(TRAMITES=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
        .reset_index()
        .sort_values(["MES_NUM", "TIPO_PROCEDIMIENTO"])
    )

    fig = px.bar(
        resumen,
        x="MES",
        y="RECAUDACION",
        color="TIPO_PROCEDIMIENTO",
        barmode="group",
        text="RECAUDACION",
        category_orders={"MES": MONTH_ORDER},
        color_discrete_map=PROCEDURE_COLORS,
        height=390,
        labels={"MES": "Mes", "RECAUDACION": "Recaudación (S/)", "TIPO_PROCEDIMIENTO": "Tipo"},
    )
    fig.update_traces(texttemplate="S/ %{y:,.0f}", textposition="outside")
    fig.update_layout(plot_bgcolor="rgba(0,0,0,0)", legend_title="Tipo")
    section_key = normalize_text(title).replace(" ", "_").replace("/", "_")
    st.plotly_chart(fig, use_container_width=True, key=f"licencias_anual_{year}_{section_key}_ingresos_mensual")

    fig_cantidad = px.bar(
        resumen,
        x="MES",
        y="TRAMITES",
        color="TIPO_PROCEDIMIENTO",
        barmode="group",
        text="TRAMITES",
        category_orders={"MES": MONTH_ORDER},
        color_discrete_map=PROCEDURE_COLORS,
        height=360,
        labels={"MES": "Mes", "TRAMITES": "Trámites", "TIPO_PROCEDIMIENTO": "Tipo"},
    )
    fig_cantidad.update_traces(textposition="outside")
    fig_cantidad.update_layout(plot_bgcolor="rgba(0,0,0,0)", legend_title="Tipo")
    st.plotly_chart(fig_cantidad, use_container_width=True, key=f"licencias_anual_{year}_{section_key}_cantidad_mensual")

    st.dataframe(
        resumen.rename(
            columns={
                "TIPO_PROCEDIMIENTO": "Tipo de trámite",
                "MES": "Mes",
                "TRAMITES": "Trámites",
                "RECAUDACION": "Recaudación",
            }
        )[["Tipo de trámite", "Mes", "Trámites", "Recaudación"]],
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tipo de trámite": st.column_config.TextColumn("Tipo de trámite", width="large"),
            "Mes": st.column_config.TextColumn("Mes"),
            "Trámites": st.column_config.NumberColumn("Trámites", format="%d"),
            "Recaudación": st.column_config.NumberColumn("Recaudación", format="S/ %.2f"),
        },
    )


def render_year_income_section(year, tramites_year):
    st.subheader("Ingresos del año")
    if tramites_year.empty:
        st.info("No hay registros detallados desde Drive para calcular ingresos por tipo.")
        return

    resumen_tipo = (
        tramites_year.groupby("TIPO_PROCEDIMIENTO", observed=False)
        .agg(TRAMITES=("FECHA_RESOLUCION", "size"), RECAUDACION=("COSTO_NUM", "sum"))
        .reset_index()
        .sort_values("RECAUDACION", ascending=False)
    )

    fig = px.bar(
        resumen_tipo,
        x="TIPO_PROCEDIMIENTO",
        y="RECAUDACION",
        color="TIPO_PROCEDIMIENTO",
        text="RECAUDACION",
        color_discrete_map=PROCEDURE_COLORS,
        height=410,
        labels={"TIPO_PROCEDIMIENTO": "Tipo de trámite", "RECAUDACION": "Recaudación (S/)"},
    )
    fig.update_traces(texttemplate="S/ %{y:,.2f}", textposition="outside")
    fig.update_layout(
        plot_bgcolor="rgba(0,0,0,0)",
        xaxis_title="Tipo de trámite",
        yaxis_title="Recaudación (S/)",
        showlegend=False,
    )
    st.plotly_chart(fig, use_container_width=True, key=f"licencias_anual_{year}_ingresos_tipo")

    st.dataframe(
        resumen_tipo.rename(
            columns={
                "TIPO_PROCEDIMIENTO": "Tipo de trámite",
                "TRAMITES": "Trámites",
                "RECAUDACION": "Recaudación",
            }
        ),
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tipo de trámite": st.column_config.TextColumn("Tipo de trámite", width="large"),
            "Trámites": st.column_config.NumberColumn("Trámites", format="%d"),
            "Recaudación": st.column_config.NumberColumn("Recaudación", format="S/ %.2f"),
        },
    )


def render_year_consolidated_section(year, detalle_year, resumen_year, tramites_year):
    st.subheader("Consolidado general")
    if not resumen_year.empty:
        tabla_resumen_anual(resumen_year)

    if tramites_year.empty:
        if not detalle_year.empty:
            tabla_detallada(detalle_year)
        return

    total = pd.DataFrame(
        [
            {
                "Año": year,
                "Total trámites": len(tramites_year),
                "Ingresos totales": float(tramites_year["COSTO_NUM"].sum()),
                "Licencias temp./indet.": int(tramites_year["ES_LICENCIA_PRINCIPAL"].sum()),
                "Otros trámites": int((~tramites_year["ES_LICENCIA_PRINCIPAL"]).sum()),
            }
        ]
    )
    st.dataframe(
        total,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año"),
            "Total trámites": st.column_config.NumberColumn("Total trámites", format="%d"),
            "Ingresos totales": st.column_config.NumberColumn("Ingresos totales", format="S/ %.2f"),
            "Licencias temp./indet.": st.column_config.NumberColumn("Licencias temp./indet.", format="%d"),
            "Otros trámites": st.column_config.NumberColumn("Otros trámites", format="%d"),
        },
    )
    tabla_detalle_tramites(tramites_year)


def render_year_licencias(year, detalle_df, resumen_df, tramites_df):
    detalle_year = filter_period(detalle_df, year)
    resumen_year = filter_period(resumen_df, year)
    tramites_year = filter_period(tramites_df, year)

    st.subheader(f"Reporte {year}")
    if detalle_year.empty and resumen_year.empty and tramites_year.empty:
        st.info(f"No hay datos para {year}.")
        return

    total_expedientes = int(resumen_year["EXPEDIENTES"].sum()) if not resumen_year.empty else 0
    recaudacion_base = float(resumen_year["RECAUDACION"].sum()) if not resumen_year.empty else 0
    ingresos_detallados = float(tramites_year["COSTO_NUM"].sum()) if not tramites_year.empty else recaudacion_base
    total_tramites = len(tramites_year) if not tramites_year.empty else total_expedientes

    show_metric_row(
        [
            ("Expedientes de licencias", f"{total_expedientes:,}"),
            ("Recaudación de licencias", f"S/ {recaudacion_base:,.2f}"),
            ("Trámites registrados", f"{total_tramites:,}"),
            ("Ingresos registrados", f"S/ {ingresos_detallados:,.2f}"),
        ]
    )
    st.markdown("---")

    render_year_license_section(year, detalle_year, tramites_year)
    st.markdown("---")

    if not tramites_year.empty:
        duplicados = tramites_year[
            tramites_year["PROCEDIMIENTO_NORMALIZADO"] == "DUPLICADO DE LICENCIA DE FUNCIONAMIENTO"
        ].copy()
        transferencias = tramites_year[
            tramites_year["PROCEDIMIENTO_NORMALIZADO"] == "TRANSFERENCIA DE LICENCIA DE FUNCIONAMIENTO"
        ].copy()
        improcedentes = tramites_year[
            tramites_year["PROCEDIMIENTO_NORMALIZADO"] == "LICENCIA DE FUNCIONAMIENTO"
        ].copy()

        render_year_group_section(year, "Duplicados", duplicados)
        st.markdown("---")

        render_year_group_section(year, "Transferencias", transferencias)
        st.markdown("---")

        render_year_group_section(year, "Improcedentes con pago", improcedentes)
        st.markdown("---")

        render_year_income_section(year, tramites_year)
        st.markdown("---")
    else:
        st.info("Para este año solo se muestra el consolidado local; no hay detalle por tipo de trámite desde Drive.")
        st.markdown("---")

    render_year_consolidated_section(year, detalle_year, resumen_year, tramites_year)


def tabla_resumen_anual(resumen_df):
    st.subheader("Tabla resumen anual")

    tabla_df = resumen_df.copy().rename(columns={
        "PERIODO": "Año",
        "EXPEDIENTES": "Nro. de expedientes",
        "RECAUDACION": "Recaudación"
    })

    st.dataframe(
        tabla_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año", width="medium"),
            "Nro. de expedientes": st.column_config.NumberColumn("Nro. de expedientes", format="%d"),
            "Recaudación": st.column_config.NumberColumn("Recaudación", format="S/ %.2f"),
        }
    )


def tabla_detallada(detalle_df):
    st.subheader("Detalle por riesgo")

    tabla_df = detalle_df.copy().rename(columns={
        "PERIODO": "Año",
        "RIESGO_DETALLE": "Riesgo",
        "EXPEDIENTES": "Expedientes",
        "COSTO": "Costo",
        "TOTAL": "Total"
    })

    tabla_df = tabla_df[["Año", "Riesgo", "Expedientes", "Costo", "Total"]]

    st.dataframe(
        tabla_df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Año": st.column_config.TextColumn("Año", width="small"),
            "Riesgo": st.column_config.TextColumn("Riesgo", width="large"),
            "Expedientes": st.column_config.NumberColumn("Expedientes", format="%d"),
            "Costo": st.column_config.NumberColumn("Costo", format="S/ %.2f"),
            "Total": st.column_config.NumberColumn("Total", format="S/ %.2f"),
        }
    )


def observaciones(resumen_df, detalle_df=None):
    st.subheader("Observaciones")

    periodo_max_exp = resumen_df.loc[resumen_df["EXPEDIENTES"].idxmax(), "PERIODO"]
    periodo_max_rec = resumen_df.loc[resumen_df["RECAUDACION"].idxmax(), "PERIODO"]
    fuente = resumen_df.attrs.get("source")
    if fuente == "drive":
        texto_fuente = "- Los datos se actualizan automáticamente desde Google Drive.\n"
    elif fuente == "mixed":
        texto_fuente = "- Se conserva el histórico local y el año actual se actualiza desde Google Drive.\n"
    else:
        texto_fuente = "- Los totales anuales se han consignado según el cuadro consolidado fuente.\n"

    texto = (
        f"- En el período analizado, el año con mayor número de licencias emitidas fue "
        f"**{periodo_max_exp}**, con **{int(resumen_df['EXPEDIENTES'].max()):,} expedientes**.\n"
        f"- El año con mayor recaudación fue **{periodo_max_rec}**, con "
        f"**S/ {float(resumen_df['RECAUDACION'].max()):,.2f}**.\n"
    )

    if detalle_df is not None and not detalle_df.empty and {"MES", "MES_NUM"}.issubset(detalle_df.columns):
        mensual_general = (
            detalle_df.groupby(["MES_NUM", "MES"], observed=False)["EXPEDIENTES"]
            .sum()
            .reset_index(name="TOTAL")
            .sort_values(["TOTAL", "MES_NUM"], ascending=[False, True])
        )
        if not mensual_general.empty:
            mes_general = mensual_general.iloc[0]
            texto += (
                f"- El mes con mayor concentración de licencias en todo el período fue "
                f"**{mes_general['MES']}**, con **{int(mes_general['TOTAL']):,} registros acumulados**.\n"
            )

        pico_por_anio = (
            detalle_df.groupby(["PERIODO", "MES_NUM", "MES"], observed=False)["EXPEDIENTES"]
            .sum()
            .reset_index(name="TOTAL")
            .sort_values(["PERIODO", "TOTAL", "MES_NUM"], ascending=[True, False, True])
            .drop_duplicates(subset=["PERIODO"])
            .sort_values("PERIODO")
        )
        for _, fila in pico_por_anio.iterrows():
            texto += (
                f"- En **{fila['PERIODO']}**, el mes con mayor número de licencias fue "
                f"**{fila['MES']}**, con **{int(fila['TOTAL']):,} registros**.\n"
            )

    resumen_ordenado = resumen_df.copy()
    resumen_ordenado["ANIO_NUM"] = resumen_ordenado["PERIODO"].astype(str).str.extract(r"(\d{4})")[0].astype(int)
    resumen_ordenado = resumen_ordenado.sort_values("ANIO_NUM")
    filas = list(resumen_ordenado.itertuples(index=False))
    for anterior, actual in zip(filas, filas[1:]):
        variacion = variacion_porcentual(anterior.EXPEDIENTES, actual.EXPEDIENTES)
        if variacion is None:
            continue
        tendencia = "disminución" if variacion < 0 else "incremento"
        texto += (
            f"- Entre **{anterior.PERIODO} y {actual.PERIODO}** se observa una "
            f"**{tendencia} de {abs(variacion):.1f}%** en licencias emitidas.\n"
        )

    if detalle_df is not None and not detalle_df.empty and "MES_NUM" in detalle_df.columns:
        max_year = str(resumen_ordenado.iloc[-1]["PERIODO"])
        detalle_max = detalle_df[detalle_df["PERIODO"].astype(str) == max_year]
        if not detalle_max.empty:
            ultimo_mes = int(detalle_max["MES_NUM"].max())
            if ultimo_mes < 12:
                texto += (
                    f"- El año **{max_year}** presenta información parcial hasta "
                    f"**{MONTH_MAP.get(ultimo_mes, 'Sin fecha')}**, por lo que su comparación "
                    "con años completos debe interpretarse con cautela.\n"
                )

    texto += texto_fuente
    st.info(texto)

def show_licencias_funcionamiento_module():
    st.header("Módulo de Licencias de Funcionamiento")
    render_drive_refresh_button()
    st.markdown("---")

    detalle_df, resumen_df, tramites_df = load_licencias_funcionamiento_data()

    if resumen_df is None or resumen_df.empty:
        st.error("No se pudieron cargar los datos.")
        return

    if resumen_df.attrs.get("source") == "drive":
        st.success("Datos actualizados desde Google Drive: hojas RESOLUCIONES 2025 y RESOLUCIONES 2026.")
    elif resumen_df.attrs.get("source") == "mixed":
        st.success("Histórico local conservado y años disponibles actualizados desde Google Drive.")

    tabs = st.tabs(["General", "2023", "2024", "2025", "2026", "Actualizar"])

    with tabs[0]:
        render_general_licencias(detalle_df, resumen_df, tramites_df)

    for tab, year in zip(tabs[1:], ["2023", "2024", "2025", "2026"]):
        with tab:
            render_year_licencias(year, detalle_df, resumen_df, tramites_df)

    with tabs[5]:
        render_resoluciones_2026_update_tools()
