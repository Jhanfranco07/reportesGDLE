import streamlit as st
import plotly.io as pio
from utils.helpers import install_excel_dataframe_download

from modules.anuncios_publicitarios import show_anuncios_publicitarios_module
from modules.comercio_ambulatorio import show_comercio_ambulatorio_module
from modules.ferias import show_ferias_module
from modules.licencias_funcionamiento import show_licencias_funcionamiento_module
from modules.pachacard import show_pachacard_module
from modules.pachambear import show_pachambear_module
from modules.pachamikuy import show_pachamikuy_module


st.set_page_config(
    page_title="Reportes GDE",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded",
)

pio.templates.default = "plotly_white"
install_excel_dataframe_download(st)

MODULES = {
    "PACHAMIKUY": {
        "description": "Ferias y actividad mensual del programa.",
        "handler": show_pachamikuy_module,
    },
    "PACHACARD": {
        "description": "Indicadores de tarjetas emitidas, categorias y comercios afiliados.",
        "handler": show_pachacard_module,
    },
    "PACHAMBEAR": {
        "description": "Reporte laboral y gestion de atenciones.",
        "handler": show_pachambear_module,
    },
    "FERIAS": {
        "description": "Ferias ambulatorias por sede, rubro y recaudacion.",
        "handler": show_ferias_module,
    },
    "COMERCIO AMBULATORIO": {
        "description": "Autorizaciones emitidas y comportamiento anual.",
        "handler": show_comercio_ambulatorio_module,
    },
    "ANUNCIOS PUBLICITARIOS": {
        "description": "Certificados, tipos de panel y recaudacion.",
        "handler": show_anuncios_publicitarios_module,
    },
    "LICENCIAS DE FUNCIONAMIENTO": {
        "description": "Expedientes por riesgo y recaudacion acumulada.",
        "handler": show_licencias_funcionamiento_module,
    },
}


def apply_professional_theme():
    st.markdown(
        """
        <style>
        :root {
            --gde-bg: #f4f7fb;
            --gde-surface: #ffffff;
            --gde-surface-soft: #f8fafc;
            --gde-border: #dde6f0;
            --gde-border-strong: #c8d4e3;
            --gde-text: #162033;
            --gde-muted: #667085;
            --gde-primary: #0f5f80;
            --gde-primary-dark: #0a445c;
            --gde-accent: #2f9e8f;
            --gde-shadow: 0 18px 45px rgba(20, 44, 72, 0.10);
        }

        html, body, [data-testid="stAppViewContainer"] {
            background:
                radial-gradient(circle at top left, rgba(47, 158, 143, 0.12), transparent 28rem),
                linear-gradient(180deg, #f8fbfd 0%, var(--gde-bg) 42%, #eef4f8 100%);
            color: var(--gde-text);
        }

        [data-testid="stAppViewContainer"] > .main {
            background: transparent;
        }

        .block-container {
            max-width: 1380px;
            padding: 2rem 2.4rem 4rem;
        }

        [data-testid="stSidebar"] {
            background: linear-gradient(180deg, #0d3447 0%, #10283a 55%, #14253a 100%);
            border-right: 1px solid rgba(255, 255, 255, 0.10);
        }

        [data-testid="stSidebar"] * {
            color: #eef7fb;
        }

        [data-testid="stSidebar"] [role="radiogroup"] label {
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 8px;
            padding: 0.48rem 0.62rem;
            margin-bottom: 0.35rem;
            background: rgba(255, 255, 255, 0.045);
            transition: background 0.16s ease, border-color 0.16s ease;
        }

        [data-testid="stSidebar"] [role="radiogroup"] label:hover {
            background: rgba(255, 255, 255, 0.10);
            border-color: rgba(255, 255, 255, 0.26);
        }

        h1, h2, h3 {
            color: var(--gde-text);
            letter-spacing: 0;
        }

        h1 {
            font-size: 2rem;
            line-height: 1.15;
            margin-bottom: 0.35rem;
        }

        h2, h3 {
            margin-top: 0.8rem;
        }

        hr {
            margin: 1.25rem 0;
            border-color: rgba(15, 95, 128, 0.14);
        }

        .gde-shell {
            border: 1px solid var(--gde-border);
            border-radius: 10px;
            padding: 1.35rem 1.45rem;
            margin-bottom: 1.25rem;
            background: rgba(255, 255, 255, 0.86);
            box-shadow: var(--gde-shadow);
        }

        .gde-kicker {
            color: var(--gde-primary);
            font-weight: 800;
            font-size: 0.78rem;
            letter-spacing: 0.08em;
            text-transform: uppercase;
            margin-bottom: 0.45rem;
        }

        .gde-title-row {
            display: flex;
            align-items: flex-end;
            justify-content: space-between;
            gap: 1rem;
            flex-wrap: wrap;
        }

        .gde-title {
            margin: 0;
            font-size: 2.05rem;
            font-weight: 850;
            color: var(--gde-text);
        }

        .gde-subtitle {
            margin: 0.5rem 0 0;
            max-width: 860px;
            color: var(--gde-muted);
            font-size: 1rem;
            line-height: 1.55;
        }

        .gde-pill {
            display: inline-flex;
            align-items: center;
            border: 1px solid rgba(47, 158, 143, 0.28);
            border-radius: 999px;
            padding: 0.45rem 0.75rem;
            color: var(--gde-primary-dark);
            background: rgba(47, 158, 143, 0.10);
            font-weight: 700;
            white-space: nowrap;
        }

        .gde-side-brand {
            padding: 0.2rem 0 1rem;
            border-bottom: 1px solid rgba(255, 255, 255, 0.14);
            margin-bottom: 1rem;
        }

        .gde-side-title {
            font-size: 1.05rem;
            font-weight: 850;
            margin: 0;
        }

        .gde-side-subtitle {
            margin: 0.3rem 0 0;
            color: rgba(238, 247, 251, 0.72);
            font-size: 0.84rem;
            line-height: 1.4;
        }

        .gde-side-note {
            margin-top: 1rem;
            padding: 0.85rem;
            border: 1px solid rgba(255, 255, 255, 0.12);
            border-radius: 8px;
            background: rgba(255, 255, 255, 0.06);
            color: rgba(238, 247, 251, 0.78);
            font-size: 0.82rem;
            line-height: 1.45;
        }

        [data-testid="stMetric"] {
            border: 1px solid var(--gde-border);
            border-radius: 8px;
            padding: 0.95rem 1rem;
            background: linear-gradient(180deg, #ffffff 0%, #f9fbfd 100%);
            box-shadow: 0 10px 28px rgba(20, 44, 72, 0.07);
        }

        [data-testid="stMetricLabel"] {
            color: var(--gde-muted);
            font-size: 0.82rem;
            font-weight: 700;
        }

        [data-testid="stMetricValue"] {
            color: var(--gde-primary-dark);
            font-weight: 850;
        }

        [data-testid="stMetricValue"] > div {
            overflow: visible;
            text-overflow: clip;
            white-space: normal;
            overflow-wrap: anywhere;
            font-size: clamp(1.35rem, 2.1vw, 2rem);
            line-height: 1.12;
        }

        .gde-metric-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(150px, 1fr));
            gap: 0.9rem;
            margin: 0.35rem 0 1rem;
        }

        .gde-metric-card {
            border: 1px solid var(--gde-border);
            border-radius: 8px;
            padding: 0.95rem 1rem;
            background: linear-gradient(180deg, #ffffff 0%, #f9fbfd 100%);
            box-shadow: 0 10px 28px rgba(20, 44, 72, 0.07);
            min-width: 0;
        }

        .gde-metric-label {
            color: var(--gde-muted);
            font-size: 0.78rem;
            font-weight: 760;
            margin-bottom: 0.45rem;
        }

        .gde-metric-value {
            color: var(--gde-primary-dark);
            font-size: clamp(1.35rem, 2.1vw, 2rem);
            font-weight: 850;
            line-height: 1.12;
            overflow-wrap: anywhere;
        }

        [data-testid="stPlotlyChart"],
        [data-testid="stDataFrame"],
        [data-testid="stTable"] {
            border: 1px solid var(--gde-border);
            border-radius: 10px;
            padding: 0.65rem;
            background: var(--gde-surface);
            box-shadow: 0 12px 32px rgba(20, 44, 72, 0.06);
        }

        div[data-baseweb="select"] > div,
        [data-testid="stSelectbox"] div[data-baseweb="select"] > div {
            border-radius: 8px;
            border-color: var(--gde-border-strong);
            background: #ffffff;
        }

        div.stButton > button {
            border-radius: 8px;
            border: 1px solid var(--gde-border-strong);
            background: #ffffff;
            color: var(--gde-primary-dark);
            font-weight: 750;
            min-height: 2.45rem;
            box-shadow: 0 6px 18px rgba(20, 44, 72, 0.06);
        }

        div.stButton > button:hover {
            border-color: var(--gde-primary);
            color: var(--gde-primary-dark);
            background: #f3fbfa;
        }

        .stAlert {
            border-radius: 8px;
        }

        @media (max-width: 768px) {
            .block-container {
                padding: 1rem 1rem 3rem;
            }

            .gde-shell {
                padding: 1rem;
            }

            .gde-title {
                font-size: 1.55rem;
            }
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_sidebar(selected_module):
    st.sidebar.markdown(
        """
        <div class="gde-side-brand">
            <p class="gde-side-title">Reportes GDE</p>
            <p class="gde-side-subtitle">Gerencia de Licencias y Desarrollo Economico</p>
        </div>
        """,
        unsafe_allow_html=True,
    )

    module = st.sidebar.radio(
        "Modulo",
        tuple(MODULES.keys()),
        index=list(MODULES.keys()).index(selected_module),
    )

    st.sidebar.markdown(
        f"""
        <div class="gde-side-note">
            <strong>{module}</strong><br>
            {MODULES[module]["description"]}
        </div>
        """,
        unsafe_allow_html=True,
    )
    return module


def render_module_selector(selected_module):
    module = st.selectbox(
        "Cambiar módulo / seleccionar reporte",
        tuple(MODULES.keys()),
        index=list(MODULES.keys()).index(selected_module),
        label_visibility="visible",
    )
    return module


def render_header(module):
    st.markdown(
        f"""
        <section class="gde-shell">
            <div class="gde-title-row">
                <div>
                    <div class="gde-kicker">Panel de gestion municipal</div>
                    <h1 class="gde-title">Reportes Estadisticos GDE</h1>
                    <p class="gde-subtitle">
                        Seguimiento operativo de licencias, comercio, ferias, programas y recaudacion.
                        La vista activa esta optimizada para lectura rapida, comparacion y toma de decisiones.
                    </p>
                </div>
                <div class="gde-pill">{module}</div>
            </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def main():
    apply_professional_theme()

    default_module = "PACHAMIKUY"
    if "active_module" not in st.session_state:
        st.session_state.active_module = default_module

    sidebar_module = render_sidebar(st.session_state.active_module)
    if sidebar_module != st.session_state.active_module:
        st.session_state.active_module = sidebar_module
        st.rerun()

    selected_module = render_module_selector(st.session_state.active_module)
    if selected_module != st.session_state.active_module:
        st.session_state.active_module = selected_module
        st.rerun()

    render_header(selected_module)

    MODULES[selected_module]["handler"]()


if __name__ == "__main__":
    main()
