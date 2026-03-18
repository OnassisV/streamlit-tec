from __future__ import annotations

from datetime import datetime, time, timedelta
from io import BytesIO
from pathlib import Path
import re
import zipfile

import pandas as pd
import streamlit as st
from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from PIL import Image, ImageChops
from streamlit.errors import StreamlitSecretNotFoundError

from app_auth import (
    ROLE_DEFINITIONS,
    clear_authenticated_user,
    describe_access_window,
    get_authenticated_user,
    set_authenticated_user,
    user_has_permission,
)
from app_storage import build_storage_backend


APP_TITLE = "Procesador TEC de Peajes"
APP_NAV_KEY = "app_selected_page"
CONTRACTOR_LOGO_PATH = Path(__file__).parent / "ChatGPT Image 18 mar 2026, 03_37_00 a.m..png"

MODULE_CATALOG = [
    {
        "page": "TEC",
        "title": "TEC",
        "eyebrow": "Modulo operativo",
        "description": "Procesamiento de bases de peaje, limpieza de placas, recuperacion de tiempos y generacion de entregables.",
        "status": "Activo",
        "accent": "#0f3d91",
        "icon_svg": '<path d="M4 19h16" /><path d="M7 16V9" /><path d="M12 16V5" /><path d="M17 16v-3" />',
    },
    {
        "page": "Relevamientos",
        "title": "Relevamientos",
        "eyebrow": "Modulo planificado",
        "description": "Registro estructurado de levantamiento de campo, evidencias, hallazgos y seguimiento por estacion o activo.",
        "status": "Proximamente",
        "accent": "#1849a9",
        "icon_svg": '<path d="M9 11l3 3L22 4" /><path d="M21 12v7a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h11" />',
    },
    {
        "page": "Auditorias",
        "title": "Auditorias",
        "eyebrow": "Modulo planificado",
        "description": "Control de auditorias operativas, observaciones, estados de cumplimiento y trazabilidad de acciones correctivas.",
        "status": "Proximamente",
        "accent": "#245cc6",
        "icon_svg": '<path d="M9 3h6l5 5v10a2 2 0 0 1-2 2H6a2 2 0 0 1-2-2V5a2 2 0 0 1 2-2h3" /><path d="M9 3v5h5" /><path d="M9 14h6" /><path d="M9 18h6" />',
    },
    {
        "page": "Satisfaccion",
        "title": "Satisfaccion",
        "eyebrow": "Modulo planificado",
        "description": "Seguimiento de experiencia de usuario, encuestas, indicadores de percepcion y resumenes ejecutivos.",
        "status": "Proximamente",
        "accent": "#2f6ddc",
        "icon_svg": '<path d="M8 14s1.5 2 4 2 4-2 4-2" /><path d="M9 9h.01" /><path d="M15 9h.01" /><path d="M12 21c4.97 0 9-3.582 9-8s-4.03-8-9-8-9 3.582-9 8c0 2.394 1.184 4.542 3.06 6.008L6 21l2.563-1.025A10.11 10.11 0 0 0 12 21Z" />',
    },
    {
        "page": "Flujogramas",
        "title": "Flujogramas",
        "eyebrow": "Modulo planificado",
        "description": "Biblioteca de procesos, mapas operativos, versionado documental y vistas navegables de flujo de trabajo.",
        "status": "Proximamente",
        "accent": "#3b7be6",
        "icon_svg": '<path d="M12 3v4" /><path d="M12 17v4" /><path d="M5 12H3" /><path d="M21 12h-2" /><rect x="8" y="7" width="8" height="10" rx="2" /><path d="M8 12H5" /><path d="M19 12h-3" />',
    },
]

MODULE_PLACEHOLDERS = {
    "Relevamientos": {
        "headline": "Levantamiento de campo con estructura comun",
        "summary": "Este espacio quedara listo para programar formularios, evidencias y seguimiento de hallazgos por sede o peaje.",
        "items": [
            "fichas por punto relevado",
            "adjuntos fotograficos y comentarios",
            "estado de observaciones y responsables",
        ],
    },
    "Auditorias": {
        "headline": "Control operativo con trazabilidad",
        "summary": "Aqui podras concentrar revisiones, no conformidades, evidencias y cierres por auditoria.",
        "items": [
            "programacion de auditorias",
            "matriz de hallazgos y severidad",
            "seguimiento de planes de accion",
        ],
    },
    "Satisfaccion": {
        "headline": "Indicadores de experiencia y percepcion",
        "summary": "El modulo servira para centralizar encuestas, cortes por canal y tableros de satisfaccion.",
        "items": [
            "carga de encuestas y bases",
            "resumenes por periodo o sede",
            "hallazgos de experiencia del usuario",
        ],
    },
    "Flujogramas": {
        "headline": "Mapa visual de procesos y documentos",
        "summary": "Quedara como repositorio de diagramas, procedimientos y navegacion por proceso.",
        "items": [
            "catalogo de procesos por area",
            "versionado de diagramas y anexos",
            "consulta rapida para operacion y control",
        ],
    },
}

EXPECTED_COLUMNS = [
    "PEAJE",
    "CASETA",
    "SENTIDO",
    "FECHA",
    "VEHICULO",
    "PLACA",
    "LLEGADA COLA",
    "LLEGADA CASETA",
    "SALIDA CASETA",
    "T. TEC",
    "T. CASETA",
]

TIME_ALIASES = ["T1", "T2", "T3"]
TIME_COLUMNS = {
    "LLEGADA COLA": "T1",
    "LLEGADA CASETA": "T2",
    "SALIDA CASETA": "T3",
}
TIME_GROUP_KEYS = ["PEAJE", "CASETA", "SENTIDO", "FECHA_DIA"]

EXPECTED_PLATE_LENGTHS = {6}
COMMON_PATTERNS = {"AAA999", "A9A999", "A99999", "AA9999", "AAAA99", "AA999"}
PLACEHOLDER_PLATES = {"AAA111", "BBB999", "ABC123", "XXX", "X"}

CONFUSION_LOOKUP = {
    "0": {"O"},
    "O": {"0"},
    "1": {"I"},
    "I": {"1"},
    "5": {"S"},
    "S": {"5"},
    "8": {"B"},
    "B": {"8"},
    "2": {"Z"},
    "Z": {"2"},
    "6": {"G"},
    "G": {"6"},
}

ACCIONES_PLACA_VALIDAS = {
    "corregir_a_normalizada",
    "corregir_a_sugerida",
    "corregir_a_coincidencia_lista",
    "corregir_a_recorte_sufijo_x",
    "corregir_manual",
    "excluir_analisis_placa",
    "mantener_observada",
    "sin_cambio",
}

ACCIONES_PLACA_AJUSTE = {
    "corregir_a_normalizada",
    "corregir_a_sugerida",
    "corregir_a_coincidencia_lista",
    "corregir_a_recorte_sufijo_x",
    "corregir_manual",
}

DEFAULT_MANUAL_RULES = [
    {
        "activo": True,
        "tipo_regla": "corregir_placa",
        "placa_origen": "D4P750UN",
        "placa_destino": "D4P750",
        "longitud_objetivo": pd.NA,
        "comentario": "Correccion manual heredada del caso Norvial.",
    },
    {
        "activo": True,
        "tipo_regla": "corregir_placa",
        "placa_origen": "TDA824UN",
        "placa_destino": "TDA824",
        "longitud_objetivo": pd.NA,
        "comentario": "Correccion manual heredada del caso Norvial.",
    },
    {
        "activo": True,
        "tipo_regla": "corregir_placa",
        "placa_origen": "D00191X",
        "placa_destino": "D0O191",
        "longitud_objetivo": pd.NA,
        "comentario": "Correccion manual heredada del caso Norvial.",
    },
    {
        "activo": True,
        "tipo_regla": "corregir_placa",
        "placa_origen": "BNI87D",
        "placa_destino": "BNI870",
        "longitud_objetivo": pd.NA,
        "comentario": "Correccion manual heredada del caso Norvial.",
    },
    {
        "activo": True,
        "tipo_regla": "corregir_placa",
        "placa_origen": "2PL067",
        "placa_destino": "ZPL067",
        "longitud_objetivo": pd.NA,
        "comentario": "Correccion manual heredada del caso Norvial.",
    },
    {
        "activo": True,
        "tipo_regla": "eliminar_placa",
        "placa_origen": "VCJV27CH",
        "placa_destino": pd.NA,
        "longitud_objetivo": pd.NA,
        "comentario": "Excluir del analisis.",
    },
    {
        "activo": True,
        "tipo_regla": "eliminar_placa",
        "placa_origen": "VCJV28CH",
        "placa_destino": pd.NA,
        "longitud_objetivo": pd.NA,
        "comentario": "Excluir del analisis.",
    },
    {
        "activo": True,
        "tipo_regla": "eliminar_por_longitud",
        "placa_origen": pd.NA,
        "placa_destino": pd.NA,
        "longitud_objetivo": 1,
        "comentario": "Excluir placas de largo 1.",
    },
]

DEFAULT_CONFIG = {
    "aplicar_limpieza_placa": True,
    "aplicar_ruido_con_respaldo": True,
    "aplicar_ruido_sin_respaldo": True,
    "aplicar_coincidencia_unica_lista": True,
    "aplicar_confusion_visual": True,
    "aplicar_recorte_sufijo_x": True,
    "aplicar_exclusion_placeholders": True,
    "aplicar_reglas_manuales": True,
    "eliminar_bordes_caseta": True,
    "aplicar_interpolacion": True,
    "aplicar_mediana_local": True,
    "aplicar_donantes": True,
    "aplicar_swap_final_t2_t3": True,
}


def inject_global_styles() -> None:
    st.markdown(
        """
        <style>
            :root {
                --app-bg: #f3f7ff;
                --app-surface: rgba(255, 255, 255, 0.92);
                --app-surface-strong: #ffffff;
                --app-ink: #102347;
                --app-muted: #5c6f93;
                --app-line: rgba(15, 61, 145, 0.12);
                --app-accent: #0f3d91;
                --app-accent-soft: #e9f0ff;
                --app-shadow: 0 18px 45px rgba(9, 29, 74, 0.12);
            }

            .stApp {
                background:
                    radial-gradient(circle at top left, rgba(50, 100, 204, 0.18), transparent 28%),
                    linear-gradient(180deg, #eef4ff 0%, #f7faff 42%, #eef3fb 100%);
                color: var(--app-ink);
            }

            [data-testid="stSidebar"] {
                background: linear-gradient(180deg, #07162f 0%, #0d2448 100%);
                border-right: 1px solid rgba(255, 255, 255, 0.08);
            }

            [data-testid="stSidebar"] * {
                color: #f4f8ff;
            }

            [data-testid="stSidebar"] .stRadio label,
            [data-testid="stSidebar"] .stCheckbox label,
            [data-testid="stSidebar"] .stMarkdown,
            [data-testid="stSidebar"] .stCaption {
                color: #f4f8ff;
            }

            [data-testid="stSidebar"] .stButton > button {
                border-radius: 18px;
                border: 1px solid rgba(255, 255, 255, 0.18);
                background: rgba(255, 255, 255, 0.08);
                color: #f4f8ff;
                box-shadow: none;
            }

            [data-testid="stSidebar"] .stButton > button:hover {
                border-color: rgba(255, 255, 255, 0.3);
                background: rgba(255, 255, 255, 0.16);
                color: #ffffff;
            }

            [data-testid="stSidebar"] .stButton > button[kind="primary"] {
                background: linear-gradient(135deg, #2a7de1 0%, #59b2ff 100%);
                border-color: transparent;
                color: #ffffff;
            }

            [data-testid="stSidebar"] .stButton > button[kind="secondary"] {
                background: rgba(255, 255, 255, 0.1);
                color: #f4f8ff;
            }

            [data-testid="stSidebar"] .stButton > button p {
                font-size: 1.05rem;
                font-weight: 600;
            }

            [data-testid="stSidebar"] .stButton > button#sidebar_users_button {
                min-height: 3.25rem;
            }

            [data-testid="stSidebar"] .stButton > button#sidebar_users_button p {
                font-size: 1.45rem;
                line-height: 1;
            }

            .sidebar-brand-wrap {
                padding: 0.35rem 0 1rem;
            }

            .sidebar-brand-kicker {
                text-transform: uppercase;
                letter-spacing: 0.14em;
                font-size: 0.68rem;
                font-weight: 700;
                color: rgba(214, 230, 255, 0.7);
                margin-bottom: 0.35rem;
            }

            .sidebar-brand-note {
                font-size: 0.82rem;
                line-height: 1.45;
                color: rgba(230, 239, 255, 0.76);
                margin-top: 0.45rem;
            }

            .sidebar-text-link {
                margin-top: 0.55rem;
            }

            .sidebar-text-link a {
                color: #8fd0ff;
                text-decoration: none;
                font-size: 0.98rem;
                font-weight: 700;
            }

            .sidebar-text-link a:hover {
                color: #d6efff;
                text-decoration: underline;
            }

            .auth-layout-title {
                font-size: 1.1rem;
                font-weight: 700;
                color: var(--app-ink);
                margin: 0.4rem 0 0.25rem;
            }

            .auth-layout-copy {
                color: var(--app-muted);
                line-height: 1.7;
                margin-bottom: 1rem;
            }

            .auth-card {
                border-radius: 26px;
                border: 1px solid rgba(15, 61, 145, 0.12);
                background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(242,247,255,0.96));
                box-shadow: 0 18px 40px rgba(8, 29, 74, 0.09);
                padding: 1.4rem 1.4rem 1rem;
                margin-bottom: 1rem;
            }

            .auth-card-kicker {
                text-transform: uppercase;
                letter-spacing: 0.14em;
                font-size: 0.72rem;
                color: var(--app-accent);
                font-weight: 700;
                margin-bottom: 0.55rem;
            }

            .auth-card-title {
                font-size: 1.45rem;
                font-weight: 700;
                color: var(--app-ink);
                margin-bottom: 0.35rem;
            }

            .auth-card-copy {
                color: var(--app-muted);
                line-height: 1.65;
                margin-bottom: 0;
            }

            .brand-panel {
                border-radius: 24px;
                border: 1px solid rgba(255,255,255,0.14);
                background: rgba(255,255,255,0.08);
                padding: 1rem;
                backdrop-filter: blur(8px);
                margin-top: 1rem;
            }

            .brand-panel-copy {
                color: rgba(240, 246, 255, 0.82);
                font-size: 0.92rem;
                line-height: 1.6;
                margin-top: 0.7rem;
            }

            .hero-panel {
                padding: 2rem 2.2rem;
                border-radius: 28px;
                background:
                    linear-gradient(135deg, rgba(5, 20, 44, 0.96) 0%, rgba(15, 61, 145, 0.94) 55%, rgba(74, 125, 218, 0.88) 100%);
                box-shadow: var(--app-shadow);
                position: relative;
                overflow: hidden;
                margin-bottom: 1.2rem;
                color: #f6f9ff;
            }

            .hero-panel::after {
                content: "";
                position: absolute;
                inset: auto -90px -90px auto;
                width: 240px;
                height: 240px;
                border-radius: 50%;
                background: radial-gradient(circle, rgba(255,255,255,0.22), rgba(255,255,255,0.02) 70%);
            }

            .hero-kicker {
                text-transform: uppercase;
                letter-spacing: 0.18em;
                font-size: 0.72rem;
                color: rgba(235, 243, 255, 0.78);
                margin-bottom: 0.7rem;
                font-weight: 700;
            }

            .hero-title {
                font-size: 2.5rem;
                line-height: 1.04;
                font-weight: 700;
                margin: 0;
                max-width: 12ch;
            }

            .hero-copy {
                font-size: 1rem;
                line-height: 1.7;
                max-width: 62ch;
                margin-top: 0.9rem;
                color: rgba(245, 248, 255, 0.86);
            }

            .metrics-strip {
                display: grid;
                grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
                gap: 0.9rem;
                margin: 1rem 0 1.3rem;
            }

            .metric-tile {
                border: 1px solid rgba(255, 255, 255, 0.16);
                background: rgba(255, 255, 255, 0.08);
                backdrop-filter: blur(8px);
                border-radius: 18px;
                padding: 1rem 1rem 0.9rem;
            }

            .metric-value {
                font-size: 1.7rem;
                font-weight: 700;
                color: #ffffff;
            }

            .metric-label {
                margin-top: 0.2rem;
                font-size: 0.85rem;
                color: rgba(238, 245, 255, 0.78);
            }

            .section-heading {
                font-size: 1.25rem;
                font-weight: 700;
                color: var(--app-ink);
                margin: 1rem 0 0.3rem;
            }

            .section-copy {
                color: var(--app-muted);
                margin-bottom: 1rem;
            }

            .module-card {
                min-height: 240px;
                border-radius: 24px;
                padding: 1.2rem;
                background: linear-gradient(180deg, rgba(255,255,255,0.98), rgba(244,248,255,0.92));
                border: 1px solid var(--app-line);
                box-shadow: 0 14px 35px rgba(12, 31, 76, 0.08);
                position: relative;
                overflow: hidden;
                margin-bottom: 0.75rem;
            }

            .module-card::before {
                content: "";
                position: absolute;
                inset: 0 0 auto 0;
                height: 5px;
                background: var(--module-accent);
            }

            .module-icon {
                width: 52px;
                height: 52px;
                border-radius: 16px;
                display: grid;
                place-items: center;
                color: var(--module-accent);
                background: rgba(15, 61, 145, 0.08);
                border: 1px solid rgba(15, 61, 145, 0.12);
            }

            .module-icon svg {
                width: 26px;
                height: 26px;
                stroke: currentColor;
                stroke-width: 1.7;
                fill: none;
                stroke-linecap: round;
                stroke-linejoin: round;
            }

            .module-eyebrow {
                margin-top: 1rem;
                text-transform: uppercase;
                letter-spacing: 0.12em;
                font-size: 0.72rem;
                font-weight: 700;
                color: var(--module-accent);
            }

            .module-title {
                font-size: 1.35rem;
                font-weight: 700;
                margin: 0.45rem 0 0.35rem;
                color: var(--app-ink);
            }

            .module-description {
                color: var(--app-muted);
                line-height: 1.6;
                font-size: 0.95rem;
                min-height: 74px;
            }

            .module-status {
                display: inline-flex;
                align-items: center;
                gap: 0.4rem;
                border-radius: 999px;
                padding: 0.38rem 0.72rem;
                background: rgba(15, 61, 145, 0.08);
                color: var(--module-accent);
                font-size: 0.82rem;
                font-weight: 700;
            }

            .placeholder-card {
                border-radius: 24px;
                border: 1px solid var(--app-line);
                background: rgba(255, 255, 255, 0.95);
                padding: 1.45rem;
                box-shadow: 0 16px 34px rgba(12, 31, 76, 0.08);
            }

            .placeholder-title {
                font-size: 1.55rem;
                font-weight: 700;
                color: var(--app-ink);
                margin-bottom: 0.35rem;
            }

            .placeholder-copy {
                color: var(--app-muted);
                line-height: 1.7;
            }

            .placeholder-list {
                margin: 0.9rem 0 0;
                padding-left: 1rem;
                color: var(--app-ink);
            }

            .placeholder-list li {
                margin-bottom: 0.45rem;
            }

            div[data-testid="stMetric"] {
                background: rgba(255,255,255,0.92);
                border: 1px solid var(--app-line);
                border-radius: 20px;
                padding: 0.8rem 1rem;
                box-shadow: 0 10px 28px rgba(12, 31, 76, 0.07);
            }

            @media (max-width: 900px) {
                .hero-title {
                    font-size: 2rem;
                }

                .hero-panel {
                    padding: 1.5rem;
                }
            }
        </style>
        """,
        unsafe_allow_html=True,
    )


def build_hero_panel(title: str, copy: str, kicker: str, metrics: list[tuple[str, str]] | None = None) -> str:
    metrics_markup = ""
    if metrics:
        metric_items = []
        for value, label in metrics:
            metric_items.append(
                f'<div class="metric-tile"><div class="metric-value">{value}</div><div class="metric-label">{label}</div></div>'
            )
        metrics_markup = f'<div class="metrics-strip">{"".join(metric_items)}</div>'

    return (
        '<section class="hero-panel">'
        f'<div class="hero-kicker">{kicker}</div>'
        f'<h1 class="hero-title">{title}</h1>'
        f'<div class="hero-copy">{copy}</div>'
        f'{metrics_markup}'
        '</section>'
    )


def build_module_card(module: dict) -> str:
    return (
        f'<section class="module-card" style="--module-accent:{module["accent"]};">'
        f'<div class="module-icon"><svg viewBox="0 0 24 24" aria-hidden="true">{module["icon_svg"]}</svg></div>'
        f'<div class="module-eyebrow">{module["eyebrow"]}</div>'
        f'<div class="module-title">{module["title"]}</div>'
        f'<div class="module-description">{module["description"]}</div>'
        f'<div class="module-status">{module["status"]}</div>'
        '</section>'
    )


@st.cache_data(show_spinner=False)
def load_contractor_logo() -> Image.Image | None:
    if not CONTRACTOR_LOGO_PATH.exists():
        return None

    image = Image.open(CONTRACTOR_LOGO_PATH).convert("RGBA")
    background = Image.new("RGBA", image.size, (255, 255, 255, 255))
    diff = ImageChops.difference(image, background)
    bbox = diff.getbbox()
    if bbox:
        left, top, right, bottom = bbox
        padding = 24
        left = max(0, left - padding)
        top = max(0, top - padding)
        right = min(image.width, right + padding)
        bottom = min(image.height, bottom + padding)
        image = image.crop((left, top, right, bottom))

    return image


def render_contractor_branding() -> None:
    logo_image = load_contractor_logo()
    if logo_image is None:
        return

    st.markdown('<div class="sidebar-brand-wrap"><div class="sidebar-brand-kicker">Empresa contratante</div></div>', unsafe_allow_html=True)
    st.image(logo_image, use_container_width=True)
    st.markdown(
        '<div class="sidebar-brand-note">Identidad visual del contratante integrada en el acceso y en la portada del aplicativo.</div>',
        unsafe_allow_html=True,
    )


def navigate_to(page: str) -> None:
    if page == "Inicio":
        st.query_params.clear()
    else:
        st.query_params["page"] = page
    st.session_state[APP_NAV_KEY] = page
    st.rerun()


def get_requested_page() -> str | None:
    raw_page = st.query_params.get("page")
    if raw_page is None:
        return None
    if isinstance(raw_page, list):
        raw_page = raw_page[0] if raw_page else None
    if raw_page is None:
        return None
    page = str(raw_page).strip()
    return page or None


def render_home_page(current_user: dict | None) -> None:
    user_label = current_user["full_name"] if current_user else "Equipo CIDATT"
    role_label = current_user["role_label"] if current_user else "Acceso directo"
    hero_col, brand_col = st.columns([1.35, 0.65], gap="large")
    with hero_col:
        st.markdown(
            build_hero_panel(
                title="Centro de control operativo y analitico",
                copy=(
                    "Una portada unificada para operar los modulos del aplicativo con una interfaz sobria, "
                    "clara y lista para crecer. Desde aqui entras a TEC y a los espacios que luego completaremos."
                ),
                kicker="Suite Operativa",
                metrics=[
                    (str(len(MODULE_CATALOG)), "modulos visibles"),
                    ("1", "modulo ya operativo"),
                    (role_label, "perfil activo"),
                ],
            ),
            unsafe_allow_html=True,
        )
    with brand_col:
        logo_image = load_contractor_logo()
        if logo_image is not None:
            st.markdown('<section class="auth-card"><div class="auth-card-kicker">Empresa contratante</div><div class="auth-card-title">Identidad institucional</div><p class="auth-card-copy">La portada incorpora la marca del contratante como referencia visual permanente.</p></section>', unsafe_allow_html=True)
            st.image(logo_image, use_container_width=True)

    st.markdown('<div class="section-heading">Modulos disponibles</div>', unsafe_allow_html=True)
    st.markdown(
        f'<div class="section-copy">Sesion iniciada como {user_label}. Selecciona un modulo para continuar.</div>',
        unsafe_allow_html=True,
    )

    for start_index in range(0, len(MODULE_CATALOG), 3):
        row_modules = MODULE_CATALOG[start_index : start_index + 3]
        columns = st.columns(len(row_modules))
        for column, module in zip(columns, row_modules):
            with column:
                st.markdown(build_module_card(module), unsafe_allow_html=True)
                button_label = "Abrir modulo" if module["page"] == "TEC" else "Ver portada"
                if st.button(button_label, key=f'home_{module["page"]}', use_container_width=True):
                    navigate_to(module["page"])


def render_placeholder_module_page(page: str) -> None:
    module = next((item for item in MODULE_CATALOG if item["page"] == page), None)
    placeholder = MODULE_PLACEHOLDERS[page]
    st.markdown(
        build_hero_panel(
            title=module["title"] if module else page,
            copy=placeholder["summary"],
            kicker=module["eyebrow"] if module else "Modulo",
            metrics=[
                (module["status"] if module else "Proximamente", "estado"),
                ("Azul marino", "linea visual"),
                ("En definicion", "siguiente etapa"),
            ],
        ),
        unsafe_allow_html=True,
    )
    items_markup = "".join(f"<li>{item}</li>" for item in placeholder["items"])
    st.markdown(
        (
            '<section class="placeholder-card">'
            f'<div class="placeholder-title">{placeholder["headline"]}</div>'
            f'<div class="placeholder-copy">{placeholder["summary"]}</div>'
            f'<ul class="placeholder-list">{items_markup}</ul>'
            '</section>'
        ),
        unsafe_allow_html=True,
    )
    if st.button("Volver al inicio", key=f"back_{page}", use_container_width=False):
        navigate_to("Inicio")


def suggest_column(columns: list[str], patterns: list[str]) -> str | None:
    normalized = {col: re.sub(r"[^a-z0-9]+", "", col.lower()) for col in columns}
    for pattern in patterns:
        target = re.sub(r"[^a-z0-9]+", "", pattern.lower())
        for col, normalized_col in normalized.items():
            if target == normalized_col or target in normalized_col:
                return col
    return None


def patron_alfanumerico(plate: str) -> str:
    return "".join("A" if char.isalpha() else "9" if char.isdigit() else char for char in plate)


def generar_candidatas_confusion(plate: str) -> list[str]:
    candidates = set()
    for idx, char in enumerate(plate):
        for alternative in CONFUSION_LOOKUP.get(char, set()):
            candidates.add(plate[:idx] + alternative + plate[idx + 1 :])
    return sorted(candidates)


def normalizar_hora(valor):
    if pd.isna(valor):
        return pd.NaT
    if isinstance(valor, pd.Timedelta):
        return valor
    if isinstance(valor, timedelta):
        return pd.to_timedelta(valor)
    if isinstance(valor, pd.Timestamp):
        return (
            pd.to_timedelta(valor.hour, unit="h")
            + pd.to_timedelta(valor.minute, unit="m")
            + pd.to_timedelta(valor.second, unit="s")
        )
    if isinstance(valor, time):
        return (
            pd.to_timedelta(valor.hour, unit="h")
            + pd.to_timedelta(valor.minute, unit="m")
            + pd.to_timedelta(valor.second, unit="s")
        )
    if isinstance(valor, (int, float)) and not isinstance(valor, bool) and 0 <= float(valor) < 2:
        return pd.to_timedelta(float(valor), unit="D")
    texto = str(valor).strip()
    if not texto or texto.lower() in {"nat", "nan", "none"}:
        return pd.NaT
    como_timedelta = pd.to_timedelta(texto, errors="coerce")
    if pd.notna(como_timedelta):
        return como_timedelta
    como_datetime = pd.to_datetime(texto, errors="coerce")
    if pd.notna(como_datetime):
        return (
            pd.to_timedelta(como_datetime.hour, unit="h")
            + pd.to_timedelta(como_datetime.minute, unit="m")
            + pd.to_timedelta(como_datetime.second, unit="s")
        )
    return pd.NaT


def formatear_hora(valor):
    if pd.isna(valor):
        return pd.NA
    total_segundos = int(valor.total_seconds())
    horas = total_segundos // 3600
    minutos = (total_segundos % 3600) // 60
    segundos = total_segundos % 60
    return f"{horas:02}:{minutos:02}:{segundos:02}"


def segundos_a_timedelta(valor):
    if pd.isna(valor):
        return pd.NaT
    return pd.to_timedelta(float(valor), unit="s")


def timedelta_a_segundos(valor):
    if pd.isna(valor):
        return pd.NA
    return int(valor.total_seconds())


def timedelta_a_minutos(valor):
    if pd.isna(valor):
        return pd.NA
    return round(valor.total_seconds() / 60, 2)


def parse_manual_rules(rules_df: pd.DataFrame) -> tuple[dict[str, str], set[str], set[int]]:
    reglas = rules_df.copy()
    if reglas.empty:
        return {}, set(), set()
    reglas = reglas[reglas["activo"].fillna(False)].copy()
    reglas["tipo_regla"] = reglas["tipo_regla"].astype(str).str.strip()
    reglas["placa_origen"] = reglas["placa_origen"].fillna("").astype(str).str.upper().str.strip()
    reglas["placa_destino"] = reglas["placa_destino"].fillna("").astype(str).str.upper().str.strip()
    reglas["longitud_objetivo"] = pd.to_numeric(reglas["longitud_objetivo"], errors="coerce")

    correcciones = {
        row["placa_origen"]: row["placa_destino"]
        for _, row in reglas.iterrows()
        if row["tipo_regla"] == "corregir_placa" and row["placa_origen"] and row["placa_destino"]
    }
    eliminar_placa = {
        row["placa_origen"]
        for _, row in reglas.iterrows()
        if row["tipo_regla"] == "eliminar_placa" and row["placa_origen"]
    }
    eliminar_longitud = {
        int(row["longitud_objetivo"])
        for _, row in reglas.iterrows()
        if row["tipo_regla"] == "eliminar_por_longitud" and pd.notna(row["longitud_objetivo"])
    }
    return correcciones, eliminar_placa, eliminar_longitud


def load_input_dataframe(uploaded_file, sheet_name: str | None) -> pd.DataFrame:
    name = uploaded_file.name.lower()
    uploaded_file.seek(0)
    if name.endswith(".csv"):
        return pd.read_csv(uploaded_file)
    if name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls"):
        return pd.read_excel(uploaded_file, sheet_name=sheet_name or 0)
    raise ValueError("Formato no soportado. Usa CSV o Excel.")


def list_excel_sheets(uploaded_file) -> list[str]:
    name = uploaded_file.name.lower()
    if not (name.endswith(".xlsx") or name.endswith(".xlsm") or name.endswith(".xls")):
        return []
    uploaded_file.seek(0)
    excel = pd.ExcelFile(uploaded_file)
    return list(excel.sheet_names)


def build_standardized_df(df_raw: pd.DataFrame, mapping: dict[str, str | None]) -> pd.DataFrame:
    df_std = pd.DataFrame(index=df_raw.index)
    for target in EXPECTED_COLUMNS:
        source = mapping.get(target)
        if source:
            df_std[target] = df_raw[source]
        elif target == "VEHICULO":
            df_std[target] = pd.RangeIndex(1, len(df_raw) + 1)
        elif target in {"PEAJE", "CASETA", "SENTIDO"}:
            df_std[target] = f"SIN_{target}"
        elif target == "FECHA":
            df_std[target] = pd.NaT
        else:
            df_std[target] = pd.NA
    return df_std


def combine_action(placa_action: str, tiempo_action: str) -> str:
    partes = []
    if placa_action and placa_action != "sin_cambio":
        partes.append(f"placa:{placa_action}")
    if tiempo_action and tiempo_action != "sin_cambio":
        partes.append(f"tiempo:{tiempo_action}")
    return "sin_cambio" if not partes else " | ".join(partes)


def limpiar_placas_peaje(sub_df: pd.DataFrame) -> pd.DataFrame:
    result = sub_df.copy()
    original = result["PLACA"].fillna("").astype(str)
    original_upper = original.str.upper()
    trimmed_upper = original_upper.str.strip()
    normalized = trimmed_upper.str.replace(r"[^A-Z0-9]", "", regex=True)
    pattern = normalized.map(patron_alfanumerico)

    has_space_hyphen_or_symbol = original_upper.str.contains(r"[\s]|-|[^A-Z0-9\s-]", regex=True)
    invalid_length = ~normalized.str.len().isin(EXPECTED_PLATE_LENGTHS)
    uncommon_pattern = ~pattern.isin(COMMON_PATTERNS)
    placeholder_plate = normalized.isin(PLACEHOLDER_PLATES)
    suspicious_mask = has_space_hyphen_or_symbol | invalid_length | uncommon_pattern | placeholder_plate

    result["PLACA_NORMALIZADA"] = normalized
    result["PLACA_SUGERIDA"] = pd.Series(pd.NA, index=result.index, dtype="string")
    result["PLACA_FINAL"] = normalized
    result["PLACA_ESTADO"] = "ok"
    result["PLACA_MOTIVO"] = "sin observaciones"

    same_peaje_frequencies = normalized.value_counts()

    for idx in result.index:
        reasons = []
        plate = normalized.at[idx]
        plate_pattern = pattern.at[idx]
        current_state = "ok"
        final_plate = plate
        suggested_plate = pd.NA

        if has_space_hyphen_or_symbol.at[idx]:
            reasons.append("contiene espacios, guiones o simbolos")
        if invalid_length.at[idx]:
            reasons.append("longitud fuera de 6 caracteres")
        if uncommon_pattern.at[idx]:
            reasons.append(f"patron poco frecuente: {plate_pattern}")
        if placeholder_plate.at[idx]:
            reasons.append("placeholder o ejemplo")

        if suspicious_mask.at[idx]:
            current_state = "sospechosa"
            candidates = [
                candidate
                for candidate in generar_candidatas_confusion(plate)
                if candidate in same_peaje_frequencies.index
            ]
            if len(candidates) == 1:
                suggested_plate = candidates[0]
                reasons.append(f"sugerencia por confusion visual: {suggested_plate}")
                if same_peaje_frequencies.get(suggested_plate, 0) > same_peaje_frequencies.get(plate, 0):
                    current_state = "autocorregida_alta_confianza"
                    final_plate = suggested_plate
                    reasons.append(f"autocorreccion aplicada hacia {suggested_plate}")
        elif original.at[idx] != trimmed_upper.at[idx]:
            current_state = "normalizada"
            reasons = ["normalizacion segura de mayusculas o espacios externos"]

        result.at[idx, "PLACA_SUGERIDA"] = suggested_plate
        result.at[idx, "PLACA_FINAL"] = final_plate
        result.at[idx, "PLACA_ESTADO"] = current_state
        result.at[idx, "PLACA_MOTIVO"] = "; ".join(reasons) if reasons else "sin observaciones"

    return result


def distancia_levenshtein(a: str, b: str) -> int:
    if a == b:
        return 0
    prev = list(range(len(b) + 1))
    for i, char_a in enumerate(a, start=1):
        current = [i]
        for j, char_b in enumerate(b, start=1):
            current.append(
                min(
                    current[j - 1] + 1,
                    prev[j] + 1,
                    prev[j - 1] + (char_a != char_b),
                )
            )
        prev = current
    return prev[-1]


def tipo_operacion_lista(origen: str, candidata: str) -> str:
    if len(origen) == len(candidata):
        mismatches = [(a, b) for a, b in zip(origen, candidata) if a != b]
        if len(mismatches) == 1:
            old_char, new_char = mismatches[0]
            if new_char in CONFUSION_LOOKUP.get(old_char, set()) or old_char in CONFUSION_LOOKUP.get(new_char, set()):
                return "sustitucion_confusion"
            return "sustitucion_1_char"
    if len(origen) == len(candidata) + 1:
        for idx in range(len(origen)):
            if origen[:idx] + origen[idx + 1 :] == candidata:
                return "quitar_1_char"
    if len(origen) + 1 == len(candidata):
        for idx in range(len(candidata)):
            if candidata[:idx] + candidata[idx + 1 :] == origen:
                return "agregar_1_char"
    return "otro"


def buscar_coincidencias_lista(
    plate: str,
    placas_unicas_base: list[str],
    frecuencia_placas_base: dict[str, int],
) -> tuple[list[dict], list[dict]]:
    tipos_fuertes = {"sustitucion_confusion", "quitar_1_char", "agregar_1_char", "sustitucion_1_char"}
    prioridad_tipo = {
        "sustitucion_confusion": 0,
        "quitar_1_char": 1,
        "agregar_1_char": 1,
        "sustitucion_1_char": 2,
        "otro": 9,
    }
    coincidencias_fuertes = []
    coincidencias_debiles = []
    for candidata in placas_unicas_base:
        if candidata == plate or abs(len(candidata) - len(plate)) > 1:
            continue
        if distancia_levenshtein(plate, candidata) != 1:
            continue
        tipo_operacion = tipo_operacion_lista(plate, candidata)
        patron_candidata = patron_alfanumerico(candidata)
        item = {
            "candidata": candidata,
            "tipo_operacion": tipo_operacion,
            "frecuencia_candidata": int(frecuencia_placas_base.get(candidata, 0)),
            "patron_candidata": patron_candidata,
            "patron_comun": patron_candidata in COMMON_PATTERNS,
        }
        coincidencias_debiles.append(item)
        if item["patron_comun"] and tipo_operacion in tipos_fuertes:
            coincidencias_fuertes.append(item)
    sort_key = lambda item: (
        prioridad_tipo.get(item["tipo_operacion"], 9),
        -item["frecuencia_candidata"],
        item["candidata"],
    )
    coincidencias_fuertes = sorted(coincidencias_fuertes, key=sort_key)
    coincidencias_debiles = sorted(coincidencias_debiles, key=sort_key)
    return coincidencias_fuertes, coincidencias_debiles


def run_plate_cleaning(
    df_in: pd.DataFrame,
    config: dict,
    manual_rules_df: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    df_peajes = {
        peaje: limpiar_placas_peaje(sub_df)
        for peaje, sub_df in df_in.groupby("PEAJE", dropna=False)
    }
    df = pd.concat(df_peajes.values(), axis=0).sort_index()

    df["REVISION_GRUPO_BLOQUE"] = pd.NA
    df["REV_CANDIDATA_SIN_SUFIJO_X"] = pd.NA
    df["REV_LISTA_CANDIDATA_TOP"] = pd.NA
    df["PLACA_AJUSTE_MANUAL"] = pd.NA
    df["PLACA_ACCION_FINAL"] = "sin_cambio"
    df["PLACA_FINAL_DECIDIDA"] = df["PLACA_FINAL"]
    df["PLACA_EXCLUIR_ANALISIS"] = False

    df_revision_placas = df[df["PLACA_ESTADO"] == "sospechosa"].copy()
    if df_revision_placas.empty:
        correcciones_manual, eliminar_manual, longitudes_eliminar = parse_manual_rules(manual_rules_df)
        if config["aplicar_reglas_manuales"]:
            for idx, row in df.iterrows():
                normalized = str(row["PLACA_NORMALIZADA"])
                if normalized in correcciones_manual:
                    df.at[idx, "PLACA_ACCION_FINAL"] = "corregir_manual"
                    df.at[idx, "PLACA_FINAL_DECIDIDA"] = correcciones_manual[normalized]
                    df.at[idx, "PLACA_AJUSTE_MANUAL"] = f"Correccion manual: {normalized} -> {correcciones_manual[normalized]}"
                elif normalized in eliminar_manual:
                    df.at[idx, "PLACA_ACCION_FINAL"] = "excluir_analisis_placa"
                    df.at[idx, "PLACA_EXCLUIR_ANALISIS"] = True
                    df.at[idx, "PLACA_AJUSTE_MANUAL"] = f"Eliminacion manual: {normalized}"
                elif len(normalized) in longitudes_eliminar:
                    df.at[idx, "PLACA_ACCION_FINAL"] = "excluir_analisis_placa"
                    df.at[idx, "PLACA_EXCLUIR_ANALISIS"] = True
                    df.at[idx, "PLACA_AJUSTE_MANUAL"] = f"Eliminacion manual por longitud {len(normalized)}"
        return {
            "df": df,
            "df_trabajo": df[~df["PLACA_EXCLUIR_ANALISIS"]].copy(),
            "df_eliminados": df[df["PLACA_EXCLUIR_ANALISIS"]].copy(),
            "df_revision_placas": df_revision_placas,
            "df_bloques_decision": pd.DataFrame(),
            "resumen_acciones_placa": df["PLACA_ACCION_FINAL"].value_counts(dropna=False).rename_axis("PLACA_ACCION_FINAL").to_frame("filas"),
        }

    df_revision_placas["REV_PATRON"] = df_revision_placas["PLACA_NORMALIZADA"].map(patron_alfanumerico)
    df_revision_placas["REV_LONGITUD"] = df_revision_placas["PLACA_NORMALIZADA"].str.len()
    df_revision_placas["REV_FLAG_SIMBOLOS"] = df_revision_placas["PLACA"].astype(str).str.upper().str.contains(r"[\s]|-|[^A-Z0-9\s-]", regex=True)
    df_revision_placas["REV_FLAG_LONGITUD"] = ~df_revision_placas["REV_LONGITUD"].isin(EXPECTED_PLATE_LENGTHS)
    df_revision_placas["REV_FLAG_PLACEHOLDER"] = df_revision_placas["PLACA_NORMALIZADA"].isin(PLACEHOLDER_PLATES)
    df_revision_placas["REV_FLAG_PATRON"] = ~df_revision_placas["REV_PATRON"].isin(COMMON_PATTERNS)

    frecuencia_global_placa = df["PLACA_NORMALIZADA"].value_counts()
    frecuencia_mismo_peaje = df.groupby(["PEAJE", "PLACA_NORMALIZADA"]).size()
    frecuencia_patron_base = df["PLACA_NORMALIZADA"].map(patron_alfanumerico).value_counts()

    df_revision_placas["REV_COINCIDENCIAS_MISMO_PEAJE"] = df_revision_placas.apply(
        lambda row: int(frecuencia_mismo_peaje.get((row["PEAJE"], row["PLACA_NORMALIZADA"]), 0)) - 1,
        axis=1,
    )
    df_revision_placas["REV_COINCIDENCIAS_BASE"] = df_revision_placas.apply(
        lambda row: int(frecuencia_global_placa.get(row["PLACA_NORMALIZADA"], 0))
        - int(frecuencia_mismo_peaje.get((row["PEAJE"], row["PLACA_NORMALIZADA"]), 0)),
        axis=1,
    )
    df_revision_placas["REV_FRECUENCIA_PATRON_BASE"] = df_revision_placas["REV_PATRON"].map(frecuencia_patron_base)

    candidatas_confusion = []
    candidatas_confusion_mismo_peaje = []
    candidatas_confusion_base = []
    for _, row in df_revision_placas.iterrows():
        candidates = []
        for candidate in generar_candidatas_confusion(row["PLACA_NORMALIZADA"]):
            global_count = int(frecuencia_global_placa.get(candidate, 0))
            if global_count:
                same_peaje_count = int(frecuencia_mismo_peaje.get((row["PEAJE"], candidate), 0))
                candidates.append((candidate, same_peaje_count, global_count))
        candidates = sorted(candidates, key=lambda item: (-item[1], -item[2], item[0]))
        if len(candidates) == 1:
            candidatas_confusion.append(candidates[0][0])
            candidatas_confusion_mismo_peaje.append(candidates[0][1])
            candidatas_confusion_base.append(candidates[0][2])
        else:
            candidatas_confusion.append(pd.NA)
            candidatas_confusion_mismo_peaje.append(pd.NA)
            candidatas_confusion_base.append(pd.NA)

    df_revision_placas["REV_CANDIDATA_CONFUSION"] = candidatas_confusion
    df_revision_placas["REV_CANDIDATA_CONF_MISMO_PEAJE"] = candidatas_confusion_mismo_peaje
    df_revision_placas["REV_CANDIDATA_CONF_BASE"] = candidatas_confusion_base
    df_revision_placas["REV_CANDIDATA_SIN_SUFIJO_X"] = df_revision_placas["PLACA_NORMALIZADA"].where(
        df_revision_placas["PLACA_NORMALIZADA"].astype(str).str.endswith("X")
        & (df_revision_placas["REV_LONGITUD"] > max(EXPECTED_PLATE_LENGTHS)),
        pd.NA,
    ).str[:-1]
    df_revision_placas["REV_SUFIJO_X_MISMO_PEAJE"] = df_revision_placas.apply(
        lambda row: int(frecuencia_mismo_peaje.get((row["PEAJE"], row["REV_CANDIDATA_SIN_SUFIJO_X"]), 0))
        if pd.notna(row["REV_CANDIDATA_SIN_SUFIJO_X"])
        else 0,
        axis=1,
    )
    df_revision_placas["REV_SUFIJO_X_BASE"] = df_revision_placas["REV_CANDIDATA_SIN_SUFIJO_X"].map(
        lambda plate: int(frecuencia_global_placa.get(plate, 0)) if pd.notna(plate) else 0
    )

    def clasificar_revision(row: pd.Series) -> pd.Series:
        normalized = row["PLACA_NORMALIZADA"]
        pattern = row["REV_PATRON"]
        pattern_frequency = int(row["REV_FRECUENCIA_PATRON_BASE"])
        length = int(row["REV_LONGITUD"])
        exact_same_peaje = int(row["REV_COINCIDENCIAS_MISMO_PEAJE"])
        exact_base = int(row["REV_COINCIDENCIAS_BASE"])
        candidate = row["REV_CANDIDATA_CONFUSION"]
        candidate_same_peaje = row["REV_CANDIDATA_CONF_MISMO_PEAJE"]
        candidate_base = row["REV_CANDIDATA_CONF_BASE"]
        suffix_x_candidate = row["REV_CANDIDATA_SIN_SUFIJO_X"]
        suffix_x_same_peaje = int(row["REV_SUFIJO_X_MISMO_PEAJE"])
        suffix_x_base = int(row["REV_SUFIJO_X_BASE"])
        has_symbols = bool(row["REV_FLAG_SIMBOLOS"])
        has_length_issue = bool(row["REV_FLAG_LONGITUD"])
        is_placeholder = bool(row["REV_FLAG_PLACEHOLDER"])
        has_suffix_x = pd.notna(suffix_x_candidate)
        suffix_x_pattern = patron_alfanumerico(str(suffix_x_candidate)) if has_suffix_x else ""

        if is_placeholder:
            return pd.Series({"REVISION_GRUPO": "placeholder_o_ejemplo", "REVISION_SUBGRUPO": normalized})
        if has_suffix_x and suffix_x_same_peaje > 0:
            return pd.Series({"REVISION_GRUPO": "sufijo_x_final_con_respaldo_local", "REVISION_SUBGRUPO": "recorte_confirmado_en_mismo_peaje"})
        if has_suffix_x and suffix_x_base > 0:
            return pd.Series({"REVISION_GRUPO": "sufijo_x_final_con_respaldo_global", "REVISION_SUBGRUPO": "recorte_confirmado_en_base"})
        if has_suffix_x and suffix_x_pattern in {"AAA999", "A9A999"}:
            return pd.Series({"REVISION_GRUPO": "sufijo_x_final_patron_peru", "REVISION_SUBGRUPO": suffix_x_pattern})
        if has_suffix_x:
            return pd.Series({"REVISION_GRUPO": "sufijo_x_final_sin_respaldo", "REVISION_SUBGRUPO": "recorte_sin_confirmacion_en_base"})
        if has_symbols and not has_length_issue and exact_same_peaje > 0:
            return pd.Series({"REVISION_GRUPO": "ruido_tipografico_con_respaldo_local", "REVISION_SUBGRUPO": "normalizacion_confirmada_en_mismo_peaje"})
        if has_symbols and not has_length_issue and exact_base > 0:
            return pd.Series({"REVISION_GRUPO": "ruido_tipografico_con_respaldo_global", "REVISION_SUBGRUPO": "normalizacion_confirmada_en_base"})
        if pd.notna(candidate):
            return pd.Series({"REVISION_GRUPO": "posible_confusion_visual", "REVISION_SUBGRUPO": pattern})
        if has_symbols and not has_length_issue:
            return pd.Series({"REVISION_GRUPO": "ruido_tipografico_sin_respaldo", "REVISION_SUBGRUPO": pattern})
        if has_length_issue:
            return pd.Series({"REVISION_GRUPO": "longitud_atipica", "REVISION_SUBGRUPO": f"longitud_{length}_patron_{pattern}"})
        if pattern_frequency >= 2:
            return pd.Series({"REVISION_GRUPO": "patron_atipico_recurrente", "REVISION_SUBGRUPO": pattern})
        return pd.Series({"REVISION_GRUPO": "patron_atipico_aislado", "REVISION_SUBGRUPO": pattern})

    df_revision_placas = pd.concat([df_revision_placas, df_revision_placas.apply(clasificar_revision, axis=1)], axis=1)

    target_blocks = {"longitud_atipica", "patron_atipico_aislado", "patron_atipico_recurrente"}
    placas_unicas_base = sorted(df["PLACA_NORMALIZADA"].dropna().astype(str).unique())
    frecuencia_placas_base = df["PLACA_NORMALIZADA"].value_counts().to_dict()

    estado_lista = []
    bloque_decision = []
    candidata_top = []
    for _, row in df_revision_placas.iterrows():
        group = row["REVISION_GRUPO"]
        plate = row["PLACA_NORMALIZADA"]
        if group not in target_blocks:
            estado_lista.append("no_aplica")
            bloque_decision.append(group)
            candidata_top.append(pd.NA)
            continue
        fuertes, debiles = buscar_coincidencias_lista(plate, placas_unicas_base, frecuencia_placas_base)
        if len(fuertes) == 1:
            estado = "coincidencia_unica"
            top = fuertes[0]
        elif len(fuertes) > 1:
            estado = "coincidencia_multiple"
            top = fuertes[0]
        elif debiles:
            estado = "sin_coincidencia_fuerte"
            top = debiles[0]
        else:
            estado = "sin_coincidencia"
            top = None
        estado_lista.append(estado)
        bloque_decision.append(f"{group}__{estado}")
        candidata_top.append(top["candidata"] if top else pd.NA)

    df_revision_placas["REV_LISTA_COINCIDENCIA_ESTADO"] = estado_lista
    df_revision_placas["REVISION_BLOQUE_DECISION"] = bloque_decision
    df_revision_placas["REV_LISTA_CANDIDATA_TOP"] = candidata_top

    conteo_peajes_por_bloque = (
        df_revision_placas.groupby(["REVISION_BLOQUE_DECISION", "PEAJE"]).size().rename("filas").reset_index()
    )
    resumen_peajes_por_bloque = (
        conteo_peajes_por_bloque.groupby("REVISION_BLOQUE_DECISION")
        .apply(lambda sub: " | ".join(f"{row['PEAJE']}: {row['filas']}" for _, row in sub.iterrows()), include_groups=False)
        .rename("peajes")
    )
    df_bloques_decision = (
        df_revision_placas.groupby("REVISION_BLOQUE_DECISION")
        .agg(
            grupo_base=("REVISION_GRUPO", "first"),
            filas=("PLACA", "size"),
            ejemplos=("PLACA", lambda s: ", ".join(s.astype(str).head(6))),
            candidatas=("REV_LISTA_CANDIDATA_TOP", lambda s: ", ".join(s.dropna().astype(str).head(6))),
        )
        .join(resumen_peajes_por_bloque)
        .reset_index()
    )

    decisiones_por_bloque = {bloque: "mantener_observada" for bloque in df_bloques_decision["REVISION_BLOQUE_DECISION"]}
    for bloque in decisiones_por_bloque:
        base = bloque.split("__", 1)[0]
        estado = bloque.split("__", 1)[1] if "__" in bloque else "base"
        if base in {"ruido_tipografico_con_respaldo_local", "ruido_tipografico_con_respaldo_global"} and config["aplicar_ruido_con_respaldo"]:
            decisiones_por_bloque[bloque] = "corregir_a_normalizada"
        elif base == "ruido_tipografico_sin_respaldo" and config["aplicar_ruido_sin_respaldo"]:
            decisiones_por_bloque[bloque] = "corregir_a_normalizada"
        elif base == "posible_confusion_visual" and config["aplicar_confusion_visual"]:
            decisiones_por_bloque[bloque] = "corregir_a_sugerida"
        elif base in {
            "sufijo_x_final_patron_peru",
            "sufijo_x_final_con_respaldo_local",
            "sufijo_x_final_con_respaldo_global",
            "sufijo_x_final_sin_respaldo",
        } and config["aplicar_recorte_sufijo_x"]:
            decisiones_por_bloque[bloque] = "corregir_a_recorte_sufijo_x"
        elif base == "placeholder_o_ejemplo" and config["aplicar_exclusion_placeholders"]:
            decisiones_por_bloque[bloque] = "excluir_analisis_placa"
        elif estado == "coincidencia_unica" and config["aplicar_coincidencia_unica_lista"]:
            decisiones_por_bloque[bloque] = "corregir_a_coincidencia_lista"

    df_bloques_decision["ACCION_ELEGIDA"] = df_bloques_decision["REVISION_BLOQUE_DECISION"].map(decisiones_por_bloque)
    df_revision_placas["PLACA_ACCION_BLOQUE"] = df_revision_placas["REVISION_BLOQUE_DECISION"].map(decisiones_por_bloque)

    df.loc[df_revision_placas.index, "REVISION_GRUPO_BLOQUE"] = df_revision_placas["REVISION_BLOQUE_DECISION"]
    df.loc[df_revision_placas.index, "REV_CANDIDATA_SIN_SUFIJO_X"] = df_revision_placas["REV_CANDIDATA_SIN_SUFIJO_X"]
    df.loc[df_revision_placas.index, "REV_LISTA_CANDIDATA_TOP"] = df_revision_placas["REV_LISTA_CANDIDATA_TOP"]

    for idx, row in df.iterrows():
        action = df_revision_placas["PLACA_ACCION_BLOQUE"].get(idx, "sin_cambio")
        normalized = str(row["PLACA_NORMALIZADA"])
        final_plate = row["PLACA_FINAL"]
        exclude_plate = False

        if action == "corregir_a_normalizada":
            final_plate = normalized
        elif action == "corregir_a_sugerida" and pd.notna(row["PLACA_SUGERIDA"]):
            final_plate = row["PLACA_SUGERIDA"]
        elif action == "corregir_a_coincidencia_lista" and pd.notna(row["REV_LISTA_CANDIDATA_TOP"]):
            final_plate = row["REV_LISTA_CANDIDATA_TOP"]
        elif action == "corregir_a_recorte_sufijo_x" and pd.notna(row["REV_CANDIDATA_SIN_SUFIJO_X"]):
            final_plate = row["REV_CANDIDATA_SIN_SUFIJO_X"]
        elif action == "excluir_analisis_placa":
            exclude_plate = True
        else:
            action = "sin_cambio" if idx not in df_revision_placas.index else "mantener_observada"

        df.at[idx, "PLACA_ACCION_FINAL"] = action
        df.at[idx, "PLACA_FINAL_DECIDIDA"] = final_plate
        df.at[idx, "PLACA_EXCLUIR_ANALISIS"] = exclude_plate

    correcciones_manual, eliminar_manual, longitudes_eliminar = parse_manual_rules(manual_rules_df)
    if config["aplicar_reglas_manuales"]:
        for idx, row in df.iterrows():
            normalized = str(row["PLACA_NORMALIZADA"])
            if normalized in correcciones_manual:
                target = correcciones_manual[normalized]
                df.at[idx, "PLACA_ACCION_FINAL"] = "corregir_manual"
                df.at[idx, "PLACA_FINAL_DECIDIDA"] = target
                df.at[idx, "PLACA_EXCLUIR_ANALISIS"] = False
                df.at[idx, "PLACA_AJUSTE_MANUAL"] = f"Correccion manual: {normalized} -> {target}"
            elif normalized in eliminar_manual:
                df.at[idx, "PLACA_ACCION_FINAL"] = "excluir_analisis_placa"
                df.at[idx, "PLACA_EXCLUIR_ANALISIS"] = True
                df.at[idx, "PLACA_AJUSTE_MANUAL"] = f"Eliminacion manual: {normalized}"
            elif len(normalized) in longitudes_eliminar and not bool(df.at[idx, "PLACA_EXCLUIR_ANALISIS"]):
                df.at[idx, "PLACA_ACCION_FINAL"] = "excluir_analisis_placa"
                df.at[idx, "PLACA_EXCLUIR_ANALISIS"] = True
                df.at[idx, "PLACA_AJUSTE_MANUAL"] = f"Eliminacion manual por longitud {len(normalized)}"

    df_eliminados = df[df["PLACA_EXCLUIR_ANALISIS"]].copy()
    df_trabajo = df[~df["PLACA_EXCLUIR_ANALISIS"]].copy()
    resumen_acciones_placa = (
        df["PLACA_ACCION_FINAL"].value_counts(dropna=False).rename_axis("PLACA_ACCION_FINAL").to_frame("filas")
    )

    return {
        "df": df,
        "df_trabajo": df_trabajo,
        "df_eliminados": df_eliminados,
        "df_revision_placas": df_revision_placas,
        "df_bloques_decision": df_bloques_decision,
        "resumen_acciones_placa": resumen_acciones_placa,
    }


def describir_faltantes_trio(row: pd.Series) -> str:
    faltantes = []
    if pd.isna(row["T1"]):
        faltantes.append("T1")
    if pd.isna(row["T2"]):
        faltantes.append("T2")
    if pd.isna(row["T3"]):
        faltantes.append("T3")
    return ", ".join(faltantes) if faltantes else "ninguno"


def evaluar_bordes_caseta(grupo: pd.DataFrame) -> pd.DataFrame:
    grupo = grupo.sort_values(
        ["TIEMPO_REFERENCIA", "T1", "T2", "T3", "_ORDEN_FILA"],
        na_position="last",
    ).copy()
    grupo["POSICION_FLUJO"] = range(1, len(grupo) + 1)
    grupo["VENTANA_COMPLETA_IDENTIFICADA"] = grupo["TIEMPOS_COMPLETOS"].any()
    grupo["BORDE_CASETA_ACCION"] = "mantener_para_tiempos"
    grupo["BORDE_CASETA_MOTIVO"] = "dentro_ventana_o_registro_completo"
    grupo["BORDE_CASETA_ELIMINAR"] = False
    grupo["BORDE_ES_INICIO"] = False
    grupo["BORDE_ES_CIERRE"] = False

    mascara_completa = grupo["TIEMPOS_COMPLETOS"]
    if not mascara_completa.any():
        grupo["BORDE_CASETA_ACCION"] = "sin_referencia_completa"
        grupo["BORDE_CASETA_MOTIVO"] = (
            "El flujo no tiene ningun vehiculo con T1, T2 y T3 completos; "
            "no se fija apertura ni cierre de manera automatica."
        )
        return grupo

    apertura_pos = int(grupo.loc[mascara_completa, "POSICION_FLUJO"].min())
    cierre_pos = int(grupo.loc[mascara_completa, "POSICION_FLUJO"].max())

    eliminar_inicio = (grupo["POSICION_FLUJO"] < apertura_pos) & (~grupo["TIEMPOS_COMPLETOS"])
    eliminar_cierre = (grupo["POSICION_FLUJO"] > cierre_pos) & (~grupo["TIEMPOS_COMPLETOS"])

    grupo.loc[eliminar_inicio, "BORDE_CASETA_ACCION"] = "eliminar_inicio_caseta"
    grupo.loc[eliminar_inicio, "BORDE_CASETA_MOTIVO"] = (
        "Registro incompleto antes del primer vehiculo con T1, T2 y T3 completos del flujo."
    )
    grupo.loc[eliminar_inicio, "BORDE_CASETA_ELIMINAR"] = True
    grupo.loc[eliminar_inicio, "BORDE_ES_INICIO"] = True

    grupo.loc[eliminar_cierre, "BORDE_CASETA_ACCION"] = "eliminar_cierre_caseta"
    grupo.loc[eliminar_cierre, "BORDE_CASETA_MOTIVO"] = (
        "Registro incompleto despues del ultimo vehiculo con T1, T2 y T3 completos del flujo."
    )
    grupo.loc[eliminar_cierre, "BORDE_CASETA_ELIMINAR"] = True
    grupo.loc[eliminar_cierre, "BORDE_ES_CIERRE"] = True

    return grupo


def patron_faltantes_tiempo(row: pd.Series, columnas: list[str]) -> str:
    faltantes = [col for col in columnas if pd.isna(row[col])]
    return "completo" if not faltantes else "+".join(faltantes)


def columnas_imputadas_tiempo(row: pd.Series, columnas_origen: list[str], columnas_finales: list[str]) -> str:
    imputadas = []
    for origen, final in zip(columnas_origen, columnas_finales):
        if pd.isna(row[origen]) and pd.notna(row[final]):
            imputadas.append(origen)
    return "ninguna" if not imputadas else "+".join(imputadas)


def interpolar_tiempos_grupo(grupo: pd.DataFrame) -> pd.DataFrame:
    grupo = grupo.sort_values("POSICION_FLUJO").copy()
    posicion = grupo["POSICION_FLUJO"].astype(float)
    for alias in TIME_ALIASES:
        segundos = grupo[alias].dt.total_seconds()
        observados = segundos.notna()
        prev_segundos = segundos.where(observados).ffill()
        next_segundos = segundos.where(observados).bfill()
        prev_posicion = posicion.where(observados).ffill()
        next_posicion = posicion.where(observados).bfill()
        denominador = next_posicion - prev_posicion
        proporcion = (posicion - prev_posicion) / denominador
        candidata_segundos = prev_segundos + (next_segundos - prev_segundos) * proporcion
        mascara_candidata = (
            grupo[alias].isna()
            & prev_segundos.notna()
            & next_segundos.notna()
            & prev_posicion.notna()
            & next_posicion.notna()
            & (denominador > 0)
        )
        grupo[f"{alias}_INTERP"] = candidata_segundos.where(mascara_candidata).map(segundos_a_timedelta)
    return grupo


def run_time_cleaning(df_trabajo: pd.DataFrame, config: dict) -> dict[str, pd.DataFrame]:
    df_tiempos_base = df_trabajo.copy()
    df_tiempos_base["PLACA_FINAL"] = df_tiempos_base["PLACA_FINAL_DECIDIDA"]
    df_tiempos_base = df_tiempos_base[
        [
            "PEAJE",
            "CASETA",
            "SENTIDO",
            "FECHA",
            "VEHICULO",
            "PLACA_FINAL",
            "PLACA_ACCION_FINAL",
            "LLEGADA COLA",
            "LLEGADA CASETA",
            "SALIDA CASETA",
            "T. TEC",
            "T. CASETA",
            "_ORDEN_FILA",
        ]
    ].copy()
    df_tiempos_base["FECHA_DIA"] = pd.to_datetime(df_tiempos_base["FECHA"], errors="coerce").dt.normalize()
    for columna_origen, alias in TIME_COLUMNS.items():
        df_tiempos_base[alias] = df_tiempos_base[columna_origen].map(normalizar_hora)
        df_tiempos_base[f"{alias}_TXT"] = df_tiempos_base[alias].map(formatear_hora)
    df_tiempos_base["TIEMPOS_COMPLETOS"] = df_tiempos_base[["T1", "T2", "T3"]].notna().all(axis=1)
    df_tiempos_base["TIEMPO_REFERENCIA"] = (
        df_tiempos_base["T1"].combine_first(df_tiempos_base["T2"]).combine_first(df_tiempos_base["T3"])
    )

    if config["eliminar_bordes_caseta"]:
        grupos_evaluados = [
            evaluar_bordes_caseta(sub_df)
            for _, sub_df in df_tiempos_base.groupby(TIME_GROUP_KEYS, sort=True, dropna=False)
        ]
        df_tiempos_bordes = pd.concat(grupos_evaluados, axis=0).sort_index()
    else:
        df_tiempos_bordes = df_tiempos_base.copy()
        df_tiempos_bordes["POSICION_FLUJO"] = df_tiempos_bordes.groupby(TIME_GROUP_KEYS, dropna=False).cumcount() + 1
        df_tiempos_bordes["VENTANA_COMPLETA_IDENTIFICADA"] = df_tiempos_bordes.groupby(TIME_GROUP_KEYS, dropna=False)["TIEMPOS_COMPLETOS"].transform("any")
        df_tiempos_bordes["BORDE_CASETA_ACCION"] = "sin_cambio"
        df_tiempos_bordes["BORDE_CASETA_MOTIVO"] = "etapa_bordes_desactivada"
        df_tiempos_bordes["BORDE_CASETA_ELIMINAR"] = False
        df_tiempos_bordes["BORDE_ES_INICIO"] = False
        df_tiempos_bordes["BORDE_ES_CIERRE"] = False

    df_tiempos_bordes["PENDIENTE_TIEMPOS_INTERNO"] = (
        ~df_tiempos_bordes["TIEMPOS_COMPLETOS"] & ~df_tiempos_bordes["BORDE_CASETA_ELIMINAR"]
    )
    df_tiempos_eliminados_borde = df_tiempos_bordes[df_tiempos_bordes["BORDE_CASETA_ELIMINAR"]].copy()
    df_tiempos_trabajo = df_tiempos_bordes[~df_tiempos_bordes["BORDE_CASETA_ELIMINAR"]].copy()

    grupos_interpolados = [
        interpolar_tiempos_grupo(sub_df)
        for _, sub_df in df_tiempos_trabajo.groupby(TIME_GROUP_KEYS, sort=True, dropna=False)
    ]
    df_inter = pd.concat(grupos_interpolados, axis=0).sort_index()
    df_inter["PATRON_FALTANTES_ORIGINAL"] = df_inter.apply(lambda row: patron_faltantes_tiempo(row, TIME_ALIASES), axis=1)
    df_inter["ELEGIBLE_INTERPOLACION"] = (
        df_inter["PENDIENTE_TIEMPOS_INTERNO"]
        & df_inter["VENTANA_COMPLETA_IDENTIFICADA"]
        & df_inter["TIEMPO_REFERENCIA"].notna()
    )

    mascara_candidata_completa = pd.Series(True, index=df_inter.index)
    for alias in TIME_ALIASES:
        mascara_candidata_completa &= df_inter[alias].notna() | df_inter[f"{alias}_INTERP"].notna()
        df_inter[f"{alias}_PROPUESTA"] = df_inter[alias].combine_first(df_inter[f"{alias}_INTERP"])
    df_inter["INTERPOLACION_COMPLETA"] = mascara_candidata_completa
    mascara_propuesta_completa = df_inter[[f"{alias}_PROPUESTA" for alias in TIME_ALIASES]].notna().all(axis=1)
    df_inter["INTERPOLACION_ORDEN_VALIDO"] = (
        mascara_propuesta_completa
        & (df_inter["T1_PROPUESTA"] <= df_inter["T2_PROPUESTA"])
        & (df_inter["T2_PROPUESTA"] <= df_inter["T3_PROPUESTA"])
    )
    mascara_incompleta = ~df_inter["TIEMPOS_COMPLETOS"]
    mascara_imputar = (
        config["aplicar_interpolacion"]
        & mascara_incompleta
        & df_inter["ELEGIBLE_INTERPOLACION"]
        & df_inter["INTERPOLACION_COMPLETA"]
        & df_inter["INTERPOLACION_ORDEN_VALIDO"]
    )
    for alias in TIME_ALIASES:
        df_inter[f"{alias}_FINAL"] = df_inter[alias]
        df_inter.loc[mascara_imputar, f"{alias}_FINAL"] = df_inter.loc[mascara_imputar, f"{alias}_PROPUESTA"]
    df_inter["TIEMPO_ACCION_INTERPOLACION"] = "sin_cambio"
    df_inter["TIEMPO_MOTIVO_INTERPOLACION"] = "registro_ya_completo"
    df_inter.loc[mascara_incompleta, "TIEMPO_ACCION_INTERPOLACION"] = "mantener_pendiente"
    df_inter.loc[mascara_incompleta, "TIEMPO_MOTIVO_INTERPOLACION"] = "pendiente_por_evaluar"
    df_inter.loc[
        mascara_incompleta & ~df_inter["VENTANA_COMPLETA_IDENTIFICADA"],
        "TIEMPO_MOTIVO_INTERPOLACION",
    ] = "flujo_sin_referencia_completa"
    df_inter.loc[
        mascara_incompleta & df_inter["VENTANA_COMPLETA_IDENTIFICADA"] & df_inter["TIEMPO_REFERENCIA"].isna(),
        "TIEMPO_MOTIVO_INTERPOLACION",
    ] = "fila_sin_tiempo_referencia"
    df_inter.loc[
        mascara_incompleta & df_inter["ELEGIBLE_INTERPOLACION"] & ~df_inter["INTERPOLACION_COMPLETA"],
        "TIEMPO_MOTIVO_INTERPOLACION",
    ] = "faltan_anclas_para_interpolar"
    df_inter.loc[
        mascara_incompleta
        & df_inter["ELEGIBLE_INTERPOLACION"]
        & df_inter["INTERPOLACION_COMPLETA"]
        & ~df_inter["INTERPOLACION_ORDEN_VALIDO"],
        "TIEMPO_MOTIVO_INTERPOLACION",
    ] = "interpolacion_rompe_orden"
    df_inter.loc[mascara_imputar, "TIEMPO_ACCION_INTERPOLACION"] = "imputar_interpolacion"
    df_inter.loc[mascara_imputar, "TIEMPO_MOTIVO_INTERPOLACION"] = "interpolacion_interna_con_ancla_previa_y_posterior"

    base_medianas_locales = df_inter[df_inter[[f"{alias}_FINAL" for alias in TIME_ALIASES]].notna().all(axis=1)].copy()
    base_medianas_locales["T_COLA_FINAL"] = base_medianas_locales["T2_FINAL"] - base_medianas_locales["T1_FINAL"]
    base_medianas_locales["T_CASETA_FINAL"] = base_medianas_locales["T3_FINAL"] - base_medianas_locales["T2_FINAL"]
    base_medianas_locales = base_medianas_locales[
        base_medianas_locales["T_COLA_FINAL"].notna()
        & base_medianas_locales["T_CASETA_FINAL"].notna()
        & (base_medianas_locales["T_COLA_FINAL"] >= pd.Timedelta(0))
        & (base_medianas_locales["T_CASETA_FINAL"] >= pd.Timedelta(0))
    ].copy()
    medianas_locales_flujo = (
        base_medianas_locales.groupby(TIME_GROUP_KEYS, dropna=False)
        .agg(
            MEDIANA_COLA_FLUJO=("T_COLA_FINAL", "median"),
            MEDIANA_CASETA_FLUJO=("T_CASETA_FINAL", "median"),
        )
        .reset_index()
    )
    df_2da = df_inter.merge(medianas_locales_flujo, on=TIME_GROUP_KEYS, how="left")
    objetivo_segunda_pasada = df_2da["TIEMPO_MOTIVO_INTERPOLACION"] == "interpolacion_rompe_orden"

    def proponer_segunda_pasada(row: pd.Series) -> pd.Series:
        resultado = {
            "SEGUNDA_T1": row["T1_FINAL"],
            "SEGUNDA_T2": row["T2_FINAL"],
            "SEGUNDA_T3": row["T3_FINAL"],
            "SEGUNDA_PASADA_MOTIVO": pd.NA,
        }
        if row["TIEMPO_MOTIVO_INTERPOLACION"] != "interpolacion_rompe_orden":
            resultado["SEGUNDA_PASADA_MOTIVO"] = "no_aplica"
            return pd.Series(resultado)
        mediana_cola = row["MEDIANA_COLA_FLUJO"]
        mediana_caseta = row["MEDIANA_CASETA_FLUJO"]
        if pd.isna(mediana_cola) or pd.isna(mediana_caseta):
            resultado["SEGUNDA_PASADA_MOTIVO"] = "medianas_locales_insuficientes"
            return pd.Series(resultado)
        t1, t2, t3 = row["T1_FINAL"], row["T2_FINAL"], row["T3_FINAL"]
        for _ in range(3):
            cambio = False
            if pd.isna(t2):
                if pd.notna(t1):
                    t2 = t1 + mediana_cola
                    cambio = True
                elif pd.notna(t3):
                    t2 = t3 - mediana_caseta
                    cambio = True
            if pd.isna(t3) and pd.notna(t2):
                t3 = t2 + mediana_caseta
                cambio = True
            if pd.isna(t1) and pd.notna(t2):
                t1 = t2 - mediana_cola
                cambio = True
            if not cambio:
                break
        resultado["SEGUNDA_T1"] = t1
        resultado["SEGUNDA_T2"] = t2
        resultado["SEGUNDA_T3"] = t3
        if pd.isna(t1) or pd.isna(t2) or pd.isna(t3):
            resultado["SEGUNDA_PASADA_MOTIVO"] = "segunda_pasada_incompleta"
        elif not (t1 <= t2 <= t3):
            resultado["SEGUNDA_PASADA_MOTIVO"] = "segunda_pasada_rompe_orden"
        else:
            resultado["SEGUNDA_PASADA_MOTIVO"] = "segunda_pasada_mediana_local"
        return pd.Series(resultado)

    propuestas_segunda = df_2da.apply(proponer_segunda_pasada, axis=1)
    df_2da = pd.concat([df_2da, propuestas_segunda], axis=1)
    mascara_segunda_recuperada = (
        config["aplicar_mediana_local"]
        & objetivo_segunda_pasada
        & df_2da["SEGUNDA_PASADA_MOTIVO"].eq("segunda_pasada_mediana_local")
    )
    for alias in TIME_ALIASES:
        df_2da[f"{alias}_FINAL_2DA"] = df_2da[f"{alias}_FINAL"]
        df_2da.loc[mascara_segunda_recuperada, f"{alias}_FINAL_2DA"] = df_2da.loc[
            mascara_segunda_recuperada,
            f"SEGUNDA_{alias}",
        ]
    df_2da["TIEMPO_ACCION_FINAL"] = df_2da["TIEMPO_ACCION_INTERPOLACION"]
    df_2da["TIEMPO_MOTIVO_FINAL"] = df_2da["TIEMPO_MOTIVO_INTERPOLACION"]
    df_2da.loc[mascara_segunda_recuperada, "TIEMPO_ACCION_FINAL"] = "imputar_mediana_local"
    df_2da.loc[mascara_segunda_recuperada, "TIEMPO_MOTIVO_FINAL"] = "segunda_pasada_mediana_local"
    df_2da.loc[
        objetivo_segunda_pasada & ~mascara_segunda_recuperada,
        "TIEMPO_MOTIVO_FINAL",
    ] = df_2da.loc[objetivo_segunda_pasada & ~mascara_segunda_recuperada, "SEGUNDA_PASADA_MOTIVO"]

    base_donantes = df_2da[df_2da[[f"{alias}_FINAL_2DA" for alias in TIME_ALIASES]].notna().all(axis=1)].copy()
    base_donantes["COLA_DONANTE"] = base_donantes["T2_FINAL_2DA"] - base_donantes["T1_FINAL_2DA"]
    base_donantes["CASETA_DONANTE"] = base_donantes["T3_FINAL_2DA"] - base_donantes["T2_FINAL_2DA"]
    base_donantes = base_donantes[
        base_donantes["COLA_DONANTE"].notna()
        & base_donantes["CASETA_DONANTE"].notna()
        & (base_donantes["COLA_DONANTE"] >= pd.Timedelta(0))
        & (base_donantes["CASETA_DONANTE"] >= pd.Timedelta(0))
    ].copy()
    donantes_caseta = (
        base_donantes.groupby(["PEAJE", "CASETA", "SENTIDO"], dropna=False)
        .agg(
            MEDIANA_COLA_DONANTE_CASETA=("COLA_DONANTE", "median"),
            MEDIANA_CASETA_DONANTE_CASETA=("CASETA_DONANTE", "median"),
            REFERENCIAS_DONANTE_CASETA=("PLACA_FINAL", "size"),
        )
        .reset_index()
    )
    donantes_dia = (
        base_donantes.groupby(["PEAJE", "FECHA_DIA", "SENTIDO"], dropna=False)
        .agg(
            MEDIANA_COLA_DONANTE_DIA=("COLA_DONANTE", "median"),
            MEDIANA_CASETA_DONANTE_DIA=("CASETA_DONANTE", "median"),
            REFERENCIAS_DONANTE_DIA=("PLACA_FINAL", "size"),
        )
        .reset_index()
    )
    donantes_peaje_sentido = (
        base_donantes.groupby(["PEAJE", "SENTIDO"], dropna=False)
        .agg(
            MEDIANA_COLA_DONANTE_PEAJE_SENTIDO=("COLA_DONANTE", "median"),
            MEDIANA_CASETA_DONANTE_PEAJE_SENTIDO=("CASETA_DONANTE", "median"),
            REFERENCIAS_DONANTE_PEAJE_SENTIDO=("PLACA_FINAL", "size"),
        )
        .reset_index()
    )
    df_3ra = df_2da.merge(donantes_caseta, on=["PEAJE", "CASETA", "SENTIDO"], how="left")
    df_3ra = df_3ra.merge(donantes_dia, on=["PEAJE", "FECHA_DIA", "SENTIDO"], how="left")
    df_3ra = df_3ra.merge(donantes_peaje_sentido, on=["PEAJE", "SENTIDO"], how="left")

    def elegir_donante_cola(row: pd.Series) -> pd.Series:
        if pd.notna(row["MEDIANA_COLA_DONANTE_CASETA"]):
            return pd.Series({
                "COLA_DONANTE_ELEGIDA": row["MEDIANA_COLA_DONANTE_CASETA"],
                "CASETA_DONANTE_ELEGIDA": row["MEDIANA_CASETA_DONANTE_CASETA"],
                "FUENTE_DONANTE": "misma_caseta_sentido",
            })
        if pd.notna(row["MEDIANA_COLA_DONANTE_DIA"]):
            return pd.Series({
                "COLA_DONANTE_ELEGIDA": row["MEDIANA_COLA_DONANTE_DIA"],
                "CASETA_DONANTE_ELEGIDA": row["MEDIANA_CASETA_DONANTE_DIA"],
                "FUENTE_DONANTE": "mismo_dia_mismo_sentido",
            })
        if pd.notna(row["MEDIANA_COLA_DONANTE_PEAJE_SENTIDO"]):
            return pd.Series({
                "COLA_DONANTE_ELEGIDA": row["MEDIANA_COLA_DONANTE_PEAJE_SENTIDO"],
                "CASETA_DONANTE_ELEGIDA": row["MEDIANA_CASETA_DONANTE_PEAJE_SENTIDO"],
                "FUENTE_DONANTE": "mismo_peaje_mismo_sentido",
            })
        return pd.Series({"COLA_DONANTE_ELEGIDA": pd.NaT, "CASETA_DONANTE_ELEGIDA": pd.NaT, "FUENTE_DONANTE": pd.NA})

    df_3ra = pd.concat([df_3ra, df_3ra.apply(elegir_donante_cola, axis=1)], axis=1)
    objetivo_t1_donante = (
        config["aplicar_donantes"]
        & (df_3ra["TIEMPO_MOTIVO_FINAL"] == "flujo_sin_referencia_completa")
        & df_3ra["T1_FINAL_2DA"].isna()
        & df_3ra["T2_FINAL_2DA"].notna()
        & df_3ra["T3_FINAL_2DA"].notna()
    )
    objetivo_t2_t3_donante = (
        config["aplicar_donantes"]
        & (df_3ra["TIEMPO_MOTIVO_FINAL"] == "flujo_sin_referencia_completa")
        & df_3ra["T1_FINAL_2DA"].notna()
        & df_3ra["T2_FINAL_2DA"].isna()
        & df_3ra["T3_FINAL_2DA"].isna()
    )
    df_3ra["T1_FINAL_3RA"] = df_3ra["T1_FINAL_2DA"]
    df_3ra["T2_FINAL_3RA"] = df_3ra["T2_FINAL_2DA"]
    df_3ra["T3_FINAL_3RA"] = df_3ra["T3_FINAL_2DA"]
    mascara_t1_recuperada = objetivo_t1_donante & df_3ra["COLA_DONANTE_ELEGIDA"].notna()
    df_3ra.loc[mascara_t1_recuperada, "T1_FINAL_3RA"] = (
        df_3ra.loc[mascara_t1_recuperada, "T2_FINAL_2DA"] - df_3ra.loc[mascara_t1_recuperada, "COLA_DONANTE_ELEGIDA"]
    )

    def proponer_t2_t3_con_donantes(row: pd.Series) -> pd.Series:
        resultado = {"T2_DONANTE_PROPUESTO": pd.NaT, "T3_DONANTE_PROPUESTO": pd.NaT, "RECUPERABLE_T2_T3": False}
        if not (
            row["TIEMPO_MOTIVO_FINAL"] == "flujo_sin_referencia_completa"
            and pd.notna(row["T1_FINAL_2DA"])
            and pd.isna(row["T2_FINAL_2DA"])
            and pd.isna(row["T3_FINAL_2DA"])
        ):
            return pd.Series(resultado)
        if pd.notna(row["COLA_DONANTE_ELEGIDA"]) and pd.notna(row["CASETA_DONANTE_ELEGIDA"]):
            t2 = row["T1_FINAL_2DA"] + row["COLA_DONANTE_ELEGIDA"]
            t3 = t2 + row["CASETA_DONANTE_ELEGIDA"]
            if row["T1_FINAL_2DA"] <= t2 <= t3:
                resultado["T2_DONANTE_PROPUESTO"] = t2
                resultado["T3_DONANTE_PROPUESTO"] = t3
                resultado["RECUPERABLE_T2_T3"] = True
        return pd.Series(resultado)

    df_3ra = pd.concat([df_3ra, df_3ra.apply(proponer_t2_t3_con_donantes, axis=1)], axis=1)
    mascara_t2_t3_recuperada = objetivo_t2_t3_donante & df_3ra["RECUPERABLE_T2_T3"]
    df_3ra.loc[mascara_t2_t3_recuperada, "T2_FINAL_3RA"] = df_3ra.loc[mascara_t2_t3_recuperada, "T2_DONANTE_PROPUESTO"]
    df_3ra.loc[mascara_t2_t3_recuperada, "T3_FINAL_3RA"] = df_3ra.loc[mascara_t2_t3_recuperada, "T3_DONANTE_PROPUESTO"]
    df_3ra["TIEMPO_ACCION_GLOBAL"] = df_3ra["TIEMPO_ACCION_FINAL"]
    df_3ra["TIEMPO_MOTIVO_GLOBAL"] = df_3ra["TIEMPO_MOTIVO_FINAL"]
    df_3ra.loc[mascara_t1_recuperada, "TIEMPO_ACCION_GLOBAL"] = "imputar_donante_cola"
    df_3ra.loc[mascara_t1_recuperada, "TIEMPO_MOTIVO_GLOBAL"] = "t1_recuperado_con_mediana_cola_donante"
    df_3ra.loc[mascara_t2_t3_recuperada, "TIEMPO_ACCION_GLOBAL"] = "imputar_donante_t2_t3"
    df_3ra.loc[mascara_t2_t3_recuperada, "TIEMPO_MOTIVO_GLOBAL"] = "t2_t3_recuperados_con_donantes"

    df_final = df_3ra.copy()
    df_final["T1_FINAL_4TA"] = df_final["T1_FINAL_3RA"]
    df_final["T2_FINAL_4TA"] = df_final["T2_FINAL_3RA"]
    df_final["T3_FINAL_4TA"] = df_final["T3_FINAL_3RA"]
    objetivo_ajuste_final = (
        config["aplicar_swap_final_t2_t3"]
        & (df_final["TIEMPO_MOTIVO_GLOBAL"] == "segunda_pasada_rompe_orden")
        & df_final["T1_FINAL_3RA"].isna()
        & df_final["T2_FINAL_3RA"].notna()
        & df_final["T3_FINAL_3RA"].notna()
        & (df_final["T2_FINAL_3RA"] > df_final["T3_FINAL_3RA"])
    )
    df_final.loc[objetivo_ajuste_final, "T2_FINAL_4TA"] = df_final.loc[objetivo_ajuste_final, "T3_FINAL_3RA"]
    df_final.loc[objetivo_ajuste_final, "T3_FINAL_4TA"] = df_final.loc[objetivo_ajuste_final, "T2_FINAL_3RA"]
    mascara_con_donante_final = objetivo_ajuste_final & df_final["COLA_DONANTE_ELEGIDA"].notna()
    df_final.loc[mascara_con_donante_final, "T1_FINAL_4TA"] = (
        df_final.loc[mascara_con_donante_final, "T2_FINAL_4TA"] - df_final.loc[mascara_con_donante_final, "COLA_DONANTE_ELEGIDA"]
    )
    mascara_recuperada_final = (
        objetivo_ajuste_final
        & df_final["T1_FINAL_4TA"].notna()
        & df_final["T2_FINAL_4TA"].notna()
        & df_final["T3_FINAL_4TA"].notna()
        & (df_final["T1_FINAL_4TA"] <= df_final["T2_FINAL_4TA"])
        & (df_final["T2_FINAL_4TA"] <= df_final["T3_FINAL_4TA"])
    )
    df_final["TIEMPO_ACCION_CIERRE"] = df_final["TIEMPO_ACCION_GLOBAL"]
    df_final["TIEMPO_MOTIVO_CIERRE"] = df_final["TIEMPO_MOTIVO_GLOBAL"]
    df_final.loc[mascara_recuperada_final, "TIEMPO_ACCION_CIERRE"] = "swap_t2_t3_y_reconstruir_t1"
    df_final.loc[mascara_recuperada_final, "TIEMPO_MOTIVO_CIERRE"] = "t2_t3_intercambiados_y_t1_reconstruido_con_cola_donante"

    df_final["TIEMPOS_COMPLETOS_CIERRE"] = df_final[["T1_FINAL_4TA", "T2_FINAL_4TA", "T3_FINAL_4TA"]].notna().all(axis=1)
    df_final["LLEGADA_COLA_FINAL"] = df_final["T1_FINAL_4TA"].map(formatear_hora)
    df_final["LLEGADA_CASETA_FINAL"] = df_final["T2_FINAL_4TA"].map(formatear_hora)
    df_final["SALIDA_CASETA_FINAL"] = df_final["T3_FINAL_4TA"].map(formatear_hora)
    df_final["T_COLA_FINAL"] = df_final["T2_FINAL_4TA"] - df_final["T1_FINAL_4TA"]
    df_final["T_CASETA_FINAL"] = df_final["T3_FINAL_4TA"] - df_final["T2_FINAL_4TA"]
    df_final["T_TEC_FINAL"] = df_final["T3_FINAL_4TA"] - df_final["T1_FINAL_4TA"]
    df_final["T_COLA_FINAL_TXT"] = df_final["T_COLA_FINAL"].map(formatear_hora)
    df_final["T_CASETA_FINAL_TXT"] = df_final["T_CASETA_FINAL"].map(formatear_hora)
    df_final["T_TEC_FINAL_TXT"] = df_final["T_TEC_FINAL"].map(formatear_hora)

    df_pendientes = df_final[~df_final["TIEMPOS_COMPLETOS_CIERRE"]].copy()
    resumen_tiempos = (
        df_final["TIEMPO_ACCION_CIERRE"].value_counts(dropna=False).rename_axis("TIEMPO_ACCION_CIERRE").to_frame("filas")
    )
    return {
        "df_tiempos_base": df_tiempos_base,
        "df_tiempos_bordes": df_tiempos_bordes,
        "df_tiempos_eliminados_borde": df_tiempos_eliminados_borde,
        "df_tiempos_final": df_final,
        "df_tiempos_pendientes": df_pendientes,
        "resumen_tiempos": resumen_tiempos,
    }


def build_export_tables(
    df_original: pd.DataFrame,
    plate_result: dict[str, pd.DataFrame],
    time_result: dict[str, pd.DataFrame],
    config: dict,
    manual_rules_df: pd.DataFrame,
) -> dict[str, pd.DataFrame]:
    df = plate_result["df"]
    df_trabajo = plate_result["df_trabajo"]
    df_eliminados_placa = plate_result["df_eliminados"]
    df_tiempos_final = time_result["df_tiempos_final"]
    df_eliminados_borde = time_result["df_tiempos_eliminados_borde"]
    df_tiempos_pendientes = time_result["df_tiempos_pendientes"]

    placa_meta_export = df_trabajo[
        ["_ORDEN_FILA", "PLACA", "PLACA_FINAL_DECIDIDA", "PLACA_ACCION_FINAL", "PLACA_AJUSTE_MANUAL"]
    ].copy()
    df_export_base = df_tiempos_final.merge(
        placa_meta_export,
        on="_ORDEN_FILA",
        how="left",
        validate="one_to_one",
        suffixes=("", "_PLACA"),
    )
    df_export_base["PLACA_FINAL"] = df_export_base["PLACA_FINAL_DECIDIDA"]
    df_export_base["TIEMPO_ACCION_FINAL"] = df_export_base["TIEMPO_ACCION_CIERRE"]
    df_export_base["TIEMPO_MOTIVO_FINAL"] = df_export_base["TIEMPO_MOTIVO_CIERRE"]
    df_export_base["ACCION_REALIZADA"] = df_export_base.apply(
        lambda row: combine_action(row["PLACA_ACCION_FINAL"], row["TIEMPO_ACCION_FINAL"]),
        axis=1,
    )
    df_export_base["T_COLA_FINAL_SEGUNDOS"] = df_export_base["T_COLA_FINAL"].map(timedelta_a_segundos)
    df_export_base["T_COLA_FINAL_MINUTOS"] = df_export_base["T_COLA_FINAL"].map(timedelta_a_minutos)
    df_export_base["T_CASETA_FINAL_SEGUNDOS"] = df_export_base["T_CASETA_FINAL"].map(timedelta_a_segundos)
    df_export_base["T_CASETA_FINAL_MINUTOS"] = df_export_base["T_CASETA_FINAL"].map(timedelta_a_minutos)
    df_export_base["T_TEC_FINAL_SEGUNDOS"] = df_export_base["T_TEC_FINAL"].map(timedelta_a_segundos)
    df_export_base["T_TEC_FINAL_MINUTOS"] = df_export_base["T_TEC_FINAL"].map(timedelta_a_minutos)
    df_export_base["REGISTRO_AJUSTADO"] = (
        df_export_base["PLACA_ACCION_FINAL"].isin(ACCIONES_PLACA_AJUSTE)
        | df_export_base["TIEMPO_ACCION_FINAL"].ne("sin_cambio")
    )
    df_export_base["PLACA_PENDIENTE_REVISION"] = df_export_base["PLACA_ACCION_FINAL"].eq("mantener_observada")

    base_limpia = df_export_base[df_export_base["TIEMPOS_COMPLETOS_CIERRE"]].copy()
    base_limpia = base_limpia[
        [
            "PEAJE",
            "CASETA",
            "SENTIDO",
            "FECHA",
            "VEHICULO",
            "PLACA_FINAL",
            "LLEGADA_COLA_FINAL",
            "LLEGADA_CASETA_FINAL",
            "SALIDA_CASETA_FINAL",
            "T_COLA_FINAL_TXT",
            "T_COLA_FINAL_SEGUNDOS",
            "T_COLA_FINAL_MINUTOS",
            "T_CASETA_FINAL_TXT",
            "T_CASETA_FINAL_SEGUNDOS",
            "T_CASETA_FINAL_MINUTOS",
            "T_TEC_FINAL_TXT",
            "T_TEC_FINAL_SEGUNDOS",
            "T_TEC_FINAL_MINUTOS",
            "ACCION_REALIZADA",
        ]
    ].rename(
        columns={
            "PLACA_FINAL": "PLACA",
            "LLEGADA_COLA_FINAL": "LLEGADA COLA",
            "LLEGADA_CASETA_FINAL": "LLEGADA CASETA",
            "SALIDA_CASETA_FINAL": "SALIDA CASETA",
            "T_COLA_FINAL_TXT": "T. COLA",
            "T_COLA_FINAL_SEGUNDOS": "T. COLA_SEGUNDOS",
            "T_COLA_FINAL_MINUTOS": "T. COLA_MINUTOS",
            "T_CASETA_FINAL_TXT": "T. CASETA",
            "T_CASETA_FINAL_SEGUNDOS": "T. CASETA_SEGUNDOS",
            "T_CASETA_FINAL_MINUTOS": "T. CASETA_MINUTOS",
            "T_TEC_FINAL_TXT": "T. TEC",
            "T_TEC_FINAL_SEGUNDOS": "T. TEC_SEGUNDOS",
            "T_TEC_FINAL_MINUTOS": "T. TEC_MINUTOS",
        }
    )

    eliminados_placa = pd.DataFrame()
    if not df_eliminados_placa.empty:
        eliminados_placa = df_eliminados_placa[
            EXPECTED_COLUMNS + ["PLACA_FINAL_DECIDIDA", "PLACA_ACCION_FINAL", "PLACA_AJUSTE_MANUAL", "REVISION_GRUPO_BLOQUE", "PLACA_MOTIVO"]
        ].copy()
        eliminados_placa["PLACA_FINAL"] = eliminados_placa["PLACA_FINAL_DECIDIDA"]
        eliminados_placa["ETAPA_ELIMINACION"] = "placa"
        eliminados_placa["ACCION_REALIZADA"] = "eliminado:" + eliminados_placa["PLACA_ACCION_FINAL"].astype(str)
        eliminados_placa["MOTIVO_ELIMINACION"] = (
            eliminados_placa["PLACA_AJUSTE_MANUAL"]
            .combine_first(eliminados_placa["REVISION_GRUPO_BLOQUE"])
            .combine_first(eliminados_placa["PLACA_MOTIVO"])
        )

    eliminados_tiempo = pd.DataFrame()
    if not df_eliminados_borde.empty:
        eliminados_tiempo = df_eliminados_borde[
            [
                "PEAJE",
                "CASETA",
                "SENTIDO",
                "FECHA",
                "VEHICULO",
                "LLEGADA COLA",
                "LLEGADA CASETA",
                "SALIDA CASETA",
                "T. TEC",
                "T. CASETA",
                "_ORDEN_FILA",
                "BORDE_CASETA_ACCION",
                "BORDE_CASETA_MOTIVO",
            ]
        ].copy()
        eliminados_tiempo = eliminados_tiempo.merge(
            placa_meta_export[["_ORDEN_FILA", "PLACA", "PLACA_FINAL_DECIDIDA"]],
            on="_ORDEN_FILA",
            how="left",
            validate="many_to_one",
        )
        eliminados_tiempo["PLACA_FINAL"] = eliminados_tiempo["PLACA_FINAL_DECIDIDA"]
        eliminados_tiempo["ETAPA_ELIMINACION"] = "tiempo_borde_caseta"
        eliminados_tiempo["ACCION_REALIZADA"] = "eliminado:" + eliminados_tiempo["BORDE_CASETA_ACCION"].astype(str)
        eliminados_tiempo["MOTIVO_ELIMINACION"] = eliminados_tiempo["BORDE_CASETA_MOTIVO"]

    columnas_eliminados = [
        "PEAJE",
        "CASETA",
        "SENTIDO",
        "FECHA",
        "VEHICULO",
        "PLACA",
        "PLACA_FINAL",
        "ETAPA_ELIMINACION",
        "ACCION_REALIZADA",
        "MOTIVO_ELIMINACION",
        "LLEGADA COLA",
        "LLEGADA CASETA",
        "SALIDA CASETA",
        "T. TEC",
        "T. CASETA",
    ]
    casos_eliminados = pd.concat(
        [
            eliminados_placa.reindex(columns=columnas_eliminados),
            eliminados_tiempo.reindex(columns=columnas_eliminados),
        ],
        ignore_index=True,
    )

    casos_pendientes = pd.DataFrame()
    if not df_tiempos_pendientes.empty:
        casos_pendientes = df_tiempos_pendientes[
            [
                "PEAJE",
                "CASETA",
                "SENTIDO",
                "FECHA",
                "VEHICULO",
                "PLACA_FINAL",
                "LLEGADA COLA",
                "LLEGADA CASETA",
                "SALIDA CASETA",
                "TIEMPO_ACCION_CIERRE",
                "TIEMPO_MOTIVO_CIERRE",
            ]
        ].copy()
        casos_pendientes = casos_pendientes.rename(
            columns={
                "PLACA_FINAL": "PLACA",
                "TIEMPO_ACCION_CIERRE": "ACCION_REALIZADA",
                "TIEMPO_MOTIVO_CIERRE": "MOTIVO",
            }
        )

    resumen_rows = []
    resumen_rows.extend(
        [
            {"seccion": "general", "indicador": "filas_entrada", "valor": len(df_original)},
            {"seccion": "general", "indicador": "filas_base_limpia", "valor": len(base_limpia)},
            {"seccion": "general", "indicador": "filas_eliminadas", "valor": len(casos_eliminados)},
            {"seccion": "general", "indicador": "filas_pendientes", "valor": len(casos_pendientes)},
        ]
    )
    for idx, row in plate_result["resumen_acciones_placa"].reset_index().iterrows():
        resumen_rows.append({"seccion": "placa", "indicador": row.iloc[0], "valor": row.iloc[1]})
    for idx, row in time_result["resumen_tiempos"].reset_index().iterrows():
        resumen_rows.append({"seccion": "tiempo", "indicador": row.iloc[0], "valor": row.iloc[1]})
    for clave, valor in config.items():
        resumen_rows.append({"seccion": "config", "indicador": clave, "valor": valor})
    reporte_resumen = pd.DataFrame(resumen_rows)

    config_df = pd.DataFrame(
        [{"clave": clave, "valor": valor} for clave, valor in config.items()]
        + [{"clave": "manual_rules_rows", "valor": len(manual_rules_df)}]
    )

    return {
        "base_limpia": base_limpia,
        "casos_eliminados": casos_eliminados,
        "casos_pendientes": casos_pendientes,
        "reporte_resumen": reporte_resumen,
        "revision_placas": plate_result["df_revision_placas"],
        "bloques_decision": plate_result["df_bloques_decision"],
        "config_usada": config_df,
        "export_base_detalle": df_export_base,
    }


def build_exact_export_package(
    df_original: pd.DataFrame,
    export_tables: dict[str, pd.DataFrame],
) -> dict[str, pd.DataFrame]:
    base_limpia = export_tables["base_limpia"].copy()
    casos_eliminados = export_tables["casos_eliminados"].copy()
    df_export_base = export_tables["export_base_detalle"].copy()

    resumen_general = pd.Series(
        {
            "filas_base_original": len(df_original),
            "filas_base_limpia": len(base_limpia),
            "filas_ajustadas": int(df_export_base["REGISTRO_AJUSTADO"].sum()),
            "filas_con_placa_pendiente_revision": int(df_export_base["PLACA_PENDIENTE_REVISION"].sum()),
            "filas_eliminadas_placa": int(casos_eliminados["ETAPA_ELIMINACION"].eq("placa").sum()),
            "filas_eliminadas_tiempo": int(casos_eliminados["ETAPA_ELIMINACION"].eq("tiempo_borde_caseta").sum()),
            "filas_eliminadas_total": len(casos_eliminados),
        }
    ).rename_axis("indicador").reset_index(name="valor")

    resumen_accion_realizada = (
        base_limpia["ACCION_REALIZADA"]
        .value_counts(dropna=False)
        .rename_axis("ACCION_REALIZADA")
        .reset_index(name="filas")
    )

    resumen_acciones_placa = (
        df_export_base["PLACA_ACCION_FINAL"]
        .value_counts(dropna=False)
        .rename_axis("PLACA_ACCION_FINAL")
        .reset_index(name="filas")
    )

    resumen_acciones_tiempo = (
        df_export_base["TIEMPO_ACCION_FINAL"]
        .value_counts(dropna=False)
        .rename_axis("TIEMPO_ACCION_FINAL")
        .reset_index(name="filas")
    )

    resumen_eliminados = (
        casos_eliminados.groupby(["ETAPA_ELIMINACION", "ACCION_REALIZADA"], dropna=False)
        .size()
        .rename("filas")
        .reset_index()
        .sort_values(["ETAPA_ELIMINACION", "filas"], ascending=[True, False])
    )

    return {
        "base_limpia": base_limpia,
        "casos_eliminados": casos_eliminados,
        "resumen_general": resumen_general,
        "resumen_accion_realizada": resumen_accion_realizada,
        "resumen_acciones_placa": resumen_acciones_placa,
        "resumen_acciones_tiempo": resumen_acciones_tiempo,
        "resumen_eliminados": resumen_eliminados,
    }


def derive_output_filenames(uploaded_name: str) -> dict[str, str]:
    stem = Path(uploaded_name).stem.strip() or "resultado"
    if stem.lower().startswith("data "):
        label = stem[5:].strip() or stem
        return {
            "report_label": label,
            "clean_excel": f"{stem}_limpio.xlsx",
            "report_excel": f"Resultados {label}.xlsx",
            "report_docx": f"Tablas {label} para informe.docx",
            "extra_excel": f"Resultados {label} complementarios.xlsx",
        }
    return {
        "report_label": stem,
        "clean_excel": f"{stem}_limpio.xlsx",
        "report_excel": f"{stem}_resultados.xlsx",
        "report_docx": f"{stem}_tablas_informe.docx",
        "extra_excel": f"{stem}_resultados_complementarios.xlsx",
    }


def calcular_cola_espera_real(grupo: pd.DataFrame) -> pd.DataFrame:
    grupo = grupo.sort_values(
        [
            "LLEGADA_COLA_FINAL_TD",
            "LLEGADA_CASETA_FINAL_TD",
            "SALIDA_CASETA_FINAL_TD",
            "_ORDEN_FILA",
        ],
        na_position="last",
    ).copy()

    activos = []
    colas = []

    for row in grupo.itertuples(index=False):
        t1 = getattr(row, "LLEGADA_COLA_FINAL_TD")
        activos = [veh for veh in activos if veh["t3"] > t1]
        cola_espera = sum(1 for veh in activos if veh["t2"] > t1)
        colas.append(cola_espera)
        activos.append(
            {
                "t2": getattr(row, "LLEGADA_CASETA_FINAL_TD"),
                "t3": getattr(row, "SALIDA_CASETA_FINAL_TD"),
            }
        )

    grupo["COLA_ESPERA_USUARIOS"] = colas
    return grupo


def format_zero_blank(value):
    if pd.isna(value):
        return ""
    if isinstance(value, str):
        return value
    try:
        return "" if float(value) == 0 else value
    except (TypeError, ValueError):
        return value


def formatear_tabla_frecuencia(tabla: pd.DataFrame) -> pd.DataFrame:
    tabla = tabla.copy()
    columnas_valor = tabla.columns[1:]
    for columna in columnas_valor:
        tabla[columna] = tabla[columna].map(format_zero_blank)
    return tabla


def agregar_tabla_docx(doc: Document, titulo: str, tabla: pd.DataFrame) -> None:
    doc.add_paragraph(titulo, style="Heading 2")
    docx_table = doc.add_table(rows=1, cols=len(tabla.columns))
    docx_table.style = "Table Grid"
    docx_table.alignment = WD_TABLE_ALIGNMENT.CENTER

    for idx, columna in enumerate(tabla.columns):
        docx_table.rows[0].cells[idx].text = str(columna)

    for fila in tabla.itertuples(index=False):
        celdas = docx_table.add_row().cells
        for idx, valor in enumerate(fila):
            celdas[idx].text = "" if pd.isna(valor) else str(valor)


def q85(serie: pd.Series) -> float:
    return float(serie.quantile(0.85))


def q95(serie: pd.Series) -> float:
    return float(serie.quantile(0.95))


def resumir_metricas(df_in: pd.DataFrame, group_cols: list[str], value_col: str, label: str) -> pd.DataFrame:
    tabla = (
        df_in.groupby(group_cols, dropna=False)[value_col]
        .agg(
            casos="size",
            promedio="mean",
            mediana="median",
            p85=q85,
            p95=q95,
            maximo="max",
        )
        .reset_index()
    )
    tabla.insert(len(group_cols), "indicador", label)
    tabla[["promedio", "mediana", "p85", "p95", "maximo"]] = tabla[
        ["promedio", "mediana", "p85", "p95", "maximo"]
    ].round(2)
    return tabla


def build_resultados_dataframe(df_export_base: pd.DataFrame) -> pd.DataFrame:
    df_resultados_base = df_export_base[df_export_base["TIEMPOS_COMPLETOS_CIERRE"]].copy()
    if df_resultados_base.empty:
        df_resultados_base["COLA_ESPERA_USUARIOS"] = pd.Series(dtype="int64")
        return df_resultados_base

    for columna in ["LLEGADA_COLA_FINAL", "LLEGADA_CASETA_FINAL", "SALIDA_CASETA_FINAL"]:
        df_resultados_base[f"{columna}_TD"] = pd.to_timedelta(df_resultados_base[columna].astype(str), errors="coerce")

    grupos_resultados = [
        calcular_cola_espera_real(sub_df)
        for _, sub_df in df_resultados_base.groupby(["PEAJE", "CASETA", "SENTIDO", "FECHA"], dropna=False)
    ]
    if not grupos_resultados:
        df_resultados_base["COLA_ESPERA_USUARIOS"] = pd.Series(dtype="int64")
        return df_resultados_base

    return pd.concat(grupos_resultados, axis=0).sort_values(
        ["PEAJE", "SENTIDO", "CASETA", "FECHA", "LLEGADA_COLA_FINAL_TD", "_ORDEN_FILA"]
    ).reset_index(drop=True)


def build_informe_package(df_export_base: pd.DataFrame) -> dict[str, object]:
    df_resultados = build_resultados_dataframe(df_export_base)

    tabla_tec_caseta_informe = pd.DataFrame(
        columns=["Peaje", "Caseta Controlada", "Sentido de Circulacion", "Tiempo de Espera en Cola - TEC", "Unidades de tiempo"]
    )
    tabla_tec_peaje_informe = pd.DataFrame(
        columns=["Peaje", "Sentido de Circulacion", "Tiempo de Espera en Cola - TEC", "Unidades de tiempo"]
    )
    tabla_cola_maxima_informe = pd.DataFrame(
        columns=["Peaje", "Caseta Controlada", "Sentido de Circulacion", "Cola maxima real", "Vehiculos evaluados"]
    )
    tabla_frecuencia_paraiso_informe = pd.DataFrame(columns=["Cantidad de usuarios en la cola", "Total"])
    tabla_frecuencia_variante_informe = pd.DataFrame(columns=["Cantidad de usuarios en la cola", "Total"])
    resumen_narrativo = pd.DataFrame({"Texto sugerido para informe": ["No hay datos suficientes para generar resultados de informe."]})

    if not df_resultados.empty:
        tabla_tec_caseta = (
            df_resultados.groupby(["PEAJE", "CASETA", "SENTIDO"], dropna=False)
            .agg(
                TEC_MINUTOS=("T_TEC_FINAL_MINUTOS", "mean"),
                VEHICULOS=("PLACA_FINAL", "size"),
            )
            .reset_index()
            .sort_values(["PEAJE", "SENTIDO", "CASETA"])
        )
        tabla_tec_caseta["TEC_MINUTOS"] = tabla_tec_caseta["TEC_MINUTOS"].round(2)
        tabla_tec_caseta_informe = tabla_tec_caseta.rename(
            columns={
                "PEAJE": "Peaje",
                "CASETA": "Caseta Controlada",
                "SENTIDO": "Sentido de Circulacion",
                "TEC_MINUTOS": "Tiempo de Espera en Cola - TEC",
            }
        )[
            ["Peaje", "Caseta Controlada", "Sentido de Circulacion", "Tiempo de Espera en Cola - TEC"]
        ].copy()
        tabla_tec_caseta_informe["Unidades de tiempo"] = "Minutos"

        tabla_tec_peaje = (
            df_resultados.groupby(["PEAJE", "SENTIDO"], dropna=False)
            .agg(
                TEC_MINUTOS=("T_TEC_FINAL_MINUTOS", "mean"),
                VEHICULOS=("PLACA_FINAL", "size"),
            )
            .reset_index()
            .sort_values(["PEAJE", "SENTIDO"])
        )
        tabla_tec_peaje["TEC_MINUTOS"] = tabla_tec_peaje["TEC_MINUTOS"].round(2)
        tabla_tec_peaje_informe = tabla_tec_peaje.rename(
            columns={
                "PEAJE": "Peaje",
                "SENTIDO": "Sentido de Circulacion",
                "TEC_MINUTOS": "Tiempo de Espera en Cola - TEC",
            }
        )[
            ["Peaje", "Sentido de Circulacion", "Tiempo de Espera en Cola - TEC"]
        ].copy()
        tabla_tec_peaje_informe["Unidades de tiempo"] = "Minutos"

        tabla_cola_maxima = (
            df_resultados.groupby(["PEAJE", "CASETA", "SENTIDO"], dropna=False)
            .agg(
                COLA_MAXIMA_REAL=("COLA_ESPERA_USUARIOS", "max"),
                VEHICULOS=("PLACA_FINAL", "size"),
            )
            .reset_index()
            .sort_values(["PEAJE", "SENTIDO", "CASETA"])
        )
        tabla_cola_maxima_informe = tabla_cola_maxima.rename(
            columns={
                "PEAJE": "Peaje",
                "CASETA": "Caseta Controlada",
                "SENTIDO": "Sentido de Circulacion",
                "COLA_MAXIMA_REAL": "Cola maxima real",
                "VEHICULOS": "Vehiculos evaluados",
            }
        )

        def construir_tabla_frecuencia(peaje: str) -> pd.DataFrame:
            sub_df = df_resultados[df_resultados["PEAJE"] == peaje].copy()
            if sub_df.empty:
                return pd.DataFrame(columns=["Cantidad de usuarios en la cola", "Total"])
            tabla = pd.pivot_table(
                sub_df,
                index="COLA_ESPERA_USUARIOS",
                columns="CASETA",
                values="PLACA_FINAL",
                aggfunc="size",
                fill_value=0,
            ).sort_index()
            tabla.columns = [f"{peaje} | {int(columna)}" for columna in tabla.columns]
            tabla["Total"] = tabla.sum(axis=1)
            tabla.loc["Total general"] = tabla.sum(axis=0)
            return tabla.reset_index().rename(columns={"COLA_ESPERA_USUARIOS": "Cantidad de usuarios en la cola"})

        tabla_frecuencia_paraiso_informe = formatear_tabla_frecuencia(construir_tabla_frecuencia("PARAISO"))
        tabla_frecuencia_variante_informe = formatear_tabla_frecuencia(construir_tabla_frecuencia("VARIANTE"))

        fila_mayor_tec_caseta = tabla_tec_caseta.sort_values("TEC_MINUTOS", ascending=False).iloc[0]
        fila_mayor_tec_peaje = tabla_tec_peaje.sort_values("TEC_MINUTOS", ascending=False).iloc[0]
        fila_mayor_cola = tabla_cola_maxima.sort_values("COLA_MAXIMA_REAL", ascending=False).iloc[0]
        resumen_narrativo = pd.DataFrame(
            {
                "Texto sugerido para informe": [
                    (
                        "El mayor TEC promedio por peaje y sentido se observo en "
                        f"{fila_mayor_tec_peaje['PEAJE']} - {fila_mayor_tec_peaje['SENTIDO']}, "
                        f"con {fila_mayor_tec_peaje['TEC_MINUTOS']:.2f} minutos."
                    ),
                    (
                        "La caseta con mayor TEC promedio fue "
                        f"{fila_mayor_tec_caseta['PEAJE']} - caseta {int(fila_mayor_tec_caseta['CASETA'])} - "
                        f"{fila_mayor_tec_caseta['SENTIDO']}, con {fila_mayor_tec_caseta['TEC_MINUTOS']:.2f} minutos."
                    ),
                    (
                        "La mayor cola de espera real se observo en "
                        f"{fila_mayor_cola['PEAJE']} - caseta {int(fila_mayor_cola['CASETA'])} - "
                        f"{fila_mayor_cola['SENTIDO']}, con {int(fila_mayor_cola['COLA_MAXIMA_REAL'])} usuarios en cola."
                    ),
                    (
                        "Las tablas de frecuencia por tamano de cola se calcularon contando solo a los usuarios "
                        "que aun no llegan a la caseta al momento en que arriba un nuevo vehiculo."
                    ),
                ]
            }
        )

    excel_sheets = {
        "tec_caseta": tabla_tec_caseta_informe,
        "tec_peaje_sentido": tabla_tec_peaje_informe,
        "cola_maxima_real": tabla_cola_maxima_informe,
        "cola_paraiso": tabla_frecuencia_paraiso_informe,
        "cola_variante": tabla_frecuencia_variante_informe,
        "texto_informe": resumen_narrativo,
    }
    return {
        "df_resultados": df_resultados,
        "excel_sheets": excel_sheets,
        "tabla_tec_caseta": tabla_tec_caseta_informe,
        "tabla_tec_peaje": tabla_tec_peaje_informe,
        "tabla_cola_maxima": tabla_cola_maxima_informe,
        "tabla_frecuencia_paraiso": tabla_frecuencia_paraiso_informe,
        "tabla_frecuencia_variante": tabla_frecuencia_variante_informe,
        "texto_informe": resumen_narrativo,
    }


def build_complementary_package(df_resultados: pd.DataFrame, df_export_base: pd.DataFrame) -> dict[str, object]:
    resumen_complementario = pd.DataFrame(columns=["indicador", "valor"])
    descriptivos_peaje_sentido = pd.DataFrame(columns=["PEAJE", "SENTIDO", "indicador", "casos", "promedio", "mediana", "p85", "p95", "maximo"])
    descriptivos_caseta = pd.DataFrame(columns=["PEAJE", "CASETA", "SENTIDO", "indicador", "casos", "promedio", "mediana", "p85", "p95", "maximo"])
    cumplimiento_3min_peaje = pd.DataFrame(columns=["PEAJE", "SENTIDO", "vehiculos", "dentro_3_min", "fuera_3_min", "tec_promedio_min", "tec_p95_min", "tec_max_min", "pct_dentro_3_min"])
    cumplimiento_3min_caseta = pd.DataFrame(columns=["PEAJE", "CASETA", "SENTIDO", "vehiculos", "dentro_3_min", "fuera_3_min", "tec_promedio_min", "tec_p95_min", "tec_max_min", "pct_dentro_3_min"])
    top_casetas_tec = pd.DataFrame(columns=descriptivos_caseta.columns)
    top_bloques_30min = pd.DataFrame(columns=["PEAJE", "CASETA", "SENTIDO", "BLOQUE_30MIN", "vehiculos", "tec_promedio_min", "tec_p95_min", "cola_promedio_veh", "cola_maxima_veh"])
    acciones_por_peaje = pd.DataFrame(columns=["PEAJE", "ACCION_REALIZADA", "filas"])
    texto_sugerido_extra = pd.DataFrame({"Texto sugerido": ["No hay datos suficientes para generar resultados complementarios."]})

    if not df_resultados.empty:
        df_resultados_extra = df_resultados.copy()
        df_resultados_extra["FECHA_DT"] = pd.to_datetime(df_resultados_extra["FECHA"], errors="coerce")
        df_resultados_extra["LLEGADA_COLA_TIMESTAMP"] = (
            df_resultados_extra["FECHA_DT"] + df_resultados_extra["LLEGADA_COLA_FINAL_TD"]
        )
        df_resultados_extra["BLOQUE_30MIN"] = (
            df_resultados_extra["LLEGADA_COLA_TIMESTAMP"].dt.floor("30min").dt.strftime("%Y-%m-%d %H:%M")
        )

        resumen_complementario = pd.DataFrame(
            {
                "indicador": [
                    "vehiculos_base_limpia",
                    "registros_ajustados",
                    "porcentaje_ajustados",
                    "tec_promedio_global_min",
                    "tec_mediana_global_min",
                    "tec_p95_global_min",
                    "cola_promedio_global_veh",
                    "cola_p95_global_veh",
                    "cola_maxima_global_veh",
                ],
                "valor": [
                    len(df_resultados_extra),
                    int(df_export_base["REGISTRO_AJUSTADO"].sum()),
                    round(100 * df_export_base["REGISTRO_AJUSTADO"].mean(), 2),
                    round(df_resultados_extra["T_TEC_FINAL_MINUTOS"].mean(), 2),
                    round(df_resultados_extra["T_TEC_FINAL_MINUTOS"].median(), 2),
                    round(df_resultados_extra["T_TEC_FINAL_MINUTOS"].quantile(0.95), 2),
                    round(df_resultados_extra["COLA_ESPERA_USUARIOS"].mean(), 2),
                    round(df_resultados_extra["COLA_ESPERA_USUARIOS"].quantile(0.95), 2),
                    int(df_resultados_extra["COLA_ESPERA_USUARIOS"].max()),
                ],
            }
        )

        descriptivos_peaje_sentido = pd.concat(
            [
                resumir_metricas(df_resultados_extra, ["PEAJE", "SENTIDO"], "T_TEC_FINAL_MINUTOS", "T_TEC_MINUTOS"),
                resumir_metricas(df_resultados_extra, ["PEAJE", "SENTIDO"], "T_COLA_FINAL_MINUTOS", "T_COLA_MINUTOS"),
                resumir_metricas(df_resultados_extra, ["PEAJE", "SENTIDO"], "T_CASETA_FINAL_MINUTOS", "T_CASETA_MINUTOS"),
            ],
            ignore_index=True,
        ).sort_values(["PEAJE", "SENTIDO", "indicador"])

        descriptivos_caseta = resumir_metricas(
            df_resultados_extra,
            ["PEAJE", "CASETA", "SENTIDO"],
            "T_TEC_FINAL_MINUTOS",
            "T_TEC_MINUTOS",
        ).sort_values(["promedio", "p95", "maximo"], ascending=[False, False, False])

        cumplimiento_3min_peaje = (
            df_resultados_extra.groupby(["PEAJE", "SENTIDO"], dropna=False)
            .agg(
                vehiculos=("PLACA_FINAL", "size"),
                dentro_3_min=("T_TEC_FINAL_MINUTOS", lambda s: int((s <= 3).sum())),
                fuera_3_min=("T_TEC_FINAL_MINUTOS", lambda s: int((s > 3).sum())),
                tec_promedio_min=("T_TEC_FINAL_MINUTOS", "mean"),
                tec_p95_min=("T_TEC_FINAL_MINUTOS", q95),
                tec_max_min=("T_TEC_FINAL_MINUTOS", "max"),
            )
            .reset_index()
        )
        cumplimiento_3min_peaje["pct_dentro_3_min"] = (
            100 * cumplimiento_3min_peaje["dentro_3_min"] / cumplimiento_3min_peaje["vehiculos"]
        ).round(2)
        cumplimiento_3min_peaje[["tec_promedio_min", "tec_p95_min", "tec_max_min"]] = cumplimiento_3min_peaje[
            ["tec_promedio_min", "tec_p95_min", "tec_max_min"]
        ].round(2)

        cumplimiento_3min_caseta = (
            df_resultados_extra.groupby(["PEAJE", "CASETA", "SENTIDO"], dropna=False)
            .agg(
                vehiculos=("PLACA_FINAL", "size"),
                dentro_3_min=("T_TEC_FINAL_MINUTOS", lambda s: int((s <= 3).sum())),
                fuera_3_min=("T_TEC_FINAL_MINUTOS", lambda s: int((s > 3).sum())),
                tec_promedio_min=("T_TEC_FINAL_MINUTOS", "mean"),
                tec_p95_min=("T_TEC_FINAL_MINUTOS", q95),
                tec_max_min=("T_TEC_FINAL_MINUTOS", "max"),
            )
            .reset_index()
        )
        cumplimiento_3min_caseta["pct_dentro_3_min"] = (
            100 * cumplimiento_3min_caseta["dentro_3_min"] / cumplimiento_3min_caseta["vehiculos"]
        ).round(2)
        cumplimiento_3min_caseta[["tec_promedio_min", "tec_p95_min", "tec_max_min"]] = cumplimiento_3min_caseta[
            ["tec_promedio_min", "tec_p95_min", "tec_max_min"]
        ].round(2)

        top_casetas_tec = descriptivos_caseta.head(10).copy()

        top_bloques_30min = (
            df_resultados_extra.groupby(["PEAJE", "CASETA", "SENTIDO", "BLOQUE_30MIN"], dropna=False)
            .agg(
                vehiculos=("PLACA_FINAL", "size"),
                tec_promedio_min=("T_TEC_FINAL_MINUTOS", "mean"),
                tec_p95_min=("T_TEC_FINAL_MINUTOS", q95),
                cola_promedio_veh=("COLA_ESPERA_USUARIOS", "mean"),
                cola_maxima_veh=("COLA_ESPERA_USUARIOS", "max"),
            )
            .reset_index()
        )
        top_bloques_30min = top_bloques_30min[top_bloques_30min["vehiculos"] >= 5].copy()
        top_bloques_30min[["tec_promedio_min", "tec_p95_min", "cola_promedio_veh"]] = top_bloques_30min[
            ["tec_promedio_min", "tec_p95_min", "cola_promedio_veh"]
        ].round(2)
        top_bloques_30min = top_bloques_30min.sort_values(
            ["tec_promedio_min", "cola_maxima_veh", "vehiculos"],
            ascending=[False, False, False],
        ).head(30)

        acciones_por_peaje = (
            df_export_base.groupby(["PEAJE", "ACCION_REALIZADA"], dropna=False)
            .size()
            .rename("filas")
            .reset_index()
            .sort_values(["PEAJE", "filas"], ascending=[True, False])
        )

        top_peaje = cumplimiento_3min_peaje.sort_values("tec_promedio_min", ascending=False).iloc[0]
        top_caseta = top_casetas_tec.iloc[0]
        texto_sugerido_extra = pd.DataFrame(
            {
                "Texto sugerido": [
                    (
                        "El "
                        f"{round(100 * df_export_base['REGISTRO_AJUSTADO'].mean(), 2):.2f}% "
                        "de los registros de la base limpia requirio algun ajuste o imputacion."
                    ),
                    (
                        "El peaje/sentido con mayor tiempo de espera promedio fue "
                        f"{top_peaje['PEAJE']} - {top_peaje['SENTIDO']}."
                    ),
                    (
                        "La caseta mas critica por TEC promedio fue "
                        f"{top_caseta['PEAJE']} - caseta {int(top_caseta['CASETA'])} - {top_caseta['SENTIDO']}."
                    ),
                    (
                        "Los bloques de 30 minutos permiten identificar periodos puntuales de mayor congestion "
                        "y priorizar revisiones operativas por caseta."
                    ),
                ]
            }
        )

    excel_sheets = {
        "resumen": resumen_complementario,
        "desc_peaje_sentido": descriptivos_peaje_sentido,
        "desc_caseta_tec": descriptivos_caseta,
        "cumpl_3min_peaje": cumplimiento_3min_peaje,
        "cumpl_3min_caseta": cumplimiento_3min_caseta,
        "top_casetas_tec": top_casetas_tec,
        "top_bloques_30m": top_bloques_30min,
        "acciones_por_peaje": acciones_por_peaje,
        "texto_sugerido": texto_sugerido_extra,
    }
    return {
        "excel_sheets": excel_sheets,
        "top_casetas_tec": top_casetas_tec,
        "top_bloques_30m": top_bloques_30min,
    }


def to_excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for sheet_name, df_sheet in sheets.items():
            safe_name = sheet_name[:31]
            df_sheet.to_excel(writer, sheet_name=safe_name, index=False)
    output.seek(0)
    return output.getvalue()


def write_report_block(
    writer: pd.ExcelWriter,
    sheet_name: str,
    title: str,
    table: pd.DataFrame,
    startrow: int,
) -> int:
    pd.DataFrame({title: []}).to_excel(writer, sheet_name=sheet_name, startrow=startrow, index=False)
    worksheet = writer.sheets[sheet_name]
    worksheet.cell(row=startrow + 1, column=1, value=title)
    table.to_excel(writer, sheet_name=sheet_name, startrow=startrow + 2, index=False)
    return startrow + len(table) + 5


def to_exact_excel_bytes(export_package: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        export_package["base_limpia"].to_excel(writer, sheet_name="base_limpia", index=False)
        export_package["casos_eliminados"].to_excel(writer, sheet_name="casos_eliminados", index=False)

        startrow = 0
        startrow = write_report_block(
            writer,
            "reporte_resumen",
            "Resumen general",
            export_package["resumen_general"],
            startrow,
        )
        startrow = write_report_block(
            writer,
            "reporte_resumen",
            "Resumen de acciones en base_limpia",
            export_package["resumen_accion_realizada"],
            startrow,
        )
        startrow = write_report_block(
            writer,
            "reporte_resumen",
            "Resumen de acciones de placa",
            export_package["resumen_acciones_placa"],
            startrow,
        )
        startrow = write_report_block(
            writer,
            "reporte_resumen",
            "Resumen de acciones de tiempo",
            export_package["resumen_acciones_tiempo"],
            startrow,
        )
        write_report_block(
            writer,
            "reporte_resumen",
            "Resumen de casos eliminados",
            export_package["resumen_eliminados"],
            startrow,
        )
    output.seek(0)
    return output.getvalue()


def to_docx_bytes(report_label: str, informe_package: dict[str, object]) -> bytes:
    doc = Document()
    doc.add_heading(f"Tablas de resultados {report_label}", level=1)
    doc.add_paragraph(
        "Base utilizada: registros finales limpios. La cola de espera se define como los usuarios "
        "que aun esperan llegar a la caseta cuando arriba un nuevo vehiculo."
    )
    agregar_tabla_docx(doc, "Tabla 1. Tiempo de Espera en Cola por caseta y sentido", informe_package["tabla_tec_caseta"])
    agregar_tabla_docx(doc, "Tabla 2. Promedio de tiempo de espera en cola por peaje segun sentido", informe_package["tabla_tec_peaje"])
    agregar_tabla_docx(doc, "Tabla 3. Cola maxima de espera real por caseta", informe_package["tabla_cola_maxima"])
    agregar_tabla_docx(
        doc,
        "Tabla 4. Frecuencia de usuarios segun tamano de cola de espera para la estacion de peaje Paraiso por caseta",
        informe_package["tabla_frecuencia_paraiso"],
    )
    agregar_tabla_docx(
        doc,
        "Tabla 5. Frecuencia de usuarios segun tamano de cola de espera para la estacion de peaje Variante por caseta",
        informe_package["tabla_frecuencia_variante"],
    )
    doc.add_paragraph("Texto de apoyo para conclusiones", style="Heading 2")
    for texto in informe_package["texto_informe"]["Texto sugerido para informe"]:
        doc.add_paragraph(texto, style="List Bullet")

    output = BytesIO()
    doc.save(output)
    output.seek(0)
    return output.getvalue()


def to_zip_bytes(sheets: dict[str, pd.DataFrame], excel_bytes: bytes) -> bytes:
    buffer = BytesIO()
    with zipfile.ZipFile(buffer, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
        zf.writestr("procesamiento_tec_output.xlsx", excel_bytes)
        for name, df_sheet in sheets.items():
            zf.writestr(f"{name}.csv", df_sheet.to_csv(index=False).encode("utf-8-sig"))
    buffer.seek(0)
    return buffer.getvalue()


def build_run_payload(
    source_name: str,
    mapping: dict[str, str | None],
    config: dict,
    export_tables: dict[str, pd.DataFrame],
) -> dict:
    return {
        "source_name": source_name,
        "input_rows": int(export_tables["reporte_resumen"].loc[
            export_tables["reporte_resumen"]["indicador"] == "filas_entrada", "valor"
        ].iloc[0]),
        "clean_rows": len(export_tables["base_limpia"]),
        "deleted_rows": len(export_tables["casos_eliminados"]),
        "pending_rows": len(export_tables["casos_pendientes"]),
        "mapping": {k: v for k, v in mapping.items() if v},
        "config": config,
        "notes": {
            "sheets": list(export_tables.keys()),
        },
    }


def process_pipeline(df_std: pd.DataFrame, config: dict, manual_rules_df: pd.DataFrame) -> dict[str, object]:
    df_std = df_std.copy()
    df_std["_ORDEN_FILA"] = range(len(df_std))
    if config["aplicar_limpieza_placa"]:
        plate_result = run_plate_cleaning(df_std, config, manual_rules_df)
    else:
        df_passthrough = limpiar_placas_peaje(df_std)
        df_passthrough["_ORDEN_FILA"] = df_std["_ORDEN_FILA"]
        df_passthrough["PLACA_ACCION_FINAL"] = "sin_cambio"
        df_passthrough["PLACA_FINAL_DECIDIDA"] = df_passthrough["PLACA_NORMALIZADA"]
        df_passthrough["PLACA_EXCLUIR_ANALISIS"] = False
        df_passthrough["PLACA_AJUSTE_MANUAL"] = pd.NA
        plate_result = {
            "df": df_passthrough,
            "df_trabajo": df_passthrough.copy(),
            "df_eliminados": pd.DataFrame(columns=df_passthrough.columns),
            "df_revision_placas": pd.DataFrame(),
            "df_bloques_decision": pd.DataFrame(),
            "resumen_acciones_placa": df_passthrough["PLACA_ACCION_FINAL"].value_counts(dropna=False).rename_axis("PLACA_ACCION_FINAL").to_frame("filas"),
        }

    time_result = run_time_cleaning(plate_result["df_trabajo"], config)
    export_tables = build_export_tables(df_std, plate_result, time_result, config, manual_rules_df)
    exact_export = build_exact_export_package(df_std, export_tables)
    informe_package = build_informe_package(export_tables["export_base_detalle"])
    complementary_package = build_complementary_package(
        informe_package["df_resultados"],
        export_tables["export_base_detalle"],
    )
    return {
        "plate_result": plate_result,
        "time_result": time_result,
        "export_tables": export_tables,
        "exact_export": exact_export,
        "informe_package": informe_package,
        "complementary_package": complementary_package,
    }


def load_streamlit_secrets() -> dict:
    try:
        return dict(st.secrets)
    except StreamlitSecretNotFoundError:
        return {}


def explain_auth_failure(reason: str) -> str:
    messages = {
        "invalid_credentials": "Usuario o contrasena no validos.",
        "disabled": "Tu usuario esta deshabilitado.",
        "not_yet_active": "Tu usuario todavia no entra en vigencia.",
        "expired": "Tu usuario ya vencio.",
        "auth_not_available": "El backend actual no tiene autenticacion habilitada.",
    }
    return messages.get(reason, "No fue posible iniciar sesion.")


def normalize_date_input(value) -> datetime | None:
    if value is None or pd.isna(value):
        return None
    if isinstance(value, pd.Timestamp):
        return value.to_pydatetime()
    if isinstance(value, datetime):
        return value
    return datetime.combine(value, time.min)


def date_to_window_start(value) -> datetime | None:
    parsed = normalize_date_input(value)
    if not parsed:
        return None
    return parsed.replace(hour=0, minute=0, second=0, microsecond=0)


def date_to_window_end(value) -> datetime | None:
    parsed = normalize_date_input(value)
    if not parsed:
        return None
    return parsed.replace(hour=23, minute=59, second=59, microsecond=0)


def date_value_or_none(value):
    parsed = normalize_date_input(value)
    return parsed.date() if parsed else None


def password_is_valid(password: str) -> bool:
    return len(password) >= 8


def build_available_users_df(users_df: pd.DataFrame) -> pd.DataFrame:
    available_users_df = users_df.copy()
    if "is_enabled" in available_users_df.columns:
        available_users_df["estado"] = available_users_df["is_enabled"].map(
            {True: "Habilitado", False: "Deshabilitado"}
        ).fillna("Sin dato")
    for column_name in ["active_from", "active_until", "last_login_at", "created_at"]:
        if column_name in available_users_df.columns:
            available_users_df[column_name] = pd.to_datetime(
                available_users_df[column_name], errors="coerce"
            ).dt.strftime("%Y-%m-%d %H:%M")
            available_users_df[column_name] = available_users_df[column_name].fillna("-")

    visible_columns = [
        column_name
        for column_name in [
            "username",
            "full_name",
            "role_name",
            "estado",
            "email",
            "phone_number",
            "active_from",
            "active_until",
            "last_login_at",
        ]
        if column_name in available_users_df.columns
    ]
    return available_users_df[visible_columns]


def render_bootstrap_admin(storage_backend) -> None:
    _, center_col, _ = st.columns([1, 0.9, 1], gap="large")
    with center_col:
        st.markdown(
            '<section class="auth-card"><div class="auth-card-kicker">Primer acceso</div><div class="auth-card-title">Crear administrador inicial</div><p class="auth-card-copy">Registra la primera cuenta con control total sobre usuarios, accesos e historial.</p></section>',
            unsafe_allow_html=True,
        )
        with st.form("bootstrap_admin_form"):
            full_name = st.text_input("Nombre completo")
            username = st.text_input("Usuario")
            email = st.text_input("Correo electronico")
            phone_number = st.text_input("Celular")
            password = st.text_input("Contrasena", type="password")
            confirm_password = st.text_input("Confirmar contrasena", type="password")
            submitted = st.form_submit_button("Crear administrador", use_container_width=True)

    if not submitted:
        return

    if not full_name.strip() or not username.strip() or not email.strip() or not phone_number.strip():
        st.error("Nombre completo, usuario, correo y celular son obligatorios.")
        return
    if password != confirm_password:
        st.error("Las contrasenas no coinciden.")
        return
    if not password_is_valid(password):
        st.error("La contrasena debe tener al menos 8 caracteres.")
        return

    try:
        storage_backend.create_initial_admin(username, full_name, password, email.strip(), phone_number.strip())
        auth_result = storage_backend.authenticate_user(username, password)
    except Exception as exc:
        st.error(f"No pude crear el administrador inicial: {exc}")
        return

    if auth_result.get("ok"):
        set_authenticated_user(auth_result["user"])
        st.session_state[APP_NAV_KEY] = "Inicio"
        st.query_params.clear()
        st.rerun()


def render_login_gate(storage_backend):
    if not storage_backend.has_users():
        render_bootstrap_admin(storage_backend)
        return None

    users_df = storage_backend.list_users().copy()
    if "id" in users_df.columns:
        users_df["id"] = pd.to_numeric(users_df["id"], errors="coerce")
        users_df = users_df[users_df["id"].notna()].copy()
    user_options = users_df.to_dict("records")

    _, center_col, _ = st.columns([1, 0.85, 1], gap="large")
    with center_col:
        st.markdown(
            '<section class="auth-card"><div class="auth-card-kicker">Ingreso</div><div class="auth-card-title">Iniciar sesion</div><p class="auth-card-copy">Ingresa tus credenciales para acceder al aplicativo.</p></section>',
            unsafe_allow_html=True,
        )
        with st.form("login_form"):
            selected_user = st.selectbox(
                "Usuario",
                options=user_options,
                format_func=lambda row: f"{row['username']} | {row['full_name']} | {row['role_name']}",
            )
            password = st.text_input("Contrasena", type="password")
            submitted = st.form_submit_button("Ingresar", use_container_width=True)

    if not submitted:
        return None

    auth_result = storage_backend.authenticate_user(selected_user["username"], password)
    if not auth_result.get("ok"):
        st.error(explain_auth_failure(auth_result.get("reason", "")))
        return None

    set_authenticated_user(auth_result["user"])
    st.session_state[APP_NAV_KEY] = "Inicio"
    st.query_params.clear()
    st.rerun()
    return None


def render_history_page(storage_backend) -> None:
    st.markdown(
        build_hero_panel(
            title="Historial de corridas",
            copy="Consulta ejecuciones recientes guardadas en la base de datos para seguimiento operativo y trazabilidad.",
            kicker="Historial",
            metrics=[
                (storage_backend.mode.upper(), "backend"),
                ("50", "maximo visible"),
            ],
        ),
        unsafe_allow_html=True,
    )
    st.dataframe(storage_backend.list_recent_runs(50), use_container_width=True)


def render_user_management_page(storage_backend, current_user: dict) -> None:
    can_manage_users = user_has_permission("manage_users", current_user)
    can_manage_activation = user_has_permission("manage_user_activation", current_user)

    st.markdown(
        build_hero_panel(
            title="Gestion de usuarios",
            copy="Administra altas, roles, vigencias y credenciales de acceso con una vista centralizada para operacion y control.",
            kicker="Administracion",
            metrics=[
                ("3", "roles base"),
                ("Correo + celular", "datos obligatorios"),
                (current_user["role_label"], "tu perfil"),
            ],
        ),
        unsafe_allow_html=True,
    )
    st.caption("Roles disponibles: administrador general, administrador operativo y analista.")

    if not (can_manage_users or can_manage_activation):
        st.error("No tienes permisos para gestionar usuarios.")
        return

    roles_info = pd.DataFrame(
        [
            {
                "rol": role_key,
                "nombre": role_data["label"],
                "descripcion": role_data["description"],
                "permisos": ", ".join(role_data["permissions"]),
            }
            for role_key, role_data in ROLE_DEFINITIONS.items()
        ]
    )
    st.dataframe(roles_info, use_container_width=True, hide_index=True)

    users_df = storage_backend.list_users().copy()
    if "id" in users_df.columns:
        users_df["id"] = pd.to_numeric(users_df["id"], errors="coerce")
        users_df = users_df[users_df["id"].notna()].copy()

    st.subheader("Usuarios actuales")
    st.caption(f"Usuarios disponibles en la base de datos: {len(users_df)}")
    st.dataframe(build_available_users_df(users_df), use_container_width=True, hide_index=True)

    if users_df.empty:
        st.info("Aun no hay usuarios registrados.")
        return

    user_options = {
        f"{row['username']} | {row['role_name']}": row.to_dict()
        for _, row in users_df.iterrows()
    }
    selected_user_label = st.selectbox("Usuario a editar", options=list(user_options.keys()))
    selected_user = user_options[selected_user_label]
    current_user_id = pd.to_numeric(pd.Series([current_user.get("id")]), errors="coerce").iloc[0]
    selected_user_id = int(selected_user["id"])
    selected_is_self = pd.notna(current_user_id) and selected_user_id == int(current_user_id)

    current_active_from = date_value_or_none(selected_user.get("active_from"))
    current_active_until = date_value_or_none(selected_user.get("active_until"))
    can_edit_access = can_manage_users or can_manage_activation

    st.subheader("Actualizar usuario")
    with st.form("edit_user_form"):
        full_name = st.text_input(
            "Nombre completo",
            value=str(selected_user.get("full_name") or ""),
            disabled=not can_manage_users,
        )
        email = st.text_input(
            "Correo electronico",
            value="" if pd.isna(selected_user.get("email")) else str(selected_user.get("email")),
            disabled=not can_manage_users,
        )
        phone_number = st.text_input(
            "Celular",
            value="" if pd.isna(selected_user.get("phone_number")) else str(selected_user.get("phone_number")),
            disabled=not can_manage_users,
        )
        role_keys = list(ROLE_DEFINITIONS.keys())
        role_index = role_keys.index(selected_user["role_key"]) if selected_user["role_key"] in role_keys else 0
        role_key = st.selectbox(
            "Rol",
            options=role_keys,
            index=role_index,
            format_func=lambda key: ROLE_DEFINITIONS[key]["label"],
            disabled=not can_manage_users,
        )
        is_enabled = st.checkbox(
            "Usuario habilitado",
            value=bool(selected_user.get("is_enabled")),
            disabled=not can_edit_access,
        )
        use_active_from = st.checkbox(
            "Restringir fecha inicial",
            value=current_active_from is not None,
            key=f"edit_use_active_from_{selected_user['id']}",
            disabled=not can_edit_access,
        )
        active_from = st.date_input(
            "Activo desde",
            value=current_active_from or datetime.now().date(),
            key=f"edit_active_from_{selected_user['id']}",
            disabled=(not can_edit_access) or (not use_active_from),
        )
        use_active_until = st.checkbox(
            "Restringir fecha final",
            value=current_active_until is not None,
            key=f"edit_use_active_until_{selected_user['id']}",
            disabled=not can_edit_access,
        )
        active_until = st.date_input(
            "Activo hasta",
            value=current_active_until or datetime.now().date(),
            key=f"edit_active_until_{selected_user['id']}",
            disabled=(not can_edit_access) or (not use_active_until),
        )
        new_password = st.text_input(
            "Nueva contrasena",
            type="password",
            disabled=not can_manage_users,
        )
        confirm_password = st.text_input(
            "Confirmar nueva contrasena",
            type="password",
            disabled=not can_manage_users,
        )
        submitted = st.form_submit_button("Guardar cambios", use_container_width=True)

    if submitted:
        if selected_is_self and can_edit_access and not is_enabled:
            st.error("No puedes deshabilitar tu propio usuario desde esta sesion.")
            return
        if selected_is_self and can_manage_users and role_key != selected_user["role_key"]:
            st.error("No puedes cambiar tu propio rol mientras estas autenticado.")
            return
        if can_manage_users and new_password:
            if new_password != confirm_password:
                st.error("Las contrasenas no coinciden.")
                return
            if not password_is_valid(new_password):
                st.error("La nueva contrasena debe tener al menos 8 caracteres.")
                return
        if can_manage_users and (not email.strip() or not phone_number.strip()):
            st.error("Correo electronico y celular son obligatorios.")
            return

        payload = {}
        if can_manage_users:
            payload.update(
                {
                    "full_name": full_name.strip(),
                    "email": email.strip(),
                    "phone_number": phone_number.strip(),
                    "role_key": role_key,
                }
            )
            if new_password:
                payload["password"] = new_password
        if can_edit_access:
            payload.update(
                {
                    "is_enabled": is_enabled,
                    "active_from": date_to_window_start(active_from) if use_active_from else None,
                    "active_until": date_to_window_end(active_until) if use_active_until else None,
                }
            )

        if payload.get("active_from") and payload.get("active_until") and payload["active_from"] > payload["active_until"]:
            st.error("La fecha inicial no puede ser mayor que la fecha final.")
            return

        try:
            storage_backend.update_user(selected_user_id, payload)
        except Exception as exc:
            st.error(f"No pude actualizar el usuario: {exc}")
            return

        if selected_is_self and can_manage_users:
            current_user["full_name"] = full_name.strip()
            current_user["email"] = email.strip()
            current_user["phone_number"] = phone_number.strip()
            set_authenticated_user(current_user)
        st.success("Usuario actualizado.")
        st.rerun()

    if not can_manage_users:
        st.info("Tu rol solo puede habilitar, deshabilitar y definir vigencia de usuarios.")
        return

    st.subheader("Crear usuario")
    with st.form("create_user_form"):
        username = st.text_input("Nuevo usuario")
        full_name = st.text_input("Nombre completo", key="create_full_name")
        email = st.text_input("Correo electronico", key="create_email")
        phone_number = st.text_input("Celular", key="create_phone_number")
        password = st.text_input("Contrasena inicial", type="password", key="create_password")
        confirm_password = st.text_input("Confirmar contrasena", type="password", key="create_confirm_password")
        role_key = st.selectbox(
            "Rol inicial",
            options=list(ROLE_DEFINITIONS.keys()),
            format_func=lambda key: ROLE_DEFINITIONS[key]["label"],
            key="create_role_key",
        )
        is_enabled = st.checkbox("Crear usuario habilitado", value=True, key="create_enabled")
        use_active_from = st.checkbox("Definir fecha inicial", value=False, key="create_use_active_from")
        active_from = st.date_input(
            "Activo desde",
            value=datetime.now().date(),
            key="create_active_from",
            disabled=not use_active_from,
        )
        use_active_until = st.checkbox("Definir fecha final", value=False, key="create_use_active_until")
        active_until = st.date_input(
            "Activo hasta",
            value=datetime.now().date(),
            key="create_active_until",
            disabled=not use_active_until,
        )
        submitted = st.form_submit_button("Crear usuario", use_container_width=True)

    if not submitted:
        return

    if password != confirm_password:
        st.error("Las contrasenas no coinciden.")
        return
    if not password_is_valid(password):
        st.error("La contrasena inicial debe tener al menos 8 caracteres.")
        return
    if not email.strip() or not phone_number.strip():
        st.error("Correo electronico y celular son obligatorios.")
        return

    payload = {
        "username": username,
        "full_name": full_name,
        "email": email.strip(),
        "phone_number": phone_number.strip(),
        "password": password,
        "role_key": role_key,
        "is_enabled": is_enabled,
        "active_from": date_to_window_start(active_from) if use_active_from else None,
        "active_until": date_to_window_end(active_until) if use_active_until else None,
        "created_by": current_user["id"],
    }
    if payload["active_from"] and payload["active_until"] and payload["active_from"] > payload["active_until"]:
        st.error("La fecha inicial no puede ser mayor que la fecha final.")
        return

    try:
        storage_backend.create_user(payload)
    except Exception as exc:
        st.error(f"No pude crear el usuario: {exc}")
        return

    st.success("Usuario creado.")
    st.rerun()


def render_processing_page(storage_backend, current_user: dict | None) -> None:
    st.markdown(
        build_hero_panel(
            title="Modulo TEC",
            copy="Carga bases, revisa columnas, controla la configuracion del pipeline y descarga los entregables resultantes desde un unico flujo de trabajo.",
            kicker="Procesamiento activo",
            metrics=[
                ("TEC", "modulo"),
                (storage_backend.mode.upper(), "persistencia"),
                (current_user["role_label"] if current_user else "Libre", "perfil"),
            ],
        ),
        unsafe_allow_html=True,
    )
    can_manage_general_config = not storage_backend.auth_enabled() or user_has_permission(
        "manage_general_config",
        current_user,
    )
    can_view_history = not storage_backend.auth_enabled() or user_has_permission("view_history", current_user)

    uploaded_file = st.file_uploader("Sube tu archivo Excel o CSV", type=["xlsx", "xls", "xlsm", "csv"])
    if not uploaded_file:
        if storage_backend.mode != "none" and can_view_history:
            st.info(f"Persistencia activa: `{storage_backend.mode}`")
            st.dataframe(storage_backend.list_recent_runs(10), use_container_width=True)
        else:
            st.info("Sube una base para comenzar.")
        return

    sheet_options = list_excel_sheets(uploaded_file)
    selected_sheet = None
    if sheet_options:
        selected_sheet = st.selectbox("Hoja a procesar", options=sheet_options, index=0)

    try:
        df_raw = load_input_dataframe(uploaded_file, selected_sheet)
    except Exception as exc:
        st.error(f"No pude leer el archivo: {exc}")
        return

    st.subheader("Vista previa")
    st.dataframe(df_raw.head(10), use_container_width=True)

    columnas = df_raw.columns.tolist()
    suggestions = {
        "PLACA": suggest_column(columnas, ["placa", "plate"]),
        "LLEGADA COLA": suggest_column(columnas, ["llegada cola", "t1", "hora cola"]),
        "LLEGADA CASETA": suggest_column(columnas, ["llegada caseta", "t2", "hora caseta"]),
        "SALIDA CASETA": suggest_column(columnas, ["salida caseta", "t3", "hora salida"]),
        "PEAJE": suggest_column(columnas, ["peaje"]),
        "CASETA": suggest_column(columnas, ["caseta"]),
        "SENTIDO": suggest_column(columnas, ["sentido"]),
        "FECHA": suggest_column(columnas, ["fecha", "date"]),
        "VEHICULO": suggest_column(columnas, ["vehiculo", "vehicle"]),
        "T. TEC": suggest_column(columnas, ["t. tec", "t tec", "tec"]),
        "T. CASETA": suggest_column(columnas, ["t. caseta", "t caseta"]),
    }

    st.subheader("Mapeo de columnas")
    col1, col2, col3 = st.columns(3)
    mapping = {}
    selectors = [
        ("PLACA", col1),
        ("LLEGADA COLA", col1),
        ("LLEGADA CASETA", col1),
        ("SALIDA CASETA", col1),
        ("PEAJE", col2),
        ("CASETA", col2),
        ("SENTIDO", col2),
        ("FECHA", col2),
        ("VEHICULO", col3),
        ("T. TEC", col3),
        ("T. CASETA", col3),
    ]
    options_required = [None] + columnas
    for target, container in selectors:
        default = suggestions[target]
        index = options_required.index(default) if default in options_required else 0
        mapping[target] = container.selectbox(
            f"Columna para {target}",
            options=options_required,
            index=index,
            key=f"map_{target}",
        )

    with st.sidebar:
        st.header("Configuracion")
        st.caption(f"Persistencia: `{storage_backend.mode}`")
        if not can_manage_general_config:
            st.caption("Tu rol no puede cambiar la configuracion general. Se aplican los valores definidos por administracion.")
        st.caption("Si desmarcas una opcion, esa etapa no se aplicara en esta corrida.")
        st.markdown("**Placas**")
        config = {
            "aplicar_limpieza_placa": st.checkbox(
                "Revisar y corregir placas automaticamente",
                value=DEFAULT_CONFIG["aplicar_limpieza_placa"],
                disabled=not can_manage_general_config,
                help="Activa toda la etapa de analisis de placas antes de trabajar los tiempos.",
            ),
            "aplicar_ruido_con_respaldo": st.checkbox(
                "Corregir espacios o simbolos cuando la placa limpia ya aparece en la base",
                value=DEFAULT_CONFIG["aplicar_ruido_con_respaldo"],
                disabled=not can_manage_general_config,
                help="Ejemplo: convertir 'CFQ 004' en 'CFQ004' cuando la version limpia ya existe en la base.",
            ),
            "aplicar_ruido_sin_respaldo": st.checkbox(
                "Corregir espacios o simbolos aunque no haya otra coincidencia",
                value=DEFAULT_CONFIG["aplicar_ruido_sin_respaldo"],
                disabled=not can_manage_general_config,
                help="Aplica una limpieza mas agresiva a placas con guiones, espacios o simbolos.",
            ),
            "aplicar_coincidencia_unica_lista": st.checkbox(
                "Corregir placas cuando hay una unica coincidencia clara en la base",
                value=DEFAULT_CONFIG["aplicar_coincidencia_unica_lista"],
                disabled=not can_manage_general_config,
                help="Busca una sola candidata muy parecida dentro de toda la base y corrige hacia esa placa.",
            ),
            "aplicar_confusion_visual": st.checkbox(
                "Corregir confusiones visuales como O/0, I/1 o S/5",
                value=DEFAULT_CONFIG["aplicar_confusion_visual"],
                disabled=not can_manage_general_config,
                help="Corrige placas que parecen tener letras y numeros confundidos visualmente.",
            ),
            "aplicar_recorte_sufijo_x": st.checkbox(
                "Quitar una X final cuando parece un sufijo extra",
                value=DEFAULT_CONFIG["aplicar_recorte_sufijo_x"],
                disabled=not can_manage_general_config,
                help="Se usa en placas donde la X final parece agregada y la version recortada tiene sentido.",
            ),
            "aplicar_exclusion_placeholders": st.checkbox(
                "Excluir placas de prueba o ejemplo",
                value=DEFAULT_CONFIG["aplicar_exclusion_placeholders"],
                disabled=not can_manage_general_config,
                help="Elimina valores como placas placeholder o ejemplos que no deben analizarse.",
            ),
            "aplicar_reglas_manuales": st.checkbox(
                "Aplicar las reglas manuales de la tabla inferior",
                value=DEFAULT_CONFIG["aplicar_reglas_manuales"],
                disabled=not can_manage_general_config,
                help="Usa las correcciones o exclusiones definidas manualmente por ti.",
            ),
        }
        st.markdown("**Tiempos**")
        config.update(
            {
                "eliminar_bordes_caseta": st.checkbox(
                    "Eliminar registros incompletos al inicio o final de cada flujo",
                    value=DEFAULT_CONFIG["eliminar_bordes_caseta"],
                    disabled=not can_manage_general_config,
                    help="Quita filas incompletas que quedan antes de que el flujo empiece bien o despues de que ya termino.",
                ),
                "aplicar_interpolacion": st.checkbox(
                    "Completar tiempos faltantes usando vehiculos vecinos del mismo flujo",
                    value=DEFAULT_CONFIG["aplicar_interpolacion"],
                    disabled=not can_manage_general_config,
                    help="Reconstruye tiempos internos cuando hay anclas antes y despues dentro del mismo flujo.",
                ),
                "aplicar_mediana_local": st.checkbox(
                    "Hacer una segunda recuperacion usando medianas del mismo flujo",
                    value=DEFAULT_CONFIG["aplicar_mediana_local"],
                    disabled=not can_manage_general_config,
                    help="Intenta completar tiempos pendientes con duraciones tipicas observadas en ese mismo flujo.",
                ),
                "aplicar_donantes": st.checkbox(
                    "Hacer una tercera recuperacion usando casos similares",
                    value=DEFAULT_CONFIG["aplicar_donantes"],
                    disabled=not can_manage_general_config,
                    help="Usa medianas de casetas, dias o sentidos parecidos como apoyo para recuperar tiempos.",
                ),
                "aplicar_swap_final_t2_t3": st.checkbox(
                    "Corregir un posible intercambio final entre llegada y salida de caseta",
                    value=DEFAULT_CONFIG["aplicar_swap_final_t2_t3"],
                    disabled=not can_manage_general_config,
                    help="Aplica un ajuste final si los tiempos de llegada y salida quedaron invertidos.",
                ),
            }
        )

    st.subheader("Reglas manuales")
    reglas_df = st.data_editor(
        pd.DataFrame(DEFAULT_MANUAL_RULES),
        num_rows="dynamic",
        use_container_width=True,
        disabled=not can_manage_general_config,
        key="manual_rules_editor",
    )

    if st.button("Procesar base", type="primary", use_container_width=True):
        faltantes = [campo for campo in ["PLACA", "LLEGADA COLA", "LLEGADA CASETA", "SALIDA CASETA"] if not mapping.get(campo)]
        if faltantes:
            st.error(f"Falta mapear estas columnas obligatorias: {', '.join(faltantes)}")
            return
        if config["eliminar_bordes_caseta"] and not mapping.get("FECHA"):
            st.error("Para procesar tiempos necesitas mapear FECHA.")
            return

        df_std = build_standardized_df(df_raw, mapping)
        if df_std["FECHA"].isna().all():
            st.warning("La columna FECHA quedo vacia; los pasos de tiempos pueden perder precision o fallar.")

        try:
            result = process_pipeline(df_std, config, reglas_df)
        except Exception as exc:
            st.exception(exc)
            return

        export_tables = result["export_tables"]
        exact_export = result["exact_export"]
        informe_package = result["informe_package"]
        complementary_package = result["complementary_package"]
        output_names = derive_output_filenames(uploaded_file.name)
        excel_bytes = to_exact_excel_bytes(exact_export)
        report_excel_bytes = to_excel_bytes(informe_package["excel_sheets"])
        report_docx_bytes = to_docx_bytes(output_names["report_label"], informe_package)
        complementary_excel_bytes = to_excel_bytes(complementary_package["excel_sheets"])
        storage_backend.save_run(
            build_run_payload(uploaded_file.name, mapping, config, export_tables)
        )

        resumen = export_tables["reporte_resumen"]
        general = resumen[resumen["seccion"] == "general"]
        c1, c2, c3, c4 = st.columns(4)
        metric_values = {row["indicador"]: row["valor"] for _, row in general.iterrows()}
        c1.metric("Filas entrada", metric_values.get("filas_entrada", 0))
        c2.metric("Base limpia", metric_values.get("filas_base_limpia", 0))
        c3.metric("Eliminados", metric_values.get("filas_eliminadas", 0))
        c4.metric("Pendientes", metric_values.get("filas_pendientes", 0))

        st.download_button(
            "Descargar Excel",
            data=excel_bytes,
            file_name=output_names["clean_excel"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Descargar Resultados Informe",
            data=report_excel_bytes,
            file_name=output_names["report_excel"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )
        st.download_button(
            "Descargar Tablas Informe",
            data=report_docx_bytes,
            file_name=output_names["report_docx"],
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
        st.download_button(
            "Descargar Resultados Complementarios",
            data=complementary_excel_bytes,
            file_name=output_names["extra_excel"],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

        tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
            ["Resumen", "Base limpia", "Eliminados", "Pendientes", "Revision placas", "Bloques"]
        )
        with tab1:
            st.dataframe(export_tables["reporte_resumen"], use_container_width=True)
        with tab2:
            st.dataframe(export_tables["base_limpia"], use_container_width=True)
        with tab3:
            st.dataframe(export_tables["casos_eliminados"], use_container_width=True)
        with tab4:
            st.dataframe(export_tables["casos_pendientes"], use_container_width=True)
        with tab5:
            st.dataframe(export_tables["revision_placas"], use_container_width=True)
        with tab6:
            st.dataframe(export_tables["bloques_decision"], use_container_width=True)

        if storage_backend.mode != "none" and can_view_history:
            st.subheader("Historial guardado")
            st.dataframe(storage_backend.list_recent_runs(20), use_container_width=True)


def main() -> None:
    st.set_page_config(page_title=APP_TITLE, layout="wide")
    inject_global_styles()
    storage_backend = build_storage_backend(load_streamlit_secrets())

    if storage_backend.auth_enabled():
        current_user = get_authenticated_user()
        if not current_user:
            with st.sidebar:
                render_contractor_branding()
            render_login_gate(storage_backend)
            return
    else:
        current_user = None

    can_open_user_management = storage_backend.auth_enabled() and (
        user_has_permission("manage_users", current_user)
        or user_has_permission("manage_user_activation", current_user)
    )

    can_view_history = storage_backend.mode != "none" and (
        not storage_backend.auth_enabled() or user_has_permission("view_history", current_user)
    )

    page_options = ["Inicio"]
    if not storage_backend.auth_enabled() or user_has_permission("process_files", current_user):
        page_options.append("TEC")
    page_options.extend(["Relevamientos", "Auditorias", "Satisfaccion", "Flujogramas"])

    requested_page = get_requested_page()
    if requested_page:
        st.session_state[APP_NAV_KEY] = requested_page

    current_page = st.session_state.get(APP_NAV_KEY, "Inicio")
    valid_pages = set(page_options)
    if can_open_user_management:
        valid_pages.add("Usuarios")
    if can_view_history:
        valid_pages.add("Historial")
    if current_page not in valid_pages:
        current_page = "Inicio"
        st.session_state[APP_NAV_KEY] = current_page

    with st.sidebar:
        if current_user:
            st.header("Sesion")
            st.write(current_user["full_name"])
            st.caption(f"Usuario: {current_user['username']}")
            st.caption(f"Rol: {current_user['role_label']}")
            st.caption(describe_access_window(current_user.get("active_from"), current_user.get("active_until")))
            action_cols = st.columns([5, 1]) if can_open_user_management else st.columns([1])
            with action_cols[0]:
                if st.button("Cerrar sesion", key="sidebar_logout_button", use_container_width=True):
                    st.session_state[APP_NAV_KEY] = "Inicio"
                    st.query_params.clear()
                    clear_authenticated_user()
                    st.rerun()
            if can_open_user_management:
                with action_cols[1]:
                    if st.button(
                        "⚙",
                        key="sidebar_users_button",
                        help="Usuarios",
                        use_container_width=True,
                        type="primary" if current_page == "Usuarios" else "secondary",
                    ):
                        navigate_to("Usuarios")

            if can_view_history:
                st.markdown(
                    '<div class="sidebar-text-link"><a href="?page=Historial">Historial</a></div>',
                    unsafe_allow_html=True,
                )
            st.divider()
        page = current_page

    if page == "Inicio":
        render_home_page(current_user)
        return
    if page == "Usuarios":
        render_user_management_page(storage_backend, current_user)
        return
    if page == "Historial":
        render_history_page(storage_backend)
        return
    if page in MODULE_PLACEHOLDERS:
        render_placeholder_module_page(page)
        return

    render_processing_page(storage_backend, current_user)


if __name__ == "__main__":
    main()
