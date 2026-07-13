# -*- coding: utf-8 -*-
import os
import datetime as dt
import textwrap

import numpy as np
import pandas as pd
import streamlit as st
import streamlit.components.v1 as components


# =============================================================================
# CONFIG
# =============================================================================
st.set_page_config(
    page_title="EMS Elektroauto",
    layout="wide",
    initial_sidebar_state="collapsed"
)

EXCEL_DATEINAME = "Eingangsdaten.xlsx"
PROFIL_START = dt.date(2025, 1, 1)

MONATE = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

JETZT = dt.datetime.now()
HEUTE = JETZT.date()

# Version des festen PV-Startprofils.
# Bei einer neuen Version wird die Startverteilung einmalig neu gesetzt.
PV_STARTPROFIL_VERSION = 1


# =============================================================================
# SESSION STATE
# =============================================================================
defaults = {
    "theme_mode": "🌙 Dunkel",

    "soc_aktuell": 20,
    "soc_ziel": 80,
    "auto_kapazitaet_kwh": 50.0,
    "ladeleistung_auto_kw": 11.0,

    "ansicht": "📆 Heute",
    "gewaehltes_datum": HEUTE,
    "monat_nr": HEUTE.month,

    # Globale Summen als Fallback:
    # Süd gesamt = 2 kWp bei 36° + 2 kWp bei 60° = 4 kWp.
    "kwp_sued": 4.0,
    "kwp_ost": 1.0,
    "kwp_nord": 0.0,
    "kwp_west": 5.0,

    # Feste PV-Startverteilung mit insgesamt 10 kWp:
    # 36°: Süd 2, Ost 1, Nord 0, West 5 kWp
    # 60°: Süd 2, Ost 0, Nord 0, West 0 kWp
    "pv_kwp_config": None,
    "neigungen_wahl": None,
    "pv_startprofil_version": 0,

    "speicher_aktiv": True,
    "speicher_kapazitaet_kwh": 10.0,
    "speicher_start_soc": 50,
    "speicher_max_ladeleistung_kw": 5.0,
    "speicher_max_entladeleistung_kw": 5.0,
    "speicher_wirkungsgrad": 90,

    "lade_modus": "🔌 Direkt",

    # Option B: Auto hat eine Ankunftszeit und eine gewünschte Abfahrtszeit.
    # Falls das Ziel bis zur Abfahrt nicht erreichbar ist, wird weitergerechnet,
    # bis Ziel-SoC erreicht wird.
    "ankunft_datum": HEUTE,
    "zeit_modus": "📅 Heute",
    "ankunftszeit": "18:00",
    "abfahrt_datum": HEUTE + dt.timedelta(days=1),
    "abfahrtszeit": "14:00",

    # alte Keys bleiben nur für Kompatibilität
    "smart_abfahrt_datum": HEUTE + dt.timedelta(days=1),
    "smart_abfahrtszeit": "14:00",

    "sim_result": None,
}

for k, v in defaults.items():
    if k not in st.session_state:
        st.session_state[k] = v


# =============================================================================
# GLOBAL CSS
# =============================================================================
st.markdown(
    """
    <style>
    header[data-testid="stHeader"], footer, div[data-testid="stSidebar"], div[data-testid="stToolbar"] {
        display: none !important;
    }

    html, body, .stApp {
        overflow-x: hidden !important;
        overflow-y: auto !important;
        min-height: 100vh !important;
    }

    .block-container {
        padding-top: 0.25rem !important;
        padding-left: 0.45rem !important;
        padding-right: 0.45rem !important;
        padding-bottom: 1rem !important;
        max-width: 100vw !important;
        min-height: 100vh !important;
        overflow: visible !important;
    }

    div[data-testid="stVerticalBlockBorderWrapper"] {
        border-radius: 22px !important;
        padding: 0.65rem !important;
        overflow: visible !important;
    }

    .app-title {
        font-size: 1.45rem;
        font-weight: 900;
        line-height: 1.15;
        margin-top: 0.1rem;
    }

    .app-subtitle {
        font-size: 0.85rem;
        margin-top: 0.15rem;
    }

    .time-pill {
        text-align: center;
        font-weight: 900;
        font-size: 0.96rem;
        border-radius: 999px;
        padding: 9px 12px;
        white-space: nowrap;
    }

    .stButton > button {
        border-radius: 12px !important;
        font-weight: 850 !important;
    }

    .stButton > button[kind="primary"] {
        background: linear-gradient(135deg, #00d1b2, #16a085) !important;
        border: 1px solid rgba(0,209,178,0.9) !important;
        color: white !important;
    }

    .stButton > button[kind="primary"]:hover {
        background: linear-gradient(135deg, #16a085, #00b894) !important;
        border-color: #00d1b2 !important;
        color: white !important;
    }

    div[data-testid="stPopover"] button {
        border-radius: 999px !important;
        font-weight: 850 !important;
        min-height: 40px !important;
    }

    .stRadio div[role="radiogroup"] {
        display: flex !important;
        flex-direction: column !important;
        gap: 0.35rem !important;
    }

    .stRadio div[role="radiogroup"] label {
        border-radius: 999px !important;
        padding: 7px 10px !important;
        min-height: 36px !important;
        margin: 0 !important;
    }

    div[data-testid="stDateInput"] input,
    div[data-testid="stTextInput"] input,
    div[data-testid="stNumberInput"] input {
        border-radius: 10px !important;
    }

    div[data-testid="stSelectbox"] label,
    div[data-testid="stRadio"] label,
    div[data-testid="stNumberInput"] label,
    div[data-testid="stSlider"] label,
    div[data-testid="stDateInput"] label,
    div[data-testid="stCheckbox"] label {
        font-size: 0.82rem !important;
    }
    </style>
    """,
    unsafe_allow_html=True
)


def apply_theme_css():
    if st.session_state.theme_mode == "🌙 Dunkel":
        st.markdown(
            """
            <style>
            .stApp {
                background: #0b111b !important;
                color: #ffffff !important;
            }

            div[data-testid="stVerticalBlockBorderWrapper"] {
                background: linear-gradient(135deg, rgba(26,33,46,0.96), rgba(13,17,23,0.98)) !important;
                border-color: rgba(255,255,255,0.12) !important;
            }

            .app-title { color: #ffffff !important; }
            .app-subtitle { color: #b8c2d6 !important; }

            .time-pill {
                color: #4A90E2 !important;
                background: rgba(255,255,255,0.075);
                border: 1px solid rgba(255,255,255,0.13);
            }

            div[data-testid="stPopover"] button {
                background: rgba(255,255,255,0.09) !important;
                border: 1px solid rgba(255,255,255,0.18) !important;
                color: white !important;
            }

            div[data-baseweb="select"] > div {
                background-color: rgba(255,255,255,0.10) !important;
                border-color: rgba(255,255,255,0.20) !important;
                color: #ffffff !important;
            }

            div[data-baseweb="select"] span {
                color: #ffffff !important;
            }

            .stRadio div[role="radiogroup"] label {
                background: rgba(255,255,255,0.09) !important;
                border: 1px solid rgba(255,255,255,0.16) !important;
                color: #ffffff !important;
            }

            div[data-testid="stDateInput"] input,
            div[data-testid="stTextInput"] input,
            div[data-testid="stNumberInput"] input {
                background: rgba(255,255,255,0.10) !important;
                color: #ffffff !important;
                border: 1px solid rgba(255,255,255,0.20) !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )
    else:
        st.markdown(
            """
            <style>
            .stApp {
                background: #eef2f7 !important;
                color: #172033 !important;
            }

            div[data-testid="stVerticalBlockBorderWrapper"] {
                background: linear-gradient(135deg, rgba(255,255,255,0.98), rgba(235,240,248,0.99)) !important;
                border-color: rgba(15,23,42,0.10) !important;
            }

            .app-title { color: #172033 !important; }
            .app-subtitle { color: #526173 !important; }

            .time-pill {
                color: #2563eb !important;
                background: rgba(255,255,255,0.88);
                border: 1px solid rgba(15,23,42,0.12);
            }

            div[data-testid="stPopover"] button {
                background: rgba(255,255,255,0.92) !important;
                border: 1px solid rgba(15,23,42,0.15) !important;
                color: #172033 !important;
            }

            .stRadio div[role="radiogroup"] label {
                background: rgba(255,255,255,0.9) !important;
                border: 1px solid rgba(15,23,42,0.12) !important;
                color: #172033 !important;
            }
            </style>
            """,
            unsafe_allow_html=True
        )


apply_theme_css()


# =============================================================================
# HELPERS
# =============================================================================
def klemme_datum_auf_2025(d: dt.date) -> dt.date:
    try:
        return d.replace(year=2025)
    except ValueError:
        return dt.date(2025, d.month, 28)


def slot_von_datum(datum: dt.date, uhrzeit_slot: int = 0) -> int:
    return (datum - PROFIL_START).days * 96 + uhrzeit_slot


def zeitachse_erstellen(start_datum: dt.date, start_uhrzeit_slot: int, n_slots: int):
    start_dt = dt.datetime.combine(start_datum, dt.time()) + dt.timedelta(minutes=15 * start_uhrzeit_slot)
    return [(start_dt + dt.timedelta(minutes=15 * i)).strftime("%d.%m %H:%M") for i in range(n_slots)]


def zeitpunkte_erstellen(start_datum: dt.date, start_uhrzeit_slot: int, n_slots: int):
    start_dt = dt.datetime.combine(start_datum, dt.time()) + dt.timedelta(minutes=15 * start_uhrzeit_slot)
    return [start_dt + dt.timedelta(minutes=15 * i) for i in range(n_slots)]


def slots_fuer_zeitraum(arr, start_slot: int, n_slots: int):
    arr = np.asarray(arr, dtype=float)
    if len(arr) == 0:
        return np.zeros(n_slots)
    return np.array([arr[(start_slot + i) % len(arr)] for i in range(n_slots)], dtype=float)


def ladeleistung_abhaengig_von_soc(soc_prozent: float, max_ladeleistung_kw: float) -> float:
    """
    Vereinfachtes AC-Lademodell mit abfallender Ladeleistung bei hohem SoC.
    """
    soc = max(0.0, min(100.0, float(soc_prozent)))
    pmax = float(max_ladeleistung_kw)

    if soc < 70.0:
        faktor = 1.00
    elif soc < 80.0:
        faktor = 1.00 - ((soc - 70.0) / 10.0) * 0.15
    elif soc < 90.0:
        faktor = 0.85 - ((soc - 80.0) / 10.0) * 0.30
    elif soc < 95.0:
        faktor = 0.55 - ((soc - 90.0) / 5.0) * 0.25
    else:
        faktor = 0.30 - ((soc - 95.0) / 5.0) * 0.20

    faktor = max(0.10, min(1.00, faktor))
    return pmax * faktor


# =============================================================================
# EXCEL
# =============================================================================
def finde_excel_datei():
    home = os.path.expanduser("~")
    pfade = [
        EXCEL_DATEINAME,
        os.path.join(os.getcwd(), EXCEL_DATEINAME),
        os.path.join(home, "Desktop", EXCEL_DATEINAME),
        os.path.join(home, "Documents", EXCEL_DATEINAME),
        os.path.join(home, "Downloads", EXCEL_DATEINAME),
    ]

    for p in pfade:
        if os.path.isfile(p):
            return p
    return None


def parse_excel(xl: pd.ExcelFile):
    sheets = xl.sheet_names

    hs_sheet = next(
        s for s in sheets
        if "haushalt" in s.lower() or "strom" in s.lower()
    )
    df_hs = xl.parse(hs_sheet, header=None)
    hs_kwh = (
        pd.to_numeric(df_hs.iloc[2:, 1], errors="coerce")
        .fillna(0)
        .values
        .astype(float)
        * 0.25 / 1000.0
    )

    waerme_sheet = next(
        (s for s in sheets if "wärme" in s.lower() or "waerme" in s.lower()),
        None
    )

    if waerme_sheet:
        df_w = xl.parse(waerme_sheet, header=None)

        rw = (
            pd.to_numeric(df_w.iloc[4:, 1], errors="coerce")
            .fillna(0)
            .values
            .astype(float)
            * (10140.0 / 1000.0)
        )
        ww = (
            pd.to_numeric(df_w.iloc[4:, 2], errors="coerce")
            .fillna(0)
            .values
            .astype(float)
            * (2433.0 / 1000.0)
        )
        ml = min(len(rw), len(ww))
        waerme_kwh = rw[:ml] + ww[:ml]
    else:
        waerme_kwh = np.zeros(len(hs_kwh))

    ml = min(len(hs_kwh), len(waerme_kwh))
    hs_kwh = hs_kwh[:ml]
    waerme_kwh = waerme_kwh[:ml]

    preis_sheet = next(
        s for s in sheets
        if "preis" in s.lower() or "börse" in s.lower() or "boerse" in s.lower()
    )
    df_p = xl.parse(preis_sheet, header=None)

    boerse = (
        pd.to_numeric(df_p.iloc[1:, 1], errors="coerce")
        .ffill()
        .bfill()
        .fillna(0)
        .values
        .astype(float)
    )

    preise_berechnet = (boerse + 21.0) * 1.19
    preisquelle = "Berechnet aus Excel: (Börsenstrompreis + 21 ct/kWh) × 1,19"

    if df_p.shape[1] >= 3:
        spalte_c = (
            pd.to_numeric(df_p.iloc[1:, 2], errors="coerce")
            .ffill()
            .bfill()
            .fillna(0)
            .values
            .astype(float)
        )
        if len(spalte_c) and np.nanmedian(spalte_c) < 2.0:
            spalte_c = spalte_c * 100.0

        if len(spalte_c) and np.nanmedian(spalte_c) > 20.0:
            preise = spalte_c
            preisquelle = "Excel-Spalte C: Strom-Kaufpreis / Arbeitspreis gesamt"
        else:
            preise = preise_berechnet
    else:
        preise = preise_berechnet

    pv_dict = {}

    for sheet in sheets:
        if "pv" not in sheet.lower() and "neigung" not in sheet.lower():
            continue

        df_pv = xl.parse(sheet, header=None)

        for col_idx in range(1, min(5, df_pv.shape[1])):
            try:
                neigung = str(df_pv.iloc[2, col_idx]).strip()
                richtung = str(df_pv.iloc[3, col_idx]).strip()

                if neigung.lower() == "nan" or richtung.lower() == "nan":
                    continue

                jahres_kwh_kwp = float(df_pv.iloc[0, col_idx])
                kurve_raw = (
                    pd.to_numeric(df_pv.iloc[5:, col_idx], errors="coerce")
                    .fillna(0)
                    .values
                    .astype(float)
                )
                kurve_kwh_kwp = kurve_raw * 0.25 / 1000.0

                pv_dict[(neigung, richtung)] = {
                    "kurve": kurve_kwh_kwp,
                    "jahres_kwh": jahres_kwh_kwp,
                }

            except Exception:
                continue

    if not pv_dict:
        raise ValueError("Keine PV-Daten in der Excel-Datei gefunden.")

    return {
        "hs": np.asarray(hs_kwh, dtype=float),
        "waerme": np.asarray(waerme_kwh, dtype=float),
        "preise": np.asarray(preise, dtype=float),
        "pv": pv_dict,
        "preisquelle": preisquelle,
    }


@st.cache_data(show_spinner=False)
def lade_excel_cached(pfad: str):
    xl = pd.ExcelFile(pfad, engine="openpyxl")
    return parse_excel(xl)


def lade_excel():
    pfad = finde_excel_datei()

    if not pfad:
        st.error(f"❌ Datei `{EXCEL_DATEINAME}` wurde nicht gefunden.")
        st.info("Bitte lege `Eingangsdaten.xlsx` in den Projektordner, Desktop, Documents oder Downloads.")
        st.stop()

    try:
        daten_local = lade_excel_cached(pfad)
        return daten_local, pfad
    except Exception as e:
        st.error(f"❌ Fehler beim Laden der Excel-Datei: {e}")
        st.stop()


daten, excel_pfad = lade_excel()

alle_neigungen = sorted(set(k[0] for k in daten["pv"].keys()), key=str)
richtungen = ["Süd", "Ost", "Nord", "West"]


def finde_neigung_mit_grad(grad: int):
    """Findet robust den Excel-Namen, z. B. '36°' oder '36° Neigung'."""
    grad_text = str(grad)
    return next(
        (
            neigung
            for neigung in alle_neigungen
            if str(neigung).strip().startswith(grad_text)
        ),
        None,
    )


neigung_36 = finde_neigung_mit_grad(36)
neigung_60 = finde_neigung_mit_grad(60)

# Gewünschte feste Startverteilung: insgesamt genau 10 kWp.
standard_pv_kwp_config = {
    neigung: {richtung: 0.0 for richtung in richtungen}
    for neigung in alle_neigungen
}

if neigung_36 is not None:
    standard_pv_kwp_config[neigung_36] = {
        "Süd": 2.0,
        "Ost": 1.0,
        "Nord": 0.0,
        "West": 5.0,
    }

if neigung_60 is not None:
    standard_pv_kwp_config[neigung_60] = {
        "Süd": 2.0,
        "Ost": 0.0,
        "Nord": 0.0,
        "West": 0.0,
    }

standard_neigungen = [
    neigung
    for neigung in (neigung_36, neigung_60)
    if neigung is not None
]

# Fallback, falls die Excel-Datei andere Bezeichnungen enthält.
if not standard_neigungen:
    standard_neigungen = list(alle_neigungen)

startprofil_neu_setzen = (
    st.session_state.get("pv_startprofil_version", 0)
    != PV_STARTPROFIL_VERSION
)

if startprofil_neu_setzen or st.session_state.pv_kwp_config is None:
    # Alte Werte der kWp-Eingabefelder entfernen, damit die neuen Startwerte
    # auch nach einem Streamlit-Code-Rerun sofort sichtbar werden.
    for widget_key in list(st.session_state.keys()):
        if widget_key.startswith(
            ("kwp_sued_", "kwp_ost_", "kwp_nord_", "kwp_west_")
        ):
            del st.session_state[widget_key]

    st.session_state.neigungen_wahl = list(standard_neigungen)
    st.session_state.pv_kwp_config = {
        neigung: dict(config)
        for neigung, config in standard_pv_kwp_config.items()
    }

    # Globale Summen als Fallback/Info synchronisieren.
    st.session_state.kwp_sued = 4.0
    st.session_state.kwp_ost = 1.0
    st.session_state.kwp_nord = 0.0
    st.session_state.kwp_west = 5.0
    st.session_state.pv_startprofil_version = PV_STARTPROFIL_VERSION

# Nur gültige und aktive Neigungen behalten.
st.session_state.neigungen_wahl = [
    neigung
    for neigung in (st.session_state.neigungen_wahl or [])
    if neigung in alle_neigungen
] or list(standard_neigungen)

# Falls Excel/Session sich geändert hat: fehlende Neigungen/Richtungen ergänzen.
for neigung in alle_neigungen:
    if neigung not in st.session_state.pv_kwp_config:
        st.session_state.pv_kwp_config[neigung] = {
            richtung: 0.0 for richtung in richtungen
        }
    for richtung in richtungen:
        st.session_state.pv_kwp_config[neigung].setdefault(richtung, 0.0)

# Ungültige Neigungen aus einer älteren Session entfernen.
st.session_state.pv_kwp_config = {
    neigung: config
    for neigung, config in st.session_state.pv_kwp_config.items()
    if neigung in alle_neigungen
}


# =============================================================================
# SVG CHARTS
# =============================================================================
def svg_line(werte, zeit_labels=None, farbe="#00d1b2", keep_above_zero=False, n_ticks=6):
    arr_real = np.asarray(werte, dtype=float)
    arr_real = np.nan_to_num(arr_real, nan=0.0, posinf=0.0, neginf=0.0)

    if len(arr_real) == 0:
        arr_real = np.zeros(2)

    n = len(arr_real)

    if zeit_labels is None or len(zeit_labels) != n:
        zeit_labels = ["" for _ in range(n)]

    arr = arr_real.copy()

    if keep_above_zero and np.max(arr) > 1e-9:
        arr = np.maximum(arr, np.max(arr) * 0.06)

    w, h = 760, 125
    pl, pr, pt, pb = 38, 12, 14, 26
    pw = w - pl - pr
    ph = h - pt - pb

    mn = min(0.0, float(np.min(arr)))
    mx = float(np.max(arr))
    if mx <= 1e-9:
        mx = 1.0

    span = max(1e-6, mx - mn)
    bottom_gap = 8.0 if keep_above_zero else 0.0

    points = []
    for i, v in enumerate(arr):
        x = pl + i / max(1, n - 1) * pw
        y = pt + ph - bottom_gap - ((v - mn) / span) * (ph - bottom_gap)
        points.append(f"{x:.1f},{y:.1f}")

    grid = ""
    for i in range(5):
        y = pt + ph * i / 4
        grid += (
            f'<line x1="{pl}" y1="{y:.1f}" x2="{pl+pw}" y2="{y:.1f}" '
            f'stroke="rgba(255,255,255,.12)" />'
        )

    ticks = ""
    ticks_count = min(n_ticks, n)

    for k in range(ticks_count):
        idx = int(round(k * (n - 1) / max(1, ticks_count - 1)))
        x = pl + idx / max(1, n - 1) * pw
        ticks += (
            f'<text x="{x:.1f}" y="{h-7}" fill="#aebbd3" font-size="10" '
            f'text-anchor="middle">{zeit_labels[idx]}</text>'
        )

    return f"""
    <svg viewBox="0 0 {w} {h}" preserveAspectRatio="none" style="width:100%;height:{h}px;display:block;">
        {grid}
        <polyline points="{' '.join(points)}" fill="none" stroke="{farbe}" stroke-width="2.7"
        stroke-linejoin="round" stroke-linecap="round"/>
        <text x="3" y="{pt+5}" fill="#aebbd3" font-size="10">{np.max(arr_real):.1f}</text>
        <text x="3" y="{pt+ph:.1f}" fill="#aebbd3" font-size="10">{np.min(arr_real):.1f}</text>
        {ticks}
    </svg>
    """


def svg_multi(series, zeit_labels=None):
    farben = {
        "PV → Auto": "#2ecc71",
        "Netz → Auto": "#6aa8ff",
    }

    prepared = {}
    n = 0

    for name, values in series.items():
        arr = np.asarray(values, dtype=float)
        arr = np.nan_to_num(arr, nan=0.0, posinf=0.0, neginf=0.0)
        prepared[name] = arr
        n = max(n, len(arr))

    if n == 0:
        n = 2
        prepared = {"PV → Auto": np.zeros(2), "Netz → Auto": np.zeros(2)}

    if zeit_labels is None or len(zeit_labels) != n:
        zeit_labels = ["" for _ in range(n)]

    all_vals = np.concatenate([v for v in prepared.values()])
    mx = float(np.max(all_vals)) if len(all_vals) else 1.0
    mn = min(0.0, float(np.min(all_vals))) if len(all_vals) else 0.0

    if mx <= 1e-9:
        mx = 1.0

    span = max(1e-6, mx - mn)

    w, h = 760, 135
    pl, pr, pt, pb = 38, 12, 26, 28
    pw = w - pl - pr
    ph = h - pt - pb

    grid = ""
    for i in range(5):
        y = pt + ph * i / 4
        grid += (
            f'<line x1="{pl}" y1="{y:.1f}" x2="{pl+pw}" y2="{y:.1f}" '
            f'stroke="rgba(255,255,255,.12)" />'
        )

    lines = ""
    for name, arr in prepared.items():
        pts = []
        for i, v in enumerate(arr):
            x = pl + i / max(1, n - 1) * pw
            y = pt + ph - ((v - mn) / span) * ph
            pts.append(f"{x:.1f},{y:.1f}")

        lines += (
            f'<polyline points="{" ".join(pts)}" fill="none" '
            f'stroke="{farben.get(name, "#fff")}" stroke-width="2.8" '
            f'stroke-linejoin="round" stroke-linecap="round"/>'
        )

    legend = ""
    lx = pl
    for name, color in farben.items():
        legend += (
            f'<circle cx="{lx}" cy="13" r="4" fill="{color}"/>'
            f'<text x="{lx+9}" y="17" fill="#dce7ff" font-size="11">{name}</text>'
        )
        lx += 125

    ticks = ""
    ticks_count = min(6, n)

    for k in range(ticks_count):
        idx = int(round(k * (n - 1) / max(1, ticks_count - 1)))
        x = pl + idx / max(1, n - 1) * pw
        ticks += (
            f'<text x="{x:.1f}" y="{h-7}" fill="#aebbd3" font-size="10" '
            f'text-anchor="middle">{zeit_labels[idx]}</text>'
        )

    return f"""
    <svg viewBox="0 0 {w} {h}" preserveAspectRatio="none" style="width:100%;height:{h}px;display:block;">
        {grid}
        {lines}
        {legend}
        <text x="3" y="{pt+5}" fill="#aebbd3" font-size="10">{mx:.2f}</text>
        <text x="3" y="{pt+ph:.1f}" fill="#aebbd3" font-size="10">{mn:.2f}</text>
        {ticks}
    </svg>
    """


# =============================================================================
# BATTERIESPEICHER
# =============================================================================
def simuliere_batteriespeicher(
    pv_kwh,
    verbrauch_kwh,
    kapazitaet_kwh,
    start_soc_prozent,
    max_ladeleistung_kw,
    max_entladeleistung_kw,
    wirkungsgrad_prozent,
):
    n = len(pv_kwh)

    soc = np.zeros(n)
    laden = np.zeros(n)
    entladen = np.zeros(n)
    netz = np.zeros(n)
    einspeisung = np.zeros(n)

    if kapazitaet_kwh <= 0:
        netz = np.maximum(verbrauch_kwh - pv_kwh, 0)
        einspeisung = np.maximum(pv_kwh - verbrauch_kwh, 0)
        return soc, laden, entladen, netz, einspeisung

    eta = max(0.01, wirkungsgrad_prozent / 100.0)
    energie = kapazitaet_kwh * start_soc_prozent / 100.0

    max_lade_slot = max_ladeleistung_kw * 0.25
    max_entlade_slot = max_entladeleistung_kw * 0.25

    for i in range(n):
        saldo = pv_kwh[i] - verbrauch_kwh[i]

        if saldo > 0:
            frei = kapazitaet_kwh - energie
            e_laden = min(saldo, max_lade_slot, frei / eta)
            energie += e_laden * eta
            laden[i] = e_laden
            einspeisung[i] = max(0.0, saldo - e_laden)

        elif saldo < 0:
            bedarf = abs(saldo)
            e_entladen = min(bedarf, max_entlade_slot, energie * eta)
            energie -= e_entladen / eta
            entladen[i] = e_entladen
            netz[i] = max(0.0, bedarf - e_entladen)

        energie = max(0.0, min(kapazitaet_kwh, energie))
        soc[i] = energie / kapazitaet_kwh * 100.0

    return soc, laden, entladen, netz, einspeisung


# =============================================================================
# TOP APP UI: EIN EINZIGER SICHTBARER RAHMEN
# =============================================================================
with st.container(border=True):
    # Erste Zeile: Titel links, Zeit / Theme / Settings rechts
    top_left, top_time, top_theme, top_settings = st.columns([5.8, 1.55, 0.42, 0.42])

    with top_left:
        st.markdown(
            """
            <div class="app-title">🔄 Energiefluss im System</div>
            <div class="app-subtitle">Klick auf PV, Akku, Haus, E-Auto oder Netz.</div>
            """,
            unsafe_allow_html=True,
        )

    with top_time:
        st.markdown(
            f"""
            <div class="time-pill">
                🕒 {JETZT.strftime('%d.%m.%Y')} · {JETZT.strftime('%H:%M')}
            </div>
            """,
            unsafe_allow_html=True,
        )

    with top_theme:
        theme_icon = "🌙" if st.session_state.theme_mode == "🌙 Dunkel" else "☀️"

        if st.button(
            theme_icon,
            width="stretch",
            help="Hell/Dunkel wechseln",
            key="theme_icon_button"
        ):
            st.session_state.theme_mode = (
                "☀️ Hell"
                if st.session_state.theme_mode == "🌙 Dunkel"
                else "🌙 Dunkel"
            )
            st.session_state.sim_result = None
            st.rerun()

    with top_settings:
        with st.popover("⚙️", width="stretch"):
            with st.form("settings_form"):
                st.markdown("### 🚗 Fahrzeug")

                soc_aktuell_new = st.slider(
                    "Aktueller SoC Auto (%)",
                    0,
                    100,
                    int(st.session_state.soc_aktuell),
                )

                soc_ziel_new = st.slider(
                    "Ziel-SoC Auto (%)",
                    0,
                    100,
                    int(st.session_state.soc_ziel),
                )

                auto_kapazitaet_new = st.number_input(
                    "Auto-Batteriekapazität (kWh)",
                    10.0,
                    150.0,
                    float(st.session_state.auto_kapazitaet_kwh),
                    5.0,
                )

                ladeleistung_new = st.number_input(
                    "AC-Ladeleistung Auto (kW)",
                    1.0,
                    22.0,
                    float(st.session_state.ladeleistung_auto_kw),
                    1.0,
                )

                # Zeitraum wird jetzt direkt links in der Hauptansicht gewählt.
                ansicht_new = st.session_state.ansicht
                gewaehltes_datum_new = st.session_state.gewaehltes_datum
                monat_nr_new = st.session_state.monat_nr

                st.markdown("---")
                st.markdown("### ☀️ PV")

                st.caption("Jede Dachneigung hat eigene kWp-Werte pro Richtung.")

                neigungen_new = st.multiselect(
                    "Aktive Dachneigungen",
                    alle_neigungen,
                    default=st.session_state.neigungen_wahl,
                    help="Wähle z. B. 36° und 60°. Jede Neigung hat unten eigene kWp-Werte.",
                )

                if not neigungen_new:
                    st.warning("Bitte mindestens eine Dachneigung auswählen.")
                    neigungen_new = st.session_state.neigungen_wahl or list(alle_neigungen)

                pv_kwp_config_new = {}

                for neigung in alle_neigungen:
                    with st.expander(f"☀️ PV {neigung}", expanded=(neigung in neigungen_new)):
                        active_neigung = neigung in neigungen_new

                        cfg_old = st.session_state.pv_kwp_config.get(
                            neigung,
                            {"Süd": 0.0, "Ost": 0.0, "Nord": 0.0, "West": 0.0},
                        )

                        c_sued, c_ost = st.columns(2)
                        c_nord, c_west = st.columns(2)

                        with c_sued:
                            kwp_sued_val = st.number_input(
                                f"kWp Süd ({neigung})",
                                0.0,
                                100.0,
                                float(cfg_old.get("Süd", 0.0)),
                                0.5,
                                key=f"kwp_sued_{neigung}",
                                disabled=not active_neigung,
                            )

                        with c_ost:
                            kwp_ost_val = st.number_input(
                                f"kWp Ost ({neigung})",
                                0.0,
                                100.0,
                                float(cfg_old.get("Ost", 0.0)),
                                0.5,
                                key=f"kwp_ost_{neigung}",
                                disabled=not active_neigung,
                            )

                        with c_nord:
                            kwp_nord_val = st.number_input(
                                f"kWp Nord ({neigung})",
                                0.0,
                                100.0,
                                float(cfg_old.get("Nord", 0.0)),
                                0.5,
                                key=f"kwp_nord_{neigung}",
                                disabled=not active_neigung,
                            )

                        with c_west:
                            kwp_west_val = st.number_input(
                                f"kWp West ({neigung})",
                                0.0,
                                100.0,
                                float(cfg_old.get("West", 0.0)),
                                0.5,
                                key=f"kwp_west_{neigung}",
                                disabled=not active_neigung,
                            )

                        pv_kwp_config_new[neigung] = {
                            "Süd": kwp_sued_val if active_neigung else 0.0,
                            "Ost": kwp_ost_val if active_neigung else 0.0,
                            "Nord": kwp_nord_val if active_neigung else 0.0,
                            "West": kwp_west_val if active_neigung else 0.0,
                        }

                # Alte globale Werte werden weiterhin gesetzt, damit alte Anzeigen/Fallbacks nicht brechen.
                kwp_sued_new = sum(cfg["Süd"] for cfg in pv_kwp_config_new.values())
                kwp_ost_new = sum(cfg["Ost"] for cfg in pv_kwp_config_new.values())
                kwp_nord_new = sum(cfg["Nord"] for cfg in pv_kwp_config_new.values())
                kwp_west_new = sum(cfg["West"] for cfg in pv_kwp_config_new.values())

                st.markdown("---")
                st.markdown("### 🔋 Hausspeicher")

                speicher_aktiv_new = st.checkbox(
                    "Hausspeicher aktiv",
                    value=bool(st.session_state.speicher_aktiv),
                )

                speicher_optionen = [0.0, 5.0, 10.0, 15.0, 20.0, 50.0]

                speicher_kapazitaet_new = st.selectbox(
                    "Speicherkapazität (kWh)",
                    speicher_optionen,
                    index=speicher_optionen.index(float(st.session_state.speicher_kapazitaet_kwh))
                    if float(st.session_state.speicher_kapazitaet_kwh) in speicher_optionen
                    else 2,
                )

                speicher_start_soc_new = st.slider(
                    "Start-SoC Speicher (%)",
                    0,
                    100,
                    int(st.session_state.speicher_start_soc),
                )

                speicher_max_ladeleistung_new = st.number_input(
                    "Max. Ladeleistung Speicher (kW)",
                    0.0,
                    20.0,
                    float(st.session_state.speicher_max_ladeleistung_kw),
                    0.5,
                )

                speicher_max_entladeleistung_new = st.number_input(
                    "Max. Entladeleistung Speicher (kW)",
                    0.0,
                    20.0,
                    float(st.session_state.speicher_max_entladeleistung_kw),
                    0.5,
                )

                speicher_wirkungsgrad_new = st.slider(
                    "Speicher-Wirkungsgrad (%)",
                    50,
                    100,
                    int(st.session_state.speicher_wirkungsgrad),
                )

                submitted = st.form_submit_button("✅ Speichern")

                if submitted:
                    st.session_state.soc_aktuell = soc_aktuell_new
                    st.session_state.soc_ziel = soc_ziel_new
                    st.session_state.auto_kapazitaet_kwh = auto_kapazitaet_new
                    st.session_state.ladeleistung_auto_kw = ladeleistung_new

                    st.session_state.ansicht = ansicht_new
                    st.session_state.gewaehltes_datum = gewaehltes_datum_new
                    st.session_state.monat_nr = monat_nr_new

                    # Wenn ein Tag in den Einstellungen gewählt wird, übernimmt die linke Zeitsteuerung diesen Tag.
                    if ansicht_new == "🗓️ Tag wählen":
                        st.session_state.zeit_modus = "🗓️ Beliebiger Tag"
                        st.session_state.ankunft_datum = gewaehltes_datum_new
                    elif ansicht_new == "📆 Heute":
                        st.session_state.zeit_modus = "📅 Heute"
                        st.session_state.ankunft_datum = HEUTE

                    st.session_state.neigungen_wahl = list(neigungen_new)
                    st.session_state.pv_kwp_config = pv_kwp_config_new

                    # Summen als Fallback/Info speichern.
                    st.session_state.kwp_sued = kwp_sued_new
                    st.session_state.kwp_ost = kwp_ost_new
                    st.session_state.kwp_nord = kwp_nord_new
                    st.session_state.kwp_west = kwp_west_new

                    st.session_state.speicher_aktiv = speicher_aktiv_new
                    st.session_state.speicher_kapazitaet_kwh = speicher_kapazitaet_new
                    st.session_state.speicher_start_soc = speicher_start_soc_new
                    st.session_state.speicher_max_ladeleistung_kw = speicher_max_ladeleistung_new
                    st.session_state.speicher_max_entladeleistung_kw = speicher_max_entladeleistung_new
                    st.session_state.speicher_wirkungsgrad = speicher_wirkungsgrad_new

                    st.session_state.sim_result = None

    # Zweite Zeile: Lademethode und Zeitsteuerung direkt links in der Hauptansicht
    mode_col, time_col, spacer_col = st.columns([0.95, 3.15, 3.5])

    alle_zeiten = [
        f"{h:02d}:{m:02d}"
        for h in range(24)
        for m in (0, 15, 30, 45)
    ]

    with mode_col:
        selected_mode = st.radio(
            "Lademethode",
            ["🔌 Direkt", "☀️ PV", "🧠 Smart"],
            index=["🔌 Direkt", "☀️ PV", "🧠 Smart"].index(st.session_state.lade_modus)
            if st.session_state.lade_modus in ["🔌 Direkt", "☀️ PV", "🧠 Smart"]
            else 0,
            label_visibility="collapsed",
            horizontal=False,
            key="lade_modus_radio",
        )

        if selected_mode != st.session_state.lade_modus:
            st.session_state.lade_modus = selected_mode
            st.session_state.sim_result = None

        start_button = st.button(
            "⚡ Laden",
            type="primary",
            width="stretch"
        )

    with time_col:
        zeit_modus_new = st.radio(
            "Zeitmodus",
            ["📅 Heute", "🗓️ Beliebiger Tag"],
            index=["📅 Heute", "🗓️ Beliebiger Tag"].index(st.session_state.zeit_modus)
            if st.session_state.zeit_modus in ["📅 Heute", "🗓️ Beliebiger Tag"]
            else 0,
            horizontal=True,
            label_visibility="collapsed",
            key="zeit_modus_radio",
        )

        if zeit_modus_new == "📅 Heute":
            ankunft_datum_new = HEUTE
            # Bei Heute soll die Ankunft automatisch jetzt sein.
            aktuelle_minute = (JETZT.minute // 15) * 15
            ankunftszeit_new = f"{JETZT.hour:02d}:{aktuelle_minute:02d}"

            if selected_mode == "🧠 Smart":
                # Der Nutzer wählt das Abfahrtsdatum und die Abfahrtszeit selbst.
                # Standardmäßig wird der nächste Tag verwendet.
                abfahrt_standard = st.session_state.abfahrt_datum
                if abfahrt_standard < HEUTE:
                    abfahrt_standard = HEUTE + dt.timedelta(days=1)

                datum_col, zeit_col = st.columns(2)

                with datum_col:
                    abfahrt_datum_new = st.date_input(
                        "📅 Abfahrtsdatum",
                        value=abfahrt_standard,
                        min_value=HEUTE,
                        key="abfahrtsdatum_heute_main",
                    )

                with zeit_col:
                    abfahrtszeit_new = st.selectbox(
                        "🏁 Abfahrtszeit",
                        alle_zeiten,
                        index=alle_zeiten.index(st.session_state.abfahrtszeit)
                        if st.session_state.abfahrtszeit in alle_zeiten
                        else alle_zeiten.index("14:00"),
                        key="abfahrtszeit_heute_main",
                    )

                st.caption(
                    f"🚗 Ankunft: heute {ankunftszeit_new} · "
                    f"🏁 Abfahrt: {abfahrt_datum_new.strftime('%d.%m.%Y')} "
                    f"{abfahrtszeit_new}"
                )
            else:
                abfahrt_datum_new = HEUTE + dt.timedelta(days=1)
                abfahrtszeit_new = "14:00"
                st.caption(f"🚗 Ankunft: heute {ankunftszeit_new}")

        else:
            c1, c2 = st.columns(2)

            with c1:
                ankunft_datum_new = st.date_input(
                    "📅 Ankunftstag",
                    value=st.session_state.ankunft_datum,
                    key="ankunftstag_main",
                )

                ankunftszeit_new = st.selectbox(
                    "🚗 Ankunftszeit",
                    alle_zeiten,
                    index=alle_zeiten.index(st.session_state.ankunftszeit)
                    if st.session_state.ankunftszeit in alle_zeiten
                    else alle_zeiten.index("18:00"),
                    key="ankunftszeit_main",
                )

            if selected_mode == "🧠 Smart":
                with c2:
                    abfahrt_datum_new = st.date_input(
                        "📅 Abfahrtstag",
                        value=st.session_state.abfahrt_datum
                        if st.session_state.abfahrt_datum >= ankunft_datum_new
                        else ankunft_datum_new + dt.timedelta(days=1),
                        key="abfahrtstag_main",
                    )

                    abfahrtszeit_new = st.selectbox(
                        "🏁 Abfahrtszeit",
                        alle_zeiten,
                        index=alle_zeiten.index(st.session_state.abfahrtszeit)
                        if st.session_state.abfahrtszeit in alle_zeiten
                        else alle_zeiten.index("14:00"),
                        key="abfahrtszeit_main",
                    )
            else:
                abfahrt_datum_new = ankunft_datum_new + dt.timedelta(days=1)
                abfahrtszeit_new = "14:00"

        # Prüfen, ob die gewählte Abfahrt wirklich nach der Ankunft liegt.
        if selected_mode == "🧠 Smart":
            ank_h_ui, ank_m_ui = map(int, ankunftszeit_new.split(":"))
            abf_h_ui, abf_m_ui = map(int, abfahrtszeit_new.split(":"))

            ankunft_test_ui = dt.datetime.combine(
                ankunft_datum_new,
                dt.time(hour=ank_h_ui, minute=ank_m_ui),
            )
            abfahrt_test_ui = dt.datetime.combine(
                abfahrt_datum_new,
                dt.time(hour=abf_h_ui, minute=abf_m_ui),
            )

            if abfahrt_test_ui <= ankunft_test_ui:
                st.error(
                    "❌ Das Abfahrtsdatum und die Abfahrtszeit müssen "
                    "nach der Ankunft liegen."
                )

        # Session-State synchronisieren
        if (
            zeit_modus_new != st.session_state.zeit_modus
            or ankunft_datum_new != st.session_state.ankunft_datum
            or ankunftszeit_new != st.session_state.ankunftszeit
            or abfahrt_datum_new != st.session_state.abfahrt_datum
            or abfahrtszeit_new != st.session_state.abfahrtszeit
        ):
            st.session_state.zeit_modus = zeit_modus_new
            st.session_state.ankunft_datum = ankunft_datum_new
            st.session_state.ankunftszeit = ankunftszeit_new
            st.session_state.abfahrt_datum = abfahrt_datum_new
            st.session_state.abfahrtszeit = abfahrtszeit_new

            # Excel-Analyse folgt dem Ankunftstag, damit Winter/Sommer-Tage direkt getestet werden können.
            if zeit_modus_new == "📅 Heute":
                st.session_state.ansicht = "📆 Heute"
                st.session_state.gewaehltes_datum = HEUTE
            else:
                st.session_state.ansicht = "🗓️ Tag wählen"
                st.session_state.gewaehltes_datum = ankunft_datum_new

            st.session_state.smart_abfahrt_datum = abfahrt_datum_new
            st.session_state.smart_abfahrtszeit = abfahrtszeit_new
            st.session_state.sim_result = None
            st.rerun()

    energiefluss_placeholder = st.empty()


# =============================================================================
# DATENZEITRAUM
# =============================================================================
if st.session_state.ansicht == "📆 Heute":
    uhrzeit_slot = (JETZT.hour * 60 + JETZT.minute) // 15
    profil_heute = klemme_datum_auf_2025(HEUTE)

    start_slot = slot_von_datum(profil_heute, uhrzeit_slot) - 288
    n_slots = 288 + 96 + 96 * 60  # 3 Tage davor + 1 Tag sichtbar + 60 Tage zum Weiterladen

    start_datum_real = HEUTE - dt.timedelta(days=3)

    labels = zeitachse_erstellen(start_datum_real, uhrzeit_slot, n_slots)
    zeitpunkte_dt = zeitpunkte_erstellen(start_datum_real, uhrzeit_slot, n_slots)

    fokus_start = 288
    fokus_ende = 288 + 95

    ansicht_titel = f"Heute – {HEUTE.strftime('%d.%m.%Y')}"

elif st.session_state.ansicht == "🗓️ Tag wählen":
    real_date = st.session_state.gewaehltes_datum
    profil_date = klemme_datum_auf_2025(real_date)

    start_slot = slot_von_datum(profil_date, 0) - 288
    n_slots = 288 + 96 + 96 * 60  # 3 Tage davor + 1 Tag sichtbar + 60 Tage zum Weiterladen

    start_datum_real = real_date - dt.timedelta(days=3)

    labels = zeitachse_erstellen(start_datum_real, 0, n_slots)
    zeitpunkte_dt = zeitpunkte_erstellen(start_datum_real, 0, n_slots)

    fokus_start = 288
    fokus_ende = 288 + 95

    ansicht_titel = f"Tag – {real_date.strftime('%d.%m.%Y')}"

else:
    monat_nr = int(st.session_state.monat_nr)
    monat_start = dt.date(2025, monat_nr, 1)

    if monat_nr == 12:
        monat_ende = dt.date(2025, 12, 31)
    else:
        monat_ende = dt.date(2025, monat_nr + 1, 1) - dt.timedelta(days=1)

    n_slots = ((monat_ende - monat_start).days + 1) * 96
    start_slot = slot_von_datum(monat_start, 0)

    start_datum_real = monat_start

    labels = zeitachse_erstellen(monat_start, 0, n_slots)
    zeitpunkte_dt = zeitpunkte_erstellen(monat_start, 0, n_slots)

    fokus_start = 0
    fokus_ende = n_slots - 1

    ansicht_titel = f"{MONATE[monat_nr]} 2025"

# =============================================================================
# LADEFENSTER: ANKUNFT BIS GEWÜNSCHTE ABFAHRT
# =============================================================================
def _parse_hhmm(text_zeit: str):
    h, m = map(int, text_zeit.split(":"))
    return h, m

ank_h, ank_m = _parse_hhmm(st.session_state.ankunftszeit)
abf_h, abf_m = _parse_hhmm(st.session_state.abfahrtszeit)

ankunft_dt = dt.datetime.combine(
    st.session_state.ankunft_datum,
    dt.time(hour=ank_h, minute=ank_m)
)
abfahrt_dt = dt.datetime.combine(
    st.session_state.abfahrt_datum,
    dt.time(hour=abf_h, minute=abf_m)
)

# Keine automatische Änderung des gewählten Abfahrtsdatums:
# Eine ungültige Abfahrt wird beim Start der Smart-Simulation gemeldet.
abfahrt_ungueltig = abfahrt_dt <= ankunft_dt

def index_fuer_datetime(ziel_dt: dt.datetime) -> int:
    raw = int(round((ziel_dt - zeitpunkte_dt[0]).total_seconds() / (15 * 60)))
    return int(min(max(raw, 0), n_slots - 1))

lade_start_slot = index_fuer_datetime(ankunft_dt)
abfahrt_slot = index_fuer_datetime(abfahrt_dt)

ankunft_text = ankunft_dt.strftime("%d.%m %H:%M")
abfahrt_text = abfahrt_dt.strftime("%d.%m %H:%M")

hs_z = slots_fuer_zeitraum(daten["hs"], start_slot, n_slots)
waerme_z = slots_fuer_zeitraum(daten["waerme"], start_slot, n_slots)
preise_z = slots_fuer_zeitraum(daten["preise"], start_slot, n_slots)

verbrauch_z = hs_z + waerme_z


# =============================================================================
# PV
# =============================================================================
kwp_map = {
    "Süd": float(st.session_state.kwp_sued),
    "Ost": float(st.session_state.kwp_ost),
    "Nord": float(st.session_state.kwp_nord),
    "West": float(st.session_state.kwp_west),
}

neigungen_wahl = st.session_state.neigungen_wahl or list(alle_neigungen)
pv_richtungen = {}

# Jede Kombination aus Dachneigung und Richtung hat jetzt eigene kWp.
for (neigung, richtung), info in daten["pv"].items():
    if neigung not in neigungen_wahl:
        continue

    cfg = st.session_state.pv_kwp_config.get(neigung, {})
    kwp = float(cfg.get(richtung, 0.0))

    if kwp > 0:
        name = f"{neigung} · {richtung}"
        pv_richtungen[name] = slots_fuer_zeitraum(info["kurve"], start_slot, n_slots) * kwp

if pv_richtungen:
    pv_z = np.sum(np.vstack(list(pv_richtungen.values())), axis=0)
else:
    pv_z = np.zeros(n_slots)


# =============================================================================
# HAUSSPEICHER
# =============================================================================
if st.session_state.speicher_aktiv:
    (
        speicher_soc,
        speicher_laden,
        speicher_entladen,
        haus_netzbezug,
        einspeisung_vor_auto,
    ) = simuliere_batteriespeicher(
        pv_kwh=pv_z,
        verbrauch_kwh=verbrauch_z,
        kapazitaet_kwh=float(st.session_state.speicher_kapazitaet_kwh),
        start_soc_prozent=float(st.session_state.speicher_start_soc),
        max_ladeleistung_kw=float(st.session_state.speicher_max_ladeleistung_kw),
        max_entladeleistung_kw=float(st.session_state.speicher_max_entladeleistung_kw),
        wirkungsgrad_prozent=float(st.session_state.speicher_wirkungsgrad),
    )
else:
    speicher_soc = np.zeros(n_slots)
    speicher_laden = np.zeros(n_slots)
    speicher_entladen = np.zeros(n_slots)
    haus_netzbezug = np.maximum(verbrauch_z - pv_z, 0)
    einspeisung_vor_auto = np.maximum(pv_z - verbrauch_z, 0)


# =============================================================================
# SIMULATION
# =============================================================================
def empty_result():
    return {
        "mode_key": "none",
        "strategie": "Noch keine Ladesimulation berechnet – E-Auto noch nicht verbunden",
        "auto_pv": np.zeros(n_slots),
        "auto_netz": np.zeros(n_slots),
        "auto_total": np.zeros(n_slots),
        "auto_soc": np.full(n_slots, float(st.session_state.soc_aktuell)),
        "auto_pv_kwh": 0.0,
        "auto_netz_kwh": 0.0,
        "auto_total_kwh": 0.0,
        "kosten": 0.0,
        "avg_ct": 0.0,
        "ziel_text": "—",
        "voll_text": "—",
        "ladezeit_text": "—",
        "final_soc_text": f"{float(st.session_state.soc_aktuell):.1f} %",
        "dauer_bis_ziel_text": "—",
    }


def result_pack(
    mode_key,
    strategie,
    auto_pv,
    auto_netz,
    auto_total,
    auto_soc,
    kosten,
    ziel_text,
    voll_text,
    geladene_energie,
    ziel_idx=None,
):
    auto_pv_kwh = float(np.sum(auto_pv))
    auto_netz_kwh = float(np.sum(auto_netz))
    auto_total_kwh = float(np.sum(auto_total))

    aktive_slots = int(np.count_nonzero(auto_total > 1e-9))
    minuten = aktive_slots * 15
    ladezeit_text = f"{minuten // 60} h {minuten % 60} min"

    if ziel_idx is not None:
        minuten_bis_ziel = max(0, int(ziel_idx - lade_start_slot) * 15)
        tage = minuten_bis_ziel // (24 * 60)
        rest_min = minuten_bis_ziel % (24 * 60)
        stunden = rest_min // 60
        minuten_rest = rest_min % 60
        if tage > 0:
            dauer_bis_ziel_text = f"{tage} Tage {stunden} h {minuten_rest} min"
        else:
            dauer_bis_ziel_text = f"{stunden} h {minuten_rest} min"
    else:
        dauer_bis_ziel_text = "—"

    if auto_total_kwh > 1e-9:
        final_soc = min(
            100.0,
            float(st.session_state.soc_aktuell)
            + geladene_energie / float(st.session_state.auto_kapazitaet_kwh) * 100.0,
        )
    else:
        final_soc = float(st.session_state.soc_aktuell)

    # Durchschnittlicher Kaufpreis nur für den gekauften Netzstrom:
    # Kosten (€) / Netzenergie (kWh) * 100 = ct/kWh
    # Wichtig: PV-Energie kostet hier 0 €, deshalb darf sie den Kaufpreis nicht verwässern.
    avg_ct = kosten / auto_netz_kwh * 100.0 if auto_netz_kwh > 1e-9 else 0.0

    return {
        "mode_key": mode_key,
        "strategie": strategie,
        "auto_pv": auto_pv,
        "auto_netz": auto_netz,
        "auto_total": auto_total,
        "auto_soc": auto_soc,
        "auto_pv_kwh": auto_pv_kwh,
        "auto_netz_kwh": auto_netz_kwh,
        "auto_total_kwh": auto_total_kwh,
        "kosten": float(kosten),
        "avg_ct": float(avg_ct),
        "ziel_text": ziel_text,
        "voll_text": voll_text,
        "ladezeit_text": ladezeit_text,
        "final_soc_text": f"{final_soc:.1f} %",
        "dauer_bis_ziel_text": dauer_bis_ziel_text,
    }


def simulate_direct_to_target():
    auto_pv = np.zeros(n_slots)
    auto_netz = np.zeros(n_slots)
    auto_total = np.zeros(n_slots)
    auto_soc = np.full(n_slots, float(st.session_state.soc_aktuell))

    soc_start = float(st.session_state.soc_aktuell)
    soc_ziel = float(st.session_state.soc_ziel)
    kap = float(st.session_state.auto_kapazitaet_kwh)
    pmax = float(st.session_state.ladeleistung_auto_kw)

    energie_ziel = max(0.0, (soc_ziel - soc_start) / 100.0 * kap)
    rest = energie_ziel
    geladen = 0.0
    kosten = 0.0
    ziel_idx = None

    for idx in range(lade_start_slot, n_slots):
        if rest <= 1e-9:
            break

        soc_vor = soc_start + geladen / kap * 100.0
        slot_max = ladeleistung_abhaengig_von_soc(soc_vor, pmax) * 0.25
        ladung = min(slot_max, rest)

        pv_rest = max(0.0, einspeisung_vor_auto[idx])
        pv_anteil = min(pv_rest, ladung)
        netz_anteil = ladung - pv_anteil

        auto_pv[idx] = pv_anteil
        auto_netz[idx] = netz_anteil
        auto_total[idx] = ladung

        kosten += netz_anteil * preise_z[idx] / 100.0
        geladen += ladung
        rest -= ladung

        auto_soc[idx] = min(100.0, soc_start + geladen / kap * 100.0)

        if rest <= 1e-9 and ziel_idx is None:
            ziel_idx = idx
            break

    last = soc_start
    for idx in range(lade_start_slot, n_slots):
        if auto_total[idx] > 1e-9:
            last = auto_soc[idx]
        else:
            auto_soc[idx] = last

    ziel_text = labels[ziel_idx] if ziel_idx is not None else "nicht innerhalb von 60 Tagen erreicht"

    if ziel_idx is not None and ziel_idx > abfahrt_slot:
        strategie_text = f"Direkt laden ab {ankunft_text} bis Ziel-SoC {int(soc_ziel)} %"
    else:
        strategie_text = f"Direkt laden ab {ankunft_text} bis Ziel-SoC {int(soc_ziel)} %"

    return result_pack(
        mode_key="direct",
        strategie=strategie_text,
        auto_pv=auto_pv,
        auto_netz=auto_netz,
        auto_total=auto_total,
        auto_soc=auto_soc,
        kosten=kosten,
        ziel_text=ziel_text,
        voll_text="nicht berechnet",
        geladene_energie=geladen,
        ziel_idx=ziel_idx,
    )


def simulate_pv_only_to_100():
    """
    Option B – Nur PV:
    Auto ist ab Ankunft verbunden. Es lädt nur mit PV-Überschuss.
    Wenn Ziel bis zur gewünschten Abfahrt nicht erreicht wird, läuft die Simulation weiter,
    bis Ziel-SoC erreicht ist oder der Simulationshorizont endet.
    """
    auto_pv = np.zeros(n_slots)
    auto_netz = np.zeros(n_slots)
    auto_total = np.zeros(n_slots)
    auto_soc = np.full(n_slots, float(st.session_state.soc_aktuell))

    soc_start = float(st.session_state.soc_aktuell)
    soc_ziel = float(st.session_state.soc_ziel)
    kap = float(st.session_state.auto_kapazitaet_kwh)
    pmax = float(st.session_state.ladeleistung_auto_kw)

    energie_ziel = max(0.0, (soc_ziel - soc_start) / 100.0 * kap)
    rest = energie_ziel
    geladen = 0.0
    ziel_idx = None

    for idx in range(lade_start_slot, n_slots):
        if rest <= 1e-9:
            break

        soc_vor = soc_start + geladen / kap * 100.0
        slot_max = ladeleistung_abhaengig_von_soc(soc_vor, pmax) * 0.25

        pv_rest = max(0.0, einspeisung_vor_auto[idx])
        ladung = min(slot_max, pv_rest, rest)

        if ladung <= 1e-9:
            continue

        auto_pv[idx] = ladung
        auto_total[idx] = ladung

        geladen += ladung
        rest -= ladung
        auto_soc[idx] = min(100.0, soc_start + geladen / kap * 100.0)

        if rest <= 1e-9 and ziel_idx is None:
            ziel_idx = idx
            break

    last = soc_start
    for idx in range(lade_start_slot, n_slots):
        if auto_total[idx] > 1e-9:
            last = auto_soc[idx]
        else:
            auto_soc[idx] = last

    ziel_text = labels[ziel_idx] if ziel_idx is not None else "nicht innerhalb von 60 Tagen erreicht"

    if ziel_idx is not None and ziel_idx > abfahrt_slot:
        strategie_text = f"Nur PV ab {ankunft_text} bis Ziel-SoC {int(soc_ziel)} %, ohne Netzstrom – Suche bis 60 Tage"
    else:
        strategie_text = f"Nur PV ab {ankunft_text} bis Ziel-SoC {int(soc_ziel)} %, ohne Netzstrom – Suche bis 60 Tage"

    return result_pack(
        mode_key="pv_only",
        strategie=strategie_text,
        auto_pv=auto_pv,
        auto_netz=auto_netz,
        auto_total=auto_total,
        auto_soc=auto_soc,
        kosten=0.0,
        ziel_text=ziel_text,
        voll_text=ziel_text if soc_ziel >= 100 else "—",
        geladene_energie=geladen,
        ziel_idx=ziel_idx,
    )


def simulate_smart_to_target():
    """
    Smart-Laden:
    - Auto ist ab Ankunft verbunden.
    - Ziel-SoC soll bis zur gewünschten Abfahrt erreicht werden.
    - Priorität 1: PV-Überschuss nutzen.
    - Priorität 2: Wenn PV bis Abfahrt nicht reicht, fehlende Energie aus dem Netz kaufen.
      Dabei werden die günstigsten Netzpreis-Slots im Fenster Ankunft -> Abfahrt gewählt.
    """
    auto_pv = np.zeros(n_slots)
    auto_netz = np.zeros(n_slots)
    auto_total = np.zeros(n_slots)
    auto_soc = np.full(n_slots, float(st.session_state.soc_aktuell))

    soc_start = float(st.session_state.soc_aktuell)
    soc_ziel = float(st.session_state.soc_ziel)
    kap = float(st.session_state.auto_kapazitaet_kwh)
    pmax = float(st.session_state.ladeleistung_auto_kw)

    energie_ziel = max(0.0, (soc_ziel - soc_start) / 100.0 * kap)

    if energie_ziel <= 1e-9:
        return result_pack(
            "smart",
            "Smart: Ziel bereits erreicht",
            auto_pv,
            auto_netz,
            auto_total,
            auto_soc,
            0.0,
            "bereits erreicht",
            "—",
            0.0,
            ziel_idx=lade_start_slot,
        )

    # Harte Smart-Grenze: nur zwischen Ankunft und gewünschter Abfahrt planen.
    ende_slot = int(min(max(abfahrt_slot, lade_start_slot + 1), n_slots))
    slots = np.arange(lade_start_slot, ende_slot + 1)

    if len(slots) == 0:
        return result_pack(
            "smart",
            "Smart: kein Ladefenster",
            auto_pv,
            auto_netz,
            auto_total,
            auto_soc,
            0.0,
            "kein Ladefenster",
            "—",
            0.0,
            ziel_idx=None,
        )

    # 1) PV zuerst über das Ladefenster einplanen.
    rest_test = energie_ziel
    geladen_test = 0.0
    pv_vorplanung = np.zeros(n_slots)

    for idx in slots:
        if rest_test <= 1e-9:
            break

        soc_vor = soc_start + geladen_test / kap * 100.0
        slot_max = ladeleistung_abhaengig_von_soc(soc_vor, pmax) * 0.25
        pv_rest = max(0.0, einspeisung_vor_auto[idx])

        pv_ladung = min(pv_rest, slot_max, rest_test)
        pv_vorplanung[idx] = pv_ladung
        geladen_test += pv_ladung
        rest_test -= pv_ladung

    # 2) Wenn PV bis Abfahrt nicht reicht, fehlende Energie aus Netz kaufen.
    #    Netz wird in den billigsten Slots innerhalb Ankunft -> Abfahrt geplant.
    fehlend = max(0.0, rest_test)
    netz_plan = np.zeros(n_slots)

    if fehlend > 1e-9:
        slots_billig = sorted(list(slots), key=lambda i: (preise_z[i], i))
        rest_plan = fehlend

        for idx in slots_billig:
            if rest_plan <= 1e-9:
                break

            # freie Ladeleistung im Slot nach PV-Nutzung
            frei = max(0.0, pmax * 0.25 - pv_vorplanung[idx])
            zusatz = min(frei, rest_plan)

            netz_plan[idx] = zusatz
            rest_plan -= zusatz

    # 3) Zeitlichen Ablauf ausführen: in jedem Slot PV zuerst, Netz nur nach Plan.
    rest = energie_ziel
    geladen = 0.0
    kosten = 0.0
    ziel_idx = None

    for idx in slots:
        if rest <= 1e-9:
            break

        soc_vor = soc_start + geladen / kap * 100.0
        slot_max = ladeleistung_abhaengig_von_soc(soc_vor, pmax) * 0.25

        pv_rest = max(0.0, einspeisung_vor_auto[idx])
        pv_anteil = min(pv_rest, slot_max, rest)

        frei = max(0.0, slot_max - pv_anteil)
        netz_anteil = min(netz_plan[idx], frei, max(0.0, rest - pv_anteil))

        ladung = pv_anteil + netz_anteil

        if ladung <= 1e-9:
            continue

        auto_pv[idx] = pv_anteil
        auto_netz[idx] = netz_anteil
        auto_total[idx] = ladung

        kosten += netz_anteil * preise_z[idx] / 100.0
        geladen += ladung
        rest -= ladung

        auto_soc[idx] = min(100.0, soc_start + geladen / kap * 100.0)

        if rest <= 1e-9 and ziel_idx is None:
            ziel_idx = idx
            break

    last = soc_start
    for idx in range(lade_start_slot, n_slots):
        if auto_total[idx] > 1e-9:
            last = auto_soc[idx]
        else:
            auto_soc[idx] = last

    if ziel_idx is None:
        ziel_text = "bis Abfahrt nicht erreicht"
        strategie_text = (
            f"Smart ab {ankunft_text} bis Ziel-SoC {int(soc_ziel)} % – "
            f"PV zuerst, Netz billig; Ziel bis Abfahrt {abfahrt_text} nicht erreicht"
        )
    else:
        ziel_text = labels[ziel_idx]
        strategie_text = (
            f"Smart ab {ankunft_text} bis Ziel-SoC {int(soc_ziel)} % – "
            f"PV zuerst, Netz billig; Ziel vor Abfahrt {abfahrt_text} erreicht"
        )

    return result_pack(
        mode_key="smart",
        strategie=strategie_text,
        auto_pv=auto_pv,
        auto_netz=auto_netz,
        auto_total=auto_total,
        auto_soc=auto_soc,
        kosten=kosten,
        ziel_text=ziel_text,
        voll_text=ziel_text if soc_ziel >= 100 else "—",
        geladene_energie=geladen,
        ziel_idx=ziel_idx,
    )


if start_button:
    if st.session_state.lade_modus == "🧠 Smart" and abfahrt_ungueltig:
        st.error(
            "❌ Die Abfahrt muss nach der Ankunft liegen. "
            "Bitte Abfahrtsdatum oder Abfahrtszeit ändern."
        )
        st.session_state.sim_result = None
    elif st.session_state.lade_modus == "🔌 Direkt":
        st.session_state.sim_result = simulate_direct_to_target()
    elif st.session_state.lade_modus == "☀️ PV":
        st.session_state.sim_result = simulate_pv_only_to_100()
    else:
        st.session_state.sim_result = simulate_smart_to_target()


sim = st.session_state.sim_result if st.session_state.sim_result is not None else empty_result()


# =============================================================================
# VISUAL PREP
# =============================================================================
sicht = slice(fokus_start, fokus_ende + 1)
zeit_labels_vis = [(l.split(" ", 1)[1] if " " in l else l) for l in labels[sicht]]

auto_pv = np.asarray(sim["auto_pv"], dtype=float)
auto_netz = np.asarray(sim["auto_netz"], dtype=float)
auto_total = np.asarray(sim["auto_total"], dtype=float)
auto_soc = np.asarray(sim["auto_soc"], dtype=float)

einspeisung_nach_auto = np.maximum(einspeisung_vor_auto - auto_pv, 0.0)
netz_mit_auto = haus_netzbezug + auto_netz

pv_sum = float(pv_z[sicht].sum())
haus_sum = float(verbrauch_z[sicht].sum())

pv_kwp_gesamt = 0.0
for _neigung in (st.session_state.neigungen_wahl or []):
    _cfg = st.session_state.pv_kwp_config.get(_neigung, {})
    pv_kwp_gesamt += float(sum(_cfg.get(_r, 0.0) for _r in ["Süd", "Ost", "Nord", "West"]))
speicher_laden_sum = float(speicher_laden[sicht].sum())
speicher_entladen_sum = float(speicher_entladen[sicht].sum())
netz_sum = float(netz_mit_auto[sicht].sum())
haus_netz_sum = float(haus_netzbezug[sicht].sum())
einspeisung_sum = float(einspeisung_nach_auto[sicht].sum())

# -------------------------------------------------------------------------
# AKTUELLE LEISTUNGSFLÜSSE FÜR DIE PFEILE
# -------------------------------------------------------------------------
# Die Karten/KPIs zeigen weiterhin kWh für den gewählten Zeitraum.
# Die Pfeile im Energiefluss zeigen aber nur, was im aktuellen 15-Minuten-Slot
# wirklich fließt. Wenn in diesem Slot kein Fluss vorhanden ist, wird kein Pfeil
# gezeichnet. Das verhindert unlogische Pfeile nur wegen Tages-/Monatssummen.
idx_flow = int(min(max(fokus_start, 0), n_slots - 1))

pv_slot = float(max(0.0, pv_z[idx_flow]))
haus_slot = float(max(0.0, verbrauch_z[idx_flow]))

pv_haus_slot = float(min(pv_slot, haus_slot))
pv_akku_slot = float(max(0.0, speicher_laden[idx_flow]))
akku_haus_slot = float(max(0.0, speicher_entladen[idx_flow]))
pv_auto_slot = float(max(0.0, auto_pv[idx_flow]))
netz_auto_slot = float(max(0.0, auto_netz[idx_flow]))
pv_netz_slot = float(max(0.0, einspeisung_nach_auto[idx_flow]))

# Umrechnung: kWh pro 15 Minuten × 4 = durchschnittliche kW-Leistung in diesem Slot
pv_haus_kw = pv_haus_slot * 4.0
pv_akku_kw = pv_akku_slot * 4.0
akku_haus_kw = akku_haus_slot * 4.0
pv_auto_kw = pv_auto_slot * 4.0
netz_auto_kw = netz_auto_slot * 4.0
pv_netz_kw = pv_netz_slot * 4.0

# Aktueller Zustand des Hausspeichers im gleichen Zeitpunkt wie die Pfeile
if len(speicher_soc) > idx_flow:
    akku_soc_aktuell = float(speicher_soc[idx_flow])
elif len(speicher_soc) > 0:
    akku_soc_aktuell = float(speicher_soc[-1])
else:
    akku_soc_aktuell = 0.0

# Für die Pfeile verwenden wir die Summe im sichtbaren Zeitraum.
# Dadurch sieht man eine vollständige Richtung, sobald an diesem Tag/Zeitraum
# Energie in diese Richtung geflossen ist. Wenn nichts geflossen ist, bleibt
# der Pfeil grau.
pv_haus_flow = float(min(pv_sum, haus_sum))
pv_akku_flow = float(speicher_laden_sum)
akku_haus_flow = float(speicher_entladen_sum)
pv_auto_flow = float(sim["auto_pv_kwh"])
netz_auto_flow = float(sim["auto_netz_kwh"])
pv_netz_flow = float(einspeisung_sum)

# Netzbezug der Hauslast separat anzeigen.
# Das ist nicht Auto-Netzstrom, sondern Strom, den das Haus aus dem Netz braucht.
netz_haus_flow = float(haus_netz_sum)


# =============================================================================
# ENERGY HTML
# =============================================================================
def render_energy_html():
    dark = st.session_state.theme_mode == "🌙 Dunkel"

    if dark:
        title_col = "#ffffff"
        text_col = "#dce7ff"
        muted_col = "#aebbd3"
        card_bg = "rgba(255,255,255,0.09)"
        card_border = "rgba(255,255,255,0.18)"
        kpi_bg = "rgba(255,255,255,0.075)"
        detail_bg = "rgba(0,0,0,0.18)"
    else:
        title_col = "#172033"
        text_col = "#334155"
        muted_col = "#64748b"
        card_bg = "rgba(255,255,255,0.88)"
        card_border = "rgba(15,23,42,0.12)"
        kpi_bg = "rgba(255,255,255,0.82)"
        detail_bg = "rgba(255,255,255,0.72)"

    mode_key = sim["mode_key"]

    auto_pv_kwh = float(sim["auto_pv_kwh"])
    auto_netz_kwh = float(sim["auto_netz_kwh"])
    auto_total_kwh = float(sim["auto_total_kwh"])

    pv_anteil = auto_pv_kwh / auto_total_kwh * 100.0 if auto_total_kwh > 1e-9 else 0.0
    netz_anteil = auto_netz_kwh / auto_total_kwh * 100.0 if auto_total_kwh > 1e-9 else 0.0

    pv_text = f"{auto_pv_kwh:.1f} kWh ({pv_anteil:.1f} %)"
    netz_text = f"{auto_netz_kwh:.1f} kWh ({netz_anteil:.1f} %)"

    ziel_erreicht_text = sim["ziel_text"]
    if sim["mode_key"] == "none":
        ziel_status_text = "kein Auto verbunden"
    elif "nicht" in str(ziel_erreicht_text).lower() or "kein" in str(ziel_erreicht_text).lower():
        ziel_status_text = "nicht erreicht"
    else:
        ziel_status_text = f"erreicht: {ziel_erreicht_text}"

    energie_bis_ziel = max(
        0.0,
        (float(st.session_state.soc_ziel) - float(st.session_state.soc_aktuell))
        / 100.0
        * float(st.session_state.auto_kapazitaet_kwh)
    )
    fehlend_bis_ziel = max(0.0, energie_bis_ziel - auto_total_kwh)

    if mode_key == "pv_only":
        car_top_cards = f"""
        <div class="detail-card"><div class="label">Benötigt bis Ziel</div><div class="value">{energie_bis_ziel:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Fehlend bis Ziel</div><div class="value">{fehlend_bis_ziel:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Dauer bis Ziel</div><div class="value">{sim["dauer_bis_ziel_text"]}</div></div>
        <div class="detail-card"><div class="label">Gesamte PV-Ladung</div><div class="value">{auto_total_kwh:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Aus PV geladen</div><div class="value">{pv_text}</div></div>
        <div class="detail-card"><div class="label">Aktive Ladezeit</div><div class="value">{sim["ladezeit_text"]}</div></div>
        <div class="detail-card"><div class="label">Ziel-SoC erreicht</div><div class="value">{sim["ziel_text"]}</div></div>
        <div class="detail-card"><div class="label">100 % erreicht</div><div class="value">{sim["voll_text"]}</div></div>
        <div class="detail-card"><div class="label">SoC am Ende</div><div class="value">{sim["final_soc_text"]}</div></div>
        """
    else:
        car_top_cards = f"""
        <div class="detail-card"><div class="label">Benötigt bis Ziel</div><div class="value">{energie_bis_ziel:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Fehlend bis Ziel</div><div class="value">{fehlend_bis_ziel:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Dauer bis Ziel</div><div class="value">{sim["dauer_bis_ziel_text"]}</div></div>
        <div class="detail-card"><div class="label">Bezahlt für Netzstrom</div><div class="value">{sim["kosten"]:.2f} €</div></div>
        <div class="detail-card"><div class="label">Ø Kaufpreis Netz</div><div class="value">{sim["avg_ct"]:.2f} ct/kWh</div></div>
        <div class="detail-card"><div class="label">Gekauft aus Netz</div><div class="value">{auto_netz_kwh:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Aktive Ladezeit</div><div class="value">{sim["ladezeit_text"]}</div></div>
        <div class="detail-card"><div class="label">Gesamte Auto-Ladung</div><div class="value">{auto_total_kwh:.1f} kWh</div></div>
        <div class="detail-card"><div class="label">Aus PV geladen</div><div class="value">{pv_text}</div></div>
        <div class="detail-card"><div class="label">Aus Netz geladen</div><div class="value">{netz_text}</div></div>
        <div class="detail-card"><div class="label">Ziel-SoC erreicht</div><div class="value">{sim["ziel_text"]}</div></div>
        <div class="detail-card"><div class="label">100 % erreicht</div><div class="value">{sim["voll_text"]}</div></div>
        <div class="detail-card"><div class="label">SoC am Ende</div><div class="value">{sim["final_soc_text"]}</div></div>
        """

    car_energy_svg = svg_multi(
        {
            "PV → Auto": auto_pv[sicht],
            "Netz → Auto": auto_netz[sicht],
        },
        zeit_labels=zeit_labels_vis,
    )

    car_soc_svg = svg_line(auto_soc[sicht], zeit_labels=zeit_labels_vis, farbe="#00d1b2")
    pv_svg = svg_line(pv_z[sicht], zeit_labels=zeit_labels_vis, farbe="#f1c40f")
    haus_svg = svg_line(verbrauch_z[sicht], zeit_labels=zeit_labels_vis, farbe="#3498db", keep_above_zero=True)
    speicher_svg = svg_line(speicher_soc[sicht], zeit_labels=zeit_labels_vis, farbe="#2ecc71")
    netz_svg = svg_line(netz_mit_auto[sicht], zeit_labels=zeit_labels_vis, farbe="#9b59b6")
    einspeisung_svg = svg_line(einspeisung_nach_auto[sicht], zeit_labels=zeit_labels_vis, farbe="#ff9f43")

    def active(v):
        return "flow-line active" if v > 0.01 else "flow-line muted"

    def hidden(v):
        return "" if v > 0.01 else "hidden"

    html = f"""
    <html>
    <head>
    <style>
    html, body {{
        margin: 0;
        padding: 0;
        background: transparent !important;
        overflow: hidden;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
    }}

    .energy-wrapper {{
        width: 100%;
        min-height: 560px;
        border-radius: 0;
        padding: 4px 0 0 0;
        background: transparent;
        overflow: hidden;
        border: none;
    }}

    .energy-sub {{
        color: {muted_col};
        font-size: .84rem;
        margin-bottom: 8px;
    }}

    .box-view {{
        display: none;
    }}

    .box-view.active {{
        display: block;
    }}

    .energy-stage {{
        position: relative;
        width: 100%;
        height: 330px;
    }}

    .energy-svg {{
        position: absolute;
        inset: 0;
        width: 100%;
        height: 100%;
        z-index: 1;
    }}

    .node {{
        position: absolute;
        z-index: 2;
        width: 150px;
        min-height: 92px;
        border-radius: 18px;
        padding: 12px;
        text-align: center;
        background: {card_bg};
        border: 1px solid {card_border};
        color: {title_col};
        box-shadow: 0 8px 18px rgba(0,0,0,.18);
        cursor: pointer;
        transition: .18s;
    }}

    .node:hover {{
        transform: translateY(-3px) scale(1.02);
        border-color: rgba(0,209,178,.9);
        box-shadow: 0 0 22px rgba(0,209,178,.28);
    }}

    .icon {{
        font-size: 2rem;
        line-height: 1;
        margin-bottom: 6px;
    }}

    .name {{
        font-weight: 850;
        font-size: .98rem;
    }}

    .value {{
        margin-top: 4px;
        font-size: .88rem;
        color: {text_col};
    }}

    .small {{
        margin-top: 2px;
        font-size: .72rem;
        color: {muted_col};
    }}

    .pv-node {{left: calc(50% - 75px); top: 5px;}}
    .battery-node {{left: 5%; top: 188px;}}
    .house-node {{left: calc(50% - 75px); top: 205px;}}
    .car-node {{right: 8%; top: 205px;}}
    .grid-node {{right: 3%; top: 22px;}}

    .flow-line {{
        fill: none;
        stroke-width: 6;
        stroke-linecap: round;
        stroke-dasharray: 14 14;
        animation: flowMove 1s linear infinite;
        filter: drop-shadow(0 0 5px rgba(255,255,255,.25));
    }}

    .active {{opacity: 1;}}
    .muted {{
        opacity: .32;
        animation: none;
        stroke: #6b7280 !important;
        stroke-dasharray: 7 11;
        filter: none;
    }}

    .pv-flow {{stroke: #f1c40f;}}
    .battery-flow {{stroke: #2ecc71;}}
    .grid-flow {{stroke: #6aa8ff;}}
    .grid-house-flow {{stroke: #ff4b4b;}}
    .car-flow {{stroke: #00d1b2;}}
    .feed-flow {{stroke: #ff9f43;}}

    @keyframes flowMove {{
        to {{stroke-dashoffset: -56;}}
    }}

    .flow-label {{
        position: absolute;
        z-index: 3;
        padding: 5px 8px;
        border-radius: 999px;
        background: rgba(0,0,0,.45);
        color: #fff;
        font-size: .75rem;
        border: 1px solid rgba(255,255,255,.16);
        white-space: nowrap;
    }}

    .hidden {{
        display: none !important;
    }}

    .label-pv-house {{left: calc(50% - 145px); top: 178px;}}
    .label-pv-battery {{left: 20%; top: 150px;}}
    .label-akku-house {{left: 28%; top: 285px;}}
    .label-pv-car {{right: 24%; top: 154px;}}
    .label-grid {{right: 19%; top: 94px;}}
    .label-grid-house {{right: 31%; top: 205px; border-color: rgba(255,75,75,.55);}}
    .label-feed {{right: 14%; top: 72px;}}

    .kpis {{
        display: grid;
        grid-template-columns: repeat(6, minmax(110px, 1fr));
        gap: 8px;
        margin-top: 10px;
    }}

    .kpi {{
        background: {kpi_bg};
        border: 1px solid {card_border};
        border-radius: 14px;
        padding: 10px 12px;
        color: {title_col};
    }}

    .kpi-label {{
        color: {muted_col};
        font-size: .78rem;
    }}

    .kpi-value {{
        font-weight: 850;
        font-size: 1.02rem;
        margin-top: 3px;
        color: {title_col};
    }}

    .detail-header {{
        display: flex;
        justify-content: space-between;
        align-items: center;
        padding: 4px 2px 10px 2px;
    }}

    .detail-title {{
        color: {title_col};
        font-size: 1.26rem;
        font-weight: 850;
    }}

    .back-btn {{
        border: 1px solid {card_border};
        background: {card_bg};
        color: {title_col};
        padding: 8px 12px;
        border-radius: 12px;
        cursor: pointer;
        font-weight: 750;
    }}

    .detail-grid {{
        display: grid;
        grid-template-columns: repeat(3, minmax(150px, 1fr));
        gap: 8px;
        margin-bottom: 10px;
    }}

    .detail-card {{
        background: {card_bg};
        border: 1px solid {card_border};
        border-radius: 16px;
        padding: 8px 10px;
        color: {title_col};
    }}

    .detail-card .label {{
        color: {muted_col};
        font-size: .80rem;
    }}

    .detail-card .value {{
        color: {title_col};
        font-size: 1.00rem;
        font-weight: 850;
        margin-top: 4px;
    }}

    .mini-chart {{
        padding: 6px 8px 2px 8px;
        border-radius: 16px;
        background: {detail_bg};
        border: 1px solid {card_border};
        margin-bottom: 6px;
    }}

    .chart-title {{
        color: {title_col};
        font-weight: 850;
        margin-bottom: 5px;
        font-size: .9rem;
    }}

    .detail-text {{
        color: {text_col};
        background: {detail_bg};
        border: 1px solid {card_border};
        border-radius: 16px;
        padding: 10px 12px;
        line-height: 1.35;
        font-size: .86rem;
    }}
    </style>
    </head>

    <body>
    <div class="energy-wrapper">
        <div class="energy-sub">
            Zeitraum: <b>{ansicht_titel}</b> · Ankunft: <b>{ankunft_text}</b>{(" · gewünschte Abfahrt: <b>" + abfahrt_text + "</b>") if st.session_state.lade_modus == "🧠 Smart" else ""}<br>
            Strategie: <b>{sim["strategie"]}</b><br>
            Farbige Pfeile zeigen Energiefluss im gewählten Zeitraum. Graue Pfeile bedeuten: in dieser Richtung gab es keinen Energiefluss.
        </div>

        <div id="view-overview" class="box-view active">
            <div class="energy-stage">
                <svg class="energy-svg" viewBox="0 0 900 375" preserveAspectRatio="none">
                    <defs>
                        <marker id="arrowYellow" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#f1c40f"/></marker>
                        <marker id="arrowGreen" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#2ecc71"/></marker>
                        <marker id="arrowBlue" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#6aa8ff"/></marker>
                        <marker id="arrowRed" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#ff4b4b"/></marker>
                        <marker id="arrowTeal" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#00d1b2"/></marker>
                        <marker id="arrowOrange" markerWidth="10" markerHeight="10" refX="8" refY="3" orient="auto"><path d="M0,0 L0,6 L9,3 z" fill="#ff9f43"/></marker>
                    </defs>

                    <path class="{active(pv_haus_flow)} pv-flow" d="M450 103 C450 145,450 174,450 212" marker-end="url(#arrowYellow)" />
                    <path class="{active(pv_akku_flow)} battery-flow" d="M420 98 C330 130,245 154,150 202" marker-end="url(#arrowGreen)" />
                    <path class="{active(akku_haus_flow)} battery-flow" d="M185 235 C280 258,350 260,430 240" marker-end="url(#arrowGreen)" />
                    <path class="{active(pv_auto_flow)} car-flow" d="M480 98 C570 130,655 168,755 212" marker-end="url(#arrowTeal)" />
                    <path class="{active(netz_auto_flow)} grid-flow" d="M760 100 C722 138,718 174,750 212" marker-end="url(#arrowBlue)" />
                    <path class="{active(netz_haus_flow)} grid-house-flow" d="M805 108 C735 165,635 205,520 230" marker-end="url(#arrowRed)" />
                    <path class="{active(pv_netz_flow)} feed-flow" d="M520 73 C650 48,735 48,815 73" marker-end="url(#arrowOrange)" />
                </svg>

                <div class="node pv-node" onclick="showEnergyView('pv')">
                    <div class="icon">☀️</div><div class="name">PV-Anlage</div>
                    <div class="value">{pv_sum:.1f} kWh</div><div class="small">anklicken</div>
                </div>

                <div class="node battery-node" onclick="showEnergyView('battery')">
                    <div class="icon">🔋</div><div class="name">Akku</div>
                    <div class="value">SoC: {akku_soc_aktuell:.0f} %</div>
                    <div class="small">+{speicher_laden_sum:.1f} / −{speicher_entladen_sum:.1f} kWh · anklicken</div>
                </div>

                <div class="node house-node" onclick="showEnergyView('house')">
                    <div class="icon">🏠</div><div class="name">Haus</div>
                    <div class="value">{haus_sum:.1f} kWh</div><div class="small">anklicken</div>
                </div>

                <div class="node car-node" onclick="showEnergyView('car')">
                    <div class="icon">🚗</div><div class="name">E-Auto</div>
                    <div class="value">{auto_total_kwh:.1f} kWh</div>
                    <div class="small">Ziel: {ziel_status_text}</div>
                </div>

                <div class="node grid-node" onclick="showEnergyView('grid')">
                    <div class="icon">⚡</div><div class="name">Netz</div>
                    <div class="value">Bezug: {netz_sum:.1f} kWh</div>
                    <div class="small">Haus {netz_haus_flow:.1f} / Auto {netz_auto_flow:.1f} kWh</div>
                </div>

                <div class="flow-label label-pv-house {hidden(pv_haus_flow)}">PV → Haus: {pv_haus_flow:.1f} kWh</div>
                <div class="flow-label label-pv-battery {hidden(pv_akku_flow)}">PV → Akku: {pv_akku_flow:.1f} kWh</div>
                <div class="flow-label label-akku-house {hidden(akku_haus_flow)}">Akku → Haus: {akku_haus_flow:.1f} kWh</div>
                <div class="flow-label label-pv-car {hidden(pv_auto_flow)}">PV → Auto: {pv_auto_flow:.1f} kWh</div>
                <div class="flow-label label-grid {hidden(netz_auto_flow)}">Netz → Auto: {netz_auto_flow:.1f} kWh</div>
                <div class="flow-label label-grid-house {hidden(netz_haus_flow)}">Netz → Haus: {netz_haus_flow:.1f} kWh</div>
                <div class="flow-label label-feed {hidden(pv_netz_flow)}">PV → Netz: {pv_netz_flow:.1f} kWh</div>
            </div>

            <div class="kpis">
                <div class="kpi"><div class="kpi-label">PV-Erzeugung</div><div class="kpi-value">{pv_sum:.1f} kWh</div></div>
                <div class="kpi"><div class="kpi-label">Hausverbrauch</div><div class="kpi-value">{haus_sum:.1f} kWh</div></div>
                <div class="kpi"><div class="kpi-label">Auto-Ladung</div><div class="kpi-value">{auto_total_kwh:.1f} kWh</div></div>
                <div class="kpi"><div class="kpi-label">Auto aus PV</div><div class="kpi-value">{auto_pv_kwh:.1f} kWh</div></div>
                <div class="kpi"><div class="kpi-label">Netz → Haus</div><div class="kpi-value">{netz_haus_flow:.1f} kWh</div></div>
                <div class="kpi"><div class="kpi-label">Ziel-SoC</div><div class="kpi-value">{ziel_status_text}</div></div>
            </div>
        </div>

        <div id="view-pv" class="box-view">
            <div class="detail-header"><div class="detail-title">☀️ PV-Informationen</div><button class="back-btn" onclick="showEnergyView('overview')">⬅ Zurück</button></div>
            <div class="detail-grid">
                <div class="detail-card"><div class="label">PV-Erzeugung</div><div class="value">{pv_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Dachneigungen</div><div class="value">{", ".join(neigungen_wahl)}</div></div>
                <div class="detail-card"><div class="label">installierte PV-Leistung</div><div class="value">{pv_kwp_gesamt:.1f} kWp</div></div>
                <div class="detail-card"><div class="label">PV → Auto</div><div class="value">{pv_text}</div></div>
                <div class="detail-card"><div class="label">PV → Netz</div><div class="value">{einspeisung_sum:.1f} kWh</div></div>
            </div>
            <div class="mini-chart"><div class="chart-title">PV-Erzeugungsverlauf</div>{pv_svg}</div>
        </div>

        <div id="view-battery" class="box-view">
            <div class="detail-header"><div class="detail-title">🔋 Akku-Informationen</div><button class="back-btn" onclick="showEnergyView('overview')">⬅ Zurück</button></div>
            <div class="detail-grid">
                <div class="detail-card"><div class="label">Akku-Zustand jetzt</div><div class="value">{akku_soc_aktuell:.0f} %</div></div>
                <div class="detail-card"><div class="label">Geladen</div><div class="value">{speicher_laden_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Entladen</div><div class="value">{speicher_entladen_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Netto</div><div class="value">{speicher_laden_sum - speicher_entladen_sum:.1f} kWh</div></div>
            </div>
            <div class="mini-chart"><div class="chart-title">SoC-Verlauf Hausspeicher</div>{speicher_svg}</div>
        </div>

        <div id="view-house" class="box-view">
            <div class="detail-header"><div class="detail-title">🏠 Haus-Informationen</div><button class="back-btn" onclick="showEnergyView('overview')">⬅ Zurück</button></div>
            <div class="detail-grid">
                <div class="detail-card"><div class="label">Hausverbrauch</div><div class="value">{haus_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Haus-Netzbezug</div><div class="value">{haus_netz_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Grundlast</div><div class="value">sichtbar</div></div>
            </div>
            <div class="mini-chart"><div class="chart-title">Hausverbrauchsverlauf</div>{haus_svg}</div>
        </div>

        <div id="view-car" class="box-view">
            <div class="detail-header"><div class="detail-title">🚗 E-Auto Lade-Informationen</div><button class="back-btn" onclick="showEnergyView('overview')">⬅ Zurück</button></div>

            <div class="detail-grid">
                {car_top_cards}
            </div>

            <div class="mini-chart">
                <div class="chart-title">Energiequelle der Auto-Ladung</div>
                {car_energy_svg}
            </div>

            <div class="mini-chart">
                <div class="chart-title">SoC-Verlauf des Elektroautos</div>
                {car_soc_svg}
            </div>

            <div class="detail-text">
                Oben sieht man, wann das Auto Strom aus PV oder Netz bekommt.
                Unten sieht man den Ladestand der Autobatterie.
                Im PV-Modus sucht das System bis zu 60 Tage weiter, bis der Ziel-SoC erreicht wird.
            </div>
        </div>

        <div id="view-grid" class="box-view">
            <div class="detail-header"><div class="detail-title">⚡ Netz-Informationen</div><button class="back-btn" onclick="showEnergyView('overview')">⬅ Zurück</button></div>
            <div class="detail-grid">
                <div class="detail-card"><div class="label">Netzbezug gesamt</div><div class="value">{netz_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Netz → Haus</div><div class="value">{netz_haus_flow:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Netz → Auto</div><div class="value">{netz_auto_flow:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Haus aus Netz</div><div class="value">{haus_netz_sum:.1f} kWh</div></div>
                <div class="detail-card"><div class="label">Auto aus Netz</div><div class="value">{netz_text}</div></div>
                <div class="detail-card"><div class="label">PV → Netz</div><div class="value">{einspeisung_sum:.1f} kWh</div></div>
            </div>
            <div class="mini-chart"><div class="chart-title">Netzbezug</div>{netz_svg}</div>
            <div class="mini-chart"><div class="chart-title">PV-Einspeisung ins Netz</div>{einspeisung_svg}</div>
        </div>
    </div>

    <script>
    function showEnergyView(name) {{
        const views = document.querySelectorAll('.box-view');
        views.forEach(v => v.classList.remove('active'));
        const target = document.getElementById('view-' + name);
        if (target) target.classList.add('active');
    }}
    </script>
    </body>
    </html>
    """

    return textwrap.dedent(html).strip()


with energiefluss_placeholder.container():
    components.html(render_energy_html(), height=760, scrolling=True)

# Keine weiteren Ausgaben.
