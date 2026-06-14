# -*- coding: utf-8 -*-

import streamlit as st
import pandas as pd
import numpy as np
import os
import datetime as dt
import plotly.graph_objects as go
import requests


# =============================================================================
# STREAMLIT CONFIG
# =============================================================================
st.set_page_config(page_title="EMS Elektroauto", layout="wide")


# =============================================================================
# KONSTANTEN
# =============================================================================
PROFIL_START = dt.date(2025, 1, 1)
MAX_PROGNOSE_TAGE = 5

MONATE = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

QUELLEN_OPTIONEN = [
    "📅 Historische Daten (Excel des letzten Jahres)",
    "☁️ Online-Wetterdaten (Simulierte API)",
    "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)"
]

AZIMUT_WINKEL = {
    "Süd": 0,
    "Ost": -90,
    "West": 90,
    "Nord": 180
}


# =============================================================================
# SESSION STATE
# =============================================================================
if "theme_mode" not in st.session_state:
    st.session_state.theme_mode = "☀️ Hell"

if "quelle" not in st.session_state:
    st.session_state.quelle = "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)"

if "ansicht" not in st.session_state:
    st.session_state.ansicht = "📆 Heute (24h ab jetzt)"

if "anlage_lat" not in st.session_state:
    st.session_state.anlage_lat = 51.71

if "anlage_lon" not in st.session_state:
    st.session_state.anlage_lon = 8.76

if "anlage_adresse" not in st.session_state:
    st.session_state.anlage_adresse = "Paderborn (Standard)"


aktuelle_zeit = dt.datetime.now()
heute = aktuelle_zeit.date()


# =============================================================================
# HILFSFUNKTIONEN DATUM
# =============================================================================
def klemme_datum_auf_2025(d: dt.date) -> dt.date:
    """
    Für historische Excel-Profile wird jedes reale Datum auf das Jahr 2025 gemappt.
    Damit kann ein Standard-Jahresprofil wiederverwendet werden.
    """
    try:
        return d.replace(year=2025)
    except ValueError:
        return dt.date(2025, d.month, 28)


def datum_kategorie(datum: dt.date) -> str:
    delta = (datum - heute).days

    if datum < heute:
        return "vergangenheit"
    elif datum == heute:
        return "heute"
    elif 0 < delta <= MAX_PROGNOSE_TAGE:
        return "prognose"
    else:
        return "zu_weit"


def slot_von_datum(datum: dt.date, uhrzeit_slot: int = 0) -> int:
    delta = (datum - PROFIL_START).days
    return delta * 96 + uhrzeit_slot


def zeitachse_erstellen(start_datum: dt.date, start_uhrzeit_slot: int, n_slots: int):
    start_dt = dt.datetime.combine(start_datum, dt.time()) + dt.timedelta(minutes=15 * start_uhrzeit_slot)
    return [
        (start_dt + dt.timedelta(minutes=15 * i)).strftime("%d.%m %H:%M")
        for i in range(n_slots)
    ]


def slots_fuer_zeitraum(arr, start_slot: int, n_slots: int):
    arr = np.asarray(arr)
    total = len(arr)

    if total == 0:
        return np.zeros(n_slots)

    return np.array([arr[(start_slot + i) % total] for i in range(n_slots)])


def uhrzeit_zu_slot(uhrzeit_text: str) -> int:
    h, m = map(int, uhrzeit_text.split(":"))
    return (h * 60 + m) // 15


# =============================================================================
# THEME
# =============================================================================
col_titel, col_zeit, col_theme = st.columns([3, 1, 1])

with col_titel:
    st.title("🔌 EMS – Steuerungslogik für ein Elektroauto")
    st.markdown("Ladeplanung basierend auf PV-Überschuss, Batteriespeicher und Börsenstrompreisen.")

with col_zeit:
    st.markdown(
        f"""
        <div style='text-align:right;font-size:1.1rem;font-weight:bold;
        color:#4A90E2;padding-top:20px;'>
        🕒 {aktuelle_zeit.strftime('%d.%m.%Y — %H:%M')}
        </div>
        """,
        unsafe_allow_html=True
    )

with col_theme:
    theme_auswahl = st.selectbox(
        "Design:",
        ["☀️ Hell", "🌙 Dunkel"],
        index=0 if st.session_state.theme_mode == "☀️ Hell" else 1,
        key="theme_input"
    )
    st.session_state.theme_mode = theme_auswahl


if st.session_state.theme_mode == "🌙 Dunkel":
    plotly_template = "plotly_dark"
    container_bg = "#1e2530"
    val_color = "#ffffff"
    label_color = "#a0aec0"

    st.markdown(
        """
        <style>
        .stApp { background-color: #0e1117 !important; color: #ffffff !important; }
        h1, h2, h3, h4, h5, h6, p, span, label { color: #ffffff !important; }
        div[data-testid="stExpander"] { background-color: #161b22 !important; }
        </style>
        """,
        unsafe_allow_html=True
    )
else:
    plotly_template = "plotly"
    container_bg = "#f0f2f6"
    val_color = "#2c3e50"
    label_color = "#7f8c8d"


st.markdown("---")


# =============================================================================
# DATENQUELLE
# =============================================================================
st.subheader("🌐 Datenquelle für PV-Erzeugung")

daten_quelle = st.radio(
    "Bitte wähle die Basis für die Berechnung der Sonnenenergie:",
    QUELLEN_OPTIONEN,
    horizontal=True,
    key="quelle_input",
    index=QUELLEN_OPTIONEN.index(st.session_state.quelle)
    if st.session_state.quelle in QUELLEN_OPTIONEN else 0
)

st.session_state.quelle = daten_quelle

if "Historische" in daten_quelle:
    st.info("📋 Historische Daten: Verwendet das Jahresprofil aus der Excel-Datei.")
elif "Simulierte" in daten_quelle:
    st.info("⚙️ Simulierte API: Historisches PV-Profil wird pauschal mit Faktor 1,2 skaliert.")
else:
    st.success("🛰️ Live-Wetterdaten: Verwendet Open-Meteo mit Strahlung und Temperatur.")

st.markdown("---")


# =============================================================================
# EXCEL LADEN
# =============================================================================
def finde_excel_datei():
    user_home = os.path.expanduser("~")

    suchpfade = [
        "Eingangsdaten.xlsx",
        "Eingangsdaten - Profile.xlsx",
        os.path.join(user_home, "Desktop", "Eingangsdaten.xlsx"),
        os.path.join(user_home, "Desktop", "Eingangsdaten - Profile.xlsx"),
        os.path.join(user_home, "Desktop", "eg", "Eingangsdaten.xlsx"),
        os.path.join(user_home, "Desktop", "psp", "Eingangsdaten.xlsx"),
        r"C:\Users\elham\Desktop\eg\Eingangsdaten.xlsx",
        r"C:\Users\elham\Desktop\eg\Eingangsdaten - Profile.xlsx",
    ]

    for p in suchpfade:
        if os.path.isfile(p):
            return p

    return None


@st.cache_data
def lade_excel_von_pfad(pfad):
    xl = pd.ExcelFile(pfad, engine="openpyxl")
    return parse_excel(xl), pfad


@st.cache_data
def lade_excel_upload(uploaded_file):
    xl = pd.ExcelFile(uploaded_file, engine="openpyxl")
    return parse_excel(xl), uploaded_file.name


def parse_excel(xl):
    sheets = xl.sheet_names

    # Haushalt
    hs_sheet = next(
        s for s in sheets
        if "haushalt" in s.lower() or "strom" in s.lower()
    )
    df_hs = xl.parse(hs_sheet, header=None)
    hs_kwh = pd.to_numeric(df_hs.iloc[2:, 1], errors="coerce").fillna(0).values * 0.25 / 1000

    # Wärme
    waerme_sheet = next(
        (s for s in sheets if "wärme" in s.lower() or "waerme" in s.lower()),
        None
    )

    if waerme_sheet:
        df_w = xl.parse(waerme_sheet, header=None)

        rw = pd.to_numeric(df_w.iloc[4:, 1], errors="coerce").fillna(0).values * (10140 / 1000)
        ww = pd.to_numeric(df_w.iloc[4:, 2], errors="coerce").fillna(0).values * (2433 / 1000)

        min_len = min(len(rw), len(ww))
        waerme_kwh = rw[:min_len] + ww[:min_len]
    else:
        waerme_kwh = np.zeros(len(hs_kwh))

    # Preise
    preis_sheet = next(
        s for s in sheets
        if "preis" in s.lower() or "börse" in s.lower() or "boerse" in s.lower()
    )
    df_p = xl.parse(preis_sheet, header=None)
    preise = pd.to_numeric(df_p.iloc[1:, 1], errors="coerce").fillna(0).values

    # PV
    pv_dict = {}

    for sheet in sheets:
        if "pv" not in sheet.lower() and "neigung" not in sheet.lower():
            continue

        df_pv = xl.parse(sheet, header=None)

        for col_idx in range(1, 5):
            try:
                neigung = str(df_pv.iloc[2, col_idx]).strip()
                richtung = str(df_pv.iloc[3, col_idx]).strip()

                if neigung.lower() == "nan" or richtung.lower() == "nan":
                    continue

                jahres_kwh_kwp = float(df_pv.iloc[0, col_idx])

                kurve_raw = pd.to_numeric(df_pv.iloc[5:, col_idx], errors="coerce").fillna(0).values

                # Annahme: Excel-Werte sind W/kWp je 15 min.
                kurve_kwh_kwp = kurve_raw * 0.25 / 1000

                pv_dict[(neigung, richtung)] = {
                    "kurve": kurve_kwh_kwp,
                    "jahres_kwh": jahres_kwh_kwp
                }

            except Exception:
                continue

    return {
        "hs": np.asarray(hs_kwh),
        "waerme": np.asarray(waerme_kwh),
        "preise": np.asarray(preise),
        "pv": pv_dict
    }


uploaded_file = st.file_uploader(
    "Optional: Eingangsdaten.xlsx hochladen",
    type=["xlsx"]
)

daten = None
gefundener_pfad = None

try:
    if uploaded_file is not None:
        daten, gefundener_pfad = lade_excel_upload(uploaded_file)
    else:
        pfad = finde_excel_datei()
        if pfad:
            daten, gefundener_pfad = lade_excel_von_pfad(pfad)
except Exception as e:
    st.error(f"❌ Fehler beim Laden der Excel-Datei: {e}")
    st.stop()

if daten is None:
    st.error("❌ Excel-Datei nicht gefunden.")
    st.info("Bitte lege `Eingangsdaten.xlsx` in den Projektordner/Desktop oder lade sie oben hoch.")
    st.stop()


# =============================================================================
# OPEN-METEO
# =============================================================================
@st.cache_data(ttl=3600)
def hole_open_meteo_live(lat: float, lon: float, start_datum: dt.date, ende_datum: dt.date) -> dict:
    heute_local = dt.date.today()
    ergebnisse = {}

    def fetch_data(url):
        try:
            r = requests.get(url, timeout=15)
            r.raise_for_status()

            data = r.json()

            times = data["hourly"]["time"]
            rads = data["hourly"]["shortwave_radiation"]
            temps = data["hourly"]["temperature_2m"]

            for t_str, rad, temp in zip(times, rads, temps):
                d = dt.date.fromisoformat(t_str[:10])
                hhmm = t_str[11:16]

                if d not in ergebnisse:
                    ergebnisse[d] = {}

                ergebnisse[d][hhmm] = {
                    "ghi": float(rad or 0.0),
                    "temp": float(temp or 15.0)
                }

        except Exception:
            pass

    # Historische Daten bis gestern
    if start_datum < heute_local:
        hist_ende = min(ende_datum, heute_local - dt.timedelta(days=1))

        url_hist = (
            "https://archive-api.open-meteo.com/v1/archive"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={start_datum.isoformat()}"
            f"&end_date={hist_ende.isoformat()}"
            "&hourly=shortwave_radiation,temperature_2m"
            "&timezone=Europe%2FBerlin"
        )
        fetch_data(url_hist)

    # Prognose ab heute
    if ende_datum >= heute_local:
        prog_start = max(start_datum, heute_local)

        url_prog = (
            "https://api.open-meteo.com/v1/forecast"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={prog_start.isoformat()}"
            f"&end_date={ende_datum.isoformat()}"
            "&hourly=shortwave_radiation,temperature_2m"
            "&timezone=Europe%2FBerlin"
        )
        fetch_data(url_prog)

    return ergebnisse


def interpoliere_stunde_auf_15min(val_aktuell, val_naechste, schritt):
    return val_aktuell + (val_naechste - val_aktuell) * (schritt / 4.0)


def berechne_leistung_physikalisch(
    ghi: float,
    temp: float,
    stunde: float,
    kwp: float,
    richtung: str,
    neigung_text: str
) -> float:
    """
    Vereinfachtes PV-Modell.
    Ergebnis: kWh pro 15-Minuten-Slot.

    Wichtig:
    Das ist keine perfekte PV-Simulation, aber deutlich sauberer als die alte Version.
    """
    if ghi <= 0 or kwp <= 0:
        return 0.0

    # Temperaturfaktor: -0,4 % pro °C über 25 °C
    eta_temp = 1 - 0.004 * (temp - 25)
    eta_temp = max(0.70, min(1.15, eta_temp))

    # Grobe Tageswinkel-Näherung
    omega = 15 * (stunde - 12.0)
    gamma_modul = AZIMUT_WINKEL.get(richtung, 0)

    # Ausrichtungsfaktor
    winkel_rad = np.radians(omega - gamma_modul)
    projektion = max(0.0, np.cos(winkel_rad))

    # Nord bekommt nur diffuse/grobe Reststrahlung
    if richtung == "Nord":
        projektion = 0.20 * max(0.0, np.cos(np.radians(omega)))

    # Dachneigung grob berücksichtigen
    try:
        neigung_zahl = float(
            str(neigung_text)
            .replace("°", "")
            .replace("Neigung", "")
            .replace(",", ".")
            .strip()
        )
    except Exception:
        neigung_zahl = 36.0

    neigungs_faktor = 1.0 - abs(neigung_zahl - 35.0) / 200.0
    neigungs_faktor = max(0.75, min(1.05, neigungs_faktor))

    # Performance Ratio
    PR = 0.82

    # 15-Minuten-Energie
    return (ghi / 1000.0) * kwp * projektion * neigungs_faktor * eta_temp * PR * 0.25


def baue_pv_kurve_aus_wetterdaten(
    wetter_dict: dict,
    start_datum_real: dt.date,
    start_uhrzeit_slot: int,
    n_slots: int,
    kwp: float,
    richtung: str,
    neigung_text: str
) -> np.ndarray:

    basis_dt = dt.datetime.combine(start_datum_real, dt.time()) + dt.timedelta(minutes=15 * start_uhrzeit_slot)

    kurve = np.zeros(n_slots)

    for i in range(n_slots):
        slot_dt = basis_dt + dt.timedelta(minutes=15 * i)

        tag = slot_dt.date()
        stunde_str = f"{slot_dt.hour:02d}:00"

        naechste_stunde_dt = slot_dt.replace(minute=0) + dt.timedelta(hours=1)
        naechster_tag = naechste_stunde_dt.date()
        naechste_stunde_str = f"{naechste_stunde_dt.hour:02d}:00"

        if tag not in wetter_dict or stunde_str not in wetter_dict[tag]:
            continue

        daten_aktuell = wetter_dict[tag][stunde_str]
        daten_naechste = wetter_dict.get(naechster_tag, {}).get(naechste_stunde_str, daten_aktuell)

        schritt = slot_dt.minute // 15

        ghi_interp = interpoliere_stunde_auf_15min(
            daten_aktuell["ghi"],
            daten_naechste["ghi"],
            schritt
        )

        temp_interp = interpoliere_stunde_auf_15min(
            daten_aktuell["temp"],
            daten_naechste["temp"],
            schritt
        )

        t_dezimal = slot_dt.hour + slot_dt.minute / 60.0

        kurve[i] = berechne_leistung_physikalisch(
            ghi=ghi_interp,
            temp=temp_interp,
            stunde=t_dezimal,
            kwp=kwp,
            richtung=richtung,
            neigung_text=neigung_text
        )

    return kurve


@st.cache_data(ttl=3600)
def suche_adressen(suchtext: str) -> list:
    if len(suchtext.strip()) < 4:
        return []

    try:
        url = "https://nominatim.openstreetmap.org/search"

        params = {
            "q": suchtext,
            "format": "json",
            "limit": 5,
            "addressdetails": 1,
            "countrycodes": "de,at,ch"
        }

        headers = {
            "User-Agent": "EMS-PV-Dashboard/1.0"
        }

        r = requests.get(url, params=params, headers=headers, timeout=8)
        r.raise_for_status()

        ergebnisse = []

        for item in r.json():
            a = item.get("address", {})

            teile = []

            if a.get("road"):
                road = a["road"]

                if a.get("house_number"):
                    road += " " + a["house_number"]

                teile.append(road)

            if a.get("postcode"):
                teile.append(a["postcode"])

            ort = a.get("city") or a.get("town") or a.get("village")
            if ort:
                teile.append(ort)

            if a.get("country"):
                teile.append(a["country"])

            anzeige = ", ".join(teile) if teile else item.get("display_name", "")[:80]

            ergebnisse.append({
                "anzeige": anzeige,
                "lat": float(item["lat"]),
                "lon": float(item["lon"])
            })

        return ergebnisse

    except Exception:
        return []


# =============================================================================
# PLOT-FUNKTIONEN
# =============================================================================
def plotly_line(df_dict, labels, titel, fokus_start=0, fokus_ende=95, ren_farben=None):
    fig = go.Figure()

    farben = ren_farben if ren_farben else [
        "#FF4B4B", "#FFA500", "#0055FF", "#00CC88", "#AA44FF", "#FF00AA"
    ]

    for i, (name, werte) in enumerate(df_dict.items()):
        fig.add_trace(go.Scatter(
            x=labels,
            y=werte,
            name=name,
            line=dict(color=farben[i % len(farben)], width=1.6),
            hovertemplate=f"<b>{name}</b><br>%{{x}}<br>%{{y:.3f}}<extra></extra>"
        ))

    f_start = max(0, fokus_start)
    f_ende = min(fokus_ende, len(labels) - 1)

    fig.update_layout(
        template=plotly_template,
        title=titel,
        height=290,
        margin=dict(l=10, r=10, t=35, b=10),
        hovermode="x unified",
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        xaxis=dict(
            rangeslider=dict(visible=True, thickness=0.08),
            type="category",
            range=[f_start, f_ende]
        ),
        dragmode="pan",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
    )

    grid_col = "rgba(255,255,255,0.1)" if st.session_state.theme_mode == "🌙 Dunkel" else "rgba(128,128,128,0.2)"

    fig.update_xaxes(showgrid=True, gridcolor=grid_col)
    fig.update_yaxes(showgrid=True, gridcolor=grid_col)

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"scrollZoom": True, "displayModeBar": False}
    )


def plotly_bar(werte, labels, name, farbe, fokus_start=0, fokus_ende=95):
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=labels,
        y=werte,
        name=name,
        marker_color=farbe,
        hovertemplate=f"<b>{name}</b><br>%{{x}}<br>%{{y:.3f}} kWh<extra></extra>"
    ))

    f_start = max(0, fokus_start)
    f_ende = min(fokus_ende, len(labels) - 1)

    fig.update_layout(
        template=plotly_template,
        height=300,
        margin=dict(l=10, r=10, t=25, b=10),
        xaxis=dict(
            rangeslider=dict(visible=True, thickness=0.08),
            type="category",
            range=[f_start, f_ende]
        ),
        dragmode="pan",
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"scrollZoom": True, "displayModeBar": False}
    )


def plotly_batteriespeicher(labels, soc, laden, entladen, fokus_start, fokus_ende):
    fig = go.Figure()

    fig.add_trace(go.Bar(
        x=labels,
        y=laden,
        name="Speicher lädt",
        marker_color="#2ecc71",
        hovertemplate="<b>Speicher lädt</b><br>%{x}<br>%{y:.3f} kWh<extra></extra>"
    ))

    fig.add_trace(go.Bar(
        x=labels,
        y=-entladen,
        name="Speicher entlädt",
        marker_color="#e67e22",
        hovertemplate="<b>Speicher entlädt</b><br>%{x}<br>%{y:.3f} kWh<extra></extra>"
    ))

    fig.add_trace(go.Scatter(
        x=labels,
        y=soc,
        name="Speicher-SoC (%)",
        mode="lines",
        yaxis="y2",
        line=dict(color="#3498db", width=3),
        hovertemplate="<b>Speicher-SoC</b><br>%{x}<br>%{y:.1f} %<extra></extra>"
    ))

    f_start = max(0, fokus_start)
    f_ende = min(fokus_ende, len(labels) - 1)

    fig.update_layout(
        template=plotly_template,
        height=360,
        margin=dict(l=10, r=10, t=35, b=10),
        barmode="relative",
        hovermode="x unified",
        xaxis=dict(
            rangeslider=dict(visible=True, thickness=0.08),
            type="category",
            range=[f_start, f_ende]
        ),
        yaxis=dict(title="Energie pro 15 min (kWh)"),
        yaxis2=dict(
            title="Ladestand Speicher (%)",
            overlaying="y",
            side="right",
            range=[0, 100]
        ),
        legend=dict(
            orientation="h",
            yanchor="bottom",
            y=1.02,
            xanchor="right",
            x=1
        ),
        paper_bgcolor="rgba(0,0,0,0)",
        plot_bgcolor="rgba(0,0,0,0)"
    )

    st.plotly_chart(
        fig,
        use_container_width=True,
        config={"scrollZoom": True, "displayModeBar": False}
    )


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
    wirkungsgrad_prozent
):
    n = len(pv_kwh)

    soc_verlauf = np.zeros(n)
    lade_verlauf = np.zeros(n)
    entlade_verlauf = np.zeros(n)
    netzbezug_verlauf = np.zeros(n)
    einspeisung_verlauf = np.zeros(n)

    if kapazitaet_kwh <= 0:
        netzbezug_verlauf = np.maximum(verbrauch_kwh - pv_kwh, 0)
        einspeisung_verlauf = np.maximum(pv_kwh - verbrauch_kwh, 0)
        return soc_verlauf, lade_verlauf, entlade_verlauf, netzbezug_verlauf, einspeisung_verlauf

    wirkungsgrad = max(0.01, wirkungsgrad_prozent / 100.0)
    speicher_energie = kapazitaet_kwh * start_soc_prozent / 100.0

    max_ladeenergie_slot = max_ladeleistung_kw * 0.25
    max_entladeenergie_slot = max_entladeleistung_kw * 0.25

    for i in range(n):
        saldo = pv_kwh[i] - verbrauch_kwh[i]

        if saldo > 0:
            freie_kapazitaet = kapazitaet_kwh - speicher_energie

            energie_zum_laden = min(
                saldo,
                max_ladeenergie_slot,
                freie_kapazitaet / wirkungsgrad
            )

            speicher_energie += energie_zum_laden * wirkungsgrad
            lade_verlauf[i] = energie_zum_laden

            einspeisung_verlauf[i] = max(0, saldo - energie_zum_laden)

        elif saldo < 0:
            strombedarf = abs(saldo)

            energie_aus_speicher = min(
                strombedarf,
                max_entladeenergie_slot,
                speicher_energie * wirkungsgrad
            )

            speicher_energie -= energie_aus_speicher / wirkungsgrad
            entlade_verlauf[i] = energie_aus_speicher

            netzbezug_verlauf[i] = max(0, strombedarf - energie_aus_speicher)

        speicher_energie = max(0, min(kapazitaet_kwh, speicher_energie))
        soc_verlauf[i] = speicher_energie / kapazitaet_kwh * 100

    return soc_verlauf, lade_verlauf, entlade_verlauf, netzbezug_verlauf, einspeisung_verlauf


# =============================================================================
# SIDEBAR
# =============================================================================
st.sidebar.header("🚗 Fahrzeug-Konfiguration")

soc_aktuell = st.sidebar.slider("Aktueller Ladestand Auto (SoC) in %", 0, 100, 20)
soc_ziel = st.sidebar.slider("Gewünschter Ziel-Ladestand Auto in %", 0, 100, 80)

auto_kapazitaet_kwh = st.sidebar.number_input(
    "Auto-Batteriekapazität (kWh)",
    min_value=10.0,
    max_value=150.0,
    value=50.0,
    step=5.0
)

ladeleistung_auto_kw = st.sidebar.number_input(
    "AC-Ladeleistung Auto (kW)",
    min_value=1.0,
    max_value=22.0,
    value=11.0,
    step=1.0
)

abfahrtszeit = st.sidebar.selectbox(
    "Geplante Abfahrtszeit",
    ["07:30", "08:00", "12:00", "16:00", "17:30", "20:00"]
)

ladestrategie = st.sidebar.radio(
    "Ladestrategie:",
    (
        "🌱 Öko / Preisoptimiert",
        "⚡ Sofort Schnellladen"
    )
)

st.sidebar.markdown("---")


# Standort nur bei Live-Modus
if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
    st.sidebar.header("📍 Anlagenstandort")

    suchtext = st.sidebar.text_input(
        "🔍 Adresse eingeben",
        value=st.session_state.anlage_adresse,
        key="adresse_suchfeld_sidebar"
    )

    st.session_state.anlage_adresse = suchtext

    if len(suchtext.strip()) >= 4:
        vorschlaege = suche_adressen(suchtext)

        if vorschlaege:
            st.sidebar.markdown("**📋 Gefundene Adressen:**")

            anzeige_liste = [v["anzeige"] for v in vorschlaege]

            auswahl = st.sidebar.radio(
                "Adresse wählen",
                anzeige_liste,
                label_visibility="collapsed",
                key="adresse_radio_sidebar"
            )

            if st.sidebar.button("✅ Diese Adresse übernehmen", key="btn_uebernehmen_sidebar"):
                gewaehlter = next(v for v in vorschlaege if v["anzeige"] == auswahl)

                st.session_state.anlage_lat = gewaehlter["lat"]
                st.session_state.anlage_lon = gewaehlter["lon"]
                st.session_state.anlage_adresse = gewaehlter["anzeige"]

                st.rerun()

        elif suchtext:
            st.sidebar.warning("⚠️ Keine Treffer.")

    st.sidebar.markdown(
        f"""
        <div style='background:#e8f4e8;border-radius:8px;padding:8px 10px;
        font-size:0.8rem;color:#1a4a1a;margin-top:6px;border-left:3px solid #2ea043'>
        <b>📡 Standort:</b><br>{st.session_state.anlage_adresse[:50]}<br>
        <span style='font-family:monospace'>
        📌 {st.session_state.anlage_lat:.4f}°N, {st.session_state.anlage_lon:.4f}°E
        </span>
        </div>
        """,
        unsafe_allow_html=True
    )

    anlage_lat = st.session_state.anlage_lat
    anlage_lon = st.session_state.anlage_lon

    st.sidebar.markdown("---")

else:
    anlage_lat = 51.71
    anlage_lon = 8.76


# Zeitraum
st.sidebar.header("📅 Zeitraum-Auswahl")

ansicht_modus = st.sidebar.radio(
    "Ansicht:",
    (
        "📆 Heute (24h ab jetzt)",
        "🗓️ Bestimmten Tag wählen",
        "📊 Monate wählen"
    ),
    key="ansicht_input_sidebar"
)

st.session_state.ansicht = ansicht_modus

datum_blockiert = False


if st.session_state.ansicht == "📆 Heute (24h ab jetzt)":
    uhrzeit_slot = (aktuelle_zeit.hour * 60 + aktuelle_zeit.minute) // 15

    profil_heute = klemme_datum_auf_2025(heute)

    start_slot = slot_von_datum(profil_heute, uhrzeit_slot) - 288
    n_slots = 288 + 96 + 672

    start_datum_profil = profil_heute - dt.timedelta(days=3)
    start_datum_real = heute - dt.timedelta(days=3)

    labels = zeitachse_erstellen(start_datum_real, uhrzeit_slot, n_slots)

    fokus_start = 288
    fokus_ende = 288 + 95

    ansicht_titel = f"Heute – {heute.strftime('%d.%m.%Y')}"

elif st.session_state.ansicht == "🗓️ Bestimmten Tag wählen":
    if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
        max_datum = heute + dt.timedelta(days=MAX_PROGNOSE_TAGE)
        st.sidebar.info(f"Live-Modus: Zukunft maximal {MAX_PROGNOSE_TAGE} Tage.")
    else:
        max_datum = dt.date(2030, 12, 31)

    gewaehltes_datum_real = st.sidebar.date_input(
        "Tag auswählen:",
        value=heute,
        min_value=dt.date(2020, 1, 1),
        max_value=max_datum
    )

    kat = datum_kategorie(gewaehltes_datum_real)

    if kat == "zu_weit" and st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
        datum_blockiert = True

    profil_datum = klemme_datum_auf_2025(gewaehltes_datum_real)

    start_slot = slot_von_datum(profil_datum, 0) - 288
    n_slots = 288 + 96 + 672

    start_datum_profil = profil_datum - dt.timedelta(days=3)
    start_datum_real = gewaehltes_datum_real - dt.timedelta(days=3)

    labels = zeitachse_erstellen(start_datum_real, 0, n_slots)

    fokus_start = 288
    fokus_ende = 288 + 95

    ansicht_titel = f"Tag: {gewaehltes_datum_real.strftime('%d.%m.%Y')}"

else:
    st.sidebar.markdown("**Monat auswählen:**")

    monat_name_wahl = st.sidebar.radio(
        "Wähle einen Monat:",
        list(MONATE.values()),
        index=heute.month - 1
    )

    gewaehlter_monat_nr = next(nr for nr, name in MONATE.items() if name == monat_name_wahl)

    monat_start = dt.date(2025, gewaehlter_monat_nr, 1)

    if gewaehlter_monat_nr == 12:
        monat_ende = dt.date(2025, 12, 31)
    else:
        monat_ende = dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)

    n_slots = ((monat_ende - monat_start).days + 1) * 96

    start_slot = slot_von_datum(monat_start, 0)

    start_datum_profil = monat_start
    start_datum_real = monat_start

    labels = zeitachse_erstellen(monat_start, 0, n_slots)

    fokus_start = 0
    fokus_ende = n_slots - 1

    ansicht_titel = f"{monat_name_wahl} ({n_slots // 96} Tage)"


if datum_blockiert:
    st.error("🚫 Datum liegt zu weit in der Zukunft.")
    st.stop()


st.sidebar.markdown("---")


# PV-Konfiguration
st.sidebar.header("☀️ PV-Konfiguration")

alle_neigungen = sorted(set(k[0] for k in daten["pv"].keys()))

if not alle_neigungen:
    st.error("❌ Keine PV-Daten in der Excel-Datei gefunden.")
    st.stop()

neigung_wahl = st.sidebar.selectbox("Dachneigung:", alle_neigungen)

kwp_sued = st.sidebar.number_input("kWp Süd", min_value=0.0, value=1.0, step=0.5)
kwp_ost = st.sidebar.number_input("kWp Ost", min_value=0.0, value=1.0, step=0.5)
kwp_nord = st.sidebar.number_input("kWp Nord", min_value=0.0, value=1.0, step=0.5)
kwp_west = st.sidebar.number_input("kWp West", min_value=0.0, value=1.0, step=0.5)

kwp_map = {
    "Süd": kwp_sued,
    "Ost": kwp_ost,
    "Nord": kwp_nord,
    "West": kwp_west
}


st.sidebar.markdown("---")
st.sidebar.header("🔋 Haus-Batteriespeicher")

speicher_aktiv = st.sidebar.checkbox("Batteriespeicher aktivieren", value=True)

speicher_kapazitaet_kwh = st.sidebar.selectbox(
    "Speicherkapazität (kWh)",
    [0.0, 5.0, 10.0, 15.0, 20.0, 50.0],
    index=2
)

speicher_start_soc = st.sidebar.slider("Start-Ladestand Speicher (%)", 0, 100, 50)

speicher_max_ladeleistung_kw = st.sidebar.number_input(
    "Max. Ladeleistung Speicher (kW)",
    min_value=0.0,
    max_value=20.0,
    value=5.0,
    step=0.5
)

speicher_max_entladeleistung_kw = st.sidebar.number_input(
    "Max. Entladeleistung Speicher (kW)",
    min_value=0.0,
    max_value=20.0,
    value=5.0,
    step=0.5
)

speicher_wirkungsgrad = st.sidebar.slider("Speicher-Wirkungsgrad (%)", 50, 100, 90)

start_button = st.sidebar.button("🔄 Optimierung berechnen")


# =============================================================================
# DATEN FÜR ZEITRAUM
# =============================================================================
hs_z = slots_fuer_zeitraum(daten["hs"], start_slot, n_slots)
waerme_z = slots_fuer_zeitraum(daten["waerme"], start_slot, n_slots)
preise_z = slots_fuer_zeitraum(daten["preise"], start_slot, n_slots)

verbrauch_gesamt_z = hs_z + waerme_z


# =============================================================================
# PV BERECHNEN
# =============================================================================
pv_richtungen = {}

if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
    datum_von = start_datum_real
    datum_bis = start_datum_real + dt.timedelta(days=int(np.ceil(n_slots / 96)) + 1)

    # API-Prognose begrenzen
    datum_bis = min(datum_bis, heute + dt.timedelta(days=MAX_PROGNOSE_TAGE))

    with st.spinner("🛰️ Lade Wetterdaten von Open-Meteo..."):
        wetter_dict = hole_open_meteo_live(anlage_lat, anlage_lon, datum_von, datum_bis)

    if not wetter_dict:
        st.warning("⚠️ Keine Wetterdaten erhalten. PV wird mit 0 kWh berechnet.")

    for (neigung, richtung), info in daten["pv"].items():
        if neigung != neigung_wahl:
            continue

        kwp = kwp_map.get(richtung, 0.0)

        if kwp > 0:
            pv_richtungen[richtung] = baue_pv_kurve_aus_wetterdaten(
                wetter_dict=wetter_dict,
                start_datum_real=start_datum_real,
                start_uhrzeit_slot=0 if st.session_state.ansicht != "📆 Heute (24h ab jetzt)" else ((aktuelle_zeit.hour * 60 + aktuelle_zeit.minute) // 15),
                n_slots=n_slots,
                kwp=kwp,
                richtung=richtung,
                neigung_text=neigung
            )

else:
    faktor = 1.2 if "Simulierte" in st.session_state.quelle else 1.0

    for (neigung, richtung), info in daten["pv"].items():
        if neigung != neigung_wahl:
            continue

        kwp = kwp_map.get(richtung, 0.0)

        if kwp > 0:
            base_kurve = slots_fuer_zeitraum(info["kurve"], start_slot, n_slots)
            pv_richtungen[richtung] = base_kurve * kwp * faktor


if pv_richtungen:
    pv_gesamt_z = np.sum(np.vstack(list(pv_richtungen.values())), axis=0)
else:
    pv_gesamt_z = np.zeros(n_slots)


# =============================================================================
# BATTERIE BERECHNEN
# =============================================================================
if speicher_aktiv:
    (
        speicher_soc_verlauf,
        speicher_laden,
        speicher_entladen,
        speicher_netzbezug,
        speicher_einspeisung
    ) = simuliere_batteriespeicher(
        pv_kwh=pv_gesamt_z,
        verbrauch_kwh=verbrauch_gesamt_z,
        kapazitaet_kwh=speicher_kapazitaet_kwh,
        start_soc_prozent=speicher_start_soc,
        max_ladeleistung_kw=speicher_max_ladeleistung_kw,
        max_entladeleistung_kw=speicher_max_entladeleistung_kw,
        wirkungsgrad_prozent=speicher_wirkungsgrad
    )
else:
    speicher_soc_verlauf = np.zeros(n_slots)
    speicher_laden = np.zeros(n_slots)
    speicher_entladen = np.zeros(n_slots)
    speicher_netzbezug = np.maximum(verbrauch_gesamt_z - pv_gesamt_z, 0)
    speicher_einspeisung = np.maximum(pv_gesamt_z - verbrauch_gesamt_z, 0)


# =============================================================================
# HAUPTBEREICH
# =============================================================================
st.success(f"✅ Datei geladen: `{gefundener_pfad}`")
st.subheader(f"📅 Analyse: {ansicht_titel}")

tag_hs = float(hs_z[fokus_start:fokus_ende + 1].sum())
tag_waerme = float(waerme_z[fokus_start:fokus_ende + 1].sum())
tag_pv = float(pv_gesamt_z[fokus_start:fokus_ende + 1].sum())
tag_verbrauch = tag_hs + tag_waerme
ueberschuss = tag_pv - tag_verbrauch

netzbezug_tag = float(speicher_netzbezug[fokus_start:fokus_ende + 1].sum())
einspeisung_tag = float(speicher_einspeisung[fokus_start:fokus_ende + 1].sum())


st.markdown(
    f"""
    <style>
    .metric-container {{
        background-color: {container_bg};
        border-radius: 10px;
        padding: 15px;
        text-align: center;
        box-shadow: 0 2px 5px rgba(0,0,0,0.1);
    }}
    .metric-value {{
        font-size: 2.1rem;
        font-weight: bold;
        color: {val_color};
        margin: 0;
    }}
    .metric-label {{
        font-size: 1.0rem;
        color: {label_color};
        margin: 5px 0 0 0;
    }}
    </style>
    """,
    unsafe_allow_html=True
)

c1, c2, c3, c4, c5, c6 = st.columns(6)

with c1:
    st.markdown(
        f'<div class="metric-container"><p class="metric-value">{tag_hs:.1f}</p><p class="metric-label">kWh Haushalt</p></div>',
        unsafe_allow_html=True
    )

with c2:
    st.markdown(
        f'<div class="metric-container"><p class="metric-value">{tag_waerme:.1f}</p><p class="metric-label">kWh Wärme</p></div>',
        unsafe_allow_html=True
    )

with c3:
    st.markdown(
        f'<div class="metric-container"><p class="metric-value">{tag_verbrauch:.1f}</p><p class="metric-label">kWh Verbrauch</p></div>',
        unsafe_allow_html=True
    )

with c4:
    st.markdown(
        f'<div class="metric-container"><p class="metric-value">{tag_pv:.1f}</p><p class="metric-label">kWh PV</p></div>',
        unsafe_allow_html=True
    )

with c5:
    color_bilanz = "#27ae60" if ueberschuss >= 0 else "#c0392b"
    bilanz_text = f"+{ueberschuss:.1f}" if ueberschuss >= 0 else f"{ueberschuss:.1f}"

    st.markdown(
        f'<div class="metric-container"><p class="metric-value" style="color:{color_bilanz};">{bilanz_text}</p><p class="metric-label">kWh Bilanz</p></div>',
        unsafe_allow_html=True
    )

with c6:
    st.markdown(
        f'<div class="metric-container"><p class="metric-value">{netzbezug_tag:.1f}</p><p class="metric-label">kWh Netzbezug</p></div>',
        unsafe_allow_html=True
    )


st.markdown("---")
st.subheader("🔮 Verbrauchsprofil")

c1, c2, c3 = st.columns(3)

with c1:
    plotly_line({"Haushalt": hs_z}, labels, "", fokus_start, fokus_ende, ["#FF4B4B"])

with c2:
    plotly_line({"Wärme": waerme_z}, labels, "", fokus_start, fokus_ende, ["#FFA500"])

with c3:
    plotly_line({"Gesamt": verbrauch_gesamt_z}, labels, "", fokus_start, fokus_ende, ["#0055FF"])


st.markdown("---")
st.subheader("☀️ PV-Erzeugung")

aktive_pv = {r: v for r, v in pv_richtungen.items() if v.sum() > 0}

if aktive_pv:
    pv_chart_dict = dict(aktive_pv)
    pv_chart_dict["Gesamt"] = pv_gesamt_z

    plotly_line(pv_chart_dict, labels, "", fokus_start, fokus_ende)
else:
    st.info("Keine PV-Erzeugung vorhanden. Prüfe kWp-Werte oder Wetterdaten.")


st.markdown("---")
st.subheader("🔋 Batteriespeicher")

if speicher_aktiv:
    plotly_batteriespeicher(
        labels=labels,
        soc=speicher_soc_verlauf,
        laden=speicher_laden,
        entladen=speicher_entladen,
        fokus_start=fokus_start,
        fokus_ende=fokus_ende
    )

    b1, b2, b3, b4, b5 = st.columns(5)

    with b1:
        st.metric(
            "In Speicher geladen",
            f"{speicher_laden[fokus_start:fokus_ende + 1].sum():.1f} kWh"
        )

    with b2:
        st.metric(
            "Aus Speicher genutzt",
            f"{speicher_entladen[fokus_start:fokus_ende + 1].sum():.1f} kWh"
        )

    with b3:
        st.metric(
            "SoC am Ende",
            f"{speicher_soc_verlauf[fokus_ende]:.1f} %"
        )

    with b4:
        st.metric(
            "Netzbezug nach Speicher",
            f"{netzbezug_tag:.1f} kWh"
        )

    with b5:
        st.metric(
            "Einspeisung",
            f"{einspeisung_tag:.1f} kWh"
        )

else:
    st.info("Batteriespeicher ist deaktiviert.")


st.markdown("---")
st.subheader("💶 Börsenstrompreise")
plotly_line({"ct/kWh": preise_z}, labels, "", fokus_start, fokus_ende, ["#AA44FF"])


# =============================================================================
# E-AUTO OPTIMIERUNG
# =============================================================================
if start_button:
    st.markdown("---")
    st.subheader("🏆 Optimierungsergebnis Elektroauto")

    energie_needed = max(
        0.0,
        (soc_ziel - soc_aktuell) / 100.0 * auto_kapazitaet_kwh
    )

    st.write(f"**Benötigte Ladeenergie:** {energie_needed:.1f} kWh")

    if energie_needed <= 0:
        st.success("✅ Ziel-SoC ist bereits erreicht. Es muss nicht geladen werden.")

    elif st.session_state.ansicht == "📊 Monate wählen":
        st.warning("Für die Monatsansicht wird keine konkrete Auto-Ladeoptimierung erstellt. Bitte Tagesansicht wählen.")

    else:
        ladeplan = np.zeros(n_slots)
        ladeenergie_slot = ladeleistung_auto_kw * 0.25

        slots_noetig = int(np.ceil(energie_needed / ladeenergie_slot))

        abfahrt_slot_tag = uhrzeit_zu_slot(abfahrtszeit)

        # Im Fokusbereich ist der betrachtete Tag von fokus_start bis fokus_start+95
        start_ladefenster = fokus_start

        # Wenn Abfahrt sehr früh ist, trotzdem bis zur Abfahrtszeit des betrachteten Tages
        ende_ladefenster = fokus_start + abfahrt_slot_tag

        if ende_ladefenster <= start_ladefenster:
            ende_ladefenster = fokus_start + 96

        moegliche_slots = np.arange(start_ladefenster, min(ende_ladefenster, n_slots))

        if len(moegliche_slots) == 0:
            st.error("Kein gültiges Ladefenster gefunden.")
        else:
            if "Sofort" in ladestrategie:
                gewaehlte_slots = moegliche_slots[:slots_noetig]

            else:
                # Öko/preisoptimiert:
                # Bewertungsfunktion:
                # niedriger Strompreis gut
                # PV-Überschuss gut
                pv_ueberschuss = np.maximum(pv_gesamt_z - verbrauch_gesamt_z, 0)

                preise_f = preise_z[moegliche_slots]
                pv_f = pv_ueberschuss[moegliche_slots]

                # Normalisieren
                preis_norm = (preise_f - np.min(preise_f)) / (np.ptp(preise_f) + 1e-9)
                pv_norm = (pv_f - np.min(pv_f)) / (np.ptp(pv_f) + 1e-9)

                # Score klein = gut
                # Preis zählt 60 %, PV-Überschuss 40 %
                score = 0.60 * preis_norm - 0.40 * pv_norm

                sort_idx = np.argsort(score)
                gewaehlte_slots = moegliche_slots[sort_idx[:slots_noetig]]

            rest = energie_needed

            for s in gewaehlte_slots:
                if rest <= 0:
                    break

                e = min(ladeenergie_slot, rest)
                ladeplan[s] = e
                rest -= e

            geladene_energie = ladeplan.sum()

            plotly_bar(
                ladeplan,
                labels,
                "Auto-Ladeenergie",
                "#00AA44",
                fokus_start,
                fokus_ende
            )

            st.metric("Geplante Auto-Ladeenergie", f"{geladene_energie:.1f} kWh")

            if geladene_energie + 1e-6 >= energie_needed:
                st.success("✅ Ziel-SoC wird innerhalb des Ladefensters erreicht.")
            else:
                st.warning(
                    f"⚠️ Ziel-SoC wird nicht vollständig erreicht. "
                    f"Es fehlen noch {energie_needed - geladene_energie:.1f} kWh."
                )

            kosten = np.sum(ladeplan * preise_z / 100.0)

            st.metric("Geschätzte Ladekosten", f"{kosten:.2f} €")

            neuer_verbrauch_mit_auto = verbrauch_gesamt_z + ladeplan
            neuer_netzbezug = np.maximum(neuer_verbrauch_mit_auto - pv_gesamt_z, 0)

            st.metric(
                "Netzbezug mit Auto ohne Speicher-Neuberechnung",
                f"{neuer_netzbezug[fokus_start:fokus_ende + 1].sum():.1f} kWh"
            )

            st.caption(
                "Hinweis: Die Auto-Ladung wird hier als zusätzlicher Ladeplan dargestellt. "
                "Für eine perfekte Gesamtbilanz müsste der Batteriespeicher anschließend "
                "mit Auto-Ladeplan nochmals neu simuliert werden."
            )


# =============================================================================
# JAHRESERTRAG PV AUS EXCEL
# =============================================================================
st.markdown("---")
st.subheader("📌 PV-Jahreswerte aus Excel")

jahreswerte = []

for (neigung, richtung), info in daten["pv"].items():
    if neigung == neigung_wahl:
        kwp = kwp_map.get(richtung, 0.0)
        jahreswerte.append({
            "Richtung": richtung,
            "kWp": kwp,
            "kWh/kWp/a laut Excel": info["jahres_kwh"],
            "geschätzter Jahresertrag": kwp * info["jahres_kwh"]
        })

if jahreswerte:
    df_jahr = pd.DataFrame(jahreswerte)
    st.dataframe(df_jahr, use_container_width=True)

    st.success(
        f"Geschätzter PV-Jahresertrag mit aktueller Konfiguration: "
        f"{df_jahr['geschätzter Jahresertrag'].sum():.1f} kWh/a"
    )
