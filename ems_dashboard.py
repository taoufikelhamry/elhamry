# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import os
import datetime as dt
import plotly.graph_objects as go
import requests

# Page Config
st.set_page_config(page_title="EMS Elektroauto", layout="wide")

aktuelle_zeit = dt.datetime.now()
heute         = aktuelle_zeit.date()
PROFIL_START = dt.date(2025, 1, 1)
PROFIL_ENDE  = dt.date(2025, 12, 31)
MAX_PROGNOSE_TAGE = 5

MONATE = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

# ── Session State Initialisierung ──────────────────────────────────────────────
if "quelle" not in st.session_state:
    st.session_state.quelle = "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)"
if "ansicht" not in st.session_state:
    st.session_state.ansicht = "📆 Heute (24h ab jetzt)"
if "anlage_lat" not in st.session_state:
    st.session_state.anlage_lat = 51.71
    st.session_state.anlage_lon = 8.76
    st.session_state.anlage_adresse = "Paderborn (Standard)"

def klemme_datum(d):
    try:
        return d.replace(year=2025)
    except ValueError:
        return dt.date(2025, d.month, 28)

def datum_kategorie(datum):
    delta = (datum - heute).days
    if datum < heute:
        return "vergangenheit"
    elif datum == heute:
        return "heute"
    elif 0 < delta <= MAX_PROGNOSE_TAGE:
        return "prognose"
    else:
        return "zu_weit"

# ── Kopfzeile ──────────────────────────────────────────────────────────────────
col_titel, col_zeit = st.columns([3, 1])
with col_titel:
    st.title("🔌 EMS – Steuerungslogik für ein Elektroauto")
    st.markdown("Ladeplanung basierend auf PV-Überschuss und Börsenstrompreisen.")
with col_zeit:
    st.markdown(
        f"<div style='text-align:right;font-size:1.1rem;font-weight:bold;"
        f"color:#4A90E2;padding-top:20px;'>🕒 {aktuelle_zeit.strftime('%d.%m.%Y — %H:%M')}</div>",
        unsafe_allow_html=True
    )
st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# DATENQUELLE AUSWAHL (MIT SESSION STATE)
# ══════════════════════════════════════════════════════════════════════════════
st.subheader("🌐 Datenquelle für PV-Erzeugung")
QUELLEN_OPTIONEN = [
    "📅 Historische Daten (Excel des letzten Jahres)",
    "☁️ Online-Wetterdaten (Simulierte API)",
    "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)"
]

daten_quelle = st.radio(
    "Bitte wähle die Basis für die Berechnung der Sonnenenergie:",
    QUELLEN_OPTIONEN,
    horizontal=True,
    key="quelle_input"
)
st.session_state.quelle = daten_quelle

if "Historische" in daten_quelle:
    st.info("📋 **Historische Daten:** Verwendet das Jahresprofil aus der Excel-Datei.")
elif "Simulierte" in daten_quelle:
    st.info("⚙️ **Simulierte API:** Wendet einen mathematischen Faktor auf die Basiskurve an.")
elif "Live" in daten_quelle:
    st.success("🛰️ **Live-Wetterdaten (Open-Meteo):** "
               "• Verwendet **Plane of Array Irradiance** (berücksichtigt Neigung und Ausrichtung) "
               "• Berücksichtigt **Temperatureffekte** (-0.4%/°C über 25°C) "
               "• Verwendet **glättete 15-Minuten-Daten** für realistischere Kurven")

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# DATEN LADEN
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def lade_excel():
    user_home = os.path.expanduser("~")
    suchpfade = [
        r"C:\Users\elham\Desktop\eg\Eingangsdaten.xlsx",
        r"C:\Users\elham\Desktop\eg\Eingangsdaten - Profile.xlsx",
        "Eingangsdaten.xlsx",
        "Eingangsdaten - Profile.xlsx",
        os.path.join(user_home, "Desktop", "Eingangsdaten.xlsx"),
        os.path.join(user_home, "Desktop", "psp", "Eingangsdaten.xlsx"),
        os.path.join(user_home, "OneDrive", "Desktop", "Eingangsdaten.xlsx"),
        os.path.join(user_home, "OneDrive - mail.uni-paderborn.de", "Bureau", "Eingangsdaten.xlsx"),
        os.path.join(user_home, "OneDrive - mail.uni-paderborn.de", "Bureau", "psp", "Eingangsdaten.xlsx"),
    ]
    pfad = None
    for p in suchpfade:
        if os.path.isfile(p):
            pfad = p
            break
    if pfad is None:
        return None, None
    
    xl     = pd.ExcelFile(pfad, engine="openpyxl")
    sheets = xl.sheet_names
    
    hs_sheet = next(s for s in sheets if "haushalt" in s.lower() or "strom" in s.lower())
    df_hs    = xl.parse(hs_sheet, header=None)
    hs_kwh   = pd.to_numeric(df_hs.iloc[2:, 1], errors="coerce").fillna(0).values * 0.25 / 1000
    
    waerme_sheet = next((s for s in sheets if "wärme" in s.lower() or "waerme" in s.lower()), None)
    if waerme_sheet:
        df_w       = xl.parse(waerme_sheet, header=None)
        rw         = pd.to_numeric(df_w.iloc[4:, 1], errors="coerce").fillna(0).values * (10140 / 1000)
        ww         = pd.to_numeric(df_w.iloc[4:, 2], errors="coerce").fillna(0).values * (2433  / 1000)
        waerme_kwh = rw + ww
    else:
        waerme_kwh = np.zeros(len(hs_kwh))
    
    preis_sheet = next(s for s in sheets if "preis" in s.lower() or "börse" in s.lower() or "boerse" in s.lower())
    df_p        = xl.parse(preis_sheet, header=None)
    preise      = pd.to_numeric(df_p.iloc[1:, 1], errors="coerce").fillna(0).values
    
    pv_dict = {}
    for sheet in sheets:
        if "pv" not in sheet.lower() and "neigung" not in sheet.lower():
            continue
        df_pv = xl.parse(sheet, header=None)
        for col_idx in range(1, 5):
            try:
                neigung  = str(df_pv.iloc[2, col_idx]).strip()
                richtung = str(df_pv.iloc[3, col_idx]).strip()
                if neigung in ("nan", "") or richtung in ("nan", ""):
                    continue
                jahres_kwh_kwp = float(df_pv.iloc[0, col_idx])
                kurve_kwh_kwp  = (pd.to_numeric(df_pv.iloc[5:, col_idx], errors="coerce")
                                  .fillna(0).values * 0.25 / 1000)
                pv_dict[(neigung, richtung)] = {
                    "kurve":      kurve_kwh_kwp,
                    "jahres_kwh": jahres_kwh_kwp
                }
            except Exception:
                continue
    return {"hs": hs_kwh, "waerme": waerme_kwh, "preise": preise, "pv": pv_dict}, pfad

daten, gefundener_pfad = lade_excel()

# ══════════════════════════════════════════════════════════════════════════════
# ERWEITERTE OPEN-METEO API (MIT NEIGUNG, TEMPERATUR, GLÄTTUNG)
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=3600)  
def hole_open_meteo_live_erweitert(lat: float, lon: float, start_datum: dt.date, ende_datum: dt.date, 
                                   neigung: float, richtung: str = "Süd") -> dict:
    """
    Holt Wetterdaten von Open-Meteo mit Berücksichtigung von:
    - Plane of Array Irradiance (berücksichtigt Neigung und Ausrichtung)
    - Temperatur (für Temperaturkorrektur der PV-Effizienz)
    - Direkte und diffuse Strahlung
    """
    heute_local = dt.date.today()
    ergebnisse = {}
    
    # Neigungswinkel konvertieren
    try:
        neigung_float = float(neigung)
    except (ValueError, TypeError):
        neigung_float = 30.0  # Default
    
    # Richtung in Azimut-Winkel umwandeln
    azimut_map = {
        "Süd": 180,
        "Südost": 135,
        "Südwest": 225,
        "Ost": 90,
        "West": 270,
        "Nord": 0
    }
    azimut = azimut_map.get(richtung, 180)
    
    # 1. Vergangenheit (Archive API)
    if start_datum < heute_local:
        hist_ende = min(ende_datum, heute_local - dt.timedelta(days=1))
        url_hist = (
            f"https://archive-api.open-meteo.com/v1/archive"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={start_datum.isoformat()}&end_date={hist_ende.isoformat()}"
            f"&hourly=shortwave_radiation,direct_radiation,diffuse_radiation,temperature_2m"
            f"&timezone=Europe%2FBerlin"
        )
        try:
            r = requests.get(url_hist, timeout=15)
            data = r.json()
            if "hourly" in data:
                times = data["hourly"]["time"]
               ghi     = data["hourly"].get("shortwave_radiation", [0] * len(times))
                temp    = data["hourly"].get("temperature_2m", [20] * len(times))
                
                for i, t_str in enumerate(times):
                    d = dt.date.fromisoformat(t_str[:10])
                    stunde_str = t_str[11:16]
                    ergebnisse.setdefault(d, {})[stunde_str] = {
                        "ghi": ghi[i] or 0.0,
                        "temp": temp[i] or 20.0
                    }
        except Exception as e:
            st.warning(f"⚠️ Historische API fehlgeschlagen: {e}")
            
    # 2. Heute & Zukunft (Forecast API) - MIT PLANE OF ARRAY
    if ende_datum >= heute_local:
        prog_start = max(start_datum, heute_local)
        url_prog = (
            f"https://api.open-meteo.com/v1/forecast"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={prog_start.isoformat()}&end_date={ende_datum.isoformat()}"
            f"&hourly=shortwave_radiation,direct_radiation,diffuse_radiation,temperature_2m"
            f"&azimuth={azimut}&surface_slope={neigung_float}"
            f"&timezone=Europe%2FBerlin"
        )
        try:
            r = requests.get(url_prog, timeout=15)
            data = r.json()
            if "hourly" in data:
                times = data["hourly"]["time"]
                ghi = data["hourly"].get("shortwave_radiation", [0] * len(times))
                direct = data["hourly"].get("direct_radiation", [0] * len(times))
                diffuse = data["hourly"].get("diffuse_radiation", [0] * len(times))
                temp = data["hourly"].get("temperature_2m", [20] * len(times))
                # Plane of Array Irradiance (POA) - das ist der entscheidende Wert!
                poa = data["hourly"].get("surface_sum", [0] * len(times))
                
                for i, t_str in enumerate(times):
                    d = dt.date.fromisoformat(t_str[:10])
                    stunde_str = t_str[11:16]
                    ergebnisse.setdefault(d, {})[stunde_str] = {
                        "ghi": ghi[i] or 0.0,
                        "poa": poa[i] or 0.0,  # WICHTIG: Einstrahlung auf die schräge Fläche
                        "direct": direct[i] or 0.0,
                        "diffuse": diffuse[i] or 0.0,
                        "temp": temp[i] or 20.0
                    }
        except Exception as e:
            st.warning(f"⚠️ Prognose-API fehlgeschlagen: {e}")
            
    return ergebnisse

# Temperatur-Korrekturfaktor für PV-Module
# Verlust von ~0.4% pro Grad über 25°C
def temp_korrektur_faktor(temperatur_celsius: float) -> float:
    """
    Berechnet den Temperatur-Korrekturfaktor für PV-Module.
    Bei 25°C = 1.0 (nominale Bedingungen)
    Bei 45°C ≈ 0.92 (8% Verlust)
    """
    # Typischer Temperaturkoeffizient für kristalline Silizium-Module
    temp_coeff = -0.004  # -0.4% pro °C
    t_ref = 25.0  # Referenztemperatur
    
    faktor = 1.0 + temp_coeff * (temperatur_celsius - t_ref)
    return max(0.7, min(1.1, faktor))  # Clamp zwischen 0.7 und 1.1

def strahlung_zu_kwh_15min_erweitert(strahlung_wm2: float, temperatur_c: float) -> float:
    """
    Wandelt Strahlung [W/m²] und Temperatur [°C] in kWh/15min um.
    Berücksichtigt Temperatur-Effekte.
    """
    # Basis-Kalkulation
    PR = 0.82  # Performance Ratio (Verluste durch Kabel, Wechselrichter, etc.)
    
    # Temperatur-Korrektur
    temp_faktor = temp_korrektur_faktor(temperatur_c)
    
    # Endwert: kWh pro 15 Minuten pro kWp
    return (strahlung_wm2 / 1000.0) * PR * temp_faktor * 0.25  # 0.25h = 15 Minuten

def baue_pv_kurve_aus_wetterdaten_erweitert(wetter_dict: dict, datum_start, n_slots: int, 
                                             kwp: float, neigung: float, richtung: str = "Süd") -> np.ndarray:
    """
    Erstellt ein 15-Minuten-Array aus Wetterdaten unter Berücksichtigung von:
    - Plane of Array Irradiance (wenn verfügbar)
    - Temperatur-Effekten
    - Glättung zwischen den Stundenwerten
    """
    if isinstance(datum_start, dt.datetime):
        basis_dt = datum_start.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        basis_dt = dt.datetime(datum_start.year, datum_start.month, datum_start.day, 0, 0, 0)
        
    kurve = np.zeros(n_slots)
    
    # Zwischenspeicher für vorherigen Wert (für Glättung)
    vorheriger_wert = 0
    
    for i in range(n_slots):
        slot_dt    = basis_dt + dt.timedelta(minutes=15 * i)
        tag        = slot_dt.date()
        stunde_str = f"{slot_dt.hour:02d}:00"
        
        if tag in wetter_dict and stunde_str in wetter_dict[tag]:
            daten = wetter_dict[tag][stunde_str]
            
            # Verwende POA (Plane of Array) wenn verfügbar, sonst GHI
            if "poa" in daten and daten["poa"] > 0:
                strahlung = daten["poa"]
            else:
                strahlung = daten.get("ghi", 0)
            
            # Temperatur aus den Daten
            temp = daten.get("temp", 20.0)
            
            # Berechne kWh für diesen Slot
            basis_wert = strahlung_zu_kwh_15min_erweitert(strahlung, temp) * kwp
            
            # Einfache Glättung (Durchschnitt mit vorherigem Wert)
            if vorheriger_wert > 0:
                glatteter_wert = (vorheriger_wert + basis_wert) / 2
            else:
                glatteter_wert = basis_wert
                
            kurve[i] = glatteter_wert
            vorheriger_wert = basis_wert
        else:
            # Kein Datenwert - interpolieren
            kurve[i] = vorheriger_wert
            
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
            "countrycodes": "de,at,ch",
        }
        headers = {"User-Agent": "EMS-PV-Dashboard/1.0"}
        r = requests.get(url, params=params, headers=headers, timeout=8)
        ergebnisse = []
        for item in r.json():
            a = item.get("address", {})
            teile = []
            if a.get("road"):       teile.append(a["road"])
            if a.get("house_number"): teile[-1] += " " + a["house_number"] if teile else teile.append(a["house_number"])
            if a.get("postcode"):   teile.append(a["postcode"])
            if a.get("city") or a.get("town") or a.get("village"):
                teile.append(a.get("city") or a.get("town") or a.get("village"))
            if a.get("country"):    teile.append(a["country"])
            anzeige = ", ".join(teile) if teile else item["display_name"][:80]
            ergebnisse.append({
                "anzeige": anzeige,
                "lat":     float(item["lat"]),
                "lon":     float(item["lon"]),
            })
        return ergebnisse
    except Exception:
        return []

# ══════════════════════════════════════════════════════════════════════════════
# HILFSFUNKTIONEN
# ══════════════════════════════════════════════════════════════════════════════
def slots_fuer_zeitraum(arr, start_slot, n_slots):
    total = len(arr)
    return np.array([arr[(start_slot + i) % total] for i in range(n_slots)])

def slot_von_datum(datum, uhrzeit_slot=0):
    delta = (datum - PROFIL_START).days
    return delta * 96 + uhrzeit_slot

def zeitachse_erstellen(start_datum, start_uhrzeit_slot, n_slots):
    start_dt = dt.datetime.combine(start_datum, dt.time()) + dt.timedelta(minutes=15 * start_uhrzeit_slot)
    start_dt = start_dt.replace(minute=(start_dt.minute // 15) * 15)
    return [(start_dt + dt.timedelta(minutes=15 * i)).strftime("%d.%m %H:%M") for i in range(n_slots)]

def plotly_line(df_dict, labels, titel, fokus_start=0, fokus_ende=95, ren_farben=None):
    fig = go.Figure()
    farben = ren_farben if ren_farben else ["#FF4B4B", "#FFA500", "#0055FF", "#00CC88", "#AA44FF", "#FF00AA"]
    for i, (name, werte) in enumerate(df_dict.items()):
        fig.add_trace(go.Scatter(
            x=labels, y=werte, name=name,
            line=dict(color=farben[i % len(farben)], width=1.5),
            hovertemplate=f"<b>{name}</b><br>%{{x}}<br>%{{y:.3f}}<extra></extra>"
        ))
    f_start = max(0, fokus_start)
    f_ende  = min(fokus_ende, len(labels) - 1)
    fig.update_layout(
        title=titel, height=280, margin=dict(l=10, r=10, t=35, b=10),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(rangeslider=dict(visible=True, thickness=0.08), type="category", range=[f_start, f_ende]),
        dragmode="pan", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    config = {"scrollZoom": True, "displayModeBar": False, "displaylogo": False}
    st.plotly_chart(fig, use_container_width=True, config=config)

def plotly_bar(werte, labels, name, farbe, fokus_start=0, fokus_ende=95):
    fig = go.Figure()
    fig.add_trace(go.Bar(
        x=labels, y=werte, name=name, marker_color=farbe,
        hovertemplate=f"<b>{name}</b><br>%{{x}}<br>%{{y:.3f}} kWh<extra></extra>"
    ))
    f_start = max(0, fokus_start)
    f_ende  = min(fokus_ende, len(labels) - 1)
    fig.update_layout(
        height=280, margin=dict(l=10, r=10, t=20, b=10),
        xaxis=dict(rangeslider=dict(visible=True, thickness=0.08), type="category", range=[f_start, f_ende]),
        dragmode="pan", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    config = {"scrollZoom": True, "displayModeBar": False, "displaylogo": False}
    st.plotly_chart(fig, use_container_width=True, config=config)

def hole_online_pv_daten(base_kurve, neigung, richtung, kwp):
    simulierter_wetter_faktor = 1.2
    return base_kurve * kwp * simulierter_wetter_faktor

# ══════════════════════════════════════════════════════════════════════════════
# SEITENLEISTE
# ══════════════════════════════════════════════════════════════════════════════
st.sidebar.header("🚗 Fahrzeug-Konfiguration")
soc_aktuell   = st.sidebar.slider("Aktueller Ladestand (SoC) in %", 0, 100, 20)
soc_ziel      = st.sidebar.slider("Gewünschter Ziel-Ladestand in %", 0, 100, 80)
abfahrtszeit  = st.sidebar.selectbox("Geplante Abfahrtszeit",
                                     ["07:30","08:00","12:00","16:00","17:30","20:00"])
ladestrategie = st.sidebar.radio("Ladestrategie:",
    ("🌱 Öko / Preisoptimiert (Lastverschiebung)", "⚡ Sofort Schnellladen"))

st.sidebar.markdown("---")

# ── Standort (Live-API) ───────────────────────────────────────────────────
if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
    st.sidebar.header("📍 Anlagenstandort")
    
    suchtext = st.sidebar.text_input(
        "🔍 Adresse eingeben",
        value=st.session_state.anlage_adresse,
        placeholder="z.B. Mersinweg 7, Paderborn",
        key="adresse_suchfeld"
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
                key="adresse_radio_live"
            )
            if st.sidebar.button("✅ Diese Adresse übernehmen", key="btn_uebernehmen_live"):
                gewaehlter = next(v for v in vorschlaege if v["anzeige"] == auswahl)
                st.session_state.anlage_lat     = gewaehlter["lat"]
                st.session_state.anlage_lon     = gewaehlter["lon"]
                st.session_state.anlage_adresse = gewaehlter["anzeige"]
                st.rerun()
        elif suchtext:
            st.sidebar.warning("⚠️ Keine Treffer.")
            
    kurz  = st.session_state.anlage_adresse
    lat_s = st.session_state.anlage_lat
    lon_s = st.session_state.anlage_lon
    st.sidebar.markdown(
        f"<div style='background:#e8f4e8;border-radius:8px;padding:8px 10px;"
        f"font-size:0.8rem;color:#1a4a1a;margin-top:6px;border-left:3px solid #2ea043'>"
        f"<b>📡 Standort:</b><br>{kurz[:50]}<br>"
        f"<span style='color:#444;font-family:monospace'>📌 {lat_s:.4f}°N, {lon_s:.4f}°E</span></div>",
        unsafe_allow_html=True
    )
    anlage_lat = lat_s
    anlage_lon = lon_s
    st.sidebar.markdown("---")
else:
    anlage_lat = 51.71
    anlage_lon = 8.76

# ── Zeitraum-Auswahl (MIT SESSION STATE) ─────────────────────────────────────
st.sidebar.header("📅 Zeitraum-Auswahl")
ansicht_modus = st.sidebar.radio(
    "Ansicht:",
    ("📆 Heute (24h ab jetzt)", "🗓️ Bestimmten Tag wählen", "📊 Monate wählen"),
    key="ansicht_input"
)
st.session_state.ansicht = ansicht_modus

ist_schaetzung = False
datum_blockiert = False

if st.session_state.ansicht == "📆 Heute (24h ab jetzt)":
    uhrzeit_slot  = (aktuelle_zeit.hour * 60 + aktuelle_zeit.minute) // 15
    profil_heute  = klemme_datum(heute)
    base_start_slot = slot_von_datum(profil_heute, uhrzeit_slot)
    start_slot      = base_start_slot - 288
    n_slots         = 288 + 96 + 672
    start_datum   = profil_heute - dt.timedelta(days=3)
    labels        = zeitachse_erstellen(start_datum, uhrzeit_slot, n_slots)
    fokus_start   = 288
    fokus_ende    = 288 + 95
    ansicht_titel = f"Heute – {aktuelle_zeit.strftime('%d.%m.%Y')}"
    gewaehltes_datum_real = heute
    if heute.year != 2025: ist_schaetzung = True
        
elif st.session_state.ansicht == "🗓️ Bestimmten Tag wählen":
    if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
        max_datum = heute + dt.timedelta(days=MAX_PROGNOSE_TAGE)
        st.sidebar.info(f"📅 Live-Modus: Zukunft max. {MAX_PROGNOSE_TAGE} Tage ({max_datum.strftime('%d.%m.%Y')})")
    else:
        max_datum = dt.date(2030, 12, 31)
        
    gewaehltes_datum = st.sidebar.date_input(
        "Tag auswählen:", value=heute,
        min_value=dt.date(2020, 1, 1), max_value=max_datum
    )
    gewaehltes_datum_real = gewaehltes_datum
    kat = datum_kategorie(gewaehltes_datum)
    if kat == "zu_weit" and st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
        st.sidebar.error(f"🚫 Datum zu weit! Max. {MAX_PROGNOSE_TAGE} Tage.")
        datum_blockiert = True
        
    profil_datum  = klemme_datum(gewaehltes_datum)
    base_start_slot = slot_von_datum(profil_datum, 0)
    start_slot      = base_start_slot - 288
    n_slots         = 288 + 96 + 672
    start_datum   = profil_datum - dt.timedelta(days=3)
    labels        = zeitachse_erstellen(start_datum, 0, n_slots)
    fokus_start   = 288
    fokus_ende    = 288 + 95
    if gewaehltes_datum.year != 2025:
        ist_schaetzung = True
        ansicht_titel = f"📅 {gewaehltes_datum.strftime('%d.%m.%Y')}"
    else:
        ansicht_titel = f"Tag: {gewaehltes_datum.strftime('%A, %d.%m.%Y')}"
else:  # Monate
    st.sidebar.markdown("**Monat auswählen:**")
    monat_name_wahl = st.sidebar.radio("Wähle einen Monat:", list(MONATE.values()), index=heute.month - 1)
    gewaehlter_monat_nr = next(nr for nr, name in MONATE.items() if name == monat_name_wahl)
    gewaehltes_datum_real = dt.date(2025, gewaehlter_monat_nr, 1)
    alle_slots_hs      = []
    alle_slots_waerme  = []
    alle_slots_preise  = []
    alle_labels        = []
    monat_start = dt.date(2025, gewaehlter_monat_nr, 1)
    monat_ende = dt.date(2025, gewaehlter_monat_nr, 31) if gewaehlter_monat_nr == 12 else dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)
    tage_im_monat = (monat_ende - monat_start).days + 1
    m_slots = tage_im_monat * 96
    m_start_slot = slot_von_datum(monat_start, 0)
    alle_slots_hs.extend(slots_fuer_zeitraum(daten["hs"],     m_start_slot, m_slots))
    alle_slots_waerme.extend(slots_fuer_zeitraum(daten["waerme"], m_start_slot, m_slots))
    alle_slots_preise.extend(slots_fuer_zeitraum(daten["preise"], m_start_slot, m_slots))
    alle_labels.extend(zeitachse_erstellen(monat_start, 0, m_slots))
    hs_z     = np.array(alle_slots_hs)
    waerme_z = np.array(alle_slots_waerme)
    preise_z = np.array(alle_slots_preise)
    n_slots  = len(hs_z)
    labels   = alle_labels
    start_slot = m_start_slot
    start_datum = monat_start
    anzahl_tage = n_slots // 96
    ansicht_titel = f"{monat_name_wahl} ({anzahl_tage} Tage)"
    fokus_start   = 0
    fokus_ende    = n_slots - 1

st.sidebar.markdown("---")

# ── PV-Konfiguration ──────────────────────────────────────────────────────────
st.sidebar.header("☀️ PV-Konfiguration")
alle_neigungen = sorted(set(k[0] for k in daten["pv"].keys())) if daten else []
neigung_wahl   = st.sidebar.selectbox("Dachneigung:", alle_neigungen) if alle_neigungen else ""
kwp_sued = st.sidebar.number_input("kWp Süd",  min_value=0.0, value=1.0, step=0.5)
kwp_ost  = st.sidebar.number_input("kWp Ost",  min_value=0.0, value=1.0, step=0.5)
kwp_nord = st.sidebar.number_input("kWp Nord", min_value=0.0, value=1.0, step=0.5)
kwp_west = st.sidebar.number_input("kWp West", min_value=0.0, value=1.0, step=0.5)
kwp_map  = {"Süd": kwp_sued, "Ost": kwp_ost, "Nord": kwp_nord, "West": kwp_west}

# WICHTIG: Neigungswinkel als float für API-Berechnung
neigung_float = 30.0  # Standardwert
if neigung_wahl:
    try:
        neigung_float = float(neigung_wahl)
    except (ValueError, TypeError):
        pass

start_button = st.sidebar.button("🔄 Optimierung berechnen")

# ══════════════════════════════════════════════════════════════════════════════
# HAUPTBEREICH
# ══════════════════════════════════════════════════════════════════════════════
if daten is None:
    st.error("❌ Excel-Datei nicht gefunden.")
    st.stop()
if datum_blockiert:
    st.error("🚫 **Datum zu weit in der Zukunft.** Bitte wählen Sie ein Datum innerhalb von 5 Tagen.")
    st.stop()

st.success(f"✅ Datei geladen: `{gefundener_pfad}`")
st.subheader(f"📅 Analyse: {ansicht_titel}")

if ist_schaetzung:
    profil_tag = klemme_datum(heute if st.session_state.ansicht == "📆 Heute (24h ab jetzt)" else gewaehltes_datum)
    st.info(f"📊 Daten-Schätzung auf Basis von {profil_tag}.")

# Datumskategorie-Badge
if st.session_state.ansicht != "📊 Monate wählen" and st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
    kat = datum_kategorie(gewaehltes_datum_real)
    kat_labels = {
        "vergangenheit": ("🕰️ Historische Echtdaten (ERA5)", "#1a6b3c", "#d4edda"),
        "heute":         ("📡 Heutige Live-Messung + Prognose", "#0d47a1", "#dce8ff"),
        "prognose":      ("🔮 Wetterprognose (< 5 Tage)", "#7b4f00", "#fff3cd"),
    }
    if kat in kat_labels:
        txt, col_text, col_bg = kat_labels[kat]
        st.markdown(
            f"<div style='background:{col_bg};border-radius:8px;padding:8px 14px;"
            f"color:{col_text};font-weight:600;display:inline-block;margin-bottom:8px'>"
            f"{txt}</div>", unsafe_allow_html=True
        )

# Verbrauchsdaten laden (für Tagesansichten)
if st.session_state.ansicht != "📊 Monate wählen":
    hs_z      = slots_fuer_zeitraum(daten["hs"],     start_slot, n_slots)
    waerme_z  = slots_fuer_zeitraum(daten["waerme"], start_slot, n_slots)
    preise_z  = slots_fuer_zeitraum(daten["preise"], start_slot, n_slots)

# ══════════════════════════════════════════════════════════════════════════════
# PV-ERZEUGUNG BERECHNEN (MIT ERWEITERTE FAKTOREN)
# ══════════════════════════════════════════════════════════════════════════════
pv_richtungen  = {}
wetter_geladen = False

if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)" and st.session_state.ansicht != "📊 Monate wählen":
    datum_von = start_datum
    datum_bis = start_datum + dt.timedelta(days=(n_slots // 96) + 1)
    datum_bis = min(datum_bis, heute + dt.timedelta(days=MAX_PROGNOSE_TAGE))
    
    with st.spinner("🛰️ Wetterdaten (inkl. Temperatur & Neigung) abgerufen..."):
        # Verwende die ERWEITERTE Funktion!
        wetter_dict = hole_open_meteo_live_erweitert(
            anlage_lat, anlage_lon, datum_von, datum_bis, 
            neigung_float, "Süd"  # Standard-Richtung
        )
    
    if wetter_dict:
        wetter_geladen = True
        tage_hist = sum(1 for d in wetter_dict if d < heute)
        tage_prog = sum(1 for d in wetter_dict if d >= heute)
        st.caption(f"✅ {len(wetter_dict)} Tage geladen: {tage_hist} Historisch, {tage_prog} Prognose")
        
        for (neigung, richtung), info in daten["pv"].items():
            if neigung != neigung_wahl: continue
            kwp = kwp_map.get(richtung, 0.0)
            if kwp > 0:
                # Verwende die ERWEITERTE Funktion mit Neigung und Richtung
                pv_kurve = baue_pv_kurve_aus_wetterdaten_erweitert(
                    wetter_dict, start_datum, n_slots, kwp, neigung_float, richtung
                )
                pv_richtungen[richtung] = pv_kurve
    else:
        st.warning("⚠️ Keine Live-Daten, Fallback zu Excel.")
        for (neigung, richtung), info in daten["pv"].items():
            if neigung == neigung_wahl:
                pv_richtungen[richtung] = slots_fuer_zeitraum(info["kurve"], start_slot, n_slots) * kwp_map.get(richtung, 0)
            
elif st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)" and st.session_state.ansicht == "📊 Monate wählen":
    st.info("ℹ️ Monatsberechnung mit Live-Daten läuft...")
    monat_start_l = dt.date(2025, gewaehlter_monat_nr, 1)
    monat_ende_l  = dt.date(2025, gewaehlter_monat_nr, 31) if gewaehlter_monat_nr == 12 else dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)
    with st.spinner("🛰️ Monatliche Wetterdaten geladen..."):
        wetter_dict = hole_open_meteo_live_erweitert(
            anlage_lat, anlage_lon, monat_start_l, monat_ende_l, 
            neigung_float, "Süd"
        )
    for (neigung, richtung), info in daten["pv"].items():
        if neigung == neigung_wahl:
            kwp = kwp_map.get(richtung, 0.0)
            if kwp > 0:
                pv_richtungen[richtung] = baue_pv_kurve_aus_wetterdaten_erweitert(
                    wetter_dict, monat_start_l, m_slots, kwp, neigung_float, richtung
                )
else:
    # Historisch / Simuliert
    for (neigung, richtung), info in daten["pv"].items():
        if neigung != neigung_wahl: continue
        kwp = kwp_map.get(richtung, 0.0)
        base_kurve = slots_fuer_zeitraum(info["kurve"], start_slot if st.session_state.ansicht != "📊 Monate wählen" else m_start_slot, 
                                         n_slots if st.session_state.ansicht != "📊 Monate wählen" else m_slots)
        if "Simulierte" in st.session_state.quelle:
            pv_richtungen[richtung] = base_kurve * kwp * 1.2
        else:
            pv_richtungen[richtung] = base_kurve * kwp

pv_gesamt_z = sum(pv_richtungen.values()) if pv_richtungen else np.zeros(n_slots)

# ── Metriken (VERGRÖSSERT) ───────────────────────────────────────────────────
if st.session_state.ansicht != "📊 Monate wählen":
    tag_hs     = float(hs_z[fokus_start : fokus_start + 96].sum())
    tag_waerm  = float(waerme_z[fokus_start : fokus_start + 96].sum())
    tag_pv     = float(pv_gesamt_z[fokus_start : fokus_start + 96].sum())
else:
    tag_hs     = float(hs_z.sum())
    tag_waerm  = float(waerme_z.sum())
    tag_pv     = float(pv_gesamt_z.sum())

ueberschuss = tag_pv - (tag_hs + tag_waerm)

st.markdown("""
<style>
.metric-container {
    background-color: #f0f2f6;
    border-radius: 10px;
    padding: 15px;
    text-align: center;
    box-shadow: 0 2px 5px rgba(0,0,0,0.1);
}
.metric-value {
    font-size: 2.5rem;
    font-weight: bold;
    color: #2c3e50;
    margin: 0;
}
.metric-label {
    font-size: 1.1rem;
    color: #7f8c8d;
    margin: 5px 0 0 0;
}
</style>
""", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
with c1:
    st.markdown(f"""
    <div class="metric-container">
        <p class="metric-value">{tag_hs:.1f} kWh</p>
        <p class="metric-label">📊 Haushaltsstrom</p>
    </div>
    """, unsafe_allow_html=True)
with c2:
    st.markdown(f"""
    <div class="metric-container">
        <p class="metric-value">{tag_waerm:.1f} kWh</p>
        <p class="metric-label">🔥 Wärmebedarf</p>
    </div>
    """, unsafe_allow_html=True)
with c3:
    st.markdown(f"""
    <div class="metric-container">
        <p class="metric-value">{tag_hs+tag_waerm:.1f} kWh</p>
        <p class="metric-label">⚡ Gesamt-Verbrauch</p>
    </div>
    """, unsafe_allow_html=True)
with c4:
    st.markdown(f"""
    <div class="metric-container">
        <p class="metric-value">{tag_pv:.1f} kWh</p>
        <p class="metric-label">☀️ PV-Erzeugung</p>
    </div>
    """, unsafe_allow_html=True)
with c5:
    val = f"+{ueberschuss:.1f} kWh" if ueberschuss >= 0 else f"{abs(ueberschuss):.1f} kWh"
    icon = "🟢" if ueberschuss >= 0 else "🔴"
    label = "PV-Überschuss" if ueberschuss >= 0 else "Netzstrombedarf"
    st.markdown(f"""
    <div class="metric-container">
        <p class="metric-value" style="color: {'#27ae60' if ueberschuss >= 0 else '#c0392b'}">{val}</p>
        <p class="metric-label">{icon} {label}</p>
    </div>
    """, unsafe_allow_html=True)

# ── TAGESÜBERSICHT (Monate)
if st.session_state.ansicht == "📊 Monate wählen" and anzahl_tage > 1:
    st.markdown("---")
    st.subheader("📊 Tagesübersicht")
    tages_daten = []
    monat_start_t = dt.date(2025, gewaehlter_monat_nr, 1)
    monat_ende_t  = dt.date(2025, gewaehlter_monat_nr, 31) if gewaehlter_monat_nr == 12 else dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)
    for i in range((monat_ende_t - monat_start_t).days + 1):
        tag_datum    = monat_start_t + dt.timedelta(days=i)
        tag_slot     = slot_von_datum(tag_datum, 0)
        t_hs         = float(slots_fuer_zeitraum(daten["hs"],     tag_slot, 96).sum())
        t_waerme     = float(slots_fuer_zeitraum(daten["waerme"], tag_slot, 96).sum())
        t_pv_val = sum(pv_richtungen[r][i*96:(i+1)*96].sum() for r in pv_richtungen) if pv_richtungen else 0.0
        tages_daten.append({
            "Datum": tag_datum.strftime("%a %d.%m"),
            "PV-Erzeugung": round(t_pv_val, 2),
            "Bilanz": round(t_pv_val - (t_hs + t_waerme), 2),
        })
    df_tage = pd.DataFrame(tages_daten).set_index("Datum")
    st.dataframe(df_tage, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# DIAGRAMME
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.subheader("🔮 Verbrauchsprofil")
st.caption("👉 Nutze den Regler unten oder ziehe das Diagramm: Links = Vergangenheit, Rechts = Zukunft")
c1, c2, c3 = st.columns(3)
with c1: plotly_line({"Haushaltsstrom": hs_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#FF4B4B"])
with c2: plotly_line({"Wärmebedarf": waerme_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#FFA500"])
with c3: plotly_line({"Gesamt-Verbrauch": hs_z + waerme_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#0055FF"])

st.markdown("---")
quelle_kurz = "Live-Wetter (Open-Meteo)" if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)" else "Historisch/Simuliert"
st.subheader(f"☀️ PV-Erzeugung – {neigung_wahl} – Quelle: {quelle_kurz}")
aktive_pv = {r: v for r, v in pv_richtungen.items() if v.sum() > 0}
if aktive_pv:
    pv_chart_dict = dict(aktive_pv)
    pv_chart_dict["Gesamt"] = pv_gesamt_z
    plotly_line(pv_chart_dict, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende)

st.markdown("---")
st.subheader("💶 Börsenstrompreise")
plotly_line({"ct/kWh": preise_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#AA44FF"])

# ── OPTIMIERUNG
if start_button:
    st.markdown("---")
    st.subheader("🏆 Optimierungsergebnis")
    kapazitaet_kwh = 50.0
    energie_needed = (soc_ziel - soc_aktuell) / 100.0 * kapazitaet_kwh
    st.write(f"**Benötigte Ladeenergie:** {energie_needed:.1f} kWh")
    if "Öko" in ladestrategie and st.session_state.ansicht != "📊 Monate wählen":
        ladeleistung_kw = 11.0
        slots_noetig    = max(1, int(np.ceil(energie_needed / (ladeleistung_kw * 0.25))))
        preise_heute    = preise_z[fokus_start : fokus_start + 96]
        guenstigste     = np.argsort(preise_heute)[:slots_noetig]
        ladeplan        = np.zeros(n_slots)
        ladeplan[fokus_start + guenstigste] = ladeleistung_kw * 0.25
        plotly_bar(ladeplan, labels, "Ladeenergie (kWh)", "#00AA44", fokus_start=fokus_start, fokus_ende=fokus_ende)