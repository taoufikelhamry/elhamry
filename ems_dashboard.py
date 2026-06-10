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
MAX_PROGNOSE_TAGE = 5

MONATE = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

# ── Session State Initialisierung ──────────────────────────────────────────────
if "quellem" not in st.session_state:
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
# DATENQUELLE AUSWAHL
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
               "Berücksichtigt GHI, Lufttemperatur (Temperaturkoeffizient) und stundengenaue Interpolation.")

st.markdown("---")

# ══════════════════════════════════════════════════════════════════════════════
# EXCEL DATEN LADEN
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
# OPEN-METEO API & PHYSIKALISCHE BERECHNUNG
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data(ttl=3600)  
def hole_open_meteo_live(lat: float, lon: float, start_datum: dt.date, ende_datum: dt.date) -> dict:
    heute_local = dt.date.today()
    ergebnisse = {}
    
    def fetch_data(url):
        try:
            r = requests.get(url, timeout=10)
            data = r.json()
            times = data["hourly"]["time"]
            rads  = data["hourly"]["shortwave_radiation"]
            temps = data["hourly"]["temperature_2m"] # Neu: Temperaturabruf
            for t_str, rad, temp in zip(times, rads, temps):
                d = dt.date.fromisoformat(t_str[:10])
                if d not in ergebnisse:
                    ergebnisse[d] = {}
                ergebnisse[d][t_str[11:16]] = {"ghi": rad or 0.0, "temp": temp or 15.0}
        except Exception as e:
            pass

    # Historische Daten (ERA5)
    if start_datum < heute_local:
        hist_ende = min(ende_datum, heute_local - dt.timedelta(days=1))
        url_hist = (
            f"https://archive-api.open-meteo.com/v1/archive"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={start_datum.isoformat()}&end_date={hist_ende.isoformat()}"
            f"&hourly=shortwave_radiation,temperature_2m"
            f"&timezone=Europe%2FBerlin"
        )
        fetch_data(url_hist)
            
    # Prognose Daten
    if ende_datum >= heute_local:
        prog_start = max(start_datum, heute_local)
        url_prog = (
            f"https://api.open-meteo.com/v1/forecast"
            f"?latitude={lat}&longitude={lon}"
            f"&start_date={prog_start.isoformat()}&end_date={ende_datum.isoformat()}"
            f"&hourly=shortwave_radiation,temperature_2m"
            f"&timezone=Europe%2FBerlin"
        )
        fetch_data(url_prog)
            
    return ergebnisse

# Neu: Geometrische Projektion anstelle statischer Faktoren
AZIMUT_WINKEL = {
    "Süd":  0,
    "Ost":  -90,
    "West": 90,
    "Nord": 180,
}

def interpoliere_stunde_auf_15min(val_aktuell, val_naechste, schritt):
    # Lineare Interpolation (0=0min, 1=15min, 2=30min, 3=45min)
    return val_aktuell + (val_naechste - val_aktuell) * (schritt / 4.0)

def berechne_leistung_physikalisch(ghi: float, temp: float, stunde: float, kwp: float, richtung: str) -> float:
    if ghi <= 0: return 0.0
    
    # Temperaturkorrektur (STC = 25°C, Gamma = -0.4%/K)
    eta_temp = 1 - 0.004 * (temp - 25)
    
    # Geometrische Näherung: Stundenwinkel (12:00 = 0°)
    omega = 15 * (stunde - 12)
    gamma_m = AZIMUT_WINKEL.get(richtung, 0)
    
    # Kosinus des Differenzwinkels approximiert den zeitlichen Verlauf
    winkel_rad = np.radians(omega - gamma_m)
    projektion = max(0, np.cos(winkel_rad))
    
    # Norddächer erhalten stark reduzierte indirekte Strahlung
    if richtung == "Nord": projektion = 0.2
    
    e_poa = ghi * projektion
    PR = 0.82
    
    return (e_poa / 1000.0) * kwp * eta_temp * PR * 0.25

def baue_pv_kurve_aus_wetterdaten(wetter_dict: dict, datum_start, n_slots: int, kwp: float, richtung: str = "Süd") -> np.ndarray:
    if isinstance(datum_start, dt.datetime):
        basis_dt = datum_start.replace(hour=0, minute=0, second=0, microsecond=0)
    else:
        basis_dt = dt.datetime(datum_start.year, datum_start.month, datum_start.day, 0, 0, 0)
        
    kurve = np.zeros(n_slots)
    
    for i in range(n_slots):
        slot_dt = basis_dt + dt.timedelta(minutes=15 * i)
        tag = slot_dt.date()
        stunde_str = f"{slot_dt.hour:02d}:00"
        
        naechste_stunde_dt = slot_dt.replace(minute=0) + dt.timedelta(hours=1)
        naechster_tag = naechste_stunde_dt.date()
        naechste_stunde_str = f"{naechste_stunde_dt.hour:02d}:00"
        
        if tag in wetter_dict and stunde_str in wetter_dict[tag]:
            daten_aktuell = wetter_dict[tag][stunde_str]
            daten_naechste = wetter_dict.get(naechster_tag, {}).get(naechste_stunde_str, daten_aktuell)
            
            schritt = (slot_dt.minute // 15)
            ghi_interp = interpoliere_stunde_auf_15min(daten_aktuell["ghi"], daten_naechste["ghi"], schritt)
            temp_interp = interpoliere_stunde_auf_15min(daten_aktuell["temp"], daten_naechste["temp"], schritt)
            
            t_dezimal = slot_dt.hour + (slot_dt.minute / 60.0)
            kurve[i] = berechne_leistung_physikalisch(ghi_interp, temp_interp, t_dezimal, kwp, richtung)
            
    return kurve

@st.cache_data(ttl=3600)
def suche_adressen(suchtext: str) -> list:
    if len(suchtext.strip()) < 4:
        return []
    try:
        url = "https://nominatim.openstreetmap.org/search"
        params = {"q": suchtext, "format": "json", "limit": 5, "addressdetails": 1, "countrycodes": "de,at,ch"}
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
            ergebnisse.append({"anzeige": anzeige, "lat": float(item["lat"]), "lon": float(item["lon"])})
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
    st.plotly_chart(fig, use_container_width=True, config={"scrollZoom": True, "displayModeBar": False})

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
    st.plotly_chart(fig, use_container_width=True, config={"scrollZoom": True, "displayModeBar": False})

# ══════════════════════════════════════════════════════════════════════════════
# SEITENLEISTE
# ══════════════════════════════════════════════════════════════════════════════
st.sidebar.header("🚗 Fahrzeug-Konfiguration")
soc_aktuell   = st.sidebar.slider("Aktueller Ladestand (SoC) in %", 0, 100, 20)
soc_ziel      = st.sidebar.slider("Gewünschter Ziel-Ladestand in %", 0, 100, 80)
abfahrtszeit  = st.sidebar.selectbox("Geplante Abfahrtszeit", ["07:30","08:00","12:00","16:00","17:30","20:00"])
ladestrategie = st.sidebar.radio("Ladestrategie:", ("🌱 Öko / Preisoptimiert (Lastverschiebung)", "⚡ Sofort Schnellladen"))

st.sidebar.markdown("---")

if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
    st.sidebar.header("📍 Anlagenstandort")
    suchtext = st.sidebar.text_input("🔍 Adresse eingeben", value=st.session_state.anlage_adresse, key="adresse_suchfeld_sidebar")
    st.session_state.anlage_adresse = suchtext
    
    if len(suchtext.strip()) >= 4:
        vorschlaege = suche_adressen(suchtext)
        if vorschlaege:
            st.sidebar.markdown("**📋 Gefundene Adressen:**")
            anzeige_liste = [v["anzeige"] for v in vorschlaege]
            auswahl = st.sidebar.radio("Adresse wählen", anzeige_liste, label_visibility="collapsed", key="adresse_radio_sidebar")
            if st.sidebar.button("✅ Diese Adresse übernehmen", key="btn_uebernehmen_sidebar"):
                gewaehlter = next(v for v in vorschlaege if v["anzeige"] == auswahl)
                st.session_state.anlage_lat = gewaehlter["lat"]
                st.session_state.anlage_lon = gewaehlter["lon"]
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
    anlage_lat, anlage_lon = 51.71, 8.76

# Zeitraum-Auswahl
st.sidebar.header("📅 Zeitraum-Auswahl")
ansicht_modus = st.sidebar.radio("Ansicht:", ("📆 Heute (24h ab jetzt)", "🗓️ Bestimmten Tag wählen", "📊 Monate wählen"), key="ansicht_input_sidebar")
st.session_state.ansicht = ansicht_modus

ist_schaetzung, datum_blockiert = False, False

if st.session_state.ansicht == "📆 Heute (24h ab jetzt)":
    uhrzeit_slot  = (aktuelle_zeit.hour * 60 + aktuelle_zeit.minute) // 15
    profil_heute  = klemme_datum(heute)
    base_start_slot = slot_von_datum(profil_heute, uhrzeit_slot)
    start_slot      = base_start_slot - 288
    n_slots         = 288 + 96 + 672
    start_datum   = profil_heute - dt.timedelta(days=3)
    labels        = zeitachse_erstellen(start_datum, uhrzeit_slot, n_slots)
    fokus_start, fokus_ende = 288, 288 + 95
    ansicht_titel = f"Heute – {aktuelle_zeit.strftime('%d.%m.%Y')}"
    gewaehltes_datum_real = heute
    if heute.year != 2025: ist_schaetzung = True
        
elif st.session_state.ansicht == "🗓️ Bestimmten Tag wählen":
    if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
        max_datum = heute + dt.timedelta(days=MAX_PROGNOSE_TAGE)
        st.sidebar.info(f"📅 Live-Modus: Zukunft max. {MAX_PROGNOSE_TAGE} Tage.")
    else:
        max_datum = dt.date(2030, 12, 31)
        
    gewaehltes_datum = st.sidebar.date_input("Tag auswählen:", value=heute, min_value=dt.date(2020, 1, 1), max_value=max_datum)
    gewaehltes_datum_real = gewaehltes_datum
    kat = datum_kategorie(gewaehltes_datum)
    if kat == "zu_weit" and st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
        datum_blockiert = True
        
    profil_datum  = klemme_datum(gewaehltes_datum)
    base_start_slot = slot_von_datum(profil_datum, 0)
    start_slot      = base_start_slot - 288
    n_slots         = 288 + 96 + 672
    start_datum   = profil_datum - dt.timedelta(days=3)
    labels        = zeitachse_erstellen(start_datum, 0, n_slots)
    fokus_start, fokus_ende = 288, 288 + 95
    ist_schaetzung = gewaehltes_datum.year != 2025
    ansicht_titel = f"Tag: {gewaehltes_datum.strftime('%d.%m.%Y')}"

else:
    st.sidebar.markdown("**Monat auswählen:**")
    monat_name_wahl = st.sidebar.radio("Wähle einen Monat:", list(MONATE.values()), index=heute.month - 1)
    gewaehlter_monat_nr = next(nr for nr, name in MONATE.items() if name == monat_name_wahl)
    gewaehltes_datum_real = dt.date(2025, gewaehlter_monat_nr, 1)
    
    monat_start = dt.date(2025, gewaehlter_monat_nr, 1)
    monat_ende = dt.date(2025, gewaehlter_monat_nr, 31) if gewaehlter_monat_nr == 12 else dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)
    m_slots = ((monat_ende - monat_start).days + 1) * 96
    m_start_slot = slot_von_datum(monat_start, 0)
    
    hs_z     = np.array(slots_fuer_zeitraum(daten["hs"], m_start_slot, m_slots))
    waerme_z = np.array(slots_fuer_zeitraum(daten["waerme"], m_start_slot, m_slots))
    preise_z = np.array(slots_fuer_zeitraum(daten["preise"], m_start_slot, m_slots))
    labels   = zeitachse_erstellen(monat_start, 0, m_slots)
    
    n_slots, start_slot, start_datum = m_slots, m_start_slot, monat_start
    ansicht_titel = f"{monat_name_wahl} ({n_slots // 96} Tage)"
    fokus_start, fokus_ende = 0, n_slots - 1

st.sidebar.markdown("---")

# PV-Konfiguration
st.sidebar.header("☀️ PV-Konfiguration")
alle_neigungen = sorted(set(k[0] for k in daten["pv"].keys())) if daten else []
neigung_wahl   = st.sidebar.selectbox("Dachneigung:", alle_neigungen) if alle_neigungen else ""
kwp_sued = st.sidebar.number_input("kWp Süd",  min_value=0.0, value=1.0, step=0.5)
kwp_ost  = st.sidebar.number_input("kWp Ost",  min_value=0.0, value=1.0, step=0.5)
kwp_nord = st.sidebar.number_input("kWp Nord", min_value=0.0, value=1.0, step=0.5)
kwp_west = st.sidebar.number_input("kWp West", min_value=0.0, value=1.0, step=0.5)
kwp_map  = {"Süd": kwp_sued, "Ost": kwp_ost, "Nord": kwp_nord, "West": kwp_west}
start_button = st.sidebar.button("🔄 Optimierung berechnen")

# ══════════════════════════════════════════════════════════════════════════════
# HAUPTBEREICH
# ══════════════════════════════════════════════════════════════════════════════
if daten is None:
    st.error("❌ Excel-Datei nicht gefunden.")
    st.stop()
if datum_blockiert:
    st.error("🚫 **Datum zu weit in der Zukunft.**")
    st.stop()

st.success(f"✅ Datei geladen: `{gefundener_pfad}`")
st.subheader(f"📅 Analyse: {ansicht_titel}")

if st.session_state.ansicht != "📊 Monate wählen":
    hs_z      = slots_fuer_zeitraum(daten["hs"],     start_slot, n_slots)
    waerme_z  = slots_fuer_zeitraum(daten["waerme"], start_slot, n_slots)
    preise_z  = slots_fuer_zeitraum(daten["preise"], start_slot, n_slots)

# PV-ERZEUGUNG BERECHNEN
pv_richtungen  = {}

if st.session_state.quelle == "🌤️ Live-Wetterdaten & Prognose (Open-Meteo API)":
    datum_von = start_datum
    datum_bis = min(start_datum + dt.timedelta(days=(n_slots // 96) + 1), heute + dt.timedelta(days=MAX_PROGNOSE_TAGE))
    
    with st.spinner("🛰️ Lade Wetterdaten (GHI & Temperatur)..."):
        wetter_dict = hole_open_meteo_live(anlage_lat, anlage_lon, datum_von, datum_bis)
    
    if wetter_dict:
        for (neigung, richtung), info in daten["pv"].items():
            if neigung != neigung_wahl: continue
            kwp = kwp_map.get(richtung, 0.0)
            if kwp > 0:
                pv_richtungen[richtung] = baue_pv_kurve_aus_wetterdaten(wetter_dict, start_datum, n_slots, kwp, richtung)
else:
    for (neigung, richtung), info in daten["pv"].items():
        if neigung != neigung_wahl: continue
        kwp = kwp_map.get(richtung, 0.0)
        base_kurve = slots_fuer_zeitraum(info["kurve"], start_slot, n_slots)
        pv_richtungen[richtung] = base_kurve * kwp * (1.2 if "Simulierte" in st.session_state.quelle else 1.0)

pv_gesamt_z = sum(pv_richtungen.values()) if pv_richtungen else np.zeros(n_slots)

# Metriken
tag_hs     = float(hs_z[fokus_start:fokus_ende+1].sum())
tag_waerm  = float(waerme_z[fokus_start:fokus_ende+1].sum())
tag_pv     = float(pv_gesamt_z[fokus_start:fokus_ende+1].sum())
ueberschuss = tag_pv - (tag_hs + tag_waerm)

st.markdown("""
<style>
.metric-container { background-color: #f0f2f6; border-radius: 10px; padding: 15px; text-align: center; box-shadow: 0 2px 5px rgba(0,0,0,0.1); }
.metric-value { font-size: 2.5rem; font-weight: bold; color: #2c3e50; margin: 0; }
.metric-label { font-size: 1.1rem; color: #7f8c8d; margin: 5px 0 0 0; }
</style>
""", unsafe_allow_html=True)

c1, c2, c3, c4, c5 = st.columns(5)
with c1: st.markdown(f'<div class="metric-container"><p class="metric-value">{tag_hs:.1f} kWh</p><p class="metric-label">📊 Haushalt</p></div>', unsafe_allow_html=True)
with c2: st.markdown(f'<div class="metric-container"><p class="metric-value">{tag_waerm:.1f} kWh</p><p class="metric-label">🔥 Wärme</p></div>', unsafe_allow_html=True)
with c3: st.markdown(f'<div class="metric-container"><p class="metric-value">{tag_hs+tag_waerm:.1f} kWh</p><p class="metric-label">⚡ Gesamt</p></div>', unsafe_allow_html=True)
with c4: st.markdown(f'<div class="metric-container"><p class="metric-value">{tag_pv:.1f} kWh</p><p class="metric-label">☀️ PV</p></div>', unsafe_allow_html=True)
with c5:
    val, icon = (f"+{ueberschuss:.1f}", "🟢") if ueberschuss >= 0 else (f"{abs(ueberschuss):.1f}", "🔴")
    st.markdown(f'<div class="metric-container"><p class="metric-value" style="color: {"#27ae60" if ueberschuss >= 0 else "#c0392b"}">{val} kWh</p><p class="metric-label">{icon} Bilanz</p></div>', unsafe_allow_html=True)

# Diagramme
st.markdown("---")
st.subheader("🔮 Verbrauchsprofil")
c1, c2, c3 = st.columns(3)
with c1: plotly_line({"Haushalt": hs_z}, labels, "", fokus_start, fokus_ende, ["#FF4B4B"])
with c2: plotly_line({"Wärme": waerme_z}, labels, "", fokus_start, fokus_ende, ["#FFA500"])
with c3: plotly_line({"Gesamt": hs_z + waerme_z}, labels, "", fokus_start, fokus_ende, ["#0055FF"])

st.markdown("---")
st.subheader(f"☀️ PV-Erzeugung")
aktive_pv = {r: v for r, v in pv_richtungen.items() if v.sum() > 0}
if aktive_pv:
    pv_chart_dict = dict(aktive_pv)
    pv_chart_dict["Gesamt"] = pv_gesamt_z
    plotly_line(pv_chart_dict, labels, "", fokus_start, fokus_ende)

st.markdown("---")
st.subheader("💶 Börsenstrompreise")
plotly_line({"ct/kWh": preise_z}, labels, "", fokus_start, fokus_ende, ["#AA44FF"])

# Optimierung
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
        plotly_bar(ladeplan, labels, "Ladeenergie (kWh)", "#00AA44", fokus_start, fokus_ende)