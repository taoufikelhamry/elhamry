# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import numpy as np
import os
import datetime as dt
import plotly.graph_objects as go

st.set_page_config(page_title="EMS Elektroauto", layout="wide")

aktuelle_zeit = dt.datetime.now()
heute         = aktuelle_zeit.date()

PROFIL_START = dt.date(2025, 1, 1)
PROFIL_ENDE  = dt.date(2025, 12, 31)

MONATE = {
    1: "Januar", 2: "Februar", 3: "März", 4: "April",
    5: "Mai", 6: "Juni", 7: "Juli", 8: "August",
    9: "September", 10: "Oktober", 11: "November", 12: "Dezember"
}

def klemme_datum(d):
    try:
        return d.replace(year=2025)
    except ValueError:
        return dt.date(2025, d.month, 28)

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

# ══════════════════════════════════════════════════════════════════════════════
# DATEN LADEN
# ══════════════════════════════════════════════════════════════════════════════
@st.cache_data
def lade_excel():
    user_home = os.path.expanduser("~")
    suchpfade = [
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
            x=labels,
            y=werte,
            name=name,
            line=dict(color=farben[i % len(farben)], width=1.5),
            hovertemplate=f"<b>{name}</b><br>%{{x}}<br>%{{y:.3f}}<extra></extra>"
        ))
    
    f_start = max(0, fokus_start)
    f_ende  = min(fokus_ende, len(labels) - 1)
    
    fig.update_layout(
        title=titel, height=280, margin=dict(l=10, r=10, t=35, b=10),
        hovermode="x unified",
        legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1),
        xaxis=dict(
            rangeslider=dict(visible=True, thickness=0.08),
            type="category",
            range=[f_start, f_ende]
        ),
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
        xaxis=dict(
            rangeslider=dict(visible=True, thickness=0.08),
            type="category",
            range=[f_start, f_ende]
        ),
        dragmode="pan", paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
    )
    fig.update_xaxes(showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    fig.update_yaxes(showgrid=True, gridcolor="rgba(128,128,128,0.2)")
    
    config = {"scrollZoom": True, "displayModeBar": False, "displaylogo": False}
    st.plotly_chart(fig, use_container_width=True, config=config)

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

# ── Zeitraum-Auswahl ──────────────────────────────────────────────────────────
st.sidebar.header("📅 Zeitraum-Auswahl")
ansicht_modus = st.sidebar.radio("Ansicht:",
    ("📆 Heute (24h ab jetzt)", "🗓️ Bestimmten Tag wählen", "📊 Monate wählen"))

ist_schaetzung = False

if ansicht_modus == "📆 Heute (24h ab jetzt)":
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
    if heute.year != 2025:
        ist_schaetzung = True

elif ansicht_modus == "🗓️ Bestimmten Tag wählen":
    gewaehltes_datum = st.sidebar.date_input(
        "Tag auswählen:", value=heute,
        min_value=dt.date(2020, 1, 1), max_value=dt.date(2030, 12, 31)
    )
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

else:  # Monate wählen
    st.sidebar.markdown("**Monat auswählen:**")
    
    # KORREKTUR: Radio-Button statt Checkboxen, damit nur ein einziger Monat gewählt werden kann
    monat_name_wahl = st.sidebar.radio("Wähle einen Monat:", list(MONATE.values()), index=heute.month - 1)
    
    # Ermittle die Nummer des gewählten Monats
    gewaehlter_monat_nr = next(nr for nr, name in MONATE.items() if name == monat_name_wahl)

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
    
    # KORREKTUR: Fokus-Fenster umfasst jetzt exakt den kompletten Monat (alle Slots von Anfang bis Ende)
    fokus_start   = 0
    fokus_ende    = n_slots - 1

st.sidebar.markdown("---")

# ── PV-Konfiguration ──────────────────────────────────────────────────────────
st.sidebar.header("☀️ PV-Konfiguration")
alle_neigungen = sorted(set(k[0] for k in daten["pv"].keys())) if daten else []
neigung_wahl   = st.sidebar.selectbox("Dachneigung:", alle_neigungen) if alle_neigungen else ""

kwp_sued = st.sidebar.number_input("Installierte kWp – Süd",  min_value=0.0, value=1.0, step=0.5)
kwp_ost  = st.sidebar.number_input("Installierte kWp – Ost",  min_value=0.0, value=1.0, step=0.5)
kwp_nord = st.sidebar.number_input("Installierte kWp – Nord", min_value=0.0, value=1.0, step=0.5)
kwp_west = st.sidebar.number_input("Installierte kWp – West", min_value=0.0, value=1.0, step=0.5)
kwp_map  = {"Süd": kwp_sued, "Ost": kwp_ost, "Nord": kwp_nord, "West": kwp_west}
start_button = st.sidebar.button("🔄 Optimierung berechnen")

# ══════════════════════════════════════════════════════════════════════════════
# HAUPTBEREICH
# ══════════════════════════════════════════════════════════════════════════════
if daten is None:
    st.error("❌ Excel-Datei nicht gefunden.")
    st.stop()

st.success(f"✅ Datei: `{gefundener_pfad}`")
st.subheader(f"📅 Analyse: {ansicht_titel}")

if ist_schaetzung:
    profil_tag = klemme_datum(heute if ansicht_modus == "📆 Heute (24h ab jetzt)" else gewaehltes_datum)
    st.info(f"📊 Daten-Schätzung auf Basis des Profiltags {profil_tag.strftime('%d.%m.%Y')} aus der Excel-Tabelle.")

# Daten extrahieren für Tagesansichten
if ansicht_modus != "📊 Monate wählen":
    hs_z      = slots_fuer_zeitraum(daten["hs"],     start_slot, n_slots)
    waerme_z  = slots_fuer_zeitraum(daten["waerme"], start_slot, n_slots)
    preise_z  = slots_fuer_zeitraum(daten["preise"], start_slot, n_slots)

pv_richtungen  = {}
for (neigung, richtung), info in daten["pv"].items():
    if neigung != neigung_wahl:
        continue
    kwp = kwp_map.get(richtung, 0.0)
    if ansicht_modus == "📊 Monate wählen":
        monat_start = dt.date(2025, gewaehlter_monat_nr, 1)
        monat_ende  = dt.date(2025, gewaehlter_monat_nr, 31) if gewaehlter_monat_nr == 12 else dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)
        tage = (monat_ende - monat_start).days + 1
        m_start_slot = slot_von_datum(monat_start, 0)
        pv_richtungen[richtung] = slots_fuer_zeitraum(info["kurve"], m_start_slot, tage * 96) * kwp
    else:
        pv_richtungen[richtung] = slots_fuer_zeitraum(info["kurve"], start_slot, n_slots) * kwp

pv_gesamt_z = sum(pv_richtungen.values()) if pv_richtungen else np.zeros(n_slots)

# Kennzahlen-Berechnung
if ansicht_modus != "📊 Monate wählen":
    tag_hs     = float(hs_z[fokus_start : fokus_start + 96].sum())
    tag_waerm  = float(waerme_z[fokus_start : fokus_start + 96].sum())
    tag_pv     = float(pv_gesamt_z[fokus_start : fokus_start + 96].sum())
else:
    tag_hs     = float(hs_z.sum())
    tag_waerm  = float(waerme_z.sum())
    tag_pv     = float(pv_gesamt_z.sum())

ueberschuss = tag_pv - (tag_hs + tag_waerm)

k1, k2, k3, k4, k5 = st.columns(5)
with k1: st.metric("📊 Haushaltsstrom",    f"{tag_hs:.1f} kWh")
with k2: st.metric("🔥 Wärmebedarf",       f"{tag_waerm:.1f} kWh")
with k3: st.metric("⚡ Gesamt-Verbrauch",   f"{tag_hs+tag_waerm:.1f} kWh")
with k4: st.metric("☀️ PV-Erzeugung",      f"{tag_pv:.1f} kWh")
with k5:
    if ueberschuss >= 0: st.metric("🟢 PV-Überschuss",  f"+{ueberschuss:.1f} kWh")
    else: st.metric("🔴 Netzstrombedarf", f"{abs(ueberschuss):.1f} kWh")

# ── TAGESÜBERSICHT
if ansicht_modus == "📊 Monate wählen" and anzahl_tage > 1:
    st.markdown("---")
    st.subheader("📊 Tagesübersicht")
    tages_daten = []
    monat_start = dt.date(2025, gewaehlter_monat_nr, 1)
    monat_ende  = dt.date(2025, gewaehlter_monat_nr, 31) if gewaehlter_monat_nr == 12 else dt.date(2025, gewaehlter_monat_nr + 1, 1) - dt.timedelta(days=1)
    tage = (monat_ende - monat_start).days + 1
    for i in range(tage):
        tag_datum    = monat_start + dt.timedelta(days=i)
        tag_slot     = slot_von_datum(tag_datum, 0)
        t_hs         = float(slots_fuer_zeitraum(daten["hs"],     tag_slot, 96).sum())
        t_waerme     = float(slots_fuer_zeitraum(daten["waerme"], tag_slot, 96).sum())
        t_pv         = float(sum(slots_fuer_zeitraum(info["kurve"], tag_slot, 96).sum() * kwp_map.get(r, 0.0) for (n, r), info in daten["pv"].items() if n == neigung_wahl))
        tages_daten.append({
            "Datum": tag_datum.strftime("%a %d.%m"), "Haushaltsstrom": round(t_hs, 2),
            "Wärmebedarf": round(t_waerme, 2), "Gesamt-Verbrauch": round(t_hs + t_waerme, 2),
            "PV-Erzeugung": round(t_pv, 2), "Bilanz (PV-Verb.)": round(t_pv - (t_hs + t_waerme), 2),
        })
    df_tage = pd.DataFrame(tages_daten).set_index("Datum")
    st.dataframe(df_tage, use_container_width=True)

# ══════════════════════════════════════════════════════════════════════════════
# DIAGRAMME (Im Monatsmodus jetzt vollständig sichtbar!)
# ══════════════════════════════════════════════════════════════════════════════
st.markdown("---")
st.subheader("🔮 Verbrauchsprofil")
st.caption("👉 Nutze den Regler unten oder ziehe das Diagramm mit der Maus: Nach LINKS für die Vergangenheit, nach RECHTS für die Zukunft!")

c1, c2, c3 = st.columns(3)
with c1:
    st.markdown("📊 **Haushaltsstrom (kWh/15min)**")
    plotly_line({"Haushaltsstrom": hs_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#FF4B4B"])
with c2:
    st.markdown("🔥 **Wärmebedarf (kWh/15min)**")
    plotly_line({"Wärmebedarf": waerme_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#FFA500"])
with c3:
    st.markdown("⚡ **Gesamt-Verbrauch (kWh/15min)**")
    plotly_line({"Gesamt-Verbrauch": hs_z + waerme_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#0055FF"])

# ── PV-ERZEUGUNG
st.markdown("---")
st.subheader(f"☀️ PV-Erzeugung – {neigung_wahl}")
aktive_pv = {r: v for r, v in pv_richtungen.items() if v.sum() > 0}
if aktive_pv:
    pv_chart_dict = dict(aktive_pv)
    pv_chart_dict["Gesamt"] = pv_gesamt_z
    plotly_line(pv_chart_dict, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende)
else:
    st.info("ℹ️ Bitte trag kWp-Werte > 0 in der Seitenleiste ein.")

# ── BÖRSENSTROMPREISE
st.markdown("---")
st.subheader("💶 Börsenstrompreise")
plotly_line({"ct/kWh": preise_z}, labels, "", fokus_start=fokus_start, fokus_ende=fokus_ende, ren_farben=["#AA44FF"])

# ── OPTIMIERUNG
if start_button:
    st.markdown("---")
    st.subheader("🏆 Optimierungsergebnis")
    if ansicht_modus == "📊 Monate wählen":
        st.warning("⚠️ Optimierung ist nur für Tagesansichten sinnvoll.")
    else:
        kapazitaet_kwh = 50.0
        energie_needed = (soc_ziel - soc_aktuell) / 100.0 * kapazitaet_kwh
        st.write(f"**Benötigte Ladeenergie:** {energie_needed:.1f} kWh")
        if "Öko" in ladestrategie:
            ladeleistung_kw = 11.0
            slots_noetig    = max(1, int(np.ceil(energie_needed / (ladeleistung_kw * 0.25))))
            
            preise_heute    = preise_z[fokus_start : fokus_start + 96]
            guenstigste     = np.argsort(preise_heute)[:slots_noetig]
            
            ladeplan        = np.zeros(n_slots)
            ladeplan[fokus_start + guenstigste] = ladeleistung_kw * 0.25
            plotly_bar(ladeplan, labels, "Ladeenergie (kWh)", "#00AA44", fokus_start=fokus_start, fokus_ende=fokus_ende)