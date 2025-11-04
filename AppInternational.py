# -*- coding: utf-8 -*-
"""
Created on Mon Nov  3 12:54:22 2025

@author: nina.reitsam
"""

# AppInternational.py
import os
from datetime import datetime
from typing import List, Dict
import pandas as pd
import streamlit as st

# =======================
# Konfiguration
# =======================
EXCEL_PATH = r"J:\02_International\08_Statistiken\DashboardApp\DashboardApp_2025.xlsx"

# Übersichtszähler NUR Land + Anzahl (ohne Thema)
SHEET_COUNTS = "Zaehlungen"   # Spalten: Land | Anzahl

# Vorgabewerte
COUNTRIES: List[str] = [
    "Österreich", "Schweiz", "Italien", "Frankreich", "USA",
    "GB", "China", "Polen", "Ungarn", "Tschechien", "Slowakei",
]
TOPICS: List[str] = [
    "Mitarbeiterentsendung",
    "Marktberatung",
    "XXX",
    "XXY",
]
EMPLOYEES: List[str] = ["Behrenz", "Glas", "Li", "Lovell", "Wind"]  # alphabetisch
OTHER_LABEL = "Sonstiges"

# =======================
# Excel-Helfer
# =======================
def ensure_workbook_exists():
    folder = os.path.dirname(EXCEL_PATH)
    if folder and not os.path.exists(folder):
        os.makedirs(folder, exist_ok=True)
    if not os.path.exists(EXCEL_PATH):
        df = pd.DataFrame(columns=["Land", "Anzahl"])
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as w:
            df.to_excel(w, sheet_name=SHEET_COUNTS, index=False)

def read_counts() -> pd.DataFrame:
    """Liest das Übersichtssheet 'Zaehlungen' (Land, Anzahl)."""
    ensure_workbook_exists()
    try:
        df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_COUNTS, engine="openpyxl")
    except Exception:
        df = pd.DataFrame(columns=["Land", "Anzahl"])
        with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="overlay") as w:
            df.to_excel(w, sheet_name=SHEET_COUNTS, index=False)
    if "Land" not in df.columns:
        df["Land"] = []
    if "Anzahl" not in df.columns:
        df["Anzahl"] = 0
    df["Land"] = df["Land"].astype(str).str.strip()
    df["Anzahl"] = pd.to_numeric(df["Anzahl"], errors="coerce").fillna(0).astype(int)
    return df

def write_counts(df: pd.DataFrame):
    """Schreibt ausschließlich das Sheet 'Zaehlungen'."""
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df.to_excel(w, sheet_name=SHEET_COUNTS, index=False)

def bump_country_total(land: str):
    """Erhöht die Anzahl in 'Zaehlungen' nur für das Land um +1."""
    df = read_counts()
    mask = (df["Land"] == land)
    if not mask.any():
        df = pd.concat([df, pd.DataFrame([{"Land": land, "Anzahl": 0}])], ignore_index=True)
        mask = (df["Land"] == land)
    df.loc[mask, "Anzahl"] = pd.to_numeric(df.loc[mask, "Anzahl"], errors="coerce").fillna(0).astype(int) + 1
    write_counts(df)

def append_detail_row(land: str, thema: str, payload: Dict[str, str]):
    """
    Schreibt eine Detailzeile in ein Sheet mit Namen = Land.
    Spalten: Zeitstempel | Mitarbeiter | Thema | Unternehmensname | Ansprechpartner | Identnummer | Bemerkung
    """
    ensure_workbook_exists()
    try:
        df_log = pd.read_excel(EXCEL_PATH, sheet_name=land, engine="openpyxl")
    except Exception:
        df_log = pd.DataFrame(columns=[
            "Zeitstempel", "Mitarbeiter", "Thema",
            "Unternehmensname", "Ansprechpartner", "Identnummer", "Bemerkung"
        ])
    new_row = pd.DataFrame([{
        "Zeitstempel": datetime.now(),
        "Mitarbeiter": (payload.get("mitarbeiter") or "").strip(),
        "Thema": (thema or "").strip(),
        "Unternehmensname": (payload.get("firma") or "").strip(),
        "Ansprechpartner": (payload.get("ansprechpartner") or "").strip(),
        "Identnummer": (payload.get("identnummer") or "").strip(),
        "Bemerkung": (payload.get("bemerkung") or "").strip(),
    }])
    df_log = pd.concat([df_log, new_row], ignore_index=True)
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as w:
        df_log.to_excel(w, sheet_name=land, index=False)

def list_detail_sheets() -> List[str]:
    """Alle Sheetnamen außer 'Zaehlungen' (Detailblätter)."""
    ensure_workbook_exists()
    try:
        xls = pd.ExcelFile(EXCEL_PATH, engine="openpyxl")
        return [s for s in xls.sheet_names if s != SHEET_COUNTS]
    except Exception:
        return []

def matrix_from_details() -> pd.DataFrame:
    """
    Baut die Matrix (Thema × Land) aus den Detailblättern.
    Freitext-Werte werden jeweils als 'Sonstiges' gebündelt.
    """
    sheets = list_detail_sheets()
    records = []
    for sheet in sheets:
        try:
            df = pd.read_excel(EXCEL_PATH, sheet_name=sheet, engine="openpyxl")
        except Exception:
            continue
        if df.empty or "Thema" not in df.columns:
            continue
        # Mapping auf definierte Listen, sonst -> Sonstiges
        mapped_land = sheet if sheet in COUNTRIES else OTHER_LABEL
        for _, row in df.iterrows():
            topic_val = str(row.get("Thema", "")).strip()
            mapped_topic = topic_val if topic_val in TOPICS else OTHER_LABEL
            records.append({"Land": mapped_land, "Thema": mapped_topic, "Anzahl": 1})

    # Wenn noch keine Details existieren: leere Matrix
    if not records:
        base = pd.MultiIndex.from_product([TOPICS + [OTHER_LABEL], COUNTRIES + [OTHER_LABEL]],
                                          names=["Thema", "Land"])
        pivot = pd.Series(0, index=base).unstack("Land")
        return pivot.reset_index()

    df_all = pd.DataFrame(records)
    # Vollständige Matrix inkl. 'Sonstiges' erzwingen
    all_idx = pd.MultiIndex.from_product([TOPICS + [OTHER_LABEL], COUNTRIES + [OTHER_LABEL]],
                                         names=["Thema", "Land"])
    grp = df_all.groupby(["Thema", "Land"], as_index=True)["Anzahl"].sum().reindex(all_idx, fill_value=0)
    pivot = grp.unstack("Land")
    # Spalten in gewünschter Reihenfolge
    pivot = pivot[COUNTRIES + [OTHER_LABEL]]
    # Index (Themen) ebenso sortieren
    pivot = pivot.reindex(TOPICS + [OTHER_LABEL])
    return pivot.reset_index()

# =======================
# Streamlit UI
# =======================
st.set_page_config(page_title="International – Beratungs-Dashboard", page_icon="🌍", layout="wide")
st.title("🌍 International – Beratungs-Dashboard")
st.caption(f"Excel: `{EXCEL_PATH}`")

# Session-State
if "view" not in st.session_state:
    st.session_state.view = "overview"   # "overview" | "select" | "form"
if "land" not in st.session_state:
    st.session_state.land = None
if "thema" not in st.session_state:
    st.session_state.thema = None
if "custom_land" not in st.session_state:
    st.session_state.custom_land = ""
if "custom_thema" not in st.session_state:
    st.session_state.custom_thema = ""
if "mitarbeiter" not in st.session_state:
    st.session_state.mitarbeiter = None

def to_overview():
    st.session_state.view = "overview"

def start_new_entry():
    st.session_state.view = "select"
    st.session_state.land = None
    st.session_state.thema = None
    st.session_state.custom_land = ""
    st.session_state.custom_thema = ""
    st.session_state.mitarbeiter = None

def to_form():
    st.session_state.view = "form"

# ---------- OVERVIEW: Matrix (Thema × Land) ----------
if st.session_state.view == "overview":
    st.subheader("Übersicht (Summe Einträge: Thema × Land)")
    try:
        matrix_df = matrix_from_details()
        st.dataframe(matrix_df, use_container_width=True, hide_index=True)
    except Exception as e:
        st.warning(f"Matrix konnte nicht geladen werden: {e}")

    st.divider()
    if st.button("Neuen Eintrag erfassen", type="primary"):
        start_new_entry()
        st.rerun()

# ---------- SELECT: Land & Thema wählen ----------
elif st.session_state.view == "select":
    st.subheader("Land wählen")
    cols_land = st.columns(3)
    for i, c in enumerate(COUNTRIES):
        with cols_land[i % 3]:
            label = f"✅ {c}" if st.session_state.land == c else c
            if st.button(label, use_container_width=True, key=f"land_{c}"):
                st.session_state.land = c
                st.session_state.custom_land = ""
                st.rerun()
    with st.expander("Anderes Land eingeben"):
        st.session_state.custom_land = st.text_input("Anderes Land", st.session_state.custom_land)
        if st.button("Land übernehmen", key="btn_custom_land"):
            if st.session_state.custom_land.strip():
                st.session_state.land = st.session_state.custom_land.strip()
                st.rerun()
    if not st.session_state.land:
        st.info("Bitte Land wählen oder eingeben.")
        st.stop()

    st.divider()
    st.subheader("Thema wählen")
    cols_topic = st.columns(3)
    for i, t in enumerate(TOPICS):
        with cols_topic[i % 3]:
            label = f"✅ {t}" if st.session_state.thema == t else t
            if st.button(label, use_container_width=True, key=f"topic_{t}"):
                st.session_state.thema = t
                st.session_state.custom_thema = ""
                st.rerun()
    with st.expander("Anderes Thema eingeben"):
        st.session_state.custom_thema = st.text_input("Anderes Thema", st.session_state.custom_thema)
        if st.button("Thema übernehmen", key="btn_custom_topic"):
            if st.session_state.custom_thema.strip():
                st.session_state.thema = st.session_state.custom_thema.strip()
                st.rerun()
    if not st.session_state.thema:
        st.info("Bitte Thema wählen oder eingeben.")
        st.stop()

    st.divider()
    if st.button("Weiter zur Detail-Erfassung", type="primary", use_container_width=True):
        try:
            # Zähler NUR nach Land (ohne Thema) pflegen:
            bump_country_total(st.session_state.land)
            st.toast("Zähler (Land) aktualisiert.")
            to_form()
            st.rerun()
        except PermissionError:
            st.error("Excel ist geöffnet/gesperrt. Bitte schließen und erneut versuchen.")
        except Exception as e:
            st.error(f"Fehler beim Zählen: {e}")

# ---------- FORM: Unternehmensdaten + Mitarbeiter ----------
elif st.session_state.view == "form":
    st.subheader("Details eintragen")
    st.caption(f"Land: **{st.session_state.land}** · Thema: **{st.session_state.thema}**")

    st.markdown("**Mitarbeiter wählen**")
    cols_emp = st.columns(min(5, len(EMPLOYEES)))
    for i, name in enumerate(EMPLOYEES):
        with cols_emp[i % len(cols_emp)]:
            label = f"✅ {name}" if st.session_state.mitarbeiter == name else name
            if st.button(label, use_container_width=True, key=f"emp_{name}"):
                st.session_state.mitarbeiter = name
                st.rerun()

    with st.form("detail_form", clear_on_submit=True):
        firma = st.text_input("Unternehmensname (Pflicht)*")
        ansprechpartner = st.text_input("Name Ansprechpartner")
        identnummer = st.text_input("Identnummer")
        bemerkung = st.text_area("Bemerkung", height=100)

        colA, colB = st.columns(2)
        submitted = colA.form_submit_button("Eintrag speichern")
        cancel    = colB.form_submit_button("Abbrechen")

    if cancel:
        st.info("Eingabe verworfen.")
        to_overview()
        st.rerun()

    if submitted:
        if not firma.strip():
            st.warning("Bitte den **Unternehmensnamen** ausfüllen.")
        elif not st.session_state.mitarbeiter:
            st.warning("Bitte **Mitarbeiter** wählen.")
        else:
            try:
                # Detailzeile in Land-Sheet; Thema wird für Matrix später genutzt.
                append_detail_row(
                    st.session_state.land,
                    st.session_state.thema,
                    {
                        "mitarbeiter": st.session_state.mitarbeiter,
                        "firma": firma,
                        "ansprechpartner": ansprechpartner,
                        "identnummer": identnummer,
                        "bemerkung": bemerkung,
                    }
                )
                st.success("Eintrag gespeichert.")
                to_overview()   # zurück zur Matrix-Startansicht
                st.rerun()
            except PermissionError:
                st.error("Excel ist geöffnet/gesperrt. Bitte schließen und erneut versuchen.")
            except Exception as e:
                st.error(f"Fehler beim Speichern: {e}")
