import streamlit as st
import pandas as pd
from datetime import datetime

# --- Streamlit Seitenkonfiguration ---
st.set_page_config(page_title="Touren-Export als JSON", layout="centered")
st.title("Dienstplan als JSON exportieren")

# --- File Uploader ---
uploaded_files = st.file_uploader(
    "Excel-Dateien hochladen (Blatt 'Touren')",
    type=["xlsx"],
    accept_multiple_files=True
)

if uploaded_files:
    json_records = []
    # Schlüsselwörter, die zum Ausschluss eines Eintrags führen
    exclusion_keywords = [
        "zippel", "insel", "paasch", "meyer",
        "ihde", "devies", "insellogistik"
    ]

    # --- Verarbeitung jeder hochgeladenen Datei ---
    for file in uploaded_files:
        # Excel-Datei einlesen (startet ab der 5. Zeile)
        df = pd.read_excel(
            file,
            sheet_name="Touren",
            skiprows=4,
            engine="openpyxl"
        )
        
        # Jede Zeile der Tabelle durchgehen
        for _, row in df.iterrows():
            datum = row.iloc[14]
            tour = row.iloc[15]
            uhrzeit = row.iloc[8]

            # Zeilen ohne gültiges Datum überspringen
            if pd.isna(datum):
                continue
            try:
                datum_dt = pd.to_datetime(datum)
            except ValueError:
                continue

            # Uhrzeit formatieren
            if pd.isna(uhrzeit):
                uhrzeit_str = "–"
            elif isinstance(uhrzeit, datetime):
                uhrzeit_str = uhrzeit.strftime("%H:%M")
            else:
                uhrzeit_str = str(uhrzeit).strip()


            # --- Fahrer-Kombinationen prüfen ---
            # Prüft Spalten (D, E) und (G, H) auf Fahrernamen
            fahrer_kombis = [(3, 4), (6, 7)] 
            fahrer_gefunden = False

            for pos in fahrer_kombis:
                nachname_raw = row.iloc[pos[0]]
                vorname_raw = row.iloc[pos[1]]

                # Wenn beide Felder leer sind, nächste Kombination prüfen
                if pd.isna(nachname_raw) and pd.isna(vorname_raw):
                    continue

                nachname = str(nachname_raw).strip().title() if pd.notna(nachname_raw) else ""
                vorname = str(vorname_raw).strip().title() if pd.notna(vorname_raw) else ""
                fahrer_name = f"{nachname}, {vorname}".strip(", ")

                # Prüfen, ob der Fahrername ein Ausschlusskriterium enthält
                fahrer_name_lower = fahrer_name.lower()
                if any(keyword in fahrer_name_lower for keyword in exclusion_keywords):
                    continue

                # JSON-Eintrag erstellen und zur Liste hinzufügen
                record = {
                    "Fahrer": fahrer_name,
                    "Datum": datum_dt.date().isoformat(),
                    "Uhrzeit": uhrzeit_str,
                    "Tour/Aufgabe": str(tour).strip() if pd.notna(tour) else ""
                }
                json_records.append(record)
                fahrer_gefunden = True

    # --- JSON-Export anbieten ---
    if json_records:
        # DataFrame aus den gesammelten Einträgen erstellen
        df_export = pd.DataFrame(json_records)
        
        # In JSON umwandeln (UTF-8 für Umlaute, eingerückt für Lesbarkeit)
        json_str = df_export.to_json(orient='records', force_ascii=False, indent=2)
        json_bytes = json_str.encode('utf-8')

        st.download_button(
            label="JSON mit allen Touren herunterladen",
            data=json_bytes,
            file_name="touren_export.json",
            mime="application/json"
        )
    else:
        st.warning("Keine gültigen Touren-Einträge in den hochgeladenen Dateien gefunden.")
