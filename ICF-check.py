import streamlit as st
import pandas as pd
from datetime import datetime
from docx import Document
from io import BytesIO

st.set_page_config(page_title="ICF Consent Report Generator", layout="centered")

st.title("📄 ICF Consent Report Generator")
st.write("Lade die drei benötigten Excel-Dateien per Drag & Drop hoch.")

# --------------------------
# Upload-Felder
# --------------------------

icf_file = st.file_uploader("ICF-Datei (mit Sheet 'ICF2')", type=["xlsx", "xls"])
consent_file = st.file_uploader("Consent-Daten", type=["xlsx", "xls"])
eos_file = st.file_uploader("EOS-Daten", type=["xlsx", "xls"])

# --------------------------
# Hilfsfunktionen
# --------------------------

def load_icf_versions(file):
    xls = pd.ExcelFile(file)
    sheet = "ICF2" if "ICF2" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)

    col_map = {}
    for c in df.columns:
        lc = c.lower()
        if "icf" in lc and "version" in lc:
            col_map[c] = "ICF Version"
        if "gültig" in lc or "valid" in lc:
            col_map[c] = "Gültig ab"

    df = df.rename(columns=col_map)
    df["Gültig ab"] = pd.to_datetime(df["Gültig ab"], errors="coerce")
    df = df.sort_values("Gültig ab").reset_index(drop=True)

    return df


def load_consents(file):
    df = pd.read_excel(file, dtype={"mnpaid": str})
    df["mnpaid"] = df["mnpaid"].astype(str)
    df["icdat"] = pd.to_datetime(df["icdat"], errors="coerce")

    if "mnp_rando_gr" not in df.columns:
        df["mnp_rando_gr"] = ""
    if "mnp_rando_v6_gr" not in df.columns:
        df["mnp_rando_v6_gr"] = ""

    return df


def load_eos(file):
    df = pd.read_excel(file, dtype={"mnpaid": str})
    df["eosdat"] = pd.to_datetime(df["eosdat"], errors="coerce")
    df["mnpaid"] = df["mnpaid"].astype(str)
    return df


def find_icf_version(icf_df, date):
    if pd.isna(date):
        return None

    rows = icf_df.reset_index(drop=True)
    for i, row in rows.iterrows():
        valid_from = row["Gültig ab"]
        next_valid = rows.iloc[i+1]["Gültig ab"] if i+1 < len(rows) else None

        if next_valid is None:
            if date >= valid_from:
                return row["ICF Version"]
        else:
            if valid_from <= date < next_valid:
                return row["ICF Version"]

    return None


# --------------------------
# Report Generator
# --------------------------

def generate_report(icf_df, consents_df, eos_df):
    eos_map = eos_df.set_index("mnpaid")["eosdat"].to_dict()
    rows = []

    for pid, group in consents_df.groupby("mnpaid"):
        group = group.sort_values("icdat")
        eos_date = eos_map.get(pid, pd.NaT)

        # Basis-Kommentartext vorbereiten
        first_row = group.iloc[0]
        r1 = first_row.get("mnp_rando_gr", "") or "-"
        r2 = first_row.get("mnp_rando_v6_gr", "") or "-"
        rando_text = f"{r1} / {r2}"

        # EOS-Text hinzufügen (falls vorhanden)
        eos_text = ""
        if not pd.isna(eos_date):
            eos_text = f"EOS ({eos_date.strftime('%d.%m.%Y')})"

        # Beide Kommentare kombinieren
        comment_block = "\n".join([x for x in [rando_text, eos_text] if x])

        # Jede Consent-Zeile einfügen
        for _, rec in group.iterrows():
            icdate = rec["icdat"]
            version = find_icf_version(icf_df, icdate)

            rows.append({
                "Patient-ID": pid,
                "Version": version or "",
                "Date": icdate.strftime("%Y-%m-%d") if not pd.isna(icdate) else "",
                "Comment": comment_block
            })

        # Re-consent Logik
        last_consent = group["icdat"].max()

        for _, icf_row in icf_df.iterrows():
            v_name = icf_row["ICF Version"]
            v_valid = icf_row["Gültig ab"]

            if v_valid > last_consent and (pd.isna(eos_date) or eos_date >= v_valid):
                already_signed = any(group["icdat"] >= v_valid)

                if not already_signed:
                    rows.append({
                        "Patient-ID": pid,
                        "Version": v_name,
                        "Date": "CHECK",
                        "Comment": comment_block
                    })

    # Word-Dokument erstellen
    doc = Document()
    doc.add_heading("Consent Report", level=1)

    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Patient-ID"
    hdr[1].text = "Version of Informed Consent Form"
    hdr[2].text = "Date of Consent"
    hdr[3].text = "Comment"

    for r in rows:
        row_cells = table.add_row().cells
        row_cells[0].text = r["Patient-ID"]
        row_cells[1].text = r["Version"]
        row_cells[2].text = r["Date"]
        row_cells[3].text = r["Comment"]

    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# --------------------------
# Button → Report generieren
# --------------------------

if icf_file and consent_file and eos_file:
    if st.button("📄 Report generieren"):
        with st.spinner("Erstelle Word-Datei..."):
            icf_df = load_icf_versions(icf_file)
            consents_df = load_consents(consent_file)
            eos_df = load_eos(eos_file)

            word_file = generate_report(icf_df, consents_df, eos_df)

        st.success("Fertig!")
        st.download_button(
            label="⬇️ Word-Datei herunterladen",
            data=word_file,
            file_name="consent_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("Bitte zuerst alle drei Dateien hochladen.")
