import streamlit as st
import pandas as pd
from docx import Document
from io import BytesIO

# ==========================
# Streamlit Setup
# ==========================

st.set_page_config(page_title="ICF Consent Report Generator", layout="centered")

st.title("📄 ICF Consent Report Generator")
st.write("Studie auswählen und benötigte Excel-Dateien per Drag & Drop hochladen.")

# ==========================
# Mapping laden (lokal)
# ==========================

MAPPING_FILE = "study_mapping.xlsx"

MEANING_TO_INTERNAL = {
    "patientenid": "mnpaid",
    "icf datum unterschrift": "icdat",
    "eos datum": "eosdat",
    "todesdatum": "dthdat",
    "patient eligible": "eligyn",
    "randomisierungsgruppe": "mnp_rando_gr",
    "randomisierungsgruppe2": "mnp_rando_v6_gr"
}

@st.cache_data
def load_study_mapping():
    df = pd.read_excel(MAPPING_FILE)
    df = df[~df.iloc[:, 0].str.lower().eq("xlsx")]
    return df

mapping_df = load_study_mapping()
available_studies = list(mapping_df.columns[1:])

selected_study = st.selectbox("📌 Studie auswählen", available_studies)

def get_mapping_for_study(mapping_df, study):
    mapping = {}
    for _, row in mapping_df.iterrows():
        meaning = str(row.iloc[0]).strip().lower()
        if meaning not in MEANING_TO_INTERNAL:
            continue
        code = row.get(study)
        if pd.isna(code):
            continue
        mapping[MEANING_TO_INTERNAL[meaning]] = str(code)
    return mapping

COLUMN_MAP = get_mapping_for_study(mapping_df, selected_study)

# ==========================
# Uploads
# ==========================

icf_file = st.file_uploader("ICF-Versionen (xlsx, Sheet 'ICF2')", type=["xlsx", "xls"])
consent_file = st.file_uploader("Consent-Daten", type=["xlsx", "xls"])
eos_file = st.file_uploader("EOS-Daten", type=["xlsx", "xls"])
elig_file = st.file_uploader("Eligibility-Daten", type=["xlsx", "xls"])

# ==========================
# Helper Loader
# ==========================

def normalize_columns(df, column_map):
    rename_dict = {v: k for k, v in column_map.items() if v in df.columns}
    return df.rename(columns=rename_dict)

def load_icf_versions(file):
    xls = pd.ExcelFile(file)
    sheet = "ICF2" if "ICF2" in xls.sheet_names else xls.sheet_names[0]
    df = pd.read_excel(xls, sheet_name=sheet)

    df = df.rename(columns={
        c: "ICF Version" for c in df.columns if "version" in c.lower()
    })
    df = df.rename(columns={
        c: "Gültig ab" for c in df.columns if "gültig" in c.lower() or "valid" in c.lower()
    })

    df["Gültig ab"] = pd.to_datetime(df["Gültig ab"], errors="coerce")
    return df.sort_values("Gültig ab").reset_index(drop=True)

def load_consents(file):
    df = pd.read_excel(file, dtype=str)
    df = normalize_columns(df, COLUMN_MAP)
    df["icdat"] = pd.to_datetime(df.get("icdat"), errors="coerce")
    return df

def load_eos(file):
    df = pd.read_excel(file, dtype=str)
    df = normalize_columns(df, COLUMN_MAP)
    df["eosdat"] = pd.to_datetime(df.get("eosdat"), errors="coerce")
    df["dthdat"] = pd.to_datetime(df.get("dthdat"), errors="coerce")
    return df

def load_elig(file):
    df = pd.read_excel(file, dtype=str)
    df = normalize_columns(df, COLUMN_MAP)
    return df

# ==========================
# Core Logic
# ==========================

def find_all_icf_versions(icf_df, date):
    if pd.isna(date):
        return []

    rows = icf_df.sort_values("Gültig ab").reset_index(drop=True)

    valid_rows = rows[rows["Gültig ab"] <= date]

    if valid_rows.empty:
        return []

    # IMPORTANT:
    # only versions with the SAME "gültig ab" as the latest applicable date
    latest_valid_date = valid_rows["Gültig ab"].max()

    return valid_rows[valid_rows["Gültig ab"] == latest_valid_date]["ICF Version"].tolist()


def generate_report(icf_df, consents_df, eos_df, elig_df):
    eos_map = eos_df.set_index("mnpaid").get("eosdat", {}).to_dict()
    dth_map = eos_df.set_index("mnpaid").get("dthdat", {}).to_dict()
    elig_map = elig_df.set_index("mnpaid").get("eligyn", {}).to_dict()

    rows = []

    for pid, group in consents_df.groupby("mnpaid"):
        group = group.sort_values("icdat")
        elig = elig_map.get(pid, "yes").lower()

        r1 = group.iloc[0].get("mnp_rando_gr", "-")
        r2 = group.iloc[0].get("mnp_rando_v6_gr", "-")
        rando_text = f"{r1} / {r2}"

        eos_date = eos_map.get(pid)
        dth_date = dth_map.get(pid)

        eos_text = ""
        if pd.notna(dth_date):
            eos_text = f"EOS (Death, {dth_date.strftime('%d.%m.%Y')})"
        elif pd.notna(eos_date):
            eos_text = f"EOS ({eos_date.strftime('%d.%m.%Y')})"

        signed_versions = {}
        for _, rec in group.iterrows():
            icdate = rec["icdat"]
        
            versions = find_all_icf_versions(icf_df, icdate)
        
            for version in versions:
                signed_versions[version] = icdate.strftime("%Y-%m-%d")

        comment = "Screening Failure" if elig == "no" else "\n".join(filter(None, [rando_text, eos_text]))
        last_consent = group["icdat"].max()

        for _, icf in icf_df.iterrows():
            v = icf["ICF Version"]
            valid_from = icf["Gültig ab"]

            if v in signed_versions:
                date = signed_versions[v]
            elif elig != "no" and valid_from > last_consent and (pd.isna(eos_date) or eos_date >= valid_from):
                date = "CHECK"
            else:
                date = "n.a."

            rows.append({
                "Patient-ID": pid,
                "Version": v,
                "Date": date,
                "Comment": comment
            })

    # Word
    doc = Document()
    table = doc.add_table(rows=1, cols=4)
    hdr = table.rows[0].cells
    hdr[0].text = "Patient-ID"
    hdr[1].text = "Version of Informed Consent Form"
    hdr[2].text = "Date of Consent"
    hdr[3].text = "Comment"


    for r in rows:
        row = table.add_row().cells
        row[0].text = r["Patient-ID"]
        row[1].text = r["Version"]
        row[2].text = r["Date"]
        row[3].text = r["Comment"]

    # --- Merge cells for Patient-ID and Comment ---
    table_rows = table.rows[1:]  # skip header
    n_rows = len(table_rows)
    start = 0
    
    while start < n_rows:
        current_pid = table_rows[start].cells[0].text
        end = start + 1
    
        while end < n_rows and table_rows[end].cells[0].text == current_pid:
            end += 1
    
        if end - start > 1:
    
            # -------- Patient-ID --------
            pid_text = table_rows[start].cells[0].text
    
            # clear text in all involved cells
            for i in range(start, end):
                table_rows[i].cells[0].text = ""
    
            merged_pid = table_rows[start].cells[0].merge(
                table_rows[end - 1].cells[0]
            )
            merged_pid.text = pid_text
    
            # -------- Comment --------
            comment_text = table_rows[start].cells[3].text
    
            for i in range(start, end):
                table_rows[i].cells[3].text = ""
    
            merged_comment = table_rows[start].cells[3].merge(
                table_rows[end - 1].cells[3]
            )
            merged_comment.text = comment_text
    
        start = end



    
    bio = BytesIO()
    doc.save(bio)
    bio.seek(0)
    return bio

# ==========================
# Run
# ==========================

if icf_file and consent_file and eos_file and elig_file:
    if st.button("📄 Report generieren"):
        icf_df = load_icf_versions(icf_file)
        cons_df = load_consents(consent_file)
        eos_df = load_eos(eos_file)
        elig_df = load_elig(elig_file)

        word = generate_report(icf_df, cons_df, eos_df, elig_df)

        st.download_button(
            "⬇️ Word-Datei herunterladen",
            data=word,
            file_name="consent_report.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
else:
    st.info("Bitte Studie auswählen und alle Dateien hochladen.")
