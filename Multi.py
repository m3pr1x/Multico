# -*- coding: utf-8 -*-
"""app.py – Générateur multiconnexion (PF1)
Fonctionne même sans openpyxl/xlsxwriter : si aucun moteur Excel n’est dispo, l’app exporte un CSV.
Dépendances minimales : pandas, streamlit (openpyxl ou xlsxwriter *optionnels*).
"""

from __future__ import annotations

import importlib.util
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ───────────────────────── CONFIG ─────────────────────────
st.set_page_config(page_title="multiconnexion", page_icon="📦")

st.title("📦 Générateur PF1 – multiconnexion")
st.markdown(
    "Chargez un fichier CSV ou Excel contenant **Numéro de compte**, **Raison sociale** et **Adresse**."\
    " Le résultat peut être exporté en Excel (si openpyxl/xlsxwriter est installé) ou, à défaut, en CSV."
)

# ───────────────────────── Détection moteurs Excel ─────────────────────────
if importlib.util.find_spec("openpyxl") is not None:
    EXCEL_ENGINE = "openpyxl"
elif importlib.util.find_spec("xlsxwriter") is not None:
    EXCEL_ENGINE = "xlsxwriter"
else:
    EXCEL_ENGINE = None  # aucun moteur Excel dispo

# ───────────────────────── UPLOAD ─────────────────────────
uploaded = st.file_uploader("📄 Fichier comptes", type=("csv", "xlsx", "xls"))

# ───────────────────────── PARAMS ─────────────────────────
col1, col2 = st.columns(2)
with col1:
    entreprise = st.text_input("🏢 Entreprise", placeholder="Ex. DALKIA")
with col2:
    vm_choice = st.radio("🗂️ ViewMasterCatalog", options=["True", "False"], horizontal=True)

# ───────────────────────── HELPERS ─────────────────────────

@st.cache_data(show_spinner=False)
def read_any(file):
    """Lit CSV ou Excel."""
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu.")

    # Excel → nécessite openpyxl
    if EXCEL_ENGINE != "openpyxl":
        raise ImportError("Le module openpyxl est requis pour lire les fichiers Excel .xlsx.")

    return pd.read_excel(file, engine="openpyxl")

def build_pf1(df: pd.DataFrame, ent: str, vm_flag: str) -> pd.DataFrame:
    req = {"Numéro de compte", "Raison sociale", "Adresse"}
    missing = req - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(sorted(missing))}")

    pf1 = pd.DataFrame(columns=[
        "uid", "name", "locName",
        "CXmIAssignedConfiguration", "pcCompoundProfile", "ViewMasterCatalog",
    ])

    for code in df["Numéro de compte"].dropna().unique():
        row = df.loc[df["Numéro de compte"] == code].iloc[0]
        rs = row["Raison sociale"]
        pf1.loc[len(pf1)] = [
            code,
            rs,
            rs,  # locName = name
            f"frx-variant-{ent}-configuration-set",
            f"PC_{ent}",
            vm_flag,
        ]
    return pf1

def export_bytes(df: pd.DataFrame):
    """Retourne (bytes, extension, mime) en fonction des moteurs dispo."""
    if EXCEL_ENGINE:
        buffer = BytesIO()
        with pd.ExcelWriter(buffer, engine=EXCEL_ENGINE) as writer:
            df.to_excel(writer, index=False)
        buffer.seek(0)
        return buffer.getvalue(), "xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    # Fallback CSV utf‑8
    return df.to_csv(index=False).encode("utf-8"), "csv", "text/csv"

# ───────────────────────── ACTION ─────────────────────────

if st.button("🚀 Générer le PF1"):
    if uploaded is None or not entreprise:
        st.warning("Veuillez déposer un fichier et saisir le nom d’entreprise.")
    else:
        try:
            df_src = read_any(uploaded)
            pf1 = build_pf1(df_src, entreprise.strip(), vm_choice)

            st.success("✅ Fichier généré ! Aperçu :")
            st.dataframe(pf1.head())

            data_bytes, ext, mime = export_bytes(pf1)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"PF1_{entreprise.replace(' ', '_')}_{ts}.{ext}"

            st.download_button(
                label=f"📥 Télécharger le fichier {ext.upper()}",
                data=data_bytes,
                file_name=filename,
                mime=mime,
            )
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
