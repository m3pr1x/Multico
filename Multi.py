# -*- coding: utf-8 -*-
"""app.py – Générateur multiconnexion (PF1)
Dépendances : pandas, streamlit, openpyxl **ou** xlsxwriter (au moins l’un des deux).
La colonne **ViewMasterCatalog** est désormais une *chaîne* "True" / "False" au lieu d’un booléen.
"""

import importlib.util
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ───────────────────────── CONFIG ─────────────────────────
st.set_page_config(page_title="multiconnexion", page_icon="📦")

st.title("📦 Outil de multiconnexion – génération PF1")
st.markdown(
    "Déposez un fichier CSV ou Excel contenant les colonnes **Numéro de compte**, **Raison sociale** et **Adresse**.\n\n"
    "Les colonnes générées seront : `uid`, `name`, `locName`, `CXmIAssignedConfiguration`, `pcCompoundProfile`, `ViewMasterCatalog`.\n"
    "• `uid` = Numéro de compte\n"
    "• `name` et `locName` = Raison sociale\n"
    "• `ViewMasterCatalog` = chaîne \"True\" ou \"False\" choisie ci‑dessous"
)

# ───────────────────────── Détection moteur Excel ─────────────────────────
if importlib.util.find_spec("openpyxl") is not None:
    EXCEL_ENGINE = "openpyxl"
elif importlib.util.find_spec("xlsxwriter") is not None:
    EXCEL_ENGINE = "xlsxwriter"
else:
    EXCEL_ENGINE = None  # sera traité plus loin

# ───────────────────────── UPLOAD ─────────────────────────
uploaded = st.file_uploader("📄 Fichier comptes", type=("csv", "xlsx", "xls"))

# ───────────────────────── PARAMÈTRES ─────────────────────────
col1, col2 = st.columns(2)
with col1:
    entreprise = st.text_input("🏢 Entreprise (ex. DALKIA)")
with col2:
    vm_choice = st.radio("🗂️ ViewMasterCatalog ?", options=["True", "False"], horizontal=True)

# ───────────────────────── UTILITAIRES ─────────────────────────

@st.cache_data(show_spinner=False)
def read_any(file):
    """Lit CSV (encodages courants) ou Excel (premier onglet) via openpyxl."""
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu ; essayez UTF‑8, Latin‑1 ou CP1252.")

    # Excel : nécessite openpyxl
    if EXCEL_ENGINE != "openpyxl":
        raise ImportError("Le module openpyxl est requis pour lire les fichiers Excel .xlsx.")

    return pd.read_excel(file, engine="openpyxl")

def build_pf1(df: pd.DataFrame, ent: str, vm_flag: str) -> pd.DataFrame:
    """Construit le DataFrame PF1.\n    - name et locName = Raison sociale\n    - ViewMasterCatalog = vm_flag ("True"/"False")"""
    required = {"Numéro de compte", "Raison sociale", "Adresse"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(sorted(missing))}")

    pf1 = pd.DataFrame(columns=[
        "uid",
        "name",
        "locName",
        "CXmIAssignedConfiguration",
        "pcCompoundProfile",
        "ViewMasterCatalog",
    ])

    for code in df["Numéro de compte"].dropna().unique():
        row = df.loc[df["Numéro de compte"] == code].iloc[0]
        raison_sociale = row["Raison sociale"]
        pf1.loc[len(pf1)] = [
            code,                                # uid
            raison_sociale,                      # name
            raison_sociale,                      # locName (identique)
            f"frx-variant-{ent}-configuration-set",
            f"PC_{ent}",
            vm_flag,                             # "True" / "False" (string)
        ]

    return pf1

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    if EXCEL_ENGINE is None:
        raise ImportError("Installez openpyxl ou xlsxwriter pour exporter le fichier Excel.")

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine=EXCEL_ENGINE) as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()

# ───────────────────────── ACTION ─────────────────────────

if st.button("🚀 Générer le PF1"):
    if uploaded is None or not entreprise:
        st.warning("Veuillez déposer un fichier et renseigner le nom d’entreprise.")
    else:
        try:
            df_src = read_any(uploaded)
            pf1_df = build_pf1(df_src, entreprise.strip(), vm_choice)

            st.success("✅ Fichier généré !")
            st.dataframe(pf1_df.head())

            filename = (
                f"B2B_Units_creation_{entreprise.replace(' ', '_')}_"
                f"{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            )

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=to_excel_bytes(pf1_df),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
