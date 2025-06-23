# -*- coding: utf-8 -*-
"""app.py – Générateur multiconnexion (PF1)
Exporte **toujours** un fichier Excel (.xlsx).\n
Exigence : au moins l’un des moteurs Excel suivants doit être installé dans l’environnement :\n• `openpyxl` (recommandé)\n• `xlsxwriter`\n
Si aucun moteur n’est disponible, l’app affiche un message d’erreur clair invitant à installer `openpyxl`.
"""

from __future__ import annotations

import importlib.util
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ───────────────────────── CONFIG ─────────────────────────
st.set_page_config(page_title="multiconnexion", page_icon="📦")

st.title("📦 Générateur PF1 – multiconnexion (export Excel)")
st.markdown(
    "Chargez un fichier CSV ou Excel contenant **Numéro de compte**, **Raison sociale** et **Adresse**.\n"
    "Le résultat sera exporté en **Excel .xlsx**. Assurez‑vous que `openpyxl` ou `xlsxwriter` est installé sur l’hôte."
)

# ───────────────────────── Détection moteur Excel ─────────────────────────
if importlib.util.find_spec("openpyxl") is not None:
    EXCEL_ENGINE = "openpyxl"
elif importlib.util.find_spec("xlsxwriter") is not None:
    EXCEL_ENGINE = "xlsxwriter"
else:
    EXCEL_ENGINE = None  # aucun moteur ; on bloquera l’export

# ───────────────────────── UPLOAD ─────────────────────────
uploaded = st.file_uploader("📄 Fichier comptes", type=("csv", "xlsx", "xls"))

# ───────────────────────── PARAMÈTRES ─────────────────────────
col1, col2 = st.columns(2)
with col1:
    entreprise = st.text_input("🏢 Entreprise", placeholder="Ex. DALKIA")
with col2:
    vm_choice = st.radio("🗂️ ViewMasterCatalog", options=["True", "False"], index=0, horizontal=True)

# ───────────────────────── UTILITAIRES ─────────────────────────

@st.cache_data(show_spinner=False)
def read_any(file):
    """Lit CSV (encodages usuels) ou Excel (.xlsx) via openpyxl."""
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu ; essayez UTF-8, Latin-1 ou CP1252.")

    # Excel : nécessite openpyxl pour la lecture
    if importlib.util.find_spec("openpyxl") is None:
        raise ImportError("Le module openpyxl est requis pour lire les fichiers Excel .xlsx.")

    return pd.read_excel(file, engine="openpyxl")

def build_pf1(df: pd.DataFrame, ent: str, vm_flag: str) -> pd.DataFrame:
    required = {"Numéro de compte", "Raison sociale", "Adresse"}
    missing = required - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(sorted(missing))}")

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
            rs,
            f"frx-variant-{ent}-configuration-set",
            f"PC_{ent}",
            vm_flag,
        ]
    return pf1

def export_excel_bytes(df: pd.DataFrame):
    """Renvoie un tuple (bytes, mime) pour l’export Excel obligatoire."""
    if EXCEL_ENGINE is None:
        raise ImportError(
            "Aucun moteur Excel disponible. Installez le paquet openpyxl :\n    pip install openpyxl"
        )

    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine=EXCEL_ENGINE) as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ───────────────────────── ACTION ─────────────────────────

if st.button("🚀 Générer le PF1 (.xlsx)"):
    if uploaded is None or not entreprise:
        st.warning("Veuillez déposer un fichier et saisir le nom d’entreprise.")
    else:
        try:
            df_src = read_any(uploaded)
            pf1 = build_pf1(df_src, entreprise.strip(), vm_choice)

            st.success("✅ Fichier généré ! Aperçu :")
            st.dataframe(pf1.head())

            data_bytes, mime = export_excel_bytes(pf1)
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"PF1_{entreprise.replace(' ', '_')}_{ts}.xlsx"

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=data_bytes,
                file_name=filename,
                mime=mime,
            )
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
