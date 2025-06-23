# -*- coding: utf-8 -*-
"""app.py – Générateur multiconnexion (PF1)
Version simplifiée : impose explicitement `openpyxl` comme moteur Excel et l’importe en début de script.
Ajoutez simplement `openpyxl` dans votre *requirements.txt* : 
```
pandas
streamlit
openpyxl
```
```bash
pip install -r requirements.txt
```
"""

from __future__ import annotations

from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ───────────────────────── DEPENDENCY CHECK ─────────────────────────
try:
    import openpyxl  # noqa: F401 – assure l’import
except ImportError as e:
    raise ImportError(
        "Le module 'openpyxl' est requis pour lire/écrire les fichiers Excel (.xlsx).\n"
        "Installez-le avec : pip install openpyxl"
    ) from e

EXCEL_ENGINE = "openpyxl"  # moteur unique désormais

# ───────────────────────── CONFIG ─────────────────────────
st.set_page_config(page_title="multiconnexion", page_icon="📦")

st.title("📦 Générateur PF1 – multiconnexion (export Excel)")
st.markdown(
    "Chargez un fichier CSV ou Excel contenant **Numéro de compte**, **Raison sociale** et **Adresse**.\n\n"
    "Le résultat sera exporté en **Excel (.xlsx)** à l’aide du moteur openpyxl."
)

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
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu.")

    return pd.read_excel(file, engine=EXCEL_ENGINE)

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

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine=EXCEL_ENGINE) as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()

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

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"PF1_{entreprise.replace(' ', '_')}_{ts}.xlsx"

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=to_excel_bytes(pf1),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as e:
            st.error(f"❌ Erreur : {e}")
