# -*- coding: utf-8 -*-
"""app.py – Générateur B2B Units (PF1)
Streamlit app à placer dans un dépôt GitHub puis à déployer sur Streamlit Cloud.
Dépendances : pandas, streamlit, openpyxl (pas besoin de xlsxwriter)
"""

import pandas as pd
from datetime import datetime
from io import BytesIO
import streamlit as st

# ───────────────────── CONFIG ─────────────────────
st.set_page_config(page_title="multiconnexion", page_icon="📦")

st.title("📦 Outil de multiconnextion)")
st.markdown(
    "Déposez un fichier CSV ou Excel contenant les colonnes **Numéro de compte**, **Raison sociale** et **Adresse**. "
    "Renseignez les champs ci‑dessous puis cliquez sur **Générer** pour télécharger un fichier Excel prêt à importer."
)

# ───────────────────── UPLOAD ─────────────────────
uploaded = st.file_uploader("📄 Fichier comptes", type=("csv", "xlsx", "xls"))

# ───────────────────── PARAMÈTRES ─────────────────────
col1, col2 = st.columns(2)
with col1:
    entreprise = st.text_input("🏢 Entreprise", key="ent")
with col2:
    fot = st.text_input("🆔 True or False", key="T or F")

# ───────────────────── UTILS ─────────────────────

@st.cache_data(show_spinner=False)
def read_any(file):
    """Lit un CSV ou Excel (premier onglet) en gérant encodages."""
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu.")
    # Excel
    return pd.read_excel(file, engine="openpyxl")

def build_pf1(df: pd.DataFrame, ent: str, fot: str) -> pd.DataFrame:
    """Construit le DataFrame PF1 à partir des colonnes obligatoires."""
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
        pf1.loc[len(pf1)] = [
            code,
            row["Raison sociale"],
            row["Adresse"],
            f"frx-variant-{ent}-configuration-set",
            f"PC_{ent}",
            fot,
        ]

    return pf1

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Renvoie un fichier Excel en mémoire (engine = openpyxl)."""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False)
    buffer.seek(0)
    return buffer.getvalue()

# ───────────────────── ACTION ─────────────────────

if st.button("🚀 Générer"):
    if uploaded is None or not entreprise or not fot:
        st.warning("Veuillez déposer un fichier et renseigner tous les champs.")
    else:
        try:
            src = read_any(uploaded)
            pf1_df = build_pf1(src, entreprise.strip(), fot.strip())

            st.success("✅ Fichier généré !")
            st.dataframe(pf1_df.head())

            filename = (
                f"B2B Units creation_{entreprise.replace(' ', '_')}_"
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
