# -*- coding: utf-8 -*-
"""Streamlit app : Générateur B2B Units (PF1)
Hébergez ce script dans un dépôt GitHub (app.py + requirements.txt) puis
connectez‑le à Streamlit Cloud pour le rendre accessible aux équipes.
"""

from __future__ import annotations

import pandas as pd
from datetime import datetime
from io import BytesIO
import streamlit as st

# ───────────────────── PAGE CONFIG ─────────────────────
st.set_page_config(page_title="Générateur B2B Units", page_icon="📦")
st.title("📦 Outil de création des B2B Units (PF1)")
st.markdown("Déposez un fichier Excel ou CSV contenant au minimum les colonnes **Numéro de compte**, **Raison sociale** et **Adresse**.\n\nVous obtiendrez un fichier *B2B Units creation_* prêt à être importé dans CXmIA.")

# ────────────────────── UPLOAD AREA ─────────────────────

def read_any(file) -> pd.DataFrame:
    """Lit un CSV ou un fichier Excel (premier onglet)."""
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu ; essayez UTF‑8, Latin‑1 ou CP1252.")
    # Excel
    return pd.read_excel(file, engine="openpyxl")

uploaded = st.file_uploader("📄 Fichier comptes (CSV ou Excel)", type=("csv", "xlsx", "xls"))

if uploaded is not None:
    try:
        df_accounts = read_any(uploaded)
        st.success(f"{len(df_accounts)} lignes importées.")
        st.dataframe(df_accounts.head())
    except Exception as err:
        st.error(f"❌ Erreur : {err}")
        df_accounts = None
else:
    df_accounts = None

# ────────────────────── PARAMÈTRES ──────────────────────

entreprise = st.text_input("🏢 Entreprise", placeholder="DALKIA / EIFFAGE / ITEC…")
fot        = st.text_input("🆔 FoT", placeholder="FoT123")

# ───────────────────── CORE FUNCTION ─────────────────────

def build_pf1(df: pd.DataFrame, ent: str, fot: str) -> pd.DataFrame:
    """Construit le DF PF1 à partir du DataFrame source."""
    # Vérif colonnes minimales
    required = {"Numéro de compte", "Raison sociale", "Adresse"}
    missing  = required - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(missing)}")

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
            code,                                # uid
            row["Raison sociale"],              # name
            row["Adresse"],                     # locName
            f"frx-variant-{ent}-configuration-set",  # CXmIAssignedConfiguration
            f"PC_{ent}",                        # pcCompoundProfile
            fot,                                 # ViewMasterCatalog
        ]

    return pf1

# ───────────────────── GENERATE & DOWNLOAD ──────────────

def to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Renvoie un fichier Excel sous forme de bytes (in‑memory)."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, index=False)
    return output.getvalue()

if st.button("🚀 Générer le fichier B2B Units"):
    if df_accounts is None or not entreprise or not fot:
        st.warning("🛈 Veuillez déposer un fichier et remplir tous les champs (Entreprise, FoT).")
    else:
        try:
            pf1_df = build_pf1(df_accounts, entreprise.strip(), fot.strip())
            st.success("✅ Fichier généré !")
            st.dataframe(pf1_df.head())

            now = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"B2B Units creation_{entreprise}_{now}.xlsx"

            st.download_button(
                label="📥 Télécharger le fichier Excel",
                data=to_excel_bytes(pf1_df),
                file_name=filename,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        except Exception as exc:
            st.error(f"❌ Erreur : {exc}")

# ──────────────────── FOOTER / AIDE ─────────────────────

with st.expander("ℹ️ Mode d’emploi et déploiement"):
    st.markdown(
        """
        **1. Installation locale**  
        ```bash
        pip install streamlit pandas openpyxl xlsxwriter
        streamlit run app.py
        ```
        **2. Déploiement GitHub / Streamlit Cloud**  
        - Créez un dépôt et déposez `app.py` et un fichier `requirements.txt` contenant :
          ```
          pandas
          streamlit
          openpyxl
          xlsxwriter
          ```
        - Connectez‑vous sur [Streamlit Cloud](https://streamlit.io/cloud), choisissez votre dépôt et référez le fichier `app.py`.  
        - C’est tout ! L’application sera disponible en ligne.
        """,
        unsafe_allow_html=True,
    )
