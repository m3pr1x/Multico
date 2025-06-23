# -*- coding: utf-8 -*-
"""app.py – Générateur tables PF1…PF6
Déployable sur Streamlit Cloud. Ajoutez au dépôt :
    requirements.txt ->
        pandas
        streamlit
        openpyxl
        postal
        libpostal
Si libpostal n’est disponible que via apt/brew, ajoutez la commande
correspondante dans le README du repo.
"""

from __future__ import annotations

import io
from datetime import datetime
from io import BytesIO
from pathlib import Path

import pandas as pd
import streamlit as st
from postal.parser import parse_address  # libpostal (wrapper Python)

# ───────────────────────── DEP CHECK ─────────────────────────
try:
    import openpyxl  # noqa: F401
except ImportError as e:
    raise ImportError("openpyxl requis : pip install openpyxl") from e

ENGINE = "openpyxl"

# ───────────────────────── PAGE SETUP ─────────────────────────
st.set_page_config(page_title="PF1‑PF6 generator", page_icon="📦", layout="wide")

st.title("📦 Générateur PF1 → PF6")
st.markdown(
    "Uploadez un fichier CSV ou Excel respectant **le template ci‑dessous**.\n"
    "Remplissez les champs, puis cliquez sur *Générer* pour télécharger les 6 fichiers Excel.")

# ───────────────────────── TEMPLATE SECTION ─────────────────────────
TEMPLATE_COLS = [
    "Numéro de compte",   # ex: 12345678901
    "Raison sociale",     # ex: ALPHA SARL
    "Adresse",            # ex: 10 Rue de la Paix 75002 Paris
    "ManagingBranch",     # ex: PARIS01
]

template_df = pd.DataFrame([{c: "" for c in TEMPLATE_COLS}])
buf_template = io.BytesIO()
template_df.to_excel(buf_template, index=False, engine=ENGINE)
buf_template.seek(0)

with st.expander("📑 Télécharger le template dfrEcu.xlsx"):
    st.download_button("📥 Template Excel", buf_template, file_name="dfrecu_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(template_df)

# ───────────────────────── UPLOAD & PARAMS ─────────────────────────
file = st.file_uploader("📄 Fichier dfrecu", type=("csv", "xlsx", "xls"))

colA, colB, colC = st.columns(3)
with colA:
    entreprise = st.text_input("🏢 Entreprise", placeholder="DALKIA").strip()
with colB:
    punchout_user_id = st.text_input("👤 punchoutUserID", placeholder="user001")
with colC:
    domain = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])

identity = st.text_input("🆔 Identity (string)")
vm_choice = st.radio("ViewMasterCatalog?", ["True", "False"], horizontal=True)

# ───────────────────────── HELPERS ─────────────────────────

def read_any(uploaded):
    name = uploaded.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                uploaded.seek(0)
                return pd.read_csv(uploaded, encoding=enc)
            except UnicodeDecodeError:
                uploaded.seek(0)
        st.error("Encodage CSV non reconnu.")
    else:
        return pd.read_excel(uploaded, engine=ENGINE)


def split_address(addr: str) -> dict:
    parts = {"num": "", "voie": "", "cp": "", "ville": "", "pays": "FR"}
    for val, label in parse_address(addr or ""):
        if label == "house_number":
            parts["num"] = val
        elif label in {"road", "footway", "path"}:
            parts["voie"] = val
        elif label == "postcode":
            parts["cp"] = val
        elif label in {"city", "town", "village", "suburb"}:
            parts["ville"] = val
        elif label == "country":
            parts["pays"] = val
    return parts


def to_excel_bytes(df):
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine=ENGINE) as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()

# ───────────────────────── CORE GENERATE ─────────────────────────

def generate_tables(df: pd.DataFrame):
    required = set(TEMPLATE_COLS)
    if missing := required - set(df.columns):
        raise ValueError(f"Colonnes manquantes : {', '.join(sorted(missing))}")

    sealed = "false"

    # PF1
    pf1 = pd.DataFrame(columns=[
        "uid", "name", "locName",
        "CXmIAssignedConfiguration", "pcCompoundProfile", "ViewMasterCatalog",
    ])

    pf2 = pd.DataFrame(columns=[
        "B2B Unit", "ADRESSE / Numéro de rue", "ADRESSE / rue",
        "ADRESSEE Code postal", "ADRESSE / Ville", "ADRESSE / Pays/Région",
        "INFORMATIONS D'ADRESSE SUPPLÉMENTAIRES / Téléphone 1",
    ])

    pf3 = pd.DataFrame(columns=[
        "B2BUnitID", "itemtype", "managingBranches", "punchoutUserID", "sealed",
    ])

    pf4 = pd.DataFrame(columns=["aliasName", "branch", "punchoutUserID", "sealed"])
    pf5 = pd.DataFrame(columns=["B2BUnitID", "punchoutUserID"])
    pf6 = pd.DataFrame(columns=["number", "domain", "identity"])

    for _, r in df.iterrows():
        code = r["Numéro de compte"]
        # PF1
        pf1.loc[len(pf1)] = [
            code, r["Raison sociale"], r["Adresse"],
            f"frx-variant-{entreprise}-configuration-set",
            f"PC_{entreprise}", vm_choice,
        ]
        # PF2
        addr = split_address(r["Adresse"])
        pf2.loc[len(pf2)] = [
            code, addr["num"], addr["voie"], addr["cp"], addr["ville"], addr["pays"], "",
        ]
        # PF3
        managing = r["ManagingBranch"]
        pf3.loc[len(pf3)] = [
            code, "PunchoutAccountAndBranchAssociation", managing, punchout_user_id, sealed,
        ]
        # PF4
        pf4.loc[len(pf4)] = [managing, managing, punchout_user_id, sealed]
        # PF5
        pf5.loc[len(pf5)] = [code, punchout_user_id]
        # PF6
        pf6.loc[len(pf6)] = [code, domain, identity]

    return pf1, pf2, pf3, pf4, pf5, pf6

# ───────────────────────── ACTION BUTTON ─────────────────────────

if st.button("🚀 Générer PF1‑PF6"):
    if not (file and entreprise and punchout_user_id and identity):
        st.warning("Veuillez charger un fichier et renseigner tous les champs.")
        st.stop()

    try:
        src = read_any(file)
        pf1, pf2, pf3, pf4, pf5, pf6 = generate_tables(src)
    except Exception as e:
        st.error(f"❌ Erreur : {e}")
        st.stop()

    st.success("✅ Tables générées !")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    for name, df in zip(
        ("PF1", "PF2", "PF3", "PF4", "PF5", "PF6"),
        (pf1, pf2, pf3, pf4, pf5, pf6),
    ):
        file_bytes = to_excel_bytes(df)
        fname = f"{name}_{entreprise}_{ts}.xlsx".replace(" ", "_")
        st.download_button(f"⬇️ {name}", file_bytes, file_name=fname,
                           mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    # Aperçu PF1
    st.subheader("Aperçu PF1")
    st.dataframe(pf1.head(20), use_container_width=True)
