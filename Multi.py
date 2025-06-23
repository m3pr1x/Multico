# -*- coding: utf-8 -*-
"""
app.py – Générateur PF1 → PF6
• Fonctionne avec ou sans libpostal.
• Dépendances Python : pandas, streamlit, openpyxl   (postal est facultatif).
"""

from __future__ import annotations

import io
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ───────────────────────── ESSAI D’IMPORT LIBPOSTAL ─────────────────────────
try:
    from postal.parser import parse_address  # type: ignore
    USE_POSTAL = True
except ImportError:
    USE_POSTAL = False

# ───────────────────────── CONFIG PAGE ─────────────────────────
st.set_page_config(page_title="PF1-PF6 generator", page_icon="📦", layout="wide")
st.title("📦 Générateur PF1 → PF6")
st.markdown(
    "Téléchargez le template, remplissez-le puis uploadez votre fichier.  "
    "Champs requis : **Numéro de compte**, **Raison sociale**, **Adresse**, **ManagingBranch**."
)

# ───────────────────────── TEMPLATE TELECHARGEABLE ─────────────────────────
TEMPLATE_COLS = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]
template_df = pd.DataFrame([{c: "" for c in TEMPLATE_COLS}])
buf_tpl = io.BytesIO()
template_df.to_excel(buf_tpl, index=False)
buf_tpl.seek(0)

with st.expander("📑 Template dfrecu.xlsx"):
    st.download_button(
        "📥 Télécharger le template",
        data=buf_tpl.getvalue(),
        file_name="dfrecu_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.dataframe(template_df)

# ───────────────────────── UPLOAD & PARAMETRES ─────────────────────────
uploaded = st.file_uploader("📄 Fichier dfrecu (CSV ou Excel)", type=("csv", "xlsx", "xls"))

col1, col2, col3 = st.columns(3)
with col1:
    entreprise = st.text_input("🏢 Entreprise", placeholder="DALKIA").strip()
with col2:
    punchout_user_id = st.text_input("👤 punchoutUserID", placeholder="user001")
with col3:
    domain = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])

identity = st.text_input("🆔 Identity")
vm_choice = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)

# ───────────────────────── FONCTIONS UTILITAIRES ─────────────────────────
def read_any(file):
    name = file.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                file.seek(0)
                return pd.read_csv(file, encoding=enc)
            except UnicodeDecodeError:
                file.seek(0)
        raise ValueError("Encodage CSV non reconnu")
    return pd.read_excel(file, engine="openpyxl")


def split_address(addr: str) -> dict:
    """Retourne {num, voie, cp, ville, pays}. Fallback regex si libpostal absent."""
    if USE_POSTAL:
        res = {"num": "", "voie": "", "cp": "", "ville": "", "pays": "FR"}
        for val, lab in parse_address(addr or ""):
            if lab == "house_number":
                res["num"] = val
            elif lab in {"road", "footway", "path"}:
                res["voie"] = val
            elif lab == "postcode":
                res["cp"] = val
            elif lab in {"city", "town", "village", "suburb"}:
                res["ville"] = val
            elif lab == "country":
                res["pays"] = val
        return res

    m = re.match(
        r"^\s*(?P<num>\d+\w?)\s+(?P<voie>.+?)\s+(?P<cp>\d{5})\s+(?P<ville>.+)$",
        addr or "",
        flags=re.I,
    )
    return {
        "num": m.group("num") if m else "",
        "voie": m.group("voie") if m else "",
        "cp": m.group("cp") if m else "",
        "ville": m.group("ville") if m else "",
        "pays": "FR",
    }


def to_excel_bytes(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()

# ───────────────────────── GÉNÉRATION DES TABLES ─────────────────────────
def build_tables(df: pd.DataFrame):
    if missing := set(TEMPLATE_COLS) - set(df.columns):
        raise ValueError(f"Colonnes manquantes : {', '.join(sorted(missing))}")

    sealed = "false"

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
        managing = r["ManagingBranch"]

        pf1.loc[len(pf1)] = [
            code, r["Raison sociale"], r["Adresse"],
            f"frx-variant-{entreprise}-configuration-set",
            f"PC_{entreprise}", vm_choice,
        ]

        a = split_address(r["Adresse"])
        pf2.loc[len(pf2)] = [code, a["num"], a["voie"], a["cp"], a["ville"], a["pays"], ""]

        pf3.loc[len(pf3)] = [code, "PunchoutAccountAndBranchAssociation", managing, punchout_user_id, sealed]
        pf4.loc[len(pf4)] = [managing, managing, punchout_user_id, sealed]
        pf5.loc[len(pf5)] = [code, punchout_user_id]
        pf6.loc[len(pf6)] = [code, domain, identity]

    return pf1, pf2, pf3, pf4, pf5, pf6

# ───────────────────────── ACTION ─────────────────────────
if st.button("🚀 Générer PF1-PF6"):
    if not (uploaded and entreprise and punchout_user_id and identity):
        st.warning("Veuillez charger un fichier et remplir tous les champs.")
        st.stop()

    try:
        src_df = read_any(uploaded)
        pf1, pf2, pf3, pf4, pf5, pf6 = build_tables(src_df)
    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()

    st.success("✅ Fichiers prêts !")
    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    for tag, df in zip(
        ("PF1", "PF2", "PF3", "PF4", "PF5", "PF6"),
        (pf1, pf2, pf3, pf4, pf5, pf6),
    ):
        st.download_button(
            f"⬇️ {tag}",
            data=to_excel_bytes(df),
            file_name=f"{tag}_{entreprise}_{timestamp}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.subheader("Aperçu PF1")
    st.dataframe(pf1.head())

    if not USE_POSTAL:
        st.info("⚠️ libpostal non détecté → découpage d’adresse basé sur RegEx (moins précis).")
