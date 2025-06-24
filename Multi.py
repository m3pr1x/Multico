# -*- coding: utf-8 -*-
"""
app.py – Générateur PF1 → PF6
• Sélecteur « OCI / cXML » :
    – OCI  → colonne OCIAssignedConfiguration, pas de PF6
    – cXML → colonne CXmIAssignedConfiguration + PF6
• Option « Personal Catalogue »
• Fallback RegEx si libpostal absent
"""

from __future__ import annotations
import io, re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ═══════════ libpostal optionnel ═══════════
try:
    from postal.parser import parse_address  # type: ignore
    USE_POSTAL = True
except ImportError:
    USE_POSTAL = False

# ═══════════ PAGE CONFIG ═══════════
st.set_page_config(page_title="PF1-PF6 generator", page_icon="📦", layout="wide")
st.title("📦 Générateur PF1 → PF6")

integration_type = st.radio("Type d’intégration", ["cXML", "OCI"], horizontal=True)

st.markdown(
    "Téléchargez le template, remplissez-le puis uploadez votre fichier.  \n"
    "Colonnes requises : **Numéro de compte**, **Raison sociale**, **Adresse**, **ManagingBranch**."
)

# ═══════════ TEMPLATE ═══════════
TEMPLATE_COLS = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]
tpl = pd.DataFrame([{c: "" for c in TEMPLATE_COLS}])
buf = io.BytesIO()
tpl.to_excel(buf, index=False)
buf.seek(0)
with st.expander("📑 Template dfrecu.xlsx"):
    st.download_button(
        "📥 Télécharger le template",
        data=buf.getvalue(),
        file_name="dfrecu_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.dataframe(tpl)

# ═══════════ UPLOAD & PARAMS ═══════════
up_file = st.file_uploader("📄 Fichier dfrecu", type=("csv", "xlsx", "xls"))

col1, col2, col3 = st.columns(3)
with col1:
    entreprise = st.text_input("🏢 Entreprise").strip()
with col2:
    punchout_user = st.text_input("👤 punchoutUserID")
with col3:
    domain = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])

identity   = st.text_input("🆔 Identity")
vm_choice  = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)

# — Personal Catalogue
pc_enabled = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
pc_name    = ""
if pc_enabled == "True":
    pc_name = st.text_input("Nom du catalogue (sans PC_)", placeholder="CATALOGUE").strip()

# ═══════════ UTILS ═══════════
def read_any(f):
    name = f.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                f.seek(0); return pd.read_csv(f, encoding=enc)
            except UnicodeDecodeError:
                f.seek(0)
        raise ValueError("Encodage CSV non reconnu.")
    return pd.read_excel(f, engine="openpyxl")

def split_address(addr: str) -> dict:
    if USE_POSTAL:
        d = {"num": "", "voie": "", "cp": "", "ville": "", "pays": "FR"}
        for val, lab in parse_address(addr or ""):
            if lab == "house_number": d["num"] = val
            elif lab in {"road", "footway", "path"}: d["voie"] = val
            elif lab == "postcode": d["cp"] = val
            elif lab in {"city", "town", "village", "suburb"}: d["ville"] = val
            elif lab == "country": d["pays"] = val
        return d
    m = re.match(r"^\s*(?P<num>\d+\w?)\s+(?P<voie>.+?)\s+(?P<cp>\d{5})\s+(?P<ville>.+)$",
                 addr or "", re.I)
    return {
        "num":   m.group("num")   if m else "",
        "voie":  m.group("voie")  if m else "",
        "cp":    m.group("cp")    if m else "",
        "ville": m.group("ville") if m else "",
        "pays": "FR",
    }

def to_xlsx(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as w:
        df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()

# ═══════════ BUILD TABLES ═══════════
def build_tables(df: pd.DataFrame):
    missing = set(TEMPLATE_COLS) - set(df.columns)
    if missing:
        raise ValueError(f"Colonnes manquantes : {', '.join(sorted(missing))}")

    sealed     = "false"
    pc_value   = f"PC_{pc_name}" if pc_enabled == "True" and pc_name else ""
    config_col = "OCIAssignedConfiguration" if integration_type == "OCI" else "CXmIAssignedConfiguration"

    # PF1 columns adaptées
    pf1 = pd.DataFrame(columns=[
        "uid", "name", "locName",
        config_col,                # colonne variable
        "pcCompoundProfile", "ViewMasterCatalog",
    ])
    pf2 = pd.DataFrame(columns=[
        "B2B Unit", "ADRESSE / Numéro de rue", "ADRESSE / rue",
        "ADRESSEE Code postal", "ADRESSE / Ville", "ADRESSE / Pays/Région",
        "INFORMATIONS D'ADRESSE SUPPLÉMENTAIRES / Téléphone 1",
    ])
    pf3 = pd.DataFrame(columns=["B2BUnitID", "itemtype", "managingBranches", "punchoutUserID", "sealed"])
    pf4 = pd.DataFrame(columns=["aliasName", "branch", "punchoutUserID", "sealed"])
    pf5 = pd.DataFrame(columns=["B2BUnitID", "punchoutUserID"])
    # PF6 seulement si cXML
    if integration_type == "cXML":
        pf6 = pd.DataFrame(columns=["number", "domain", "identity"])

    for _, r in df.iterrows():
        code = r["Numéro de compte"]; man = r["ManagingBranch"]
        addr = split_address(r["Adresse"])

        # PF1
        pf1.loc[len(pf1)] = [
            code, r["Raison sociale"], r["Adresse"],
            f"frx-variant-{entreprise}-configuration-set",
            pc_value, vm_choice,
        ]
        # PF2
        pf2.loc[len(pf2)] = [code, addr["num"], addr["voie"], addr["cp"], addr["ville"], addr["pays"], ""]
        # PF3-5
        pf3.loc[len(pf3)] = [code, "PunchoutAccountAndBranchAssociation", man, punchout_user, sealed]
        pf4.loc[len(pf4)] = [man, man, punchout_user, sealed]
        pf5.loc[len(pf5)] = [code, punchout_user]
        # PF6 (cXML)
        if integration_type == "cXML":
            pf6.loc[len(pf6)] = [code, domain, identity]

    if integration_type == "cXML":
        return pf1, pf2, pf3, pf4, pf5, pf6
    return pf1, pf2, pf3, pf4, pf5  # OCI

# ═══════════ ACTION ═══════════
if st.button("🚀 Générer"):
    if not (up_file and entreprise and punchout_user and identity
            and (pc_enabled == "False" or pc_name)):
        st.warning("Remplissez tous les champs requis.")
        st.stop()

    try:
        df_src = read_any(up_file)
        tables = build_tables(df_src)
    except Exception as err:
        st.error(f"❌ {err}")
        st.stop()

    st.success("✅ Fichiers prêts !")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    tags = ["PF1", "PF2", "PF3", "PF4", "PF5"] + (["PF6"] if integration_type == "cXML" else [])
    for tag, df in zip(tags, tables):
        st.download_button(
            f"⬇️ {tag}",
            data=to_xlsx(df),
            file_name=f"{tag}_{entreprise}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.subheader("Aperçu PF1")
    st.dataframe(tables[0].head())

    if not USE_POSTAL:
        st.info("libpostal non détecté → découpage d’adresse via RegEx.")
