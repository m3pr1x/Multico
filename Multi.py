# -*- coding: utf-8 -*-
"""app.py – Générateur PF1 → PF6  (Personal Catalogue option)
Fonctionne avec ou sans libpostal. Si « Personal Catalogue » est réglé sur False, la colonne
`pcCompoundProfile` reste vide dans PF1 ; sinon elle prend la valeur `PC_<nom saisi>`.
"""

from __future__ import annotations

import io
import re
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ─────────── ESSAI D’IMPORT LIBPOSTAL ───────────
try:
    from postal.parser import parse_address  # type: ignore
    USE_POSTAL = True
except ImportError:
    USE_POSTAL = False

# ─────────── CONFIG PAGE ───────────
st.set_page_config(page_title="PF1-PF6 generator", page_icon="📦", layout="wide")
st.title("📦 Générateur PF1 → PF6")
st.markdown(
    "Téléchargez le template, remplissez‑le puis uploadez votre fichier.  "
    "Colonnes requises : **Numéro de compte**, **Raison sociale**, **Adresse**, **ManagingBranch**."
)

# ─────────── TEMPLATE ───────────
TEMPLATE_COLS = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]
_template = pd.DataFrame([{c: "" for c in TEMPLATE_COLS}])
buf_tpl = io.BytesIO(); _template.to_excel(buf_tpl, index=False); buf_tpl.seek(0)
with st.expander("📑 Template dfrecu.xlsx"):
    st.download_button("📥 Télécharger", buf_tpl.getvalue(), file_name="dfrecu_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(_template)

# ─────────── UPLOAD & PARAMS ───────────
uploaded = st.file_uploader("📄 Fichier dfrecu", type=("csv", "xlsx", "xls"))

c1, c2, c3 = st.columns(3)
with c1:
    entreprise = st.text_input("🏢 Entreprise", placeholder="DALKIA").strip()
with c2:
    punchout_user_id = st.text_input("👤 punchoutUserID", placeholder="user001")
with c3:
    domain = st.selectbox("🌐 Domain", ["NetworkID", "DUNS"])

identity = st.text_input("🆔 Identity")
vm_choice = st.radio("ViewMasterCatalog", ["True", "False"], horizontal=True)

# —— Personal catalogue option ——
pc_enabled = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
pc_name = ""
if pc_enabled == "True":
    pc_name = st.text_input("Nom du catalogue personnel (sans PC_)", placeholder="CATALOGUE")

# ─────────── UTILS ───────────

def read_any(f):
    n = f.name.lower()
    if n.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                f.seek(0); return pd.read_csv(f, encoding=enc)
            except UnicodeDecodeError: f.seek(0)
        raise ValueError("Encodage CSV non reconnu")
    return pd.read_excel(f, engine="openpyxl")


def split_address(addr: str):
    if USE_POSTAL:
        res = {k: "" for k in ("num","voie","cp","ville")}; res["pays"] = "FR"
        for v,l in parse_address(addr or ""):
            if l=="house_number": res["num"]=v
            elif l in {"road","footway","path"}: res["voie"]=v
            elif l=="postcode": res["cp"]=v
            elif l in {"city","town","village","suburb"}: res["ville"]=v
            elif l=="country": res["pays"]=v
        return res
    m=re.match(r"^\s*(?P<num>\d+\w?)\s+(?P<voie>.+?)\s+(?P<cp>\d{5})\s+(?P<ville>.+)$",addr or "",re.I)
    return {"num":m.group("num") if m else "","voie":m.group("voie") if m else "","cp":m.group("cp") if m else "","ville":m.group("ville") if m else "","pays":"FR"}


def to_xlsx(df):
    b=BytesIO();
    with pd.ExcelWriter(b,engine="openpyxl") as w: df.to_excel(w,index=False)
    b.seek(0);return b.getvalue()

# ─────────── BUILD TABLES ───────────

def build_tables(df: pd.DataFrame):
    if (miss := set(TEMPLATE_COLS)-set(df.columns)): raise ValueError(f"Colonnes manquantes : {', '.join(miss)}")
    sealed="false"
    pf1=pd.DataFrame(columns=["uid","name","locName","CXmIAssignedConfiguration","pcCompoundProfile","ViewMasterCatalog"])
    pf2=pd.DataFrame(columns=["B2B Unit","ADRESSE / Numéro de rue","ADRESSE / rue","ADRESSEE Code postal","ADRESSE / Ville","ADRESSE / Pays/Région","INFORMATIONS D'ADRESSE SUPPLÉMENTAIRES / Téléphone 1"])
    pf3=pd.DataFrame(columns=["B2BUnitID","itemtype","managingBranches","punchoutUserID","sealed"])
    pf4=pd.DataFrame(columns=["aliasName","branch","punchoutUserID","sealed"])
    pf5=pd.DataFrame(columns=["B2BUnitID","punchoutUserID"])
    pf6=pd.DataFrame(columns=["number","domain","identity"])

    profile_val = f"PC_{pc_name.strip()}" if pc_enabled=="True" and pc_name.strip() else ""

    for _,r in df.iterrows():
        code=r["Numéro de compte"]; man=r["ManagingBranch"]
        # PF1
        pf1.loc[len(pf1)] = [code,r["Raison sociale"],r["Adresse"],f"frx-variant-{entreprise}-configuration-set",profile_val,vm_choice]
        # PF2
        a=split_address(r["Adresse"])
        pf2.loc[len(pf2)]=[code,a["num"],a["voie"],a["cp"],a["ville"],a["pays"],""]
        # PF3‑PF6
        pf3.loc[len(pf3)]=[code,"PunchoutAccountAndBranchAssociation",man,punchout_user_id,sealed]
        pf4.loc[len(pf4)]=[man,man,punchout_user_id,sealed]
        pf5.loc[len(pf5)]=[code,punchout_user_id]
        pf6.loc[len(pf6)]=[code,domain,identity]
    return pf1,pf2,pf3,pf4,pf5,pf6

# ─────────── ACTION ───────────
if st.button("🚀 Générer PF1‑PF6"):
    if not (uploaded and entreprise and punchout_user_id and identity and (pc_enabled=="False" or pc_name.strip())):
        st.warning("Veuillez charger un fichier et compléter tous les champs (y compris le nom du catalogue).")
        st.stop()
    try:
        src=read_any(uploaded)
        tables=build_tables(src)
    except Exception as e:
        st.error(f"❌ {e}"); st.stop()

    st.success("✅ Fichiers prêts !")
    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    for tag,df in zip(("PF1","PF2","PF3","PF4","PF5","PF6"),tables):
        st.download_button(f"⬇️ {tag}",to_xlsx(df),file_name=f"{tag}_{entreprise}_{ts}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.subheader("Aperçu PF1")
    st.dataframe(tables[0].head())

    if not USE_POSTAL:
        st.info("libpostal non détecté → découpage d’adresse via RegEx (moins précis).")
