# -*- coding: utf-8 -*-
"""
Générateur PF1 → PF6
• Option « Personal Catalogue » : si cochée, pcCompoundProfile = PC_<nom_saisi>, sinon champ vide.
• Dépendances : pandas, streamlit, openpyxl   (postal facultatif).
"""

from __future__ import annotations
import io, re
from datetime import datetime
from io import BytesIO
import pandas as pd
import streamlit as st

# ─── libpostal (optionnel) ───
try:
    from postal.parser import parse_address  # type: ignore
    USE_POSTAL = True
except ImportError:
    USE_POSTAL = False

# ─── Page ───
st.set_page_config(page_title="PF1-PF6 generator", page_icon="📦", layout="wide")
st.title("📦 Générateur PF1 → PF6")
st.markdown(
    "Colonnes requises : **Numéro de compte**, **Raison sociale**, **Adresse**, **ManagingBranch**."
)

# ─── Template téléchargeable ───
TEMPLATE_COLS = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]
tpl = pd.DataFrame([{c: "" for c in TEMPLATE_COLS}])
buf = io.BytesIO(); tpl.to_excel(buf, index=False); buf.seek(0)
with st.expander("📑 Template dfrecu.xlsx"):
    st.download_button("📥 Télécharger", buf.getvalue(), file_name="dfrecu_template.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.dataframe(tpl)

# ─── Upload & paramètres ───
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

# Personal Catalogue
pc_enabled = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
pc_name    = st.text_input("Nom du catalogue (sans PC_)",
                           disabled=(pc_enabled == "False")).strip()

# ─── Utils ───
def read_any(f):
    n=f.name.lower()
    if n.endswith(".csv"):
        for enc in ("utf-8","latin1","cp1252"):
            try: f.seek(0); return pd.read_csv(f,encoding=enc)
            except UnicodeDecodeError: f.seek(0)
        raise ValueError("Encodage CSV non reconnu.")
    return pd.read_excel(f, engine="openpyxl")

def split_address(addr:str):
    if USE_POSTAL:
        d={"num":"","voie":"","cp":"","ville":"","pays":"FR"}
        for v,l in parse_address(addr or ""):
            if l=="house_number": d["num"]=v
            elif l in {"road","footway","path"}: d["voie"]=v
            elif l=="postcode": d["cp"]=v
            elif l in {"city","town","village","suburb"}: d["ville"]=v
            elif l=="country": d["pays"]=v
        return d
    m=re.match(r"^\s*(?P<num>\d+\w?)\s+(?P<voie>.+?)\s+(?P<cp>\d{5})\s+(?P<ville>.+)$",
               addr or "", re.I)
    return {
        "num": m.group("num") if m else "",
        "voie": m.group("voie") if m else "",
        "cp": m.group("cp") if m else "",
        "ville": m.group("ville") if m else "",
        "pays": "FR",
    }

def to_xlsx(df:pd.DataFrame)->bytes:
    b=BytesIO()
    with pd.ExcelWriter(b, engine="openpyxl") as w: df.to_excel(w, index=False)
    b.seek(0); return b.getvalue()

# ─── Build tables ───
def build(df:pd.DataFrame):
    miss=set(TEMPLATE_COLS)-set(df.columns)
    if miss: raise ValueError(f"Colonnes manquantes : {', '.join(miss)}")

    sealed = "false"
    pc_value = f"PC_{pc_name}" if pc_enabled=="True" and pc_name else ""

    pf1 = pd.DataFrame(columns=["uid","name","locName","CXmIAssignedConfiguration","pcCompoundProfile","ViewMasterCatalog"])
    pf2 = pd.DataFrame(columns=["B2B Unit","ADRESSE / Numéro de rue","ADRESSE / rue","ADRESSEE Code postal","ADRESSE / Ville","ADRESSE / Pays/Région","INFORMATIONS D'ADRESSE SUPPLÉMENTAIRES / Téléphone 1"])
    pf3 = pd.DataFrame(columns=["B2BUnitID","itemtype","managingBranches","punchoutUserID","sealed"])
    pf4 = pd.DataFrame(columns=["aliasName","branch","punchoutUserID","sealed"])
    pf5 = pd.DataFrame(columns=["B2BUnitID","punchoutUserID"])
    pf6 = pd.DataFrame(columns=["number","domain","identity"])

    for _, r in df.iterrows():
        code = r["Numéro de compte"]; man = r["ManagingBranch"]
        # PF1
        pf1.loc[len(pf1)] = [
            code, r["Raison sociale"], r["Adresse"],
            f"frx-variant-{entreprise}-configuration-set",
            pc_value,
            vm_choice,
        ]
        # PF2
        a=split_address(r["Adresse"])
        pf2.loc[len(pf2)] = [code,a["num"],a["voie"],a["cp"],a["ville"],a["pays"],""]
        # PF3-PF6
        pf3.loc[len(pf3)] = [code,"PunchoutAccountAndBranchAssociation",man,punchout_user,sealed]
        pf4.loc[len(pf4)] = [man,man,punchout_user,sealed]
        pf5.loc[len(pf5)] = [code,punchout_user]
        pf6.loc[len(pf6)] = [code,domain,identity]
    return pf1,pf2,pf3,pf4,pf5,pf6

# ─── Action ───
if st.button("🚀 Générer PF1-PF6"):
    if not (uploaded and entreprise and punchout_user and identity and
            (pc_enabled=="False" or pc_name)):
        st.warning("Veuillez remplir tous les champs obligatoires.")
        st.stop()

    try:
        src=read_any(uploaded)
        tables=build(src)
    except Exception as e:
        st.error(f"❌ {e}"); st.stop()

    st.success("✅ Fichiers créés")
    ts=datetime.now().strftime("%Y%m%d_%H%M%S")
    for tag,df in zip(("PF1","PF2","PF3","PF4","PF5","PF6"),tables):
        st.download_button(f"⬇️ {tag}", to_xlsx(df),
                           file_name=f\"{tag}_{entreprise}_{ts}.xlsx\")

    st.dataframe(tables[0].head())
    if not USE_POSTAL:
        st.info("libpostal non détecté → découpage d’adresse via RegEx.")
