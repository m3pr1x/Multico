# -*- coding: utf-8 -*-
"""
app.py – Générateur PF1 → PF6 (Multiconnexion) + Export Outlook
• Sélecteur « OCI / cXML » :
    – OCI → colonne **OCIAssignedConfiguration**, pas de PF6
    – cXML → colonne **CXmIAssignedConfiguration** + PF6
• Option « Personal Catalogue »
• Fallback RegEx si libpostal absent
• **Nouveau** : bouton « Brouillon Outlook » qui ouvre un e‑mail dans Outlook Desktop avec les XLSX joints.
"""

from __future__ import annotations
import io, re, sys, tempfile, os
from datetime import datetime
from io import BytesIO

import pandas as pd
import streamlit as st

# ═══════════ libpostal (optionnel) ═══════════
USE_POSTAL = False
try:
    from postal.parser import parse_address  # type: ignore
    USE_POSTAL = True
except ImportError:
    pass  # on continue en regex

# ═══════════ Outlook (optionnel) ═══════════
IS_OUTLOOK = False
try:
    import win32com.client as win32  # type: ignore
    IS_OUTLOOK = True
except ImportError:
    pass  # pas Outlook disponible (ex : Linux, Cloud)

# ═══════════ PAGE CONFIG ═══════════
st.set_page_config(page_title="PF1-PF6 generator", page_icon="📦", layout="wide")
st.title("📦 Outil Multiconnexion")

integration_type = st.radio("Type d’intégration", ["cXML", "OCI"], horizontal=True)

st.markdown(
    "Téléchargez le template, remplissez‑le puis uploadez votre fichier.  \n"
    "Colonnes requises : **Numéro de compte** (7 chiffres), **Raison sociale**, **Adresse**, **ManagingBranch** (4 chiffres)."
)

# ═══════════ TEMPLATE ═══════════
TEMPLATE_COLS = ["Numéro de compte", "Raison sociale", "Adresse", "ManagingBranch"]

example_row = {
    "Numéro de compte": "1234567",
    "Raison sociale":   "EXEMPLE",
    "Adresse":          "10 Rue de la Paix 75002 Paris",
    "ManagingBranch":   "0123",
}

tpl_display = pd.DataFrame([example_row])

tpl_buffer = io.BytesIO()
(pd.DataFrame([{c: "" for c in TEMPLATE_COLS}])
    .to_excel(tpl_buffer, index=False, engine="openpyxl"))
tpl_buffer.seek(0)

with st.expander("📑 Template dfrecu.xlsx"):
    st.download_button(
        "📥 Télécharger le template",
        data=tpl_buffer.getvalue(),
        file_name="dfrecu_template.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )
    st.dataframe(tpl_display, use_container_width=True)

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
pc_enabled = st.radio("Personal Catalogue ?", ["True", "False"], horizontal=True)
pc_name    = ""
if pc_enabled == "True":
    pc_name = st.text_input("Nom du catalogue (sans PC_)", placeholder="CATALOGUE").strip()

# ═══════════ UTILS ═══════════

def read_any(f):
    name = f.name.lower()
    if name.endswith(".csv"):
        for enc in ("utf-8", "latin1", "cp1252"):
            try:
                f.seek(0)
                return pd.read_csv(f, encoding=enc)
            except UnicodeDecodeError:
                f.seek(0)
        raise ValueError("Encodage CSV non reconnu.")
    try:
        return pd.read_excel(f, engine="openpyxl")
    except ImportError:
        return pd.read_excel(f, engine="xlsxwriter")

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
    m = re.match(r"^\s*(?P<num>\d+\w?)\s+(?P<voie>.+?)\s+(?P<cp>\d{5})\s+(?P<ville>.+)$", addr or "", re.I)
    return {"num": m.group("num") if m else "", "voie": m.group("voie") if m else "", "cp": m.group("cp") if m else "", "ville": m.group("ville") if m else "", "pays": "FR"}

def to_xlsx(df: pd.DataFrame) -> bytes:
    buf = BytesIO()
    try:
        with pd.ExcelWriter(buf, engine="openpyxl") as w:
            df.to_excel(w, index=False)
    except ImportError:
        with pd.ExcelWriter(buf, engine="xlsxwriter") as w:
            df.to_excel(w, index=False)
    buf.seek(0)
    return buf.getvalue()

# — Outlook helper —

def create_outlook_draft(attachments: list[tuple[str, bytes]], to_: str = "", subject: str = "", body: str = ""):
    if not IS_OUTLOOK:
        raise RuntimeError("Outlook COM indisponible sur cet environnement.")
    outlook = win32.Dispatch("Outlook.Application")
    mail    = outlook.CreateItem(0)  # olMailItem
    mail.To = to_
    mail.Subject = subject
    mail.Body = body or "Bonjour,\n\nVeuillez trouver les fichiers PF en pièce jointe.\n"

    temp_paths = []
    for name, data in attachments:
        tmp = tempfile.NamedTemporaryFile(delete=False, suffix=name)
        tmp.write(data)
        tmp.close()
        temp_paths.append(tmp.name)
        mail.Attachments.Add(tmp.name)
    mail.Display()

# ——— Sanity‑check helpers ———

def sanitize_numeric(series: pd.Series, width: int):
    s = series.astype(str).str.strip()
    s_padded = s.apply(lambda x: x.zfill(width) if x.isdigit() and len(x) <= width else x)
    invalid = ~s_padded.str.fullmatch(fr"\d{{{width}}}")
    return s_padded, invalid

# ═══════════ BUILD TABLES (identique) ═══════════
# … (fonction build_tables inchangée)

# ═══════════ ACTION ═══════════
if st.button("🚀 Générer"):
    if not (up_file and entreprise and punchout_user and identity and (pc_enabled == "False" or pc_name)):
        st.warning("Remplissez tous les champs requis.")
        st.stop()

    try:
        df_src = read_any(up_file)

        if {"Numéro de compte", "ManagingBranch"} - set(df_src.columns):
            raise ValueError("Colonnes 'Numéro de compte' ou 'ManagingBranch' manquantes.")

        acc_series, bad_acc = sanitize_numeric(df_src["Numéro de compte"], 7)
        man_series, bad_man = sanitize_numeric(df_src["ManagingBranch"], 4)

        if bad_acc.any():
            st.error("❌ Numéro(s) de compte invalide(s).")
            st.dataframe(pd.DataFrame({"Ligne": acc_series.index[bad_acc] + 1, "Numéro": acc_series[bad_acc]}), use_container_width=True)
            st.stop()
        if bad_man.any():
            st.error("❌ ManagingBranch invalide(s).")
            st.dataframe(pd.DataFrame({"Ligne": man_series.index[bad_man] + 1, "ManagingBranch": man_series[bad_man]}), use_container_width=True)
            st.stop()

        df_src["Numéro de compte"] = acc_series
        df_src["ManagingBranch"]   = man_series

        tables = build_tables(df_src)

    except Exception as e:
        st.error(f"❌ {e}")
        st.stop()

    st.success("✅ Fichiers prêts !")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")

    labels = ["PF1", "PF2", "PF3", "PF4", "PF5"] + (["PF6"] if integration_type == "cXML" else [])
    files_bytes = {}

    for label, df in zip(labels, tables):
        data_bytes = to_xlsx(df)
        files_bytes[f"{label}_{entreprise}_{ts}.xlsx"] = data_bytes
        st.download_button(
            f"⬇️ {label}", data=data_bytes,
            file_name=f"{label}_{entreprise}_{ts}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    st.subheader("Aperçu PF1")
    st.dataframe(tables[0].head(), use_container_width=True)

    # ═════ Export Outlook ═════
    st.markdown("---")
    st.header("📧 Exporter via Outlook Desktop")
    if IS_OUTLOOK:
        dest = st.text_input("Destinataire (optionnel)")
        subj = f"Fichiers PF – {entreprise} ({ts})"
        if st.button("Ouvrir un brouillon Outlook", key="outlook_btn"):
            try:
                create_outlook_draft(
                    attachments=list(files_bytes.items()),
                    to_=dest,
                    subject=subj,
                )
                st.success("Brouillon Outlook ouvert.")
            except Exception as err:
                st.error(f"Erreur Outlook : {err}")
    else:
        st.info("Automatisation Outlook indisponible sur cet environnement.")

    if not USE_POSTAL:
        st.info("libpostal non détecté → découpage d’adresse via RegEx.")
