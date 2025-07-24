import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Extracteur de feuilles Excel EHPAD", layout="centered")

st.title("📊 Extracteur de feuilles Excel - EHPAD")
st.markdown("Dépose un fichier `.xlsm`, choisis ton type d'export, et récupère automatiquement les fichiers formatés pour Qlik Sense.")

uploaded_file = st.file_uploader("Déposez ici votre fichier Excel", type=["xlsm"])

export_type = st.radio("Que voulez-vous exporter ?", [
    "Les 3 feuilles d'historique (Global / Local / Projection)",
    "Seulement la feuille Export_Qlik"
])

def formater_excel(df, sheetname):
    """Retourne un BytesIO contenant le fichier Excel bien formaté"""
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheetname, index=False)

        workbook = writer.book
        worksheet = writer.sheets[sheetname]
        format_pct = workbook.add_format({'num_format': '0.00%'})

        for idx, col in enumerate(df.columns):
            # Largeur auto
            max_len = max([len(str(col))] + [len(str(val)) for val in df[col]])
            worksheet.set_column(idx, idx, max_len + 2)

            # Format pourcentages
            if "Pourcent" in col or "Marge" in col:
                worksheet.set_column(idx, idx, max_len + 2, format_pct)

    buffer.seek(0)
    return buffer

def analyser_et_log(df, nom_feuille):
    """Retourne le texte à écrire dans le log pour une feuille"""
    lignes_utiles = len(df.dropna(how="all"))
    log = f"{nom_feuille} : {lignes_utiles} lignes utiles exportées.\n"
    if df.isnull().any().any():
        log += f"  ⚠ Contient des valeurs manquantes.\n"
    if "Année" in df.columns:
        if df["Année"].min() < 2020 or df["Année"].max() > 2026:
            log += f"  ⚠ Années hors limites [2020-2026].\n"
    log += "\n"
    return log, lignes_utiles

if uploaded_file and st.button("Lancer l'export"):

    xls = pd.ExcelFile(uploaded_file)
    zip_buffer = BytesIO()
    log_txt = ""
    total_lignes = 0

    with zipfile.ZipFile(zip_buffer, "a") as zf:

        if export_type == "Les 3 feuilles d'historique (Global / Local / Projection)":
            feuilles_cibles = {
                "Historique_Global": "Historique_Global",
                "Historique_Local": "Historique_Local",
                "Historique_Projection": "Historique_Projection"
            }
        else:
            feuilles_cibles = {"Export_Qlik": "Export_Qlik"}

        for nom, feuille in feuilles_cibles.items():
            try:
                df = xls.parse(feuille)
                buffer = formater_excel(df, "Export")
                zf.writestr(f"{nom}.xlsx", buffer.read())

                log_entry, lignes = analyser_et_log(df, nom)
                log_txt += log_entry
                total_lignes += lignes
            except Exception as e:
                log_txt += f"{nom} :  Erreur lors du traitement ({str(e)})\n\n"

        # Ajout du rapport log
        log_txt += f"Total lignes exportées : {total_lignes}\n"
        zf.writestr("log_export.txt", log_txt)

    st.success("✅ Export terminé avec succès.")
    st.download_button(
        "📦 Télécharger les fichiers exportés (.zip)",
        data=zip_buffer.getvalue(),
        file_name=f"export_Qlik_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip"
    )

    with st.expander("📄 Aperçu du rapport d'export"):
        st.text(log_txt)
