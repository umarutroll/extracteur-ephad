import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from datetime import datetime

st.set_page_config(page_title="Extracteur de feuilles Excel EHPAD", layout="centered")

st.title("📊 Extracteur de feuilles Excel - EHPAD")
st.markdown("Dépose un fichier `.xlsm`, choisis ton type d'export, et récupère automatiquement les fichiers formatés pour Qlik Sense.")

uploaded_file = st.file_uploader("📁 Déposez ici votre fichier Excel", type=["xlsm"])

export_type = st.radio("Que voulez-vous exporter ?", [
    "Les 3 feuilles d'historique (Global / Local / Projection)",
    "Seulement la feuille Export_Qlik"
])

def formater_excel(df, sheetname):
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheetname, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheetname]
        format_pct = workbook.add_format({'num_format': '0.00%'})
        for idx, col in enumerate(df.columns):
            max_len = max([len(str(col))] + [len(str(val)) for val in df[col]])
            worksheet.set_column(idx, idx, max_len + 2)
            if "Pourcent" in col or "Marge" in col:
                worksheet.set_column(idx, idx, max_len + 2, format_pct)
    buffer.seek(0)
    return buffer

def analyser_et_log(df, nom_feuille):
    lignes_utiles = len(df.dropna(how="all"))
    log = f"{nom_feuille} : {lignes_utiles} lignes utiles exportées.\n"

    if df.isnull().any().any():
        log += f"  ⚠ Contient des valeurs manquantes.\n"

    # Score de qualité
    pct_nan = df.isnull().sum().sum() / df.size
    log += f"  🔎 Score de complétude : {round((1 - pct_nan) * 100)}%\n"

    # Alerte métier
    if "Résultat" in df.columns and (df["Résultat"] < 0).any():
        log += f"  ⚠ Résultats négatifs détectés\n"
    if "Marge" in df.columns and (df["Marge"] < 0.1).any():
        log += f"  📉 Marges inférieures à 10%\n"

    # Vérification année
    if "Annee" in df.columns:
        if df["Annee"].min() < 2022 or df["Annee"].max() > 2026:
            log += f"  ⚠ Années hors bornes attendues [2022-2026]\n"

    log += "\n"
    return log, lignes_utiles

if uploaded_file and st.button("🚀 Lancer l'export"):

    xls = pd.ExcelFile(uploaded_file)
    zip_buffer = BytesIO()
    log_txt = f"Export réalisé le {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
    total_lignes = 0
    dernier_df = None  # pour l'aperçu plus bas

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
                dernier_df = df.copy()  # garder pour l'aperçu
                buffer = formater_excel(df, "Export")
                zf.writestr(f"{nom}.xlsx", buffer.read())

                log_entry, lignes = analyser_et_log(df, nom)
                log_txt += log_entry
                total_lignes += lignes
            except Exception as e:
                log_txt += f"{nom} :  ❌ Erreur lors du traitement ({str(e)})\n\n"

        log_txt += f"✅ Total lignes exportées : {total_lignes}\n"
        zf.writestr("log_export.txt", log_txt)

    st.success("✅ Export terminé avec succès.")
    st.download_button(
        "📦 Télécharger les fichiers exportés (.zip)",
        data=zip_buffer.getvalue(),
        file_name=f"export_Qlik_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
        mime="application/zip"
    )

    st.download_button("📄 Télécharger uniquement le log", data=log_txt, file_name="log_export.txt")

    with st.expander("📋 Aperçu du rapport d'export"):
        st.text(log_txt)

    if dernier_df is not None:
        with st.expander("🔍 Aperçu des données extraites"):
            st.dataframe(dernier_df.head(10))

st.markdown("---")
st.caption("Développé par Rémy Laguerre – CY Tech – Stage Pilotage Financier – Juillet 2025")

