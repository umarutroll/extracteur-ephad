import streamlit as st
import pandas as pd
import zipfile
from io import BytesIO
from datetime import datetime

# -------------------- Fonctions Utilitaires --------------------

def formater_excel(df, sheetname):
    """
    Formate un DataFrame dans un fichier Excel en mémoire (BytesIO), avec :
    - Largeurs de colonnes automatiques
    - Format de pourcentage appliqué sur les colonnes pertinentes
    """
    buffer = BytesIO()
    with pd.ExcelWriter(buffer, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name=sheetname, index=False)
        workbook = writer.book
        worksheet = writer.sheets[sheetname]

        # Format personnalisé pour les pourcentages
        format_pct = workbook.add_format({'num_format': '0.00%'})
        
        # Ajustement automatique des colonnes + format si besoin
        for idx, col in enumerate(df.columns):
            max_len = max([len(str(col))] + [len(str(val)) for val in df[col]])
            worksheet.set_column(idx, idx, max_len + 2)

            if "Pourcent" in col or "Marge" in col:
                worksheet.set_column(idx, idx, max_len + 2, format_pct)

    buffer.seek(0)
    return buffer

def analyser_et_log(df, nom_feuille):
    """
    Analyse le contenu d'une feuille exportée :
    - Compte les lignes utiles
    - Vérifie les valeurs manquantes
    - Détecte les marges faibles ou résultats négatifs
    - Retourne un log au format texte
    """
    lignes_utiles = len(df.dropna(how="all"))
    log = f"{nom_feuille} : {lignes_utiles} lignes utiles exportées.\n"

    # Vérification des données manquantes
    if df.isnull().any().any():
        log += f"  ⚠ Contient des valeurs manquantes.\n"

    # Score de complétude des données
    pct_nan = df.isnull().sum().sum() / df.size
    log += f"  🔎 Score de complétude : {round((1 - pct_nan) * 100)}%\n"

    # Alertes métier spécifiques
    if "Résultat" in df.columns and (df["Résultat"] < 0).any():
        log += f"  ⚠ Résultats négatifs détectés\n"
    if "Marge" in df.columns and (df["Marge"] < 0.1).any():
        log += f"  📉 Marges inférieures à 10%\n"

    # Vérification des années dans les bornes attendues
    if "Annee" in df.columns:
        if df["Annee"].min() < 2022 or df["Annee"].max() > 2026:
            log += f"  ⚠ Années hors bornes attendues [2022-2026]\n"

    log += "\n"
    return log, lignes_utiles


# -------------------- Interface Streamlit --------------------

# Configuration de la page Streamlit
st.set_page_config(page_title="Extracteur de feuilles Excel EHPAD", layout="centered")

# Titre principal de l'app
st.title("📊 Extracteur de feuilles Excel - EHPAD")
st.markdown(
    "Dépose un fichier `.xlsm`, choisis les feuilles à exporter, "
    "et récupère automatiquement les fichiers formatés pour Qlik Sense."
)

# -------- 1. Upload du fichier --------
uploaded_file = st.file_uploader("📁 Déposez ici votre fichier Excel", type=["xlsm"])

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)  # Lecture sans chargement complet
        st.success("✅ Fichier chargé avec succès !")

        # -------- 2. Détection des feuilles --------
        feuilles_disponibles = xls.sheet_names
        st.markdown("### 🗂️ Feuilles détectées dans le fichier :")
        st.write(feuilles_disponibles)

        # Sélection multiple des feuilles à exporter
        feuilles_cibles = st.multiselect(
            "Sélectionnez les feuilles à exporter :",
            options=["Historique_Global", "Historique_Local", "Historique_Projection", "Export_Qlik"],
            default=["Export_Qlik"]
        )

        if feuilles_cibles:

            # -------- 3. Bouton pour lancer l’export --------
            if st.button("🚀 Lancer l'export"):

                zip_buffer = BytesIO()
                log_txt = f"Export réalisé le {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n\n"
                total_lignes = 0
                dernier_df = None  # Pour afficher un aperçu à la fin

                with zipfile.ZipFile(zip_buffer, "a") as zf:
                    for feuille in feuilles_cibles:
                        try:
                            if feuille not in feuilles_disponibles:
                                log_txt += f"{feuille} : ❌ Feuille non trouvée dans le fichier.\n\n"
                                continue

                            df = xls.parse(feuille)
                            dernier_df = df.copy()

                            buffer = formater_excel(df, "Export")
                            zf.writestr(f"{feuille}.xlsx", buffer.read())

                            log_entry, lignes = analyser_et_log(df, feuille)
                            log_txt += log_entry
                            total_lignes += lignes

                        except Exception as e:
                            log_txt += f"{feuille} : ❌ Erreur lors du traitement ({str(e)})\n\n"

                    # Résumé global
                    log_txt += f"✅ Total lignes exportées : {total_lignes}\n"
                    zf.writestr("log_export.txt", log_txt)

                # -------- 4. Téléchargements & aperçus --------
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

        else:
            st.warning("Veuillez sélectionner au moins une feuille à exporter.")

    except Exception as e:
        st.error(f"❌ Erreur lors de la lecture du fichier : {str(e)}")

# -------- Pied de page --------
st.markdown("---")
st.caption("Développé par Rémy Laguerre – CY Tech – Stage Pilotage Financier – Juillet 2025")
