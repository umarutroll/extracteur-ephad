# Extracteur de feuille Excel de type :  Historique
Application Streamlit permettant d'extraire des feuilles Exel de type *historique* et de générer des fichier .xlsm formatés, accompagné d'un rapport d'export automatique 


## Objectif : Simplifier la récupération et la préparation de données issus de fichier Excel complexe, pour une utilisation dans Qlik Sense ou tout autre outil de visualisation


**Prérequis pour utiliser** : Utilisation local via streamlit ou streamlit cloud



**Période de travail** : 23 juillet 2025 pendant mon stage de fin d'étude se finissant le 15 décembre 2025, dans une société possédant plusieurs Ephad dont j'étais chargé de construire un de pilotage financier. 





##  Fonctionnalités

-  **Téléversement d’un fichier** `.xlsm`
-  **Sélection du type d’export** :
  - Les 3 feuilles d’historique (`Global`, `Local`, `Projection`)
  - Ou une seule feuille `Export_Qlik`
-  **Mise en forme automatique des fichiers Excel générés** :
  - Largeurs de colonnes ajustées
  - Formatage des pourcentages
- **Rapport de synthèse** inclus dans un fichier `log_export.txt`

---

##  Évolutions envisagées

- Détection automatique des feuilles présentes dans le fichier
- Interface plus flexible avec extraction **à la carte**
- Rapport enrichi (type de valeurs manquantes, analyse plus poussée)

---

##  Technologies utilisées

- Python
- Streamlit
- Pandas
- XlsxWriter
- openpyxl



<img width="1126" height="575" alt="image" src="https://github.com/user-attachments/assets/411c894a-daa1-48b1-9648-1580708ac000" />



<img width="1118" height="589" alt="image" src="https://github.com/user-attachments/assets/a9e263a8-132b-418a-9a0d-62c548e5fcc0" />
