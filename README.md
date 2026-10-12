# Pharmacie : tableau de bord interactif

Bonus du Rush 2 (MBA promo 2028, groupe les 3 M) : version web de l'outil `Pharmacie_Analyse.xlsx` pour le
propriétaire et le pharmacien. Pages : activité, environnement, stocks et alertes, scan des ventes, recherche,
calendrier, paramètres.

Lancer en local :

    python3 -m pip install -r requirements.txt
    python3 -m streamlit run app.py

Contenu : `app.py` (l'application), `Pharmacie_Analyse.xlsx` (le classeur lu par l'application), `externe/`
(données publiques : Réseau Sentinelles, Météo-France, RNSA, vacances scolaires, jours fériés).
Les saisies (stock, scans) restent sur le poste et ne sont pas publiées.
