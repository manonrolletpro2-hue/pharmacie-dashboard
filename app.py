# -*- coding: utf-8 -*-
"""
Outil de gestion de la pharmacie (propriétaire et pharmacien).

Lancement :   streamlit run app.py
Dépendances : voir requirements.txt

Source unique : Pharmacie_Analyse.xlsx. L'outil reprend les onglets du classeur, page par page.
Les graphiques reprennent ceux de l'onglet Activité du classeur ;
les autres pages reprennent les tableaux des onglets Stocks, Scan, Médicament,
Recherche, Calendrier et Statistiques.
Saisies (stock, scans, correspondance code-barres) : dossier saisies/.
"""

import calendar
import datetime as dt
import io
import math
import re
import sqlite3
from pathlib import Path

import numpy as np
import pandas as pd
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import streamlit as st
import streamlit.components.v1 as components

# ---------------------------------------------------------------------------
# Paramètres
# ---------------------------------------------------------------------------
ICI = Path(__file__).resolve().parent
CLASSEUR = ICI / "Pharmacie_Analyse.xlsx"
SAISIES = ICI / "saisies"
F_BASE = SAISIES / "pharmacie.db"
DOSSIER_EXTERNE = ICI / "externe"

ATC = ["M01AB", "M01AE", "N02BA", "N02BE", "N05B", "N05C", "R03", "R06"]
JOURS = ["Lundi", "Mardi", "Mercredi", "Jeudi", "Vendredi", "Samedi", "Dimanche"]
JOURS_COURTS = ["Lun.", "Mar.", "Mer.", "Jeu.", "Ven.", "Sam.", "Dim."]
MOIS = ["Janvier", "Février", "Mars", "Avril", "Mai", "Juin", "Juillet", "Août",
        "Septembre", "Octobre", "Novembre", "Décembre"]
MOIS_COURTS = ["Janv.", "Févr.", "Mars", "Avr.", "Mai", "Juin", "Juil.", "Août",
               "Sept.", "Oct.", "Nov.", "Déc."]

# Palette : bleu marine et bleu vif
MARINE = "#1b2a3d"
BLEU = "#2d62b8"
BLEU_CLAIR = "#c4d3ea"
TEXTE2, GRIS = "#4b5a6b", "#8a96a3"
GRILLE, AXE = "#e6ebf1", "#c9d2dc"
SEQ = ["#eaf1fc", "#cfe0f9", "#a9c6f2", "#7aa6ea", "#4d86e2", "#2f6fdf", "#2457b5", "#16324f"]
ROUGE, ORANGE, VERT = "#c0392b", "#d9822b", "#2e8b57"

st.set_page_config(page_title="Pharmacie", layout="wide")

st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap');
    .stApp *:not([data-testid="stIconMaterial"]) {
        font-family: 'Inter', system-ui, -apple-system, 'Segoe UI', sans-serif !important;}
    .stApp {background: #f5f6f8;}
    #MainMenu, footer, [data-testid="stToolbar"], [data-testid="stDecoration"],
    [data-testid="stAppDeployButton"], [data-testid="stMainMenu"] {display: none !important;}
    header[data-testid="stHeader"] {background: transparent;}
    .block-container {padding-top: 2rem; padding-bottom: 2rem; max-width: 1440px;}

    /* Barre latérale */
    section[data-testid="stSidebar"] {background: #111c2d; border-right: 1px solid #0b1320;}
    section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] {
        background: transparent; border-radius: 4px; margin: 1px 8px; padding: 7px 12px;
        border-left: 3px solid transparent;}
    section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"] span {
        color: #aab4c3 !important; font-weight: 500; font-size: .88rem;}
    section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"]:hover {background: rgba(255,255,255,.05);}
    section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"][aria-current="page"] {
        background: rgba(255,255,255,.07); border-left-color: #3d7be0;}
    section[data-testid="stSidebar"] [data-testid="stSidebarNavLink"][aria-current="page"] span {color: #ffffff !important;}
    section[data-testid="stSidebar"] [data-testid="stSidebarCollapseButton"] * {color: #aab4c3;}
    [data-testid="stNavSectionHeader"], [data-testid="stNavSectionHeader"] * {color: #6f7b8c !important;
        font-size: .7rem !important; font-weight: 600 !important; letter-spacing: .08em; text-transform: uppercase;}
    [data-testid="stSidebarHeader"] img, [data-testid="stLogo"] {height: 2.4rem !important; max-width: 100%;}

    /* Cartes */
    div[class*="st-key-carte"], div[class*="st-key-filtres"] {background: #ffffff; border: 1px solid #e2e5ea;
        border-radius: 6px; padding: 14px 18px 8px 18px;}
    /* Cartes côte à côte : même hauteur */
    [data-testid="stColumn"] > [data-testid="stVerticalBlock"] {height: 100%;}
    [data-testid="stColumn"] > [data-testid="stVerticalBlock"] > [data-testid="stLayoutWrapper"]:last-child {flex-grow: 1;}
    [data-testid="stLayoutWrapper"] > div[class*="st-key-carte"] {height: 100%;}
    .titre-carte {font-weight: 600; font-size: .86rem; color: #1b2a3d; margin-bottom: 4px;}
    .kpi {background: #fff; border: 1px solid #e2e5ea; border-radius: 6px; padding: 12px 16px; min-height: 78px;
        overflow: hidden;}
    .kpi .lib {font-size: .76rem; color: #6b7685; font-weight: 500; white-space: nowrap; overflow: hidden;
        text-overflow: ellipsis;}
    .kpi .val {font-size: 1.45rem; font-weight: 600; color: #1b2a3d; line-height: 1.3; margin-top: 6px;
        white-space: nowrap; font-variant-numeric: tabular-nums;}
    .kpi.rouge {border-left: 3px solid #c0392b;} .kpi.orange {border-left: 3px solid #d9822b;}
    .kpi.vert {border-left: 3px solid #2e8b57;} .kpi.gris {border-left: 3px solid #a3acb8;}
    div[data-testid="stForm"] {border: none; padding: 0;}
    table.cal {width: 100%; border-collapse: separate; border-spacing: 4px; table-layout: fixed; margin-bottom: 8px;}
    table.cal th {font-size: .78rem; font-weight: 600; color: #4b5563; padding: 6px 0; background: #f5f6f8;
        border-radius: 3px; text-align: center;}
    table.cal th.we {background: #e6e8ec;}
    table.cal td {height: 64px; border: 1px solid #e2e5ea; border-radius: 3px; padding: 6px 8px; vertical-align: top;}
    table.cal td.vide {border-color: transparent; background: transparent;}
    table.cal td.vide.we {background: #f2f2f2;}
    table.cal .num {font-size: .75rem; color: #6b7685; font-weight: 600;}
    table.cal .v {font-size: 1rem; font-weight: 600; color: #1b2a3d; text-align: right; margin-top: 8px;}
    table.env {width: 100%; border-collapse: collapse; table-layout: fixed; margin: 4px 0 10px 0;}
    table.env th, table.env td {padding: 8px 12px; border-bottom: 1px solid #eef0f3; text-align: left;
        vertical-align: middle; font-size: .85rem;}
    table.env thead th {font-size: .76rem; font-weight: 600; color: #6b7685; background: #f5f6f8;}
    table.env thead th:first-child, table.env tbody th {width: 18%;}
    table.env tbody th {font-weight: 600; color: #1b2a3d;}
    table.env td b {display: block; font-size: .95rem;}
    table.env td span {display: block; font-size: .74rem; color: #6b7685; white-space: nowrap;}
    table.env td.hausse b {color: #C00000;}
    table.env td.baisse b {color: #2E5FA3;}
    table.env tr.total th, table.env tr.total td {background: #f5f6f8; border-bottom: none;}
    .cal-tete {font-size: .74rem; font-weight: 600; color: #6b7685; text-align: center; padding: 4px 0;
        background: #f5f6f8; border-radius: 3px;}
    .cal-tete.we {background: #e6e8ec;}
    .cal-vide {height: 58px;} .cal-vide.we {height: 58px; background: #f2f2f2; border-radius: 3px;}
    div[class*="st-key-cal_j_"] button {height: 58px; border: 1px solid #e2e5ea; border-radius: 3px;
        padding: 2px 6px; line-height: 1.15;}
    div[class*="st-key-cal_j_"] button p {font-size: .82rem; text-align: right; width: 100%;}
    div[class*="st-key-cal_s_"] button {height: 58px; color: #6b7685; border: 1px solid #e2e5ea; padding: 0 2px;}
    div[class*="st-key-cal_s_"] button p {font-size: .72rem; white-space: nowrap;}
    div[class*="st-key-cal_j_"] button {position: relative;}
    div[class*="st-key-cal_j_"] button::before {position: absolute; top: 4px; left: 6px; font-size: .64rem;
        font-weight: 700; color: #1b2a3d; font-style: normal;}
    .cal-legende .trait {display: inline-block; width: 22px; height: 3px; vertical-align: middle; margin: 0 6px 0 0;}
    .cal-legende .trait.vac {background: #1b2a3d;} .cal-legende .trait.ferie {background: #C00000; margin-left: 22px;}
    .cal-legende .leg-prev {font-style: italic; font-weight: 600; color: #4b5563; margin: 0 6px 0 22px;}
    .cal-vac {font-size: .74rem; font-weight: 600; color: #4b5563; margin-bottom: 2px;}
    .cal-env {font-size: .82rem; color: #1b2a3d; line-height: 1.5;}
    .temp-mois {float: right; font-weight: 500; color: #6b7685; font-size: .8rem;}
    span.pic {display: inline-block; font-size: .68rem; font-weight: 600; padding: 1px 6px; border-radius: 3px;
        margin: 1px 3px 1px 0; white-space: nowrap;}
    span.pic {background: #eef0f3; color: #1b2a3d;}
    .ligne-synth span.pic {font-size: .74rem;}
    .cal-legende {font-size: .74rem; color: #6b7685; margin-top: 6px;}
    table.bon {width: 100%; border-collapse: collapse; margin: 4px 0 12px 0;}
    table.bon th, table.bon td {padding: 9px 12px; border-bottom: 1px solid #eef0f3; font-size: .86rem; text-align: right;}
    table.bon th:first-child {text-align: left;}
    table.bon thead th {background: #1b2a3d; color: #fff; font-weight: 600; font-size: .78rem;}
    table.bon tbody th {font-weight: 600; color: #1b2a3d;}
    table.bon td.qte {font-weight: 700; color: #1b2a3d;}
    table.bon tfoot th, table.bon tfoot td {border-top: 2px solid #1b2a3d; border-bottom: none; font-weight: 700;}
    table.bon td:last-child, table.bon th:last-child {text-align: center;}
    table.bon tr.l-rupture th, table.bon tr.l-rupture td {background: #F8CBAD;}
    table.bon tr.l-commande th, table.bon tr.l-commande td {background: #FCE4D6;}
    span.alerte {display: inline-block; min-width: 110px; padding: 3px 8px; border-radius: 3px; font-size: .74rem;
        font-weight: 700; text-align: center;}
    span.alerte.rupture {background: #C00000; color: #fff;} span.alerte.commande {background: #FF6600; color: #fff;}
    span.alerte.ok {background: #E2EFDA; color: #375623;} span.alerte.saisir {background: #FFF2CC; color: #7F6000;}
    .source {font-size: .7rem; color: #8a94a3; margin-top: 6px; line-height: 1.35;}
    table.env td.hausse {color: #C00000; font-weight: 600;} table.env td.baisse {color: #2E5FA3; font-weight: 600;}
    table.env.tous td span {display: inline; font-weight: 500; color: #6b7685;}
    .ligne-synth {padding: 10px 0; border-bottom: 1px solid #eef0f3; font-size: .78rem; color: #6b7685;}
    .ligne-synth b {display: block; font-size: 1.15rem; font-weight: 600; color: #1b2a3d; margin-top: 2px;}
    [data-testid="stWidgetLabel"] p {font-weight: 500; color: #4b5563; font-size: .8rem;}
    </style>
    """,
    unsafe_allow_html=True,
)


# ---------------------------------------------------------------------------
# Lecture du classeur
# ---------------------------------------------------------------------------
def _ligne(feuille, valeur, colonne=0):
    serie = feuille.iloc[:, colonne].astype(str).str.strip()
    trouve = serie[serie == valeur]
    if trouve.empty:
        raise ValueError("Repère %s introuvable dans le classeur" % valeur)
    return trouve.index[0]


def _bloc(feuille, debut, nb_colonnes, col0=0):
    """Tableau à partir de la ligne d'en-tête `debut`, jusqu'à la première ligne vide."""
    zone = feuille.iloc[:, col0:col0 + nb_colonnes]
    entete = [str(x) for x in zone.iloc[debut]]
    lignes = []
    for i in range(debut + 1, len(zone)):
        ligne = zone.iloc[i]
        if pd.isna(ligne.iloc[0]):
            break
        lignes.append(list(ligne))
    return pd.DataFrame(lignes, columns=entete)


@st.cache_data(show_spinner="Chargement des données")
def charger(chemin, horodatage):
    xls = pd.ExcelFile(chemin, engine="openpyxl")

    hist = pd.read_excel(xls, "Historique", usecols="A:Y")
    hist["Date"] = pd.to_datetime(hist["Date"])
    hist["Total"] = hist[ATC].sum(axis=1)
    jours = hist[hist["Jour entier"] == "Oui"].copy()
    jours["Jour"] = pd.Categorical(jours["Jour"], categories=JOURS, ordered=True)

    brut = pd.read_excel(xls, "Calculs", header=None)
    heures = _bloc(brut, _ligne(brut, "Heure"), 10)
    heures = heures[heures["Heure"].astype(str).str.match(r"^\d{2}:\d{2}$")]
    heures[ATC + ["Total"]] = heures[ATC + ["Total"]].astype(float)

    # Onglet Environnement : ventes par semaine (colonnes A à L) et grippe en France (colonne masquée U),
    # quel que soit le facteur choisi dans le classeur
    env = pd.read_excel(xls, "Environnement", header=None)
    l_sem = _ligne(env, "Semaine")
    hebdo = _bloc(env, l_sem, 12)
    hebdo = hebdo[["Lundi"] + ATC + ["Total"]].copy()
    hebdo["Grippe"] = pd.to_numeric(env.iloc[l_sem + 1:l_sem + 1 + len(hebdo), 20], errors="coerce").values
    hebdo["Lundi"] = pd.to_datetime(hebdo["Lundi"])
    for c in ATC + ["Total"]:
        hebdo[c] = hebdo[c].astype(float)
    annee_moy = _bloc(env, l_sem, 4, col0=15)
    annee_moy.columns = ["Semaine", "Repère", "Facteur moyen", "Ventes moyennes"]

    # Données externes complémentaires (dossier externe/)
    externe = {}
    for nom in ["pollens_hebdo", "meteo_hebdo", "epidemies_hebdo", "bronchiolite_saison", "villes"]:
        f = DOSSIER_EXTERNE / (nom + ".csv")
        if f.exists():
            externe[nom] = pd.read_csv(f, sep=";", parse_dates=["Lundi"] if nom.endswith("hebdo") else None,
                                       dtype={"Département": str})
    for nom in ["jours_feries", "vacances_scolaires"]:
        f = DOSSIER_EXTERNE / (nom + ".csv")
        if f.exists():
            externe[nom] = pd.read_csv(f, sep=";", parse_dates=["Date"] if nom == "jours_feries" else ["Début", "Fin"])

    return {"hist": hist, "jours": jours, "heures": heures, "annee_moy": annee_moy, "hebdo": hebdo,
            "externe": externe}


# ---------------------------------------------------------------------------
# Base de données de la pharmacie (SQLite, fichier saisies/pharmacie.db)
#   produits     : code CIP13 -> groupe ATC (comme la table de correspondance de l'onglet Scan)
#   inventaires  : stock compté en rayon et date du comptage, par groupe
#   mouvements   : chaque scan ou saisie, en entrée (réception) ou en sortie (vente)
# Stock en rayon = dernier inventaire + entrées - sorties enregistrées depuis cet inventaire.
# ---------------------------------------------------------------------------
NON_REFERENCE = "Non référencé"


def base():
    SAISIES.mkdir(exist_ok=True)
    con = sqlite3.connect(str(F_BASE))
    con.executescript("""
        CREATE TABLE IF NOT EXISTS produits (cip13 TEXT PRIMARY KEY, libelle TEXT, groupe TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS inventaires (groupe TEXT PRIMARY KEY, stock REAL NOT NULL, date TEXT NOT NULL);
        CREATE TABLE IF NOT EXISTS parametres (cle TEXT PRIMARY KEY, valeur TEXT);
        CREATE TABLE IF NOT EXISTS mouvements (id INTEGER PRIMARY KEY AUTOINCREMENT, date TEXT NOT NULL,
            cip13 TEXT, groupe TEXT NOT NULL, sens TEXT NOT NULL, quantite REAL NOT NULL);
    """)
    return con


def requete(sql, params=()):
    with base() as con:
        return pd.read_sql_query(sql, con, params=params)


def executer(sql, params=()):
    with base() as con:
        con.execute(sql, params)


PARAMETRES_DEFAUT = {"ville": "Montpellier", "nom": "Pharmacie", "titulaire": "", "adresse": "", "telephone": "",
                     "email": ""}


def parametre(cle):
    r = requete("SELECT valeur FROM parametres WHERE cle = ?", (cle,))
    return PARAMETRES_DEFAUT.get(cle, "") if r.empty else r.iloc[0, 0]


def enregistrer_parametre(cle, valeur):
    executer("INSERT OR REPLACE INTO parametres (cle, valeur) VALUES (?, ?)", (cle, str(valeur)))


def maintenant():
    return dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


def code_cip13(code):
    """Un Data Matrix GS1 commence par 01 suivi du GTIN-14, dont le CIP13 forme la fin (comme l'onglet Scan)."""
    code = str(code).strip().replace(" ", "")
    if len(code) >= 16 and code.startswith("01"):
        return code[3:16]
    return code


def produit_du_code(cip):
    r = requete("SELECT libelle, groupe FROM produits WHERE cip13 = ?", (cip,))
    return None if r.empty else (r.iloc[0, 0], r.iloc[0, 1])


def enregistrer_mouvement(groupe, sens, quantite, cip=None):
    executer("INSERT INTO mouvements (date, cip13, groupe, sens, quantite) VALUES (?, ?, ?, ?, ?)",
             (maintenant(), cip, groupe, sens, float(quantite)))


def stock_en_rayon():
    """Stock de départ (dernier inventaire), entrées et sorties depuis, par groupe."""
    inv = requete("SELECT groupe, stock, date FROM inventaires").set_index("groupe").reindex(ATC)
    mv = requete("SELECT m.groupe, m.sens, m.quantite FROM mouvements m "
                 "LEFT JOIN inventaires i ON i.groupe = m.groupe WHERE i.date IS NULL OR m.date >= i.date")
    entrees = mv[mv["sens"] == "Entrée"].groupby("groupe")["quantite"].sum().reindex(ATC, fill_value=0)
    sorties = mv[mv["sens"] == "Sortie"].groupby("groupe")["quantite"].sum().reindex(ATC, fill_value=0)
    t = pd.DataFrame({"Stock de départ": pd.to_numeric(inv["stock"]).astype(float), "Entrées": entrees.astype(float),
                      "Sorties": sorties.astype(float)}, index=ATC)
    connu = t["Stock de départ"].notna() | (t["Entrées"] > 0) | (t["Sorties"] > 0)
    t["Stock en rayon"] = np.where(connu, t["Stock de départ"].fillna(0) + t["Entrées"] - t["Sorties"], np.nan)
    t["Inventaire du"] = pd.to_datetime(inv["date"]).dt.strftime("%d/%m/%Y")
    return t


def tendance_du_mois(mois_num):
    """
    Ventes par jour du mois, année par année (mois complets), et tendance : droite ajustée sur les années
    étudiées, prolongée d'un an. Même calcul que l'onglet Stocks du classeur (PREVISION).
    """
    j = JOURS_DATA[(JOURS_DATA["Date"].dt.month == mois_num) & (JOURS_DATA["Mois entier"] == "Oui")]
    par_an = j.groupby(j["Date"].dt.year)[ATC].mean()
    annees = par_an.index.values.astype(float)
    cible = annees.max() + 1
    sortie = pd.DataFrame(index=ATC)
    for g in ATC:
        if len(annees) >= 3:
            pente, origine = np.polyfit(annees, par_an[g].values, 1)
            sortie.loc[g, "Prévision par jour"] = max(0.0, pente * cible + origine)
            sortie.loc[g, "Tendance"] = pente / par_an[g].mean() if par_an[g].mean() else np.nan
        else:
            sortie.loc[g, "Prévision par jour"] = par_an[g].mean()
            sortie.loc[g, "Tendance"] = np.nan
    return sortie, par_an, int(cible)


def etat_stock(mois_num):
    """Reprend les formules de l'onglet Stocks du classeur, avec la tendance du mois sur les années étudiées."""
    t = stock_en_rayon()
    nb_jours = calendar.monthrange(2019, mois_num)[1]
    t.index.name = "Code ATC"
    tendance, _, _ = tendance_du_mois(mois_num)
    t["Ventes par jour"] = tendance["Prévision par jour"]
    t["Tendance"] = tendance["Tendance"]
    t["Autonomie (jours)"] = t["Stock en rayon"] / t["Ventes par jour"].replace(0, np.nan)
    t["Stock minimum conseillé"] = (t["Ventes par jour"] * nb_jours).map(lambda v: math.ceil(v) if pd.notna(v) else np.nan)
    t["Quantité à commander"] = (t["Stock minimum conseillé"] - t["Stock en rayon"]).clip(lower=0).map(
        lambda v: math.ceil(v) if pd.notna(v) else np.nan)
    t["Alerte"] = np.where(t["Stock en rayon"].isna(), "À saisir",
                  np.where(t["Stock en rayon"] <= 0, "RUPTURE",
                  np.where(t["Quantité à commander"] > 0, "À COMMANDER", "OK")))
    return t, nb_jours


# ---------------------------------------------------------------------------
# Mise en forme
# ---------------------------------------------------------------------------
def nb(x, d=0):
    if x is None or pd.isna(x):
        return ""
    return ("{:,.%df}" % d).format(x).replace(",", " ").replace(".", ",")


def pct(x, d=0):
    return "" if x is None or pd.isna(x) else nb(100 * x, d) + " %"


def kpi(colonne, icone, valeur, libelle, ton=""):
    """Indicateur chiffré (l'icône est conservée dans les appels mais n'est plus affichée)."""
    colonne.markdown('<div class="kpi %s"><div class="lib">%s</div><div class="val">%s</div></div>'
                     % (ton, libelle, valeur), unsafe_allow_html=True)


def carte(cle, titre):
    """Carte blanche arrondie avec un titre centré ; renvoie le conteneur."""
    c = st.container(key="carte_" + cle)
    c.markdown('<div class="titre-carte">%s</div>' % titre, unsafe_allow_html=True)
    return c


def graphique(conteneur, fig, hauteur=300, legende=False, marge_haut=10):
    fig.update_layout(
        height=hauteur, template="plotly_white",
        font=dict(family="Inter, system-ui, -apple-system, Segoe UI, sans-serif", color=TEXTE2, size=12),
        margin=dict(l=6, r=6, t=marge_haut, b=6), paper_bgcolor="rgba(0,0,0,0)", plot_bgcolor="rgba(0,0,0,0)",
        hoverlabel=dict(bgcolor="white", font_size=12), showlegend=legende, bargap=0.35,
        legend=dict(orientation="h", yanchor="top", y=-0.15, xanchor="center", x=0.5, title=None),
    )
    fig.update_xaxes(showgrid=False, linecolor=AXE, tickfont=dict(color=GRIS))
    fig.update_yaxes(gridcolor=GRILLE, zeroline=False, tickfont=dict(color=GRIS))
    conteneur.plotly_chart(fig, theme=None, config={"displayModeBar": False})


def barres(x, y, couleur=BLEU, d=1, horizontal=False, surligne=None, etiquettes=True):
    couleurs = couleur if surligne is None else [BLEU if s else BLEU_CLAIR for s in surligne]
    if horizontal:
        fig = go.Figure(go.Bar(x=y, y=x, orientation="h", marker_color=couleurs,
                               text=[nb(v, d) for v in y], textposition="outside",
                               textfont=dict(color=TEXTE2, size=11),
                               hovertemplate="%{y} : %{x:,.1f}<extra></extra>"))
        fig.update_xaxes(range=[0, max(y) * 1.32], showticklabels=False, showgrid=False)
        fig.update_yaxes(autorange="reversed", showgrid=False)
    else:
        fig = go.Figure(go.Bar(x=x, y=y, marker_color=couleurs,
                               text=[nb(v, d) for v in y] if etiquettes else None, textposition="outside",
                               textfont=dict(color=TEXTE2, size=10), cliponaxis=False,
                               hovertemplate="%{x} : %{y:,.1f}<extra></extra>"))
        fig.update_yaxes(range=[0, np.nanmax(y) * 1.18] if len(y) else None)
        fig.update_xaxes(tickangle=0, tickfont=dict(size=10))
    return fig


def telecharger(df, nom, cle=None):
    tampon = io.StringIO()
    df.to_csv(tampon, sep=";", decimal=",")
    st.download_button("Exporter en CSV", tampon.getvalue().encode("utf-8-sig"), file_name=nom,
                       mime="text/csv", key=cle)


def texte(df, formats):
    """Copie affichable : nombres au format français, case vide au lieu d'une valeur manquante."""
    out = df.copy()
    for c, d in formats.items():
        out[c] = out[c].map(lambda v, d=d: nb(v, d))
    return out


# Couleurs reprises des mises en forme conditionnelles du classeur
STYLE_ALERTE = {
    "RUPTURE": "background-color:#C00000; color:#FFFFFF; font-weight:700",
    "À COMMANDER": "background-color:#FF6600; color:#FFFFFF; font-weight:700",
    "OK": "background-color:#E2EFDA; color:#375623; font-weight:700",
    "À saisir": "background-color:#FFF2CC; color:#7F6000; font-weight:700",
}
FOND_LIGNE = {"RUPTURE": "background-color:#F8CBAD", "À COMMANDER": "background-color:#FCE4D6"}


def style_stock(df):
    """Ligne colorée selon l'alerte (onglet Stocks) et case Alerte pleine."""
    def ligne(r):
        fond = FOND_LIGNE.get(r["Alerte"], "")
        return [STYLE_ALERTE.get(r["Alerte"], "") if c == "Alerte" else fond for c in r.index]
    return df.style.apply(ligne, axis=1)


def couleur_echelle(v, vmin, vmax):
    """Couleur entre blanc et #8EA9DB, comme l'échelle de couleur du classeur."""
    if pd.isna(v) or vmax == vmin:
        return "#FFFFFF"
    t = (v - vmin) / (vmax - vmin)
    r, g, b = [round(255 + (c - 255) * t) for c in (0x8E, 0xA9, 0xDB)]
    return "#%02X%02X%02X" % (r, g, b)


# ---------------------------------------------------------------------------
# Données
# ---------------------------------------------------------------------------
if not CLASSEUR.exists():
    st.error("Classeur introuvable : %s" % CLASSEUR)
    st.stop()
D = charger(str(CLASSEUR), CLASSEUR.stat().st_mtime + sum(
    f.stat().st_mtime for f in DOSSIER_EXTERNE.glob("*.csv")))
JOURS_DATA, HIST, HEURES = D["jours"], D["hist"], D["heures"]
VILLES = D["externe"].get("villes", pd.DataFrame(columns=["Ville"]))
VILLE = parametre("ville") if parametre("ville") in set(VILLES["Ville"]) else PARAMETRES_DEFAUT["ville"]
INFO_VILLE = VILLES[VILLES["Ville"] == VILLE].iloc[0] if (VILLES["Ville"] == VILLE).any() else None
REGION = INFO_VILLE["Région"] if INFO_VILLE is not None else "France"
ZONE = INFO_VILLE["Zone scolaire"] if INFO_VILLE is not None and "Zone scolaire" in INFO_VILLE else None
ANNEES = sorted(JOURS_DATA["Date"].dt.year.unique().tolist())
MOIS_PROCHAIN = (pd.Period(JOURS_DATA.loc[JOURS_DATA["Mois entier"] == "Oui", "Période"].max(), "M") + 1).month



# ===========================================================================
# Activité (onglet Activite)
# ===========================================================================
def page_activite():
    filtres = st.container(key="filtres")
    f = filtres.columns(3)
    an = f[0].selectbox("Année", ["Toutes"] + ANNEES)
    mois = f[1].selectbox("Mois", ["Tous"] + MOIS)
    med = f[2].selectbox("Médicament", ["Tous"] + ATC)
    col = "Total" if med == "Tous" else med
    mn = None if mois == "Tous" else MOIS.index(mois) + 1

    j = JOURS_DATA
    if an != "Toutes":
        j = j[j["Date"].dt.year == an]
    if mn:
        j = j[j["Date"].dt.month == mn]
    st.write("")
    if j.empty:
        st.warning("Aucune journée pour cette sélection.")
        return

    # Mêmes règles que l'onglet Activité : la saisonnalité suit l'année et le médicament,
    # l'année par année suit le mois et le médicament, le reste suit les trois filtres.
    j_an = JOURS_DATA if an == "Toutes" else JOURS_DATA[JOURS_DATA["Date"].dt.year == an]
    saison = j_an.groupby(j_an["Date"].dt.month)[col].mean()
    j_mois = JOURS_DATA if not mn else JOURS_DATA[JOURS_DATA["Date"].dt.month == mn]
    par_an = j_mois.groupby(j_mois["Date"].dt.year)[col].mean()
    par_jour = j.groupby("Jour", observed=True)[col].mean()
    h = HEURES.set_index("Heure")[col]
    h = h[h > 0]
    k = st.columns(5)
    kpi(k[0], "", nb(j[col].sum()), "Total vendu")
    kpi(k[1], "", nb(j[col].mean(), 1), "Ventes par jour")
    if mn is None:
        kpi(k[2], "", MOIS_COURTS[saison.idxmax() - 1], "Mois le plus fort")
    else:
        kpi(k[2], "", str(par_an.idxmax()), "Meilleure année")
    kpi(k[3], "", JOURS_COURTS[JOURS.index(par_jour.idxmax())], "Jour le plus fort")
    kpi(k[4], "", h.idxmax()[:2] + " h", "Heure de pointe")
    st.write("")

    o1, o2 = st.tabs(["Évolution", "Quand vend-on ?"])
    with o1:
        c1, c2 = st.columns([2, 1])
        with c1:
            cont = carte("evolution", "Ventes par jour, mois par mois")
            m = JOURS_DATA[JOURS_DATA["Mois entier"] == "Oui"].groupby("Période")[col].mean()
            x = pd.PeriodIndex(m.index, freq="M").to_timestamp()
            fig = go.Figure(go.Scatter(x=x, y=m.values, mode="lines", line=dict(color=BLEU, width=2),
                                       fill="tozeroy", fillcolor="rgba(45,98,184,.06)",
                                       hovertemplate="%{x|%m/%Y} : %{y:,.1f} par jour<extra></extra>"))
            fig.update_yaxes(rangemode="tozero")
            graphique(cont, fig, 330)
        with c2:
            cont = carte("repartition", "Total vendu par médicament")
            rep = j[ATC].sum().sort_values(ascending=False)
            graphique(cont, barres(list(rep.index), list(rep.values), d=0, horizontal=True,
                                   surligne=[med in ("Tous", g) for g in rep.index]), 330)
    with o2:
        c1, c2, c3 = st.columns(3)
        with c1:
            cont = carte("saison", "Selon le mois")
            fig = barres(MOIS, list(saison.reindex(range(1, 13)).values), d=0,
                         surligne=[mn in (None, i) for i in range(1, 13)])
            fig.update_xaxes(tickvals=MOIS, ticktext=[m[0] for m in MOIS])
            graphique(cont, fig, 300)
        with c2:
            cont = carte("semaine", "Selon le jour de la semaine")
            graphique(cont, barres(JOURS_COURTS, list(par_jour.reindex(JOURS).values)), 300)
        with c3:
            cont = carte("heure", "Selon l'heure")
            lib = [x[:2] + " h" for x in h.index]
            fig = barres(lib, list(h.values), d=0, etiquettes=False)
            fig.update_xaxes(tickvals=lib, ticktext=[str(int(x[:2])) for x in h.index])
            graphique(cont, fig, 300)


# ===========================================================================
# Environnement (onglet Environnement, complété par la gastro, la varicelle, la température et les pollens)
# ===========================================================================
FACTEURS = {
    # nom : (colonne, libellé si les ventes montent, libellé si elles baissent)
    "Grippe": ("Grippe", "Montent avec la grippe", "Baissent pendant la grippe"),
    "Gastro": ("Gastro", "Montent avec la gastro", "Baissent pendant la gastro"),
    "Varicelle": ("Varicelle", "Montent avec la varicelle", "Baissent pendant la varicelle"),
    "Température": ("Température", "Montent avec la chaleur", "Montent avec le froid"),
    # graminées : ce sont elles qui suivent les allergies saisonnières (le cyprès domine le total dans le Sud)
    "Pollens": ("Graminées", "Montent avec les pollens", "Baissent avec les pollens"),
}


def periode_semaine(libelle):
    """« Semaine 6 » -> « Début février » : période de l'année correspondant au numéro de semaine."""
    try:
        n = int(str(libelle).split()[-1])
        lundi = dt.date.fromisocalendar(2019, min(n, 52), 1) + dt.timedelta(days=3)  # milieu de semaine
    except (ValueError, IndexError):
        return str(libelle)
    moment = "Début" if lundi.day <= 10 else ("Mi-" if lundi.day <= 20 else "Fin")
    return "%s%s%s" % (moment, "" if moment == "Mi-" else " ", MOIS[lundi.month - 1].lower())


def semaines_environnement(ville=None):
    """Ventes hebdomadaires du classeur, avec les épidémies de la région, la météo et les pollens de la ville."""
    ville = ville or VILLE
    ext = D["externe"]
    h = D["hebdo"].copy()  # contient la grippe en France (onglet Environnement), remplacée si possible
    info = VILLES[VILLES["Ville"] == ville]
    if "epidemies_hebdo" in ext and not info.empty:
        e = ext["epidemies_hebdo"]
        e = e[e["Région"] == int(info["Code région"].iloc[0])].drop(columns="Région")
        h = h.drop(columns="Grippe").merge(e, on="Lundi", how="left")
    for nom in ["meteo_hebdo", "pollens_hebdo"]:
        if nom in ext:
            h = h.merge(ext[nom][ext[nom]["Ville"] == ville].drop(columns="Ville"), on="Lundi", how="left")
    h["N° semaine"] = h["Lundi"].dt.isocalendar().week.astype(int).clip(upper=52)
    return h


def lien(x, y, monte, baisse):
    """Même calcul que l'onglet Environnement : coefficient de détermination (RSQ) sur les semaines communes."""
    ok = x.notna() & y.notna()
    if ok.sum() < 10:
        return "", np.nan
    r = np.corrcoef(x[ok], y[ok])[0, 1]
    if abs(r) < 0.2:  # même règle que l'onglet Environnement du classeur
        return "Pas de lien", r ** 2
    return (monte if r > 0 else baisse), r ** 2


def deux_courbes(cont, h, colonne, titre_haut, couleur, groupe, hauteur=420):
    """Une année moyenne : le facteur en haut, les ventes de la semaine en bas (mêmes semaines)."""
    am = h.groupby("N° semaine")[[colonne, groupe]].mean().reset_index()
    fig = make_subplots(rows=2, cols=1, shared_xaxes=True, vertical_spacing=0.16,
                        subplot_titles=(titre_haut, "Ventes moyennes de la semaine, %s" % (
                            "tous médicaments" if groupe == "Total" else groupe)))
    fig.add_scatter(x=am["N° semaine"], y=am[colonne], mode="lines", line=dict(color=couleur, width=2),
                    row=1, col=1, hovertemplate="Semaine %{x} : %{y:,.1f}<extra></extra>")
    fig.add_scatter(x=am["N° semaine"], y=am[groupe], mode="lines", line=dict(color=BLEU, width=2),
                    fill="tozeroy", fillcolor="rgba(45,98,184,.06)", row=2, col=1,
                    hovertemplate="Semaine %{x} : %{y:,.1f}<extra></extra>")
    reperes = D["annee_moy"][D["annee_moy"]["Repère"].notna()]
    fig.update_xaxes(tickvals=reperes["Semaine"], ticktext=reperes["Repère"], row=2, col=1)
    fig.update_yaxes(rangemode="tozero", row=2, col=1)
    fig.update_annotations(font=dict(size=12, color=TEXTE2), x=0, xanchor="left", yanchor="bottom", yshift=6)
    graphique(cont, fig, hauteur, marge_haut=34)


def sources_facteurs():
    """Source de chaque donnée externe, citée sous les tableaux et les graphiques."""
    region = REGION
    return {
        "Grippe": "Source : Réseau Sentinelles (Inserm, Sorbonne Université), syndromes grippaux, incidence "
                  "hebdomadaire pour 100 000 habitants, région %s, 2014-2019." % region,
        "Gastro": "Source : Réseau Sentinelles (Inserm, Sorbonne Université), diarrhées aiguës, incidence "
                  "hebdomadaire pour 100 000 habitants, région %s, 2014-2019." % region,
        "Varicelle": "Source : Réseau Sentinelles (Inserm, Sorbonne Université), varicelle, incidence hebdomadaire "
                     "pour 100 000 habitants, région %s, 2014-2019." % region,
        "Température": "Source : Météo-France, données climatologiques de base quotidiennes, stations à moins de "
                       "25 km de %s, 2014-2019." % VILLE,
        "Pollens": "Source : RNSA (Réseau national de surveillance aérobiologique), pollens de graminées, "
                   "capteur de %s, 2014-2019." % VILLE,
        "Bronchiolite": "Source : Santé publique France (Odissé, réseau OSCOUR), passages aux urgences pour "
                        "bronchiolite des moins de 1 an, région %s, depuis fin 2019." % region,
    }


def environnement_tous(h, disponibles, noms):
    """Vue d'ensemble : effet de chaque facteur sur chaque médicament, puis moment du pic moyen de chaque facteur."""
    cont = carte("lien_tous", "Effet de chaque facteur sur les ventes")
    tete = "".join("<th>%s</th>" % noms[f] for f in disponibles)
    corps = ""
    for m in ATC + ["Total"]:
        cases = ""
        for f in disponibles:
            colonne, monte, baisse = FACTEURS[f]
            txt, r2 = lien(h[colonne], h[m], monte, baisse)
            effet = "" if txt in ("Pas de lien", "") else ("baisse" if txt.startswith("Baissent") else "hausse")
            libelle = {"Montent avec le froid": "Hausse au froid",
                       "Montent avec la chaleur": "Hausse au chaud"}.get(txt, effet.capitalize())
            cases += '<td class="%s">%s</td>' % (effet, "%s <span>%s</span>" % (libelle, pct(r2, 0)) if effet else "")
        corps += '<tr class="%s"><th>%s</th>%s</tr>' % ("total" if m == "Total" else "",
                                                       "Toutes ventes" if m == "Total" else m, cases)
    sources = sources_facteurs()
    cont.markdown('<table class="env tous"><thead><tr><th>Médicament</th>%s</tr></thead><tbody>%s</tbody></table>'
                  '<div class="source">Hausse ou baisse : sens dans lequel les ventes évoluent avec le facteur ; à côté, '
                  'part des hausses et baisses des ventes qui vont de pair avec lui. Case vide : pas de lien.</div>'
                  '<div class="source">%s</div>' % (tete, corps, " ".join(sources[f] for f in disponibles)),
                  unsafe_allow_html=True)
    st.write("")
    cont = carte("pics_tous", "Pic moyen de chaque facteur")
    colonnes = cont.columns(len(disponibles) + 1)
    for c, f in zip(colonnes, disponibles):
        n = h.groupby("N° semaine")[FACTEURS[f][0]].mean().idxmax()
        c.markdown('<div class="ligne-synth">%s<b>%s</b></div>' % ("Chaleur" if f == "Température" else noms[f],
                                                                     periode_semaine("Semaine %d" % n)),
                   unsafe_allow_html=True)
    n = h.groupby("N° semaine")["Total"].mean().idxmax()
    colonnes[-1].markdown('<div class="ligne-synth">Ventes (tous médicaments)<b>%s</b></div>'
                          % periode_semaine("Semaine %d" % n), unsafe_allow_html=True)


def page_environnement():
    h = semaines_environnement()
    noms = {"Grippe": "Grippe", "Gastro": "Gastro-entérite", "Varicelle": "Varicelle", "Température": "Température",
            "Pollens": "Pollens"}
    disponibles = [f for f in noms if FACTEURS[f][0] in h]
    filtres = st.container(key="filtres")
    c1, c2, _ = filtres.columns([1, 1, 2])
    facteur = c1.selectbox("Facteur", ["Tous"] + disponibles, format_func=lambda f: noms.get(f, f), key="env_facteur")
    g = c2.selectbox("Médicament", ["Total"] + ATC, key="env_med", format_func=lambda x: "Tous" if x == "Total" else x,
                     disabled=facteur == "Tous")
    st.write("")
    if facteur == "Tous":
        environnement_tous(h, disponibles, noms)
        return
    colonne, monte, baisse = FACTEURS[facteur]
    source = sources_facteurs()[facteur]

    c1, c2 = st.columns([2, 3])
    with c1:
        cont = carte("lien_env", "Effet sur les ventes")
        corps = ""
        for m in ATC + ["Total"]:
            txt, r2 = lien(h[colonne], h[m], monte, baisse)
            effet = "" if txt in ("Pas de lien", "") else ("baisse" if txt.startswith("Baissent") else "hausse")
            libelle = {"Montent avec le froid": "Hausse par temps froid",
                       "Montent avec la chaleur": "Hausse par temps chaud"}.get(txt, effet.capitalize())
            corps += '<tr class="%s"><th>%s</th><td class="%s">%s</td><td class="%s"><b>%s</b></td></tr>' % (
                "total" if m == "Total" else "", "Toutes ventes" if m == "Total" else m, effet,
                libelle, effet, pct(r2, 1) if effet else "")
        cont.markdown('<table class="env"><thead><tr><th>Médicament</th><th>Ventes</th><th>Part des variations</th>'
                      '</tr></thead><tbody>%s</tbody></table><div class="source">%s</div>' % (corps, source),
                      unsafe_allow_html=True)
    with c2:
        titres = {"Grippe": "Grippe en %s" % REGION, "Gastro": "Gastro-entérite en %s" % REGION,
                  "Varicelle": "Varicelle en %s" % REGION, "Température": "Température à %s (°C)" % VILLE,
                  "Pollens": "Pollens de graminées à %s" % VILLE}
        couleurs_ = {"Grippe": ROUGE, "Gastro": ORANGE, "Varicelle": "#8e44ad", "Température": ORANGE, "Pollens": VERT}
        cont = carte("annee_moy", "Une année moyenne, semaine par semaine")
        deux_courbes(cont, h, colonne, titres[facteur], couleurs_[facteur], g, hauteur=400)
        if facteur != "Température":
            p1, p2, _ = cont.columns([1, 1, 1])
            pic = h.groupby("N° semaine")[colonne].mean().idxmax()
            p1.markdown('<div class="ligne-synth">Pic %s<b>%s</b></div>' % (
                "des pollens" if facteur == "Pollens" else "de l'épidémie", periode_semaine("Semaine %d" % pic)),
                unsafe_allow_html=True)
            pic_v = h.groupby("N° semaine")[g].mean().idxmax()
            p2.markdown('<div class="ligne-synth">Pic des ventes<b>%s</b></div>'
                        % periode_semaine("Semaine %d" % pic_v), unsafe_allow_html=True)
        cont.markdown('<div class="source">%s Ventes : historique de la pharmacie.</div>' % source,
                      unsafe_allow_html=True)


def bon_de_commande_xlsx(commande, mois):
    """Bon de commande au format Excel, mis en forme, prêt à envoyer au grossiste."""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    maintenant_ = dt.datetime.now()
    wb = Workbook()
    ws = wb.active
    ws.title = "Bon de commande"
    marine, gris = "1B2A3D", "F5F6F8"
    fin = Side(style="thin", color="D5D9E0")
    ws.column_dimensions["A"].width = 16
    for c in "BCD":
        ws.column_dimensions[c].width = 22

    ws.merge_cells("A1:D1")
    ws["A1"] = "BON DE COMMANDE"
    ws["A1"].font = Font(size=18, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill("solid", fgColor=marine)
    ws["A1"].alignment = Alignment(horizontal="left", vertical="center", indent=1)
    ws.row_dimensions[1].height = 34
    adresse = ", ".join(x for x in [parametre("adresse"), VILLE] if x)
    infos = [("Pharmacie", parametre("nom")), ("Adresse", adresse)]
    infos += [(k_, v) for k_, v in [("Pharmacien titulaire", parametre("titulaire")),
                                    ("Téléphone", parametre("telephone")), ("E-mail", parametre("email"))] if v]
    infos += [("N° de commande", "BC-%s" % maintenant_.strftime("%Y%m%d-%H%M")),
              ("Date", maintenant_.strftime("%d/%m/%Y")),
              ("Mois préparé", "%s (%d jours)" % (mois, calendar.monthrange(2019, MOIS.index(mois) + 1)[1]))]
    for i, (k_, v) in enumerate(infos, start=3):
        ws.cell(i, 1, k_).font = Font(bold=True, color="6B7685")
        ws.cell(i, 2, v).font = Font(bold=True, color=marine)

    entete = ["Code ATC", "Stock en rayon", "Stock minimum conseillé", "Quantité à commander"]
    ligne0 = len(infos) + 4
    for j, titre in enumerate(entete, start=1):
        c = ws.cell(ligne0, j, titre)
        c.font = Font(bold=True, color="FFFFFF")
        c.fill = PatternFill("solid", fgColor=marine)
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        c.border = Border(top=fin, bottom=fin, left=fin, right=fin)
    ws.row_dimensions[ligne0].height = 30
    for i, (g, r) in enumerate(commande.iterrows(), start=ligne0 + 1):
        valeurs = [g, r["Stock en rayon"], r["Stock minimum conseillé"], r["Quantité à commander"]]
        for j, v in enumerate(valeurs, start=1):
            c = ws.cell(i, j, v if j == 1 else int(round(v)))
            c.border = Border(top=fin, bottom=fin, left=fin, right=fin)
            c.alignment = Alignment(horizontal="left" if j == 1 else "right", indent=1)
            c.number_format = "# ##0"
            if (i - ligne0) % 2 == 0:
                c.fill = PatternFill("solid", fgColor=gris)
        ws.cell(i, 4).font = Font(bold=True, color=marine)
    total = ligne0 + len(commande) + 1
    ws.cell(total, 1, "Total").font = Font(bold=True)
    c = ws.cell(total, 4, "=SUM(D%d:D%d)" % (ligne0 + 1, total - 1))
    c.font = Font(bold=True, size=12, color=marine)
    c.number_format = "# ##0"
    c.alignment = Alignment(horizontal="right", indent=1)
    for j in range(1, 5):
        ws.cell(total, j).border = Border(top=Side(style="medium", color=marine))
    ws.cell(total + 3, 1, "Visa du pharmacien :").font = Font(bold=True, color="6B7685")
    ws.cell(total + 3, 3, "Date de réception :").font = Font(bold=True, color="6B7685")
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.print_options.horizontalCentered = True
    tampon = io.BytesIO()
    wb.save(tampon)
    return tampon.getvalue()


# ===========================================================================
# Stocks et alertes (onglet Stocks)
# ===========================================================================
ORDRE_ALERTE = {"RUPTURE": 0, "À COMMANDER": 1, "À saisir": 2, "OK": 3}


def page_stocks():
    filtres = st.container(key="filtres")
    c1, _ = filtres.columns([1, 3])
    mois = c1.selectbox("Mois à préparer", MOIS, index=MOIS_PROCHAIN - 1)
    t, nb_jours = etat_stock(MOIS.index(mois) + 1)
    st.write("")

    # Les urgences d'abord : rupture, à commander, à saisir, OK
    t = t.iloc[t["Alerte"].map(ORDRE_ALERTE).argsort(kind="stable")]
    cont = carte("etat", "État du stock, %s" % mois.lower())
    lignes = ""
    for g, r in t.iterrows():
        a_ = r["Alerte"]
        tend = ("%+.0f %%/an" % (100 * r["Tendance"])).replace(".", ",") if pd.notna(r["Tendance"]) else ""
        lignes += '<tr class="%s"><th>%s</th><td>%s</td><td>%s</td><td>%s</td><td class="qte">%s</td><td><span class="alerte %s">%s</span></td></tr>' % (
            {"RUPTURE": "l-rupture", "À COMMANDER": "l-commande"}.get(a_, ""), g, tend, nb(r["Stock en rayon"]),
            ("%s j" % nb(r["Autonomie (jours)"])) if pd.notna(r["Autonomie (jours)"]) else "",
            nb(r["Quantité à commander"]) if r["Quantité à commander"] > 0 else "",
            {"RUPTURE": "rupture", "À COMMANDER": "commande", "OK": "ok"}.get(a_, "saisir"), a_)
    de_mois = ("d'" if mois[0] in "AEIOUÉ" else "de ") + mois.lower()
    cont.markdown('<table class="bon"><thead><tr><th>Médicament</th><th>Tendance %s</th><th>En rayon</th>'
                  '<th>Autonomie</th><th>À commander</th><th>Alerte</th></tr></thead><tbody>%s</tbody></table>'
                  % (de_mois, lignes), unsafe_allow_html=True)
    commande = t[t["Quantité à commander"] > 0][["Stock en rayon", "Stock minimum conseillé", "Quantité à commander"]]
    if not commande.empty:
        with cont:
            b1, _ = st.columns([1, 2])
            b1.download_button("Télécharger le bon de commande", bon_de_commande_xlsx(commande.sort_index(), mois),
                               file_name="bon_de_commande_%s.xlsx" % dt.date.today().isoformat(),
                               mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                               type="primary", width="stretch")

    # Tendance du mois pour un médicament : ventes par jour de ce mois, année par année (données de la page Activité)
    tendance, par_an, cible = tendance_du_mois(MOIS.index(mois) + 1)
    cont = carte("tendance_mois", "Ventes %s, année par année" % de_mois)
    c1, _ = cont.columns([1, 3])
    med = c1.selectbox("Médicament", list(t.index), key="stock_med")
    annees = [int(a_) for a_ in par_an.index]
    valeurs = list(par_an[med].values)
    pente, origine = np.polyfit(annees, valeurs, 1) if len(annees) >= 3 else (0.0, float(np.mean(valeurs)))
    fig = go.Figure()
    fig.add_bar(x=[str(a_) for a_ in annees], y=valeurs, name="Ventes par jour", marker_color=BLEU,
                text=[nb(v, 1) for v in valeurs], textposition="outside", textfont=dict(color=TEXTE2, size=11),
                hovertemplate="%{x} : %{y:.1f} par jour<extra></extra>")
    fig.add_bar(x=[str(cible)], y=[tendance.loc[med, "Prévision par jour"]], name="Prévision", marker_color=BLEU_CLAIR,
                marker_pattern_shape="/", text=[nb(tendance.loc[med, "Prévision par jour"], 1)], textposition="outside",
                textfont=dict(color=TEXTE2, size=11), hovertemplate="%{x} (prévision) : %{y:.1f} par jour<extra></extra>")
    xs = annees + [cible]
    fig.add_scatter(x=[str(a_) for a_ in xs], y=[max(0, pente * a_ + origine) for a_ in xs], mode="lines",
                    name="Tendance", line=dict(color=MARINE, width=1.5, dash="dot"), hoverinfo="skip")
    fig.update_yaxes(rangemode="tozero", range=[0, max(valeurs + [tendance.loc[med, "Prévision par jour"]]) * 1.2])
    graphique(cont, fig, 300, legende=True)
    cont.markdown('<div class="source">Ventes par jour %s (mois complets, historique de la pharmacie). '
                  "Prévision %d : tendance des années prolongée d'un an ; elle fixe la quantité à commander.</div>"
                  % (de_mois, cible), unsafe_allow_html=True)

    cont = carte("inventaire", "Inventaire : stock compté en rayon")
    with cont.form("form_inventaire", clear_on_submit=True):
        cols = st.columns(4)
        for i, g in enumerate(ATC):
            cols[i % 4].number_input(g, min_value=0.0, step=1.0, format="%.0f", key="inv_" + g, value=None,
                                     placeholder="Compté")
        if st.form_submit_button("Enregistrer l'inventaire", type="primary"):
            for g in ATC:
                v = st.session_state.get("inv_" + g)
                if v is not None:
                    executer("INSERT OR REPLACE INTO inventaires (groupe, stock, date) VALUES (?, ?, ?)",
                             (g, float(v), maintenant()))
            st.rerun()


# ===========================================================================
# Scan (onglet Scan) : chaque scan est enregistré automatiquement dans le stock
# ===========================================================================
def valider_scan(cip, sens, quantite):
    produit = produit_du_code(cip)
    if produit is None:
        enregistrer_mouvement(NON_REFERENCE, sens, quantite, cip)
        st.session_state["scan_message"] = ("Code %s non référencé : mouvement enregistré, sans effet sur le stock "
                                            "des 8 groupes suivis." % cip, "À saisir")
        return
    libelle, groupe = produit
    avant, _ = etat_stock(MOIS_PROCHAIN)
    enregistrer_mouvement(groupe, sens, quantite, cip)
    apres, _ = etat_stock(MOIS_PROCHAIN)
    a, b = avant.loc[groupe], apres.loc[groupe]
    message = "%s (%s) : %s%s. Stock en rayon %s → %s." % (
        libelle, groupe, "+" if sens == "Entrée" else "-", nb(quantite), nb(a["Stock en rayon"]) or "0",
        nb(b["Stock en rayon"]))
    if a["Alerte"] != b["Alerte"]:
        message += " Statut : %s → %s." % (a["Alerte"], b["Alerte"])
    st.session_state["scan_message"] = (message, b["Alerte"])


def page_scan():
    c1, c2 = st.columns([1, 1])
    with c1:
        cont = carte("scanner", "Scanner un produit")
        sens = cont.radio("Type de mouvement", ["Entrée en stock", "Vente"], key="scan_sens", horizontal=True)
        sens = "Entrée" if sens == "Entrée en stock" else "Sortie"
        with cont.form("form_scan", clear_on_submit=True):
            f1, f2 = st.columns([3, 1])
            code = f1.text_input("Code scanné", placeholder="En attente de scan…", key="zone_scan")
            qte = f2.number_input("Boîtes", min_value=1, step=1, value=1)
            if st.form_submit_button("Valider", type="primary") and code.strip():
                valider_scan(code_cip13(code), sens, qte)
                st.rerun()
        # La douchette se comporte comme un clavier : elle écrit le code puis valide (Entrée).
        # Le curseur est placé automatiquement dans la zone de scan, à l'ouverture et après chaque scan.
        components.html("""<script>
            const focus = () => {
                const z = window.parent.document.querySelector('input[aria-label="Code scanné"]');
                if (z && window.parent.document.activeElement !== z) { z.focus(); }
            };
            focus(); setTimeout(focus, 300); setTimeout(focus, 1000);
        </script>""", height=0)
        if "scan_message" in st.session_state:
            message, alerte = st.session_state["scan_message"]
            {"RUPTURE": cont.error, "À COMMANDER": cont.warning, "À saisir": cont.info}.get(alerte, cont.success)(message)

    with c2:
        cont = carte("mouvements", "Derniers mouvements")
        mv = requete("SELECT m.id, m.date, m.cip13, p.libelle, m.groupe, m.sens, m.quantite FROM mouvements m "
                     "LEFT JOIN produits p ON p.cip13 = m.cip13 ORDER BY m.id DESC LIMIT 200")
        if mv.empty:
            cont.caption("Aucun mouvement.")
        else:
            aff = pd.DataFrame({"Date": pd.to_datetime(mv["date"]).dt.strftime("%d/%m/%Y %H:%M"),
                                "Code CIP13": mv["cip13"].fillna(""), "Produit": mv["libelle"].fillna(""),
                                "Groupe ATC": mv["groupe"], "Mouvement": mv["sens"], "Boîtes": mv["quantite"].map(nb)})
            cont.dataframe(aff.style.map(lambda v: "color:%s; font-weight:600" % (VERT if v == "Entrée" else ROUGE),
                                         subset=["Mouvement"]), width="stretch", hide_index=True, height=300)
            if cont.button("Annuler le dernier mouvement"):
                executer("DELETE FROM mouvements WHERE id = ?", (int(mv["id"].iloc[0]),))
                st.session_state.pop("scan_message", None)
                st.rerun()


# ===========================================================================
# Recherche (onglet Recherche) : sélection de groupes et d'un jour ou d'une période
# ===========================================================================
def page_recherche():
    jours = JOURS_DATA
    filtres = st.container(key="filtres")
    f1, f2, f3 = filtres.columns([2, 1.2, 1.6])
    sel = f1.multiselect("Code ATC", ATC, default=ATC, key="rech_g")
    dmin, dmax = jours["Date"].min().date(), jours["Date"].max().date()
    plage = f2.date_input("Date (jour ou période)", value=((pd.Timestamp(dmax) - pd.DateOffset(months=1)).date(), dmax),
                          min_value=dmin, max_value=dmax, format="DD/MM/YYYY")
    plage = plage if isinstance(plage, tuple) else (plage,)
    d1 = pd.Timestamp(plage[0])
    d2 = pd.Timestamp(plage[1]) if len(plage) == 2 else d1
    jours_sel = f3.pills("Jour", JOURS, selection_mode="multi", default=JOURS, key="rech_j",
                         format_func=lambda x: x[:3] + ".") or []
    st.write("")

    f = jours[(jours["Date"] >= d1) & (jours["Date"] <= d2) & (jours["Jour"].isin(jours_sel))]
    detail = f.melt(id_vars=["Date", "Jour"], value_vars=sel, var_name="Code ATC", value_name="Quantité vendue")
    if detail.empty:
        st.warning("Aucune vente pour cette sélection.")
        return
    detail = detail.sort_values(["Date", "Code ATC"])

    nb_jours = detail["Date"].nunique()
    total = detail["Quantité vendue"].sum()
    k = st.columns(3)
    kpi(k[0], "", nb(total, 1), "Total vendu")
    kpi(k[1], "", nb(nb_jours), "Journées")
    kpi(k[2], "", nb(total / nb_jours, 1), "Ventes par jour")
    st.write("")

    cont = carte("detail", "Ventes jour par jour")
    tab = detail.pivot_table(index=["Date", "Jour"], columns="Code ATC", values="Quantité vendue",
                             aggfunc="sum", observed=True).reindex(columns=sel)
    tab["Total"] = tab.sum(axis=1)
    tab = tab.reset_index()
    tab["Date"] = tab["Date"].dt.strftime("%d/%m/%Y")
    tab["Jour"] = tab["Jour"].astype(str)
    cont.dataframe(tab.style.format({c: (lambda v: nb(v, 1)) for c in sel + ["Total"]})
                   .set_properties(subset=["Total"], **{"font-weight": "700"}),
                   width="stretch", hide_index=True, height=min(420, 36 * len(tab) + 40))
    with cont:
        telecharger(tab.set_index("Date"), "ventes_%s_%s.csv" % (d1.strftime("%Y%m%d"), d2.strftime("%Y%m%d")))


# ===========================================================================
# Calendrier (onglet Calendrier) : ventes réelles, prévisions et environnement de la semaine
# ===========================================================================
@st.cache_data(show_spinner=False)
def profil_prevision(horodatage):
    """
    Prévision d'une journée, par médicament :
      moyenne historique des ventes du même jour de la semaine, dans le même mois
      x niveau récent (moyenne des 12 derniers mois complets / moyenne de tout l'historique).
    """
    j = JOURS_DATA.copy()
    j["mois"], j["jsem"] = j["Date"].dt.month, j["Date"].dt.weekday
    profil = j.groupby(["mois", "jsem"])[ATC].mean()
    fin = j.loc[j["Mois entier"] == "Oui", "Date"].max()
    recent = j[j["Date"] > fin - pd.DateOffset(months=12)][ATC].mean()
    niveau = (recent / j[ATC].mean()).fillna(1)
    return profil * niveau


def prevision(jour):
    p = profil_prevision(CLASSEUR.stat().st_mtime).loc[(jour.month, jour.weekday())]
    return p.reindex(ATC)


# Une semaine est « de pic » quand la moyenne de saison atteint 60 % du maximum annuel pour les épidémies,
# 40 % pour les pollens (arbres puis graminées) ; froid et chaleur = les 15 % de semaines les plus froides
# et les plus chaudes de l'année.
PICS = {
    "grippe": ("Grippe", "Pic de grippe"),
    "gastro": ("Gastro", "Pic de gastro"),
    "varicelle": ("Varicelle", "Pic de varicelle"),
    "bronchiolite": ("Bronchiolite", "Pic de bronchiolite"),
    "pollens": ("Pollens allergisants", "Pic de pollens"),
    "froid": ("Température", "Froid"),
    "chaleur": ("Température", "Chaleur"),
}


@st.cache_data(show_spinner=False)
def saisons_environnement(horodatage, ville):
    """Moyenne de chaque facteur par n° de semaine, et semaines de pic, pour la ville choisie."""
    h = semaines_environnement(ville)
    saison = h.groupby("N° semaine")[[c for c in ["Grippe", "Gastro", "Varicelle", "Température",
                                                  "Pollens allergisants", "Graminées"] if c in h]].mean()
    pics = {}
    for cle, col in [("grippe", "Grippe"), ("gastro", "Gastro"), ("varicelle", "Varicelle")]:
        if col in saison:
            pics[cle] = set(saison.index[saison[col] >= 0.6 * saison[col].max()])
    # bronchiolite : profil de saison 2020-2026 de la région (pas de période commune avec les ventes)
    info = VILLES[VILLES["Ville"] == ville]
    if "bronchiolite_saison" in D["externe"] and not info.empty:
        b = D["externe"]["bronchiolite_saison"]
        b = b[b["Région"] == int(info["Code région"].iloc[0])].set_index("N° semaine")["Bronchiolite"]
        if len(b):
            pics["bronchiolite"] = set(b.index[b >= 0.6 * b.max()])
    if "Pollens allergisants" in saison:
        # deux vagues : arbres (février à avril) puis graminées (mai-juin), chacune à 40 % de son maximum
        p = saison["Pollens allergisants"]
        pics["pollens"] = set(saison.index[p >= 0.4 * p.max()])
        if "Graminées" in saison:
            gr = saison["Graminées"]
            pics["pollens"] |= set(saison.index[gr >= 0.4 * gr.max()])
    if "Température" in saison:
        t = saison["Température"]
        pics["froid"] = set(saison.index[t <= t.quantile(0.15)])
        pics["chaleur"] = set(saison.index[t >= t.quantile(0.85)])
    return saison, pics


@st.cache_data(show_spinner=False)
def medicaments_qui_montent(horodatage, ville):
    """Pour chaque pic, les médicaments dont les ventes montent avec ce facteur (même calcul que la page Environnement)."""
    h = semaines_environnement(ville)
    sortie = {}
    for cle, (col, monte, baisse) in [("grippe", ("Grippe", "grippe", "")), ("gastro", ("Gastro", "gastro", "")),
                                      ("varicelle", ("Varicelle", "varicelle", "")),
                                      ("froid", ("Température", "", "froid")), ("chaleur", ("Température", "chaleur", "")),
                                      ("pollens", ("Graminées", "pollens", ""))]:
        if col not in h:
            continue
        liste = []
        for g in ATC:
            txt, _ = lien(h[col], h[g], "monte", "baisse")
            if (monte and txt == "monte") or (baisse and txt == "baisse"):
                liste.append(g)
        sortie[cle] = liste
    return sortie


def pics_semaine(lundi):
    """Périodes de pic de la semaine."""
    _, pics = saisons_environnement(CLASSEUR.stat().st_mtime, VILLE)
    n = min(int(pd.Timestamp(lundi).isocalendar()[1]), 52)
    return [cle for cle in ["grippe", "gastro", "varicelle", "bronchiolite", "pollens", "froid", "chaleur"]
            if n in pics.get(cle, set())]


@st.cache_data(show_spinner=False)
def temperatures_du_mois(horodatage, ville):
    """Température moyenne de chaque mois à la ville choisie, 2014-2019 (semaines rattachées au mois de leur jeudi)."""
    if "meteo_hebdo" not in D["externe"]:
        return pd.Series(dtype=float)
    m = D["externe"]["meteo_hebdo"]
    m = m[m["Ville"] == ville]
    return m.groupby((m["Lundi"] + pd.Timedelta(days=3)).dt.month)["Température"].mean()


def temperature_du_mois(mois_num):
    return temperatures_du_mois(CLASSEUR.stat().st_mtime, VILLE).get(mois_num, np.nan)


def jour_ferie(jour):
    f = D["externe"].get("jours_feries")
    if f is None:
        return None
    t = f.loc[f["Date"] == pd.Timestamp(jour), "Jour férié"]
    return t.iloc[0] if len(t) else None


def vacances(jour):
    v = D["externe"].get("vacances_scolaires")
    if v is None or ZONE is None:
        return None
    ts = pd.Timestamp(jour)
    t = v.loc[(v["Zone"] == ZONE) & (v["Début"] <= ts) & (v["Fin"] >= ts), "Vacances"]
    return t.iloc[0] if len(t) else None


def etiquettes(liste):
    return "".join('<span class="pic %s">%s</span>' % (c, PICS[c][1]) for c in liste)


def page_calendrier():
    aujourd_hui = dt.date.today()
    if "cal_mois" not in st.session_state:
        st.session_state["cal_annee"], st.session_state["cal_mois"] = aujourd_hui.year, MOIS[aujourd_hui.month - 1]
        st.session_state["cal_sel"] = ("jour", aujourd_hui)
    annees = list(range(ANNEES[0], max(ANNEES[-1], aujourd_hui.year) + 2))
    def aller_au(jour):
        st.session_state["cal_annee"], st.session_state["cal_mois"] = jour.year, MOIS[jour.month - 1]
        st.session_state["cal_sel"] = ("jour", jour)

    filtres = st.container(key="filtres")
    c1, c2, c3, _, c5 = filtres.columns([1, 1, 1, 1, 1], vertical_alignment="bottom")
    an = c1.selectbox("Année", annees, key="cal_annee")
    mo = c2.selectbox("Mois", MOIS, key="cal_mois")
    gsel = c3.selectbox("Médicament", ["Tous"] + ATC, key="cal_med")
    c5.button("Aujourd'hui", on_click=aller_au, args=(aujourd_hui,), width="stretch")
    st.write("")

    mn = MOIS.index(mo) + 1
    col = "Total" if gsel == "Tous" else gsel
    reel_jour = JOURS_DATA.set_index("Date")
    semaines = calendar.Calendar().monthdatescalendar(an, mn)

    # Valeur affichée dans chaque case : vente réelle si elle existe, sinon prévision
    valeurs = {}
    for sem in semaines:
        for d in sem:
            ts = pd.Timestamp(d)
            if ts in reel_jour.index:
                valeurs[d] = (reel_jour.loc[ts, col], True)
            else:
                p = prevision(d)
                valeurs[d] = (p.sum() if col == "Total" else p[col], False)
    du_mois = [v for d, (v, _) in valeurs.items() if d.month == mn]
    vmin, vmax = (min(du_mois), max(du_mois)) if du_mois else (0, 0)

    sel = st.session_state.get("cal_sel", ("jour", aujourd_hui))
    styles = []
    if True:
        t_mois = temperature_du_mois(mn)
        cont = carte("calendrier", "%s %d, %s%s" % (
            mo, an, "tous médicaments" if gsel == "Tous" else gsel,
            '<span class="temp-mois">Température moyenne de saison : %s °C</span>' % nb(t_mois, 1)
            if pd.notna(t_mois) else ""))
        largeurs = [0.5] + [1] * 7 + [1.5]
        entete = cont.columns(largeurs, gap="small")
        entete[0].markdown('<div class="cal-tete">Sem.</div>', unsafe_allow_html=True)
        for i, j in enumerate(JOURS_COURTS):
            entete[i + 1].markdown('<div class="cal-tete%s">%s</div>' % (" we" if i >= 5 else "", j),
                                   unsafe_allow_html=True)
        entete[8].markdown('<div class="cal-tete">Repères de la semaine</div>', unsafe_allow_html=True)
        for sem in semaines:
            ligne = cont.columns(largeurs, vertical_alignment="center", gap="small")
            lundi = sem[0]
            cle_s = "cal_s_%s" % lundi.strftime("%Y%m%d")
            if ligne[0].button("S%d" % lundi.isocalendar()[1], key=cle_s, width="stretch"):
                st.session_state["cal_sel"] = ("semaine", lundi)
                st.rerun()
            if sel == ("semaine", lundi):
                styles.append(".st-key-%s button {background:#1b2a3d !important; color:#fff !important;}" % cle_s)
            for i, d in enumerate(sem):
                cle = "cal_j_%s" % d.strftime("%Y%m%d")
                if d.month != mn:
                    ligne[i + 1].markdown('<div class="cal-vide%s"></div>' % (" we" if i >= 5 else ""),
                                          unsafe_allow_html=True)
                    continue
                v, est_reel = valeurs[d]
                etiquette = "%d  \n%s%s" % (d.day, "" if est_reel else "≈ ", nb(v, 0 if v >= 100 else 1))
                if ligne[i + 1].button(etiquette, key=cle, width="stretch"):
                    st.session_state["cal_sel"] = ("jour", d)
                    st.rerun()
                fond = couleur_echelle(v, vmin, vmax)
                regle = ".st-key-%s button {background:%s; %s}" % (
                    cle, fond, "" if est_reel else "font-style: italic; color:#4b5563;")
                if jour_ferie(d):
                    regle += (' .st-key-%s button::before {content: "Férié";}'
                              " .st-key-%s button {border-bottom: 3px solid #C00000 !important;}" % (cle, cle))
                elif vacances(d):
                    regle += " .st-key-%s button {border-bottom: 3px solid #1b2a3d !important;}" % cle
                if d == aujourd_hui:
                    regle += " .st-key-%s button {box-shadow: inset 0 0 0 2px #2d62b8;}" % cle
                if sel == ("jour", d) or (sel[0] == "semaine" and sel[1] == lundi):
                    regle += " .st-key-%s button {box-shadow: inset 0 0 0 2px #1b2a3d; font-weight:700;}" % cle
                styles.append(regle)
            vac = [v for v in dict.fromkeys(map(vacances, sem)) if v]
            ligne[8].markdown('<div class="cal-env">%s%s</div>' % (
                "".join('<div class="cal-vac">%s</div>' % v for v in vac), etiquettes(pics_semaine(lundi))),
                unsafe_allow_html=True)
        cont.markdown('<div class="cal-legende"><span class="trait vac"></span>Vacances scolaires (%s)'
                      '<span class="trait ferie"></span>Jour férié<span class="leg-prev">≈ 80,3</span>Prévision</div>'
                      '<div class="source">Sources des pics et de la température : Réseau Sentinelles (grippe, '
                      "gastro-entérite, varicelle, région %s, 2014-2019) ; Santé publique France, Odissé (bronchiolite, "
                      "région %s, depuis fin 2019) ; Météo-France (température, %s, 2014-2019) ; RNSA (pollens, %s, "
                      "2014-2019). Jours fériés : calendrier.api.gouv.fr (Etalab). Vacances scolaires : ministère de "
                      "l'Éducation nationale, data.education.gouv.fr.</div>" % (
                          (ZONE or "").replace("Zone", "zone"), REGION, REGION, VILLE, VILLE), unsafe_allow_html=True)
    st.markdown("<style>%s</style>" % "\n".join(styles), unsafe_allow_html=True)

    if True:
        # Détail de la sélection : prévision par médicament, réel si connu, environnement
        if sel[0] == "jour":
            jours_sel = [sel[1]]
            titre = "%s %d %s %d" % (JOURS[sel[1].weekday()], sel[1].day, MOIS[sel[1].month - 1].lower(), sel[1].year)
        else:
            jours_sel = [sel[1] + dt.timedelta(days=i) for i in range(7)]
            fin = jours_sel[-1]
            titre = "Semaine du %d %s au %d %s %d" % (sel[1].day, MOIS[sel[1].month - 1].lower(), fin.day,
                                                     MOIS[fin.month - 1].lower(), fin.year)
        prev = sum(prevision(d) for d in jours_sel)
        connus = [pd.Timestamp(d) for d in jours_sel if pd.Timestamp(d) in reel_jour.index]
        cont = carte("detail_cal", titre)
        d1, d2 = cont.columns([3, 2], gap="large")
        tab = pd.DataFrame({"Prévision": prev})
        if len(connus) == len(jours_sel):
            tab["Réel"] = reel_jour.loc[connus, ATC].sum()
            tab["Écart"] = tab["Réel"] / tab["Prévision"] - 1
        tab.loc["Total"] = tab.sum(numeric_only=True)
        if "Écart" in tab:
            tab.loc["Total", "Écart"] = tab.loc["Total", "Réel"] / tab.loc["Total", "Prévision"] - 1
        tab.index.name = "Médicament"
        fmt = {"Prévision": lambda v: nb(v, 1), "Réel": lambda v: nb(v, 1),
               "Écart": lambda v: ("%+.0f %%" % (100 * v)).replace(".", ",")}
        d1.dataframe(tab.style.format({k_: f for k_, f in fmt.items() if k_ in tab})
                       .set_properties(subset=pd.IndexSlice[["Total"], :], **{"font-weight": "700"}),
                       width="stretch", height=36 * len(tab) + 40)
        pics_sem = pics_semaine(jours_sel[0] - dt.timedelta(days=jours_sel[0].weekday()))
        html = ""
        feries = ["%s (%d %s)" % (jour_ferie(d), d.day, MOIS[d.month - 1].lower()) if len(jours_sel) > 1
                  else jour_ferie(d) for d in jours_sel if jour_ferie(d)]
        if feries:
            html += '<div class="ligne-synth">Jour férié<b>%s</b></div>' % ", ".join(feries)
        vacs = list(dict.fromkeys(v for v in map(vacances, jours_sel) if v))
        if vacs:
            html += '<div class="ligne-synth">Vacances scolaires (%s)<b>%s</b></div>' % (
                ZONE.replace("Zone", "zone"), ", ".join(vacs))
        montent = medicaments_qui_montent(CLASSEUR.stat().st_mtime, VILLE)
        for c in pics_sem:
            meds = montent.get(c, [])
            html += '<div class="ligne-synth">%s<b>%s</b></div>' % (
                etiquettes([c]), ("Ventes en hausse : " + ", ".join(meds)) if meds else (
                    "Repère de saison" if c == "bronchiolite" else "Pas d'effet sur les ventes"))
        if not pics_sem:
            html += '<div class="ligne-synth">Pics de la semaine<b>Aucun</b></div>'
        d2.markdown(html, unsafe_allow_html=True)


# ===========================================================================
# Paramètres : localisation de la pharmacie et nom affiché sur le bon de commande
# ===========================================================================
def page_parametres():
    cont = carte("parametres", "Profil de la pharmacie")
    villes = list(VILLES["Ville"])
    with cont.form("form_parametres"):
        c1, c2 = st.columns(2)
        nom = c1.text_input("Nom de la pharmacie", value=parametre("nom"))
        titulaire = c2.text_input("Pharmacien titulaire", value=parametre("titulaire"))
        adresse = c1.text_input("Adresse", value=parametre("adresse"))
        ville = c2.selectbox("Ville", villes, index=villes.index(VILLE) if VILLE in villes else 0)
        telephone = c1.text_input("Téléphone", value=parametre("telephone"), placeholder="04 67 00 00 00")
        email = c2.text_input("Adresse e-mail", value=parametre("email"), placeholder="contact@pharmacie.fr")
        if st.form_submit_button("Enregistrer", type="primary"):
            erreurs = []
            if email.strip() and not re.match(r"^[^@\s]+@[^@\s]+\.[^@\s]+$", email.strip()):
                erreurs.append("L'adresse e-mail n'est pas valide.")
            if telephone.strip() and len(re.sub(r"\D", "", telephone)) != 10:
                erreurs.append("Le numéro de téléphone doit contenir 10 chiffres.")
            if erreurs:
                for e in erreurs:
                    st.error(e)
            else:
                enregistrer_parametre("nom", nom.strip() or PARAMETRES_DEFAUT["nom"])
                enregistrer_parametre("titulaire", titulaire.strip())
                enregistrer_parametre("adresse", adresse.strip())
                enregistrer_parametre("ville", ville)
                enregistrer_parametre("telephone", telephone.strip())
                enregistrer_parametre("email", email.strip())
                st.rerun()

    if INFO_VILLE is not None:
        cont = carte("localisation", "Localisation")
        c1, c2, _ = cont.columns([1, 1, 2])
        c1.markdown('<div class="ligne-synth">Région<b>%s</b></div>' % INFO_VILLE["Région"], unsafe_allow_html=True)
        c2.markdown('<div class="ligne-synth">Département<b>%s</b></div>' % INFO_VILLE["Département"],
                    unsafe_allow_html=True)


# ===========================================================================
# Navigation
# ===========================================================================
st.logo(str(ICI / "logo.svg"), size="large")

navigation = st.navigation({
    "Propriétaire": [
        st.Page(page_activite, title="Activité", url_path="activite", default=True),
        st.Page(page_environnement, title="Environnement", url_path="environnement"),
    ],
    "Pharmacien": [
        st.Page(page_stocks, title="Stocks et alertes", url_path="stocks"),
        st.Page(page_scan, title="Scan des ventes", url_path="scan"),
    ],
    "Outils communs": [
        st.Page(page_recherche, title="Recherche", url_path="recherche"),
        st.Page(page_calendrier, title="Calendrier", url_path="calendrier"),
    ],
    "Réglages": [
        st.Page(page_parametres, title="Paramètres", url_path="parametres"),
    ],
})

navigation.run()
