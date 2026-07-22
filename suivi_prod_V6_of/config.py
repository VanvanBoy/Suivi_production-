# -*- coding: utf-8 -*-
"""
Constantes de configuration de l'application Suivi de production.
"""

EXCEL_PATH = r"G:\Drive partagés\11_Data\IHM\Instructions IHM\Suivi_prod_par_modele.xlsx"

# Correspondance entre les clés "stage" utilisées cote IHM
# et les noms de colonnes réelles dans la table suivi_production (MySQL)
STAGE_TO_DBCOL = {
    "picking":        "picking_tension",
    "pack":           "soudure_pack",
    "nappe":          "soudure_nappe",
    "bms":            "soudure_bms",
    "wrap":           "wrap",
    "fermeture_batt": "fermeture_batt",
    "capa":           "test_capa",
    "emb":            "emballage",
    "exp":            "expedition",
    "recherche":      "recherche",
    "recyclage":      "recyclage",
    "tri_test":       "tri_test",
    "banc_somfy":     "banc_somfy",
    "fin_ligne":      "fin_ligne",
}

ALLOWED_STAGE_KEYS = set(STAGE_TO_DBCOL.keys())

# --- Chemins réseau (lecteur partagé) -----------------------------------
# Centralisés ici pour n'avoir qu'un seul endroit à modifier si l'arborescence
# du lecteur partagé change (auparavant recopiés en dur dans plusieurs onglets).

# Logo VoltR affiché dans les onglets bms, fermeture, nappe
VOLTR_LOGO_PATH = r"G:\Drive partagés\11_Data\IHM\Executable\IHM_suivi_prod_beta\Suivi_prod_rsc\voltr_logo.jpg"

# Onglet picking - remplacement de cellule
REMPLACEMENT_DIR = r"G:\Drive partagés\4_Production\8_Picking\remplacement\Remplacement en cours"
TEMPLATE_REMPLACEMENT_PATH = r"G:\Drive partagés\11_Data\IHM\Executable\IHM_suivi_prod_beta\Suivi_prod_rsc\Template remplacement cellule (13).xlsx"

# Onglet capa - dossiers de dépouillement des résultats de cyclage
CAPA_DOSSIER_ENTRANT = r"G:\Drive partagés\Résultat cyclage\1_Résultats de cyclage\Fichier en cours"
CAPA_DOSSIER_EXPLOITES = r"G:\Drive partagés\Résultat cyclage\1_Résultats de cyclage\Fichiers traités\Batteries"
CAPA_DOSSIER_KO = r"G:\Drive partagés\Résultat cyclage\1_Résultats de cyclage\Fichiers NOK\Batteries KO"

# URL du formulaire Google de déclaration de non-conformité, ouvert dans le
# navigateur depuis (quasiment) tous les onglets de contrôle.
NON_CONFORMITE_FORM_URL = "https://docs.google.com/forms/d/e/1FAIpQLSeDivu0XsxeXnRhJrf1AyoVaywsDtKyPdaCJ9_-EfSQ-3-x7A/viewform?usp=sf_link"

# Onglet banc Somfy - fichier de mesures du banc de test physique (bouton
# "Charger les valeurs de test")
BANC_TEST_SOMFY_XLSM_PATH = r"G:\Drive partagés\2_Sales\2_Clients\Somfy\1_Fourniture\1_Sur mesure\10_Banc de Tests\Banc de test fin de ligne\IHM_BANC_TEST_SOMFY.xlsm"
