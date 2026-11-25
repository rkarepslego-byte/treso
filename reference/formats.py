"""Paramètres d'import des deux banques (table tbl_ParamImport, onglet masqué Param_Import).

Les valeurs ci-dessous correspondent aux exports réels du client (BPGO et CIC). En cas de
changement de format par une banque, seule la table tbl_ParamImport du classeur doit être
adaptée (aide : python -m outils.analyser_releve fichier).
"""

# Ordre et libellés des paramètres tels qu'affichés dans tbl_ParamImport
# (clé M, libellé affiché, description)
PARAMETRES_IMPORT = [
    ("Extensions", "Extensions acceptées", "Liste séparée par « ; » (csv;txt;xlsx;xls)"),
    ("Separateur", "Séparateur CSV", "Caractère séparant les colonnes d'un CSV (; , ou TAB)"),
    ("Encodage", "Encodage CSV", "1252 = Windows (ANSI), 65001 = UTF-8"),
    ("LignesIgnorees", "Lignes à ignorer avant l'en-tête", "Nombre de lignes de présentation au début du fichier"),
    ("Feuille", "Feuille Excel", "Nom de l'onglet à lire (vide = premier onglet)"),
    ("ColDate", "Colonne date d'opération", "Nom exact de la colonne (casse et accents indifférents)"),
    ("ColDateValeur", "Colonne date de valeur", "Facultatif"),
    ("FormatDate", "Format de date", "Format des dates texte : dd/MM/yyyy, dd/MM/yy, yyyy-MM-dd..."),
    ("ColsLibelle", "Colonnes libellé", "Une ou plusieurs colonnes séparées par « ; » (concaténées)"),
    ("ModeMontant", "Mode montant", "Montant signé | Débit/Crédit | Montant + Sens"),
    ("ColMontant", "Colonne montant", "Mode « Montant signé » ou « Montant + Sens »"),
    ("ColDebit", "Colonne débit", "Mode « Débit/Crédit »"),
    ("ColCredit", "Colonne crédit", "Mode « Débit/Crédit »"),
    ("ColSens", "Colonne sens", "Mode « Montant + Sens » (ex. D / C)"),
    ("ValeursDebit", "Valeurs signifiant un débit", "Mode « Montant + Sens », séparées par « ; »"),
    ("InverserSigne", "Inverser le signe", "Oui si la banque exporte les débits en positif (mode Montant signé)"),
    ("SeparateurDecimal", "Séparateur décimal", "« , » (France) ou « . »"),
    ("ColReference", "Colonne référence", "Facultatif : n° d'opération unique fourni par la banque"),
]

# Formats RÉELS des exports du client (analysés le 25/11/2026 avec outils/analyser_releve.py) :
# * Banque 1 = BPGO (Banque Populaire Grand Ouest) : CSV « ; » Windows-1252, 4 lignes de présentation
#   (code banque / compte / ligne vide / solde de fin de période), débits négatifs et crédits « +… »
#   dans deux colonnes, référence d'opération, pied « Solde en début de période ».
# * Banque 2 = CIC : classeur .xlsx « comptes_N.xlsx » ; 1er onglet = récapitulatif des comptes,
#   opérations du compte courant dans l'onglet « Cpt 14270 00020886002 » (4 lignes de présentation,
#   débits négatifs / crédits positifs), pied « Solde au … ».
FORMAT_B1 = {
    "Code": "B1",
    "Extensions": "csv;txt",
    "Separateur": ";",
    "Encodage": 1252,
    "LignesIgnorees": 4,
    "Feuille": "",
    "ColDate": "Date comptable",
    "ColDateValeur": "Date de valeur",
    "FormatDate": "dd/MM/yyyy",
    "ColsLibelle": "Libelle simplifie;Informations complementaires",
    "ModeMontant": "Débit/Crédit",
    "ColMontant": "",
    "ColDebit": "Debit",
    "ColCredit": "Credit",
    "ColSens": "",
    "ValeursDebit": "D;DEBIT",
    "InverserSigne": "Non",
    "SeparateurDecimal": ",",
    "ColReference": "Reference",
}

FORMAT_B2 = {
    "Code": "B2",
    "Extensions": "xlsx;xls",
    "Separateur": ";",
    "Encodage": 1252,
    "LignesIgnorees": 4,
    "Feuille": "Cpt 14270 00020886002",
    "ColDate": "Date",
    "ColDateValeur": "Valeur",
    "FormatDate": "dd/MM/yyyy",
    "ColsLibelle": "Libellé",
    "ModeMontant": "Débit/Crédit",
    "ColMontant": "",
    "ColDebit": "Débit",
    "ColCredit": "Crédit",
    "ColSens": "",
    "ValeursDebit": "D;DEBIT",
    "InverserSigne": "Non",
    "SeparateurDecimal": ",",
    "ColReference": "",
}


def vers_parametres_moteur(fmt: dict, nom_banque: str, dossier: str = "") -> dict:
    """Convertit une colonne de tbl_ParamImport en record de paramètres (comme le fait le VBA)."""
    p = dict(fmt)
    p["Banque"] = nom_banque
    p["Dossier"] = dossier
    p["Extensions"] = ["." + e.strip().lower() for e in fmt["Extensions"].split(";") if e.strip()]
    p["Encodage"] = int(fmt["Encodage"])
    p["LignesIgnorees"] = int(fmt["LignesIgnorees"])
    p["ValeursDebit"] = [v.strip().upper() for v in fmt["ValeursDebit"].split(";") if v.strip()]
    p["InverserSigne"] = str(fmt["InverserSigne"]).strip().lower() == "oui"
    return p
