"""Paramètres d'import des formats de DÉMONSTRATION.

Ces paramètres décrivent les deux formats fictifs livrés en attendant les vrais exports :

* Banque 1 : CSV séparateur « ; », encodage Windows-1252, 3 lignes d'en-tête de banque,
  une colonne « Montant » signée au format français (1 234,56).
* Banque 2 : classeur Excel .xlsx, colonnes Débit / Crédit séparées (débits positifs),
  libellé réparti sur deux colonnes.

Les mêmes valeurs alimentent la table tbl_ParamImport du classeur (onglet masqué
Param_Import). Quand les vrais exports seront fournis, seule cette table (et au besoin
la fonction Power Query fnNormaliserReleve) devra être adaptée.
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

FORMAT_B1 = {
    "Code": "B1",
    "Extensions": "csv;txt",
    "Separateur": ";",
    "Encodage": 1252,
    "LignesIgnorees": 3,
    "Feuille": "",
    "ColDate": "Date opération",
    "ColDateValeur": "Date valeur",
    "FormatDate": "dd/MM/yyyy",
    "ColsLibelle": "Libellé",
    "ModeMontant": "Montant signé",
    "ColMontant": "Montant (EUR)",
    "ColDebit": "",
    "ColCredit": "",
    "ColSens": "",
    "ValeursDebit": "D;DEBIT",
    "InverserSigne": "Non",
    "SeparateurDecimal": ",",
    "ColReference": "",
}

FORMAT_B2 = {
    "Code": "B2",
    "Extensions": "xlsx;xls;csv",
    "Separateur": ";",
    "Encodage": 65001,
    "LignesIgnorees": 0,
    "Feuille": "",
    "ColDate": "Date",
    "ColDateValeur": "",
    "FormatDate": "dd/MM/yyyy",
    "ColsLibelle": "Libellé opération;Détail",
    "ModeMontant": "Débit/Crédit",
    "ColMontant": "",
    "ColDebit": "Débit",
    "ColCredit": "Crédit",
    "ColSens": "",
    "ValeursDebit": "D;DEBIT",
    "InverserSigne": "Non",
    "SeparateurDecimal": ",",
    "ColReference": "N° opération",
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
