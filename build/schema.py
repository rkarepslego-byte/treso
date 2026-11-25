"""Structure du classeur : onglets, tables, colonnes, formules et noms définis.

Source unique de vérité pour le générateur (generer_classeur.py) et pour le
contrôle du code VBA (tests/outils/lint_vba.py vérifie que chaque colonne et chaque
nom utilisés par le VBA existent ici).
"""

from dataclasses import dataclass, field

# ---------------------------------------------------------------------------
# Onglets (nom affiché, nom de code VBA, visible ?)
# ---------------------------------------------------------------------------
F_TDB = "Tableau de bord"
F_SAISIE = "Saisie"
F_CHARGES = "Charges fixes"
F_SUIVI = "Suivi prévisionnel"
F_RAPPRO = "Rapprochement"
F_ATRAITER = "À traiter"
F_BANQUE = "Opérations bancaires"
F_PARAM = "Paramètres"
F_AIDE = "Aide"
F_IMPORT_B1 = "Import_B1"
F_IMPORT_B2 = "Import_B2"
F_HISTO = "Historique"
F_CALC = "Calculs"
F_LISTES = "Listes"
F_PARAM_IMPORT = "Param_Import"

ONGLETS = [
    (F_TDB, "shTableauDeBord", True),
    (F_SAISIE, "shSaisie", True),
    (F_CHARGES, "shCharges", True),
    (F_SUIVI, "shSuivi", True),
    (F_RAPPRO, "shRapprochement", True),
    (F_ATRAITER, "shATraiter", True),
    (F_BANQUE, "shBanque", True),
    (F_PARAM, "shParametres", True),
    (F_AIDE, "shAide", True),
    (F_IMPORT_B1, "shImportB1", False),
    (F_IMPORT_B2, "shImportB2", False),
    (F_HISTO, "shHistorique", False),
    (F_CALC, "shCalculs", False),
    (F_LISTES, "shListes", False),
    (F_PARAM_IMPORT, "shParamImport", False),
]

MOT_DE_PASSE = "Treso2026"
NB_LIGNES_LIBRES = 30

# Formats (le code de format est enregistré en notation anglaise dans le fichier ;
# Excel l'affiche selon les paramètres régionaux : 1 234,56 €)
FMT_EUR = '#,##0.00\\ "€";[Red]\\-#,##0.00\\ "€"'
FMT_EUR0 = '#,##0\\ "€";[Red]\\-#,##0\\ "€"'
FMT_DATE = "dd/mm/yyyy"
FMT_DATEHEURE = "dd/mm/yyyy hh:mm"
FMT_TEXTE = "@"
FMT_PCT = "0%"
FMT_ENTIER = "0"

# Genres de colonnes
SAISIE = "saisie"      # renseignée par l'utilisateur (cellule déverrouillée)
SYSTEME = "systeme"    # renseignée par les macros (verrouillée)
FORMULE = "formule"    # colonne calculée (verrouillée)


@dataclass
class Colonne:
    nom: str
    genre: str
    fmt: str | None = None
    largeur: float = 12
    formule: str | None = None
    liste: str | None = None       # nom défini servant de liste déroulante
    masquee: bool = False


@dataclass
class Table:
    nom: str
    onglet: str
    ligne_entete: int      # numéro de ligne Excel (1 = première ligne)
    col_debut: int         # 0 = A, 1 = B
    colonnes: list[Colonne] = field(default_factory=list)
    style: str = "Table Style Light 9"
    lignes_libres: int = 0

    def noms(self) -> list[str]:
        return [c.nom for c in self.colonnes]

    def index(self, nom: str) -> int:
        return self.noms().index(nom)


# ---------------------------------------------------------------------------
# Formules réutilisées
# ---------------------------------------------------------------------------
_D0_PREV = "IF([@Banque]=B1_Nom,B1_DateSolde,IF([@Banque]=B2_Nom,B2_DateSolde,\"\"))"

PREVISIONS = Table("tbl_Previsions", F_SAISIE, 10, 1, [
    Colonne("ID", SYSTEME, FMT_TEXTE, 10),
    Colonne("Date prévue", SAISIE, FMT_DATE, 12),
    Colonne("Banque", SAISIE, FMT_TEXTE, 17, liste="L_Banques"),
    Colonne("Libellé", SAISIE, FMT_TEXTE, 34),
    Colonne("Catégorie", SAISIE, FMT_TEXTE, 19, liste="L_Categories"),
    Colonne("Type", SAISIE, FMT_TEXTE, 16, liste="L_Types"),
    Colonne("Montant", SAISIE, FMT_EUR, 14),
    Colonne("Commentaire", SAISIE, FMT_TEXTE, 26),
    Colonne("Statut", SAISIE, FMT_TEXTE, 11, liste="L_StatutsPrev"),
    Colonne("Montant signé", FORMULE, FMT_EUR, 14,
            '=IF([@Montant]="","",IF(OR([@Type]="Recette",[@Type]="Virement entrant"),ABS([@Montant]),-ABS([@Montant])))'),
    Colonne("Suivi", FORMULE, None, 12,
            '=IF(AND([@ID]="",[@Libellé]="",[@Montant]=""),"",IF([@Statut]="Pointé","Pointé",'
            'IF([@Statut]="Annulé","Annulé",IF(OR([@[ID proposé]]<>"",AND([@[Date prévue]]<>"",'
            '[@[Date prévue]]<Date_Reference)),"À contrôler","Prévu"))))'),
    Colonne("Contrôle", FORMULE, None, 24,
            '=IF(AND([@ID]="",[@Libellé]="",[@Montant]=""),"",IF([@[Date prévue]]="","Date manquante",'
            'IF(AND([@Banque]<>B1_Nom,[@Banque]<>B2_Nom),"Banque inconnue",IF([@Montant]="","Montant manquant",'
            'IF(AND([@Statut]="Pointé",[@[ID opération]]=""),"Pointé sans opération",'
            'IF(AND([@Statut]<>"Pointé",[@[ID opération]]<>""),"Lien bancaire incohérent",'
            'IF(AND(LEFT([@Type],8)="Virement",[@[Lien virement]]=""),"Virement sans contrepartie",'
            'IF([@[Date prévue]]<=' + _D0_PREV + ',"Antérieure au solde initial",""))))))))'),
    Colonne("Origine", SYSTEME, FMT_TEXTE, 12),
    Colonne("Réf charge", SYSTEME, FMT_TEXTE, 10),
    Colonne("ID opération", SYSTEME, FMT_TEXTE, 12),
    Colonne("ID proposé", SYSTEME, FMT_TEXTE, 12),
    Colonne("Date réelle", SYSTEME, FMT_DATE, 12),
    Colonne("Montant réel", SYSTEME, FMT_EUR, 14),
    Colonne("Lien virement", SYSTEME, FMT_TEXTE, 12),
    Colonne("Empreinte", SYSTEME, FMT_TEXTE, 10, masquee=True),
    Colonne("Montant à venir", FORMULE, FMT_EUR, 12,
            '=IF(OR([@Montant]="",[@[Date prévue]]="",[@Statut]="Pointé",[@Statut]="Annulé",'
            '[@[ID opération]]<>"",[@[ID proposé]]<>""),0,IF([@Banque]=B1_Nom,IF([@[Date prévue]]>B1_DateSolde,'
            '[@[Montant signé]],0),IF([@Banque]=B2_Nom,IF([@[Date prévue]]>B2_DateSolde,[@[Montant signé]],0),0)))',
            masquee=True),
    Colonne("Date effective", FORMULE, FMT_DATE, 12,
            '=IF([@[Montant à venir]]=0,"",MAX([@[Date prévue]],Date_Reference))', masquee=True),
    Colonne("Clé tri", FORMULE, FMT_ENTIER, 12,
            '=IF([@[Date effective]]="","",([@[Date effective]]-40000)*100000+50000+ROW())', masquee=True),
    Colonne("Inclus banque", FORMULE, FMT_ENTIER, 8,
            '=IF(AND([@[Montant à venir]]<>0,OR(Suivi_Banque="Toutes",Suivi_Banque=[@Banque])),1,0)', masquee=True),
    Colonne("Inclus vue", FORMULE, FMT_ENTIER, 8,
            '=IF(AND([@[Inclus banque]]=1,[@[Date effective]]>=Suivi_Du,[@[Date effective]]<=Suivi_Au,'
            'Suivi_Afficher<>"Réel uniquement"),1,0)', masquee=True),
    Colonne("Rang vue", FORMULE, FMT_ENTIER, 8,
            '=IF([@[Inclus vue]]=1,COUNTIFS(tbl_Banque[Inclus vue],1,tbl_Banque[Clé tri],"<"&[@[Clé tri]])'
            '+COUNTIFS(tbl_Previsions[Inclus vue],1,tbl_Previsions[Clé tri],"<"&[@[Clé tri]])+1,"")', masquee=True),
], lignes_libres=NB_LIGNES_LIBRES)

CHARGES = Table("tbl_Charges", F_CHARGES, 7, 1, [
    Colonne("ID", SYSTEME, FMT_TEXTE, 9),
    Colonne("Libellé", SAISIE, FMT_TEXTE, 30),
    Colonne("Banque", SAISIE, FMT_TEXTE, 17, liste="L_Banques"),
    Colonne("Catégorie", SAISIE, FMT_TEXTE, 18, liste="L_Categories"),
    Colonne("Type", SAISIE, FMT_TEXTE, 11, liste="L_TypesCharge"),
    Colonne("Montant", SAISIE, FMT_EUR, 13),
    Colonne("Première échéance", SAISIE, FMT_DATE, 13),
    Colonne("Fréquence", SAISIE, FMT_TEXTE, 14, liste="L_Frequences"),
    Colonne("Date de fin", SAISIE, FMT_DATE, 12),
    Colonne("Actif", SAISIE, FMT_TEXTE, 8, liste="L_OuiNon"),
    Colonne("Commentaire", SAISIE, FMT_TEXTE, 26),
    Colonne("Pas", FORMULE, FMT_ENTIER, 6,
            '=IF([@Fréquence]="Mensuelle",1,IF([@Fréquence]="Trimestrielle",3,IF([@Fréquence]="Semestrielle",6,'
            'IF([@Fréquence]="Annuelle",12,""))))', masquee=True),
    Colonne("Prochaine échéance", FORMULE, FMT_DATE, 13,
            '=IF(OR([@Libellé]="",[@[Première échéance]]="",[@Pas]=""),"",IF([@Actif]="Non","Inactive",'
            'IF([@[Première échéance]]>=Date_Reference,[@[Première échéance]],'
            'IF(EDATE([@[Première échéance]],[@Pas]*INT(DATEDIF([@[Première échéance]],Date_Reference,"m")/[@Pas]))'
            '>=Date_Reference,EDATE([@[Première échéance]],[@Pas]*INT(DATEDIF([@[Première échéance]],'
            'Date_Reference,"m")/[@Pas])),EDATE([@[Première échéance]],[@Pas]*(INT(DATEDIF([@[Première échéance]],'
            'Date_Reference,"m")/[@Pas])+1))))))'),
    Colonne("Équivalent mensuel", FORMULE, FMT_EUR, 14,
            '=IF(OR([@Montant]="",[@Pas]="",[@Actif]="Non"),"",IF([@Type]="Recette",1,-1)*ABS([@Montant])/[@Pas])'),
    Colonne("Échéances à venir", FORMULE, FMT_ENTIER, 11,
            '=IF([@ID]="","",COUNTIFS(tbl_Previsions[Réf charge],[@ID],tbl_Previsions[Suivi],"Prévu"))'),
], lignes_libres=NB_LIGNES_LIBRES)

BANQUE = Table("tbl_Banque", F_BANQUE, 7, 1, [
    Colonne("ID", SYSTEME, FMT_TEXTE, 11),
    Colonne("Banque", SYSTEME, FMT_TEXTE, 16),
    Colonne("Date", SYSTEME, FMT_DATE, 11),
    Colonne("Date valeur", SYSTEME, FMT_DATE, 11),
    Colonne("Libellé", SYSTEME, FMT_TEXTE, 40),
    Colonne("Montant", SYSTEME, FMT_EUR, 14),
    Colonne("Référence", SYSTEME, FMT_TEXTE, 12),
    Colonne("Catégorie", SAISIE, FMT_TEXTE, 18, liste="L_Categories"),
    Colonne("Commentaire", SAISIE, FMT_TEXTE, 24),
    Colonne("Statut", SYSTEME, FMT_TEXTE, 13),
    Colonne("ID prévision", SYSTEME, FMT_TEXTE, 11),
    Colonne("Mode", SYSTEME, FMT_TEXTE, 11),
    Colonne("ID import", SYSTEME, FMT_TEXTE, 10),
    Colonne("Fichier", SYSTEME, FMT_TEXTE, 18),
    Colonne("Date import", SYSTEME, FMT_DATEHEURE, 15),
    Colonne("Clé", SYSTEME, FMT_TEXTE, 12, masquee=True),
    Colonne("Inclus", FORMULE, FMT_ENTIER, 7,
            '=IF([@ID]="",0,IF([@Banque]=B1_Nom,IF([@Date]>B1_DateSolde,1,0),IF([@Banque]=B2_Nom,'
            'IF([@Date]>B2_DateSolde,1,0),0)))', masquee=True),
    Colonne("Clé tri", FORMULE, FMT_ENTIER, 12,
            '=IF([@Inclus]=1,([@Date]-40000)*100000+ROW(),"")', masquee=True),
    Colonne("Inclus banque", FORMULE, FMT_ENTIER, 8,
            '=IF(AND([@Inclus]=1,OR(Suivi_Banque="Toutes",Suivi_Banque=[@Banque])),1,0)', masquee=True),
    Colonne("Inclus vue", FORMULE, FMT_ENTIER, 8,
            '=IF(AND([@[Inclus banque]]=1,[@Date]>=Suivi_Du,[@Date]<=Suivi_Au,Suivi_Afficher<>"Prévisions uniquement"),1,0)',
            masquee=True),
    Colonne("Rang vue", FORMULE, FMT_ENTIER, 8,
            '=IF([@[Inclus vue]]=1,COUNTIFS(tbl_Banque[Inclus vue],1,tbl_Banque[Clé tri],"<"&[@[Clé tri]])'
            '+COUNTIFS(tbl_Previsions[Inclus vue],1,tbl_Previsions[Clé tri],"<"&[@[Clé tri]])+1,"")', masquee=True),
])

PROPOSITIONS = Table("tbl_Propositions", F_RAPPRO, 8, 1, [
    Colonne("Valider", SAISIE, FMT_TEXTE, 9, liste="L_OuiNon"),
    Colonne("Niveau", SYSTEME, FMT_TEXTE, 11),
    Colonne("Groupe", SYSTEME, FMT_ENTIER, 8),
    Colonne("Suggestion", SYSTEME, FMT_TEXTE, 11),
    Colonne("Banque", SYSTEME, FMT_TEXTE, 15),
    Colonne("Date opération", SYSTEME, FMT_DATE, 12),
    Colonne("Libellé bancaire", SYSTEME, FMT_TEXTE, 34),
    Colonne("Montant", SYSTEME, FMT_EUR, 13),
    Colonne("Date prévue", SYSTEME, FMT_DATE, 12),
    Colonne("Libellé prévu", SYSTEME, FMT_TEXTE, 28),
    Colonne("Écart (jours)", SYSTEME, FMT_ENTIER, 9),
    Colonne("Similarité", SYSTEME, FMT_PCT, 10),
    Colonne("Score", SYSTEME, FMT_PCT, 8),
    Colonne("ID opération", SYSTEME, FMT_TEXTE, 11),
    Colonne("ID prévision", SYSTEME, FMT_TEXTE, 11),
])

ATRAITER = Table("tbl_ATraiter", F_ATRAITER, 8, 1, [
    Colonne("ID opération", SYSTEME, FMT_TEXTE, 11),
    Colonne("Date", SYSTEME, FMT_DATE, 11),
    Colonne("Banque", SYSTEME, FMT_TEXTE, 15),
    Colonne("Libellé", SYSTEME, FMT_TEXTE, 36),
    Colonne("Montant", SYSTEME, FMT_EUR, 13),
    Colonne("Suggestion", SYSTEME, FMT_TEXTE, 30),
    Colonne("Catégorie", SAISIE, FMT_TEXTE, 18, liste="L_Categories"),
    Colonne("Commentaire", SAISIE, FMT_TEXTE, 24),
    Colonne("Action", SAISIE, FMT_TEXTE, 30, liste="L_Actions"),
    Colonne("ID prévision", SAISIE, FMT_TEXTE, 12),
])

CATEGORIES = Table("tbl_Categories", F_PARAM, 5, 5, [
    Colonne("Catégorie", SAISIE, FMT_TEXTE, 24),
    Colonne("Type", SAISIE, FMT_TEXTE, 12, liste="L_TypesCategorie"),
], lignes_libres=0)

REGLES = Table("tbl_Regles", F_PARAM, 5, 8, [
    Colonne("Mot-clé", SAISIE, FMT_TEXTE, 18),
    Colonne("Catégorie", SAISIE, FMT_TEXTE, 22, liste="L_Categories"),
], lignes_libres=0)

HISTORIQUE = Table("tbl_Historique", F_HISTO, 1, 0, [
    Colonne("ID import", SYSTEME, FMT_TEXTE, 11),
    Colonne("Date et heure", SYSTEME, FMT_DATEHEURE, 16),
    Colonne("Banque", SYSTEME, FMT_TEXTE, 16),
    Colonne("Fichier", SYSTEME, FMT_TEXTE, 30),
    Colonne("Lignes lues", SYSTEME, FMT_ENTIER, 10),
    Colonne("Ajoutées", SYSTEME, FMT_ENTIER, 10),
    Colonne("Doublons ignorés", SYSTEME, FMT_ENTIER, 10),
    Colonne("Antérieures au solde initial", SYSTEME, FMT_ENTIER, 12),
    Colonne("Erreur", SYSTEME, FMT_TEXTE, 50),
    Colonne("Utilisateur", SYSTEME, FMT_TEXTE, 14),
])

REJETS = Table("tbl_Rejets", F_LISTES, 1, 14, [
    Colonne("ID opération", SYSTEME, FMT_TEXTE, 12),
    Colonne("ID prévision", SYSTEME, FMT_TEXTE, 12),
    Colonne("Date", SYSTEME, FMT_DATEHEURE, 16),
])

COLONNES_TRANSIT = ["Banque", "Code", "Date", "DateValeur", "Libelle", "LibelleNorm", "Montant",
                    "Reference", "Fichier", "Ligne", "Occurrence", "Cle", "Erreur"]
FMT_TRANSIT = {"Date": FMT_DATE, "DateValeur": FMT_DATE, "Montant": FMT_EUR}


def _table_transit(nom, onglet):
    return Table(nom, onglet, 1, 0, [
        Colonne(n, SYSTEME, FMT_TRANSIT.get(n), 14) for n in COLONNES_TRANSIT
    ], style="Table Style Light 1")


STAGE_B1 = _table_transit("tbl_Stage_B1", F_IMPORT_B1)
STAGE_B2 = _table_transit("tbl_Stage_B2", F_IMPORT_B2)

PARAM_IMPORT = Table("tbl_ParamImport", F_PARAM_IMPORT, 1, 0, [
    Colonne("Clé", SYSTEME, FMT_TEXTE, 16),
    Colonne("Paramètre", SYSTEME, FMT_TEXTE, 34),
    Colonne("Banque 1", SAISIE, FMT_TEXTE, 26),
    Colonne("Banque 2", SAISIE, FMT_TEXTE, 26),
    Colonne("Description", SYSTEME, FMT_TEXTE, 70),
])

CODE_M = Table("tbl_CodeM", F_PARAM_IMPORT, 24, 0, [
    Colonne("Requête", SYSTEME, FMT_TEXTE, 22),
    Colonne("Table de chargement", SYSTEME, FMT_TEXTE, 18),
    Colonne("Feuille", SYSTEME, FMT_TEXTE, 12),
    Colonne("Formule M", SYSTEME, FMT_TEXTE, 120),
])

TABLES = [PREVISIONS, CHARGES, BANQUE, PROPOSITIONS, ATRAITER, CATEGORIES, REGLES, HISTORIQUE,
          REJETS, STAGE_B1, STAGE_B2, PARAM_IMPORT, CODE_M]
TABLES_PAR_NOM = {t.nom: t for t in TABLES}

# ---------------------------------------------------------------------------
# Paramètres (onglet Paramètres) : nom défini, libellé, valeur par défaut, format, aide
# ---------------------------------------------------------------------------
PARAMETRES = [
    ("Société", [
        ("Societe", "Nom de la société", "Société Exemple SAS", None, "Affiché en titre du tableau de bord."),
    ]),
    ("Banque 1", [
        ("B1_Nom", "Nom de la banque", "Banque Alpha", None, "Nom affiché partout dans l'outil."),
        ("B1_IBAN", "IBAN / n° de compte", "FR76 0000 0000 0000 0000 0000 001", None, "Information uniquement."),
        ("B1_SoldeInit", "Solde initial", 25000, FMT_EUR, "Solde du relevé bancaire à la date ci-dessous."),
        ("B1_DateSolde", "Date du solde initial", None, FMT_DATE,
         "Le solde inclut toutes les opérations jusqu'à cette date incluse : elles ne sont pas réimportées."),
    ]),
    ("Banque 2", [
        ("B2_Nom", "Nom de la banque", "Banque Beta", None, "Nom affiché partout dans l'outil."),
        ("B2_IBAN", "IBAN / n° de compte", "FR76 0000 0000 0000 0000 0000 002", None, "Information uniquement."),
        ("B2_SoldeInit", "Solde initial", 8000, FMT_EUR, "Solde du relevé bancaire à la date ci-dessous."),
        ("B2_DateSolde", "Date du solde initial", None, FMT_DATE,
         "Le solde inclut toutes les opérations jusqu'à cette date incluse : elles ne sont pas réimportées."),
    ]),
    ("Rapprochement", [
        ("Tolerance_Jours", "Tolérance de date (jours)", 7, FMT_ENTIER,
         "Écart maximal entre la date prévue et la date bancaire."),
        ("Ecart_Certain", "Écart max. pour une correspondance certaine (jours)", 3, FMT_ENTIER,
         "Une correspondance unique est « certaine » si l'écart est inférieur ou égal..."),
        ("Seuil_Similarite", "... ou si la similarité des libellés atteint", 0.4, FMT_PCT,
         "Sinon elle est « incertaine » et doit être choisie manuellement."),
        ("Auto_Valider", "Valider automatiquement les correspondances certaines", "Non", None,
         "Oui / Non. Les correspondances incertaines ne sont jamais validées automatiquement."),
    ]),
    ("Prévisions", [
        ("Horizon_Mois", "Horizon de génération des charges fixes (mois)", 12, FMT_ENTIER,
         "Les échéances sont générées jusqu'à cette date."),
        ("Date_Ref_Saisie", "Date de calcul (vide = aujourd'hui)", None, FMT_DATE,
         "Laisser vide en utilisation normale. Sert aux tests et démonstrations."),
    ]),
    ("Dossiers et sauvegardes", [
        ("Dossier_Racine", "Dossier racine (vide = dossier du classeur)", None, None,
         "Les relevés sont lus dans <racine>\\Imports\\Banque1 et \\Banque2."),
        ("Dossier_B1", "Dossier des relevés Banque 1 (facultatif)", None, None, "Remplace le dossier par défaut."),
        ("Dossier_B2", "Dossier des relevés Banque 2 (facultatif)", None, None, "Remplace le dossier par défaut."),
        ("Nb_Sauvegardes", "Nombre de sauvegardes conservées", 30, FMT_ENTIER, "0 = toutes conservées."),
    ]),
]

LISTES = {
    "L_Types": ["Dépense", "Recette", "Virement sortant", "Virement entrant"],
    "L_TypesForm": ["Dépense", "Recette", "Virement interne"],
    "L_TypesCharge": ["Dépense", "Recette"],
    "L_TypesCategorie": ["Dépense", "Recette", "Les deux"],
    "L_StatutsPrev": ["Prévu", "Annulé"],
    "L_Frequences": ["Mensuelle", "Trimestrielle", "Semestrielle", "Annuelle"],
    "L_OuiNon": ["Oui", "Non"],
    "L_Actions": ["Conserver comme non prévue", "Virement interne", "Rapprocher avec la prévision indiquée",
                  "Créer charge mensuelle", "Créer charge trimestrielle", "Créer charge semestrielle",
                  "Créer charge annuelle"],
    "L_Afficher": ["Tout", "Prévisions uniquement", "Réel uniquement"],
    "L_Horizon": [30, 60, 90],
}

CATEGORIES_DEFAUT = [
    ("Ventes / Clients", "Recette"), ("Subventions", "Recette"), ("Autres recettes", "Recette"),
    ("Salaires", "Dépense"), ("Charges sociales", "Dépense"), ("Loyer", "Dépense"),
    ("Énergie", "Dépense"), ("Télécom / Internet", "Dépense"), ("Assurances", "Dépense"),
    ("Fournisseurs", "Dépense"), ("Fournitures", "Dépense"), ("Informatique", "Dépense"),
    ("Honoraires", "Dépense"), ("Impôts et taxes", "Dépense"), ("TVA", "Dépense"),
    ("Emprunts", "Dépense"), ("Frais bancaires", "Dépense"), ("Véhicules / Carburant", "Dépense"),
    ("Déplacements", "Dépense"), ("Divers", "Les deux"), ("Virement interne", "Les deux"),
    ("Non catégorisé", "Les deux"),
]

REGLES_DEFAUT = [
    ("URSSAF", "Charges sociales"), ("DGFIP", "Impôts et taxes"), ("IMPOTS", "Impôts et taxes"),
    ("EDF", "Énergie"), ("ENGIE", "Énergie"), ("ORANGE", "Télécom / Internet"), ("SFR", "Télécom / Internet"),
    ("FREE", "Télécom / Internet"), ("AXA", "Assurances"), ("MAAF", "Assurances"), ("OVH", "Informatique"),
    ("MICROSOFT", "Informatique"), ("TOTAL", "Véhicules / Carburant"), ("PEAGE", "Véhicules / Carburant"),
    ("SNCF", "Déplacements"), ("COMMISSION", "Frais bancaires"), ("FRAIS", "Frais bancaires"),
    ("COTIS", "Frais bancaires"), ("SALAIRE", "Salaires"), ("LOYER", "Loyer"),
]
