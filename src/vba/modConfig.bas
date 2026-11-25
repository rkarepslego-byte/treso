Attribute VB_Name = "modConfig"
'===============================================================================
' modConfig - Constantes de l'outil de trésorerie
'-------------------------------------------------------------------------------
' Toutes les chaînes (noms d'onglets, de tables, statuts...) sont centralisées ici.
' Elles doivent rester identiques à celles du générateur (build/schema.py) et du
' moteur de référence (reference/moteur.py).
'===============================================================================
Option Explicit

Public Const VERSION_OUTIL As String = "1.0.0"

' Mot de passe de protection des feuilles et des fonctions d'administration.
' A personnaliser pour chaque client (voir docs/03_Guide_technique.md).
Public Const MOT_DE_PASSE As String = "Treso2026"

' --- Onglets visibles ----------------------------------------------------------
Public Const F_TDB As String = "Tableau de bord"
Public Const F_SAISIE As String = "Saisie"
Public Const F_CHARGES As String = "Charges fixes"
Public Const F_SUIVI As String = "Suivi prévisionnel"
Public Const F_SYNTHESE As String = "Synthèse mensuelle"
Public Const F_RAPPRO As String = "Rapprochement"
Public Const F_ATRAITER As String = "À traiter"
Public Const F_BANQUE As String = "Opérations bancaires"
Public Const F_PARAM As String = "Paramètres"
Public Const F_AIDE As String = "Aide"

' --- Onglets techniques (masqués) ---------------------------------------------
Public Const F_IMPORT_B1 As String = "Import_B1"
Public Const F_IMPORT_B2 As String = "Import_B2"
Public Const F_HISTO As String = "Historique"
Public Const F_CALC As String = "Calculs"
Public Const F_LISTES As String = "Listes"
Public Const F_PARAM_IMPORT As String = "Param_Import"

' --- Tables --------------------------------------------------------------------
Public Const T_PREV As String = "tbl_Previsions"
Public Const T_CHARGES As String = "tbl_Charges"
Public Const T_BANQUE As String = "tbl_Banque"
Public Const T_PROPOS As String = "tbl_Propositions"
Public Const T_ATRAITER As String = "tbl_ATraiter"
Public Const T_CATEG As String = "tbl_Categories"
Public Const T_REGLES As String = "tbl_Regles"
Public Const T_HISTO As String = "tbl_Historique"
Public Const T_REJETS As String = "tbl_Rejets"
Public Const T_PARAM_IMPORT As String = "tbl_ParamImport"
Public Const T_CODE_M As String = "tbl_CodeM"
Public Const T_STAGE_B1 As String = "tbl_Stage_B1"
Public Const T_STAGE_B2 As String = "tbl_Stage_B2"

' --- Requêtes Power Query -----------------------------------------------------
Public Const Q_PARAMETRES As String = "Parametres_Import"
Public Const Q_IMPORT_B1 As String = "Import_Banque1"
Public Const Q_IMPORT_B2 As String = "Import_Banque2"

' --- Statuts des prévisions ----------------------------------------------------
Public Const P_PREVU As String = "Prévu"
Public Const P_POINTE As String = "Pointé"
Public Const P_ANNULE As String = "Annulé"
Public Const P_A_CONTROLER As String = "À contrôler"

' --- Statuts des opérations bancaires -----------------------------------------
Public Const OB_A_TRAITER As String = "À traiter"
Public Const OB_PROPOSE As String = "Proposé"
Public Const OB_POINTE As String = "Pointé"
Public Const OB_NON_PREVUE As String = "Non prévue"
Public Const OB_VIREMENT As String = "Virement interne"

' --- Types d'opération ---------------------------------------------------------
Public Const TY_DEPENSE As String = "Dépense"
Public Const TY_RECETTE As String = "Recette"
Public Const TY_VIR_SORTANT As String = "Virement sortant"
Public Const TY_VIR_ENTRANT As String = "Virement entrant"
Public Const TY_VIR_FORM As String = "Virement interne"

' --- Origines, niveaux, modes --------------------------------------------------
Public Const ORI_SAISIE As String = "Saisie"
Public Const ORI_CHARGE As String = "Charge fixe"
Public Const NIV_CERTAINE As String = "Certaine"
Public Const NIV_INCERTAINE As String = "Incertaine"
Public Const MODE_VALIDE As String = "Validé"
Public Const MODE_AUTO As String = "Automatique"
Public Const MODE_MANUEL As String = "Manuel"

' --- Actions de la liste « À traiter » -----------------------------------------
Public Const ACT_CONSERVER As String = "Conserver comme non prévue"
Public Const ACT_VIREMENT As String = "Virement interne"
Public Const ACT_RAPPROCHER As String = "Rapprocher avec la prévision indiquée"
Public Const ACT_CHARGE_PREFIXE As String = "Créer charge "

' --- Divers --------------------------------------------------------------------
Public Const CAT_NON_CATEGORISE As String = "Non catégorisé"
Public Const CAT_VIREMENT As String = "Virement interne"
Public Const NB_LIGNES_LIBRES As Long = 30
Public Const SUGGERE As String = "Suggérée"

' Mots ignorés pour la similarité des libellés (identique à reference/moteur.py)
Public Const MOTS_VIDES As String = "|PRLV|PRELEVEMENT|SEPA|VIR|VIREMENT|CB|CARTE|PAIEMENT|FACTURE|FACT|ECHEANCE|ECH|REF|EUR|DE|DU|DES|LA|LE|LES|ET|AU|AUX|EN|SA|SAS|SARL|EURL|INST|RECU|EMIS|FAVEUR|ORDRE|MOTIF|NUM|NO|"
