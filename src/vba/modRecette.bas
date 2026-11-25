Attribute VB_Name = "modRecette"
'===============================================================================
' modRecette - Recette automatique (appelée par install/Recette.ps1)
'-------------------------------------------------------------------------------
' Exécute, sans aucune boîte de dialogue, la chaîne complète sur les relevés de
' démonstration : actualisation (Power Query + import + charges fixes +
' rapprochement) puis validation des correspondances certaines.
' Renvoie les indicateurs clés sous la forme « Nom=valeur;Nom=valeur;... » afin que
' le script PowerShell les compare aux valeurs attendues calculées par le moteur de
' référence (dist/recette_attendue.json).
' A n'utiliser que sur une COPIE du classeur de démonstration.
'===============================================================================
Option Explicit

Private Function Nb(ByVal v As Variant) As String
    Nb = Trim$(Str$(Round(Nombre(v), 2)))
End Function

Public Function Recette_Executer(ByVal dateReference As Double) As String
    Dim res As String, noms As Variant, k As Long, lo As ListObject
    On Error GoTo Erreur
    gSilencieux = True
    gJournal = ""
    gDerniereErreur = ""

    DebutTraitement
    EcrireNom "Date_Ref_Saisie", dateReference
    FinTraitement
    Application.Calculate

    modImport.Actualiser
    If gDerniereErreur = "" Then modRapprochement.ToutValiderCertaines
    Application.Calculate

    res = "Erreur=" & Replace(Replace(gDerniereErreur, ";", ","), "=", ":")
    noms = Array("SoldeReel_B1", "SoldeReel_B2", "SoldeReel_Total", "Prev_B1_30", "Prev_B2_30", "Prev_Total_30", _
                 "Prev_B1_90", "Prev_B2_90", "Prev_Total_90", "Depenses_A_Venir", "Recettes_Attendues", _
                 "NbATraiter", "NbPropositions", "NbOpsProposees", "NbEchues")
    For k = LBound(noms) To UBound(noms)
        res = res & ";" & noms(k) & "=" & Nb(ValeurNom(CStr(noms(k))))
    Next k
    Set lo = T(T_BANQUE)
    res = res & ";NbOperations=" & CStr(Application.WorksheetFunction.CountA(lo.ListColumns("ID").DataBodyRange))
    res = res & ";NbPointees=" & CStr(Application.WorksheetFunction.CountIf(lo.ListColumns("Statut").DataBodyRange, OB_POINTE))
    Set lo = T(T_HISTO)
    res = res & ";NbFichiersHistorique=" & CStr(Application.WorksheetFunction.CountA(lo.ListColumns("Fichier").DataBodyRange))
    gSilencieux = False
    Recette_Executer = res
    Exit Function
Erreur:
    gSilencieux = False
    TerminerTraitement
    Recette_Executer = "Erreur=" & Replace(Err.Description, ";", ",")
End Function

' Second passage : réimporter les mêmes relevés ne doit rien ajouter (doublons)
Public Function Recette_Reimporter() As String
    Dim lo As ListObject, avant As Long, apres As Long
    On Error GoTo Erreur
    gSilencieux = True
    gDerniereErreur = ""
    Set lo = T(T_BANQUE)
    avant = Application.WorksheetFunction.CountA(lo.ListColumns("ID").DataBodyRange)
    modImport.Actualiser
    Application.Calculate
    apres = Application.WorksheetFunction.CountA(lo.ListColumns("ID").DataBodyRange)
    gSilencieux = False
    Recette_Reimporter = "Erreur=" & Replace(gDerniereErreur, ";", ",") & ";Avant=" & avant & ";Apres=" & apres & _
                         ";SoldeReel_Total=" & Nb(ValeurNom("SoldeReel_Total"))
    Exit Function
Erreur:
    gSilencieux = False
    TerminerTraitement
    Recette_Reimporter = "Erreur=" & Replace(Err.Description, ";", ",")
End Function

Public Function Recette_Journal() As String
    Recette_Journal = gJournal
End Function
