Attribute VB_Name = "modTableauDeBord"
'===============================================================================
' modTableauDeBord - Horizon du graphique (30 / 60 / 90 jours) et navigation
'===============================================================================
Option Explicit

Public Sub Horizon30()
    ChoisirHorizon 30
End Sub

Public Sub Horizon60()
    ChoisirHorizon 60
End Sub

Public Sub Horizon90()
    ChoisirHorizon 90
End Sub

Public Sub ChoisirHorizon(ByVal nbJours As Long)
    Dim msgErr As String
    On Error GoTo Erreur
    DebutTraitement
    EcrireNom "Horizon_Jours", nbJours
    FinTraitement
    Application.Calculate
    AppliquerHorizon
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "le changement d'horizon", msgErr
End Sub

' Redimensionne les séries du graphique sur J+0 .. J+horizon.
' Ordre des séries (défini par le générateur) : 1 = global, 2 = banque 1, 3 = banque 2.
Public Sub AppliquerHorizon()
    Dim ws As Worksheet, ch As Chart, n As Long, noms As Variant, k As Long
    On Error GoTo Fin
    Set ws = ThisWorkbook.Worksheets(F_TDB)
    If ws.ChartObjects.Count = 0 Then Exit Sub
    n = CLng(Nombre(ValeurNom("Horizon_Jours")))
    If n <> 30 And n <> 60 And n <> 90 Then n = 30
    noms = Array("Calc_SoldeGlobal", "Calc_SoldeB1", "Calc_SoldeB2")
    Set ch = ws.ChartObjects(1).Chart
    If ch.SeriesCollection.Count > 0 Then
        If ch.SeriesCollection(1).Points.Count = n + 1 Then Exit Sub   ' déjà à la bonne taille
    End If
    ws.Unprotect MOT_DE_PASSE
    For k = 0 To 2
        If ch.SeriesCollection.Count > k Then
            ch.SeriesCollection(k + 1).Values = ThisWorkbook.Names(CStr(noms(k))).RefersToRange.Resize(n + 1, 1)
            ch.SeriesCollection(k + 1).XValues = ThisWorkbook.Names("Calc_Dates").RefersToRange.Resize(n + 1, 1)
        End If
    Next k
Fin:
    If Not ws Is Nothing Then ProtegerFeuille ws
End Sub

'-------------------------------------------------------------------------------
' Navigation
'-------------------------------------------------------------------------------
Private Sub Aller(ByVal nomFeuille As String)
    On Error Resume Next
    ThisWorkbook.Worksheets(nomFeuille).Activate
    ThisWorkbook.Worksheets(nomFeuille).Range("A1").Select
End Sub

Public Sub AllerTableauDeBord()
    Aller F_TDB
End Sub

Public Sub AllerSaisie()
    Aller F_SAISIE
End Sub

Public Sub AllerCharges()
    Aller F_CHARGES
End Sub

Public Sub AllerSuivi()
    Aller F_SUIVI
End Sub

Public Sub AllerSynthese()
    Aller F_SYNTHESE
End Sub

Public Sub AllerRapprochement()
    Aller F_RAPPRO
End Sub

Public Sub AllerATraiter()
    Aller F_ATRAITER
End Sub

Public Sub AllerBanque()
    Aller F_BANQUE
End Sub

Public Sub AllerParametres()
    Aller F_PARAM
End Sub

Public Sub AllerAide()
    Aller F_AIDE
End Sub

'-------------------------------------------------------------------------------
' Suivi prévisionnel et opérations bancaires
'-------------------------------------------------------------------------------
Public Sub ReinitialiserFiltresSuivi()
    On Error Resume Next
    DebutTraitement
    ThisWorkbook.Names("Suivi_Banque").RefersToRange.Value = "Toutes"
    ThisWorkbook.Names("Suivi_Du").RefersToRange.Formula = "=Date_Reference-30"
    ThisWorkbook.Names("Suivi_Au").RefersToRange.Formula = "=Date_Reference+90"
    ThisWorkbook.Names("Suivi_Afficher").RefersToRange.Value = "Tout"
    FinTraitement
    Application.Calculate
End Sub

Public Sub FiltrerOperationsATraiter()
    Dim lo As ListObject
    On Error Resume Next
    DebutTraitement
    Set lo = T(T_BANQUE)
    lo.Range.AutoFilter Field:=C(Colonnes(lo), "Statut"), Criteria1:=OB_A_TRAITER
    FinTraitement
End Sub

Public Sub AfficherToutesOperations()
    Dim lo As ListObject
    On Error Resume Next
    DebutTraitement
    Set lo = T(T_BANQUE)
    If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    FinTraitement
End Sub
