Attribute VB_Name = "modAdmin"
'===============================================================================
' modAdmin - Fonctions réservées à l'administrateur (Alt+F8)
'-------------------------------------------------------------------------------
'   Admin_AfficherOngletsTechniques / Admin_MasquerOngletsTechniques
'   Admin_DeprotegerFeuilles / Admin_ProtegerFeuilles
'   Admin_ReinstallerRequetes   : recrée les requêtes Power Query depuis tbl_CodeM
'   Admin_VerifierIntegrite     : contrôle de cohérence des données
'   Admin_ViderDonnees          : remise à zéro (mise en service chez le client)
' et procédures appelées automatiquement (ouverture, renommage des banques).
'===============================================================================
Option Explicit
Option Compare Binary

Private Function FeuillesTechniques() As Variant
    FeuillesTechniques = Array(F_IMPORT_B1, F_IMPORT_B2, F_HISTO, F_CALC, F_LISTES, F_PARAM_IMPORT)
End Function

Private Function MotDePasseOk() As Boolean
    Dim saisie As String
    saisie = InputBox("Mot de passe administrateur :", "Administration")
    MotDePasseOk = (saisie = MOT_DE_PASSE)
    If Not MotDePasseOk And saisie <> "" Then Afficher "Mot de passe incorrect.", vbExclamation, "Administration"
End Function

Public Sub Admin_AfficherOngletsTechniques()
    Dim nom As Variant
    If Not MotDePasseOk() Then Exit Sub
    On Error Resume Next
    ThisWorkbook.Unprotect MOT_DE_PASSE
    For Each nom In FeuillesTechniques()
        ThisWorkbook.Worksheets(CStr(nom)).Visible = xlSheetVisible
    Next nom
End Sub

Public Sub Admin_MasquerOngletsTechniques()
    Dim nom As Variant
    On Error Resume Next
    ThisWorkbook.Worksheets(F_TDB).Activate
    For Each nom In FeuillesTechniques()
        ThisWorkbook.Worksheets(CStr(nom)).Visible = xlSheetVeryHidden
    Next nom
End Sub

Public Sub Admin_DeprotegerFeuilles()
    If Not MotDePasseOk() Then Exit Sub
    DeprotegerTout
    Afficher "Feuilles déprotégées. Pensez à lancer Admin_ProtegerFeuilles après vos modifications.", vbInformation, "Administration"
End Sub

Public Sub Admin_ProtegerFeuilles()
    ProtegerTout
End Sub

'-------------------------------------------------------------------------------
' Requêtes Power Query
'-------------------------------------------------------------------------------
Public Function RequetesInstallees() As Boolean
    Dim q As Object
    On Error Resume Next
    Set q = ThisWorkbook.Queries(Q_IMPORT_B1)
    If q Is Nothing Then Exit Function
    Set q = Nothing
    Set q = ThisWorkbook.Queries(Q_IMPORT_B2)
    RequetesInstallees = Not (q Is Nothing)
End Function

Public Sub Admin_ReinstallerRequetes()
    Dim msgErr As String
    On Error GoTo Erreur
    If Afficher("Recréer les requêtes Power Query à partir du code M stocké dans l'onglet Param_Import ?" & vbCrLf & _
              "(les données déjà importées dans le journal bancaire ne sont pas touchées)", _
              vbOKCancel + vbQuestion, "Administration") <> vbOK Then Exit Sub
    DebutTraitement "Installation des requêtes..."
    InstallerRequetes
    modImport.MettreAJourParametresPQ
    FinTraitement
    Afficher "Requêtes Power Query installées. Lancez « Actualiser les banques » depuis le tableau de bord.", vbInformation, "Administration"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "l'installation des requêtes", msgErr
End Sub

Private Sub InstallerRequetes()
    Dim lo As ListObject, cM As Object, dM As Variant, i As Long, nomQ As String, formule As String
    Dim q As Object, cible As String, nomFeuille As String
    Set lo = T(T_CODE_M)
    Set cM = Colonnes(lo)
    dM = LireTable(lo)
    For i = 1 To NbLignes(dM)
        nomQ = Texte(dM(i, C(cM, "Requête")))
        formule = CStr(dM(i, C(cM, "Formule M")))
        If nomQ <> "" And formule <> "" Then
            Set q = Nothing
            On Error Resume Next
            Set q = ThisWorkbook.Queries(nomQ)
            On Error GoTo 0
            If q Is Nothing Then
                ThisWorkbook.Queries.Add nomQ, formule
            ElseIf q.Formula <> formule Then
                q.Formula = formule
            End If
        End If
    Next i
    For i = 1 To NbLignes(dM)
        nomQ = Texte(dM(i, C(cM, "Requête")))
        cible = Texte(dM(i, C(cM, "Table de chargement")))
        nomFeuille = Texte(dM(i, C(cM, "Feuille")))
        If nomQ <> "" And cible <> "" And nomFeuille <> "" Then ChargerRequete nomQ, cible, nomFeuille
    Next i
End Sub

' Relie (ou relie de nouveau) une table de transit à sa requête Power Query
Private Sub ChargerRequete(ByVal nomQ As String, ByVal nomTable As String, ByVal nomFeuille As String)
    Dim ws As Worksheet, lo As ListObject, qt As QueryTable, existante As ListObject
    Set ws = Feuille(nomFeuille)
    For Each lo In ws.ListObjects
        If StrComp(lo.Name, nomTable, vbTextCompare) = 0 Then Set existante = lo
    Next lo
    If Not existante Is Nothing Then
        On Error Resume Next
        Set qt = existante.QueryTable
        On Error GoTo 0
        If Not qt Is Nothing Then Exit Sub
        existante.Delete
    End If
    ws.Cells.Clear
    Set lo = ws.ListObjects.Add(SourceType:=0, Source:=Array( _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=" & nomQ & ";Extended Properties="""""), _
        Destination:=ws.Range("A1"))
    lo.Name = nomTable
    With lo.QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & nomQ & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With
End Sub

'-------------------------------------------------------------------------------
' Ouverture du classeur
'-------------------------------------------------------------------------------
Public Sub VerifierInstallation()
    On Error Resume Next
    If Not RequetesInstallees() Then
        If Afficher("Les requêtes Power Query d'import bancaire sont absentes de ce classeur." & vbCrLf & _
                  "Les installer maintenant ?", vbYesNo + vbQuestion, "Trésorerie") = vbYes Then
            DebutTraitement "Installation des requêtes..."
            InstallerRequetes
            modImport.MettreAJourParametresPQ
            FinTraitement
        End If
    End If
End Sub

' Réassocie chaque bouton (forme nommée « btn_NomDeMacro ») à sa macro
Public Sub AssignerBoutons()
    Dim nom As Variant, ws As Worksheet, sh As Shape, macro As String
    On Error Resume Next
    For Each nom In FeuillesProtegees()
        Set ws = ThisWorkbook.Worksheets(CStr(nom))
        For Each sh In ws.Shapes
            If Left$(sh.Name, 4) = "btn_" Then
                macro = Mid$(sh.Name, 5)
                If InStr(1, sh.OnAction, macro, vbTextCompare) = 0 Then sh.OnAction = macro
            End If
        Next sh
    Next nom
End Sub

'-------------------------------------------------------------------------------
' Renommage des banques : répercute un changement de nom dans toutes les tables
'-------------------------------------------------------------------------------
Public Sub VerifierRenommageBanques()
    Dim code As Variant, actuel As String, precedent As String
    For Each code In Array("B1", "B2")
        actuel = NomBanque(CStr(code))
        precedent = Texte(ValeurNom(CStr(code) & "_NomPrecedent"))
        If precedent = "" Then
            EcrireNom CStr(code) & "_NomPrecedent", actuel
        ElseIf actuel <> precedent And actuel <> "" Then
            If Afficher("La banque « " & precedent & " » a été renommée en « " & actuel & " »." & vbCrLf & _
                      "Mettre à jour toutes les opérations, prévisions et charges existantes ?", _
                      vbYesNo + vbQuestion, "Renommage de banque") = vbYes Then
                RemplacerBanque precedent, actuel
                EcrireNom CStr(code) & "_NomPrecedent", actuel
            End If
        End If
    Next code
End Sub

Private Sub RemplacerBanque(ByVal ancien As String, ByVal nouveau As String)
    Dim nomTable As Variant, lo As ListObject, cols As Object, data As Variant, i As Long, n As Long
    For Each nomTable In Array(T_PREV, T_CHARGES, T_BANQUE, T_PROPOS, T_ATRAITER)
        Set lo = T(CStr(nomTable))
        Set cols = Colonnes(lo)
        data = LireTable(lo)
        n = NbLignes(data)
        If n > 0 And cols.Exists("Banque") Then
            For i = 1 To n
                If Texte(data(i, cols("Banque"))) = ancien Then data(i, cols("Banque")) = nouveau
            Next i
            EcrireTable lo, data, n
        End If
    Next nomTable
    ' Formulaire de saisie
    If Texte(ValeurNom("Form_Banque")) = ancien Then EcrireNom "Form_Banque", nouveau
    If Texte(ValeurNom("Form_BanqueDest")) = ancien Then EcrireNom "Form_BanqueDest", nouveau
    If Texte(ValeurNom("Suivi_Banque")) = ancien Then EcrireNom "Suivi_Banque", nouveau
End Sub

'-------------------------------------------------------------------------------
' Contrôle d'intégrité
'-------------------------------------------------------------------------------
Public Sub Admin_VerifierIntegrite()
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long
    Dim idsB As Object, idsP As Object, clesB As Object, i As Long, id As String, lien As String
    Dim rapport As String, nbPb As Long

    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    Set loP = T(T_PREV)
    Set cP = Colonnes(loP)
    dP = LireTable(loP)
    nP = NbLignes(dP)
    Set idsB = CreateObject("Scripting.Dictionary")
    Set idsP = CreateObject("Scripting.Dictionary")
    Set clesB = CreateObject("Scripting.Dictionary")

    For i = 1 To nB
        id = Texte(dB(i, C(cB, "ID")))
        If id <> "" Then
            If idsB.Exists(id) Then Ajouter rapport, nbPb, "Identifiant d'opération en double : " & id
            idsB(id) = i
            If clesB.Exists(Texte(dB(i, C(cB, "Clé")))) Then Ajouter rapport, nbPb, "Clé technique en double : " & id
            clesB(Texte(dB(i, C(cB, "Clé")))) = True
        End If
    Next i
    For i = 1 To nP
        id = Texte(dP(i, C(cP, "ID")))
        If id <> "" Then
            If idsP.Exists(id) Then Ajouter rapport, nbPb, "Identifiant de prévision en double : " & id
            idsP(id) = i
        End If
    Next i
    For i = 1 To nP
        id = Texte(dP(i, C(cP, "ID")))
        lien = Texte(dP(i, C(cP, "ID opération")))
        If Texte(dP(i, C(cP, "Statut"))) = P_POINTE Then
            If lien = "" Or Not idsB.Exists(lien) Then
                Ajouter rapport, nbPb, "Prévision " & id & " pointée sans opération bancaire valide"
            ElseIf Texte(dB(idsB(lien), C(cB, "ID prévision"))) <> id Then
                Ajouter rapport, nbPb, "Lien incohérent entre la prévision " & id & " et l'opération " & lien
            End If
        End If
        lien = Texte(dP(i, C(cP, "Lien virement")))
        If lien <> "" And Not idsP.Exists(lien) Then Ajouter rapport, nbPb, "Virement " & id & " sans contrepartie (" & lien & ")"
    Next i
    For i = 1 To nB
        id = Texte(dB(i, C(cB, "ID")))
        lien = Texte(dB(i, C(cB, "ID prévision")))
        If Texte(dB(i, C(cB, "Statut"))) = OB_POINTE Then
            If lien = "" Or Not idsP.Exists(lien) Then
                Ajouter rapport, nbPb, "Opération " & id & " pointée sans prévision valide"
            ElseIf Texte(dP(idsP(lien), C(cP, "ID opération"))) <> id Then
                Ajouter rapport, nbPb, "Lien incohérent entre l'opération " & id & " et la prévision " & lien
            End If
        End If
    Next i
    If nbPb = 0 Then
        Afficher "Aucune incohérence détectée (" & idsB.Count & " opérations bancaires, " & idsP.Count & " prévisions).", _
               vbInformation, "Contrôle d'intégrité"
    Else
        Afficher nbPb & " incohérence(s) détectée(s) :" & vbCrLf & vbCrLf & rapport, vbExclamation, "Contrôle d'intégrité"
    End If
End Sub

Private Sub Ajouter(ByRef rapport As String, ByRef nb As Long, ByVal ligne As String)
    nb = nb + 1
    If nb <= 30 Then rapport = rapport & "- " & ligne & vbCrLf
    If nb = 31 Then rapport = rapport & "- ..." & vbCrLf
End Sub

'-------------------------------------------------------------------------------
' Remise à zéro (mise en service chez le client)
'-------------------------------------------------------------------------------
Public Sub Admin_ViderDonnees()
    Dim saisie As String, nomTable As Variant, lo As ListObject, vide() As Variant, msgErr As String, sauvegarde As String
    On Error GoTo Erreur
    If Not MotDePasseOk() Then Exit Sub
    saisie = InputBox("ATTENTION : toutes les opérations bancaires importées, prévisions, charges fixes," & vbCrLf & _
                      "propositions et l'historique des imports vont être supprimés." & vbCrLf & _
                      "Les paramètres sont conservés et une sauvegarde est faite avant." & vbCrLf & vbCrLf & _
                      "Tapez VIDER pour confirmer :", "Remise à zéro")
    If saisie <> "VIDER" Then Exit Sub
    DebutTraitement "Remise à zéro..."
    sauvegarde = modImport.SauvegarderClasseur()
    For Each nomTable In Array(T_BANQUE, T_PROPOS, T_ATRAITER, T_HISTO, T_REJETS)
        Set lo = T(CStr(nomTable))
        EcrireTable lo, Empty, 0
    Next nomTable
    For Each nomTable In Array(T_PREV, T_CHARGES)
        Set lo = T(CStr(nomTable))
        ReDim vide(1 To NB_LIGNES_LIBRES, 1 To lo.ListColumns.Count)
        EcrireTable lo, vide, NB_LIGNES_LIBRES
    Next nomTable
    FinTraitement
    Application.Calculate
    Afficher "Données vidées. Sauvegarde préalable : " & sauvegarde & vbCrLf & vbCrLf & _
           "Renseignez maintenant les soldes initiaux et leur date dans l'onglet Paramètres.", vbInformation, "Remise à zéro"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la remise à zéro", msgErr
End Sub
