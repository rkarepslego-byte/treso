Attribute VB_Name = "modOutils"
'===============================================================================
' modOutils - Fonctions utilitaires communes
'-------------------------------------------------------------------------------
' * accès aux tables structurées (lecture / réécriture en bloc)
' * protection des feuilles et encadrement des traitements
' * normalisation des libellés et similarité (identique à reference/moteur.py)
' * dossiers d'import et de sauvegarde
'===============================================================================
Option Explicit
Option Compare Binary

Private mProfondeur As Long

'-------------------------------------------------------------------------------
' Accès aux objets
'-------------------------------------------------------------------------------
Public Function T(ByVal nomTable As String) As ListObject
    Dim ws As Worksheet, lo As ListObject
    For Each ws In ThisWorkbook.Worksheets
        For Each lo In ws.ListObjects
            If StrComp(lo.Name, nomTable, vbTextCompare) = 0 Then
                Set T = lo
                Exit Function
            End If
        Next lo
    Next ws
    Err.Raise vbObjectError + 1000, "modOutils.T", "La table « " & nomTable & " » est introuvable dans le classeur."
End Function

Public Function Feuille(ByVal nom As String) As Worksheet
    On Error Resume Next
    Set Feuille = ThisWorkbook.Worksheets(nom)
    On Error GoTo 0
    If Feuille Is Nothing Then
        Err.Raise vbObjectError + 1000, "modOutils.Feuille", "L'onglet « " & nom & " » est introuvable."
    End If
End Function

Public Function ValeurNom(ByVal nom As String) As Variant
    Dim rg As Range
    On Error Resume Next
    Set rg = ThisWorkbook.Names(nom).RefersToRange
    On Error GoTo 0
    If rg Is Nothing Then
        Err.Raise vbObjectError + 1000, "modOutils.ValeurNom", "Le nom défini « " & nom & " » est introuvable."
    End If
    ValeurNom = rg.Cells(1, 1).Value2
End Function

Public Sub EcrireNom(ByVal nom As String, ByVal valeur As Variant)
    ThisWorkbook.Names(nom).RefersToRange.Cells(1, 1).Value = valeur
End Sub

Public Function DateReference() As Double
    DateReference = DateVal(ValeurNom("Date_Reference"))
    If DateReference = 0 Then DateReference = CDbl(Date)
End Function

'-------------------------------------------------------------------------------
' Conversions tolérantes
'-------------------------------------------------------------------------------
Public Function Texte(ByVal v As Variant) As String
    If IsError(v) Then
        Texte = ""
    ElseIf IsEmpty(v) Or IsNull(v) Then
        Texte = ""
    Else
        Texte = Trim$(CStr(v))
    End If
End Function

Public Function Nombre(ByVal v As Variant) As Double
    If IsError(v) Then
        Nombre = 0
    ElseIf IsEmpty(v) Or IsNull(v) Then
        Nombre = 0
    ElseIf IsNumeric(v) Then
        Nombre = CDbl(v)
    Else
        Nombre = 0
    End If
End Function

Public Function DateVal(ByVal v As Variant) As Double
    If IsError(v) Then
        DateVal = 0
    ElseIf IsEmpty(v) Or IsNull(v) Then
        DateVal = 0
    ElseIf IsNumeric(v) Then
        DateVal = Int(CDbl(v))
    ElseIf IsDate(v) Then
        DateVal = Int(CDbl(CDate(v)))
    Else
        DateVal = 0
    End If
End Function

Public Function EstOui(ByVal v As Variant) As Boolean
    EstOui = (UCase$(Texte(v)) = "OUI")
End Function

Public Function DateISO(ByVal d As Double) As String
    DateISO = Format$(CDate(d), "yyyy-mm-dd")
End Function

Public Function Cents(ByVal m As Double) As Double
    Dim v As Double
    v = Int(Abs(m) * 100# + 0.5)
    If m < 0 Then Cents = -v Else Cents = v
End Function

Public Function CentsTexte(ByVal m As Double) As String
    CentsTexte = Format$(Cents(m), "0")
End Function

Public Function AjouterMois(ByVal d As Double, ByVal n As Long) As Double
    AjouterMois = CDbl(Application.WorksheetFunction.EDate(d, n))
End Function

Public Function FormatId(ByVal prefixe As String, ByVal n As Long, ByVal largeur As Long) As String
    FormatId = prefixe & Format$(n, String$(largeur, "0"))
End Function

' Plus grand numéro utilisé dans une colonne d'identifiants (ex. "P-000123" -> 123)
Public Function MaxNumero(data As Variant, ByVal nb As Long, ByVal colonne As Long, ByVal prefixe As String) As Long
    Dim i As Long, s As String, v As Long
    For i = 1 To nb
        s = Texte(data(i, colonne))
        If Left$(s, Len(prefixe)) = prefixe Then
            If IsNumeric(Mid$(s, Len(prefixe) + 1)) Then
                v = CLng(Mid$(s, Len(prefixe) + 1))
                If v > MaxNumero Then MaxNumero = v
            End If
        End If
    Next i
End Function

'-------------------------------------------------------------------------------
' Tables : lecture et réécriture en bloc
'-------------------------------------------------------------------------------
' Dictionnaire en-tête -> numéro de colonne
Public Function Colonnes(lo As ListObject) As Object
    Dim d As Object, lc As ListColumn
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = vbTextCompare
    For Each lc In lo.ListColumns
        d(lc.Name) = lc.Index
    Next lc
    Set Colonnes = d
End Function

Public Function C(cols As Object, ByVal nom As String) As Long
    If Not cols.Exists(nom) Then
        Err.Raise vbObjectError + 1002, "modOutils.C", "Colonne « " & nom & " » introuvable : la structure du classeur a été modifiée."
    End If
    C = cols(nom)
End Function

' Renvoie les valeurs de la table (1..n, 1..nbCol) ou Empty si la table est vide
Public Function LireTable(lo As ListObject) As Variant
    Dim v As Variant, a() As Variant
    If lo.ListRows.Count = 0 Then
        LireTable = Empty
        Exit Function
    End If
    v = lo.DataBodyRange.Value2
    If Not IsArray(v) Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = v
        v = a
    End If
    LireTable = v
End Function

Public Function NbLignes(data As Variant) As Long
    If IsArray(data) Then NbLignes = UBound(data, 1) Else NbLignes = 0
End Function

' Ligne vide = aucune des colonnes indiquées n'est renseignée
Public Function LigneVide(data As Variant, ByVal i As Long, ParamArray colonnesTest() As Variant) As Boolean
    Dim k As Long
    For k = LBound(colonnesTest) To UBound(colonnesTest)
        If Texte(data(i, colonnesTest(k))) <> "" Then
            LigneVide = False
            Exit Function
        End If
    Next k
    LigneVide = True
End Function

' Réécrit entièrement le contenu de la table avec les n premières lignes de data.
' Les colonnes calculées (formules) sont conservées et recopiées sur toute la hauteur.
' La table garde toujours au moins une ligne.
Public Sub EcrireTable(lo As ListObject, data As Variant, ByVal n As Long)
    Dim nbCol As Long, j As Long, i As Long, nAncien As Long, nCible As Long
    Dim estFormule() As Boolean, formules() As String, col() As Variant

    nbCol = lo.ListColumns.Count
    ReDim estFormule(1 To nbCol)
    ReDim formules(1 To nbCol)

    On Error Resume Next
    If lo.ShowAutoFilter Then
        If lo.AutoFilter.FilterMode Then lo.AutoFilter.ShowAllData
    End If
    On Error GoTo 0

    nAncien = lo.ListRows.Count
    If nAncien > 0 Then
        For j = 1 To nbCol
            If lo.DataBodyRange.Cells(1, j).HasFormula Then
                estFormule(j) = True
                formules(j) = lo.DataBodyRange.Cells(1, j).FormulaR1C1
            End If
        Next j
    End If

    nCible = n
    If nCible < 1 Then nCible = 1
    If nAncien > nCible Then
        lo.DataBodyRange.Rows(nCible + 1).Resize(nAncien - nCible).Delete xlShiftUp
    ElseIf nAncien < nCible Then
        lo.Resize lo.Range.Resize(nCible + 1, nbCol)
    End If

    For j = 1 To nbCol
        If Not estFormule(j) Then
            ReDim col(1 To nCible, 1 To 1)
            For i = 1 To nCible
                If i <= n Then col(i, 1) = data(i, j) Else col(i, 1) = Empty
            Next i
            lo.ListColumns(j).DataBodyRange.Value2 = col
        End If
    Next j
    For j = 1 To nbCol
        If estFormule(j) Then lo.ListColumns(j).DataBodyRange.FormulaR1C1 = formules(j)
    Next j
End Sub

' Copie la ligne i de source vers la ligne k de cible (mêmes colonnes)
Public Sub CopierLigne(source As Variant, ByVal i As Long, cible As Variant, ByVal k As Long)
    Dim j As Long
    For j = LBound(source, 2) To UBound(source, 2)
        cible(k, j) = source(i, j)
    Next j
End Sub

' Ajoute des lignes à la fin d'une table sans ligne de saisie libre (journal, historique...)
Public Sub AjouterLignes(lo As ListObject, nouvelles As Variant, ByVal nNouv As Long, ByVal colonneCle As Long)
    Dim data As Variant, n As Long, i As Long, k As Long, total() As Variant, nbCol As Long
    If nNouv = 0 Then Exit Sub
    data = LireTable(lo)
    n = NbLignes(data)
    nbCol = lo.ListColumns.Count
    ReDim total(1 To n + nNouv, 1 To nbCol)
    For i = 1 To n
        If Texte(data(i, colonneCle)) <> "" Then
            k = k + 1
            CopierLigne data, i, total, k
        End If
    Next i
    For i = 1 To nNouv
        k = k + 1
        CopierLigne nouvelles, i, total, k
    Next i
    EcrireTable lo, total, k
End Sub

'-------------------------------------------------------------------------------
' Tri d'indices selon des clés texte (tri rapide, comparaison binaire)
'-------------------------------------------------------------------------------
Public Sub TrierIndices(cles() As String, idx() As Long, ByVal bas As Long, ByVal haut As Long)
    Dim i As Long, j As Long, pivot As String, tmp As Long
    If bas >= haut Then Exit Sub
    i = bas
    j = haut
    pivot = cles(idx((bas + haut) \ 2))
    Do While i <= j
        Do While StrComp(cles(idx(i)), pivot, vbBinaryCompare) < 0
            i = i + 1
        Loop
        Do While StrComp(cles(idx(j)), pivot, vbBinaryCompare) > 0
            j = j - 1
        Loop
        If i <= j Then
            tmp = idx(i)
            idx(i) = idx(j)
            idx(j) = tmp
            i = i + 1
            j = j - 1
        End If
    Loop
    If bas < j Then TrierIndices cles, idx, bas, j
    If i < haut Then TrierIndices cles, idx, i, haut
End Sub

Public Function CleDate(ByVal d As Double) As String
    CleDate = Format$(CLng(d), "000000")
End Function

'-------------------------------------------------------------------------------
' Libellés : normalisation et similarité (miroir de reference/moteur.py)
'-------------------------------------------------------------------------------
Public Function SansAccents(ByVal s As String) As String
    Const SRC As String = "ÀÁÂÃÄÅàáâãäåÇçÈÉÊËèéêëÌÍÎÏìíîïÑñÒÓÔÕÖòóôõöÙÚÛÜùúûüÝýÿ"
    Const DST As String = "AAAAAAaaaaaaCcEEEEeeeeIIIIiiiiNnOOOOOoooooUUUUuuuuYyy"
    Dim i As Long, p As Long, ch As String, r As String
    s = Replace(s, "Œ", "OE")
    s = Replace(s, "œ", "oe")
    s = Replace(s, "Æ", "AE")
    s = Replace(s, "æ", "ae")
    s = Replace(s, "ß", "ss")
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        p = InStr(1, SRC, ch, vbBinaryCompare)
        If p > 0 Then r = r & Mid$(DST, p, 1) Else r = r & ch
    Next i
    SansAccents = r
End Function

Public Function NormaliserTexte(ByVal s As String) As String
    Dim i As Long, ch As String, r As String
    s = UCase$(SansAccents(s))
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            r = r & ch
        Else
            r = r & " "
        End If
    Next i
    NormaliserTexte = Application.WorksheetFunction.Trim(r)
End Function

Private Function MotsSignificatifs(ByVal libelle As String, mots() As String) As Long
    Dim brut() As String, i As Long, n As Long, m As String, norm As String
    norm = NormaliserTexte(libelle)
    If norm = "" Then
        MotsSignificatifs = 0
        Exit Function
    End If
    brut = Split(norm, " ")
    ReDim mots(0 To UBound(brut))
    For i = 0 To UBound(brut)
        m = brut(i)
        If Len(m) >= 2 Then
            If Not (m Like String$(Len(m), "#")) Then
                If InStr(1, MOTS_VIDES, "|" & m & "|", vbBinaryCompare) = 0 Then
                    mots(n) = m
                    n = n + 1
                End If
            End If
        End If
    Next i
    MotsSignificatifs = n
End Function

Private Function MotsEgaux(ByVal a As String, ByVal b As String) As Boolean
    If a = b Then
        MotsEgaux = True
    ElseIf Len(a) >= 4 And Len(b) >= 4 Then
        MotsEgaux = (Left$(a, Len(b)) = b) Or (Left$(b, Len(a)) = a)
    Else
        MotsEgaux = False
    End If
End Function

' Similarité de deux libellés entre 0 et 1 :
' max(Dice sur les bigrammes des mots significatifs, part des mots du plus court retrouvés)
Public Function Similarite(ByVal a As String, ByVal b As String) As Double
    Dim ma() As String, mb() As String, na As Long, nb As Long
    Dim compte As Object, i As Long, k As Long, g As String
    Dim totA As Long, totB As Long, commun As Long, dice As Double, inclusion As Double
    Dim mCourt() As String, mLong() As String, nc As Long, nl As Long, j As Long, trouves As Long

    na = MotsSignificatifs(a, ma)
    nb = MotsSignificatifs(b, mb)
    If na = 0 Or nb = 0 Then
        Similarite = 0
        Exit Function
    End If

    Set compte = CreateObject("Scripting.Dictionary")
    For i = 0 To nb - 1
        For k = 1 To Len(mb(i)) - 1
            g = Mid$(mb(i), k, 2)
            compte(g) = compte(g) + 1
            totB = totB + 1
        Next k
    Next i
    For i = 0 To na - 1
        For k = 1 To Len(ma(i)) - 1
            g = Mid$(ma(i), k, 2)
            totA = totA + 1
            If compte.Exists(g) Then
                If compte(g) > 0 Then
                    commun = commun + 1
                    compte(g) = compte(g) - 1
                End If
            End If
        Next k
    Next i
    If totA > 0 And totB > 0 Then dice = 2# * commun / (totA + totB)

    If na <= nb Then
        mCourt = ma: nc = na: mLong = mb: nl = nb
    Else
        mCourt = mb: nc = nb: mLong = ma: nl = na
    End If
    For i = 0 To nc - 1
        For j = 0 To nl - 1
            If MotsEgaux(mCourt(i), mLong(j)) Then
                trouves = trouves + 1
                Exit For
            End If
        Next j
    Next i
    inclusion = trouves / nc

    If dice > inclusion Then Similarite = Round(dice, 4) Else Similarite = Round(inclusion, 4)
End Function

'-------------------------------------------------------------------------------
' Protection et encadrement des traitements
'-------------------------------------------------------------------------------
Public Function FeuillesProtegees() As Variant
    FeuillesProtegees = Array(F_TDB, F_SAISIE, F_CHARGES, F_SUIVI, F_RAPPRO, F_ATRAITER, F_BANQUE, F_PARAM, F_AIDE)
End Function

Public Sub ProtegerFeuille(ws As Worksheet)
    On Error Resume Next
    ws.Protect Password:=MOT_DE_PASSE, DrawingObjects:=True, Contents:=True, Scenarios:=True, _
        UserInterfaceOnly:=True, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowFiltering:=True
    ws.EnableSelection = xlNoRestrictions
End Sub

Public Sub ProtegerTout()
    Dim nom As Variant
    On Error Resume Next
    For Each nom In FeuillesProtegees()
        ProtegerFeuille ThisWorkbook.Worksheets(CStr(nom))
    Next nom
End Sub

Public Sub DeprotegerTout()
    Dim nom As Variant
    On Error Resume Next
    For Each nom In FeuillesProtegees()
        ThisWorkbook.Worksheets(CStr(nom)).Unprotect MOT_DE_PASSE
    Next nom
End Sub

Public Sub DebutTraitement(Optional ByVal message As String = "")
    If mProfondeur = 0 Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.Calculation = xlCalculationManual
        DeprotegerTout
    End If
    mProfondeur = mProfondeur + 1
    If message <> "" Then Application.StatusBar = message
End Sub

Public Sub FinTraitement()
    mProfondeur = mProfondeur - 1
    If mProfondeur <= 0 Then TerminerTraitement
End Sub

' A appeler dans les gestionnaires d'erreur : remet Excel dans un état normal
Public Sub TerminerTraitement()
    On Error Resume Next
    mProfondeur = 0
    Application.Calculation = xlCalculationAutomatic
    ProtegerTout
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

Public Sub AfficherStatut(ByVal message As String)
    Application.StatusBar = message
    DoEvents
End Sub

Public Sub AfficherErreur(ByVal contexte As String, ByVal description As String)
    If description = "" Then description = "Erreur inconnue."
    MsgBox "Une erreur est survenue pendant : " & contexte & vbCrLf & vbCrLf & description & vbCrLf & vbCrLf & _
        "Les données déjà enregistrées n'ont pas été supprimées." & vbCrLf & _
        "En cas de doute, une sauvegarde est disponible dans le dossier « Sauvegardes ».", _
        vbExclamation, "Trésorerie - erreur"
End Sub

'-------------------------------------------------------------------------------
' Dossiers
'-------------------------------------------------------------------------------
Private Function SansBarreFinale(ByVal s As String) As String
    s = Trim$(s)
    Do While Len(s) > 3 And Right$(s, 1) = "\"
        s = Left$(s, Len(s) - 1)
    Loop
    SansBarreFinale = s
End Function

Public Function DossierExiste(ByVal chemin As String) As Boolean
    On Error Resume Next
    DossierExiste = ((GetAttr(chemin) And vbDirectory) = vbDirectory)
    If Err.Number <> 0 Then DossierExiste = False
    Err.Clear
End Function

Public Sub CreerDossier(ByVal chemin As String)
    Dim pos As Long
    chemin = SansBarreFinale(chemin)
    If chemin = "" Then Exit Sub
    If DossierExiste(chemin) Then Exit Sub
    pos = InStrRev(chemin, "\")
    If pos > 3 Then CreerDossier Left$(chemin, pos - 1)
    On Error Resume Next
    MkDir chemin
    On Error GoTo 0
    If Not DossierExiste(chemin) Then
        Err.Raise vbObjectError + 1003, "modOutils.CreerDossier", "Impossible de créer le dossier : " & chemin
    End If
End Sub

' Chemin local du classeur (gère les classeurs synchronisés OneDrive / SharePoint)
Public Function CheminLocalClasseur() As String
    Dim p As String
    p = ThisWorkbook.Path
    If LCase$(Left$(p, 4)) = "http" Then p = ConvertirCheminOneDrive(p)
    CheminLocalClasseur = p
End Function

Private Function ConvertirCheminOneDrive(ByVal url As String) As String
    Dim reste As String, pos As Long, variable As Variant, base As String, candidat As String
    reste = url
    If InStr(1, reste, "d.docs.live.net", vbTextCompare) > 0 Then
        pos = InStr(9, reste, "/")
        pos = InStr(pos + 1, reste, "/")
        If pos > 0 Then reste = Mid$(reste, pos + 1) Else reste = ""
    ElseIf InStr(1, reste, "/Documents/", vbTextCompare) > 0 Then
        reste = Mid$(reste, InStr(1, reste, "/Documents/", vbTextCompare) + 11)
    ElseIf LCase$(Right$(reste, 10)) = "/documents" Then
        reste = ""
    End If
    reste = Replace(reste, "/", "\")
    reste = Replace(reste, "%20", " ")
    For Each variable In Array("OneDriveCommercial", "OneDriveConsumer", "OneDrive")
        base = Environ$(CStr(variable))
        If base <> "" Then
            If reste = "" Then candidat = base Else candidat = base & "\" & reste
            If DossierExiste(candidat) Then
                ConvertirCheminOneDrive = candidat
                Exit Function
            End If
        End If
    Next variable
    ConvertirCheminOneDrive = ""
End Function

Public Function DossierRacine() As String
    Dim s As String
    s = Texte(ValeurNom("Dossier_Racine"))
    If s = "" Then
        If ThisWorkbook.Path = "" Then
            Err.Raise vbObjectError + 1004, "modOutils.DossierRacine", _
                "Le classeur n'a jamais été enregistré : enregistrez-le d'abord dans son dossier définitif."
        End If
        s = CheminLocalClasseur()
    End If
    If s = "" Then
        Err.Raise vbObjectError + 1004, "modOutils.DossierRacine", _
            "Impossible de déterminer le dossier local du classeur (OneDrive / SharePoint)." & vbCrLf & _
            "Indiquez un dossier local dans l'onglet Paramètres, rubrique « Dossier racine »."
    End If
    DossierRacine = SansBarreFinale(s)
End Function

Public Function DossierBanque(ByVal code As String) As String
    Dim s As String
    s = Texte(ValeurNom("Dossier_" & code))
    If s = "" Then
        If code = "B1" Then s = DossierRacine() & "\Imports\Banque1" Else s = DossierRacine() & "\Imports\Banque2"
    End If
    DossierBanque = SansBarreFinale(s)
End Function

Public Function DossierSauvegardes() As String
    DossierSauvegardes = DossierRacine() & "\Sauvegardes"
End Function

'-------------------------------------------------------------------------------
' Banques
'-------------------------------------------------------------------------------
Public Function NomBanque(ByVal code As String) As String
    NomBanque = Texte(ValeurNom(code & "_Nom"))
End Function

' Date du solde initial d'une banque désignée par son nom (0 si banque inconnue)
Public Function DateSoldeBanque(ByVal nom As String) As Double
    If nom = "" Then
        DateSoldeBanque = 0
    ElseIf nom = NomBanque("B1") Then
        DateSoldeBanque = DateVal(ValeurNom("B1_DateSolde"))
    ElseIf nom = NomBanque("B2") Then
        DateSoldeBanque = DateVal(ValeurNom("B2_DateSolde"))
    Else
        DateSoldeBanque = 0
    End If
End Function

Public Function BanqueConnue(ByVal nom As String) As Boolean
    BanqueConnue = (nom <> "" And (nom = NomBanque("B1") Or nom = NomBanque("B2")))
End Function
