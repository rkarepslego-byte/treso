Attribute VB_Name = "modRapprochement"
'===============================================================================
' modRapprochement - Moteur de rapprochement prévisions / opérations bancaires
'-------------------------------------------------------------------------------
' Critères (dans l'ordre) : même banque, même montant au centime, écart de dates
' <= tolérance (Paramètres, 7 jours par défaut), similarité des libellés.
'
' * Composante « 1 opération - 1 prévision » : correspondance CERTAINE si l'écart
'   est <= « Écart max certaine » OU si les libellés sont similaires (>= seuil).
' * Plusieurs candidats : correspondance INCERTAINE, jamais validée automatiquement.
'   Le moteur indique la paire la plus probable (« Suggérée ») ; cette prévision est
'   provisoirement neutralisée dans le prévisionnel pour ne pas compter deux fois
'   l'opération déjà présente dans le réel.
' * Une opération sans candidat va dans la liste « À traiter ».
' Le rapprochement est enregistré dans les tables (ID prévision / ID opération) :
' il est conservé d'une actualisation à l'autre.
' Miroir de reference/moteur.py : rechercher / valider / delier.
'===============================================================================
Option Explicit
Option Compare Binary

'-------------------------------------------------------------------------------
' Boutons
'-------------------------------------------------------------------------------
Public Sub RechercherCorrespondances()
    Dim nbCert As Long, nbIncert As Long, nbAuto As Long, nbATraiter As Long, msgErr As String
    On Error GoTo Erreur
    DebutTraitement "Recherche des correspondances..."
    modPrevisions.CompleterSaisiesInterne
    RechercherInterne nbCert, nbIncert, nbAuto
    nbATraiter = modATraiter.ConstruireListeATraiterInterne()
    FinTraitement
    Application.Calculate
    MsgBox ResumeRecherche(nbCert, nbIncert, nbAuto, nbATraiter), vbInformation, "Rapprochement"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la recherche des correspondances", msgErr
End Sub

Public Function ResumeRecherche(ByVal nbCert As Long, ByVal nbIncert As Long, ByVal nbAuto As Long, _
                                ByVal nbATraiter As Long) As String
    Dim s As String
    s = "Propositions de rapprochement : " & vbCrLf & _
        "   - correspondances certaines : " & nbCert & vbCrLf & _
        "   - correspondances incertaines (à choisir) : " & nbIncert & vbCrLf
    If nbAuto > 0 Then s = s & "   - validées automatiquement : " & nbAuto & vbCrLf
    s = s & "Nouvelles opérations bancaires à traiter : " & nbATraiter
    ResumeRecherche = s
End Function

Public Sub ValiderChoix()
    Dim lo As ListObject, cR As Object, dR As Variant, nR As Long, msgErr As String, resultat As String
    On Error GoTo Erreur
    Set lo = T(T_PROPOS)
    Set cR = Colonnes(lo)
    dR = LireTable(lo)
    nR = NbLignes(dR)
    DebutTraitement "Validation des rapprochements..."
    resultat = AppliquerChoix(dR, nR, cR)
    FinTraitement
    Application.Calculate
    MsgBox resultat, vbInformation, "Rapprochement"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la validation des rapprochements", msgErr
End Sub

Public Sub ToutValiderCertaines()
    Dim lo As ListObject, cR As Object, dR As Variant, nR As Long, i As Long, nb As Long
    Dim msgErr As String, resultat As String
    On Error GoTo Erreur
    Set lo = T(T_PROPOS)
    Set cR = Colonnes(lo)
    dR = LireTable(lo)
    nR = NbLignes(dR)
    For i = 1 To nR
        If Texte(dR(i, C(cR, "Niveau"))) = NIV_CERTAINE Then
            dR(i, C(cR, "Valider")) = "Oui"
            nb = nb + 1
        End If
    Next i
    If nb = 0 Then
        MsgBox "Aucune correspondance certaine à valider.", vbInformation, "Rapprochement"
        Exit Sub
    End If
    If MsgBox("Valider les " & nb & " correspondance(s) certaine(s) ?" & vbCrLf & _
              "Les correspondances incertaines ne sont jamais validées automatiquement.", _
              vbOKCancel + vbQuestion, "Rapprochement") <> vbOK Then Exit Sub
    DebutTraitement "Validation des rapprochements..."
    resultat = AppliquerChoix(dR, nR, cR)
    FinTraitement
    Application.Calculate
    MsgBox resultat, vbInformation, "Rapprochement"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la validation des rapprochements", msgErr
End Sub

' Annule le rapprochement (ou le traitement) de la ligne sélectionnée
Public Sub AnnulerRapprochement()
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long
    Dim id As String, iB As Long, iP As Long, i As Long, r As Long, statutOp As String, msgErr As String
    Dim idOp As String, idPrev As String

    On Error GoTo Erreur
    Set loB = T(T_BANQUE)
    Set loP = T(T_PREV)
    Set cB = Colonnes(loB)
    Set cP = Colonnes(loP)
    If ActiveSheet.Name = loB.Parent.Name Then
        If Intersect(ActiveCell, loB.DataBodyRange) Is Nothing Then GoTo HorsTable
        r = ActiveCell.Row - loB.HeaderRowRange.Row
        id = Texte(loB.DataBodyRange.Cells(r, C(cB, "ID")).Value)
    ElseIf ActiveSheet.Name = loP.Parent.Name Then
        If Intersect(ActiveCell, loP.DataBodyRange) Is Nothing Then GoTo HorsTable
        r = ActiveCell.Row - loP.HeaderRowRange.Row
        id = Texte(loP.DataBodyRange.Cells(r, C(cP, "ID")).Value)
    Else
        GoTo HorsTable
    End If
    If id = "" Then GoTo HorsTable

    dB = LireTable(loB)
    nB = NbLignes(dB)
    dP = LireTable(loP)
    nP = NbLignes(dP)
    For i = 1 To nB
        If Texte(dB(i, C(cB, "ID"))) = id Then iB = i
    Next i
    For i = 1 To nP
        If Texte(dP(i, C(cP, "ID"))) = id Then iP = i
    Next i
    If iB > 0 Then
        statutOp = Texte(dB(iB, C(cB, "Statut")))
        If statutOp = OB_POINTE Then
            For i = 1 To nP
                If Texte(dP(i, C(cP, "ID"))) = Texte(dB(iB, C(cB, "ID prévision"))) Then iP = i
            Next i
        ElseIf statutOp = OB_NON_PREVUE Or statutOp = OB_VIREMENT Then
            If MsgBox("L'opération " & id & " a été traitée (« " & statutOp & " »)." & vbCrLf & _
                      "La remettre dans la liste des opérations à traiter ?", vbOKCancel + vbQuestion, "Annulation") <> vbOK Then Exit Sub
            DebutTraitement
            dB(iB, C(cB, "Statut")) = OB_A_TRAITER
            dB(iB, C(cB, "Mode")) = ""
            EcrireTable loB, dB, nB
            RechercherInterne
            modATraiter.ConstruireListeATraiterInterne
            FinTraitement
            Application.Calculate
            Exit Sub
        Else
            MsgBox "L'opération " & id & " n'est pas rapprochée.", vbInformation, "Annulation"
            Exit Sub
        End If
    ElseIf iP > 0 Then
        If Texte(dP(iP, C(cP, "Statut"))) <> P_POINTE Then
            MsgBox "La prévision " & id & " n'est pas pointée.", vbInformation, "Annulation"
            Exit Sub
        End If
        For i = 1 To nB
            If Texte(dB(i, C(cB, "ID"))) = Texte(dP(iP, C(cP, "ID opération"))) Then iB = i
        Next i
    End If

    idOp = "?"
    idPrev = "?"
    If iB > 0 Then idOp = Texte(dB(iB, C(cB, "ID")))
    If iP > 0 Then idPrev = Texte(dP(iP, C(cP, "ID")))
    If MsgBox("Annuler le rapprochement entre l'opération " & idOp & " et la prévision " & idPrev & " ?", _
              vbOKCancel + vbQuestion, "Annulation") <> vbOK Then Exit Sub

    DebutTraitement "Annulation du rapprochement..."
    If iP > 0 Then
        dP(iP, C(cP, "Statut")) = P_PREVU
        dP(iP, C(cP, "ID opération")) = ""
        dP(iP, C(cP, "Date réelle")) = Empty
        dP(iP, C(cP, "Montant réel")) = Empty
        EcrireTable loP, dP, nP
    End If
    If iB > 0 Then
        dB(iB, C(cB, "Statut")) = OB_A_TRAITER
        dB(iB, C(cB, "ID prévision")) = ""
        dB(iB, C(cB, "Mode")) = ""
        EcrireTable loB, dB, nB
    End If
    RechercherInterne
    modATraiter.ConstruireListeATraiterInterne
    FinTraitement
    Application.Calculate
    Exit Sub
HorsTable:
    MsgBox "Sélectionnez d'abord une cellule de la ligne concernée dans le tableau des prévisions (onglet Saisie)" & vbCrLf & _
           "ou dans le tableau des opérations bancaires.", vbInformation, "Annulation"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "l'annulation du rapprochement", msgErr
End Sub

'-------------------------------------------------------------------------------
' Liaison d'une opération et d'une prévision (dans les tableaux en mémoire)
'-------------------------------------------------------------------------------
Public Sub Lier(dB As Variant, cB As Object, ByVal iB As Long, dP As Variant, cP As Object, ByVal iP As Long, _
                ByVal mode As String)
    dP(iP, C(cP, "Statut")) = P_POINTE
    dP(iP, C(cP, "ID opération")) = Texte(dB(iB, C(cB, "ID")))
    dP(iP, C(cP, "ID proposé")) = ""
    dP(iP, C(cP, "Date réelle")) = DateVal(dB(iB, C(cB, "Date")))
    dP(iP, C(cP, "Montant réel")) = Nombre(dB(iB, C(cB, "Montant")))
    dB(iB, C(cB, "Statut")) = OB_POINTE
    dB(iB, C(cB, "ID prévision")) = Texte(dP(iP, C(cP, "ID")))
    dB(iB, C(cB, "Mode")) = mode
    If Texte(dB(iB, C(cB, "Catégorie"))) = "" Then dB(iB, C(cB, "Catégorie")) = Texte(dP(iP, C(cP, "Catégorie")))
End Sub

'-------------------------------------------------------------------------------
' Application des choix de l'onglet Rapprochement (« Oui » / « Non »)
'-------------------------------------------------------------------------------
Private Function AppliquerChoix(dR As Variant, ByVal nR As Long, cR As Object) As String
    Dim vusO As Object, vusP As Object, conflits As String, i As Long, choix As String
    Dim idO As String, idP As String, nbLiees As Long, nbRejets As Long, nbIgnorees As Long
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long, indexB As Object
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long, indexP As Object
    Dim iB As Long, iP As Long, rejets() As Variant, loRej As ListObject, cRej As Object
    Dim nbCert As Long, nbIncert As Long, nbAuto As Long, nbATraiter As Long

    Set vusO = CreateObject("Scripting.Dictionary")
    Set vusP = CreateObject("Scripting.Dictionary")
    For i = 1 To nR
        If EstOui(dR(i, C(cR, "Valider"))) Then
            idO = Texte(dR(i, C(cR, "ID opération")))
            idP = Texte(dR(i, C(cR, "ID prévision")))
            If vusO.Exists(idO) Then conflits = conflits & "- l'opération " & idO & " est validée plusieurs fois" & vbCrLf
            If vusP.Exists(idP) Then conflits = conflits & "- la prévision " & idP & " est validée plusieurs fois" & vbCrLf
            vusO(idO) = True
            vusP(idP) = True
        End If
    Next i
    If conflits <> "" Then
        AppliquerChoix = "Aucune validation n'a été enregistrée car certains choix sont contradictoires :" & vbCrLf & _
                         conflits & vbCrLf & "Une opération ne peut être rapprochée que d'une seule prévision (et inversement)."
        Exit Function
    End If

    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    Set loP = T(T_PREV)
    Set cP = Colonnes(loP)
    dP = LireTable(loP)
    nP = NbLignes(dP)
    Set indexB = CreateObject("Scripting.Dictionary")
    Set indexP = CreateObject("Scripting.Dictionary")
    For i = 1 To nB
        If Texte(dB(i, C(cB, "ID"))) <> "" Then indexB(Texte(dB(i, C(cB, "ID")))) = i
    Next i
    For i = 1 To nP
        If Texte(dP(i, C(cP, "ID"))) <> "" Then indexP(Texte(dP(i, C(cP, "ID")))) = i
    Next i

    Set loRej = T(T_REJETS)
    Set cRej = Colonnes(loRej)
    ReDim rejets(1 To nR + 1, 1 To loRej.ListColumns.Count)
    For i = 1 To nR
        choix = UCase$(Texte(dR(i, C(cR, "Valider"))))
        idO = Texte(dR(i, C(cR, "ID opération")))
        idP = Texte(dR(i, C(cR, "ID prévision")))
        If choix = "OUI" Then
            iB = 0
            iP = 0
            If indexB.Exists(idO) Then iB = indexB(idO)
            If indexP.Exists(idP) Then iP = indexP(idP)
            If iB = 0 Or iP = 0 Then
                nbIgnorees = nbIgnorees + 1
            ElseIf Texte(dB(iB, C(cB, "Statut"))) <> OB_A_TRAITER And Texte(dB(iB, C(cB, "Statut"))) <> OB_PROPOSE Then
                nbIgnorees = nbIgnorees + 1
            ElseIf Texte(dP(iP, C(cP, "Statut"))) = P_POINTE Or Texte(dP(iP, C(cP, "Statut"))) = P_ANNULE _
                   Or Texte(dP(iP, C(cP, "ID opération"))) <> "" Then
                nbIgnorees = nbIgnorees + 1
            Else
                Lier dB, cB, iB, dP, cP, iP, MODE_VALIDE
                nbLiees = nbLiees + 1
            End If
        ElseIf choix = "NON" Then
            nbRejets = nbRejets + 1
            rejets(nbRejets, C(cRej, "ID opération")) = idO
            rejets(nbRejets, C(cRej, "ID prévision")) = idP
            rejets(nbRejets, C(cRej, "Date")) = CDbl(Now)
        End If
    Next i

    EcrireTable loB, dB, nB
    EcrireTable loP, dP, nP
    AjouterLignes loRej, rejets, nbRejets, C(cRej, "ID opération")
    RechercherInterne nbCert, nbIncert, nbAuto
    nbATraiter = modATraiter.ConstruireListeATraiterInterne()

    AppliquerChoix = nbLiees & " rapprochement(s) validé(s), " & nbRejets & " proposition(s) rejetée(s)." & vbCrLf & _
                     IIf(nbIgnorees > 0, nbIgnorees & " choix ignoré(s) (déjà traités entre-temps)." & vbCrLf, "") & vbCrLf & _
                     ResumeRecherche(nbCert, nbIncert, nbAuto, nbATraiter)
End Function

'-------------------------------------------------------------------------------
' Recherche des correspondances
'-------------------------------------------------------------------------------
Public Sub RechercherInterne(Optional ByRef nbCertaines As Long, Optional ByRef nbIncertaines As Long, _
                             Optional ByRef nbAuto As Long)
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long
    Dim rejets As Object, loRej As ListObject, cRej As Object, dRej As Variant, i As Long
    Dim autoValider As Boolean, nProps As Long, k As Long, tours As Long
    Dim pGroupe() As Long, pNiveau() As String, pSugg() As Boolean, pB() As Long, pP() As Long
    Dim pEcart() As Long, pSim() As Double, pScore() As Double

    nbAuto = 0
    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    Set loP = T(T_PREV)
    Set cP = Colonnes(loP)
    dP = LireTable(loP)
    nP = NbLignes(dP)

    Set rejets = CreateObject("Scripting.Dictionary")
    Set loRej = T(T_REJETS)
    Set cRej = Colonnes(loRej)
    dRej = LireTable(loRej)
    For i = 1 To NbLignes(dRej)
        rejets(Texte(dRej(i, C(cRej, "ID opération"))) & "|" & Texte(dRej(i, C(cRej, "ID prévision")))) = True
    Next i
    autoValider = EstOui(ValeurNom("Auto_Valider"))

    Do
        tours = tours + 1
        CalculerPropositions dB, nB, cB, dP, nP, cP, rejets, nProps, pGroupe, pNiveau, pSugg, pB, pP, _
                             pEcart, pSim, pScore, nbCertaines, nbIncertaines
        If Not autoValider Or nbCertaines = 0 Or tours > 20 Then Exit Do
        For k = 1 To nProps
            If pNiveau(k) = NIV_CERTAINE Then
                Lier dB, cB, pB(k), dP, cP, pP(k), MODE_AUTO
                nbAuto = nbAuto + 1
            End If
        Next k
    Loop

    EcrireTable loB, dB, nB
    EcrireTable loP, dP, nP
    EcrirePropositions dB, cB, dP, cP, nProps, pGroupe, pNiveau, pSugg, pB, pP, pEcart, pSim, pScore
End Sub

Private Sub EcrirePropositions(dB As Variant, cB As Object, dP As Variant, cP As Object, ByVal nProps As Long, _
                               pGroupe() As Long, pNiveau() As String, pSugg() As Boolean, pB() As Long, pP() As Long, _
                               pEcart() As Long, pSim() As Double, pScore() As Double)
    Dim lo As ListObject, cR As Object, sortie() As Variant, k As Long
    Set lo = T(T_PROPOS)
    Set cR = Colonnes(lo)
    ReDim sortie(1 To IIf(nProps < 1, 1, nProps), 1 To lo.ListColumns.Count)
    For k = 1 To nProps
        sortie(k, C(cR, "Niveau")) = pNiveau(k)
        sortie(k, C(cR, "Groupe")) = pGroupe(k)
        If pSugg(k) Then sortie(k, C(cR, "Suggestion")) = SUGGERE
        sortie(k, C(cR, "Banque")) = Texte(dB(pB(k), C(cB, "Banque")))
        sortie(k, C(cR, "Date opération")) = DateVal(dB(pB(k), C(cB, "Date")))
        sortie(k, C(cR, "Libellé bancaire")) = Texte(dB(pB(k), C(cB, "Libellé")))
        sortie(k, C(cR, "Montant")) = Nombre(dB(pB(k), C(cB, "Montant")))
        sortie(k, C(cR, "Date prévue")) = DateVal(dP(pP(k), C(cP, "Date prévue")))
        sortie(k, C(cR, "Libellé prévu")) = Texte(dP(pP(k), C(cP, "Libellé")))
        sortie(k, C(cR, "Écart (jours)")) = pEcart(k)
        sortie(k, C(cR, "Similarité")) = pSim(k)
        sortie(k, C(cR, "Score")) = pScore(k)
        sortie(k, C(cR, "ID opération")) = Texte(dB(pB(k), C(cB, "ID")))
        sortie(k, C(cR, "ID prévision")) = Texte(dP(pP(k), C(cP, "ID")))
    Next k
    EcrireTable lo, sortie, nProps
End Sub

Private Function Racine(parent As Object, ByVal x As String) As String
    If Not parent.Exists(x) Then parent(x) = x
    Do While parent(x) <> x
        parent(x) = parent(parent(x))
        x = parent(x)
    Loop
    Racine = x
End Function

' Coeur du moteur : met à jour les statuts « Proposé » / « ID proposé » dans les tableaux
' en mémoire et renvoie les propositions (une ligne par paire candidate).
Private Sub CalculerPropositions(dB As Variant, ByVal nB As Long, cB As Object, _
                                 dP As Variant, ByVal nP As Long, cP As Object, rejets As Object, _
                                 ByRef nProps As Long, pGroupe() As Long, pNiveau() As String, pSugg() As Boolean, _
                                 pB() As Long, pP() As Long, pEcart() As Long, pSim() As Double, pScore() As Double, _
                                 ByRef nbCertaines As Long, ByRef nbIncertaines As Long)
    Dim tolerance As Long, seuil As Double, ecartCertain As Long
    Dim i As Long, j As Long, k As Long, m As Long, banque As String, cle As String, statutP As String
    Dim index As Object, liste() As String, d0 As Double, dOp As Double, dPrev As Double, ecart As Long
    Dim nE As Long, eB() As Long, eP() As Long, eEcart() As Long, eSim() As Double, eScore() As Double
    Dim parent As Object, comps As Object, clesComp() As String, idxComp() As Long, racines As Variant
    Dim nComp As Long, num As Long, aretes() As String, nA As Long, opsVus As Object, prevsVus As Object
    Dim ordre() As Long, prisO As Object, prisP As Object, sugg As Object, rO As String, rP As String
    Dim niveau As String, a As Long, cleMin As String, cleE As String

    tolerance = CLng(Nombre(ValeurNom("Tolerance_Jours")))
    seuil = Nombre(ValeurNom("Seuil_Similarite"))
    ecartCertain = CLng(Nombre(ValeurNom("Ecart_Certain")))
    nProps = 0
    nbCertaines = 0
    nbIncertaines = 0

    ' 1. Réinitialisation des propositions précédentes
    For i = 1 To nB
        If Texte(dB(i, C(cB, "Statut"))) = OB_PROPOSE Then dB(i, C(cB, "Statut")) = OB_A_TRAITER
    Next i
    For j = 1 To nP
        dP(j, C(cP, "ID proposé")) = ""
    Next j

    ' 2. Index des prévisions candidates par « banque|montant en centimes »
    Set index = CreateObject("Scripting.Dictionary")
    For j = 1 To nP
        banque = Texte(dP(j, C(cP, "Banque")))
        statutP = Texte(dP(j, C(cP, "Statut")))
        dPrev = DateVal(dP(j, C(cP, "Date prévue")))
        If Texte(dP(j, C(cP, "ID"))) <> "" And BanqueConnue(banque) And dPrev > 0 _
           And statutP <> P_POINTE And statutP <> P_ANNULE And Texte(dP(j, C(cP, "ID opération"))) = "" _
           And Nombre(dP(j, C(cP, "Montant"))) <> 0 Then
            If dPrev > DateSoldeBanque(banque) Then
                cle = banque & "|" & CentsTexte(modPrevisions.MontantSigne(Texte(dP(j, C(cP, "Type"))), Nombre(dP(j, C(cP, "Montant")))))
                index(cle) = index(cle) & CStr(j) & ";"
            End If
        End If
    Next j

    ' 3. Arêtes candidates
    For i = 1 To nB
        banque = Texte(dB(i, C(cB, "Banque")))
        dOp = DateVal(dB(i, C(cB, "Date")))
        If (Texte(dB(i, C(cB, "Statut"))) = OB_A_TRAITER Or Texte(dB(i, C(cB, "Statut"))) = OB_PROPOSE) _
           And Texte(dB(i, C(cB, "ID"))) <> "" And BanqueConnue(banque) Then
            d0 = DateSoldeBanque(banque)
            If dOp > d0 Then
                cle = banque & "|" & CentsTexte(Nombre(dB(i, C(cB, "Montant"))))
                If index.Exists(cle) Then
                    liste = Split(index(cle), ";")
                    For k = LBound(liste) To UBound(liste)
                        If liste(k) <> "" Then
                            j = CLng(liste(k))
                            dPrev = DateVal(dP(j, C(cP, "Date prévue")))
                            ecart = CLng(Abs(dOp - dPrev))
                            If ecart <= tolerance Then
                                If Not rejets.Exists(Texte(dB(i, C(cB, "ID"))) & "|" & Texte(dP(j, C(cP, "ID")))) Then
                                    nE = nE + 1
                                    ReDim Preserve eB(1 To nE)
                                    ReDim Preserve eP(1 To nE)
                                    ReDim Preserve eEcart(1 To nE)
                                    ReDim Preserve eSim(1 To nE)
                                    ReDim Preserve eScore(1 To nE)
                                    eB(nE) = i
                                    eP(nE) = j
                                    eEcart(nE) = ecart
                                    eSim(nE) = Similarite(Texte(dB(i, C(cB, "Libellé"))), Texte(dP(j, C(cP, "Libellé"))))
                                    eScore(nE) = Round(0.5 * (1 - ecart / (tolerance + 1)) + 0.5 * eSim(nE), 4)
                                End If
                            End If
                        End If
                    Next k
                End If
            End If
        End If
    Next i
    If nE = 0 Then Exit Sub

    ' 4. Composantes connexes (union-find)
    Set parent = CreateObject("Scripting.Dictionary")
    For k = 1 To nE
        rO = Racine(parent, "O" & eB(k))
        rP = Racine(parent, "P" & eP(k))
        If rO <> rP Then parent(rP) = rO
    Next k
    Set comps = CreateObject("Scripting.Dictionary")
    For k = 1 To nE
        cle = Racine(parent, "O" & eB(k))
        comps(cle) = comps(cle) & CStr(k) & ";"
    Next k

    ' Ordre des composantes : plus petite (date, ID) d'opération
    nComp = comps.Count
    racines = comps.Keys
    ReDim clesComp(1 To nComp)
    ReDim idxComp(1 To nComp)
    For m = 1 To nComp
        idxComp(m) = m
        aretes = Split(comps(racines(m - 1)), ";")
        cleMin = ""
        For a = LBound(aretes) To UBound(aretes)
            If aretes(a) <> "" Then
                k = CLng(aretes(a))
                cleE = CleDate(DateVal(dB(eB(k), C(cB, "Date")))) & Texte(dB(eB(k), C(cB, "ID")))
                If cleMin = "" Or StrComp(cleE, cleMin, vbBinaryCompare) < 0 Then cleMin = cleE
            End If
        Next a
        clesComp(m) = cleMin
    Next m
    TrierIndices clesComp, idxComp, 1, nComp

    ReDim pGroupe(1 To nE)
    ReDim pNiveau(1 To nE)
    ReDim pSugg(1 To nE)
    ReDim pB(1 To nE)
    ReDim pP(1 To nE)
    ReDim pEcart(1 To nE)
    ReDim pSim(1 To nE)
    ReDim pScore(1 To nE)

    For num = 1 To nComp
        aretes = Split(comps(racines(idxComp(num) - 1)), ";")
        nA = 0
        For a = LBound(aretes) To UBound(aretes)
            If aretes(a) <> "" Then nA = nA + 1
        Next a
        ReDim ordre(1 To nA)
        Set opsVus = CreateObject("Scripting.Dictionary")
        Set prevsVus = CreateObject("Scripting.Dictionary")
        m = 0
        For a = LBound(aretes) To UBound(aretes)
            If aretes(a) <> "" Then
                m = m + 1
                k = CLng(aretes(a))
                ordre(m) = k
                opsVus(eB(k)) = True
                prevsVus(eP(k)) = True
            End If
        Next a

        ' Tri par pertinence : score décroissant, écart, date prévue, ID prévision, ID opération
        CleTriPertinence ordre, nA, dB, cB, dP, cP, eB, eP, eEcart, eScore

        Set sugg = CreateObject("Scripting.Dictionary")
        If opsVus.Count = 1 And prevsVus.Count = 1 Then
            k = ordre(1)
            If eEcart(k) <= ecartCertain Or eSim(k) >= seuil Then niveau = NIV_CERTAINE Else niveau = NIV_INCERTAINE
            sugg(k) = True
        Else
            niveau = NIV_INCERTAINE
            Set prisO = CreateObject("Scripting.Dictionary")
            Set prisP = CreateObject("Scripting.Dictionary")
            For m = 1 To nA
                k = ordre(m)
                If Not prisO.Exists(eB(k)) And Not prisP.Exists(eP(k)) Then
                    sugg(k) = True
                    prisO(eB(k)) = True
                    prisP(eP(k)) = True
                End If
            Next m
        End If
        If niveau = NIV_CERTAINE Then nbCertaines = nbCertaines + 1 Else nbIncertaines = nbIncertaines + 1

        ' Ordre d'affichage : opération (date, ID), puis score décroissant, date prévue, ID prévision
        CleTriAffichage ordre, nA, dB, cB, dP, cP, eB, eP, eScore
        For m = 1 To nA
            k = ordre(m)
            nProps = nProps + 1
            pGroupe(nProps) = num
            pNiveau(nProps) = niveau
            pSugg(nProps) = sugg.Exists(k)
            pB(nProps) = eB(k)
            pP(nProps) = eP(k)
            pEcart(nProps) = eEcart(k)
            pSim(nProps) = eSim(k)
            pScore(nProps) = eScore(k)
            dB(eB(k), C(cB, "Statut")) = OB_PROPOSE
            If sugg.Exists(k) Then dP(eP(k), C(cP, "ID proposé")) = Texte(dB(eB(k), C(cB, "ID")))
        Next m
    Next num
End Sub

Private Function CleScore(ByVal score As Double) As String
    CleScore = Format$(10000 - CLng(Round(score * 10000, 0)), "00000")
End Function

' Trie ordre(1..n) (numéros d'arêtes) par pertinence décroissante
Private Sub CleTriPertinence(ordre() As Long, ByVal n As Long, dB As Variant, cB As Object, _
                             dP As Variant, cP As Object, eB() As Long, eP() As Long, eEcart() As Long, eScore() As Double)
    Dim m As Long, k As Long, cl() As String, idx() As Long, copie() As Long
    ReDim cl(1 To n)
    ReDim idx(1 To n)
    ReDim copie(1 To n)
    For m = 1 To n
        k = ordre(m)
        idx(m) = m
        copie(m) = k
        cl(m) = CleScore(eScore(k)) & Format$(eEcart(k), "000") & CleDate(DateVal(dP(eP(k), C(cP, "Date prévue")))) & _
                Texte(dP(eP(k), C(cP, "ID"))) & "|" & Texte(dB(eB(k), C(cB, "ID")))
    Next m
    TrierIndices cl, idx, 1, n
    For m = 1 To n
        ordre(m) = copie(idx(m))
    Next m
End Sub

Private Sub CleTriAffichage(ordre() As Long, ByVal n As Long, dB As Variant, cB As Object, _
                            dP As Variant, cP As Object, eB() As Long, eP() As Long, eScore() As Double)
    Dim m As Long, k As Long, cl() As String, idx() As Long, copie() As Long
    ReDim cl(1 To n)
    ReDim idx(1 To n)
    ReDim copie(1 To n)
    For m = 1 To n
        k = ordre(m)
        idx(m) = m
        copie(m) = k
        cl(m) = CleDate(DateVal(dB(eB(k), C(cB, "Date")))) & Texte(dB(eB(k), C(cB, "ID"))) & "|" & CleScore(eScore(k)) & _
                CleDate(DateVal(dP(eP(k), C(cP, "Date prévue")))) & Texte(dP(eP(k), C(cP, "ID")))
    Next m
    TrierIndices cl, idx, 1, n
    For m = 1 To n
        ordre(m) = copie(idx(m))
    Next m
End Sub
