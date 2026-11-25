Attribute VB_Name = "modPrevisions"
'===============================================================================
' modPrevisions - Saisie des opérations prévisionnelles et charges fixes
'-------------------------------------------------------------------------------
' * formulaire de saisie (onglet Saisie) et virements internes (2 lignes liées)
' * complétion des lignes saisies directement dans le tableau (ID, valeurs par défaut)
' * lignes de saisie libres en bas des tableaux (les tables d'une feuille protégée
'   ne peuvent pas s'agrandir d'elles-mêmes)
' * génération des échéances des charges fixes, SANS effet rétroactif
'   (miroir de reference/moteur.py : generer_echeances)
'===============================================================================
Option Explicit
Option Compare Binary

'-------------------------------------------------------------------------------
' Outils
'-------------------------------------------------------------------------------
Public Function LigneBlanche(data As Variant, ByVal i As Long, colonnesTest As Variant) As Boolean
    Dim k As Long
    For k = LBound(colonnesTest) To UBound(colonnesTest)
        If Texte(data(i, colonnesTest(k))) <> "" Then
            LigneBlanche = False
            Exit Function
        End If
    Next k
    LigneBlanche = True
End Function

Public Function MontantSigne(ByVal typeOp As String, ByVal montant As Double) As Double
    If typeOp = TY_RECETTE Or typeOp = TY_VIR_ENTRANT Then
        MontantSigne = Abs(montant)
    Else
        MontantSigne = -Abs(montant)
    End If
End Function

Public Function Empreinte(ByVal d As Double, ByVal banque As String, ByVal libelle As String, _
                          ByVal typeOp As String, ByVal montant As Double) As String
    Empreinte = DateISO(d) & "|" & banque & "|" & libelle & "|" & typeOp & "|" & CentsTexte(Abs(montant))
End Function

' Colonnes qui rendent une ligne « non vide » : toutes les colonnes de saisie, pour ne jamais
' perdre une ligne commencée par l'utilisateur lors de la réécriture du tableau.
Public Function ColonnesTestPrev(cP As Object) As Variant
    ColonnesTestPrev = Array(C(cP, "ID"), C(cP, "Date prévue"), C(cP, "Banque"), C(cP, "Libellé"), C(cP, "Catégorie"), _
                             C(cP, "Type"), C(cP, "Montant"), C(cP, "Commentaire"), C(cP, "Statut"))
End Function

Public Function ColonnesTestCharges(cC As Object) As Variant
    ColonnesTestCharges = Array(C(cC, "ID"), C(cC, "Libellé"), C(cC, "Banque"), C(cC, "Catégorie"), C(cC, "Type"), _
                                C(cC, "Montant"), C(cC, "Première échéance"), C(cC, "Fréquence"), C(cC, "Date de fin"), _
                                C(cC, "Commentaire"))
End Function

' Réécrit une table : lignes non vides (dans l'ordre), puis lignes ajoutées, puis lignes libres
Public Sub ReecrireAvecLignesLibres(lo As ListObject, data As Variant, ByVal n As Long, colonnesTest As Variant, _
                                    Optional ajouts As Variant, Optional ByVal nAjouts As Long = 0)
    Dim nbCol As Long, i As Long, k As Long, total() As Variant
    nbCol = lo.ListColumns.Count
    ReDim total(1 To n + nAjouts + NB_LIGNES_LIBRES, 1 To nbCol)
    For i = 1 To n
        If Not LigneBlanche(data, i, colonnesTest) Then
            k = k + 1
            CopierLigne data, i, total, k
        End If
    Next i
    For i = 1 To nAjouts
        k = k + 1
        CopierLigne ajouts, i, total, k
    Next i
    EcrireTable lo, total, k + NB_LIGNES_LIBRES
End Sub

'-------------------------------------------------------------------------------
' Complétion des saisies (ID, valeurs par défaut, montants positifs)
'-------------------------------------------------------------------------------
Public Sub CompleterSaisiesInterne()
    Dim lo As ListObject, cP As Object, dP As Variant, nP As Long, i As Long, mx As Long
    Dim cC As Object, dC As Variant, nC As Long, montant As Double

    Set lo = T(T_PREV)
    Set cP = Colonnes(lo)
    dP = LireTable(lo)
    nP = NbLignes(dP)
    mx = MaxNumero(dP, nP, C(cP, "ID"), "P-")
    For i = 1 To nP
        If Not LigneBlanche(dP, i, ColonnesTestPrev(cP)) Then
            If Texte(dP(i, C(cP, "ID"))) = "" Then
                mx = mx + 1
                dP(i, C(cP, "ID")) = FormatId("P-", mx, 6)
            End If
            If Texte(dP(i, C(cP, "Origine"))) = "" Then dP(i, C(cP, "Origine")) = ORI_SAISIE
            If Texte(dP(i, C(cP, "Statut"))) = "" Then dP(i, C(cP, "Statut")) = P_PREVU
            If Texte(dP(i, C(cP, "Type"))) = "" Then dP(i, C(cP, "Type")) = TY_DEPENSE
            montant = Nombre(dP(i, C(cP, "Montant")))
            If montant < 0 Then dP(i, C(cP, "Montant")) = Abs(montant)
        End If
    Next i
    ReecrireAvecLignesLibres lo, dP, nP, ColonnesTestPrev(cP)

    Set lo = T(T_CHARGES)
    Set cC = Colonnes(lo)
    dC = LireTable(lo)
    nC = NbLignes(dC)
    mx = MaxNumero(dC, nC, C(cC, "ID"), "CF-")
    For i = 1 To nC
        If Not LigneBlanche(dC, i, ColonnesTestCharges(cC)) Then
            If Texte(dC(i, C(cC, "ID"))) = "" Then
                mx = mx + 1
                dC(i, C(cC, "ID")) = FormatId("CF-", mx, 3)
            End If
            If Texte(dC(i, C(cC, "Actif"))) = "" Then dC(i, C(cC, "Actif")) = "Oui"
            If Texte(dC(i, C(cC, "Type"))) = "" Then dC(i, C(cC, "Type")) = TY_DEPENSE
            If Texte(dC(i, C(cC, "Fréquence"))) = "" Then dC(i, C(cC, "Fréquence")) = "Mensuelle"
            montant = Nombre(dC(i, C(cC, "Montant")))
            If montant < 0 Then dC(i, C(cC, "Montant")) = Abs(montant)
        End If
    Next i
    ReecrireAvecLignesLibres lo, dC, nC, ColonnesTestCharges(cC)
End Sub

'-------------------------------------------------------------------------------
' Formulaire de saisie
'-------------------------------------------------------------------------------
Public Sub AjouterOperation()
    Dim d As Double, banque As String, dest As String, libelle As String, categorie As String
    Dim montant As Double, typeOp As String, commentaire As String, msgErr As String
    Dim lo As ListObject, cP As Object, dP As Variant, nP As Long, mx As Long
    Dim ajouts() As Variant, nAjouts As Long, id1 As String, id2 As String

    On Error GoTo Erreur
    d = DateVal(ValeurNom("Form_Date"))
    banque = Texte(ValeurNom("Form_Banque"))
    dest = Texte(ValeurNom("Form_BanqueDest"))
    libelle = Texte(ValeurNom("Form_Libelle"))
    categorie = Texte(ValeurNom("Form_Categorie"))
    montant = Abs(Nombre(ValeurNom("Form_Montant")))
    typeOp = Texte(ValeurNom("Form_Type"))
    commentaire = Texte(ValeurNom("Form_Commentaire"))

    If d = 0 Then msgErr = msgErr & "- la date prévue est obligatoire" & vbCrLf
    If Not BanqueConnue(banque) Then msgErr = msgErr & "- choisissez une banque dans la liste" & vbCrLf
    If libelle = "" Then msgErr = msgErr & "- le libellé est obligatoire" & vbCrLf
    If montant = 0 Then msgErr = msgErr & "- le montant doit être supérieur à zéro" & vbCrLf
    If typeOp <> TY_DEPENSE And typeOp <> TY_RECETTE And typeOp <> TY_VIR_FORM Then
        msgErr = msgErr & "- choisissez le type : Dépense, Recette ou Virement interne" & vbCrLf
    End If
    If typeOp = TY_VIR_FORM Then
        If Not BanqueConnue(dest) Or dest = banque Then
            msgErr = msgErr & "- pour un virement interne, choisissez une banque de destination différente" & vbCrLf
        End If
    End If
    If msgErr = "" And d > 0 And BanqueConnue(banque) Then
        If d <= DateSoldeBanque(banque) Then
            msgErr = msgErr & "- la date est antérieure ou égale à la date du solde initial de " & banque & vbCrLf
        End If
    End If
    If msgErr <> "" Then
        Afficher "L'opération n'a pas été ajoutée :" & vbCrLf & vbCrLf & msgErr, vbExclamation, "Saisie"
        Exit Sub
    End If

    DebutTraitement
    Set lo = T(T_PREV)
    Set cP = Colonnes(lo)
    dP = LireTable(lo)
    nP = NbLignes(dP)
    mx = MaxNumero(dP, nP, C(cP, "ID"), "P-")
    ReDim ajouts(1 To 2, 1 To lo.ListColumns.Count)

    id1 = FormatId("P-", mx + 1, 6)
    ajouts(1, C(cP, "ID")) = id1
    ajouts(1, C(cP, "Date prévue")) = d
    ajouts(1, C(cP, "Banque")) = banque
    ajouts(1, C(cP, "Libellé")) = libelle
    ajouts(1, C(cP, "Catégorie")) = categorie
    ajouts(1, C(cP, "Montant")) = montant
    ajouts(1, C(cP, "Commentaire")) = commentaire
    ajouts(1, C(cP, "Statut")) = P_PREVU
    ajouts(1, C(cP, "Origine")) = ORI_SAISIE
    nAjouts = 1
    If typeOp = TY_VIR_FORM Then
        id2 = FormatId("P-", mx + 2, 6)
        If categorie = "" Then ajouts(1, C(cP, "Catégorie")) = CAT_VIREMENT
        ajouts(1, C(cP, "Type")) = TY_VIR_SORTANT
        ajouts(1, C(cP, "Lien virement")) = id2
        CopierLigne ajouts, 1, ajouts, 2
        ajouts(2, C(cP, "ID")) = id2
        ajouts(2, C(cP, "Banque")) = dest
        ajouts(2, C(cP, "Type")) = TY_VIR_ENTRANT
        ajouts(2, C(cP, "Lien virement")) = id1
        nAjouts = 2
    Else
        ajouts(1, C(cP, "Type")) = typeOp
    End If
    ReecrireAvecLignesLibres lo, dP, nP, ColonnesTestPrev(cP), ajouts, nAjouts

    EcrireNom "Form_Libelle", Empty
    EcrireNom "Form_Montant", Empty
    EcrireNom "Form_Commentaire", Empty
    EcrireNom "Form_Categorie", Empty
    FinTraitement
    Application.Calculate
    If nAjouts = 2 Then
        AfficherStatut "Virement interne ajouté (" & id1 & " / " & id2 & ")."
    Else
        AfficherStatut "Opération ajoutée (" & id1 & ")."
    End If
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "l'ajout de l'opération", msgErr
End Sub

Public Sub EffacerFormulaire()
    On Error GoTo Erreur
    DebutTraitement
    EcrireNom "Form_Libelle", Empty
    EcrireNom "Form_Montant", Empty
    EcrireNom "Form_Commentaire", Empty
    EcrireNom "Form_Categorie", Empty
    EcrireNom "Form_BanqueDest", Empty
    FinTraitement
    Exit Sub
Erreur:
    TerminerTraitement
End Sub

Public Sub ValiderSaisies()
    Dim msgErr As String
    On Error GoTo Erreur
    DebutTraitement "Validation des saisies..."
    CompleterSaisiesInterne
    FinTraitement
    Application.Calculate
    Afficher "Les lignes saisies ont été complétées (identifiants, valeurs par défaut).", vbInformation, "Saisie"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la validation des saisies", msgErr
End Sub

'-------------------------------------------------------------------------------
' Suppression et tri
'-------------------------------------------------------------------------------
Public Sub SupprimerLigne()
    Dim lo As ListObject, cP As Object, dP As Variant, nP As Long, r As Long, i As Long
    Dim id As String, lien As String, rep As VbMsgBoxResult, msgErr As String, rLien As Long
    Dim garder() As Boolean, total() As Variant, k As Long

    On Error GoTo Erreur
    Set lo = T(T_PREV)
    If ActiveSheet.Name <> lo.Parent.Name Or lo.DataBodyRange Is Nothing Then GoTo HorsTable
    If Intersect(ActiveCell, lo.DataBodyRange) Is Nothing Then GoTo HorsTable
    r = ActiveCell.Row - lo.HeaderRowRange.Row
    Set cP = Colonnes(lo)
    dP = LireTable(lo)
    nP = NbLignes(dP)
    If LigneBlanche(dP, r, ColonnesTestPrev(cP)) Then Exit Sub
    id = Texte(dP(r, C(cP, "ID")))

    If Texte(dP(r, C(cP, "Statut"))) = P_POINTE Then
        Afficher "La prévision " & id & " est pointée avec une opération bancaire." & vbCrLf & _
               "Annulez d'abord le rapprochement (bouton « Annuler le rapprochement »).", vbExclamation, "Suppression"
        Exit Sub
    End If

    DebutTraitement
    If Texte(dP(r, C(cP, "Origine"))) = ORI_CHARGE Then
        rep = Afficher("Cette échéance provient de la charge fixe " & Texte(dP(r, C(cP, "Réf charge"))) & "." & vbCrLf & _
                     "Elle va être marquée « Annulé » (elle ne sera pas recréée)." & vbCrLf & vbCrLf & _
                     "Pour arrêter toute la série, modifiez la charge fixe (Actif = Non ou date de fin).", _
                     vbOKCancel + vbQuestion, "Suppression")
        If rep = vbOK Then
            dP(r, C(cP, "Statut")) = P_ANNULE
            dP(r, C(cP, "ID proposé")) = ""
            EcrireTable lo, dP, nP
        End If
        FinTraitement
        Exit Sub
    End If

    ReDim garder(1 To nP)
    For i = 1 To nP
        garder(i) = True
    Next i
    garder(r) = False
    lien = Texte(dP(r, C(cP, "Lien virement")))
    If lien <> "" Then
        For i = 1 To nP
            If Texte(dP(i, C(cP, "ID"))) = lien Then rLien = i
        Next i
        If rLien > 0 Then
            If Texte(dP(rLien, C(cP, "Statut"))) = P_POINTE Then
                FinTraitement
                Afficher "La contrepartie " & lien & " de ce virement est pointée : annulez d'abord son rapprochement.", vbExclamation, "Suppression"
                Exit Sub
            End If
            garder(rLien) = False
        End If
        rep = Afficher("Supprimer le virement interne " & id & " et sa contrepartie " & lien & " ?", vbOKCancel + vbQuestion, "Suppression")
    Else
        rep = Afficher("Supprimer définitivement la prévision " & id & " (" & Texte(dP(r, C(cP, "Libellé"))) & ") ?", _
                     vbOKCancel + vbQuestion, "Suppression")
    End If
    If rep <> vbOK Then
        FinTraitement
        Exit Sub
    End If

    ReDim total(1 To nP, 1 To lo.ListColumns.Count)
    For i = 1 To nP
        If garder(i) Then
            k = k + 1
            CopierLigne dP, i, total, k
        End If
    Next i
    ReecrireAvecLignesLibres lo, total, k, ColonnesTestPrev(cP)
    modRapprochement.RechercherInterne
    FinTraitement
    Application.Calculate
    Exit Sub
HorsTable:
    Afficher "Sélectionnez d'abord une cellule de la ligne à supprimer dans le tableau des prévisions.", vbInformation, "Suppression"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la suppression", msgErr
End Sub

Public Sub TrierPrevisions()
    Dim lo As ListObject, cP As Object, dP As Variant, nP As Long, i As Long, k As Long
    Dim cles() As String, idx() As Long, total() As Variant, msgErr As String

    On Error GoTo Erreur
    DebutTraitement "Tri des prévisions..."
    Set lo = T(T_PREV)
    Set cP = Colonnes(lo)
    dP = LireTable(lo)
    nP = NbLignes(dP)
    ReDim cles(1 To nP)
    ReDim idx(1 To nP)
    For i = 1 To nP
        idx(i) = i
        If LigneBlanche(dP, i, ColonnesTestPrev(cP)) Then
            cles(i) = "9" & Format$(i, "000000")
        Else
            cles(i) = "0" & CleDate(DateVal(dP(i, C(cP, "Date prévue")))) & Texte(dP(i, C(cP, "ID"))) & Format$(i, "000000")
        End If
    Next i
    TrierIndices cles, idx, 1, nP
    ReDim total(1 To nP, 1 To lo.ListColumns.Count)
    For i = 1 To nP
        k = k + 1
        CopierLigne dP, idx(i), total, k
    Next i
    ReecrireAvecLignesLibres lo, total, k, ColonnesTestPrev(cP)
    FinTraitement
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "le tri", msgErr
End Sub

'-------------------------------------------------------------------------------
' Charges fixes : génération des échéances
'-------------------------------------------------------------------------------
Public Sub GenererEcheances()
    Dim nbCreees As Long, nbSupprimees As Long, msgErr As String
    On Error GoTo Erreur
    DebutTraitement "Génération des échéances..."
    CompleterSaisiesInterne
    GenererEcheancesInterne nbCreees, nbSupprimees
    modRapprochement.RechercherInterne
    FinTraitement
    Application.Calculate
    Afficher "Échéances des charges fixes mises à jour." & vbCrLf & _
           nbCreees & " échéance(s) créée(s), " & nbSupprimees & " échéance(s) future(s) supprimée(s)." & vbCrLf & vbCrLf & _
           "Les échéances passées, pointées, annulées ou modifiées à la main ne sont jamais modifiées.", _
           vbInformation, "Charges fixes"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la génération des échéances", msgErr
End Sub

Private Function PasFrequence(ByVal frequence As String) As Long
    Select Case frequence
        Case "Mensuelle": PasFrequence = 1
        Case "Trimestrielle": PasFrequence = 3
        Case "Semestrielle": PasFrequence = 6
        Case "Annuelle": PasFrequence = 12
        Case Else: PasFrequence = 0
    End Select
End Function

Private Function DemiCreneau(ByVal pas As Long) As Long
    Select Case pas
        Case 1: DemiCreneau = 14
        Case 3: DemiCreneau = 44
        Case 6: DemiCreneau = 89
        Case Else: DemiCreneau = 181
    End Select
End Function

Private Function Couverte(ByVal listeDates As String, ByVal d As Double, ByVal demi As Long) As Boolean
    Dim morceaux() As String, k As Long
    If listeDates = "" Then Exit Function
    morceaux = Split(listeDates, ";")
    For k = LBound(morceaux) To UBound(morceaux)
        If morceaux(k) <> "" Then
            If Abs(CDbl(morceaux(k)) - d) <= demi Then
                Couverte = True
                Exit Function
            End If
        End If
    Next k
End Function

' Génération sans effet rétroactif :
'  - une échéance figée (pointée, annulée, proposée, passée ou modifiée à la main)
'    n'est jamais modifiée ni supprimée et « couvre » son créneau ;
'  - une échéance non figée est conservée (même ID) si elle correspond exactement
'    à une échéance théorique de la définition actuelle de la charge, sinon supprimée ;
'  - les échéances théoriques manquantes sont créées.
Public Sub GenererEcheancesInterne(Optional ByRef nbCreees As Long, Optional ByRef nbSupprimees As Long)
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long
    Dim loC As ListObject, cC As Object, dC As Variant, nC As Long
    Dim refDate As Double, tolerance As Long, finHorizon As Double, debutFenetre As Double
    Dim garder() As Boolean, figees As Object, libres As Object, i As Long, j As Long, r As Long
    Dim refCharge As String, d As Double, estFigee As Boolean
    Dim idCharge As String, banque As String, libelle As String, typeOp As String, categorie As String
    Dim montant As Double, premiere As Double, fin As Double, finCharge As Double, debut As Double
    Dim pas As Long, demi As Long, k As Long, empTheo As String, trouve As Boolean
    Dim candidats() As String, mx As Long, nbCol As Long
    Dim ajouts() As Variant, nAjouts As Long, total() As Variant, nTotal As Long

    nbCreees = 0
    nbSupprimees = 0
    Set loP = T(T_PREV)
    Set cP = Colonnes(loP)
    dP = LireTable(loP)
    nP = NbLignes(dP)
    Set loC = T(T_CHARGES)
    Set cC = Colonnes(loC)
    dC = LireTable(loC)
    nC = NbLignes(dC)
    nbCol = loP.ListColumns.Count

    refDate = DateReference()
    tolerance = CLng(Nombre(ValeurNom("Tolerance_Jours")))
    finHorizon = AjouterMois(refDate, CLng(Nombre(ValeurNom("Horizon_Mois"))))
    debutFenetre = refDate - tolerance

    ' 1. Classement des échéances existantes
    Set figees = CreateObject("Scripting.Dictionary")
    Set libres = CreateObject("Scripting.Dictionary")
    ReDim garder(1 To IIf(nP < 1, 1, nP))
    For i = 1 To nP
        garder(i) = True
        refCharge = Texte(dP(i, C(cP, "Réf charge")))
        If Texte(dP(i, C(cP, "Origine"))) = ORI_CHARGE And refCharge <> "" Then
            d = DateVal(dP(i, C(cP, "Date prévue")))
            estFigee = False
            If Texte(dP(i, C(cP, "Statut"))) = P_POINTE Or Texte(dP(i, C(cP, "Statut"))) = P_ANNULE Then estFigee = True
            If Texte(dP(i, C(cP, "ID opération"))) <> "" Or Texte(dP(i, C(cP, "ID proposé"))) <> "" Then estFigee = True
            If d < refDate Then estFigee = True
            If Texte(dP(i, C(cP, "Empreinte"))) <> Empreinte(d, Texte(dP(i, C(cP, "Banque"))), _
                    Texte(dP(i, C(cP, "Libellé"))), Texte(dP(i, C(cP, "Type"))), Nombre(dP(i, C(cP, "Montant")))) Then
                estFigee = True
            End If
            If estFigee Then
                figees(refCharge) = figees(refCharge) & CStr(d) & ";"
            Else
                libres(refCharge) = libres(refCharge) & CStr(i) & ";"
                garder(i) = False
            End If
        End If
    Next i

    ' 2. Échéances théoriques de chaque charge active
    mx = MaxNumero(dP, nP, C(cP, "ID"), "P-")
    ReDim ajouts(1 To nbCol, 1 To 1)
    For j = 1 To nC
        idCharge = Texte(dC(j, C(cC, "ID")))
        pas = PasFrequence(Texte(dC(j, C(cC, "Fréquence"))))
        montant = Abs(Nombre(dC(j, C(cC, "Montant"))))
        banque = Texte(dC(j, C(cC, "Banque")))
        premiere = DateVal(dC(j, C(cC, "Première échéance")))
        If idCharge <> "" And EstOui(dC(j, C(cC, "Actif"))) And pas > 0 And montant > 0 _
           And BanqueConnue(banque) And premiere > 0 Then
            libelle = Texte(dC(j, C(cC, "Libellé")))
            typeOp = Texte(dC(j, C(cC, "Type")))
            If typeOp <> TY_RECETTE Then typeOp = TY_DEPENSE
            categorie = Texte(dC(j, C(cC, "Catégorie")))
            fin = DateVal(dC(j, C(cC, "Date de fin")))
            demi = DemiCreneau(pas)
            debut = premiere
            If DateSoldeBanque(banque) + 1 > debut Then debut = DateSoldeBanque(banque) + 1
            If debutFenetre > debut Then debut = debutFenetre
            finCharge = finHorizon
            If fin > 0 And fin < finCharge Then finCharge = fin
            k = 0
            Do
                d = AjouterMois(premiere, k * pas)
                k = k + 1
                If d > finCharge Then Exit Do
                If d >= debut Then
                    If Not Couverte(CStr(figees(idCharge)), d, demi) Then
                        empTheo = Empreinte(d, banque, libelle, typeOp, montant)
                        trouve = False
                        If CStr(libres(idCharge)) <> "" Then
                            candidats = Split(CStr(libres(idCharge)), ";")
                            For r = LBound(candidats) To UBound(candidats)
                                If candidats(r) <> "" Then
                                    i = CLng(candidats(r))
                                    If Not garder(i) Then
                                        If Texte(dP(i, C(cP, "Empreinte"))) = empTheo Then
                                            garder(i) = True
                                            dP(i, C(cP, "Catégorie")) = categorie
                                            trouve = True
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next r
                        End If
                        If Not trouve Then
                            nAjouts = nAjouts + 1
                            ReDim Preserve ajouts(1 To nbCol, 1 To nAjouts)
                            mx = mx + 1
                            ajouts(C(cP, "ID"), nAjouts) = FormatId("P-", mx, 6)
                            ajouts(C(cP, "Date prévue"), nAjouts) = d
                            ajouts(C(cP, "Banque"), nAjouts) = banque
                            ajouts(C(cP, "Libellé"), nAjouts) = libelle
                            ajouts(C(cP, "Catégorie"), nAjouts) = categorie
                            ajouts(C(cP, "Type"), nAjouts) = typeOp
                            ajouts(C(cP, "Montant"), nAjouts) = montant
                            ajouts(C(cP, "Statut"), nAjouts) = P_PREVU
                            ajouts(C(cP, "Origine"), nAjouts) = ORI_CHARGE
                            ajouts(C(cP, "Réf charge"), nAjouts) = idCharge
                            ajouts(C(cP, "Empreinte"), nAjouts) = empTheo
                        End If
                    End If
                End If
            Loop
        End If
    Next j

    ' 3. Réécriture : lignes conservées + nouvelles échéances + lignes libres
    For i = 1 To nP
        If Not garder(i) Then nbSupprimees = nbSupprimees + 1
    Next i
    nbCreees = nAjouts
    ReDim total(1 To nP + nAjouts + 1, 1 To nbCol)
    For i = 1 To nP
        If garder(i) Then
            nTotal = nTotal + 1
            CopierLigne dP, i, total, nTotal
        End If
    Next i
    For r = 1 To nAjouts
        nTotal = nTotal + 1
        For j = 1 To nbCol
            total(nTotal, j) = ajouts(j, r)
        Next j
    Next r
    ReecrireAvecLignesLibres loP, total, nTotal, ColonnesTestPrev(cP)
End Sub
