Attribute VB_Name = "modATraiter"
'===============================================================================
' modATraiter - « Nouvelles opérations bancaires à traiter »
'-------------------------------------------------------------------------------
' Opérations bancaires sans prévision correspondante. Pour chacune, l'utilisateur
' choisit une catégorie, un commentaire et une action :
'   - Conserver comme non prévue (dépense / recette réelle non prévue)
'   - Virement interne
'   - Rapprocher avec la prévision indiquée (montant différent, rapprochement manuel)
'   - Créer charge mensuelle / trimestrielle / semestrielle / annuelle
' puis clique « Appliquer les traitements ».
' Miroir de reference/moteur.py : a_traiter / suggestion_virement / traiter.
'===============================================================================
Option Explicit
Option Compare Binary

Public Function ActionsEnAttente() As Long
    Dim lo As ListObject, cA As Object, dA As Variant, i As Long
    Set lo = T(T_ATRAITER)
    Set cA = Colonnes(lo)
    dA = LireTable(lo)
    For i = 1 To NbLignes(dA)
        If Texte(dA(i, C(cA, "Action"))) <> "" And Texte(dA(i, C(cA, "ID opération"))) <> "" Then
            ActionsEnAttente = ActionsEnAttente + 1
        End If
    Next i
End Function

'-------------------------------------------------------------------------------
' Construction de la liste
'-------------------------------------------------------------------------------
Public Sub ActualiserListe()
    Dim n As Long, msgErr As String
    On Error GoTo Erreur
    DebutTraitement "Mise à jour de la liste..."
    n = ConstruireListeATraiterInterne()
    FinTraitement
    AfficherStatut n & " opération(s) à traiter."
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "la mise à jour de la liste des opérations à traiter", msgErr
End Sub

' Chaque mot du mot-clé doit commencer un mot du libellé (« COTIS » couvre « COTISATION »)
Public Function RegleCorrespond(ByVal motCle As String, ByVal libelle As String) As Boolean
    Dim mk() As String, ml() As String, a As Long, b As Long, trouve As Boolean, nk As String, nl As String
    nk = NormaliserTexte(motCle)
    nl = NormaliserTexte(libelle)
    If nk = "" Or nl = "" Then Exit Function
    mk = Split(nk, " ")
    ml = Split(nl, " ")
    For a = LBound(mk) To UBound(mk)
        trouve = False
        For b = LBound(ml) To UBound(ml)
            If Left$(ml(b), Len(mk(a))) = mk(a) Then
                trouve = True
                Exit For
            End If
        Next b
        If Not trouve Then Exit Function
    Next a
    RegleCorrespond = True
End Function

' Indice de la règle la plus précise (plus grand nombre de mots) qui s'applique au libellé, 0 sinon
Public Function TrouverRegle(ByVal libelle As String, dRg As Variant, ByVal nRg As Long, cRg As Object) As Long
    Dim i As Long, nb As Long, nbMax As Long, mot As String
    For i = 1 To nRg
        mot = NormaliserTexte(Texte(dRg(i, C(cRg, "Mot-clé"))))
        If mot <> "" And Texte(dRg(i, C(cRg, "Catégorie"))) <> "" Then
            nb = UBound(Split(mot, " ")) + 1
            If nb > nbMax Then
                If RegleCorrespond(mot, libelle) Then
                    TrouverRegle = i
                    nbMax = nb
                End If
            End If
        End If
    Next i
End Function

' Mot-clé appris : 3 premiers mots significatifs sans chiffre (références, dates exclues)
Public Function MotCleAppris(ByVal libelle As String) As String
    Dim mots() As String, n As Long, k As Long, res As String, nb As Long
    n = MotsSignificatifs(libelle, mots)
    For k = 0 To n - 1
        If Not (mots(k) Like "*#*") Then
            If res <> "" Then res = res & " "
            res = res & mots(k)
            nb = nb + 1
            If nb = 3 Then Exit For
        End If
    Next k
    MotCleAppris = res
End Function

Private Function EstVirement(ByVal libelle As String) As Boolean
    Dim mots() As String, k As Long, nl As String
    nl = NormaliserTexte(libelle)
    If nl = "" Then Exit Function
    mots = Split(nl, " ")
    For k = LBound(mots) To UBound(mots)
        If Left$(mots(k), 3) = "VIR" Then
            EstVirement = True
            Exit Function
        End If
    Next k
End Function

' Prévision en attente, même banque et même sens, dans la tolérance, au libellé similaire mais de
' montant différent : rapprochement manuel probable (jamais de traitement automatique dans ce cas)
Public Function PrevisionProche(dB As Variant, cB As Object, ByVal i As Long, dP As Variant, ByVal nP As Long, _
                                cP As Object, ByVal tolerance As Long, ByVal seuil As Double) As String
    Dim j As Long, banque As String, mtOp As Double, mtP As Double, dPrev As Double, ecart As Long, sim As Double
    Dim cle As String, cleMin As String, statutP As String
    banque = Texte(dB(i, C(cB, "Banque")))
    mtOp = Nombre(dB(i, C(cB, "Montant")))
    For j = 1 To nP
        statutP = Texte(dP(j, C(cP, "Statut")))
        dPrev = DateVal(dP(j, C(cP, "Date prévue")))
        If Texte(dP(j, C(cP, "ID"))) <> "" And Texte(dP(j, C(cP, "Banque"))) = banque And dPrev > DateSoldeBanque(banque) _
           And statutP <> P_POINTE And statutP <> P_ANNULE And Texte(dP(j, C(cP, "ID opération"))) = "" _
           And Nombre(dP(j, C(cP, "Montant"))) <> 0 Then
            mtP = modPrevisions.MontantSigne(Texte(dP(j, C(cP, "Type"))), Nombre(dP(j, C(cP, "Montant"))))
            If (mtP < 0) = (mtOp < 0) And Cents(mtP) <> Cents(mtOp) Then
                ecart = CLng(Abs(DateVal(dB(i, C(cB, "Date"))) - dPrev))
                If ecart <= tolerance Then
                    sim = Similarite(Texte(dB(i, C(cB, "Libellé"))), Texte(dP(j, C(cP, "Libellé"))))
                    If sim >= seuil Then
                        cle = Format$(10000 - CLng(Round(sim * 10000, 0)), "00000") & Format$(ecart, "000") & _
                              Texte(dP(j, C(cP, "ID")))
                        If cleMin = "" Or StrComp(cle, cleMin, vbBinaryCompare) < 0 Then
                            cleMin = cle
                            PrevisionProche = Texte(dP(j, C(cP, "ID")))
                        End If
                    End If
                End If
            End If
        End If
    Next j
End Function

' Opérations « à traiter » (postérieures au solde initial) triées par date puis ID
Private Function ListeATraiter(dB As Variant, ByVal nB As Long, cB As Object, liste() As Long) As Long
    Dim i As Long, n As Long, banque As String, cles() As String, idx() As Long, copie() As Long, k As Long
    If nB = 0 Then Exit Function
    ReDim liste(1 To nB)
    ReDim cles(1 To nB)
    For i = 1 To nB
        banque = Texte(dB(i, C(cB, "Banque")))
        If Texte(dB(i, C(cB, "Statut"))) = OB_A_TRAITER And Texte(dB(i, C(cB, "ID"))) <> "" And BanqueConnue(banque) Then
            If DateVal(dB(i, C(cB, "Date"))) > DateSoldeBanque(banque) Then
                n = n + 1
                liste(n) = i
                cles(n) = CleDate(DateVal(dB(i, C(cB, "Date")))) & Texte(dB(i, C(cB, "ID")))
            End If
        End If
    Next i
    If n > 1 Then
        ReDim idx(1 To n)
        ReDim copie(1 To n)
        For k = 1 To n
            idx(k) = k
            copie(k) = liste(k)
        Next k
        TrierIndices cles, idx, 1, n
        For k = 1 To n
            liste(k) = copie(idx(k))
        Next k
    End If
    ListeATraiter = n
End Function

' Après la recherche des correspondances : virements internes non prévus, puis règles automatiques
Public Sub TraitementsAutomatiquesInterne(Optional ByRef nbVirements As Long, Optional ByRef nbRegles As Long)
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long
    Dim loRg As ListObject, cRg As Object, dRg As Variant, nRg As Long
    Dim liste() As Long, n As Long, k As Long, q As Long, i As Long, x As Long, meilleur As Long, r As Long
    Dim ecartCertain As Long, tolerance As Long, seuil As Double, ecart As Long, cle As String, cleMin As String
    Dim action As String, modifie As Boolean, op As Variant

    nbVirements = 0
    nbRegles = 0
    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    n = ListeATraiter(dB, nB, cB, liste)
    If n = 0 Then Exit Sub
    ecartCertain = CLng(Nombre(ValeurNom("Ecart_Certain")))
    tolerance = CLng(Nombre(ValeurNom("Tolerance_Jours")))
    seuil = Nombre(ValeurNom("Seuil_Similarite"))

    If EstOui(ValeurNom("Virements_Auto")) Then
        For k = 1 To n
            i = liste(k)
            If Texte(dB(i, C(cB, "Statut"))) = OB_A_TRAITER And EstVirement(Texte(dB(i, C(cB, "Libellé")))) Then
                meilleur = 0
                cleMin = ""
                For q = 1 To n
                    x = liste(q)
                    If x <> i And Texte(dB(x, C(cB, "Banque"))) <> Texte(dB(i, C(cB, "Banque"))) _
                       And Texte(dB(x, C(cB, "Statut"))) = OB_A_TRAITER And EstVirement(Texte(dB(x, C(cB, "Libellé")))) Then
                        If Cents(Nombre(dB(x, C(cB, "Montant")))) = -Cents(Nombre(dB(i, C(cB, "Montant")))) Then
                            ecart = CLng(Abs(DateVal(dB(x, C(cB, "Date"))) - DateVal(dB(i, C(cB, "Date")))))
                            If ecart <= ecartCertain Then
                                cle = Format$(ecart, "000") & Texte(dB(x, C(cB, "ID")))
                                If cleMin = "" Or StrComp(cle, cleMin, vbBinaryCompare) < 0 Then
                                    cleMin = cle
                                    meilleur = x
                                End If
                            End If
                        End If
                    End If
                Next q
                If meilleur > 0 Then
                    For Each op In Array(i, meilleur)
                        dB(op, C(cB, "Statut")) = OB_VIREMENT
                        If Texte(dB(op, C(cB, "Catégorie"))) = "" Then dB(op, C(cB, "Catégorie")) = CAT_VIREMENT
                        dB(op, C(cB, "Mode")) = MODE_AUTO
                    Next op
                    nbVirements = nbVirements + 1
                    modifie = True
                End If
            End If
        Next k
    End If

    If EstOui(ValeurNom("Traitement_Auto")) Then
        Set loRg = T(T_REGLES)
        Set cRg = Colonnes(loRg)
        dRg = LireTable(loRg)
        nRg = NbLignes(dRg)
        Set loP = T(T_PREV)
        Set cP = Colonnes(loP)
        dP = LireTable(loP)
        nP = NbLignes(dP)
        For k = 1 To n
            i = liste(k)
            If Texte(dB(i, C(cB, "Statut"))) = OB_A_TRAITER Then
                r = TrouverRegle(Texte(dB(i, C(cB, "Libellé"))), dRg, nRg, cRg)
                If r > 0 Then
                    action = Texte(dRg(r, C(cRg, "Action automatique")))
                    If action = ACT_CONSERVER Or action = ACT_VIREMENT Then
                        If PrevisionProche(dB, cB, i, dP, nP, cP, tolerance, seuil) = "" Then
                            If action = ACT_CONSERVER Then
                                dB(i, C(cB, "Statut")) = OB_NON_PREVUE
                            Else
                                dB(i, C(cB, "Statut")) = OB_VIREMENT
                            End If
                            dB(i, C(cB, "Catégorie")) = Texte(dRg(r, C(cRg, "Catégorie")))
                            dB(i, C(cB, "Mode")) = MODE_AUTO
                            nbRegles = nbRegles + 1
                            modifie = True
                        End If
                    End If
                End If
            End If
        Next k
    End If
    If modifie Then EcrireTable loB, dB, nB
End Sub

' Reconstruit la liste. Les catégories et commentaires déjà saisis sont d'abord
' enregistrés dans le journal bancaire pour ne rien perdre.
Public Function ConstruireListeATraiterInterne() As Long
    Dim loA As ListObject, cA As Object, dA As Variant, nA As Long
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long, indexB As Object
    Dim loRg As ListObject, cRg As Object, dRg As Variant, nRg As Long
    Dim i As Long, j As Long, k As Long, n As Long, id As String, modifie As Boolean
    Dim cles() As String, idx() As Long, liste() As Long, sortie() As Variant
    Dim tolerance As Long, meilleur As Long, meilleurCle As String, cleJ As String, banque As String
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long, seuil As Double, proche As String, r As Long

    Set loA = T(T_ATRAITER)
    Set cA = Colonnes(loA)
    dA = LireTable(loA)
    nA = NbLignes(dA)
    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    Set loRg = T(T_REGLES)
    Set cRg = Colonnes(loRg)
    dRg = LireTable(loRg)
    nRg = NbLignes(dRg)
    tolerance = CLng(Nombre(ValeurNom("Tolerance_Jours")))
    seuil = Nombre(ValeurNom("Seuil_Similarite"))
    Set loP = T(T_PREV)
    Set cP = Colonnes(loP)
    dP = LireTable(loP)
    nP = NbLignes(dP)

    Set indexB = CreateObject("Scripting.Dictionary")
    For i = 1 To nB
        If Texte(dB(i, C(cB, "ID"))) <> "" Then indexB(Texte(dB(i, C(cB, "ID")))) = i
    Next i

    ' 1. Sauvegarde des brouillons (catégorie / commentaire)
    For k = 1 To nA
        id = Texte(dA(k, C(cA, "ID opération")))
        If indexB.Exists(id) Then
            i = indexB(id)
            If Texte(dB(i, C(cB, "Statut"))) = OB_A_TRAITER Then
                If Texte(dA(k, C(cA, "Catégorie"))) <> "" And Texte(dA(k, C(cA, "Catégorie"))) <> Texte(dB(i, C(cB, "Catégorie"))) Then
                    dB(i, C(cB, "Catégorie")) = Texte(dA(k, C(cA, "Catégorie")))
                    modifie = True
                End If
                If Texte(dA(k, C(cA, "Commentaire"))) <> "" And Texte(dA(k, C(cA, "Commentaire"))) <> Texte(dB(i, C(cB, "Commentaire"))) Then
                    dB(i, C(cB, "Commentaire")) = Texte(dA(k, C(cA, "Commentaire")))
                    modifie = True
                End If
            End If
        End If
    Next k
    If modifie Then EcrireTable loB, dB, nB

    ' 2. Opérations à traiter, triées par date puis ID
    If nB > 0 Then
        ReDim liste(1 To nB)
        ReDim cles(1 To nB)
    End If
    For i = 1 To nB
        banque = Texte(dB(i, C(cB, "Banque")))
        If Texte(dB(i, C(cB, "Statut"))) = OB_A_TRAITER And Texte(dB(i, C(cB, "ID"))) <> "" And BanqueConnue(banque) Then
            If DateVal(dB(i, C(cB, "Date"))) > DateSoldeBanque(banque) Then
                n = n + 1
                liste(n) = i
                cles(n) = CleDate(DateVal(dB(i, C(cB, "Date")))) & Texte(dB(i, C(cB, "ID")))
            End If
        End If
    Next i
    ReDim idx(1 To IIf(n < 1, 1, n))
    For k = 1 To n
        idx(k) = k
    Next k
    If n > 1 Then TrierIndices cles, idx, 1, n

    ReDim sortie(1 To IIf(n < 1, 1, n), 1 To loA.ListColumns.Count)
    For k = 1 To n
        i = liste(idx(k))
        sortie(k, C(cA, "ID opération")) = Texte(dB(i, C(cB, "ID")))
        sortie(k, C(cA, "Date")) = DateVal(dB(i, C(cB, "Date")))
        sortie(k, C(cA, "Banque")) = Texte(dB(i, C(cB, "Banque")))
        sortie(k, C(cA, "Libellé")) = Texte(dB(i, C(cB, "Libellé")))
        sortie(k, C(cA, "Montant")) = Nombre(dB(i, C(cB, "Montant")))
        sortie(k, C(cA, "Commentaire")) = Texte(dB(i, C(cB, "Commentaire")))
        If Texte(dB(i, C(cB, "Catégorie"))) <> "" Then
            sortie(k, C(cA, "Catégorie")) = Texte(dB(i, C(cB, "Catégorie")))
        Else
            r = TrouverRegle(Texte(dB(i, C(cB, "Libellé"))), dRg, nRg, cRg)
            If r > 0 Then sortie(k, C(cA, "Catégorie")) = Texte(dRg(r, C(cRg, "Catégorie")))
        End If
        proche = PrevisionProche(dB, cB, i, dP, nP, cP, tolerance, seuil)
        ' Virement interne probable : opération de montant opposé dans l'autre banque
        meilleur = 0
        meilleurCle = ""
        For j = 1 To n
            If liste(j) <> i Then
                If Texte(dB(liste(j), C(cB, "Banque"))) <> Texte(dB(i, C(cB, "Banque"))) Then
                    If Cents(Nombre(dB(liste(j), C(cB, "Montant")))) = -Cents(Nombre(dB(i, C(cB, "Montant")))) Then
                        If Abs(DateVal(dB(liste(j), C(cB, "Date"))) - DateVal(dB(i, C(cB, "Date")))) <= tolerance Then
                            cleJ = Format$(Abs(DateVal(dB(liste(j), C(cB, "Date"))) - DateVal(dB(i, C(cB, "Date")))), "000") & _
                                   Texte(dB(liste(j), C(cB, "ID")))
                            If meilleurCle = "" Or StrComp(cleJ, meilleurCle, vbBinaryCompare) < 0 Then
                                meilleurCle = cleJ
                                meilleur = liste(j)
                            End If
                        End If
                    End If
                End If
            End If
        Next j
        If proche <> "" Then
            sortie(k, C(cA, "Suggestion")) = "Prévision proche (montant différent) : " & proche
            sortie(k, C(cA, "ID prévision")) = proche
        ElseIf meilleur > 0 Then
            sortie(k, C(cA, "Suggestion")) = "Virement interne probable avec " & Texte(dB(meilleur, C(cB, "ID"))) & _
                " (" & Texte(dB(meilleur, C(cB, "Banque"))) & ")"
        End If
    Next k
    EcrireTable loA, sortie, n
    ConstruireListeATraiterInterne = n
End Function

'-------------------------------------------------------------------------------
' Application des traitements
'-------------------------------------------------------------------------------
Public Sub AppliquerTraitements()
    Dim resultat As String, msgErr As String
    On Error GoTo Erreur
    If ActionsEnAttente() = 0 Then
        Afficher "Aucune action n'a été choisie dans la colonne « Action ».", vbInformation, "À traiter"
        Exit Sub
    End If
    DebutTraitement "Application des traitements..."
    resultat = AppliquerTraitementsInterne()
    FinTraitement
    Application.Calculate
    Afficher resultat, vbInformation, "À traiter"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "l'application des traitements", msgErr
End Sub

Private Function FrequenceAction(ByVal action As String) As String
    Select Case LCase$(Trim$(Mid$(action, Len(ACT_CHARGE_PREFIXE) + 1)))
        Case "mensuelle": FrequenceAction = "Mensuelle"
        Case "trimestrielle": FrequenceAction = "Trimestrielle"
        Case "semestrielle": FrequenceAction = "Semestrielle"
        Case "annuelle": FrequenceAction = "Annuelle"
        Case Else: FrequenceAction = ""
    End Select
End Function

Private Function PasMois(ByVal frequence As String) As Long
    Select Case frequence
        Case "Mensuelle": PasMois = 1
        Case "Trimestrielle": PasMois = 3
        Case "Semestrielle": PasMois = 6
        Case Else: PasMois = 12
    End Select
End Function

' Mémorise un traitement manuel sous forme de règle (mot-clé -> catégorie, action). Renvoie 1 si une
' règle a été créée ou mise à jour. Miroir de reference/moteur.py : apprendre.
Private Function Apprendre(ByVal libelle As String, ByVal categorie As String, ByVal action As String, _
                           dRg As Variant, ByVal nRg As Long, cRg As Object, ajoutsRg As Variant, _
                           ByRef nAjoutsRg As Long, clesRg As Object) As Long
    Dim cle As String, pos As Variant
    If Not EstOui(ValeurNom("Apprendre_Regles")) Then Exit Function
    If categorie = "" Or categorie = CAT_NON_CATEGORISE Then Exit Function
    cle = MotCleAppris(libelle)
    If cle = "" Then Exit Function
    If clesRg.Exists(cle) Then
        pos = clesRg(cle)
        If pos > 0 Then
            dRg(pos, C(cRg, "Catégorie")) = categorie
            dRg(pos, C(cRg, "Action automatique")) = action
        Else
            ajoutsRg(-pos, C(cRg, "Catégorie")) = categorie
            ajoutsRg(-pos, C(cRg, "Action automatique")) = action
        End If
    Else
        nAjoutsRg = nAjoutsRg + 1
        ajoutsRg(nAjoutsRg, C(cRg, "Mot-clé")) = cle
        ajoutsRg(nAjoutsRg, C(cRg, "Catégorie")) = categorie
        ajoutsRg(nAjoutsRg, C(cRg, "Action automatique")) = action
        clesRg(cle) = -nAjoutsRg
    End If
    Apprendre = 1
End Function

Public Function AppliquerTraitementsInterne() As String
    Dim loA As ListObject, cA As Object, dA As Variant, nA As Long
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long, indexB As Object
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long, indexP As Object
    Dim loC As ListObject, cC As Object, dC As Variant, nC As Long, mxC As Long
    Dim ajoutsC() As Variant, nAjoutsC As Long, k As Long, i As Long, j As Long
    Dim id As String, action As String, categorie As String, commentaire As String, idPrev As String
    Dim nbOk As Long, erreurs As String, frequence As String, montant As Double
    Dim nbCreees As Long, nbSuppr As Long
    Dim loRg As ListObject, cRg As Object, dRg As Variant, nRg As Long, ajoutsRg() As Variant, nAjoutsRg As Long
    Dim clesRg As Object, nbAppris As Long

    Set loRg = T(T_REGLES)
    Set cRg = Colonnes(loRg)
    dRg = LireTable(loRg)
    nRg = NbLignes(dRg)
    Set clesRg = CreateObject("Scripting.Dictionary")
    For i = 1 To nRg
        If NormaliserTexte(Texte(dRg(i, C(cRg, "Mot-clé")))) <> "" Then clesRg(NormaliserTexte(Texte(dRg(i, C(cRg, "Mot-clé"))))) = i
    Next i

    Set loA = T(T_ATRAITER)
    Set cA = Colonnes(loA)
    dA = LireTable(loA)
    nA = NbLignes(dA)
    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    Set loP = T(T_PREV)
    Set cP = Colonnes(loP)
    dP = LireTable(loP)
    nP = NbLignes(dP)
    Set loC = T(T_CHARGES)
    Set cC = Colonnes(loC)
    dC = LireTable(loC)
    nC = NbLignes(dC)
    mxC = MaxNumero(dC, nC, C(cC, "ID"), "CF-")
    ReDim ajoutsC(1 To IIf(nA < 1, 1, nA), 1 To loC.ListColumns.Count)
    ReDim ajoutsRg(1 To IIf(nA < 1, 1, nA), 1 To loRg.ListColumns.Count)

    Set indexB = CreateObject("Scripting.Dictionary")
    For i = 1 To nB
        If Texte(dB(i, C(cB, "ID"))) <> "" Then indexB(Texte(dB(i, C(cB, "ID")))) = i
    Next i
    Set indexP = CreateObject("Scripting.Dictionary")
    For j = 1 To nP
        If Texte(dP(j, C(cP, "ID"))) <> "" Then indexP(Texte(dP(j, C(cP, "ID")))) = j
    Next j

    For k = 1 To nA
        id = Texte(dA(k, C(cA, "ID opération")))
        action = Texte(dA(k, C(cA, "Action")))
        If id <> "" And action <> "" Then
            categorie = Texte(dA(k, C(cA, "Catégorie")))
            commentaire = Texte(dA(k, C(cA, "Commentaire")))
            If Not indexB.Exists(id) Then
                erreurs = erreurs & "- " & id & " : opération introuvable" & vbCrLf
            Else
                i = indexB(id)
                If Texte(dB(i, C(cB, "Statut"))) <> OB_A_TRAITER Then
                    erreurs = erreurs & "- " & id & " : déjà traitée" & vbCrLf
                Else
                    If categorie <> "" Then dB(i, C(cB, "Catégorie")) = categorie
                    If commentaire <> "" Then dB(i, C(cB, "Commentaire")) = commentaire
                    If action = ACT_CONSERVER Then
                        dB(i, C(cB, "Statut")) = OB_NON_PREVUE
                        If Texte(dB(i, C(cB, "Catégorie"))) = "" Then dB(i, C(cB, "Catégorie")) = CAT_NON_CATEGORISE
                        nbOk = nbOk + 1
                        nbAppris = nbAppris + Apprendre(Texte(dB(i, C(cB, "Libellé"))), Texte(dB(i, C(cB, "Catégorie"))), _
                                                        action, dRg, nRg, cRg, ajoutsRg, nAjoutsRg, clesRg)
                    ElseIf action = ACT_VIREMENT Then
                        dB(i, C(cB, "Statut")) = OB_VIREMENT
                        If Texte(dB(i, C(cB, "Catégorie"))) = "" Then dB(i, C(cB, "Catégorie")) = CAT_VIREMENT
                        nbOk = nbOk + 1
                        nbAppris = nbAppris + Apprendre(Texte(dB(i, C(cB, "Libellé"))), Texte(dB(i, C(cB, "Catégorie"))), _
                                                        action, dRg, nRg, cRg, ajoutsRg, nAjoutsRg, clesRg)
                    ElseIf action = ACT_RAPPROCHER Then
                        idPrev = Texte(dA(k, C(cA, "ID prévision")))
                        If Not indexP.Exists(idPrev) Then
                            erreurs = erreurs & "- " & id & " : indiquez un ID de prévision existant (colonne « ID prévision »)" & vbCrLf
                        Else
                            j = indexP(idPrev)
                            If Texte(dP(j, C(cP, "Banque"))) <> Texte(dB(i, C(cB, "Banque"))) Then
                                erreurs = erreurs & "- " & id & " : la prévision " & idPrev & " concerne une autre banque" & vbCrLf
                            ElseIf Texte(dP(j, C(cP, "Statut"))) = P_POINTE Or Texte(dP(j, C(cP, "Statut"))) = P_ANNULE _
                                   Or Texte(dP(j, C(cP, "ID opération"))) <> "" Then
                                erreurs = erreurs & "- " & id & " : la prévision " & idPrev & " est déjà pointée ou annulée" & vbCrLf
                            Else
                                modRapprochement.Lier dB, cB, i, dP, cP, j, MODE_MANUEL
                                nbOk = nbOk + 1
                            End If
                        End If
                    ElseIf Left$(action, Len(ACT_CHARGE_PREFIXE)) = ACT_CHARGE_PREFIXE Then
                        frequence = FrequenceAction(action)
                        If frequence = "" Then
                            erreurs = erreurs & "- " & id & " : action inconnue" & vbCrLf
                        Else
                            dB(i, C(cB, "Statut")) = OB_NON_PREVUE
                            If Texte(dB(i, C(cB, "Catégorie"))) = "" Then dB(i, C(cB, "Catégorie")) = CAT_NON_CATEGORISE
                            montant = Nombre(dB(i, C(cB, "Montant")))
                            nAjoutsC = nAjoutsC + 1
                            mxC = mxC + 1
                            ajoutsC(nAjoutsC, C(cC, "ID")) = FormatId("CF-", mxC, 3)
                            ajoutsC(nAjoutsC, C(cC, "Libellé")) = Texte(dB(i, C(cB, "Libellé")))
                            ajoutsC(nAjoutsC, C(cC, "Banque")) = Texte(dB(i, C(cB, "Banque")))
                            ajoutsC(nAjoutsC, C(cC, "Catégorie")) = Texte(dB(i, C(cB, "Catégorie")))
                            If montant < 0 Then
                                ajoutsC(nAjoutsC, C(cC, "Type")) = TY_DEPENSE
                            Else
                                ajoutsC(nAjoutsC, C(cC, "Type")) = TY_RECETTE
                            End If
                            ajoutsC(nAjoutsC, C(cC, "Montant")) = Abs(montant)
                            ajoutsC(nAjoutsC, C(cC, "Première échéance")) = AjouterMois(DateVal(dB(i, C(cB, "Date"))), PasMois(frequence))
                            ajoutsC(nAjoutsC, C(cC, "Fréquence")) = frequence
                            ajoutsC(nAjoutsC, C(cC, "Actif")) = "Oui"
                            ajoutsC(nAjoutsC, C(cC, "Commentaire")) = "Créée depuis l'opération " & id
                            nbOk = nbOk + 1
                        End If
                    Else
                        erreurs = erreurs & "- " & id & " : action inconnue « " & action & " »" & vbCrLf
                    End If
                End If
            End If
        End If
    Next k

    EcrireTable loB, dB, nB
    EcrireTable loP, dP, nP
    If nbAppris > 0 Then
        modPrevisions.ReecrireAvecLignesLibres loRg, dRg, nRg, Array(C(cRg, "Mot-clé"), C(cRg, "Catégorie"), _
                                               C(cRg, "Action automatique")), ajoutsRg, nAjoutsRg
    End If
    If nAjoutsC > 0 Then
        modPrevisions.ReecrireAvecLignesLibres loC, dC, nC, modPrevisions.ColonnesTestCharges(cC), ajoutsC, nAjoutsC
        modPrevisions.GenererEcheancesInterne nbCreees, nbSuppr
    End If
    modRapprochement.RechercherInterne
    ConstruireListeATraiterInterne

    AppliquerTraitementsInterne = nbOk & " opération(s) traitée(s)." & _
        IIf(nAjoutsC > 0, vbCrLf & nAjoutsC & " charge(s) fixe(s) créée(s) (" & nbCreees & " échéance(s) générée(s)).", "") & _
        IIf(nbAppris > 0, vbCrLf & nbAppris & " règle(s) mémorisée(s) : les opérations similaires seront traitées " & _
            "automatiquement (Paramètres > Règles automatiques).", "") & _
        IIf(erreurs <> "", vbCrLf & vbCrLf & "Lignes non traitées :" & vbCrLf & erreurs, "")
End Function
