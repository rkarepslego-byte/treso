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

Private Function CategorieSuggeree(ByVal libelle As String, dRg As Variant, ByVal nRg As Long, cRg As Object) As String
    Dim i As Long, libNorm As String, mot As String
    libNorm = " " & NormaliserTexte(libelle) & " "
    For i = 1 To nRg
        mot = NormaliserTexte(Texte(dRg(i, C(cRg, "Mot-clé"))))
        If mot <> "" And Texte(dRg(i, C(cRg, "Catégorie"))) <> "" Then
            If InStr(1, libNorm, mot, vbBinaryCompare) > 0 Then
                CategorieSuggeree = Texte(dRg(i, C(cRg, "Catégorie")))
                Exit Function
            End If
        End If
    Next i
End Function

' Reconstruit la liste. Les catégories et commentaires déjà saisis sont d'abord
' enregistrés dans le journal bancaire pour ne rien perdre.
Public Function ConstruireListeATraiterInterne() As Long
    Dim loA As ListObject, cA As Object, dA As Variant, nA As Long
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long, indexB As Object
    Dim loRg As ListObject, cRg As Object, dRg As Variant, nRg As Long
    Dim i As Long, j As Long, k As Long, n As Long, id As String, modifie As Boolean
    Dim cles() As String, idx() As Long, liste() As Long, sortie() As Variant
    Dim tolerance As Long, meilleur As Long, meilleurCle As String, cleJ As String, banque As String

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
            sortie(k, C(cA, "Catégorie")) = CategorieSuggeree(Texte(dB(i, C(cB, "Libellé"))), dRg, nRg, cRg)
        End If
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
        If meilleur > 0 Then
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

Public Function AppliquerTraitementsInterne() As String
    Dim loA As ListObject, cA As Object, dA As Variant, nA As Long
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long, indexB As Object
    Dim loP As ListObject, cP As Object, dP As Variant, nP As Long, indexP As Object
    Dim loC As ListObject, cC As Object, dC As Variant, nC As Long, mxC As Long
    Dim ajoutsC() As Variant, nAjoutsC As Long, k As Long, i As Long, j As Long
    Dim id As String, action As String, categorie As String, commentaire As String, idPrev As String
    Dim nbOk As Long, erreurs As String, frequence As String, montant As Double
    Dim nbCreees As Long, nbSuppr As Long

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
                    ElseIf action = ACT_VIREMENT Then
                        dB(i, C(cB, "Statut")) = OB_VIREMENT
                        If Texte(dB(i, C(cB, "Catégorie"))) = "" Then dB(i, C(cB, "Catégorie")) = CAT_VIREMENT
                        nbOk = nbOk + 1
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
    If nAjoutsC > 0 Then
        modPrevisions.ReecrireAvecLignesLibres loC, dC, nC, modPrevisions.ColonnesTestCharges(cC), ajoutsC, nAjoutsC
        modPrevisions.GenererEcheancesInterne nbCreees, nbSuppr
    End If
    modRapprochement.RechercherInterne
    ConstruireListeATraiterInterne

    AppliquerTraitementsInterne = nbOk & " opération(s) traitée(s)." & _
        IIf(nAjoutsC > 0, vbCrLf & nAjoutsC & " charge(s) fixe(s) créée(s) (" & nbCreees & " échéance(s) générée(s)).", "") & _
        IIf(erreurs <> "", vbCrLf & vbCrLf & "Lignes non traitées :" & vbCrLf & erreurs, "")
End Function
