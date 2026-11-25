Attribute VB_Name = "modImport"
'===============================================================================
' modImport - Actualisation : sauvegarde, import des relevés, dédoublonnage
'-------------------------------------------------------------------------------
' Chaîne complète lancée par le bouton « Actualiser les banques » :
'   1. sauvegarde du classeur (dossier Sauvegardes)
'   2. transmission des paramètres d'import à Power Query (requête Parametres_Import)
'   3. actualisation des deux requêtes d'import (tables de transit masquées)
'   4. ajout au journal bancaire des seules opérations nouvelles (clé technique)
'   5. génération des échéances des charges fixes
'   6. recherche des correspondances et liste des opérations à traiter
' Le journal bancaire (tbl_Banque) n'est JAMAIS vidé : en cas d'erreur de lecture
' d'un relevé, rien n'est supprimé et l'historique garde la trace de l'erreur.
'===============================================================================
Option Explicit
Option Compare Binary

Public Sub Actualiser()
    Dim rep As VbMsgBoxResult, n As Long, code As Variant, erreurLecture As String
    Dim rapport As String, msgErr As String, sauvegarde As String, resultatTraitements As String
    Dim nbCreees As Long, nbSuppr As Long, nbCert As Long, nbIncert As Long, nbAuto As Long, nbATraiter As Long

    On Error GoTo Erreur
    n = modATraiter.ActionsEnAttente()
    If n > 0 Then
        rep = Afficher(n & " traitement(s) choisi(s) dans l'onglet « À traiter » n'ont pas encore été appliqués." & vbCrLf & vbCrLf & _
                     "Oui : les appliquer avant l'actualisation" & vbCrLf & _
                     "Non : les ignorer" & vbCrLf & "Annuler : ne rien faire", vbYesNoCancel + vbQuestion, "Actualisation")
        If rep = vbCancel Then Exit Sub
        If rep = vbYes Then
            DebutTraitement "Application des traitements en attente..."
            resultatTraitements = modATraiter.AppliquerTraitementsInterne()
            FinTraitement
        End If
    End If

    DebutTraitement "Actualisation : préparation..."
    modAdmin.VerifierRenommageBanques
    CreerDossier DossierBanque("B1")
    CreerDossier DossierBanque("B2")
    CreerDossier DossierSauvegardes()
    modPrevisions.CompleterSaisiesInterne

    AfficherStatut "Actualisation : sauvegarde du classeur..."
    sauvegarde = SauvegarderClasseur()

    AfficherStatut "Actualisation : paramètres d'import..."
    MettreAJourParametresPQ

    For Each code In Array("B1", "B2")
        AfficherStatut "Actualisation : lecture des relevés " & NomBanque(CStr(code)) & "..."
        If RafraichirTransit(CStr(code), erreurLecture) Then
            rapport = rapport & IntegrerTransit(CStr(code))
        Else
            rapport = rapport & NomBanque(CStr(code)) & " : LECTURE IMPOSSIBLE - " & erreurLecture & vbCrLf & _
                      "   (aucune donnée existante n'a été modifiée)" & vbCrLf
            JournaliserErreur CStr(code), erreurLecture
        End If
    Next code

    AfficherStatut "Actualisation : charges fixes..."
    modPrevisions.GenererEcheancesInterne nbCreees, nbSuppr
    AfficherStatut "Actualisation : rapprochement..."
    modRapprochement.RechercherInterne nbCert, nbIncert, nbAuto
    nbATraiter = modATraiter.ConstruireListeATraiterInterne()
    FinTraitement
    Application.Calculate
    ThisWorkbook.Worksheets(F_TDB).Activate

    Afficher "Actualisation terminée." & vbCrLf & vbCrLf & rapport & vbCrLf & _
           modRapprochement.ResumeRecherche(nbCert, nbIncert, nbAuto, nbATraiter) & vbCrLf & vbCrLf & _
           IIf(resultatTraitements <> "", "Traitements appliqués : " & resultatTraitements & vbCrLf & vbCrLf, "") & _
           "Sauvegarde : " & sauvegarde & vbCrLf & _
           "Pensez à enregistrer le classeur.", vbInformation, "Trésorerie"
    Exit Sub
Erreur:
    msgErr = Err.Description
    TerminerTraitement
    AfficherErreur "l'actualisation", msgErr
End Sub

'-------------------------------------------------------------------------------
' Sauvegarde
'-------------------------------------------------------------------------------
Public Function SauvegarderClasseur() As String
    Dim dossier As String, base As String, ext As String, cible As String, pos As Long
    dossier = DossierSauvegardes()
    CreerDossier dossier
    pos = InStrRev(ThisWorkbook.Name, ".")
    If pos > 0 Then
        base = Left$(ThisWorkbook.Name, pos - 1)
        ext = Mid$(ThisWorkbook.Name, pos)
    Else
        base = ThisWorkbook.Name
        ext = ".xlsm"
    End If
    cible = dossier & "\" & base & "_" & Format$(Now, "yyyymmdd_hhnnss") & ext
    ThisWorkbook.SaveCopyAs cible
    PurgerSauvegardes dossier, base, ext
    SauvegarderClasseur = cible
End Function

' Conserve les N sauvegardes les plus récentes (N = paramètre, 0 = tout conserver)
Private Sub PurgerSauvegardes(ByVal dossier As String, ByVal base As String, ByVal ext As String)
    Dim nbGarder As Long, fichiers() As String, n As Long, f As String, idx() As Long, k As Long
    nbGarder = CLng(Nombre(ValeurNom("Nb_Sauvegardes")))
    If nbGarder <= 0 Then Exit Sub
    f = Dir$(dossier & "\" & base & "_????????_??????" & ext)
    Do While f <> ""
        n = n + 1
        ReDim Preserve fichiers(1 To n)
        fichiers(n) = f
        f = Dir$()
    Loop
    If n <= nbGarder Then Exit Sub
    ReDim idx(1 To n)
    For k = 1 To n
        idx(k) = k
    Next k
    TrierIndices fichiers, idx, 1, n
    On Error Resume Next
    For k = 1 To n - nbGarder
        Kill dossier & "\" & fichiers(idx(k))
    Next k
End Sub

'-------------------------------------------------------------------------------
' Paramètres transmis à Power Query
'-------------------------------------------------------------------------------
Private Function MTexte(ByVal s As String) As String
    MTexte = """" & Replace(s, """", """""") & """"
End Function

Private Function MListe(ByVal s As String, ByVal extensions As Boolean) As String
    Dim morceaux() As String, k As Long, v As String, r As String
    morceaux = Split(s, ";")
    For k = LBound(morceaux) To UBound(morceaux)
        v = Trim$(morceaux(k))
        If v <> "" Then
            If extensions Then
                v = LCase$(v)
                If Left$(v, 1) <> "." Then v = "." & v
            Else
                v = NormaliserTexte(v)
            End If
            If r <> "" Then r = r & ", "
            r = r & MTexte(v)
        End If
    Next k
    MListe = "{" & r & "}"
End Function

Private Function MEntier(ByVal v As Variant, ByVal libelle As String, ByVal banque As String) As String
    If Not IsNumeric(Texte(v)) Or Texte(v) = "" Then
        Err.Raise vbObjectError + 1100, "modImport", "Paramètre d'import invalide pour " & banque & " : « " & libelle & " » doit être un nombre entier."
    End If
    MEntier = CStr(CLng(Texte(v)))
End Function

' Construit le record M d'une banque à partir d'une colonne de tbl_ParamImport
Private Function RecordBanque(ByVal code As String, valeurs As Object) As String
    Dim nom As String, mode As String, s As String
    nom = NomBanque(code)
    mode = valeurs("ModeMontant")
    If mode <> "Montant signé" And mode <> "Débit/Crédit" And mode <> "Montant + Sens" Then
        Err.Raise vbObjectError + 1101, "modImport", "Paramètre d'import invalide pour " & nom & _
                  " : « Mode montant » doit valoir Montant signé, Débit/Crédit ou Montant + Sens."
    End If
    s = "[Code = " & MTexte(code) & _
        ", Banque = " & MTexte(nom) & _
        ", Dossier = " & MTexte(DossierBanque(code)) & _
        ", Extensions = " & MListe(valeurs("Extensions"), True) & _
        ", Separateur = " & MTexte(valeurs("Separateur")) & _
        ", Encodage = " & MEntier(valeurs("Encodage"), "Encodage CSV", nom) & _
        ", LignesIgnorees = " & MEntier(valeurs("LignesIgnorees"), "Lignes à ignorer", nom) & _
        ", Feuille = " & MTexte(valeurs("Feuille")) & _
        ", ColDate = " & MTexte(valeurs("ColDate")) & _
        ", ColDateValeur = " & MTexte(valeurs("ColDateValeur")) & _
        ", FormatDate = " & MTexte(valeurs("FormatDate")) & _
        ", ColsLibelle = " & MTexte(valeurs("ColsLibelle")) & _
        ", ModeMontant = " & MTexte(mode) & _
        ", ColMontant = " & MTexte(valeurs("ColMontant")) & _
        ", ColDebit = " & MTexte(valeurs("ColDebit")) & _
        ", ColCredit = " & MTexte(valeurs("ColCredit")) & _
        ", ColSens = " & MTexte(valeurs("ColSens")) & _
        ", ValeursDebit = " & MListe(valeurs("ValeursDebit"), False) & _
        ", InverserSigne = " & IIf(EstOui(valeurs("InverserSigne")), "true", "false") & _
        ", SeparateurDecimal = " & MTexte(valeurs("SeparateurDecimal")) & _
        ", ColReference = " & MTexte(valeurs("ColReference")) & "]"
    RecordBanque = s
End Function

Public Function FormuleParametresM() As String
    Dim lo As ListObject, cI As Object, dI As Variant, i As Long, v1 As Object, v2 As Object, cle As String
    Set lo = T(T_PARAM_IMPORT)
    Set cI = Colonnes(lo)
    dI = LireTable(lo)
    Set v1 = CreateObject("Scripting.Dictionary")
    Set v2 = CreateObject("Scripting.Dictionary")
    For i = 1 To NbLignes(dI)
        cle = Texte(dI(i, C(cI, "Clé")))
        If cle <> "" Then
            v1(cle) = Texte(dI(i, C(cI, "Banque 1")))
            v2(cle) = Texte(dI(i, C(cI, "Banque 2")))
        End If
    Next i
    FormuleParametresM = "// Généré automatiquement par la macro Actualiser à partir de l'onglet Param_Import." & vbCrLf & _
        "// Ne pas modifier ici : modifier la table tbl_ParamImport." & vbCrLf & _
        "let" & vbCrLf & _
        "    Parametres = [" & vbCrLf & _
        "        Banque1 = " & RecordBanque("B1", v1) & "," & vbCrLf & _
        "        Banque2 = " & RecordBanque("B2", v2) & vbCrLf & _
        "    ]" & vbCrLf & _
        "in" & vbCrLf & _
        "    Parametres"
End Function

Public Sub MettreAJourParametresPQ()
    Dim formule As String, q As Object
    formule = FormuleParametresM()
    On Error Resume Next
    Set q = ThisWorkbook.Queries(Q_PARAMETRES)
    On Error GoTo 0
    If q Is Nothing Then
        ThisWorkbook.Queries.Add Q_PARAMETRES, formule, "Paramètres d'import (générés par VBA)"
    ElseIf q.Formula <> formule Then
        q.Formula = formule
    End If
End Sub

'-------------------------------------------------------------------------------
' Actualisation d'une table de transit Power Query
'-------------------------------------------------------------------------------
Public Function RafraichirTransit(ByVal code As String, ByRef messageErreur As String) As Boolean
    Dim lo As ListObject, qt As QueryTable
    messageErreur = ""
    On Error GoTo Echec
    If code = "B1" Then Set lo = T(T_STAGE_B1) Else Set lo = T(T_STAGE_B2)
    On Error Resume Next
    Set qt = lo.QueryTable
    On Error GoTo Echec
    If qt Is Nothing Then
        messageErreur = "la table de transit n'est pas reliée à Power Query." & vbCrLf & _
                        "   Lancez la macro Admin_ReinstallerRequetes (voir le guide technique)."
        Exit Function
    End If
    qt.BackgroundQuery = False
    qt.Refresh BackgroundQuery:=False
    RafraichirTransit = True
    Exit Function
Echec:
    messageErreur = Err.Description
    If InStr(1, messageErreur, "Firewall", vbTextCompare) > 0 Or InStr(1, messageErreur, "confidentialit", vbTextCompare) > 0 Then
        messageErreur = messageErreur & vbCrLf & "   Données > Obtenir des données > Options de requête > Classeur actif > " & _
                        "Confidentialité : cochez « Ignorer les niveaux de confidentialité »."
    End If
    RafraichirTransit = False
End Function

'-------------------------------------------------------------------------------
' Intégration des lignes nouvelles dans le journal bancaire
'-------------------------------------------------------------------------------
Public Function IntegrerTransit(ByVal code As String) As String
    Dim loS As ListObject, cS As Object, dS As Variant, nS As Long
    Dim loB As ListObject, cB As Object, dB As Variant, nB As Long
    Dim loH As ListObject, cH As Object, dH As Variant, nH As Long
    Dim cles As Object, stats As Object, dejaImportes As Object, fichiersOrdre As Collection
    Dim i As Long, k As Long, nom As String, d0 As Double, idImport As String, maxOB As Long, horodatage As Double
    Dim fichier As String, erreur As String, cle As String, dOp As Double, s As Variant
    Dim nouv() As Variant, nNouv As Long, histo() As Variant, nHisto As Long
    Dim totAjout As Long, totDoublons As Long, totAnt As Long, totErreurs As String, alerte As String

    If code = "B1" Then Set loS = T(T_STAGE_B1) Else Set loS = T(T_STAGE_B2)
    Set cS = Colonnes(loS)
    dS = LireTable(loS)
    nS = NbLignes(dS)
    Set loB = T(T_BANQUE)
    Set cB = Colonnes(loB)
    dB = LireTable(loB)
    nB = NbLignes(dB)
    Set loH = T(T_HISTO)
    Set cH = Colonnes(loH)
    dH = LireTable(loH)
    nH = NbLignes(dH)

    nom = NomBanque(code)
    d0 = DateSoldeBanque(nom)
    idImport = FormatId("IMP-", MaxNumero(dH, nH, C(cH, "ID import"), "IMP-") + 1, 5)
    maxOB = MaxNumero(dB, nB, C(cB, "ID"), "OB-")
    horodatage = CDbl(Now)

    Set cles = CreateObject("Scripting.Dictionary")
    For i = 1 To nB
        cle = Texte(dB(i, C(cB, "Clé")))
        If cle <> "" Then cles(cle) = True
    Next i
    Set dejaImportes = CreateObject("Scripting.Dictionary")
    For i = 1 To nH
        If Texte(dH(i, C(cH, "Banque"))) = nom And Nombre(dH(i, C(cH, "Ajoutées"))) > 0 Then
            dejaImportes(Texte(dH(i, C(cH, "Fichier")))) = True
        End If
    Next i

    ' stats(fichier) = Array(lues, ajoutées, doublons, antérieures, erreur)
    Set stats = CreateObject("Scripting.Dictionary")
    Set fichiersOrdre = New Collection
    ReDim nouv(1 To IIf(nS < 1, 1, nS), 1 To loB.ListColumns.Count)
    For i = 1 To nS
        fichier = Texte(dS(i, C(cS, "Fichier")))
        erreur = Texte(dS(i, C(cS, "Erreur")))
        cle = Texte(dS(i, C(cS, "Cle")))
        If fichier = "" And erreur <> "" Then fichier = "(dossier)"
        If fichier <> "" Then
            If Not stats.Exists(fichier) Then
                stats(fichier) = Array(0, 0, 0, 0, "")
                fichiersOrdre.Add fichier
            End If
            s = stats(fichier)
            If erreur <> "" Then
                s(4) = erreur
            ElseIf cle <> "" Then
                s(0) = s(0) + 1
                dOp = DateVal(dS(i, C(cS, "Date")))
                If dOp <= d0 Then
                    s(3) = s(3) + 1
                ElseIf cles.Exists(cle) Then
                    s(2) = s(2) + 1
                Else
                    s(1) = s(1) + 1
                    cles(cle) = True
                    nNouv = nNouv + 1
                    maxOB = maxOB + 1
                    nouv(nNouv, C(cB, "ID")) = FormatId("OB-", maxOB, 6)
                    nouv(nNouv, C(cB, "Banque")) = nom
                    nouv(nNouv, C(cB, "Date")) = dOp
                    If DateVal(dS(i, C(cS, "DateValeur"))) > 0 Then nouv(nNouv, C(cB, "Date valeur")) = DateVal(dS(i, C(cS, "DateValeur")))
                    nouv(nNouv, C(cB, "Libellé")) = Texte(dS(i, C(cS, "Libelle")))
                    nouv(nNouv, C(cB, "Montant")) = Round(Nombre(dS(i, C(cS, "Montant"))), 2)
                    nouv(nNouv, C(cB, "Référence")) = Texte(dS(i, C(cS, "Reference")))
                    nouv(nNouv, C(cB, "Statut")) = OB_A_TRAITER
                    nouv(nNouv, C(cB, "ID import")) = idImport
                    nouv(nNouv, C(cB, "Fichier")) = fichier
                    nouv(nNouv, C(cB, "Date import")) = horodatage
                    nouv(nNouv, C(cB, "Clé")) = cle
                End If
            End If
            stats(fichier) = s
        End If
    Next i

    AjouterLignes loB, nouv, nNouv, C(cB, "ID")

    ReDim histo(1 To IIf(stats.Count < 1, 1, stats.Count), 1 To loH.ListColumns.Count)
    For k = 1 To fichiersOrdre.Count
        fichier = fichiersOrdre(k)
        s = stats(fichier)
        nHisto = nHisto + 1
        histo(nHisto, C(cH, "ID import")) = idImport
        histo(nHisto, C(cH, "Date et heure")) = horodatage
        histo(nHisto, C(cH, "Banque")) = nom
        histo(nHisto, C(cH, "Fichier")) = fichier
        histo(nHisto, C(cH, "Lignes lues")) = s(0)
        histo(nHisto, C(cH, "Ajoutées")) = s(1)
        histo(nHisto, C(cH, "Doublons ignorés")) = s(2)
        histo(nHisto, C(cH, "Antérieures au solde initial")) = s(3)
        histo(nHisto, C(cH, "Erreur")) = s(4)
        histo(nHisto, C(cH, "Utilisateur")) = Environ$("USERNAME")
        totAjout = totAjout + s(1)
        totDoublons = totDoublons + s(2)
        totAnt = totAnt + s(3)
        If s(4) <> "" Then totErreurs = totErreurs & "   ! " & s(4) & vbCrLf
        If s(1) > 0 And dejaImportes.Exists(fichier) Then
            alerte = alerte & "   ! Le fichier « " & fichier & " » avait déjà été importé et contient " & s(1) & _
                     " ligne(s) nouvelle(s) : vérifiez qu'il ne s'agit pas d'un changement de format ou de paramètres." & vbCrLf
        End If
    Next k
    AjouterLignes loH, histo, nHisto, C(cH, "ID import")

    If fichiersOrdre.Count = 0 Then
        IntegrerTransit = nom & " : aucun relevé trouvé dans " & DossierBanque(code) & vbCrLf
    Else
        IntegrerTransit = nom & " : " & totAjout & " nouvelle(s) opération(s), " & totDoublons & " doublon(s) ignoré(s)" & _
            IIf(totAnt > 0, ", " & totAnt & " antérieure(s) au solde initial", "") & vbCrLf & totErreurs & alerte
    End If
End Function

Private Sub JournaliserErreur(ByVal code As String, ByVal message As String)
    Dim loH As ListObject, cH As Object, dH As Variant, histo() As Variant
    On Error Resume Next
    Set loH = T(T_HISTO)
    Set cH = Colonnes(loH)
    dH = LireTable(loH)
    ReDim histo(1 To 1, 1 To loH.ListColumns.Count)
    histo(1, C(cH, "ID import")) = FormatId("IMP-", MaxNumero(dH, NbLignes(dH), C(cH, "ID import"), "IMP-") + 1, 5)
    histo(1, C(cH, "Date et heure")) = CDbl(Now)
    histo(1, C(cH, "Banque")) = NomBanque(code)
    histo(1, C(cH, "Fichier")) = "(requête)"
    histo(1, C(cH, "Erreur")) = Left$(message, 250)
    histo(1, C(cH, "Utilisateur")) = Environ$("USERNAME")
    AjouterLignes loH, histo, 1, C(cH, "ID import")
End Sub
