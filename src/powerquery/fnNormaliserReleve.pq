// Transforme UN fichier de relevé (CSV/TXT ou Excel) en table normalisée :
//   Date | DateValeur | Libelle | Montant | Reference | Ligne | LibelleNorm
// Tous les réglages viennent du record p (une colonne de tbl_ParamImport).
// C'est la fonction à adapter si un format bancaire réel ne rentre pas dans les paramètres.
(contenu as binary, nomFichier as text, extension as text, p as record) as table =>
let
    Ext = Text.Lower(extension),
    EstTexte = Ext = ".csv" or Ext = ".txt",
    Separateur = if Text.Upper(p[Separateur]) = "TAB" then "#(tab)" else p[Separateur],

    // 1. Lecture brute (les lignes de présentation de la banque sont sautées)
    Brut =
        if EstTexte then
            let
                Lignes = Lines.FromBinary(contenu, QuoteStyle.Csv, false, p[Encodage]),
                Utiles = List.Skip(Lignes, p[LignesIgnorees]),
                Tableau = Csv.Document(Text.Combine(Utiles, "#(lf)"), [Delimiter = Separateur, QuoteStyle = QuoteStyle.Csv])
            in
                Tableau
        else
            let
                Classeur = Excel.Workbook(contenu, null, true),
                Feuilles = Table.SelectRows(Classeur, each [Kind] = "Sheet"),
                Feuille =
                    if p[Feuille] = "" then Feuilles{0}[Data]
                    else Table.SelectRows(Feuilles, each [Name] = p[Feuille]){0}[Data]
            in
                Table.Skip(Feuille, p[LignesIgnorees]),

    // 2. En-têtes et recherche des colonnes (casse, accents et espaces indifférents)
    AvecEntetes = Table.PromoteHeaders(Brut, [PromoteAllScalars = true]),
    Colonnes = Table.ColumnNames(AvecEntetes),
    ColonnesNorm = List.Transform(Colonnes, fnTexteNormalise),
    Trouver = (attendu as text) as nullable text =>
        if Text.Trim(attendu) = "" then null
        else
            let
                Position = List.PositionOf(ColonnesNorm, fnTexteNormalise(attendu))
            in
                if Position < 0 then
                    error Error.Record(
                        "Colonne introuvable",
                        "La colonne « " & attendu & " » est absente du fichier " & nomFichier
                            & ". Colonnes trouvées : " & Text.Combine(Colonnes, ", ")
                    )
                else Colonnes{Position},

    Mode = p[ModeMontant],
    SepDecimal = p[SeparateurDecimal],
    ColDate = Trouver(p[ColDate]),
    ColDateValeur = Trouver(p[ColDateValeur]),
    ColLibelles = List.Transform(
        List.Select(List.Transform(Text.Split(p[ColsLibelle], ";"), Text.Trim), each _ <> ""),
        Trouver
    ),
    ColReference = Trouver(p[ColReference]),
    ColMontant = if Mode = "Débit/Crédit" then null else Trouver(p[ColMontant]),
    ColDebit = if Mode = "Débit/Crédit" then Trouver(p[ColDebit]) else null,
    ColCredit = if Mode = "Débit/Crédit" then Trouver(p[ColCredit]) else null,
    ColSens = if Mode = "Montant + Sens" then Trouver(p[ColSens]) else null,
    // Force l'évaluation : une colonne manquante doit être signalée même si le fichier est vide
    Verification = Text.Combine(
        List.Transform(
            {ColDate, ColDateValeur, ColReference, ColMontant, ColDebit, ColCredit, ColSens} & ColLibelles,
            each if _ = null then "" else _
        ),
        "|"
    ),

    Valeur = (ligne as record, colonne as nullable text) as any =>
        if colonne = null then null else Record.Field(ligne, colonne),
    Texte = (v as any) as text => if v = null then "" else Text.Trim(Text.From(v, "fr-FR")),

    // 3. Conversion ligne par ligne
    Indexe = Table.AddIndexColumn(AvecEntetes, "__Ligne__", p[LignesIgnorees] + 2, 1),
    Enregistrements = Table.TransformRows(
        Indexe,
        (r) =>
            let
                DateOp = fnLireDate(Valeur(r, ColDate), p[FormatDate]),
                DateVal = fnLireDate(Valeur(r, ColDateValeur), p[FormatDate]),
                Libelle = Text.Combine(
                    List.Select(List.Transform(ColLibelles, (c) => Texte(Valeur(r, c))), each _ <> ""),
                    " "
                ),
                Montant =
                    if Mode = "Débit/Crédit" then
                        let
                            Debit = fnLireMontant(Valeur(r, ColDebit), SepDecimal),
                            Credit = fnLireMontant(Valeur(r, ColCredit), SepDecimal)
                        in
                            if Debit = null and Credit = null then null
                            else (if Credit = null then 0 else Number.Abs(Credit))
                                - (if Debit = null then 0 else Number.Abs(Debit))
                    else if Mode = "Montant + Sens" then
                        let
                            MontantBrut = fnLireMontant(Valeur(r, ColMontant), SepDecimal),
                            Sens = fnTexteNormalise(Texte(Valeur(r, ColSens)))
                        in
                            if MontantBrut = null then null
                            else if List.Contains(p[ValeursDebit], Sens) then -Number.Abs(MontantBrut)
                            else Number.Abs(MontantBrut)
                    else
                        let
                            MontantBrut = fnLireMontant(Valeur(r, ColMontant), SepDecimal)
                        in
                            if MontantBrut = null then null else if p[InverserSigne] then -MontantBrut else MontantBrut
            in
                [
                    Date = DateOp,
                    DateValeur = DateVal,
                    Libelle = Libelle,
                    Montant = if Montant = null then null else Number.Round(Montant, 2),
                    Reference = Texte(Valeur(r, ColReference)),
                    Ligne = r[__Ligne__]
                ]
    ),
    Tableau = Table.FromRecords(
        Enregistrements,
        type table [
            Date = nullable date, DateValeur = nullable date, Libelle = text,
            Montant = nullable number, Reference = text, Ligne = Int64.Type
        ]
    ),
    // Les lignes sans date ou sans montant (titres, totaux, soldes) sont ignorées
    Valides = Table.SelectRows(Tableau, each [Date] <> null and [Montant] <> null and [Montant] <> 0),
    Resultat = Table.AddColumn(Valides, "LibelleNorm", each fnTexteNormalise([Libelle]), type text)
in
    if Verification = null then error "Paramètres de colonnes invalides" else Resultat
