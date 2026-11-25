// Lit tous les relevés déposés dans le dossier d'une banque et produit la table de transit :
//   Banque | Code | Date | DateValeur | Libelle | LibelleNorm | Montant | Reference
//   | Fichier | Ligne | Occurrence | Cle | Erreur
// * Un fichier illisible produit une ligne « Erreur » sans bloquer les autres fichiers.
// * Occurrence = rang d'une ligne parmi les lignes strictement identiques d'un même fichier
//   (deux paiements identiques le même jour restent deux opérations distinctes).
// * Cle = Code banque | date | montant en centimes | libellé normalisé | référence | occurrence
//   Le VBA n'ajoute au journal bancaire que les clés qu'il ne connaît pas encore.
(p as record) as table =>
let
    TypeSortie = type table [
        Banque = text, Code = text, Date = nullable date, DateValeur = nullable date,
        Libelle = text, LibelleNorm = text, Montant = nullable number, Reference = text,
        Fichier = text, Ligne = nullable Int64.Type, Occurrence = nullable Int64.Type,
        Cle = text, Erreur = nullable text
    ],
    NomsColonnes = Table.ColumnNames(#table(TypeSortie, {})),
    LigneErreur = (fichier as text, message as text) as table =>
        #table(TypeSortie, {{p[Banque], p[Code], null, null, "", "", null, "", fichier, null, null, "", message}}),
    Dossier = Text.Trim(p[Dossier]),
    Racine = Text.Lower(if Text.EndsWith(Dossier, "\") then Dossier else Dossier & "\"),
    ContenuDossier = try Folder.Files(Dossier),

    Lire = () as table =>
        let
            Fichiers = Table.SelectRows(
                ContenuDossier[Value],
                each Text.Lower([Folder Path]) = Racine
                    and List.Contains(p[Extensions], Text.Lower([Extension]))
                    and not Text.StartsWith([Name], "~$")
            ),
            Tries = Table.Sort(Fichiers, {{"Name", Order.Ascending}}),
            ParFichier = List.Transform(
                Table.ToRecords(Tries),
                (f) =>
                    let
                        Essai = try Table.Buffer(fnNormaliserReleve(f[Content], f[Name], f[Extension], p)),
                        Lignes =
                            if Essai[HasError] then
                                LigneErreur(f[Name], "Fichier " & f[Name] & " illisible : " & Essai[Error][Message])
                            else
                                Table.FromRecords(
                                    Table.TransformRows(
                                        Essai[Value],
                                        (r) => [
                                            Banque = p[Banque], Code = p[Code], Date = r[Date],
                                            DateValeur = r[DateValeur], Libelle = r[Libelle],
                                            LibelleNorm = r[LibelleNorm], Montant = r[Montant],
                                            Reference = r[Reference], Fichier = f[Name], Ligne = r[Ligne],
                                            Occurrence = null, Cle = "", Erreur = null
                                        ]
                                    ),
                                    TypeSortie
                                )
                    in
                        Lignes
            ),
            Tout = Table.Combine({#table(TypeSortie, {})} & ParFichier),
            Erreurs = Table.SelectRows(Tout, each [Erreur] <> null),
            Valides = Table.SelectRows(Tout, each [Erreur] = null),
            AvecCentimes = Table.AddColumn(
                Table.RemoveColumns(Valides, {"Occurrence", "Cle"}),
                "Centimes",
                each Number.Round([Montant] * 100, 0, RoundingMode.AwayFromZero),
                type number
            ),
            Groupes = Table.Group(
                AvecCentimes,
                {"Fichier", "Date", "Centimes", "LibelleNorm", "Reference"},
                {{"Lignes", each Table.AddIndexColumn(Table.Sort(_, {{"Ligne", Order.Ascending}}), "Occurrence", 1, 1), type table}}
            ),
            AvecOccurrence =
                if Table.RowCount(Groupes) = 0 then Table.AddColumn(AvecCentimes, "Occurrence", each null)
                else Table.Combine(Groupes[Lignes]),
            AvecCle = Table.AddColumn(
                AvecOccurrence,
                "Cle",
                each Text.Combine(
                    {
                        [Code], Date.ToText([Date], "yyyy-MM-dd"), Text.From(Int64.From([Centimes])),
                        [LibelleNorm], [Reference], Text.From([Occurrence])
                    },
                    "|"
                ),
                type text
            ),
            Lignes = Table.SelectColumns(AvecCle, NomsColonnes, MissingField.UseNull),
            Final = Table.Combine({Lignes, Erreurs}),
            Trie = Table.Sort(Final, {{"Fichier", Order.Ascending}, {"Ligne", Order.Ascending}})
        in
            Trie,

    Resultat =
        if Dossier = "" then LigneErreur("", "Dossier d'import non paramétré pour " & p[Banque] & " : lancez « Actualiser » depuis le tableau de bord.")
        else if ContenuDossier[HasError] then LigneErreur("", "Dossier introuvable ou inaccessible : " & Dossier)
        else Lire(),
    Types = Table.TransformColumnTypes(
        Resultat,
        {
            {"Banque", type text}, {"Code", type text}, {"Date", type date}, {"DateValeur", type date},
            {"Libelle", type text}, {"LibelleNorm", type text}, {"Montant", type number},
            {"Reference", type text}, {"Fichier", type text}, {"Ligne", Int64.Type},
            {"Occurrence", Int64.Type}, {"Cle", type text}, {"Erreur", type text}
        }
    )
in
    Types
