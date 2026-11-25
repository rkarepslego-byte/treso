// Convertit une valeur de relevé en date : accepte les dates/dates-heures Excel,
// les numéros de série et les textes au format indiqué (ex. "dd/MM/yyyy").
(valeur as any, formatDate as text) as nullable date =>
let
    Resultat =
        if valeur = null then null
        else if Value.Is(valeur, type date) then valeur
        else if Value.Is(valeur, type datetime) then Date.From(valeur)
        else if Value.Is(valeur, type datetimezone) then Date.From(DateTimeZone.RemoveZone(valeur))
        else if Value.Is(valeur, type number) then Date.From(valeur)
        else
            let
                Texte = Text.Trim(Text.From(valeur)),
                Essai1 = try Date.FromText(Texte, [Format = formatDate, Culture = "fr-FR"]) otherwise null,
                Essai2 = if Essai1 <> null then Essai1 else try Date.From(DateTime.FromText(Texte, "fr-FR")) otherwise null,
                Essai3 = if Essai2 <> null then Essai2 else try Date.FromText(Texte, [Format = "yyyy-MM-dd"]) otherwise null
            in
                if Texte = "" then null else Essai3
in
    Resultat
