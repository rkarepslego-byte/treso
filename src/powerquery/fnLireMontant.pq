// Convertit une valeur de relevé en nombre : accepte les nombres, et les textes du type
// "1 234,56", "-1.234,56", "1234.56", "(1 234,56)", "1 234,56 €", "123,45-".
(valeur as any, separateurDecimal as text) as nullable number =>
let
    Resultat =
        if valeur = null then null
        else if Value.Is(valeur, type number) then valeur
        else
            let
                Texte = Text.Trim(Text.From(valeur, "fr-FR")),
                Negatif = Text.Contains(Texte, "-") or Text.StartsWith(Texte, "("),
                Chiffres = Text.Select(Texte, {"0".."9", ",", "."}),
                Normalise =
                    if separateurDecimal = "," then
                        Text.Replace(Text.Remove(Chiffres, {"."}), ",", ".")
                    else
                        Text.Remove(Chiffres, {","}),
                Nombre = if Normalise = "" or Normalise = "." then null else Number.FromText(Normalise, "en-US")
            in
                if Nombre = null then null else if Negatif then -Nombre else Nombre
in
    Resultat
