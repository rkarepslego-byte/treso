// Normalise un libellé : majuscules, sans accents, uniquement lettres/chiffres séparés
// par un espace simple. Sert à comparer les libellés et à construire la clé technique.
(texte as nullable text) as text =>
let
    Source = if texte = null then "" else texte,
    // Astuce classique : la page de code 1251 translittère les lettres accentuées (é -> e)
    SansAccents = Text.FromBinary(Text.ToBinary(Source, 1251), TextEncoding.Ascii),
    Majuscules = Text.Upper(SansAccents),
    Caracteres = List.Transform(
        Text.ToList(Majuscules),
        each if (_ >= "A" and _ <= "Z") or (_ >= "0" and _ <= "9") then _ else " "
    ),
    Mots = List.Select(Text.Split(Text.Combine(Caracteres), " "), each _ <> ""),
    Resultat = Text.Combine(Mots, " ")
in
    Resultat
