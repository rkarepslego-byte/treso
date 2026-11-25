"""Contrôle statique du code VBA (aucun compilateur VBA n'existe hors de Windows).

Vérifie :
  * encodage Windows-1252, longueur des lignes, présence de « Option Explicit » ;
  * équilibre des blocs (Sub/Function, If, For, Do, With, Select, While) ;
  * variables non déclarées (équivalent de l'erreur de compilation « Variable non définie ») ;
  * appels de procédures entre modules (modX.Procedure) et macros des boutons ;
  * noms de colonnes utilisés via C(cX, "...") et noms définis (ValeurNom, EcrireNom, Names)
    comparés à la structure du classeur (build/schema.py et classeur généré).
"""

from __future__ import annotations

import glob
import os
import re
import sys

RACINE = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.insert(0, RACINE)

from build import schema as S  # noqa: E402

MOTS_CLES = set("""
and as boolean byref byval call case const currency date dim do double each else elseif empty end
error exit explicit false for function goto if in integer is let like long loop mod new next not nothing
null object on option or preserve private public redim resume select set single step string sub then to
true until variant wend while with xor binary compare compare text resume next byte paramarray optional
static type property get friend withevents raiseevent stop me debug print lbound ubound attribute
""".split())

FONCTIONS_VBA = set("""
abs array asc cbool cbyte ccur cdate cdbl cdec chr chrw cint clng csng cstr cvar date dateadd datediff
datepart dateserial datevalue day dir doevents environ eof err error exp filedatetime filelen fix format
getattr hex hour iif instr instrrev int isarray isdate isempty iserror ismissing isnull isnumeric isobject
join kill lcase left len log ltrim mid minute mkdir month msgbox now oct replace right rnd round rtrim
second sgn shell space split sqr str strcomp strconv string timer timeserial trim typename ucase val
weekday year createobject getobject inputbox vartype filecopy name rmdir chdir curdir
format$ left$ right$ mid$ trim$ lcase$ ucase$ string$ dir$ environ$ chr$ space$ ltrim$ rtrim$ str$
application thisworkbook activesheet activecell activeworkbook worksheets sheets range cells intersect
union selection
""".split())

CONSTANTES = re.compile(r"^(vb|xl|msg)[a-z0-9]*$", re.I)

TYPES_OBJETS = set("""
listobject listcolumn listrow worksheet workbook range chart chartobject shape querytable object collection
vbmsgboxresult dictionary name names workbookquery
""".split())


def lire_modules():
    mods = {}
    for f in sorted(glob.glob(os.path.join(RACINE, "src", "vba", "*.bas")) + glob.glob(os.path.join(RACINE, "src", "vba", "*.cls"))):
        with open(f, encoding="utf-8") as fh:
            mods[os.path.splitext(os.path.basename(f))[0]] = fh.read()
    return mods


def sans_chaines_ni_commentaires(ligne: str) -> str:
    sortie, dans_chaine, i = [], False, 0
    while i < len(ligne):
        ch = ligne[i]
        if dans_chaine:
            if ch == '"':
                if i + 1 < len(ligne) and ligne[i + 1] == '"':
                    i += 2
                    continue
                dans_chaine = False
                sortie.append('"')
            i += 1
            continue
        if ch == '"':
            dans_chaine = True
            sortie.append('"')
        elif ch == "'":
            break
        else:
            sortie.append(ch)
        i += 1
    s = "".join(sortie)
    if re.match(r"^\s*Rem\b", s, re.I):
        return ""
    return s


def lignes_logiques(texte: str):
    """(numéro, ligne sans chaînes ni commentaires, ligne brute) avec continuations « _ » jointes."""
    brutes = texte.splitlines()
    i = 0
    while i < len(brutes):
        debut = i
        ligne = brutes[i]
        while ligne.rstrip().endswith(" _") and i + 1 < len(brutes):
            i += 1
            ligne = ligne.rstrip()[:-1] + " " + brutes[i].strip()
        yield debut + 1, sans_chaines_ni_commentaires(ligne), ligne
        i += 1


def entete_procedure(ligne: str):
    """(nom, liste des paramètres) si la ligne déclare une procédure, sinon None."""
    m = re.match(r"^\s*(Public |Private |Friend )?(Static )?(Sub|Function|Property Get|Property Let|Property Set)\s+(\w+)\s*\(",
                 ligne, re.I)
    if not m:
        return None
    i, prof, debut = m.end(), 1, m.end()
    while i < len(ligne) and prof:
        if ligne[i] == "(":
            prof += 1
        elif ligne[i] == ")":
            prof -= 1
        i += 1
    interieur = ligne[debut:i - 1]
    params, prof, morceau = [], 0, ""
    for ch in interieur + ",":
        if ch == "(":
            prof += 1
        elif ch == ")":
            prof -= 1
        if ch == "," and prof == 0:
            if morceau.strip():
                params.append(morceau.strip())
            morceau = ""
        else:
            morceau += ch
    return m, params


def instructions(ligne: str):
    """Découpe une ligne logique (sans chaînes) en instructions séparées par « : »."""
    morceaux = [p.strip() for p in re.split(r":(?!=)", ligne)]
    return [p for p in morceaux if p]


def verifier(mods: dict, noms_definis: set | None = None) -> list[str]:
    erreurs = []
    procedures = {}       # nom (minuscule) -> (module, public, nb_params_obligatoires)
    globales = set()      # identifiants de niveau module (publics ou privés par module)
    privees_par_module = {m: set() for m in mods}

    # Passe 1 : déclarations de niveau module et procédures
    for nom_mod, texte in mods.items():
        if "Option Explicit" not in texte:
            erreurs.append(f"{nom_mod} : « Option Explicit » manquant")
        try:
            texte.encode("cp1252")
        except UnicodeEncodeError as e:
            erreurs.append(f"{nom_mod} : caractère non encodable en Windows-1252 ({e})")
        dans_proc = False
        for num, phys in enumerate(texte.splitlines(), start=1):
            if len(phys) > 1000:
                erreurs.append(f"{nom_mod}:{num} : ligne physique trop longue (limite VBA 1023)")
        suite = 0
        for num, phys in enumerate(texte.splitlines(), start=1):
            suite = suite + 1 if phys.rstrip().endswith(" _") else 0
            if suite > 24:
                erreurs.append(f"{nom_mod}:{num} : plus de 24 continuations de ligne (limite VBA)")
        for num, ligne, brute in lignes_logiques(texte):
            ep = entete_procedure(ligne)
            if ep:
                m, params = ep
                dans_proc = True
                public = (m.group(1) or "Public ").strip().lower() != "private"
                obligatoires = [p for p in params if not re.match(r"^(Optional|ParamArray)\b", p, re.I)]
                procedures[m.group(4).lower()] = (nom_mod, public, len(obligatoires), len(params))
                continue
            if re.match(r"^\s*End (Sub|Function|Property)\b", ligne, re.I):
                dans_proc = False
                continue
            if not dans_proc:
                m = re.match(r"^\s*(Public|Private|Dim|Global)\s+(Const\s+)?(.*)$", ligne, re.I)
                if m and not re.match(r"^\s*(Public|Private)\s+(Sub|Function|Property|Declare|Type|Enum)\b", ligne, re.I):
                    for decl in m.group(3).split(","):
                        ident = re.match(r"\s*(\w+)", decl)
                        if ident:
                            if m.group(1).lower() == "private" or m.group(1).lower() == "dim":
                                privees_par_module[nom_mod].add(ident.group(1).lower())
                            else:
                                globales.add(ident.group(1).lower())

    # Passe 2 : blocs, variables, appels
    for nom_mod, texte in mods.items():
        pile = []
        locales: set[str] = set()
        proc_courante = None
        for num, ligne, brute in lignes_logiques(texte):
            for instr in instructions(ligne):
                low = instr.lower()
                if re.match(r"^attribute\b", low) or re.match(r"^option\b", low):
                    continue
                ep = entete_procedure(low)
                if ep:
                    m, params = ep
                    pile.append(("proc", num))
                    proc_courante = m.group(4).lower()
                    locales = {proc_courante}
                    for p in params:
                        p = re.sub(r"^(optional|byval|byref|paramarray)\s+", "", p.strip())
                        p = re.sub(r"^(optional|byval|byref|paramarray)\s+", "", p)
                        ident = re.match(r"(\w+)", p)
                        if ident:
                            locales.add(ident.group(1))
                    continue
                if re.match(r"^end (sub|function|property)\b", low):
                    if not pile or pile[-1][0] != "proc":
                        erreurs.append(f"{nom_mod}:{num} : « {instr} » sans début correspondant (pile={pile[-1:] })")
                    else:
                        pile.pop()
                    proc_courante = None
                    continue
                # blocs
                if re.match(r"^if\b.*\bthen$", low):
                    pile.append(("if", num))
                elif re.match(r"^end if\b", low):
                    if not pile or pile[-1][0] != "if":
                        erreurs.append(f"{nom_mod}:{num} : End If inattendu")
                    else:
                        pile.pop()
                elif re.match(r"^for\b", low):
                    pile.append(("for", num))
                elif re.match(r"^next\b", low):
                    if not pile or pile[-1][0] != "for":
                        erreurs.append(f"{nom_mod}:{num} : Next inattendu")
                    else:
                        pile.pop()
                elif re.match(r"^do\b", low):
                    pile.append(("do", num))
                elif re.match(r"^loop\b", low):
                    if not pile or pile[-1][0] != "do":
                        erreurs.append(f"{nom_mod}:{num} : Loop inattendu")
                    else:
                        pile.pop()
                elif re.match(r"^while\b", low):
                    pile.append(("while", num))
                elif re.match(r"^wend\b", low):
                    if not pile or pile[-1][0] != "while":
                        erreurs.append(f"{nom_mod}:{num} : Wend inattendu")
                    else:
                        pile.pop()
                elif re.match(r"^with\b", low):
                    pile.append(("with", num))
                elif re.match(r"^end with\b", low):
                    if not pile or pile[-1][0] != "with":
                        erreurs.append(f"{nom_mod}:{num} : End With inattendu")
                    else:
                        pile.pop()
                elif re.match(r"^select case\b", low):
                    pile.append(("select", num))
                elif re.match(r"^end select\b", low):
                    if not pile or pile[-1][0] != "select":
                        erreurs.append(f"{nom_mod}:{num} : End Select inattendu")
                    else:
                        pile.pop()
                elif re.match(r"^(elseif|else)\b", low):
                    if not pile or pile[-1][0] != "if":
                        if not (re.match(r"^else\b", low) and pile and pile[-1][0] == "select"):
                            erreurs.append(f"{nom_mod}:{num} : {instr.split()[0]} hors d'un bloc If")

                if proc_courante is None:
                    continue
                # déclarations locales
                md = re.match(r"^(dim|static|const|redim(\s+preserve)?)\s+(.*)$", low)
                if md and md.group(1).startswith(("dim", "static", "const")):
                    profondeur = 0
                    morceau = ""
                    for ch in md.group(3) + ",":
                        if ch == "(":
                            profondeur += 1
                        elif ch == ")":
                            profondeur -= 1
                        if ch == "," and profondeur == 0:
                            ident = re.match(r"\s*(\w+)", morceau)
                            if ident:
                                locales.add(ident.group(1))
                            morceau = ""
                        else:
                            morceau += ch
                    if md.group(1).startswith("const"):
                        continue
                    # la partie « As Type » et les dimensions ne contiennent pas de variables à contrôler
                    continue
                # identifiants utilisés
                travail = re.sub(r'"[^"]*"', '""', instr)
                travail = re.sub(r"\bAs\s+(New\s+)?\w+", " ", travail, flags=re.I)
                travail = re.sub(r"\w+\s*:=", " ", travail)          # arguments nommés
                travail = re.sub(r"\.\s*\w+", " ", travail)            # membres d'objet
                travail = re.sub(r"#[\d/]+#", " ", travail)
                travail = re.sub(r"\b\d+(\.\d+)?#?\b", " ", travail)
                travail = re.sub(r"&H[0-9A-F]+", " ", travail, flags=re.I)
                if re.match(r"^\s*(goto|on error goto|resume)\b", travail, re.I):
                    continue
                if re.match(r"^\s*\w+\s*$", travail) and instr.endswith(":") is False and re.match(r"^\s*[A-Za-z]\w*\s*$", travail) \
                        and travail.strip().lower() in ("erreur", "fin", "echec", "horstable"):
                    continue
                for ident in re.findall(r"\b[A-Za-z_]\w*\$?", travail):
                    li = ident.lower()
                    if li in MOTS_CLES or li in FONCTIONS_VBA or CONSTANTES.match(li) or li in TYPES_OBJETS:
                        continue
                    if li in locales or li in globales or li in privees_par_module[nom_mod] or li in procedures:
                        continue
                    if li in {m.lower() for m in mods}:   # préfixe de module : modX.Proc
                        continue
                    if li in ("erreur", "fin", "echec", "horstable"):
                        continue
                    erreurs.append(f"{nom_mod}:{num} : identifiant non déclaré « {ident} » dans {proc_courante}")
                # appels qualifiés modX.Proc
                for mod_appel, proc in re.findall(r"\b(mod\w+)\.(\w+)", instr):
                    info = procedures.get(proc.lower())
                    if mod_appel not in mods:
                        erreurs.append(f"{nom_mod}:{num} : module inconnu {mod_appel}")
                    elif info is None or info[0] != mod_appel:
                        erreurs.append(f"{nom_mod}:{num} : procédure {mod_appel}.{proc} introuvable")
                    elif not info[1]:
                        erreurs.append(f"{nom_mod}:{num} : procédure {mod_appel}.{proc} privée")
        if pile:
            erreurs.append(f"{nom_mod} : blocs non fermés {pile}")

    # Colonnes C(cX, "...")
    tables_par_variable = {
        "cP": "tbl_Previsions", "cB": "tbl_Banque", "cC": "tbl_Charges", "cA": "tbl_ATraiter", "cR": "tbl_Propositions",
        "cRej": "tbl_Rejets", "cH": "tbl_Historique", "cS": "tbl_Stage_B1", "cI": "tbl_ParamImport", "cM": "tbl_CodeM",
        "cRg": "tbl_Regles",
    }
    for nom_mod, texte in mods.items():
        for num, ligne in enumerate(texte.splitlines(), start=1):
            for var, col in re.findall(r'\bC\((c\w+),\s*"([^"]+)"\)', ligne):
                table = tables_par_variable.get(var)
                if table is None:
                    erreurs.append(f"{nom_mod}:{num} : variable de colonnes inconnue {var}")
                elif col not in S.TABLES_PAR_NOM[table].noms():
                    erreurs.append(f"{nom_mod}:{num} : colonne « {col} » absente de {table}")
            for col in re.findall(r'cols\.Exists\("([^"]+)"\)', ligne):
                pass
            if noms_definis is not None:
                for nom in re.findall(r'(?:ValeurNom|EcrireNom|Names)\(\s*"([^"]+)"', ligne) + \
                        re.findall(r'EcrireNom\s+"([^"]+)"', ligne):
                    if nom not in noms_definis:
                        erreurs.append(f"{nom_mod}:{num} : nom défini « {nom} » absent du classeur")

    # Constantes de modConfig cohérentes avec le schéma
    config = mods.get("modConfig", "")
    for cst, val in re.findall(r'Public Const (T_\w+) As String = "([^"]+)"', config):
        if val not in S.TABLES_PAR_NOM:
            erreurs.append(f"modConfig : {cst} = {val} absent du schéma")
    onglets = {n for n, _c, _v in S.ONGLETS}
    for cst, val in re.findall(r'Public Const (F_\w+) As String = "([^"]+)"', config):
        if val not in onglets:
            erreurs.append(f"modConfig : {cst} = {val} n'est pas un onglet du classeur")
    return erreurs, procedures


def macros_boutons() -> set[str]:
    with open(os.path.join(RACINE, "build", "generer_classeur.py"), encoding="utf-8") as f:
        return set(re.findall(r'self\.bouton\([^)]*?"([A-Za-z0-9_]+)"\s*,\s*\d+', f.read()))


if __name__ == "__main__":
    errs, procs = verifier(lire_modules())
    for b in sorted(macros_boutons()):
        info = procs.get(b.lower())
        if info is None or not info[1] or info[2] != 0:
            errs.append(f"bouton : macro {b} introuvable ou non publique / avec paramètres")
    print("\n".join(errs) if errs else "Aucune anomalie.")
    sys.exit(1 if errs else 0)
