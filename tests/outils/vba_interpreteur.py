"""Interpréteur d'un sous-ensemble de VBA, écrit pour exécuter les modules de src/vba
hors d'Excel (avec un modèle objet Excel simulé : tests/outils/excel_simule.py).

Sous-ensemble pris en charge (celui utilisé par l'outil) : modules standard, Sub /
Function, ByRef / ByVal / Optional / ParamArray, Dim / ReDim [Preserve] / Const,
If (bloc et sur une ligne), For / For Each / Do / While, Select Case, Exit, GoTo,
étiquettes, On Error GoTo / Resume Next / GoTo 0, objet Err, tableaux à bornes
quelconques, Variant, chaînes, opérateurs (dont Like, \\, Mod, &), fonctions
intégrées courantes, Scripting.Dictionary, Collection.

Ce n'est pas un VBA complet : son seul but est de vérifier que le code de l'outil
s'exécute sans erreur et produit les mêmes résultats que le moteur de référence.
"""

from __future__ import annotations

import calendar
import glob as _glob
import math
import os
import re
from datetime import datetime, timedelta

# ---------------------------------------------------------------------------
# Valeurs
# ---------------------------------------------------------------------------


class _Special:
    def __init__(self, nom):
        self.nom = nom

    def __repr__(self):
        return self.nom


EMPTY = _Special("Empty")
NOTHING = _Special("Nothing")
NULL = _Special("Null")
MISSING = _Special("Missing")


class VDate(float):
    """Valeur de type Date (numéro de série Excel)."""


class VBAErreur(Exception):
    def __init__(self, numero, description, source=""):
        super().__init__(f"{numero}: {description}")
        self.numero = numero
        self.description = description
        self.source = source


def erreur(numero, description):
    raise VBAErreur(numero, description)


class VArray:
    def __init__(self, bornes, defaut=EMPTY):
        self.bornes = list(bornes)          # [(lb, ub), ...]
        self.defaut = defaut
        self.data = {}

    @staticmethod
    def depuis_liste(valeurs, base=0):
        a = VArray([(base, base + len(valeurs) - 1)])
        for i, v in enumerate(valeurs):
            a.data[(base + i,)] = v
        return a

    def _cle(self, indices):
        if len(indices) != len(self.bornes):
            erreur(9, "L'indice n'appartient pas à la sélection (dimensions)")
        cle = []
        for i, (lb, ub) in zip(indices, self.bornes):
            i = cint(i)
            if i < lb or i > ub:
                erreur(9, f"L'indice n'appartient pas à la sélection ({i} hors de {lb}..{ub})")
            cle.append(i)
        return tuple(cle)

    def get(self, indices):
        return self.data.get(self._cle(indices), self.defaut)

    def set(self, indices, valeur):
        self.data[self._cle(indices)] = valeur

    def copie(self):
        a = VArray(self.bornes, self.defaut)
        a.data = dict(self.data)
        return a

    def redim(self, bornes, preserve):
        if preserve and self.bornes:
            if len(bornes) != len(self.bornes) or any(b != a for a, b in zip(self.bornes[:-1], bornes[:-1])):
                erreur(9, "ReDim Preserve : seule la dernière dimension peut changer")
            anciens = self.data
            self.bornes = list(bornes)
            self.data = {k: v for k, v in anciens.items() if all(lb <= x <= ub for x, (lb, ub) in zip(k, bornes))}
        else:
            self.bornes = list(bornes)
            self.data = {}

    def lignes(self):
        """Valeurs en liste de listes (tableaux 2D) ou liste (1D)."""
        if len(self.bornes) == 1:
            lb, ub = self.bornes[0]
            return [self.get((i,)) for i in range(lb, ub + 1)]
        (l1, u1), (l2, u2) = self.bornes
        return [[self.get((i, j)) for j in range(l2, u2 + 1)] for i in range(l1, u1 + 1)]


def copier_valeur(v):
    return v.copie() if isinstance(v, VArray) else v


# ---------------------------------------------------------------------------
# Conversions
# ---------------------------------------------------------------------------
def est_vide(v):
    return v is EMPTY or v is MISSING


def vers_nombre(v):
    if v is EMPTY or v is None:
        return 0
    if isinstance(v, bool):
        return -1 if v else 0
    if isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip()
        if s == "":
            erreur(13, "Incompatibilité de type (chaîne vide)")
        try:
            return float(s.replace(",", ".")) if ("." in s or "," in s or "e" in s.lower()) else int(s)
        except ValueError:
            erreur(13, f"Incompatibilité de type : « {v} » n'est pas un nombre")
    erreur(13, f"Incompatibilité de type : {v!r}")


def arrondi_bancaire(x):
    return int(round(x))


def cint(v):
    return arrondi_bancaire(vers_nombre(v))


def cdbl(v):
    return float(vers_nombre(v))


def cstr(v):
    if v is EMPTY:
        return ""
    if isinstance(v, bool):
        return "Vrai" if v else "Faux"
    if isinstance(v, VDate):
        d = serie_vers_datetime(v)
        return d.strftime("%d/%m/%Y") if float(v).is_integer() else d.strftime("%d/%m/%Y %H:%M:%S")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else repr(v)
    if isinstance(v, int):
        return str(v)
    if isinstance(v, str):
        return v
    if v is NULL:
        erreur(94, "Utilisation incorrecte de Null")
    erreur(13, f"Incompatibilité de type (CStr) : {v!r}")


def cbool(v):
    if isinstance(v, bool):
        return v
    if v is EMPTY:
        return False
    if isinstance(v, str):
        if v.lower() in ("true", "vrai"):
            return True
        if v.lower() in ("false", "faux"):
            return False
    return vers_nombre(v) != 0


def serie_vers_datetime(s):
    return datetime(1899, 12, 30) + timedelta(days=float(s))


def datetime_vers_serie(d):
    return (d - datetime(1899, 12, 30)).total_seconds() / 86400.0


def cdate(v):
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return VDate(v)
    if isinstance(v, str):
        s = v.strip()
        for fmt in ("%d/%m/%Y", "%d/%m/%y", "%Y-%m-%d", "%d/%m/%Y %H:%M:%S"):
            try:
                return VDate(datetime_vers_serie(datetime.strptime(s, fmt)))
            except ValueError:
                pass
    erreur(13, f"Incompatibilité de type (CDate) : {v!r}")


def est_date(v):
    try:
        if isinstance(v, str):
            cdate(v)
            return True
        return isinstance(v, VDate)
    except VBAErreur:
        return False


def est_numerique(v):
    if isinstance(v, bool):
        return True
    if isinstance(v, (int, float)):
        return True
    if isinstance(v, str):
        s = v.strip()
        if not s:
            return False
        try:
            float(s.replace(",", "."))
            return True
        except ValueError:
            return False
    return v is EMPTY


def conv_type(type_, v):
    if type_ is None or type_ == "variant":
        return copier_valeur(v)
    if type_ in ("long", "integer", "byte", "longlong"):
        return cint(v)
    if type_ in ("double", "single", "currency"):
        return cdbl(v)
    if type_ == "string":
        return cstr(v)
    if type_ == "boolean":
        return cbool(v)
    if type_ == "date":
        return cdate(v) if not isinstance(v, VDate) else v
    return v


def defaut_type(type_):
    return {"long": 0, "integer": 0, "byte": 0, "longlong": 0, "double": 0.0, "single": 0.0, "currency": 0.0,
            "string": "", "boolean": False, "date": VDate(0)}.get(type_, NOTHING if type_ in OBJETS else EMPTY)


OBJETS = {"object", "listobject", "listcolumn", "worksheet", "workbook", "range", "chart", "shape", "querytable",
          "collection", "dictionary"}

# ---------------------------------------------------------------------------
# Lexer
# ---------------------------------------------------------------------------
_RE_TOKEN = re.compile(r"""
    (?P<ws>[ \t]+)
  | (?P<str>"(?:[^"]|"")*")
  | (?P<num>\d+\.\d*(?:[eE][+-]?\d+)?[#!]?|\d+(?:[eE][+-]?\d+)?[#&!]?|\.\d+)
  | (?P<hex>&H[0-9A-Fa-f]+&?)
  | (?P<id>[A-Za-z_][A-Za-z0-9_]*\$?)
  | (?P<op>:=|<>|<=|>=|[&+\-*/\\^=<>(),.:;])
""", re.X)


def tokeniser(ligne):
    toks, pos = [], 0
    while pos < len(ligne):
        m = _RE_TOKEN.match(ligne, pos)
        if not m:
            raise SyntaxError(f"caractère inattendu : {ligne[pos:pos + 20]!r}")
        pos = m.end()
        k = m.lastgroup
        t = m.group(k)
        if k == "ws":
            continue
        if k == "str":
            toks.append(("str", t[1:-1].replace('""', '"')))
        elif k == "num":
            t = t.rstrip("#&!")
            toks.append(("num", float(t) if ("." in t or "e" in t.lower()) else int(t)))
        elif k == "hex":
            toks.append(("num", int(t[2:].rstrip("&"), 16)))
        elif k == "id":
            toks.append(("id", t))
        else:
            toks.append(("op", t))
    return toks


def sans_commentaire(ligne):
    dans = False
    for i, ch in enumerate(ligne):
        if ch == '"':
            dans = not dans
        elif ch == "'" and not dans:
            return ligne[:i]
    return ligne


def lignes_logiques(source):
    phys = source.replace("\r\n", "\n").split("\n")
    res, i = [], 0
    while i < len(phys):
        num = i + 1
        ligne = sans_commentaire(phys[i])
        while ligne.rstrip().endswith(" _") and i + 1 < len(phys):
            i += 1
            ligne = ligne.rstrip()[:-1] + " " + sans_commentaire(phys[i]).strip()
        i += 1
        if ligne.strip():
            res.append((num, ligne))
    return res


# ---------------------------------------------------------------------------
# Expressions (AST : tuples)
# ---------------------------------------------------------------------------
PRIO_BIN = {
    "imp": 1, "eqv": 2, "xor": 3, "or": 4, "and": 5,
    "=": 7, "<>": 7, "<": 7, ">": 7, "<=": 7, ">=": 7, "like": 7, "is": 7,
    "&": 8, "+": 9, "-": 9, "mod": 10, "\\": 11, "*": 12, "/": 12, "^": 14,
}


class AnalyseurExpr:
    def __init__(self, toks, pos=0):
        self.toks = toks
        self.pos = pos

    def voir(self, k=0):
        i = self.pos + k
        return self.toks[i] if i < len(self.toks) else (None, None)

    def est(self, valeur, k=0):
        t = self.voir(k)
        return t[0] in ("op", "id") and isinstance(t[1], str) and t[1].lower() == valeur

    def prendre(self, valeur=None):
        t = self.voir()
        if valeur is not None and not self.est(valeur):
            raise SyntaxError(f"attendu « {valeur} », trouvé {t}")
        self.pos += 1
        return t

    def fin(self):
        return self.pos >= len(self.toks)

    def expression(self, prio_min=0):
        gauche = self.unaire()
        while True:
            t = self.voir()
            if t[0] is None:
                break
            op = t[1].lower() if isinstance(t[1], str) else None
            if t[0] not in ("op", "id") or op not in PRIO_BIN:
                break
            p = PRIO_BIN[op]
            if p < prio_min:
                break
            self.pos += 1
            droite = self.expression(p + 1 if op != "^" else p)
            gauche = ("bin", op, gauche, droite)
        return gauche

    def unaire(self):
        if self.est("not"):
            self.pos += 1
            return ("not", self.expression(6))
        if self.est("-"):
            self.pos += 1
            return ("neg", self.expression(13))
        if self.est("+"):
            self.pos += 1
            return self.expression(13)
        return self.postfixe()

    def arguments(self):
        """Après « ( » : liste d'arguments jusqu'à « ) » (vides et nommés admis)."""
        args = []
        if self.est(")"):
            self.pos += 1
            return args
        while True:
            if self.est(",") or self.est(")"):
                args.append(("missing",))
            elif self.voir()[0] == "id" and self.est(":=", 1):
                nom = self.prendre()[1].lower()
                self.prendre(":=")
                args.append(("nomme", nom, self.expression()))
            else:
                args.append(self.expression())
            if self.est(","):
                self.pos += 1
                continue
            self.prendre(")")
            return args

    def postfixe(self):
        t = self.prendre()
        if t[0] == "num":
            e = ("lit", t[1])
        elif t[0] == "str":
            e = ("lit", t[1])
        elif t[0] == "op" and t[1] == "(":
            e = ("paren", self.expression())
            self.prendre(")")
        elif t[0] == "op" and t[1] == ".":
            nom = self.prendre()[1]
            e = ("membre", ("with",), nom.lower())
        elif t[0] == "id":
            if t[1].lower() == "new":
                classe = self.prendre()[1].lower()
                return ("new", classe)
            e = ("id", t[1].lower().rstrip("$"))
        else:
            raise SyntaxError(f"expression inattendue : {t}")
        while True:
            if self.est("("):
                self.pos += 1
                e = ("appel", e, self.arguments())
            elif self.est("."):
                self.pos += 1
                nom = self.prendre()[1]
                e = ("membre", e, nom.lower().rstrip("$"))
            else:
                return e


# ---------------------------------------------------------------------------
# Instructions
# ---------------------------------------------------------------------------
def decouper_instructions(toks):
    """Découpe une ligne en instructions séparées par « : » (hors If sur une ligne)."""
    if toks and toks[0][0] == "id" and toks[0][1].lower() == "if":
        # If ... Then <instructions> : traité en bloc par l'analyseur
        prof, i_then = 0, None
        for i, t in enumerate(toks):
            if t == ("op", "("):
                prof += 1
            elif t == ("op", ")"):
                prof -= 1
            elif t[0] == "id" and t[1].lower() == "then" and prof == 0:
                i_then = i
                break
        if i_then is not None and i_then < len(toks) - 1:
            return [toks]
    res, cour, prof = [], [], 0
    for t in toks:
        if t == ("op", "("):
            prof += 1
        elif t == ("op", ")"):
            prof -= 1
        if t == ("op", ":") and prof == 0:
            if cour:
                res.append(cour)
            cour = []
        else:
            cour.append(t)
    if cour:
        res.append(cour)
    return res


def mot(toks, i=0):
    return toks[i][1].lower() if len(toks) > i and toks[i][0] == "id" else None


class Procedure:
    def __init__(self, module, nom, est_fonction, params, corps, public, type_retour):
        self.module = module
        self.nom = nom
        self.est_fonction = est_fonction
        self.params = params      # [(nom, byref, optionnel, defaut_expr, type, paramarray, est_tableau)]
        self.corps = corps
        self.public = public
        self.type_retour = type_retour


def analyser_declaration(toks):
    """« x As Long », « a() As String », « n As Long = 3 » -> (nom, type, tableau, bornes)."""
    p = AnalyseurExpr(toks)
    nom = p.prendre()[1].lower().rstrip("$")
    tableau, bornes = False, None
    if p.est("("):
        p.pos += 1
        tableau = True
        if not p.est(")"):
            bornes = []
            while True:
                b1 = p.expression()
                if p.est("to"):
                    p.pos += 1
                    b2 = p.expression()
                    bornes.append((b1, b2))
                else:
                    bornes.append((("lit", 0), b1))
                if p.est(","):
                    p.pos += 1
                    continue
                break
        p.prendre(")")
    type_ = None
    if p.est("as"):
        p.pos += 1
        if p.est("new"):
            p.pos += 1
        type_ = p.prendre()[1].lower()
    return nom, type_, tableau, bornes, p


def separer_virgules(toks):
    res, cour, prof = [], [], 0
    for t in toks:
        if t == ("op", "("):
            prof += 1
        elif t == ("op", ")"):
            prof -= 1
        if t == ("op", ",") and prof == 0:
            res.append(cour)
            cour = []
        else:
            cour.append(t)
    res.append(cour)
    return res


class AnalyseurModule:
    def __init__(self, nom_module, source):
        self.nom = nom_module
        self.lignes = []
        for num, ligne in lignes_logiques(source):
            toks = tokeniser(ligne)
            if not toks:
                continue
            # étiquette « Nom: » en début de ligne
            if len(toks) >= 2 and toks[0][0] == "id" and toks[1] == ("op", ":") and \
                    toks[0][1].lower() not in ("else",):
                self.lignes.append((num, [("label", toks[0][1].lower())]))
                toks = toks[2:]
                if not toks:
                    continue
            for instr in decouper_instructions(toks):
                self.lignes.append((num, instr))
        self.i = 0
        self.procedures = {}
        self.variables = []     # (nom, type, tableau, bornes, public, const_expr)

    def analyser(self):
        while self.i < len(self.lignes):
            num, toks = self.lignes[self.i]
            m0 = mot(toks)
            if m0 in ("attribute", "option"):
                self.i += 1
                continue
            public = True
            k = 0
            if m0 in ("public", "private", "global", "friend"):
                public = m0 != "private"
                k = 1
            if mot(toks, k) == "static":
                k += 1
            if mot(toks, k) in ("sub", "function"):
                self.procedure(toks, k, public)
                continue
            # déclarations de module
            if mot(toks, k) == "const" or m0 == "const":
                k = k + 1 if mot(toks, k) == "const" else 1
                for morceau in separer_virgules(toks[k:]):
                    nom, type_, _t, _b, p = analyser_declaration(morceau)
                    p.prendre("=")
                    self.variables.append((nom, type_, False, None, public, p.expression()))
            elif m0 in ("dim", "public", "private", "global"):
                k = max(k, 1)
                for morceau in separer_virgules(toks[k:]):
                    nom, type_, tab, bornes, _p = analyser_declaration(morceau)
                    self.variables.append((nom, type_, tab, bornes, public, None))
            else:
                raise SyntaxError(f"{self.nom}:{num} : instruction hors procédure")
            self.i += 1

    def procedure(self, toks, k, public):
        est_fonction = mot(toks, k) == "function"
        nom = toks[k + 1][1].lower()
        p = AnalyseurExpr(toks, k + 2)
        params = []
        p.prendre("(")
        if not p.est(")"):
            prof, debut, morceaux = 0, p.pos, []
            j = p.pos
            while j < len(toks):
                t = toks[j]
                if t == ("op", "("):
                    prof += 1
                elif t == ("op", ")"):
                    if prof == 0:
                        break
                    prof -= 1
                elif t == ("op", ",") and prof == 0:
                    morceaux.append(toks[debut:j])
                    debut = j + 1
                j += 1
            morceaux.append(toks[debut:j])
            p.pos = j
            for m in morceaux:
                byref, opt, pa = True, False, False
                while m and m[0][0] == "id" and m[0][1].lower() in ("optional", "byval", "byref", "paramarray"):
                    w = m[0][1].lower()
                    if w == "optional":
                        opt = True
                    elif w == "byval":
                        byref = False
                    elif w == "paramarray":
                        pa = True
                    m = m[1:]
                nom_p, type_, tab, _b, pp = analyser_declaration(m)
                defaut = None
                if pp.est("="):
                    pp.pos += 1
                    defaut = pp.expression()
                params.append((nom_p, byref, opt, defaut, type_, pa, tab))
        p.prendre(")")
        type_retour = None
        if p.est("as"):
            p.pos += 1
            type_retour = p.prendre()[1].lower()
        self.i += 1
        corps = self.bloc(("end sub", "end function"))
        self.i += 1  # End Sub / End Function
        self.procedures[nom] = Procedure(self.nom, nom, est_fonction, params, corps, public, type_retour)

    def debut(self, toks):
        a, b = mot(toks), mot(toks, 1)
        return f"{a} {b}" if b else (a or "")

    def bloc(self, fins):
        """Analyse des instructions jusqu'à l'un des mots de fin (non consommé)."""
        instrs = []
        while self.i < len(self.lignes):
            num, toks = self.lignes[self.i]
            d = self.debut(toks)
            premier = mot(toks)
            if any(d.startswith(f) or premier == f for f in fins):
                return instrs
            instrs.append(self.instruction(num, toks))
        raise SyntaxError(f"{self.nom} : fin de bloc manquante {fins}")

    def instruction(self, num, toks):
        if toks[0][0] == "label":
            self.i += 1
            return ("label", toks[0][1], num)
        m0 = mot(toks)
        if m0 == "if":
            return self.si(num, toks)
        if m0 == "for":
            return self.pour(num, toks)
        if m0 == "do":
            return self.faire(num, toks)
        if m0 == "while":
            self.i += 1
            cond = AnalyseurExpr(toks, 1).expression()
            corps = self.bloc(("wend",))
            self.i += 1
            return ("do", "while", cond, None, None, corps, num)
        if m0 == "select":
            return self.selon(num, toks)
        if m0 == "with":
            self.i += 1
            obj = AnalyseurExpr(toks, 1).expression()
            corps = self.bloc(("end with",))
            self.i += 1
            return ("with", obj, corps, num)
        self.i += 1
        return self.simple(num, toks)

    def simple(self, num, toks):
        m0 = mot(toks)
        if m0 in ("dim", "static"):
            decl = []
            for morceau in separer_virgules(toks[1:]):
                nom, type_, tab, bornes, _p = analyser_declaration(morceau)
                decl.append((nom, type_, tab, bornes))
            return ("dim", decl, num)
        if m0 == "const":
            decl = []
            for morceau in separer_virgules(toks[1:]):
                nom, type_, _t, _b, p = analyser_declaration(morceau)
                p.prendre("=")
                decl.append((nom, p.expression()))
            return ("const", decl, num)
        if m0 == "redim":
            k = 1
            preserve = False
            if mot(toks, 1) == "preserve":
                preserve, k = True, 2
            decl = []
            for morceau in separer_virgules(toks[k:]):
                nom, type_, _t, bornes, _p = analyser_declaration(morceau)
                decl.append((nom, bornes))
            return ("redim", preserve, decl, num)
        if m0 == "exit":
            return ("exit", mot(toks, 1), num)
        if m0 == "goto":
            return ("goto", toks[1][1].lower(), num)
        if m0 == "on" and mot(toks, 1) == "error":
            if mot(toks, 2) == "resume":
                return ("onerror", "next", None, num)
            if mot(toks, 2) == "goto":
                cible = toks[3][1]
                if cible == 0:
                    return ("onerror", "zero", None, num)
                return ("onerror", "goto", str(cible).lower(), num)
        avec_set = m0 == "set"
        if avec_set:
            toks = toks[1:]
        if m0 == "call":
            p = AnalyseurExpr(toks, 1)
            return ("appel", p.expression(), [], num)
        p = AnalyseurExpr(toks)
        cible = p.postfixe()
        if p.est("="):
            p.pos += 1
            valeur = p.expression()
            if not p.fin():
                raise SyntaxError(f"ligne {num} : fin d'affectation inattendue")
            return ("affecter", cible, valeur, num, avec_set)
        if p.fin():
            if cible[0] == "appel" and len(cible[2]) > 1:
                raise SyntaxError(f"{self.nom}:{num} : appel avec plusieurs arguments entre parenthèses sans « Call »")
            return ("appel", cible, None, num)
        # appel sans parenthèses : Proc a, b, nom:=c
        args = []
        reste = toks[p.pos:]
        for morceau in separer_virgules(reste):
            if not morceau:
                args.append(("missing",))
            elif morceau[0][0] == "id" and len(morceau) > 1 and morceau[1] == ("op", ":="):
                args.append(("nomme", morceau[0][1].lower(), AnalyseurExpr(morceau, 2).expression()))
            else:
                args.append(AnalyseurExpr(morceau).expression())
        return ("appel", cible, args, num)

    def si(self, num, toks):
        prof, i_then = 0, None
        for i, t in enumerate(toks):
            if t == ("op", "("):
                prof += 1
            elif t == ("op", ")"):
                prof -= 1
            elif t[0] == "id" and t[1].lower() == "then" and prof == 0:
                i_then = i
                break
        cond = AnalyseurExpr(toks[1:i_then]).expression()
        reste = toks[i_then + 1:]
        self.i += 1
        if reste:  # If sur une ligne
            prof, i_else = 0, None
            for i, t in enumerate(reste):
                if t == ("op", "("):
                    prof += 1
                elif t == ("op", ")"):
                    prof -= 1
                elif t[0] == "id" and t[1].lower() == "else" and prof == 0:
                    i_else = i
                    break
            alors = reste if i_else is None else reste[:i_else]
            sinon = [] if i_else is None else reste[i_else + 1:]
            b_alors = [self.simple(num, ins) for ins in decouper_instructions(alors)]
            b_sinon = [self.simple(num, ins) for ins in decouper_instructions(sinon)] if sinon else []
            return ("if", [(cond, b_alors)], b_sinon, num)
        branches = [(cond, self.bloc(("elseif", "else", "end if")))]
        sinon = []
        while True:
            n2, t2 = self.lignes[self.i]
            m = mot(t2)
            if m == "elseif":
                j = next(k for k, t in enumerate(t2) if t[0] == "id" and t[1].lower() == "then")
                c2 = AnalyseurExpr(t2[1:j]).expression()
                self.i += 1
                branches.append((c2, self.bloc(("elseif", "else", "end if"))))
            elif m == "else":
                self.i += 1
                sinon = self.bloc(("end if",))
            else:
                self.i += 1   # End If
                return ("if", branches, sinon, num)

    def pour(self, num, toks):
        self.i += 1
        if mot(toks, 1) == "each":
            var = toks[2][1].lower()
            coll = AnalyseurExpr(toks, 4).expression()
            corps = self.bloc(("next",))
            self.i += 1
            return ("foreach", var, coll, corps, num)
        var = toks[1][1].lower()
        p = AnalyseurExpr(toks, 3)
        debut = p.expression()
        p.prendre("to")
        fin = p.expression()
        pas = ("lit", 1)
        if p.est("step"):
            p.pos += 1
            pas = p.expression()
        corps = self.bloc(("next",))
        self.i += 1
        return ("for", var, debut, fin, pas, corps, num)

    def faire(self, num, toks):
        self.i += 1
        pre_type, pre = None, None
        if mot(toks, 1) in ("while", "until"):
            pre_type = mot(toks, 1)
            pre = AnalyseurExpr(toks, 2).expression()
        corps = self.bloc(("loop",))
        _n, tl = self.lignes[self.i]
        self.i += 1
        post_type, post = None, None
        if mot(tl, 1) in ("while", "until"):
            post_type = mot(tl, 1)
            post = AnalyseurExpr(tl, 2).expression()
        return ("do", pre_type, pre, post_type, post, corps, num)

    def selon(self, num, toks):
        expr = AnalyseurExpr(toks, 2).expression()
        self.i += 1
        cas, sinon = [], None
        while True:
            n2, t2 = self.lignes[self.i]
            if self.debut(t2) == "end select":
                self.i += 1
                return ("select", expr, cas, sinon, num)
            if mot(t2) != "case":
                raise SyntaxError(f"{self.nom}:{n2} : Case attendu")
            self.i += 1
            if mot(t2, 1) == "else":
                sinon = self.bloc(("case", "end select"))
                continue
            valeurs = []
            for morceau in separer_virgules(t2[1:]):
                if morceau and morceau[0][0] == "id" and morceau[0][1].lower() == "is":
                    op = morceau[1][1]
                    valeurs.append(("is", op, AnalyseurExpr(morceau, 2).expression()))
                else:
                    pe = AnalyseurExpr(morceau)
                    a = pe.expression()
                    if pe.est("to"):
                        pe.pos += 1
                        valeurs.append(("to", a, pe.expression()))
                    else:
                        valeurs.append(("eq", a))
            cas.append((valeurs, self.bloc(("case", "end select"))))


# ---------------------------------------------------------------------------
# Exécution
# ---------------------------------------------------------------------------
class Cellule:
    __slots__ = ("valeur", "type", "tableau")

    def __init__(self, valeur=EMPTY, type_=None, tableau=False):
        self.valeur = valeur
        self.type = type_
        self.tableau = tableau

    def affecter(self, v):
        if isinstance(v, VArray):
            self.valeur = v.copie()
        elif self.tableau and v is EMPTY:
            self.valeur = v
        else:
            self.valeur = conv_type(self.type, v)


class SautGoto(Exception):
    def __init__(self, etiquette):
        self.etiquette = etiquette


class Sortie(Exception):
    def __init__(self, genre):
        self.genre = genre


class Cadre:
    def __init__(self, proc, module):
        self.proc = proc
        self.module = module
        self.vars = {}
        self.mode_erreur = "zero"
        self.etiquette_erreur = None
        self.with_pile = []


class ObjetErr:
    def __init__(self):
        self.number = 0
        self.description = ""
        self.source = ""

    def membre(self, nom, args, interp):
        if nom == "number":
            return self.number
        if nom == "description":
            return self.description
        if nom == "source":
            return self.source
        if nom == "clear":
            self.number, self.description = 0, ""
            return EMPTY
        if nom == "raise":
            numero = interp_val(args[0]) if args else 0
            source = cstr(interp_val(args[1])) if len(args) > 1 and args[1] is not MISSING else ""
            desc = cstr(interp_val(args[2])) if len(args) > 2 and args[2] is not MISSING else "Erreur définie par l'application"
            raise VBAErreur(numero, desc, source)
        erreur(438, f"Err.{nom} non pris en charge")


def interp_val(v):
    return v


class Dictionnaire:
    """Scripting.Dictionary"""

    def __init__(self):
        self.d = {}
        self.comparemode = 0

    def _k(self, k):
        if isinstance(k, str) and self.comparemode == 1:
            return k.lower()
        if isinstance(k, float) and k.is_integer():
            return int(k)
        return k

    def defaut(self, args):
        k = self._k(args[0])
        if k not in self.d:
            self.d[k] = (args[0], EMPTY)
        return self.d[k][1]

    def defaut_set(self, args, valeur):
        k = self._k(args[0])
        orig = self.d[k][0] if k in self.d else args[0]
        self.d[k] = (orig, copier_valeur(valeur))

    def membre(self, nom, args, interp):
        if nom == "exists":
            return self._k(args[0]) in self.d
        if nom == "count":
            return len(self.d)
        if nom == "keys":
            return VArray.depuis_liste([v[0] for v in self.d.values()])
        if nom == "items":
            return VArray.depuis_liste([v[1] for v in self.d.values()])
        if nom == "item":
            return self.defaut(args)
        if nom == "add":
            k = self._k(args[0])
            if k in self.d:
                erreur(457, "Cette clé est déjà associée à un élément de cette collection")
            self.d[k] = (args[0], copier_valeur(args[1]))
            return EMPTY
        if nom == "remove":
            self.d.pop(self._k(args[0]), None)
            return EMPTY
        if nom == "removeall":
            self.d.clear()
            return EMPTY
        if nom == "comparemode":
            return self.comparemode
        erreur(438, f"Dictionary.{nom} non pris en charge")

    def membre_set(self, nom, valeur):
        if nom == "comparemode":
            self.comparemode = cint(valeur)
            return
        erreur(438, f"Dictionary.{nom} (affectation)")

    def iterer(self):
        return [v[0] for v in self.d.values()]


class Collection:
    def __init__(self):
        self.items = []

    def defaut(self, args):
        return self.items[cint(args[0]) - 1]

    def membre(self, nom, args, interp):
        if nom == "add":
            self.items.append(copier_valeur(args[0]))
            return EMPTY
        if nom == "count":
            return len(self.items)
        if nom == "item":
            return self.defaut(args)
        erreur(438, f"Collection.{nom}")

    def iterer(self):
        return list(self.items)


def like(s, motif):
    rx, i = "", 0
    while i < len(motif):
        c = motif[i]
        if c == "*":
            rx += ".*"
        elif c == "?":
            rx += "."
        elif c == "#":
            rx += r"\d"
        elif c == "[":
            j = motif.index("]", i)
            contenu = motif[i + 1:j]
            rx += "[^" + re.escape(contenu[1:]) + "]" if contenu.startswith("!") else "[" + contenu + "]"
            i = j
        else:
            rx += re.escape(c)
        i += 1
    return re.fullmatch(rx, s, re.S) is not None


def format_vba(v, fmt):
    if fmt is MISSING or fmt == "":
        return cstr(v)
    if re.fullmatch(r"0+", fmt):
        n = vers_nombre(v)
        n = int(math.floor(abs(n) + 0.5)) * (1 if n >= 0 else -1)
        s = str(abs(n)).rjust(len(fmt), "0")
        return ("-" if n < 0 else "") + s
    if any(x in fmt.lower() for x in ("yyyy", "dd", "hh", "mm")):
        d = serie_vers_datetime(vers_nombre(v))
        f = fmt
        f = f.replace("yyyy", "%Y").replace("dd", "%d").replace("hh", "%H").replace("nn", "%M").replace("ss", "%S")
        f = f.replace("mm", "%m")
        return d.strftime(f)
    erreur(5, f"Format non pris en charge : {fmt}")


class Interpreteur:
    def __init__(self, sources: dict[str, str], monde):
        """sources : {nom_module: code} ; monde : objet fournissant Application, ThisWorkbook, etc."""
        self.modules = {}
        self.procs = {}           # nom -> Procedure (public)
        self.procs_module = {}    # (module, nom) -> Procedure
        self.vars_module = {}     # module -> {nom: Cellule}
        self.publics = {}         # nom -> Cellule
        self.err = ObjetErr()
        self.monde = monde
        self.dir_iter = []
        self.appels = 0
        analyses = {}
        for nom, src in sources.items():
            a = AnalyseurModule(nom.lower(), src)
            a.analyser()
            analyses[nom.lower()] = a
            self.modules[nom.lower()] = a
            for pnom, proc in a.procedures.items():
                self.procs_module[(nom.lower(), pnom)] = proc
                if proc.public:
                    self.procs.setdefault(pnom, proc)
        # variables et constantes de module
        for nom, a in analyses.items():
            vm = {}
            self.vars_module[nom] = vm
            for vnom, type_, tab, bornes, public, const in a.variables:
                c = Cellule(defaut_type(type_), type_, tab)
                if tab:
                    c.valeur = VArray([], defaut_type(type_)) if bornes is None else EMPTY
                vm[vnom] = c
                if public:
                    self.publics[vnom] = c
        for nom, a in analyses.items():
            for vnom, type_, tab, bornes, public, const in a.variables:
                if const is not None:
                    cadre = Cadre(None, nom)
                    self.vars_module[nom][vnom].valeur = self.eval(const, cadre)

    # -- appels ---------------------------------------------------------------
    def appeler(self, nom, *args, module=None):
        proc = self.procs_module[(module.lower(), nom.lower())] if module else self.procs[nom.lower()]
        return self.executer_proc(proc, [Cellule(copier_valeur(a)) for a in args])

    def executer_proc(self, proc, cellules_args, nommes=None):
        self.appels += 1
        cadre = Cadre(proc, proc.module)
        nommes = nommes or {}
        for k, (pnom, byref, opt, defaut, type_, pa, tab) in enumerate(proc.params):
            if pa:
                restes = [c.valeur for c in cellules_args[k:]]
                cadre.vars[pnom] = Cellule(VArray.depuis_liste(restes))
                break
            if pnom in nommes:
                c = nommes[pnom]
            elif k < len(cellules_args) and cellules_args[k] is not None:
                c = cellules_args[k]
            else:
                if not opt:
                    erreur(449, f"Argument obligatoire manquant : {proc.nom}({pnom})")
                if defaut is not None:
                    c = Cellule(self.eval(defaut, cadre), type_)
                elif type_ in (None, "variant"):
                    c = Cellule(MISSING)
                else:
                    c = Cellule(defaut_type(type_), type_)
            if not byref or c is None:
                c = Cellule(copier_valeur(c.valeur), type_, tab)
                if type_ and not tab and c.valeur is not MISSING:
                    c.valeur = conv_type(type_, c.valeur)
            elif type_ and not tab and c.type is None and c.valeur is not MISSING and not isinstance(c.valeur, VArray):
                # argument ByRef typé : VBA exige le même type ; on convertit la valeur reçue
                pass
            cadre.vars[pnom] = c
        if proc.est_fonction:
            cadre.vars[proc.nom] = Cellule(defaut_type(proc.type_retour), proc.type_retour)
        try:
            self.executer_corps(proc.corps, cadre)
        except Sortie as s:
            if s.genre not in ("sub", "function"):
                raise
        return cadre.vars[proc.nom].valeur if proc.est_fonction else EMPTY

    def executer_corps(self, corps, cadre):
        """Corps de procédure : gère GoTo / étiquettes et gestionnaires d'erreurs."""
        etiquettes = {ins[1]: k for k, ins in enumerate(corps) if ins[0] == "label"}
        k = 0
        while k < len(corps):
            try:
                self.executer_bloc(corps[k:k + 1], cadre)
                k += 1
            except SautGoto as g:
                if g.etiquette not in etiquettes:
                    raise SyntaxError(f"étiquette inconnue {g.etiquette}")
                k = etiquettes[g.etiquette] + 1
            except VBAErreur as e:
                if cadre.mode_erreur == "goto":
                    self.err.number, self.err.description, self.err.source = e.numero, e.description, e.source
                    cadre.mode_erreur = "gestionnaire"
                    k = etiquettes[cadre.etiquette_erreur] + 1
                else:
                    raise

    def executer_bloc(self, instrs, cadre):
        for ins in instrs:
            try:
                self.executer(ins, cadre)
            except VBAErreur as e:
                if cadre.mode_erreur == "next":
                    self.err.number, self.err.description = e.numero, e.description
                    continue
                raise
            except (Sortie, SautGoto):
                raise
            except RecursionError:
                raise
            except Exception as e:  # erreur Python dans le monde simulé -> erreur VBA
                if isinstance(e, SyntaxError):
                    raise
                ve = VBAErreur(1004, f"{type(e).__name__}: {e} (ligne {ins[-1]} de {cadre.module})")
                if cadre.mode_erreur == "next":
                    self.err.number, self.err.description = ve.numero, ve.description
                    continue
                raise ve from e

    def executer(self, ins, cadre):
        genre = ins[0]
        if genre == "label":
            return
        if genre == "dim":
            for nom, type_, tab, bornes in ins[1]:
                c = Cellule(defaut_type(type_), type_, tab)
                if tab:
                    c.valeur = VArray([], defaut_type(type_)) if bornes is None else \
                        VArray([(cint(self.eval(a, cadre)), cint(self.eval(b, cadre))) for a, b in bornes], defaut_type(type_))
                    if bornes is None:
                        c.valeur.bornes = None
                cadre.vars[nom] = c
            return
        if genre == "const":
            for nom, expr in ins[1]:
                cadre.vars[nom] = Cellule(self.eval(expr, cadre))
            return
        if genre == "redim":
            for nom, bornes in ins[2]:
                c = self.cellule(nom, cadre, creer=True)
                b = [(cint(self.eval(x, cadre)), cint(self.eval(y, cadre))) for x, y in bornes]
                defaut = defaut_type(c.type)
                if isinstance(c.valeur, VArray) and c.valeur.bornes:
                    c.valeur.redim(b, ins[1])
                    c.valeur.defaut = defaut
                else:
                    c.valeur = VArray(b, defaut)
            return
        if genre == "affecter":
            valeur = self.eval(ins[2], cadre)
            est_objet = hasattr(valeur, "membre") or valeur is NOTHING
            if ins[4] and not est_objet:
                raise VBAErreur(424, f"Objet requis (Set d'une valeur non objet, ligne {ins[3]} de {cadre.module})")
            if not ins[4] and est_objet and ins[1][0] == "id":
                raise VBAErreur(91, f"Affectation d'un objet sans « Set » (ligne {ins[3]} de {cadre.module})")
            self.affecter(ins[1], valeur, cadre)
            return
        if genre == "appel":
            self.instruction_appel(ins[1], ins[2], cadre)
            return
        if genre == "if":
            for cond, bloc in ins[1]:
                if cbool(self.eval(cond, cadre)):
                    self.executer_bloc(bloc, cadre)
                    return
            self.executer_bloc(ins[2], cadre)
            return
        if genre == "for":
            _g, var, debut, fin, pas, corps, _n = ins
            c = self.cellule(var, cadre)
            v = vers_nombre(self.eval(debut, cadre))
            f = vers_nombre(self.eval(fin, cadre))
            p = vers_nombre(self.eval(pas, cadre))
            c.affecter(v)
            while (vers_nombre(c.valeur) <= f) if p >= 0 else (vers_nombre(c.valeur) >= f):
                try:
                    self.executer_bloc(corps, cadre)
                except Sortie as s:
                    if s.genre == "for":
                        return
                    raise
                c.affecter(vers_nombre(c.valeur) + p)
            return
        if genre == "foreach":
            _g, var, coll, corps, _n = ins
            obj = self.eval(coll, cadre)
            if isinstance(obj, VArray):
                elements = [obj.data.get(k, obj.defaut) for k in sorted(obj.data.keys())] \
                    if len(obj.bornes) == 1 else list(obj.data.values())
                if len(obj.bornes) == 1:
                    lb, ub = obj.bornes[0]
                    elements = [obj.get((i,)) for i in range(lb, ub + 1)]
            else:
                elements = obj.iterer()
            c = self.cellule(var, cadre)
            for e in elements:
                c.valeur = e
                try:
                    self.executer_bloc(corps, cadre)
                except Sortie as s:
                    if s.genre == "for":
                        return
                    raise
            return
        if genre == "do":
            _g, pre_t, pre, post_t, post, corps, _n = ins
            garde = 0
            while True:
                garde += 1
                if garde > 5_000_000:
                    raise RuntimeError("boucle infinie probable")
                if pre_t == "while" and not cbool(self.eval(pre, cadre)):
                    return
                if pre_t == "until" and cbool(self.eval(pre, cadre)):
                    return
                try:
                    self.executer_bloc(corps, cadre)
                except Sortie as s:
                    if s.genre == "do":
                        return
                    raise
                if post_t == "while" and not cbool(self.eval(post, cadre)):
                    return
                if post_t == "until" and cbool(self.eval(post, cadre)):
                    return
        if genre == "select":
            _g, expr, cas, sinon, _n = ins
            v = self.eval(expr, cadre)
            for valeurs, bloc in cas:
                for item in valeurs:
                    if item[0] == "eq" and egal(v, self.eval(item[1], cadre)):
                        self.executer_bloc(bloc, cadre)
                        return
                    if item[0] == "to" and comparer(v, self.eval(item[1], cadre)) >= 0 and \
                            comparer(v, self.eval(item[2], cadre)) <= 0:
                        self.executer_bloc(bloc, cadre)
                        return
                    if item[0] == "is" and cbool(operation(item[1], v, self.eval(item[2], cadre))):
                        self.executer_bloc(bloc, cadre)
                        return
            if sinon is not None:
                self.executer_bloc(sinon, cadre)
            return
        if genre == "with":
            cadre.with_pile.append(self.eval(ins[1], cadre))
            try:
                self.executer_bloc(ins[2], cadre)
            finally:
                cadre.with_pile.pop()
            return
        if genre == "exit":
            raise Sortie(ins[1])
        if genre == "goto":
            raise SautGoto(ins[1])
        if genre == "onerror":
            cadre.mode_erreur = ins[1]
            cadre.etiquette_erreur = ins[2]
            if ins[1] in ("zero", "next", "goto"):
                self.err.number, self.err.description = 0, ""
            return
        raise SyntaxError(f"instruction non prise en charge : {genre}")

    # -- variables -------------------------------------------------------------
    def cellule(self, nom, cadre, creer=False):
        if nom in cadre.vars:
            return cadre.vars[nom]
        vm = self.vars_module.get(cadre.module, {})
        if nom in vm:
            return vm[nom]
        if nom in self.publics:
            return self.publics[nom]
        if creer:
            c = Cellule()
            cadre.vars[nom] = c
            return c
        raise VBAErreur(0, f"VARIABLE_INCONNUE:{nom}")

    def cellule_ou_none(self, nom, cadre):
        try:
            return self.cellule(nom, cadre)
        except VBAErreur:
            return None

    # -- affectation ------------------------------------------------------------
    def affecter(self, cible, valeur, cadre):
        if cible[0] == "id":
            c = self.cellule_ou_none(cible[1], cadre)
            if c is None:
                raise SyntaxError(f"affectation à une variable non déclarée : {cible[1]} ({cadre.module})")
            c.affecter(valeur)
            return
        if cible[0] == "appel":
            base, args = cible[1], cible[2]
            vals = [self.eval(a, cadre) for a in args]
            if base[0] == "id":
                c = self.cellule_ou_none(base[1], cadre)
                if c is not None:
                    v = c.valeur
                    if isinstance(v, VArray):
                        t = c.type
                        v.set(vals, conv_type(t, valeur) if t and t != "variant" else copier_valeur(valeur))
                        return
                    if hasattr(v, "defaut_set"):
                        v.defaut_set(vals, valeur)
                        return
                    raise VBAErreur(13, f"affectation indexée sur {base[1]} ({v!r})")
            obj = self.eval(base, cadre)
            if hasattr(obj, "defaut_set"):
                obj.defaut_set(vals, valeur)
                return
            if base[0] == "membre":
                o = self.eval(base[1], cadre)
                o.membre_appel_set(base[2], vals, valeur)
                return
            raise VBAErreur(438, f"affectation non prise en charge : {cible}")
        if cible[0] == "membre":
            obj = self.eval(cible[1], cadre)
            obj.membre_set(cible[2], valeur)
            return
        raise SyntaxError(f"cible d'affectation : {cible}")

    # -- appels sous forme d'instruction --------------------------------------
    def instruction_appel(self, cible, args, cadre):
        if args is None:
            # « Proc » seul, ou « obj.Methode », ou « Proc(x) »
            self.eval(cible, cadre, instruction=True)
            return
        self.eval(("appel", cible, args), cadre, instruction=True)

    # -- évaluation ------------------------------------------------------------
    def args_cellules(self, proc, args, cadre):
        cellules, nommes = [], {}
        for k, a in enumerate(args):
            if a[0] == "missing":
                cellules.append(None)
                continue
            if a[0] == "nomme":
                nommes[a[1]] = Cellule(self.eval(a[2], cadre))
                continue
            byref, ptype, ptab = True, None, False
            if k < len(proc.params):
                byref, ptype, ptab = proc.params[k][1], proc.params[k][4], proc.params[k][6]
            if byref and a[0] == "id":
                c = self.cellule_ou_none(a[1], cadre)
                if c is not None:
                    if ptype not in (None, "variant") and not (ptype == "object" and c.type in OBJETS) \
                            and (c.type != ptype or bool(c.tableau) != bool(ptab)):
                        raise SyntaxError(f"Type d'argument ByRef incompatible : {a[1]} ({c.type}"
                                          f"{'()' if c.tableau else ''}) passé à {proc.nom}.{proc.params[k][0]} "
                                          f"({ptype}{'()' if ptab else ''}) [{cadre.module}]")
                    cellules.append(c)
                    continue
            cellules.append(Cellule(self.eval(a, cadre)))
        return cellules, nommes

    def trouver_proc(self, nom, cadre):
        if (cadre.module, nom) in self.procs_module:
            return self.procs_module[(cadre.module, nom)]
        return self.procs.get(nom)

    def eval(self, e, cadre, instruction=False):
        g = e[0]
        if g == "lit":
            return e[1]
        if g == "paren":
            return self.eval(e[1], cadre)
        if g == "neg":
            v = self.eval(e[1], cadre)
            return -vers_nombre(v)
        if g == "not":
            v = self.eval(e[1], cadre)
            if isinstance(v, bool):
                return not v
            return ~cint(v)
        if g == "bin":
            op = e[1]
            if op in ("and", "or"):
                a = self.eval(e[2], cadre)
                b = self.eval(e[3], cadre)
                if isinstance(a, bool) and isinstance(b, bool):
                    return (a and b) if op == "and" else (a or b)
                return (cint(a) & cint(b)) if op == "and" else (cint(a) | cint(b))
            if op == "is":
                return self.eval(e[2], cadre) is self.eval(e[3], cadre)
            return operation(op, self.eval(e[2], cadre), self.eval(e[3], cadre))
        if g == "new":
            if e[1] == "collection":
                return Collection()
            erreur(429, f"New {e[1]}")
        if g == "id":
            return self.eval_id(e[1], [], cadre, sans_parentheses=True)
        if g == "appel":
            base, args = e[1], e[2]
            if base[0] == "id":
                return self.eval_id(base[1], args, cadre)
            if base[0] == "membre":
                # module.Proc(args) ?
                if base[1][0] == "id" and base[1][1] in self.modules and self.cellule_ou_none(base[1][1], cadre) is None \
                        and (base[1][1], base[2]) in self.procs_module:
                    proc = self.procs_module.get((base[1][1], base[2]))
                    if proc is None:
                        erreur(438, f"{base[1][1]}.{base[2]} introuvable")
                    cel, nom = self.args_cellules(proc, args, cadre)
                    return self.executer_proc(proc, cel, nom)
                obj = self.eval(base[1], cadre)
                vals = [self.eval_arg(a, cadre) for a in args]
                return appeler_membre(obj, base[2], vals, self)
            obj = self.eval(base, cadre)
            vals = [self.eval_arg(a, cadre) for a in args]
            if hasattr(obj, "defaut"):
                return obj.defaut(vals)
            if isinstance(obj, VArray):
                return obj.get(vals)
            erreur(13, f"appel sur {obj!r}")
        if g == "membre":
            if e[1][0] == "with":
                obj = cadre.with_pile[-1]
            else:
                if e[1][0] == "id" and e[1][1] in self.modules and self.cellule_ou_none(e[1][1], cadre) is None \
                        and (e[1][1], e[2]) in self.procs_module:
                    proc = self.procs_module.get((e[1][1], e[2]))
                    if proc is None:
                        erreur(438, f"{e[1][1]}.{e[2]} introuvable")
                    return self.executer_proc(proc, [])
                obj = self.eval(e[1], cadre)
            return appeler_membre(obj, e[2], None, self)
        raise SyntaxError(f"expression : {e}")

    def eval_arg(self, a, cadre):
        if a[0] == "missing":
            return MISSING
        if a[0] == "nomme":
            return ("__nomme__", a[1], self.eval(a[2], cadre))
        return self.eval(a, cadre)

    def eval_id(self, nom, args, cadre, sans_parentheses=False):
        c = self.cellule_ou_none(nom, cadre)
        if c is not None and not (cadre.proc and nom == cadre.proc.nom and args and cadre.proc.est_fonction):
            v = c.valeur
            if not args:
                return v
            vals = [self.eval_arg(a, cadre) for a in args]
            if isinstance(v, VArray):
                return v.get(vals)
            if hasattr(v, "defaut"):
                return v.defaut(vals)
            erreur(13, f"{nom} n'est pas un tableau")
        proc = self.trouver_proc(nom, cadre)
        if proc is not None:
            cel, nommes = self.args_cellules(proc, args, cadre)
            return self.executer_proc(proc, cel, nommes)
        vals = [self.eval_arg(a, cadre) for a in args]
        return self.intrinseque(nom, vals, cadre)

    # -- fonctions intégrées -----------------------------------------------------
    def intrinseque(self, nom, a, cadre):
        m = self.monde
        consts = CONSTANTES
        if nom in consts:
            return consts[nom]
        if nom == "err":
            return self.err
        if nom == "true":
            return True
        if nom == "false":
            return False
        if nom == "empty":
            return EMPTY
        if nom == "nothing":
            return NOTHING
        if nom == "null":
            return NULL
        if nom in ("application",):
            return m.application
        if nom == "thisworkbook":
            return m.classeur
        if nom == "activesheet":
            return m.feuille_active()
        if nom == "activecell":
            return m.cellule_active
        if nom == "workbooks":
            return m.workbooks
        if nom == "intersect":
            return m.intersect(a[0], a[1])
        f = FONCTIONS.get(nom)
        if f is not None:
            return f(self, *a)
        erreur(0, f"FONCTION_INCONNUE:{nom}")


def egal(a, b):
    return comparer(a, b) == 0


def comparer(a, b):
    if a is EMPTY and isinstance(b, str):
        a = ""
    if b is EMPTY and isinstance(a, str):
        b = ""
    if isinstance(a, str) and isinstance(b, str):
        return (a > b) - (a < b)
    if isinstance(a, str) or isinstance(b, str):
        # comparaison chaîne / nombre : VBA convertit la chaîne si possible
        try:
            x, y = vers_nombre(a), vers_nombre(b)
        except VBAErreur:
            x, y = cstr(a), cstr(b)
        return (x > y) - (x < y)
    x, y = vers_nombre(a), vers_nombre(b)
    return (x > y) - (x < y)


def operation(op, a, b):
    if op == "&":
        return cstr(a if a is not NULL else "") + cstr(b if b is not NULL else "")
    if op == "+":
        if isinstance(a, str) and isinstance(b, str):
            return a + b
        r = vers_nombre(a) + vers_nombre(b)
        return VDate(r) if isinstance(a, VDate) or isinstance(b, VDate) else r
    if op == "-":
        r = vers_nombre(a) - vers_nombre(b)
        if isinstance(a, VDate) and not isinstance(b, VDate):
            return VDate(r)
        return r
    if op == "*":
        return vers_nombre(a) * vers_nombre(b)
    if op == "/":
        y = vers_nombre(b)
        if y == 0:
            erreur(11, "Division par zéro")
        return vers_nombre(a) / y
    if op == "\\":
        y = cint(b)
        if y == 0:
            erreur(11, "Division par zéro")
        x = cint(a)
        return int(x / y)
    if op == "mod":
        y = cint(b)
        x = cint(a)
        return int(math.fmod(x, y))
    if op == "^":
        return vers_nombre(a) ** vers_nombre(b)
    if op in ("=", "<>", "<", ">", "<=", ">="):
        c = comparer(a, b)
        return {"=": c == 0, "<>": c != 0, "<": c < 0, ">": c > 0, "<=": c <= 0, ">=": c >= 0}[op]
    if op == "like":
        return like(cstr(a), cstr(b))
    if op == "xor":
        return cbool(a) != cbool(b)
    raise SyntaxError(f"opérateur {op}")


def appeler_membre(obj, nom, args, interp):
    if obj is NOTHING or obj is EMPTY or obj is None:
        erreur(91, f"Variable objet non définie (.{nom})")
    if not hasattr(obj, "membre"):
        erreur(424, f"Objet requis pour .{nom} ({obj!r})")
    positionnels, nommes = [], {}
    for x in (args or []):
        if isinstance(x, tuple) and len(x) == 3 and x[0] == "__nomme__":
            nommes[x[1]] = x[2]
        else:
            positionnels.append(x)
    if nommes:
        return obj.membre(nom, positionnels, interp, **({"nommes": nommes}))
    return obj.membre(nom, positionnels if args is not None else [], interp)


# ---------------------------------------------------------------------------
# Constantes et fonctions intégrées
# ---------------------------------------------------------------------------
CONSTANTES = {
    "vbcrlf": "\r\n", "vblf": "\n", "vbcr": "\r", "vbtab": "\t", "vbnullstring": "",
    "vbinformation": 64, "vbexclamation": 48, "vbquestion": 32, "vbcritical": 16,
    "vbokonly": 0, "vbokcancel": 1, "vbabortretryignore": 2, "vbyesnocancel": 3, "vbyesno": 4, "vbretrycancel": 5,
    "vbok": 1, "vbcancel": 2, "vbabort": 3, "vbretry": 4, "vbignore": 5, "vbyes": 6, "vbno": 7,
    "vbbinarycompare": 0, "vbtextcompare": 1, "vbobjecterror": -2147221504, "vbdirectory": 16,
    "vbempty": 0, "vbinteger": 2, "vblong": 3, "vbsingle": 4, "vbdouble": 5, "vbdate": 7, "vbstring": 8,
    "vbboolean": 11, "vbvariant": 12, "vbarray": 8192,
    "xlcalculationmanual": -4135, "xlcalculationautomatic": -4105, "xlshiftup": -4162, "xlnorestrictions": 0,
    "xlsheetvisible": -1, "xlsheethidden": 0, "xlsheetveryhidden": 2, "xlcmdsql": 2, "xlinsertdeletecells": 1,
    "xlsrcquery": 3, "xlsrcrange": 1,
}


def _len(i, s):
    return len(cstr(s))


def _mid(i, s, debut, longueur=MISSING):
    s = cstr(s)
    d = cint(debut) - 1
    if longueur is MISSING:
        return s[d:]
    return s[d:d + cint(longueur)]


def _instr(i, *a):
    a = [x for x in a]
    if len(a) >= 3 and not isinstance(a[0], str):
        debut, s1, s2 = cint(a[0]), cstr(a[1]), cstr(a[2])
        mode = cint(a[3]) if len(a) > 3 else 0
    else:
        debut, s1, s2 = 1, cstr(a[0]), cstr(a[1])
        mode = cint(a[2]) if len(a) > 2 else 0
    if mode == 1:
        s1, s2 = s1.lower(), s2.lower()
    if s2 == "":
        return debut
    return s1.find(s2, debut - 1) + 1


def _instrrev(i, s, cherche, debut=MISSING):
    s, c = cstr(s), cstr(cherche)
    return s.rfind(c) + 1


def _replace(i, s, a, b, *rest):
    return cstr(s).replace(cstr(a), cstr(b))


def _split(i, s, sep=" ", *rest):
    s = cstr(s)
    if s == "":
        a = VArray([(0, -1)])
        return a
    return VArray.depuis_liste(s.split(cstr(sep)))


def _ubound(i, a, dim=1):
    if not isinstance(a, VArray):
        erreur(13, "UBound sur une valeur qui n'est pas un tableau")
    if not a.bornes:
        erreur(9, "UBound sur un tableau non dimensionné")
    return a.bornes[cint(dim) - 1][1]


def _lbound(i, a, dim=1):
    if not isinstance(a, VArray) or not a.bornes:
        erreur(9, "LBound sur un tableau non dimensionné")
    return a.bornes[cint(dim) - 1][0]


def _round(i, x, n=0):
    x = vers_nombre(x)
    n = cint(n)
    return round(x, n) if n else float(round(x))


def _string(i, n, ch):
    return cstr(ch)[0] * cint(n)


def _str(i, x):
    x = vers_nombre(x)
    s = cstr(float(x)) if not isinstance(x, int) else str(x)
    return s if s.startswith("-") else " " + s


def _val(i, s):
    m = re.match(r"\s*[-+]?\d*\.?\d+", cstr(s))
    return float(m.group(0)) if m else 0.0


def _vartype(i, v):
    if v is EMPTY:
        return 0
    if isinstance(v, bool):
        return 11
    if isinstance(v, str):
        return 8
    if isinstance(v, VDate):
        return 7
    if isinstance(v, int):
        return 3
    if isinstance(v, float):
        return 5
    if isinstance(v, VArray):
        return 8204
    return 9


def _trim(i, s):
    return cstr(s).strip(" ")


def _dir(interp, motif=MISSING, attributs=0):
    if motif is MISSING:
        if not interp.dir_iter:
            return ""
        return interp.dir_iter.pop(0)
    chemin = interp.monde.chemin_local(cstr(motif))
    if cint(attributs) & 16:
        interp.dir_iter = []
        return os.path.basename(chemin.rstrip("/")) if os.path.isdir(chemin) else ""
    trouves = sorted(os.path.basename(p) for p in _glob.glob(chemin) if os.path.isfile(p))
    interp.dir_iter = trouves[1:]
    return trouves[0] if trouves else ""


def _getattr(interp, chemin):
    p = interp.monde.chemin_local(cstr(chemin))
    if os.path.isdir(p):
        return 16
    if os.path.exists(p):
        return 0
    erreur(53, "Fichier introuvable")


def _mkdir(interp, chemin):
    p = interp.monde.chemin_local(cstr(chemin))
    if os.path.exists(p):
        erreur(75, "Erreur d'accès au chemin")
    os.mkdir(p)
    return EMPTY


def _kill(interp, chemin):
    os.remove(interp.monde.chemin_local(cstr(chemin)))
    return EMPTY


def _filedatetime(interp, chemin):
    p = interp.monde.chemin_local(cstr(chemin))
    return VDate(datetime_vers_serie(datetime.fromtimestamp(os.path.getmtime(p))))


def _iif(i, c, a, b):
    return a if cbool(c) else b


def _strcomp(i, a, b, mode=0):
    a, b = cstr(a), cstr(b)
    if cint(mode) == 1:
        a, b = a.lower(), b.lower()
    return (a > b) - (a < b)


def _array(i, *a):
    return VArray.depuis_liste(list(a))


def _isarray(i, v):
    return isinstance(v, VArray)


def _createobject(i, classe):
    if cstr(classe).lower() == "scripting.dictionary":
        return Dictionnaire()
    erreur(429, f"CreateObject({classe})")


def _timeserial(i, h, mi, s):
    return VDate((cint(h) * 3600 + cint(mi) * 60 + cint(s)) / 86400.0)


FONCTIONS = {
    "len": _len, "mid": _mid, "left": lambda i, s, n: cstr(s)[:cint(n)],
    "right": lambda i, s, n: cstr(s)[len(cstr(s)) - cint(n):] if cint(n) else "",
    "trim": _trim, "ltrim": lambda i, s: cstr(s).lstrip(" "), "rtrim": lambda i, s: cstr(s).rstrip(" "),
    "ucase": lambda i, s: cstr(s).upper(), "lcase": lambda i, s: cstr(s).lower(),
    "instr": _instr, "instrrev": _instrrev, "replace": _replace, "split": _split,
    "join": lambda i, a, sep=" ": cstr(sep).join(cstr(x) for x in a.lignes()),
    "ubound": _ubound, "lbound": _lbound, "cstr": lambda i, v: cstr(v), "clng": lambda i, v: cint(v),
    "cint": lambda i, v: cint(v), "cdbl": lambda i, v: cdbl(v), "cbool": lambda i, v: cbool(v),
    "cdate": lambda i, v: cdate(v), "int": lambda i, v: math.floor(vers_nombre(v)),
    "fix": lambda i, v: math.trunc(vers_nombre(v)), "abs": lambda i, v: abs(vers_nombre(v)),
    "sgn": lambda i, v: (vers_nombre(v) > 0) - (vers_nombre(v) < 0),
    "round": _round, "isnumeric": lambda i, v: est_numerique(v), "isdate": lambda i, v: est_date(v),
    "isempty": lambda i, v: v is EMPTY, "iserror": lambda i, v: False, "isnull": lambda i, v: v is NULL,
    "ismissing": lambda i, v: v is MISSING, "isarray": _isarray, "isobject": lambda i, v: hasattr(v, "membre"),
    "string": _string, "str": _str, "val": _val, "format": lambda i, v, f=MISSING: format_vba(v, f),
    "iif": _iif, "strcomp": _strcomp, "array": _array, "vartype": _vartype,
    "environ": lambda i, n: {"username": "testeur"}.get(cstr(n).lower(), ""),
    "now": lambda i: VDate(i.monde.maintenant), "date": lambda i: VDate(math.floor(i.monde.maintenant)),
    "doevents": lambda i: 0, "dir": _dir, "getattr": _getattr, "mkdir": _mkdir, "kill": _kill,
    "filedatetime": _filedatetime, "createobject": _createobject, "timeserial": _timeserial,
    "msgbox": lambda i, *a: i.monde.msgbox(*a), "inputbox": lambda i, *a: i.monde.inputbox(*a),
    "chr": lambda i, n: chr(cint(n)), "chrw": lambda i, n: chr(cint(n)), "space": lambda i, n: " " * cint(n),
    "typename": lambda i, v: type(v).__name__,
}
