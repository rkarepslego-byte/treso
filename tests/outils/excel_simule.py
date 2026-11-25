"""Modèle objet Excel simulé pour exécuter le VBA de l'outil avec vba_interpreteur.

Le classeur simulé est construit à partir du VRAI classeur généré (lu avec openpyxl) :
mêmes tables, mêmes colonnes, mêmes noms définis. Seuls les éléments nécessaires au
code de l'outil sont simulés : ListObject / ListColumns / DataBodyRange / Resize / Delete,
noms définis, requêtes Power Query (l'actualisation d'une table de transit est faite
par le moteur de référence à partir des fichiers du dossier indiqué dans la requête
Parametres_Import générée par le VBA), Application, fichiers.
"""

from __future__ import annotations

import os
import re
from datetime import date, datetime

import openpyxl

from tests.outils.vba_interpreteur import (EMPTY, MISSING, NOTHING, VArray, VBAErreur, cint, cstr, datetime_vers_serie,
                                           erreur, serie_vers_datetime, vers_nombre)


def vers_excel(v):
    if v is None:
        return EMPTY
    if isinstance(v, datetime):
        return datetime_vers_serie(v)
    if isinstance(v, date):
        return datetime_vers_serie(datetime(v.year, v.month, v.day))
    return v


class Objet:
    """Base : accès aux membres par nom en minuscules."""

    def membre(self, nom, args, interp, nommes=None):
        f = getattr(self, "m_" + nom, None)
        if f is not None:
            return f(*args, **(nommes or {}))
        p = getattr(self, "p_" + nom, MISSING)
        if p is not MISSING:
            if args:
                erreur(450, f"{type(self).__name__}.{nom} : arguments inattendus")
            return p
        erreur(438, f"{type(self).__name__} ne gère pas la propriété ou méthode « {nom} »")

    def membre_set(self, nom, valeur):
        f = getattr(self, "s_" + nom, None)
        if f is not None:
            return f(valeur)
        if hasattr(self, "p_" + nom):
            setattr(self, "p_" + nom, valeur)
            return
        erreur(438, f"{type(self).__name__} : affectation de « {nom} » non gérée")


class Collection(Objet):
    def __init__(self, elements, cle=lambda e: e.p_name):
        self.elements = elements
        self.cle = cle

    def item(self, x):
        if isinstance(x, str):
            for e in self.elements:
                if str(self.cle(e)).lower() == x.lower():
                    return e
            erreur(9, f"L'indice n'appartient pas à la sélection ({x})")
        return self.elements[cint(x) - 1]

    def m_item(self, x):
        return self.item(x)

    def defaut(self, args):
        return self.item(args[0])

    @property
    def p_count(self):
        return len(self.elements)

    def iterer(self):
        return list(self.elements)


# ---------------------------------------------------------------------------
# Tables
# ---------------------------------------------------------------------------
class Table(Objet):
    def __init__(self, feuille, nom, entetes, lignes, formules, ligne_entete, col_debut):
        self.feuille = feuille
        self.p_name = nom
        self.entetes = entetes
        self.lignes = lignes            # liste de listes (valeurs « Value2 »)
        self.formules = formules        # [bool] par colonne
        self.ligne_entete = ligne_entete
        self.col_debut = col_debut
        self.p_showautofilter = True
        self.requete = None             # pour les tables de transit
        self.feuille_ws = None

    # -- membres
    def m_listcolumns(self, *args):
        cols = Collection([Colonne(self, j) for j in range(len(self.entetes))])
        return cols.item(args[0]) if args else cols

    @property
    def p_listrows(self):
        return ObjetCompte(len(self.lignes))

    @property
    def p_databodyrange(self):
        if not self.lignes:
            return NOTHING
        return Plage(self, 0, len(self.lignes), 0, len(self.entetes))

    @property
    def p_range(self):
        return Plage(self, -1, len(self.lignes) + 1, 0, len(self.entetes))

    @property
    def p_headerrowrange(self):
        return Plage(self, -1, 1, 0, len(self.entetes))

    @property
    def p_parent(self):
        return self.feuille

    @property
    def p_autofilter(self):
        return Filtre()

    @property
    def p_querytable(self):
        if self.requete is None:
            erreur(1004, "Cette table n'est pas reliée à une requête")
        return TableRequete(self)

    @property
    def p_sourcetype(self):
        return 3 if self.requete else 1

    def m_resize(self, plage):
        n = plage.nr - 1
        while len(self.lignes) < n:
            self.lignes.append([EMPTY] * len(self.entetes))
        del self.lignes[n:]
        return EMPTY

    def m_delete(self):
        self.feuille.tables.remove(self)
        return EMPTY

    # -- accès Python
    def dicts(self):
        return [dict(zip(self.entetes, l)) for l in self.lignes]


class ObjetCompte(Objet):
    def __init__(self, n):
        self.p_count = n


class Filtre(Objet):
    p_filtermode = False

    def m_showalldata(self):
        return EMPTY


class Colonne(Objet):
    def __init__(self, table, j):
        self.table = table
        self.j = j
        self.p_name = table.entetes[j]
        self.p_index = j + 1

    @property
    def p_databodyrange(self):
        if not self.table.lignes:
            return NOTHING
        return Plage(self.table, 0, len(self.table.lignes), self.j, 1)


class Plage(Objet):
    """Zone rectangulaire d'une table ; r0 = -1 désigne la ligne d'en-tête."""

    def __init__(self, table, r0, nr, c0, nc):
        self.table, self.r0, self.nr, self.c0, self.nc = table, r0, nr, c0, nc

    def _val(self, r, c):
        if r < 0:
            return self.table.entetes[c]
        if r >= len(self.table.lignes):
            return EMPTY
        return self.table.lignes[r][c]

    def _lire(self):
        if self.nr == 1 and self.nc == 1:
            return self._val(self.r0, self.c0)
        a = VArray([(1, self.nr), (1, self.nc)])
        for i in range(self.nr):
            for j in range(self.nc):
                a.data[(i + 1, j + 1)] = self._val(self.r0 + i, self.c0 + j)
        return a

    def _ecrire(self, v):
        for i in range(self.nr):
            for j in range(self.nc):
                r, c = self.r0 + i, self.c0 + j
                if r < 0:
                    continue
                if self.table.formules[c]:
                    erreur(1004, f"écriture d'une valeur dans la colonne calculée « {self.table.entetes[c]} »")
                x = v.get((i + 1, j + 1)) if isinstance(v, VArray) else v
                if x is MISSING or x is None:
                    x = EMPTY
                self.table.lignes[r][c] = x

    p_value2 = property(_lire)
    p_value = property(_lire)

    def s_value2(self, v):
        self._ecrire(v)

    def s_value(self, v):
        self._ecrire(v)

    @property
    def p_hasformula(self):
        return self.table.formules[self.c0]

    @property
    def p_formular1c1(self):
        return "=<<formule>>" if self.table.formules[self.c0] else cstr(self._val(self.r0, self.c0))

    def s_formular1c1(self, v):
        if not all(self.table.formules[c] for c in range(self.c0, self.c0 + self.nc)):
            erreur(1004, "formule écrite dans une colonne de données")

    @property
    def p_row(self):
        return self.table.ligne_entete + 1 + self.r0

    @property
    def p_count(self):
        return self.nr * self.nc

    def m_cells(self, i, j):
        return Plage(self.table, self.r0 + cint(i) - 1, 1, self.c0 + cint(j) - 1, 1)

    def m_rows(self, k):
        return Plage(self.table, self.r0 + cint(k) - 1, 1, self.c0, self.nc)

    def m_columns(self, k):
        return Plage(self.table, self.r0, self.nr, self.c0 + cint(k) - 1, 1)

    def m_resize(self, nr, nc=MISSING):
        return Plage(self.table, self.r0, cint(nr), self.c0, self.nc if nc is MISSING else cint(nc))

    def m_delete(self, shift=MISSING):
        if self.r0 < 0:
            erreur(1004, "suppression de l'en-tête")
        del self.table.lignes[self.r0:self.r0 + self.nr]
        return EMPTY

    def m_clearcontents(self):
        self._ecrire(EMPTY)
        return EMPTY


class TableRequete(Objet):
    def __init__(self, table):
        self.table = table
        self.p_backgroundquery = True

    def m_refresh(self, *a, **k):
        self.table.feuille.classeur.monde.rafraichir(self.table)
        return EMPTY


# ---------------------------------------------------------------------------
# Feuilles, noms, classeur, application
# ---------------------------------------------------------------------------
class Feuille(Objet):
    def __init__(self, classeur, nom):
        self.classeur = classeur
        self.p_name = nom
        self.tables = []
        self.p_visible = -1
        self.p_enableselection = 0
        self.protegee = False

    def m_listobjects(self, *args):
        c = Collection(self.tables)
        return c.item(args[0]) if args else c

    def m_protect(self, *a, **k):
        self.protegee = True
        return EMPTY

    def m_unprotect(self, *a, **k):
        self.protegee = False
        return EMPTY

    def m_activate(self):
        self.classeur.monde.active = self
        return EMPTY

    def m_chartobjects(self, *a):
        return Collection([])

    @property
    def p_shapes(self):
        return Collection([])

    def m_range(self, ref):
        return CelluleNommee(None, EMPTY)


class CelluleNommee(Objet):
    def __init__(self, nom, valeur):
        self.nom = nom
        self.valeur = valeur

    p_value2 = property(lambda self: self.valeur)
    p_value = property(lambda self: self.valeur)

    def s_value2(self, v):
        self.valeur = EMPTY if v is MISSING else v

    s_value = s_value2

    def s_formula(self, v):
        self.valeur = v

    def m_cells(self, i, j):
        return self

    def m_select(self):
        return EMPTY

    @property
    def p_refersto(self):
        return self.nom


class Nom(Objet):
    def __init__(self, cellule):
        self.p_refersToRange = cellule

    @property
    def p_referstorange(self):
        return self.p_refersToRange


class Requete(Objet):
    def __init__(self, nom, formule):
        self.p_name = nom
        self.p_formula = formule


class Classeur(Objet):
    def __init__(self, monde, chemin):
        self.monde = monde
        self.feuilles = []
        self.noms = {}
        self.requetes = []
        self.p_path = os.path.dirname(chemin)
        self.p_name = os.path.basename(chemin)
        self.p_saved = True
        self.copies = []

    def m_worksheets(self, *args):
        c = Collection(self.feuilles)
        return c.item(args[0]) if args else c

    def m_names(self, nom):
        if nom.lower() not in self.noms:
            erreur(1004, f"nom défini introuvable : {nom}")
        return Nom(self.noms[nom.lower()])

    def m_queries(self, *args):
        c = CollectionRequetes(self.requetes)
        return c.item(args[0]) if args else c

    def m_savecopyas(self, chemin):
        p = self.monde.chemin_local(cstr(chemin))
        with open(p, "wb") as f:
            f.write(b"copie")
        self.copies.append(p)
        return EMPTY

    def m_unprotect(self, *a):
        return EMPTY

    def m_protect(self, *a, **k):
        return EMPTY

    def m_close(self, **k):
        return EMPTY


class CollectionRequetes(Collection):
    def __init__(self, requetes):
        super().__init__(requetes)

    def m_add(self, nom, formule, description=MISSING):
        r = Requete(nom, formule)
        self.elements.append(r)
        return r


class FonctionsFeuille(Objet):
    def m_edate(self, d, n):
        dt = serie_vers_datetime(vers_nombre(d))
        from reference.moteur import edate
        r = edate(dt.date(), cint(n))
        return datetime_vers_serie(datetime(r.year, r.month, r.day))

    def m_trim(self, s):
        return " ".join(cstr(s).split(" ")).strip() if s is not EMPTY else ""

    def m_counta(self, plage):
        v = plage._lire()
        valeurs = v.data.values() if isinstance(v, VArray) else [v]
        return sum(1 for x in valeurs if x is not EMPTY and x != "")

    def m_countif(self, plage, critere):
        v = plage._lire()
        valeurs = v.data.values() if isinstance(v, VArray) else [v]
        return sum(1 for x in valeurs if x == critere)


class Application(Objet):
    def __init__(self, monde):
        self.monde = monde
        self.p_calculation = -4105
        self.p_screenupdating = True
        self.p_enableevents = True
        self.p_statusbar = False
        self.p_displayalerts = True
        self.p_worksheetfunction = FonctionsFeuille()

    def m_calculate(self):
        self.monde.calculer()
        return EMPTY

    def m_ontime(self, *a, **k):
        return EMPTY

    def m_getopenfilename(self, *a, **k):
        return self.monde.fichier_a_ouvrir if self.monde.fichier_a_ouvrir else False


class ClasseurExterne(Objet):
    """Classeur ouvert par Workbooks.Open (import de prévisions)."""

    def __init__(self, chemin):
        wb = openpyxl.load_workbook(chemin, data_only=True)
        ws = wb.worksheets[0]
        lignes = [[vers_excel(v) for v in r] for r in ws.iter_rows(values_only=True)]
        self.feuille = FeuilleExterne(lignes)

    def m_worksheets(self, *args):
        return self.feuille

    def m_close(self, *a, **k):
        return EMPTY


class FeuilleExterne(Objet):
    def __init__(self, lignes):
        self.lignes = lignes

    @property
    def p_usedrange(self):
        nc = max(len(l) for l in self.lignes)
        a = VArray([(1, len(self.lignes)), (1, nc)])
        for i, l in enumerate(self.lignes):
            for j in range(nc):
                a.data[(i + 1, j + 1)] = l[j] if j < len(l) else EMPTY
        return ObjetValeur(a)


class ObjetValeur(Objet):
    def __init__(self, v):
        self.p_value2 = v
        self.p_value = v


class Classeurs(Objet):
    def m_open(self, chemin, **k):
        return ClasseurExterne(cstr(chemin))


# ---------------------------------------------------------------------------
# Monde simulé
# ---------------------------------------------------------------------------
class Monde:
    def __init__(self, chemin_xlsm, valeurs_calculees, maintenant, rafraichir_transit=None):
        self.application = Application(self)
        self.classeur = Classeur(self, chemin_xlsm)
        self.workbooks = Classeurs()
        self.maintenant = maintenant
        self.valeurs_calculees = {k.lower(): v for k, v in valeurs_calculees.items()}
        self.rafraichir_transit = rafraichir_transit
        self.messages = []
        self.reponses_inputbox = []
        self.fichier_a_ouvrir = None
        self.active = None
        self.cellule_active = NOTHING
        self._charger(chemin_xlsm)

    def chemin_local(self, chemin):
        return chemin.replace("\\", "/")

    def _charger(self, chemin):
        wb = openpyxl.load_workbook(chemin)
        cellules = {}
        for ws in wb.worksheets:
            f = Feuille(self.classeur, ws.title)
            self.classeur.feuilles.append(f)
            for nom_table, ref in ws.tables.items():
                (c1, r1), (c2, r2) = [openpyxl.utils.cell.coordinate_from_string(x) for x in ref.split(":")]
                j1 = openpyxl.utils.cell.column_index_from_string(c1)
                j2 = openpyxl.utils.cell.column_index_from_string(c2)
                entetes = [ws.cell(r1, j).value for j in range(j1, j2 + 1)]
                lignes, formules = [], [False] * len(entetes)
                for r in range(r1 + 1, r2 + 1):
                    ligne = []
                    for k, j in enumerate(range(j1, j2 + 1)):
                        v = ws.cell(r, j).value
                        if isinstance(v, str) and v.startswith("="):
                            formules[k] = True
                            v = None
                        ligne.append(vers_excel(v))
                    lignes.append(ligne)
                t = Table(f, nom_table, entetes, lignes, formules, r1, j1)
                if nom_table in ("tbl_Stage_B1", "tbl_Stage_B2"):
                    t.requete = True
                f.tables.append(t)
        for nom, dn in wb.defined_names.items():
            dest = list(dn.destinations)
            if len(dest) != 1:
                continue
            feuille, coord = dest[0]
            coord = coord.replace("$", "")
            if ":" in coord:
                continue
            v = wb[feuille][coord].value
            if isinstance(v, str) and v.startswith("="):
                v = self.valeurs_calculees.get(nom.lower(), EMPTY)
            self.classeur.noms[nom.lower()] = CelluleNommee(nom, vers_excel(v))
        # requêtes présentes dans le classeur livré
        for nom in ("Parametres_Import", "Import_Banque1", "Import_Banque2"):
            self.classeur.requetes.append(Requete(nom, ""))

    def table(self, nom):
        for f in self.classeur.feuilles:
            for t in f.tables:
                if t.p_name.lower() == nom.lower():
                    return t
        raise KeyError(nom)

    def nom(self, nom):
        return self.classeur.noms[nom.lower()].valeur

    def rafraichir(self, table):
        formule = next(r.p_formula for r in self.classeur.requetes if r.p_name == "Parametres_Import")
        self.rafraichir_transit(self, table, formule)

    def calculer(self):
        pass

    def feuille_active(self):
        return self.active

    def intersect(self, a, b):
        if a is NOTHING or b is NOTHING or not isinstance(a, Plage) or not isinstance(b, Plage) or a.table is not b.table:
            return NOTHING
        if b.r0 <= a.r0 < b.r0 + b.nr and b.c0 <= a.c0 < b.c0 + b.nc:
            return a
        return NOTHING

    def msgbox(self, *a):
        self.messages.append(a[0])
        return 1

    def inputbox(self, *a):
        return self.reponses_inputbox.pop(0) if self.reponses_inputbox else ""
