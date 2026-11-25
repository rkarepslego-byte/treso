"""Générateur du classeur Tresorerie.xlsm.

Usage :
    python -m build.generer_classeur                 # dist/ (classeur + relevés de démonstration)
    python -m build.generer_classeur --date 2026-11-25 --sortie dist

Le classeur est construit avec XlsxWriter, puis post-traité (build/integration.py) pour
y ajouter les requêtes Power Query. Le projet VBA est compilé par build/vba_project.py.
"""

from __future__ import annotations

import argparse
import glob
import os
import re
import shutil
from datetime import date, datetime, timedelta

import xlsxwriter
from xlsxwriter.utility import xl_col_to_name, xl_rowcol_to_cell

from build import datamashup as dm
from build import donnees_demo as dd
from build import integration
from build import schema as S
from build import vba_project as vp
from reference import moteur as m
from reference.formats import FORMAT_B1, FORMAT_B2, PARAMETRES_IMPORT

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC_PQ = os.path.join(RACINE, "src", "powerquery")
SRC_VBA = os.path.join(RACINE, "src", "vba")

# Palette
BLEU_NUIT = "#1E3A5F"
BLEU = "#2563EB"
BLEU_CLAIR = "#EAF1FB"
AMBRE = "#D97706"
VERT = "#15803D"
ROUGE = "#B91C1C"
GRIS = "#64748B"
GRIS_CLAIR = "#F1F5F9"
BORDURE = "#CBD5E1"
JAUNE_SAISIE = "#FFFBEB"
POLICE = "Segoe UI"

ORDRE_REQUETES = ["Parametres_Import", "fnTexteNormalise", "fnLireMontant", "fnLireDate",
                  "fnNormaliserReleve", "fnImporterBanque", "Import_Banque1", "Import_Banque2"]


# ---------------------------------------------------------------------------
# Paramètres Power Query (même texte que modImport.FormuleParametresM)
# ---------------------------------------------------------------------------
def _m_texte(s) -> str:
    return '"' + str(s).replace('"', '""') + '"'


def _m_liste(s: str, extensions: bool) -> str:
    vals = []
    for v in str(s).split(";"):
        v = v.strip()
        if not v:
            continue
        if extensions:
            v = v.lower()
            v = v if v.startswith(".") else "." + v
        else:
            v = m.normaliser_texte(v)
        vals.append(_m_texte(v))
    return "{" + ", ".join(vals) + "}"


def _record_banque(code: str, nom: str, dossier: str, f: dict) -> str:
    return ("[Code = " + _m_texte(code) + ", Banque = " + _m_texte(nom) + ", Dossier = " + _m_texte(dossier)
            + ", Extensions = " + _m_liste(f["Extensions"], True) + ", Separateur = " + _m_texte(f["Separateur"])
            + ", Encodage = " + str(int(f["Encodage"])) + ", LignesIgnorees = " + str(int(f["LignesIgnorees"]))
            + ", Feuille = " + _m_texte(f["Feuille"]) + ", ColDate = " + _m_texte(f["ColDate"])
            + ", ColDateValeur = " + _m_texte(f["ColDateValeur"]) + ", FormatDate = " + _m_texte(f["FormatDate"])
            + ", ColsLibelle = " + _m_texte(f["ColsLibelle"]) + ", ModeMontant = " + _m_texte(f["ModeMontant"])
            + ", ColMontant = " + _m_texte(f["ColMontant"]) + ", ColDebit = " + _m_texte(f["ColDebit"])
            + ", ColCredit = " + _m_texte(f["ColCredit"]) + ", ColSens = " + _m_texte(f["ColSens"])
            + ", ValeursDebit = " + _m_liste(f["ValeursDebit"], False)
            + ", InverserSigne = " + ("true" if str(f["InverserSigne"]).lower() == "oui" else "false")
            + ", SeparateurDecimal = " + _m_texte(f["SeparateurDecimal"])
            + ", ColReference = " + _m_texte(f["ColReference"]) + "]")


def formule_parametres_m(nom_b1: str, nom_b2: str, dossier_b1: str = "", dossier_b2: str = "") -> str:
    return ("// Généré automatiquement par la macro Actualiser à partir de l'onglet Param_Import.\r\n"
            "// Ne pas modifier ici : modifier la table tbl_ParamImport.\r\n"
            "let\r\n    Parametres = [\r\n"
            "        Banque1 = " + _record_banque("B1", nom_b1, dossier_b1, FORMAT_B1) + ",\r\n"
            "        Banque2 = " + _record_banque("B2", nom_b2, dossier_b2, FORMAT_B2) + "\r\n"
            "    ]\r\nin\r\n    Parametres")


def requetes_power_query(nom_b1: str, nom_b2: str) -> list[dm.Requete]:
    req = []
    for nom in ORDRE_REQUETES:
        if nom == "Parametres_Import":
            code = formule_parametres_m(nom_b1, nom_b2)
            req.append(dm.Requete(nom, code, type_resultat="Record"))
            continue
        with open(os.path.join(SRC_PQ, nom + ".pq"), encoding="utf-8") as f:
            code = f.read().strip()
        if nom == "Import_Banque1":
            req.append(dm.Requete(nom, code, "tbl_Stage_B1", S.COLONNES_TRANSIT))
        elif nom == "Import_Banque2":
            req.append(dm.Requete(nom, code, "tbl_Stage_B2", S.COLONNES_TRANSIT))
        else:
            req.append(dm.Requete(nom, code, type_resultat="Function"))
    return req


def formule_explicite(formule: str, table: str) -> str:
    """[@Col] / [@[Col avec espace]] -> Table[[#This Row],[Col]] (notation enregistrée dans le fichier)."""
    formule = re.sub(r"\[@\[([^\]]+)\]\]", lambda mm: f"{table}[[#This Row],[{mm.group(1)}]]", formule)
    formule = re.sub(r"\[@([^\[\]]+)\]", lambda mm: f"{table}[[#This Row],[{mm.group(1)}]]", formule)
    return formule


# ---------------------------------------------------------------------------
# Projet VBA
# ---------------------------------------------------------------------------
def modules_vba() -> list[dict]:
    with open(os.path.join(SRC_VBA, "ThisWorkbook.cls"), encoding="utf-8") as f:
        mods = [vp.module_document("ThisWorkbook", f.read(), True)]
    for _nom, code, _visible in S.ONGLETS:
        mods.append(vp.module_document(code, "", False))
    for chemin in sorted(glob.glob(os.path.join(SRC_VBA, "*.bas"))):
        with open(chemin, encoding="utf-8") as f:
            mods.append(vp.module_standard(f.read()))
    return mods


# ---------------------------------------------------------------------------
# Générateur
# ---------------------------------------------------------------------------
class Generateur:
    def __init__(self, chemin: str, etat: m.Etat, avec_vba: bool = True):
        self.chemin = chemin
        self.etat = etat
        self.avec_vba = avec_vba
        self.wb = xlsxwriter.Workbook(chemin, {"use_future_functions": True, "strings_to_numbers": False})
        self.wb.set_properties({"title": "Trésorerie", "subject": "Outil de gestion de trésorerie",
                                "comments": "Généré par build/generer_classeur.py"})
        self.f = {}
        self.ws = {}
        self._formats()

    # -- formats ---------------------------------------------------------------
    def fmt(self, **kw):
        base = {"font_name": POLICE, "font_size": 10}
        base.update(kw)
        return self.wb.add_format(base)

    def _formats(self):
        f = self.f
        f["titre"] = self.fmt(bold=True, font_size=18, font_color=BLEU_NUIT)
        f["sous_titre"] = self.fmt(font_size=10, font_color=GRIS, italic=True)
        f["section"] = self.fmt(bold=True, font_size=11, font_color="white", bg_color=BLEU_NUIT, indent=1, valign="vcenter")
        f["texte"] = self.fmt(text_wrap=True, valign="top")
        f["texte_gris"] = self.fmt(font_color=GRIS, text_wrap=True, valign="top", font_size=9)
        f["gras"] = self.fmt(bold=True)
        f["libelle"] = self.fmt(font_color=BLEU_NUIT, bold=True, valign="vcenter")
        f["entete_saisie"] = self.fmt(bold=True, font_color="white", bg_color=BLEU, text_wrap=True, valign="vcenter")
        f["entete_calc"] = self.fmt(bold=True, font_color="white", bg_color=GRIS, text_wrap=True, valign="vcenter")
        f["saisie"] = self.fmt(locked=False, bg_color=JAUNE_SAISIE, border=1, border_color=BORDURE)
        f["saisie_date"] = self.fmt(locked=False, bg_color=JAUNE_SAISIE, border=1, border_color=BORDURE, num_format=S.FMT_DATE)
        f["saisie_eur"] = self.fmt(locked=False, bg_color=JAUNE_SAISIE, border=1, border_color=BORDURE, num_format=S.FMT_EUR)
        f["saisie_pct"] = self.fmt(locked=False, bg_color=JAUNE_SAISIE, border=1, border_color=BORDURE, num_format=S.FMT_PCT)
        f["saisie_ent"] = self.fmt(locked=False, bg_color=JAUNE_SAISIE, border=1, border_color=BORDURE, num_format="0")
        f["calc"] = self.fmt(bg_color=GRIS_CLAIR, border=1, border_color=BORDURE)
        f["eur"] = self.fmt(num_format=S.FMT_EUR)
        f["date"] = self.fmt(num_format=S.FMT_DATE)
        f["dateheure"] = self.fmt(num_format=S.FMT_DATEHEURE)
        # cartes du tableau de bord
        f["carte_titre"] = self.fmt(bold=True, font_color=GRIS, bg_color=BLEU_CLAIR, align="center", valign="vcenter",
                                    font_size=9, top=2, top_color=BLEU)
        f["carte_valeur"] = self.fmt(bold=True, font_size=18, font_color=BLEU_NUIT, bg_color=BLEU_CLAIR, align="center",
                                     valign="vcenter", num_format=S.FMT_EUR0)
        f["carte_valeur_n"] = self.fmt(bold=True, font_size=18, font_color=BLEU_NUIT, bg_color=BLEU_CLAIR, align="center",
                                       valign="vcenter", num_format="0")
        f["carte_detail"] = self.fmt(font_color=GRIS, bg_color=BLEU_CLAIR, align="center", valign="vcenter", font_size=8,
                                     bottom=1, bottom_color=BORDURE)
        f["carte_detail_date"] = self.fmt(font_color=GRIS, bg_color=BLEU_CLAIR, align="center", valign="vcenter",
                                          font_size=8, bottom=1, bottom_color=BORDURE, num_format='"au "dd/mm/yyyy')
        f["tab_entete"] = self.fmt(bold=True, font_color="white", bg_color=BLEU_NUIT, align="center", border=1,
                                   border_color=BORDURE)
        f["tab_texte"] = self.fmt(border=1, border_color=BORDURE)
        f["tab_eur"] = self.fmt(border=1, border_color=BORDURE, num_format=S.FMT_EUR)
        f["tab_eur_gras"] = self.fmt(border=1, border_color=BORDURE, num_format=S.FMT_EUR, bold=True)
        f["tab_date"] = self.fmt(border=1, border_color=BORDURE, num_format="ddd dd/mm/yyyy")
        f["rouge"] = self.wb.add_format({"font_color": ROUGE, "bold": True})
        f["fond_rouge"] = self.wb.add_format({"bg_color": "#FEE2E2", "font_color": ROUGE})
        f["fond_orange"] = self.wb.add_format({"bg_color": "#FEF3C7", "font_color": "#92400E"})
        f["fond_vert"] = self.wb.add_format({"bg_color": "#DCFCE7", "font_color": "#166534"})
        f["fond_bleu"] = self.wb.add_format({"bg_color": "#DBEAFE", "font_color": "#1E40AF"})
        f["fond_violet"] = self.wb.add_format({"bg_color": "#F3E8FF", "font_color": "#6B21A8"})
        f["gris_barre"] = self.wb.add_format({"font_color": "#94A3B8", "font_strikeout": True})

    # -- outils ----------------------------------------------------------------
    def bouton(self, ws, cellule: str, texte: str, macro: str, largeur=150, hauteur=30, couleur=BLEU, dx=0, dy=0):
        ws.insert_textbox(cellule, texte, {
            "width": largeur, "height": hauteur, "x_offset": dx, "y_offset": dy,
            "font": {"name": POLICE, "bold": True, "color": "white", "size": 10},
            "align": {"vertical": "middle", "horizontal": "center"},
            "fill": {"color": couleur}, "line": {"none": True},
            "description": "btn:" + macro, "object_position": 3,
        })

    def nom(self, nom: str, onglet: str, ligne0: int, col0: int, ligne1: int | None = None, col1: int | None = None):
        ref = "'" + onglet + "'!" + xl_rowcol_to_cell(ligne0, col0, True, True)
        if ligne1 is not None:
            ref += ":" + xl_rowcol_to_cell(ligne1, col1, True, True)
        self.wb.define_name(nom, "=" + ref)

    def entete_onglet(self, ws, titre: str, sous_titre: str, largeur_fusion: int = 10):
        ws.hide_gridlines(2)
        ws.set_row(0, 8)
        ws.set_column(0, 0, 2)
        ws.write(1, 1, titre, self.f["titre"])
        ws.set_row(1, 30)
        ws.write(2, 1, sous_titre, self.f["sous_titre"])

    # -- tables ----------------------------------------------------------------
    def valeur(self, v):
        return v

    def ecrire_table(self, ws, t: S.Table, lignes: list[list], nb_libres: int = 0):
        donnees = [list(r) for r in lignes]
        nbcol = len(t.colonnes)
        donnees += [[None] * nbcol for _ in range(nb_libres)]
        if not donnees:
            donnees = [[None] * nbcol]
        r0 = t.ligne_entete - 1
        c0 = t.col_debut
        colonnes = []
        for j, c in enumerate(t.colonnes):
            props = {"num_format": c.fmt} if c.fmt else {}
            props["locked"] = c.genre != S.SAISIE
            fmt_cellule = self.fmt(**props)
            spec = {"header": c.nom, "format": fmt_cellule,
                    "header_format": self.f["entete_saisie"] if c.genre == S.SAISIE else self.f["entete_calc"]}
            if c.formule:
                spec["formula"] = formule_explicite(c.formule, t.nom)
            colonnes.append(spec)
            options = {"hidden": True} if c.masquee else {}
            ws.set_column(c0 + j, c0 + j, c.largeur, fmt_cellule, options)
            if c.liste:
                ws.data_validation(r0 + 1, c0 + j, r0 + 20000, c0 + j, {
                    "validate": "list", "source": "=" + c.liste, "ignore_blank": True,
                    "error_title": "Valeur non autorisée", "error_message": "Choisissez une valeur dans la liste.",
                })
            elif c.genre == S.SAISIE and c.fmt == S.FMT_DATE:
                ws.data_validation(r0 + 1, c0 + j, r0 + 20000, c0 + j, {
                    "validate": "date", "criteria": "between", "minimum": date(2000, 1, 1), "maximum": date(2100, 12, 31),
                    "error_title": "Date invalide", "error_message": "Saisissez une date (jj/mm/aaaa)."})
            elif c.genre == S.SAISIE and c.fmt == S.FMT_EUR:
                ws.data_validation(r0 + 1, c0 + j, r0 + 20000, c0 + j, {
                    "validate": "decimal", "criteria": "between", "minimum": -1e12, "maximum": 1e12,
                    "error_title": "Montant invalide", "error_message": "Saisissez un nombre."})
        data_sans_formules = []
        for r in donnees:
            data_sans_formules.append([None if t.colonnes[j].formule else r[j] for j in range(nbcol)])
        ws.add_table(r0, c0, r0 + len(donnees), c0 + nbcol - 1, {
            "name": t.nom, "style": t.style, "columns": colonnes, "data": data_sans_formules, "autofilter": True,
        })
        ws.set_row(r0, 32)
        return r0, c0, len(donnees)

    def plage_colonne(self, t: S.Table, nom_col: str, nb_lignes: int = 20000) -> str:
        j = t.col_debut + t.index(nom_col)
        r = t.ligne_entete
        return f"{xl_col_to_name(j)}{r + 1}:{xl_col_to_name(j)}{r + nb_lignes}"

    def mise_en_forme_statut(self, ws, t: S.Table, nom_col: str, regles: dict, nb_lignes: int = 20000):
        plage = self.plage_colonne(t, nom_col, nb_lignes)
        for valeur, fmt in regles.items():
            ws.conditional_format(plage, {"type": "cell", "criteria": "==", "value": f'"{valeur}"', "format": self.f[fmt]})

    # =========================================================================
    def generer(self):
        e = self.etat
        for nom, code, visible in S.ONGLETS:
            ws = self.wb.add_worksheet(nom)
            ws.set_vba_name(code) if self.avec_vba else None
            if not visible:
                ws.very_hidden()
            self.ws[nom] = ws
        if self.avec_vba:
            self.wb.set_vba_name("ThisWorkbook")
        self.onglet_parametres()
        self.onglet_listes()
        self.onglet_param_import()
        self.onglet_saisie()
        self.onglet_charges()
        self.onglet_banque()
        self.onglet_rapprochement()
        self.onglet_atraiter()
        self.onglets_transit()
        self.onglet_historique()
        self.onglet_calculs()
        self.onglet_suivi()
        self.onglet_synthese()
        self.onglet_tableau_de_bord()
        self.onglet_aide()
        for nom, _code, visible in S.ONGLETS:
            if visible:
                ws = self.ws[nom]
                ws.set_landscape()
                ws.set_paper(9)
                ws.fit_to_pages(1, 0)
                ws.set_margins(0.4, 0.4, 0.5, 0.5)
                ws.set_footer("&L&8" + nom + "&R&8Page &P / &N")
                self.ws[nom].protect(S.MOT_DE_PASSE, {"autofilter": True, "format_columns": True, "format_rows": True,
                                                      "select_locked_cells": True, "select_unlocked_cells": True})
        self.ws[S.F_TDB].activate()
        if self.avec_vba:
            chemin_bin = self.chemin + ".vbaProject.bin"
            with open(chemin_bin, "wb") as fbin:
                fbin.write(vp.construire_projet(modules_vba()))
            self.wb.add_vba_project(chemin_bin)
        self.wb.close()
        if self.avec_vba:
            os.remove(self.chemin + ".vbaProject.bin")
        noms = [b.nom for b in e.params.banques]
        integration.post_traiter(self.chemin, requetes_power_query(noms[0], noms[1]))

    # -------------------------------------------------------------------------
    def onglet_parametres(self):
        ws = self.ws[S.F_PARAM]
        self.entete_onglet(ws, "Paramètres", "Cellules jaunes : à renseigner. Les autres cellules sont protégées.")
        ws.set_column(1, 1, 50)
        ws.set_column(2, 2, 38)
        ws.set_column(3, 3, 70)
        ws.set_column(4, 4, 3)
        valeurs = {
            "B1_Nom": self.etat.params.banques[0].nom, "B2_Nom": self.etat.params.banques[1].nom,
            "B1_SoldeInit": self.etat.params.banques[0].solde_initial,
            "B2_SoldeInit": self.etat.params.banques[1].solde_initial,
            "B1_DateSolde": self.etat.params.banques[0].date_solde,
            "B2_DateSolde": self.etat.params.banques[1].date_solde,
            "Tolerance_Jours": self.etat.params.tolerance, "Ecart_Certain": self.etat.params.ecart_certain,
            "Seuil_Similarite": self.etat.params.seuil_similarite,
            "Auto_Valider": "Oui" if self.etat.params.auto_valider else "Non",
            "Horizon_Mois": self.etat.params.horizon_mois,
            "Apprendre_Regles": "Oui" if self.etat.params.apprendre_regles else "Non",
            "Traitement_Auto": "Oui" if self.etat.params.traitement_auto else "Non",
            "Virements_Auto": "Oui" if self.etat.params.virements_auto else "Non",
            "Seuil_Alerte": self.etat.params.seuil_alerte,
        }
        if getattr(self.etat, "societe", ""):
            valeurs["Societe"] = self.etat.societe
        if getattr(self.etat, "date_calcul_fixe", False):
            valeurs["Date_Ref_Saisie"] = self.etat.params.date_ref
        ligne = 4
        for section, params in S.PARAMETRES:
            ws.merge_range(ligne, 1, ligne, 3, section, self.f["section"])
            ws.set_row(ligne, 20)
            ligne += 1
            for nom, libelle, defaut, fmt, aide in params:
                ws.write(ligne, 1, libelle, self.f["libelle"])
                v = valeurs.get(nom, defaut)
                style = {S.FMT_DATE: "saisie_date", S.FMT_EUR: "saisie_eur", S.FMT_PCT: "saisie_pct",
                         S.FMT_ENTIER: "saisie_ent"}.get(fmt, "saisie")
                if isinstance(v, date):
                    ws.write_datetime(ligne, 2, datetime(v.year, v.month, v.day), self.f[style])
                elif v is None:
                    ws.write_blank(ligne, 2, None, self.f[style])
                else:
                    ws.write(ligne, 2, v, self.f[style])
                ws.write(ligne, 3, aide, self.f["texte_gris"])
                self.nom(nom, S.F_PARAM, ligne, 2)
                if nom in ("Auto_Valider", "Apprendre_Regles", "Traitement_Auto", "Virements_Auto", "Actualiser_Ouverture"):
                    ws.data_validation(ligne, 2, ligne, 2, {"validate": "list", "source": ["Oui", "Non"]})
                elif nom in ("B1_DateSolde", "B2_DateSolde", "Date_Ref_Saisie"):
                    ws.data_validation(ligne, 2, ligne, 2, {"validate": "date", "criteria": "between",
                                                            "minimum": date(2000, 1, 1), "maximum": date(2100, 12, 31)})
                elif fmt == S.FMT_ENTIER:
                    ws.data_validation(ligne, 2, ligne, 2, {"validate": "integer", "criteria": "between",
                                                            "minimum": 0, "maximum": 3650})
                ws.set_row(ligne, 30)
                ligne += 1
            ligne += 1
        # Catégories et règles
        ws.set_column(5, 6, 22)
        ws.set_column(7, 7, 3)
        ws.write(3, 5, "Catégories", self.f["gras"])
        ws.write(3, 8, "Règles automatiques (mot-clé du libellé bancaire -> catégorie, action)", self.f["gras"])
        cats = [[c, t] for c, t in (getattr(self.etat, "categories", None) or S.CATEGORIES_DEFAUT)] + [[None, None]] * 28
        self.ecrire_table(ws, S.CATEGORIES, cats)
        source = self.etat.regles or [m.Regle(k, c, a) for k, c, a in S.REGLES_DEFAUT]
        regles = [[r.mot_cle, r.categorie, r.action or None] for r in source] + [[None, None, None]] * 30
        self.ecrire_table(ws, S.REGLES, regles)
        c_cat = S.CATEGORIES.col_debut
        self.wb.define_name("L_Categories",
                            f"=OFFSET('{S.F_PARAM}'!${xl_col_to_name(c_cat)}${S.CATEGORIES.ligne_entete + 1},0,0,"
                            f"MAX(1,COUNTA(tbl_Categories[Catégorie])),1)")
        ws.freeze_panes(4, 0)

    def onglet_listes(self):
        ws = self.ws[S.F_LISTES]
        col = 0
        ws.write(0, col, "L_Banques")
        ws.write_formula(1, col, "=B1_Nom", None, self.etat.params.banques[0].nom)
        ws.write_formula(2, col, "=B2_Nom", None, self.etat.params.banques[1].nom)
        self.nom("L_Banques", S.F_LISTES, 1, col, 2, col)
        col += 1
        ws.write(0, col, "L_BanquesSuivi")
        ws.write(1, col, "Toutes")
        ws.write_formula(2, col, "=B1_Nom", None, self.etat.params.banques[0].nom)
        ws.write_formula(3, col, "=B2_Nom", None, self.etat.params.banques[1].nom)
        self.nom("L_BanquesSuivi", S.F_LISTES, 1, col, 3, col)
        col += 1
        for nom, valeurs in S.LISTES.items():
            ws.write(0, col, nom)
            for i, v in enumerate(valeurs):
                ws.write(1 + i, col, v)
            self.nom(nom, S.F_LISTES, 1, col, len(valeurs), col)
            col += 1
        ws.write(0, 12, "Nom précédent B1")
        ws.write(1, 12, self.etat.params.banques[0].nom)
        self.nom("B1_NomPrecedent", S.F_LISTES, 1, 12)
        ws.write(2, 12, "Nom précédent B2")
        ws.write(3, 12, self.etat.params.banques[1].nom)
        self.nom("B2_NomPrecedent", S.F_LISTES, 3, 12)
        rejets = [[o, p, None] for o, p in sorted(self.etat.rejets)]
        self.ecrire_table(ws, S.REJETS, rejets)

    def onglet_param_import(self):
        ws = self.ws[S.F_PARAM_IMPORT]
        lignes = [[cle, lib, str(FORMAT_B1[cle]), str(FORMAT_B2[cle]), desc] for cle, lib, desc in PARAMETRES_IMPORT]
        self.ecrire_table(ws, S.PARAM_IMPORT, lignes)
        code_m = []
        for r in requetes_power_query(*[b.nom for b in self.etat.params.banques]):
            feuille = {"tbl_Stage_B1": S.F_IMPORT_B1, "tbl_Stage_B2": S.F_IMPORT_B2}.get(r.table or "", "")
            code_m.append([r.nom, r.table or "", feuille, r.code])
        self.ecrire_table(ws, S.CODE_M, code_m)

    # -------------------------------------------------------------------------
    def lignes_previsions(self) -> list[list]:
        t = S.PREVISIONS
        res = []
        for p in self.etat.previsions:
            r = [None] * len(t.colonnes)
            v = {"ID": p.id, "Date prévue": p.date, "Banque": p.banque, "Libellé": p.libelle, "Catégorie": p.categorie,
                 "Type": p.type, "Montant": abs(p.montant), "Commentaire": p.commentaire or None, "Statut": p.statut,
                 "Origine": p.origine, "Réf charge": p.ref_charge or None, "ID opération": p.id_operation or None,
                 "ID proposé": p.id_propose or None, "Date réelle": p.date_reelle, "Montant réel": p.montant_reel,
                 "Lien virement": p.lien_virement or None, "Empreinte": p.empreinte or None}
            for k, val in v.items():
                r[t.index(k)] = val
            res.append(r)
        return res

    def onglet_saisie(self):
        ws = self.ws[S.F_SAISIE]
        self.entete_onglet(ws, "Saisie des opérations prévisionnelles",
                           "Saisissez une opération dans le formulaire puis cliquez « Ajouter », ou complétez directement "
                           "les lignes vides en bas du tableau. Dépenses : montant positif, le signe est appliqué automatiquement.")
        champs = [("Date prévue", "Form_Date", "saisie_date", None), ("Banque", "Form_Banque", "saisie", "L_Banques"),
                  ("Libellé", "Form_Libelle", "saisie", None), ("Catégorie", "Form_Categorie", "saisie", "L_Categories"),
                  ("Type", "Form_Type", "saisie", "L_TypesForm"), ("Montant", "Form_Montant", "saisie_eur", None),
                  ("Commentaire", "Form_Commentaire", "saisie", None),
                  ("Banque destination (virement)", "Form_BanqueDest", "saisie", "L_Banques")]
        ws.merge_range(4, 1, 5, 1, "NOUVELLE\nOPÉRATION", self.fmt(bold=True, font_color="white", bg_color=BLEU_NUIT,
                                                                     align="center", valign="vcenter", text_wrap=True))
        for j, (lib, nom, style, liste) in enumerate(champs):
            col = 2 + j
            ws.write(4, col, lib, self.fmt(bold=True, font_color=BLEU_NUIT, font_size=9, text_wrap=True, valign="bottom"))
            if nom == "Form_Date":
                d = self.etat.params.date_ref
                ws.write_datetime(5, col, datetime(d.year, d.month, d.day), self.f[style])
            elif nom == "Form_Banque":
                ws.write(5, col, self.etat.params.banques[0].nom, self.f[style])
            elif nom == "Form_Type":
                ws.write(5, col, "Dépense", self.f[style])
            else:
                ws.write_blank(5, col, None, self.f[style])
            self.nom(nom, S.F_SAISIE, 5, col)
            if liste:
                ws.data_validation(5, col, 5, col, {"validate": "list", "source": "=" + liste})
        ws.set_row(4, 28)
        ws.set_row(5, 22)
        self.bouton(ws, "C8", "＋ Ajouter l'opération", "AjouterOperation", 170, 28, VERT)
        self.bouton(ws, "C8", "Effacer", "EffacerFormulaire", 80, 28, GRIS, dx=178)
        self.bouton(ws, "F8", "Valider les saisies du tableau", "ValiderSaisies", 200, 28, BLEU, dx=20)
        self.bouton(ws, "H8", "Trier par date", "TrierPrevisions", 110, 28, BLEU, dx=10)
        self.bouton(ws, "I8", "Supprimer la ligne sélectionnée", "SupprimerLigne", 200, 28, ROUGE, dx=40)
        self.bouton(ws, "K8", "Annuler le rapprochement", "AnnulerRapprochement", 180, 28, AMBRE, dx=30)
        self.bouton(ws, "M8", "◄ Tableau de bord", "AllerTableauDeBord", 160, 28, BLEU_NUIT, dx=20)
        self.bouton(ws, "O8", "Importer des prévisions…", "ImporterPrevisions", 190, 28, BLEU, dx=20)
        ws.set_row(7, 26)
        t = S.PREVISIONS
        r0, c0, n = self.ecrire_table(ws, t, self.lignes_previsions(), t.lignes_libres)
        self.mise_en_forme_statut(ws, t, "Suivi", {"À contrôler": "fond_orange", "Pointé": "fond_vert",
                                                   "Annulé": "gris_barre"})
        ws.conditional_format(self.plage_colonne(t, "Contrôle"), {"type": "cell", "criteria": "!=", "value": '""',
                                                                  "format": self.f["fond_rouge"]})
        ws.freeze_panes(t.ligne_entete, 3)

    def onglet_charges(self):
        ws = self.ws[S.F_CHARGES]
        self.entete_onglet(ws, "Charges fixes et recettes récurrentes",
                           "Les échéances sont générées automatiquement à chaque actualisation (horizon : voir Paramètres). "
                           "Modifier une charge ne modifie jamais les échéances passées, pointées ou modifiées à la main.")
        self.bouton(ws, "B5", "Générer les échéances", "GenererEcheances", 180, 28, VERT)
        self.bouton(ws, "D5", "◄ Tableau de bord", "AllerTableauDeBord", 160, 28, BLEU_NUIT, dx=10)
        ws.write(4, 6, "Charges fixes mensuelles nettes :", self.f["libelle"])
        ws.write_formula(4, 8, '=SUMIFS(tbl_Charges[Équivalent mensuel],tbl_Charges[Actif],"Oui")',
                         self.fmt(bold=True, num_format=S.FMT_EUR, font_size=12, font_color=BLEU_NUIT))
        ws.set_row(4, 28)
        t = S.CHARGES
        lignes = []
        for c in self.etat.charges:
            r = [None] * len(t.colonnes)
            v = {"ID": c.id, "Libellé": c.libelle, "Banque": c.banque, "Catégorie": c.categorie, "Type": c.type,
                 "Montant": c.montant, "Première échéance": c.premiere, "Fréquence": c.frequence, "Date de fin": c.fin,
                 "Actif": "Oui" if c.actif else "Non", "Commentaire": c.commentaire or None}
            for k, val in v.items():
                r[t.index(k)] = val
            lignes.append(r)
        self.ecrire_table(ws, t, lignes, t.lignes_libres)
        ws.freeze_panes(t.ligne_entete, 3)

    def onglet_banque(self):
        ws = self.ws[S.F_BANQUE]
        self.entete_onglet(ws, "Opérations bancaires importées",
                           "Journal des opérations réelles importées des relevés (jamais supprimées). "
                           "Seules les colonnes Catégorie et Commentaire sont modifiables.")
        self.bouton(ws, "B5", "Afficher les opérations à traiter", "FiltrerOperationsATraiter", 240, 28, AMBRE)
        self.bouton(ws, "D5", "Tout afficher", "AfficherToutesOperations", 110, 28, BLEU, dx=60)
        self.bouton(ws, "F5", "Annuler le rapprochement / traitement", "AnnulerRapprochement", 290, 28, ROUGE, dx=20)
        self.bouton(ws, "I5", "◄ Tableau de bord", "AllerTableauDeBord", 160, 28, BLEU_NUIT, dx=80)
        ws.set_row(4, 28)
        t = S.BANQUE
        lignes = []
        for o in self.etat.operations:
            r = [None] * len(t.colonnes)
            v = {"ID": o.id, "Banque": o.banque, "Date": o.date, "Date valeur": o.date_valeur, "Libellé": o.libelle,
                 "Montant": o.montant, "Référence": o.reference or None, "Catégorie": o.categorie or None,
                 "Commentaire": o.commentaire or None, "Statut": o.statut, "ID prévision": o.id_prevision or None,
                 "Mode": o.mode or None, "ID import": o.id_import or None, "Fichier": o.fichier or None,
                 "Date import": None, "Clé": o.cle or None}
            for k, val in v.items():
                r[t.index(k)] = val
            lignes.append(r)
        self.ecrire_table(ws, t, lignes)
        self.mise_en_forme_statut(ws, t, "Statut", {"À traiter": "fond_orange", "Proposé": "fond_bleu",
                                                    "Pointé": "fond_vert", "Non prévue": "fond_violet",
                                                    "Virement interne": "gris_barre"}, 20000)
        ws.freeze_panes(t.ligne_entete, 3)

    def onglet_rapprochement(self):
        ws = self.ws[S.F_RAPPRO]
        self.entete_onglet(ws, "Rapprochement : propositions à valider",
                           "Correspondances trouvées entre les prévisions et les opérations bancaires "
                           "(même banque, même montant, écart de date ≤ tolérance).")
        ws.merge_range(3, 1, 4, 10,
                       "Mode d'emploi : saisissez « Oui » dans la colonne Valider pour la bonne paire, « Non » pour rejeter "
                       "une proposition, puis cliquez « Valider les choix ». Une correspondance « Incertaine » regroupe "
                       "plusieurs candidats (même Groupe) : choisissez-en une seule. La ligne « Suggérée » est la plus probable.",
                       self.fmt(text_wrap=True, valign="top", font_color=GRIS, font_size=9, bg_color=GRIS_CLAIR))
        ws.set_row(3, 22)
        ws.set_row(4, 22)
        self.bouton(ws, "B6", "Rechercher les correspondances", "RechercherCorrespondances", 240, 28, BLEU)
        self.bouton(ws, "D6", "Valider les choix", "ValiderChoix", 140, 28, VERT, dx=80)
        self.bouton(ws, "G6", "Tout valider (certaines)", "ToutValiderCertaines", 180, 28, VERT, dx=0)
        self.bouton(ws, "I6", "◄ Tableau de bord", "AllerTableauDeBord", 160, 28, BLEU_NUIT, dx=60)
        ws.set_row(5, 28)
        t = S.PROPOSITIONS
        lignes = []
        for pr in self.etat.propositions:
            r = [None] * len(t.colonnes)
            v = {"Valider": pr.valider or None, "Niveau": pr.niveau, "Groupe": pr.groupe,
                 "Suggestion": "Suggérée" if pr.suggere else None, "Banque": pr.op.banque, "Date opération": pr.op.date,
                 "Libellé bancaire": pr.op.libelle, "Montant": pr.op.montant, "Date prévue": pr.prev.date,
                 "Libellé prévu": pr.prev.libelle, "Écart (jours)": pr.ecart, "Similarité": pr.sim, "Score": pr.score,
                 "ID opération": pr.op.id, "ID prévision": pr.prev.id}
            for k, val in v.items():
                r[t.index(k)] = val
            lignes.append(r)
        self.ecrire_table(ws, t, lignes)
        self.mise_en_forme_statut(ws, t, "Niveau", {"Certaine": "fond_vert", "Incertaine": "fond_orange"})
        self.mise_en_forme_statut(ws, t, "Suggestion", {"Suggérée": "fond_bleu"})
        ws.freeze_panes(t.ligne_entete, 2)

    def onglet_atraiter(self):
        ws = self.ws[S.F_ATRAITER]
        self.entete_onglet(ws, "Nouvelles opérations bancaires à traiter",
                           "Opérations bancaires sans prévision correspondante.")
        ws.merge_range(3, 1, 4, 10,
                       "Pour chaque opération : vérifiez la catégorie, ajoutez un commentaire si besoin, choisissez une Action "
                       "puis cliquez « Appliquer les traitements ». Pour rapprocher d'une prévision de montant différent, "
                       "choisissez « Rapprocher avec la prévision indiquée » et saisissez son ID (colonne ID de l'onglet Saisie).",
                       self.fmt(text_wrap=True, valign="top", font_color=GRIS, font_size=9, bg_color=GRIS_CLAIR))
        ws.set_row(3, 22)
        ws.set_row(4, 22)
        self.bouton(ws, "B6", "Appliquer les traitements", "AppliquerTraitements", 200, 28, VERT)
        self.bouton(ws, "D6", "Actualiser la liste", "ActualiserListe", 150, 28, BLEU, dx=30)
        self.bouton(ws, "G6", "◄ Tableau de bord", "AllerTableauDeBord", 160, 28, BLEU_NUIT, dx=0)
        ws.set_row(5, 28)
        t = S.ATRAITER
        lignes = []
        if not self.etat.regles:
            self.etat.regles = [m.Regle(k, c, a) for k, c, a in S.REGLES_DEFAUT]
        for o in m.a_traiter(self.etat):
            r = [None] * len(t.colonnes)
            proche = m.prevision_proche(self.etat, o)
            sugg = m.suggestion_virement(self.etat, o)
            categorie = o.categorie
            if not categorie:
                regle = m.trouver_regle(self.etat, o.libelle)
                categorie = regle.categorie if regle else ""
            if proche:
                texte = f"Prévision proche (montant différent) : {proche}"
            elif sugg:
                texte = f"Virement interne probable avec {sugg} ({self.etat.op(sugg).banque})"
            else:
                texte = None
            v = {"ID opération": o.id, "Date": o.date, "Banque": o.banque, "Libellé": o.libelle, "Montant": o.montant,
                 "Suggestion": texte, "Catégorie": categorie or None, "Commentaire": o.commentaire or None,
                 "ID prévision": proche or None}
            for k, val in v.items():
                r[t.index(k)] = val
            lignes.append(r)
        self.ecrire_table(ws, t, lignes)
        ws.freeze_panes(t.ligne_entete, 2)

    def onglets_transit(self):
        for t in (S.STAGE_B1, S.STAGE_B2):
            self.ecrire_table(self.ws[t.onglet], t, [])

    def onglet_historique(self):
        t = S.HISTORIQUE
        lignes = []
        for h in self.etat.historique:
            lignes.append([h["id"], h["date"], h["banque"], h["fichier"], h["lues"], h["ajoutees"], h["doublons"],
                           h["anterieures"], h["erreur"] or None, None])
        self.ecrire_table(self.ws[S.F_HISTO], t, lignes)

    # -------------------------------------------------------------------------
    CALCULS = [
        ("Date_Reference", '=IF(Date_Ref_Saisie="",TODAY(),Date_Ref_Saisie)', S.FMT_DATE),
        ("SoldeReel_B1", '=B1_SoldeInit+SUMIFS(tbl_Banque[Montant],tbl_Banque[Banque],B1_Nom,tbl_Banque[Date],">"&B1_DateSolde)', S.FMT_EUR),
        ("SoldeReel_B2", '=B2_SoldeInit+SUMIFS(tbl_Banque[Montant],tbl_Banque[Banque],B2_Nom,tbl_Banque[Date],">"&B2_DateSolde)', S.FMT_EUR),
        ("SoldeReel_Total", "=SoldeReel_B1+SoldeReel_B2", S.FMT_EUR),
        ("DernierReleve_B1", '=IF(COUNTIFS(tbl_Banque[Banque],B1_Nom)=0,"",MAXIFS(tbl_Banque[Date],tbl_Banque[Banque],B1_Nom))', S.FMT_DATE),
        ("DernierReleve_B2", '=IF(COUNTIFS(tbl_Banque[Banque],B2_Nom)=0,"",MAXIFS(tbl_Banque[Date],tbl_Banque[Banque],B2_Nom))', S.FMT_DATE),
        ("DernierImport", '=IF(COUNT(tbl_Historique[Date et heure])=0,"",MAX(tbl_Historique[Date et heure]))', S.FMT_DATEHEURE),
        ("DateHorizon", "=Date_Reference+Horizon_Jours", S.FMT_DATE),
        ("SoldePrev_B1", "=INDEX(Calc_SoldeB1,Horizon_Jours+1)", S.FMT_EUR),
        ("SoldePrev_B2", "=INDEX(Calc_SoldeB2,Horizon_Jours+1)", S.FMT_EUR),
        ("SoldePrev_Total", "=INDEX(Calc_SoldeGlobal,Horizon_Jours+1)", S.FMT_EUR),
        ("Prev_B1_30", "=INDEX(Calc_SoldeB1,31)", S.FMT_EUR),
        ("Prev_B1_60", "=INDEX(Calc_SoldeB1,61)", S.FMT_EUR),
        ("Prev_B1_90", "=INDEX(Calc_SoldeB1,91)", S.FMT_EUR),
        ("Prev_B2_30", "=INDEX(Calc_SoldeB2,31)", S.FMT_EUR),
        ("Prev_B2_60", "=INDEX(Calc_SoldeB2,61)", S.FMT_EUR),
        ("Prev_B2_90", "=INDEX(Calc_SoldeB2,91)", S.FMT_EUR),
        ("Prev_Total_30", "=INDEX(Calc_SoldeGlobal,31)", S.FMT_EUR),
        ("Prev_Total_60", "=INDEX(Calc_SoldeGlobal,61)", S.FMT_EUR),
        ("Prev_Total_90", "=INDEX(Calc_SoldeGlobal,91)", S.FMT_EUR),
        ("Depenses_A_Venir", '=-SUMIFS(tbl_Previsions[Montant à venir],tbl_Previsions[Type],"Dépense",tbl_Previsions[Date prévue],"<="&DateHorizon)', S.FMT_EUR),
        ("Recettes_Attendues", '=SUMIFS(tbl_Previsions[Montant à venir],tbl_Previsions[Type],"Recette",tbl_Previsions[Date prévue],"<="&DateHorizon)', S.FMT_EUR),
        ("PointBas", '=MINIFS(Calc_SoldeGlobal,Calc_Jours,"<="&Horizon_Jours)', S.FMT_EUR),
        ("PointBas_Date", "=INDEX(Calc_Dates,MATCH(PointBas,Calc_SoldeGlobal,0))", S.FMT_DATE),
        ("NbPropositions", '=COUNTIFS(tbl_Propositions[ID opération],"?*")', "0"),
        ("NbOpsProposees", '=COUNTIFS(tbl_Banque[Statut],"Proposé",tbl_Banque[Inclus],1)', "0"),
        ("NbATraiter", '=COUNTIFS(tbl_Banque[Statut],"À traiter",tbl_Banque[Inclus],1)', "0"),
        ("NbEchues", '=COUNTIFS(tbl_Previsions[Suivi],"À contrôler")', "0"),
        ("NonPrevues_Montant", '=-(SUMIFS(tbl_Banque[Montant],tbl_Banque[Statut],"Non prévue",tbl_Banque[Montant],"<0",tbl_Banque[Inclus],1)'
                               '+SUMIFS(tbl_Banque[Montant],tbl_Banque[Statut],"À traiter",tbl_Banque[Montant],"<0",tbl_Banque[Inclus],1))', S.FMT_EUR),
        ("NonPrevues_Nb", '=COUNTIFS(tbl_Banque[Statut],"Non prévue",tbl_Banque[Montant],"<0",tbl_Banque[Inclus],1)'
                          '+COUNTIFS(tbl_Banque[Statut],"À traiter",tbl_Banque[Montant],"<0",tbl_Banque[Inclus],1)', "0"),
        ("Suivi_Ouverture", '=IF(Suivi_Banque="Toutes",B1_SoldeInit+B2_SoldeInit,IF(Suivi_Banque=B1_Nom,B1_SoldeInit,'
                            'IF(Suivi_Banque=B2_Nom,B2_SoldeInit,0)))', S.FMT_EUR),
        ("Suivi_NbLignes", "=COUNT(tbl_Banque[Rang vue])+COUNT(tbl_Previsions[Rang vue])", "0"),
        ("Alerte_B1_Date", '=IF(MINIFS(Calc_Dates,Calc_SoldeB1,"<0")=0,"",MINIFS(Calc_Dates,Calc_SoldeB1,"<0"))', S.FMT_DATE),
        ("Alerte_B2_Date", '=IF(MINIFS(Calc_Dates,Calc_SoldeB2,"<0")=0,"",MINIFS(Calc_Dates,Calc_SoldeB2,"<0"))', S.FMT_DATE),
        ("Alerte_Global_Date", '=IF(MINIFS(Calc_Dates,Calc_SoldeGlobal,"<"&Seuil_Alerte)=0,"",'
                               'MINIFS(Calc_Dates,Calc_SoldeGlobal,"<"&Seuil_Alerte))', S.FMT_DATE),
        ("Alerte_Date", '=IF(COUNT(Alerte_B1_Date,Alerte_B2_Date,Alerte_Global_Date)=0,"",'
                        'MIN(Alerte_B1_Date,Alerte_B2_Date,Alerte_Global_Date))', S.FMT_DATE),
        ("Alerte_Texte", '=IF(Alerte_Date="","","⚠ ALERTE TRÉSORERIE (90 jours) : "&IF(Alerte_B1_Date<>"",B1_Nom&'
                         '" passe en négatif ; ","")&IF(Alerte_B2_Date<>"",B2_Nom&" passe en négatif ; ","")&'
                         'IF(Alerte_Global_Date<>"","trésorerie globale sous le seuil d\'alerte ; ","")&"premier jour concerné :")', None),
    ]

    def onglet_calculs(self):
        ws = self.ws[S.F_CALC]
        ws.set_column(0, 0, 22)
        ws.set_column(1, 1, 16)
        ws.set_column(3, 12, 13)
        for i, (nom, formule, fmt) in enumerate(self.CALCULS):
            ws.write(i, 0, nom)
            ws.write_formula(i, 1, formule, self.fmt(**({"num_format": fmt} if fmt else {})))
            self.nom(nom, S.F_CALC, i, 1)
        entetes = ["Jour", "Date", "Cumul B1", "Solde B1", "Cumul B2", "Solde B2", "Solde global", "Flux B1",
                   "Flux B2", "Flux global"]
        for j, h in enumerate(entetes):
            ws.write(0, 3 + j, h, self.f["gras"])
        feur = self.fmt(num_format=S.FMT_EUR)
        fdate = self.fmt(num_format=S.FMT_DATE)
        for k in range(91):
            r = 1 + k
            ligne = r + 1
            ws.write_number(r, 3, k)
            ws.write_formula(r, 4, f"=Date_Reference+D{ligne}", fdate)
            ws.write_formula(r, 5, f'=SUMIFS(tbl_Previsions[Montant à venir],tbl_Previsions[Banque],B1_Nom,'
                                   f'tbl_Previsions[Date prévue],"<="&E{ligne})', feur)
            ws.write_formula(r, 6, f"=SoldeReel_B1+F{ligne}", feur)
            ws.write_formula(r, 7, f'=SUMIFS(tbl_Previsions[Montant à venir],tbl_Previsions[Banque],B2_Nom,'
                                   f'tbl_Previsions[Date prévue],"<="&E{ligne})', feur)
            ws.write_formula(r, 8, f"=SoldeReel_B2+H{ligne}", feur)
            ws.write_formula(r, 9, f"=G{ligne}+I{ligne}", feur)
            if k == 0:
                ws.write_formula(r, 10, f"=F{ligne}", feur)
                ws.write_formula(r, 11, f"=H{ligne}", feur)
            else:
                ws.write_formula(r, 10, f"=F{ligne}-F{ligne - 1}", feur)
                ws.write_formula(r, 11, f"=H{ligne}-H{ligne - 1}", feur)
            ws.write_formula(r, 12, f"=K{ligne}+L{ligne}", feur)
        for nom, col in (("Calc_Jours", 3), ("Calc_Dates", 4), ("Calc_SoldeB1", 6), ("Calc_SoldeB2", 8),
                         ("Calc_SoldeGlobal", 9), ("Calc_FluxB1", 10), ("Calc_FluxB2", 11), ("Calc_FluxGlobal", 12)):
            self.nom(nom, S.F_CALC, 1, col, 91, col)

    # -------------------------------------------------------------------------
    NB_LIGNES_SUIVI = 1000

    def onglet_suivi(self):
        ws = self.ws[S.F_SUIVI]
        self.entete_onglet(ws, "Suivi prévisionnel chronologique",
                           "Opérations réelles importées puis prévisions restant à réaliser, avec le solde après chaque "
                           "opération. Les prévisions échues non pointées sont placées à la date du jour.")
        filtres = [("Banque", "Suivi_Banque", "Toutes", "L_BanquesSuivi", "saisie"),
                   ("Du", "Suivi_Du", "=Date_Reference-30", None, "saisie_date"),
                   ("Au", "Suivi_Au", "=Date_Reference+90", None, "saisie_date"),
                   ("Afficher", "Suivi_Afficher", "Tout", "L_Afficher", "saisie")]
        for j, (lib, nom, defaut, liste, style) in enumerate(filtres):
            col = 2 + j
            ws.write(3, col, lib, self.fmt(bold=True, font_color=BLEU_NUIT, font_size=9))
            if defaut.startswith("="):
                ws.write_formula(4, col, defaut, self.f[style])
            else:
                ws.write(4, col, defaut, self.f[style])
            self.nom(nom, S.F_SUIVI, 4, col)
            if liste:
                ws.data_validation(4, col, 4, col, {"validate": "list", "source": "=" + liste})
            else:
                ws.data_validation(4, col, 4, col, {"validate": "date", "criteria": "between",
                                                    "minimum": date(2000, 1, 1), "maximum": date(2100, 12, 31)})
        ws.write(4, 1, "Filtres :", self.f["libelle"])
        self.bouton(ws, "G4", "Réinitialiser les filtres", "ReinitialiserFiltresSuivi", 170, 26, BLEU, dy=4)
        self.bouton(ws, "I4", "◄ Tableau de bord", "AllerTableauDeBord", 160, 26, BLEU_NUIT, dx=20, dy=4)
        ws.write_formula(5, 2, '=IF(Suivi_NbLignes>' + str(self.NB_LIGNES_SUIVI) + ',"Attention : seules les '
                         + str(self.NB_LIGNES_SUIVI) + ' premières lignes sont affichées (réduisez la période).","")',
                         self.f["rouge"])
        larg = {1: 6, 2: 12, 3: 16, 4: 40, 5: 18, 6: 14, 7: 18, 8: 21, 9: 10, 10: 11}
        for c, w in larg.items():
            ws.set_column(c, c, w)
        ws.set_column(12, 15, 10, None, {"hidden": True})
        entetes = ["N°", "Date", "Banque", "Libellé", "Catégorie", "Montant", "Statut", "Solde après opération",
                   "Source", "ID"]
        r0 = 7
        for j, h in enumerate(entetes):
            ws.write(r0, 1 + j, h, self.f["tab_entete"])
        ws.set_row(r0, 22)
        f_txt = self.fmt(border=1, border_color=BORDURE, font_size=9)
        f_eur = self.fmt(border=1, border_color=BORDURE, font_size=9, num_format=S.FMT_EUR)
        f_solde = self.fmt(border=1, border_color=BORDURE, font_size=9, num_format=S.FMT_EUR, bold=True)
        f_dat = self.fmt(border=1, border_color=BORDURE, font_size=9, num_format=S.FMT_DATE)
        for k in range(1, self.NB_LIGNES_SUIVI + 1):
            r = r0 + k
            L = r + 1  # numéro de ligne Excel
            ws.write_number(r, 12, k)
            ws.write_formula(r, 13, f'=IFERROR(MATCH(M{L},tbl_Banque[Rang vue],0),"")')
            ws.write_formula(r, 14, f'=IF(N{L}<>"","",IFERROR(MATCH(M{L},tbl_Previsions[Rang vue],0),""))')
            ws.write_formula(r, 15, f'=IF(N{L}<>"",INDEX(tbl_Banque[Clé tri],N{L}),IF(O{L}<>"",INDEX(tbl_Previsions[Clé tri],O{L}),""))')

            def choix(col_b, col_p, texte=False):
                suffixe = '&""' if texte else ""
                return (f'=IF(N{L}<>"",INDEX(tbl_Banque[{col_b}],N{L}){suffixe},'
                        f'IF(O{L}<>"",INDEX(tbl_Previsions[{col_p}],O{L}){suffixe},""))')
            ws.write_formula(r, 1, f'=IF(P{L}="","",M{L})', f_txt)
            ws.write_formula(r, 2, choix("Date", "Date prévue"), f_dat)
            ws.write_formula(r, 3, choix("Banque", "Banque", True), f_txt)
            ws.write_formula(r, 4, choix("Libellé", "Libellé", True), f_txt)
            ws.write_formula(r, 5, choix("Catégorie", "Catégorie", True), f_txt)
            ws.write_formula(r, 6, choix("Montant", "Montant à venir"), f_eur)
            ws.write_formula(r, 7, f'=IF(N{L}<>"",IF(INDEX(tbl_Banque[Statut],N{L})="Pointé","Pointé",'
                                   f'IF(INDEX(tbl_Banque[Statut],N{L})="Proposé","Pointé (à valider)",'
                                   f'IF(INDEX(tbl_Banque[Statut],N{L})="À traiter","Réel à traiter",'
                                   f'IF(INDEX(tbl_Banque[Statut],N{L})="Virement interne","Virement interne","Réel non prévu")))),'
                                   f'IF(O{L}<>"",INDEX(tbl_Previsions[Suivi],O{L}),""))', f_txt)
            ws.write_formula(r, 8, f'=IF(P{L}="","",Suivi_Ouverture+SUMIFS(tbl_Banque[Montant],tbl_Banque[Inclus banque],1,'
                                   f'tbl_Banque[Clé tri],"<="&P{L})+SUMIFS(tbl_Previsions[Montant à venir],'
                                   f'tbl_Previsions[Inclus banque],1,tbl_Previsions[Clé tri],"<="&P{L}))', f_solde)
            ws.write_formula(r, 9, f'=IF(N{L}<>"","Banque",IF(O{L}<>"","Prévision",""))', f_txt)
            ws.write_formula(r, 10, choix("ID", "ID", True), f_txt)
        plage = f"B{r0 + 2}:K{r0 + 1 + self.NB_LIGNES_SUIVI}"
        ws.conditional_format(f"I{r0 + 2}:I{r0 + 1 + self.NB_LIGNES_SUIVI}",
                              {"type": "cell", "criteria": "<", "value": 0, "format": self.f["fond_rouge"]})
        ws.conditional_format(plage, {"type": "formula", "criteria": f'=$H{r0 + 2}="À contrôler"', "format": self.f["fond_orange"]})
        ws.conditional_format(plage, {"type": "formula", "criteria": f'=$J{r0 + 2}="Prévision"',
                                      "format": self.wb.add_format({"italic": True, "font_color": "#1E40AF"})})
        ws.freeze_panes(r0 + 1, 0)

    # -------------------------------------------------------------------------
    NB_CATEGORIES_SYNTHESE = 50
    MOIS_PASSES = 5
    MOIS_FUTURS = 6

    def onglet_synthese(self):
        ws = self.ws[S.F_SYNTHESE]
        self.entete_onglet(ws, "Synthèse mensuelle par catégorie",
                           "Mois passés : opérations bancaires réelles. Mois en cours et à venir : réel + prévisions "
                           "restant à réaliser (les prévisions échues sont comptées dans le mois en cours). "
                           "Les virements internes s'annulent sur « Toutes ».")
        ws.write(4, 1, "Banque :", self.f["libelle"])
        ws.write(4, 2, "Toutes", self.f["saisie"])
        ws.data_validation(4, 2, 4, 2, {"validate": "list", "source": "=L_BanquesSuivi"})
        self.nom("Synth_Banque", S.F_SYNTHESE, 4, 2)
        self.bouton(ws, "E5", "◄ Tableau de bord", "AllerTableauDeBord", 160, 26, BLEU_NUIT, dy=-2)
        nb_mois = self.MOIS_PASSES + 1 + self.MOIS_FUTURS
        ws.set_column(1, 1, 26)
        ws.set_column(2, 1 + nb_mois, 12.5)
        ws.set_column(2 + nb_mois, 2 + nb_mois, 14)
        r_ent = 6
        f_ent = self.fmt(bold=True, font_color="white", bg_color=BLEU_NUIT, align="center", valign="vcenter",
                         num_format="mmm yyyy", border=1, border_color=BORDURE)
        f_nat = self.fmt(italic=True, font_color=GRIS, align="center", font_size=8)
        ws.write(r_ent, 1, "Catégorie", f_ent)
        ws.write(r_ent - 1, 1, "", f_nat)
        crit_b = 'IF(Synth_Banque="Toutes","*",Synth_Banque)'
        for k in range(nb_mois):
            col = 2 + k
            decalage = k - self.MOIS_PASSES
            ws.write_formula(r_ent, col, f"=EDATE(DATE(YEAR(Date_Reference),MONTH(Date_Reference),1),{decalage})", f_ent)
            ws.write(r_ent - 1, col, "Réel" if decalage < 0 else ("Réel + prévu" if decalage == 0 else "Prévu"), f_nat)
        ws.write(r_ent, 2 + nb_mois, "Total", f_ent)
        fmt_sans_zero = '#,##0\\ "€";[Red]\\-#,##0\\ "€";""'
        f_cat = self.fmt(font_size=9)
        f_val = self.fmt(font_size=9, num_format=fmt_sans_zero)
        f_tot = self.fmt(border=1, border_color=BORDURE, font_size=9, num_format=S.FMT_EUR0, bold=True, bg_color=GRIS_CLAIR)
        f_lib_tot = self.fmt(border=1, border_color=BORDURE, font_size=9, bold=True, bg_color=GRIS_CLAIR)
        f_tot_cat = self.fmt(font_size=9, num_format=fmt_sans_zero, bold=True, bg_color=GRIS_CLAIR)
        lettre = xl_col_to_name

        def flux(col, crit_cat):
            m = f"{lettre(col)}${r_ent + 1}"
            fin = f"EDATE({m},1)"
            return (f'SUMIFS(tbl_Banque[Montant],tbl_Banque[Banque],{crit_b},tbl_Banque[Inclus],1{crit_cat("tbl_Banque")},'
                    f'tbl_Banque[Date],">="&{m},tbl_Banque[Date],"<"&{fin})'
                    f'+SUMIFS(tbl_Previsions[Montant à venir],tbl_Previsions[Banque],{crit_b}{crit_cat("tbl_Previsions")},'
                    f'tbl_Previsions[Date effective],">="&{m},tbl_Previsions[Date effective],"<"&{fin})')

        # Lignes de synthèse en haut, catégories ensuite
        r_solde, r_net, r_autres = r_ent + 1, r_ent + 2, r_ent + 3
        premiere = r_ent + 4
        derniere = premiere + self.NB_CATEGORIES_SYNTHESE - 1
        ws.write(r_solde, 1, "Solde de fin de mois", f_lib_tot)
        ws.write(r_net, 1, "Flux net du mois", f_lib_tot)
        ws.write(r_autres, 1, "Sans catégorie / autres", f_lib_tot)
        ouverture = ('IF(Synth_Banque="Toutes",B1_SoldeInit+B2_SoldeInit,IF(Synth_Banque=B1_Nom,B1_SoldeInit,'
                     'IF(Synth_Banque=B2_Nom,B2_SoldeInit,0)))')
        for k in range(nb_mois):
            col = 2 + k
            c = lettre(col)
            m = f"{c}${r_ent + 1}"
            ws.write_formula(r_solde, col,
                             f'={ouverture}+SUMIFS(tbl_Banque[Montant],tbl_Banque[Banque],{crit_b},tbl_Banque[Inclus],1,'
                             f'tbl_Banque[Date],"<"&EDATE({m},1))+SUMIFS(tbl_Previsions[Montant à venir],'
                             f'tbl_Previsions[Banque],{crit_b},tbl_Previsions[Date effective],"<"&EDATE({m},1))', f_tot)
            ws.write_formula(r_net, col, "=" + flux(col, lambda t: ""), f_tot)
            ws.write_formula(r_autres, col, f"={c}{r_net + 1}-SUM({c}{premiere + 1}:{c}{derniere + 1})", f_tot)
        ws.write_blank(r_solde, 2 + nb_mois, None, f_tot)
        for r in (r_autres, r_net):
            ws.write_formula(r, 2 + nb_mois, f"=SUM({lettre(2)}{r + 1}:{lettre(1 + nb_mois)}{r + 1})", f_tot)
        for i in range(self.NB_CATEGORIES_SYNTHESE):
            r = premiere + i
            ws.write_formula(r, 1, f'=IFERROR(INDEX(tbl_Categories[Catégorie],{i + 1})&"","")', f_cat)
            for k in range(nb_mois):
                col = 2 + k
                ws.write_formula(r, col, f'=IF($B{r + 1}="","",' + flux(col, lambda t, rr=r: f",{t}[Catégorie],$B{rr + 1}") + ")",
                                 f_val)
            ws.write_formula(r, 2 + nb_mois, f'=IF($B{r + 1}="","",SUM({lettre(2)}{r + 1}:{lettre(1 + nb_mois)}{r + 1}))', f_tot_cat)
        # bordures uniquement sur les catégories renseignées
        ws.conditional_format(premiere, 1, derniere, 2 + nb_mois,
                              {"type": "formula", "criteria": f'=$B{premiere + 1}<>""',
                               "format": self.wb.add_format({"border": 1, "border_color": BORDURE})})
        plage = f"{lettre(2)}{r_solde + 1}:{lettre(2 + nb_mois)}{derniere + 1}"
        ws.conditional_format(plage, {"type": "cell", "criteria": "<", "value": 0,
                                      "format": self.wb.add_format({"font_color": ROUGE})})
        col_cour = lettre(2 + self.MOIS_PASSES)
        ws.conditional_format(f"{col_cour}{premiere + 1}:{col_cour}{derniere + 1}",
                              {"type": "formula", "criteria": f'=$B{premiere + 1}<>""',
                               "format": self.wb.add_format({"bg_color": "#FFFBEB"})})
        ws.freeze_panes(r_ent + 1, 2)

    # -------------------------------------------------------------------------
    def carte(self, ws, ligne: int, col: int, titre, formule_valeur: str, formule_detail: str | None,
              fmt_valeur="carte_valeur", fmt_detail="carte_detail", largeur=4):
        if titre.startswith("="):
            ws.merge_range(ligne, col, ligne, col + largeur - 1, "", self.f["carte_titre"])
            ws.write_formula(ligne, col, titre, self.f["carte_titre"])
        else:
            ws.merge_range(ligne, col, ligne, col + largeur - 1, titre, self.f["carte_titre"])
        ws.merge_range(ligne + 1, col, ligne + 1, col + largeur - 1, "", self.f[fmt_valeur])
        ws.write_formula(ligne + 1, col, formule_valeur, self.f[fmt_valeur])
        ws.merge_range(ligne + 2, col, ligne + 2, col + largeur - 1, "", self.f[fmt_detail])
        if formule_detail:
            ws.write_formula(ligne + 2, col, formule_detail, self.f[fmt_detail])
        cellule = xl_rowcol_to_cell(ligne + 1, col)
        ws.conditional_format(cellule, {"type": "cell", "criteria": "<", "value": 0,
                                        "format": self.wb.add_format({"font_color": ROUGE})})

    def onglet_tableau_de_bord(self):
        ws = self.ws[S.F_TDB]
        ws.hide_gridlines(2)
        ws.set_column(0, 0, 2)
        ws.set_column(1, 15, 10.5)
        ws.set_row(0, 8)
        ws.merge_range(1, 1, 1, 11, "", self.f["titre"])
        ws.write_formula(1, 1, '="Trésorerie - "&Societe', self.f["titre"])
        ws.set_row(1, 32)
        ws.write(2, 1, "Situation au", self.f["sous_titre"])
        ws.write_formula(2, 2, "=Date_Reference", self.fmt(font_color=GRIS, italic=True, num_format="dd/mm/yyyy", align="left"))
        ws.merge_range(2, 4, 2, 5, "Dernière actualisation :", self.f["sous_titre"])
        ws.merge_range(2, 6, 2, 8, "", self.fmt(font_color=GRIS, italic=True, num_format="dd/mm/yyyy hh:mm", align="left"))
        ws.write_formula(2, 6, '=IF(DernierImport="","jamais",DernierImport)',
                         self.fmt(font_color=GRIS, italic=True, num_format="dd/mm/yyyy hh:mm", align="left"))
        # Boutons
        self.bouton(ws, "B5", "⟳ Actualiser les banques", "Actualiser", 190, 34, VERT)
        self.bouton(ws, "D5", "Rapprochement", "AllerRapprochement", 120, 34, BLEU, dx=55)
        self.bouton(ws, "F5", "À traiter", "AllerATraiter", 95, 34, AMBRE, dx=40)
        self.bouton(ws, "H5", "Saisir une opération", "AllerSaisie", 150, 34, BLEU, dx=0)
        self.bouton(ws, "J5", "Charges fixes", "AllerCharges", 110, 34, BLEU, dx=20)
        self.bouton(ws, "L5", "Suivi prévisionnel", "AllerSuivi", 140, 34, BLEU, dx=0)
        self.bouton(ws, "N5", "Synthèse", "AllerSynthese", 85, 34, BLEU, dx=20)
        self.bouton(ws, "P5", "Aide", "AllerAide", 55, 34, GRIS, dx=0)
        ws.set_row(4, 30)
        # Bandeau d'alerte (vide en l'absence d'alerte)
        f_alerte = self.fmt(bold=True, font_color=ROUGE, valign="vcenter", indent=1)
        ws.merge_range(5, 1, 5, 11, "", f_alerte)
        ws.write_formula(5, 1, "=Alerte_Texte", f_alerte)
        f_alerte_d = self.fmt(bold=True, font_color=ROUGE, valign="vcenter", num_format="dd/mm/yyyy", align="left")
        ws.merge_range(5, 12, 5, 14, "", f_alerte_d)
        ws.write_formula(5, 12, "=Alerte_Date", f_alerte_d)
        ws.conditional_format(5, 1, 5, 14, {"type": "formula", "criteria": '=Alerte_Texte<>""',
                                            "format": self.wb.add_format({"bg_color": "#FEE2E2"})})
        ws.set_row(5, 24)

        ws.merge_range(7, 1, 7, 15, "SOLDES RÉELS (opérations bancaires importées)", self.f["section"])
        ws.set_row(7, 20)
        self.carte(ws, 8, 1, "=B1_Nom", "=SoldeReel_B1", '=IF(DernierReleve_B1="","aucune opération importée",DernierReleve_B1)',
                   fmt_detail="carte_detail_date", largeur=5)
        self.carte(ws, 8, 6, "=B2_Nom", "=SoldeReel_B2", '=IF(DernierReleve_B2="","aucune opération importée",DernierReleve_B2)',
                   fmt_detail="carte_detail_date", largeur=5)
        self.carte(ws, 8, 11, "TRÉSORERIE GLOBALE", "=SoldeReel_Total", '="2 comptes consolidés"', largeur=5)

        ws.merge_range(12, 1, 12, 9, "", self.f["section"])
        ws.write_formula(12, 1, '="SOLDES PRÉVISIONNELS À J+"&Horizon_Jours', self.f["section"])
        ws.merge_range(12, 10, 12, 11, "Horizon :", self.fmt(bold=True, font_color="white", bg_color=BLEU_NUIT,
                                                             align="right", valign="vcenter"))
        ws.write_blank(12, 12, None, self.f["section"])
        ws.write_blank(12, 13, None, self.f["section"])
        ws.write_blank(12, 14, None, self.f["section"])
        ws.write_number(12, 15, 30, self.fmt(locked=False, bold=True, align="center", valign="vcenter", font_color=BLEU_NUIT,
                                             bg_color=JAUNE_SAISIE, num_format='0" j"'))
        ws.data_validation(12, 15, 12, 15, {"validate": "list", "source": "=L_Horizon"})
        self.nom("Horizon_Jours", S.F_TDB, 12, 15)
        ws.set_row(12, 22)
        self.carte(ws, 13, 1, "=B1_Nom", "=SoldePrev_B1", "=DateHorizon", fmt_detail="carte_detail_date", largeur=5)
        self.carte(ws, 13, 6, "=B2_Nom", "=SoldePrev_B2", "=DateHorizon", fmt_detail="carte_detail_date", largeur=5)
        self.carte(ws, 13, 11, "TRÉSORERIE GLOBALE PRÉVUE", "=SoldePrev_Total", "=DateHorizon",
                   fmt_detail="carte_detail_date", largeur=5)
        self.bouton(ws, "M13", "30 j", "Horizon30", 60, 20, BLEU, dx=8, dy=1)
        self.bouton(ws, "N13", "60 j", "Horizon60", 60, 20, BLEU, dx=8, dy=1)
        self.bouton(ws, "O13", "90 j", "Horizon90", 60, 20, BLEU, dx=8, dy=1)

        ws.merge_range(17, 1, 17, 15, "À SURVEILLER", self.f["section"])
        ws.set_row(17, 20)
        self.carte(ws, 18, 1, '="Dépenses à venir (J+"&Horizon_Jours&")"', "=Depenses_A_Venir",
                   '="hors virements internes"', largeur=5)
        self.carte(ws, 18, 6, '="Recettes attendues (J+"&Horizon_Jours&")"', "=Recettes_Attendues",
                   '="hors virements internes"', largeur=5)
        self.carte(ws, 18, 11, '="Point bas global (J+"&Horizon_Jours&")"', "=PointBas", "=PointBas_Date",
                   fmt_detail="carte_detail_date", largeur=5)
        self.carte(ws, 22, 1, "En attente de rapprochement", "=NbOpsProposees",
                   '=NbPropositions&" proposition(s) à valider"', fmt_valeur="carte_valeur_n", largeur=4)
        self.carte(ws, 22, 5, "Nouvelles opérations à traiter", "=NbATraiter", '="opérations bancaires non prévues"',
                   fmt_valeur="carte_valeur_n", largeur=4)
        self.carte(ws, 22, 9, "Dépenses bancaires non prévues", "=NonPrevues_Montant",
                   '=NonPrevues_Nb&" opération(s) depuis le solde initial"', largeur=4)
        self.carte(ws, 22, 13, "Prévisions à contrôler", "=NbEchues", '="échues ou à valider"',
                   fmt_valeur="carte_valeur_n", largeur=3)
        for r in (9, 14, 19, 23):
            ws.set_row(r, 34)

        # Graphique
        ws.merge_range(26, 1, 26, 15, "", self.f["section"])
        ws.write_formula(26, 1, '="ÉVOLUTION PRÉVISIONNELLE DE LA TRÉSORERIE (J+0 à J+"&Horizon_Jours&")"', self.f["section"])
        ws.set_row(26, 20)
        h = 30
        graph = self.wb.add_chart({"type": "line"})
        for nom, col, couleur, epaisseur in (("Trésorerie globale", "J", BLEU_NUIT, 2.75), ("=B1", "G", BLEU, 1.75),
                                             ("=B2", "I", AMBRE, 1.75)):
            nom_serie = nom
            if nom == "=B1":
                nom_serie = f"='{S.F_LISTES}'!$A$2"
            elif nom == "=B2":
                nom_serie = f"='{S.F_LISTES}'!$A$3"
            graph.add_series({
                "name": nom_serie,
                "categories": f"='{S.F_CALC}'!$E$2:$E${2 + h}",
                "values": f"='{S.F_CALC}'!${col}$2:${col}${2 + h}",
                "line": {"color": couleur, "width": epaisseur},
                "marker": {"type": "none"},
                "smooth": False,
            })
        graph.set_x_axis({"date_axis": True, "num_format": "dd/mm", "num_font": {"size": 8, "color": GRIS},
                          "line": {"color": BORDURE}, "major_gridlines": {"visible": False}})
        graph.set_y_axis({"num_format": '#,##0 "€"', "num_font": {"size": 8, "color": GRIS},
                          "major_gridlines": {"visible": True, "line": {"color": "#E2E8F0"}}, "line": {"none": True}})
        graph.set_legend({"position": "top", "font": {"size": 9}})
        graph.set_chartarea({"border": {"none": True}})
        graph.set_plotarea({"border": {"none": True}})
        graph.show_hidden_data()
        graph.set_size({"width": 1180, "height": 330})
        ws.insert_chart("B28", graph, {"x_offset": 4, "y_offset": 4, "object_position": 3})

        # Projection 30 / 60 / 90
        r = 46
        ws.merge_range(r, 1, r, 15, "PROJECTION À 30 / 60 / 90 JOURS", self.f["section"])
        ws.set_row(r, 20)
        ws.merge_range(r + 1, 1, r + 1, 3, "Compte", self.f["tab_entete"])
        for j, lib in enumerate(["Solde réel", "J+30", "J+60", "J+90"]):
            ws.merge_range(r + 1, 4 + 2 * j, r + 1, 5 + 2 * j, lib, self.f["tab_entete"])
        lignes = [("=B1_Nom", "SoldeReel_B1", "Prev_B1_30", "Prev_B1_60", "Prev_B1_90", "tab_eur"),
                  ("=B2_Nom", "SoldeReel_B2", "Prev_B2_30", "Prev_B2_60", "Prev_B2_90", "tab_eur"),
                  ('="Trésorerie globale"', "SoldeReel_Total", "Prev_Total_30", "Prev_Total_60", "Prev_Total_90",
                   "tab_eur_gras")]
        for i, (lib, *noms, style) in enumerate(lignes):
            rr = r + 2 + i
            ws.merge_range(rr, 1, rr, 3, "", self.f["tab_texte"])
            ws.write_formula(rr, 1, lib, self.f["tab_texte"])
            for j, nom in enumerate(noms):
                ws.merge_range(rr, 4 + 2 * j, rr, 5 + 2 * j, "", self.f[style])
                ws.write_formula(rr, 4 + 2 * j, "=" + nom, self.f[style])
            ws.conditional_format(rr, 4, rr, 11, {"type": "cell", "criteria": "<", "value": 0,
                                                  "format": self.wb.add_format({"font_color": ROUGE})})

        # Soldes jour par jour
        r = 52
        ws.merge_range(r, 1, r, 15, "", self.f["section"])
        ws.write_formula(r, 1, '="SOLDES PRÉVISIONNELS JOUR PAR JOUR (J+0 à J+"&Horizon_Jours&")"', self.f["section"])
        ws.set_row(r, 20)
        entetes = [("Date", 2), ("Mouvements " + "B1", 2), ("Solde B1", 2), ("Mouvements B2", 2), ("Solde B2", 2),
                   ("Solde global", 2)]
        col = 1
        titres = ['="Date"', '="Mouvements "&B1_Nom', '="Solde "&B1_Nom', '="Mouvements "&B2_Nom', '="Solde "&B2_Nom',
                  '="Solde global"']
        f_ent = self.fmt(bold=True, font_color="white", bg_color=BLEU_NUIT, align="center", valign="vcenter",
                         text_wrap=True, border=1, border_color=BORDURE, font_size=9)
        for (lib, larg), titre in zip(entetes, titres):
            ws.merge_range(r + 1, col, r + 1, col + larg - 1, "", f_ent)
            ws.write_formula(r + 1, col, titre, f_ent)
            col += larg
        ws.set_row(r + 1, 30)
        styles = {"date": self.fmt(num_format="ddd dd/mm/yyyy", font_size=9),
                  "eur": self.fmt(num_format=S.FMT_EUR, font_size=9),
                  "solde": self.fmt(num_format=S.FMT_EUR, font_size=9, bold=True)}
        sources = [("E", "date"), ("K", "eur"), ("G", "solde"), ("L", "eur"), ("I", "solde"), ("J", "solde")]
        for k in range(91):
            rr = r + 2 + k
            col = 1
            for src, style in sources:
                ws.merge_range(rr, col, rr, col + 1, "", styles[style])
                ws.write_formula(rr, col, f"=IF({k}>Horizon_Jours,\"\",'{S.F_CALC}'!{src}{2 + k})", styles[style])
                col += 2
        premiere, derniere = r + 3, r + 93  # numéros de ligne Excel
        ws.conditional_format(r + 2, 1, r + 92, 12, {"type": "formula", "criteria": f"=$B{premiere}<>\"\"",
                                                     "format": self.wb.add_format({"border": 1, "border_color": BORDURE})})
        for c0 in (5, 9, 11):
            ws.conditional_format(r + 2, c0, r + 92, c0 + 1, {"type": "cell", "criteria": "<", "value": 0,
                                                              "format": self.f["fond_rouge"]})
        ws.freeze_panes(6, 0)

    def onglet_aide(self):
        ws = self.ws[S.F_AIDE]
        self.entete_onglet(ws, "Aide - utilisation quotidienne", "Le guide complet est fourni avec l'outil (docs/).")
        ws.set_column(1, 1, 4)
        ws.set_column(2, 2, 110)
        etapes = [
            ("1", "Saisir les opérations prévues", "Onglet Saisie : formulaire « Nouvelle opération » puis « Ajouter ». "
             "Dépense ou recette : saisir un montant positif, le signe est appliqué automatiquement. Un virement entre les "
             "deux banques crée deux lignes liées (sortie / entrée) et n'est compté ni en dépense ni en recette."),
            ("2", "Enregistrer les charges fixes", "Onglet Charges fixes : loyer, salaires, abonnements... Les échéances "
             "futures sont générées automatiquement. Modifier un montant ne change pas les échéances déjà réalisées."),
            ("3", "Déposer les relevés bancaires", "Exporter les relevés depuis le site de chaque banque (CSV ou Excel) et "
             "les enregistrer dans les dossiers Imports\\Banque1 et Imports\\Banque2 situés à côté du classeur. "
             "Il n'est pas nécessaire de supprimer les anciens relevés : les doublons sont détectés."),
            ("4", "Actualiser", "Tableau de bord : bouton « Actualiser les banques ». Le classeur est sauvegardé, les "
             "relevés sont lus, les nouvelles opérations ajoutées et les correspondances recherchées."),
            ("5", "Valider les rapprochements", "Onglet Rapprochement : « Oui » pour valider une correspondance, « Non » "
             "pour la rejeter, puis « Valider les choix ». Le bouton « Tout valider (certaines) » valide en une fois les "
             "correspondances certaines. Les correspondances incertaines doivent toujours être choisies."),
            ("6", "Traiter les opérations non prévues", "Onglet À traiter : catégorie, commentaire et action (conserver, "
             "virement interne, rapprocher d'une prévision, créer une charge récurrente) puis « Appliquer les traitements »."),
            ("7", "Consulter", "Tableau de bord (soldes réels et prévisionnels, graphique 30/60/90 jours, soldes jour par "
             "jour) et Suivi prévisionnel (chronologie filtrable par banque et période)."),
            ("!", "En cas de problème", "Un message explique l'erreur. Aucune donnée importée n'est jamais supprimée et une "
             "sauvegarde est créée avant chaque actualisation (dossier Sauvegardes)."),
        ]
        r = 4
        for num, titre, texte in etapes:
            ws.write(r, 1, num, self.fmt(bold=True, font_color="white", bg_color=BLEU if num != "!" else AMBRE,
                                         align="center", valign="top", font_size=12))
            ws.write(r, 2, titre, self.fmt(bold=True, font_color=BLEU_NUIT, font_size=11))
            ws.write(r + 1, 2, texte, self.f["texte"])
            ws.set_row(r + 1, 44)
            r += 3
        ws.write(r, 2, "Légende des couleurs : cellules jaunes = saisie ; en-têtes bleus = colonnes modifiables ; "
                       "en-têtes gris = colonnes calculées ou gérées par l'outil.", self.f["texte_gris"])
        self.bouton(ws, "C" + str(r + 3), "◄ Tableau de bord", "AllerTableauDeBord", 160, 28, BLEU_NUIT)


# ---------------------------------------------------------------------------
def construire(dossier_sortie: str, ref: date | None = None, nom_fichier: str = "Tresorerie.xlsm") -> str:
    ref = ref or date.today()
    os.makedirs(dossier_sortie, exist_ok=True)
    etat = dd.etat_initial(ref)
    chemin = os.path.join(dossier_sortie, nom_fichier)
    Generateur(chemin, etat).generer()
    f1, f2 = dd.releves_demo(etat)
    for sous in ("Imports", "Sauvegardes"):
        p = os.path.join(dossier_sortie, sous)
        if os.path.isdir(p):
            shutil.rmtree(p)
    dd.ecrire_releves(dossier_sortie, f1, f2)
    os.makedirs(os.path.join(dossier_sortie, "Sauvegardes"), exist_ok=True)
    with open(os.path.join(dossier_sortie, "Sauvegardes", "LISEZMOI.txt"), "w", encoding="utf-8-sig") as f:
        f.write("Copies datées du classeur créées automatiquement avant chaque actualisation.\r\n"
                "Le nombre de copies conservées se règle dans l'onglet Paramètres.\r\n")
    ecrire_recette_attendue(dossier_sortie, ref)
    return chemin


def resultats_recette(ref: date, dossier_releves: str) -> dict:
    """Valeurs attendues de la recette Windows (install/Recette.ps1) : même chaîne que
    modRecette.Recette_Executer, calculée par le moteur de référence à partir des
    fichiers de relevés réellement livrés."""
    from reference.formats import vers_parametres_moteur
    etat = dd.etat_initial(ref)
    p1 = vers_parametres_moteur(FORMAT_B1, dd.NOM_B1)
    p2 = vers_parametres_moteur(FORMAT_B2, dd.NOM_B2)
    f1, f2 = dd.lignes_brutes_depuis_fichiers(dossier_releves, p1, p2)
    m.importer(etat, m.preparer_import(f1, p1), datetime.now())
    m.importer(etat, m.preparer_import(f2, p2), datetime.now())
    m.generer_echeances(etat)
    m.rechercher(etat)
    m.traitements_automatiques(etat)
    m.valider_certaines(etat)
    b1, b2 = dd.NOM_B1, dd.NOM_B2
    r = {
        "SoldeReel_B1": m.solde_reel(etat, b1), "SoldeReel_B2": m.solde_reel(etat, b2),
        "SoldeReel_Total": round(m.solde_reel(etat, b1) + m.solde_reel(etat, b2), 2),
        "Prev_B1_30": m.solde_previsionnel(etat, b1, 30), "Prev_B2_30": m.solde_previsionnel(etat, b2, 30),
        "Prev_Total_30": round(m.solde_previsionnel(etat, b1, 30) + m.solde_previsionnel(etat, b2, 30), 2),
        "Prev_B1_90": m.solde_previsionnel(etat, b1, 90), "Prev_B2_90": m.solde_previsionnel(etat, b2, 90),
        "Prev_Total_90": round(m.solde_previsionnel(etat, b1, 90) + m.solde_previsionnel(etat, b2, 90), 2),
        "Depenses_A_Venir": m.total_a_venir(etat, m.TYPE_DEPENSE, 30),
        "Recettes_Attendues": m.total_a_venir(etat, m.TYPE_RECETTE, 30),
        "NbATraiter": len(m.a_traiter(etat)), "NbPropositions": len(etat.propositions),
        "NbOpsProposees": len({pr.op.id for pr in etat.propositions}),
        "NbEchues": sum(1 for p in etat.previsions if m.statut_suivi(etat.params, p) == m.PREV_A_CONTROLER),
        "NbOperations": len(etat.operations),
        "NbPointees": sum(1 for o in etat.operations if o.statut == m.OB_POINTE),
        "NbFichiersHistorique": len(etat.historique),
    }
    return r


def ecrire_recette_attendue(dossier_sortie: str, ref: date) -> None:
    import json
    attendu = {"date_reference": ref.isoformat(),
               "date_reference_serie": (ref - date(1899, 12, 30)).days,
               "valeurs": resultats_recette(ref, dossier_sortie)}
    with open(os.path.join(dossier_sortie, "recette_attendue.json"), "w", encoding="utf-8") as f:
        json.dump(attendu, f, ensure_ascii=False, indent=2)


def main():
    ap = argparse.ArgumentParser(description="Génère le classeur de trésorerie et les relevés de démonstration.")
    ap.add_argument("--sortie", default=os.path.join(RACINE, "dist"))
    ap.add_argument("--date", help="date de référence des données de démonstration (AAAA-MM-JJ)")
    a = ap.parse_args()
    ref = datetime.strptime(a.date, "%Y-%m-%d").date() if a.date else None
    print(construire(a.sortie, ref))


if __name__ == "__main__":
    main()
