"""Analyse d'un export bancaire réel et proposition des paramètres d'import.

Usage :
    python -m outils.analyser_releve chemin/vers/releve.csv [--banque "Nom"] [--json]

Détecte l'encodage, le séparateur, la ligne d'en-tête, les colonnes (date, date de
valeur, libellé, montant ou débit/crédit, sens, référence), le format de date et le
séparateur décimal ; puis relit le fichier avec ces paramètres grâce au moteur de
référence (miroir de Power Query) et affiche un contrôle : lignes reconnues, lignes
écartées, période, totaux, lignes identiques.

Les paramètres proposés se reportent dans la table tbl_ParamImport (onglet masqué
Param_Import) et dans reference/formats.py.
"""

from __future__ import annotations

import argparse
import csv
import io
import json
import os
import re
import sys
from collections import Counter
from datetime import date, datetime

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RACINE not in sys.path:
    sys.path.insert(0, RACINE)

from reference import moteur as m  # noqa: E402
from reference.formats import PARAMETRES_IMPORT, vers_parametres_moteur  # noqa: E402

SEPARATEURS = [";", ",", "\t", "|"]
MOTS_ENTETE = ("DATE", "LIBELLE", "MONTANT", "DEBIT", "CREDIT", "VALEUR", "OPERATION", "REFERENCE",
               "DETAIL", "SENS", "DESCRIPTION", "INTITULE", "AMOUNT", "SOLDE")
RE_DATES = [
    (re.compile(r"^\d{1,2}/\d{1,2}/\d{4}$"), "dd/MM/yyyy"),
    (re.compile(r"^\d{1,2}/\d{1,2}/\d{2}$"), "dd/MM/yy"),
    (re.compile(r"^\d{4}-\d{2}-\d{2}"), "yyyy-MM-dd"),
    (re.compile(r"^\d{1,2}-\d{1,2}-\d{4}$"), "dd-MM-yyyy"),
    (re.compile(r"^\d{1,2}\.\d{1,2}\.\d{4}$"), "dd.MM.yyyy"),
]


# ---------------------------------------------------------------------------
# Lecture brute
# ---------------------------------------------------------------------------
def detecter_encodage(octets: bytes) -> tuple[int, str]:
    if octets.startswith(b"\xef\xbb\xbf"):
        return 65001, "utf-8-sig"
    try:
        octets.decode("utf-8")
        return 65001, "utf-8-sig"
    except UnicodeDecodeError:
        return 1252, "cp1252"


def detecter_separateur(lignes: list[str]) -> str:
    meilleur, score_max = ";", -1
    for sep in SEPARATEURS:
        comptes = [len(next(csv.reader([l], delimiter=sep))) for l in lignes if l.strip()]
        if not comptes:
            continue
        frequent, nb = Counter(comptes).most_common(1)[0]
        score = nb * (frequent - 1)
        if frequent > 1 and score > score_max:
            meilleur, score_max = sep, score
    return meilleur


def lire_brut(chemin: str) -> tuple[list[list], dict]:
    """Lignes brutes + paramètres de lecture détectés (encodage, séparateur, feuille)."""
    infos = {"Separateur": ";", "Encodage": 1252, "Feuille": ""}
    ext = os.path.splitext(chemin)[1].lower()
    if ext in (".csv", ".txt"):
        octets = open(chemin, "rb").read()
        code, enc = detecter_encodage(octets)
        texte = octets.decode(enc)
        sep = detecter_separateur(texte.splitlines()[:60])
        infos.update(Encodage=code, Separateur=sep)
        return [list(r) for r in csv.reader(io.StringIO(texte), delimiter=sep)], infos
    if ext in (".xlsx", ".xlsm"):
        import openpyxl
        wb = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
        ws = wb.worksheets[0]
        lignes = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
        return lignes, infos
    raise ValueError(f"Format non pris en charge par l'analyse : {ext} (convertir en .xlsx ou .csv)")


# ---------------------------------------------------------------------------
# Détection des colonnes
# ---------------------------------------------------------------------------
def _norm(v) -> str:
    return m.normaliser_texte("" if v is None else str(v))


def detecter_entete(lignes: list[list]) -> int:
    meilleur, score_max = 0, -1
    for i, ligne in enumerate(lignes[:40]):
        cellules = [_norm(c) for c in ligne if c not in (None, "")]
        score = sum(1 for c in cellules if any(mot in c for mot in MOTS_ENTETE))
        if score > score_max and score >= 2:
            meilleur, score_max = i, score
    return meilleur


def _colonne(entetes: list[str], inclure, exclure=()) -> list[int]:
    res = []
    for i, h in enumerate(entetes):
        n = _norm(h)
        if any(re.search(r"\b" + mot + r"\b", n) or mot in n.split() for mot in inclure) and not any(x in n for x in exclure):
            res.append(i)
    return res


def detecter_colonnes(entetes: list[str]) -> dict:
    c = {}
    dates = _colonne(entetes, ["DATE"], exclure=["VALEUR"])
    pref = [i for i in dates if "OPERATION" in _norm(entetes[i]) or _norm(entetes[i]) == "DATE"]
    c["date"] = (pref or dates or [None])[0]
    c["date_valeur"] = (_colonne(entetes, ["VALEUR"]) or [None])[0]
    c["libelles"] = _colonne(entetes, ["LIBELLE", "DETAIL", "DESCRIPTION", "INTITULE", "MOTIF", "NATURE"],
                             exclure=["DATE"])
    c["debit"] = (_colonne(entetes, ["DEBIT", "DEBITS"], exclure=["DATE"]) or [None])[0]
    c["credit"] = (_colonne(entetes, ["CREDIT", "CREDITS"], exclure=["DATE"]) or [None])[0]
    c["montant"] = (_colonne(entetes, ["MONTANT", "AMOUNT", "SOMME"], exclure=["DATE"]) or [None])[0]
    c["sens"] = (_colonne(entetes, ["SENS"]) or [None])[0]
    refs = _colonne(entetes, ["REFERENCE", "REF", "NUMERO"], exclure=["DATE"])
    refs += [i for i, h in enumerate(entetes) if _norm(h) in ("N OPERATION", "NO OPERATION", "N OPE") and i not in refs]
    c["reference"] = (refs or [None])[0]
    return c


def detecter_format_date(valeurs: list) -> tuple[str, bool]:
    """(format, valeurs Excel natives ?)"""
    textes = [str(v).strip() for v in valeurs if v not in (None, "") and not isinstance(v, (date, datetime))]
    natifs = sum(1 for v in valeurs if isinstance(v, (date, datetime)))
    if natifs and not textes:
        return "dd/MM/yyyy", True
    compte = Counter()
    for t in textes:
        for regex, fmt in RE_DATES:
            if regex.match(t):
                compte[fmt] += 1
                break
    if not compte:
        return "dd/MM/yyyy", False
    fmt = compte.most_common(1)[0][0]
    if fmt == "dd/MM/yyyy":
        premiers = [int(t.split("/")[0]) for t in textes if RE_DATES[0][0].match(t)]
        seconds = [int(t.split("/")[1]) for t in textes if RE_DATES[0][0].match(t)]
        if seconds and max(seconds) > 12 and max(premiers) <= 12:
            fmt = "MM/dd/yyyy"
    return fmt, False


def detecter_decimal(valeurs: list) -> str:
    virgule = point = 0
    for v in valeurs:
        if v in (None, "") or isinstance(v, (int, float)):
            continue
        t = re.sub(r"[^\d,.\-]", "", str(v))
        if re.search(r",\d{1,2}-?$", t):
            virgule += 1
        elif re.search(r"\.\d{1,2}-?$", t):
            point += 1
    return "." if point > virgule else ","


# ---------------------------------------------------------------------------
# Analyse complète
# ---------------------------------------------------------------------------
def analyser(chemin: str, nom_banque: str = "Banque", code: str = "B1") -> dict:
    lignes, infos = lire_brut(chemin)
    i_entete = detecter_entete(lignes)
    entetes = ["" if h is None else str(h).strip() for h in lignes[i_entete]]
    donnees = [l for l in lignes[i_entete + 1:] if any(c not in (None, "") for c in l)]
    col = detecter_colonnes(entetes)
    alertes = []

    def valeurs(i):
        return [l[i] if i is not None and i < len(l) else None for l in donnees]

    if col["date"] is None:
        alertes.append("Colonne de date introuvable : à renseigner à la main.")
    if col["debit"] is not None and col["credit"] is not None:
        mode = "Débit/Crédit"
        echantillon = valeurs(col["debit"]) + valeurs(col["credit"])
    elif col["montant"] is not None and col["sens"] is not None:
        mode = "Montant + Sens"
        echantillon = valeurs(col["montant"])
    elif col["montant"] is not None:
        mode = "Montant signé"
        echantillon = valeurs(col["montant"])
    else:
        mode = "Montant signé"
        echantillon = []
        alertes.append("Colonne de montant introuvable : à renseigner à la main.")
    if not col["libelles"]:
        alertes.append("Colonne de libellé introuvable : à renseigner à la main.")

    fmt_date, _natif = detecter_format_date(valeurs(col["date"]) if col["date"] is not None else [])
    sep_dec = detecter_decimal(echantillon)
    valeurs_debit = "D;DEBIT"
    if mode == "Montant + Sens":
        distincts = sorted({_norm(v) for v in valeurs(col["sens"]) if v not in (None, "")})
        debits = [v for v in distincts if v.startswith("D") or v in ("-",)]
        valeurs_debit = ";".join(debits) or "D;DEBIT"
        alertes.append(f"Valeurs de la colonne sens trouvées : {', '.join(distincts)} (débits retenus : {valeurs_debit}).")

    if col["reference"] is not None:
        refs = [str(v).strip() for v in valeurs(col["reference"]) if v not in (None, "")]
        if refs and all(r.isdigit() for r in refs) and [int(r) for r in refs] == list(range(int(refs[0]), int(refs[0]) + len(refs))):
            alertes.append(f"La colonne « {entetes[col['reference']]} » ressemble à un simple numéro de ligne : "
                           "elle n'est pas retenue comme référence (elle changerait d'un export à l'autre).")
            col["reference"] = None
        else:
            alertes.append(f"La colonne « {entetes[col['reference']]} » entre dans la clé anti-doublon : vérifier qu'elle "
                           "reste identique d'un export à l'autre pour une même opération.")

    ext = os.path.splitext(chemin)[1].lower().lstrip(".")
    params = {
        "Code": code,
        "Extensions": "csv;txt" if ext in ("csv", "txt") else "xlsx;xls",
        "Separateur": "TAB" if infos["Separateur"] == "\t" else infos["Separateur"],
        "Encodage": infos["Encodage"],
        "LignesIgnorees": i_entete,
        "Feuille": infos["Feuille"],
        "ColDate": entetes[col["date"]] if col["date"] is not None else "",
        "ColDateValeur": entetes[col["date_valeur"]] if col["date_valeur"] is not None else "",
        "FormatDate": fmt_date,
        "ColsLibelle": ";".join(entetes[i] for i in col["libelles"]),
        "ModeMontant": mode,
        "ColMontant": entetes[col["montant"]] if col["montant"] is not None and mode != "Débit/Crédit" else "",
        "ColDebit": entetes[col["debit"]] if mode == "Débit/Crédit" else "",
        "ColCredit": entetes[col["credit"]] if mode == "Débit/Crédit" else "",
        "ColSens": entetes[col["sens"]] if mode == "Montant + Sens" else "",
        "ValeursDebit": valeurs_debit,
        "InverserSigne": "Non",
        "SeparateurDecimal": sep_dec,
        "ColReference": entetes[col["reference"]] if col["reference"] is not None else "",
    }

    # Contrôle : relecture avec les paramètres proposés (même logique que Power Query)
    controle = controler(chemin, lignes, params, nom_banque, len(donnees))
    if mode == "Montant signé" and controle["nb_lignes"] and controle["nb_debits"] == 0:
        alertes.append("Aucun montant négatif : si la banque exporte les débits en positif, régler "
                       "« Inverser le signe » ou utiliser une colonne sens / débit-crédit.")
    if controle["lignes_ecartees"]:
        alertes.append(f"{len(controle['lignes_ecartees'])} ligne(s) non vide(s) écartée(s) (sans date ou sans montant) : "
                       "vérifier qu'il s'agit bien de lignes de titre, de total ou de solde.")
    if controle["doublons_internes"]:
        alertes.append(f"{controle['doublons_internes']} ligne(s) strictement identique(s) dans le fichier : "
                       "elles seront importées comme opérations distinctes (occurrences).")
    return {"fichier": os.path.basename(chemin), "entetes": entetes, "params": params, "alertes": alertes,
            "controle": controle}


def controler(chemin: str, lignes: list[list], params: dict, nom_banque: str, nb_donnees: int) -> dict:
    p = vers_parametres_moteur(params, nom_banque)
    if p["Separateur"] == "TAB":
        p["Separateur"] = "\t"
    try:
        res = m.calculer_cles(m.normaliser_releve(lignes, os.path.basename(chemin), p))
        erreur = ""
    except ValueError as e:
        res, erreur = [], str(e)
    reconnues = {l.ligne for l in res}
    ecartees = []
    for k, ligne in enumerate(lignes[params["LignesIgnorees"] + 1:]):
        num = params["LignesIgnorees"] + 2 + k
        if num not in reconnues and any(c not in (None, "") for c in ligne):
            ecartees.append((num, [c for c in ligne if c not in (None, "")]))
    return {
        "erreur": erreur,
        "nb_lignes": len(res),
        "nb_lignes_donnees": nb_donnees,
        "nb_debits": sum(1 for l in res if l.montant < 0),
        "total_debits": round(sum(l.montant for l in res if l.montant < 0), 2),
        "total_credits": round(sum(l.montant for l in res if l.montant > 0), 2),
        "du": min((l.date for l in res), default=None),
        "au": max((l.date for l in res), default=None),
        "doublons_internes": sum(1 for l in res if l.occurrence and l.occurrence > 1),
        "apercu": [(l.date.isoformat(), l.libelle, l.montant, l.reference) for l in res[:8]],
        "lignes_ecartees": ecartees[:20],
    }


def afficher(r: dict) -> None:
    print(f"\n=== {r['fichier']} ===")
    print("Colonnes trouvées :", " | ".join(r["entetes"]))
    print("\nParamètres proposés (table tbl_ParamImport) :")
    libelles = {cle: lib for cle, lib, _d in PARAMETRES_IMPORT}
    for cle, val in r["params"].items():
        if cle in libelles:
            print(f"  {libelles[cle]:<36} {val}")
    c = r["controle"]
    print("\nContrôle de relecture :")
    if c["erreur"]:
        print("  ERREUR :", c["erreur"])
    print(f"  {c['nb_lignes']} opération(s) reconnue(s) sur {c['nb_lignes_donnees']} ligne(s) de données")
    if c["du"]:
        print(f"  période du {c['du']:%d/%m/%Y} au {c['au']:%d/%m/%Y}")
    print(f"  total débits {c['total_debits']:,.2f}  total crédits {c['total_credits']:,.2f}")
    for a in c["apercu"]:
        print("   ", a)
    for num, cellules in c["lignes_ecartees"]:
        print(f"  ligne {num} écartée : {cellules}")
    if r["alertes"]:
        print("\nPoints à vérifier :")
        for a in r["alertes"]:
            print("  -", a)


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("fichiers", nargs="+")
    ap.add_argument("--banque", default="Banque")
    ap.add_argument("--code", default="B1")
    ap.add_argument("--json", action="store_true")
    a = ap.parse_args(argv)
    resultats = [analyser(f, a.banque, a.code) for f in a.fichiers]
    if a.json:
        print(json.dumps(resultats, ensure_ascii=False, indent=2, default=str))
    else:
        for r in resultats:
            afficher(r)


if __name__ == "__main__":
    main()
