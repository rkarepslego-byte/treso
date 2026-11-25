"""Prépare le classeur de trésorerie du client à partir de ses propres fichiers.

Entrées :
  * l'export CSV de la banque 1 (BPGO) et l'export Excel de la banque 2 (CIC) ;
  * l'ancien classeur de trésorerie (un onglet par mois, ex. « DEC26 »).

Traitement :
  1. soldes de départ : solde de fin d'export de chaque banque, ramené au soir de la date de
     départ (les opérations postérieures restent dans les relevés et seront importées) ;
  2. charges fixes reprises de l'onglet du mois modèle (outils/migrer_treso.py), puis calées sur
     l'historique bancaire : prochaine échéance après la date de départ (mois trouvé dans les
     relevés pour les charges trimestrielles, semestrielles ou annuelles), doublons probables
     désactivés, ligne de salaires globale remplacée par un salaire par personne quand les
     virements individuels d'un même jour totalisent exactement son montant ;
  3. catégories de l'ancien classeur ; règles de catégorisation apprises (tiers de l'ancien
     classeur et libellés bancaires des charges), conservées seulement si elles reconnaissent au
     moins une opération réelle ;
  4. génération du classeur, copie des relevés dans Imports\\Banque1 et Imports\\Banque2,
     contrôle de bout en bout avec le moteur de référence et rapport (rapport_reprise.txt).

Le classeur produit contient des données réelles : le dossier de sortie (par défaut client/)
est exclu du dépôt git.

Usage :
  python -m outils.preparer_client --bpgo export.csv --cic comptes.xlsx --ancien TRESO.xlsx \\
      --onglet DEC26 --annee 2026 --mois 12 --depart 2026-11-24 [--sortie client]
"""

from __future__ import annotations

import argparse
import os
import shutil
import sys
from collections import Counter, defaultdict
from datetime import date, datetime, timedelta

RACINE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if RACINE not in sys.path:
    sys.path.insert(0, RACINE)

from reference import moteur as m  # noqa: E402
from reference.formats import FORMAT_B1, FORMAT_B2, vers_parametres_moteur  # noqa: E402
from outils.migrer_treso import CATEGORIES_CLIENT, categorie, lire_mois  # noqa: E402

MOIS_FREQ = {"Mensuelle": 1, "Trimestrielle": 3, "Semestrielle": 6, "Annuelle": 12}
TOLERANCE_MONTANT = 0.15     # écart relatif admis pour reconnaître une charge à montant variable
SIMILARITE_MIN = 0.30
MOTS_TROP_GENERAUX = {"VIR", "PRLV", "SEPA", "INST", "PRET", "LOYER", "FRAIS", "PAIEMENT", "CARTE", "FACTURE",
                      "FAC", "SALAIRE", "SALAIRES", "GLOBAL", "PREL", "DIVERS", "BPGO", "CIC", "ROCHE", "VERS",
                      "VIRT", "VIREMENT", "REGLEMENT", "REMBOURSEMENT", "REMB", "TPE", "SOLUTION", "SERVICE",
                      "SERVICES", "TECH", "NEW", "GROUP", "FRANCE", "SARL", "SAS", "CONSEILS", "PROFORMA",
                      "TRANSPORT", "TRANSPORTS", "INTERNATIONAL", "EUROPE", "INDUSTRIE", "OUEST", "SUR",
                      "CHQ", "REM", "PSC", "DONT", "CAP", "PART", "SOC", "POUR", "COMPTE", "INTERNE"}


# ---------------------------------------------------------------------------
# Relevés
# ---------------------------------------------------------------------------
def lire_releve(chemin: str, params: dict) -> list[m.LigneImport]:
    lignes = m.preparer_import({os.path.basename(chemin): m.lire_fichier(chemin, params)}, params)
    erreurs = [l for l in lignes if l.erreur]
    if erreurs:
        raise ValueError(f"{os.path.basename(chemin)} : {erreurs[0].erreur}")
    return lignes


def solde_fin_export(chemin: str) -> float:
    """Solde de fin indiqué dans l'export (ligne « Solde en fin de période » ou « Solde au … »)."""
    if chemin.lower().endswith((".csv", ".txt")):
        for l in open(chemin, encoding="cp1252").read().splitlines():
            if l.lower().startswith("solde en fin de p"):
                return float(l.split(";")[3].replace(" ", "").replace(",", "."))
    else:
        import openpyxl
        wb = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
        ws = wb[FORMAT_B2["Feuille"]] if FORMAT_B2["Feuille"] in wb.sheetnames else wb.worksheets[0]
        for r in ws.iter_rows(values_only=True):
            for j, c in enumerate(r):
                if isinstance(c, str) and c.strip().lower().startswith("solde au"):
                    v = next((x for x in r[j + 1:] if isinstance(x, (int, float))), None)
                    if v is not None:
                        wb.close()
                        return float(v)
        wb.close()
    raise ValueError(f"Solde de fin introuvable dans {os.path.basename(chemin)}")


def solde_au(lignes, solde_fin: float, depart: date) -> float:
    """Solde au soir de « depart » = solde de fin d'export - opérations postérieures."""
    return round(solde_fin - sum(l.montant for l in lignes if l.date > depart), 2)


# ---------------------------------------------------------------------------
# Calage des charges fixes sur l'historique
# ---------------------------------------------------------------------------
def _date_mois(annee: int, mois: int, jour: int) -> date:
    import calendar
    while mois > 12:
        annee, mois = annee + 1, mois - 12
    return date(annee, mois, min(jour, calendar.monthrange(annee, mois)[1]))


def correspondances(charge: m.Charge, historique: list[m.LigneImport]) -> list[tuple[m.LigneImport, bool]]:
    """Opérations passées reconnues pour la charge : (ligne, montant exact ?)."""
    res = []
    for l in historique:
        if l.banque != charge.banque or (l.montant < 0) != (charge.type == m.TYPE_DEPENSE):
            continue
        exact = m.cents(abs(l.montant)) == m.cents(charge.montant)
        proche = abs(abs(l.montant) - charge.montant) <= TOLERANCE_MONTANT * charge.montant
        if exact or (proche and m.similarite(charge.libelle, l.libelle) >= SIMILARITE_MIN):
            res.append((l, exact))
    return res


def caler_charges(charges: list[m.Charge], historique: list[m.LigneImport], depart: date) -> list[str]:
    """Fixe la première échéance de chaque charge après « depart » ; renvoie les remarques."""
    remarques = []
    passe = [l for l in historique if l.date <= depart]
    trouvees = {id(c): correspondances(c, passe) for c in charges}
    # Doublons probables : charge sans montant exact dont toutes les opérations reconnues sont déjà
    # reconnues au montant exact par une autre charge
    exactes = {id(l) for c in charges for l, ex in trouvees[id(c)] if ex}
    for c in charges:
        t = trouvees[id(c)]
        if t and not any(ex for _, ex in t) and all(id(l) in exactes for l, _ in t):
            autre = next(o for o in charges if o is not c and any(ex and l is t[0][0] for l, ex in trouvees[id(o)]))
            if abs(autre.premiere.day - c.premiere.day) > 3 or autre.frequence != c.frequence:
                trouvees[id(c)] = []      # montant voisin d'une autre ligne : pas un doublon
                continue
            c.actif = False
            c.commentaire = f"Désactivée à la reprise : doublon probable de « {autre.libelle} » ({autre.montant:.2f})"
            remarques.append(f"{c.libelle} ({c.montant:.2f}) : doublon probable de {autre.libelle} "
                             f"({autre.montant:.2f}) - désactivée")
    for c in charges:
        n = MOIS_FREQ[c.frequence]
        jour = c.premiere.day
        t = sorted(trouvees[id(c)], key=lambda x: x[0].date)
        if not t and n > 1:
            # montant variable (ex. cotisations trimestrielles) : le libellé seul suffit à situer le mois
            t = [(l, False) for l in passe if l.banque == c.banque and l.montant < 0
                 and m.similarite(c.libelle, l.libelle) >= SIMILARITE_MIN]
            t.sort(key=lambda x: x[0].date)
        if len(t) >= 2 and n > 1:
            ecarts = sorted((b[0].date - a[0].date).days for a, b in zip(t, t[1:]) if (b[0].date - a[0].date).days > 5)
            if ecarts and ecarts[len(ecarts) // 2] <= 40:
                remarques.append(f"{c.libelle} : prélevée chaque mois dans les relevés (fréquence « {c.frequence} » "
                                 "dans l'ancien classeur) - passée en mensuelle")
                c.frequence, n = "Mensuelle", 1
        if t:
            derniere = t[-1][0].date
            prochaine = _date_mois(derniere.year, derniere.month + n, jour)
            while prochaine <= depart:
                prochaine = _date_mois(prochaine.year, prochaine.month + n, jour)
            # Pour une charge mensuelle, l'échéance du mois de départ non encore passée reste due
            c.premiere = prochaine
            if n > 1:
                remarques.append(f"{c.libelle} ({c.frequence.lower()}) : dernier paiement le {derniere:%d/%m/%Y}, "
                                 f"prochaine échéance prévue le {prochaine:%d/%m/%Y}"
                                 + (" (un seul paiement trouvé dans les relevés : à confirmer)" if len(t) == 1 else ""))
        else:
            mois_depart = _date_mois(depart.year, depart.month, jour)
            if n == 1:
                c.premiere = mois_depart if mois_depart > depart else _date_mois(depart.year, depart.month + 1, jour)
            remarques.append(f"{c.libelle} ({c.montant:.2f}, {c.banque}) : aucun paiement retrouvé dans les relevés, "
                             f"première échéance le {c.premiere:%d/%m/%Y} (à vérifier)")
    return remarques


def detecter_salaires(charges: list[m.Charge], historique: list[m.LigneImport], depart: date) -> list[str]:
    """Remplace une ligne de salaires globale par un salaire par personne si, un même jour, des
    virements nominatifs totalisent exactement son montant."""
    remarques = []
    for c in list(charges):
        if "SALAIRE" not in m.normaliser_texte(c.libelle):
            continue
        par_jour = defaultdict(list)
        for l in historique:
            if (l.banque == c.banque and l.montant < 0 and l.date <= depart
                    and m.normaliser_texte(l.libelle).startswith("VIR ")
                    and not any(ch.isdigit() for ch in l.libelle)):
                par_jour[l.date].append(l)
        cible = m.cents(c.montant)
        for jour in sorted(par_jour, reverse=True):
            lot = _sous_ensemble(par_jour[jour], cible)
            if lot:
                charges.remove(c)
                for l in sorted(lot, key=lambda x: x.libelle):
                    nom = " ".join(w.capitalize() for w in m.normaliser_texte(l.libelle).split()[1:])
                    charges.append(m.Charge("", f"Salaire {nom}", c.banque, c.categorie, m.TYPE_DEPENSE,
                                            abs(l.montant), c.premiere, "Mensuelle",
                                            commentaire=f"Montant du virement du {jour:%d/%m/%Y} ; remplace « {c.libelle} »"))
                remarques.append(f"« {c.libelle} » ({c.montant:.2f}) remplacée par {len(lot)} salaires individuels "
                                 f"(virements du {jour:%d/%m/%Y}, total identique au centime)")
                break
    return remarques


def _sous_ensemble(lignes, cible: int):
    """Sous-ensemble de lignes dont la somme des montants vaut exactement « cible » (en centimes)."""
    if len(lignes) > 24:
        return None
    sommes = {0: ()}
    for i, l in enumerate(lignes):
        v = m.cents(abs(l.montant))
        for s, idx in list(sommes.items()):
            if s + v <= cible and s + v not in sommes:
                sommes[s + v] = idx + (i,)
    if cible in sommes and len(sommes[cible]) >= 2:
        return [lignes[i] for i in sommes[cible]]
    return None


# ---------------------------------------------------------------------------
# Règles de catégorisation
# ---------------------------------------------------------------------------
def tiers_ancien_classeur(chemin: str) -> list[tuple[str, str]]:
    """Couples (tiers, catégorie) de tous les onglets mensuels de l'ancien classeur."""
    import openpyxl
    wb = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
    res = []
    for ws in wb.worksheets:
        lignes = [list(r) for r in ws.iter_rows(max_col=14, values_only=True)]
        try:
            i0 = next(i for i, l in enumerate(lignes) if l and len(l) > 6
                      and str(l[1] or "").strip().lower().startswith("date") and str(l[6] or "").strip().upper() == "BANQUE")
        except StopIteration:
            continue
        for l in lignes[i0 + 1:]:
            l = (l + [None] * 14)[:14]
            tiers, type_ = str(l[2] or "").strip(), str(l[4] or "").strip()
            if tiers and type_:
                res.append((tiers, categorie(type_)))
    wb.close()
    return res


def _mot_cle(texte: str) -> str:
    """Comme l'apprentissage du classeur (3 premiers mots significatifs sans chiffre), sans les mots
    trop généraux (VIR, PRLV, LOYER…) qui ne caractérisent pas le tiers."""
    mots = [w for w in m.mots_significatifs(texte) if not any(c.isdigit() for c in w) and w not in MOTS_TROP_GENERAUX]
    return " ".join(mots[:3])


def construire_regles(couples: list[tuple[str, str]], charges: list[m.Charge],
                      historique: list[m.LigneImport], deja: list[m.Regle] = ()) -> list[m.Regle]:
    libelles = [l.libelle for l in historique]
    votes: dict[str, Counter] = defaultdict(Counter)
    for tiers, cat in couples:
        cle = _mot_cle(tiers)
        if cle:
            votes[cle][cat] += 1
    for c in charges:
        for l, ex in correspondances(c, historique):
            cle = _mot_cle(l.libelle)
            if cle:
                votes[cle][c.categorie] += 3
    couverts = {i for i, lib in enumerate(libelles) if any(m.regle_correspond(r, lib) for r in deja)}
    candidates = []
    for cle, compteur in sorted(votes.items(), key=lambda kv: (len(kv[0].split()), kv[0])):
        mots = cle.split()
        if any(w in MOTS_TROP_GENERAUX or len(w) < 3 for w in mots) or (len(mots) == 1 and len(mots[0]) < 5):
            continue
        r = m.Regle(cle, compteur.most_common(1)[0][0], "")
        reconnus = {i for i, lib in enumerate(libelles) if m.regle_correspond(r, lib)}
        # au moins une opération réelle, pas d'opération déjà couverte par une règle de départ,
        # et pas un mot si courant qu'il couvrirait une grande partie des relevés
        if not reconnus or reconnus & couverts or len(reconnus) > max(40, len(libelles) // 20):
            continue
        # une règle plus courte de même catégorie couvre déjà ces opérations : inutile
        if any(c.categorie == r.categorie and reconnus <= ens for c, ens in candidates):
            continue
        candidates.append((r, reconnus))
    return [r for r, _ in candidates]


# ---------------------------------------------------------------------------
# Construction
# ---------------------------------------------------------------------------
def preparer(bpgo: str, cic: str, ancien: str, onglet: str, annee: int, mois: int, depart: date,
             sortie: str, ref: date | None = None, societe: str = "") -> dict:
    from build import schema as S
    from build.generer_classeur import Generateur

    ref = ref or date.today()
    p1 = vers_parametres_moteur(FORMAT_B1, "BPGO")
    p2 = vers_parametres_moteur(FORMAT_B2, "CIC")
    l1, l2 = lire_releve(bpgo, p1), lire_releve(cic, p2)
    s1 = solde_au(l1, solde_fin_export(bpgo), depart)
    s2 = solde_au(l2, solde_fin_export(cic), depart)
    historique = l1 + l2

    reprise = lire_mois(ancien, onglet, annee, mois)
    charges = reprise.charges
    remarques = detecter_salaires(charges, historique, depart)
    remarques += caler_charges(charges, historique, depart)
    for i, c in enumerate(sorted(charges, key=lambda c: (c.premiere, c.banque, c.libelle)), start=1):
        c.id = f"CF-{i:03d}"
    charges.sort(key=lambda c: c.id)

    params = m.Parametres(banques=[m.Banque("B1", "BPGO", s1, depart), m.Banque("B2", "CIC", s2, depart)],
                          date_ref=ref)
    etat = m.Etat(params)
    etat.charges = charges
    etat.societe = societe
    etat.categories = [(lib, t) for _code, lib, t in CATEGORIES_CLIENT] + [("Virement interne", "Les deux"),
                                                                           ("Non catégorisé", "Les deux")]
    regles_defaut = [m.Regle(k, c, a) for k, c, a in S.REGLES_CLIENT]
    appris = construire_regles(tiers_ancien_classeur(ancien), charges, [l for l in historique if l.date <= depart],
                               regles_defaut)
    connues = {m.normaliser_texte(r.mot_cle) for r in regles_defaut}
    etat.regles = regles_defaut + [r for r in appris if m.normaliser_texte(r.mot_cle) not in connues]
    for p in reprise.previsions:
        if p.date > depart:
            p.id = m.prochain_id("P-", [x.id for x in etat.previsions], 6)
            etat.previsions.append(p)
    m.generer_echeances(etat)

    os.makedirs(sortie, exist_ok=True)
    chemin = os.path.join(sortie, "Tresorerie.xlsm")
    Generateur(chemin, etat).generer()
    for sous, src in (("Banque1", bpgo), ("Banque2", cic)):
        d = os.path.join(sortie, "Imports", sous)
        if os.path.isdir(d):
            shutil.rmtree(d)
        os.makedirs(d)
        shutil.copy2(src, os.path.join(d, os.path.basename(src)))
    os.makedirs(os.path.join(sortie, "Sauvegardes"), exist_ok=True)

    controle = simuler_premiere_actualisation(etat, bpgo, cic, p1, p2)
    rapport = {"soldes": (s1, s2), "charges": charges, "regles": etat.regles, "remarques": remarques,
               "ignorees": reprise.ignorees, "controle": controle, "chemin": chemin, "etat": etat}
    ecrire_rapport(os.path.join(sortie, "rapport_reprise.txt"), rapport, depart)
    return rapport


def simuler_premiere_actualisation(etat: m.Etat, bpgo, cic, p1, p2) -> dict:
    """Rejoue sur une copie la chaîne du bouton « Actualiser les banques »."""
    import copy
    e = copy.deepcopy(etat)
    r1 = m.importer(e, m.preparer_import({os.path.basename(bpgo): m.lire_fichier(bpgo, p1)}, p1), datetime.now())
    r2 = m.importer(e, m.preparer_import({os.path.basename(cic): m.lire_fichier(cic, p2)}, p2), datetime.now())
    m.generer_echeances(e)
    m.rechercher(e)
    auto = m.traitements_automatiques(e)
    certaines = [p for p in e.propositions if p.niveau == m.NIVEAU_CERTAINE]
    return {"import_b1": r1, "import_b2": r2, "auto": auto,
            "propositions": [(p.op.date, p.op.libelle, p.op.montant, p.prev.libelle, p.niveau) for p in e.propositions],
            "nb_certaines": len(certaines), "a_traiter": [(o.date, o.banque, o.libelle, o.montant) for o in m.a_traiter(e)],
            "solde_reel": (m.solde_reel(e, "BPGO"), m.solde_reel(e, "CIC")),
            "prev_30": (m.solde_previsionnel(e, "BPGO", 30), m.solde_previsionnel(e, "CIC", 30)),
            "prev_90": (m.solde_previsionnel(e, "BPGO", 90), m.solde_previsionnel(e, "CIC", 90))}


def ecrire_rapport(chemin: str, r: dict, depart: date) -> None:
    c = r["controle"]
    eur = lambda v: f"{v:>13,.2f} €".replace(",", " ").replace(".", ",")  # noqa: E731
    out = [f"REPRISE DES DONNÉES - classeur généré le {datetime.now():%d/%m/%Y %H:%M}", "",
           f"Soldes de départ au {depart:%d/%m/%Y} au soir :",
           f"  BPGO {eur(r['soldes'][0])}", f"  CIC  {eur(r['soldes'][1])}", "",
           f"Charges fixes : {len(r['charges'])} ({sum(1 for x in r['charges'] if x.actif)} actives)"]
    for x in r["charges"]:
        out.append(f"  {x.id}  {x.premiere:%d/%m/%Y}  {x.banque:<4} {x.frequence:<13} {eur(x.montant)}  "
                   f"{'' if x.actif else '[inactive] '}{x.libelle}")
    out += ["", "Remarques de calage :"] + [f"  - {t}" for t in r["remarques"]]
    out += ["", "Lignes de l'ancien classeur non reprises (sans jour ou sans banque) :"]
    out += [f"  - ligne {n} : {lib} ({mt:.2f}) - {raison}" for n, lib, mt, raison in r["ignorees"]] or ["  (aucune)"]
    out += ["", f"Règles de catégorisation : {len(r['regles'])}"]
    out += [f"  {x.mot_cle:<30} -> {x.categorie}{' / ' + x.action if x.action else ''}" for x in r["regles"]]
    out += ["", "Simulation de la première actualisation :",
            f"  BPGO : {_resume_import(c['import_b1'])}", f"  CIC  : {_resume_import(c['import_b2'])}",
            f"  Traitements automatiques : {c['auto']}",
            f"  Propositions de rapprochement : {len(c['propositions'])} dont {c['nb_certaines']} certaine(s)"]
    out += [f"    {d:%d/%m} {mt:>10.2f}  {lib[:45]:<45} <-> {prev[:35]} [{st}]" for d, lib, mt, prev, st in c["propositions"]]
    out += [f"  Opérations à traiter : {len(c['a_traiter'])}"]
    out += [f"    {d:%d/%m} {b:<4} {mt:>10.2f}  {lib[:70]}" for d, b, lib, mt in c["a_traiter"]]
    out += [f"  Solde réel après import : BPGO {eur(c['solde_reel'][0])} | CIC {eur(c['solde_reel'][1])}",
            f"  Solde prévisionnel à 30 j : BPGO {eur(c['prev_30'][0])} | CIC {eur(c['prev_30'][1])}",
            f"  Solde prévisionnel à 90 j : BPGO {eur(c['prev_90'][0])} | CIC {eur(c['prev_90'][1])}"]
    with open(chemin, "w", encoding="utf-8-sig", newline="\r\n") as f:
        f.write("\n".join(out) + "\n")


def _resume_import(rapport: dict) -> str:
    tot = defaultdict(int)
    for v in rapport.values():
        for k in ("ajoutees", "doublons", "anterieures"):
            tot[k] += v.get(k, 0) or 0
    return f"{tot['ajoutees']} nouvelle(s), {tot['anterieures']} antérieure(s) au solde de départ, {tot['doublons']} doublon(s)"


def main(argv=None):
    ap = argparse.ArgumentParser(description=__doc__.split("\n")[0])
    ap.add_argument("--bpgo", required=True)
    ap.add_argument("--cic", required=True)
    ap.add_argument("--ancien", required=True)
    ap.add_argument("--onglet", required=True)
    ap.add_argument("--annee", type=int, required=True)
    ap.add_argument("--mois", type=int, required=True)
    ap.add_argument("--depart", required=True, help="date du solde de départ (AAAA-MM-JJ)")
    ap.add_argument("--sortie", default=os.path.join(RACINE, "client"))
    ap.add_argument("--date", help="date de référence (AAAA-MM-JJ, défaut : aujourd'hui)")
    ap.add_argument("--societe", default="", help="nom affiché en titre du tableau de bord")
    a = ap.parse_args(argv)
    lire = lambda s: datetime.strptime(s, "%Y-%m-%d").date()  # noqa: E731
    r = preparer(a.bpgo, a.cic, a.ancien, a.onglet, a.annee, a.mois, lire(a.depart), a.sortie,
                 lire(a.date) if a.date else None, a.societe)
    print(open(os.path.join(a.sortie, "rapport_reprise.txt"), encoding="utf-8-sig").read())
    print(r["chemin"])


if __name__ == "__main__":
    main()
