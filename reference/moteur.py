"""Moteur de référence en Python de l'outil de trésorerie.

Ce module reproduit, de façon exécutable et testable, la logique implémentée dans
le classeur Excel :

* normalisation des relevés bancaires et clés techniques  -> Power Query (src/powerquery)
* import incrémental avec détection des doublons           -> VBA modImport
* génération des échéances des charges fixes               -> VBA modPrevisions
* moteur de rapprochement                                   -> VBA modRapprochement
* calcul des soldes réels / prévisionnels                   -> formules Excel

Il sert de spécification exécutable : les tests (tests/) vérifient les scénarios
métier sur ce moteur, puis vérifient que les formules du classeur (recalculées par
LibreOffice) donnent exactement les mêmes soldes.

Toute évolution de la logique VBA doit être répercutée ici (et inversement).
"""

from __future__ import annotations

import calendar
import csv
import io
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta

# ---------------------------------------------------------------------------
# Constantes métier (identiques à modConfig.bas)
# ---------------------------------------------------------------------------
PREV_PREVU = "Prévu"
PREV_POINTE = "Pointé"
PREV_ANNULE = "Annulé"
PREV_A_CONTROLER = "À contrôler"

OB_A_TRAITER = "À traiter"
OB_PROPOSE = "Proposé"
OB_POINTE = "Pointé"
OB_NON_PREVUE = "Non prévue"
OB_VIREMENT = "Virement interne"

TYPE_DEPENSE = "Dépense"
TYPE_RECETTE = "Recette"
TYPE_VIR_SORTANT = "Virement sortant"
TYPE_VIR_ENTRANT = "Virement entrant"
TYPES_NEGATIFS = (TYPE_DEPENSE, TYPE_VIR_SORTANT)

ORIGINE_SAISIE = "Saisie"
ORIGINE_CHARGE = "Charge fixe"

NIVEAU_CERTAINE = "Certaine"
NIVEAU_INCERTAINE = "Incertaine"

MODE_VALIDE = "Validé"
MODE_AUTO = "Automatique"
MODE_MANUEL = "Manuel"

FREQUENCES = {"Mensuelle": 1, "Trimestrielle": 3, "Semestrielle": 6, "Annuelle": 12}
# Demi-largeur (en jours) du "créneau" d'une échéance : une échéance théorique est
# considérée comme déjà couverte si une échéance figée de la même charge se trouve
# à moins de cette distance.
DEMI_CRENEAU = {1: 14, 3: 44, 6: 89, 12: 181}

# Table de translittération des accents (identique à modUtils.SansAccents)
ACCENTS_SRC = "ÀÁÂÃÄÅàáâãäåÇçÈÉÊËèéêëÌÍÎÏìíîïÑñÒÓÔÕÖòóôõöÙÚÛÜùúûüÝýÿ"
ACCENTS_DST = "AAAAAAaaaaaaCcEEEEeeeeIIIIiiiiNnOOOOOoooooUUUUuuuuYyy"
LIGATURES = {"Œ": "OE", "œ": "oe", "Æ": "AE", "æ": "ae", "ß": "ss"}

# Mots ignorés pour la similarité des libellés (identique à modUtils.MOTS_VIDES)
MOTS_VIDES = {
    "PRLV", "PRELEVEMENT", "SEPA", "VIR", "VIREMENT", "CB", "CARTE", "PAIEMENT",
    "FACTURE", "FACT", "ECHEANCE", "ECH", "REF", "EUR", "DE", "DU", "DES", "LA",
    "LE", "LES", "ET", "AU", "AUX", "EN", "SA", "SAS", "SARL", "EURL", "INST",
    "RECU", "EMIS", "FAVEUR", "ORDRE", "MOTIF", "NUM", "NO",
}


# ---------------------------------------------------------------------------
# Outils
# ---------------------------------------------------------------------------
def edate(d: date, mois: int) -> date:
    """Équivalent de la fonction Excel MOIS.DECALER / EDATE."""
    total = d.month - 1 + mois
    annee = d.year + total // 12
    m = total % 12 + 1
    jour = min(d.day, calendar.monthrange(annee, m)[1])
    return date(annee, m, jour)


def sans_accents(s: str) -> str:
    out = []
    for c in s:
        if c in LIGATURES:
            out.append(LIGATURES[c])
            continue
        i = ACCENTS_SRC.find(c)
        out.append(ACCENTS_DST[i] if i >= 0 else c)
    return "".join(out)


def normaliser_texte(s: str | None) -> str:
    """Majuscules, sans accents, uniquement A-Z/0-9 séparés par un espace."""
    s = sans_accents(s or "").upper()
    out = [c if ("A" <= c <= "Z" or "0" <= c <= "9") else " " for c in s]
    return " ".join("".join(out).split())


def mots_significatifs(libelle: str) -> list[str]:
    return [
        m for m in normaliser_texte(libelle).split()
        if len(m) >= 2 and not m.isdigit() and m not in MOTS_VIDES
    ]


def _bigrammes(mots: list[str]) -> list[str]:
    return [m[i:i + 2] for m in mots for i in range(len(m) - 1)]


def _mots_egaux(a: str, b: str) -> bool:
    if a == b:
        return True
    return len(a) >= 4 and len(b) >= 4 and (a.startswith(b) or b.startswith(a))


def similarite(a: str, b: str) -> float:
    """Similarité de libellés entre 0 et 1.

    max( Dice sur les bigrammes de caractères des mots significatifs,
         proportion des mots du libellé le plus court retrouvés dans l'autre ).
    """
    ma, mb = mots_significatifs(a), mots_significatifs(b)
    if not ma or not mb:
        return 0.0
    ba, bb = _bigrammes(ma), _bigrammes(mb)
    dice = 0.0
    if ba and bb:
        reste = list(bb)
        commun = 0
        for g in ba:
            if g in reste:
                reste.remove(g)
                commun += 1
        dice = 2.0 * commun / (len(ba) + len(bb))
    court, long_ = (ma, mb) if len(ma) <= len(mb) else (mb, ma)
    trouves = sum(1 for m in court if any(_mots_egaux(m, n) for n in long_))
    inclusion = trouves / len(court)
    return round(max(dice, inclusion), 4)


def cents(montant: float) -> int:
    """Montant en centimes (arrondi commercial)."""
    v = abs(montant) * 100 + 0.5
    return int(v) * (1 if montant >= 0 else -1)


# ---------------------------------------------------------------------------
# Modèle de données
# ---------------------------------------------------------------------------
@dataclass
class Banque:
    code: str           # "B1" / "B2" (utilisé dans les clés techniques)
    nom: str            # nom affiché
    solde_initial: float
    date_solde: date    # le solde initial inclut toutes les opérations <= cette date


@dataclass
class Parametres:
    banques: list[Banque]
    date_ref: date
    tolerance: int = 7
    seuil_similarite: float = 0.40
    ecart_certain: int = 3
    auto_valider: bool = False
    horizon_mois: int = 12

    def banque(self, nom: str) -> Banque | None:
        for b in self.banques:
            if b.nom == nom:
                return b
        return None


@dataclass
class OpBanque:
    id: str
    banque: str
    date: date
    libelle: str
    montant: float
    reference: str = ""
    date_valeur: date | None = None
    categorie: str = ""
    commentaire: str = ""
    statut: str = OB_A_TRAITER
    id_prevision: str = ""
    mode: str = ""
    id_import: str = ""
    fichier: str = ""
    cle: str = ""


@dataclass
class Prevision:
    id: str
    date: date
    banque: str
    libelle: str
    categorie: str
    type: str
    montant: float              # saisi en valeur absolue
    commentaire: str = ""
    statut: str = PREV_PREVU    # Prévu / Annulé / Pointé
    origine: str = ORIGINE_SAISIE
    ref_charge: str = ""
    id_operation: str = ""
    id_propose: str = ""
    date_reelle: date | None = None
    montant_reel: float | None = None
    lien_virement: str = ""
    empreinte: str = ""

    @property
    def montant_signe(self) -> float:
        return -abs(self.montant) if self.type in TYPES_NEGATIFS else abs(self.montant)


@dataclass
class Charge:
    id: str
    libelle: str
    banque: str
    categorie: str
    type: str
    montant: float
    premiere: date
    frequence: str
    fin: date | None = None
    actif: bool = True
    commentaire: str = ""


@dataclass
class Proposition:
    groupe: int
    niveau: str
    suggere: bool
    op: OpBanque
    prev: Prevision
    ecart: int
    sim: float
    score: float
    valider: str = ""


@dataclass
class Etat:
    """État complet du classeur (équivalent des tables Excel)."""
    params: Parametres
    operations: list[OpBanque] = field(default_factory=list)
    previsions: list[Prevision] = field(default_factory=list)
    charges: list[Charge] = field(default_factory=list)
    rejets: set[tuple[str, str]] = field(default_factory=set)
    propositions: list[Proposition] = field(default_factory=list)
    historique: list[dict] = field(default_factory=list)

    # -- accès ---------------------------------------------------------------
    def op(self, id_: str) -> OpBanque | None:
        return next((o for o in self.operations if o.id == id_), None)

    def prev(self, id_: str) -> Prevision | None:
        return next((p for p in self.previsions if p.id == id_), None)


def prochain_id(prefixe: str, ids, largeur: int) -> str:
    mx = 0
    for i in ids:
        if i and i.startswith(prefixe):
            try:
                mx = max(mx, int(i[len(prefixe):]))
            except ValueError:
                pass
    return f"{prefixe}{mx + 1:0{largeur}d}"


# ---------------------------------------------------------------------------
# Règles de calcul (formules Excel)
# ---------------------------------------------------------------------------
def date_solde(params: Parametres, nom_banque: str) -> date | None:
    b = params.banque(nom_banque)
    return b.date_solde if b else None


def op_incluse(params: Parametres, op: OpBanque) -> bool:
    """Colonne « Inclus » de tbl_Banque : l'opération n'est pas déjà dans le solde initial."""
    d0 = date_solde(params, op.banque)
    return d0 is not None and op.date > d0


def montant_a_venir(params: Parametres, p: Prevision) -> float:
    """Colonne « Montant à venir » de tbl_Previsions."""
    d0 = date_solde(params, p.banque)
    if d0 is None or not p.id:
        return 0.0
    if p.statut in (PREV_POINTE, PREV_ANNULE) or p.id_operation or p.id_propose:
        return 0.0
    if p.date <= d0:
        return 0.0
    return p.montant_signe


def statut_suivi(params: Parametres, p: Prevision) -> str:
    if p.statut == PREV_POINTE:
        return PREV_POINTE
    if p.statut == PREV_ANNULE:
        return PREV_ANNULE
    if p.id_propose or p.date < params.date_ref:
        return PREV_A_CONTROLER
    return PREV_PREVU


def solde_reel(etat: Etat, nom_banque: str) -> float:
    b = etat.params.banque(nom_banque)
    total = b.solde_initial + sum(
        o.montant for o in etat.operations if o.banque == nom_banque and o.date > b.date_solde
    )
    return round(total, 2)


def solde_previsionnel(etat: Etat, nom_banque: str, horizon_jours: int) -> float:
    limite = etat.params.date_ref + timedelta(days=horizon_jours)
    total = solde_reel(etat, nom_banque) + sum(
        montant_a_venir(etat.params, p) for p in etat.previsions
        if p.banque == nom_banque and p.date <= limite
    )
    return round(total, 2)


def total_a_venir(etat: Etat, type_: str, horizon_jours: int) -> float:
    """Dépenses à venir (valeur positive) ou recettes attendues, virements exclus."""
    limite = etat.params.date_ref + timedelta(days=horizon_jours)
    s = sum(
        montant_a_venir(etat.params, p) for p in etat.previsions
        if p.type == type_ and p.date <= limite
    )
    return round(abs(s), 2)


def serie_journaliere(etat: Etat, nom_banque: str, nb_jours: int) -> list[float]:
    """Solde prévisionnel de fin de journée de J+0 à J+nb_jours."""
    return [solde_previsionnel(etat, nom_banque, j) for j in range(nb_jours + 1)]


# ---------------------------------------------------------------------------
# Import des relevés (miroir de Power Query + modImport)
# ---------------------------------------------------------------------------
@dataclass
class LigneImport:
    """Une ligne de la table de transit tbl_Stage_Bx (sortie Power Query)."""
    banque: str
    code: str
    date: date | None
    date_valeur: date | None
    libelle: str
    libelle_norm: str
    montant: float | None
    reference: str
    fichier: str
    ligne: int | None
    occurrence: int | None = None
    cle: str = ""
    erreur: str | None = None


_FORMATS_DATE = {
    "dd/MM/yyyy": "%d/%m/%Y", "dd/MM/yy": "%d/%m/%y", "yyyy-MM-dd": "%Y-%m-%d",
    "dd-MM-yyyy": "%d-%m-%Y", "dd.MM.yyyy": "%d.%m.%Y", "MM/dd/yyyy": "%m/%d/%Y",
}


def lire_date(v, format_date: str) -> date | None:
    if v is None or v == "":
        return None
    if isinstance(v, datetime):
        return v.date()
    if isinstance(v, date):
        return v
    t = str(v).strip()
    for fmt in (_FORMATS_DATE.get(format_date, "%d/%m/%Y"), "%d/%m/%Y", "%Y-%m-%d"):
        try:
            return datetime.strptime(t, fmt).date()
        except ValueError:
            continue
    return None


def lire_montant(v, sep_decimal: str) -> float | None:
    if v is None or v == "":
        return None
    if isinstance(v, (int, float)):
        return float(v)
    t = str(v).strip()
    neg = "-" in t or t.startswith("(")
    chiffres = "".join(c for c in t if c in "0123456789,.")
    if sep_decimal == ",":
        chiffres = chiffres.replace(".", "").replace(",", ".")
    else:
        chiffres = chiffres.replace(",", "")
    if chiffres in ("", "."):
        return None
    n = float(chiffres)
    return -n if neg else n


def _trouver_colonne(entetes: list[str], attendu: str, fichier: str) -> int | None:
    if not attendu:
        return None
    cible = normaliser_texte(attendu)
    for i, h in enumerate(entetes):
        if normaliser_texte(h) == cible:
            return i
    raise ValueError(
        f"La colonne « {attendu} » est absente du fichier {fichier}. "
        f"Colonnes trouvées : {', '.join(entetes)}"
    )


def normaliser_releve(lignes: list[list], fichier: str, p: dict) -> list[LigneImport]:
    """Équivalent de fnNormaliserReleve : lignes brutes -> lignes normalisées."""
    reste = lignes[p["LignesIgnorees"]:]
    if not reste:
        return []
    entetes = ["" if h is None else str(h).strip() for h in reste[0]]
    donnees = reste[1:]
    i_date = _trouver_colonne(entetes, p["ColDate"], fichier)
    i_dv = _trouver_colonne(entetes, p["ColDateValeur"], fichier)
    i_libs = [_trouver_colonne(entetes, c.strip(), fichier) for c in p["ColsLibelle"].split(";") if c.strip()]
    i_ref = _trouver_colonne(entetes, p["ColReference"], fichier)
    mode = p["ModeMontant"]
    sep = p["SeparateurDecimal"]
    if mode == "Débit/Crédit":
        i_deb = _trouver_colonne(entetes, p["ColDebit"], fichier)
        i_cre = _trouver_colonne(entetes, p["ColCredit"], fichier)
    else:
        i_mt = _trouver_colonne(entetes, p["ColMontant"], fichier)
        i_sens = _trouver_colonne(entetes, p["ColSens"], fichier) if mode == "Montant + Sens" else None

    def cell(row, i):
        return row[i] if i is not None and i < len(row) else None

    res = []
    for n, row in enumerate(donnees):
        d = lire_date(cell(row, i_date), p["FormatDate"])
        dv = lire_date(cell(row, i_dv), p["FormatDate"]) if i_dv is not None else None
        lib = " ".join(
            str(cell(row, i)).strip() for i in i_libs
            if cell(row, i) not in (None, "") and str(cell(row, i)).strip() != ""
        )
        if mode == "Débit/Crédit":
            deb = lire_montant(cell(row, i_deb), sep)
            cre = lire_montant(cell(row, i_cre), sep)
            m = None if deb is None and cre is None else abs(cre or 0) - abs(deb or 0)
        elif mode == "Montant + Sens":
            mt = lire_montant(cell(row, i_mt), sep)
            sens = normaliser_texte(str(cell(row, i_sens) or ""))
            m = None if mt is None else (-abs(mt) if sens in p["ValeursDebit"] else abs(mt))
        else:
            mt = lire_montant(cell(row, i_mt), sep)
            m = None if mt is None else (-mt if p["InverserSigne"] else mt)
        ref = "" if cell(row, i_ref) is None else str(cell(row, i_ref)).strip()
        if d is None or m is None or m == 0:
            continue
        res.append(LigneImport(
            banque=p["Banque"], code=p["Code"], date=d, date_valeur=dv, libelle=lib,
            libelle_norm=normaliser_texte(lib), montant=round(m, 2), reference=ref,
            fichier=fichier, ligne=p["LignesIgnorees"] + 2 + n,
        ))
    return res


def lire_fichier(chemin: str, p: dict) -> list[list]:
    """Lit un fichier CSV/TXT ou XLSX en lignes brutes (équivalent Csv.Document/Excel.Workbook)."""
    low = chemin.lower()
    if low.endswith((".csv", ".txt")):
        enc = "cp1252" if p["Encodage"] == 1252 else "utf-8-sig"
        with open(chemin, encoding=enc, newline="") as f:
            texte = f.read()
        return [list(r) for r in csv.reader(io.StringIO(texte), delimiter=p["Separateur"])]
    import openpyxl
    wb = openpyxl.load_workbook(chemin, data_only=True, read_only=True)
    ws = wb[p["Feuille"]] if p["Feuille"] else wb.worksheets[0]
    lignes = [list(r) for r in ws.iter_rows(values_only=True)]
    wb.close()
    return lignes


def calculer_cles(lignes: list[LigneImport]) -> list[LigneImport]:
    """Occurrence (n-ième ligne identique d'un même fichier) + clé technique."""
    lignes = sorted(lignes, key=lambda l: (l.fichier, l.ligne))
    compteur: dict[tuple, int] = {}
    for l in lignes:
        k = (l.fichier, l.date, cents(l.montant), l.libelle_norm, l.reference)
        compteur[k] = compteur.get(k, 0) + 1
        l.occurrence = compteur[k]
        l.cle = "|".join([
            l.code, l.date.isoformat(), str(cents(l.montant)), l.libelle_norm,
            l.reference, str(l.occurrence),
        ])
    return lignes


def preparer_import(fichiers: dict[str, list[list]], p: dict) -> list[LigneImport]:
    """Équivalent de fnImporterBanque : {nom_fichier: lignes brutes} -> table de transit."""
    lignes: list[LigneImport] = []
    erreurs: list[LigneImport] = []
    for nom in sorted(fichiers):
        try:
            lignes.extend(normaliser_releve(fichiers[nom], nom, p))
        except Exception as e:  # noqa: BLE001 - miroir du « try ... otherwise » M
            erreurs.append(LigneImport(p["Banque"], p["Code"], None, None, "", "", None, "", nom, None, erreur=str(e)))
    return calculer_cles(lignes) + erreurs


def importer(etat: Etat, transit: list[LigneImport], horodatage: datetime) -> dict:
    """Ajoute au journal bancaire les lignes nouvelles (modImport.IntegrerTransit)."""
    cles = {o.cle for o in etat.operations}
    id_import = prochain_id("IMP-", [h["id"] for h in etat.historique], 5)
    par_fichier: dict[str, dict] = {}
    for l in sorted(transit, key=lambda x: (x.fichier, x.ligne or 0)):
        st = par_fichier.setdefault(l.fichier, {
            "id": id_import, "date": horodatage, "banque": l.banque, "fichier": l.fichier,
            "lues": 0, "ajoutees": 0, "doublons": 0, "anterieures": 0, "erreur": "",
        })
        if l.erreur:
            st["erreur"] = l.erreur
            continue
        st["lues"] += 1
        d0 = date_solde(etat.params, l.banque)
        if d0 is not None and l.date <= d0:
            st["anterieures"] += 1
            continue
        if l.cle in cles:
            st["doublons"] += 1
            continue
        nouvel_id = prochain_id("OB-", [o.id for o in etat.operations], 6)
        etat.operations.append(OpBanque(
            id=nouvel_id, banque=l.banque, date=l.date, libelle=l.libelle, montant=l.montant,
            reference=l.reference, date_valeur=l.date_valeur, id_import=id_import,
            fichier=l.fichier, cle=l.cle,
        ))
        cles.add(l.cle)
        st["ajoutees"] += 1
    etat.historique.extend(par_fichier.values())
    return par_fichier


# ---------------------------------------------------------------------------
# Charges fixes (miroir de modPrevisions.GenererEcheances)
# ---------------------------------------------------------------------------
def empreinte(p: Prevision) -> str:
    return "|".join([p.date.isoformat(), p.banque, p.libelle, p.type, str(cents(abs(p.montant)))])


def _est_figee(params: Parametres, p: Prevision) -> bool:
    return (
        p.statut in (PREV_POINTE, PREV_ANNULE)
        or bool(p.id_operation) or bool(p.id_propose)
        or p.date < params.date_ref
        or p.empreinte != empreinte(p)
    )


def generer_echeances(etat: Etat) -> dict:
    """(Re)génère les échéances des charges fixes, sans effet rétroactif.

    * Une échéance figée (pointée, annulée, proposée, passée ou modifiée à la main)
      n'est jamais modifiée ni supprimée ; elle « couvre » son créneau.
    * Une échéance non figée est conservée telle quelle (même ID) si elle correspond
      exactement à une échéance théorique de la définition actuelle de la charge ;
      sinon elle est supprimée.
    * Les échéances théoriques manquantes sont créées.
    """
    params = etat.params
    fin_horizon = edate(params.date_ref, params.horizon_mois)
    debut_fenetre = params.date_ref - timedelta(days=params.tolerance)

    figees: dict[str, list[Prevision]] = {}
    libres: dict[str, list[Prevision]] = {}
    for p in etat.previsions:
        if p.origine == ORIGINE_CHARGE and p.ref_charge:
            (figees if _est_figee(params, p) else libres).setdefault(p.ref_charge, []).append(p)

    a_supprimer: set[int] = set()
    a_creer: list[Prevision] = []
    for lst in libres.values():
        a_supprimer.update(id(p) for p in lst)

    for c in etat.charges:
        if not c.id or not c.actif or c.frequence not in FREQUENCES or not c.montant:
            continue
        b = params.banque(c.banque)
        if b is None:
            continue
        pas = FREQUENCES[c.frequence]
        demi = DEMI_CRENEAU[pas]
        debut = max(c.premiere, b.date_solde + timedelta(days=1), debut_fenetre)
        fin = min(fin_horizon, c.fin) if c.fin else fin_horizon
        fig = figees.get(c.id, [])
        disponibles = list(libres.get(c.id, []))
        k = 0
        while True:
            d = edate(c.premiere, k * pas)
            k += 1
            if d > fin:
                break
            if d < debut or any(abs((f.date - d).days) <= demi for f in fig):
                continue
            theorique = Prevision(
                id="", date=d, banque=c.banque, libelle=c.libelle, categorie=c.categorie,
                type=c.type, montant=abs(c.montant), origine=ORIGINE_CHARGE, ref_charge=c.id,
            )
            theorique.empreinte = empreinte(theorique)
            existante = next((e for e in disponibles if e.empreinte == theorique.empreinte), None)
            if existante is not None:
                disponibles.remove(existante)
                a_supprimer.discard(id(existante))
                existante.categorie = c.categorie
            else:
                a_creer.append(theorique)

    conservees = [p for p in etat.previsions if id(p) not in a_supprimer]
    # Les numéros ne sont jamais réutilisés (max sur TOUTES les lignes, y compris supprimées)
    tous_ids = [x.id for x in etat.previsions]
    for t in a_creer:
        t.id = prochain_id("P-", tous_ids, 6)
        tous_ids.append(t.id)
        conservees.append(t)
    etat.previsions = conservees
    return {"creees": len(a_creer), "supprimees": len(a_supprimer)}


def ajouter_virement(etat: Etat, d: date, source: str, destination: str, montant: float,
                     libelle: str, commentaire: str = "") -> tuple[Prevision, Prevision]:
    """Un virement interne = deux prévisions liées (sortant / entrant)."""
    id1 = prochain_id("P-", [x.id for x in etat.previsions], 6)
    p1 = Prevision(id1, d, source, libelle, "Virement interne", TYPE_VIR_SORTANT, abs(montant), commentaire)
    etat.previsions.append(p1)
    id2 = prochain_id("P-", [x.id for x in etat.previsions], 6)
    p2 = Prevision(id2, d, destination, libelle, "Virement interne", TYPE_VIR_ENTRANT, abs(montant), commentaire)
    etat.previsions.append(p2)
    p1.lien_virement, p2.lien_virement = id2, id1
    return p1, p2


# ---------------------------------------------------------------------------
# Rapprochement (miroir de modRapprochement)
# ---------------------------------------------------------------------------
def _op_candidate(params: Parametres, o: OpBanque) -> bool:
    return o.statut in (OB_A_TRAITER, OB_PROPOSE) and op_incluse(params, o)


def _prev_candidate(params: Parametres, p: Prevision) -> bool:
    d0 = date_solde(params, p.banque)
    return (
        bool(p.id) and d0 is not None and p.date > d0
        and p.statut not in (PREV_POINTE, PREV_ANNULE)
        and not p.id_operation and bool(p.montant)
    )


def lier(o: OpBanque, p: Prevision, mode: str) -> None:
    p.statut = PREV_POINTE
    p.id_operation = o.id
    p.id_propose = ""
    p.date_reelle = o.date
    p.montant_reel = o.montant
    o.statut = OB_POINTE
    o.id_prevision = p.id
    o.mode = mode
    if not o.categorie:
        o.categorie = p.categorie


def delier(etat: Etat, id_: str) -> bool:
    """Annule un rapprochement à partir de l'ID d'une opération ou d'une prévision."""
    o = etat.op(id_)
    p = etat.prev(id_)
    if o is not None and o.id_prevision:
        p = etat.prev(o.id_prevision)
    elif p is not None and p.id_operation:
        o = etat.op(p.id_operation)
    else:
        return False
    if p is not None:
        p.statut = PREV_PREVU
        p.id_operation = ""
        p.date_reelle = None
        p.montant_reel = None
    if o is not None:
        o.statut = OB_A_TRAITER
        o.id_prevision = ""
        o.mode = ""
    rechercher(etat)
    return True


def rechercher(etat: Etat) -> list[Proposition]:
    """Recherche des correspondances : propositions certaines / incertaines."""
    params = etat.params
    for o in etat.operations:
        if o.statut == OB_PROPOSE:
            o.statut = OB_A_TRAITER
    for p in etat.previsions:
        p.id_propose = ""

    ops = [o for o in etat.operations if _op_candidate(params, o)]
    prevs = [p for p in etat.previsions if _prev_candidate(params, p)]

    # Arêtes : même banque, même montant au centime, écart de dates <= tolérance
    aretes = []
    for o in ops:
        co = cents(o.montant)
        for p in prevs:
            if p.banque != o.banque or cents(p.montant_signe) != co:
                continue
            ecart = abs((o.date - p.date).days)
            if ecart > params.tolerance or (o.id, p.id) in etat.rejets:
                continue
            sim = similarite(o.libelle, p.libelle)
            score = round(0.5 * (1 - ecart / (params.tolerance + 1)) + 0.5 * sim, 4)
            aretes.append([o, p, ecart, sim, score])

    # Composantes connexes (union-find)
    parent: dict[str, str] = {}

    def trouver(x):
        while parent.setdefault(x, x) != x:
            parent[x] = parent[parent[x]]
            x = parent[x]
        return x

    for o, p, *_ in aretes:
        ra, rb = trouver("O" + o.id), trouver("P" + p.id)
        if ra != rb:
            parent[rb] = ra

    composantes: dict[str, list] = {}
    for a in aretes:
        composantes.setdefault(trouver("O" + a[0].id), []).append(a)

    def cle_comp(lst):
        return min((a[0].date, a[0].id) for a in lst)

    propositions: list[Proposition] = []
    for num, lst in enumerate(sorted(composantes.values(), key=cle_comp), start=1):
        ids_o = {a[0].id for a in lst}
        ids_p = {a[1].id for a in lst}
        lst.sort(key=lambda a: (-a[4], a[2], a[1].date, a[1].id, a[0].id))
        suggeres = set()
        if len(ids_o) == 1 and len(ids_p) == 1:
            o, p, ecart, sim, _score = lst[0]
            niveau = NIVEAU_CERTAINE if (ecart <= params.ecart_certain or sim >= params.seuil_similarite) else NIVEAU_INCERTAINE
            suggeres.add((o.id, p.id))
        else:
            niveau = NIVEAU_INCERTAINE
            pris_o, pris_p = set(), set()
            for o, p, *_ in lst:
                if o.id not in pris_o and p.id not in pris_p:
                    suggeres.add((o.id, p.id))
                    pris_o.add(o.id)
                    pris_p.add(p.id)
        for o, p, ecart, sim, score in sorted(lst, key=lambda a: (a[0].date, a[0].id, -a[4], a[1].date, a[1].id)):
            propositions.append(Proposition(num, niveau, (o.id, p.id) in suggeres, o, p, ecart, sim, score))
        for o, p, *_ in lst:
            o.statut = OB_PROPOSE
            if (o.id, p.id) in suggeres:
                p.id_propose = o.id

    etat.propositions = propositions
    if params.auto_valider:
        certaines = [pr for pr in propositions if pr.niveau == NIVEAU_CERTAINE]
        if certaines:
            for pr in certaines:
                lier(pr.op, pr.prev, MODE_AUTO)
            return rechercher(etat)
    return propositions


class ConflitValidation(Exception):
    pass


def valider(etat: Etat, choix: list[tuple[str, str, str]]) -> dict:
    """Applique les choix (ID opération, ID prévision, "Oui"/"Non")."""
    oui = [(o, p) for o, p, c in choix if c == "Oui"]
    vus_o, vus_p, conflits = {}, {}, []
    for o, p in oui:
        if o in vus_o:
            conflits.append(f"L'opération {o} est validée deux fois")
        if p in vus_p:
            conflits.append(f"La prévision {p} est validée deux fois")
        vus_o[o] = vus_p[p] = True
    if conflits:
        raise ConflitValidation("; ".join(conflits))
    liees = 0
    for o_id, p_id in oui:
        o, p = etat.op(o_id), etat.prev(p_id)
        if o is None or p is None or o.statut not in (OB_A_TRAITER, OB_PROPOSE):
            continue
        if p.statut in (PREV_POINTE, PREV_ANNULE) or p.id_operation:
            continue
        lier(o, p, MODE_VALIDE)
        liees += 1
    rejets = 0
    for o_id, p_id, c in choix:
        if c == "Non":
            etat.rejets.add((o_id, p_id))
            rejets += 1
    rechercher(etat)
    return {"liees": liees, "rejets": rejets}


def valider_certaines(etat: Etat) -> dict:
    choix = [(pr.op.id, pr.prev.id, "Oui") for pr in etat.propositions if pr.niveau == NIVEAU_CERTAINE]
    return valider(etat, choix)


# ---------------------------------------------------------------------------
# Opérations à traiter (miroir de modATraiter)
# ---------------------------------------------------------------------------
def a_traiter(etat: Etat) -> list[OpBanque]:
    return sorted(
        (o for o in etat.operations if o.statut == OB_A_TRAITER and op_incluse(etat.params, o)),
        key=lambda o: (o.date, o.id),
    )


def suggestion_virement(etat: Etat, o: OpBanque) -> str:
    """Cherche dans l'autre banque une opération de montant opposé (virement interne probable)."""
    candidats = [
        x for x in a_traiter(etat)
        if x.banque != o.banque and cents(x.montant) == -cents(o.montant)
        and abs((x.date - o.date).days) <= etat.params.tolerance
    ]
    candidats.sort(key=lambda x: (abs((x.date - o.date).days), x.id))
    return candidats[0].id if candidats else ""


def traiter(etat: Etat, id_op: str, action: str, categorie: str = "", commentaire: str = "",
            id_prevision: str = "") -> str:
    o = etat.op(id_op)
    if o is None or o.statut != OB_A_TRAITER:
        return "ignorée"
    if categorie:
        o.categorie = categorie
    if commentaire:
        o.commentaire = commentaire
    if action == "Conserver comme non prévue":
        o.statut = OB_NON_PREVUE
        if not o.categorie:
            o.categorie = "Non catégorisé"
    elif action == "Virement interne":
        o.statut = OB_VIREMENT
        o.categorie = o.categorie or "Virement interne"
    elif action == "Rapprocher avec la prévision indiquée":
        p = etat.prev(id_prevision)
        if p is None or p.banque != o.banque or p.statut in (PREV_POINTE, PREV_ANNULE) or p.id_operation:
            return "erreur"
        lier(o, p, MODE_MANUEL)
    elif action.startswith("Créer charge "):
        freq = {"mensuelle": "Mensuelle", "trimestrielle": "Trimestrielle",
                "semestrielle": "Semestrielle", "annuelle": "Annuelle"}[action.split()[-1]]
        o.statut = OB_NON_PREVUE
        if not o.categorie:
            o.categorie = "Non catégorisé"
        etat.charges.append(Charge(
            id=prochain_id("CF-", [c.id for c in etat.charges], 3), libelle=o.libelle,
            banque=o.banque, categorie=o.categorie,
            type=TYPE_DEPENSE if o.montant < 0 else TYPE_RECETTE, montant=abs(o.montant),
            premiere=edate(o.date, FREQUENCES[freq]), frequence=freq,
        ))
        generer_echeances(etat)
    else:
        return "ignorée"
    return "ok"


# ---------------------------------------------------------------------------
# Import de prévisions depuis un autre classeur (miroir de modPrevisions.ImporterPrevisions)
# ---------------------------------------------------------------------------
# Mots-clés recherchés dans les en-têtes (normalisés) ; l'ordre de test compte :
# une colonne n'est attribuée qu'à un seul rôle.
MOTS_COLONNES_PREVISIONS = [
    ("date", ["DATE", "ECHEANCE"]),
    ("debit", ["DEBIT", "DEPENSE", "DEPENSES", "SORTIE", "SORTIES", "DECAISSEMENT", "DECAISSEMENTS"]),
    ("credit", ["CREDIT", "RECETTE", "RECETTES", "ENTREE", "ENTREES", "ENCAISSEMENT", "ENCAISSEMENTS"]),
    ("montant", ["MONTANT", "SOMME", "VALEUR"]),
    ("banque", ["BANQUE", "COMPTE"]),
    ("categorie", ["CATEGORIE", "RUBRIQUE", "POSTE", "NATURE"]),
    ("type", ["TYPE", "SENS"]),
    ("commentaire", ["COMMENTAIRE", "COMMENTAIRES", "NOTE", "NOTES", "REMARQUE", "REMARQUES", "OBSERVATION"]),
    ("libelle", ["LIBELLE", "DESIGNATION", "DESCRIPTION", "INTITULE", "OPERATION", "OBJET", "TIERS"]),
]


def detecter_colonnes_previsions(entetes: list) -> dict:
    roles: dict[str, int] = {}
    pris: set[int] = set()
    for role, mots in MOTS_COLONNES_PREVISIONS:
        for j, h in enumerate(entetes):
            if j in pris:
                continue
            n = normaliser_texte("" if h is None else str(h))
            if n and any(mot in n.split() for mot in mots):
                roles[role] = j
                pris.add(j)
                break
    return roles


def ligne_entete_previsions(lignes: list[list]) -> int:
    for i, l in enumerate(lignes[:20]):
        r = detecter_colonnes_previsions(l)
        if "date" in r and ("montant" in r or "debit" in r or "credit" in r):
            return i
    return -1


def importer_previsions(etat: Etat, lignes: list[list], banque_defaut: str) -> dict:
    """Ajoute les prévisions lues ; renvoie les compteurs (importées, ignorées, doublons, antérieures)."""
    res = {"importees": 0, "ignorees": 0, "doublons": 0, "anterieures": 0}
    i0 = ligne_entete_previsions(lignes)
    if i0 < 0:
        raise ValueError("En-tête introuvable (colonnes Date et Montant / Débit / Crédit attendues)")
    r = detecter_colonnes_previsions(lignes[i0])
    existantes = {(p.date, p.banque, p.libelle, cents(p.montant_signe)) for p in etat.previsions}

    def cell(l, role):
        j = r.get(role)
        return l[j] if j is not None and j < len(l) else None

    for l in lignes[i0 + 1:]:
        d = lire_date(cell(l, "date"), "dd/MM/yyyy")
        if "debit" in r or "credit" in r:
            deb = lire_montant(cell(l, "debit"), ",") or 0.0
            cre = lire_montant(cell(l, "credit"), ",") or 0.0
            mt = abs(cre) - abs(deb)
        else:
            mt = lire_montant(cell(l, "montant"), ",") or 0.0
        typ = normaliser_texte(str(cell(l, "type") or ""))
        if typ and ("RECETTE" in typ or "ENCAISS" in typ or typ in ("C", "CREDIT", "ENTREE")):
            mt = abs(mt)
        elif typ and ("DEPENSE" in typ or "DECAISS" in typ or typ in ("D", "DEBIT", "SORTIE")):
            mt = -abs(mt)
        lib = str(cell(l, "libelle") or "").strip() or str(cell(l, "categorie") or "").strip()
        if d is None or not mt or not lib:
            if any(c not in (None, "") for c in l):
                res["ignorees"] += 1
            continue
        nom_b = banque_defaut
        b_txt = normaliser_texte(str(cell(l, "banque") or ""))
        for b in etat.params.banques:
            nb = normaliser_texte(b.nom)
            if b_txt and (nb in b_txt or b_txt in nb or b_txt == b.code):
                nom_b = b.nom
        if d <= date_solde(etat.params, nom_b):
            res["anterieures"] += 1
            continue
        cle = (d, nom_b, lib, cents(mt))
        if cle in existantes:
            res["doublons"] += 1
            continue
        existantes.add(cle)
        p = Prevision(prochain_id("P-", [x.id for x in etat.previsions], 6), d, nom_b, lib,
                      str(cell(l, "categorie") or "").strip(), TYPE_RECETTE if mt > 0 else TYPE_DEPENSE, abs(mt),
                      str(cell(l, "commentaire") or "").strip(), origine="Import")
        etat.previsions.append(p)
        res["importees"] += 1
    return res
