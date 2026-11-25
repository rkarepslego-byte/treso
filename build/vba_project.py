"""Construction d'un fichier vbaProject.bin à partir des sources VBA (.bas/.cls).

Implémente les spécifications publiques de Microsoft :
  * [MS-CFB]  Compound File Binary Format (version 3, secteurs de 512 octets)
  * [MS-OVBA] Office VBA File Format (compression, flux dir / PROJECT / modules)

Le projet ne contient que le code source compressé (pas de p-code) : à la première
ouverture, Excel compile le code à partir des sources ([MS-OVBA] 2.3.4.1 : le flux
_VBA_PROJECT ne contient pas de cache de performance, version 0xFFFF).
"""

from __future__ import annotations

import random
import struct
import uuid

# ---------------------------------------------------------------------------
# Compression MS-OVBA (2.4.1)
# ---------------------------------------------------------------------------
_TAILLE_BLOC = 4096


def _nb_bits(position: int) -> int:
    bits = 0
    while (1 << bits) < position:
        bits += 1
    return max(bits, 4)


def _compresser_bloc(bloc: bytes) -> bytes:
    n = len(bloc)
    sortie = bytearray()
    index: dict[bytes, list[int]] = {}
    i = 0
    while i < n:
        pos_flag = len(sortie)
        sortie.append(0)
        flag = 0
        for bit in range(8):
            if i >= n:
                break
            meilleur_len, meilleur_off = 0, 0
            if i >= 1 and i + 3 <= n:
                bits = _nb_bits(i)
                max_len = (0xFFFF >> bits) + 3
                for j in reversed(index.get(bloc[i:i + 3], [])[-64:]):
                    longueur = 0
                    while i + longueur < n and longueur < max_len and bloc[j + longueur] == bloc[i + longueur]:
                        longueur += 1
                    if longueur > meilleur_len:
                        meilleur_len, meilleur_off = longueur, i - j
                        if longueur == max_len:
                            break
            if meilleur_len >= 3:
                bits = _nb_bits(i)
                jeton = ((meilleur_off - 1) << (16 - bits)) | (meilleur_len - 3)
                sortie += struct.pack("<H", jeton)
                flag |= 1 << bit
                for k in range(i, i + meilleur_len):
                    if k + 3 <= n:
                        index.setdefault(bloc[k:k + 3], []).append(k)
                i += meilleur_len
            else:
                if i + 3 <= n:
                    index.setdefault(bloc[i:i + 3], []).append(i)
                sortie.append(bloc[i])
                i += 1
        sortie[pos_flag] = flag
    if n == _TAILLE_BLOC and len(sortie) > _TAILLE_BLOC:
        return struct.pack("<H", 0x3FFF) + bloc  # bloc non compressé
    return struct.pack("<H", 0xB000 | (len(sortie) + 2 - 3)) + bytes(sortie)


def compresser(donnees: bytes) -> bytes:
    sortie = bytearray([0x01])
    for pos in range(0, len(donnees), _TAILLE_BLOC):
        sortie += _compresser_bloc(donnees[pos:pos + _TAILLE_BLOC])
    return bytes(sortie)


def decompresser(donnees: bytes) -> bytes:
    """Décompression ([MS-OVBA] 2.4.1.3.1), utilisée pour l'auto-contrôle."""
    if donnees[0] != 0x01:
        raise ValueError("signature de conteneur compressé invalide")
    sortie = bytearray()
    pos = 1
    while pos < len(donnees):
        entete = struct.unpack("<H", donnees[pos:pos + 2])[0]
        taille = (entete & 0x0FFF) + 3
        compresse = entete & 0x8000
        if (entete >> 12) & 0x7 != 0b011:
            raise ValueError("signature de bloc invalide")
        fin = pos + taille
        pos += 2
        debut_bloc = len(sortie)
        if not compresse:
            sortie += donnees[pos:pos + 4096]
            pos += 4096
            continue
        while pos < fin:
            flag = donnees[pos]
            pos += 1
            for bit in range(8):
                if pos >= fin:
                    break
                if flag & (1 << bit):
                    jeton = struct.unpack("<H", donnees[pos:pos + 2])[0]
                    pos += 2
                    bits = _nb_bits(len(sortie) - debut_bloc)
                    longueur = (jeton & (0xFFFF >> bits)) + 3
                    decalage = (jeton >> (16 - bits)) + 1
                    for _ in range(longueur):
                        sortie.append(sortie[-decalage])
                else:
                    sortie.append(donnees[pos])
                    pos += 1
    return bytes(sortie)


# ---------------------------------------------------------------------------
# Chiffrement des propriétés CMG / DPB / GC ([MS-OVBA] 2.4.3.2)
# ---------------------------------------------------------------------------
def chiffrer(project_id: str, donnees: bytes, graine: int | None = None) -> str:
    graine = random.randint(0, 255) if graine is None else graine
    version = 2
    proj_key = sum(project_id.encode("ascii")) & 0xFF
    non_chiffre1 = proj_key
    chiffre1 = graine ^ proj_key      # ProjKeyEnc
    chiffre2 = graine ^ version       # VersionEnc
    sortie = bytearray([graine, graine ^ version, graine ^ proj_key])
    nb_ignores = (graine & 6) // 2

    def encoder(octet):
        nonlocal non_chiffre1, chiffre1, chiffre2
        enc = octet ^ ((chiffre2 + non_chiffre1) & 0xFF)
        sortie.append(enc)
        chiffre2 = chiffre1
        chiffre1 = enc
        non_chiffre1 = octet

    for _ in range(nb_ignores):
        encoder(random.randint(0, 255))
    for octet in struct.pack("<I", len(donnees)):
        encoder(octet)
    for octet in donnees:
        encoder(octet)
    return sortie.hex().upper()


def dechiffrer(texte_hex: str) -> bytes:
    data = bytes.fromhex(texte_hex)
    graine, version_enc, proj_key_enc = data[0], data[1], data[2]
    if graine ^ version_enc != 2:
        raise ValueError("version invalide")
    proj_key = graine ^ proj_key_enc
    non_chiffre1, chiffre1, chiffre2 = proj_key, proj_key_enc, version_enc
    pos = 3
    nb_ignores = (graine & 6) // 2

    def decoder():
        nonlocal pos, non_chiffre1, chiffre1, chiffre2
        enc = data[pos]
        pos += 1
        octet = enc ^ ((chiffre2 + non_chiffre1) & 0xFF)
        chiffre2 = chiffre1
        chiffre1 = enc
        non_chiffre1 = octet
        return octet

    for _ in range(nb_ignores):
        decoder()
    longueur = struct.unpack("<I", bytes(decoder() for _ in range(4)))[0]
    return bytes(decoder() for _ in range(longueur))


# ---------------------------------------------------------------------------
# Flux dir ([MS-OVBA] 2.3.4.2)
# ---------------------------------------------------------------------------
CODEPAGE = "cp1252"


def _rec(ident: int, donnees: bytes) -> bytes:
    return struct.pack("<HI", ident, len(donnees)) + donnees


def _rec_double(ident: int, texte: str, ident_unicode: int) -> bytes:
    return (_rec(ident, texte.encode(CODEPAGE))
            + struct.pack("<HI", ident_unicode, len(texte.encode("utf-16-le"))) + texte.encode("utf-16-le"))


def _flux_dir(modules: list[dict]) -> bytes:
    d = bytearray()
    d += _rec(0x0001, struct.pack("<I", 1))            # PROJECTSYSKIND : Win32
    d += _rec(0x0002, struct.pack("<I", 0x0409))       # PROJECTLCID
    d += _rec(0x0014, struct.pack("<I", 0x0409))       # PROJECTLCIDINVOKE
    d += _rec(0x0003, struct.pack("<H", 1252))         # PROJECTCODEPAGE
    d += _rec(0x0004, b"VBAProject")                   # PROJECTNAME
    d += _rec_double(0x0005, "", 0x0040)               # PROJECTDOCSTRING
    d += _rec_double(0x0006, "", 0x003D)               # PROJECTHELPFILEPATH
    d += _rec(0x0007, struct.pack("<I", 0))            # PROJECTHELPCONTEXT
    d += _rec(0x0008, struct.pack("<I", 0))            # PROJECTLIBFLAGS
    d += struct.pack("<HIIH", 0x0009, 4, 1, 0)        # PROJECTVERSION
    d += _rec_double(0x000C, "", 0x003C)               # PROJECTCONSTANTS
    # Référence : OLE Automation (stdole)
    libid = b"*\\G{00020430-0000-0000-C000-000000000046}#2.0#0#C:\\Windows\\System32\\stdole2.tlb#OLE Automation"
    d += _rec_double(0x0016, "stdole", 0x003E)         # REFERENCENAME
    d += _rec(0x000D, struct.pack("<I", len(libid)) + libid + struct.pack("<IH", 0, 0))  # REFERENCEREGISTERED
    # Modules
    d += _rec(0x000F, struct.pack("<H", len(modules)))  # PROJECTMODULES
    d += _rec(0x0013, struct.pack("<H", 0xFFFF))       # PROJECTCOOKIE
    for mod in modules:
        nom = mod["nom"]
        d += _rec(0x0019, nom.encode(CODEPAGE))                                     # MODULENAME
        d += _rec(0x0047, nom.encode("utf-16-le"))                                  # MODULENAMEUNICODE
        d += _rec_double(0x001A, nom, 0x0032)                                       # MODULESTREAMNAME
        d += _rec_double(0x001C, "", 0x0048)                                        # MODULEDOCSTRING
        d += _rec(0x0031, struct.pack("<I", 0))                                     # MODULEOFFSET
        d += _rec(0x001E, struct.pack("<I", 0))                                     # MODULEHELPCONTEXT
        d += _rec(0x002C, struct.pack("<H", 0xFFFF))                                # MODULECOOKIE
        d += struct.pack("<HI", 0x0022 if mod["document"] else 0x0021, 0)          # MODULETYPE
        d += struct.pack("<HI", 0x002B, 0)                                          # Terminator
    d += struct.pack("<HI", 0x0010, 0)                                              # dir Terminator
    return bytes(d)


# ---------------------------------------------------------------------------
# Modules
# ---------------------------------------------------------------------------
ATTRIBUTS_DOCUMENT = (
    'Attribute VB_Name = "{nom}"\r\n'
    'Attribute VB_Base = "{base}"\r\n'
    "Attribute VB_GlobalNameSpace = False\r\n"
    "Attribute VB_Creatable = False\r\n"
    "Attribute VB_PredeclaredId = True\r\n"
    "Attribute VB_Exposed = True\r\n"
    "Attribute VB_TemplateDerived = False\r\n"
    "Attribute VB_Customizable = True\r\n"
)
BASE_CLASSEUR = "0{00020819-0000-0000-C000-000000000046}"
BASE_FEUILLE = "0{00020820-0000-0000-C000-000000000046}"


def _crlf(texte: str) -> str:
    texte = texte.replace("\r\n", "\n").replace("\r", "\n")
    if not texte.endswith("\n"):
        texte += "\n"
    return texte.replace("\n", "\r\n")


def module_standard(source: str) -> dict:
    source = _crlf(source.lstrip("\ufeff"))
    premiere = source.split("\r\n", 1)[0]
    if not premiere.startswith('Attribute VB_Name = "'):
        raise ValueError("Un module standard doit commencer par Attribute VB_Name")
    nom = premiere.split('"')[1]
    return {"nom": nom, "document": False, "source": source}


def module_document(nom: str, code: str, classeur: bool) -> dict:
    entete = ATTRIBUTS_DOCUMENT.format(nom=nom, base=BASE_CLASSEUR if classeur else BASE_FEUILLE)
    return {"nom": nom, "document": True, "source": entete + _crlf(code) if code.strip() else entete}


# ---------------------------------------------------------------------------
# Fichier composé ([MS-CFB])
# ---------------------------------------------------------------------------
FREESECT, ENDOFCHAIN, FATSECT, NOSTREAM = 0xFFFFFFFF, 0xFFFFFFFE, 0xFFFFFFFD, 0xFFFFFFFF
SECTEUR, MINI, SEUIL_MINI = 512, 64, 4096


class _Entree:
    def __init__(self, nom: str, type_: int, donnees: bytes = b""):
        self.nom, self.type, self.donnees = nom, type_, donnees
        self.enfants: list[_Entree] = []
        self.gauche = self.droite = self.enfant = NOSTREAM
        self.couleur = 1
        self.debut = 0
        self.index = 0


def _cle_cfb(nom: str):
    return (len(nom.encode("utf-16-le")), nom.upper())


def _arbre(entrees: list[_Entree]) -> int:
    """Arbre binaire équilibré ; la profondeur maximale est colorée en rouge (arbre rouge-noir valide)."""
    entrees = sorted(entrees, key=lambda e: _cle_cfb(e.nom))
    profondeurs: dict[int, int] = {}

    def construire(lst, prof):
        if not lst:
            return None
        milieu = len(lst) // 2
        racine = lst[milieu]
        profondeurs[id(racine)] = prof
        g = construire(lst[:milieu], prof + 1)
        d = construire(lst[milieu + 1:], prof + 1)
        racine.gauche = g.index if g else NOSTREAM
        racine.droite = d.index if d else NOSTREAM
        return racine

    racine = construire(entrees, 0)
    if racine is None:
        return NOSTREAM
    pmax = max(profondeurs.values())
    parfait = len(entrees) == (1 << (pmax + 1)) - 1
    for e in entrees:
        e.couleur = 0 if (not parfait and pmax > 0 and profondeurs[id(e)] == pmax) else 1
    return racine.index


def ecrire_cfb(racine: _Entree) -> bytes:
    # Numérotation des entrées (parcours en largeur)
    toutes: list[_Entree] = []
    file = [racine]
    while file:
        e = file.pop(0)
        e.index = len(toutes)
        toutes.append(e)
        file.extend(e.enfants)
    for e in toutes:
        if e.type in (1, 5):
            e.enfant = _arbre(e.enfants)
    racine.couleur = 1

    flux = [e for e in toutes if e.type == 2]
    petits = [e for e in flux if len(e.donnees) < SEUIL_MINI]
    grands = [e for e in flux if len(e.donnees) >= SEUIL_MINI]

    # Mini-flux
    mini_flux = bytearray()
    mini_fat: list[int] = []
    for e in petits:
        nb = (len(e.donnees) + MINI - 1) // MINI
        e.debut = len(mini_fat)
        mini_fat.extend(range(e.debut + 1, e.debut + nb))
        mini_fat.append(ENDOFCHAIN)
        mini_flux += e.donnees + b"\0" * (nb * MINI - len(e.donnees))

    secteurs: list[bytes] = []
    fat: list[int] = []

    def allouer(donnees: bytes) -> int:
        nb = max(1, (len(donnees) + SECTEUR - 1) // SECTEUR)
        debut = len(secteurs)
        donnees = donnees + b"\0" * (nb * SECTEUR - len(donnees))
        for k in range(nb):
            secteurs.append(donnees[k * SECTEUR:(k + 1) * SECTEUR])
            fat.append(debut + k + 1 if k < nb - 1 else ENDOFCHAIN)
        return debut

    for e in grands:
        e.debut = allouer(e.donnees)
    racine.debut = allouer(bytes(mini_flux)) if mini_flux else ENDOFCHAIN
    taille_racine = len(mini_flux)

    premier_minifat, nb_minifat = ENDOFCHAIN, 0
    if mini_fat:
        nb_entrees = ((len(mini_fat) + 127) // 128) * 128
        donnees_mf = struct.pack(f"<{nb_entrees}I", *(mini_fat + [FREESECT] * (nb_entrees - len(mini_fat))))
        premier_minifat = allouer(donnees_mf)
        nb_minifat = nb_entrees // 128

    # Répertoire
    rep = bytearray()
    for e in toutes:
        nom = e.nom.encode("utf-16-le")
        rep += nom + b"\0" * (64 - len(nom))
        rep += struct.pack("<HBB", len(nom) + 2, e.type, e.couleur)
        rep += struct.pack("<III", e.gauche, e.droite, e.enfant)
        rep += b"\0" * 16 + struct.pack("<I", 0) + b"\0" * 16
        if e.type == 5:
            rep += struct.pack("<IQ", racine.debut, taille_racine)
        elif e.type == 2:
            rep += struct.pack("<IQ", e.debut, len(e.donnees))
        else:
            rep += struct.pack("<IQ", 0, 0)
    while len(rep) % SECTEUR:
        rep += b"\0" * 64 + struct.pack("<HBB", 0, 0, 0) + struct.pack("<III", NOSTREAM, NOSTREAM, NOSTREAM) \
            + b"\0" * 36 + struct.pack("<IQ", 0, 0)
    premier_rep = allouer(bytes(rep))

    # FAT (elle se décrit elle-même)
    nb_fat = 1
    while (len(secteurs) + nb_fat + 127) // 128 > nb_fat:
        nb_fat += 1
    debut_fat = len(secteurs)
    fat.extend([FATSECT] * nb_fat)
    total = ((len(fat) + 127) // 128) * 128
    fat.extend([FREESECT] * (total - len(fat)))
    donnees_fat = struct.pack(f"<{total}I", *fat)
    for k in range(nb_fat):
        secteurs.append(donnees_fat[k * SECTEUR:(k + 1) * SECTEUR])
    if nb_fat > 109:
        raise ValueError("fichier trop volumineux pour cet écrivain CFB simplifié")

    entete = bytearray()
    entete += bytes.fromhex("D0CF11E0A1B11AE1") + b"\0" * 16
    entete += struct.pack("<HHHHH", 0x003E, 0x0003, 0xFFFE, 9, 6) + b"\0" * 6
    entete += struct.pack("<IIIIIIIII", 0, nb_fat, premier_rep, 0, SEUIL_MINI,
                          premier_minifat, nb_minifat, ENDOFCHAIN, 0)
    difat = list(range(debut_fat, debut_fat + nb_fat)) + [FREESECT] * (109 - nb_fat)
    entete += struct.pack("<109I", *difat)
    assert len(entete) == 512
    return bytes(entete) + b"".join(secteurs)


# ---------------------------------------------------------------------------
# Projet complet
# ---------------------------------------------------------------------------
def construire_projet(modules: list[dict], project_id: str | None = None) -> bytes:
    project_id = project_id or "{" + str(uuid.uuid4()).upper() + "}"
    lignes = [f'ID="{project_id}"']
    for mod in modules:
        if mod["document"]:
            lignes.append(f"Document={mod['nom']}/&H00000000")
    for mod in modules:
        if not mod["document"]:
            lignes.append(f"Module={mod['nom']}")
    lignes += [
        'Name="VBAProject"',
        'HelpContextID="0"',
        'VersionCompatible32="393222000"',
        f'CMG="{chiffrer(project_id, bytes(4))}"',
        f'DPB="{chiffrer(project_id, bytes(1))}"',
        f'GC="{chiffrer(project_id, bytes([0xFF]))}"',
        "",
        "[Host Extender Info]",
        "&H00000001={3832D640-CF90-11CF-8E43-00A0C911005A};VBE;&H00000000",
        "",
        "[Workspace]",
    ]
    lignes += [f"{mod['nom']}=0, 0, 0, 0, C" for mod in modules]
    flux_project = ("\r\n".join(lignes) + "\r\n").encode(CODEPAGE)

    wm = bytearray()
    for mod in modules:
        wm += mod["nom"].encode(CODEPAGE) + b"\0" + mod["nom"].encode("utf-16-le") + b"\0\0"
    wm += b"\0\0"

    racine = _Entree("Root Entry", 5)
    vba = _Entree("VBA", 1)
    racine.enfants = [vba, _Entree("PROJECT", 2, flux_project), _Entree("PROJECTwm", 2, bytes(wm))]
    vba.enfants = [
        _Entree("_VBA_PROJECT", 2, bytes([0xCC, 0x61, 0xFF, 0xFF, 0x00, 0x00, 0x00])),
        _Entree("dir", 2, compresser(_flux_dir(modules))),
    ]
    for mod in modules:
        vba.enfants.append(_Entree(mod["nom"], 2, compresser(mod["source"].encode(CODEPAGE))))
    return ecrire_cfb(racine)
