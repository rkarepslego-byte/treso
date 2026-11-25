"""Intégration des requêtes Power Query dans un classeur .xlsx/.xlsm.

Format décrit par la spécification publique [MS-QDEFF] (Query Definition File Format) :
la partie XML personnalisée customXml/itemN.xml contient un élément <DataMashup>
dont le contenu base64 est :

    Version(4) | Taille(4) + paquet OPC (Formulas/Section1.m, Config/Package.xml)
    | Taille(4) + Permissions (XML) | Taille(4) + Métadonnées | Taille(4) + liaisons

Les requêtes chargées dans une feuille sont reliées à une table par une connexion
(xl/connections.xml) et une partie queryTable (xl/queryTables/queryTableN.xml).

La structure suit celle produite par Excel et par la bibliothèque open source de
Microsoft « connected-workbooks » (licence MIT), dont provient le bloc
PERMISSION_BINDINGS ci-dessous (bloc DPAPI non vérifiable sur un autre poste : Excel
utilise alors les permissions par défaut, comportement prévu par [MS-QDEFF] 2.6).
"""

from __future__ import annotations

import base64
import io
import json
import re
import struct
import uuid
import zipfile
from datetime import datetime, timezone
from xml.sax.saxutils import escape

BOM = "﻿"

PERMISSION_BINDINGS = bytes.fromhex(
    "01000000d08c9ddf0115d1118c7a00c04fc297eb010000006cea9252aefc814689434c4718f5bafe0000000002000000"
    "000003660000c000000010000000cd8006aae450758b223bad9b04fc692c0000000004800000a0000000100000000"
    "7efb8a64d9cdfc9e5f81f02435c50f250000000834e08154d775c7a2d5d0384294b610cde3f730f1b8b44eccf705f9"
    "5943dd3387b9c7548a59df5468858118faa03a616ad113c1cc2090ffb15c1e138b16d8bc3a7dd329fb6a31456d5047"
    "e35b312be1314000000e75303e3f2864339cb1a080959d82968f1540c9b"
)


class Requete:
    def __init__(self, nom: str, code: str, table: str | None = None, colonnes: list[str] | None = None,
                 type_resultat: str = "Table"):
        self.nom = nom
        self.code = code.strip()
        self.table = table              # nom de la table Excel cible (None = connexion seule)
        self.colonnes = colonnes or []
        self.type_resultat = type_resultat


def section_m(requetes: list[Requete]) -> str:
    morceaux = ["section Section1;"]
    for r in requetes:
        morceaux.append(f"shared {r.nom} = {r.code};")
    return "\r\n\r\n".join(morceaux) + "\r\n"


def _zip(parties: dict[str, bytes]) -> bytes:
    tampon = io.BytesIO()
    with zipfile.ZipFile(tampon, "w", zipfile.ZIP_DEFLATED) as z:
        for nom, contenu in parties.items():
            info = zipfile.ZipInfo(nom, date_time=(1980, 1, 1, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            z.writestr(info, contenu)
    return tampon.getvalue()


def _entree(type_: str, valeur: str) -> str:
    return f'<Entry Type="{type_}" Value="{escape(valeur, {chr(34): "&quot;"})}" />'


def _metadonnees(requetes: list[Requete]) -> str:
    maintenant = datetime.now(timezone.utc).strftime("%Y-%m-%dT%H:%M:%S.0000000Z")
    items = ["<Item><ItemLocation><ItemType>AllFormulas</ItemType><ItemPath /></ItemLocation><StableEntries /></Item>"]
    for r in requetes:
        e = [_entree("IsPrivate", "l0")]
        if r.table:
            ident = [f"Section1/{r.nom}/AutoRemovedColumns1.{{{c},{i}}}" for i, c in enumerate(r.colonnes)]
            relations = json.dumps({
                "columnCount": len(r.colonnes), "keyColumnNames": [], "queryRelationships": [],
                "columnIdentities": ident, "ColumnCount": len(r.colonnes), "KeyColumnNames": [],
                "ColumnIdentities": ident, "RelationshipInfo": [],
            }, ensure_ascii=False, separators=(",", ":"))
            e += [
                _entree("FillEnabled", "l1"), _entree("FillObjectType", "sTable"),
                _entree("FillToDataModelEnabled", "l0"), _entree("BufferNextRefresh", "l1"),
                _entree("ResultType", "sTable"), _entree("NameUpdatedAfterFill", "l0"),
                _entree("NavigationStepName", "sNavigation"), _entree("FillTarget", "s" + r.table),
                _entree("FilledCompleteResultToWorksheet", "l1"), _entree("AddedToDataModel", "l0"),
                _entree("FillCount", "l0"), _entree("FillErrorCode", "sUnknown"), _entree("FillErrorCount", "l0"),
                _entree("FillLastUpdated", "d" + maintenant),
                _entree("FillColumnNames", "s" + json.dumps(r.colonnes, ensure_ascii=False, separators=(",", ":"))),
                _entree("FillStatus", "sComplete"), _entree("RelationshipInfoContainer", "s" + relations),
            ]
        else:
            e += [
                _entree("FillEnabled", "l0"), _entree("FillObjectType", "sConnectionOnly"),
                _entree("FillToDataModelEnabled", "l0"), _entree("BufferNextRefresh", "l1"),
                _entree("ResultType", "s" + r.type_resultat), _entree("NavigationStepName", "sNavigation"),
                _entree("AddedToDataModel", "l0"), _entree("FillErrorCode", "sUnknown"),
                _entree("FillLastUpdated", "d" + maintenant), _entree("FillStatus", "sComplete"),
            ]
        items.append(
            f"<Item><ItemLocation><ItemType>Formula</ItemType><ItemPath>Section1/{escape(r.nom)}</ItemPath>"
            f"</ItemLocation><StableEntries>{''.join(e)}</StableEntries></Item>"
        )
    return (BOM + '<?xml version="1.0" encoding="utf-8"?><LocalPackageMetadataFile '
            'xmlns:xsd="http://www.w3.org/2001/XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">'
            f"<Items>{''.join(items)}</Items></LocalPackageMetadataFile>")


def datamashup_xml(requetes: list[Requete]) -> bytes:
    """Contenu de customXml/itemN.xml (UTF-16 avec BOM, comme Excel)."""
    paquet = _zip({
        "[Content_Types].xml": (BOM + '<?xml version="1.0" encoding="utf-8"?><Types xmlns="http://schemas.openxmlformats.org/'
                                'package/2006/content-types"><Default Extension="xml" ContentType="text/xml" />'
                                '<Default Extension="m" ContentType="application/x-ms-m" /></Types>').encode("utf-8"),
        "Config/Package.xml": (BOM + '<?xml version="1.0" encoding="utf-8"?><Package xmlns:xsd="http://www.w3.org/2001/'
                               'XMLSchema" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><Version>2.99.202.0'
                               '</Version><MinVersion>2.21.0.0</MinVersion><Culture>fr-FR</Culture></Package>').encode("utf-8"),
        "Formulas/Section1.m": section_m(requetes).encode("utf-8"),
    })
    permissions = (BOM + '<?xml version="1.0" encoding="utf-8"?><PermissionList xmlns:xsd="http://www.w3.org/2001/XMLSchema" '
                   'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><CanEvaluateFuturePackages>false'
                   '</CanEvaluateFuturePackages><FirewallEnabled>true</FirewallEnabled></PermissionList>').encode("utf-8")
    meta_xml = _metadonnees(requetes).encode("utf-8")
    contenu_meta = _zip({})  # paquet OPC vide (22 octets)
    meta = struct.pack("<II", 0, len(meta_xml)) + meta_xml + struct.pack("<I", len(contenu_meta)) + contenu_meta
    binaire = (struct.pack("<I", 0)
               + struct.pack("<I", len(paquet)) + paquet
               + struct.pack("<I", len(permissions)) + permissions
               + struct.pack("<I", len(meta)) + meta
               + struct.pack("<I", len(PERMISSION_BINDINGS)) + PERMISSION_BINDINGS)
    xml = ('<?xml version="1.0" encoding="utf-16"?><DataMashup xmlns="http://schemas.microsoft.com/DataMashup">'
           + base64.b64encode(binaire).decode("ascii") + "</DataMashup>")
    return b"\xff\xfe" + xml.encode("utf-16-le")


def lire_section_m(xml_utf16: bytes) -> str:
    """Relit Section1.m depuis une partie DataMashup (contrôle)."""
    texte = xml_utf16.decode("utf-16")
    binaire = base64.b64decode(re.search(r"<DataMashup[^>]*>(.*)</DataMashup>", texte, re.S).group(1))
    taille = struct.unpack("<I", binaire[4:8])[0]
    with zipfile.ZipFile(io.BytesIO(binaire[8:8 + taille])) as z:
        return z.read("Formulas/Section1.m").decode("utf-8")


# ---------------------------------------------------------------------------
# Parties SpreadsheetML (connexions, queryTables)
# ---------------------------------------------------------------------------
NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"


def connections_xml(requetes_chargees: list[Requete]) -> str:
    cx = []
    for i, r in enumerate(requetes_chargees, start=1):
        cx.append(
            f'<connection id="{i}" keepAlive="1" name="Requête - {escape(r.nom)}" '
            f'description="Connexion à la requête « {escape(r.nom)} » du classeur." type="5" refreshedVersion="8" '
            f'background="1" saveData="1"><dbPr connection="Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;'
            f'Location={escape(r.nom)};Extended Properties=&quot;&quot;" command="SELECT * FROM [{escape(r.nom)}]"/>'
            f"</connection>"
        )
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<connections xmlns="{NS_MAIN}">'
            + "".join(cx) + "</connections>")


def query_table_xml(index: int, id_connexion: int, colonnes: list[str]) -> str:
    champs = "".join(
        f'<queryTableField id="{i}" name="{escape(c)}" tableColumnId="{i}"/>' for i, c in enumerate(colonnes, start=1)
    )
    return (f'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<queryTable xmlns="{NS_MAIN}" '
            f'name="ExternalData_{index}" connectionId="{id_connexion}" autoFormatId="16" applyNumberFormats="0" '
            f'applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0" '
            f'applyWidthHeightFormats="0"><queryTableRefresh nextId="{len(colonnes) + 1}">'
            f'<queryTableFields count="{len(colonnes)}">{champs}</queryTableFields></queryTableRefresh></queryTable>')


def item_props_xml() -> str:
    return ('<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n'
            f'<ds:datastoreItem ds:itemID="{{{str(uuid.uuid4()).upper()}}}" '
            'xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs>'
            '<ds:schemaRef ds:uri="http://schemas.microsoft.com/DataMashup"/></ds:schemaRefs></ds:datastoreItem>')
