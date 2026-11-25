"""Post-traitement du fichier produit par XlsxWriter.

* ajoute les requêtes Power Query (DataMashup, connexions, queryTables) et relie les
  tables de transit tbl_Stage_B1 / tbl_Stage_B2 à leur requête ;
* transforme les zones de texte « boutons » en boutons arrondis associés à une macro
  (attribut macro="[0]!NomMacro", la macro Workbook_Open réassocie aussi les boutons
  par sécurité).
"""

from __future__ import annotations

import re
import zipfile
from xml.sax.saxutils import escape

from build import datamashup as dm

REL_CONNECTIONS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/connections"
REL_CUSTOMXML = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml"
REL_CUSTOMXMLPROPS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps"
REL_QUERYTABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/queryTable"
CT_CONNECTIONS = "application/vnd.openxmlformats-officedocument.spreadsheetml.connections+xml"
CT_QUERYTABLE = "application/vnd.openxmlformats-officedocument.spreadsheetml.queryTable+xml"
CT_CUSTOMXMLPROPS = "application/vnd.openxmlformats-officedocument.customXmlProperties+xml"


def _ajouter_relation(rels: str, type_: str, cible: str) -> tuple[str, str]:
    ids = [int(x) for x in re.findall(r'Id="rId(\d+)"', rels)]
    rid = f"rId{max(ids, default=0) + 1}"
    rels = rels.replace("</Relationships>", f'<Relationship Id="{rid}" Type="{type_}" Target="{cible}"/></Relationships>')
    return rels, rid


def _ajouter_override(ct: str, partie: str, type_: str) -> str:
    return ct.replace("</Types>", f'<Override PartName="{partie}" ContentType="{type_}"/></Types>')


def post_traiter(chemin: str, requetes: list[dm.Requete], boutons: bool = True) -> None:
    with zipfile.ZipFile(chemin) as z:
        parties = {n: z.read(n) for n in z.namelist()}

    if requetes:
        _integrer_power_query(parties, requetes)
    if boutons:
        _boutons(parties)

    with zipfile.ZipFile(chemin, "w", zipfile.ZIP_DEFLATED) as z:
        ordre = ["[Content_Types].xml"] + [n for n in parties if n != "[Content_Types].xml"]
        for n in ordre:
            z.writestr(n, parties[n])


def _integrer_power_query(parties: dict, requetes: list[dm.Requete]) -> None:
    ct = parties["[Content_Types].xml"].decode("utf-8")
    wb_rels = parties["xl/_rels/workbook.xml.rels"].decode("utf-8")
    wb = parties["xl/workbook.xml"].decode("utf-8")

    # 1. DataMashup
    parties["customXml/item1.xml"] = dm.datamashup_xml(requetes)
    parties["customXml/itemProps1.xml"] = dm.item_props_xml().encode("utf-8")
    parties["customXml/_rels/item1.xml.rels"] = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/'
        f'package/2006/relationships"><Relationship Id="rId1" Type="{REL_CUSTOMXMLPROPS}" Target="itemProps1.xml"/>'
        "</Relationships>").encode("utf-8")
    ct = _ajouter_override(ct, "/customXml/itemProps1.xml", CT_CUSTOMXMLPROPS)
    wb_rels, _ = _ajouter_relation(wb_rels, REL_CUSTOMXML, "../customXml/item1.xml")

    # 2. Connexions
    chargees = [r for r in requetes if r.table]
    parties["xl/connections.xml"] = dm.connections_xml(chargees).encode("utf-8")
    ct = _ajouter_override(ct, "/xl/connections.xml", CT_CONNECTIONS)
    wb_rels, _ = _ajouter_relation(wb_rels, REL_CONNECTIONS, "connections.xml")

    # 3. Tables de transit -> queryTables
    noms_feuilles = re.findall(r'<sheet name="([^"]+)" sheetId="\d+"(?: state="\w+")? r:id="(rId\d+)"', wb)
    rel_feuilles = dict(re.findall(r'Id="(rId\d+)" Type="[^"]+/worksheet" Target="worksheets/([^"]+)"', wb_rels))
    noms_definis = []
    for idx, r in enumerate(chargees, start=1):
        fichier_table = None
        for n, contenu in parties.items():
            if n.startswith("xl/tables/table") and n.endswith(".xml") and f'name="{r.table}"'.encode() in contenu:
                fichier_table = n
        if fichier_table is None:
            raise ValueError(f"table {r.table} introuvable")
        table = parties[fichier_table].decode("utf-8")
        table = table.replace("<table ", '<table tableType="queryTable" ', 1)

        def col(m, compteur=[0]):  # noqa: B006 - compteur local à la table
            compteur[0] += 1
            i = compteur[0]
            return f'<tableColumn id="{m.group(1)}" uniqueName="{i}" queryTableFieldId="{i}" '
        compteur_local = [0]
        table = re.sub(r'<tableColumn id="(\d+)" ', lambda m: col(m, compteur_local), table)
        parties[fichier_table] = table.encode("utf-8")
        ref = re.search(r'ref="([A-Z]+\d+:[A-Z]+\d+)"', table).group(1)

        nom_qt = f"xl/queryTables/queryTable{idx}.xml"
        parties[nom_qt] = dm.query_table_xml(idx, idx, r.colonnes).encode("utf-8")
        ct = _ajouter_override(ct, "/" + nom_qt, CT_QUERYTABLE)
        nom_rels = fichier_table.replace("xl/tables/", "xl/tables/_rels/") + ".rels"
        parties[nom_rels] = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n<Relationships xmlns="http://schemas.openxmlformats.org/'
            f'package/2006/relationships"><Relationship Id="rId1" Type="{REL_QUERYTABLE}" '
            f'Target="../queryTables/queryTable{idx}.xml"/></Relationships>').encode("utf-8")

        # Feuille contenant la table -> nom défini ExternalData_N
        feuille_xml = None
        for n, contenu in parties.items():
            if n.startswith("xl/worksheets/_rels/") and fichier_table.split("/")[-1].encode() in contenu:
                feuille_xml = n.split("/")[-1].replace(".rels", "")
        position = None
        for pos, (nom_f, rid) in enumerate(noms_feuilles):
            if rel_feuilles.get(rid) == feuille_xml:
                position = pos
                nom_feuille = nom_f
        if position is None:
            raise ValueError(f"feuille de la table {r.table} introuvable")
        debut, fin = ref.split(":")
        abs_ref = re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", debut) + ":" + re.sub(r"([A-Z]+)(\d+)", r"$\1$\2", fin)
        noms_definis.append(f'<definedName name="ExternalData_{idx}" localSheetId="{position}" hidden="1">'
                            f"{escape(nom_feuille)}!{abs_ref}</definedName>")

    if "<definedNames>" in wb:
        wb = wb.replace("<definedNames>", "<definedNames>" + "".join(noms_definis), 1)
    else:
        wb = wb.replace("</sheets>", "</sheets><definedNames>" + "".join(noms_definis) + "</definedNames>", 1)

    parties["[Content_Types].xml"] = ct.encode("utf-8")
    parties["xl/_rels/workbook.xml.rels"] = wb_rels.encode("utf-8")
    parties["xl/workbook.xml"] = wb.encode("utf-8")


def _boutons(parties: dict) -> None:
    """Zones de texte décrites « btn:Macro » -> boutons arrondis nommés btn_Macro reliés à la macro."""
    for n in list(parties):
        if not (n.startswith("xl/drawings/drawing") and n.endswith(".xml")):
            continue
        xml = parties[n].decode("utf-8")

        def remplacer(m):
            bloc = m.group(0)
            macro = re.search(r'descr="btn:([A-Za-z0-9_]+)"', bloc)
            if not macro:
                return bloc
            nom = macro.group(1)
            bloc = bloc.replace('<xdr:sp macro=""', f'<xdr:sp macro="[0]!{nom}"', 1)
            bloc = re.sub(r'(<xdr:cNvPr id="\d+" )name="[^"]*" descr="btn:[A-Za-z0-9_]+"', rf'\1name="btn_{nom}"', bloc)
            bloc = bloc.replace('prst="rect"', 'prst="roundRect"', 1)
            return bloc

        xml = re.sub(r"<xdr:sp macro=\"\".*?</xdr:sp>", remplacer, xml, flags=re.S)
        parties[n] = xml.encode("utf-8")
