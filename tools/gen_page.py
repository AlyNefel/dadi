# -*- coding: utf-8 -*-
"""Génère index.html à partir des livrables (xlsx / docx) pour que la page reflète exactement les fichiers."""
import json, os, sys
from openpyxl import load_workbook
from docx import Document
from docx.table import Table
from docx.text.paragraph import Paragraph

ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
LIV = os.path.join(ROOT, "livrables")


def sheet_rows(ws):
    rows = [[("" if v is None else v) for v in r] for r in ws.iter_rows(values_only=True)]
    rows = [r for r in rows if any(str(x).strip() for x in r)]
    return rows[0], rows[1:]


def docx_blocks(path, skip=2):
    d = Document(path)
    out, n = [], 0
    for el in d.element.body.iterchildren():
        tag = el.tag.split("}")[1]
        if tag == "p":
            p = Paragraph(el, d)
            t = p.text.strip()
            if not t:
                continue
            n += 1
            if n <= skip:
                continue
            st = p.style.name
            if st.startswith("Heading"):
                out.append({"t": "h" + st.split()[-1], "v": t})
            elif "List" in st:
                if out and out[-1]["t"] == "ul":
                    out[-1]["v"].append(t)
                else:
                    out.append({"t": "ul", "v": [t]})
            else:
                bold = all(r.bold for r in p.runs if r.text.strip()) and len(p.runs) > 0
                out.append({"t": "b" if bold else "p", "v": t})
        elif tag == "tbl":
            tb = Table(el, d)
            rows = [[c.text for c in r.cells] for r in tb.rows]
            out.append({"t": "table", "h": rows[0], "r": rows[1:]})
    return out


wb = load_workbook(os.path.join(LIV, "1_Base_prospection_CQ_7_zones.xlsx"))
data = {"sheets": {}}
for name in ["Synthèse zones", "Entreprises", "Contacts", "Offres", "Offres cabinets", "Pistes", "Lisez-moi"]:
    h, r = sheet_rows(wb[name])
    data["sheets"][name] = {"h": h, "r": r}
data["synthese"] = docx_blocks(os.path.join(LIV, "2_Synthese_opportunites_CQ.docx"))
data["referentiel"] = docx_blocks(os.path.join(LIV, "3_Referentiel_competences_CQ.docx"))
data["grille"] = docx_blocks(os.path.join(LIV, "4_Grille_evaluation_technique_CQ.docx"))
data["support"] = docx_blocks(os.path.join(LIV, "5_Support_OOS_OOT_validation.docx"))

tpl = open(os.path.join(ROOT, "tools", "page_template.html"), encoding="utf-8").read()
js = json.dumps(data, ensure_ascii=False).replace("</", "<\\/")
open(os.path.join(ROOT, "index.html"), "w", encoding="utf-8").write(tpl.replace("/*__DATA__*/null", js))
print("index.html", os.path.getsize(os.path.join(ROOT, "index.html")), "octets")
