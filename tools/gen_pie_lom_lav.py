# -*- coding: utf-8 -*-
"""Genera data/pie_lom_lav.json: corrispondenze lavorazione PIE -> lavorazione LOM
(Modulo B). Fonte: RAFFRONTO/Corrispondenze_Analisi_PIE_LOM_rev7.xlsx
(pairing automatico + verifica semantica LLM + controllo prezzi, rev.7).
Include SOLO le coppie con controllo prezzi "coerente" (scelta di Gianluca 12/06/2026).
Formato: {codPIE: [codLOM, parte(A|C), prezzoLOM, banda(A|M|B|MB)]}
Richiede: pip install python-calamine --break-system-packages
"""
import json, os
from python_calamine import CalamineWorkbook

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
SRC = os.path.join(os.path.dirname(BASE), "RAFFRONTO", "Corrispondenze_Analisi_PIE_LOM_rev7.xlsx")
BANDA = {"ALTA":"A","MEDIA":"M","BASSA":"B","MOLTO BASSA":"MB"}

lomA = set(v[0] for v in json.load(open(os.path.join(BASE,"data","voci_A.json"),encoding="utf-8"))["voci"])
lomC = set(v[0] for v in json.load(open(os.path.join(BASE,"data","voci_C.json"),encoding="utf-8"))["voci"])

wb = CalamineWorkbook.from_path(SRC)
rows = wb.get_sheet_by_name("Corrispondenze rev.2").to_python(skip_empty_area=True)[1:]
out = {}
for r in rows:
    if str(r[12]) != "coerente":
        continue
    pie = str(r[1]).replace("PIE26_","").strip()
    lom = str(r[3]).replace("LOM261.","").strip()
    part = "A" if lom in lomA else ("C" if lom in lomC else None)
    if not part:
        continue
    out[pie] = [lom, part, round(float(r[10]),2), BANDA.get(str(r[7]),"?")]
dest = os.path.join(BASE,"data","pie_lom_lav.json")
json.dump(out, open(dest,"w",encoding="utf-8"), ensure_ascii=False, separators=(",",":"))
print("scritto", dest, "| coppie:", len(out), "| KB:", os.path.getsize(dest)//1024)
