# -*- coding: utf-8 -*-
"""Estrae il Prezzario Piemonte 2026 (listino + analisi prezzi) in JSON per la dashboard."""
import json, os, re, sys
from python_calamine import CalamineWorkbook

SRC = "/sessions/lucid-blissful-bell/mnt/LOMBARDIA Prezzario Regionale dei lavori pubblici - edizione 2026"
PIE = os.path.join(SRC, "01 PIEMONTE 26")
OUT = os.path.join(SRC, "dashboard", "data")
PFX = "PIE26_"
NSHARD = 96

def strip(c):
    c = str(c or "").strip()
    return c[len(PFX):] if c.startswith(PFX) else c

def num(v):
    try:
        return float(v)
    except (TypeError, ValueError):
        return None

def clean(s):
    return re.sub(r"\s*\n\s*", " | ", str(s or "").strip())

def djb2(s, n):
    h = 5381
    for ch in s:
        h = ((h * 33) ^ ord(ch)) & 0xFFFFFFFF
    return h % n

# ---- transcodifica 2025 -> 2026 (invertita) ----
wb = CalamineWorkbook.from_path(os.path.join(PIE, "tabella-transcodifica-2025-2026.xls"))
rows = wb.get_sheet_by_name(wb.sheet_names[0]).to_python(skip_empty_area=True)
old_by_new = {}
for r in rows:
    c25 = str(r[0] or "").replace("\x00", "").strip()
    c26 = str(r[2] or "").replace("\x00", "").strip() if len(r) > 2 else ""
    if re.match(r"^\d{2}\.", c25) and re.match(r"^\d{2}\.", c26):
        old_by_new.setdefault(c26, []).append(c25)
print("transcodifica: mappature", sum(len(v) for v in old_by_new.values()), "su", len(old_by_new), "codici 2026")

# ---- analisi: prima passata per sapere quali voci hanno l'analisi ----
wb = CalamineWorkbook.from_path(os.path.join(PIE, "AP_2026.xlsx"))
ap_rows = wb.get_sheet_by_name("Worksheet").to_python(skip_empty_area=True)
has_an = set()
for r in ap_rows[1:]:
    c0 = str(r[0] or "").strip()
    if c0.startswith(PFX):
        has_an.add(strip(c0))
print("voci con analisi:", len(has_an))

# ---- listino ----
wb = CalamineWorkbook.from_path(os.path.join(PIE, "Listino_2026.xlsx"))
rows = wb.get_sheet_by_name("Worksheet").to_python(skip_empty_area=True)
paths, path_idx = [], {}
voci = []
details = {}
listino_prezzi = {}
for r in rows[1:]:
    cod = strip(r[0])
    if not cod:
        continue
    tip, cap = clean(r[2]), clean(r[4])
    voce, art = clean(r[6]), clean(r[8])
    decl = (voce + (" — " + art if art else "")).strip()
    um = str(r[10] or "").strip()
    psg = num(r[11])   # prezzo senza SG e UI
    prezzo = num(r[12])
    scost = num(r[13])
    sgp, sg = num(r[14]), num(r[15])
    uip, ui = num(r[16]), num(r[17])
    inc, mano = num(r[18]), num(r[19])
    key = tip + "|" + cap if cap else tip
    if key not in path_idx:
        path_idx[key] = len(paths)
        paths.append(key)
    c25 = " | ".join(old_by_new.get(cod, [])[:4])
    voci.append([
        cod, decl, um,
        round(prezzo, 4) if prezzo else 0,
        round(inc * 100, 1) if inc else 0,
        c25,
        path_idx[key],
        None, None, None,   # niente elevata/modesta/sicurezza per PIE
        "",
        1 if cod in has_an else 0,
    ])
    listino_prezzi[cod] = (psg, prezzo)
    d = {
        "imp": round(psg, 4) if psg else None,
        "sgp": sgp, "sgv": round(sg, 4) if sg else None,
        "uip": uip, "uiv": round(ui, 4) if ui else None,
        "scost": round(scost * 100, 2) if scost is not None else None,
        "mano": round(mano, 4) if mano else None,
        "annc": clean(r[5]), "annv": clean(r[7]), "anna": clean(r[9]),
        "cam": str(r[1] or "").strip(),
    }
    details[cod] = {k: v for k, v in d.items() if v not in ("", None)}
print(f"listino: {len(voci)} voci, {len(paths)} percorsi")

with open(os.path.join(OUT, "voci_P.json"), "w", encoding="utf-8") as f:
    json.dump({"paths": paths, "voci": voci}, f, ensure_ascii=False, separators=(",", ":"))

# ---- analisi prezzi ----
rows = ap_rows
analisi = {}
cur, cat = None, ""
sgq = uiq = None
for r in rows[1:]:
    c0, c3 = str(r[0] or "").strip(), str(r[3] or "").strip()
    c5 = str(r[5] or "").strip()
    if c0.startswith(PFX):
        cur = strip(c0)
        cat = ""
        analisi[cur] = {"ris": [], "som": None, "sg": None, "ui": None, "tot": None,
                        "sgq": None, "uiq": None}
        continue
    if cur is None:
        continue
    if c3.startswith(PFX):
        qty, pu = num(r[8]), num(r[9])
        imp = round(qty * pu, 4) if (qty is not None and pu is not None) else None
        analisi[cur]["ris"].append([cat, strip(c3), clean(r[5]), str(r[7] or "").strip(), qty, pu, imp])
        continue
    if c3 and not c3.startswith("Totale"):
        cat = c3
        continue
    if c5.startswith("Spese generali"):
        analisi[cur]["sgq"] = num(r[8])
    elif c5.startswith("Utili"):
        analisi[cur]["uiq"] = num(r[8])

# calcola importi e verifica col listino
ok = ko = 0
for cod, an in analisi.items():
    som = round(sum(x[6] or 0 for x in an["ris"]), 4)
    sgq = an.pop("sgq") or 0.16
    uiq = an.pop("uiq") or 0.10
    sg = round(som * sgq, 4)
    ui = round((som + sg) * uiq, 4)
    an["som"], an["sg"], an["ui"] = som, sg, ui
    an["tot"] = round(som + sg + ui, 4)
    ref = listino_prezzi.get(cod)
    if ref and ref[1]:
        if abs(an["tot"] - ref[1]) <= max(0.01, ref[1] * 0.001):
            ok += 1
        else:
            ko += 1
            if ko <= 5:
                print(f"  SCOSTAMENTO {cod}: analisi {an['tot']} vs listino {ref[1]}")
print(f"analisi: {len(analisi)} blocchi | verifica prezzi: ok {ok}, scostamenti {ko}")

# ---- shard ----
os.makedirs(os.path.join(OUT, "det"), exist_ok=True)
shards = {}
matched = 0
for cod, det in details.items():
    an = analisi.get(cod)
    if an:
        det["an"] = an
        matched += 1
    key = f"p{djb2(cod, NSHARD):02d}"
    shards.s