# -*- coding: utf-8 -*-
"""Genera data/paniere.json per la card Monitoraggio prezzi materiali.
Fonte: RAFFRONTO/Serie_storica_PIE_2022-2026.xlsx (foglio Serie storica),
serie ufficiale prezzario Piemonte 2022-2026, agganciata per Codice 2025.
Integrazioni 2026 da 01 PIEMONTE 26/AP_2026.xlsx (analisi prezzi) per le voci
non piu a listino (gasolio, benzina).
"""
import json, os
from python_calamine import CalamineWorkbook

BASE = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
ROOT = os.path.dirname(BASE)
XLSX = os.path.join(ROOT, "RAFFRONTO", "Serie_storica_PIE_2022-2026.xlsx")
OUT  = os.path.join(BASE, "data", "paniere.json")

# Valori 2026 dalle ANALISI PREZZI 2026 (AP_2026.xlsx) per voci dismesse dal listino.
# codice 2025 -> (prezzo2026, codice risorsa 2026, nota)
OVERRIDE_2026 = {
    "01.P23.E80.005": (1.65, "PR195.G030.000.000", "Valore 2026 dalle analisi prezzi PIE 2026 (risorsa PR195.G030): voce non piu a listino."),
    "01.P23.E70.010": (1.55, "PR195.B020.000.000", "Valore 2026 dalle analisi prezzi PIE 2026 (risorsa PR195.B020): voce non piu a listino."),
    "01.P10.A50.005": (0.67, "PR015.B005.000.010", "Valore 2026 dalle analisi prezzi PIE 2026 (risorsa PR015.B005, bitume tal quale penetrazione 50/70): voce non piu a listino; prezzo componente non ufficiale."),
    "01.P10.A57.005": (0.72, "PR015.B010.000.000", "Valore 2026 dalle analisi prezzi PIE 2026 (risorsa PR015.B010, bitume modificato tipo hard 45/80-70): voce non piu a listino; prezzo componente non ufficiale."),
}

# Note descrittive per codice 2025 (senza modificare i valori)
NOTE_C25 = {
    "01.P25.A60.005": "Nel prezzario 2026 le voci ponteggi sono state ristrutturate: il prezzo per i primi 30 giorni cambia sensibilmente rispetto alle edizioni precedenti; il confronto storico non e direttamente omogeneo (verificare sulla pubblicazione ufficiale).",
}

PANIERE = [
("BITUMI","01.P10.A50.005","Bitume semisolido 50/70 per pavimentazioni stradali"),
("BITUMI","01.P10.A57.005","Bitume semisolido Hard 45/80-70"),
("BITUMI","01.P10.B00","Emulsione bituminosa"),
("BITUMI","01.P10.C70","Conglomerati bituminosi plastico freddo in sacchi"),
("BITUMI","01.P10.F00","Prodotti impermeabilizzanti"),
("BITUMI","30.A80","Conglomerati bituminosi CAM"),
("BITUMI","01.P10.H59","Guaina impermeabile barriera al vapore bituminosa"),
("CEMENTI e CLS","01.P02.A90","Calce in sacchi"),
("CEMENTI e CLS","01.P02.A80","Calce in zolle"),
("CEMENTI e CLS","01.P02.A05","Cementi comuni conformi alla norma UNI EN 197-1"),
("DISCARICHE","29.P15.A40.010","Rifiuti C&D contenenti sostanze pericolose (CER 17 09 03)"),
("DISCARICHE","29.P15.A35.010","Materiali a base di gesso (CER 17 08 02)"),
("DISCARICHE","29.P15.A30.005","Materiali isolanti contenenti amianto (CER 17 06 01)"),
("DISCARICHE","29.P15.A30.015","Materiali isolanti diversi (CER 17 06 04)"),
("DISCARICHE","29.P20.A10.005","Rifiuti biodegradabili (CER 20 02 01)"),
("DISCARICHE","29.P15.A40.015","Rifiuti misti C&D (CER 17 09 04)"),
("DISCARICHE","29.P15.A25.010","Terra e rocce conformi col. A (CER 17 05 04)"),
("ELEMENTI_CLS","01.P05.A20","Blocchi forati in cls REI 120 - 20x20x50"),
("ELEMENTI_CLS","01.P05.A25","Blocchi forati in cls REI 180 - 25x20x50"),
("ELEMENTI_CLS","01.P05.A10","Blocchi forati in cls REI 90 - 12x20x50"),
("ELEMENTI_CLS","01.P05.C30","Condotto prefabbricato sezione ovoidale in cls vibrocompresso"),
("ELEMENTI_CLS","01.P05.B50","Cordoli in cemento pressato retti o curvi"),
("ELEMENTI_CLS","01.P05.B70","Elemento dissuasore di traffico in cls C30/37"),
("ELEMENTI_CLS","30.P51.A00","Barriere tipo new jersey in c.a. prefabbricato C35/45"),
("ELEMENTI_CLS","01.P05.F00","Travetti prefabbricati in c.a.p. sezione a T rovesciata"),
("ELEMENTI_CLS","01.P05.C25","Tubi autoportanti in cls vibrocompresso (1,00 kN)"),
("ELEMENTI_CLS","01.P05.C40","Tubi autoportanti in cls vibrocompresso (1,30 kN)"),
("ELEMENTI_CLS","01.P05.C20","Tubi in conglomerato cementizio vibrocompressi"),
("FERRO ACCIAIO","01.P12.A05","Acciaio laminato a caldo B450A/B450C"),
("FERRO ACCIAIO","01.P12.A07","Acciaio trafilato a freddo B450A/B450C"),
("FERRO ACCIAIO","01.P12.C00","Barre in acciaio ad aderenza migliorata B450A/B450C"),
("FERRO ACCIAIO","18.P03.B41","Barriera di sicurezza acciaio Corten - bordo laterale rilevato"),
("FERRO ACCIAIO","18.P03.B43","Barriera di sicurezza acciaio Corten - bordo ponte"),
("FERRO ACCIAIO","01.P13.E50","Ghisa grigia per chiusini carreggiabili tipo Torino"),
("FERRO ACCIAIO","01.P13.E55","Ghisa per bocchette apribili con suggello incernierato"),
("FERRO ACCIAIO","01.P13.E62","Ghisa sferoidale UNI EN 124 per griglie e chiusini"),
("FERRO ACCIAIO","01.P12.M30","Rete elettrosaldata"),
("FERRO ACCIAIO","01.P13.L30","Barriera stradale acciaio"),
("FERRO ACCIAIO","01.P13.L20","Barriera stradale di sicurezza (guard-rail) marcata CE"),
("INERTI","01.P03.B20","Misto frantumato stabilizzato di cava 0/30"),
("INERTI","01.P03.A10","Aggregati non frantumati per cls UNI EN 12620"),
("INERTI","01.P03.B30","Ciottoli serpentinosi/silicei per gabbioni e drenaggi"),
("INERTI","01.P03.B50","Ciottoli serpentinosi per ciottolati"),
("INERTI","01.P03.A80","Ghiaia e ghiaietto vagliati e lavati"),
("INERTI","01.P18.N35","Guide curve di granito"),
("INERTI","01.P03.C80","Massi di cava di natura granitica per difese idrauliche"),
("INERTI","01.P03.C60","Misto granulare di cava o di fiume"),
("INERTI","01.P03.B10","Pisello lavato"),
("INERTI","01.P03.A10.005","Sabbia granita"),
("INERTI","01.P03.A50","Sabbia fine"),
("INERTI","30.P05.A30.005","Sabbia vagliata (riciclo)"),
("INERTI","30.P05.A70","Cocciopesto"),
("INERTI","01.P03.D10","Pozzolana"),
("INERTI","01.P03.A10.025","Pietrisco"),
("INERTI","01.P03.A10.015","Pietrischetto"),
("ISOLANTI","30.P50.G05","Lana di roccia in pannelli semirigidi"),
("ISOLANTI","30.P50.A00","Pannelli EPS"),
("ISOLANTI","30.P50.B00","Pannello in polistirene espanso estruso (XPS)"),
("LATERIZI","01.P04.F50","Blocchi in laterizio alleggerito porizzato portante"),
("LATERIZI","30.P20.B00","Blocchi in laterizio portanti antisismici"),
("LATERIZI","01.P04.F10","Blocchi semipieni portanti alte prestazioni termiche"),
("LATERIZI","01.P04.F02","Laterizio alveolato termoacustico f180"),
("LATERIZI","01.P04.B70","Mattoni copriferro"),
("LATERIZI","01.P04.A10","Mattoni faccia-vista sabbiati"),
("LATERIZI","01.P04.A30","Mattoni forati"),
("LATERIZI","01.P04.B30","Mattoni multifori 6x12x24 faccia a vista"),
("LATERIZI","01.P04.A80","Mattoni pieni a paramento 6x12x24"),
("LATERIZI","01.P04.A60","Mattoni pieni comuni 6x12x24"),
("LATERIZI","01.P04.A20","Mattoni semipieni"),
("LATERIZI","01.P04.B80","Tavelle forate perret"),
("LATERIZI","01.P04.C20","Tavelle spaccabili a doppia parete"),
("LATERIZI","01.P04.C30","Tavelloni forati 6x25"),
("LATERIZI","01.P07.C00","Tegole in gres"),
("LEGNAMI","01.P15.C10","Compensato mm 4 di spessore"),
("LEGNAMI","01.P15.A25","Pannelli di abete per armatura"),
("LEGNAMI","01.P15.C15","Multistrato in pioppo"),
("LEGNAMI","01.P15.A20","Tavole di abete per armatura e ponteggi"),
("PAVIMENTI","01.P11.C95","Elementi trafilati in klinker ceramico per piscine"),
("PAVIMENTI","01.P11.C40","Lastre in cls multistrato vibrocompresso UNI EN 1339"),
("PAVIMENTI","01.P11.E10","Linoleum su juta con trattamento protettivo"),
("PAVIMENTI","01.P11.B44","Marmette autobloccanti forate in cls"),
("PAVIMENTI","01.P11.C35","Masselli in cls multistrato UNI EN 1338"),
("PAVIMENTI","01.P07.B10","Mattonelle in gres"),
("PAVIMENTI","01.P11.E55","Pavimento vinilico omogeneo marmorizzato"),
("PAVIMENTI","01.P07.B30","Piastrelle in gres"),
("PAVIMENTI","01.P07.B40","Piastrelle in gres ceramico"),
("PAVIMENTI","01.P16.A00","Prefinito per palchetto sp. 10 mm"),
("PAVIMENTI","01.P16.A01","Prefinito per palchetto sp. 14 mm"),
("PAVIMENTI","01.P16.A02","Prefinito per palchetto sp. 15 mm"),
("SERRAMENTI","01.P20.I00","Telaio serramenti esterni in alluminio a taglio termico"),
("SERRAMENTI","01.P20.L00","Telaio serramenti esterni in legno"),
("TUBAZIONI","01.P08.A26","Curve a 15 gradi in PVC rigido UNI EN 1401"),
("TUBAZIONI","01.P08.A25","Curve a 45 gradi in PVC rigido UNI EN 1329"),
("TUBAZIONI","01.P08.A24","Curve a 45 gradi in PVC serie normale"),
("TUBAZIONI","07.P06.G05","Tubazione PE100 per acqua potabile UNI EN 12201-2"),
("TUBAZIONI","30.P35.L10","Tubi multistrato PEAD riciclato per scarichi"),
("TUBAZIONI","01.P08.A13","Tubo PVC termoresistente 95 gradi - 1 m"),
("TUBAZIONI","01.P08.A15","Tubo PVC termoresistente 95 gradi - 2 m"),
("TUBAZIONI","01.P08.A18","Tubo PVC termoresistente 95 gradi - 3 m"),
("TUBAZIONI","01.P08.A19","Tubo PVC rigido scarichi non a pressione UNI EN 1401 - 6 m"),
("COMBUSTIBILI ED ENERGIA","01.P23.E80","Gasolio"),
("COMBUSTIBILI ED ENERGIA","01.P23.E70","Benzina"),
("COMBUSTIBILI ED ENERGIA","01.P23.F15.005","Gas propano (in bombole)"),
("COMBUSTIBILI ED ENERGIA","01.P99.D45.005","Energia elettrica"),
("COMBUSTIBILI ED ENERGIA","01.P99.D01.010","Acqua"),
("LEGANTI","01.P02.A05.005","Cemento sfuso 32,5"),
("LEGANTI","01.P02.A05.010","Cemento in sacchi 32,5"),
("LEGANTI","01.P02.A05.015","Cemento sfuso 42,5"),
("LEGANTI","01.P02.A05.020","Cemento in sacchi 42,5"),
("LEGANTI","01.P02.A05.045","Cemento sfuso 52,5"),
("LEGANTI","01.P02.A05.050","Cemento in sacchi 52,5"),
("LEGANTI","01.P02.B30.005","Calce spenta"),
("PONTEGGI","01.P25.A60.005","Ponteggio tubolare esterno tubo-giunto - primi 30 giorni"),
]

def num(x):
    if x is None or x == "": return None
    try:
        return round(float(x), 4)
    except (TypeError, ValueError):
        return None

def main():
    wb = CalamineWorkbook.from_path(XLSX)
    rows = wb.get_sheet_by_name("Serie storica").to_python()
    idx = {}
    for r in rows[1:]:
        c25 = str(r[1]).strip() if r[1] is not None else ""
        if c25:
            idx.setdefault(c25, []).append(r)
    anni = ["mar 2022", "lug 2022", "2023", "2024", "2025", "2026"]
    out_voci = []
    cats_order = []
    cov = {"ok": 0, "parz": 0, "ass": 0}
    for cat, code, label in PANIERE:
        if cat not in cats_order: cats_order.append(cat)
        matched = []
        for c25, rs in idx.items():
            if c25 == code or c25.startswith(code + "."):
                matched.extend(rs)
        items = []
        nota = ""
        for r in matched:
            serie = [num(r[i]) for i in range(5, 11)]
            c25v = str(r[1]).strip()
            c26v = str(r[2]).strip() if r[2] is not None else ""
            dsv = str(r[3]).strip() if r[3] is not None else ""
            if c25v in OVERRIDE_2026 and serie[5] is None:
                val, code26, n = OVERRIDE_2026[c25v]
                serie[5] = val; c26v = code26; nota = n
                dsv = (dsv + " - 2026 da analisi prezzi PIE").strip(" -")
            if all(v is None for v in serie): continue
            items.append({"c25": c25v, "c26": c26v, "ds": dsv,
                          "um": (str(r[4]).strip() if r[4] is not None else ""), "s": serie})
        items.sort(key=lambda x: x["c25"])
        # Monitoraggio: tenere solo le voci con prezzo nell'edizione 2026 vigente
        items = [it for it in items if it["s"][5] is not None]
        if not items: continue
        if not nota and code in NOTE_C25: nota = NOTE_C25[code]
        cov["ok"] += 1
        rec = {"cat": cat, "code": code, "label": label, "stato": "ok", "items": items}
        if nota: rec["nota"] = nota
        out_voci.append(rec)
    data = {"anni": anni, "cats": [c for c in cats_order if any(v["cat"]==c for v in out_voci)],
        "fonte": "Prezzario Regione Piemonte - serie storica ufficiale 2022-2026 (mar 2022, lug 2022 straord., 2023, 2024, 2025, 2026), ricostruita con transcodifiche ufficiali. Integrazioni 2026 da analisi prezzi PIE 2026 dove la voce non e piu a listino. Elaborazione ANCE Piemonte Valle d'Aosta.",
        "voci": out_voci}
    with open(OUT, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, separators=(",", ":"))
    print("Materiali:", len(out_voci), "| ok:", cov["ok"], "parziali:", cov["parz"], "assenti:", cov["ass"])
    print("Varianti:", sum(len(v["items"]) for v in out_voci), "| KB:", round(os.path.getsize(OUT)/1024,1))

if __name__ == "__main__":
    main()
