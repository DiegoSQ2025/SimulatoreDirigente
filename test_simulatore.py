#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Protocollo di verifica obbligatorio (AGENTS.md) — dirigente CCNL Terziario.

1. Replica Python (addizionale_lib + regole dirigente 2026) vs motore 'formulas'
   sull'xlsx reale: confronto cella per cella, tolleranza 0,01 euro.
2. Scan errori su tutte le celle calcolate.
3. Casi limite addizionali MEF.
4. Benchmark cedolino reale 01 (dirigente liv. 5, Lazio/Roma, 2004->2026):
   comp 8.291,38, trattenute 4.107,91, netto 4.184,00.
5. Benchmark cedolino reale 03 (dirigente 03, stesso datore, RME 12.857,14):
   comp 13.612,95, base IVS arrotondata 13.613,00, trattenute 6.817,53, netto 6.796,00.
"""
import os
import shutil
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import addizionale_lib as ADL
from openpyxl import load_workbook

SRC = os.path.join(HERE, "Simulatore_Dirigente_CCNL_Commercio.xlsx")
TMP = "/tmp/test_dir.xlsx"
PYBOOK = "[TEST_DIR.XLSX]"
TOLL = 0.01
ERRORS = {"#N/A", "#VALUE!", "#REF!", "#NAME?", "#DIV/0!", "#NUM!", "#NULL!"}

MINIMI = {"2026": 4660.0, "2027": 4920.0, "2028": 5140.0}

MENS = 14
NG = 26
INPS = 0.0919
AGG_SOG = 56224.0
AGG_ALQ = 0.01
MAXC = 122295.0
MN_BASE, MN_D = 59224.54, 0.02
FD_BASE, FD_D = 45940.0, 0.0187
AP_ANN = 464.81
AP_M = round(AP_ANN / 12, 2)
MNF_DIP = 10.83


def r2(x):
    return round(x + 1e-12, 2)


def scaglioni(imp):
    if imp <= 28000:
        return imp * 0.23
    if imp <= 50000:
        return 28000 * 0.23 + (imp - 28000) * 0.33
    return 28000 * 0.23 + 22000 * 0.33 + (imp - 50000) * 0.43


def detrazione13(rld):
    if rld <= 15000.0:
        ded = 1955.0
    elif rld <= 28000.0:
        ded = 1910.0 + 1190.0 * (28000.0 - rld) / 13000.0
    elif rld <= 50000.0:
        ded = 1910.0
    elif rld <= 75000.0:
        ded = 1910.0 * (75000.0 - rld) / 25000.0
    else:
        ded = 0.0
    return ded


def py_calc(anno=2026, supc=1317.48, smm=350.0, supi=1656.80572, scatti=0.0,
            ggRet=25, ggFer=1, ggFest=1, fringe=136.4, fringeCash=0.0, ov=2755.77,
            regRes=253.95, comRes=52.97, acconto=27.33, arrotT=0.03, arrotN=0.53,
            oneri=0.0, regione="Lazio", comune="Roma"):
    """Regole dirigente 2026 riconciliate coi cedolini reali 01 e 03 (Lazio/Roma)."""
    minimo = MINIMI[str(anno)]
    RME = minimo + supc + smm + supi + scatti
    giorn = RME / NG
    comp = r2(ggRet * giorn) + r2(ggFer * giorn) + r2(ggFest * giorn) + r2(fringeCash)
    base_m = float(int(comp + 0.5))   # il gestionale ARROTONDA all'euro (ced.03: 13.612,95 -> 13.613,00)
    ivsM = r2(base_m * INPS)
    aggM = r2(max(0.0, base_m - AGG_SOG / 12) * AGG_ALQ)
    negriM = r2(MN_BASE / 12 * MN_D)
    fasdacM = r2(FD_BASE / 12 * FD_D)
    pastM = AP_M
    formM = MNF_DIP
    RAL = RME * MENS
    ralBase = (minimo + supc + smm) * MENS
    supiA = supi * MENS
    ivsA = min(RAL, MAXC) * INPS
    aggA = max(0.0, min(RAL, MAXC) - AGG_SOG) * AGG_ALQ   # Circ. INPS 6/2026: 1% solo fino al massimale 122.295
    negriA = MN_BASE * MN_D
    fasdacA = FD_BASE * FD_D
    rld = RAL - ivsA - aggA - negriA
    imp = rld + (fringe + fringeCash) * 12
    irpl = scaglioni(imp)
    ded = detrazione13(rld)
    dede = min(ded, irpl)
    irpn = max(0.0, irpl - dede - oneri)
    regA = ADL.addizionale_regionale(imp, regione)
    comA = ADL.addizionale_comunale(imp, regione, comune)
    # trattenuta IRPEF in busta = NETTA: lorda annua/14 meno detrazione art. 13 annua/14
    # (override = importo gia' trattenuto dal cedolino reale; nei cedolini campione detrazione 0, RLD > 75.000)
    irplM = ov if ov > 0 else r2(irpl / MENS)
    dedM = 0.0 if ov > 0 else r2(dede / MENS)
    irpm = ov if ov > 0 else r2(irplM - dedM)
    regM = regRes if regRes > 0 else r2(regA / 12)
    comM = comRes if comRes > 0 else r2(comA / 12)
    tratt = ivsM + aggM + negriM + fasdacM + pastM + formM + irpm + regM + comM + acconto + arrotT
    nettoM = comp - tratt + arrotN
    nettoA = RAL - (ivsA + aggA + negriA + fasdacA + AP_ANN + formM * 12 + irpn + regA + comA)
    return {"RME": RME, "giorn": giorn, "comp": comp, "ivsM": ivsM, "aggM": aggM,
            "negriM": negriM, "fasdacM": fasdacM, "pastM": pastM, "formM": formM,
            "irplM": irplM, "dedM": dedM, "irpm": irpm, "regM": regM, "comM": comM,
            "tratt": tratt, "nettoM": nettoM,
            "RAL": RAL, "ralBase": ralBase, "supiA": supiA,
            "ivsA": ivsA, "aggA": aggA, "negriA": negriA, "fasdacA": fasdacA,
            "rld": rld, "imp": imp, "irpl": irpl, "ded": ded, "dede": dede, "irpn": irpn,
            "regA": regA, "comA": comA, "nettoA": nettoA, "nettoM14": nettoA / MENS}


IN_LABELS = {
    "anno": "Anno di paga",
    "supc": "SUP.CONTR \u2014 superminimo da contratto",
    "smm": "SMM 2016 (mensile)",
    "supi": "S.MINIMO \u2014 superminimo INDIVIDUALE personale",
    "scatti": "Scatti pre-2004 (mensile)",
    "ggRet": "Giorni retribuiti",
    "ggFer": "Giorni di ferie godute",
    "ggFest": "Giorni festivita'/permessi",
    "fringe": "Fringe benefit (mensile",
    "fringeCash": "Indennita' denaro (mensile",
    "ov": "IRPEF lorda del mese \u2014 override",
    "regRes": "Add. regionale del mese \u2014 residuo",
    "comRes": "Add. comunale del mese \u2014 residuo",
    "acconto": "Acconto addizionali extra",
    "arrotT": "Arrotondamento trattenute (cedolino)",
    "arrotN": "Arrotondamento netto (cedolino)",
    "oneri": "Altri oneri/detrazioni annuali",
}

REG_LABELS = {
    "regione": "Regione (menu a discesa)",
    "comune": "Comune (menu a discesa)",
}

OUT_LABELS = {
    "RME": "RME \u2014 retribuzione mensile effettiva",
    "giorn": "Retribuzione giornaliera (RME/26",
    "comp": "TOTALE COMPETENZE",
    "ivsM": "\u2212 Contributo IVS 9,19%",
    "aggM": "\u2212 Contributo IVS agg. 1%",
    "negriM": "\u2212 Fondo Mario Negri 2% (base",
    "fasdacM": "\u2212 FASDAC 1,87% (base",
    "pastM": "\u2212 A. Pastore quota dirigente",
    "formM": "\u2212 Fondo formazione Mario Negri",
    "irplM": "IRPEF lorda del mese (annua/14",
    "dedM": "Detrazione lavoro dip. art. 13 del mese",
    "irpm": "IRPEF netta del mese",
    "regM": "\u2212 Addizionale regionale del mese",
    "comM": "\u2212 Addizionale comunale del mese",
    "tratt": "TOTALE TRATTENUTE",
    "nettoM": "NETTO DEL MESE",
    "RAL": "RAL \u2014 retribuzione annua effettiva",
    "ralBase": "Retr. base di contratto annua CCNL",
    "supiA": "Superminimo individuale annuo",
    "ivsA": "IVS 9,19% annua",
    "aggA": "IVS agg. 1% annua",
    "negriA": "Fondo Mario Negri 2% annuo",
    "fasdacA": "FASDAC 1,87% annuo",
    "rld": "RLD \u2014 reddito di riferimento",
    "imp": "Imponibile IRPEF annuo",
    "irpl": "IRPEF lorda annua",
    "ded": "Detrazione lavoro dipendente (art. 13 TUIR)",
    "dede": "Detrazione effettiva (capienza imposta",
    "irpn": "IRPEF NETTA annua",
    "regA": "Addizionale regionale annua",
    "comA": "Addizionale comunale annua",
    "nettoA": "NETTO ANNUO in busta",
}


# ---------------------------------------------------------------- righe/etichette
def row_of(ws, prefix, sheet):
    hits = [c.row for (c,) in ws.iter_rows(min_col=1, max_col=1)
            if isinstance(c.value, str) and c.value.strip().startswith(prefix)]
    assert len(hits) == 1, f"etichetta ambigua o assente in {sheet}: {prefix!r} -> {hits}"
    return hits[0]


def find_rows():
    wb = load_workbook(SRC)
    rows = {}
    for k, prefix in IN_LABELS.items():
        sheet = "Simulazione"
        rows[k] = (sheet, row_of(wb[sheet], prefix, sheet))
    for k, prefix in REG_LABELS.items():
        rows[k] = ("Parametri", row_of(wb["Parametri"], prefix, "Parametri"))
    for k, prefix in OUT_LABELS.items():
        rows[k] = ("Simulazione", row_of(wb["Simulazione"], prefix, "Simulazione"))
    wb.close()
    return rows


# ---------------------------------------------------------------- motore xlsx
from formulas import ExcelModel

SHEETS = {"Simulazione": "SIMULAZIONE", "Parametri": "PARAMETRI"}


def scrub(path):
    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.data_validations.dataValidation = []
    try:
        wb.defined_names.clear()
    except Exception:
        pass
    wb.save(path)


def build_model():
    shutil.copy(SRC, TMP)
    scrub(TMP)
    xl = ExcelModel().loads(TMP).finish()
    sol0 = xl.calculate()
    keys = {k.upper(): k for k in sol0}
    return xl, keys


def xls_calc(case, rows, model):
    xl, keys = model

    def key(sheet, rowno):
        up = f"'{PYBOOK}{SHEETS[sheet]}'!B{rowno}".upper()
        assert up in keys, f"cella non nel modello: {up}"
        return keys[up]

    defaults = dict(anno=2026, supc=1317.48, smm=350.0, supi=1656.80572, scatti=0.0,
                    ggRet=25, ggFer=1, ggFest=1, fringe=136.4, fringeCash=0.0, ov=2755.77,
                    regRes=253.95, comRes=52.97, acconto=27.33, arrotT=0.03,
                    arrotN=0.53, oneri=0.0, regione="Lazio", comune="Roma")
    defaults.update(case)
    defaults["anno"] = int(case.get("anno", 2026))
    inputs = {key(sheet, rowno): defaults[k] for k, (sheet, rowno) in rows.items()
              if k in IN_LABELS or k in REG_LABELS}
    sol = xl.calculate(inputs=inputs)

    def val(k):
        v = sol[key(*rows[k])]
        try:
            return float(v.value[0, 0])
        except Exception:
            return float(v.value)

    out = {k: val(k) for k in OUT_LABELS}
    bad = []
    for k, v in sol.items():
        sval = getattr(v, "value", v)
        try:
            sval = sval[0, 0]
        except Exception:
            pass
        if isinstance(sval, str) and sval.strip() in ERRORS:
            bad.append((k, sval))
    return out, bad


CASES = [
    dict(nome="preset cedolino reale (default)", case={}),
    dict(nome="mese pulito (override a 0)", case=dict(ov=0, regRes=0, comRes=0, acconto=0,
                                                      arrotT=0, arrotN=0)),
    dict(nome="mese pieno 26gg auto", case=dict(ov=0, regRes=0, comRes=0, acconto=0,
                                                arrotT=0, arrotN=0, ggRet=26, ggFer=0, ggFest=0)),
    dict(nome="solo ferie 3gg", case=dict(ggRet=23, ggFer=3)),
    dict(nome="fringe 300", case=dict(fringe=300.0)),
    dict(nome="Lombardia/Milano", case=dict(regione="Lombardia", comune="Milano")),
    dict(nome="anno 2027 (minimo 4.920)", case=dict(anno=2027)),
    dict(nome="anno 2028 + oneri conguaglio", case=dict(anno=2028, oneri=530.0)),
    dict(nome="scatti pre-2004 +258,22", case=dict(scatti=258.22)),
    dict(nome="superminimi Others", case=dict(supc=1500.0, supi=2000.0, smm=0.0)),
    dict(nome="superminimo ad personam 1.000/mese (mese pieno, pulito)", case=dict(
        supi=1000.0, ggRet=26, ggFer=0, ggFest=0, fringe=0.0, ov=0, regRes=0,
        comRes=0, acconto=0, arrotT=0, arrotN=0)),
    dict(nome="superminimo ad personam 1.500/mese (mese pieno, pulito)", case=dict(
        supi=1500.0, ggRet=26, ggFer=0, ggFest=0, fringe=0.0, ov=0, regRes=0,
        comRes=0, acconto=0, arrotT=0, arrotN=0)),
    dict(nome="superminimo ad personam 2.500/mese (mese pieno, pulito)", case=dict(
        supi=2500.0, ggRet=26, ggFer=0, ggFest=0, fringe=0.0, ov=0, regRes=0,
        comRes=0, acconto=0, arrotT=0, arrotN=0)),
    dict(nome="fringe benefit 100/mese x 12 (mese pieno, pulito)", case=dict(
        fringe=100.0, ggRet=26, ggFer=0, ggFest=0, ov=0, regRes=0,
        comRes=0, acconto=0, arrotT=0, arrotN=0)),
    dict(nome="cedolino 03 (dirigente 03: RME 12.857, autovettura)", case=dict(
        supc=0.0, supi=7847.14, ggRet=26, ggFer=0, ggFest=1, fringe=0.0, fringeCash=261.30,
        ov=4810.85, regRes=345.82, comRes=64.12, acconto=36.43, arrotT=0.13, arrotN=0.58)),
]

SOGLIE = [
    ("Lazio imp=28000 esatto flat", ADL.addizionale_regionale(28000, "Lazio"), 28000 * 0.0173),
    ("Lazio imp=28000.01 scaglioni + detrazione", ADL.addizionale_regionale(28000.01, "Lazio"),
     15000 * 0.0173 + 13000.01 * 0.0333 - 60),
    ("Lazio imp=30000 detrazione inclusa", ADL.addizionale_regionale(30000, "Lazio"),
     15000 * 0.0173 + 15000 * 0.0333 - 60),
    ("Lazio imp=30000.01 senza detrazione", ADL.addizionale_regionale(30000.01, "Lazio"),
     15000 * 0.0173 + 15000.01 * 0.0333),
    ("Roma alla soglia esenzione", ADL.addizionale_comunale(14000, "Lazio", "Roma"), 0),
    ("Roma primo euro sopra esenzione", ADL.addizionale_comunale(14000.01, "Lazio", "Roma"),
     14000.01 * 0.009),
    ("Milano unica 0,8% sopra esenzione 23.000", ADL.addizionale_comunale(40000, "Lombardia", "Milano"),
     40000 * 0.008),
    ("Abruzzo tre scaglioni", ADL.addizionale_regionale(120000, "Abruzzo"),
     28000 * 0.0167 + 22000 * 0.0287 + 70000 * 0.0333),
    ("FVG 20k", ADL.addizionale_regionale(20000, "Friuli-Venezia Giulia"), 20000 * 0.0123),
    ("VdA esente fino a 15.000", ADL.addizionale_regionale(15000, "Valle d'Aosta"), 0),
    ("Trento scaglione 2", ADL.addizionale_regionale(100000, "Trentino-Alto Adige (Trento)"),
     0.0123 * 50000 + 0.0173 * 20000),
    ("Bolzano con detrazione a 40k", round(ADL.addizionale_regionale(40000, "Trentino-Alto Adige (Bolzano)"), 2),
     round(40000 * 0.0123 - 430.50, 2)),
]

CED01 = [
    ("ced.01: RME 7.984,29", "RME", 7984.29),
    ("ced.01: retr. giornaliera 307,08791", "giorn", 307.08791),
    ("ced.01: TOTALE COMPETENZE 8.291,38", "comp", 8291.38),
    ("ced.01: IVS 9,19% 761,94", "ivsM", 761.94),
    ("ced.01: IVS agg. 1% 36,06", "aggM", 36.06),
    ("ced.01: Fondo Mario Negri 98,71", "negriM", 98.71),
    ("ced.01: FASDAC 71,59", "fasdacM", 71.59),
    ("ced.01: A. Pastore 38,73", "pastM", 38.73),
    ("ced.01: formazione 10,83", "formM", 10.83),
    ("ced.01: IRPEF lorda 2.755,77 (override)", "irpm", 2755.77),
    ("ced.01: add. reg. residuo 253,95", "regM", 253.95),
    ("ced.01: add. com. residuo 52,97", "comM", 52.97),
    ("ced.01: TOTALE TRATTENUTE 4.107,91", "tratt", 4107.91),
    ("ced.01: NETTO DEL MESE 4.184,00", "nettoM", 4184.00),
    ("ced.01: RAL 111.780,00", "RAL", 111780.00),
    ("ced.01: RAL base di contratto 88.584,72", "ralBase", 88584.72),
    ("ced.01: superminimo individuale annuo 23.195,28", "supiA", 23195.28),
    ("ced.01: Detrazione art. 13 = 0 (RLD > 75k)", "ded", 0.0),
    ("ced.01: addizionale regionale annua Lazio", "regA", 3136.76),
    ("ced.01: addizionale comunale annua Roma", "comA", 912.64),
]

D03 = dict(anno=2026, supc=0.0, supi=7847.14, ggRet=26, ggFer=0, ggFest=1,
           fringe=0.0, fringeCash=261.30, ov=4810.85, regRes=345.82, comRes=64.12,
           acconto=36.43, arrotT=0.14, arrotN=0.58)   # arrotT 0,14 = arrotondamento 0,13 del cedolino + 0,01 di arrotondamento IVS del gestionale (vedi README)
CED03 = [
    ("ced.03: RME 12.857,14", "RME", 12857.14),
    ("ced.03: retr. giornaliera 494,50538", "giorn", 494.50538),
    ("ced.03: TOTALE COMPETENZE 13.612,95 (26gg + fest. + autovettura)", "comp", 13612.95),
    ("ced.03: base IVS arrotondata 13.613,00 -> IVS 1.251,03 (cedolino 1.251,04, scarto 0,01)", "ivsM", 1251.03),
    ("ced.03: IVS agg. 1% 89,28", "aggM", 89.28),
    ("ced.03: Fondo Mario Negri 98,71", "negriM", 98.71),
    ("ced.03: FASDAC 71,59", "fasdacM", 71.59),
    ("ced.03: A. Pastore 38,73", "pastM", 38.73),
    ("ced.03: formazione 10,83", "formM", 10.83),
    ("ced.03: IRPEF lorda 4.810,85 (override)", "irpm", 4810.85),
    ("ced.03: add. reg. 345,82", "regM", 345.82),
    ("ced.03: add. com. 64,12", "comM", 64.12),
    ("ced.03: TOTALE TRATTENUTE 6.817,53", "tratt", 6817.53),
    ("ced.03: NETTO DEL MESE 6.796,00", "nettoM", 6796.00),
    ("ced.03: RAL 179.999,96", "RAL", 179999.96),
    ("ced.03: RAL base di contratto 5.010 x 14 = 70.140,00", "ralBase", 70140.00),
    ("ced.03: superminimo individuale annuo 109.859,96", "supiA", 109859.96),
    ("ced.03: imponibile IRPEF annuo (RLD + autovettura x 12) = 170.051,45 (1% IVS capato al massimale)", "imp", 170051.45),
]

# Terzo benchmark: elaborazione di controllo di un gestionale paghe (gennaio 2027,
# visuale annua, superminimo retro-ingegnerizzato per netto annuo target 44.000,00).
# Stesse regole condivise, struttura ridotta: SENZA FASDAC, SENZA Pastore, SENZA
# detrazione art. 13, formazione non annualizzata (10,83 in totale).
D_SIM27 = dict(anno=2027, supc=0.0, smm=0.0, supi=550.11117, ggRet=26, ggFer=0,
               ggFest=0, fringe=0.0, fringeCash=0.0, ov=0.0, regRes=0.0, comRes=0.0,
               acconto=0.0, arrotT=0.0, arrotN=0.0, oneri=0.0,
               regione="Lazio", comune="Roma")
SIM27 = [
    ("sim27: RME 5.470,11 (4.920 + 550,11117)", "RME", 5470.11),
    ("sim27: RAL 76.581,56 = rate x 14 (allegato 76.581,55)", "RAL", 76581.56),
    ("sim27: RAL base contratto 4.920 x 14 = 68.880,00", "ralBase", 68880.00),
    ("sim27: superminimo ad personam annuo 7.701,56", "supiA", 7701.56),
    ("sim27: IVS 9,19% annua 7.037,85 (allegato 7.037,89 su base annua arrotondata)", "ivsA", 7037.85),
    ("sim27: IVS agg. 1% annua 203,58 (76.582 - 56.224)", "aggA", 203.58),
    ("sim27: Fondo Mario Negri 2% annuo 1.184,49", "negriA", 1184.49),
    ("sim27: RLD 68.155,64 (allegato 68.144,76: deduce anche formazione 10,83)", "rld", 68155.64),
    ("sim27: IRPEF lorda annua 21.506,93 (scaglioni 23/33/43)", "irpl", 21506.93),
    ("sim27: detrazione art. 13 522,91 (RLD < 75.000: l'allegato NON la applica)", "ded", 522.91),
    ("sim27: IRPEF netta annua 20.984,02", "irpn", 20984.02),
    ("sim27: add. regionale annua 2.029,58 (regola marg Lazio)", "regA", 2029.58),
    ("sim27: add. comunale annua 613,40 (Roma 0,90%)", "comA", 613.40),
    ("sim27: netto annuo modello completo 43.074,79 (allegato 44.000 con struttura ridotta)", "nettoA", 43074.79),
]
# Verifiche sulla STRUTTURA dell'allegato (stesso imponibile, sue omissioni):
# riproduce esattamente i suoi numeri con le nostre formule condivise.
SIM27_ALLEGATO = [
    ("sim27 allegato: IRPEF lorda 23/33/43 su imponibile 68.144,76 = 21.502,25",
     scaglioni(68144.76), 21502.25),
    ("sim27 allegato: add. regionale marg Lazio su 68.144,76 = 2.029,22",
     ADL.addizionale_regionale(68144.76, "Lazio"), 2029.22),
    ("sim27 allegato: add. comunale Roma 0,90% su 68.144,76 = 613,30",
     ADL.addizionale_comunale(68144.76, "Lazio", "Roma"), 613.30),
    ("sim27 allegato: IVS 9,19% su base annua arrotondata 76.582 = 7.037,89",
     round(76582 * 0.0919, 2), 7037.89),
    ("sim27 allegato: IVS agg. 1% su 76.582 - 56.224 = 203,58",
     round((76582 - 56224) * 0.01, 2), 203.58),
    ("sim27 allegato: TFR quota anno 76.581,55 / 13,5 = 5.672,71",
     round(76581.55 / 13.5, 2), 5672.71),
]
_NETTO_ALLEGATO = round(76581.55638 - (7037.89 + 203.58 + 1184.49 + 10.83
                                       + 21502.25 + 2029.22 + 613.30), 2)
SIM27_ALLEGATO.append(
    ("sim27 allegato: TOTALE TRATTENUTE sua struttura -> netto 44.000,00",
     _NETTO_ALLEGATO, 44000.00))


def main():
    rows = find_rows()
    model = build_model()
    fails, tot_checked = [], 0
    for entry in CASES:
        nome, case = entry["nome"], entry["case"]
        c = dict(anno=2026, supc=1317.48, smm=350.0, supi=1656.80572, scatti=0.0,
                 ggRet=25, ggFer=1, ggFest=1, fringe=136.4, fringeCash=0.0, ov=2755.77,
                 regRes=253.95, comRes=52.97, acconto=27.33, arrotT=0.03,
                 arrotN=0.53, oneri=0.0, regione="Lazio", comune="Roma")
        c.update(case)
        ref = py_calc(**c)
        out, bad = xls_calc(c, rows, model)
        if bad:
            fails.append(f"{nome}: errori celle {bad[:5]} (tot {len(bad)})")
        row_bad = []
        for k in OUT_LABELS:
            tot_checked += 1
            if abs(ref[k] - out[k]) > TOLL:
                row_bad.append(f"{k}: py={ref[k]:.2f} xls={out[k]:.2f}")
        if row_bad:
            fails.append(f"{nome}: " + "; ".join(row_bad))
            print(f"  FAIL {nome}")
        else:
            print(f"  OK   {nome}  (comp={ref['comp']:.2f} tratt={ref['tratt']:.2f} "
                  f"nettoM={ref['nettoM']:.2f} nettoA={ref['nettoA']:.2f})")

    for nome, got, exp in SOGLIE:
        if abs(got - exp) > TOLL:
            fails.append(f"{nome}: got={got:.4f} exp={exp:.4f}")
            print(f"  FAIL soglia {nome}")
        else:
            print(f"  OK   soglia {nome} = {got:.2f}")

    refz = py_calc()
    for nome, k, exp in CED01:
        if abs(refz[k] - exp) > TOLL:
            fails.append(f"{nome}: got={refz[k]:.2f} exp={exp:.2f}")
            print(f"  FAIL {nome}")
        else:
            print(f"  OK   {nome} = {refz[k]:.2f}")

    refz3 = py_calc(**D03)
    for nome, k, exp in CED03:
        if abs(refz3[k] - exp) > TOLL:
            fails.append(f"{nome}: got={refz3[k]:.2f} exp={exp:.2f}")
            print(f"  FAIL {nome}")
        else:
            print(f"  OK   {nome} = {refz3[k]:.2f}")

    refsim = py_calc(**D_SIM27)
    for nome, k, exp in SIM27:
        if abs(refsim[k] - exp) > TOLL:
            fails.append(f"{nome}: got={refsim[k]:.2f} exp={exp:.2f}")
            print(f"  FAIL {nome}")
        else:
            print(f"  OK   {nome} = {refsim[k]:.2f}")

    for nome, got, exp in SIM27_ALLEGATO:
        if abs(got - exp) > TOLL:
            fails.append(f"{nome}: got={got:.2f} exp={exp:.2f}")
            print(f"  FAIL {nome}")
        else:
            print(f"  OK   {nome} = {got:.2f}")

    print()
    if fails:
        print(f"ESITO: {len(fails)} FALLIMENTI ({len(CASES)} casi x {len(OUT_LABELS)} "
              f"confronti, {tot_checked} totale; {len(SOGLIE)} soglie; "
              f"{len(CED01)+len(CED03)+len(SIM27)+len(SIM27_ALLEGATO)} "
              f"verifiche benchmark reali)")
        for f in fails:
            print(" -", f)
        return 1
    print(f"ESITO: TUTTI I TEST PASSATI ({len(CASES)} casi, {tot_checked} confronti "
          f"cella per cella tolleranza {TOLL} \u20ac, {len(SOGLIE)} soglie, "
          f"{len(CED01)+len(CED03)+len(SIM27)+len(SIM27_ALLEGATO)} "
          f"verifiche benchmark reali: {len(CED01)} ced.01 + {len(CED03)} ced.03 "
          f"+ {len(SIM27)+len(SIM27_ALLEGATO)} elab. controllo 2027 netto 44.000)")

    print()
    print("=== BENCHMARK (2026, GG 25+1+1 come cedolino, senza override) ===")
    p = py_calc(ov=0, regRes=0, comRes=0, acconto=0, arrotT=0, arrotN=0)
    print(f"auto: irpm={p['irpm']:.2f} regM={p['regM']:.2f} comM={p['comM']:.2f} "
          f"tratt={p['tratt']:.2f} nettoM={p['nettoM']:.2f}")
    for anno in (2026, 2027, 2028):
        p = py_calc(anno=anno, ov=0, regRes=0, comRes=0, acconto=0, arrotT=0, arrotN=0,
                    ggRet=26, ggFer=0, ggFest=0)
        print(f"{anno} mese pieno: RME={p['RME']:.2f} RAL={p['RAL']:.2f} imp={p['imp']:.2f} "
              f"irpl={p['irpl']:.2f} irpn={p['irpn']:.2f} nettoA={p['nettoA']:.2f} "
              f"netto/14={p['nettoA']/14:.2f}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
