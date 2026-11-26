#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Genera 'Simulatore_Dirigente_CCNL_Commercio.xlsx' (regole 2026, dirigente).

Single source of truth delle formule Excel: non modificare l'xlsx a mano.
Genera anche il blocco dati comuni/regioni (elenchi MEF 2026) dentro
simulatore_dirigente.html tra i marker __COMUNI_DATA__ (vedi inject_html_data).

Regole dirigente derivate dal cedolino Zucchetti (Dirig.Terz. livello 5,
Lazio/Roma, IVS 9,19% + agg 1%) e dal TU CCNL Dirigenti Terziario 12/4/2023:
- RME = MINIMO + SUP.CONTR + SMM 2016 + S.MINIMO; retribuzione giornaliera =
  mensile/26 (art. 4 c. 6 TU); 14 mensilita' = 13a dicembre + 14a giugno (art. 8).
- IVS dipendente 9,19% + 1% sopra 56.224 (annui); massimale 122.295.
  Nessun CIGS/FIS/EST per i dirigenti.
- Fondo Mario Negri (base convenzionale 59.224,54): dip 2%, azienda 12,86%+2,52%,
  regime deroga art. 20 c. 7 D.Lgs. 252/2005 -> deducibile per intero.
- FASDAC "Mario Besusso" (base 45.940): dip 1,87%, azienda 5,29%+2,78% pensionati.
- A. Pastore: dirigente 464,81/anno (non deducibile, detraibile 19% max 530 a
  conguaglio); azienda da cedolino 392,20/mese (ufficiale 4.856,45/anno).
- IRPEF lorda per mensilita' = IRPEF annua a scaglioni / 14 (verificata sul
  cedolino: 38.580,79/14 = 2.755,77).
"""
import os

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

import addizionale_lib as ADL

MONEY = '#,##0.00 "\u20ac"'
PCT = '0.00%'
HERE = os.path.dirname(os.path.abspath(__file__))
XLSX = os.path.join(HERE, "Simulatore_Dirigente_CCNL_Commercio.xlsx")
HTML = os.path.join(HERE, "simulatore_dirigente.html")

REG_ROW0 = 4             # prima riga enti regionali (foglio Addizionali)
BIG = 1e15


def section(ws, row, text, ncols=4):
    c = ws.cell(row=row, column=1, value=text)
    c.font = Font(bold=True, size=12, color="FFFFFF")
    fill = PatternFill("solid", fgColor="1F4E78")
    for col in range(1, ncols + 1):
        ws.cell(row=row, column=col).fill = fill
    return row + 1


def put(ws, row, label, value=None, note=None, fmt=None, yellow=False, bold=False):
    ws.cell(row=row, column=1, value=label)
    if value is not None:
        c = ws.cell(row=row, column=2, value=value)
        if fmt:
            c.number_format = fmt
        if yellow:
            c.fill = PatternFill("solid", fgColor="FFF2CC")
            c.font = Font(bold=True)
    if note is not None:
        n = ws.cell(row=row, column=3, value=note)
        n.font = Font(size=9, color="555555", italic=True)
    if bold:
        ws.cell(row=row, column=1).font = Font(bold=True)
    return row + 1


# ============================================================ formule addizionali
def marg_xls(imp, s1, a1, s2, a2, s3, a3, s4, a4):
    """Marginalita' a 4 scaglioni: fase i = (fino a S_i, aliquota A_i)."""
    return ("MIN({i},{s1})*{a1}"
            "+MAX(0,MIN({i},{s2})-{s1})*{a2}"
            "+MAX(0,MIN({i},{s3})-{s2})*{a3}"
            "+MAX(0,MIN({i},{s4})-{s3})*{a4}").format(
        i=imp, s1=s1, a1=a1, s2=s2, a2=a2, s3=s3, a3=a3, s4=s4, a4=a4)


def reg_add_xls(imp):
    p = {k: P[k] for k in ('REGRULE', 'REG_S1', 'REG_A1', 'REG_S2', 'REG_A2',
                           'REG_S3', 'REG_A3', 'REG_S4', 'REG_A4',
                           'REG_DET', 'REG_DET_DA', 'REG_DET_A')}
    m = marg_xls(imp, p['REG_S1'], p['REG_A1'], p['REG_S2'], p['REG_A2'],
                 p['REG_S3'], p['REG_A3'], p['REG_S4'], p['REG_A4'])
    m2 = "MIN({i},{s1})*{a1}+MAX(0,{i}-{s1})*{a2}".format(
        i=imp, s1=p['REG_S1'], a1=p['REG_A1'], a2=p['REG_A2'])
    trento = "IF({i}<=30000,0,IF({i}-30000<=50000,0.0123*({i}-30000),0.0123*50000+0.0173*({i}-80000)))".format(i=imp)
    r = p['REGRULE']
    chain = m
    for cond, true in (
        ('"trento"', trento),
        ('"vda"', "IF({i}<=15000,0,0.0123*{i})".format(i=imp)),
        ('"fvg"', "IF({i}<=15000,0.007*{i},0.0123*{i})".format(i=imp)),
        ('"umbria"', "IF({i}<=28000,0.0123*{i},{m})".format(i=imp, m=m)),
        ('"lazio"', "IF({i}<=28000,0.0173*{i},{m2})".format(i=imp, m2=m2)),
    ):
        assert cond.count('(') == cond.count(')') and true.count('(') == true.count(')')
        chain = "IF({r}={c},{t},{f})".format(r=r, c=cond, t=true, f=chain)
    detr = "IF(AND({d}>0,{i}>{da},{i}<={a}),{d},0)".format(
        d=p['REG_DET'], i=imp, da=p['REG_DET_DA'], a=p['REG_DET_A'])
    return "=MAX(0,{chain}-{detr})".format(chain=chain, detr=detr)


def com_add_xls(imp):
    m = marg_xls(imp, P['COM_S1'], P['COM_A1'], P['COM_S2'], P['COM_A2'],
                 P['COM_S3'], P['COM_A3'], P['COM_S4'], P['COM_A4'])
    return "=IF(OR({e}=0,{i}>{e}),{m},0)".format(e=P['COM_ESE'], i=imp, m=m)


# ============================================================ wb
wb = Workbook()

# ------------------------------------------------------------ LEGGIMI
wsl = wb.active
wsl.title = "Leggimi"
wsl.column_dimensions['A'].width = 112
lines = [
    ("SIMULATORE CEDOLINO DIRIGENTE \u2014 DAL LORDO AL NETTO (regole 2026)", "title"),
    ("CCNL Dirigenti Terziario, Distribuzione e Servizi (Confcommercio\u2013Manageritalia, TU 12/4/2023)", "sub"),
    ("", None),
    ("COME FUNZIONA", "h"),
    ("1. Foglio 'Parametri': TUTTE le variabili (IVS, fondi Negri/FASDAC/Pastore, IRPEF, minimi CCNL 2026-2028).", None),
    ("2. Scegli REGIONE e COMUNE (menu a discesa, dati MEF 2026): addizionali calcolate in automatico.", None),
    ("3. Foglio 'Simulazione': retribuzione in 4 voci, giorni del mese, prospetto annuo e 'cedolino del mese'.", None),
    ("4. Foglio 'Confronto anni': minimi CCNL 2026/2027/2028 con gli stessi parametri.", None),
    ("", None),
    ("REGOLE DIRIGENTE (TU CCNL Dirigenti + cedolino Zucchetti di riferimento)", "h"),
    ("- Retribuzione mensile in 4 voci: MINIMO CCNL (4.660,00 dal 1/1/2026), superminimo contrattuale", None),
    ("  (SUP.CONTR), SMM 2016 (350,00) e superminimo individuale (S.MINIMO): 4.660+1.317,48+350+1.656,81 = 7.984,29.", None),
    ("- Retribuzione GIORNALIERA = mensile / 26 (art. 4 c. 6 TU): nel cedolino 7.984,29/26 = 307,08791 \u20ac.", None),
    ("- 14 mensilita' (art. 8 TU): 13\u00aa a dicembre, 14\u00aa a giugno, retribuzione di fatto.", None),
    ("- Scatti di anzianita' abrogati dal 1/7/2004: resta solo il +258,22 dei premi pre-2004 (art. 9 TU).", None),
    ("- IVS dipendente: 9,19% sull'imponibile + 1% sulla quota annua sopra 56.224 \u20ac, massimale 122.295 \u20ac.", None),
    ("  Per i dirigenti NON si applicano CIGS, FIS bilaterale e Fondo EST (assenti nel cedolino).", None),
    ("- IRPEF 2026: 23% fino a 28.000 \u20ac, 33% da 28.000 a 50.000 \u20ac, 43% oltre (L. 199/2025); detrazione", None),
    ("  art. 13 TUIR come per i dipendenti. L'imponibile dirigente e' sopra 75.000 \u20ac: nessuna somma", None),
    ("  integrativa del cuneo L. 207/2024 ne il trattamento integrativo (fuori fascia).", None),
    ("- IRPEF lorda del MESE = IRPEF annua a scaglioni / 14 (verificata: cedolino 38.580,79/14 = 2.755,77).", None),
    ("- Fondo Mario Negri (previdenza complementare, deroga art. 20 c. 7 D.Lgs. 252/2005): contributi su base", None),
    ("  convenzionale 59.224,54 \u20ac/anno; dirigente 2% (= 98,71/mese), azienda 12,86% + 2,52% integrativo.", None),
    ("  DEDUCIBILE dal reddito per intero, senza tetto di 5.164,46 \u20ac.", None),
    ("- FASDAC 'Mario Besusso' (sanita', NON deducibile): base 45.940 \u20ac/anno; dirigente 1,87% (= 71,59/mese),", None),
    ("  azienda 5,29% + 2,78% gestione pensionati.", None),
    ("- Associazione A. Pastore: dirigente 464,81 \u20ac/anno (= 38,73/mese; nel cedolino base 387,34 x 10%). NON", None),
    ("  deducibile ma detraibile al 19% (max 530 \u20ac) a conguaglio. Azienda: 4.856,45 \u20ac/anno", None),
    ("  (da Q4 2025 incluso premio infortuni 560,00); il cedolino riporta 392,20 \u20ac/mese: composizione gestionale.", None),
    ("- Fondo formazione Mario Negri: 10,83 \u20ac/mese dirigente + 24,17 azienda (valori cedolino; il rinnovo 2026", None),
    ("  potrebbe averli aggiornati: verifica il fondo).", None),
    ("- TFR: art. 2120 c.c. con devoluzione ai fondi (art. 43 TU). NON e' trattenuta: non riduce il netto.", None),
    ("- ADDIZIONALI: menu REGIONE (21 enti MEF) e COMUNE (7.897 comuni, dati 2026); modellate aliquote uniche e", None),
    ("  a scaglioni, soglie di esenzione, regole speciali Lazio/Umbria/FVG/VdA/Trento/Bolzano.", None),
    ("", None),
    ("CONCILIAZIONE COL CEDOLINO ZUCCHETTI (benchmark esatto)", "h"),
    ("- TOTALE COMPETENZE 8.291,38 = round2(307,08791 x 25) + round2(x 1 ferie) + round2(x 1 festivita');", None),
    ("- trattenute 4.107,91 = IVS 761,94 + agg 36,06 + Negri 98,71 + FASDAC 71,59 + Pastore 38,73 + formazione", None),
    ("  10,83 + IRPEF 2.755,77 + add.reg residuo 253,95 + add.com residuo 52,97 + acconto 27,33 + arrotond. 0,03;", None),
    ("- NETTO 4.184,00 = 8.291,38 \u2212 4.107,91 + arrotondato netto 0,53. I residui di addizionali (saldi 2025) e", None),
    ("  l'acconto 2026 sono input facoltativi: ponili a zero per un mese 'pulito'.", None),
    ("- L'imponibile IRPEF del cedolino (7.704,45 x 14 = 107.862,30) non e' ricostruibile in modo deterministico", None),
    ("  con questo modello: Zucchetti annualizza con elementi di conguaglio propri (differenza ~586 \u20ac/mese sul", None),
    ("  calcolo automatico). Per replicare il cedolino inserisci l'IRPEF lorda del mese 2.755,77 come override.", None),
    ("", None),
    ("FONTI E VIGENZA", "h"),
    ("- TU CCNL Dirigenti Terziario 12/4/2023 (artt. 4, 8, 9, 27, 28, 29, 43) \u2014 manageritalia.it.", None),
    ("- Rinnovo CCNL 5/11/2025: minimo 4.660 (2026), 4.920 (2027), 5.140 (2028), aumenti art. 6 gi\u00e0 inclusi.", None),
    ("- FASDAC fasdac.it; Fondo Mario Negri fondonegri.it; Associazione A. Pastore associazionepastore.it.", None),
    ("- INPS Circ. 6/2026 \u2014 soglia 1%: 56.224 \u20ac; massimale 122.295 \u20ac.", None),
    ("- IRPEF: L. 199/2025 (scaglioni); detrazione art. 13 TUIR, L. 207/2024.", None),
    ("- Addizionali: elenchi generali 2026 MEF (Dipartimento delle Finanze), foglio 'Addizionali'.", None),
    ("", None),
    ("DISCLAIMER: strumento di simulazione a scopo informativo; non sostituisce il cedolino ufficiale n\u00e9 un parere di", None),
    ("consulente del lavoro. Verifica sempre addizionali locali e casi particolari.", "sub"),
]
r = 1
for text, style in lines:
    c = wsl.cell(row=r, column=1, value=text if text else None)
    if style == "title":
        c.font = Font(bold=True, size=14)
    elif style == "h":
        c.font = Font(bold=True, size=11, color="1F4E78")
    elif style == "sub":
        c.font = Font(italic=True, size=10, color="555555")
    r += 1

# ------------------------------------------------------------ PARAMETRI
ws = wb.create_sheet("Parametri")
ws.column_dimensions['A'].width = 58
ws.column_dimensions['B'].width = 16
ws.column_dimensions['C'].width = 60
P = {}

r = section(ws, 1, "RETRIBUZIONE DIRIGENTE \u2014 2026 (TU CCNL Dirigenti + rinnovo 5/11/2025)")
r = put(ws, r, "MINIMO CCNL mensile 2026", 4660.0, "aumento art. 6 (320) gi\u00e0 incluso; 4.920 (2027), 5.140 (2028)", MONEY, True); P['MIN26'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Mensilita' CCNL (art. 8 TU: 13\u00aa dic. + 14\u00aa giu.)", 14, "", None, True); P['MENS'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Giorni per mensilita' (art. 4 c. 6 TU)", 26, "retrib. giornaliera = mensile / 26", None, True); P['NG'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Valore scatto pre-2004", 258.22, "art. 9 TU: abrogati dal 1/7/2004, resta solo questo", MONEY, True); P['SC04'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "CONTRIBUTI INPS IVS (DIRIGENTE) \u2014 2026")
r = put(ws, r, "Aliquota INPS dipendente (IVS)", 0.0919, "INPS Circ. 6/2026; identica per i dirigenti", PCT, True); P['INPS'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Soglia annua aliquota aggiuntiva 1%", 56224, "INPS Circ. 6/2026", MONEY, True); P['AGG_SOG'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota aggiuntiva IVS sopra soglia", 0.01, "", PCT, True); P['AGG_ALQ'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Massimale imponibile contributivo annuo", 122295, "INPS Circ. 6/2026", MONEY, True); P['MAXC'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Oneri datore IVS (solo informativo)", 0.2607, "stima quota azienda; nessun CIGS/FIS/EST per dirigenti", PCT, True); P['DATORE'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "FONDO MARIO NEGRI (art. 27 TU \u2014 previdenza complementare, regime deroga)")
r = put(ws, r, "Base convenzionale annua", 59224.54, "fondonegri.it; 4.935,38/mese nel cedolino", MONEY, True); P['MN_BASE'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota dirigente", 0.02, "raddoppiata dal 1/1/2026 (era 1%)", PCT, True); P['MN_D'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota azienda base", 0.1286, "", PCT, True); P['MN_D12'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota azienda integrativa", 0.0252, "", PCT, True); P['MN_D2A'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Fondo formazione \u2014 quota dirigente (mensile)", 10.83, "valore cedolino; possibile aggiornamento col rinnovo", MONEY, True); P['MNF_DIP'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Fondo formazione \u2014 quota azienda (mensile)", 24.17, "", MONEY, True); P['MNF_DIT'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "FASDAC 'MARIO BESUSSO' (art. 29 TU \u2014 assistenza sanitaria)")
r = put(ws, r, "Base convenzionale annua", 45940.0, "fasdac.it; 3.828,33/mese nel cedolino", MONEY, True); P['FD_BASE'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota dirigente", 0.0187, "", PCT, True); P['FD_D'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota azienda", 0.0529, "", PCT, True); P['FD_D12'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota azienda gestione pensionati", 0.0278, "", PCT, True); P['FD_D2A'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "ASSOCIAZIONE A. PASTORE (art. 28 TU \u2014 previdenza/assistenza)")
r = put(ws, r, "Quota dirigente annua", 464.81, "associazionepastore.it; base 387,34 x 10% nel cedolino", MONEY, True); P['AP_ANN'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Quota dirigente mensile", f"=ROUND({P['AP_ANN']}/12,2)", "non deducibile; detraibile 19% (max 530 \u20ac) a conguaglio", MONEY); P['AP_M'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Quota azienda mensile (da cedolino)", 392.20, "informativo; ufficiale 2026: 4.856,45/anno (404,70/mese)", MONEY, True); P['AP_DIT'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "SCAGLIONI IRPEF 2026 (L. 199/2025)")
r = put(ws, r, "Fine scaglione 1 \u2014 Aliquota 1", 28000, "fino a 28.000 \u20ac", MONEY, True); P['S1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota scaglione 1", 0.23, "", PCT, True); P['A1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Fine scaglione 2 \u2014 Aliquota 2", 50000, "28.000\u201350.000 \u20ac (33%)", MONEY, True); P['S2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota scaglione 2", 0.33, "", PCT, True); P['A2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Aliquota scaglione 3", 0.43, "oltre 50.000 \u20ac", PCT, True); P['A3'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "DETRAZIONE LAVORO DIPENDENTE (art. 13 co. 1 TUIR, L. 207/2024)")
r = put(ws, r, "Detrazione fascia 1 (fino a reddito...)", 1955, "red. \u2264 15.000 \u20ac", MONEY, True); P['D1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Soglia fine fascia 1", 15000, "", MONEY, True); P['DS1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Detrazione base fasce 2-3", 1910, "", MONEY, True); P['DB'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Incremento 1.190 (fascia 15.000\u201328.000)", 1190, "", MONEY, True); P['DINC'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Soglia fine fascia 2", 28000, "", MONEY, True); P['DS2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Divisore fascia 2 (28.000\u221215.000)", 13000, "", MONEY, True); P['DDIV2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Soglia fine fascia 3 (detrazione fissa)", 50000, "", MONEY, True); P['DS3'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Fascia 4: detrazione decrescente fino a", 75000, "da 1.910 \u20ac a 0 tra 50.000 e 75.000", MONEY, True); P['DS4'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Divisore fascia 4 (75.000\u221250.000)", 25000, "", MONEY, True); P['DDIV4'] = f"Parametri!$B${r-1}"

r += 1
r = section(ws, r, "ADDIZIONALI \u2014 REGIONE E COMUNE (elenchi MEF 2026, foglio Addizionali)")
r = put(ws, r, "Regione (menu a discesa)", ADL.DEFAULT_REGIONE, "menu a discesa: 21 enti MEF", None, True)
cell_region = r - 1
r = put(ws, r, "Comune (menu a discesa)", ADL.DEFAULT_COMUNE, "dipendente dalla regione scelta", None, True)
cell_comune = r - 1
cell_NOTA = r
r = put(ws, r, "\u2014 dati applicati (calcolati dal foglio Addizionali) \u2014", None, None)
r = put(ws, r, "Regola regionale (codice)", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},2,0),\"std\")", "", None)
P['REGRULE'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 soglia 1", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},3,0),{ADL.BIG})", "", MONEY); P['REG_S1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 aliquota 1", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},4,0),0)", "", PCT); P['REG_A1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 soglia 2", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},5,0),{ADL.BIG})", "", MONEY); P['REG_S2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 aliquota 2", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},6,0),0)", "", PCT); P['REG_A2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 soglia 3", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},7,0),{ADL.BIG})", "", MONEY); P['REG_S3'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 aliquota 3", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},8,0),0)", "", PCT); P['REG_A3'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 soglia 4", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},9,0),{ADL.BIG})", "", MONEY); P['REG_S4'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 aliquota 4", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},10,0),0)", "", PCT); P['REG_A4'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 detrazione", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},11,0),0)", "", MONEY); P['REG_DET'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 detrazione da (escluso)", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},12,0),0)", "", MONEY); P['REG_DET_DA'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 detrazione fino a (incluso)", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},13,0),0)", "", MONEY); P['REG_DET_A'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Reg. \u2014 nota", f"=IFERROR(VLOOKUP($B${cell_region},Addizionali!$A${REG_ROW0}:$N${REG_ROW0+20},14,0),\"\")", "", None)
P['REG_NOTA'] = f"Parametri!$B${r-1}"

cell_com_start = r  # le righe COMUNE si riempiono dopo la creazione del foglio Addizionali

# ------------------------------------------------------------ ADDIZIONALI (dati MEF)
wsa = wb.create_sheet("Addizionali")
wsa.column_dimensions['A'].width = 30
for col in "BCDEFGHIJKLM":
    wsa.column_dimensions[col].width = 12
wsa.column_dimensions['N'].width = 66
section(wsa, 1, "ADDIZIONALI MEF 2026 \u2014 ELENCHI GENERALI (Dipartimento delle Finanze)", 14)
wsa.cell(row=2, column=1, value="Fonte: www1.finanze.gov.it > fiscalita' locale > elenchi generali (comuni e regioni), 2026. "
                                "Aliquote in frazione (0.009 = 0,90%). Comuni con nota 2025: delibera 2026 non ancora pubblicata.").font = Font(size=9, italic=True, color="555555")
for i, h in enumerate(["Ente", "Regola", "Soglia 1", "Aliq. 1", "Soglia 2", "Aliq. 2",
                       "Soglia 3", "Aliq. 3", "Soglia 4", "Aliq. 4", "Detrazione", "Detr. da", "Detr. a", "Nota"], start=1):
    c = wsa.cell(row=3, column=i, value=h)
    c.font = Font(bold=True, size=9); c.fill = PatternFill("solid", fgColor="DDEBF7")
row = REG_ROW0
for ente in sorted(ADL._REGIONALI.keys(), key=str.lower):
    d = ADL._REGIONALI[ente]
    wsa.cell(row=row, column=1, value=ente).font = Font(bold=True)
    wsa.cell(row=row, column=2, value=d['regola'])
    for j, (s, a) in enumerate(d['brackets']):
        cs = wsa.cell(row=row, column=3 + j * 2, value=s); cs.number_format = '#,##0'
        ca = wsa.cell(row=row, column=4 + j * 2, value=a); ca.number_format = PCT
    wsa.cell(row=row, column=11, value=d['detr'] or 0).number_format = MONEY
    wsa.cell(row=row, column=12, value=d['detr_da'] or 0).number_format = '#,##0'
    wsa.cell(row=row, column=13, value=d['detr_a'] or 0).number_format = '#,##0'
    wsa.cell(row=row, column=14, value=d['nota']).font = Font(size=9, italic=True)
    row += 1
regtbl_end = row - 1
wb.defined_names.add(DefinedName("RegElenco", attr_text=f"Addizionali!$A${REG_ROW0}:$A${regtbl_end}"))

com_hdr = regtbl_end + 4
wsa.cell(row=com_hdr - 1, column=1, value="COMUNI (7.897) \u2014 menu del foglio Parametri").font = Font(bold=True, size=11, color="1F4E78")
for i, h in enumerate(["Chiave (regione|comune)", "Comune", "Provincia", "Regola", "Aliq. 1", "Soglia 1",
                       "Aliq. 2", "Soglia 2", "Aliq. 3", "Soglia 3", "Aliq. 4", "Soglia 4", "Esenzione", "Nota"], start=1):
    c = wsa.cell(row=com_hdr, column=i, value=h)
    c.font = Font(bold=True, size=9); c.fill = PatternFill("solid", fgColor="DDEBF7")
com_row0 = com_hdr + 1
row = com_row0
region_bounds = {}
import csv as _csv
with open(os.path.join(HERE, "data", "addizionali_comuni_2026.csv"), encoding="utf-8") as f:
    rdr = _csv.reader((ln for ln in f if not ln.startswith("#")), delimiter=";")
    next(rdr)
    for reg_row in rdr:
        regione, prov, nome, regola = ADL._norm_regione(reg_row[0]), reg_row[1], reg_row[2], reg_row[3]
        a1, s1, a2, s2, a3, s3, a4, s4, ese = reg_row[4:13]
        nota = reg_row[13]
        wsa.cell(row=row, column=1, value=f"{regione}|{nome}")
        wsa.cell(row=row, column=2, value=nome)
        wsa.cell(row=row, column=3, value=prov)
        wsa.cell(row=row, column=4, value=regola)
        for j, (v, fmt) in enumerate(((a1, PCT), (s1, '#,##0'), (a2, PCT), (s2, '#,##0'),
                                      (a3, PCT), (s3, '#,##0'), (a4, PCT), (s4, '#,##0'),
                                      (ese, '#,##0.00'))):
            c = wsa.cell(row=row, column=5 + j, value=float(v))
            if fmt:
                c.number_format = fmt
        if nota:
            wsa.cell(row=row, column=14, value=nota).font = Font(size=8, italic=True, color="777777")
        if regione not in region_bounds:
            region_bounds[regione] = [row, row]
        region_bounds[regione][1] = row
        row += 1
comtbl_end = row - 1
for i, ente in enumerate(sorted(ADL._REGIONALI.keys(), key=str.lower), start=1):
    a, b = region_bounds[ente]
    wb.defined_names.add(DefinedName(f"Reg{i:02d}", attr_text=f"Addizionali!$B${a}:$B${b}"))

# righe COMUNE del foglio Parametri (ora che il range e' noto)
ws = wb["Parametri"]
COMTBL = f"Addizionali!$A${com_row0}:$N${comtbl_end}"
CKEY = f"$B${cell_region}&\"|\"&$B${cell_comune}"
r = cell_com_start
r = put(ws, r, "Com. \u2014 regola/nota", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},4,0),\"n\")", "", None)
P['COM_R'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 aliquota 1", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},5,0),0)", "", PCT); P['COM_A1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 soglia 1", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},6,0),{ADL.BIG})", "", MONEY); P['COM_S1'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 aliquota 2", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},7,0),0)", "", PCT); P['COM_A2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 soglia 2", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},8,0),{ADL.BIG})", "", MONEY); P['COM_S2'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 aliquota 3", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},9,0),0)", "", PCT); P['COM_A3'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 soglia 3", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},10,0),{ADL.BIG})", "", MONEY); P['COM_S3'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 aliquota 4", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},11,0),0)", "", PCT); P['COM_A4'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 soglia 4", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},12,0),{ADL.BIG})", "", MONEY); P['COM_S4'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 esenzione fino a", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},13,0),0)", "sopra la soglia: aliquota su tutto l'imponibile", MONEY); P['COM_ESE'] = f"Parametri!$B${r-1}"
r = put(ws, r, "Com. \u2014 nota", f"=IFERROR(VLOOKUP({CKEY},{COMTBL},14,0),\"\")", "", None)
P['COM_NOTA'] = f"Parametri!$B${r-1}"
ws.cell(row=cell_NOTA, column=3, value="se il comune non compare, imposta manualmente le righe 'Com. \u2014 aliquota 1' e 'Com. \u2014 esenzione'")

dv_reg = DataValidation(type="list", formula1=f"Addizionali!$A${REG_ROW0}:$A${REG_ROW0+20}", allow_blank=False)
ws.add_data_validation(dv_reg)
dv_reg.add(ws.cell(row=cell_region, column=2))
dv_com = DataValidation(type="list", formula1='INDIRECT("Reg"&TEXT(MATCH($B${r},Addizionali!$A${s}:$A${e},0),"00"))'.format(r=cell_region, s=REG_ROW0, e=REG_ROW0+20), allow_blank=True)
ws.add_data_validation(dv_com)
dv_com.add(ws.cell(row=cell_comune, column=2))

# tabella minimi CCNL (per il menu Anno del foglio Simulazione)
r += 1
r = section(ws, r, "MINIMI CCNL DIRIGENTE \u2014 RINNOVO 5/11/2025 (lordi/mese)")
hdr_row = r
for col, h in ((1, "Anno"), (2, "Minimo mensile"), (3, "Aumento art. 6"), (4, "Nota")):
    ws.cell(row=hdr_row, column=col, value=h).font = Font(bold=True, size=9)
r += 1
tbl_start = r
MINIMI = [
    ("2026", 4660.0, 320.0, "1/1/2026 \u2014 aumento art. 6 (320) incluso nel minimo"),
    ("2027", 4920.0, 260.0, "1/1/2027"),
    ("2028", 5140.0, 220.0, "1/1/2028"),
]
for anno, minimo, aumento, nota in MINIMI:
    c = ws.cell(row=r, column=1, value=anno); c.number_format = "@"
    c = ws.cell(row=r, column=2, value=minimo); c.number_format = MONEY
    ca = ws.cell(row=r, column=3, value=aumento); ca.number_format = MONEY
    ws.cell(row=r, column=4, value=nota).font = Font(size=9, italic=True)
    r += 1
tbl_end = r - 1
P['TBL'] = f"Parametri!$A${tbl_start}:$B${tbl_end}"

# ------------------------------------------------------------ SIMULAZIONE
ws = wb.create_sheet("Simulazione")
ws.column_dimensions['A'].width = 56
ws.column_dimensions['B'].width = 18
ws.column_dimensions['C'].width = 62

r = section(ws, 1, "RETRIBUZIONE DI FATTO (celle gialle: modificarle)", 3)
r = put(ws, r, "Anno di paga", 2026, "minimo CCNL dal menu (rinnovo 5/11/2025)", None, True)
cell_anno = r - 1
ws.cell(row=cell_anno, column=2).number_format = "0"
dv_anno = DataValidation(type="list", formula1='"2026,2027,2028"', allow_blank=False)
ws.add_data_validation(dv_anno); dv_anno.add(ws.cell(row=cell_anno, column=2))
r = put(ws, r, "MINIMO CCNL mensile", f"=IFERROR(VLOOKUP(TEXT($B${cell_anno},\"0\"),{P['TBL']},2,0),VLOOKUP($B${cell_anno},{P['TBL']},2,0))", "dal foglio Parametri", MONEY)
c_minimo = r - 1
r = put(ws, r, "SUP.CONTR \u2014 superminimo contrattuale (mensile)", 1317.48, "Allegato C TU / valore cedolino", MONEY, True)
c_supc = r - 1
r = put(ws, r, "SMM 2016 (mensile)", 350.0, "SUPERMM 2016 nel cedolino", MONEY, True)
c_smm = r - 1
r = put(ws, r, "S.MINIMO \u2014 superminimo individuale (mensile)", 1656.80572, "valore esatto cedolino (display 1.656,81)", MONEY, True)
c_supi = r - 1
r = put(ws, r, "Scatti pre-2004 (mensile)", 0.0, "art. 9 TU: abrogati dal 2004; qui 0 di default", MONEY, True)
c_scatti = r - 1
r = put(ws, r, "RME \u2014 retribuzione mensile effettiva", f"=B{c_minimo}+B{c_supc}+B{c_smm}+B{c_supi}+B{c_scatti}", "", MONEY, False, True)
c_rme = r - 1
r = put(ws, r, "Retribuzione giornaliera (RME/26, art. 4 c. 6 TU)", f"=$B${c_rme}/{P['NG']}", "cedolino: 307,08791", MONEY)
c_giorn = r - 1

r += 1
r = section(ws, r, "GIORNI E VOCI DEL MESE", 3)
r = put(ws, r, "Giorni retribuiti (Retribuzione)", 25, "cedolino: 25,00000 GG", None, True)
c_ggret = r - 1
r = put(ws, r, "Giorni di ferie godute", 1, "Z00250", None, True)
c_ggfer = r - 1
r = put(ws, r, "Giorni festivita'/permessi", 1, "Z01138", None, True)
c_ggfest = r - 1
r = put(ws, r, "Fringe benefit mensile (in natura, tassabile)", 136.4, "voce in natura: tassabile, non e' flusso di cassa", MONEY, True)
c_fringe = r - 1

r += 1
r = section(ws, r, "OVERRIDE DA CEDOLINO (0 = calcolo automatico del modello annuale)", 3)
r = put(ws, r, "IRPEF lorda del mese \u2014 override", 2755.77, "cedolino F02010; 0 = IRPEF annua/14 di questo foglio", MONEY, True)
c_irpov = r - 1
r = put(ws, r, "Add. regionale del mese \u2014 residuo conguaglio", 253.95, "cedolino F09110 (saldo 2025 Lazio); 0 = rateizzata", MONEY, True)
c_regres = r - 1
r = put(ws, r, "Add. comunale del mese \u2014 residuo conguaglio", 52.97, "cedolino F09130 (saldo 2025 Roma); 0 = rateizzata", MONEY, True)
c_comres = r - 1
r = put(ws, r, "Acconto addizionali extra (da cedolino)", 27.33, "cedolino F09140 (acconto 2026 Roma)", MONEY, True)
c_acconto = r - 1
r = put(ws, r, "Arrotondamento trattenute (Zucchetti)", 0.03, "voce ZP9960; 0 = niente", MONEY, True)
c_arrotT = r - 1
r = put(ws, r, "Arrotondamento netto (Zucchetti)", 0.53, "pie' pagina cedolino; 0 = niente", MONEY, True)
c_arrotN = r - 1
r = put(ws, r, "Altri oneri/detrazioni annuali (es. 19% Pastore, max 530)", 0.0, "detrazioni a conguaglio, assenti nel cedolino", MONEY, True)
c_oneri = r - 1

r += 1
r = section(ws, r, "CEDOLINO DEL MESE (mensilita' ordinaria o 13\u00aa/14\u00aa: stessa rata di trattenute)", 3)
r = put(ws, r, "VOCE", "Importo", None)
voci_hdr = r - 1
ws.cell(row=voci_hdr, column=1).font = Font(bold=True); ws.cell(row=voci_hdr, column=2).font = Font(bold=True)
r = put(ws, r, "Retribuzione (gg x retr. giornaliera)", f"=ROUND($B${c_ggret}*$B${c_giorn},2)", "Z00001", MONEY, True)
c_ret = r - 1
r = put(ws, r, "Ferie godute (gg x retr. giornaliera)", f"=ROUND($B${c_ggfer}*$B${c_giorn},2)", "Z00250", MONEY)
c_fer = r - 1
r = put(ws, r, "Festivita'/permessi (gg x retr. giornaliera)", f"=ROUND($B${c_ggfest}*$B${c_giorn},2)", "Z01138", MONEY)
c_fest = r - 1
r = put(ws, r, "TOTALE COMPETENZE", f"=B{c_ret}+B{c_fer}+B{c_fest}", "cedolino: 8.291,38", MONEY, True)
c_comp = r - 1
r = put(ws, r, "\u2212 Contributo IVS 9,19%", f"=ROUND(FLOOR(B{c_comp},1)*{P['INPS']},2)", "Z00000: base gia' arrotondata all'euro dal gestionale (8.291,00)", MONEY)
c_ivsm = r - 1
r = put(ws, r, "\u2212 Contributo IVS agg. 1%", f"=ROUND(MAX(0,FLOOR(B{c_comp},1)-{P['AGG_SOG']}/12)*{P['AGG_ALQ']},2)", "Z00010: soglia mensile 56.224/12 = 4.685,33", MONEY)
c_aggm = r - 1
r = put(ws, r, "\u2212 Fondo Mario Negri 2% (base convenzionale)", f"=ROUND({P['MN_BASE']}/12*{P['MN_D']},2)", "Z20100: 59.224,54/12 x 2% = 98,71", MONEY)
c_negrim = r - 1
r = put(ws, r, "\u2212 FASDAC 1,87% (base convenzionale)", f"=ROUND({P['FD_BASE']}/12*{P['FD_D']},2)", "Z31055: 45.940/12 x 1,87% = 71,59", MONEY)
c_fasdaqm = r - 1
r = put(ws, r, "\u2212 A. Pastore quota dirigente", f"={P['AP_M']}", "464,81/12 = 38,73", MONEY)
c_pastm = r - 1
r = put(ws, r, "\u2212 Fondo formazione Mario Negri", f"={P['MNF_DIP']}", "quota fissa 10,83", MONEY)
c_formm = r - 1
r = put(ws, r, "IRPEF lorda a debito del mese", f"=IF($B${c_irpov}>0,$B${c_irpov},0)", "F02010; override attivo se > 0, altrimenti annua/14", MONEY, True)
c_irpm = r - 1
ws.cell(row=c_irpm, column=2).value = None  # formula definitiva assegnata dopo (dipende da IRPEF annua)
r = put(ws, r, "\u2212 Addizionale regionale del mese", f"=IF($B${c_regres}>0,$B${c_regres},0)", "F09110: residuo o rateizzata", MONEY, True)
c_regm = r - 1
ws.cell(row=c_regm, column=2).value = None
r = put(ws, r, "\u2212 Addizionale comunale del mese", "=0", "F09130: residuo o rateizzata", MONEY)
c_comm = r - 1
ws.cell(row=c_comm, column=2).value = None
r = put(ws, r, "\u2212 Acconto addizionali extra", f"=$B${c_acconto}", "F09140: trattenuta additiva del mese", MONEY)
c_accm = r - 1
r = put(ws, r, "\u2212 Arrotondamento trattenute", f"=$B${c_arrotT}", "ZP9960", MONEY)
r = put(ws, r, "TOTALE TRATTENUTE", f"=B{c_ivsm}+B{c_aggm}+B{c_negrim}+B{c_fasdaqm}+B{c_pastm}+B{c_formm}+B{c_irpm}+B{c_regm}+B{c_comm}+B{c_accm}+$B${c_arrotT}", "cedolino: 4.107,91", MONEY, True)
c_tratt = r - 1
r = put(ws, r, "NETTO DEL MESE", f"=B{c_comp}-B{c_tratt}+B{c_arrotN}", "cedolino: 4.184,00 (8.291,38 \u2212 4.107,91 + 0,53)", MONEY, True)
c_nettoM = r - 1
r = put(ws, r, "(Fringe benefit del mese: in natura, non in cassa)", f"=$B${c_fringe}", "informativo: tassabile, non tocca il netto", MONEY)

r += 1
r = section(ws, r, "CALCOLO ANNUO (modello a scaglioni \u2014 base IRPEF del gestionale non ricostruita)", 3)
r = put(ws, r, "RAL \u2014 retribuzione annua lorda (RME x 14)", f"=$B${c_rme}*{P['MENS']}", "", MONEY, True)
c_RAL = r - 1
r = put(ws, r, "IVS 9,19% annua (fino al massimale)", f"=MIN($B${c_RAL},{P['MAXC']})*{P['INPS']}", "", MONEY)
c_ivsa = r - 1
r = put(ws, r, "IVS agg. 1% annua (sopra 56.224)", f"=MAX(0,$B${c_RAL}-{P['AGG_SOG']})*{P['AGG_ALQ']}", "", MONEY)
c_agga = r - 1
r = put(ws, r, "Fondo Mario Negri 2% annuo (base convenzionale)", f"={P['MN_BASE']}*{P['MN_D']}", "deducibile per intero (regime deroga)", MONEY)
c_negria = r - 1
r = put(ws, r, "FASDAC 1,87% annuo", f"={P['FD_BASE']}*{P['FD_D']}", "non deducibile", MONEY)
c_fasdca = r - 1
r = put(ws, r, "A. Pastore annuo", f"={P['AP_ANN']}", "detrazione 19% solo a conguaglio (input oneri)", MONEY)
c_pasta = r - 1
r = put(ws, r, "Fondo formazione annuo", f"={P['MNF_DIP']}*12", "quota dirigente x 12 mesi", MONEY)
r = put(ws, r, "RLD — reddito di riferimento: RAL − IVS − agg. − Negri", f"=$B${c_RAL}-$B${c_ivsa}-$B${c_agga}-$B${c_negria}", "base detrazioni e deduzione Negri", MONEY, True)
c_rld = r - 1
r = put(ws, r, "Imponibile IRPEF annuo (RLD + fringe x 12)", f"=$B${c_rld}+$B${c_fringe}*12", "fringe in natura: tassabile per i dirigenti", MONEY, True)
c_imp = r - 1
r = put(ws, r, "IRPEF lorda annua", f"=IF($B${c_imp}<={P['S1']},$B${c_imp}*{P['A1']},{P['S1']}*{P['A1']}+IF($B${c_imp}<={P['S2']},($B${c_imp}-{P['S1']})*{P['A2']},({P['S2']}-{P['S1']})*{P['A2']}+($B${c_imp}-{P['S2']})*{P['A3']}))", "L. 199/2025: 23/33/43%", MONEY)
c_irpla = r - 1
ws.cell(row=c_irpm, column=2).value = f"=IF($B${c_irpov}>0,$B${c_irpov},ROUND($B${c_irpla}/{P['MENS']},2))"
r = put(ws, r, "Oneri/detrazioni annuali (input)", f"=$B${c_oneri}", "riducono l'imposta netta a conguaglio", MONEY)
c_dedon = r - 1
r = put(ws, r, "Detrazione lavoro dipendente (art. 13 TUIR)", f"=IF($B${c_rld}<={P['DS1']},{P['D1']},IF($B${c_rld}<={P['DS2']},{P['DB']}+{P['DINC']}*({P['DS2']}-$B${c_rld})/{P['DDIV2']},IF($B${c_rld}<={P['DS3']},{P['DB']},IF($B${c_rld}<={P['DS4']},{P['DB']}*({P['DS4']}-$B${c_rld})/{P['DDIV4']},0))))", "sul reddito di riferimento RLD", MONEY)
c_ded = r - 1
r = put(ws, r, "Detrazione effettiva (capienza imposta)", f"=MIN($B${c_ded},$B${c_irpla})", "", MONEY)
c_dede = r - 1
r = put(ws, r, "IRPEF NETTA annua", f"=MAX(0,$B${c_irpla}-$B${c_dede}-$B${c_dedon})", "", MONEY, True)
c_irpna = r - 1
r = put(ws, r, "Addizionale regionale annua", reg_add_xls(f"$B${c_imp}"), "regola MEF della regione scelta", MONEY)
c_rega = r - 1
ws.cell(row=c_regm, column=2).value = f"=IF($B${c_regres}>0,$B${c_regres},ROUND($B${c_rega}/12,2))"
r = put(ws, r, "Addizionale comunale annua", com_add_xls(f"$B${c_imp}"), "regola MEF del comune scelto", MONEY)
c_coma = r - 1
ws.cell(row=c_comm, column=2).value = f"=IF($B${c_comres}>0,$B${c_comres},ROUND($B${c_coma}/12,2))"
r = put(ws, r, "Totale trattenute annue", f"=$B${c_ivsa}+$B${c_agga}+$B${c_negria}+$B${c_fasdca}+$B${c_pasta}+$B${c_formm}*12+$B${c_irpna}+$B${c_rega}+$B${c_coma}", "senza frange e arrotondamenti di cassa", MONEY)
c_tta = r - 1
r = put(ws, r, "NETTO ANNUO in busta", f"=$B${c_RAL}-$B${c_tta}", "", MONEY, True)
c_nettoa = r - 1
r = put(ws, r, "Netto medio per mensilita' (annuo / 14)", f"=$B${c_nettoa}/{P['MENS']}", "", MONEY)

r += 1
r = section(ws, r, "VOCI A CARICO AZIENDA (informativo: non toccano il netto)", 3)
r = put(ws, r, "Fondo Mario Negri azienda base 12,86%", f"=ROUND({P['MN_BASE']}/12*{P['MN_D12']},2)", "art. 27 TU", MONEY)
r = put(ws, r, "Fondo Mario Negri azienda integrativo 2,52%", f"=ROUND({P['MN_BASE']}/12*{P['MN_D2A']},2)", "", MONEY)
r = put(ws, r, "FASDAC azienda 5,29%", f"=ROUND({P['FD_BASE']}/12*{P['FD_D12']},2)", "art. 29 TU", MONEY)
r = put(ws, r, "FASDAC gestione pensionati 2,78%", f"=ROUND({P['FD_BASE']}/12*{P['FD_D2A']},2)", "", MONEY)
r = put(ws, r, "A. Pastore azienda (da cedolino)", f"={P['AP_DIT']}", "ufficiale 2026: 4.856,45/anno", MONEY)
r = put(ws, r, "Fondo formazione Mario Negri azienda", f"={P['MNF_DIT']}", "", MONEY)
r = put(ws, r, "Oneri datore IVS (stima 26,07%)", f"=MIN($B${c_RAL},{P['MAXC']})*{P['DATORE']}", "solo informativo", MONEY)
c_costo = r - 1
r = put(ws, r, "Costo azienda annuo indicativo (RAL + oneri + fondi)", f"=$B${c_RAL}+$B${c_costo}+($B${c_negria}+12*({P['MN_BASE']}/12*{P['MN_D2A']}+{P['FD_BASE']}/12*{P['FD_D12']}+{P['FD_BASE']}/12*{P['FD_D2A']})+12*{P['AP_DIT']}+12*{P['MNF_DIT']})", "senza INAIL e voci ricorrenti", MONEY)

# ------------------------------------------------------------ CONFRONTO ANNI
ws = wb.create_sheet("Confronto anni")
ws.column_dimensions['A'].width = 10
for col in range(2, 18):
    ws.column_dimensions[get_column_letter(col)].width = 14
r = section(ws, 1, "CONFRONTO MINIMI CCNL 2026-2028 \u2014 dal lordo al netto con i parametri correnti", 16)
headers = ["Anno", "Minimo mese", "RME mese", "RAL", "IVS 9,19%", "Agg. 1%", "Negri 2%", "FASDAC 1,87%",
           "RLD", "Imponibile", "IRPEF lorda", "Detrazione", "IRPEF netta", "Add. reg.", "Add. com.", "Netto annuo"]
for i, h in enumerate(headers, start=1):
    c = ws.cell(row=2, column=i, value=h)
    c.font = Font(bold=True, size=9)
    c.fill = PatternFill("solid", fgColor="DDEBF7")
    c.alignment = Alignment(wrap_text=True, horizontal="center")

SUPC = f"Simulazione!$B${c_supc}"
SMM = f"Simulazione!$B${c_smm}"
SUPI = f"Simulazione!$B${c_supi}"
Fringe = f"Simulazione!$B${c_fringe}"
for i, (anno, minimo, _, _) in enumerate(MINIMI):
    row = 3 + i
    ws.cell(row=row, column=1, value=anno).font = Font(bold=True)
    cc = ws.cell(row=row, column=2, value=minimo); cc.number_format = MONEY
    ws.cell(row=row, column=3, value=f"=B{row}+{SUPC}+{SMM}+{SUPI}")
    ws.cell(row=row, column=4, value=f"=C{row}*{P['MENS']}")
    ws.cell(row=row, column=5, value=f"=MIN(D{row},{P['MAXC']})*{P['INPS']}")
    ws.cell(row=row, column=6, value=f"=MAX(0,D{row}-{P['AGG_SOG']})*{P['AGG_ALQ']}")
    ws.cell(row=row, column=7, value=f"={P['MN_BASE']}*{P['MN_D']}")
    ws.cell(row=row, column=8, value=f"={P['FD_BASE']}*{P['FD_D']}")
    ws.cell(row=row, column=9, value=f"=D{row}-E{row}-F{row}-G{row}")
    ws.cell(row=row, column=10, value=f"=I{row}+{Fringe}*12")
    imp = f"J{row}"
    ws.cell(row=row, column=11, value=(f"=IF({imp}<={P['S1']},{imp}*{P['A1']},{P['S1']}*{P['A1']}+"
        f"IF({imp}<={P['S2']},({imp}-{P['S1']})*{P['A2']},({P['S2']}-{P['S1']})*{P['A2']}+({imp}-{P['S2']})*{P['A3']}))"))
    rld = f"I{row}"
    ws.cell(row=row, column=12, value=(f"=IF({rld}<={P['DS1']},{P['D1']},IF({rld}<={P['DS2']},{P['DB']}+{P['DINC']}*({P['DS2']}-{rld})/{P['DDIV2']},"
        f"IF({rld}<={P['DS3']},{P['DB']},IF({rld}<={P['DS4']},{P['DB']}*({P['DS4']}-{rld})/{P['DDIV4']},0))))"))
    ws.cell(row=row, column=13, value=f"=MAX(0,K{row}-L{row})")
    ws.cell(row=row, column=14, value=reg_add_xls(imp))
    ws.cell(row=row, column=15, value=com_add_xls(imp))
    ws.cell(row=row, column=16, value=f"=D{row}-(E{row}+F{row}+G{row}+H{row}+M{row}+N{row}+O{row})-{P['AP_ANN']}-{P['MNF_DIP']}*12")
    for col in range(2, 17):
        ws.cell(row=row, column=col).number_format = MONEY
    ws.cell(row=row, column=16).font = Font(bold=True)
ws.freeze_panes = "B3"

wb.save(XLSX)
print("OK xlsx:", XLSX)
print({k: v for k, v in P.items()})

# ------------------------------------------------------------ HTML data sync
def inject_html_data():
    import json
    prov = {}
    for (reg, com), d in ADL._COMUNI.items():
        prov[d['prov']] = reg
    reg_s = {}
    for ente, d in ADL._REGIONALI.items():
        reg_s[ente] = {"r": d['regola'],
                       "b": [[s if s < 1e14 else 0, a] for s, a in d['brackets']],
                       "det": d['detr'], "dda": d['detr_da'], "da": d['detr_a'], "nota": d['nota']}
    rows = []
    for (reg, com), d in sorted(ADL._COMUNI.items()):
        b = "|".join("{:.4f},{:.0f}".format(a, s) for s, a in d['brackets'])
        fl = []
        if "2025" in d['nota']:
            fl.append("y25")
        if "particolari" in d['nota'] or "incoerente" in d['nota'] or "non riconosciuta" in d['nota']:
            fl.append("!")
        if d['regola'] == "n":
            fl.append("n")
        rows.append(f"{reg}|{com}|{d['prov']}|{d['regola']}|{b}|{d['esenzione']:.0f}|{','.join(fl)}")
    html = open(HTML, encoding="utf-8").read()
    start = html.index("/*__COMUNI_DATA_START__*/")
    END = "/*__COMUNI_DATA_END__*/"
    end = html.rindex(END) + len(END)  # consuma TUTTI i marker END (idempotenza)
    nuevo = (f"/*__COMUNI_DATA_START__*/\nconst PROV_RAW={json.dumps(prov, ensure_ascii=False, sort_keys=True)};\n"
             f"const REG26={json.dumps(reg_s, ensure_ascii=False, sort_keys=True)};\n"
             f"const COMUNI={json.dumps(rows, ensure_ascii=False)};\n"
             + END)
    tmp = HTML + ".tmp"  # scrittura atomica: mai troncare l'HTML prima della concatenazione completa
    open(tmp, "w", encoding="utf-8").write(html[:start] + nuevo + html[end:])
    os.replace(tmp, HTML)
    print("OK html data:", len(rows), "comuni")


if __name__ == "__main__":
    inject_html_data()
