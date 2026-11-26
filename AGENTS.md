# AGENTS.md — Istruzioni operative per agenti AI

## Scopo del progetto
Simulatore di cedolino paga (dal lordo al netto) per **dirigenti del_CCNL Terziario, Distribuzione e Servizi (Commercio — Confcommercio–Manageritalia, TU 12/4/2023)**.
Due deliverable sincronizzati, stesso motore di calcolo:
- `Simulatore_Dirigente_CCNL_Commercio.xlsx` — fogli di calcolo con formule vive (generato da `build_payslip.py`, formule Excel native);
- `simulatore_dirigente.html` — versione browser autonoma (vanilla JS, nessuna dipendenza runtime).

Progetto gemello: `DiegoSQ2025/SimulatoreCedolino` (versione dipendente). Condivide `addizionale_lib.py`, `gen_addizionali.py`, `data/*.csv` (copie verbatim). Attenzione: le regole dirigente sono diverse (vedi sotto) — non copiare a mano formule dal progetto dipendente.

## Struttura dei file

### Sorgenti (riproducibilità: MODIFICARE IL GENERATORE, MAI L'XLSX A MANO)
- `build_payslip.py` — genera l'xlsx e inietta i dati MEF nell'HTML tra i marcatori `/*__COMUNI_DATA_START__*/` e `/*__COMUNI_DATA_END__*/`. SINGLE SOURCE OF TRUTH delle formule Excel.
- `addizionale_lib.py` — regole addizionali regionali/comunali (copiato verbatim dal progetto dipendente; non diversificare).
- `gen_addizionali.py` — rigenera `data/*.csv` dagli elenchi MEF.
- `test_simulatore.py` — protocollo di verifica (vedi sotto). `test_html.js` — suite jsdom.
- `data/*.csv` — versionati; le righe che iniziano per `#` sono commenti.

### Excel (5 fogli)
1. **Leggimi** — istruzioni, ipotesi di calcolo, fonti, riconciliazione del cedolino Zucchetti di riferimento, disclaimer.
2. **Parametri** — TUTTE le variabili (celle gialle = input). Righe-chiave (verificare a ogni rebuild):
   - B2-B5: minimo 2026 4.660 · mensilità 14 · giorni retribuzione 26 · scatto pre-2004 258,22
   - B8-B12: INPS 9,19% · soglia agg. 1% (56.224) · aliquota agg. 1% · massimale 122.295 · datore
   - B15-B20: Fondo Mario Negri: base convenzionale 59.224,54 · dirigente 2% · dirigente 2026 (annuo 1.184,49) · fondi doppi · formazione dirigente 10,83 · formazione ditta 24,17
   - B23-B26: FASDAC: base convenzionale 45.940 · dirigente 1,87% · azienda 5,29%+2,78%
   - B29-B31: A. Pastore: annuo dirigente 464,81 · mensile 38,73 · azienda 392,20
   - B34-B38: scaglioni IRPEF 2026 23/33/43 (soglie 28.000 / 50.000)
   - B41-B49: detrazione art. 13 (sul RLD; 0 oltre 75.000 — qui sempre 0)
   - B55-B78: look-up addizionali regionali/comunali (VLOOKUP sul foglio Addizionali)
   - A82:B84: minimi tabellari 2026/2027/2028 (4.660/4.920/5.140 — aumenti art. 6 già nel triennio)
3. **Addizionali** — dati MEF 2026: 21 enti (riga 4) e comuni (da riga 28); defined names `Reg01`…`Reg21` (`RegElenco`).
4. **Simulazione** — input (anno, sup.contr, SMM 2016, S.MINIMO, scatti, GG retribuite/ferie/festività, fringe, regione/comune, override IRPEF/residui/acconto/arrotondamenti) → cedolino del mese → catena annuale (RME→RAL→IVS→RLD→imponibile→IRPEF a scaglioni→netto annuo) + voci a carico azienda. Il cedolino mensile usa le override; la catena annuale usa gli scaglioni.
5. **Confronto anni** — 2026/2027/2028 × 16 colonne con gli stessi input.

### HTML
Struttura: `TAB3` (minimi 2026/2027/2028), `PAR` (inps .0919, aggSog 56224, maxC 122295, mens 14, ng 26, mnBase 59224.54, mnD .02, fdBase 45940, fdD .0187, apAnn 464.81, apM 38.73, apDit 392.20, formazione 10.83/24.17, d1 1.955…), `PROV_RAW`/`REG26`/`COMUNI` (dati MEF iniettati, NON modificare a mano), `irpefLorda` (scaglioni), `detrazioneArt13` (sul RLD), `calc(...)` (ricorda il mensile basato sul FLOOR all'euro + rateizzazione addizionali ÷12 e IRPEF ÷14), `run()` che aggiorna i KPI (`#kpiNettoM`, `#kpiNote`, `#kpiTratt`, `#kpiRAL`, `#kpiNettoA`) e rigenera `#mensile`, `#annuale`, `#confronto`. Menu `#regione`/`#comune` (21 enti, default Lazio/Roma); `#custom` → aliquote manuali `#regPct`/`#comPct`; override: `#irpOv`, `#regRes`, `#comRes`, `#acconto`, `#arrotT`, `#arrotN`. I default della pagina sono impostati sul **preset identico al cedolino Zucchetti** (Liv. 5, GG 25+1+1, fringe 136,40, override IRPEF 2.755,77, residui 253,95/52,97, acconto 27,33, arrotondamenti 0,03/0,53) → netto 4.184,00.
ACCESSIBILITÀ: ogni input/select deve avere una label (`for`/`id`) o `aria-label` (usato per `#regPct`/`#comPct`).

## Regole dirigente vigenti (vigenza 2026 — riconciliate al centesimo con cedolino Zucchetti Dirigente Terziario Liv. 5, Lazio/Roma, GG 25+1+1)
- **Retribuzione in 4 voci** (celle Parametri B2-B5): MINIMO tabellare 2026 **4.660,00** + **SUP.CONTR = superminimo CONTRATTUALE nazionale** (Allegato C TU: elemento tabellare, uguale per tutti i dirigenti; 1.317,48) + **SMM 2016 350,00** (transitorio nazionale) + **S.MINIMO = superminimo INDIVIDUALE personale** (negoziabile azienda+dirigente, LA VARIABILE da inserire; default cedolino 1.656,80572, display 1.656,81; anche 0) → **RME 7.984,29**.
- **26 GIORNATE standard al mese** (art. 4 c.6 TU): retribuzione giornaliera = RME/26 (307,08791); ferie e festività sostituiscono le giornate una a una.
- **RAL in due righe distinte** (blocco annuale, HTML ed Excel): `Retr. base di contratto annua CCNL` = (minimo+SUP.CONTR+SMM 2016)×14 (88.584,72 nel cedolino) e `Superminimo individuale annuo` = S.MINIMO×14 (23.195,28); `RAL — retribuzione annua effettiva` = somma (111.780,00). NON nascondere la scomposizione: è richiesta dall'utente.
- **14 mensilità** (art. 8 TU): 12 + quattordicesima a giugno + tredicesima a dicembre. RAL = RME × 14.
- **Scatti di anzianità abrogati dal 1/7/2004** (art. 9 TU): solo residui pre-2004 (+258,22 €/mese), input default 0.
- **IVS INPS dirigente**: 9,19% + 1% aggregativa su (RAL − 56.224), massimale 122.295. NO CIGS/FIS/EST (dirigente non è dipendente di livello).
- **Fondo Mario Negri (art. 27, previdenza complementare)**: convenzionale **59.224,54 €/anno** (uguale per tutti). Dirigente **2% dal 1/1/2026** (era 1%: 592,25 → 1.184,49 €/anno; fonti fondonegri.it, payroll.it). Azienda 12,86% + 2,52% integrativo 2026. **Regime deroga art. 20 c.7 D.Lgs. 252/2005: contributi completamente deducibili senza tetto 5.300 €, confermato dal fondo (fondonegri.it).**
- **FASDAC "Mario Besusso" (art. 29, sanità integrativa)**: convenzionale **45.940 €/anno**. Dirigente **1,87%** = 71,59/mese (invariato 2026). Azienda 5,29% (include premio Long Term Care 206,60 €/anno) + 2,78% gestione pensionati. Nota: per aziende NON aderenti alle organizzazioni imprenditoriali il contributo pensionati è **3,08%** (non 2,78%). NON deducibile.
- **Associazione Antonio Pastore (art. 28)**: dirigente **464,81 €/anno** (38,73/mese) NON deducibile ma **detraibile 19% max 530 €** (L. 124/2013, da periodo d'imposta 2014) a conguaglio — non modellata automaticamente, richiede oneri (input). Azienda: da cedolino 392,20/mese; ufficiale 4.856,45 €/anno (base Manageritalia 4.706,45 + infortuni 560,00 da Q4 2025).
- **Fondo formazione Mario Negri**: 10,83 dip + 24,17 ditta (valori cedolino, possibili aggiornamenti col rinnovo).
- **IRPEF 2026**: 23% ≤28.000 · 33% 28.000–50.000 · 43% >50.000 (L. 199/2025). Detrazione art. 13 sul RLD (qui sempre 0: RLD > 75.000). **Cuneo L. 207/24 (TI + somma integrativa) NON applicabile al dirigente (RLD > 75k)** — non modellato, per costruzione a scaglioni standard.
- **IRPEF lorda del mese = IRPEF annua (scaglioni) ÷ 14**, NON ÷ 12. Modello annuale: RAL → ivsA (su massimale) → aggA → negriA (deducibile) → RLD → imp = RLD + fringe×12 → scaglioni → IRPEF netta. FASDAC e Pastore NON deducibili (fuori dalla catena RLD).
- **Rateizzazione mensile addizionali = annua ÷ 12** (12 rate, anche con 14 mensilità). Le buste 13ª/14ª portano solo il netto in busta (nessuna addizionale extra).
- NON inclusi: TFR (art. 43 TU: devoluzione o liquidazione a tassazione separata — fuori scopo), trattamento integrativo/somma cuneo L. 207/24, detrazioni per casi particolari (vedi Leggimi).

## Riconciliazione cedolino Zucchetti 01 (al centesimo)
MINIMO 4.660,00 + SUP.CONTR 1.317,48 + SMM 2016 350,00 + S.MINIMO 1.656,81 = **RME 7.984,29** = 307,08791 × 26 gg.
COMPETENZE: retribuzione 25 gg 7.677,20 + ferie 307,09 + festiv. 307,09 = **8.291,38**.
Trattenute: IVS **761,94** + agg **36,06** + Negri **98,71** + FASDAC **71,59** + Pastore **38,73** + formazione **10,83** + IRPEF lorda **2.755,77** (override: annualizzazione Zucchetti) + add.reg. residuo **253,95** + add.com. residuo **52,97** + acconto addiz. **27,33** + arrotondamento trattenute **0,03** = **4.107,91**.
**Netto = 8.291,38 − 4.107,91 + 0,53 (arrotondamento netto) = 4.184,00** ✓.

### Secondo caso di verifica: cedolino Zucchetti "dirigente 03" (stesso datore cod. azienda 000294, RME 12.857,14)
- **Intestazione SENZA SUP.CONTR**: MINIMO 4.660 + AUM.RETR. 500 + SMM 2016 350 + S.MINIMO 6.489,99714 + Ad Pers. 857,14286 = 12.857,14 = 494,50538 x 26 gg. Nel simulatore: supc=0 e supi=7.847,14 (AUM.RETR. + S.MINIMO + Ad Pers. sono componenti individuali da sommare). NON dichiarare piu' SUP.CONTR "uguale per tutti"
- **Fringe INCASSATO** (nuovo input `fringeCash`/"Fringe benefit INCASSATO"): l'indennita' autovettura 261,30 (Z.000463) entra in COMPETENZE, base IVS e IRPEF (diverso dal fringe in natura del ced.01, che resta fuori cassa e IVS). comp = retr. + ferie + fest. + fringeCash
- Riconciliato: comp 13.612,95 (26gg + 1 fest. 494,51 + 261,30), base IVS 13.613,00 (ROUND), IVS 1.251,03 (cedolino 1.251,04, scarto 0,01), agg 89,28, trattenute 6.817,53 (con arrotT 0,14 = ZP9960 0,13 + 0,01 IVS), NETTO 6.796,00; RAL 179.999,96 = ralBase 70.140,00 + supiA 109.859,96
- `test_simulatore.py`: caso in CASES ("cedolino 03") + blocco `ZUCCHETTI_03` (18 verifiche, dict D03); `test_html.js`: scenario 1bis ced.03.
- Peculiarita': nei DUE cedolini Zucchetti le trattenute residue hanno un +0,01 (F09110 residuo 345,81 vs trattenuta 345,82) e l'imponibile IRPEF porta un FISSO +309,78/mese (probabile valuta auto forfettizzata della flotta aziendale) non modellabile: coperti da override

### Non riconciliabili (documentati in Leggimi e nel codice, gestiti con override)
- **Imponibile IRPEF Zucchetti 7.704,45/mese** (107.862,30 annualizzato vs modello 101.404,17: gap ~586 €/mese) — Zucchetti annualizza con conguagli propri; si replica con l'override "IRPEF lorda del mese" = 2.755,77.
- **Quota TFR a fondi 611,87** vs retrib. utile TFR 8.819,98 — regole proprie del fondo TFR dirigenti (probabile componente infortuni Pastore 560).
- **Pastore ditta 392,20/mese** vs ufficiale 4.856,45/anno (404,70/mese): il cedolino usa una componente non allineata all'ultima tabella; gestito in nota.
- **Formazione Negri 10,83/24,17**: possibili aggiornamenti col rinnovo CCNL.

## Insidie note (bug già risolti o scoperti — NON reintrodurli)
1. **IRPEF mese ≠ annua/12**: per i dirigenti (14 mensilità) è **annua/14**. Per le addizionali invece è **annua/12** in entrambe le versioni (dipendente e dirigente è /12). Confondere i due divisorismi altera il netto di centinaia di euro
2. **BASE IVS mensile ARROTONDATA all'euro (Zucchetti)**: il gestionale applica la percentuale sulla base mensile arrotondata all'euro (ced.01: 8.291,38 → 8.291,00 → 761,94; ced.03: 13.612,95 → 13.613,00 → 8.928,00 x 1% = 89,28). Il cedolino 03 dimostra che è ROUND è non FLOOR (13.612,95 troncato sarebbe 13.612): in Excel `=ROUND(ROUND(base,0)*aliquota,2)`; in Python `int(comp + 0.5)`; in JS `Math.round(comp)`. Nota: su base 13.613,00 il cedolino 03 mostra IVS 1.251,04 mentre 13.613 x 9,19% = 1.251,0347 → 1.251,03 (scarto 0,01 del gestionale, assorbito negli arrotondamenti ZP9960: nel caso ced.03 l'input arrotT usato è 0,14 non 0,13)
3. **Arrotondamenti Zucchetti**: ZP9960 arrotondamento 0,03 è una voce DENTRO LE TRATTENUTE; l'arrotondamento netto 0,53 si aggiunge al netto. Servono due input separati
4. **S.MINIMO esatto**: 1.656,80572 (display 1.656,81) — usare il valore pieno per combaciare col rate 307,08791 × 26
5. **Pastore annuo**: usare **464,81** (AP_ANN/apAnn) e non 38,73×12 = 464,76 — altrimenti netto annuo di 5 centesimi sbagliato
6. **Etichette che iniziano per "="**: openpyxl le scrive come formule — nel test cercare prefix diversa (es. "= TOTALE COMPETENZE" funziona nel test perché l'etichetta è scritta come stringa con "=" davanti? No: in questo progetto le etichette mensili usa "= TOTALE COMPETENZE" ecc. — assicurarsi che openpyxl le scriva come stringhe (le etichette con "=" iniziale su XLSX devono essere ante con `'`: check il builder)
7. **Indirizzi riga NON stabili**: test e riferimenti sempre via ricerca-etichetta (`row_of`/startswith + asserzione unicità), mai righe hardcoded
8. **Data validation**: `formula1` NON deve iniziare con `=`
9. **Catene IF lunghe**: costruire per accumulazione con assert di bilanciamento parentesi (pattern del progetto dipendente)
10. **Iniezione dati HTML**: marker END consumati con `rindex(END) + len(END)`, scrittura atomica (`.tmp` + `os.replace`); dopo ogni rebuild verificare 1 marker START + 1 END e `git diff` vuoto sull'HTML (idempotenza)
11. **jsdom**: le trattenute nel DOM sono negativo (`-761,94`); nelle asserzioni usare `Math.abs(rowValue(...))`. `aria-label` per `#regPct`/`#comPct`
12. **Monitoraggi**: delibere comunali 2026 (`gen_addizionali.py`), Legge di Bilancio 2027, circolare INPS valori gennaio, tabelle fondi dopo ogni rinnovo CCNL (fondonegri.it, fasdac.it, associazionepastore.it)

## Protocollo di verifica (obbligatorio prima di qualsiasi push)
1. **`python3 test_simulatore.py`** deve stampare "TUTTI I TEST PASSATI": replica Python (`py_calc`, addizionale_lib) vs motore `formulas` sull'xlsx (11 casi × 30 output, tolleranza 0,01 €), 12 casi soglia addizionali, **38 verifiche assolute Zucchetti (20 ced.01 + 18 ced.03)** (comp 8.291,38 · IVS 761,94 · Negri 98,71 · FASDAC 71,59 · Pastore 38,73 · IRPEF 2.755,77 · trattenute 4.107,91 · netto 4.184,00 · RAL 111.780 · regA 3.136,76 · comA 912,64)
2. **Scan errori**: nessuna cella `#REF!`/`#VALUE!`/ecc.
3. **`node test_html.js`** (jsdom da `/tmp/node_modules`): 21 enti, default Lazio/Roma, preset Zucchetti completo al centesimo, mese pulito, 26 gg, Lombardia/Milano, 2027, nessun NaN/Infinity, accessibilità label
4. **Benchmark dopo ogni rebuild** (2026 Lazio/Roma, GG 25+1+1, fringe 136,40): preset pubblico netto 4.184,00; auto (senza conguagli) netto 4.378,66 (IRPEF 2.557,41 = annua/14, add.reg. 261,40 = annua/12); mese pieno 26 gg netto 4.158,89; RAL 111.780, imponibile 101.404,17, netto annuo 58.460,33 (→ 4.175,74/mese); 2027 netto annuo 60.185,43; 2028 netto annuo 61.645,12

## Convenzioni e stile
- Celle input gialle `FFF2CC` grassetto; formati `#,##0.00 "€"`, `0.00%`
- Etichette di riga stabili e univoche; niente dipendenze runtime (HTML vanilla; Excel formule native); Python: solo `openpyxl` (+ `formulas` per i test)
- Commit sintetici in italiano, push modulare su `main` (repo nuovo, primo commit diretto) — successivi rilasci via branch `vibe/*` + PR
- Disclaimer normativo obbligatorio in Leggimi/HTML/README
- Dati addizionali: fonte MEF; aliquote in frazione decimale (0.009 = 0,90%)
