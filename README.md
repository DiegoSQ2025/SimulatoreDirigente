# Simulatore Cedolino Dirigente — CCNL Terziario, Distribuzione e Servizi (2026)

Tool di simulazione **dal lordo al netto** per **dirigenti del CCNL Terziario, Distribuzione e Servizi**
(Commercio — Confcommercio–Manageritalia, Testo Unico 12/4/2023). Due versioni con lo stesso motore di calcolo:

- **`simulatore_dirigente.html`** — versione browser autonoma: aprire con doppio clic (nessuna installazione, nessuna dipendenza).
- **`Simulatore_Dirigente_CCNL_Commercio.xlsx`** — fogli di calcolo con formule vive (Excel/LibreOffice).

## Regole dirigente modellate (vigenza 2026)

| Voce | Regola | Base 2026 |
|---|---|---|
| Retribuzione | MINIMO + sup.contrattuale + SMM 2016 + superminimo individuale | 4.660 + 1.317,48 + 350 + S.MINIMO → **RME 7.984,29** |
| Retribuzione giornaliera | RME ÷ 26 (art. 4 c.6 TU) | 307,08791 |
| Mensilità | **14** (art. 8 TU: quattordicesima giugno + tredicesima dicembre) | RAL = RME × 14 |
| IVS INPS | 9,19% + 1% aggiuntiva sopra 56.224 € e **fino al massimale 122.295 €** (Circ. INPS 6/2026: oltre il massimale nessuna contribuzione IVS, 1% compreso — verificato con superminimo ad personam 1.000/1.500/2.500 €/mese; nel MESE l'1% resta dovuto come da cedolino, conguaglio annuo Zucchetti) (dirigenti: niente CIGS/FIS/EST) | base mensile **arrotondata all'euro** (confermato su DUE cedolini: 8.291,38→8.291,00 e 13.612,95→13.613,00: è arrotondamento, non troncamento) |
| [Fondo Mario Negri](https://www.fondonegri.it/) (art. 27, previdenza compl.) | su convenzionale 59.224,54 €/anno; **dirigente 2% da 1/1/2026** (era 1%: 592,25 → 1.184,49 €/anno); **regime deroga art. 20 c.7 D.Lgs. 252/2005: deducibile senza tetto** | 98,71 €/mese |
| [FASDAC "Mario Besusso"](https://www.fasdac.it/) (art. 29, sanità) | su convenzionale 45.940 €/anno; dirigente 1,87%; azienda 5,29% + 2,78% (aziende NON aderenti: 3,08%); non deducibile | 71,59 €/mese |
| [Associazione A. Pastore](https://www.associazionepastore.it/) (art. 28) | dirigente 464,81 €/anno, non deducibile ma **detraibile 19% max 530 €** (L. 124/2013) a conguaglio; azienda 4.856,45 €/anno (base 4.706,45 + infortuni 560 da Q4 2025) | 38,73 €/mese |
| Fondo formazione Mario Negri | 10,83 €/mese dirigente + 24,17 azienda (valori cedolino, possibili aggiornamenti) | 10,83 €/mese |
| IRPEF 2026 | scaglioni 23% / 33% / 43% (L. 199/2025); detrazione art. 13 sul RLD (0 oltre 75.000 €); **IRPEF del mese = IRPEF annua ÷ 14** | — |
| Addizionali | elenchi generali MEF 2026: 21 enti regionali + 7.897 comuni; **rate mensile = annua ÷ 12** | Lazio 1,73/3,33% ; Roma 0,9% ≤14.000 |

**SUP.CONTR vs S.MINIMO — differenza chiave**

- **SUP.CONTR (minimo contrattuale integrativo)**: voce *di contratto*, si somma al MINIMO art. 5 ed è uguale per tutti i dirigenti (parte tabellare del TU/allegato C) — 1.317,48 €/mese nel cedolino 01. **Nel cedolino 03 è assente**: se la tua azienda mostra voci tipo "AUM.RETR." o "Ad Pers.", poni SUP.CONTR a 0 e sommale al campo S.MINIMO (sono componenti individuali). **NON è un superminimo**: l'unico superminimo è l'S.MINIMO.
- **S.MINIMO (superminimo individuale)**: il superminimo *personale* negoziato tra azienda e singolo dirigente → **la variabile da inserire manualmente** (1.656,81 €/mese nel cedolino 01; 7.847,14 €/mese nel cedolino 03 = AUM.RETR. 500 + S.MINIMO 6.489,99714 + Ad Pers. 857,14286; può essere 0).

La pagina mostra ora entrambe le RAL separatamente: **RAL base di contratto** = (minimo + SUP.CONTR + SMM 2016) × 14 (88.584,72 € nel cedolino 01; 70.140,00 € nel cedolino 03) e **superminimo individuale** × 14 (23.195,28 € ced.01; 109.859,96 € ced.03) → RAL di fatto 111.780,00 € (ced.01) e 179.999,96 € (ced.03). Le giornate standard del mese dirigente sono **26** (art. 4 c.6 TU: ogni giornata vale RME ÷ 26; ferie e festività le sostituiscono una a una).

**Fringe benefit (art. 51 TUIR) e indennità personale in denaro: due voci concettualmente diverse**, entrambe variabili manuali:

- **fringe benefit in natura** (136,40 €/mese nel cedolino 01): beneficio in specie (art. 51), tassabile a IRPEF ma **non incassato** — fuori dalle competenze e dalla base IVS (entra però nella base TFR);
- **indennità personale in denaro** (es. autovettura 261,30 €/mese nel cedolino 03, voce Z.000463): NON è un fringe benefit — è un emolumento/compensazione monetaria (tipicamente opt-out dell’auto aziendale), retribuzione piena: entra **in competenze, base IVS, IRPEF e base TFR** come lo stipendio.

Minimi CCNL del triennio già pronti: 4.660 (2026) · 4.920 (2027) · 5.140 (2028) (aumenti art. 6 TU già inclusi, rinnovo 5/11/2025).

## Riconciliazione con il cedolino Zucchetti di riferimento

Il preset della pagina replica **al centesimo** il cedolino dirigente Zucchetti allegato (Liv. 5, Lazio/Roma, GG 25+1+1, fringe 136,40 €):

- competenze **8.291,38** = 25 gg × 307,08791 + ferie 307,09 + festività 307,09;
- trattenute **4.107,91** = IVS 761,94 (base 8.291,00: arrotondata all'euro, classe 8.291,00 → 761,94 come Zucchetti) + agg. 36,06 + Negri 98,71 + FASDAC 71,59 + Pastore 38,73 + formazione 10,83 + IRPEF lorda 2.755,77 + add.reg. residuo 253,95 + add.com. residuo 52,97 + acconto 27,33 + arrotondamento 0,03 (voce ZP9960 dentro le trattenute);
- **NETTO 4.184,00** = 8.291,38 − 4.107,91 + 0,53 (arrotondamento netto).

Differenze note, gestite con **override dedicati** (documentate nel foglio Leggimi):

- **imponibile IRPEF Zucchetti 7.704,45/mese** (107.862,30 annualizzato): il gestionale annualizza con conguagli propri non ricostruibili → usare l'override "IRPEF lorda del mese" = 2.755,77;
- addizionali 2025 "Residuo" (253,95 + 52,97) e acconto 27,33: saldi di conguaglio 2025, input facoltativi;
- quota TFR a fondi 611,87 (regole proprie del fondo TFR dirigenti, fuori scopo);
- azienda Pastore da cedolino 392,20/mese vs tabella ufficiale 4.856,45/anno (informative, non toccano il netto).

## Secondo caso di verifica — cedolino "dirigente 03" (stesso datore, superminimo alto)

Il modello si replica **al centesimo** anche sul cedolino di un secondo dirigente (Zucchetti, Lazio/Roma, 26 gg retribuiti + 1 festività non goduta):

- intestazione: MINIMO 4.660 + AUM.RETR. 500 + SMM 2016 350 + S.MINIMO 6.489,99714 + Ad Pers. 857,14286 = **RME 12.857,14** = 494,50538 × 26 gg (niente SUP.CONTR);
- voci: 26 gg × 494,50538, festività non goduta 494,51, **indennità autovettura incassata 261,30** → **TOTALE COMPETENZE 13.612,95**;
- base IVS **13.613,00** = competenze ARROTONDATE all'euro (prova definitiva del comportamento ROUND, non FLOOR):
  IVS 1.251,03 calcolate dal modello (nel cedolino 1.251,04: scarto di 0,01 del gestionale, assorbito negli arrotondamenti ZP9960)
  + agg. 89,28 + Negri 98,71 + FASDAC 71,59 + Pastore 38,73 + formazione 10,83 + IRPEF lorda 4.810,85 (override) +
  add.reg. 345,82 + add.com. 64,12 + acconto 36,43 + arrotond. 0,14 = **TOTALE TRATTENUTE 6.817,53**;
- **NETTO 6.796,00** = 13.612,95 − 6.817,53 + 0,58 ✓; RAL 179.999,96 (base contratto 70.140,00 + superminimo individuale 109.859,96).

Nel simulatore: supc=0, supi=7847.14, ggRet=26, ggFer=0, ggFest=1, fringe in natura 0, indennità in denaro 261,30,
override IRPEF lorda 4.810,85, residui 345,82/64,12, acconto 36,43, arrotondamenti 0,14/0,58.

## Struttura dell'Excel (5 fogli)

1. *Leggimi* — istruzioni, ipotesi di calcolo, fonti, riconciliazione cedolino, disclaimer.
2. *Parametri* — tutte le variabili (celle gialle = modificabili): retribuzione, IVS, fondi Negri/FASDAC/Pastore, scaglioni IRPEF, detrazione art. 13, look-up addizionali, minimi triennio.
3. *Addizionali* — dati MEF 2026: 21 enti e 7.897 comuni.
4. *Simulazione* — input, cedolino del mese, catena annuale (RME→RAL→RLD→imponibile→IRPEF→netto), voci a carico azienda.
5. *Confronto anni* — 2026/2027/2028 calcolati d'ufficio con gli stessi input.

## Riprodurre/sviluppare (per agenti AI e sviluppatori)

- `build_payslip.py` — genera l'xlsx e inietta i dati MEF nell'HTML (tra i marcatori `/*__COMUNI_DATA_START__*/` e `/*__COMUNI_DATA_END__*/`). Single source of truth delle formule Excel.
- `gen_addizionali.py` — rigenera `data/*.csv` dagli elenchi MEF (sorgenti e formato nel docstring).
- `addizionale_lib.py` — regole addizionali MEF (condiviso con il progetto gemello dipendente).
- `test_simulatore.py` — protocollo di verifica obbligatorio: replica Python vs motore `formulas` (confronti cella per cella, tolleranza 0,01 €), scan errori, 14 casi soglia, **verifiche assolute su ENTRAMBI i cedolini Zucchetti (ced.01 e ced.03)**; stampa i benchmark aggiornati. Richiede `pip install formulas openpyxl`.
- `test_html.js` — suite jsdom per l'HTML (richiede `npm i jsdom` in `/tmp`).
- `data/addizionali_comuni_2026.csv` (7.897 comuni) e `data/addizionali_regionali_2026.csv` (21 enti).
- Dettagli operativi e insidie note: leggere **AGENTS.md** prima di qualsiasi modifica.

Comandi:

```bash
python3 build_payslip.py      # genera xlsx + inietta dati MEF nell'HTML
python3 test_simulatore.py    # replica Python vs motore Excel + verifiche Zucchetti
node test_html.js             # suite jsdom sull'HTML
```

## Verifica operativa

`python3 test_simulatore.py` → **TUTTI I TEST PASSATI** (14 casi, 420 confronti cella per cella, 12 soglie addizionali, 38 verifiche Zucchetti: 20 sul cedolino 01 + 18 sul cedolino 03; i casi includono superminimo ad personam a 1.000/1.500/2.500 €/mese).
`node test_html.js` → **TUTTI I TEST HTML PASSATI** (preset cedolino 01, scenario cedolino 03, mese pulito, 26 gg, Lombardia/Milano, anno 2027, sanity NaN, accessibilità label).

Benchmark (2026 Lazio/Roma, GG 25+1+1, fringe 136,40): preset cedolino netto **4.184,00**; senza conguagli netto 4.378,66 (IRPEF 2.557,41 = annua/14, addizionali rateizzate /12); mese pieno 26 gg netto 4.158,89, netto annuo 59.233,39; default RAL 111.780 → netto annuo 58.460,33 (4.175,74/mese); 2027: 60.185,43; 2028: 61.645,12.

## Fonti (vigenza 2026)

[Testo Unico CCNL Dirigenti Terziario 12/4/2023 (Manageritalia)](https://www.manageritalia.it/wp-content/uploads/2024/07/TU-ccnl-Dirigenti-terziario-12-aprile-2023.pdf) (art. 4 c.6, 6, 8, 9, 27, 28, 29, 43) ·
[contratto Commercio, sezione Dirigenti](https://www.manageritalia.it/dirigenti/contratto-commercio/) ·
rinnovo CCNL 5/11/2025 (minimi/aliquote 2026–2028) ·
[fondonegri.it](https://www.fondonegri.it/) (convenzionale 59.224,54, quota dirigente 2% da 2026, azienda 12,86%+2,52%, regime deroga senza tetto) ·
[fasdac.it](https://www.fasdac.it/) (convenzionale 45.940, dirigente 1,87%, azienda 5,29%+2,78%), [associazionepastore.it](https://www.associazionepastore.it/) (464,81 €/anno, detraibilità 19% max 530) ·
L. 30/12/2025 n. 199 (scaglioni 2026) · D.Lgs. 252/2005 art. 20 c.7 · L. 124/2013 ·
MEF Dipartimento delle Finanze — fiscalità locale, elenchi generali addizionali 2026.

> Strumento a scopo informativo: non sostituisce il cedolino ufficiale né un parere di consulente del lavoro.
