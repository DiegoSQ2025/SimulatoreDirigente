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
| IVS INPS | 9,19% + 1% aggregativa sopra 56.224 €, massimale 122.295 (dirigenti: niente CIGS/FIS/EST) | base mensile **arrotondata all'euro** (comportamento Zucchetti) |
| [Fondo Mario Negri](https://www.fondonegri.it/) (art. 27, previdenza compl.) | su convenzionale 59.224,54 €/anno; **dirigente 2% da 1/1/2026** (era 1%: 592,25 → 1.184,49 €/anno); **regime deroga art. 20 c.7 D.Lgs. 252/2005: deducibile senza tetto** | 98,71 €/mese |
| [FASDAC "Mario Besusso"](https://www.fasdac.it/) (art. 29, sanità) | su convenzionale 45.940 €/anno; dirigente 1,87%; azienda 5,29% + 2,78% (aziende NON aderenti: 3,08%); non deducibile | 71,59 €/mese |
| [Associazione A. Pastore](https://www.associazionepastore.it/) (art. 28) | dirigente 464,81 €/anno, non deducibile ma **detraibile 19% max 530 €** (L. 124/2013) a conguaglio; azienda 4.856,45 €/anno (base 4.706,45 + infortuni 560 da Q4 2025) | 38,73 €/mese |
| Fondo formazione Mario Negri | 10,83 €/mese dirigente + 24,17 azienda (valori cedolino, possibili aggiornamenti) | 10,83 €/mese |
| IRPEF 2026 | scaglioni 23% / 33% / 43% (L. 199/2025); detrazione art. 13 sul RLD (0 oltre 75.000 €); **IRPEF del mese = IRPEF annua ÷ 14** | — |
| Addizionali | elenchi generali MEF 2026: 21 enti regionali + 7.897 comuni; **rate mensile = annua ÷ 12** | Lazio 1,73/3,33% ; Roma 0,9% ≤14.000 |

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
- `test_simulatore.py` — protocollo di verifica obbligatorio: replica Python vs motore `formulas` (confronti cella per cella, tolleranza 0,01 €), scan errori, 14 casi soglia, **18 verifiche assolute sul cedolino Zucchetti**; stampa i benchmark aggiornati. Richiede `pip install formulas openpyxl`.
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

`python3 test_simulatore.py` → **TUTTI I TEST PASSATI** (10 casi, 280 confronti cella per cella, 12 soglie addizionali, 18 verifiche Zucchetti).
`node test_html.js` → **TUTTI I TEST HTML PASSATI** (preset Zucchetti, mese pulito, 26 gg, Lombardia/Milano, anno 2027, sanity NaN, accessibilità label).

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
