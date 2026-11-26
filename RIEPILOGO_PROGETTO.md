# RIEPILOGO PROGETTO — Simulatore Dirigente CCNL Terziario

> **Documento di sintesi** per sviluppi futuri e agenti AI. Le istruzioni operative dettagliate (protocollo di verifica, insidie tecniche, convenzioni) sono in **[AGENTS.md](AGENTS.md)**; il manuale utente in **[README.md](README.md)**.
>
> Ultimo aggiornamento: 26/11/2025 · Stato: **stabile, tutti i test verdi** · Branch: `main`

---

## 1. Obiettivo

Calcolatore **dalla RAL al NETTO** per **dirigenti del CCNL Terziario, Distribuzione e Servizi (Commercio — Confcommercio–Manageritalia, Testo Unico 12/4/2023)**, pulito ed efficace per le simulazioni:

- l'utente inserisce **RAL (scomposta) + benefit + indennità** e ottiene subito il netto, come un vero ufficio paghe;
- il mese è **sempre pieno a 26 GG** (art. 4 c.6 TU) nella pagina HTML — niente variabili di giorni (richiesta esplicita dell'utente);
- il KPI principale in alto è il **netto mensile medio = netto annuale ÷ 14**;
- la replica puntuale dei cedolini reali (ferie, festività, conguagli) resta nel foglio Excel e nelle suite di test.

**Filosofia**: un unico motore di calcolo, tre implementazioni sincronizzate (Python ↔ Excel ↔ JS), verifiche al centesimo contro cedolini reali.

---

## 2. Deliverable

| File | Ruolo | Note |
|---|---|---|
| `Simulatore_Dirigente_CCNL_Commercio.xlsx` | Foglio di calcolo (5 fogli), formule Excel native | Generato da `build_payslip.py` — **mai modificato a mano** |
| `simulatore_dirigente.html` | Simulatore browser autonomo (vanilla JS, ~1 MB con dati MEF iniettati: 7.897 comuni) | Nessuna dipendenza runtime, funziona offline |
| `build_payslip.py` | **Generatore = single source of truth** delle formule Excel + iniezione dati MEF nell'HTML | Idempotente, scrittura atomica |
| `test_simulatore.py` | Suite Python: py_calc vs motore `formulas` + verifiche benchmark reali | 15 casi, 480 confronti, 12 soglie, 59 verifiche benchmark reali |
| `test_html.js` | Suite jsdom sull'HTML (77 controlli) | Richiede jsdom in `/tmp/node_modules` |
| `addizionale_lib.py`, `gen_addizionali.py`, `data/*.csv` | Regole e dati MEF addizionali (21 enti regionali, 7.897 comuni) | Copie verbatim dal progetto gemello |
| `README.md` | Manuale utente + benchmark | |
| `AGENTS.md` | Protocollo operativo per agenti AI (obbligatorio prima di ogni push) | |
| `RIEPILOGO_PROGETTO.md` | Questo documento | |

---

## 3. Fonti ufficiali verificate

| Fonte | Uso |
|---|---|
| [TU CCNL Dirigenti Terziario 12/4/2023 (Manageritalia)](https://www.manageritalia.it/wp-content/uploads/2024/07/TU-ccnl-Dirigenti-terziario-12-aprile-2023.pdf) | Art. 4 c.6 (giornaliera = mensile/26), art. 5 (minimo), art. 8 (14 mensilità), art. 9 (scatti abrogati 1/7/2004), art. 27-29 (fondi), art. 43 (TFR) |
| [ilccnl.it — tabelle Dirigenti Commercio](https://ilccnl.it/contratto/ccnl/dirigenti---commercio) | Minimo 4.660 (2026) / 4.920 (2027) / 5.140 (2028); divisore giornaliero 26; unico minimo "Dirigente" senza scaglioni di livello |
| [fondonegri.it](https://www.fondonegri.it/) | Fondo Mario Negri: basi e aliquote 2026 |
| [fasdac.it](https://www.fasdac.it/) | FASDAC "Mario Besusso": basi e aliquote 2026 |
| [associazionepastore.it](https://www.associazionepastore.it/) | Associazione A. Pastore: importi fissi 2026 |
| Elenchi MEF addizionali 2026 | Via `gen_addizionali.py` → `data/*.csv` |
| Cedolini reali ("dirigente 01", "dirigente 03") | Riconciliazione al centesimo (benchmark assoluti nei test) |
| Circ. INPS 6/2026 (art. 3-ter D.L. 384/1992) | Massimale IVS 122.295: nessuna contribuzione oltre, **1% incluso** |

---

## 4. Regole di calcolo verificate (vigenza 2026)

### 4.1 Retribuzione (4 voci → RME)
```
RME = MINIMO + SUP.CONTR + SMM 2016 + S.MINIMO + scatti pre-2004
```
| Voce | Natura | Default (ced.01) |
|---|---|---|
| MINIMO (art. 5) | tabellare CCNL, dal menu anno: 4.660 (2026) / 4.920 (2027) / 5.140 (2028) | 4.660,00 |
| **SUP.CONTR** | **superminimo da contratto** — quota fissa che supera il minimo tabellare, concordata individualmente nella lettera di assunzione per raggiungere la RAL pattuita (definizione utente) | 1.317,48 |
| SMM 2016 | aumento nazionale transitorio 2016, fisso per tutti | 350,00 |
| **S.MINIMO** | **superminimo individuale** — negoziato col singolo dirigente | 1.656,80572 |
| scatti pre-2004 | solo dirigenti assunti prima del 1/7/2004 (+258,22/mese) | 0 |

- **Nell'HTML SUP.CONTR e S.MINIMO sono UNITE in un'unica variabile "Superminimo ad personam"** (default 0: tutti gli importi variabili partono da zero, li inserisce l'utente; per replicare ced.01 inserire 2.974,28572, per ced.03 7.847,14). In Excel/py restano due voci distinte.
- **Giornaliera** = RME ÷ 26 (art. 4 c.6 TU) — ced.01: 307,08791.
- **RAL** = RME × **14 mensilità** (art. 8: 13ª dicembre, 14ª giugno). Presentata SEMPRE in due righe distinte (richiesta utente): *base di contratto* (minimo+SUP.CONTR+SMM)×14 e *superminimo individuale*×14.
- **Ced.03 (secondo dirigente)**: senza SUP.CONTR; intestazione = 4.660 + AUM.RETR. 500 + SMM 350 + S.MINIMO 6.489,99714 + Ad Pers. 857,14286 = RME 12.857,14. Nel simulatore: supc=0, supi=7.847,14.

### 4.2 Contributi (base mensile = ROUND all'euro delle competenze, dimostrato da ced.03)
| Voce | Regola | Valore ced.01 |
|---|---|---|
| IVS INPS 9,19% | su base mensile **arrotondata all'euro** (ROUND, non FLOOR) | 761,94 (base 8.291,00) |
| IVS agg. 1% | su (mensile − 4.685,33 = 56.224/12) | 36,06 |
| **Massimale 122.295** | **capa SIA il 9,19% annuo CHE l'1% annuo** (Circ. INPS 6/2026); il mensile resta non cappato come da cedolini — conguaglio annuale dominio del gestionale | — |
| Mario Negri 2% | su convenzionale fisso 59.224,54 (→ 4.935,38/mese); dal 1/1/2026 (era 1%); **deducibile integralmente** (art. 20 c.7 D.Lgs. 252/2005, regime deroga, senza tetto 5.300) | 98,71 |
| FASDAC 1,87% | su convenzionale 45.940; **non deducibile** | 71,59 |
| A. Pastore | fisso 464,81 €/anno (38,73/mese); non deducibile ma detraibile 19% max 530 € a conguaglio (via "oneri") | 38,73 |
| Formazione Negri | 10,83 dip + 24,17 ditta (valori cedolino) | 10,83 |
| Azienda | Negri 12,86%+2,52% · FASDAC 5,29%+2,78% (3,08% se non aderente) · Pastore 392,20/mese da cedolino (ufficiale 4.856,45/anno) | — |

- **Dirigenti: NO CIGS/FIS/EST.**

### 4.3 Fisco
- **IRPEF 2026**: scaglioni 23% ≤28.000 · 33% 28.000–50.000 · 43% >50.000 (L. 199/2025); detrazione art. 13 sul RLD (0 oltre 75.000 — sempre il caso dirigente).
- **Catena annuale**: RAL → ivsA (sul massimale) → aggA (sul massimale) → Negri (deducibile) → RLD → imponibile = RLD + fringe natura×12 + indennità denaro×12 → scaglioni → IRPEF netta → netto annuo.
- **IRPEF lorda del mese = annua ÷ 14** (non ÷12!); **addizionali = annua ÷ 12** (anche con 14 mensilità).
- **Cuneo L. 207/24 non applicabile** al dirigente (RLD > 75k).
- **Addizionali**: da elenchi MEF 2026 (21 enti, 7.897 comuni); verificati Lazio 3,33% / Roma 0,9% (3.136,76 + 912,64 annui sul caso ced.01) e Lombardia/Milano.

### 4.4 Fringe benefit vs indennità personale (distinzione concettuale, richiesta utente)
| | Fringe benefit **in natura** (art. 51 TUIR, ced.01: 136,40) | Indennità personale **in denaro** (ced.03: autovettura 261,30) |
|---|---|---|
| Competenze (cassa) | NO | SÌ |
| Base IVS | NO | SÌ |
| IRPEF | SÌ (sul seuone annuo) | SÌ |
| Base TFR | SÌ | SÌ |

**L'indennità auto NON è un fringe benefit**: è un emolumento/compensazione monetaria (come un rimborso/allowance personale).

### 4.5 Fuori scopo (documentati in Leggimi)
TFR (art. 43: devoluzione a fondi o liquidazione a tassazione separata), trattamento integrativo/cuneo L. 207/24, detrazioni particolari.

---

## 5. Benchmark numerici chiave (mese pieno standard 26 GG, salvo indicazione)

| Caso | Competenze | IVS 9,19% | Trattenute | **Netto mese** | RAL | **Netto annuo** | Netto/14 |
|---|---|---|---|---|---|---|---|
| **Default HTML** (TUTTO a zero: solo minimo 4.660, Lazio/Roma 2026; detrazione art. 13 attiva 1.301,11) | 4.660,00 | 428,25 | 1.962,88 (IRPEF netta) | 2.697,12 (mese tipo) | 65.240,00 | **38.477,92** | **2.748,42** |
| Stile ced.03 (supi 7.847,14 + indennità 261,30) | 13.118,44 | 1.205,54 | 6.755,03 | 6.363,41 | 179.999,96 | **93.186,70** | 6.656,19 |
| Anno 2027 (minimo 4.920) | 8.244,29 | 757,62 | 4.019,86 | 4.224,43 | 115.420,00 | 60.185,43 | 4.298,96 |
| Lombardia/Milano 2026 | 7.984,29 | 733,73 | 3.749,72 | 4.234,57 | 111.780,00 | 60.040,90 | 4.288,64 |
| Superminimo ad personam 1.000 / 1.500 / 2.500 | — | — | — | 3.795,74 / 4.029,54 / 4.492,64 (payslip) | 102.584,72 / 109.584,72 / 123.584,72 | 54.102,44 / 57.419,93 / 64.124,26 | 3.864,46 / 4.101,42 / 4.580,30 |
| Fringe 100/mese × 12 (caso permanente) | invariato | invariato | — | 4.117,80 (vs 4.158,89 senza fringe) | invariata 111.780,00 | 58.666,63 (vs 59.233,39) | +566,76 tasse/anno, vantaggio economico 633,24/anno |

**Repliche esatte dei cedolini** (nelle suite, con giorni reali e override):
- **ced.01** (27 GG pagate 25+1+1, override conguaglio): competenze 8.291,38 → trattenute 4.107,91 → **netto 4.184,00** ✓ (20 verifiche).
- **ced.03** (26 GG + 1 festività, override): competenze 13.612,95 → trattenute 6.817,53 → **netto 6.796,00** ✓ (18 verifiche).
- **elab. controllo gennaio 2027** (visuale annua, superminimo 550,11117 per netto target 44.000): RAL 76.581,55 ✓, IVS 7.037,89 su base annua arrotondata ✓, agg 203,58 ✓, Negri 1.184,49 ✓, IRPEF lorda 21.502,25 = scaglioni 23/33/43 sul suo imponibile 68.144,76 ✓, add.reg. Lazio marg **2.029,22** ✓ e add.com. Roma **613,30** ✓ AL CENTESIMO (terza validazione indipendente), TFR quota 5.672,71 = retrib. utile/13,5 ✓ (21 verifiche). Struttura ridotta dell'allegato: senza FASDAC (859,08), senza Pastore (464,81), senza detrazione art. 13 (522,91), formazione non annualizzata → suo netto 44.000 vs nostro 43.074,79 con gli stessi input, spiegato al centesimo; con il modello completo servono 713,16/mese di superminimo per netto 44.000.
- Aliquota marginale verificata: +500/mese di superminimo → +3.317,49/anno netto = **52,61% marginale** (0,1019 contributi + 0,8981 × 0,4723 fisco, esatto).

---

## 6. Architettura e flusso

```
build_payslip.py  (single source of truth)
   ├─ genera  Simulatore_Dirigente_CCNL_Commercio.xlsx  (5 fogli, formule vive)
   └─ inietta dati MEF in simulatore_dirigente.html  (marker /*__COMUNI_DATA__*/, atomico, idempotente)

Verifica a triangolazione (obbligatoria prima di ogni push):
   py_calc (test_simulatore.py)  ↔  motore `formulas` sull'xlsx  ↔  jsdom sull'HTML
   + verifiche assolute al centesimo sui TRE benchmark reali
```

Regola d'oro: **modificare il generatore, mai l'xlsx a mano**; riferimenti alle celle sempre via ricerca-etichetta (mai righe hardcoded).

---

## 7. Decisioni prese (traccia delle richieste utente)

1. **14 mensilità** per i dirigenti commercio (art. 8 TU) — IRPEF ÷14, addizionali ÷12.
2. **SUP.CONTR = superminimo da contratto** (definizione utente finale): quota fissa oltre il minimo concordata nella lettera di assunzione; prima era stato etichettato "minimo contrattuale integrativo" (commit `2506756`), poi unificato con S.MINIMO nell'unica variabile HTML "Superminimo ad personam" con importi a zero (commit corrente).
3. **RAL in due righe**: base di contratto CCNL e superminimo personale separate, mai sommarle in una voce (commit `38779d5`).
4. **Superminimo, fringe e indennità = variabili a inserimento manuale** con calcoli completi (IVS/IRPEF/TFR come da regole §4.4).
5. **Menu a tendina anno + minimo lordo** (4.660/4.920/5.140, commit `2506756`).
6. **Bug reale trovato e corretto**: IVS agg. 1% annua non cappata al massimale 122.295 → corretto in tutte e tre le implementazioni (commit `2506756`). Nota: anche il repo gemello SimulatoreCedolino ha questo bug, ma è **read-only** su richiesta esplicita.
7. **Terminologia fringe/indennità** corretta ovunque (xls+html, commit `7b25f2e`): l'indennità auto è emolumento in denaro, non fringe benefit.
8. **Terminologia parlata nei parametri** (commit `1c96cbe`): sezioni "Il salario fisso / Benefit e indennità / Dove paghi le tasse / Numeri veri del cedolino".
9. **Fix layout mobile** (commit `71fe93c`): etichette brevi su una riga + `.help` in corpo piccolo; colonne impilate sotto 640px.
10. **Sezione giorni rimossa** (commit `ae8edcb`): mese fisso 26 GG pieni nel HTML; giorni reali solo in Excel/py_calc per la replica dei cedolini.
11. **KPI in alto = netto annuale ÷ 14** (commit `ae8edcb`): vista simulatore RAL→NETTO; il "NETTO DEL MESE" nella tabella resta la vista payslip (addizionali ÷12).
12. **Bug reale trovato e corretto (verifica utente sui fondi)**: la trattenuta IRPEF mensile era la LORDA mentre KPI e prospetto annuo usano la NETTA: con detrazione art. 13 attiva il netto mensile non riconciliava col netto annuale/14 (75,10 €/mese fantasma nel caso 2027 Lazio/Roma). Ora tutti e tre i motori trattenanno la NETTA (lorda/14 − detrazione/14), con righe separate lorda/detrazione/netta e riga-ponte esplicita mese tipo ↔ media annuale (fondi/addizionali su 12 rate). Le repliche dei cedolini reali restano identiche (nei cedolini campione detrazione = 0, RLD > 75.000). Verifica fondi riconfermata al centesimo: Negri 59.224,54/12 × 2% = 98,71, FASDAC 45.940/12 × 1,87% = 71,59, Pastore 464,81/12 = 38,73.
13. **Terzo benchmark (elaborazione di controllo gennaio 2027, netto target 44.000)**: nessun errore nel nostro modello — l'elaborazione esterna conferma al centesimo IVS, agg. 1%, Negri, scaglioni 23/33/43, addizionali Lazio/Roma, 14 mensilità e TFR/13,5. Le differenze sul netto (44.000 vs 43.074,79) sono omissioni della struttura ridotta dell'allegato: senza FASDAC, senza Pastore, senza detrazione art. 13, formazione non annualizzata. Con il modello completo il superminimo per netto 44.000 è 713,16 €/mese.

---

## 8. Elementi aperti e in monitoraggio

### Proposta non ancora confermata dall'utente
Ristrutturazione a **vista annuale RAL come principale** + dettaglio mensile opzionale + input "giornate non retribuite/anno" + netto ÷14 e ÷12 (il punto 10-11 ne implementa solo lo spirito).

### Non riconciliabili con il gestionale reale (documentati, gestiti con override)
- Imponibile IRPEF con **+309,78/mese fisso** in entrambi i cedolini (probabile valuta auto forfettizzata di flotta) — non modellabile.
- Quota TFR a fondi 611,87 vs retrib. utile TFR 8.819,98 (regole proprie del fondo TFR dirigenti).
- Pastore ditta: cedolino 392,20/mese vs ufficiale 4.856,45/anno (404,70/mese).
- Formazione Negri 10,83/24,17: possibili aggiornamenti col rinnovo.
- Quirk del gestionale: ced.03 IVS 1.251,04 vs calcolato 1.251,03 (0,01 assorbito nella voce arrotondamento).

### Aggiornamenti normativi da monitorare
- **Accordo integrativo CCNL previsto per il 15/05/2026** — nuovi valori possibili.
- **Credito welfare contrattuale 2026–2028: minimo 1.500,00 €/anno** (fonte Confcommercio, rinnovo CCNL 5/11/2025; verificato durante il controllo dei fondi dirigenti).
- Delibere comunali 2026 (`gen_addizionali.py`), Legge di Bilancio 2027, circolari INPS di gennaio, tabelle fondi dopo ogni rinnovo (fondonegri.it, fasdac.it, associazionepastore.it).

---

## 9. Per gli agenti AI che riprenderanno il progetto

1. **Leggere prima AGENTS.md**: contiene il protocollo di verifica obbligatorio, le insidie tecniche (bug già risolti — non reintrodurli) e le convenzioni di stile.
2. **Prima di ogni push**: `python3 test_simulatore.py` ("TUTTI I TEST PASSATI — 15 casi, 480 confronti, 12 soglie, 59 verifiche benchmark reali") + `node test_html.js` ("TUTTI I TEST HTML PASSATI" — 77 controlli). Jsdom si trova in `/tmp/node_modules` (reinstallare se la sandbox è stata resettata: `npm i --prefix /tmp jsdom`).
3. **Rebuild Excel** dopo ogni modifica al generatore: `python3 build_payslip.py` (idempotente — `git diff` vuoto sull'HTML dopo il rebuild se non è stato toccato).
4. **Repo gemello `DiegoSQ2025/SimulatoreCedolino`** (versione dipendente): **READ-ONLY, non toccare mai nulla** — directive esplicita dell'utente. Condivide `addizionale_lib.py`/`gen_addizionali.py`/`data/*.csv` ma le regole dirigente sono diverse: non copiare formule tra i due progetti.
5. **Push su `main`**: `git -c credential.helper='!gh auth git-credential' push` (main non protetto; commit sintetici in italiano).
6. **Terminologia vincolante** (richieste esplicite dell'utente, non modificare): SUP.CONTR = "superminimo da contratto"; HTML = unica variabile "Superminimo ad personam — SUP.CONTR + S.MINIMO"; etichette corte "Fringe benefit" e "Indennità denaro"; TUTTI gli importi variabili a zero all'avvio (li inserisce l'utente); mese = 26 GG fissi nella pagina HTML; KPI in alto = netto annuale ÷ 14.

---

## 10. Cronologia dei rilasci (main)

| Commit | Contenuto |
|---|---|
| `544b3cf`, `0bb2504` | Build iniziale: generator, xlsx 5 fogli, HTML con dati MEF, suite test |
| `38779d5` | RAL in due righe, 26 GG espliciti, chiarimento SUP.CONTR vs S.MINIMO |
| `7e59f6a` | Integrazione e validazione cedolino 03 (ROUND base IVS, indennità auto, 18 verifiche) |
| `7b25f2e` | Nomenclatura fringe vs indennità personale corretta (xls+html) |
| `1c96cbe` | Parametri con terminologia parlata, sezioni, fix visibilità |
| `2506756` | Menu anno+minimo art. 5, SUP.CONTR rinominato, **cap massimale IVS agg. 1%**, casi superminimo 1.000/1.500/2.500 |
| `992c548` | Caso fringe 100/mese × 12 permanente (suite: 15 casi, 480 confronti) |
| `71fe93c` | Fix layout mobile: etichette compatte + `.help`, colonne impilate <640px |
| `ae8edcb` | **Mese fisso 26 GG pieni** (sezione giorni rimossa), KPI = netto annuale ÷ 14, override a zero |
| corrente | SUP.CONTR = superminimo da contratto; HTML: unica variabile "Superminimo ad personam", importi variabili a ZERO, etichette "Fringe benefit" e "Indennità denaro" |
| `fondo-verifica` | **IRPEF netta in busta** (lorda − detrazione art. 13, tutte e tre le implementazioni), riga-ponte mese tipo ↔ media annuale, credito welfare 2026–2028 documentato (1.500 €/anno); suite: 15 casi, 480 confronti, 38 verifiche cedolini reali + 77 controlli HTML |
| `sim27` | Terzo benchmark: elaborazione di controllo gennaio 2027 (netto target 44.000) replicata — regole condivise confermate al centesimo; differenze spiegate (FASDAC/Pastore/detrazione omesse dall'allegato); suite: 59 verifiche benchmark reali |
