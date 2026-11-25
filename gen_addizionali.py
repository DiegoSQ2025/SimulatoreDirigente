#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Trasforma gli elenchi MEF 2026 in CSV normalizzati per addizionale_lib,
build_payslip.py e simulatore_cedolino.html.

Nota: questo script consuma dati intermedi non versionati nel repo
(/tmp/comuni.json = elenco compatto comuni MEF 2026; /tmp/reg26.json =
aliquote enti regionali 2026). Per un aggiornamento dai sorgenti ufficiali:
1) scaricare https://www1.finanze.gov.it/finanze2/dipartimentopolitichefiscali/
   fiscalitalocale/nuova_addcomirpef/download/download.php?anno=2026 (comuni)
   e addregirpef/download/download.php?anno=2026 (regioni);
2) adattare il parsing al formato corrente (colonne: CODICE, COMUNE, PR,
   MULTIALIQ, ALIQUOTA, FASCIA..., IMPORTO_ESENTE).
Le aliquote sono frazioni decimali (0.009 = 0,90%).
"""
import json

OUT_COM = "/workspace/data/addizionali_comuni_2026.csv"
OUT_REG = "/workspace/data/addizionali_regionali_2026.csv"
BIG = 1e15


def parse_spec(spec):
    """Ritorna (regola, brackets[(soglia, aliquota)], esenzione, note)."""
    parts = spec.split("+")
    flags = list(parts[1:])
    main = parts[0]
    esenz = 0.0
    for p in list(flags):
        if p.startswith("e"):
            esenz = float(p[1:].replace(",", "."))
            flags.remove(p)
    note = []
    if "y25" in flags:
        note.append("aliquote 2025 (delibera 2026 non ancora pubblicata)")
    if "!" in flags:
        note.append("casi particolari/esenzioni categoriali: verificare delibera")
    if main == "n":
        return "n", [(BIG, 0.0)] * 4, 0.0, ["comune senza addizionale o delibera non presente"]
    if main.startswith("u"):
        a = float(main[1:].replace(",", ".")) / 100.0
        return "u", [(BIG, a), (BIG, 0.0), (BIG, 0.0), (BIG, 0.0)], esenz, note
    if main.startswith("b"):
        brackets = []
        for t in main[1:].split(";"):
            soglia_s, alq_s = t.split(":", 1)
            soglia = float(soglia_s.replace(",", ".")) if soglia_s else BIG
            aliquota = float(alq_s.replace(",", ".")) / 100.0
            if brackets and soglia <= brackets[-1][0] and brackets[-1][0] < BIG:
                return "n", [(BIG, 0.0)] * 4, 0.0, ["regola a scaglioni incoerente: impostare manualmente (vedi delibera MEF)"]
            brackets.append((soglia, aliquota))
        while len(brackets) < 4:
            brackets.append((BIG, 0.0))
        return "b", brackets, esenz, note
    return "n", [(BIG, 0.0)] * 4, 0.0, ["regola non riconosciuta: impostare manualmente"]


def main():
    prov = json.load(open("/tmp/prov.json", encoding="utf-8"))
    comuni_raw = json.load(open("/tmp/comuni.json", encoding="utf-8"))

    n_anom = 0
    rows = []
    regioni_seen = set()
    for item in comuni_raw:
        codice, nome, pr, spec, _pop = item.split("|")
        regione = prov.get(pr)
        if not regione:
            raise SystemExit("Provincia sconosciuta: " + pr)
        regioni_seen.add(regione)
        regola, brackets, esenz, note = parse_spec(spec)
        if regola == "n":
            n_anom += 1
        rows.append((regione, nome, pr, regola, brackets, esenz, "; ".join(note)))
    rows.sort(key=lambda r: (r[0], r[3], r[1]))

    with open(OUT_COM, "w", encoding="utf-8", newline="") as f:
        f.write("# Sorgente: MEF Dipartimento delle Finanze, elenchi generali addizionale comunale IRPEF 2026 (7.897 comuni, nov. 2026)\n")
        f.write("# Aliquote come frazione decimale (0.009 = 0,90%). 'y25' = delibera 2026 non pubblicata, si usa aliquota 2025.\n")
        f.write("Regione;Provincia;Comune;Regola;A1;S1;A2;S2;A3;S3;A4;S4;Esenzione;Nota\n")
        for regione, nome, pr, regola, brackets, esenz, nota in rows:
            if regola == "n":
                f.write("{};{};{};n;0;{};0;{};0;{};0;{};{:g};{}\n".format(
                    regione, pr, nome, BIG, BIG, BIG, BIG, esenz, nota))
            else:
                (s1, a1), (s2, a2), (s3, a3), (s4, a4) = brackets
                f.write("{};{};{};{};{:.6f};{:.0f};{:.6f};{:.0f};{:.6f};{:.0f};{:.6f};{:.0f};{:g};{}\n".format(
                    regione, pr, nome, regola, a1, s1, a2, s2, a3, s3, a4, s4, esenz, nota))

    conf = {
        "Abruzzo": ("std", [(28000, 1.67), (50000, 2.87), (BIG, 3.33)], (0, 0, 0), ""),
        "Basilicata": ("std", [(BIG, 1.23)], (0, 0, 0), ""),
        "Calabria": ("std", [(BIG, 1.73)], (0, 0, 0), ""),
        "Campania": ("std", [(15000, 1.73), (28000, 2.96), (50000, 3.20), (BIG, 3.33)], (0, 0, 0), "detrazioni regionali per figli (da due) e disabilita non modellate"),
        "Emilia-Romagna": ("std", [(15000, 1.33), (28000, 1.93), (50000, 2.78), (BIG, 3.33)], (0, 0, 0), ""),
        "Friuli-Venezia Giulia": ("fvg", [(BIG, 1.23)], (0, 0, 0), "0,70% fino a 15.000, oltre 1,23% sull'intero imponibile"),
        "Lazio": ("lazio", [(15000, 1.73), (BIG, 3.33)], (60, 28000, 30000), "L.R. 20/2025: 1,73% fino a 28.000; oltre scaglioni 1,73/3,33; detrazione 60 euro tra 28.001 e 30.000"),
        "Liguria": ("std", [(28000, 1.23), (50000, 3.18), (BIG, 3.23)], (0, 0, 0), ""),
        "Lombardia": ("std", [(15000, 1.23), (28000, 1.58), (50000, 1.72), (BIG, 1.73)], (0, 0, 0), ""),
        "Marche": ("std", [(15000, 1.23), (28000, 1.53), (50000, 1.70), (BIG, 1.73)], (0, 0, 0), "aliquota ridotta per figli disabili non modellata"),
        "Molise": ("std", [(15000, 2.03), (28000, 2.23), (BIG, 3.63)], (0, 0, 0), "aliquote rideterminate per disavanzo sanitario"),
        "Piemonte": ("std", [(15000, 1.62), (28000, 2.68), (50000, 3.31), (BIG, 3.33)], (0, 0, 0), "detrazioni figli (da tre) e disabilita non modellate"),
        "Puglia": ("std", [(15000, 1.33), (28000, 2.13), (50000, 3.23), (BIG, 3.33)], (0, 0, 0), "aliquote rideterminate per disavanzo sanitario"),
        "Sardegna": ("std", [(BIG, 1.23)], (0, 0, 0), "detrazione figli minorenni non modellate"),
        "Sicilia": ("std", [(BIG, 1.23)], (0, 0, 0), ""),
        "Toscana": ("std", [(15000, 1.42), (28000, 1.43), (50000, 3.32), (BIG, 3.33)], (0, 0, 0), ""),
        "Trentino-Alto Adige (Bolzano)": ("bolzano", [(50000, 1.23), (BIG, 1.73)], (430.50, 0, 90000), "detrazione 430,50 euro fino a 90.000; ulteriore detrazione oltre 50.000 non modellata"),
        "Trentino-Alto Adige (Trento)": ("trento", [(50000, 1.23), (BIG, 1.73)], (0, 0, 0), "deduzione 30.000 euro; detrazione 246 euro per figlio non modellata"),
        "Umbria": ("umbria", [(15000, 1.73), (28000, 3.02), (50000, 3.12), (BIG, 3.33)], (150, 28000, 50000), "1,23% fino a 28.000; oltre scaglioni con detrazione 150 euro fino a 50.000"),
        "Valle d'Aosta": ("vda", [(BIG, 1.23)], (0, 0, 0), "nessuna addizionale fino a 15.000; oltre 1,23% sull'intero imponibile"),
        "Veneto": ("std", [(BIG, 1.23)], (0, 0, 0), ""),
    }
    enti = sorted(conf.keys(), key=str.lower)
    with open(OUT_REG, "w", encoding="utf-8", newline="") as f:
        f.write("# Sorgente: MEF Dipartimento delle Finanze, addizionale regionale IRPEF 2026 (21 enti, pubblico gen-giu 2026)\n")
        f.write("Ente;Regola;S1;A1;S2;A2;S3;A3;S4;A4;Detrazione;Detr_da;Detr_a;Nota\n")
        for ente in enti:
            regola, br, (det, dda, daa), nota = conf[ente]
            ph = [(s, a / 100.0) for s, a in br]
            while len(ph) < 4:
                ph.append((BIG, 0.0))
            (s1, a1), (s2, a2), (s3, a3), (s4, a4) = ph
            f.write("{};{};{:.0f};{:.6f};{:.0f};{:.6f};{:.0f};{:.6f};{:.0f};{:.6f};{:g};{:g};{:g};{}\n".format(
                ente, regola, s1, a1, s2, a2, s3, a3, s4, a4, det, dda, daa, nota))
    print("comuni:", len(rows), "| senza addizionale/anomali:", n_anom,
          "| enti:", len(enti), "| regioni viste:", len(regioni_seen))
    assert len(regioni_seen) == 21, regioni_seen


if __name__ == "__main__":
    main()
