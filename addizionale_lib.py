#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Libreria addizionali regionali e comunali IRPEF 2026.

Fonte: Dipartimento delle Finanze (MEF), elenchi generali 2026 delle
addizionali regionali e comunali all'IRPEF (fiscalitalocale).
- Comuni: 7.897 enti, dati MEF aggiornati (novembre 2026); per i comuni
  privi di delibera 2026 in pubblicazione si usa l'aliquota 2025 (flag "y25"
  nella colonna Nota). Nel dubbio: www1.finanze.gov.it -> fiscalita locale.
- Regioni: 21 enti (regioni + province autonome di Bolzano e Trento).
Peresso le regole speciali modellate: Lazio (1,73% fino a 28.000; oltre
scaglioni 1,73/3,33 e detrazione 60 euro tra 28.001 e 30.000; L.R. 20/2025),
Umbria, Friuli-Venezia Giulia, Valle d'Aosta, Trento, Bolzano.
Non modellate: detrazioni regionali per figli/disabilita (Campania, Piemonte,
Sardegna, Marche, Bolzano, Trento) e casi particolari comunali (nota "!").
Le aliquote sono memorizzate come frazione decimale (0.009 = 0,90%).
"""
import csv
import os

DATA_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "data")
BIG = 1e15


def _skip_comments(lines):
    return (ln for ln in lines if not ln.lstrip().startswith("#"))


def _norm_regione(regione):
    return regione.replace(" · ", " (") + ")" if " · " in regione else regione


def load_comuni():
    path = os.path.join(DATA_DIR, "addizionali_comuni_2026.csv")
    comuni = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(_skip_comments(f), delimiter=";"):
            regione = _norm_regione(row["Regione"])
            comuni[(regione, row["Comune"])] = {
                "prov": row["Provincia"],
                "regola": row["Regola"],
                "brackets": [
                    (float(row["S1"]), float(row["A1"])),
                    (float(row["S2"]), float(row["A2"])),
                    (float(row["S3"]), float(row["A3"])),
                    (float(row["S4"]), float(row["A4"])),
                ],
                "esenzione": float(row["Esenzione"] or 0),
                "nota": row["Nota"],
            }
    return comuni


def load_regionali():
    path = os.path.join(DATA_DIR, "addizionali_regionali_2026.csv")
    regionali = {}
    with open(path, encoding="utf-8") as f:
        for row in csv.DictReader(_skip_comments(f), delimiter=";"):
            regionali[row["Ente"]] = {
                "regola": row["Regola"],
                "brackets": [
                    (float(row["S1"]), float(row["A1"])),
                    (float(row["S2"]), float(row["A2"])),
                    (float(row["S3"]), float(row["A3"])),
                    (float(row["S4"]), float(row["A4"])),
                ],
                "detr": float(row["Detrazione"] or 0),
                "detr_da": float(row["Detr_da"] or 0),
                "detr_a": float(row["Detr_a"] or 0),
                "nota": row["Nota"],
            }
    return regionali


def marg(imp, brackets):
    tot, lower = 0.0, 0.0
    for soglia, aliquota in brackets:
        tot += max(0.0, min(imp, soglia) - lower) * aliquota
        lower = soglia
    return tot


def addizionale_regionale(imp, ente, regionali=None):
    regionali = regionali if regionali is not None else _REGIONALI
    reg = regionali.get(ente)
    if reg is None:
        return 0.0
    rule = reg["regola"]
    b = reg["brackets"]
    if rule == "lazio":
        if imp <= 28000:
            t = 0.0173 * imp
        else:
            t = marg(imp, b)
        d = reg["detr"] if (reg["detr"] and reg["detr_da"] < imp <= reg["detr_a"]) else 0.0
        return max(0.0, t - d)
    if rule == "umbria":
        if imp <= 28000:
            t = 0.0123 * imp
        else:
            t = marg(imp, b)
        d = reg["detr"] if (reg["detr"] and reg["detr_da"] < imp <= reg["detr_a"]) else 0.0
        return max(0.0, t - d)
    if rule == "fvg":
        return 0.0070 * imp if imp <= 15000 else 0.0123 * imp
    if rule == "vda":
        return 0.0 if imp <= 15000 else 0.0123 * imp
    if rule == "trento":
        base = imp - 30000
        if base <= 0:
            return 0.0
        return 0.0123 * base if base <= 50000 else 0.0123 * 50000 + 0.0173 * (base - 50000)
    if rule == "bolzano":
        d = reg["detr"] if (reg["detr"] and imp <= reg["detr_a"]) else 0.0
        return max(0.0, marg(imp, b) - d)
    t = marg(imp, b)
    if reg["detr"] and reg["detr_da"] < imp <= reg["detr_a"]:
        t -= reg["detr"]
    return max(0.0, t)


def addizionale_comunale(imp, regione, comune, comuni=None):
    comuni = comuni if comuni is not None else _COMUNI
    c = comuni.get((regione, comune))
    if c is None or c["regola"] == "n":
        return 0.0
    if c["esenzione"] > 0 and imp <= c["esenzione"]:
        return 0.0
    return marg(imp, c["brackets"])


def trattamento_integrativo(rdl, irpl, ded_art13, add_reg, add_com,
                            spese_extra=0.0, mesi=12):
    """Trattamento integrativo (ex bonus Renzi) 2026, prorato sui mesi.

    - reddito complessivo <= 15.000: 1.200 euro se imposta lorda > detrazione
      art. 13 (capienza); sotto soglia non vale se non si copre un'imposta.
    - 15.001-28.000: spetta solo con detrazioni/oneri aggiuntivi (spese
      sanitarie, mutuo, ristrutturazioni, carichi famiglia...) che azzerano
      (o superano) l'imposta: importo = differenza, fino a 1.200 euro.
    - oltre 28.000: nessun trattamento integrativo.
    Erogato esente, fuori imponibile.
    """
    if mesi >= 12:
        fatt = 1.0
    else:
        fatt = mesi / 12.0
    if rdl <= 15000:
        ti = 1200.0 if irpl > ded_art13 else 0.0
    elif rdl <= 28000 and spese_extra > 0:
        residuo = max(0.0, irpl - min(ded_art13, irpl) - add_reg - add_com)
        ti = min(1200.0, max(0.0, spese_extra - residuo))
    else:
        ti = 0.0
    return ti * fatt


_REGIONALI = load_regionali()
_COMUNI = load_comuni()

ENTI = list(_REGIONALI.keys())
DEFAULT_REGIONE = "Lazio"
DEFAULT_COMUNE = "Roma"
