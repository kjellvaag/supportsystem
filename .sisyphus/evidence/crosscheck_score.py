# -*- coding: utf-8 -*-
"""Kryssjekk: Python-motpart av Basic ScoreDok vs kontrakt.score.

Speiler EXAKT løkkestrukturen i basic/ModFelles.bas sin ScoreDok (inkludert
Normaliser og Levenshtein), og sammenligner mot kontrakt.score (orakelet).
"""
import os
import sys

# Sørg for at prosjektroten er på sys.path (kontrakt.py ligger der)
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__)))))

import kontrakt

# --- Python-motpart av Basic ModFelles -------------------------------------

def normaliser_basic(t):
    """Speiler Basic Normaliser: LCase, fjern bindestrek, kollaps whitespace, Trim."""
    if t is None:
        return ""
    s = str(t).lower().replace("-", "")
    resultat = ""
    siste = True
    for tegn in s:
        if tegn in (" ", "\t", "\n", "\r"):
            if not siste:
                resultat += " "
                siste = True
        else:
            resultat += tegn
            siste = False
    return resultat.strip()


def levenshtein_basic(a, b):
    """Speiler Basic Levenshtein (to-rader DP)."""
    if a == b:
        return 0
    if len(a) < len(b):
        a, b = b, a
    if not b:
        return len(a)
    la, lb = len(a), len(b)
    forrige = list(range(lb + 1))
    for i in range(1, la + 1):
        gjeldende = [i]
        for j in range(1, lb + 1):
            kostnad = 0 if a[i - 1] == b[j - 1] else 1
            g = gjeldende[-1] + 1
            if forrige[j] + 1 < g:
                g = forrige[j] + 1
            if forrige[j - 1] + kostnad < g:
                g = forrige[j - 1] + kostnad
            gjeldende.append(g)
        forrige = gjeldende
    return forrige[lb]


V_TITTEL, V_NOKKEL, V_PROBLEM, V_LOSNING = 4, 3, 2, 1


def scoredok_basic(query, tittel, nokkelord, problem, losning):
    """Speiler Basic ScoreDok eksakt (samme løkkestruktur som kontrakt.score)."""
    nq = normaliser_basic(query)
    if nq == "":
        return 0
    felt = [normaliser_basic(tittel), normaliser_basic(nokkelord),
            normaliser_basic(problem), normaliser_basic(losning)]
    vekt = [V_TITTEL, V_NOKKEL, V_PROBLEM, V_LOSNING]

    term = nq.split(" ")

    total = 0
    for t in term:
        if t != "":
            for k in range(4):
                if felt[k] != "":
                    if t in felt[k]:
                        total += vekt[k]

    if total == 0:
        for t in term:
            if t != "":
                for k in range(4):
                    if felt[k] != "":
                        ord_liste = felt[k].split(" ")
                        for w in ord_liste:
                            if w != "":
                                if levenshtein_basic(t, w) <= 2:
                                    total += vekt[k]
                                    break
    return total


# --- Testsett --------------------------------------------------------------

def doc(title="", keywords="", problem="", solution=""):
    return {"title": title, "keywords": keywords, "problem": problem, "solution": solution}


TESTS = [
    # (query, doc, forventet, forklaring)
    ("wifi", doc(title="Wi-Fi nede"), 4, "eksakt delstreng i tittel (bindestrek fjernet)"),
    ("wify", doc(title="Wi-Fi nede"), 4, "fuzzy (Levenshtein 1) i tittel"),
    ("wifi", doc(keywords="wifi trådløs nettverk"), 3, "eksakt delstreng i nøkkelord"),
    ("wifi", doc(problem="wifi sluttet å virke"), 2, "eksakt delstreng i problem"),
    ("wifi", doc(solution="start wifi-ruteren på nytt"), 1, "eksakt delstreng i løsning"),
    ("wifi trådløs", doc(title="Wi-Fi nede", keywords="wifi trådløs nettverk"), 10,
     "to søkeord: wifi->tittel(4)+nøkkelord(3), trådløs->nøkkelord(3) = 10"),
    ("wify", doc(title="Wi-Fi nede", keywords="wifi trådløs", problem="nettverk",
                 solution="restart"), 7, "fuzzy i tittel(4)+nøkkelord(3); ingen i problem/løsning"),
    ("", doc(title="Wi-Fi nede"), 0, "tomt søk gir 0"),
    ("xyzabc", doc(title="Wi-Fi nede"), 0, "ingen treff, for lang avstand til fuzzy"),
]

ok = True
linjer = []
linjer.append("Kryssjekk av Basic ScoreDok (Python-motpart) mot kontrakt.score")
linjer.append("=" * 70)
linjer.append("")

for query, d, forventet, forklaring in TESTS:
    basic = scoredok_basic(query, d.get("title", ""), d.get("keywords", ""),
                           d.get("problem", ""), d.get("solution", ""))
    orakel = kontrakt.score(query, d)
    status = "OK" if (basic == orakel == forventet) else "AVVIK"
    if status != "OK":
        ok = False
    linjer.append(f"[{status}] query={query!r}  forventet={forventet}  "
                  f"basic={basic}  orakel={orakel}")
    linjer.append(f"        doc={d!r}  ({forklaring})")

linjer.append("")
linjer.append("=" * 70)

# De to påkrevde sjekkene eksplisitt:
wifi_doc = {"title": "Wi-Fi nede"}
wifi_basic = scoredok_basic("wifi", "Wi-Fi nede", "", "", "")
wifi_orakel = kontrakt.score("wifi", wifi_doc)
wify_basic = scoredok_basic("wify", "Wi-Fi nede", "", "", "")
wify_orakel = kontrakt.score("wify", wifi_doc)

linjer.append(f"score('wifi', {{'title':'Wi-Fi nede'}}) -> basic={wifi_basic}, orakel={wifi_orakel} (forventet 4)")
linjer.append(f"score('wify', {{'title':'Wi-Fi nede'}}) -> basic={wify_basic}, orakel={wify_orakel} (forventet 4)")
linjer.append("")
linjer.append(f"Normaliser('Wi-Fi nede') -> basic={normaliser_basic('Wi-Fi nede')!r}, "
              f"orakel={kontrakt.normalize('Wi-Fi nede')!r}")
linjer.append("")
linjer.append("RESULTAT: " + ("ALLE MATCHER" if ok else "AVVIK FUNNET"))

print("\n".join(linjer))
sys.exit(0 if ok else 1)
