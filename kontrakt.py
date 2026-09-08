# -*- coding: utf-8 -*-
"""
Kontrakt for supportsystem-generatorene.

Én kilde til sannhet for celle-referanser, konstanter, scoring-orakelet
og mål-profiler. Ingen verdier endres her – kun sentralisert.
"""

# ---------------- Kapasiteter ----------------
MAX_S = 501   # maksimalt antall rader i Saker (inkl. overskrift)
MAX_K = 301   # maksimalt antall rader i Kunnskapsbase (inkl. overskrift)
N_KB = 25     # maks antall KB-treff som vises
N_SAK = 30    # maks antall sakstreff som vises

# ---------------- Arkfaner og codeName-mapping ----------------
SHEETS = [
    "Dashboard",
    "Saker",
    "Søk KB",
    "Søk saker",
    "Kunnskapsbase",
    "Oppsett",
    "Instruksjoner",
]

CODE_NAMES = {
    "Dashboard": "Sheet1",
    "Saker": "Sheet2",
    "Søk KB": "Sheet3",
    "Søk saker": "Sheet4",
    "Kunnskapsbase": "Sheet5",
    "Oppsett": "Sheet6",
    "Instruksjoner": "Sheet7",
}

# ---------------- Overskrifter ----------------
SAKER_HEADERS = [
    "Saksnr", "Opprettet", "Innmelder", "E-post", "Kanal", "Kategori",
    "Prioritet", "Tittel", "Beskrivelse", "Auto-ansvarlig", "Tildelt til",
    "Ansvarlig", "Status", "SLA start frist", "SLA løsning frist",
    "Påbegynt", "Løst", "SLA start", "SLA løsning", "Løsningstid timer",
    "Løsning", "KB-ref", "Til KB",
]

KB_HEADERS = [
    "KB-ID", "Tittel", "Kategori", "Problem", "Løsning", "Nøkkelord",
    "Opprettet av", "Dato", "Antall bruk", "Vurdering", "Basert på sak",
]

# ---------------- Søke-celler og speil-tabeller ----------------
SOK_KB_INPUT = "'Søk KB'!$B$4"
SOK_KB_KAT = "'Søk KB'!$B$5"
SOK_SAK_INPUT = "'Søk saker'!$B$4"
# Skjult speil-tabell (statusnavn | SANN/USANN) på Søk saker – avkryssingsboksene
# er koblet til kolonne AB, og Saker!Y slår opp statusen her.
SOK_SAK_STATUS = "'Søk saker'!$AA$5:$AB$14"

# ---------------- Hjelpekolonner ----------------
SAKER_HELPER_COL = "Y"    # Saker!Y – skjult hjelpekolonne for sakssøk
KB_HELPER_COL = "L"       # Kunnskapsbase!L – skjult hjelpekolonne for KB-søk
SOK_SAK_MIRROR_STATUS = "AA"
SOK_SAK_MIRROR_VALUE = "AB"

# ---------------- Dashboard KPI-er ----------------
# Hver tuple er (label_celle, verdi_celle, label_tekst, formel)
KPIS = [
    ("A4", "B4", "Åpne saker",
     '=COUNTIFS(Saker!$C$2:$C$501,"<>",Saker!$M$2:$M$501,"<>Løst",Saker!$M$2:$M$501,"<>Lukket")'),
    ("C4", "D4", "Forfalte løsningsfrister", '=COUNTIFS(Saker!$S$2:$S$501,"FORFALT")'),
    ("E4", "F4", "Forfalte startfrister", '=COUNTIFS(Saker!$R$2:$R$501,"FORFALT")'),
    ("G4", "H4", "Løst denne måneden",
     '=COUNTIFS(Saker!$C$2:$C$501,"<>",Saker!$Q$2:$Q$501,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))'),
    ("I4", "J4", "Snitt løsningstid (timer)", '=IFERROR(ROUND(AVERAGE(Saker!$T$2:$T$501),1),"-")'),
    ("K4", "L4", "KB-artikler", '=COUNTA(Kunnskapsbase!$B$2:$B$301)'),
]

# Hjelpetabeller under diagrammene
DASH_STATUS_RANGE = "A27:B32"      # status | antall (fast, 6 statuser)
DASH_KATEGORI_RANGE = "A35:B44"    # kategori | antall (topp-10, sortert synkende)
DASH_ANSVARLIG_RANGE = "C27:D36"   # ansvarlig | antall (topp-10, sortert synkende)

DASH_STATUS_START = 27
DASH_STATUS_END = 32
DASH_KATEGORI_START = 35
DASH_KATEGORI_END = 44
DASH_ANSVARLIG_START = 27
DASH_ANSVARLIG_END = 36

# Topp-N-grense for kategori/ansvarlig-diagrammene.
# Hjelpetabellene viser inntil DASH_TOPP_N rader (sortert synkende etter antall);
# kandidater uten saker utelates, og radene brukes dynamisk.
DASH_TOPP_N = 10

# ---------------- Diagrammer ----------------
CHART_H = 9.5   # cm
CHART_W = 16    # cm
CHART_ANCHORS = ("A8", "F8", "F26")

# ---------------- Søkeresultat-rader ----------------
SOK_KB_HEADER_ROW = 8
SOK_KB_RESULT_START = SOK_KB_HEADER_ROW + 1          # 9
SOK_KB_RESULT_END = SOK_KB_HEADER_ROW + N_KB         # 33

SOK_SAK_HEADER_ROW = 7
SOK_SAK_RESULT_START = SOK_SAK_HEADER_ROW + 1        # 8
SOK_SAK_RESULT_END = SOK_SAK_HEADER_ROW + N_SAK      # 37

SOK_KB_COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]
SOK_SAK_COLS = ["A", "B", "F", "H", "U", "V"]

# ---------------- Søkekolonner i kildedata ----------------
# For scoring-orakelet (nøkler matcher FELT_VEKTER)
KB_SEARCH_FIELDS = {
    "title": "B",      # Tittel
    "keywords": "F",   # Nøkkelord
    "problem": "D",    # Problem
    "solution": "E",   # Løsning
}

SAKER_SEARCH_FIELDS = {
    "title": "H",      # Tittel
    "keywords": None,  # ikke egen kolonne for saker
    "problem": "I",    # Beskrivelse
    "solution": "U",   # Løsning
}

# ---------------- Scoring-orakel ----------------
FELT_VEKTER = {
    "title": 4,
    "keywords": 3,
    "problem": 2,
    "solution": 1,
}


def normalize(s):
    """Normaliser tekst for søk: små bokstaver, bindestrek FJERNES (som SUBSTITUTE(x,"-","")), kollaps whitespace."""
    if s is None:
        return ""
    t = str(s).lower().replace("-", "")
    import re
    return re.sub(r"\s+", " ", t).strip()


def levenshtein(a, b):
    """Standard dynamisk-programmering redigeringsavstand."""
    if a == b:
        return 0
    if len(a) < len(b):
        a, b = b, a
    if not b:
        return len(a)

    prev = list(range(len(b) + 1))
    for i, ca in enumerate(a, start=1):
        curr = [i]
        for j, cb in enumerate(b, start=1):
            cost = 0 if ca == cb else 1
            curr.append(min(curr[-1] + 1, prev[j] + 1, prev[j - 1] + cost))
        prev = curr
    return prev[-1]


def score(query, doc):
    """
    Score et dokument mot et søk.

    doc er en dict med valgfrie nøkler blant title/keywords/problem/solution.
    For hvert søkeord: gi feltets vekt hvis ordet er delstreng i feltet.
    Hvis ingen delstreng-treff: fuzzy fallback der hvert ord i feltet må ha
    Levenshtein-avstand <= 2 fra søkeordet.
    """
    terms = normalize(query).split()
    if not terms:
        return 0

    total = 0
    for term in terms:
        for felt, vekt in FELT_VEKTER.items():
            tekst = normalize(doc.get(felt, ""))
            if term in tekst:
                total += vekt

    if total == 0:
        for term in terms:
            for felt, vekt in FELT_VEKTER.items():
                tekst = normalize(doc.get(felt, ""))
                ord_i_tekst = tekst.split()
                for w in ord_i_tekst:
                    if levenshtein(term, w) <= 2:
                        total += vekt
                        break

    return total


# ---------------- Profiler ----------------
FALLBACK_PROFILE = {
    "timestamp_mechanism": "formula",
    "search_mechanism": "formula",
    "set_codename": False,
}

EXCEL_PROFILE = {
    "timestamp_mechanism": "macro",
    "search_mechanism": "macro",
    "set_codename": True,
}

CALC_PROFILE = {
    "timestamp_mechanism": "macro",
    "search_mechanism": "macro",
    "set_codename": True,
}
