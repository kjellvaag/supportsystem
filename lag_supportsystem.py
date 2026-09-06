# -*- coding: utf-8 -*-
"""
Genererer et komplett supportsystem i Excel for en driftsavdeling.

Innhold i arbeidsboken:
  - Dashboard      : KPI-er og diagrammer (oppdateres automatisk)
  - Saker          : Sakslogg med unikt saksnr, auto-tildeling, SLA-frister
  - Søk KB         : Søk i kunnskapsbasen (FAQ)
  - Søk saker      : Søk i tidligere løste saker
  - Kunnskapsbase  : FAQ/artikler med løsninger og nøkkelord
  - Oppsett        : Team, kategorier, SLA-tider, statuser, kanaler
  - Instruksjoner  : Brukerveiledning

Kjør:  python lag_supportsystem.py          (med eksempelsaker og KB-artikler)
       python lag_supportsystem.py ren      (tomt system, klart til bruk)
Gir:   Supportsystem.xlsx i samme mappe
"""
import sys
import datetime
import os
import re
import zipfile

# Med eksempeldata som standard; "ren" som argument genererer et tomt system.
MED_EKSEMPLER = not any(a.lower() in ("ren", "clean", "--ren", "--clean") for a in sys.argv[1:])
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import FormulaRule
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import SeriesLabel
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.utils import get_column_letter

OUT = "Supportsystem.xlsx"
DT = "DD.MM.YYYY HH:MM"
DONLY = "DD.MM.YYYY"

# ---------------- Farger og stiler ----------------
DARK = "1F4E79"
TITLE_FONT = Font(size=16, bold=True, color=DARK)
SECTION_FONT = Font(size=12, bold=True, color=DARK)
NOTE_FONT = Font(italic=True, color="808080")
HEADER_FILL = PatternFill("solid", fgColor=DARK)
HEADER_FONT = Font(color="FFFFFF", bold=True)
GREEN_FILL = PatternFill("solid", fgColor="C6EFCE")
GREEN_FONT = Font(color="006100")
RED_FILL = PatternFill("solid", fgColor="FFC7CE")
RED_FONT = Font(color="9C0006", bold=True)
ORANGE_FILL = PatternFill("solid", fgColor="FCE4D6")
ORANGE_FONT = Font(color="974706")
YELLOW_FILL = PatternFill("solid", fgColor="FFEB9C")
BLUE_FILL = PatternFill("solid", fgColor="DDEBF7")
GRAY_FILL = PatternFill("solid", fgColor="D9D9D9")
INPUT_FILL = PatternFill("solid", fgColor="FFF2CC")
KPI_LABEL_FONT = Font(color="FFFFFF", bold=True)
KPI_VALUE_FONT = Font(size=20, bold=True, color=DARK)
THIN = Side(style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")

wb = Workbook()
ws_dash = wb.active
ws_dash.title = "Dashboard"
ws_saker = wb.create_sheet("Saker")
ws_sok_kb = wb.create_sheet("Søk KB")
ws_sok_sak = wb.create_sheet("Søk saker")
ws_kb = wb.create_sheet("Kunnskapsbase")
ws_oppsett = wb.create_sheet("Oppsett")
ws_instr = wb.create_sheet("Instruksjoner")

ws_dash.sheet_properties.tabColor = "70AD47"
ws_saker.sheet_properties.tabColor = "4472C4"
ws_sok_kb.sheet_properties.tabColor = "7030A0"
ws_sok_sak.sheet_properties.tabColor = "7030A0"
ws_kb.sheet_properties.tabColor = "ED7D31"
ws_oppsett.sheet_properties.tabColor = "808080"
ws_instr.sheet_properties.tabColor = "2E9E9B"

for ws in (ws_dash, ws_sok_kb, ws_sok_sak, ws_oppsett, ws_instr):
    ws.sheet_view.showGridLines = False

# Søke-layout (deles mellom Saker/KB-hjelpekolonner og søkearkene)
N_KB = 25    # maks antall KB-treff som vises
N_SAK = 30   # maks antall sakstreff som vises
SOK_KB_INPUT = "'Søk KB'!$B$4"
SOK_KB_KAT = "'Søk KB'!$B$5"
SOK_SAK_INPUT = "'Søk saker'!$B$4"
# Skjult speil-tabell (statusnavn | SANN/USANN) på Søk saker – avkryssingsboksene
# er koblet til kolonne AB, og Saker!Y slår opp statusen her.
SOK_SAK_STATUS = "'Søk saker'!$AA$5:$AB$14"

# Tving full omberegning ved åpning (SLA/NOW-formler)
# + iterativ beregning for automatisk tidsstempel i Saker!B (sirkulær referanse)
try:
    wb.calculation.fullCalcOnLoad = True
    wb.calculation.iterate = True
    wb.calculation.iterateCount = 100
    wb.calculation.iterateDelta = 0.0001
except Exception:
    pass

# ============================================================
# OPPSSETT – grunndata (frie områder med dynamiske navn: legg til/fjern rader direkte)
# ============================================================
o = ws_oppsett
o["A1"] = "OPPSETT – grunndata for supportsystemet"
o["A1"].font = TITLE_FONT
o["A2"] = ("Tilpass listene under direkte i Excel. Skriv i første ledige rad for å legge til, "
           "eller slett en rad for å fjerne. Formler og rullegardiner oppdateres automatisk.")
o["A2"].font = NOTE_FONT

team = [
    ("Ola Nordmann", "Nettverksansvarlig", "ola.nordmann@firma.no"),
    ("Kari Hansen", "Brukerstøtte", "kari.hansen@firma.no"),
    ("Per Olsen", "Maskinvareansvarlig", "per.olsen@firma.no"),
    ("Anne Berg", "Systemansvarlig", "anne.berg@firma.no"),
    ("Erik Dahl", "Driftstekniker", "erik.dahl@firma.no"),
    ("Mona Lie", "Driftsleder", "mona.lie@firma.no"),
]
kategorier = [
    ("Nettverk", "Ola Nordmann"),
    ("Passord og tilgang", "Kari Hansen"),
    ("E-post og kalender", "Anne Berg"),
    ("Maskinvare", "Per Olsen"),
    ("Programvare", "Anne Berg"),
    ("Skriver og utskrift", "Erik Dahl"),
    ("Mobil og telefoni", "Erik Dahl"),
    ("Annet", "Mona Lie"),
]
sla = [("Kritisk", 1, 4), ("Høy", 4, 8), ("Medium", 8, 24), ("Lav", 24, 72)]
statuser = ["Ny", "Under behandling", "Venter på bruker", "Venter på leverandør", "Løst", "Lukket"]
kanaler = ["E-post", "Telefon", "Personlig oppmøte", "Chat", "Selvbetjening"]


def skriv_blokk(ws, start_col, header_row, headers, data):
    """Skriver header + data som et fritt område (ingen tabell – LO/OO-kompatibelt)."""
    for j, h in enumerate(headers):
        c = ws.cell(row=header_row, column=start_col + j, value=h)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.border = BORDER
    for i, row in enumerate(data):
        for j, v in enumerate(row):
            c = ws.cell(row=header_row + 1 + i, column=start_col + j, value=v)
            c.border = BORDER


# Frie områder (kan utvides direkte i Excel – dynamiske navn følger med)
skriv_blokk(o, 1, 3, ("Ansatt", "Rolle", "E-post"), team)
skriv_blokk(o, 5, 3, ("Kategori", "Standard ansvarlig (auto)"), kategorier)
skriv_blokk(o, 8, 3, ("Prioritet", "Påbegynn innen (timer)", "Løs innen (timer)"), sla)
skriv_blokk(o, 12, 3, ("Status",), [(v,) for v in statuser])
skriv_blokk(o, 14, 3, ("Kanal",), [(v,) for v in kanaler])

o["H9"] = "Timer regnes fra sak opprettes (24/7)."
o["H9"].font = NOTE_FONT

for col, w in {"A": 18, "B": 22, "C": 26, "D": 3, "E": 22, "F": 24, "G": 3,
               "H": 12, "I": 22, "J": 18, "K": 3, "L": 22, "M": 3, "N": 20}.items():
    o.column_dimensions[col].width = w

# Dynamiske navn (OFFSET + COUNTA): vokser/krymper med dataene – virker også i LibreOffice/OpenOffice
names = {
    "Ansatte": "OFFSET(Oppsett!$A$4,0,0,COUNTA(Oppsett!$A$4:$A$50),1)",
    "Kategorier": "OFFSET(Oppsett!$E$4,0,0,COUNTA(Oppsett!$E$4:$E$50),1)",
    "Prioriteter": "OFFSET(Oppsett!$H$4,0,0,COUNTA(Oppsett!$H$4:$H$50),1)",
    "Statuser": "OFFSET(Oppsett!$L$4,0,0,COUNTA(Oppsett!$L$4:$L$50),1)",
    "Kanaler": "OFFSET(Oppsett!$N$4,0,0,COUNTA(Oppsett!$N$4:$N$50),1)",
    # Kategori-matrise og SLA-matrise brukes av VLOOKUP i Saker
    "KategoriMatrise": "OFFSET(Oppsett!$E$4,0,0,COUNTA(Oppsett!$E$4:$E$50),2)",
    "SLAMatrise": "OFFSET(Oppsett!$H$4,0,0,COUNTA(Oppsett!$H$4:$H$50),3)",
}
for n, ref in names.items():
    dn = DefinedName(n, attr_text=ref)
    try:
        wb.defined_names[n] = dn
    except Exception:
        try:
            wb.defined_names.add(dn)
        except Exception:
            wb.defined_names.append(dn)

# ============================================================
# SAKER – sakslogg (fritt område med fargekoding – ingen Excel-tabell)
# ============================================================
s = ws_saker
headers = ["Saksnr", "Opprettet", "Innmelder", "E-post", "Kanal", "Kategori",
           "Prioritet", "Tittel", "Beskrivelse", "Auto-ansvarlig", "Tildelt til",
           "Ansvarlig", "Status", "SLA start frist", "SLA løsning frist",
           "Påbegynt", "Løst", "SLA start", "SLA løsning", "Løsningstid timer",
           "Løsning", "KB-ref", "Til KB"]
for j, h in enumerate(headers, start=1):
    c = s.cell(row=1, column=j, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.border = BORDER

MAX_S = 501  # datarader 2..501
for r in range(2, MAX_S + 1):
    # Auto-tidsstempel: fylles når Innmelder eller Tittel tastes, låses deretter.
    # Kan overskrives ved å skrive en dato manuelt. Krever iterativ beregning (slått på ovenfor).
    s.cell(row=r, column=2, value=f'=IF($C{r}&$H{r}="","",IF($B{r}="",NOW(),$B{r}))')
    s.cell(row=r, column=1, value=f'=IF($B{r}="","","SAK-"&TEXT(ROW()-1,"0000"))')
    s.cell(row=r, column=10, value=f'=IF($F{r}="","",IFERROR(VLOOKUP($F{r},KategoriMatrise,2,FALSE),"– ikke tildelt –"))')
    s.cell(row=r, column=12, value=f'=IF($B{r}="","",IF($K{r}<>"",$K{r},$J{r}))')
    s.cell(row=r, column=14, value=f'=IF(OR($B{r}="",$G{r}=""),"",$B{r}+VLOOKUP($G{r},SLAMatrise,2,FALSE)/24)')
    s.cell(row=r, column=15, value=f'=IF(OR($B{r}="",$G{r}=""),"",$B{r}+VLOOKUP($G{r},SLAMatrise,3,FALSE)/24)')
    s.cell(row=r, column=18, value=f'=IF($B{r}="","",IF($P{r}<>"",IF($P{r}<=$N{r},"OK","Påbegynt for sent"),IF(NOW()>$N{r},"FORFALT","Innen frist")))')
    s.cell(row=r, column=19, value=f'=IF($B{r}="","",IF($Q{r}<>"",IF($Q{r}<=$O{r},"OK","Løst for sent"),IF(NOW()>$O{r},"FORFALT","Innen frist")))')
    s.cell(row=r, column=20, value=f'=IF(OR($B{r}="",$Q{r}=""),"",ROUND(($Q{r}-$B{r})*24,1))')
    # Skjult hjelpekolonne for søk (Y) – markerer rader i resultatet.
    # Tester C&H (statiske brukerceller), ALDRI B (formel/sirkulær referanse) –
    # ellers dør søket i LO/OO uten iterative referanser. Bindestreker
    # normaliseres bort i både søkeord og tekst, slik at «wifi» treffer «Wi-Fi».
    # Status slås opp i speil-tabellen på «Søk saker» (AA5:AB14), der
    # avkryssingsboksene skriver SANN/USANN – ukjent/blank status = ikke med.
    nål_s = f'TRIM(SUBSTITUTE({SOK_SAK_INPUT},"-",""))'
    tekst_s = f'SUBSTITUTE($H{r}&" "&$I{r}&" "&$U{r},"-","")'
    match_s = (f'AND(ISNUMBER(SEARCH({nål_s},{tekst_s})),'
               f'IFERROR(VLOOKUP($M{r},{SOK_SAK_STATUS},2,FALSE)=TRUE,FALSE))')
    s.cell(row=r, column=25, value=f'=IF(OR($C{r}&$H{r}="",TRIM({SOK_SAK_INPUT})=""),"",IF({match_s},MAX($Y$1:$Y{r-1})+1,""))')
    s.cell(row=r, column=25).font = Font(color="808080")
    for col in (2, 14, 15, 16, 17):
        s.cell(row=r, column=col).number_format = DT
    s.cell(row=r, column=20).number_format = "0.0"

# Eksempeldata
D = datetime.datetime
eksempler = [
    # B opprettet, C innmelder, D epost, E kanal, F kategori, G prioritet, H tittel, I beskrivelse,
    # K tildelt, M status, P påbegynt, Q løst, U løsning, V kbref, W tilkb
    (D(2026, 9, 1, 9, 15), "Morten Vik", "morten.vik@firma.no", "Telefon", "Nettverk", "Høy",
     "Wi-Fi nede i møterom 2", "Ingen kan koble til trådløsnettet i møterom 2. Accesspunktet blinker rødt.",
     "", "Løst", D(2026, 9, 1, 9, 40), D(2026, 9, 1, 10, 30),
     "Restartet accesspunktet og oppdaterte firmware. Nettet er oppe igjen. Se KB-001.", "KB-001", "Nei"),
    (D(2026, 9, 1, 10, 5), "Silje Aas", "silje.aas@firma.no", "E-post", "Passord og tilgang", "Medium",
     "Glemt VPN-passord", "Bruker har glemt VPN-passordet og er låst ute av hjemmekontor.",
     "", "Løst", D(2026, 9, 1, 10, 20), D(2026, 9, 1, 10, 35),
     "Tilbakestilte passord i AD og huket av for endring ved neste pålogging. Se KB-002.", "KB-002", "Nei"),
    (D(2026, 9, 2, 8, 50), "Jonas Bakke", "jonas.bakke@firma.no", "E-post", "E-post og kalender", "Høy",
     "Outlook starter ikke", "Outlook krasjer ved oppstart etter nattens oppdatering.",
     "", "Under behandling", D(2026, 9, 2, 9, 30), None,
     "Tester i sikker modus og oppretter ny profil. Se KB-004.", "", ""),
    (D(2026, 9, 2, 11, 0), "Tone Lie (HR)", "tone.lie@firma.no", "E-post", "Maskinvare", "Lav",
     "Ny PC til nyansatt", "Nyansatt starter 15.09. Trenger laptop med standard programvareinstallasjon.",
     "", "Ny", None, None, "", "", ""),
    (D(2026, 9, 2, 13, 20), "Kontoret 3. etg", "kontor3@firma.no", "Personlig oppmøte", "Skriver og utskrift", "Medium",
     "Papirstopp hele tiden", "Skriveren i 3. etasje kjører fast papir ved dobbeltsidig utskrift.",
     "", "Løst", D(2026, 9, 2, 14, 0), D(2026, 9, 2, 14, 45),
     "Renset valsene og byttet til riktig papirtype. Se KB-003.", "KB-003", "Nei"),
    (D(2026, 9, 3, 8, 0), "Ingrid Mo", "ingrid.mo@firma.no", "Chat", "Programvare", "Kritisk",
     "Teams krasjer i videomøter", "Teams krasjer ved videosamtaler for flere brukere etter siste oppdatering.",
     "", "Under behandling", D(2026, 9, 3, 8, 20), None,
     "Oppdaterer grafikkdriver og Teams på testmaskin. Følger opp iht. KB-006.", "KB-006", ""),
    (D(2026, 9, 3, 9, 45), "Roar Eggen", "roar.eggen@firma.no", "Telefon", "Mobil og telefoni", "Medium",
     "iPhone synkroniserer ikke e-post", "E-post sluttet å synkronisere etter passordbytte i går.",
     "", "Venter på bruker", D(2026, 9, 3, 10, 0), None,
     "Bruker er bedt om å fjerne og legge til e-postkontoen på nytt. Venter på tilbakemelding.", "", ""),
    (D(2026, 9, 3, 12, 30), "Linda Foss", "linda.foss@firma.no", "E-post", "Passord og tilgang", "Høy",
     "Trenger tilgang til fellesmappe", "Ny i teamet og mangler tilgang til fellesområdet på serveren.",
     "", "Løst", D(2026, 9, 3, 13, 0), D(2026, 9, 3, 15, 10),
     "Fikk godkjenning fra leder og la bruker inn i riktig AD-gruppe. Se KB-005.", "KB-005", "Nei"),
    (D(2026, 9, 4, 8, 10), "Ottar Viken", "ottar.viken@firma.no", "Selvbetjening", "Programvare", "Lav",
     "Hjelp til Excel-formel", "Trenger hjelp med FINN.RAD i budsjettarket.",
     "", "Løst", D(2026, 9, 4, 8, 30), D(2026, 9, 4, 9, 0),
     "Viste brukeren FINN.RAD og XLOOKUP med et konkret eksempel.", "", "Ja"),
    (D(2026, 9, 4, 9, 5), "Vaktmester Per", "per.vakt@firma.no", "Telefon", "Nettverk", "Medium",
     "Ødelagt nettverkskabel i 3. etg", "Kabel er kuttet ved renovering. To brukere er uten kablet nett.",
     "", "Ny", None, None, "", "", ""),
]
for i, ex in enumerate(eksempler if MED_EKSEMPLER else []):
    r = 2 + i
    (b, c, d_, e, f, g, h, i_, k, m, p, q, u, v, w) = ex
    s.cell(row=r, column=2, value=b).number_format = DT  # eksempelsaker: statisk dato (overskriver formelen)
    s.cell(row=r, column=3, value=c)
    s.cell(row=r, column=4, value=d_)
    s.cell(row=r, column=5, value=e)
    s.cell(row=r, column=6, value=f)
    s.cell(row=r, column=7, value=g)
    s.cell(row=r, column=8, value=h).alignment = WRAP_TOP
    s.cell(row=r, column=9, value=i_).alignment = WRAP_TOP
    if k:
        s.cell(row=r, column=11, value=k)
    s.cell(row=r, column=13, value=m)
    if p:
        s.cell(row=r, column=16, value=p).number_format = DT
    if q:
        s.cell(row=r, column=17, value=q).number_format = DT
    s.cell(row=r, column=21, value=u).alignment = WRAP_TOP
    if v:
        s.cell(row=r, column=22, value=v)
    if w:
        s.cell(row=r, column=23, value=w)
    s.row_dimensions[r].height = 42

# Ingen tabell-objekt – LO/OO-kompatibelt. AutoFilter gir likevel filtrering.
s.auto_filter.ref = f"A{1}:W{MAX_S}"
s.freeze_panes = "A2"
s.row_dimensions[1].height = 30

widths = {"A": 10, "B": 17, "C": 18, "D": 24, "E": 17, "F": 20, "G": 10, "H": 30, "I": 42,
          "J": 16, "K": 15, "L": 15, "M": 19, "N": 17, "O": 17, "P": 17, "Q": 17,
          "R": 16, "S": 15, "T": 12, "U": 46, "V": 9, "W": 8, "X": 2}
for col, w in widths.items():
    s.column_dimensions[col].width = w
s.column_dimensions["Y"].hidden = True  # hjelpekolonne for søk

# Rullegardinvalidering
def add_dv(ws, formula, rng, inline=False):
    dv = DataValidation(type="list", formula1=formula, allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(rng)
    return dv

add_dv(s, "=Kanaler", "E2:E501")
add_dv(s, "=Kategorier", "F2:F501")
add_dv(s, "=Prioriteter", "G2:G501")
add_dv(s, "=Ansatte", "K2:K501")
add_dv(s, "=Statuser", "M2:M501")
add_dv(s, '"Ja,Nei"', "W2:W501", inline=True)

# NB: showErrorMessage=False (rådgivende, ikke blokkerende). Med hard stopp avvises
# datoer som ikke parses av brukerens lokale Excel/LO-innstillinger (f.eks. en-US
# maskin + norsk datoformat), og feltet framstår som "uedigerbart".
dv_date = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
dv_date.promptTitle = "Dato og klokkeslett"
dv_date.prompt = ("'Opprettet' fylles automatisk når du skriver Innmelder eller Tittel. "
                  "Vil du overstyre: skriv dato og tid slik: 04.09.2026 09:30")
dv_date.showInputMessage = True
dv_date.error = "Ugyldig dato. Bruk formatet 04.09.2026 09:30"
dv_date.errorTitle = "Datoformat"
dv_date.showErrorMessage = False
s.add_data_validation(dv_date)
dv_date.add("B2:B501")
dv_date.add("P2:P501")
dv_date.add("Q2:Q501")

# Betinget formatering
def cf(ws, rng, formula, fill=None, font=None, stop=False):
    ws.conditional_formatting.add(rng, FormulaRule(formula=[formula], fill=fill, font=font, stopIfTrue=stop))

cf(s, "M2:M501", '$M2="Løst"', GREEN_FILL, GREEN_FONT)
cf(s, "M2:M501", '$M2="Lukket"', GRAY_FILL, None)
cf(s, "M2:M501", '$M2="Ny"', BLUE_FILL, None)
cf(s, "M2:M501", 'LEFT($M2,6)="Venter"', YELLOW_FILL, None)
cf(s, "G2:G501", '$G2="Kritisk"', RED_FILL, RED_FONT)
for col in ("R", "S"):
    rng = f"{col}2:{col}501"
    cf(s, rng, f'ISNUMBER(SEARCH("FORFALT",{col}2))', RED_FILL, RED_FONT, stop=True)
    cf(s, rng, f'{col}2="OK"', GREEN_FILL, GREEN_FONT)
    cf(s, rng, f'{col}2="Påbegynt for sent"', ORANGE_FILL, ORANGE_FONT)
    cf(s, rng, f'{col}2="Løst for sent"', ORANGE_FILL, ORANGE_FONT)

# ============================================================
# KUNNSKAPSBASE – FAQ (fritt område)
# ============================================================
k = ws_kb
kb_headers = ["KB-ID", "Tittel", "Kategori", "Problem", "Løsning", "Nøkkelord",
              "Opprettet av", "Dato", "Antall bruk", "Vurdering", "Basert på sak"]
for j, h in enumerate(kb_headers, start=1):
    c = k.cell(row=1, column=j, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT
    c.border = BORDER

MAX_K = 301
for r in range(2, MAX_K + 1):
    k.cell(row=r, column=1, value=f'=IF($B{r}<>"","KB-"&TEXT(ROW()-1,"000"),"")')
    k.cell(row=r, column=8).number_format = DONLY
    # Skjult hjelpekolonne for søk (L) – samme bindestrek-normalisering som i Saker!Y
    match_k = (f'AND(ISNUMBER(SEARCH(TRIM(SUBSTITUTE({SOK_KB_INPUT},"-","")),'
               f'SUBSTITUTE($B{r}&" "&$D{r}&" "&$E{r}&" "&$F{r},"-",""))),'
               f'IF({SOK_KB_KAT}="",1,$C{r}={SOK_KB_KAT}))')
    k.cell(row=r, column=12, value=f'=IF(OR($B{r}="",{SOK_KB_INPUT}=""),"",IF({match_k},MAX($L$1:$L{r-1})+1,""))')
    k.cell(row=r, column=12).font = Font(color="808080")

kb_data = [
    ("Wi-Fi-problemer – kan ikke koble til trådløst nettverk", "Nettverk",
     "Ingen nettforbindelse på trådløsnettet. Accesspunktet kan blinke rødt.",
     "1. Sjekk at Wi-Fi er slått på hos brukeren.\n2. Start accesspunktet på nytt (strøm av/på, vent 30 sek).\n3. Sjekk at riktig SSID og passord brukes.\n4. Oppdater firmware på accesspunktet.\n5. Vedvarende feil: kontakt nettverksansvarlig.",
     "wifi trådløs nettverk tilkobling accesspunkt internett",
     "Ola Nordmann", D(2026, 9, 1), 12, 5, "SAK-0001"),
    ("Tilbakestille passord (Windows/VPN)", "Passord og tilgang",
     "Bruker er låst ute eller har glemt passordet sitt.",
     "1. Verifiser brukerens identitet.\n2. Tilbakestill passordet i Active Directory.\n3. Huk av «Bruker må endre passord ved neste pålogging».\n4. For VPN: synkroniser passordet i VPN-portalen.\n5. Be brukeren teste pålogging.",
     "passord glemt låst vpn innlogging ad",
     "Kari Hansen", D(2026, 9, 1), 25, 5, "SAK-0002"),
    ("Papirstopp i skriver", "Skriver og utskrift",
     "Papir kjører seg fast, ofte ved dobbeltsidig utskrift.",
     "1. Åpne alle luker og fjern fastkjørt papir forsiktig.\n2. Rens valsene med en lofri klut.\n3. Bruk riktig papirtype/grammatur.\n4. Sjekk at papiret ikke er fuktig.\n5. Ved gjentatte stopp: bestill service.",
     "skriver papirstopp utskrift papir fastkjørt",
     "Erik Dahl", D(2026, 9, 2), 8, 4, "SAK-0005"),
    ("Outlook treg eller krasjer ved oppstart", "E-post og kalender",
     "Outlook krasjer eller fryser ved oppstart, ofte etter oppdatering.",
     "1. Start Outlook i sikker modus (outlook.exe /safe).\n2. Deaktiver tillegg ett og ett for å finne synderen.\n3. Opprett en ny e-postprofil.\n4. Kjør hurtigreparasjon av Office.\n5. Sjekk PST/OST-fil med SCANPST ved mistanke om korrupsjon.",
     "outlook e-post krasjer treg profil tillegg fryser",
     "Anne Berg", D(2026, 9, 2), 6, 4, ""),
    ("Gi tilgang til fellesmappe", "Passord og tilgang",
     "Bruker mangler tilgang til en delt nettverksmappe.",
     "1. Innhent skriftlig godkjenning fra mappe-eier eller leder.\n2. Legg brukeren inn i riktig AD-gruppe (aldri direkte på mappen).\n3. Be brukeren logge av og på for å oppdatere tilgangen.\n4. Verifiser tilgangen sammen med brukeren.",
     "tilgang fellesmappe rettigheter mappe ad-gruppe delt",
     "Kari Hansen", D(2026, 9, 3), 15, 5, "SAK-0008"),
    ("Microsoft Teams krasjer ved videosamtale", "Programvare",
     "Teams krasjer eller fryser under videosamtaler, gjerne etter oppdatering.",
     "1. Start PC-en på nytt.\n2. Tøm Teams-bufferen (%appdata%\\Microsoft\\Teams).\n3. Oppdater grafikkdriveren.\n4. Installer Teams på nytt.\n5. Test med nettleserversjonen for å isolere problemet.",
     "teams krasjer video møte samtale kamera",
     "Anne Berg", D(2026, 9, 3), 4, 4, "SAK-0006"),
]
for i, row in enumerate(kb_data if MED_EKSEMPLER else []):
    r = 2 + i
    for j, v in enumerate(row, start=2):
        c = k.cell(row=r, column=j, value=v)
        if j == 8:
            c.number_format = DONLY
        if j in (4, 5, 6):
            c.alignment = WRAP_TOP
    k.row_dimensions[r].height = 105

# Ingen tabell-objekt. AutoFilter gir likevel filtrering.
k.auto_filter.ref = f"A{1}:K{MAX_K}"
k.freeze_panes = "A2"
k.row_dimensions[1].height = 30
for col, w in {"A": 9, "B": 38, "C": 20, "D": 48, "E": 68, "F": 34, "G": 16,
               "H": 12, "I": 12, "J": 11, "K": 14, "L": 2}.items():
    k.column_dimensions[col].width = w
k.column_dimensions["L"].hidden = True  # hjelpekolonne for søk

add_dv(k, "=Kategorier", "C2:C301")
add_dv(k, '"1,2,3,4,5"', "J2:J301", inline=True)
# Rådgivende (ikke blokkerende) – samme grunn som dv_date i Saker.
dv_kbd = DataValidation(type="date", operator="greaterThan", formula1="DATE(2020,1,1)", allow_blank=True)
dv_kbd.promptTitle = "Dato"
dv_kbd.prompt = "Skriv dato slik: 04.09.2026"
dv_kbd.showInputMessage = True
dv_kbd.showErrorMessage = False
k.add_data_validation(dv_kbd)
dv_kbd.add("H2:H301")

# ============================================================
# SØK KB – søk i kunnskapsbasen (egen fane)
# INDEX/MATCH mot skjult hjelpekolonne (Kunnskapsbase!L) – virker i Excel, LO og OO
# ============================================================
q = ws_sok_kb
q["A1"] = "Søk i kunnskapsbasen (FAQ)"
q["A1"].font = TITLE_FONT
q["A2"] = "Tips: Søk her før du åpner en ny sak – kanskje løsningen finnes allerede."
q["A2"].font = NOTE_FONT

KB_COLS = ["A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K"]   # KB-kolonner

q["A4"] = "Søkeord:"
q["B4"].fill = INPUT_FILL
q["B4"].border = BORDER
q["B4"].font = Font(bold=True)
q["A5"] = "Kategori (valgfri):"
q["B5"].fill = INPUT_FILL
q["B5"].border = BORDER
add_dv(q, "=Kategorier", "B5")

btn = q["D4"]
btn.value = "SØK I KB →"
btn.fill = HEADER_FILL
btn.font = Font(color="FFFFFF", bold=True)
btn.alignment = Alignment(horizontal="center", vertical="center")
btn.border = BORDER
btn.hyperlink = "#'Søk KB'!B4"
q.row_dimensions[4].height = 24

for j, h in enumerate(kb_headers, start=1):
    c = q.cell(row=8, column=j, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT

# Antall treff (fra hjelpekolonnen MAX)
q["A7"] = ('=IF($B$4="","Skriv søkeord ovenfor og trykk «SØK I KB →».",'
           'IF(MAX(Kunnskapsbase!$L$2:$L$301)=0,"Ingen treff – prøv andre søkeord.",'
           'MAX(Kunnskapsbase!$L$2:$L$301)&" treff i kunnskapsbasen:"))')
q["A7"].font = NOTE_FONT

for i in range(1, N_KB + 1):
    r = 8 + i
    idx = f"IFERROR(MATCH({i},Kunnskapsbase!$L$2:$L$301,0)+1,\"\")"
    for j, col in enumerate(KB_COLS, start=1):
        q.cell(row=r, column=j, value=f'=IF({idx}="","",INDEX(Kunnskapsbase!${col}:${col},{idx}))')
    q.cell(row=r, column=8).number_format = DONLY
    q.row_dimensions[r].height = 30

sist = 8 + N_KB + 2
q.cell(row=sist, column=1, value="Merk: Søket er ikke avhengig av store/små bokstaver, og bindestreker ignoreres – «wifi» finner «Wi-Fi». Tips: bruk korte enkeltord.").font = NOTE_FONT
q.cell(row=sist + 1, column=1, value=f"Viser inntil {N_KB} treff – snevr inn søket med flere ord ved mange treff. Se også fanen «Søk saker» for løste saker.").font = NOTE_FONT

for col, w in {"A": 14, "B": 34, "C": 20, "D": 48, "E": 68, "F": 34, "G": 16,
               "H": 12, "I": 12, "J": 11, "K": 14}.items():
    q.column_dimensions[col].width = w

# ============================================================
# SØK SAKER – søk i tidligere løste saker (egen fane)
# INDEX/MATCH mot skjult hjelpekolonne (Saker!Y)
# ============================================================
q2 = ws_sok_sak
q2["A1"] = "Søk i tidligere saker"
q2["A1"].font = TITLE_FONT
q2["A2"] = "Søker i tittel, beskrivelse og løsning. Kryss av for statusene som skal tas med (standard: Løst og Lukket)."
q2["A2"].font = NOTE_FONT

SAK_COLS = ["A", "B", "F", "H", "U", "V"]   # Saksnr, Opprettet, Kategori, Tittel, Løsning, KB-ref

q2["A4"] = "Søkeord:"
q2["B4"].fill = INPUT_FILL
q2["B4"].border = BORDER
q2["B4"].font = Font(bold=True)

btn2 = q2["D4"]
btn2.value = "SØK I SAKER →"
btn2.fill = HEADER_FILL
btn2.font = Font(color="FFFFFF", bold=True)
btn2.alignment = Alignment(horizontal="center", vertical="center")
btn2.border = BORDER
btn2.hyperlink = "#'Søk saker'!B4"
q2.row_dimensions[4].height = 24

# Statusfilter – EKTE avkryssingsbokser (skjemakontroller) i toppområdet,
# to kolonner fra kolonne H. openpyxl kan ikke opprette form controls, så
# boksene injiseres i xlsx-pakken av injiser_avkryssingsbokser() etter
# wb.save() – samme klassiske VML-struktur som Excel selv skriver, og som
# LibreOffice/OpenOffice også leser. Hver boks skriver SANN/USANN til en
# koblet celle i skjult kolonne AB; statusnavnene speiles fra Oppsett i
# skjult kolonne AA, og Saker!Y slår opp der (ukjent/blank = ikke med).
# Standard avkrysset: kun Løst og Lukket.
q2["H1"] = "Statuser med i søket:"
q2["H1"].font = Font(bold=True)
for i in range(10):  # speil inntil 10 statuser (Oppsett!L4:L13)
    q2.cell(row=5 + i, column=27, value=f'=IF(Oppsett!$L{4 + i}="","",Oppsett!$L{4 + i})')
for i, st in enumerate(statuser[:10]):
    q2.cell(row=5 + i, column=28, value=st in ("Løst", "Lukket"))  # startverdi for koblet celle
q2.column_dimensions["AA"].hidden = True   # speil: statusnavn
q2.column_dimensions["AB"].hidden = True   # speil: koblede celler (SANN/USANN)

sok_headers = ["Saksnr", "Opprettet", "Kategori", "Tittel", "Løsning", "KB-ref"]
for j, h in enumerate(sok_headers, start=1):
    c = q2.cell(row=7, column=j, value=h)
    c.fill = HEADER_FILL
    c.font = HEADER_FONT

q2["A6"] = ('=IF($B$4="","Skriv søkeord ovenfor og trykk «SØK I SAKER →».",'
            'IF(MAX(Saker!$Y$2:$Y$501)=0,"Ingen saker med valgte statuser matcher søket.",'
            'MAX(Saker!$Y$2:$Y$501)&" treff:"))')
q2["A6"].font = NOTE_FONT

for i in range(1, N_SAK + 1):
    r = 7 + i
    idx = f"IFERROR(MATCH({i},Saker!$Y$2:$Y$501,0)+1,\"\")"
    for j, col in enumerate(SAK_COLS, start=1):
        q2.cell(row=r, column=j, value=f'=IF({idx}="","",INDEX(Saker!${col}:${col},{idx}))')
    q2.cell(row=r, column=2).number_format = DT
    q2.row_dimensions[r].height = 30

sist2 = 7 + N_SAK + 2
q2.cell(row=sist2, column=1, value="Merk: Søket er ikke avhengig av store/små bokstaver, og bindestreker ignoreres – «wifi» finner «Wi-Fi». Tips: bruk korte enkeltord.").font = NOTE_FONT
q2.cell(row=sist2 + 1, column=1, value=f"Viser inntil {N_SAK} treff – snevr inn søket med flere ord ved mange treff. Se også fanen «Søk KB» for ferdige artikler.").font = NOTE_FONT

SOK_SAK_BREDDER = {"A": 14, "B": 18, "C": 20, "D": 48, "E": 68, "F": 12,
                   "G": 2, "H": 22, "I": 8, "J": 2, "K": 22, "L": 8}
for col, w in SOK_SAK_BREDDER.items():
    q2.column_dimensions[col].width = w

# ============================================================
# DASHBOARD – KPI-er øverst, tre kompakte diagrammer under (to på rad 8, ett på rad 26)
# ============================================================
d = ws_dash
d["A1"] = "DASHBOARD – Driftsstøtte"
d["A1"].font = TITLE_FONT
d["A2"] = "Alle tall og diagrammer oppdateres automatisk ut fra arkfanen Saker."
d["A2"].font = NOTE_FONT

# KPI-rad: 6 nøkkeltall i rad 4–5.
# NB: KPI-ene kan IKKE teste på kolonne B (Opprettet) – den er formelstyrt, og
# openpyxl skriver ingen bufrede verdier, så ved åpning (før omberegning) ser
# formelcellene tomme ut for COUNTIFS, og ~490 tomme rader telles med. C
# (Innmelder) er fritekst med statiske verdier og telles korrekt alltid.
kpis = [
    ("A4", "B4", "Åpne saker",
     '=COUNTIFS(Saker!$C$2:$C$501,"<>",Saker!$M$2:$M$501,"<>Løst",Saker!$M$2:$M$501,"<>Lukket")'),
    ("C4", "D4", "Forfalte løsningsfrister", '=COUNTIFS(Saker!$S$2:$S$501,"FORFALT")'),
    ("E4", "F4", "Forfalte startfrister", '=COUNTIFS(Saker!$R$2:$R$501,"FORFALT")'),
    ("G4", "H4", "Løst denne måneden",
     '=COUNTIFS(Saker!$C$2:$C$501,"<>",Saker!$Q$2:$Q$501,">="&DATE(YEAR(TODAY()),MONTH(TODAY()),1))'),
    ("I4", "J4", "Snitt løsningstid (timer)", '=IFERROR(ROUND(AVERAGE(Saker!$T$2:$T$501),1),"-")'),
    ("K4", "L4", "KB-artikler", '=COUNTA(Kunnskapsbase!$B$2:$B$301)'),
]
for lc, vc, label, formula in kpis:
    d[lc] = label
    d[lc].fill = HEADER_FILL
    d[lc].font = KPI_LABEL_FONT
    d[vc].fill = HEADER_FILL
    row = int(lc[1:])
    d[lc.replace(str(row), str(row + 1))] = formula
    d[lc.replace(str(row), str(row + 1))].font = KPI_VALUE_FONT
    d[lc.replace(str(row), str(row + 1))].alignment = Alignment(horizontal="center")

# Hjelpetabeller (kilde for diagrammer) – kolonne A–B, under kakediagrammet
d["A26"] = "Saker per status"
d["A26"].font = SECTION_FONT
for i, st in enumerate(statuser):
    d.cell(row=27 + i, column=1, value=st).border = BORDER
    d.cell(row=27 + i, column=2, value=f'=COUNTIFS(Saker!$M$2:$M$501,$A{27 + i})').border = BORDER

d["A34"] = "Saker per kategori"
d["A34"].font = SECTION_FONT
for i, (kat, _) in enumerate(kategorier):
    d.cell(row=35 + i, column=1, value=kat).border = BORDER
    d.cell(row=35 + i, column=2, value=f'=COUNTIFS(Saker!$F$2:$F$501,$A{35 + i})').border = BORDER

d["C26"] = "Saker per ansvarlig"
d["C26"].font = SECTION_FONT
for i, (navn, _, _) in enumerate(team):
    d.cell(row=27 + i, column=3, value=navn).border = BORDER
    d.cell(row=27 + i, column=4, value=f'=COUNTIFS(Saker!$L$2:$L$501,$C{27 + i})').border = BORDER

# Diagrammer: auto-layout av legend (IKKE ManualLayout – LO/OO tegner da legenden
# oppå diagrammet). Legend nederst, og dataseriene får eksplisitt norsk navn slik
# at legenden aldri viser "Series1". Kake til venstre (A8) med tabeller under
# (A26+), kategori tett inntil til høyre (F8), ansvarlig under kaka (A43). Alt
# skal være synlig uten scrolling.
CHART_H = 9.5   # cm (~18 rader) – litt høyere så bunnlegenden ikke kolliderer med kaka
CHART_W = 16    # cm

# Diagram 1: kakediagram for status – legend nederst, prosent på sektorene
pie = PieChart()
pie.title = "Saker per status"
pie.add_data(Reference(d, min_col=2, min_row=27, max_row=32), titles_from_data=False)
pie.series[0].tx = SeriesLabel(v="Saker")
pie.set_categories(Reference(d, min_col=1, min_row=27, max_row=32))
pie.height = CHART_H
pie.width = CHART_W
pie.legend.position = "b"
pie.legend.overlay = False
d.add_chart(pie, "A8")

# Diagram 2: vertikale søyler for kategori – varyColors gir hver søyle egen farge
# og gjør at legenden viser kategorinavn (ikke serienavnet). Akser skjules fordi
# legenden og tabellen bærer verdiene.
bar1 = BarChart()
bar1.type = "col"
bar1.title = "Saker per kategori"
bar1.add_data(Reference(d, min_col=2, min_row=35, max_row=42), titles_from_data=False)
bar1.series[0].tx = SeriesLabel(v="Saker")
bar1.set_categories(Reference(d, min_col=1, min_row=35, max_row=42))
bar1.height = CHART_H
bar1.width = CHART_W
bar1.varyColors = True
bar1.legend.position = "b"
bar1.legend.overlay = False
bar1.x_axis.delete = True
bar1.y_axis.delete = True
d.add_chart(bar1, "F8")

# Diagram 3: vertikale søyler for ansvarlig – under kategoridiagrammet til høyre
bar2 = BarChart()
bar2.type = "col"
bar2.title = "Saker per ansvarlig"
bar2.add_data(Reference(d, min_col=4, min_row=27, max_row=32), titles_from_data=False)
bar2.series[0].tx = SeriesLabel(v="Saker")
bar2.set_categories(Reference(d, min_col=3, min_row=27, max_row=32))
bar2.height = CHART_H
bar2.width = CHART_W
bar2.varyColors = True
bar2.legend.position = "b"
bar2.legend.overlay = False
bar2.x_axis.delete = True
bar2.y_axis.delete = True
d.add_chart(bar2, "F26")

for col, w in {"A": 22, "B": 12, "C": 22, "D": 12, "E": 22, "F": 12, "G": 22,
               "H": 12, "I": 24, "J": 12, "K": 14, "L": 12, "M": 0, "N": 0,
               "O": 24, "P": 12}.items():
    d.column_dimensions[col].width = w

# Tett rammeverk: smale skillerader mellom grafradene
d.row_dimensions[7].height = 3
d.row_dimensions[25].height = 12
d.row_dimensions[33].height = 12

# ============================================================
# INSTRUKSJONER
# ============================================================
g = ws_instr
innhold = [
    ("title", "Slik bruker du supportsystemet"),
    ("t", "Systemet dekker: ticket-håndtering med unike saksnumre, kunnskapsbase (FAQ), "
          "automatisk tildeling av saker per kategori og SLA-styring med tidsfrister."),
    ("b", ""),
    ("h", "1. Arkfanene"),
    ("t", "• Saker – her logger du alle saker (tickets). Hver rad = én sak med unikt saksnr (SAK-0001, SAK-0002, ...)."),
    ("t", "• Kunnskapsbase – ferdige løsningsartikler (FAQ) slik at brukere/drift kan løse kjente problemer raskt."),
    ("t", "• Søk KB / Søk saker – søk i kunnskapsbasen og i tidligere saker (to separate faner) for å finne løsningen før du starter."),
    ("t", "• Dashboard – oversikt over åpne saker, forfalte frister, kategorier og belastning per ansatt."),
    ("t", "• Oppsett – team, kategorier, SLA-tider, statuser og kanaler. Tilpass til din avdeling."),
    ("b", ""),
    ("h", "2. Registrere ny sak (ticket-håndtering)"),
    ("t", "1) Gå til Saker og fyll ut: Innmelder, E-post, Kanal, Kategori, Prioritet, Tittel, Beskrivelse."),
    ("t", "   Kom saken inn på e-post eller telefon? Lim inn innholdet i Beskrivelse og velg riktig Kanal."),
    ("t", "2) Opprettet (dato+klokkeslett) fylles automatisk idet du skriver Innmelder eller Tittel, og låses."),
    ("t", "   Vil du overstyre (f.eks. sak som egentlig kom i går)? Skriv dato manuelt slik: 04.09.2026 09:30."),
    ("t", "   NB: Overskriver du, forsvinner auto-formelen i den cellen – kopier formelen fra en tom rad for å få den tilbake."),
    ("t", "3) Saksnr lages automatisk (SAK-0001 osv.) – bruk nummeret i all videre kommunikasjon."),
    ("b", ""),
    ("h", "3. Automatisk tildeling"),
    ("t", "• Kolonnen Auto-ansvarlig foreslår riktig person basert på Kategori (koblingen ligger i Oppsett-arket)."),
    ("t", "• Vil du overstyre? Velg en annen person i kolonnen Tildelt til. Kolonnen Ansvarlig viser hvem som faktisk eier saken."),
    ("b", ""),
    ("h", "4. Prioritet og SLA-styring"),
    ("t", "• SLA-tider (Oppsett): Kritisk = start 1t / løs 4t · Høy = 4t/8t · Medium = 8t/24t · Lav = 24t/72t."),
    ("t", "• SLA start frist og SLA løsning frist regnes ut automatisk fra Opprettet + Prioritet."),
    ("t", "• Fyll inn Påbegynt og Løst (dato+kl) etter hvert som du jobber med saken."),
    ("t", "• Farger: RØD «FORFALT» = fristen er passert og saken må gripes nå. ORANSJE = løst/påbegynt for sent. GRØNN «OK» = innen frist."),
    ("t", "• Forfalte saker telles også på Dashboard."),
    ("b", ""),
    ("h", "5. Løse saken og lagre løsningen"),
    ("t", "• Skriv alltid hva som var feilen og hvordan den ble løst i kolonnen Løsning."),
    ("t", "• Sett Status = Løst når brukeren er fornøyd, og Lukket når saken er helt ferdig."),
    ("t", "• Er løsningen beskrevet i en KB-artikkel? Skriv artikkelnummeret i KB-ref (f.eks. KB-001)."),
    ("b", ""),
    ("h", "6. Bygge kunnskapsbasen (FAQ)"),
    ("t", "• Kan problemet ramme flere? Sett Til KB = Ja på saken, og opprett en artikkel i Kunnskapsbase."),
    ("t", "• Kopier/utdyp løsningen, og fyll inn Nøkkelord – de gjør artikkelen søkbar (f.eks. «wifi trådløs nettverk»)."),
    ("t", "• KB-ID lages automatisk. Bruk Antall bruk og Vurdering for å se hvilke artikler som hjelper mest."),
    ("b", ""),
    ("h", "7. Finne løsninger på lignende saker"),
    ("t", "• Søk ligger i to faner: «Søk KB» for kunnskapsbasen (kan avgrenses med kategori) og «Søk saker» for tidligere saker."),
    ("t", "• I «Søk saker» krysser du av i boksene øverst for hvilke statuser som tas med – standard er kun Løst og Lukket."),
    ("t", "• Skriv søkeord i det gule feltet og trykk «SØK I KB →» / «SØK I SAKER →» (knappen setter fokus i feltet)."),
    ("t", "• Resultatene kommer opp automatisk idet du skriver – ingen behov for Enter eller knappetrykk."),
    ("t", "• Fant du en løsning? Bruk KB-ref/Saksnr for å se hele saken/artikkelen."),
    ("b", ""),
    ("h", "8. Tilpass Oppsett direkte i Excel"),
    ("t", "• Alle lister i Oppsett er frie områder med blå overskriftsrader."),
    ("t", "• Legg til: skriv i første ledige rad under listen – formelene følger med automatisk."),
    ("t", "• Fjern: merk raden og trykk Delete, eller høyreklikk > Slett rader."),
    ("t", "• Rullegardinlister, auto-tildeling, SLA-frister og dashboard følger med automatisk."),
    ("b", ""),
    ("h", "9. Krav og tips"),
    ("t", "• Systemet virker i Excel 2007+, LibreOffice Calc og Apache OpenOffice Calc."),
    ("t", "• Søkearkene bruker INDEX og MATCH mot skjulte hjelpekolonner – ingen FILTER/HSTACK kreves."),
    ("t", "• Bruk AutoFilter (pilene i overskriftsraden) på Saker/Kunnskapsbase for å sortere og filtrere."),
    ("t", "• Skal flere jobbe i filen samtidig? Legg den på SharePoint/OneDrive og åpne i Excel med sameksistens."),
    ("t", "• Arbeidsboken har iterativ beregning slått på – det trengs for at auto-tidsstempelet skal låse seg selv."),
    ("t", "  I LibreOffice/OpenOffice må du slå på «Iterative referanser» manuelt (Verktøy > Innstillinger > Calc > Beregn)."),
    ("t", "• Ta backup av filen jevnlig (se egen backup-rutine i prosjektmappen)."),
]
r = 1
for kind, text in innhold:
    c = g.cell(row=r, column=1, value=text)
    if kind == "title":
        c.font = TITLE_FONT
    elif kind == "h":
        c.font = SECTION_FONT
    else:
        c.alignment = Alignment(wrap_text=True, vertical="top")
        if text:
            import math
            g.row_dimensions[r].height = max(15, 15 * math.ceil(len(text) / 115))
    r += 1
g.column_dimensions["A"].width = 118

# ============================================================
# AVKRYSSINGSBOKSER (skjemakontroller) på Søk saker
# openpyxl kan ikke opprette form controls, så boksene injiseres direkte i
# xlsx-pakken etter wb.save(). Strukturen er kopiert XML-for-XML fra en fil
# Excel selv laget (ref2): ark-XMLen får <drawing/>, <legacyDrawing/> og et
# <mc:AlternateContent><mc:Choice Requires="x14"><controls>-blokk der hver
# kontroll ligger i sitt eget mc:AlternateContent. Geometrien Excel faktisk
# bruker ligger i controlPr-<anchor> i ark-XMLen (EMU-verdier), IKKE i VML-en.
# from/to i <controls> skrives UTEN xdr-prefiks – med xdr:-prefiks nekter
# Excel å åpne filen. I drawing2.xml bruker Excel derimot xdr:-prefiks.
# VML-en er fallback-rendering for LibreOffice/OpenOffice.
# ============================================================
import re, os, zipfile

def _esc_xml(t):
    return t.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _ank(from_to_prefiks, kol, rad, til_kol, til_rad, row_off):
    """Bygger <from>/<to>-anker. from_to_prefiks: 'xdr:' for drawing2.xml (som
    Excel skriver der), '' (tom) for <controls> i ark-XMLen (Excel skriver
    from/to UPPREFIXET der, mens barna col/colOff/row/rowOff er xdr:)."""
    f = "<{p}from><xdr:col>{k}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>0</xdr:rowOff></{p}from>"
    t = "<{p}to><xdr:col>{k}</xdr:col><xdr:colOff>368300</xdr:colOff><xdr:row>{r}</xdr:row><xdr:rowOff>{o}</xdr:rowOff></{p}to>"
    return f.format(p=from_to_prefiks, k=kol, r=rad), t.format(p=from_to_prefiks, k=til_kol, r=til_rad, o=row_off)


def injiser_avkryssingsbokser(filnavn, arknavn, bokser):
    """Skriver ekte avkryssingsbokser inn i en allerede lagret xlsx.

    Strukturen er kopiert nøyaktig fra en fil Excel selv laget (sjekket ut
    XML-for-XML): ark-XMLen får <drawing/>, <legacyDrawing/> og et
    <mc:AlternateContent><mc:Choice Requires="x14"><controls>-blokk der hver
    kontroll ligger i sitt eget mc:AlternateContent. Geometrien Excel faktisk
    bruker ligger i controlPr-<anchor> i ark-XMLen (EMU), IKKE i VML.
    from/to i <controls> skrives UTEN xdr-prefiks – med prefiks nekter Excel å
    åpne filen. I drawing2.xml bruker Excel derimot xdr:-prefiks på from/to.
    bokser: liste med dicts – tekst, kol, rad, left_pt, top_pt, til_rad,
    row_off_to, vml_rest ('8, 58, 2, 5'), lenke ('$AB$5'), avkrysset.
    """
    zin = zipfile.ZipFile(filnavn, "r")
    deler = {n: zin.read(n) for n in zin.namelist()}
    zin.close()

    # Finn arkets xml-fil via workbook.xml + workbook-relasjonene
    wbxml = deler["xl/workbook.xml"].decode("utf-8")
    wrels = deler["xl/_rels/workbook.xml.rels"].decode("utf-8")
    sheet_tag = re.search(r'<sheet[^>]*name="%s"[^>]*/?>' % re.escape(arknavn), wbxml).group(0)
    rid = re.search(r'r:id="(rId\d+)"', sheet_tag).group(1)
    rel_tag = re.search(r'<Relationship[^>]*Id="%s"[^>]*/>' % rid, wrels).group(0)
    target = re.search(r'Target="([^"]+)"', rel_tag).group(1)
    arkfil = target.lstrip("/") if target.startswith("/") else "xl/" + target

    def ledig(mal):
        i = 1
        while mal.format(i) in deler:
            i += 1
        return mal.format(i)

    # Relasjonsfil for arket (utvid eksisterende eller lag ny)
    relfil = "xl/worksheets/_rels/" + arkfil.split("/")[-1] + ".rels"
    if relfil in deler:
        neste = max(int(x) for x in re.findall(r'Id="rId(\d+)"', deler[relfil].decode("utf-8"))) + 1
    else:
        neste = 1
    nye_rel = []

    drw_fil = ledig("xl/drawings/drawing{}.xml")
    rid_drw = f"rId{neste}"
    neste += 1
    vml_fil = ledig("xl/drawings/vmlDrawing{}.vml")
    rid_vml = f"rId{neste}"
    neste += 1
    nye_rel.append((rid_drw,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing",
                    "../drawings/" + drw_fil.split("/")[-1]))
    nye_rel.append((rid_vml,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing",
                    "../drawings/" + vml_fil.split("/")[-1]))

    shapes, kontroller, drw_ankre = [], [], []
    for i, b in enumerate(bokser):
        sid = 1025 + i
        cp_fil = ledig("xl/ctrlProps/ctrlProp{}.xml")
        rid_cp = f"rId{neste}"
        neste += 1
        nye_rel.append((rid_cp,
                        "http://schemas.openxmlformats.org/officeDocument/2006/relationships/ctrlProp",
                        "../ctrlProps/" + cp_fil.split("/")[-1]))
        lenke_full = f"'{arknavn}'!{b['lenke']}"
        # ctrlProp: checked-attributt kun for avkryssede bokser (som Excel)
        chk_attr = ' checked="Checked"' if b["avkrysset"] else ""
        deler[cp_fil] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                         '<formControlPr xmlns="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main" '
                         f'objectType="CheckBox"{chk_attr} fmlaLink="{_esc_xml(lenke_full)}" lockText="1" noThreeD="1"/>').encode("utf-8")
        kol, rad = b["kol"], b["rad"]
        til_kol = kol + 1
        # Ark-XML (controls): from/to uprefikset; drawing2.xml: xdr: prefiks
        af, at = _ank("", kol, rad, til_kol, b["til_rad"], b["row_off_to"])
        df, dt = _ank("xdr:", kol, rad, til_kol, b["til_rad"], b["row_off_to"])
        # VML-shape (fallback-rendering for LO/OO) – Anchor-verdier som Excel
        vml_checked = "   <x:Checked>1</x:Checked>\n" if b["avkrysset"] else ""
        shapes.append(f'''<v:shape id="_x0000_s{sid}" type="#_x0000_t201" style='position:absolute;
  margin-left:{b["left_pt"]}pt;margin-top:{b["top_pt"]}pt;width:150pt;height:17pt;z-index:{i + 1};
  mso-wrap-style:tight' filled="f" fillcolor="windowText [64]" stroked="f"
  strokecolor="window [65]" strokeweight="3e-5mm" o:insetmode="auto">
  <v:fill color2="window [65]"/>
  <v:path shadowok="t" strokeok="t" fillok="t"/>
  <o:lock v:ext="edit" rotation="t"/>
  <v:textbox style='mso-direction-alt:auto' o:singleclick="f">
   <div style='text-align:left'><font face="Segoe UI" size="160" color="auto">{_esc_xml(b["tekst"])}</font></div>
  </v:textbox>
  <x:ClientData ObjectType="Checkbox">
   <x:SizeWithCells/>
   <x:Anchor>
    {kol}, 0, {rad}, 0, {b["vml_rest"]}</x:Anchor>
   <x:AutoFill>False</x:AutoFill>
   <x:AutoLine>False</x:AutoLine>
   <x:TextVAlign>Center</x:TextVAlign>
{vml_checked}   <x:FmlaLink>{_esc_xml(lenke_full)}</x:FmlaLink>
   <x:NoThreeD/>
  </x:ClientData>
 </v:shape>''')
        # x14-kontroll i ark-XMLen (fra/to uprefikset – ellers nekter Excel)
        kontroller.append(
            f'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'<mc:Choice Requires="x14"><control shapeId="{sid}" r:id="{rid_cp}" name="Check Box {i + 1}">'
            f'<controlPr defaultSize="0" autoFill="0" autoLine="0" autoPict="0">'
            f'<anchor moveWithCells="1">{af}{at}</anchor></controlPr></control>'
            f'</mc:Choice></mc:AlternateContent>')
        # Skjult shape i drawing-parten (speiler ankeret; koblet til VML via spid)
        guid = "%012x" % sid
        drw_ankre.append(
            f'<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
            f'<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">'
            f'<xdr:twoCellAnchor editAs="oneCell">{df}{dt}'
            f'<xdr:sp macro="" textlink=""><xdr:nvSpPr><xdr:cNvPr id="{sid}" name="Check Box {i + 1}" hidden="1">'
            f'<a:extLst><a:ext uri="{{63B3BB69-23CF-44E3-9099-C40C66FF867C}}"><a14:compatExt spid="_x0000_s{sid}"/></a:ext>'
            f'<a:ext uri="{{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}}"><a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{{CB9B2148-2D25-A1F6-351E-{guid}}}"/></a:ext></a:extLst></xdr:cNvPr>'
            f'<xdr:cNvSpPr/></xdr:nvSpPr>'
            f'<xdr:spPr bwMode="auto"><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></a:xfrm>'
            f'<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln><a:effectLst/>'
            f'<a:extLst><a:ext uri="{{909E8E84-426E-40DD-AFC4-6F175D3DCCD1}}"><a14:hiddenFill><a:solidFill><a:srgbClr val="000000" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="64"/></a:solidFill></a14:hiddenFill></a:ext>'
            f'<a:ext uri="{{91240B29-F687-4F45-9708-019B960494DF}}"><a14:hiddenLine w="1"><a:solidFill><a:srgbClr val="FFFFFF" mc:Ignorable="a14" a14:legacySpreadsheetColorIndex="65"/></a:solidFill><a:miter lim="800000"/><a:headEnd/><a:tailEnd type="none" w="med" len="med"/></a14:hiddenLine></a:ext>'
            f'<a:ext uri="{{AF507438-7753-43E0-B8FC-AC1667EBCBE1}}"><a14:hiddenEffects><a:effectLst><a:outerShdw dist="35921" dir="2700000" algn="ctr" rotWithShape="0"><a:srgbClr val="808080"/></a:outerShdw></a:effectLst></a14:hiddenEffects></a:ext></a:extLst></xdr:spPr>'
            f'<xdr:txBody><a:bodyPr vertOverflow="clip" wrap="square" lIns="36576" tIns="32004" rIns="0" bIns="32004" anchor="ctr" upright="1"/>'
            f'<a:lstStyle/><a:p><a:pPr algn="l" rtl="0"><a:defRPr sz="1000"/></a:pPr>'
            f'<a:r><a:rPr lang="en-US" sz="800" b="0" i="0" u="none" strike="noStrike" baseline="0">'
            f'<a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="Segoe UI"/><a:cs typeface="Segoe UI"/></a:rPr>'
            f'<a:t>{_esc_xml(b["tekst"])}</a:t></a:r></a:p></xdr:txBody></xdr:sp><xdr:clientData/></xdr:twoCellAnchor>'
            f'</mc:Choice><mc:Fallback/></mc:AlternateContent>')

    deler[vml_fil] = ('<xml xmlns:v="urn:schemas-microsoft-com:vml"\n'
                      ' xmlns:o="urn:schemas-microsoft-com:office:office"\n'
                      ' xmlns:x="urn:schemas-microsoft-com:office:excel">\n'
                      ' <o:shapelayout v:ext="edit">\n'
                      '  <o:idmap v:ext="edit" data="1"/>\n'
                      ' </o:shapelayout><v:shapetype id="_x0000_t201" coordsize="21600,21600" o:spt="201"\n'
                      '  path="m,l,21600r21600,l21600,xe">\n'
                      '  <v:stroke joinstyle="miter"/>\n'
                      '  <v:path shadowok="f" o:extrusionok="f" strokeok="f" fillok="f" o:connecttype="rect"/>\n'
                      '  <o:lock v:ext="edit" shapetype="t"/>\n'
                      ' </v:shapetype>' + "".join(shapes) + "</xml>").encode("utf-8")

    deler[drw_fil] = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                      '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" '
                      'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">'
                      + "".join(drw_ankre) + "</xdr:wsDr>").encode("utf-8")

    # Ark-XML: nødvendige namespaces på rot + <drawing>, <legacyDrawing>,
    # <controls> (x14) – samme rekkefølge og struktur som Excel skriver.
    sheet = deler[arkfil].decode("utf-8")
    for ns in ('xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"',
               'xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"',
               'xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"'):
        prefix = ns.split("=")[0]
        i_ws = sheet.index("<worksheet")
        rot_tag = sheet[i_ws:sheet.index(">", i_ws)]
        if prefix + "=" not in rot_tag:
            i_ws = sheet.index("<worksheet")
            sheet = sheet[:i_ws] + "<worksheet " + ns + sheet[i_ws + len("<worksheet"):]
    kontrollblokk = ('<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
                     '<mc:Choice Requires="x14"><controls>' + "".join(kontroller) +
                     '</controls></mc:Choice></mc:AlternateContent>')
    sheet = sheet.replace("</worksheet>",
                          f'<drawing r:id="{rid_drw}"/><legacyDrawing r:id="{rid_vml}"/>{kontrollblokk}</worksheet>')
    deler[arkfil] = sheet.encode("utf-8")

    rel_linjer = "".join(f'<Relationship Id="{i_}" Type="{t}" Target="{g}"/>' for i_, t, g in nye_rel)
    if relfil in deler:
        rxml = deler[relfil].decode("utf-8").replace("</Relationships>", rel_linjer + "</Relationships>")
    else:
        rxml = ('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n'
                '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
                + rel_linjer + "</Relationships>")
    deler[relfil] = rxml.encode("utf-8")

    ct = deler["[Content_Types].xml"].decode("utf-8")
    if 'Extension="vml"' not in ct:
        ct = ct.replace('<Default Extension="xml"',
                        '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
                        '<Default Extension="xml"', 1)
    overrides = "".join(f'<Override PartName="/{n}" ContentType="application/vnd.ms-excel.controlproperties+xml"/>'
                        for n in deler if n.startswith("xl/ctrlProps/"))
    overrides += f'<Override PartName="/{drw_fil}" ContentType="application/vnd.openxmlformats-officedocument.drawing+xml"/>'
    ct = ct.replace("</Types>", overrides + "</Types>")
    deler["[Content_Types].xml"] = ct.encode("utf-8")

    midlertidig = filnavn + ".tmp"
    with zipfile.ZipFile(midlertidig, "w", zipfile.ZIP_DEFLATED) as zut:
        for n, data in deler.items():
            zut.writestr(n, data)
    os.replace(midlertidig, filnavn)


# Boks-layout: verdier er tatt rett fra filen Excel selv laget for dette arket
# (boksene CB7-CB12 i referansen, som satt i H2/K2, H3/K3, H4/K4):
#   (kol, rad, left_pt, top_pt, til_rad, row_off_to, vml_rest)
# row_off_to i EMU (31750 for rad 1-2, 215900 for rad 3); vml_rest er siste
# fire tall i VML-Anchor (kol2, kol2-offset, rad2, rad2-offset) som Excel satte.
_BOKS_DEF = [
    (7,  1, 1001, 21.0, 2, 31750, "8, 58, 2, 5"),
    (10, 1, 1177, 21.0, 2, 31750, "11, 58, 2, 5"),
    (7,  2, 1001, 35.5, 3, 31750, "8, 58, 3, 5"),
    (10, 2, 1177, 35.5, 3, 31750, "11, 58, 3, 5"),
    (7,  3, 1001, 50.0, 3, 215900, "8, 58, 3, 34"),
    (10, 3, 1177, 50.0, 3, 215900, "11, 58, 3, 34"),
]
avkryssingsbokser = []
for i, st in enumerate(statuser[:6]):
    kol, rad, left_pt, top_pt, til_rad, roff, vml_rest = _BOKS_DEF[i]
    avkryssingsbokser.append({
        "tekst": st,
        "kol": kol,
        "rad": rad,
        "left_pt": left_pt,
        "top_pt": top_pt,
        "til_rad": til_rad,
        "row_off_to": roff,
        "vml_rest": vml_rest,
        "lenke": f"$AB${5 + i}",
        "avkrysset": st in ("Løst", "Lukket"),
    })

wb.save(OUT)
injiser_avkryssingsbokser(OUT, "Søk saker", avkryssingsbokser)
modus = "med eksempeldata" if MED_EKSEMPLER else "tomt (uten eksempeldata)"
print(f"OK – {OUT} generert, {modus}")
