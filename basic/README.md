# Basic-makroer for LibreOffice / OpenOffice (`.ods`)

Denne mappen inneholder StarBasic-kildekoden for makroversjonen av
supportsystemet (`Supportsystem.ods`). Dette er kildekoden som T11 pakker inn
i `.ods`-filen — her forfattes den bare (kildeforfatting, ingen kjøring).

## Moduler

| Fil | Innhold (Sub/Function) | Formål |
|---|---|---|
| `ModFelles.bas` | `Normaliser`, `Levenshtein`, `ScoreDok` | Felles hjelpefunksjoner: normalisering, Levenshtein, scoring. Matcher `kontrakt.score` eksakt. |
| `ModTidsstempel.bas` | `StemplTidsstempel` | Tidsstempel i Saker!B (knapp-triggeret, LO har ingen `Worksheet_Change`). |
| `ModSok.bas` | `SokKB`, `SokSaker` + hjelpere (`ErBedre`, `SorterTreff`, `DatoSomTall`, `MarkerTreff`, m.fl.) | Søk i kunnskapsbase og saker, rangert (score → dato → rad). |
| `ModDashboard.bas` | `OppdaterDashboard`, `TellSaker` | Reberegner dashboard-hjelpetabellene fra Saker. |

## Dialekt og nøkkelfakta (StarBasic/UNO)

- Gjeldende dokument heter **`ThisComponent`** (ikke `Application`/`Workbook`).
- Ark: `ThisComponent.Sheets.getByName("Saker")`.
- **`getCellByPosition(col, rad)` er 0-basert** i både kolonne og rad.
  Saker!B2 = `getCellByPosition(1, 1)`. Alternativt `getCellRangeByName("B2")`
  (1-basert A1-notasjon).
- **Rask bulk-lesing**: `arr = oSheet.getCellRangeByName("A2:W501").getDataArray()`
  gir en 0-basert 2D-array (rad, kolonne). Brukes i stedet for celle-for-celle,
  som er tregt i LO Basic (~200 ms/500 rader).
- **`getDataArray()` returnerer formel-resultater** (ikke formlene), og datoer
  som Double (internt serienummer) — derfor brukes `CDbl` for tiebreak.
- **`RGB(red, green, blue)` i StarBasic = `red*65536 + green*256 + blue`**
  (0xRRGGBB), som er UNO-fargeformatet direkte (motsatt av VBA, der RGB er
  0xBBGGRR). `CellBackColor = 16771996` = `RGB(255, 235, 156)` = `FFEB9C`
  (lys gul, samme fyll som resten av arbeidsboken).
- **Fet skrift**: `CharWeight = 150` (`com.sun.star.awt.FontWeight.BOLD`).
- **Ingen `Worksheet_Change`/`Application.*`** — søk og tidsstempel er
  knapp-triggeret Subs. Se «Knappetilordning» under.
- `InStr` i StarBasic skiller ikke store/små bokstaver (i motsetning til VBA,
  som trenger `vbTextCompare`); Normaliser bruker likevel `LCase` for å
  matche `kontrakt.normalize` eksakt.

## Scoring (må matche `kontrakt.score` eksakt)

- Vekter: tittel=4, nøkkelord=3, problem=2, løsning=1.
- `Normaliser` **fjerner** bindestreker (`Replace(t, "-", "")`), gjør om til
  små bokstaver og kollapser whitespace-runs. Krav:
  `Normaliser("Wi-Fi nede") = "wifi nede"`.
- Eksakt delstreng-treff først; fuzzy (Levenshtein ≤ 2) kun når total score
  er 0. `Exit For` bryter kun ord-løkka (flere felt kan gi vekt).

## Hvordan T11 pakker dette inn i `.ods`

`soffice` kan ikke lese `.bas`-filer direkte; Basic-makroer i en `.ods` ligger
som XML i zip-stien `Basic/`. T11 må:

1. Generere `.xlsx`-skjelettet (via `lag_supportsystem.py --target calc`) og
   konvertere det til `.ods` med `soffice --headless --convert-to ods`.
2. For hver modul (i rekkefølge) legge innholdet av `.bas`-filen i
   `Basic/Standard/ModuleN.xml` inne i et `<script:source>`-element:
   ```xml
   <script:module xmlns:script="http://openoffice.org/2000/script"
       script:name="ModFelles" script:language="StarBasic">
     <script:source encoding="UTF-8"><![CDATA[ ...kildekoden... ]]></script:source>
   </script:module>
   ```
3. Legge til `Basic/Standard/script-lb.xml` (bibliotekbeskrivelse) og
   `Basic/script-lc.xml` (bibliotekbeholder, `script:name="Standard"`).

Merk: Modulinnholdet skrives som UTF-8 (norske tegn: æ, ø, å, «», →, —).

## Knappetilordning

StarBasic har ingen `Worksheet_Change`. Tilordning gjøres i Calc-UI eller ved
etterarbeid:

- **Søk**: Sett en trykknapp (Skjema > Trykknapp) over «SØK I KB →» (D4) og
  «SØK I SAKER →» (D4), knytt hendelsen «Utfør handling» til henholdsvis
  `ModSok.SokKB` og `ModSok.SokSaker`.
- **Tidsstempel**: Knapp på Saker-arket knyttet til `ModTidsstempel.StemplTidsstempel`.
- **Dashboard**: Knapp på Dashboard-arket knyttet til `ModDashboard.OppdaterDashboard`.

Alternativt kan makroene tilordnes dokumenthendelser (Verktøy > Tilpass >
Hendelser) for sanntids-/auto-oppførsel — se kommentaren «BEST-EFFORT
LISTENER» i `ModSok.bas`. Dette er bevisst utsatt til T11/T12.

## Diagram-oppdatering

Diagrammene i `.ods`-en er levende Calc-diagrammer som refererer direkte til
hjelpetabellene (A27:B32, A35:B42, C27:D32). `OppdaterDashboard` skriver bare
nye tall til disse cellene, og diagrammene oppdateres automatisk — ingen
UNO-diagram-API nødvendig. Diagramobjekter flyttes aldri (hard begrensning).

## Status og verifikasjon

- Kildekoden er **forfattet headless** (soffice er ikke installert i dette
  miljøet). Den er derfor **ikke kompilert/verifisert** ennå — det skjer i
  T12/T15 når `.ods`-pakkingen og `soffice` er på plass.
- `Option Explicit` er bevisst utelatt for å unngå at én enkelt
  udeklarert-variabel-typo bryter hele biblioteket under headless-forfatting;
  T12/T15 bør vurdere å slå det på etter kompilering.
- Scoring-kryssjekk (Python-motpart mot `kontrakt.score`) ligger i
  `.sisyphus/evidence/task-10-score.txt`.

## Se også

- `kontrakt.py` — ark-rekkefølge, celle-referanser, scoring-orakel.
- `vba/*.bas` — tilsvarende VBA-implementasjon (samme logikk, annen dialekt).
- `.sisyphus/evidence/task-10-basic.txt` — modul-/funksjonsoversikt.
