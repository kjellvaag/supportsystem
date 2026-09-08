# T3 VBA-injeksjon — oppskrift og bevis

Denne mappen inneholder proof-of-concept for å injisere en ferdig
`vbaProject.bin` (fra Excel COM) inn i en openpyxl-generert `.xlsx`, slik at
resultatet blir en gyldig `.xlsm` som Excel åpner uten reparasjon og der
makroene er kjørbare.

## Filer

| Fil | Formål |
|-----|--------|
| `lag_donor.py` | Lager `donor.xlsm` med `Sub SayHello()` via Excel COM/pywin32 og ekstraherer `xl/vbaProject.bin`. |
| `injiser_vba.py` | Lager minimal `.xlsx` med openpyxl, injiserer `vbaProject.bin` og patcher XML, verifiserer med både openpyxl og Excel COM. |
| `donor.xlsm` | Midlertidig donor-fil (genereres). |
| `spike_inn.xlsx` | Minimal openpyxl-arbeidsbok før injeksjon (genereres). |
| `spike.xlsm` | Ferdig makro-aktivert arbeidsbok (genereres). |
| `vbaProject.bin` | Ekstrahert VBA-binær (genereres). |

## Forutsetninger

- Python 3.9+ med `openpyxl` og `pywin32` (`pip install openpyxl pywin32`).
- Microsoft Excel installert (her: Office 365 C2R x64, Office16).
- Excel må tillate programmatisk tilgang til VB-prosjektet:
  - Registry: `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1` (DWORD).
  - Alternativt: Fil > Alternativer > Senter for klarering av klarering > Innstillinger > Makroinnstillinger > "Klarer tilgang til VBA-prosjektobjektmodellen".

## Kjøring

```powershell
python lag_donor.py    # lager donor.xlsm + vbaProject.bin
python injiser_vba.py  # lager spike.xlsm og verifiserer
```

## Injeksjonsoppskrift

Gitt en openpyxl-lagret `.xlsx` og en ekstrahert `vbaProject.bin`:

### 1. Legg `vbaProject.bin` inn i zip-pakken under navnet `xl/vbaProject.bin`.

### 2. Patch `xl/_rels/workbook.xml.rels`

Legg til en `Relationship` som peker på `vbaProject.bin`.  Bruk neste ledige
`rId` (her `rId4` fordi standard openpyxl-arbeidsbok bruker `rId1`–`rId3`):

```xml
<Relationship Id="rId4"
  Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject"
  Target="vbaProject.bin"/>
```

### 3. Patch `[Content_Types].xml`

a) Endre `ContentType` for `xl/workbook.xml` fra vanlig `.xlsx` til
   makro-aktivert workbook:

```xml
<!-- FØR -->
<Override PartName="/xl/workbook.xml"
  ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>

<!-- ETTER -->
<Override PartName="/xl/workbook.xml"
  ContentType="application/vnd.ms-excel.sheet.macroEnabled.main+xml"/>
```

b) Legg til `Override` for selve VBA-binæren:

```xml
<Override PartName="/xl/vbaProject.bin"
  ContentType="application/vnd.ms-office.vbaProject"/>
```

### 4. Skriv zip-en på nytt med `.xlsm`-endelse.

openpyxl må åpnes med `keep_vba=True` for å bevare `vbaProject.bin` i
fremtidige runder:

```python
wb = load_workbook("spike.xlsm", keep_vba=True)
```

## Verifisering

Spikens `injiser_vba.py` sjekker to ting:

1. `openpyxl.load_workbook(..., keep_vba=True)` inneholder
   `xl/vbaProject.bin` i `wb.vba_archive.namelist()`.
2. Excel COM åpner `spike.xlsm` uten reparasjon og `VBProject.VBComponents`
   inneholder modulen `SpikeModul` med `Sub SayHello()`.

## Kjente forhold

- Makroer er deaktivert som standard i Excel; brukeren må aktivere redigering
  for å kjøre `SayHello`.
- `vbaProject.bin` inneholder koblinger til ark-codenavn.  For produksjon må
  ark-codenavn (`codeName`) fryses i `kontrakt.py` slik at donor matcher
  målarbeidsboken.
- For å unngå hengende Excel-prosesser setter skriptene `Visible=False`,
  `DisplayAlerts=False` og kaller `app.Quit()` i en `finally`-blokk.

## Oppsummering

Injeksjonen fungerer.  Med riktig patching av `[Content_Types].xml` og
`xl/_rels/workbook.xml.rels` åpnes openpyxl-arbeidsboken i Excel som en
makro-aktivert `.xlsm` uten reparasjonsdialog, og makroen `SayHello` er
syntaks-messig til stede og kan kjøres etter aktivering.
