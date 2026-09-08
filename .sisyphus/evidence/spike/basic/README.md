# Spike T4: LibreOffice Basic-injeksjon i .ods

Denne mappen inneholder resultatet av task T4: en proof-of-concept som viser
at vi kan injisere en StarBasic-modul i en openpyxl-generert `.xlsx` som er
konvertert til `.ods`, og at den resulterende `.ods`-filen fortsatt åpnes
av LibreOffice uten reparasjon.

## Hva som er bevist

- openpyxl kan generere en makrofri `.xlsx`.
- `soffice --headless --convert-to ods` konverterer `.xlsx` → `.ods`.
- Vi kan skrive `Basic/`-XML inn i `.ods`-zipen og patche `META-INF/manifest.xml`.
- Etter injeksjon tåler `.ods`-filen en ny headless konvertering
  `.ods` → `.xlsx`; LibreOffice ser altså ikke noen korrupsjon.
- Basic-modulen overlever en ODS → ODS-runde (bekreftet ved at
  `Basic/Standard/Module1.xml` fortsatt finnes og er uendret).

## Filer

| Fil | Beskrivelse |
|---|---|
| `injiser_basic.py` | Gjenbrukbart skript for å injisere én `.bas`-modul i en `.ods`. |
| `artifacts/min.xlsx` | Minimal openpyxl-arbeidsbok brukt i spiken. |
| `artifacts/injected.ods` | `.ods` etter Basic-injeksjon. |
| `artifacts/roundtrip/injected.xlsx` | Resultat av `.ods` → `.xlsx` med soffice. |

## Injeksjonsoppskrift

### 1. Lag minimal `.xlsx`

```python
from openpyxl import Workbook
wb = Workbook()
ws = wb.active
ws.title = "Demo"
ws["A1"] = "hello"
ws["A2"] = "world"
wb.save("min.xlsx")
```

### 2. Konverter til `.ods`

```powershell
& "C:\Program Files\LibreOffice\program\soffice.exe" `
  --headless --convert-to ods --outdir . min.xlsx
```

En ren `.ods` fra soffice inneholder **ingen** `Basic/`-mappe.
`META-INF/manifest.xml` ser slik ut (kun relevante linjer):

```xml
<manifest:file-entry manifest:full-path="/"
  manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/>
<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
<!-- ... ingen Basic-linjer ... -->
```

### 3. Legg til Basic-XML

Inne i `.ods`-zipen (zip-rewrite, se `injiser_basic.py`) legges tre filer:

**`Basic/script-lc.xml`** — bibliotekbeholder (library container):

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:library-container PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "script.dtd">
<script:library-container xmlns:script="http://openoffice.org/2000/script">
  <script:library script:name="Standard" script:language="StarBasic"/>
</script:library-container>
```

**`Basic/Standard/script-lb.xml`** — biblioteksindeks (library index):

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "script.dtd">
<script:library xmlns:script="http://openoffice.org/2000/script"
    script:name="Standard" script:language="StarBasic" script:readonly="false">
  <script:module script:name="Module1" script:language="StarBasic"/>
</script:library>
```

**`Basic/Standard/Module1.xml`** — selve kilden:

```xml
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "script.dtd">
<script:module xmlns:script="http://openoffice.org/2000/script"
    script:name="Module1" script:language="StarBasic">
<script:source encoding="UTF-8"><![CDATA[Sub Hello()
    MsgBox "hei"
End Sub
]]></script:source>
</script:module>
```

### 4. Patch `META-INF/manifest.xml`

**Ja, manifest.xml MÅ patches.** Uten disse linjene feiler headless-konvertering
(tomt output-dir, ingen `.xlsx` produsert).

Legg til rett før `</manifest:manifest>`:

```xml
<manifest:file-entry manifest:full-path="Basic/" manifest:media-type="application/binary"/>
<manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="Basic/Standard/" manifest:media-type="application/binary"/>
<manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
<manifest:file-entry manifest:full-path="Basic/Standard/Module1.xml" manifest:media-type="text/xml"/>
```

### 5. Verifiser med soffice round-trip

```powershell
& "C:\Program Files\LibreOffice\program\soffice.exe" `
  --headless --convert-to xlsx --outdir roundtrip injected.ods
```

Sjekk at `roundtrip/injected.xlsx` finnes og kan åpnes med openpyxl:

```python
from openpyxl import load_workbook
wb = load_workbook("roundtrip/injected.xlsx")
print(wb.sheetnames)   # ['Demo']
print(wb["Demo"]["A1"].value)  # 'hello'
```

## Bruk av `injiser_basic.py`

```powershell
# Spike med innebygd Sub Hello
python .sisyphus/evidence/spike/basic/injiser_basic.py min.ods injected.ods --verify

# Injisere ekte produksjonsmodul (f.eks. ModSok.bas)
python .sisyphus/evidence/spike/basic/injiser_basic.py `
  min.ods injected.ods --bas basic/ModSok.bas --modul ModSok --verify

# Overstyre soffice-sti
$env:SOFFICE = "C:\Program Files\LibreOffice\program\soffice.exe"
python .sisyphus/evidence/spike/basic/injiser_basic.py min.ods injected.ods --verify
```

soffice-sti leses i prioritert rekkefølge:

1. `--soffice` argument
2. `SOFFICE` miljøvariabel
3. `C:\Program Files\LibreOffice\program\soffice.exe`

## Hva som er IKKE i denne spiken

- **Ingen knappetilordning**. Å koble makroer til skjema-knapper eller
  dokumenthendelser er deferred til T11/T12.
- **Ingen produksjons-Basic**. Bare `Sub Hello()` for å bevise at
  injeksjonsmekanismen fungerer.
- **Ingen `.ods`-generering fra supportsystem-skjelettet**. Det er T11s jobb.

## Lærdommer til T11

- Zip-rewrite-mønsteret fra `build_common.py` (`injiser_avkryssingsbokser` /
  `injiser_vba_prosjekt`) fungerer utmerket også for `.ods`:
  les alle parts, muter dict, skriv ny zip, `os.replace`.
- `META-INF/manifest.xml` må ha `manifest:file-entry` for hver enkelt
  `Basic/...`-fil (og katalogene). Dette er det viktigste detaljen.
- `Basic/script-lc.xml` og `Basic/Standard/script-lb.xml` trenger kun å peke
  på bibliotek/modul; selve kilden ligger i `Basic/Standard/<Modul>.xml`.
- For produksjon: plasser hver `.bas`-fil i `Basic/Standard/ModFelles.xml`,
  `ModTidsstempel.xml`, `ModSok.xml`, `ModDashboard.xml`, og oppdater
  `script-lb.xml` med alle modulene.
