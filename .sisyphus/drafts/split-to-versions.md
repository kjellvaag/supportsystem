# Draft: Splitte supportsystemet i to versjoner (Excel + LO/OO) med makroer

## Krav (bekreftet av bruker)
- Splitte løsningen i to separate, optimaliserte versjoner (MS Excel + LO/OO)
- **Prioritet**: Excel først, LO/OO etterpå
- **Byggmiljø**: Har Excel installert på byggmaskinen
- **Kodestruktur**: Én generator med `--target excel|calc` (felles kjerne + branch)
- **Søk-optimalisering**: sanntid type-ahead + relevans-rangering + fuzzy-søk + markering av trefford
- **Dashboard-dynamikk**: auto-resize diagrammer + auto-reposisjonering + auto-oppfriskning

## Tekniske realiteter (fakta, ikke beslutninger)
- Excel-makroer = VBA, lagres i `.xlsm`/`.xlsb`
- LO/OO-makroer = StarBasic/LibreOffice Basic, lagres i `.ods` (eller brukerprofil)
- VBA og Basic er IKKE kompatible — to helt forskjellige språk og objektmodeller
- openpyxl kan ikke skrive makroer — krever injeksjon av `vbaProject.bin` (Excel) eller Basic-bibliotek (ODS)
- Dagens løsning er 100 % makrofri, ~1016 linjer i `lag_supportsystem.py`
- Avkryssingsbokser injiseres allerede via rå zip/XML-manipulering post-save
- Backup-subsystemet (backup.py/service/config) peker på `Supportsystem.xlsx`

## Nåværende søk (formelbasert)
- INDEX/MATCH mot skjulte hjelpekolonner (Saker!Y, Kunnskapsbase!L)
- SEARCH + SUBSTITUTE for bindestrek-normalisering («wifi» treffer «Wi-Fi»)
- Søk KB: input B4 + kategorifilter B5; Søk saker: input B4 + status-avkryssingsbokser
- Krever at formler regnes om (F9/ved åpning) — ikke sanntid type-ahead

## Nåværende dashboard (statisk)
- 6 KPI-er (rad 4–5); 3 diagrammer (kake A8, kategori F8, ansvarlig F26)
- Faste størrelser (CHART_H/CHART_W) og faste hjelpetabell-områder (A26:D42)
- Ingen dynamisk skalering/reposisjonering ved voksende lister

## Forskningsfunn (fra Oracle + librarian)

### Excel VBA-injeksjon
- `.xlsm` = ZIP med `xl/vbaProject.bin`; krever patching av `[Content_Types].xml` (Override + macroEnabled main+xml) og `xl/_rels/workbook.xml.rels` (vbaProject-relasjon)
- **pyOpenVBA** (v1.0.0, 2026): ren-Python-bibliotek som lager `.xlsm` FRA SCRATCH (`ExcelFile.create_new()` + `add_module()`). Ingen Excel nødvendig.
- xlsxwriter har `add_vba_project()`; openpyxl bevarer kun VBA på round-trip (`keep_vba=True`)
- **Anbefalt (siden Excel finnes)**: author VBA i Excel → ekstraher `vbaProject.bin` → sjekk inn → generator injiserer. Ulempe: binær blob i repo + manuelt steg ved makro-endring.

### LO/OO Basic i .ods
- Basic lagres som tekst-XML: `Basic/script-lc.xml`, `Basic/Standard/script-lb.xml`, `Basic/Standard/Module1.xml`. Kan genereres fra Python og zip-injiseres.
- **openpyxl kan ikke skrive .ods** → LO-track: skjelett som `.xlsx` → `soffice --headless --convert-to ods` → injisere Basic-XML. Ny byggavhengighet: `soffice`.

### Event-drevet søk + dynamisk dashboard
- **Excel**: `Worksheet_Change` + `Application.EnableEvents`-guard + `Application.OnTime`-debounce. Diagrammer via `ChartObjects(i).Left/Top/Width/Height`. Trivielt.
- **LO/OO**: ingen `Worksheet_Change`-ekvivalent — `addModifyListener` (XModifyListener) + module-level-variabler (hindre GC). Ingen persistent event-binding uten GUI. Debounce = `Wait`-loop (blokkerer). Diagrammer via DrawPage shapes (UNO-komplekst). **LO sanntid type-ahead er skjørere** — vurder knapp-triggeret søk / AutoOpen-listener som fallback.
- **Performance**: Excel VBA ~50ms/500 rader; LO Basic ~200ms+ (bruk `getDataArray()`).
- **Makrosikkerhet**: makroer deaktivert som standard i begge → krever Trust Center (Excel) / Macro Security-nivå (LO).

## Arkitektur-anbefaling (fra Oracle)
- **Felles kjerne + to tynne target-wrapper** — ~90 % av koden deles (ark, stiler, DV, formler, KPI-er, statiske diagrammer); kun makrolag + pakking divergerer.
- **Frys en "celle-kontrakt"** (MAX_S/MAX_K, headers, KPI-celler, hjelpetabell-områder, diagram-ankre) tidlig.
- **Makroer erstatter de to skjøreste mekanismene**: sirkulær referanse-tidsstempel + skjulte hjelpekolonner (Saker!Y, Kunnskapsbase!L).
- **Behold makrofri `.xlsx` som fallback** (nesten gratis, sikrer mot blokkerte makroer).

## Risiko / fallgruver
1. **Makroer deaktivert som standard** → fallback nødvendig.
2. **`vbaProject.bin` sheet-codename-kobling** — donor må matche openpyxl-generert ark-struktur eksakt.
3. **Dynamisk plassering vs diagram-ankring** — begrens «dynamisk» til å bygge om dataserier, ikke flytte diagramobjekter vilkårlig.
4. **backup.py hardkoder `.xlsx`** (linje 76 + 101) → må generaliseres for `.xlsm`/`.ods`.
5. **Migrering** av eksisterende data: kopier kun statiske kolonner; med makroer blir B (tidsstempel) statisk verdi.

## Åpne spørsmål (avklares med bruker)
- Fallback makrofri versjon: beholde eller droppe?
- VBA-authoring: manuell i Excel (anbefalt) vs pyOpenVBA ren-Python?
- Migrering av eksisterende data: trengs eller fresh start?
- Teststrategi (prosjektet har ingen test-suite i dag)
