# Notatblokk: supportsystem-macro-split

## Konvensjoner (MÅ følges i ALLE implementasjoner)
- **Språk**: koden/kommentarer/tekst er **norsk**, men Excel-formel-funksjonsnavn forblir **engelske** med komma-separator (openpyxl-krav).
- **openpyxl-begrensninger**: ingen `ws.add_table` (LO/OO-inkompatibelt), ingen M365-funksjoner (`FILTER`, `HSTACK`, `XLOOKUP`), ingen `ManualLayout` på diagrammer.
- **Sirkulær referanse i Saker!B** er en hard AGENTS.md-begrensning — behold i fallback, erstatt KUN i target-branch for excel/calc.
- **KPI/COUNTIFS må IKKE teste kolonne B** (formel-kolonne) — fallback avhenger av denne regelen.
- **Excel-formler skrives engelsk/komma**; norske UI-tekster beholdes.
- **Makrofri `.xlsx` er byte-stabil regresjons-gate** — delt kjerne må ikke endre fallback-utdata.
- **`codeName` settes eksplisitt** i generatoren per `kontrakt.py` fordi VBA-donor matcher ark via codenames.

## Arkitekturavgjørelser (låst)
- Én generator (`lag_supportsystem.py`) med `--target excel|calc`; default target er **fallback .xlsx** (bakoverkompatibelt).
- Felles kjerne (`build_common.py`) returnerer `Workbook`; profiler (`fallback`/`excel`/`calc`) styrer divergens (tidsstempel-mekanisme, søk-mekanisme).
- Excel: VBA source i `.bas` → bygg donor via Excel-COM/pywin32 → ekstraher `vbaProject.bin` → injiser i `.xlsm`.
- LO/OO: Basic source som tekst-XML → generer `.xlsx`-skjelett → `soffice --headless --convert-to ods` → injiser Basic-XML → `.ods`.
- Søk: tittel=4, nøkkelord=3, problem=2, løsning=1; fuzzy = Levenshtein ≤ 2; trefford-markering = helcelle-formatering.
- Dashboard-dynamikk: rebuild dataserier + resize faste diagrammer; IKKE vilkårlig reposisjonering.

## Risikoårvåkenhet
- Makroer deaktivert som standard — README må forklare aktivering.
- `vbaProject.bin` har sheet-codename-kobling; `codeName` må fryses.
- LO/OO sanntid type-ahead er skjør — bruk knapp-triggeret søk + best-effort listener.
- `soffice` må være på PATH for LO-tracken.
- `backup.py` må bevare kildefilens endelse (ikke hardkod `.xlsx`).

## Gotchas
- openpyxl kan **ikke** skrive `.ods` eller makroer.
- VBA og Basic er inkompatible — to separate makroimplementasjoner.
- LO Basic er treg (~200 ms/500 rader) — bruk `getDataArray()` i stedet for celle-for-celle.


## T1 — Miljøvalidering + baseline (2026-09-07)
- Python 3.11.9, openpyxl 3.1.5 — OK. Ingen pywin32 installert (Excel VBA-donor-sporet T3/T6/T7 krever \pip install pywin32\).
- Excel ER installert (C2R x64, Office16, EXCEL.EXE funnet) — COM vil fungere når pywin32 er på plass.
- **soffice mangler helt** (ikke på PATH, ikke på standardstier) — BLOCKER for Wave 4 (LO/OO-sporet). Må installeres/legges på PATH før T4/T11/T15.
- Generator verifisert: \python lag_supportsystem.py\ og \... ren\ gir EXIT 0; endelig tilstand etter T1 = eksempeldata-variant.
- Baseline SHA-256 (\Supportsystem.xlsx\, eksempeldata): E622CC538C86229AA2E218F9976C75D33EC5DACAE6F76C63C799B343D980853B
- Zip har 31 oppføringer inkl. chart1-3.xml, drawings, vmlDrawing1.vml (form controls) og 6 ctrlProps — VBA-donor/komponent-fangst må dekke disse.
- Evidens: .sisyphus/evidence/baseline/{miljoe.txt,sha256.txt,filelist.txt}, .sisyphus/evidence/task-1-{miljoe,baseline}.txt

## T3 — Spike: VBA-injeksjon (proof of concept) (2026-09-07)
- **Status: OK**. Injeksjonen er bevist: donor `.xlsm` via Excel COM/pywin32, ekstraksjon av `xl/vbaProject.bin`, post-save injeksjon i openpyxl `.xlsx`, fil åpnes i Excel uten reparasjon og `Sub SayHello()` er tilstede.
- **Miljøkrav**: pywin32 installert; Excel COM krever `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1` for programmatisk VBProject-tilgang.
- **Injeksjonsoppskrift**:
  1. Injiser `vbaProject.bin` i zip-stien `xl/vbaProject.bin`.
  2. Patch `xl/_rels/workbook.xml.rels` med `<Relationship Id="rIdN" Type="http://schemas.microsoft.com/office/2006/relationships/vbaProject" Target="vbaProject.bin"/>`.
  3. Patch `[Content_Types].xml`:
     - endre `/xl/workbook.xml` til `application/vnd.ms-excel.sheet.macroEnabled.main+xml`;
     - legg til `<Override PartName="/xl/vbaProject.bin" ContentType="application/vnd.ms-office.vbaProject"/>`.
  4. Lagre som `.xlsm`; bruk `openpyxl.load_workbook(..., keep_vba=True)` for videre runder.
- **Verifisering**: `wb.vba_archive.namelist()` inneholder `xl/vbaProject.bin`; Excel COM lister `VBComponents` og finner `SayHello`.
- **Arv til T6/T7/T8**: VBA-donor må matche ark-codenavn (`codeName` fra `kontrakt.py`) fordi `vbaProject.bin` har hardkodet sheet-codename-mapping.
- **Evidens**: `.sisyphus/evidence/task-3-vba-spike.txt`, `.sisyphus/evidence/spike/vba/{lag_donor.py,injiser_vba.py,README.md,spike.xlsm,vbaProject.bin,donor.xlsm}`.

## T2 — Trekk ut felles kjerne build_common.py (2026-09-07)
- **Status: OK (etter fix)**. `build_common.py` (925 linjer) med `build_skeleton(profile)` + `injiser_avkryssingsbokser`; `lag_supportsystem.py` = tynt skall (65 linjer).
- **KRITISK LÆRING — byte-stabilitet er UMULIG for openpyxl**: SHA-256 av `.xlsx` endres hver kjøring fordi openpyxl skriver `dcterms:created`/`modified` i `docProps/core.xml` + zip-oppførings-tidsstempler. **Korrekt regresjonsgate = innholds-stabil**: sammenlign alle zip-oppføringer byte-for-byte, men normaliser tidsstempelfelt i `docProps/core.xml` og `docProps/app.xml` (+ ignorer zip-oppførings-tidsstempler).
- **Gotcha — VML-whitespace**: refaktorering av avkryssingsboks-f-strengen introduserte 1 mellomrom ekstra på `<x:ClientData`-linjen (3 vs 2 mellomrom). Må bevare eksakt whitespace i VML-malen.
- **Gotcha — formel-anførselstegn**: `IF($P{r}<>"",...)` ble under refaktorering ødelagt til `IF($P{r}<>",...)` (mistet ett anførselstegn). Verifiser alltid `<>""` (dobbelt anførselstegn) i formel-strenger.
- **Innholds-paritet VERIFISERT**: 0 ulikheter (etter tidsstempel-normalisering) mot original. `ren`-modus også verifisert.
- T1s SHA-256-baseline (`E622CC53...`) er en byte-hash som IKKE kan reproduseres — erstattes av innholds-sammenligning (normalisert). Referanse for innhold = git HEAD `Supportsystem.xlsx` (115965 bytes).

## T5 — Frys celle-kontrakt `kontrakt.py` (2026-09-07)
- **Status: OK**. Ny fil `kontrakt.py` er eneste kilde til sannhet for: `MAX_S`/`MAX_K`/`N_KB`/`N_SAK`, ark-rekkefølge + `codeName`-mapping, Saker/KB-overskrifter, KPI-celler + formler, dashboard-hjelpetabeller, diagram-ankre, hjelpe-/speil-kolonner, søkeinput og resultatrader; pluss scoring-orakel (`normalize`, `levenshtein`, `score`) og `FALLBACK_PROFILE`/`EXCEL_PROFILE`/`CALC_PROFILE`.
- `build_common.py` importerer nå konstanter fra `kontrakt.py` og setter `ws.sheet_properties.codeName` **betinget** (`if profile.get("set_codename")`). Fallback får dermed IKKE `codeName`, i tråd med opprinnelig openpyxl-utdata.
- `lag_supportsystem.py` bruker `dict(kontrakt.FALLBACK_PROFILE)` og fyller på `med_eksempler`; ingen `--target`-switch ennå (T8/T11).
- **Innholdsstabilitet VERIFISERT**: 0 ulikheter mot git HEAD `Supportsystem.xlsx` etter normalisering av `docProps/core.xml`/`docProps/app.xml` + `TotalTime`.
- **Orakel-resultater**: `score("wifi", doc) = 3` (delstreng i nøkkelord), `score("wify", doc) = 7` (fuzzy-treff i tittel + nøkkelord). Fuzzy-fallback aktiverer korrekt for skrivefeilen.
- Evidens: `.sisyphus/evidence/task-5-orakel.txt`, `.sisyphus/evidence/task-5-fuzzy.txt`, `.sisyphus/evidence/task-5-stabil.txt`.

## KORREKSJON — normalize må FJERNE bindestreker (IKKE erstatte med mellomrom) (2026-09-07)
- **BUG funnet under verifisering av T5**: `normalize` brukte `.replace("-", " ")` (bindestrek → mellomrom), men originalen bruker `SUBSTITUTE(x,"-","")` (bindestrek FJERNES). Konsekvens: «wifi» treffer «Wi-Fi» kun via fuzzy (dist 2), ikke som EKSAKT delstreng.
- **RETTET** i `kontrakt.py`: `.replace("-", "")`. Nå er `normalize("Wi-Fi nede") == "wifi nede"` og `score("wifi", {"title":"Wi-Fi nede"}) == 4` via eksakt treff.
- **OBS for T6 (VBA) og T10 (Basic)**: `Normaliser()`/`Normaliser` MÅ fjerne bindestreker (f.eks. `Replace(s, "-", "")` i VBA, `Replace(s, "-", "")` i Basic), IKKE erstatte med mellomrom. Match mot `kontrakt.normalize`-oppførsel eksakt.
- Fuzzy-fallback er fortsatt korrekt (Levenshtein ≤ 2), men skal kun trigge når det IKKE finnes eksakt delstreng-treff.

## KORREKSJON 2 — VBA-trefffarge er feil (lyseblå, ikke FFEB9C) (2026-09-07)
- `vba/ModSok.bas` har `Const FARG_TREFF As Long = 15652731` med kommentar "RGB(255,235,156)". Men 15652731 = 0xEED77B = BGR(238,215,123) = RGB(123,215,238) = **lyseblå**, IKKE FFEB9C (lysegul).
- Korrekt VBA-konstant for FFEB9C: `RGB(255,235,156)` = 10284031 (VBA RGB = r + g*256 + b*65536 = 0xBBGGRR).
- `basic/ModSok.bas` bruker korrekt `16771996` (StarBasic RGB = 0xRRGGBB, direkte UNO-format; 16771996 = 0xFFEB9C... NB: StarBasic RGB(255,235,156) = 255*65536+235*256+156 = 16771996).
- **ÅPEN OPPGAVE (mindre, kosmetisk)**: fiks `vba/ModSok.bas` sin `FARG_TREFF` til 10284031 og gjenbygg donor + `vbaProject.bin` (kjør `python vba/build_donor.py`). Kan gjøres i F2/F3-review eller som oppfølging.


## T9: Excel-makroveiledning i README (2026-09-07)
- **Status: OK**. README.md oppdatert med ny seksjon, kun additivt. Eksisterende fallback-/backup-/feilsøkingsdokumentasjon er ikke rørt.
- **Nye deler**: intro-avsnitt om filvarianter, Krav-bullet for `.xlsm` (Excel 2007+, makroer må aktiveres/klareres), `--target excel`/`--target calc`-kommandoer under «Generere filen på nytt», ny `## Hvilken fil skal dere bruke?`-seksjon (tre-fil-modell, tabell xlsx/xlsm/ods) med underseksjonene `### Makrofri .xlsx (fallback)`, `### Excel (.xlsm) med makroer` (sanntidssøk med rangering/fuzzy/markering, dynamisk dashboard, makrobasert tidsstempel, ingen iterativ beregning) og `#### Slik aktiverer du makroer i Excel` (Opphev blokkering, klarerte plasseringer, Trust Center-sti). `.ods`-varianten beskrives eksplisitt som «kommende», ikke implementert. Ny feilsøkings-rad for makroer som ikke kjører.
- **Lærdom**: dokumentasjonen er skrevet før `--target`-switchen er implementert (T8/T11). Når den lander må README sjekkes mot faktisk CLI (f.eks. om `ren` kombineres som `--target excel ren`). Unngå en/em-bindestrek i ny tekst.
- **Evidens**: `.sisyphus/evidence/task-9-readme.txt`

## T7 — Bygg donor `.xlsm` + ekstraher `vbaProject.bin` (2026-09-07)
- **Status: OK**. Donor bygget, codeNames verifisert, `vba/vbaProject.bin` ekstrahert (82 944 bytes).
- **Fil**: `vba/build_donor.py` — gjenbrukbart bygge-skript via Excel COM/pywin32.
  - Setter `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1` programmatisk.
  - Lager 7 ark i kontrakt-rekkefølge; verifiserer codeNames `Sheet1`..`Sheet7`.
  - Importerer `ModFelles.bas`, `ModSok.bas`, `ModDashboard.bas`, `ModTidsstempel.bas`.
  - Legger hendelsesbehandlere i `Sheet2`, `Sheet3`, `Sheet4` og `ThisWorkbook`.
  - Lagrer `vba/donor.xlsm`, ekstraherer `xl/vbaProject.bin`.
  - Verifiserer med både openpyxl (`keep_vba=True`) og Excel COM.
- **codeName-map (final)**:
  - Dashboard -> Sheet1
  - Saker -> Sheet2
  - Søk KB -> Sheet3
  - Søk saker -> Sheet4
  - Kunnskapsbase -> Sheet5
  - Oppsett -> Sheet6
  - Instruksjoner -> Sheet7
- **Kompilerings-/encoding-status**: `.bas`-filene er Windows-1252/Latin-1 (ikke UTF-8) og importeres OK til Excel. Ingen compile-feil observert under bygg/verifisering.
- **Donor `.xlsm`**: forblir på disk som byggeledd, men er lagt til `.gitignore` og skal ikke committes.
- **Dokumentasjon**: `vba/README.md` beskriver gjenoppbyggingsprosedyre.
- **Evidens**: `.sisyphus/evidence/task-7-donor.txt`

## T8 — Implementer `--target excel` (injiser VBA → `.xlsm`) (2026-09-07)
- **Status: OK**. `python lag_supportsystem.py --target excel` produserer `Supportsystem.xlsm` med `vba/vbaProject.bin`, `macroEnabled` content type og codeNames `Sheet1`..`Sheet7`.
- **Divergens implementert** i `build_common.py`:
  - `timestamp_mechanism == "macro"`: `fullCalcOnLoad=True`, ingen iterativ beregning; `Saker!B` forblir tom (sample-data fyller statiske tidsstempler).
  - `search_mechanism == "macro"`: `Saker!Y`, `Kunnskapsbase!L`, `Søk KB`-resultatrader og `Søk saker`-resultatrader forblir tomme; statiske prompt i `A7`/`A6`. Status-speilet `AA5:AB14` beholdes for avkryssingsboksene.
- **Ny VBA-injeksjon**: `build_common.injiser_vba_prosjekt(filnavn, bin_path)` følger T3-oppskriften: injiserer `xl/vbaProject.bin`, patcher `[Content_Types].xml` (macroEnabled + vbaProject-override) og `xl/_rels/workbook.xml.rels`.
- **CLI**: `lag_supportsystem.py` støtter `--target excel`, `--target calc` (NotImplementedError), og `ren` kombinert med target. Fallback uten `--target` er uendret.
- **Verifisering**:
  - openpyxl: `vbaProject.bin` tilstede, macroEnabled, codeNames korrekte, ingen formel i `Saker!B`/`Y`/`L`, ingen INDEX/MATCH i søkeresultatrader.
  - Excel COM: åpner uten reparasjon; `ModFelles`, `ModSok`, `ModDashboard`, `ModTidsstempel` tilstede.
  - Fallback-innholdsstabilitet: 0 ulikheter mot git HEAD `Supportsystem.xlsx` etter tidsstempel-normalisering.
- **Evidens**: `.sisyphus/evidence/task-8-xlsm.txt`, `.sisyphus/evidence/task-8-profil.txt`

## T10 — Forfatt Basic-makroer (tekst) (2026-09-07)
- **Status: OK**. Fire StarBasic-kildemoduler + README forfattet headless (ingen kompilering — soffice mangler, skjer i T12/T15).
- **Filer**: `basic/ModFelles.bas` (Normaliser, Levenshtein, ScoreDok), `basic/ModTidsstempel.bas` (StemplTidsstempel), `basic/ModSok.bas` (SokKB, SokSaker + hjelpere: Type Treff, SomTekst, DatoSomTall, ErBedre, SorterTreff, ErTillatt, MarkerTreff, SkrivCelle, KopierDatoformat), `basic/ModDashboard.bas` (OppdaterDashboard, TellSaker), `basic/README.md`.
- **Nøkkelvalg (StarBasic/UNO)**:
  - `ThisComponent`; `Sheets.getByName`; `getCellByPosition(col, rad)` og `getDataArray()` er **0-basert** — alle KOL-konstanter er 0-baserte.
  - Bulk-lesing via `getCellRangeByName("A2:W501").getDataArray()`; skriving via `setDataArray` (sammenhengende) / celle-for-celle (spredte kolonner A,B,F,H,U,V).
  - **Normaliser FJERNER bindestrek** (`Replace(t,"-","")`) — KORREKSJON fra T5 er ivaretatt.
  - **MarkerTreff**: `CellBackColor = 16771996` (=RGB(255,235,156)=FFEB9C) + `CharWeight = 150`. StarBasic `RGB` = 0xRRGGBB = UNO-format direkte (motsatt av VBA). VBA-konstanten 15652731 i `vba/ModSok.bas` samsvarer IKKE med kommentaren FFEB9C (gir lys blå) — Basic bruker korrekt FFEB9C.
  - **Knapp-triggeret** (LO har ingen Worksheet_Change); best-effort listener dokumentert i kommentar i ModSok.bas (utsatt til T11/T12).
  - **Diagrammer**: levende Calc-diagrammer leser hjelpetabellene direkte → OppdaterDashboard skriver kun tall, ingen UNO-diagram-API; diagramobjekter flyttes aldri.
  - `Option Explicit` bevisst utelatt (headless); T12/T15 kan slå det på etter kompilering.
- **Scoring-kryssjekk**: Python-motpart (speiler Basic-løkkene) == `kontrakt.score` i 9/9 tilfeller. `score("wifi",{"title":"Wi-Fi nede"})=4` (eksakt), `score("wify",{"title":"Wi-Fi nede"})=4` (fuzzy); `Normaliser("Wi-Fi nede")="wifi nede"`.
- **Datoformat**: B-kolonnen i Saker har allerede "DD.MM.YYYY HH:MM" fra generatoren (build_common linje 249–250, alle profiler) → StemplTidsstempel skriver kun `Now()`. Søkeresultat-datoer (Søk KB H, Søk saker B) har IKKE format i makroversjonen → kopieres fra referanseceller (KB!H2 / Saker!B2) via KopierDatoformat.
- **Evidens**: `.sisyphus/evidence/task-10-basic.txt`, `.sisyphus/evidence/task-10-score.txt`, `.sisyphus/evidence/crosscheck_score.py`.

## T13 — Generaliser backup.py for .xlsm/.ods/.xlsx (2026-09-07)
- **Status: OK**. `backup.py` hardkodet ikke lenger `.xlsx`; backupen bevarer kildefilens egen endelse.
- **Endringer**:
  - `backup.py`: docstring nå "Backup-rutine for Supportsystem-filen (.xlsx/.xlsm/.ods)".
  - `backup.py` STANDARD_CONFIG: kommentar over `kildefil` om at den kan peke på `.xlsm`/`.ods`.
  - `kjor_backup`: `ekst = os.path.splitext(os.path.basename(kilde))[1]`; `backupnavn = f"{navn_grunn}_{stempel}{ekst}"`; sender `ekst` til `roter(...)`.
  - `roter(mappe, navn_grunn, behold, cfg, ekst)`: filtrerer `f.startswith(navn_grunn + "_")` og `os.path.splitext(f)[1].lower() == ekst.lower()` — kun kopier med samme endelse roteres.
  - `backup_service.py`: `_svc_description_` er nå endelsesnøytral ("Supportsystem-filen (.xlsx/.xlsm/.ods)").
  - `backup_config.json`: uendret (JSON har ingen kommentarer; README dekker det fra T9).
- **Verifisert**: 3 kjeringer med `kildefil=Supportsystem.xlsm` og `behold_antall=2` → `.xlsm`-backuper opprettet, rotasjon beholdt 2 nyeste, eksisterende `.xlsx`-backuper urørt (6 før/etter). Dummy `.ods`-test med `behold_antall=1` → riktig endelse + rotasjon. Config gjenopprettet identisk; testfiler ryddet.
- **Lærdom**: rotasjon per endelse betyr at blandede `.xlsx`/`.xlsm`/`.ods`-backuper i samme mappe ikke forstyrrer hverandre — behold_antall gjelder per kilde-endelse, ikke per mappe totalt.
- **Evidens**: `.sisyphus/evidence/task-13-backup.txt`, `.sisyphus/evidence/task-13-rotasjon.txt`

## T14 — Strukturell test-suite `test_struktur.py` (2026-09-07)
- **Status: OK**. `test_struktur.py` implementert med 17 tester: 16 PASS, 1 SKIP (.ods-placeholder).
- **Kjøring**: `python -m pytest test_struktur.py -v` gir `16 passed, 1 skipped`.
- **Fallback `.xlsx`**: testet ark-rekkefølge mot `kontrakt.SHEETS`, sirkulær formel i `Saker!B2`, hjelpekolonner `Y`/`L`, KPI-formler, definerte navn, datavalideringer, 3 diagrammer, zip-presens av `xl/ctrlProps/` + `xl/drawings/vmlDrawing1.vml`, og at `codeName` er `None`.
- **Excel `.xlsm`**: testet `vbaProject.bin`, `macroEnabled`+`vbaProject` i `[Content_Types].xml`, `codeName` mot `kontrakt.CODE_NAMES`, at `Saker!B2` ikke er formel, at hjelpe-/søkekolonner er tomme, og at status-speilet `AA5`/`AB5` er formel/boolsk.
- **Merking av KPI-celler**: malen legger faktisk KPI-formlene én rad under label-cellen (f.eks. `A5` for label `A4`), ikke i `B5`/`D5`/... som planen antok. Testene validerer de faktiske formelcellene via `kontrakt.KPIS`.
- **`.ods` placeholder**: `@pytest.mark.skip(reason="T11 utsatt: LibreOffice/soffice ikke installert")` dokumenterer at T11 skal sjekke `basic/Standard/Module1.xml`, `script-lc.xml` og `script-lb.xml` inne i `Supportsystem.ods`.
- **Evidens**: `.sisyphus/evidence/task-14-tester.txt`

## T4 — Spike: Basic-injeksjon (proof of concept) (2026-09-07)
- **Status: OK**. Basic-injeksjon i `.ods` er bevist via zip-rewrite etter `soffice --headless --convert-to ods`.
- **soffice-sti**: `C:\Program Files\LibreOffice\program\soffice.exe` (ikke på PATH).
- **Injeksjonsoppskrift**:
  1. Generer `.xlsx`-skjelett med openpyxl.
  2. Konverter til `.ods` med `soffice --headless --convert-to ods`.
  3. Zip-rewrite: legg til `Basic/script-lc.xml`, `Basic/Standard/script-lb.xml`, `Basic/Standard/Module1.xml`.
  4. Patch `META-INF/manifest.xml` med `manifest:file-entry` for `Basic/`, `Basic/script-lc.xml`, `Basic/Standard/`, `Basic/Standard/script-lb.xml`, `Basic/Standard/Module1.xml`.
  5. **Manifest-patching er påkrevd**: uten det feiler headless-konvertering (tomt output-dir).
- **Verifisering**: `.ods` → `.xlsx` med soffice produserer gyldig fil; openpyxl åpner den og ser `['Demo']` + data. Makroen overlever også `.ods` → `.ods` round-trip.
- **Gjenbrukbart skript**: `.sisyphus/evidence/spike/basic/injiser_basic.py` støtter `--bas`, `--modul`, `--bib`, `--verify`, `--soffice`/`SOFFICE`.
- **Dokumentasjon**: `.sisyphus/evidence/spike/basic/README.md`.
- **Blokkerere**: ingen. T11 er ikke lenger blokkert av soffice/manglende LO-spike.
- **Evidens**: `.sisyphus/evidence/task-4-basic-spike.txt`.

## T11 — Implementer `--target calc` (LibreOffice/OpenOffice .ods) (2026-09-07)
- **Status: OK**. `python lag_supportsystem.py --target calc` produserer `Supportsystem.ods` med Basic-makroer.
- **Pipeline** (i `lag_supportsystem.py`):
  1. Bygg skjelett med `build_common.build_skeleton(dict(kontrakt.CALC_PROFILE) | {"med_eksempler": MED_EKSEMPLER})`.
  2. Lagre temp `.xlsx`, injiser avkryssingsbokser.
  3. Konverter til `.ods` med `soffice --headless --convert-to ods` (sti overridbar via `SOFFICE`-miljøvariabel, default `C:\Program Files\LibreOffice\program\soffice.exe`).
  4. Poll på output-fil (soffice er en launcher).
  5. Injiser Basic-kilde fra `basic/ModFelles.bas`, `basic/ModTidsstempel.bas`, `basic/ModSok.bas`, `basic/ModDashboard.bas` som `Basic/Standard/Module1.xml`..`Module4.xml` via ny `build_common.injiser_basic(ods_path, moduler)`.
  6. Manifest patches med `manifest:file-entry` for `Basic/`, `Basic/script-lc.xml`, `Basic/Standard/`, `Basic/Standard/script-lb.xml` og hver `ModuleN.xml` (påkrevd, arv fra T4).
- **Nye funksjoner i `build_common.py`**:
  - `injiser_basic(ods_path, moduler, bib_navn="Standard")`: zip-rewrite av .ods med Basic-bibliotek, moduler og manifest-patching.
  - `_les_bas(sti)`, `_xml_escape(t)`, `_basic_library_container_xml`, `_basic_library_index_xml`, `_basic_module_xml`, `_patch_ods_manifest`.
  - Encoding: prøver UTF-8, faller tilbake til cp1252; XML-escape av `& < >` før innbygging i `<script:source>`; XML-parts skrives som UTF-8.
- **Test**: `test_struktur.py::test_ods_calc_macro_parts` er un-skippet og sjekker `Module1..4.xml`, `script-lb.xml`, `script-lc.xml` + manifest-entries. I tillegg rettet `test_xlsx_saker_b2_circular_formula` til å sjekke `B12` (første tomme rad etter eksempeldata), siden `B2..B11` overskrives med statiske datoer.
- **Verifisering**:
  - `Supportsystem.ods` inneholder alle Basic-parts og manifest-entries.
  - ODS → XLSX round-trip med soffice: exit 0, 7 ark beholdes.
  - Fallback `.xlsx` innholdsstabil mot git HEAD (0 ulikheter etter normalisering av tidsstempler).
  - `python -m pytest test_struktur.py`: **17 passed**.
- **Oppdateringer**: `README.md` oppdatert til å beskrive `.ods`-varianten som implementert; fjernet "kommende"/"ikke implementert ennå".
- **Evidens**: `.sisyphus/evidence/task-11-ods.txt`.

## T12 — LO/OO README + makro-sikkerhetsveiledning (2026-09-07)
- **Status: OK**. README.md oppdatert; `.ods`-seksjonen er nå komplett og nøyaktig for makroversjonen.
- **Endringer**:
  - Intro: «planlagt versjon for LibreOffice og OpenOffice (.ods)» → «makroaktivert versjon ... der søk og tidsstempel styres fra knapper» (siste gjenværende «planlagt»-referanse fjernet).
  - «Viktig for LibreOffice / OpenOffice» (Krav): ny innledning som avgrenser iterative-referanser-punktlisten til den makrofrie `.xlsx` og peker på `.ods`-seksjonen.
  - `.ods`-seksjonen omskrevet: søk er KNAPP-triggeret («SØK I KB →» / «SØK I SAKER →»), tidsstempel er KNAPP-triggeret via `StemplTidsstempel`, dashboard via `OppdaterDashboard`; eksplisitt at makroversjonen IKKE trenger iterative referanser (fallback gjør det); statusfilter via speiltabell `AA5:AB14` (AB5:AB14 kan redigeres SANN/USANN direkte hvis boksene vises grå); ny underseksjon «Slik aktiverer du makroer i LibreOffice og OpenOffice» (Verktøy > Innstillinger > LibreOffice > Sikkerhet > Makrosikkerhet, Medium = spør; OpenOffice tilsvarende; klarerte plasseringer; verifiseringssteg).
  - Feilsøking: ny rad for `.ods` (makroer ikke aktivert / knapp ikke klikket).
- **Faktasjekk mot kode**: `StemplTidsstempel` låser eksisterende stempler (fyller kun tom B der C/H er utfylt); `ModSok` leser `AA5:AB14` (AA=statusnavn, AB=SANN/USANN); `kontrakt.SOK_SAK_STATUS`; D4-knappceller i build_common.
- **OBS (åpen tråd for T15/F-reviews)**: generert `.ods` inneholder Basic-modulene men INGEN makro-tilknyttede kontroller ennå (`<office:scripts/>` er tomt i content.xml). Knappetilordning er dokumentert i `basic/README.md` (Skjema > Trykknapp + «Utfør handling» → ModSok.SokKB osv.) og følger planens design. README beskriver tiltenkt knapp-basert bruk slik planen og arkets egne prompt-tekster angir; hvis T15/etterarbeid binder knapper automatisk inn i `.ods`, må README-verifiseringssteget (klikk «SØK I KB →») stemme med faktisk fil.
- **Evidens**: `.sisyphus/evidence/task-12-readme.txt`.

## T15 — soffice-valideringspipeline `valider_soffice.py` (2026-09-07)
- **Status: OK**. Ny fil `valider_soffice.py` validerer at genererte filer åpnes rent i LibreOffice via headless round-trip-konvertering.
- **Funksjon**: `Supportsystem.ods` → `.xlsx` og `Supportsystem.xlsx` → `.ods` med `soffice --headless --convert-to <mal> --outdir <temp>`, poller på utdatafil (soffice er en launcher), sjekker exit-kode 0 + at utdata finnes. Temp-mappe via `tempfile.mkdtemp()` ryddes i `finally` — kildefiler overskrives aldri.
- **Konfigurasjon**: `SOFFICE = os.environ.get("SOFFICE", r"C:\Program Files\LibreOffice\program\soffice.exe")` — sti overstyres med miljøvariabel, ikke hardkodet.
- **Manglende kildefil**: skriptet feiler med tydelig melding (begge filer skal finnes etter generering) — valgt fremfor å hoppe over.
- **Verifisert**: begge round-trips PASS (exit 0; ods→xlsx 113 234 bytes, xlsx→ods 125 237 bytes), EXITCODE=0. Ingen soffice-prosesser eller temp-mapper etterlatt.
- **README**: ny Feilsøking-rad dokumenterer `python valider_soffice.py` som headless åpne-validering (krever LibreOffice, `SOFFICE`-overstyring).
- **Evidens**: `.sisyphus/evidence/task-15-soffice.txt`.

## F3 — Reell manuell QA (2026-09-07)
- **Status: APPROVE**. Alle 8 scenarioer PASS: fallback .xlsx, Excel .xlsm, Calc .ods, backup-endelsesbevaring, edge-cases, kryss-integrasjon, test_struktur.py (17 passed), valider_soffice.py (PASS).
- Evidence-filer lagret i `.sisyphus/evidence/final-qa/`, inkl. `f3-manuell-qa.txt`.
- Backup-config ble midlertidig endret og gjenopprettet identisk; rotasjon fungerer per endelse.
- Norske tegn (æøå) overlever i .xlsx/.xlsm/.ods (sistnevnte verifisert via soffice-konvertering til xlsx).
- VBA (cp1252) og Basic (utf-8) kildefiler bruker begge `Replace(..., "-", "")` for bindestrek-fjerning, i tråd med `kontrakt.normalize`.
- Viktig: PowerShell `>`-redirection lager UTF-16-filer; bruk `Out-File -Encoding utf8` for QA-evidens som skal leses tilbake som tekst.



## F2 — Kodekvalitetsreview (2026-09-07)
- VERDIKT: APPROVE.
- Bygg: alle targets PASS; pytest: 17/17 PASS; valider_soffice: PASS (med transient FAIL på første parallellkjøring av ODS→XLSX, re-kjørt OK).
- Issues funnet: ubrukt import Table/TableStyleInfo i build_common.py:15; mange hardkodede 501/301/ankre i build_common.py som kunne sentraliseres fra kontrakt.py; korrupt tegn (U+FFFD) i test_struktur.py:263 ('ennå'); _BOKS_DEF-magiske tall og 'Løst'/'Lukket' hardkodet i lag_supportsystem.py.
- Kjente tråder: VBA FARG_TREFF rettet til 10284031 i ModSok.bas (kilde OK, binær avhenger av donor-rebuild); .ods mangler fortsatt automatisk knappetilknytning (åpen, dokumentert).
- AI-slop: ingen alvorlige tegn; kommentarer og abstraksjoner er forsvarlige.
- Evidence: .sisyphus/evidence/final/f2-kodekvalitet.txt

