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
