# Splitte supportsystemet i to makro-versjoner (Excel VBA + LO/OO Basic)

## TL;DR

> **Kort oppsummering**: Del dagens makrofrie `lag_supportsystem.py` i en **felles kjerne** + to målplattformer med hvert sitt makrospråk: MS Excel (`.xlsm` med VBA) og LibreOffice/OpenOffice (`.ods` med Basic). Makroene erstatter de skjøreste mekanismene (sirkulær referanse-tidsstempel og skjulte hjelpekolonner) og gir sanntid søk med rangering/fuzzy/markering samt dynamisk dashboard. Den makrofrie `.xlsx` beholdes som fallback.
>
> **Leveranser**:
> - `build_common.py` — delt kjerne (ark, stiler, formler, DV, KPI-er, diagrammer)
> - `kontrakt.py` — frosset celle-kontrakt (konstanter + scorings-orakel + målprofiler)
> - `Supportsystem.xlsm` — Excel-versjon med VBA (søk + dashboard-dynamikk + tidsstempel)
> - `Supportsystem.ods` — LO/OO-versjon med Basic (tilsvarende)
> - `Supportsystem.xlsx` — makrofri fallback (uendret, byte-stabil)
> - Oppdatert `backup.py` (generaliserer filendelse), strukturell test-suite, README per mål
>
> **Estimert innsats**: Large (XL)
> **Parallell kjøring**: JA — 4 bølger med maks 4–5 parallelle oppgaver
> **Kritisk sti**: T2 (kjerne) → T5 (kontrakt-gate) → T6→T7→T8 (Excel-track) → T13/T14 (backup+tester)

---

## Context

### Opprinnelig forespørsel
Brukeren ønsker å skille løsningen i to versjoner — én for MS Excel og én for LibreOffice/OpenOffice — og optimalisere dem hver for seg, spesielt med makroer og avanserte teknikker for (1) søk og (2) dynamisk oppdatering av innhold og plassering på dashboardet.

### Intervju-oppsummering
**Nøkkelbeslutninger**:
- Prioritet: **Excel først, LO/OO etterpå**.
- Byggmaskin har **Excel installert**.
- Kodestruktur: **én generator** med `--target excel|calc` (felles kjerne + branch).
- Søk-optimalisering: sanntid type-ahead + relevans-rangering + fuzzy + markering av trefford (alle).
- Dashboard-dynamikk: auto-resize + auto-reposisjonering + auto-oppfriskning (alle).
- **Behold makrofri `.xlsx`** som fallback.
- VBA authoring: **i Excel → ekstraher `vbaProject.bin`** → injiser (`.bas` eksporteres for review).
- **Ingen datamigrering** (fresh start med eksempeldata).
- Teststrategi: **strukturell + soffice** (ingen Excel COM smoke-test).

**Forskningsfunn**:
- VBA (Excel `.xlsm`) og Basic (LO/OO `.ods`) er inkompatible — to separate makroimplementasjoner.
- openpyxl kan **ikke** skrive makroer og **ikke** skrive `.ods`. LO-tracken må: generere `.xlsx`-skjelett → `soffice --headless --convert-to ods` → injisere Basic-XML (`Basic/Standard/Module1.xml` + `script-lc.xml` + `script-lb.xml`).
- Excel VBA-injeksjon: legg til `xl/vbaProject.bin` + patch `[Content_Types].xml` (macroEnabled main+xml + vbaProject-override) + `xl/_rels/workbook.xml.rels`.
- Excel event-drevet søk er trivielt (`Worksheet_Change` + `Application.EnableEvents`-guard + `OnTime`-debounce). LO/OO har **ingen** ekvivalent; krever `addModifyListener` (XModifyListener) med module-level-variabler (hindre GC); debounce via `Wait`-loop blokkerer tolkeren. **LO type-ahead er skjør** → knapp-triggeret søk + best-effort-listener som design.
- Makroer deaktivert som standard i begge apper.
- `vbaProject.bin` har sheet-codename-kobling → generatoren må sette `ws.sheet_properties.codeName` eksplisitt.

### Metis-gjennomgang
**Identifiserte gap (adressert)**:
1. **De-risking manglet** → lagt inn Wave 1-spikes (VBA + Basic-injeksjon) før full bygging.
2. **Kontrakt-gate manglet** → T5 er en hard, verifisert gate mellom kjerne og makro-authoring.
3. **`codeName`-kobling** → eksplisitt satt i kontrakten/generatoren.
4. **Søk-algoritme udefinert** → definert i T5 (scoringsfunksjon + fuzzy-spesifikasjon + ren-Python-orakel).
5. **Fallback-byte-stabilitet** → T1 fanger baseline; T14 verifiserer regresjon.
6. **backup-endring for tidlig** → utsatt til T13 (etter at nye artefakter finnes).
7. **`backup.py` har 3 (ikke 2) `.xlsx`-referanser** (linje 28, 76, 101) + `backup_service.py`-beskrivelse → dekket i T13.
8. **Sirkulær referanse er en hard AGENTS.md-begrensning** → beholdes i delt kjerne for fallback; makro-versjonene erstatter den KUN i sin target-branch.
9. **Dynamisk plassering skjør** → begrenset til «bygg om dataserier + resizer faste diagramobjekter», IKKE vilkårlig flytting.

---

## Work Objectives

### Kjerne-mål
Én generator som produserer tre artefakter fra én delt kjerne: makrofri fallback `.xlsx`, makro-aktivert `.xlsm` (VBA) og makro-aktivert `.ods` (Basic) — med sanntid/rangert søk og dynamisk dashboard i de to makro-versjonene.

### Konkrete leveranser
- `build_common.py` (delt kjerne), `kontrakt.py` (celle-kontrakt + scorings-orakel), `.bas`-kilder, `vbaProject.bin`-artefakt, Basic-modul-XML.
- `Supportsystem.xlsx` (fallback), `Supportsystem.xlsm` (Excel), `Supportsystem.ods` (LO/OO).
- `test_struktur.py` (strukturell test-suite), oppdatert `backup.py`/`backup_config.json`, README per mål.

### Definition of Done
- [ ] `python lag_supportsystem.py` → `.xlsx` åpner i Excel/LO/OO uten reparasjon; **innholds-stabil** mot original (identiske zip-oppføringer, ignorer docProps-tidsstempler — byte-identisk er umulig pga. openpyxl-tidsstempler).
- [ ] `python lag_supportsystem.py --target excel` → `.xlsm` har `vbaProject.bin` + makroene kjører (spike-verifisert).
- [ ] `python lag_supportsystem.py --target calc` → `.ods` har Basic-moduler + åpner i LO uten reparasjon.
- [ ] `python test_struktur.py` → alle asserts passerer.
- [ ] `python backup.py` med `.xlsm`/`.ods`-kildefil → backup + rotasjon virker.

### Must Have
- Felles kjerne som er **innholds-stabil** for fallback (identiske zip-oppføringer, ignorer tidsstempler).
- Søk med rangering, fuzzy og trefford-markering (helcelle-formatering) i begge makro-versjoner.
- Dashboard som bygger om dataserier og resizer diagrammer ved voksende lister.
- README med makro-sikkerhetsveiledning for begge mål.

### Must NOT Have (Guardrails)
- **IKKE** fjern sirkulær referanse eller KPI/COUNTIFS «ikke test kolonne B»-begrensningen fra **delt kjerne** (fallback avhenger av dem).
- **IKKE** implementer vilkårlig flytting av diagramobjekter (kun ombygging av dataserier + resize).
- **IKKE** la LO-tracken blokkere ferdigstillelse av Excel-tracken.
- **IKKE** legg til nye ark, KPI-er eller data-funksjoner utover det som trengs for søk/dashboard.
- **IKKE** bruk M365-funksjoner (`FILTER`, `HSTACK`, `XLOOKUP`), Excel-tabeller eller `ManualLayout` — eksisterende AGENTS.md-begrensninger gjelder fortsatt.

---

## Verification Strategy (MANDATORY)

> **NULL MANUELL INTERVENSJON** — all verifisering er agent-utført.

### Testbeslutning
- **Infrastruktur finnes**: NEI (ingen test-suite i dag) — settes opp i T14.
- **Automatiserte tester**: Ja (strukturell, tests-etter-implementering).
- **Ramme**: `pytest` (eller ren Python `assert`) for openpyxl read-back; `soffice --headless` for `.ods`-validitet.
- **Ingen TDD** (makro-kjøring krever app, håndteres av spike + strukturelle asserts + ren-Python-orakel).

### QA-policy
Hver oppgave har agent-utførte QA-scenarioer med konkret verktøy, steg, asserts og bevissti. Evidence til `.sisyphus/evidence/task-{N}-{slug}.{ext}`.

- **Struktur/xlsx**: Python + openpyxl read-back (assert ark, formler, navn, DV, diagrammer, codeName).
- **`.xlsm`-makro**: openpyxl `keep_vba=True` (assert `vbaProject.bin`) + spike-verifisert kjøring.
- **`.ods`-makro**: unzip (assert Basic-XML-deler) + `soffice --headless` åpningsvalidering.
- **Søkelogikk**: ren-Python scorings-orakel i `kontrakt.py` (assert rangering/fuzzy for kjente spørringer).
- **Backup**: `python backup.py` + assert backup-fil og rotasjon.

---

## Execution Strategy

### Parallelle bølger

```
Wave 1 (start umiddelbart — grunnmur + de-risking, MAX PARALLELL):
├── T1: Miljøvalidering + baseline [quick]
├── T2: Trekk ut felles kjerne build_common.py [unspecified-high]
├── T3: Spike VBA-injeksjon [unspecified-high]
└── T4: Spike Basic-injeksjon [unspecified-high]

Wave 2 (etter T2 — HARD kontrakt-gate):
└── T5: Frys celle-kontrakt kontrakt.py + codeName + scorings-orakel [unspecified-high]

Wave 3 (etter T5 + T3 — Excel-track):
├── T6: Forfatt VBA (.bas) [ultrabrain]
├── T7: Bygg donor .xlsm + ekstraher vbaProject.bin [unspecified-high]
├── T8: Implementer --target excel (injiser VBA) [unspecified-high]
└── T9: Excel README + makro-sikkerhet [writing]

Wave 4 (etter T5 + T4 — LO/OO-track):
├── T10: Forfatt Basic-makroer (tekst) [deep]
├── T11: Implementer --target calc (soffice + injiser Basic) [unspecified-high]
└── T12: LO/OO README + makro-sikkerhet [writing]

Wave 5 (etter T8 + T11 — backup + verifisering):
├── T13: Generaliser backup.py [unspecified-low]
├── T14: Strukturell test-suite [unspecified-high]
└── T15: soffice valideringspipeline [unspecified-low]

Wave FINAL (etter ALLE — 4 parallelle reviews, deretter bruker-ok):
├── F1: Plan-overholdelse (oracle)
├── F2: Kodekvalitet (unspecified-high)
├── F3: Reell manuell QA (unspecified-high)
└── F4: Scope-fidelitet (deep)
-> Presenter resultater -> eksplisitt bruker-ok

Kritisk sti: T2 → T5 → T6 → T7 → T8 → T13/T14 → F1–F4
Parallell speedup: ~55 % raskere enn sekvensiell
Maks samtidig: 4 (Wave 1)
```

### Avhengighetsmatrise (full)

| Oppgave | Blokkeres av | Blokkerer |
|---|---|---|
| T1 | — | (ingen; baseline) |
| T2 | — | T5 |
| T3 | — | T7 |
| T4 | — | T11 |
| T5 | T2 | T6, T9, T10, T12 |
| T6 | T5 | T7 |
| T7 | T6, T3 | T8 |
| T8 | T7 | T13, T14 |
| T9 | T5 | (ingen) |
| T10 | T5 | T11 |
| T11 | T10, T4 | T13, T14, T15 |
| T12 | T5 | (ingen) |
| T13 | T8, T11 | (ingen) |
| T14 | T8, T11 | (ingen) |
| T15 | T11 | (ingen) |
| F1–F4 | T1–T15 | (slutt) |

### Agent-dispatch-oppsummering

- **Wave 1**: 4 oppgaver — T1 `quick`, T2/T3/T4 `unspecified-high`
- **Wave 2**: 1 oppgave — T5 `unspecified-high`
- **Wave 3**: 4 oppgaver — T6 `ultrabrain`, T7/T8 `unspecified-high`, T9 `writing`
- **Wave 4**: 3 oppgaver — T10 `deep`, T11 `unspecified-high`, T12 `writing`
- **Wave 5**: 3 oppgaver — T13/T15 `unspecified-low`, T14 `unspecified-high`
- **FINAL**: 4 reviews — F1 `oracle`, F2/F3 `unspecified-high`, F4 `deep`

---

## TODOs

- [x] 1. Miljøvalidering + baseline-fangst

  **What to do**:
  - Verifiser miljøet: `soffice --version` (må finnes for LO-tracken), `python --version` (3.9+), `pip show openpyxl`, og Excel-COM-tilgjengelighet (`pip show pywin32` + `python -c "import win32com.client; win32com.client.Dispatch('Excel.Application')"` — må lykkes for VBA-authoring).
  - Kjør `python lag_supportsystem.py` og `python lag_supportsystem.py ren` for å sikre at dagens generator fortsatt virker.
  - Fang baseline: lagre SHA-256 av `Supportsystem.xlsx` (og en unzip-filoversikt) i `.sisyphus/evidence/baseline/`. Dette er regresjons-gaten for fallback.
  - Dokumentér mangler som evt. avdekkes (f.eks. manglende soffice) i `.sisyphus/evidence/baseline/miljoe.txt`.

  **Must NOT do**:
  - Ikke endre `lag_supportsystem.py` i denne oppgaven (kun observasjon + baseline).
  - Ikke konkluder «OK» hvis `soffice` mangler — rapporter som blokkering for Wave 4.

  **Recommended Agent Profile**:
  - **Category**: `quick` — ren miljøverifisering, ingen logikkendring.
    - Reason: Enkel diagnostikk med få trinn; krever ikke dyp domeneforståelse.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: Forstå `.xlsx`-struktur ved baseline-unzip og openpyxl-versjonsjekk.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 1 (med T2, T3, T4)
  - **Blocks**: (ingen; leverer baseline brukt av T14)
  - **Blocked By**: None

  **References**:
  - `AGENTS.md` (Commands-seksjonen): generatorkommandoer og krav til verifisering.
  - `README.md` (Krav-seksjonen): Python 3.9+ + openpyxl-krav.
  - `lag_supportsystem.py:37` (`OUT = "Supportsystem.xlsx"`): hvilken fil som er leveransen som skal baselines.

  **Acceptance Criteria**:
  - [ ] `.sisyphus/evidence/baseline/miljoe.txt` inneholder versjoner for soffice/python/openpyxl/pywin32 + Excel-COM-status.
  - [ ] SHA-256 av `Supportsystem.xlsx` lagret i `.sisyphus/evidence/baseline/sha256.txt`.

  **QA Scenarios**:
  ```
  Scenario: Miljø er komplett
    Tool: Bash (PowerShell)
    Preconditions: ingen
    Steps:
      1. kjør 'soffice --version'
      2. kjør 'python --version'
      3. kjør 'python -c "import openpyxl,win32com.client; print(openpyxl.__version__)"'
    Expected Result: soffice-versjon printes (ikke «command not found»); python ≥3.9; openpyxl importerbar + win32com importerbar.
    Failure Indicators: soffice mangler → merk som Wave 4-blokkering; win32com ImportError → merk som Wave 3-blokkering.
    Evidence: .sisyphus/evidence/task-1-miljoe.txt

  Scenario: Baseline fanges
    Tool: Bash (PowerShell)
    Steps:
      1. python lag_supportsystem.py
      2. Get-FileHash Supportsytem.xlsx -Algorithm SHA256
    Expected Result: hash skrevet til baseline/sha256.txt; generator exit 0.
    Evidence: .sisyphus/evidence/task-1-baseline.txt
  ```

  **Commit**: YES (grupperes med T2)
  - Message: `chore: baseline og miljøvalidering`
  - Files: `.sisyphus/evidence/baseline/*`

- [x] 2. Trekk ut felles kjerne `build_common.py`

  **What to do**:
  - Refaktorer `lag_supportsystem.py` slik at alt som er felles på tvers av mål (ark-opprettelse, stiler/farger, Oppsett-grunndata + dynamiske navn, Saker/KB-header + statiske data + betinget formatering + DV, KPI-formler, diagramdefinisjoner, Instruksjoner-base, sample-data, avkryssingsboks-injeksjon) flyttes til `build_common.py` med en funksjon `build_skeleton(profile: dict) -> Workbook`.
  - Innfør en `profile`-datastruktur som styrer divergensene: `timestamp_mechanism` (`formula` vs `macro`), `search_mechanism` (`formula` vs `macro`). Fallback-profilen må gjengi dagens utdata **byte-identisk**.
  - `lag_supportsystem.py` blir et tynt skall som kaller `build_common.py` med fallback-profilen (bakoverkompatibelt).
  - Behold **sirkulær referanse i Saker!B** og **hjelpekolonnene Y/L** KUN i fallback-profilen — flytt dem inn i profil-betinget kode, men IKKE slett dem fra fallback-banen.

  **Must NOT do**:
  - Ikke endre utdata for fallback (byte-stabil mot T1-baseline).
  - Ikke fjern sirkulær referanse, `wb.calculation.iterate`, eller «ikke test kolonne B»-kommentarene fra delt kode som fallback bruker.
  - Ikke introduser `ws.add_table`, M365-funksjoner eller `ManualLayout`.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — stor, presisjonskrevende refaktorering med regresjonsrisiko.
    - Reason: Krever dyp forståelse av hele generatoren + disiplin rundt byte-stabilitet.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: openpyxl-spesifikk refaktorering (Workbook, DefinedName, DataValidation, Chart).

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 1 (med T1, T3, T4)
  - **Blocks**: T5
  - **Blocked By**: None

  **References**:
  - `lag_supportsystem.py:64-72` — ark-opprettelse og rekkefølge (må bevares eksakt).
  - `lag_supportsystem.py:165-184` — dynamiske navn (`names`-dict) som må flyttes uendret.
  - `lag_supportsystem.py:202-228` — Saker-formlene inkl. sirkulær referanse (B) og hjelpekolonne Y.
  - `lag_supportsystem.py:370-378` — KB-formlene inkl. hjelpekolonne L.
  - `lag_supportsystem.py:584-593` — KPI-formler (COUNTIFS, «ikke kolonne B»-regelen).
  - `lag_supportsystem.py:796-981` — avkryssingsboks-injeksjon (må flyttes uendret, fallback bruker den).

  **Acceptance Criteria**:
  - [ ] `python lag_supportsystem.py` produserer `Supportsystem.xlsx` som er **innholds-stabil** mot original (identiske zip-oppføringer, ignorer tidsstempelfelt i docProps). NB: byte-identisk SHA-256 er umulig for openpyxl (tidsstempler endres hver kjøring).
  - [ ] `build_common.py` eksisterer med `build_skeleton(profile)`.

  **QA Scenarios**:
  ```
  Scenario: Fallback er innholds-stabil
    Tool: Bash (PowerShell) — python
    Steps:
      1. python lag_supportsystem.py
      2. kjør python-script som unzipper Supportsytem.xlsx og git HEAD-versjonen, normaliserer docProps-tidsstempler, og sammenligner alle oppføringer
    Expected Result: 0 ulikheter (identiske zip-oppføringer etter normalisering).
    Failure Indicators: noen oppføring ulik (utenom tidsstempler) → regresjon, avvis.
    Evidence: .sisyphus/evidence/task-2-byte-stabil.txt
  ```

  **Commit**: YES
  - Message: `refactor(generator): trekk ut felles kjerne build_common.py`
  - Files: `build_common.py`, `lag_supportsystem.py`

- [x] 3. Spike: VBA-injeksjon (proof of concept)

  **What to do**:
  - Lag et minimalt openpyxl-ark, og bevis at en VBA-makro kan injiseres og KJØRES i Excel.
  - Author en triviell makro (`Sub SayHello(): MsgBox "ok": End Sub`) via Excel-COM (pywin32) inn i en donor `.xlsm`, ekstraher `xl/vbaProject.bin` fra donoren.
  - Injiser binæren i et openpyxl-generert skjelett: patch `[Content_Types].xml` (Override for `/xl/vbaProject.bin` + `application/vnd.ms-excel.sheet.macroEnabled.main+xml` på workbook) og `xl/_rels/workbook.xml.rels` (vbaProject-relasjon).
  - Dokumentér nøyaktige steg + XML-patches i `.sisyphus/evidence/spike/`.

  **Must NOT do**:
  - Ikke bygg den fullstendige VBA-en her — kun proof of concept.
  - Ikke hardkod sheet-navn/codenames som senere skal komme fra `kontrakt.py`.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — teknisk risikofylt mekanikk (binær OLE + XML-patching).
    - Reason: Krever presis pakke-manipulering og COM-kunnskap.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: openpyxl + zip/XML-struktur i `.xlsm`.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 1 (med T1, T2, T4)
  - **Blocks**: T7
  - **Blocked By**: None

  **References**:
  - `lag_supportsystem.py:781-982` — eksisterende zip/XML-post-save-mønster (injiser_avkryssingsbokser) som inspirasjon for XML-patching.
  - Forskningsfunn (librarian): `[Content_Types].xml` Override + macroEnabled main+xml; `xl/_rels/workbook.xml.rels` vbaProject-relasjon.

  **Acceptance Criteria**:
  - [ ] `.sisyphus/evidence/spike/vba/` inneholder donor-trinndokumentasjon + injeksjonsskript.
  - [ ] Injisert `.xlsm` åpner i Excel uten reparasjonsprompt, og `SayHello` kjører (verifisert via COM `Application.Run` eller manuell notis).

  **QA Scenarios**:
  ```
  Scenario: VBA-injeksjon verifiseres
    Tool: Bash (PowerShell) + Excel COM
    Steps:
      1. python <injeksjonsskript>
      2. openpyxl.load_workbook('spike.xlsm', keep_vba=True) → assert 'vbaProject.bin' i wb.vba_archive
      3. COM: Excel.Application → Open spike.xlsm → Application.Run('SayHello')
    Expected Result: vbaProject.bin finnes; SayHello kjører uten feil (returnerer/viser "ok").
    Failure Indicators: 'vbaProject.bin' mangler, eller Excel repareringsprompt, eller Run() kaster.
    Evidence: .sisyphus/evidence/task-3-vba-spike.txt
  ```

  **Commit**: YES (kun spike-artefakter, ikke binær i hovedrepo ennå)
  - Message: `spike(vba): proof of concept for vbaProject.bin-injeksjon`
  - Files: `.sisyphus/evidence/spike/vba/*`

- [x] 4. Spike: Basic-injeksjon (proof of concept)

  **What to do**:
  - Lag et minimalt `.xlsx`, konverter med `soffice --headless --convert-to ods`, og bevis at en Basic-makro kan injiseres i `.ods`-pakken.
  - Injiser Basic-tekst-XML: `Basic/script-lc.xml`, `Basic/Standard/script-lb.xml`, `Basic/Standard/Module1.xml` (triviell `Sub Hello`).
  - Verifiser at `.ods` åpner i LibreOffice uten reparasjon (headless konvertering tilbake eller `soffice --headless --convert-to xlsx` som åpningsvalidering).
  - Dokumentér XML-strukturen i `.sisyphus/evidence/spike/basic/`.

  **Must NOT do**:
  - Ikke forfatt den fullstendige Basic-logikken her.
  - Ikke anta at event-binding kan gjøres uten GUI — dokumentér hva som er mulig headless.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — ODF-pakke + Basic-XML-mekanikk.
    - Reason: Krever ODF-strukturforståelse + soffice-headless-verifisering.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: `.ods`-format og soffice-konvertering.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 1 (med T1, T2, T3)
  - **Blocks**: T11
  - **Blocked By**: None

  **References**:
  - Forskningsfunn (librarian): `Basic/script-lc.xml` (bibliotek-indeks), `Basic/Standard/script-lb.xml` (modul-indeks), `Basic/Standard/Module1.xml` (kilde `<script:source>`).
  - `README.md:31-43` — dagens LO/OO-forutsetninger (iterative referanser, avkryssingsbokser).

  **Acceptance Criteria**:
  - [ ] `.sisyphus/evidence/spike/basic/` inneholder injeksjonsskript + mal-XML.
  - [ ] Injisert `.ods` inneholder `Basic/Standard/Module1.xml` og konverterer headless uten feil.

  **QA Scenarios**:
  ```
  Scenario: Basic-injeksjon verifiseres
    Tool: Bash (PowerShell) + soffice headless
    Steps:
      1. python <injeksjonsskript>
      2. Expand-Archive (eller python zipfile) → assert 'Basic/Standard/Module1.xml' finnes
      3. soffice --headless --convert-to xlsx spike.ods
    Expected Result: Module1.xml finnes med korrekt <script:source>; konvertering exit 0 (ingen reparasjon).
    Failure Indicators: Module1.xml mangler, eller soffice konvertering feiler.
    Evidence: .sisyphus/evidence/task-4-basic-spike.txt
  ```

  **Commit**: YES
  - Message: `spike(basic): proof of concept for Basic-injeksjon i .ods`
  - Files: `.sisyphus/evidence/spike/basic/*`

- [x] 5. Frys celle-kontrakt `kontrakt.py` (HARD gate)

  **What to do**:
  - Opprett `kontrakt.py` som den ENESTE kilden til sannhet for cellereferanser og konstanter begge makroforfattere og generatoren koder mot. Innhold:
    - `MAX_S = 501`, `MAX_K = 301`; ark-navn OG rekkefølge OG `codeName` (f.eks. `Dashboard`=Sheet1, `Saker`=Sheet2, `Søk KB`=Sheet3, `Søk saker`=Sheet4, `Kunnskapsbase`=Sheet5, `Oppsett`=Sheet6, `Instruksjoner`=Sheet7).
    - Header-layouts (Saker 24 kolonner, KB 11 kolonner), KPI-celler (`A4:L5`), hjelpetabell-områder (`A26:D42`), diagram-ankre (`A8`/`F8`/`F26`), hjelpekolonner (`Saker!Y`, `Kunnskapsbase!L`), speilkolonner (`Søk saker!AA`/`AB`).
    - **Scorings-orakel** som ren-Python-referanse (se algoritme under).
    - Tre målprofiler: `fallback` (formel-tidsstempel + formel-søk), `excel` (makro-tidsstempel + makro-søk), `calc` (makro-tidsstempel + makro-søk).
  - Sett `ws.sheet_properties.codeName` eksplisitt i `build_common.py` per kontrakt (openpyxl auto-codenames matcher ikke en hånd-authored donor).
  - Definer **scoringsalgoritmen skriftlig** (må være identisk i Python/VBA/Basic):
    - `normalize(s)` = lowercase, FJERN bindestreker (`-` → tom, som `SUBSTITUTE(x,"-","")` i originalen), kollaps flerrom, trim. NB: «wifi» skal treffe «Wi-Fi» som EKSAKT treff (bindestrek fjernes, ikke erstattes med mellomrom).
    - Felt-vekter: tittel=4, nøkkelord=3, problem/beskrivelse=2, løsning=1.
    - `score(query, doc)` = sum over query-termer av sum over felt der termen er substring av `normalize(felt)`.
    - Rangering: score synkende → dato synkende → rad-nr stigende.
    - Fuzzy-fallback: hvis score=0 for alle, match termer mot ord med Levenshtein-avstand ≤ 2 (samme vekt).

  **Must NOT do**:
  - Ikke la magiske tall overleve utenfor `kontrakt.py` etter denne oppgaven.
  - Ikke endre `codeName` på en måte som bryter fallback (fallback tolererer eksplisitte codenames — de må bare være stabile).

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — presisjonsarbeid som definerer hele kontrakten.
    - Reason: Feil her forplanter seg til både VBA og Basic; krever grundig gjennomlesning av generatoren.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: openpyxl `codeName`, `DefinedName`, cellekart.

  **Parallelization**:
  - **Can Run In Parallel**: NO (hard gate — kjører alene i Wave 2)
  - **Blocks**: T6, T9, T10, T12
  - **Blocked By**: T2

  **References**:
  - `lag_supportsystem.py:64-72` — ark-navn og opprettelsesrekkefølge (grunnlag for codeName-kontrakt).
  - `lag_supportsystem.py:190-195` — Saker-header (24 kolonner).
  - `lag_supportsystem.py:361-367` — KB-header (11 kolonner).
  - `lag_supportsystem.py:584-593` — KPI-celler og formler.
  - `lag_supportsystem.py:604-675` — hjelpetabeller + diagram-ankre.
  - `lag_supportsystem.py:201,369` — `MAX_S`/`MAX_K`-verdier.

  **Acceptance Criteria**:
  - [ ] `kontrakt.py` eksporterer alle konstanter + `score(query, doc)` + `PROFILES`-dict.
  - [ ] openpyxl read-back av generert `.xlsx` bekrefter `codeName` på hvert ark = kontraktens verdi.

  **QA Scenarios**:
  ```
  Scenario: Scorings-orakel rangerer kjente spørringer riktig
    Tool: Bash (PowerShell) — python
    Steps:
      1. python -c "from kontrakt import score; docs=[{'title':'Wi-Fi nede','solution':'restartet AP','dato':'2026-09-01','rad':1},{'title':'Glemt VPN-passord','solution':'tilbakestilt i AD','dato':'2026-09-01','rad':2}]; print([score('wifi', d) for d in docs])"
    Expected Result: første doc får høyere score enn andre (wifi matcher tittel, vekt 4).
    Failure Indicators: andre doc får ≥ første; eller -normalisering (wifi vs Wi-Fi) slår feil.
    Evidence: .sisyphus/evidence/task-5-orakel.txt

  Scenario: Fuzzy-fallback fanger skrivefeil
    Tool: Bash — python
    Steps:
      1. python -c "from kontrakt import score; print(score('wify', {'title':'Wi-Fi nede','solution':'','dato':'2026-09-01','rad':1}))"
    Expected Result: score > 0 (Levenshtein(wify, wifi) ≤ 2).
    Evidence: .sisyphus/evidence/task-5-fuzzy.txt
  ```

  **Commit**: YES
  - Message: `refactor(generator): frys celle-kontrakt og scorings-orakel`
  - Files: `kontrakt.py`, `build_common.py`

- [x] 6. Forfatt VBA-kilde (`.bas`) — tidsstempel + søk + dashboard

  **What to do**:
  - Skriv VBA-kilde som `.bas`-filer (reviewbar kilde som bygges inn i donor i T7). Moduler:
    - `ModFelles`: `Normaliser(t)` (lowercase, `-`→mellomrom, kollaps, trim), `Levenshtein(a,b)`, `ScoreDok(query, feltTittel, feltNokkelord, feltProblem, feltLosning)` — implementerer `kontrakt.py`-algoritmen eksakt (tittel=4/nøkkelord=3/problem=2/løsning=1, fuzzy ≤2).
    - `ModTidsstempel`: `Worksheet_Change`-logikk på Saker — når Innmelder (C) eller Tittel (H) fylles og Opprettet (B) er tom, sett B = Now(). Bruk `Application.EnableEvents = False/True` + feilhåndterer som alltid gjenoppretter.
    - `ModSok`: les inputcelle, iterér dataradene (last inn i array for ytelse), beregn score, ranger, skriv topp-N til resultatområdet, marker treffrader (helcelle `Bold` + fyll), tøm forrige resultat. Debounce med `Application.OnTime`.
    - `ModDashboard`: `OppdaterDashboard` — bygg om diagram-dataserier fra Oppsett-listene (voksende kategori/ansatt/status-antall) og resizer faste diagramobjekter; knytt til `Workbook_SheetActivate` for auto-oppfriskning.
  - Håndter `ren`/tom-modus (0 rader), æøå (Unicode), bindestrek-normalisering («wifi» treffer «Wi-Fi»).

  **Must NOT do**:
  - Ikke flytt diagramobjekter vilkårlig — kun rebuild av dataserier + resize.
  - Ikke bruk M365-spesifikke VBA-anrop; hold deg til Excel 2007+-kompatibel VBA.
  - Ikke hardkod celleadresser — les fra `kontrakt.py` (vedlikeholdes manuelt synkronisert).

  **Recommended Agent Profile**:
  - **Category**: `ultrabrain` — ikke-triviell logikk (scoring, fuzzy, event-håndtering, debounce).
    - Reason: Søkerangering + re-entrancy-guard + ytelse krever nøye logikkdesign.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: kontekst om celleadresser/formler som VBA-en opererer på.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 3 (med T9; T7/T8 avhenger av denne)
  - **Blocks**: T7
  - **Blocked By**: T5

  **References**:
  - `lag_supportsystem.py:220-224` — dagens søkelogikk (normalisering + SEARCH) som VBA-en skal replikere.
  - `lag_supportsystem.py:374-377` — KB-søklogikk.
  - `lag_supportsystem.py:202-212` — tidsstempel-formel (nå erstattet av `ModTidsstempel`).
  - `lag_supportsystem.py:604-675` — hjelpetabeller/diagrammer som `ModDashboard` bygger om.
  - `kontrakt.py` — scoringsalgoritme + cellekontrakt.

  **Acceptance Criteria**:
  - [ ] `.bas`-filer finnes: `ModFelles.bas`, `ModTidsstempel.bas`, `ModSok.bas`, `ModDashboard.bas`.
  - [ ] Hver `.bas` har `Option Explicit` og kommentarer på norsk.

  **QA Scenarios**:
  ```
  Scenario: Scorings-funksjon i VBA matcher orakel
    Tool: Excel COM (pywin32)
    Steps:
      1. Importer ModFelles.bas i donor og kjør ?ScoreDok("wifi","Wi-Fi nede","","","") i Immediate-vinduet
    Expected Result: returnerer 4 (tittel-vekt); ?ScoreDok("wify", ...) returnerer > 0 (fuzzy).
    Failure Indicators: returnerer 0 eller feil verdi.
    Evidence: .sisyphus/evidence/task-6-score.txt

  Scenario: Tidsstempel skrives ved input
    Tool: Excel COM
    Steps:
      1. Sett Saker!C2 = "Test" → sjekk Saker!B2
    Expected Result: B2 får Now() (ikke tom).
    Evidence: .sisyphus/evidence/task-6-tidsstempel.txt
  ```

  **Commit**: YES
  - Message: `feat(excel): VBA-kilde for tidsstempel, søk og dashboard`
  - Files: `vba/*.bas`

- [x] 7. Bygg donor `.xlsm` + ekstraher `vbaProject.bin`

  **What to do**:
  - Bruk Excel-COM (pywin32) til å bygge en donor `.xlsm` fra `.bas`-filene i T6 (ikke manuell GUI): `VBProject.VBComponents.Add(1).CodeModule.AddFromString(...)` per modul, lagre som donor `.xlsm`.
  - Ekstraher `xl/vbaProject.bin` fra donoren og sjekk inn i `vba/vbaProject.bin` (binær artefakt). Eksporter også `.bas`-kildene (allerede i repo) for review.
  - Dokumentér i `vba/README.md` hvordan donoren gjenbygges når makroer endres (gjenta COM-trinnene + re-ekstraksjon).

  **Must NOT do**:
  - Ikke commit donoren `.xlsm` som helhet i produktlinjen (kun den ekstraherte binæren + `.bas`).
  - Ikke endre makro-logikken her (kun pakking/ekstraksjon).

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — COM-automatisering + binær håndtering.
    - Reason: Krever pywin32 COM + forståelse av vbaProject.bin-pakken.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: `.xlsm`-pakke og openpyxl `keep_vba`.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 3 (med T9)
  - **Blocks**: T8
  - **Blocked By**: T6, T3

  **References**:
  - `vba/*.bas` (T6) — kildene som bygges inn.
  - `.sisyphus/evidence/spike/vba/` (T3) — donor-/ekstraksjonsoppskriften.

  **Acceptance Criteria**:
  - [ ] `vba/vbaProject.bin` finnes (binær) + `vba/README.md` beskriver gjenbyggingsprosedyren.

  **QA Scenarios**:
  ```
  Scenario: Donor bygges og binær ekstraheres
    Tool: Bash (PowerShell) — python COM-skript
    Steps:
      1. python <build_donor.py>
      2. assert vba/vbaProject.bin > 0 bytes
      3. openpyxl.load_workbook('donor.xlsm', keep_vba=True) → assert vba_archive
    Expected Result: binær ekstrahert og lesbar.
    Failure Indicators: binær tom/mangler, eller COM feiler (Excel ikke tilgjengelig).
    Evidence: .sisyphus/evidence/task-7-donor.txt
  ```

  **Commit**: YES
  - Message: `feat(excel): donor-bygging + vbaProject.bin-artefakt`
  - Files: `vba/vbaProject.bin`, `vba/README.md`, `vba/build_donor.py`

- [x] 8. Implementer `--target excel` (injiser VBA → `.xlsm`)

  **What to do**:
  - I `lag_supportsystem.py`: legg til `--target excel`. Bruk `build_common.build_skeleton(profile='excel')` — profil med `timestamp_mechanism='macro'` (B som statisk tom celle, INGEN sirkulær referanse, INGEN hjelpekolonner Y/L) og `search_mechanism='macro'` (søkeark som rene resultatområder uten INDEX/MATCH).
  - Injiser `vba/vbaProject.bin` post-save: patch `[Content_Types].xml` (Override `/xl/vbaProject.bin` + `application/vnd.ms-excel.sheet.macroEnabled.main+xml`) og `xl/_rels/workbook.xml.rels` (vbaProject-relasjon). Lagre som `Supportsystem.xlsm`.
  - Følg mønsteret fra `injiser_avkryssingsbokser` (zip/XML post-save).

  **Must NOT do**:
  - Ikke fjern sirkulær referanse eller hjelpekolonner fra `build_common.py` sin fallback-bane.
  - Ikke behold `wb.calculation.iterate` som en nødvendighet for excel-profilen (makro-tidsstempel trenger den ikke), men endre kun i excel-profilen.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — generator-utvidelse + pakke-injeksjon.
    - Reason: Krever både profil-betinget logikk og `.xlsm`-pakking.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: openpyxl + zip/XML-injeksjon.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 3 (med T9)
  - **Blocks**: T13, T14
  - **Blocked By**: T7

  **References**:
  - `lag_supportsystem.py:781-982` — eksisterende post-save zip/XML-injeksjonsmønster.
  - `lag_supportsystem.py:95-103` — `wb.calculation`-innstillinger som må profileres.
  - `vba/vbaProject.bin` (T7) — binæren som injiseres.

  **Acceptance Criteria**:
  - [ ] `python lag_supportsystem.py --target excel` produserer `Supportsystem.xlsm`.
  - [ ] openpyxl `keep_vba=True` viser `vbaProject.bin`; `[Content_Types].xml` har macroEnabled + vbaProject-override.

  **QA Scenarios**:
  ```
  Scenario: .xlsm genereres med VBA
    Tool: Bash (PowerShell) — python
    Steps:
      1. python lag_supportsystem.py --target excel
      2. openpyxl.load_workbook('Supportsytem.xlsm', keep_vba=True) → assert vba_archive har 'vbaProject.bin'
      3. unzip → assert '[Content_Types].xml' inneholder 'macroEnabled' og 'vbaProject'
    Expected Result: alle asserts passerer.
    Failure Indicators: vbaProject.bin mangler, eller content-type ikke macroEnabled.
    Evidence: .sisyphus/evidence/task-8-xlsm.txt

  Scenario: Excel-profilen dropper sirkulær referanse
    Tool: Bash — python
    Steps:
      1. read-back Saker!B2 i .xlsm → assert IKKE sirkulær formel
      2. assert hjelpekolonne Y/L ikke har formler (eller kolonnene fjernet)
    Expected Result: B2 er tom/statisk; ingen Y/L-formler.
    Evidence: .sisyphus/evidence/task-8-profil.txt
  ```

  **Commit**: YES
  - Message: `feat(excel): --target excel med vbaProject.bin-injeksjon`
  - Files: `lag_supportsystem.py`

- [x] 9. Excel README + makro-sikkerhetsveiledning

  **What to do**:
  - Oppdater `README.md` med en Excel-spesifikk seksjon: hvordan aktivere makroer (Trust Center → Macro Settings, «Unblock» fil i Fil-egenskaper), hvilken fil som brukes når, og at makro-søk/dashboard-dynamikk kun virker med makroer aktivert.
  - Dokumentér forskjellen mot fallback: makro-versjonen har sanntid søk med rangering/fuzzy/markering og dynamisk dashboard; fallback er makrofri og alltid-trygg.
  - Notér at Excel 2007+ kreves, og at `.xlsm` må åpnes fra en klarert plassering.

  **Must NOT do**:
  - Ikke fjern eksisterende fallback-dokumentasjon.

  **Recommended Agent Profile**:
  - **Category**: `writing` — dokumentasjon.
    - Reason: Ren prose/veiledning, ingen kode.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: terminologi om Excel-makrosikkerhet.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 3 (med T6/T7/T8)
  - **Blocks**: (ingen)
  - **Blocked By**: T5

  **References**:
  - `README.md:21-43` — eksisterende krav/seksjoner å bygge på.
  - Forskningsfunn: Excel Trust Center / Macro Security / Trusted Locations.

  **Acceptance Criteria**:
  - [ ] README har en `### Excel (.xlsm) — makroer`-seksjon med aktiveringssteg.

  **QA Scenarios**:
  ```
  Scenario: README nevner makro-aktivering
    Tool: Grep
    Steps:
      1. grep -i "makro\|Trust Center\|Unblock" README.md
    Expected Result: minst én treff med konkret aktiveringsveiledning.
    Evidence: .sisyphus/evidence/task-9-readme.txt
  ```

  **Commit**: YES
  - Message: `docs: Excel-makroveiledning i README`
  - Files: `README.md`

- [x] 10. Forfatt Basic-makroer (tekst) — tidsstempel + søk + dashboard

  **What to do**:
  - Skriv Basic-kilde som tekst (skal bli `Basic/Standard/Module1.xml` m.m.). Moduler:
    - `ModFelles`: `Normaliser(t)`, `Levenshtein(a,b)`, `ScoreDok(...)` — samme algoritme som `kontrakt.py` (tittel=4/nøkkelord=3/problem=2/løsning=1, fuzzy ≤2).
    - `ModTidsstempel`: siden LO/OO mangler `Worksheet_Change`, implementer tidsstempel som en knapp-kjørt `StemplTidsstempel` (iterer Saker, sett B=Now der C/H er fylt og B tom) — KAN evt. registreres som `addModifyListener` via `AutoOpen` med module-level-variabel for å hindre GC (best-effort).
    - `ModSok`: **knapp-triggeret** søk (knyttet til «SØK I KB →» / «SØK I SAKER →»-knappene) — les input, `getDataArray()` for ytelse, score, ranger, skriv topp-N, marker treffrader (helcelle). Best-effort `AutoOpen`-registrert listener som fallback for type-ahead.
    - `ModDashboard`: `OppdaterDashboard` — bygg om diagram-dataserier og resizer faste diagramobjekter (via DrawPage shapes); knapp-/ark-aktivering.
  - Håndter `ren`/tom-modus, æøå, bindestrek-normalisering. Bruk `getDataArray()` (ikke celle-for-celle).

  **Must NOT do**:
  - Ikke anta persistent event-binding uten GUI — design mot knapp-triggeret + AutoOpen-listener som fallback.
  - Ikke flytt diagramobjekter vilkårlig; kun rebuild av dataserier + resize.

  **Recommended Agent Profile**:
  - **Category**: `deep` — StarBasic/UNO-kompleksitet + listener-GC-fallgruver.
    - Reason: Krever nøye håndtering av UNO-API, listener-levetid og array-basert ytelse.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: `.ods`-struktur og Calc-objektmodell.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 4 (med T12; T11 avhenger av denne)
  - **Blocks**: T11
  - **Blocked By**: T5

  **References**:
  - `lag_supportsystem.py:220-224` + `374-377` — søkelogikk å replikere.
  - `lag_supportsystem.py:202-212` — tidsstempel å erstatte.
  - `kontrakt.py` — scoringsalgoritme + cellekontrakt.
  - Forskningsfunn: `addModifyListener` + module-level-variabler; `getDataArray()` for ytelse.

  **Acceptance Criteria**:
  - [ ] Basic-kilde foreligger som tekst for `ModFelles`, `ModTidsstempel`, `ModSok`, `ModDashboard`.

  **QA Scenarios**:
  ```
  Scenario: Basic-kilde inneholder søk + dashboard-makroer
    Tool: Grep
    Steps:
      1. grep "Sub" i basic-kildene
    Expected Result: Sub-er for tidsstempel, søk og dashboard finnes.
    Evidence: .sisyphus/evidence/task-10-basic.txt

  Scenario: Scoringslogikk matcher kontrakt (strukturelt)
    Tool: Grep
    Steps:
      1. grep -i "levenshtein\|score" i basic-kildene
    Expected Result: fuzzy + scoring-logikk til stede.
    Evidence: .sisyphus/evidence/task-10-score.txt
  ```

  **Commit**: YES
  - Message: `feat(calc): Basic-kilde for tidsstempel, søk og dashboard`
  - Files: `basic/*`

- [x] 11. Implementer `--target calc` (soffice + injiser Basic → `.ods`)

  **What to do**:
  - I `lag_supportsystem.py`: legg til `--target calc`. Bruk `build_common.build_skeleton(profile='calc')` (makro-tidsstempel + makro-søk, som excel-profilen).
  - Pipeline: generer skjelett som midlertidig `.xlsx` → `soffice --headless --convert-to ods` → zip-injiser Basic-XML (`Basic/script-lc.xml`, `Basic/Standard/script-lb.xml`, `Basic/Standard/Module1.xml` med kilden fra T10) → lagre som `Supportsystem.ods`.
  - Sett event-binding der det er mulig headless (eller dokumentér at knapp-triggeret brukes).

  **Must NOT do**:
  - Ikke fjern `.xlsx`-skjelett-koden fra felles kjerne.
  - Ikke hardkod soffice-sti — bruk `soffice` fra PATH med klar feilmelding hvis mangler.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — soffice-konvertering + ODS-pakke-injeksjon.
    - Reason: Krever konverteringspipeline + zip/XML-injeksjon av Basic.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: `.ods`-format og soffice headless.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 4 (med T12)
  - **Blocks**: T13, T14, T15
  - **Blocked By**: T10, T4

  **References**:
  - `.sisyphus/evidence/spike/basic/` (T4) — injeksjonsoppskriften.
  - `basic/*` (T10) — kildene som injiseres.

  **Acceptance Criteria**:
  - [ ] `python lag_supportsystem.py --target calc` produserer `Supportsystem.ods`.
  - [ ] unzip viser `Basic/Standard/Module1.xml` + `Basic/script-lc.xml` + `Basic/Standard/script-lb.xml`.

  **QA Scenarios**:
  ```
  Scenario: .ods genereres med Basic-moduler
    Tool: Bash (PowerShell) — python
    Steps:
      1. python lag_supportsystem.py --target calc
      2. python -c "import zipfile; print([n for n in zipfile.ZipFile('Supportsytem.ods').namelist() if n.startswith('Basic/')])"
    Expected Result: Basic/Standard/Module1.xml, script-lc.xml, script-lb.xml listes.
    Failure Indicators: Basic/-deler mangler, eller konvertering feiler (soffice mangler).
    Evidence: .sisyphus/evidence/task-11-ods.txt
  ```

  **Commit**: YES
  - Message: `feat(calc): --target calc med soffice-konvertering + Basic-injeksjon`
  - Files: `lag_supportsystem.py`

- [x] 12. LO/OO README + makro-sikkerhetsveiledning

  **What to do**:
  - Oppdater `README.md` med LO/OO-seksjon: aktiver makroer (Tools → Options → Security → Macro Security), bruk av `.ods`, at søk er knapp-triggeret (og evt. best-effort type-ahead), og at iterative referanser IKKE lenger trengs i makro-versjonen (men fortsatt i fallback).
  - Notér at OpenOffice-støtte er best-effort (testet primært i LibreOffice).

  **Must NOT do**:
  - Ikke fjern fallback-dokumentasjonen eller Excel-seksjonen.

  **Recommended Agent Profile**:
  - **Category**: `writing` — dokumentasjon.
    - Reason: Ren veiledning.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: LO/OO makrosikkerhet-terminologi.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 4 (med T10/T11)
  - **Blocks**: (ingen)
  - **Blocked By**: T5

  **References**:
  - `README.md:31-43` — eksisterende LO/OO-forutsetninger (iterative referanser).
  - Forskningsfunn: LO Macro Security-nivåer.

  **Acceptance Criteria**:
  - [ ] README har `### LibreOffice/OpenOffice (.ods) — makroer`-seksjon med aktiveringssteg.

  **QA Scenarios**:
  ```
  Scenario: README nevner LO-makrosikkerhet
    Tool: Grep
    Steps:
      1. grep -i "macro security\|makrosikkerhet\|.ods" README.md
    Expected Result: minst én treff med aktiveringsveiledning.
    Evidence: .sisyphus/evidence/task-12-readme.txt
  ```

  **Commit**: YES
  - Message: `docs: LO/OO-makroveiledning i README`
  - Files: `README.md`

- [x] 13. Generaliser `backup.py` for `.xlsm`/`.ods`/`.xlsx`

  **What to do**:
  - `backup.py:76`: `backupnavn = f"{navn_grunn}_{stempel}.xlsx"` → bruk kildefilens faktiske endelse: `ext = os.path.splitext(os.path.basename(kilde))[1]` og `backupnavn = f"{navn_grunn}_{stempel}{ext}"`.
  - `backup.py:101`: `f.endswith(".xlsx")` i `roter()` → match mot samme `ext` (send inn endelse eller avled fra `navn_grunn`/kildefil).
  - `backup.py:28`: behold `"kildefil": "Supportsystem.xlsx"` som standard, men legg til kommentar om at den kan peke til `.xlsm`/`.ods`.
  - `backup_service.py:35`: oppdater tjenestebeskrivelsen (nevner «Supportsystem.xlsx») til å være endelsesnøytral.
  - `backup_config.json`: behold, men dokumentér at `kildefil` kan byttes til aktiv variant.

  **Must NOT do**:
  - Ikke bygg et multi-fil-orkestreringssystem — kun én aktiv kildefil per config.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-low` — liten, avgrenset endring.
    - Reason: To-linjers logikkendring + docstring.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: filendelses-kontekst.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 5 (med T14, T15)
  - **Blocks**: (ingen)
  - **Blocked By**: T8, T11

  **References**:
  - `backup.py:76` — hardkodet `.xlsx` i backupnavn.
  - `backup.py:96-111` — `roter()` med `.endswith(".xlsx")`.
  - `backup.py:27-34` — STANDARD_CONFIG.
  - `backup_service.py:34-38` — tjenestebeskrivelse.

  **Acceptance Criteria**:
  - [ ] `backup.py` bevarer kildefilens endelse i backup-navnet og rotasjonen.

  **QA Scenarios**:
  ```
  Scenario: Backup av .xlsm bevarer endelse
    Tool: Bash (PowerShell)
    Steps:
      1. sett backup_config.json kildefil="Supportsystem.xlsm"; sørg for at filen finnes
      2. python backup.py
      3. sjekk backups/-mappen for Supportsytem_*.xlsm
    Expected Result: ny .xlsm-backup opprettet (ikke .xlsx), rotasjon virker.
    Failure Indicators: backup får .xlsx-endelse, eller rotasjon matcher ikke.
    Evidence: .sisyphus/evidence/task-13-backup.txt

  Scenario: Rotasjon sletter kun riktig endelse
    Tool: Bash
    Steps:
      1. python backup.py med behold_antall=1 to ganger → assert kun 1 .xlsm-backup igjen
    Expected Result: eldste .xlsm-kopi rotert bort.
    Evidence: .sisyphus/evidence/task-13-rotasjon.txt
  ```

  **Commit**: YES
  - Message: `feat(backup): generaliser filendelse i backup og rotasjon`
  - Files: `backup.py`, `backup_service.py`

- [x] 14. Strukturell test-suite (`test_struktur.py`)

  **What to do**:
  - Skriv `test_struktur.py` (pytest) med openpyxl read-back for alle tre mål:
    - **Fallback (.xlsx)**: ark-navn + rekkefølge + codeName; `Saker!B2` inneholder sirkulær formel; `Saker!Y2` + `Kunnskapsbase!L2` hjelpeformler; KPI-formler (`A5` etc.); dynamiske navn (`Ansatte` etc.); DV-rules; diagram-ankre; avkryssingsboks-deler i zip.
    - **Excel (.xlsm)**: `keep_vba=True` → `vbaProject.bin` i `vba_archive`; `[Content_Types].xml` macroEnabled + vbaProject-override; Saker!B2 IKKE sirkulær.
    - **LO (.ods)**: unzip → `Basic/Standard/Module1.xml`, `script-lc.xml`, `script-lb.xml`.
  - Kjørbar med `python -m pytest test_struktur.py`.

  **Must NOT do**:
  - Ikke test makro-kjøring her (krever app) — kun struktur/nærvær.

  **Recommended Agent Profile**:
  - **Category**: `unspecified-high` — bred testdekning over tre artefakter.
    - Reason: Krever systematisk openpyxl/zip-assert-design.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: openpyxl read-back + zip-inspeksjon.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 5 (med T13, T15)
  - **Blocks**: (ingen)
  - **Blocked By**: T8, T11

  **References**:
  - `kontrakt.py` — forventede ark-navn/codeName/konstanter.
  - `lag_supportsystem.py:64-72` — ark-rekkefølge.
  - `lag_supportsystem.py:202-215` — sirkulær referanse + hjelpekolonne.

  **Acceptance Criteria**:
  - [ ] `python -m pytest test_struktur.py` → alle tester passerer (0 failures).

  **QA Scenarios**:
  ```
  Scenario: Test-suite passerer for alle tre mål
    Tool: Bash (PowerShell)
    Steps:
      1. python lag_supportsystem.py
      2. python lag_supportsystem.py --target excel
      3. python lag_supportsystem.py --target calc
      4. python -m pytest test_struktur.py -v
    Expected Result: alle asserts PASS, exit 0.
    Failure Indicators: noen test FAIL → regresjon.
    Evidence: .sisyphus/evidence/task-14-tester.txt
  ```

  **Commit**: YES
  - Message: `test: strukturell test-suite for xlsx/xlsm/ods`
  - Files: `test_struktur.py`

- [x] 15. soffice valideringspipeline

  **What to do**:
  - Skriv et valideringsskript (`valider_soffice.py`) som:
    - Konverterer generert `.ods` til `.xlsx` headless (`soffice --headless --convert-to xlsx`) og asserter exit 0 (åpningsvalidering uten reparasjon).
    - Konverterer fallback `.xlsx` til `.ods` headless og asserter at LO kan åpne den (kryss-kompatibilitet).
  - Dokumentér kommandoen i README (Feilsøking).

  **Must NOT do**:
  - Ikke overskriv kildefilene ved konvertering (bruk midlertidig mappe).

  **Recommended Agent Profile**:
  - **Category**: `unspecified-low` — enkel konverteringsvalidering.
    - Reason: Skript med få steg.
  - **Skills**: [`spreadsheets`]
    - `spreadsheets`: soffice headless-konvertering.

  **Parallelization**:
  - **Can Run In Parallel**: YES — Wave 5 (med T13, T14)
  - **Blocks**: (ingen)
  - **Blocked By**: T11

  **References**:
  - `.sisyphus/evidence/spike/basic/` (T4) — soffice-konverteringsmønster.
  - `README.md:21-43` — kryss-kompatibilitetskrav.

  **Acceptance Criteria**:
  - [ ] `python valider_soffice.py` → exit 0 for både `.ods` og `.xlsx`.

  **QA Scenarios**:
  ```
  Scenario: .ods og .xlsx åpner headless uten feil
    Tool: Bash (PowerShell)
    Steps:
      1. python valider_soffice.py
      2. echo $LASTEXITCODE
    Expected Result: exit 0; konverterte filer opprettet i tmp.
    Failure Indicators: soffice feiler (reparasjon/korrupt).
    Evidence: .sisyphus/evidence/task-15-soffice.txt
  ```

  **Commit**: YES
  - Message: `test: soffice valideringspipeline`
  - Files: `valider_soffice.py`

---

## Final Verification Wave (MANDATORY — etter ALLE implementasjonsoppgaver)

> 4 review-agenter kjører PARALLELT. ALLE må GODKJENNE. Presenter konsoliderte resultater for bruker og få eksplisitt «ok» før avslutning.

- [x] F1. **Plan-overholdelse** — `oracle`
  Les planen ende-til-ende. For hver «Must Have»: verifiser implementasjonen finnes (les fil, kjør kommando). For hver «Must NOT Have»: søk kodebasen for forbudte mønstre — avvis med fil:linje. Sjekk evidence-filer finnes. Sammenlign leveranser mot plan.
  Output: `Must Have [N/N] | Must NOT Have [N/N] | Tasks [N/N] | VERDICT: APPROVE/REJECT`

- [x] F2. **Kodekvalitet** — `unspecified-high`
  Kjør `python test_struktur.py` + `python lag_supportsystem.py` (alle tre targets). Gjennomgå endrede filer for: udøde referanser, hardkodede literaler utenfor `kontrakt.py`, ubrukt kode, AI-slop (overkommentering, overabstraksjon, generiske navn).
  Output: `Build [PASS/FAIL] | Tests [N pass/N fail] | Files [N clean/N issues] | VERDICT`

- [x] F3. **Reell manuell QA** — `unspecified-high`
  Start fra ren tilstand. Kjør alle QA-scenarioer fra alle oppgaver — følg steg, fang bevis. Test kryss-oppgave-integrasjon (fallback + .xlsm + .ods sameksisterer). Test edge-cases: `ren` (tom) modus, æøå, bindestrek-normalisering, voksende lister. Lagre til `.sisyphus/evidence/final-qa/`.
  Output: `Scenarios [N/N pass] | Integration [N/N] | Edge [N tested] | VERDICT`

- [x] F4. **Scope-fidelitet** — `deep`
  For hver oppgave: les «What to do», les faktisk diff. Verifiser 1:1 — alt i spec er bygget, ingenting utover spec. Sjekk «Must NOT do»-overholdelse. Oppdag kryss-kontaminering. Flagg urapporterte endringer.
  Output: `Tasks [N/N compliant] | Contamination [CLEAN/N issues] | Unaccounted [CLEAN/N files] | VERDICT`

---

## Commit Strategy

- **1**: `refactor(generator): trekk ut felles kjerne og celle-kontrakt` — `build_common.py`, `kontrakt.py`, `lag_supportsystem.py`
- **2**: `feat(excel): VBA-makroer + .xlsm-generering` — `vba/*.bas`, `vbaProject.bin`, `--target excel`
- **3**: `feat(calc): Basic-makroer + .ods-generering` — `basic/*.xml`, `--target calc`
- **4**: `feat(backup): generaliser filendelse` — `backup.py`, `backup_config.json`, `backup_service.py`
- **5**: `test: strukturell test-suite + docs` — `test_struktur.py`, `README.md`

---

## Success Criteria

### Verifikasjonskommandoer
```bash
python lag_supportsystem.py                  # → Supportsytem.xlsx (fallback, byte-stabil)
python lag_supportsystem.py --target excel   # → Supportsytem.xlsm (VBA)
python lag_supportsystem.py --target calc    # → Supportsytem.ods (Basic)
python test_struktur.py                       # → alle asserts passerer
python backup.py                              # → backup + rotasjon for aktiv kildefil
```

### Sluttliste
- [ ] Alle «Must Have» til stede
- [ ] Alle «Must NOT Have» fraværende
- [ ] Strukturell test-suite passerer
- [ ] Alle tre artefakter åpner i sine mål-apper uten reparasjon
