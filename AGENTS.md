# AGENTS.md

Excel-based support/ticket system (Norwegian) for an operations department. The "product" is the generated `Supportsystem.xlsx` — there is no server, no app, no tests, no CI, and this is not a git repo. `README.md` (Norwegian) is the full user manual; keep it in sync when behavior changes.

## Arbeidsregel (gjelder ALLE økter, også utenfor dette prosjektet)

- **Ikke kjør samme kommando/undersøkelse om og om igjen** i søken etter svar på et problem. Hvis en diagnose ikke gir ny informasjon etter 1–2 forsøk: stopp, tenk på hva som faktisk mangler, endre tilnærming, be om avklaring, eller lever det du har med tydelig status. Gjentatte identiske verktøykall uten fremgang oppleves som en loop og er uakseptabelt.
- Samle verifisering i **én** kjøring (f.eks. ett PowerShell-script som sjekker alt) fremfor mange små identiske kall.

## Commands

- `pip install openpyxl` — only dependency of the generator (no requirements.txt exists)
- `python lag_supportsystem.py` — regenerate `Supportsystem.xlsx` **with** sample data (10 tickets, 6 KB articles)
- `python lag_supportsystem.py ren` — regenerate empty (`ren`/`clean`/`--ren`/`--clean` all work)
- `python backup.py` — one backup now (exit code 0/1); `python backup.py loop` — interval loop
- `python backup_service.py install|start|stop|restart|remove` — Windows service (needs `pip install pywin32` + admin); `installer_tjeneste.bat` / `avinstaller_tjeneste.bat` wrap this
- `installer_oppgave.bat` — scheduled task alternative (no admin); test with `schtasks /run /tn "SupportsystemBackup"`
- Verification = run the generator, then open the xlsx in Excel/LibreOffice and check it opens without repair prompts and formulas compute. There is no automated test suite.

## Hard constraints when editing `lag_supportsystem.py`

Violating these silently breaks the workbook for users — they exist because the file must work in **Excel 2007+, LibreOffice Calc, and OpenOffice Calc**, macro-free:

- **Never use M365/365-only functions** (`FILTER`, `HSTACK`, `XLOOKUP`, spill behavior). Allowed: `VLOOKUP`, `INDEX`, `MATCH`, `COUNTIFS`, `OFFSET`, `COUNTA`, `SEARCH`, `TEXT`. Search uses INDEX/MATCH against hidden helper columns for this reason.
- **No Excel Tables** (`ws.add_table`) — LibreOffice-incompatible. Data areas are free ranges; lists in Oppsett grow via dynamic defined names (`OFFSET`+`COUNTA`, see `names` dict ~line 157). Data validation references these names (`=Kanaler`), never direct cross-sheet ranges.
- **Formulas are written in English with comma separators** (openpyxl requirement) while all UI text/literals are Norwegian (`"Løst"`, `"FORFALT"`). Do not translate formula function names; do not "fix" the mismatch.
- **The circular reference in Saker!B is intentional** — auto-timestamp `=IF($C&$H="","",IF($B="",NOW(),$B))` locks via iterative calculation, which the script enables (`wb.calculation.iterate`, ~line 89). Never remove the circularity or the calc settings.
- **KPI/COUNTIFS formulas must not test column B** (Opprettet) — it is formula-driven and openpyxl writes no cached values, so on first open (before recalc) COUNTIFS sees those cells as empty and counts ~490 blank rows. Test static columns like C (Innmelder) instead. Same class of bug applies anywhere a COUNTIFS/VLOOKUP reads a formula column.
- **Hidden helper columns `Saker!Y` and `Kunnskapsbase!L` power the two search sheets** (`Søk KB` / `Søk saker`, inputs `B4`/`B5` and `B4`). Deleting/reordering them breaks search (README troubleshooting points here).
- **Row capacities are hardcoded**: `MAX_S = 501` (tickets), `MAX_K = 301` (KB articles); formula/DV/CF ranges all use these literals. Changing capacity means updating every `501`/`301` range consistently.
- **Chart legends: never `ManualLayout`** (`openpyxl.chart.layout`) — LibreOffice/OpenOffice draw the legend on top of the plot. Use auto layout (`legend.position = "b"` + `overlay=False`) and always set `series[0].tx = SeriesLabel(v="...")` or legends show "Series1". Dashboard layout: charts share `CHART_H`/`CHART_W` (8.5×16 cm) anchored A8 (pie), F8 (category), F26 (responsible); helper tables live in A26:D42 under the pie (charts read their data from those cells); separator rows 7/25/33 are 3 px. Bar charts use `varyColors = True` + `x_axis/y_axis.delete = True` so each bar gets its own color, the legend lists category names (not "Saker"), and axes are hidden while gridlines stay.
- **Close `Supportsystem.xlsx` in Excel before running the generator** or it fails with Permission denied.
- Changing the built-in lists (`team`, `kategorier`, `sla`, ~line 106-126) + regenerating replaces the whole file — tell users to edit the **Oppsett** sheet live instead; regeneration is for template changes only.

## Backup subsystem

- `backup_config.json` is auto-created with `STANDARD_CONFIG` defaults if missing, and missing keys are backfilled on every run — safe to delete for a reset, safe to edit partially.
- `behold_antall: 0` = keep all copies; rotation happens per backup run, not at startup. `nettverk_backupmappe` empty = local only; logging never crashes a run (`logg()` swallows OSError).
- Service and scheduled task share the name **SupportsystemBackup**. The service re-reads config only at start — restart it after editing `backup_config.json` or `backup.py`.
- All Python/bat source, comments, log output, and the xlsx UI are **Norwegian** — keep new strings Norwegian to match.
