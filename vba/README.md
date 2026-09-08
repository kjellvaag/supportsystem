# VBA-bygging for supportsystemet

Denne mappen inneholder VBA-kildekoden og byggeverktøyet for Excel-makroversjonen
(`Supportsystem.xlsm`).

## Filer

| Fil | Formål |
|---|---|
| `ModFelles.bas` | Felles hjelpefunksjoner: normalisering, Levenshtein, scoring. |
| `ModSok.bas` | Sanntidssøk i kunnskapsbase og saker med debounce. |
| `ModDashboard.bas` | Oppdaterer dashboard-hjelpetabeller og diagrammer. |
| `ModTidsstempel.bas` | Auto-tidsstempel i Saker via `Worksheet_Change`. |
| `build_donor.py` | Bygger donor `.xlsm` via Excel COM/pywin32 og ekstraherer `vbaProject.bin`. |
| `vbaProject.bin` | Ekstrahert VBA-binær, klar til injeksjon i `.xlsm` (T8). |
| `donor.xlsm` | Midlertidig donor-fil (genereres, skal ikke committes). |

## Forutsetninger

- Python 3.9+ med `openpyxl` og `pywin32`:
  ```powershell
  pip install openpyxl pywin32
  ```
- Microsoft Excel 2007 eller nyere.
- Skriptet setter selv registry-nøkkelen
  `HKCU\Software\Microsoft\Office\16.0\Excel\Security\AccessVBOM = 1`
  for programmatisk tilgang til VB-prosjektet.

## Bygge på nytt

Når du har endret en eller flere `.bas`-filer:

```powershell
python vba/build_donor.py
```

Skriptet gjør følgende:

1. Oppretter en tom makroaktivert arbeidsbok.
2. Lager nøyaktig 7 ark i kontrakt-rekkefølge:
   Dashboard, Saker, Søk KB, Søk saker, Kunnskapsbase, Oppsett, Instruksjoner.
3. Sikrer at arkenes codeNames er `Sheet1`..`Sheet7` (slik at binæren matcher
   den genererte `.xlsm`-en i T8).
4. Importerer de fire standardmodulene.
5. Legger hendelsesbehandlere i arkmodulene og `ThisWorkbook`:
   - `Sheet2` (Saker): `Worksheet_Change` → `ModTidsstempel.TidsstempelEndring`
   - `Sheet3` (Søk KB): `Worksheet_Change` → `ModSok.HåndterSøkKBEndring`
   - `Sheet4` (Søk saker): `Worksheet_Change` → `ModSok.HåndterSøkSakerEndring`
   - `ThisWorkbook`: `Workbook_SheetActivate` → `ModDashboard.OppdaterDashboard`
6. Lagrer donor-filen `vba/donor.xlsm`.
7. Ekstraherer `xl/vbaProject.bin` til `vba/vbaProject.bin`.
8. Verifiserer at binæren finnes, har innhold, og at modulene og
   hendelsesbehandlerne er på plass.

## Hva som committes

- `vba/*.bas` — VBA-kildekoden.
- `vba/build_donor.py` — gjenbrukbart bygge-skript.
- `vba/vbaProject.bin` — den ekstraherte binæren (bygge-artefakt).

`vba/donor.xlsm` genereres av skriptet og skal **ikke** committes; den er kun
et mellomledd for å produsere `vbaProject.bin`.

## Se også

- `kontrakt.py` — ark-rekkefølge og codeName-mapping.
- `.sisyphus/evidence/task-7-donor.txt` — siste bygge-logg.
- T3-oppskrift i `.sisyphus/evidence/spike/vba/README.md` — generell
  injeksjonsoppskrift for `vbaProject.bin`.
