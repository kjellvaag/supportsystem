# Supportsystem i Excel

Et komplett supportsystem for driftsavdelinger, bygget i ren Excel (uten makroer).
Systemet dekker ticket-håndtering, kunnskapsbase (FAQ), automatisk tildeling av
saker og SLA-styring – med søk i tidligere løsninger.

Systemet leveres i flere filvarianter: en makrofri versjon (`.xlsx`) som er
hoveddokumentasjonen nedenfor, en makroaktivert Excel-versjon (`.xlsm`) med
sanntidssøk og dynamisk dashboard, og en makroaktivert versjon for LibreOffice
og OpenOffice (`.ods`) der søk og tidsstempel styres fra knapper. Se «Hvilken
fil skal dere bruke?» for å velge riktig fil.

## Innhold i mappen

| Fil | Beskrivelse |
|---|---|
| `Supportsystem.xlsx` | Selve supportsystemet – dette er filen dere bruker |
| `lag_supportsystem.py` | Skriptet som genererer Excel-filen |
| `backup.py` | Backup-logikk (enkeltkjøring eller loop) |
| `backup_service.py` | Windows-tjeneste som kjører backup-loop |
| `backup_config.json` | Backup-innstillinger (intervall, mål, rotasjon) |
| `installer_tjeneste.bat` / `avinstaller_tjeneste.bat` | Install/fjern Windows-tjeneste (krever administrator) |
| `installer_oppgave.bat` | Opprett planlagt oppgave (kjører som din bruker, ingen admin) |
| `backups\` | Mappe med backup-kopier og backup.log |
| `README.md` | Denne filen |

## Krav

- **Excel 2007 eller nyere**, **LibreOffice Calc** eller **Apache OpenOffice Calc**.
  Systemet bruker kun standardfunksjoner (`VLOOKUP`, `INDEX`, `MATCH`, `COUNTIFS`,
  `OFFSET`) – ingen `FILTER`, `HSTACK` eller andre M365-spesifikke funksjoner.
- **Python 3.9+** med `openpyxl` – kun hvis du vil regenerere eller endre malen:
  ```
  pip install openpyxl
  ```
- **Makroversjonen** (`Supportsystem.xlsm`) krever **Excel 2007 eller nyere**
  med makroer aktivert og klarert. Den makrofrie `Supportsystem.xlsx` fungerer i
  Excel, LibreOffice Calc og Apache OpenOffice Calc uten noe oppsett. Se
  «Hvilken fil skal dere bruke?» for en samlet oversikt.

### Viktig for LibreOffice / OpenOffice

Punktlisten under gjelder den makrofrie `Supportsystem.xlsx` når den brukes i
LibreOffice/OpenOffice. Makroversjonen `Supportsystem.ods` er beskrevet under
«LibreOffice og OpenOffice (.ods) med makroer» og trenger ikke iterative
referanser, der settes tidsstempelet av en makro.

Auto-tidsstempelet i Saker (kolonne B) bruker sirkulær referanse og krever at
**iterative referanser** er slått på:
- LibreOffice: Verktøy > Innstillinger > LibreOffice Calc > Beregn > Iterative referanser
- OpenOffice: Verktøy > Innstillinger > OpenOffice.org Calc > Beregn > Iterasjoner

Statusavkryssingen på **Søk saker** bruker klassiske Excel-skjemakontroller
(avkryssingsbokser) som injiseres i filen av generatoren. Standard er kun
**Løst** og **Lukket** avkrysset. De tilsvarende SANN/USANN-verdiene står i
skjult kolonne AB og kan endres direkte der (SANN/USANN) med samme effekt på
søket – nyttig i LibreOffice/OpenOffice, som kan vise boksene som grå
elementer avhengig av versjon.

## Generere filen på nytt

```
python lag_supportsystem.py        # med eksempelsaker og KB-artikler (standard)
python lag_supportsystem.py ren    # tomt system, klart til bruk
```

Filen `Supportsystem.xlsx` må være lukket i Excel mens skriptet kjører.
Begge variantene inneholder alle formler, lister, fargekoder, dashboard og
instruksjoner – eneste forskjellen er om eksempeldataene (10 saker og 6
KB-artikler) er fylt inn.

Skriptet kan også generere makroversjonene:

```
python lag_supportsystem.py --target excel    # Supportsystem.xlsm (Excel med VBA-makroer)
python lag_supportsystem.py --target calc     # Supportsystem.ods (LibreOffice/OpenOffice med Basic-makroer)
```

Uten `--target` genereres den makrofrie `Supportsystem.xlsx` som før. Også
`Supportsystem.xlsm` må være lukket i Excel mens skriptet kjører. Se
«Hvilken fil skal dere bruke?» for hva som kreves av makroversjonene.

> Tips: Bruk standardversjonen først for å se hvordan systemet er ment å brukes,
> og generer deretter en `ren` versjon når dere skal i produksjon.

## Hvilken fil skal dere bruke?

Systemet finnes i tre varianter med samme arkstruktur og innhold. Forskjellen
ligger i hvordan tre funksjoner er løst: tidsstempel, søk og dashboard.

| Fil | Program | Makroer | Passer for |
|---|---|---|---|
| `Supportsystem.xlsx` | Excel 2007+, LibreOffice Calc, Apache OpenOffice Calc | Nei (fallback) | Alle miljøer, også blandede. Null oppsett |
| `Supportsystem.xlsm` | Excel 2007+ | Ja (VBA) | Avdelinger som bare bruker Excel og kan klarere makroer |
| `Supportsystem.ods` | LibreOffice / OpenOffice | Ja (Basic) | Avdelinger som bruker LibreOffice/OpenOffice og kan klarere makroer |

### Makrofri `.xlsx` (fallback)

Dette er originalversjonen uten makroer, og den dere får med `python
lag_supportsystem.py` uten `--target`. Alt fungerer med formler alene:
tidsstemplet i Saker bruker sirkulær referanse med iterativ beregning, søket
bruker skjulte hjelpekolonner med INDEX/MATCH, og dashboardet oppdateres når
arket beregnes på nytt. Filen krever verken makroer eller klarering og kan
deles fritt, også til LibreOffice- og OpenOffice-brukere.

### Excel (`.xlsm`) med makroer

`python lag_supportsystem.py --target excel` genererer `Supportsystem.xlsm`, en
makroaktivert variant for Excel. Innholdet i arkene er det samme som i
fallbacken, men tre ting fungerer annerledes når makroene er på:

- **Sanntidssøk**: «Søk KB» og «Søk saker» oppdaterer resultatene mens dere
  skriver. Treffene sorteres etter relevans (tittel, nøkkelord, problem og
  løsning vektes ulikt), og inntil to tegn skrivefeil tolereres. Treffordene
  markeres med farge i resultatradene.
- **Dynamisk dashboard**: diagrammene på Dashboard oppdateres straks dataene
  endres, uten at dere må trykke F9 eller lagre og åpne filen på nytt.
- **Makrobasert tidsstempel**: «Opprettet» i Saker settes og låses av en makro
  når en ny sak fylles inn. Makroversjonen trenger derfor **ikke** iterativ
  beregning, i motsetning til den makrofrie fallbacken.

Funksjonene over virker bare når makroene er **aktivert og klarert**. Hvis
makroene er av, oppfører filen seg omtrent som fallbacken, men søket og
dashboardet blir statiske og tidsstempelet må settes manuelt.

#### Slik aktiverer du makroer i Excel

1. **Fjern eventuell blokkering av filen**: høyreklikk `Supportsystem.xlsm` i
   Filutforsker, velg **Egenskaper**, og under **Generelt** setter du kryss for
   **Opphev blokkering** hvis feltet finnes. Velg **OK**.
2. **Legg filen på en klarert plassering**: for eksempel en lokal mappe, eller
   en nettverksmappe som er lagt til under Fil > Alternativer >
   Sikkerhetssenter > Innstillinger for sikkerhetssenteret > **Klarerte
   plasseringer**.
3. **Slå på makroer** i sikkerhetssenteret: Fil > Alternativer >
   Sikkerhetssenter > Innstillinger for sikkerhetssenteret > **Innstillinger
   for makroer** > **Aktiver alle makroer**. Alternativt kan dere velge
   **Deaktiver alle makroer med varsling** og bekrefte varslingen hver gang
   filen åpnes, men da må varslingen ikke overses.
4. Åpne `Supportsystem.xlsm` på nytt. Skriv et par bokstaver i søkefeltet på
   «Søk KB» eller «Søk saker» og kontroller at resultatene kommer opp mens dere
   skriver. Først da er makroene klare til bruk.

Excel 2007 eller nyere kreves, og makroene må være klarert. Hvis sikkerhetssenteret
blokkerer makroer uten å spørre, virker ingenting av det over.

### LibreOffice og OpenOffice (`.ods`) med makroer

`python lag_supportsystem.py --target calc` genererer `Supportsystem.ods`, en
Basic-makrovariant for LibreOffice Calc og Apache OpenOffice Calc. Innholdet i
arkene er det samme som i de andre variantene, men tre funksjoner er løst med
Basic-makroer:

- **Søk er knapp-triggeret**: LibreOffice/OpenOffice har ingen
  `Worksheet_Change`-hendelse, så søket starter ved å klikke på knappen
  **«SØK I KB →»** på «Søk KB» eller **«SØK I SAKER →»** på «Søk saker».
  Skriv søkeordet i det gule feltet først, og klikk deretter på knappen.
  Treffene rangeres etter relevans (tittel, nøkkelord, problem og løsning
  vektes ulikt), inntil to tegn skrivefeil tolereres, og treffordene markeres
  med farge. I Excel-versjonen (`.xlsm`) er søket direkte og oppdaterer
  resultatene mens dere skriver. `.ods`-versjonen er knappestyrt. En
  best-effort-dokumenthendelse kan også utløse søk under skriving, men den
  skal dere ikke regne med, knappen er den garanterte måten.
- **Tidsstempel er knapp-triggeret**: «Opprettet» i Saker (kolonne B) settes
  av makroen `StemplTidsstempel` fra en knapp på Saker-arket, igjen fordi
  LibreOffice/OpenOffice mangler `Worksheet_Change`. Fyll ut saken og klikk på
  knappen. Makroen stempler alle rader som har Innmelder eller Tittel utfylt
  og fortsatt tom «Opprettet», og overskriver ikke eksisterende tidsstempler.
- **Dashboard**: diagrammene på Dashboard oppdateres fra knappen på
  Dashboard-arket. Knappen kjører makroen `OppdaterDashboard`, som regner om
  hjelpetabellene diagrammene leser fra.

Makroversjonen trenger **ikke iterative referanser**: makroen erstatter den
sirkulære referansen som fallbacken (`.xlsx`) bruker for tidsstempelet. Blir
`.ods`-filen åpnet med makroer skrudd av, virker verken søk, tidsstempel eller
dashboard-oppdatering; bruk da den makrofrie `Supportsystem.xlsx` med
iterative referanser slått på.

Statusfilteret på «Søk saker» leser den skjulte speiltabellen `AA5:AB14`, der
kolonne AA inneholder statusnavnene og kolonne AB SANN/USANN-verdiene.
Avkryssingsboksene som generatoren legger inn, er koblet til kolonne AB og kan
vises som grå elementer i LibreOffice/OpenOffice avhengig av versjon. Virker
boksene ikke, kan dere endre `AB5:AB14` direkte mellom SANN og USANN med samme
effekt på søket.

#### Slik aktiverer du makroer i LibreOffice og OpenOffice

1. Åpne `Supportsystem.ods`.
2. Sett makrosikkerheten til et nivå som spør eller tillater:
   LibreOffice: **Verktøy > Innstillinger > LibreOffice > Sikkerhet >
   Makrosikkerhet**, og velg **Medium** slik at dere blir spurt om å tillate
   makroer når filen åpnes. Apache OpenOffice har tilsvarende valg under
   **Verktøy > Innstillinger > Apache OpenOffice > Sikkerhet >
   Makrosikkerhet**.
3. Åpne `Supportsystem.ods` på nytt og bekreft at makroene skal kjøres når
   dere blir spurt. Alternativt kan dere legge mappen som filen ligger i, til
   de klarerte plasseringene i samme dialogvindu, så kjører makroene uten
   spørsmål.
4. Kontroller at makroene virker: skriv et søkeord på «Søk KB» og klikk på
   **«SØK I KB →»**. Kommer resultatene opp, er makroene klare til bruk.

## Arkfanene i systemet

| Ark | Funksjon |
|---|---|
| **Dashboard** | KPI-er (åpne saker, forfalte frister, snitt løsningstid) og diagrammer – oppdateres automatisk. Kategori- og ansvarlig-diagrammene viser **Topp 10** saker sortert synkende (kun kategorier/ansvarlige som faktisk har saker) |
| **Saker** | Saksloggen. Hver rad er en sak med unikt saksnr, auto-tidsstempel, auto-tildeling og SLA-frister |
| **Søk KB** | Søk i kunnskapsbasen (FAQ) – skriv søkeord i det gule feltet, evt. avgrens med kategori |
| **Søk saker** | Søk i tidligere saker (tittel, beskrivelse og løsning) – avkryssingsbokser øverst styrer hvilke statuser som tas med (standard: Løst og Lukket) |

Begge søk ignorerer store/små bokstaver og bindestreker – «wifi» finner «Wi-Fi».
| **Kunnskapsbase** | FAQ-artikler med problem, løsning og nøkkelord |
| **Oppsett** | Team, kategorier, SLA-tider, statuser og kanaler – tilpass til din avdeling |
| **Instruksjoner** | Fullstendig brukerveiledning (på norsk) |

## Slik dekkes funksjonene

- **Ticket-håndtering** – unikt saksnr (`SAK-0001` …) genereres automatisk;
  kanal (e-post/telefon/chat …) dokumenteres per sak.
- **Automatisk tidsstempel** – «Opprettet» fylles med dato/klokkeslett idet
  Innmelder eller Tittel skrives, og låses. Kan overskrives manuelt.
  (Teknikk: sirkulær referanse + iterativ beregning, ingen makroer.)
- **Automatisk tildeling** – «Auto-ansvarlig» settes ut fra kategori
  (koblingen vedlikeholdes i Oppsett-arket) og kan overstyres i «Tildelt til».
- **SLA-styring** – frister beregnes fra prioritet (Kritisk 1t/4t, Høy 4t/8t,
  Medium 8t/24t, Lav 24t/72t). Fargekoder: rød = FORFALT, oransje = løst/påbegynt
  for sent, grønn = OK. Forfalte saker telles på Dashboard.
- **Kunnskapsbase (FAQ)** – auto KB-ID, nøkkelord for søk, telling av bruk og
  vurdering. Marker saker med «Til KB» = Ja for å bygge basen over tid.
- **Søk i løsninger** – fritekstsøk i både kunnskapsbase og alle løste saker.

## Tilpasning

- **Løpende drift** (team, kategorier, SLA-tider, statuser, kanaler):
  endre direkte i **Oppsett**-arket i Excel.
- **Backup-innstillinger**: rediger `backup_config.json` (intervall i minutter,
  lokal mappe, nettverkssti, antall kopier å beholde, loggfil).

## Automatisk backup

Backupen lager tidsstemplet kopi av `Supportsystem.xlsx` til en lokal
undermappe (`backups`) og/eller en nettverksdisk, og sletter eldste kopier
automatisk. Alle hendelser logges i `backups\backup.log`.

### Valg av kjøremodus

| Modus | Hvordan | Krever admin? |
|---|---|---|
| **Manuelt** | `python backup.py` → én kopi nå | Nei |
| **Planlagt oppgave** | `installer_oppgave.bat` → kjører loop ved oppstart/pålogging | Nei |
| **Windows-tjeneste** | `installer_tjeneste.bat` → kjører uansett om noen er logget på | Ja (én gang) |

### Tjeneste

```
pip install pywin32            # enkelt kjøring
installer_tjeneste.bat         # install + start (som administrator)
avinstaller_tjeneste.bat       # stopp + fjern (som administrator)
```

Tjenesten heter **SupportsystemBackup** og oppdateres som `backup.py` på nytt
uten avinstallering – bare restart tjenesten hvis du endret config/py-kode.

### Planlagt oppgave

`installer_oppgave.bat` oppretter oppgaven **SupportsystemBackup** som kjører
`backup.py loop` ved oppstart, alltid som din bruker (dvs. tilgang til
nettverksdisk/filen på dine premisser). Test den med:
`schtasks /run /tn "SupportsystemBackup"`

### Konfigurasjon (backup_config.json)

```json
{
  "kildefil": "Supportsystem.xlsx",
  "lokal_backupmappe": "backups",
  "nettverk_backupmappe": "",    // f.eks. "\\\\server\\felles\\backup"
  "intervall_minutter": 60,
  "behold_antall": 10,
  "loggfil": "backups\\backup.log"
}
```
- **nettverk_backupmappe** – la tom for kun lokal backup; sett UNC-sti for å ta
  andre kopi dit (krever at kontoen som kjører tjenesten har skrivetilgang).
- **behold_antall** – hvor mange nyeste kopier som beholdes per mappe; 0 =
  behold alle.

Rotasjonen skjer hver gang backup kjøres, ikke ved oppstart – en pauset tjeneste
sletter ikke historikk.
- **Permanent endring av malen**: rediger listene øverst i
  `lag_supportsystem.py` og kjør skriptet på nytt. Eksempel – legge til en
  kategori: legg til `("Sikkerhet", "Anne Berg")` i `kategorier`-listen.
- Skriptet har plass til **500 saker** og **300 KB-artikler** (`MAX_S`/`MAX_K`).

## Feilsøking

| Problem | Løsning |
|---|---|
| Excel ber om å «reparere» filen | Filen er generert med gammelt skript – kjør skriptet på nytt (funksjonsprefikser er fikset) |
| `#SPILL!` i Søk-arket | Kan ikke lenger forekomme – søket bruker INDEX/MATCH, ikke FILTER |
| Søket virker ikke | Sjekk at skjulte hjelpekolonner ikke er slettet (Saker!Y, Kunnskapsbase!L) |
| Auto-tidsstempel feiler i LO/OO | Slå på iterative referanser (se over) |
| Auto-tidsstempel står stille | Trykk F9 for å beregne på nytt; iterativ beregning skal være slått på automatisk i filen |
| «Opprettet» ble overskrevet ved en feil | Kopier formelen fra en hvilken som helst tom rad under |
| `Permission denied` ved kjøring | Lukk `Supportsystem.xlsx` i Excel og kjør igjen |
| Vil bekrefte at filene åpnes rent i LibreOffice/OpenOffice | Kjør `python valider_soffice.py` – konverterer `.ods` og `.xlsx` headless frem og tilbake og rapporterer PASS/FAIL (krever LibreOffice; sti overstyres med `SOFFICE`) |
| Søk reagerer ikke mens du skriver i `.xlsm` | Makroene er ikke aktivert eller klarert, se «Slik aktiverer du makroer i Excel» |
| Søk, tidsstempel eller dashboard virker ikke i `.ods` | Makroene er ikke aktivert, eller knappen er ikke klikket. Se «Slik aktiverer du makroer i LibreOffice og OpenOffice» |
| Tjenesten starter ikke | Se Windows Event Log + backups\backup.log; sjekk at `pip install pywin32` er kjørt |
| Nettverksbackup feiler | Sjekk at kontoen (tjeneste/oppgave) har skrivetilgang til UNC-stien |
