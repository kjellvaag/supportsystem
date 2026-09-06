# Supportsystem i Excel

Et komplett supportsystem for driftsavdelinger, bygget i ren Excel (uten makroer).
Systemet dekker ticket-håndtering, kunnskapsbase (FAQ), automatisk tildeling av
saker og SLA-styring – med søk i tidligere løsninger.

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

### Viktig for LibreOffice / OpenOffice

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

> Tips: Bruk standardversjonen først for å se hvordan systemet er ment å brukes,
> og generer deretter en `ren` versjon når dere skal i produksjon.

## Arkfanene i systemet

| Ark | Funksjon |
|---|---|
| **Dashboard** | KPI-er (åpne saker, forfalte frister, snitt løsningstid) og diagrammer – oppdateres automatisk |
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
| Tjenesten starter ikke | Se Windows Event Log + backups\backup.log; sjekk at `pip install pywin32` er kjørt |
| Nettverksbackup feiler | Sjekk at kontoen (tjeneste/oppgave) har skrivetilgang til UNC-stien |
