T12 — LO/OO README + makro-sikkerhetsveiledning
================================================
Status: OK (2026-09-07)

Endringer i README.md (kun additivt + én intro-presisering):
1. Intro (linje 7-11): "en planlagt versjon for LibreOffice og OpenOffice
   (.ods)" erstattet med "en makroaktivert versjon for LibreOffice og
   OpenOffice (.ods) der søk og tidsstempel styres fra knapper" — fjerner
   siste gjenværende "planlagt"-referanse til .ods (T11 fjernet
   "kommende"/"ikke implementert" i selve seksjonen).
2. "### Viktig for LibreOffice / OpenOffice" (under Krav): ny innledning som
   avgrenser punktlisten til den makrofrie .xlsx i LO/OO og peker på at
   .ods-makroversjonen ikke trenger iterative referanser.
3. "### LibreOffice og OpenOffice (.ods) med makroer" omskrevet og utvidet:
   - Søk er KNAPP-triggeret: klikk «SØK I KB →» / «SØK I SAKER →» etter å ha
     skrevet søkeordet; LO/OO har ingen Worksheet_Change. .xlsm er direkte
     (live), .ods er knappestyrt; best-effort-dokumenthendelse dokumentert
     men eksplisitt ikke å regne med.
   - Tidsstempel er KNAPP-triggeret via makroen StemplTidsstempel på
     Saker-arket; stempler rader med Innmelder/Tittel utfylt og tom
     «Opprettet», overskriver ikke eksisterende stempler.
   - Dashboard oppdateres fra knapp på Dashboard-arket (OppdaterDashboard).
   - Eksplisitt: makroversjonen trenger IKKE iterative referanser; fallback
     .xlsx gjør det fortsatt. Åpnet med makroer av = søk/tidsstempel/dashboard
     virker ikke, bruk da fallbacken.
   - Statusfilter: speiltabell AA5:AB14 (AA = statusnavn, AB = SANN/USANN);
     AB5:AB14 kan redigeres direkte hvis avkryssingsboksene vises grå.
   - Ny underseksjon "#### Slik aktiverer du makroer i LibreOffice og
     OpenOffice": Verktøy > Innstillinger > LibreOffice > Sikkerhet >
     Makrosikkerhet (Medium = spør), OpenOffice tilsvarende; alternativt
     klarerte plasseringer; verifiseringssteg med «SØK I KB →».
4. Feilsøking: ny rad "Søk, tidsstempel eller dashboard virker ikke i .ods".

Verifisering:
- Ingen gjenværende "planlagt versjon for LibreOffice"/"kommende"/
  "ikke implementert" i README (Select-String).
- Ingen en/em-bindestreker i ny tekst (linje 43-46 og 147-200 sjekket).
- Excel- og fallback-seksjonene, backup- og feilsøkingsdokumentasjonen er
  ikke fjernet eller omskrevet.
- Faktasjekk mot kode: basic/ModTidsstempel.bas (StemplTidsstempel, låser
  eksisterende stempler), basic/ModSok.bas (leser AA5:AB14, AA=statusnavn,
  AB=SANN/USANN), kontrakt.SOK_SAK_STATUS = 'Søk saker'!$AA$5:$AB$14,
  build_common D4/D4-knappceller «SØK I KB →»/«SØK I SAKER →».
- Merknad: generert .ods inneholder Basic-modulene (Module1-4.xml) men ingen
  makro-tilknyttede kontroller ennå (<office:scripts/> tom); knappetilordning
  er dokumentert i basic/README.md og følger planens design
  (knapp-triggeret søk/tidsstempel). README beskriver tiltenkt bruk slik
  planen og arkets egne prompt-tekster («trykk «SØK I KB →»») angir.