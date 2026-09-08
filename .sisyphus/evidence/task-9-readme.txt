T9: Excel README + makro-sikkerhetsveiledning
Status: OK (2026-09-07)

Endringer i README.md (kun additive, ingen eksisterende tekst slettet/omskrevet):

1. Nytt intro-avsnitt etter første avsnitt:
   - Forklarer at systemet leveres i flere filvarianter (.xlsx makrofri,
     .xlsm makroaktivert Excel, .ods planlagt for LO/OO), henviser til
     «Hvilken fil skal dere bruke?».

2. Krav-seksjonen: ny bullet:
   - Makroversjonen (Supportsystem.xlsm) krever Excel 2007+ med makroer
     aktivert og klarert; makrofri .xlsx fungerer i Excel/LO/OO uten oppsett.

3. «Generere filen på nytt»: ny kodeblokk + avsnitt:
   - python lag_supportsystem.py --target excel  -> Supportsystem.xlsm
   - python lag_supportsystem.py --target calc   -> Supportsystem.ods (kommende)
   - Uten --target genereres makrofri .xlsx som før; .xlsm må også være lukket
     i Excel under kjøring.

4. Ny seksjon «## Hvilken fil skal dere bruke?» (plassert mellom «Generere
   filen på nytt» og «Arkfanene i systemet»):
   - Tre-fil-modell i tabell: .xlsx (fallback, makrofri), .xlsm (Excel med
     VBA-makroer), .ods (LO/OO med Basic-makroer, planlagt, ikke implementert).
   - ### Makrofri .xlsx (fallback): oppsummerer formelbaserte mekanismer
     (sirkulær referanse + iterativ beregning, INDEX/MATCH-hjelpekolonner).
   - ### Excel (.xlsm) med makroer: dokumenterer forskjeller vs fallback:
     sanntidssøk med relevansrangering og fuzzy (inntil 2 tegn) og
     trefford-markering; dynamisk dashboard; makrobasert tidsstempel som
     IKKE krever iterativ beregning. Uten makroer faller den tilbake til
     statisk søk/dashboard og manuelt tidsstempel.
   - #### Slik aktiverer du makroer i Excel: konkrete steg med menystier:
     (1) Egenskaper > Generelt > Opphev blokkering; (2) klarert plassering via
     Fil > Alternativer > Sikkerhetssenter > Klarerte plasseringer;
     (3) Innstillinger for makroer > Aktiver alle makroer (ev. Deaktiver alle
     makroer med varsling); (4) verifiser med tastetrykk i søkefelt.
     Krav: Excel 2007 eller nyere; makroer må være klarert.
   - ### LibreOffice og OpenOffice (.ods) med makroer (kommende): eksplisitt
     merket som planlagt/ikke implementert; LO/OO-brukere bruker inntil videre
     makrofri .xlsx med iterative referanser.

5. Feilsøking: ny rad:
   - «Søk reagerer ikke mens du skriver i .xlsm» -> makroer ikke aktivert/
     klarert, se «Slik aktiverer du makroer i Excel».

Merknad: Dokumentasjonen beskriver --target excel/calc slik planen spesifiserer;
switchen er ennå ikke implementert i lag_supportsystem.py (gjenstår T8/T11).
Når den lander bør README verifiseres mot faktisk CLI-oppførsel.

Verifisering:
- grep-bekreftet tilstedeværelse av nøkkelmarkører (--target excel/calc,
  «Excel (.xlsm) med makroer», «Opphev blokkering», «Aktiver alle makroer»,
  «Makrobasert tidsstempel», «kommende», «ikke iterativ»).
- Ingen en/em-bindestrek (U+2013/U+2014) i nye linjer (linje 7-148 sjekket).
- Eksisterende fallback-, backup- og feilsøkingsdokumentasjon er intakt
  (README gikk fra 165 til 255 linjer, alle endringer er tillegg).
