# -*- coding: utf-8 -*-
"""Validering av Supportsystem-filene i LibreOffice (headless).

Konverterer Supportsystem.ods -> .xlsx og Supportsystem.xlsx -> .ods med
soffice --headless, og bekrefter at konverteringen lykkes (exit-kode 0) og
at utdatafilen faktisk blir produsert. Konverteringen skjer i en midlertidig
mappe - kildefilene overskrives aldri.

Bruksomrade: kjor etter generering for a bekrefte at filene apnes rent i
LibreOffice/OpenOffice (round-trip uten korrupsjon).

Krav: LibreOffice installert. Sti til soffice kan overstyres med
miljovariabelen SOFFICE (default: C:\\Program Files\\LibreOffice\\program\\soffice.exe).
"""

import os
import shutil
import subprocess
import sys
import tempfile
import time

SOFFICE = os.environ.get(
    "SOFFICE", r"C:\Program Files\LibreOffice\program\soffice.exe"
)

# Kilder som skal valideres: (filnavn, konverteringsmal, forventet utdata-endelse)
KILDER = [
    ("Supportsystem.ods", "xlsx", ".xlsx"),
    ("Supportsystem.xlsx", "ods", ".ods"),
]

POLL_INTERVALL = 0.5  # sekunder
POLL_MAKS = 60  # sekunder totalt


def vent_pa_fil(sti, tidsfrist):
    """Poller til filen finnes eller tidsfristen gar ut (soffice er en launcher)."""
    while time.time() < tidsfrist:
        if os.path.exists(sti) and os.path.getsize(sti) > 0:
            return True
        time.sleep(POLL_INTERVALL)
    return False


def konverter(soffice, kilde, mal, utdata_navn, utdir):
    """Kjorer soffice headless-konvertering og returnerer (ok, melding)."""
    if not os.path.exists(soffice):
        return False, "soffice ikke funnet: %s" % soffice

    utdata_sti = os.path.join(utdir, utdata_navn)
    if os.path.exists(utdata_sti):
        os.remove(utdata_sti)

    kommando = [
        soffice,
        "--headless",
        "--convert-to", mal,
        "--outdir", utdir,
        kilde,
    ]
    try:
        prosess = subprocess.run(
            kommando,
            capture_output=True,
            text=True,
            timeout=POLL_MAKS + 30,
        )
    except subprocess.TimeoutExpired:
        return False, "soffice tidsavbrudd ved konvertering av %s" % kilde

    tidsfrist = time.time() + POLL_MAKS
    if not vent_pa_fil(utdata_sti, tidsfrist):
        detaljer = (prosess.stdout or "").strip() or (prosess.stderr or "").strip()
        return False, "utdatafil %s ble ikke produsert (exit %s). %s" % (
            utdata_navn, prosess.returncode, detaljer
        )

    if prosess.returncode != 0:
        return False, "soffice returnerte exit-kode %s for %s" % (
            prosess.returncode, kilde
        )

    return True, "OK (exit %s, %d bytes)" % (
        prosess.returncode, os.path.getsize(utdata_sti)
    )


def main():
    mangler = [navn for navn, _, _ in KILDER if not os.path.exists(navn)]
    if mangler:
        print("FEIL: manglende kildefil(er): %s" % ", ".join(mangler))
        print("Kjor generatoren forst (python lag_supportsystem.py [--target calc]).")
        return 1

    if not os.path.exists(SOFFICE):
        print("FEIL: soffice ikke funnet pa %s" % SOFFICE)
        print("Sett SOFFICE-miljovariabelen til riktig sti hvis LibreOffice er installert annet sted.")
        return 1

    utdir = tempfile.mkdtemp(prefix="valider_soffice_")
    resultater = []
    try:
        for kilde, mal, endelse in KILDER:
            utdata_navn = os.path.splitext(os.path.basename(kilde))[0] + endelse
            print("Konverterer %s -> %s ..." % (kilde, utdata_navn))
            ok, melding = konverter(SOFFICE, kilde, mal, utdata_navn, utdir)
            resultater.append((kilde, ok, melding))
            print("  %s: %s" % ("PASS" if ok else "FAIL", melding))
    finally:
        shutil.rmtree(utdir, ignore_errors=True)

    print()
    feil = [r for r in resultater if not r[1]]
    if feil:
        print("OPPSUMMERING: FAIL (%d av %d konverteringer feilet)" % (
            len(feil), len(resultater)
        ))
        for kilde, _, melding in feil:
            print("  - %s: %s" % (kilde, melding))
        return 1

    print("OPPSUMMERING: PASS (alle %d round-trip-konverteringer lyktes)" % len(resultater))
    return 0


if __name__ == "__main__":
    sys.exit(main())