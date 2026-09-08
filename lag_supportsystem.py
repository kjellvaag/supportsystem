# -*- coding: utf-8 -*-
"""
Genererer et komplett supportsystem i Excel for en driftsavdeling.

Innhold i arbeidsboken:
  - Dashboard      : KPI-er og diagrammer (oppdateres automatisk)
  - Saker          : Sakslogg med unikt saksnr, auto-tildeling, SLA-frister
  - Søk KB         : Søk i kunnskapsbasen (FAQ)
  - Søk saker      : Søk i tidligere løste saker
  - Kunnskapsbase  : FAQ/artikler med løsninger og nøkkelord
  - Oppsett        : Team, kategorier, SLA-tider, statuser, kanaler
  - Instruksjoner  : Brukerveiledning

Kjør:  python lag_supportsystem.py                 (makrofri .xlsx med eksempeldata)
       python lag_supportsystem.py ren             (makrofri .xlsx, tomt)
       python lag_supportsystem.py --target excel  (Excel .xlsm med VBA-makroer)
Gir:   Supportsystem.xlsx / Supportsystem.xlsm i samme mappe
"""
import os
import shutil
import subprocess
import sys
import tempfile
import time

import build_common
import kontrakt

# Overridbar sti til LibreOffice (brukt av --target calc).
SOFFICE = os.environ.get("SOFFICE", r"C:\Program Files\LibreOffice\program\soffice.exe")

args = sys.argv[1:]

# Med eksempeldata som standard; "ren" som argument genererer et tomt system.
MED_EKSEMPLER = not any(a.lower() in ("ren", "clean", "--ren", "--clean") for a in args)

# Target-switch: excel = VBA-makroer (.xlsm), calc = LO/OO Basic (.ods, T11).
target = None
if "--target" in args:
    ti = args.index("--target")
    if ti + 1 < len(args):
        target = args[ti + 1].lower()

if target == "excel":
    OUT = "Supportsystem.xlsm"
    profile = dict(kontrakt.EXCEL_PROFILE)
elif target == "calc":
    OUT = "Supportsystem.ods"
    profile = dict(kontrakt.CALC_PROFILE)
else:
    OUT = "Supportsystem.xlsx"
    profile = dict(kontrakt.FALLBACK_PROFILE)

profile["med_eksempler"] = MED_EKSEMPLER

wb = build_common.build_skeleton(profile)

# Boks-layout: verdier er tatt rett fra filen Excel selv laget.
_BOKS_DEF = [
    (7,  1, 1001, 21.0, 2, 31750, "8, 58, 2, 5"),
    (10, 1, 1177, 21.0, 2, 31750, "11, 58, 2, 5"),
    (7,  2, 1001, 35.5, 3, 31750, "8, 58, 3, 5"),
    (10, 2, 1177, 35.5, 3, 31750, "11, 58, 3, 5"),
    (7,  3, 1001, 50.0, 3, 215900, "8, 58, 3, 34"),
    (10, 3, 1177, 50.0, 3, 215900, "11, 58, 3, 34"),
]

avkryssingsbokser = []
for i, st in enumerate(build_common.statuser[:6]):
    kol, rad, left_pt, top_pt, til_rad, roff, vml_rest = _BOKS_DEF[i]
    avkryssingsbokser.append({
        "tekst": st,
        "kol": kol,
        "rad": rad,
        "left_pt": left_pt,
        "top_pt": top_pt,
        "til_rad": til_rad,
        "row_off_to": roff,
        "vml_rest": vml_rest,
        "lenke": f"$AB${5 + i}",
        "avkrysset": st in ("Løst", "Lukket"),
    })

if target == "calc":
    # LO/OO-sporet: .xlsx-skjelett -> soffice til .ods -> injiser Basic.
    if not os.path.exists(SOFFICE):
        raise FileNotFoundError(f"soffice mangler: {SOFFICE}")

    with tempfile.TemporaryDirectory(prefix="supportsystem_calc_") as tmp:
        tmp_xlsx = os.path.join(tmp, "calc_skall.xlsx")
        wb.save(tmp_xlsx)
        build_common.injiser_avkryssingsbokser(tmp_xlsx, "Søk saker", avkryssingsbokser)

        cmd = [SOFFICE, "--headless", "--convert-to", "ods", "--outdir", tmp, tmp_xlsx]
        result = subprocess.run(cmd, capture_output=True, text=True)
        if result.returncode != 0:
            raise RuntimeError(
                f"soffice-konvertering til .ods feilet (exit {result.returncode}):\n"
                f"stdout: {result.stdout}\nstderr: {result.stderr}"
            )

        tmp_ods = os.path.join(tmp, "calc_skall.ods")
        for _ in range(60):
            if os.path.exists(tmp_ods):
                break
            time.sleep(0.5)
        else:
            raise RuntimeError(f"soffice returnerte OK, men {tmp_ods} ble ikke produsert")

        moduler = {
            "Module1": "basic/ModFelles.bas",
            "Module2": "basic/ModTidsstempel.bas",
            "Module3": "basic/ModSok.bas",
            "Module4": "basic/ModDashboard.bas",
        }
        build_common.injiser_basic(tmp_ods, moduler)
        shutil.move(tmp_ods, OUT)
else:
    wb.save(OUT)
    build_common.injiser_avkryssingsbokser(OUT, "Søk saker", avkryssingsbokser)
    if target == "excel":
        build_common.injiser_vba_prosjekt(OUT, "vba/vbaProject.bin")

modus = "med eksempeldata" if MED_EKSEMPLER else "tomt (uten eksempeldata)"
print(f"OK – {OUT} generert, {modus}")
