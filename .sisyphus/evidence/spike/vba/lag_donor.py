#!/usr/bin/env python3
"""T3-spike: lag en donor .xlsm med Excel COM og ekstraher vbaProject.bin.

Bruker pywin32 fordi openpyxl ikke kan skrive makroer.  Donoren inneholder
kun `Sub SayHello()`; binæren brukes deretter av `injiser_vba.py`.
"""

import os
import sys
import zipfile

BASE = os.path.dirname(os.path.abspath(__file__))
EVIDENCE = os.path.join(BASE, "..", "..", "task-3-vba-spike.txt")
DONOR = os.path.join(BASE, "donor.xlsm")
BIN_OUT = os.path.join(BASE, "vbaProject.bin")


def logg(linje):
    print(linje)
    with open(EVIDENCE, "a", encoding="utf-8") as f:
        f.write(linje + "\n")


def lag_donor():
    import win32com.client as win32

    app = win32.Dispatch("Excel.Application")
    logg(f"Excel versjon: {app.Version}")
    logg(f"Excel bitness: {app.Bitness if hasattr(app, 'Bitness') else 'ukjent'}")

    wb = None
    try:
        app.Visible = False
        app.DisplayAlerts = False

        wb = app.Workbooks.Add()
        # 1 = vbext_ct_StdModule
        modul = wb.VBProject.VBComponents.Add(1)
        modul.Name = "SpikeModul"
        modul.CodeModule.AddFromString(
            'Sub SayHello()\n    MsgBox "VBA-injeksjon fungerer"\nEnd Sub\n'
        )

        # xlOpenXMLWorkbookMacroEnabled = 52
        wb.SaveAs(DONOR, FileFormat=52)
        logg(f"Donor lagret: {DONOR}")
    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        app.Quit()
        logg("Excel avsluttet etter donor-eksport")


def ekstraher_bin():
    with zipfile.ZipFile(DONOR, "r") as z:
        data = z.read("xl/vbaProject.bin")
    with open(BIN_OUT, "wb") as f:
        f.write(data)
    logg(f"Ekstraherte xl/vbaProject.bin -> {BIN_OUT} ({len(data)} bytes)")


if __name__ == "__main__":
    os.makedirs(os.path.dirname(EVIDENCE), exist_ok=True)
    logg("=== lag_donor.py start ===")
    lag_donor()
    ekstraher_bin()
    logg("=== lag_donor.py ferdig ===")
