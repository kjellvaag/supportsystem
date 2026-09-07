#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Bygg VBA-donor for supportsystemet.

Lager en makroaktivert Excel-arbeidsbok (donor .xlsm) via Excel-COM/pywin32,
importerer de fire standardmodulene, legger hendelsesbehandlere i arkmodulene
og ThisWorkbook, verifiserer ark-codenavn, og ekstraherer xl/vbaProject.bin.

Bruk:
    python vba/build_donor.py

Krever:
    - pywin32 (pip install pywin32)
    - Microsoft Excel 2007+
    - AccessVBOM=1 (skriptet setter registry-nøkkelen selv)
"""

import os
import sys
import zipfile
import winreg

# Sikre at prosjektroten er på sys.path slik at kontrakt.py kan importeres
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
if ROOT not in sys.path:
    sys.path.insert(0, ROOT)

import kontrakt

VBA_DIR = os.path.join(ROOT, "vba")
DONOR_PATH = os.path.join(VBA_DIR, "donor.xlsm")
BIN_OUT = os.path.join(VBA_DIR, "vbaProject.bin")
EVIDENCE_DIR = os.path.join(ROOT, ".sisyphus", "evidence")
EVIDENCE_FILE = os.path.join(EVIDENCE_DIR, "task-7-donor.txt")

BAS_MODULES = [
    "ModFelles.bas",
    "ModSok.bas",
    "ModDashboard.bas",
    "ModTidsstempel.bas",
]

# Konstant fra VBA-ext: vbext_ct_StdModule = 1, vbext_ct_Document = 100
VBEXT_CT_STDMODULE = 1
VBEXT_CT_DOCUMENT = 100

# xlOpenXMLWorkbookMacroEnabled = 52
XL_OPENXML_WORKBOOK_MACROENABLED = 52

ACCESS_VBOM_KEY = r"Software\Microsoft\Office\16.0\Excel\Security"
ACCESS_VBOM_VALUE = "AccessVBOM"


def logg(linje):
    """Skriv til stdout og evidence-fil."""
    print(linje)
    os.makedirs(EVIDENCE_DIR, exist_ok=True)
    with open(EVIDENCE_FILE, "a", encoding="utf-8") as f:
        f.write(linje + "\n")


def sett_access_vbom():
    """Sett HKCU AccessVBOM=1 slik at Excel tillater programmatisk VBProject-tilgang."""
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, ACCESS_VBOM_KEY) as key:
            winreg.SetValueEx(key, ACCESS_VBOM_VALUE, 0, winreg.REG_DWORD, 1)
        logg("AccessVBOM satt til 1 i HKCU")
        return True
    except OSError as e:
        logg(f"ADVARSEL: kunne ikke sette AccessVBOM: {e}")
        return False


def les_access_vbom():
    """Les nåværende verdi av AccessVBOM (returner None hvis ikke satt)."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, ACCESS_VBOM_KEY, 0, winreg.KEY_READ) as key:
            value, _ = winreg.QueryValueEx(key, ACCESS_VBOM_VALUE)
            return value
    except (OSError, FileNotFoundError):
        return None


def finn_arkkomponent(vbproj, codename):
    """Returner VBComponent for ark med gitt codeName."""
    for comp in vbproj.VBComponents:
        if comp.Type == VBEXT_CT_DOCUMENT:
            # For dokumentkomponenter er Name property codenavnet
            if comp.Name == codename:
                return comp
    raise RuntimeError(f"Finner ikke arkmodul med codeName={codename}")


def finn_arkkomponent_etter_indeks(vbproj, indeks):
    """Returner det i-te ark-komponentet (1-basert, Sheet1=1).

    Filtrerer bort ThisWorkbook som også er type Document.
    """
    dok_comp = [
        c for c in vbproj.VBComponents
        if c.Type == VBEXT_CT_DOCUMENT and c.Name != "ThisWorkbook"
    ]
    if 1 <= indeks <= len(dok_comp):
        return dok_comp[indeks - 1]
    raise RuntimeError(f"Finner ikke ark-komponent {indeks}")


def bygg_donor():
    import win32com.client as win32

    app = None
    wb = None

    # Sikre AccessVBOM
    gammel = les_access_vbom()
    logg(f"AccessVBOM før: {gammel}")
    sett_access_vbom()

    try:
        app = win32.Dispatch("Excel.Application")
        logg(f"Excel versjon: {app.Version}")
        logg(f"Excel bitness: {app.Bitness if hasattr(app, 'Bitness') else 'ukjent'}")

        app.Visible = False
        app.DisplayAlerts = False

        wb = app.Workbooks.Add()

        # --- 1) Sikre nøyaktig 7 ark i kontrakt-rekkefølge ---------------------
        # Workbooks.Add() lager typisk 1 ark. Gi navn til eksisterende ark først,
        # legg deretter til resten etter hverandre slik at rekkefølgen blir riktig.
        # Bruk posisjonelle argumenter for Sheets.Add (Before, After) fordi
        # noen Excel/COM-kombinasjoner tolker navngitte argumenter feil.
        for idx, navn in enumerate(kontrakt.SHEETS, start=1):
            if idx <= wb.Sheets.Count:
                ark = wb.Sheets(idx)
            else:
                last = wb.Sheets(wb.Sheets.Count)
                ark = wb.Sheets.Add(None, last)
            ark.Name = navn

        # Fjern eventuelle overflødige ark (skal ikke skje, men defensivt)
        while wb.Sheets.Count > len(kontrakt.SHEETS):
            wb.Sheets(wb.Sheets.Count).Delete()

        logg(f"Ark etter oppsett ({wb.Sheets.Count} stk):")
        for i in range(1, wb.Sheets.Count + 1):
            logg(f"  [{i}] {wb.Sheets(i).Name}")

        # --- 2) Verifiser / rett codeNames ------------------------------------
        # Excel tildeler Sheet1..Sheet7 i opprettelses-/navngivingsrekkefølgen.
        # Noen Excel/COM-konfigurasjoner tillater ikke å skrive ws.CodeName
        # eller VBComponent.Name via dispatch; vi forsøker å rette, men godtar
        # verdiene hvis de allerede stemmer.
        for idx, navn in enumerate(kontrakt.SHEETS, start=1):
            forventet = kontrakt.CODE_NAMES[navn]
            comp = finn_arkkomponent_etter_indeks(wb.VBProject, idx)
            faktisk = comp.Name
            if faktisk != forventet:
                logg(f"Retter codeName for '{navn}': {faktisk!r} -> {forventet}")
                try:
                    comp.Name = forventet
                except Exception as e:
                    logg(f"  Kunne ikke sette codeName via COM ({e}); stoler på at Excel tildelte riktig navn")
            else:
                logg(f"codeName OK: '{navn}' = {faktisk}")

        # --- 3) Importer standardmoduler --------------------------------------
        for bas in BAS_MODULES:
            path = os.path.join(VBA_DIR, bas)
            if not os.path.isfile(path):
                raise FileNotFoundError(f"Mangler VBA-modul: {path}")
            comp = wb.VBProject.VBComponents.Import(path)
            logg(f"Importerte {bas} som {comp.Name}")

        # --- 4) Legg hendelsesbehandlere i arkmoduler og ThisWorkbook ---------
        hendelser = [
            ("Sheet2", 'Private Sub Worksheet_Change(ByVal Target As Range)\n    ModTidsstempel.TidsstempelEndring Target\nEnd Sub\n'),
            ("Sheet3", 'Private Sub Worksheet_Change(ByVal Target As Range)\n    ModSok.HåndterSøkKBEndring Target\nEnd Sub\n'),
            ("Sheet4", 'Private Sub Worksheet_Change(ByVal Target As Range)\n    ModSok.HåndterSøkSakerEndring Target\nEnd Sub\n'),
        ]

        for codename, kode in hendelser:
            comp = finn_arkkomponent(wb.VBProject, codename)
            comp.CodeModule.AddFromString(kode)
            logg(f"La til Worksheet_Change i {codename}")

        tw = wb.VBProject.VBComponents("ThisWorkbook")
        tw.CodeModule.AddFromString(
            'Private Sub Workbook_SheetActivate(ByVal Sh As Object)\n'
            '    If Sh.Name = "Dashboard" Then ModDashboard.OppdaterDashboard\n'
            'End Sub\n'
        )
        logg("La til Workbook_SheetActivate i ThisWorkbook")

        # --- 5) Verifiser komponenter og codename-map -------------------------
        komponenter = [c.Name for c in wb.VBProject.VBComponents]
        logg("VBComponents etter import:")
        for navn in sorted(komponenter):
            logg(f"  - {navn}")

        forventede_moduler = {"ModFelles", "ModSok", "ModDashboard", "ModTidsstempel"}
        mangler = forventede_moduler - set(komponenter)
        if mangler:
            raise RuntimeError(f"Mangler standardmoduler: {mangler}")

        logg("Endelig codeName-map:")
        for navn in kontrakt.SHEETS:
            ark = wb.Sheets(navn)
            logg(f"  '{navn}' -> {ark.CodeName}")
            if ark.CodeName != kontrakt.CODE_NAMES[navn]:
                raise RuntimeError(
                    f"codeName for '{navn}' er {ark.CodeName}, forventet {kontrakt.CODE_NAMES[navn]}"
                )

        # --- 6) Lagre donor .xlsm ---------------------------------------------
        if os.path.isfile(DONOR_PATH):
            os.remove(DONOR_PATH)
        wb.SaveAs(DONOR_PATH, FileFormat=XL_OPENXML_WORKBOOK_MACROENABLED)
        logg(f"Donor lagret: {DONOR_PATH}")

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception as e:
                logg(f"ADVARSEL ved lukking av workbook: {e}")
        if app is not None:
            try:
                app.Quit()
            except Exception as e:
                logg(f"ADVARSEL ved avslutting av Excel: {e}")
        logg("Excel avsluttet")


def ekstraher_bin():
    """Trekk xl/vbaProject.bin ut fra donor .xlsm."""
    if not os.path.isfile(DONOR_PATH):
        raise FileNotFoundError(f"Donor finnes ikke: {DONOR_PATH}")

    with zipfile.ZipFile(DONOR_PATH, "r") as z:
        data = z.read("xl/vbaProject.bin")

    with open(BIN_OUT, "wb") as f:
        f.write(data)

    logg(f"Ekstraherte xl/vbaProject.bin -> {BIN_OUT} ({len(data)} bytes)")


def verifiser_donor():
    """Sjekk at donor .xlsm og bin-filen er gyldige."""
    from openpyxl import load_workbook

    # 1) openpyxl skal se vbaProject.bin
    wb = load_workbook(DONOR_PATH, keep_vba=True)
    namelist = wb.vba_archive.namelist()
    har_bin = "xl/vbaProject.bin" in namelist
    logg(f"openpyxl keep_vba: xl/vbaProject.bin tilstede = {har_bin}")
    if not har_bin:
        raise RuntimeError("xl/vbaProject.bin mangler i donor ifølge openpyxl")

    # 2) Filstørrelse
    bin_size = os.path.getsize(BIN_OUT)
    logg(f"{BIN_OUT} størrelse = {bin_size} bytes")
    if bin_size == 0:
        raise RuntimeError("vbaProject.bin er tom")

    # 3) Excel COM skal liste de fire standardmodulene
    import win32com.client as win32
    app = None
    wb_com = None
    try:
        app = win32.Dispatch("Excel.Application")
        app.Visible = False
        app.DisplayAlerts = False
        wb_com = app.Workbooks.Open(DONOR_PATH)
        komponenter = {c.Name for c in wb_com.VBProject.VBComponents}
        forventede = {"ModFelles", "ModSok", "ModDashboard", "ModTidsstempel"}
        logg(f"Excel COM komponenter: {sorted(komponenter & forventede)}")
        if not forventede.issubset(komponenter):
            mangler = forventede - komponenter
            raise RuntimeError(f"Excel COM finner ikke modulene: {mangler}")

        # Verifiser at hendelsesbehandlere ligger i arkmodulene
        for codename in ("Sheet2", "Sheet3", "Sheet4"):
            comp = None
            for c in wb_com.VBProject.VBComponents:
                if c.Name == codename:
                    comp = c
                    break
            if comp is None:
                raise RuntimeError(f"Finner ikke komponent {codename}")
            tekst = comp.CodeModule.Lines(1, comp.CodeModule.CountOfLines)
            if "Worksheet_Change" not in tekst:
                raise RuntimeError(f"Worksheet_Change mangler i {codename}")
            logg(f"{codename} har Worksheet_Change")

        tw = wb_com.VBProject.VBComponents("ThisWorkbook")
        tw_tekst = tw.CodeModule.Lines(1, tw.CodeModule.CountOfLines)
        if "Workbook_SheetActivate" not in tw_tekst:
            raise RuntimeError("Workbook_SheetActivate mangler i ThisWorkbook")
        logg("ThisWorkbook har Workbook_SheetActivate")

    finally:
        if wb_com is not None:
            wb_com.Close(SaveChanges=False)
        if app is not None:
            app.Quit()
        logg("Excel avsluttet etter verifisering")


def hoved():
    # Rydd gammel evidence
    if os.path.isfile(EVIDENCE_FILE):
        os.remove(EVIDENCE_FILE)
    os.makedirs(EVIDENCE_DIR, exist_ok=True)

    logg("=== build_donor.py start ===")
    bygg_donor()
    ekstraher_bin()
    verifiser_donor()
    logg("=== build_donor.py ferdig ===")


if __name__ == "__main__":
    hoved()
