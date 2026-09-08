# -*- coding: utf-8 -*-
"""
Strukturell test-suite for genererte supportsystem-arbeidsbøker.

Validerer fallback .xlsx og Excel .xlsm via openpyxl og zip-inspeksjon.
Makroer kjøres IKKE – testene sjekker kun struktur/påstand.

Kjøring:
    python -m pytest test_struktur.py -v
    python test_struktur.py               # fallback hvis pytest mangler
"""
import datetime
import os
import sys
import zipfile

from openpyxl import load_workbook

import kontrakt

ROOT = os.path.dirname(os.path.abspath(__file__))
XLSX_FALLBACK = os.path.join(ROOT, "Supportsystem.xlsx")
XLSM_EXCEL = os.path.join(ROOT, "Supportsystem.xlsm")
ODS_CALC = os.path.join(ROOT, "Supportsystem.ods")

SAKER_B2_FORMULA = '=IF($C2&$H2="","",IF($B2="",NOW(),$B2))'
# Med eksempeldata overskrives B2..B11 med statiske datoer; sjekk første tomme rad (B12).
SAKER_B12_FORMULA = '=IF($C12&$H12="","",IF($B12="",NOW(),$B12))'

# Pytest er ønsket, men testene må kunne kjøre uten det.
try:
    import pytest
except Exception:  # pragma: no cover
    pytest = None

if pytest is None:  # pragma: no cover
    class _DummyPytest:
        class _Mark:
            @staticmethod
            def skip(reason=""):
                def decorator(func):
                    func.__skip__ = reason
                    return func
                return decorator

        @staticmethod
        def mark(*args, **kwargs):
            return _DummyPytest._Mark()

    pytest = _DummyPytest()


# ---------------------------------------------------------------------------
# Hjelpefunksjoner
# ---------------------------------------------------------------------------
def _load_xlsx():
    return load_workbook(XLSX_FALLBACK)


def _load_xlsm():
    return load_workbook(XLSM_EXCEL, keep_vba=True)


def _has_list_dv(ws, col, max_row):
    """Sjekk at det finnes en list-basert datavalidering for kolonnen."""
    expected_fragment = f"{col}2:{col}{max_row}"
    for dv in ws.data_validations.dataValidation:
        if dv.type != "list":
            continue
        sqref = getattr(dv, "sqref", "") or ""
        if expected_fragment in sqref or f"{col}2" in sqref:
            return True
    return False


# ---------------------------------------------------------------------------
# Fallback .xlsx
# ---------------------------------------------------------------------------
def test_xlsx_sheets_order():
    wb = _load_xlsx()
    assert wb.sheetnames == kontrakt.SHEETS, (
        f"Ark-rekkefølgen matcher ikke kontrakt.SHEETS: {wb.sheetnames}"
    )


def test_xlsx_saker_b2_circular_formula():
    wb = _load_xlsx()
    # Eksempeldata overskriver B2..B11 med statiske datoer; sirkelformelen skal
    # fortsatt ligge i første tomme rad (B12) og resten av området.
    value = wb["Saker"]["B12"].value
    assert isinstance(value, str) and value.startswith("="), (
        f"Forventet formel i Saker!B12, fikk {value!r}"
    )
    assert value == SAKER_B12_FORMULA, (
        f"Saker!B12 formel mismatch:\n  fikk:     {value}\n  forventet:{SAKER_B12_FORMULA}"
    )


def test_xlsx_helper_columns_have_formulas():
    wb = _load_xlsx()
    saker_y = wb["Saker"]["Y2"].value
    kb_l = wb["Kunnskapsbase"]["L2"].value
    assert isinstance(saker_y, str) and saker_y.startswith("="), (
        f"Forventet formel i Saker!Y2, fikk {saker_y!r}"
    )
    assert isinstance(kb_l, str) and kb_l.startswith("="), (
        f"Forventet formel i Kunnskapsbase!L2, fikk {kb_l!r}"
    )


def test_xlsx_dashboard_kpi_formulas():
    """
    Dashboard-KPI-formlene ligger i malen én rad under label-cellen
    (f.eks. A5 for label A4), så vi leser verdi-cellen fra kontrakt.KPIS.
    """
    wb = _load_xlsx()
    ws = wb["Dashboard"]
    allowed_prefixes = ("=COUNTIFS", "=IFERROR", "=COUNTA")
    for label_cell, _value_cell, _label, expected_formula in kontrakt.KPIS:
        col = label_cell[0]
        row = int("".join(ch for ch in label_cell if ch.isdigit()))
        value_cell = f"{col}{row + 1}"
        value = ws[value_cell].value
        assert isinstance(value, str) and value.startswith("="), (
            f"KPI-celle {value_cell} skal inneholde formel, fikk {value!r}"
        )
        assert value.startswith(allowed_prefixes), (
            f"KPI-celle {value_cell} har uventet formelstart: {value[:40]}"
        )
        assert value == expected_formula, (
            f"KPI-celle {value_cell} formel mismatch:\n  fikk:     {value}\n  forventet:{expected_formula}"
        )


def test_xlsx_defined_names_exist():
    wb = _load_xlsx()
    required = (
        "Ansatte",
        "Kategorier",
        "Prioriteter",
        "Statuser",
        "Kanaler",
        "KategoriMatrise",
        "SLAMatrise",
    )
    names = set(wb.defined_names)
    missing = [n for n in required if n not in names]
    assert not missing, f"Mangler definerte navn: {missing}"


def test_xlsx_data_validations():
    wb = _load_xlsx()
    saker = wb["Saker"]
    kb = wb["Kunnskapsbase"]
    for col in ("E", "F", "G", "K", "M", "W"):
        assert _has_list_dv(saker, col, kontrakt.MAX_S), (
            f"Mangler list-datavalidering for Saker!{col}2:{col}{kontrakt.MAX_S}"
        )
    for col in ("C", "J"):
        assert _has_list_dv(kb, col, kontrakt.MAX_K), (
            f"Mangler list-datavalidering for Kunnskapsbase!{col}2:{col}{kontrakt.MAX_K}"
        )


def test_xlsx_dashboard_has_three_charts():
    wb = _load_xlsx()
    assert len(wb["Dashboard"]._charts) == 3, (
        f"Forventet 3 diagrammer på Dashboard, fikk {len(wb['Dashboard']._charts)}"
    )


def test_xlsx_zip_has_form_controls():
    assert os.path.exists(XLSX_FALLBACK), f"{XLSX_FALLBACK} finnes ikke"
    with zipfile.ZipFile(XLSX_FALLBACK, "r") as z:
        names = z.namelist()
        ctrlprops = [n for n in names if n.startswith("xl/ctrlProps/")]
        assert ctrlprops, "Mangler xl/ctrlProps/-deler (avkryssingsbokser)"
        assert "xl/drawings/vmlDrawing1.vml" in names, (
            "Mangler xl/drawings/vmlDrawing1.vml"
        )


def test_xlsx_no_codenames():
    wb = _load_xlsx()
    for ws in wb.worksheets:
        assert ws.sheet_properties.codeName is None, (
            f"{ws.title} skal ikke ha codeName, har {ws.sheet_properties.codeName!r}"
        )


# ---------------------------------------------------------------------------
# Excel .xlsm
# ---------------------------------------------------------------------------
def test_xlsm_vba_project_present():
    wb = _load_xlsm()
    archive = wb.vba_archive
    assert archive is not None, "Ingen vba_archive i .xlsm"
    assert "xl/vbaProject.bin" in archive.namelist(), (
        "xl/vbaProject.bin finnes ikke i .xlsm"
    )


def test_xlsm_content_types_macro_enabled():
    assert os.path.exists(XLSM_EXCEL), f"{XLSM_EXCEL} finnes ikke"
    with zipfile.ZipFile(XLSM_EXCEL, "r") as z:
        ct = z.read("[Content_Types].xml").decode("utf-8")
    assert "macroEnabled" in ct, "[Content_Types].xml mangler macroEnabled"
    assert "vbaProject" in ct, "[Content_Types].xml mangler vbaProject"


def test_xlsm_sheets_have_codenames():
    wb = _load_xlsm()
    for ws in wb.worksheets:
        expected = kontrakt.CODE_NAMES[ws.title]
        actual = ws.sheet_properties.codeName
        assert actual == expected, (
            f"codeName for {ws.title}: forventet {expected!r}, fikk {actual!r}"
        )


def test_xlsm_saker_b2_not_formula():
    wb = _load_xlsm()
    value = wb["Saker"]["B2"].value
    assert not (isinstance(value, str) and value.startswith("=")), (
        f"Saker!B2 skal IKKE være formel i .xlsm, fikk {value!r}"
    )
    assert value is None or isinstance(value, datetime.datetime), (
        f"Saker!B2 forventet None/datetime i .xlsm, fikk {type(value).__name__}: {value!r}"
    )


def test_xlsm_helper_columns_empty():
    wb = _load_xlsm()
    assert wb["Saker"]["Y2"].value is None, "Saker!Y2 skal være tom i .xlsm"
    assert wb["Kunnskapsbase"]["L2"].value is None, "Kunnskapsbase!L2 skal være tom i .xlsm"


def test_xlsm_search_result_rows_empty():
    wb = _load_xlsm()
    assert wb["Søk KB"]["A9"].value is None, "Søk KB!A9 skal være tom i .xlsm"
    assert wb["Søk saker"]["A8"].value is None, "Søk saker!A8 skal være tom i .xlsm"


def test_xlsm_status_mirror_and_checkbox():
    wb = _load_xlsm()
    aa5 = wb["Søk saker"]["AA5"].value
    ab5 = wb["Søk saker"]["AB5"].value
    assert isinstance(aa5, str) and aa5.startswith("="), (
        f"Søk saker!AA5 skal være status-speil-formel, fikk {aa5!r}"
    )
    assert isinstance(ab5, bool), (
        f"Søk saker!AB5 skal være boolsk avkryssings-celle, fikk {ab5!r}"
    )


# ---------------------------------------------------------------------------
# LO/OO .ods (Basic-makroer)
# ---------------------------------------------------------------------------
def test_ods_calc_macro_parts():
    """
    Sjekker at Supportsystem.ods inneholder Basic-moduler og manifest-parts.
    """
    assert os.path.exists(ODS_CALC), f"{ODS_CALC} finnes ikke enn�"
    with zipfile.ZipFile(ODS_CALC, "r") as z:
        names = z.namelist()
        for i in range(1, 5):
            assert f"Basic/Standard/Module{i}.xml" in names, (
                f"Mangler Basic/Standard/Module{i}.xml i {ODS_CALC}"
            )
        assert "Basic/Standard/script-lb.xml" in names, "Mangler Basic/Standard/script-lb.xml"
        assert "Basic/script-lc.xml" in names, "Mangler Basic/script-lc.xml"
        manifest = z.read("META-INF/manifest.xml").decode("utf-8")
    assert 'manifest:full-path="Basic/"' in manifest, "Manifest mangler Basic/"
    assert 'manifest:full-path="Basic/script-lc.xml"' in manifest, "Manifest mangler script-lc.xml"
    assert 'manifest:full-path="Basic/Standard/"' in manifest, "Manifest mangler Basic/Standard/"
    assert 'manifest:full-path="Basic/Standard/script-lb.xml"' in manifest, "Manifest mangler script-lb.xml"



# ---------------------------------------------------------------------------
# Fallback-kjøring uten pytest
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    # Hvis pytest er tilgjengelig, la det ta over for riktig oppsummering.
    if pytest is not None and getattr(pytest, "main", None) is not None:
        sys.exit(pytest.main([__file__, "-v"]))

    import traceback

    tests = [
        obj for name, obj in globals().items()
        if name.startswith("test_") and callable(obj)
    ]
    passed = skipped = failed = 0
    for test in tests:
        skip_reason = getattr(test, "__skip__", None)
        if skip_reason is not None:
            print(f"SKIP {test.__name__}: {skip_reason}")
            skipped += 1
            continue
        try:
            test()
            print(f"PASS {test.__name__}")
            passed += 1
        except Exception as exc:  # pragma: no cover
            print(f"FAIL {test.__name__}: {exc}")
            traceback.print_exc()
            failed += 1
    print(f"\n{passed} passed, {skipped} skipped, {failed} failed")
    sys.exit(0 if failed == 0 else 1)
