"""
Test suite for D'Casa Honduras .dat file parser.

Imports all parsing logic from parser.py and validates the output
against Resultado esperado.xlsx.

Run with: pytest test_parser.py -v
"""

import pytest
from datetime import datetime
from pathlib import Path

import openpyxl

from parser import (
    BASE_DIR,
    SPANISH_MONTHS_FULL,
    SPANISH_DAYS,
    parse_dat_date,
    semana_del_mes,
    _num,
    parse_invr0601,
    parse_venr15,
    parse_facr12,
    parse_notr03,
    parse_invr29,
    load_variables,
    build_dataset,
)


# ---------------------------------------------------------------------------
# Load expected output
# ---------------------------------------------------------------------------

def load_expected() -> dict:
    """Load Resultado esperado.xlsx and return a dict keyed by (documento, linea)."""
    wb = openpyxl.load_workbook(BASE_DIR / "Resultado esperado.xlsx")
    ws = wb["MAIN"]
    rows = [r for r in ws.iter_rows(values_only=True) if any(v is not None for v in r)]
    headers = rows[0]

    result = {}
    for row in rows[1:]:
        d = dict(zip(headers, row))
        # Normalise the AÑO header which may appear garbled depending on file encoding
        for k in list(d.keys()):
            if k and k.startswith("A") and "O" in k and len(k) <= 5 and k != "AÑO":
                d["AÑO"] = d.pop(k)
                break
        key = (d["DOCUMENTO"], d["LINEA FAC"])
        result[key] = d
    return result


# ---------------------------------------------------------------------------
# Pytest fixtures
# ---------------------------------------------------------------------------

@pytest.fixture(scope="module")
def dataset():
    return build_dataset()


@pytest.fixture(scope="module")
def expected():
    return load_expected()


@pytest.fixture(scope="module")
def output_by_key(dataset):
    return {(r["DOCUMENTO"], r["LINEA FAC"]): r for r in dataset}


# ---------------------------------------------------------------------------
# Tests
# ---------------------------------------------------------------------------

class TestRowCount:
    def test_total_rows(self, dataset, expected):
        """Parsed output must have the same number of rows as the expected file."""
        assert len(dataset) == len(expected), (
            f"Row count mismatch: parsed={len(dataset)}, expected={len(expected)}"
        )

    def test_no_duplicate_keys(self, dataset):
        """Each (documento, linea) pair must appear exactly once."""
        keys = [(r["DOCUMENTO"], r["LINEA FAC"]) for r in dataset]
        assert len(keys) == len(set(keys)), "Duplicate (documento, linea) keys found"


class TestDocumentKeys:
    def test_all_expected_keys_present(self, output_by_key, expected):
        """Every (documento, linea) in the expected file must appear in the output."""
        missing = set(expected.keys()) - set(output_by_key.keys())
        assert not missing, f"Missing {len(missing)} keys: {sorted(missing)[:10]}"

    def test_no_extra_keys(self, output_by_key, expected):
        """Output must not contain keys absent from the expected file."""
        extra = set(output_by_key.keys()) - set(expected.keys())
        assert not extra, f"Extra {len(extra)} keys: {sorted(extra)[:10]}"


class TestStringFields:
    """Exact-match tests for all string columns."""

    STRING_COLS = [
        "DOCUMENTO", "PROVEEDOR", "CODIGO BARRA", "DESCRIPCION",
        "UNID/VENTA", "MODULO", "ESTACION", "MES", "DIA",
        "DEPARTAMENTO", "MUNICIPIO", "NOMBRE CLIENTE", "TIPO CLIENTE",
        "VENDEDOR", "RUTA", "TIPO VENTA", "RAZON", "RUTA-PRESUPUESTO",
    ]

    @pytest.mark.parametrize("col", STRING_COLS)
    def test_column(self, col, output_by_key, expected):
        mismatches = []
        for key, exp_row in expected.items():
            out_row = output_by_key.get(key)
            if out_row is None:
                continue
            exp_val = exp_row.get(col, "")
            out_val = out_row.get(col, "")
            if exp_val != out_val:
                mismatches.append(
                    f"key={key}: expected={repr(exp_val)}, got={repr(out_val)}"
                )
        assert not mismatches, (
            f"Column '{col}' has {len(mismatches)} mismatches:\n"
            + "\n".join(mismatches[:5])
        )


class TestIntegerFields:
    """Exact-match tests for integer columns."""

    INT_COLS = [
        "LINEA FAC", "BODEGA", "COD PROV", "LINEA",
        "CANTIDAD", "ALTERNO", "AÑO", "SEMANA",
        "CODIGO CLIENTE", "COD VENDEDOR",
    ]

    @pytest.mark.parametrize("col", INT_COLS)
    def test_column(self, col, output_by_key, expected):
        mismatches = []
        for key, exp_row in expected.items():
            out_row = output_by_key.get(key)
            if out_row is None:
                continue
            exp_val = exp_row.get(col)
            out_val = out_row.get(col)
            if exp_val != out_val:
                mismatches.append(
                    f"key={key}: expected={exp_val}, got={out_val}"
                )
        assert not mismatches, (
            f"Column '{col}' has {len(mismatches)} mismatches:\n"
            + "\n".join(mismatches[:5])
        )


class TestFloatFields:
    """Approximate-match tests for numeric columns (1% relative tolerance)."""

    FLOAT_COLS = ["COSTO", "PRECIO", "DESCUENTO", "PESO"]
    TOL = 0.01

    @pytest.mark.parametrize("col", FLOAT_COLS)
    def test_column(self, col, output_by_key, expected):
        mismatches = []
        for key, exp_row in expected.items():
            out_row = output_by_key.get(key)
            if out_row is None:
                continue
            exp_val = float(exp_row.get(col) or 0)
            out_val = float(out_row.get(col) or 0)
            tol = max(self.TOL, 0.01 * abs(exp_val))
            if abs(exp_val - out_val) > tol:
                mismatches.append(
                    f"key={key}: expected={exp_val}, got={out_val}"
                )
        assert not mismatches, (
            f"Column '{col}' has {len(mismatches)} mismatches:\n"
            + "\n".join(mismatches[:5])
        )


class TestDateField:
    def test_fecha(self, output_by_key, expected):
        """FECHA must match as a datetime (date part only)."""
        mismatches = []
        for key, exp_row in expected.items():
            out_row = output_by_key.get(key)
            if out_row is None:
                continue
            exp_d = exp_row.get("FECHA")
            out_d = out_row.get("FECHA")
            if exp_d and out_d and exp_d.date() != out_d.date():
                mismatches.append(f"key={key}: expected={exp_d.date()}, got={out_d.date()}")
        assert not mismatches, (
            f"FECHA has {len(mismatches)} mismatches:\n" + "\n".join(mismatches[:5])
        )


class TestOrdenField:
    def test_orden(self, output_by_key, expected):
        """ORDEN must match; returns use 0, sales use the order string."""
        mismatches = []
        for key, exp_row in expected.items():
            out_row = output_by_key.get(key)
            if out_row is None:
                continue
            exp_val = exp_row.get("ORDEN")
            out_val = out_row.get("ORDEN")
            if str(exp_val) != str(out_val):
                mismatches.append(
                    f"key={key}: expected={repr(exp_val)}, got={repr(out_val)}"
                )
        assert not mismatches, (
            f"ORDEN has {len(mismatches)} mismatches:\n" + "\n".join(mismatches[:5])
        )


# ---------------------------------------------------------------------------
# Unit tests for helper functions
# ---------------------------------------------------------------------------

class TestHelpers:
    def test_parse_dat_date(self):
        assert parse_dat_date("03.MAR.26") == datetime(2026, 3, 3)

    def test_parse_dat_date_december(self):
        assert parse_dat_date("31.DIC.25") == datetime(2025, 12, 31)

    def test_semana_week1(self):
        assert semana_del_mes(datetime(2026, 3, 1)) == 1
        assert semana_del_mes(datetime(2026, 3, 7)) == 1

    def test_semana_week2(self):
        assert semana_del_mes(datetime(2026, 3, 8))  == 2
        assert semana_del_mes(datetime(2026, 3, 14)) == 2

    def test_semana_week3(self):
        assert semana_del_mes(datetime(2026, 3, 15)) == 3

    def test_num_with_commas(self):
        assert _num("3,317.59") == pytest.approx(3317.59)

    def test_num_empty(self):
        assert _num("") == 0.0

    def test_spanish_day_tuesday(self):
        assert SPANISH_DAYS[datetime(2026, 3, 3).weekday()] == "MARTES"

    def test_spanish_month_march(self):
        assert SPANISH_MONTHS_FULL[3] == "MARZO"


# ---------------------------------------------------------------------------
# Smoke tests for individual parsers
# ---------------------------------------------------------------------------

class TestParsers:
    def test_parse_invr0601_ventas_count(self):
        assert len(parse_invr0601(BASE_DIR / "lp60.dat", "FACTURACION", "VENTA")) > 0

    def test_parse_invr0601_returns_count(self):
        assert len(parse_invr0601(BASE_DIR / "lp59.dat", "DEVOLUCION", "NOTA DE CREDITO")) > 0

    def test_parse_invr0601_sample_row(self):
        sales = parse_invr0601(BASE_DIR / "lp60.dat", "FACTURACION", "VENTA")
        row = next(
            (r for r in sales if r["documento"] == "000-002-01-00053402" and r["linea"] == 6),
            None,
        )
        assert row is not None
        assert row["bodega"]     == 4
        assert row["cod_prov"]   == 108
        assert row["linea_prod"] == 36
        assert row["cod_barra"]  == "7503002 163610"
        assert row["cantidad"]   == 1
        assert row["costo"]      == pytest.approx(217.85, abs=0.01)
        assert row["precio"]     == pytest.approx(294.56, abs=0.01)
        assert row["descuento"]  == pytest.approx(13.94,  abs=0.01)
        assert row["alterno"]    == 87320

    def test_parse_venr15_sample(self):
        d = parse_venr15(BASE_DIR / "lp33.dat")
        assert "000-002-01-00053402" in d
        assert d["000-002-01-00053402"]["fecha"]   == datetime(2026, 3, 3)
        assert d["000-002-01-00053402"]["cliente"] == 30312102

    def test_parse_facr12_sample(self):
        d = parse_facr12(BASE_DIR / "lp35.dat")
        assert "000-002-01-00053402" in d
        assert d["000-002-01-00053402"]["cliente"] == 30312102
        assert d["000-002-01-00053402"]["orden"]   == "T6961T0"

    def test_parse_notr03_sample(self):
        d = parse_notr03(BASE_DIR / "lp31.dat")
        assert "000-002-06-00025691" in d
        assert d["000-002-06-00025691"]["fecha"]   == datetime(2026, 3, 3)
        assert d["000-002-06-00025691"]["cliente"] == 803339

    def test_parse_invr29_sample(self):
        d = parse_invr29(BASE_DIR / "lp51.dat")
        assert d[("000-002-01-00053402", 6)] == "WALMART CDS"

    def test_parse_invr29_return_razon(self):
        d = parse_invr29(BASE_DIR / "lp51.dat")
        assert d[("000-002-06-00025691", 1)] == "DEVOLUCION TEG."

    def test_load_variables_proveedor(self):
        v = load_variables(BASE_DIR / "Variables.xlsx")
        assert v["prov_lookup"].get(108) == "HENKEL LA LUZ"
        assert v["prov_lookup"].get(115) == "SUPER DE ALIMENTOS"
        assert v["prov_lookup"].get(102) == "SUMMA INDUSTRIAL"

    def test_load_variables_bodega_estacion(self):
        v = load_variables(BASE_DIR / "Variables.xlsx")
        assert v["bodega_est"].get(4)  == "TEG"
        assert v["bodega_est"].get(3)  == "SPS"
        assert v["bodega_est"].get(45) == "TEG"

    def test_load_variables_peso(self):
        v = load_variables(BASE_DIR / "Variables.xlsx")
        assert v["peso_lookup"].get(87320) == pytest.approx(7.82, abs=0.01)

    def test_load_variables_client_ruta(self):
        v = load_variables(BASE_DIR / "Variables.xlsx")
        cli = v["client_ruta"].get(30312102)
        assert cli is not None
        assert cli["nombre"]    == "WALMART"
        assert cli["tipo"]      == "SUPERMERCADOS"
        assert cli["ruta_pres"] == "HN11"


# ---------------------------------------------------------------------------
# End-to-end spot checks
# ---------------------------------------------------------------------------

class TestBuildDatasetSamples:
    def test_walmart_teg_row(self, dataset):
        row = next(
            (r for r in dataset
             if r["DOCUMENTO"] == "000-002-01-00053402" and r["LINEA FAC"] == 6),
            None,
        )
        assert row is not None, "Expected row not found in dataset"
        assert row["BODEGA"]           == 4
        assert row["COD PROV"]         == 108
        assert row["PROVEEDOR"]        == "HENKEL LA LUZ"
        assert row["LINEA"]            == 36
        assert row["CODIGO BARRA"]     == "7503002 163610"
        assert row["DESCRIPCION"]      == "XTREME GEL ATTRA.12/250g"
        assert row["CANTIDAD"]         == 1
        assert row["COSTO"]            == pytest.approx(217.85, abs=0.01)
        assert row["PRECIO"]           == pytest.approx(294.56, abs=0.01)
        assert row["DESCUENTO"]        == pytest.approx(13.94,  abs=0.01)
        assert row["ALTERNO"]          == 87320
        assert row["MODULO"]           == "FACTURACION"
        assert row["ESTACION"]         == "TEG"
        assert row["MES"]              == "MARZO"
        assert row["AÑO"]              == 2026
        assert row["FECHA"]            == datetime(2026, 3, 3)
        assert row["DIA"]              == "MARTES"
        assert row["SEMANA"]           == 1
        assert row["CODIGO CLIENTE"]   == 30312102
        assert row["DEPARTAMENTO"]     == "FRANCISCO MORAZAN"
        assert row["MUNICIPIO"]        == "DISTRITO CENTRAL"
        assert row["NOMBRE CLIENTE"]   == "WALMART"
        assert row["TIPO CLIENTE"]     == "SUPERMERCADOS"
        assert row["COD VENDEDOR"]     == 10
        assert row["VENDEDOR"]         == "ETNA VENEGAS WAI"
        assert row["RUTA"]             == "Supermercado"
        assert row["TIPO VENTA"]       == "VENTA"
        assert row["ORDEN"]            == "T6961T0"
        assert row["RAZON"]            == "WALMART CDS"
        assert row["PESO"]             == pytest.approx(7.82, abs=0.01)
        assert row["RUTA-PRESUPUESTO"] == "HN11"

    def test_return_nota_credito_row(self, dataset):
        row = next(
            (r for r in dataset
             if r["DOCUMENTO"] == "000-002-06-00025691" and r["LINEA FAC"] == 1),
            None,
        )
        assert row is not None, "Expected return row not found"
        assert row["TIPO VENTA"]  == "NOTA DE CREDITO"
        assert row["MODULO"]      == "DEVOLUCION"
        assert row["RAZON"]       == "DEVOLUCION TEG."
        assert row["ORDEN"]       == 0
        assert row["COD PROV"]    == 115
        assert row["PROVEEDOR"]   == "SUPER DE ALIMENTOS"
        assert row["CANTIDAD"]    == 50
        assert row["FECHA"]       == datetime(2026, 3, 3)
