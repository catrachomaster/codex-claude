# D'Casa Honduras — DAT File Parser & Test Suite

A Python test suite that parses fixed-width `.dat` export files from a Honduran business ERP system, enriches them with lookup tables, and validates the output against an expected Excel result.

## Overview

The pipeline reads raw ERP reports, joins them with reference data, and produces a 34-column enriched dataset (2,142 rows) matching `Resultado esperado.xlsx`.

```
lp60.dat  (sales)    ┐
lp59.dat  (returns)  ┤─► line items
                     │
lp33.dat  (invoice diary)   ┐
lp35.dat  (invoice-client)  ├─► date + client + order per document
lp31.dat  (credit notes)    ┘
                     │
lp51.dat  (article movements) ──► RAZON per document-line
                     │
Variables.xlsx ──────────────► lookups: vendors, providers, weights, routes
                     │
                     ▼
            Enriched dataset  ──► validated against Resultado esperado.xlsx
```

## Source files

| File | Report | Contents |
|------|--------|----------|
| `lp60.dat` | INVR0601 VENTAS | Sales line items (7,320 lines) |
| `lp59.dat` | INVR0601 DEVOLUCIONES | Returns in identical format (438 lines) |
| `lp33.dat` | VENR15 Diario de Facturas | Date + client per sales document |
| `lp35.dat` | FACR12 Relacion Factura Cliente | Client + order number per document |
| `lp31.dat` | NOTR03 Reporte de Notas | Date + client for credit note documents |
| `lp51.dat` | INVR29 Movimiento por Articulos | Transaction reason (RAZON) per doc-line |
| `Variables.xlsx` | — | Lookup sheets: routes, vendors, providers, weights |
| `Resultado esperado.xlsx` | — | Expected output (2,142 rows × 34 columns) |

## Output columns

`DOCUMENTO`, `LINEA FAC`, `BODEGA`, `COD PROV`, `PROVEEDOR`, `LINEA`, `CODIGO BARRA`, `DESCRIPCION`, `UNID/VENTA`, `CANTIDAD`, `COSTO`, `PRECIO`, `DESCUENTO`, `ALTERNO`, `MODULO`, `ESTACION`, `MES`, `AÑO`, `FECHA`, `DIA`, `SEMANA`, `CODIGO CLIENTE`, `DEPARTAMENTO`, `MUNICIPIO`, `NOMBRE CLIENTE`, `TIPO CLIENTE`, `COD VENDEDOR`, `VENDEDOR`, `RUTA`, `TIPO VENTA`, `ORDEN`, `RAZON`, `PESO`, `RUTA-PRESUPUESTO`

## Requirements

```
pip install pytest openpyxl
```

## Running the tests

```
pytest test_parser.py -v
```

61 tests across 8 test classes — all passing.

## Test classes

| Class | What it checks |
|-------|---------------|
| `TestRowCount` | Total row count and no duplicate keys |
| `TestDocumentKeys` | Every expected `(documento, linea)` key is present and no extras |
| `TestStringFields` | Exact match on 18 string columns |
| `TestIntegerFields` | Exact match on 10 integer columns |
| `TestFloatFields` | Numeric match (1% relative tolerance) on COSTO, PRECIO, DESCUENTO, PESO |
| `TestDateField` | FECHA date equality |
| `TestOrdenField` | Order number match (returns use `0`) |
| `TestHelpers` | Unit tests for date parsing, week-of-month, Spanish locale helpers |
| `TestParsers` | Smoke tests for each individual parser function |
| `TestBuildDatasetSamples` | Spot-checks two fully-enriched rows end-to-end |
