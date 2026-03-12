# D'Casa Honduras — ERP Report Parser

Parses fixed-width `.dat` exports from a Honduran business ERP system, enriches each line item with lookup tables, and writes the result to Excel.

## How it works

```
lp60.dat  (sales)    ┐
lp59.dat  (returns)  ┤── line items
                     │
lp33.dat  (invoice diary)   ┐
lp35.dat  (invoice-client)  ├── date + client + order per document
lp31.dat  (credit notes)    ┘
                     │
lp51.dat  (article movements) ── RAZON per document-line
                     │
Variables.xlsx ────────── lookups: vendors, providers, weights, routes
                     │
                     ▼
            Resultado.xlsx  (2,142 rows × 34 columns)
```

## Files

| File | Report | Description |
|------|--------|-------------|
| `lp60.dat` | INVR0601 VENTAS | Sales line items |
| `lp59.dat` | INVR0601 DEVOLUCIONES | Returns (same format) |
| `lp33.dat` | VENR15 Diario de Facturas | Date + client per sales document |
| `lp35.dat` | FACR12 Relacion Factura Cliente | Client + order number per document |
| `lp31.dat` | NOTR03 Reporte de Notas | Date + client for credit note documents |
| `lp51.dat` | INVR29 Movimiento por Articulos | Transaction reason (RAZON) per doc-line |
| `Variables.xlsx` | — | Lookup sheets: routes, vendors, providers, weights |

## Output columns

`DOCUMENTO` · `LINEA FAC` · `BODEGA` · `COD PROV` · `PROVEEDOR` · `LINEA` · `CODIGO BARRA` · `DESCRIPCION` · `UNID/VENTA` · `CANTIDAD` · `COSTO` · `PRECIO` · `DESCUENTO` · `ALTERNO` · `MODULO` · `ESTACION` · `MES` · `AÑO` · `FECHA` · `DIA` · `SEMANA` · `CODIGO CLIENTE` · `DEPARTAMENTO` · `MUNICIPIO` · `NOMBRE CLIENTE` · `TIPO CLIENTE` · `COD VENDEDOR` · `VENDEDOR` · `RUTA` · `TIPO VENTA` · `ORDEN` · `RAZON` · `PESO` · `RUTA-PRESUPUESTO`

## Setup

```
pip install openpyxl pytest
```

## Usage

```
python parser.py                      # writes Resultado.xlsx
python parser.py --out report.xlsx    # custom output path
```

## Tests

```
pytest test_parser.py -v
```

61 tests across 8 classes — all passing.

| Test class | What it checks |
|------------|---------------|
| `TestRowCount` | 2,142 rows, no duplicate keys |
| `TestDocumentKeys` | Every expected `(documento, linea)` present, no extras |
| `TestStringFields` | Exact match on 18 string columns |
| `TestIntegerFields` | Exact match on 10 integer columns |
| `TestFloatFields` | 1% relative tolerance on COSTO, PRECIO, DESCUENTO, PESO |
| `TestDateField` | FECHA date equality |
| `TestOrdenField` | Order number match (returns → `0`) |
| `TestHelpers` | Unit tests for date parsing, week-of-month, Spanish locale |
| `TestParsers` | Smoke tests for each individual parser function |
| `TestBuildDatasetSamples` | End-to-end spot checks on two fully-enriched rows |
