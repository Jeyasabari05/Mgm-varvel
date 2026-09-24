# MGM-Varvel RF / RK Product Selector

Catalogue-based selector for **RF** (parallel helical) and **RK** (helical-bevel) gear motors. The final PDF follows **DS-021-0046** (the supplied Product Information sheet). Extra catalogue columns are stored internally but are **not** printed.

## Run

```text
cd C:\Users\jeyas\Downloads\Attachments\Desktop\mgm-varvel-selector
py -3 -m pip install -r requirements.txt
py -3 main.py
```

Requires Python 3.10+. Tests:

```text
py -3 -m unittest tests.test_selection -v
```

## How selection works

1. Choose series **RF** or **RK** (records are never mixed).
2. Choose rated motor power **P1 (kW)** from tabulated values.
3. The overall gear ratio **i** list shows only ratios that exist for that power and series.
4. If two sizes share the same P1 + i, pick the configuration (model).
5. Optional catalogue options: mounting **M1–M6**, shaft **AC/AS/AD/ACC**, **SMX/BAX**, terminal box, cable entry.
6. **Generate PDF report** writes only DS-021-0046 fields.

Invalid P1 + i combinations are rejected. There is no nearest-ratio fallback.

## Documents used

| File | Role |
|---|---|
| `Downloads\DS-021-0046.pdf` | Output contract (fields, order, units, disclaimer) |
| `Downloads\RF-RK Catalogue - Version 8-2025-compressed.pdf` | Technical source (selection tables pp.37–45 RF, pp.71–78 RK) |

The DS sample product is **IR FPM 63, i = 35.74, 1.5 kW**. That IR row is **not** in the RF-RK tables. Closest RF check case: **RF, 1.5 kW, i = 36.53** → n2 = 39.4 rpm, M2 = 349 Nm, FS = 1.10.

## Values not in the RF-RK catalogue

These DS fields are printed as `Required value not found in catalogue.` — they were not invented:

- Efficiency (%)
- Rated current (A)
- Power factor (Cos ø)
- Finished paint colour (RAL not listed in RF-RK)
- RF solid-shaft dimensions (AS/AD table found for RK only, p.94)

Weight is the tabulated **gear unit mass without motor** (catalogue footnote).

Lubricant is **ISO VG 320** (RF-RK p.13 / p.22), not ISO VG 220 from the IR sample sheet.

## Project layout

- `mgm_selector/` — catalogue, lookup, PDF, Tkinter UI
- `data/` — DS field list, field map, size/oil/shaft lookups
- `scratch/products_flat.json` — parsed RF/RK gearmotor selection tables
- `tests/` — power/ratio, RF vs RK, invalid combo, PDF field filter
