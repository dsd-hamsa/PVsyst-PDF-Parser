# PVsyst PDF Parser (V3)

V3 is a fast, monitoring-oriented PVsyst report parser.

It is designed to produce a single JSON payload that contains:
- Raw inverter IDs (`INV01`, `INV02`, …) as stable keys
- A human-friendly inverter `description`
- A per-inverter `combined_configuration` array that consolidates MPPT allocation + the config fields you need for monitoring

V3 is implemented in `pvsyst_parser.py`.

## What’s New in V3

- **Text-only parsing (faster):** uses `pdfplumber` only (no Camelot/table extraction).
- **Monitoring-friendly output:** per inverter, a single `combined_configuration` list that includes MPPT → config mapping plus config details.
- **Stable IDs + friendly names:** JSON keys remain raw inverter IDs; `description` provides a display label.
- **`config_id` naming:** MPPT associations reference `config_id` instead of `array_id`.
- **Current handling:** `i_mpp_a` in `combined_configuration` is scaled to the MPPT based on strings-per-config and strings-per-MPPT.
- **Multiple module models (per config/MPPT):** module manufacturer/model is tracked per array configuration (and therefore can vary by inverter/MPPT). See `module_types` + `module_type_id` in the output.
- **Single-configuration fallback:** supports reports with no `Array #` blocks (one uniform config).
- **Industry heuristics:** can infer MPPT topology for common inverter families:
  - SMA Core1: 6 MPPT, max 2 strings/MPPT
  - CHINT / CPS: 3 MPPT, max 6 strings/MPPT

## Installation

### Prerequisites

- Python 3.9+ recommended

### Dependencies

```bash
pip install pdfplumber fastapi uvicorn
```

## CLI Usage

Parse a PVsyst PDF and write outputs (text + JSON) into an output directory:

```bash
python pvsyst_parser.py "path/to/report.pdf" --output-dir "./out"
```

Outputs:
- `*_analysis.txt`
- `*_structured.json`

## API Usage (FastAPI)

V3 API entry point is `app.py`.

Run on port **8888**:

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8888
```

Endpoints:
- `POST /api/parse` (multipart form field `file`)
- `GET /api/health`

Example:

```bash
curl -X POST "http://localhost:8888/api/parse" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@your_pvsyst_report.pdf"
```

Notes:
- `app.py` runs the V3 parse pipeline and returns JSON **without writing files**.

## Web UI Usage

Open `index.html` in a browser.

By default it calls:
- `http://localhost:8888/api/parse`

If deploying elsewhere, update `API_BASE` in `index.html`.

## Output Schema (V3)

### Top-level keys

- `metadata`
- `pv_module`
- `inverter` (global inverter info from PVsyst equipment table, if present)
- `module_types` (distinct PV module types detected)
- `inverter_types` (distinct inverter types detected)
- `array_configurations` (keyed by `config_id`)
- `associations` (keyed by raw inverter ID)
- `inverter_summary` (keyed by raw inverter ID)
- `system_monthly_production`
- `system_monthly_globhor`
- `orientations`

### The monitoring-friendly view

For each inverter `INVxx`, look at:

- `inverter_summary[INVxx].description`
- `inverter_summary[INVxx].combined_configuration[]`
- `inverter_summary[INVxx].pv_modules` (all PV module types feeding that inverter)

Notes:
- `inverter_summary[INVxx].pv_module` is only populated when that inverter uses exactly one module type; otherwise use `pv_modules`.

Each entry in `combined_configuration` is one MPPT row and includes:
- `mppt`
- `config_id`
- `module_type_id` (points to `module_types[]`)
- allocation: `strings`, `modules`, `dc_kwp`
- config fields: `tilt`, `azimuth`, `modules_in_series`, `u_mpp_v`, `i_mpp_a`

### About `i_mpp_a`

- `array_configurations[config_id].i_mpp_a` represents the **total current for the full configuration** (all strings in parallel).
- `combined_configuration[].i_mpp_a` represents the **total current for that MPPT**, computed as:

`(config_i_mpp_a / config_strings_total) * strings_on_that_mppt`

## Single-Configuration Reports (No `Array #` blocks)

Some PVsyst reports represent a site as one uniform configuration and do not include separate `Array #n` blocks.

V3 detects this and synthesizes one `config_id = "1"`, then distributes strings across MPPTs using the inferred inverter model topology.

V3 also records these diagnostic fields inside `array_configurations["1"]`:
- `inferred_inverters_reported`
- `inferred_inverters_required`
- `inferred_inverters_used`

## Files

- `pvsyst_parser.py` — core V3 parser
- `app.py` — V3 FastAPI service
- `index.html` — V3 web UI
