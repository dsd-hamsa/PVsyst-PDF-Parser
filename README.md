# PVsyst PDF Parser (V3.0.2)

V3.0.2 is a fast, monitoring-oriented PVsyst report parser.

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
- **Validation:** Cross-validates parsed data against authoritative "Total Inverter Power" sections
- **Enhanced Debugging:** Detailed warnings for parsing edge cases with array block inspection

## Installation

### Prerequisites

- Python 3.9+ recommended

### Dependencies

```bash
pip install -r requirements.txt
```

**Note**: V3 uses text-only parsing with `pdfplumber` for faster, more reliable extraction.

## CLI Usage

Parse a PVsyst PDF and write outputs (text + JSON) into an output directory:

```bash
python3 pvsyst_parser.py "path/to/report.pdf" --output-dir "./out"
```

Outputs:
- `*_analysis_v3.txt`
- `*_structured_v3.json`

Optional: generate an additional PowerTrack patch JSON (per inverter):

```bash
python3 pvsyst_parser.py "path/to/report.pdf" --powertrack-patch
```

You can override the output path:

```bash
python3 pvsyst_parser.py "path/to/report.pdf" --powertrack-patch \
  --powertrack-patch-path "./out/site_powertrack_patch.json"
```

## PowerTrack Patch Output

When `--powertrack-patch` is enabled, the parser writes a second JSON file whose top-level keys are `PV0`, `PV1`, ... (PowerTrack-style keys), with one patch object per inverter.

### PowerTrack keys (PV0/PV1/...)

- Default mapping: `INV01` -> `PV0`, `INV02` -> `PV1`, ... (derived from the numeric suffix minus 1).
- If an inverter ID has no usable numeric suffix (or would collide), the next available `PV{n}` is assigned.
- Inverters are processed in sorted order by raw inverter ID, so key assignment is deterministic for a given report.

### Patch schema

Each `PV{n}` entry looks like:

```json
{
  "PV0": {
    "description": "Inv 01 - (125.0 kW) - SMA Sunny Tripower CORE1",
    "pvConfig": {
      "inverters": [
        {
          "numOfStrings": 2,
          "panelsPerString": 28,
          "wattsPerPanel": 540,
          "inverterKw": 125.0,
          "azimuth": 180.0,
          "tilt": 20.0,
          "dcSize": 30.24,
          "mppVoltage": 950.0,
          "mppAmps": 26.2,
          "mppWatts": 24890.0
        }
      ],
      "monthlyOutput": {
        "jan": 0,
        "feb": 0,
        "mar": 0,
        "apr": 0,
        "may": 0,
        "jun": 0,
        "jul": 0,
        "aug": 0,
        "sep": 0,
        "oct": 0,
        "nov": 0,
        "dec": 0
      },
      "degrade": 0.5
    }
  }
}
```

Notes:

- `pvConfig.inverters[]` is MPPT-level (one entry per `inverter_summary[INVxx].combined_configuration` row).
- `monthlyOutput` is annual energy split by month in kWh (rounded to integers) with keys `jan..dec`.
- `degrade` comes from the PVsyst Array Losses thermal loss percent when available (units are percent, e.g. `0.5` means `0.5%`).
- `mppVoltage`/`mppAmps`/`mppWatts` are included only when the underlying values are present (nulls are omitted).

## API Usage

V3 API entry point is `app.py`.

Run on port **8000**:

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8000
```

Endpoints:
- `POST /api/parse` (multipart form field `file`)
- `GET /api/health`

Example:

```bash
curl -X POST "http://localhost:8000/api/parse" \
  -H "accept: application/json" \
  -H "Content-Type: multipart/form-data" \
  -F "file=@your_pvsyst_report.pdf"
```

Notes:
- `app.py` runs the V3 parse pipeline and returns JSON **without writing files**.

## Web UI Usage

Start the API server (see above), then open `index.html` in a browser.

When opened from disk (file://), the UI defaults to `http://localhost:8000/api`.

Tip: you can override the backend URL with a query param, for example:

`index.html?apiBase=http://localhost:8000` (or `.../api`)

```bash
uvicorn app:app --reload --port 8000
```

**V3 Interface:** Modern UI that displays combined MPPT configurations and detects all inverter types.

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

## V3 Features & Enhancements

### Single-Array Round-Robin Allocation

For reports with single configurations, V3 distributes strings round-robin across all available MPPT endpoints:

- **Pattern**: INV01-MPPT1, INV02-MPPT1, INV03-MPPT1, ..., INV01-MPPT2, INV02-MPPT2, etc.
- **Benefits**: Fair distribution, optimal capacity utilization, no false "over-limit" warnings
- **Validation**: Cross-checks against "Total Inverter Power" section for accuracy

### Validation & Debugging

V3 includes built-in validation that compares parsed data against authoritative sections:

- **Inverter Count Validation**: Matches parsed count with "Total Inverter Power" section
- **Warning Output**: Clear messages for mismatches with debugging details
- **Edge Case Handling**: Helps identify regex failures in complex PVsyst formats

## Files

- `pvsyst_parser.py` — V3 core parsing logic with validation
- `app.py` — V3 FastAPI web application
- `index.html` — V3 web interface
- `requirements.txt` — Dependencies
- `README.md` — This file

## Dependencies

- **pdfplumber**: PDF text extraction (primary parsing engine)
- **fastapi**: Web API framework
- **uvicorn**: ASGI server

**V3 Changes**: Removed Camelot table extraction dependency for faster, more reliable text-only parsing.

## Troubleshooting

### Common Issues

1. **Text extraction fails**: Ensure pdfplumber can read your PDF
2. **Web interface not loading**: Verify uvicorn is running and port 8888 is accessible
3. **Inverter count mismatch warnings**: Check PDF's "Total Inverter Power" section and array headers for consistency

### PDF Compatibility

- Tested with PVsyst V7.x and V8.x reports
- Works with standard PVsyst PDF exports
- **V3 Enhancements**: Better handling of single-array reports and validation warnings for edge cases
- May require adjustments for heavily customized reports (use validation warnings for debugging)

## License

MIT License (whatever that means)

## Contributing

Contributions welcome! Please:

1. Fork the repository
2. Create a feature branch
3. Add tests for new functionality
4. Submit a pull request

## Support

For issues or questions:

- Open an issue on GitHub
- Check the troubleshooting section above
- Ensure your PVsyst PDF is a standard export format
