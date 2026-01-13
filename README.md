# PVsyst PDF Parser (V3)

A comprehensive parser for PVsyst PDF reports that extracts structured data including arrays, orientations, inverter configurations, and monthly production data. Features both a command-line interface and a web-based interface for easy PDF analysis.

## Features

- **Complex Notation Parsing**: Handles sophisticated inverter/MPPT configurations like "INV02-05, 7,8 MPPT 1-5"
- **Monthly Production**: Estimates monthly energy production per inverter based on module counts *(as an average; future updates will include azimuth & tilt variances)*
- **Structured Output**: Generates clean JSON and text reports with separated array configurations and associations
- **Web Interface**: Upload PDFs through a modern web interface
- **Cross-Version Compatibility**: Works with different PVsyst versions (V7, V7.4, V8.x)
- **Single-Array Support**: Special handling for reports with uniform configurations and round-robin string allocation
- **Validation**: Cross-validates parsed data against authoritative "Total Inverter Power" sections
- **Enhanced Debugging**: Detailed warnings for parsing edge cases with array block inspection

## Installation

### Prerequisites

- Python 3.7+
- pip

### Install Dependencies

```bash
pip install pdfplumber fastapi uvicorn
```

**Note**: V3 uses text-only parsing with pdfplumber for faster, more reliable extraction. No table extraction dependencies required.

## Usage

### Command Line Interface

Parse a PVsyst PDF and generate reports:

```bash
python pvsyst_parser.py "path/to/your/pvsyst_report.pdf" --output-dir "./output"
```

This will generate:
- `*_analysis_v3.txt`: Comprehensive text report
- `*_structured_v3.json`: Structured JSON data with V3 schema

**V3 Features:**
- Faster text-only parsing (no table extraction)
- Monitoring-friendly JSON structure
- Round-robin string allocation for single arrays
- Validation against "Total Inverter Power" sections

### Web Interface

Start the web server:

```bash
uvicorn app:app --reload --host 0.0.0.0 --port 8888
```

Open your browser to `http://localhost:8888` and upload a PVsyst PDF through the web interface.

**V3 Interface:** Modern UI that displays combined MPPT configurations and detects all inverter types.

### API Usage

The FastAPI backend provides endpoints:

- `POST /api/parse`: Upload and parse a PDF
- `GET /api/health`: Health check

Example API call:

```bash
curl -X POST "http://localhost:8000/api/parse" \
     -H "accept: application/json" \
     -H "Content-Type: multipart/form-data" \
     -F "file=@your_pvsyst_report.pdf"
```

## Output Structure

### JSON Output

```json
{
  "metadata": {
    "total_arrays": 3,
    "total_inverters": 2,
    "total_system_capacity_kwp": 150.5,
    "total_annual_production_kwh": 225000
  },
  "pv_module": {
    "manufacturer": "Hanwha Q Cells",
    "model": "Q.Peak-Duo-XL-G11S.3",
    "unit_nom_power_w": 595
  },
  "inverter": {
    "manufacturer": "SMA",
    "model": "Sunny Tripower_Core1 62-US-41",
    "unit_nom_power_kw": 62.5
  },
  "array_configurations": {
    "1": {
      "array_id": "1",
      "inverter_ids": [
        "INV01"
      ],
      "inverter_id": "INV01",
      "mppt_count": 1,
      "mppt_share_percent": 35.0,
      "inverter_unit_fraction": 0.3,
      "number_of_modules": 51,
      "nominal_stc_kwp_from_module": 27.795,
      "nominal_stc_kwp": 27.8,
      "strings": 3,
      "modules_in_series": 17,
      "u_mpp_v": 646.0,
      "i_mpp_a": 39.0,
      "orientation_id": 1,
      "tilt": 9.0,
      "azimuth_pvsyst_deg": 0.0,
      "azimuth_deg": 180.0,
      "azimuth_compass_deg": 180.0
    }
  },
  "associations": {
    "INV01": {
      "MPPT 1": {
        "array_id": "1",
        "strings": 3,
        "modules": 51,
        "dc_kwp": 27.795
      }
    }
  },
  "inverter_summary": {
    "INV01": {
      "capacity_kwp": 80.1,
      "annual_production_kwh": 130246.0,
      "specific_production_kwh_per_kwp": 1626.0,
      "monthly_production": {
        "January": 7717.0,
        "February": 8528.0,
        "March": 11929.0,
        "April": 13186.0,
        "May": 13777.0,
        "June": 13123.0,
        "July": 14111.0,
        "August": 13440.0,
        "September": 11136.0,
        "October": 9208.0,
        "November": 7144.0,
        "December": 6947.0
      }
    }
  }
}
```

## V3 Features & Enhancements

### What's New in V3

- **Text-Only Parsing**: Uses pdfplumber only for faster, more reliable extraction (no Camelot table dependencies)
- **Monitoring-Friendly Output**: Per-inverter `combined_configuration` array consolidates MPPT allocation with config fields
- **Single-Configuration Support**: Handles reports with uniform arrays using round-robin string distribution
- **Industry Heuristics**: Automatic MPPT topology detection for common inverters (SMA Core1: 6 MPPT, CHINT/CPS: 3 MPPT)
- **Validation System**: Cross-validates parsed inverter counts against "Total Inverter Power" sections with debugging output
- **Enhanced Debugging**: Detailed warnings for parsing mismatches with array block inspection

### V3 JSON Schema

```json
{
  "metadata": {
    "version": "v3",
    "total_arrays": 1,
    "total_inverters": 12,
    "total_expanded_combinations": 72
  },
  "array_configurations": {
    "1": {
      "config_id": "1",
      "inverter_ids": ["INV01", "INV02", ..., "INV12"],
      "mppt_ids": ["MPPT 1", "MPPT 2", ..., "MPPT 6"],
      "inferred_single_config": true,
      "inferred_mppt_per_inverter": 6,
      "inferred_strings_per_mppt_max": 2
    }
  },
  "associations": {
    "INV01": {
      "MPPT 1": {"config_id": "1", "strings": 2, "modules": 34, "dc_kwp": 18.36},
      "MPPT 2": {"config_id": "1", "strings": 2, ...}
    }
  },
  "inverter_summary": {
    "INV01": {
      "description": "Inv 01 - (62.5 kW) - SMA Sunny Tripower_Core1 62-US-41",
      "combined_configuration": [
        {
          "mppt": "MPPT 1",
          "config_id": "1",
          "strings": 2,
          "modules": 34,
          "dc_kwp": 18.36,
          "tilt": 9.0,
          "azimuth": 205.0,
          "modules_in_series": 17
        }
      ]
    }
  }
}
```

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

## Key Capabilities

### Inverter Range Parsing

Supports complex notation:
- `INV01`: Single inverter
- `INV02-05`: Range of inverters
- `INV02-05, 7,8`: Mixed ranges and singles
- `INV 9-11,13`: Space-separated ranges

### MPPT Configuration

Handles MPPT assignments:
- `MPPT 1-3`: Range of MPPTs
- `MPPT 1,2,4`: Specific MPPTs
- Automatic expansion of inverter × MPPT combinations

### Monthly Production Allocation

- Extracts system-level monthly production from PVsyst tables
- Allocates production to individual inverters based on module count ratios
- Provides per-inverter monthly energy estimates

## Dependencies

- **pdfplumber**: PDF text extraction (primary parsing engine)
- **fastapi**: Web API framework
- **uvicorn**: ASGI server

**V3 Changes**: Removed Camelot table extraction dependency for faster, more reliable text-only parsing.

## Development

### Project Structure

```
.
├── pvsyst_parser.py   # V3 core parsing logic with validation
├── app.py             # V3 FastAPI web application
├── index.html         # V3 web interface
├── requirements.txt   # Dependencies
└── README.md          # This file
```

### Adding New Features

The parser is modular and extensible. Key classes:

- `PVsystParser`: Main parser class
- Methods for extracting different sections (arrays, orientations, monthly data)
- Flexible text parsing that adapts to PVsyst version changes

## Troubleshooting

### Common Issues

1. **Table extraction fails**: Ensure camelot dependencies are installed with `[cv]` extra
2. **Text extraction issues**: Check that pdfplumber can read your PDF
3. **Web interface not loading**: Verify uvicorn is running and port 8000 is accessible

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