# PVsyst PDF Parser

A comprehensive parser for PVsyst PDF reports that extracts structured data including arrays, orientations, inverter configurations, and monthly production data. Features both a command-line interface and a web-based interface for easy PDF analysis.

## Features

- **Complex Notation Parsing**: Handles sophisticated inverter/MPPT configurations like "INV02-05, 7,8 MPPT 1-5"
- **Monthly Production**: Estimates monthly energy production per inverter based on module counts *(as an average; future updates will include azimuth & tilt variances)*
- **Structured Output**: Generates clean JSON and text reports with separated array configurations and associations
- **Web Interface**: Upload PDFs through a modern web interface
- **Cross-Version Compatibility**: Works with different PVsyst versions (V7, V7.4, V8.x)
- **Table Extraction**: Uses camelot for accurate table parsing from PDF reports

## Installation

### Prerequisites

- Python 3.7+
- pip

### Install Dependencies

```bash
pip install camelot-py[cv] pdfplumber fastapi uvicorn
```

**Note**: `camelot-py[cv]` includes OpenCV for better table detection. On some systems, you may need additional dependencies:

```bash
# Ubuntu/Debian
sudo apt-get install python3-tk ghostscript

# macOS
brew install ghostscript tcl-tk
```

## Usage

### Command Line Interface

Parse a PVsyst PDF and generate reports:

```bash
python pvsyst_parser.py "path/to/your/pvsyst_report.pdf"
```

Generate an additional PowerTrack patch JSON (per inverter):

```bash
python3 pvsyst_parser.py "path/to/your/pvsyst_report.pdf" --powertrack-patch
```

This writes `<pdf_stem>_powertrack_patch.json` alongside the normal outputs. The JSON is keyed as `PV0`, `PV1`, ... (derived from `INV01` -> `PV0`, `INV02` -> `PV1`, etc).

Optional: specify output directory:

```bash
python pvsyst_parser.py "report.pdf" "/path/to/output/dir"
```

Optional: specify PowerTrack patch output path:

```bash
python3 pvsyst_parser.py "report.pdf" --output-dir "/path/to/output/dir" --powertrack-patch --powertrack-patch-path "/path/to/output/dir/pt_patch.json"
```

This will generate:
- `report.txt`: Comprehensive text report
- `report.json`: Structured JSON data

### Web Interface

Start the web server:

```bash
uvicorn app:app --reload
```

Open your browser to `http://localhost:8000` and upload a PVsyst PDF through the web interface.

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

- **camelot-py**: PDF table extraction
- **pdfplumber**: PDF text extraction
- **fastapi**: Web API framework
- **uvicorn**: ASGI server
- **opencv-python**: Image processing for table detection

## Development

### Project Structure

```
.
├── pvsyst_parser.py   # Core parsing logic
├── app.py             # FastAPI web application
├── index.html         # Web interface
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
- May require adjustments for heavily customized reports (like array headers)

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
