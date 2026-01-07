"""API for parsing PVsyst PDF files (V3).

This version is compatible with `pvsyst_parser.py` and returns the V3 JSON
structure without writing output files to disk.
"""

import tempfile
from pathlib import Path

from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse

from pvsyst_parser_v3 import PVsystParser

app = FastAPI(title="PVsyst Parser API (V3)")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/parse")
async def parse_pvsyst_pdf(file: UploadFile = File(...)):
    """Parse an uploaded PVsyst PDF file and return parsed data."""

    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    # Save upload to a temp file
    try:
        suffix = Path(file.filename).suffix
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            contents = await file.read()
            tmp.write(contents)
            tmp_path = tmp.name
    except Exception as e:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Failed to save file: {e}") from e

    parser = PVsystParser()

    try:
        # Run the V3 parsing pipeline without generating output files.
        blocks = parser.extract_text_blocks(tmp_path)

        parser.sections = parser.identify_sections(blocks)
        parser.section_contents = parser.extract_section_contents(blocks, parser.sections)

        parser.extract_equipment_info(blocks)
        parser.orientations = parser.extract_orientations(blocks)

        if "Array Losses" in parser.section_contents and parser.section_contents["Array Losses"]:
            try:
                parser.array_losses = parser.parse_array_losses_section(
                    parser.section_contents["Array Losses"][0]
                )
            except Exception:
                parser.array_losses = {}

        parser.arrays = parser.parse_arrays_from_text(blocks, interactive=False)
        parser.inverter_types = parser._collect_inverter_types()
        parser.calculate_monthly_production(blocks)

        data = parser.to_dict()

    except Exception as e:  # noqa: BLE001
        raise HTTPException(status_code=500, detail=f"Parsing failed: {e}") from e
    finally:
        try:
            Path(tmp_path).unlink(missing_ok=True)
        except OSError:
            pass

    return JSONResponse(content=data)


@app.get("/api/health")
def health():
    return {"status": "ok"}
