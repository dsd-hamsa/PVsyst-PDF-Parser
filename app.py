from fastapi import FastAPI, UploadFile, File, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from fastapi.responses import JSONResponse
import tempfile
from pathlib import Path

from pvsyst_parser import PVsystParser  # your script

app = FastAPI(title="PVsyst Parser API")

# Allow JS frontends (e.g., GitHub Pages) to talk to this API
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],  # tighten if you want
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)


@app.post("/api/parse")
async def parse_pvsyst_pdf(file: UploadFile = File(...)):
    if not file.filename or not file.filename.lower().endswith(".pdf"):
        raise HTTPException(status_code=400, detail="Please upload a PDF file.")

    # Save upload to a temp file
    try:
        suffix = Path(file.filename).suffix
        with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
            contents = await file.read()
            tmp.write(contents)
            tmp_path = tmp.name
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Failed to save file: {e}")

    # Run your parser
    parser = PVsystParser()
    try:
        data = parser.parse_pdf(tmp_path, generate_outputs=False)
    except Exception as e:
        raise HTTPException(status_code=500, detail=f"Parsing failed: {e}")
    finally:
        # clean up uploaded file
        try:
            Path(tmp_path).unlink(missing_ok=True)
        except Exception:
            pass

    return JSONResponse(content=data)


@app.get("/api/health")
def health():
    return {"status": "ok"}
