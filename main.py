# main.py
from __future__ import annotations

import os
import json
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Body, Response, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from firebase_admin import credentials, initialize_app

from reports.catalog.generate_full_report_fn import generate_catalog_report

# ---------------------------------------------------------------------
# Firebase Admin initialization
# ---------------------------------------------------------------------
FIREBASE_CREDENTIALS_JSON = os.environ.get("FIREBASE_KEY")
if not FIREBASE_CREDENTIALS_JSON:
    raise RuntimeError(
        "FIREBASE_KEY env var not set. Paste your service account JSON here."
    )

cred_dict = json.loads(FIREBASE_CREDENTIALS_JSON)
cred = credentials.Certificate(cred_dict)
initialize_app(cred)

# ---------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------
app = FastAPI(title="Catalog Report Service")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# ---------------------------------------------------------------------
# Health
# ---------------------------------------------------------------------
@app.get("/health")
def health():
    return {"ok": True}

# ---------------------------------------------------------------------
# Current-month generator (This endpoint is unchanged)
# ---------------------------------------------------------------------
@app.post("/generate")
def generate_endpoint(
    class_no: int = Body(..., embed=True),
    division: str = Body(..., embed=True),
    return_inline: Optional[bool] = Body(False, embed=True),
    selected_month: Optional[int] = Body(None, embed=True),
    selected_year: Optional[int] = Body(None, embed=True),
):
    assets_dir = Path("assets")
    div = (division or "").strip().upper()
    if not div:
        raise HTTPException(status_code=400, detail="division is required")

    result = generate_catalog_report(
        class_no=class_no,
        division=div,
        return_bytes=True,
        assets_dir=assets_dir,
        selected_month=selected_month,
        selected_year=selected_year,
    )
    if not result.get("ok"):
        raise HTTPException(status_code=400, detail=result.get("error", "Unknown error"))

    data: bytes = result["bytes"]
    disp_type = "inline" if return_inline else "attachment"
    suffix = ""
    if isinstance(selected_year, int) and isinstance(selected_month, int):
        suffix = f"_{selected_year}-{str(selected_month).zfill(2)}"
    filename = f"catalog_{class_no}-{div}{suffix}.xlsx"
    headers = {"Content-Disposition": f'{disp_type}; filename="{filename}"'}
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )

# ---------------------------------------------------------------------
# --- CHANGE: Historical generator updated to accept POST requests ---
# ---------------------------------------------------------------------
@app.post("/generate-historical-report") # Changed from @app.get to @app.post
def generate_historical_report(
    # Changed from Query to Body to read JSON from the request
    class_no: int = Body(...),
    division: str = Body(...),
    selected_year: int = Body(...),
    selected_month: int = Body(...),
):
    assets_dir = Path("assets")
    div = (division or "").strip().upper()
    if not div:
        raise HTTPException(status_code=400, detail="division is required")

    result = generate_catalog_report(
        class_no=class_no,
        division=div,
        return_bytes=True,
        assets_dir=assets_dir,
        selected_month=selected_month, # Use the variables from the Body
        selected_year=selected_year,   # Use the variables from the Body
    )
    if not result.get("ok") or not result.get("bytes"):
        raise HTTPException(status_code=400, detail=result.get("error", "Unknown error"))

    data: bytes = result["bytes"]
    filename = f"catalog_{class_no}-{div}_{selected_year}-{str(selected_month).zfill(2)}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )