# main.py
from __future__ import annotations

import os
import json
from pathlib import Path
from typing import Optional

from fastapi import FastAPI, Body, Response, HTTPException, Query
from fastapi.middleware.cors import CORSMiddleware
from firebase_admin import credentials, initialize_app

from reports.catalog.generate_full_report_fn import generate_catalog_report

# ---------------------------------------------------------------------
# Firebase Admin initialization (from environment variable)
# ---------------------------------------------------------------------
FIREBASE_CREDENTIALS_JSON = os.environ.get("FIREBASE_KEY")
if not FIREBASE_CREDENTIALS_JSON:
    raise RuntimeError(
        "FIREBASE_KEY env var not set. Paste your service account JSON here."
    )

# Load credentials from JSON string
cred_dict = json.loads(FIREBASE_CREDENTIALS_JSON)
cred = credentials.Certificate(cred_dict)
initialize_app(cred)

# ---------------------------------------------------------------------
# FastAPI app
# ---------------------------------------------------------------------
app = FastAPI(title="Catalog Report Service")

# CORS for Flutter Web and other clients
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # for development; restrict in production
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
# Current-month generator (existing)
# POST body example: { "class_no": 10, "division": "A", "return_inline": false }
# ---------------------------------------------------------------------
@app.post("/generate")
def generate_endpoint(
    class_no: int = Body(..., embed=True),
    division: str = Body(..., embed=True),
    return_inline: Optional[bool] = Body(False, embed=True),
    selected_month: Optional[int] = Body(None, embed=True),  # optional pass-through
    selected_year: Optional[int] = Body(None, embed=True),   # optional pass-through
):
    """
    Generate the full catalog (Front, Catalog, Back) and return an .xlsx file.
    If selected_month and selected_year are provided, the Front Page will display that month/year;
    otherwise it uses the server's current month/year.
    """
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

    data: bytes = result["bytes"]  # type: ignore[assignment]
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
# Historical generator (for url_launcher GET in Flutter)
# URL shape:
#   /generate-historical-report?class=10&div=A&year=2025&month=8
# ---------------------------------------------------------------------
@app.get("/generate-historical-report")
def generate_historical_report(
    class_no: int = Query(..., alias="class"),
    division: str = Query(..., alias="div"),
    year: int = Query(..., ge=1900),
    month: int = Query(..., ge=1, le=12),
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
        selected_month=month,
        selected_year=year,
    )
    if not result.get("ok") or not result.get("bytes"):
        raise HTTPException(status_code=400, detail=result.get("error", "Unknown error"))

    data: bytes = result["bytes"]  # type: ignore[assignment]
    filename = f"catalog_{class_no}-{div}_{year}-{str(month).zfill(2)}.xlsx"
    headers = {"Content-Disposition": f'attachment; filename="{filename}"'}
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
