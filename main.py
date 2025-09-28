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

# Firebase Admin initialization (from environment variable)
FIREBASE_CREDENTIALS_JSON = os.environ.get("FIREBASE_KEY")
if not FIREBASE_CREDENTIALS_JSON:
    raise RuntimeError(
        "FIREBASE_KEY env var not set. Paste your service account JSON here."
    )

# Load credentials from JSON string
cred_dict = json.loads(FIREBASE_CREDENTIALS_JSON)
cred = credentials.Certificate(cred_dict)
initialize_app(cred)

app = FastAPI(title="Catalog Report Service")

# CORS for Flutter Web and other clients; tighten origins later if needed
app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],          # for development; replace with explicit origins in production
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

@app.get("/health")
def health():
    return {"ok": True}

@app.post("/generate")
def generate_endpoint(
    class_no: int = Body(..., embed=True),
    division: str = Body(..., embed=True),
    return_inline: Optional[bool] = Body(False, embed=True),
):
    """
    Generate the full catalog (Front, Catalog, Back) and return an .xlsx file.
    Body example: { "class_no": 10, "division": "A" }
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
    )
    if not result.get("ok"):
        raise HTTPException(status_code=400, detail=result.get("error", "Unknown error"))

    data: bytes = result["bytes"]
    disp_type = "inline" if return_inline else "attachment"
    filename = f"catalog_{class_no}-{div}.xlsx"
    headers = {"Content-Disposition": f'{disp_type}; filename="{filename}"'}
    return Response(
        content=data,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        headers=headers,
    )
