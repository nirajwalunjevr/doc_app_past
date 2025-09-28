from __future__ import annotations
from io import BytesIO
from pathlib import Path
from typing import Optional, Dict, Any, List
from openpyxl import Workbook

from .excel_generator_fn import generate_catalog_excel_fn
from .front_page_fn import add_front_page_fn
from .back_page_fn import add_back_page_fn

from firebase_admin import firestore  # initialized in main.py


def generate_catalog_report(
    class_no: str | int,
    division: str,
    *,
    save_path: Optional[Path | str] = None,
    return_bytes: bool = True,
    assets_dir: Optional[Path] = None,
    # NEW: optional historical override
    selected_month: Optional[int] = None,   # 1..12
    selected_year: Optional[int] = None,    # e.g., 2025
) -> Dict[str, Any]:
    """
    Generates the catalog workbook. If selected_month/year are provided,
    the Front Page will render that month/year; otherwise it uses current month/year.
    """
    try:
        db = firestore.client()
        class_division_str = f"{class_no}-{division.upper()}"

        # Fetch catalog meta for class/division
        doc = db.collection('catalog').document(class_division_str).get()
        teacher_name = "N/A"; month = 1; year = 2025
        subjects: List[Dict[str, Any]] = []; doc_data: Dict[str, Any] = {}
        if doc.exists:
            doc_data = doc.to_dict() or {}
            teacher_name = doc_data.get('classTeacher', 'N/A')
            month = doc_data.get('month', 1)
            year = doc_data.get('year', 2025)
            subjects = doc_data.get('subjects', []) or []

        # Class/Division Marathi mapping
        class_map_mr = {
            "1": "१ ली", "2": "२ री", "3": "३ री", "4": "४ थी", "5": "५ वी",
            "6": "६ वी", "7": "७ वी", "8": "८ वी", "9": "९ वी", "10": "१० वी"
        }
        class_name_mr = class_map_mr.get(str(class_no), str(class_no))
        division_map_mr = {"A": "अ", "B": "ब", "C": "क", "D": "ड"}
        division_name_mr = division_map_mr.get(division.upper(), division.upper())

        # Prepare report_data. Note: month/year kept for compatibility.
        report_data = {
            "teacher_name": teacher_name,
            "month": month,
            "year": year,
            "class_name_mr": class_name_mr,
            "division_name_mr": division_name_mr,
            "division": division.upper(),
            # NEW: optional historical override for Front Page header
            "selected_month": selected_month,
            "selected_year": selected_year,
        }

        # Students for this class/division (active)
        students_ref = db.collection('catalog/global/students')
        query = students_ref.where('status', '==', 'active').where('classDivision', '==', class_division_str)
        docs = sorted(list(query.stream()), key=lambda x: x.to_dict().get('rollNo', 0))
        students = [d.to_dict() for d in docs]
        if not students:
            return {"ok": False, "error": f"No active students found for {class_division_str}."}

        # Build workbook
        wb: Workbook = generate_catalog_excel_fn(class_no, division, students)
        wb.active.title = "Catalog"
        wb = add_front_page_fn(wb, report_data, assets_dir=assets_dir)
        wb = add_back_page_fn(wb, class_no=str(class_no), division=division.upper(), subjects=subjects, catalog_doc=doc_data)

        # Show all worksheets in page layout view
        for ws in wb.worksheets:
            ws.sheet_view.view = "pageLayout"

        # Output
        out: Dict[str, Any] = {"ok": True, "bytes": None, "path": None, "error": None}
        if save_path:
            save_path = str(save_path)
            Path(save_path).parent.mkdir(parents=True, exist_ok=True)
            wb.save(save_path)
            out["path"] = save_path
        if return_bytes:
            buf = BytesIO()
            wb.save(buf)
            out["bytes"] = buf.getvalue()
            buf.close()
        return out
    except Exception as e:
        return {"ok": False, "error": str(e), "bytes": None, "path": None}
