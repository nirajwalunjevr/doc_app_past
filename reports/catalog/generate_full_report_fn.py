from __future__ import annotations
from io import BytesIO
from pathlib import Path
from typing import Optional, Dict, Any, List
from openpyxl import Workbook

from .excel_generator_fn import generate_catalog_excel_fn
from .front_page_fn import add_front_page_fn
from .back_page_fn import add_back_page_fn

from firebase_admin import firestore  # initialized in main.py


def _coerce_timestamp_to_datetime(v: Any):
    """
    Firestore Admin returns a Timestamp for 'dob' in studentsData.
    Convert to Python datetime if possible; otherwise return as-is.
    """
    try:
        # Firestore Timestamp has .to_datetime()
        if hasattr(v, "to_datetime"):
            return v.to_datetime()
        return v
    except Exception:
        return v


def generate_catalog_report(
    class_no: str | int,
    division: str,
    *,
    save_path: Optional[Path | str] = None,
    return_bytes: bool = True,
    assets_dir: Optional[Path] = None,
    selected_month: Optional[int] = None,   # 1..12
    selected_year: Optional[int] = None,    # e.g., 2025
) -> Dict[str, Any]:
    """
    Generates the catalog workbook.

    Historical mode:
      - If selected_month/year are provided, load historical snapshot from:
          roster_records/{classNo}-{DIV}_{YYYY}-{MM}
        and use studentsData array as the roster in that stored order.

    Live mode:
      - Otherwise, read active students for the classDivision.
    """
    try:
        db = firestore.client()
        class_division_str = f"{class_no}-{division.upper()}"

        # Catalog meta for front/back pages (class teacher, subjects)
        doc = db.collection('catalog').document(class_division_str).get()
        teacher_name = "N/A"; month = 1; year = 2025
        subjects: List[Dict[str, Any]] = []; doc_data: Dict[str, Any] = {}
        if doc.exists:
            doc_data = doc.to_dict() or {}
            teacher_name = doc_data.get('classTeacher', 'N/A')
            month = doc_data.get('month', 1)
            year = doc_data.get('year', 2025)
            subjects = doc_data.get('subjects', []) or []

        # Marathi mappings for Front Page labeling
        class_map_mr = {
            "1": "१ ली", "2": "२ री", "3": "३ री", "4": "४ थी", "5": "५ वी",
            "6": "६ वी", "7": "७ वी", "8": "८ वी", "9": "९ वी", "10": "१० वी"
        }
        class_name_mr = class_map_mr.get(str(class_no), str(class_no))
        division_map_mr = {"A": "अ", "B": "ब", "C": "क", "D": "ड"}
        division_name_mr = division_map_mr.get(division.upper(), division.upper())

        report_data = {
            "teacher_name": teacher_name,
            "month": month,
            "year": year,
            "class_name_mr": class_name_mr,
            "division_name_mr": division_name_mr,
            "division": division.upper(),
            "selected_month": selected_month,
            "selected_year": selected_year,
        }

        students: List[Dict[str, Any]] = []

        if isinstance(selected_month, int) and 1 <= selected_month <= 12 and isinstance(selected_year, int) and selected_year > 0:
            # Historical mode: read frozen snapshot
            mm = str(selected_month).zfill(2)
            record_id = f"{class_division_str}_{selected_year}-{mm}"
            rec_doc = db.collection('roster_records').document(record_id).get()
            if not rec_doc.exists:
                return {
                    "ok": False,
                    "error": f"No historical roster found: roster_records/{record_id}",
                    "bytes": None,
                    "path": None,
                }
            rec = rec_doc.to_dict() or {}
            data_list = rec.get("studentsData", [])
            if not isinstance(data_list, list) or not data_list:
                return {
                    "ok": False,
                    "error": f"Historical roster for {record_id} has no studentsData.",
                    "bytes": None,
                    "path": None,
                }

            # Preserve array order and coerce timestamps
            students = []
            for item in data_list:
                if not isinstance(item, dict):
                    continue
                s = dict(item)
                if "dob" in s:
                    s["dob"] = _coerce_timestamp_to_datetime(s["dob"])
                students.append(s)
        else:
            # Live mode: query active students for this classDivision
            students_ref = db.collection('catalog/global/students')
            query = students_ref.where('status', '==', 'active').where('classDivision', '==', class_division_str)
            docs = list(query.stream())
            # Sort by rollNo if present
            docs.sort(key=lambda x: (x.to_dict().get('rollNo', 10_000_000)))
            students = [d.to_dict() for d in docs]

        if not students:
            return {"ok": False, "error": f"No students to print for {class_division_str}.", "bytes": None, "path": None}

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
