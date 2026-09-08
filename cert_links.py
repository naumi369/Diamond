"""Official lab report lookup / PDF helpers.

GIA: public Report Check page (PDF is offered on that page; bulk API needs GIA account).
IGI: public viewpdf endpoint is commonly available for many reports.
"""

from __future__ import annotations

from typing import Any, Optional, Tuple
from urllib.parse import quote

import pandas as pd


def _clean_id(val: Any) -> Optional[str]:
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(val, float) and val == int(val):
        s = str(int(val))
    elif isinstance(val, int) and not isinstance(val, bool):
        s = str(val)
    else:
        s = str(val).strip()
    if s.lower() in ("", "nan", "none", "null", "<na>"):
        return None
    if s.endswith(".0") and s.replace(".", "", 1).isdigit():
        s = s[:-2]
    return s


def detect_lab(lab: Any, cert: Any) -> str:
    """Return GIA / IGI / OTHER."""
    lab_s = (str(lab) if lab is not None else "").strip().upper()
    cert_s = _clean_id(cert) or ""
    if "IGI" in lab_s:
        return "IGI"
    if "GIA" in lab_s:
        return "GIA"
    # Heuristic: many IGI lab-grown reports start with LG
    if cert_s.upper().startswith("LG"):
        return "IGI"
    if cert_s.isdigit() and len(cert_s) >= 8:
        return "GIA"  # most numeric long IDs in this inventory are GIA
    return "OTHER"


def gia_report_check_url(cert: Any) -> Optional[str]:
    c = _clean_id(cert)
    if not c:
        return None
    return f"https://www.gia.edu/report-check?reportno={quote(c)}"


def igi_report_pdf_url(cert: Any) -> Optional[str]:
    """IGI public PDF viewer URL (works for many IGI reports)."""
    c = _clean_id(cert)
    if not c:
        return None
    return f"https://api.igi.org/viewpdf.php?r={quote(c)}"


def igi_verify_url(cert: Any) -> Optional[str]:
    c = _clean_id(cert)
    if not c:
        return None
    # Public verify landing; user can enter number if needed
    return f"https://www.igi.org/verify-your-report/?r={quote(c)}"


def official_links(lab: Any, cert: Any) -> Tuple[Optional[str], Optional[str], str]:
    """
    Returns (open_url, pdf_url_or_none, note).
    open_url: page to view/verify the report
    pdf_url: direct PDF when publicly available (mainly IGI)
    """
    c = _clean_id(cert)
    if not c:
        return None, None, "No certificate number"
    kind = detect_lab(lab, c)
    if kind == "GIA":
        return (
            gia_report_check_url(c),
            None,  # GIA PDF is on Report Check page / Report Check Plus — not a stable public direct link
            "Open GIA Report Check — use Download PDF on their site",
        )
    if kind == "IGI":
        return (
            igi_verify_url(c),
            igi_report_pdf_url(c),
            "IGI verify page + direct PDF link when available",
        )
    # Unknown lab: try both styles
    return (
        gia_report_check_url(c),
        igi_report_pdf_url(c),
        "Lab unknown — tried GIA Report Check and IGI PDF link",
    )
