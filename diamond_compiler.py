#!/usr/bin/env python3
"""
Diamond Inventory Compiler
--------------------------
Analyzes miscellaneous diamond Excel lists from different suppliers/customers
and normalizes them into a single standardized format.

Standard columns:
  stock_no, certificate_no, lab, shape, weight, mm_size, color, clarity,
  polish, symmetry, fluorescence, cut, price_per_ct, amount,
  video_link, image_link, certificate_link, source_file, notes

Usage:
  python diamond_compiler.py [--input-dir DIR] [--output-dir DIR] [--db] [--download-videos]
"""

import argparse
import hashlib
import logging
import os
import re
import sqlite3
import sys
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple
from urllib.parse import urlparse

import pandas as pd
import requests

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

STANDARD_COLUMNS = [
    "stock_no",
    "certificate_no",
    "lab",
    "shape",
    "weight",
    "mm_size",
    "color",
    "clarity",
    "polish",
    "symmetry",
    "fluorescence",
    "cut",
    "price_per_ct",
    "amount",
    "video_link",
    "image_link",
    "certificate_link",
    "source_file",
    "notes",
]

# Common column name mappings (lowercased source -> standard)
COLUMN_ALIASES = {
    # Stock / Ref
    "stock_no": "stock_no",
    "stock#": "stock_no",
    "stock no": "stock_no",
    "stone no": "stock_no",
    "ref.no": "stock_no",
    "ref.no.": "stock_no",
    "ref no": "stock_no",
    "ref": "stock_no",
    "srno.": "stock_no",
    "sr no": "stock_no",
    # Certificate
    "certificate_no": "certificate_no",
    "certificate no": "certificate_no",
    "certno": "certificate_no",
    "cert.no": "certificate_no",
    "cert no": "certificate_no",
    "report no": "certificate_no",
    "report#": "certificate_no",
    "cert": "certificate_no",
    "certificate no.": "certificate_no",
    "certificate_no.": "certificate_no",
    # Lab
    "lab": "lab",
    "certificate": "lab",  # often "GIA"
    "cert.": "lab",
    # Shape
    "shape": "shape",
    # Weight
    "weight": "weight",
    "cts.": "weight",
    "cts": "weight",
    "carats": "weight",
    "ct": "weight",
    # MM size
    "mm_size": "mm_size",
    "mm size": "mm_size",
    "measurement": "mm_size",
    "measurements": "mm_size",
    "size": "mm_size",
    # Color
    "color": "color",
    "colour": "color",
    # Clarity
    "clarity": "clarity",
    # Polish
    "polish": "polish",
    "pol.": "polish",
    "pol": "polish",
    # Symmetry
    "symmetry": "symmetry",
    "sym.": "symmetry",
    "sym": "symmetry",
    # Fluorescence
    "fluorescence": "fluorescence",
    "flour": "fluorescence",
    "fl": "fluorescence",
    "fluo": "fluorescence",
    # Cut
    "cut": "cut",
    # Price
    "price_per_ct": "price_per_ct",
    "price/ct.": "price_per_ct",
    "price/ct": "price_per_ct",
    "price ct": "price_per_ct",
    "rate": "price_per_ct",
    "p.ct": "price_per_ct",
    "avg p.ct": "price_per_ct",
    # Amount
    "amount": "amount",
    "total vl": "amount",
    "value": "amount",
    # Links
    "video_link": "video_link",
    "video link": "video_link",
    "video": "video_link",
    "image_link": "image_link",
    "image link": "image_link",
    "image": "image_link",
    "certificate_link": "certificate_link",
    "certificate link": "certificate_link",
    "cert link": "certificate_link",
}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
log = logging.getLogger("diamond_compiler")


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def normalize_col(name: Any) -> str:
    """Normalize a column header for matching."""
    if name is None or (isinstance(name, float) and pd.isna(name)):
        return ""
    s = str(name).strip().lower()
    s = re.sub(r"\s+", " ", s)
    return s


def clean_value(val: Any) -> Optional[str]:
    """Convert cell value to clean string or None.

    Certificate / stock style IDs never keep a trailing '.0'.
    """
    if val is None:
        return None
    try:
        if pd.isna(val):
            return None
    except (TypeError, ValueError):
        pass

    # Integers / whole floats → plain digit string (no .0)
    if isinstance(val, bool):
        s = str(val)
    elif isinstance(val, int) and not isinstance(val, bool):
        s = str(val)
    elif isinstance(val, float):
        if val == int(val):
            s = str(int(val))
        else:
            s = str(val)
    else:
        s = str(val).strip()

    if s.lower() in ("", "nan", "none", "null", "-", "<na>", "nat"):
        return None

    # "12345.0", "12345.000", scientific-ish whole numbers
    if re.match(r"^-?\d+\.0+$", s):
        s = s.split(".")[0]
    # Excel sometimes gives "1.535750439e+09"
    if re.match(r"^-?\d+(\.\d+)?[eE][+]\d+$", s):
        try:
            f = float(s)
            if f == int(f):
                s = str(int(f))
        except ValueError:
            pass
    return s


def to_float(val: Any) -> Optional[float]:
    """Try to parse a number."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    try:
        s = str(val).replace(",", "").strip()
        return float(s)
    except (ValueError, TypeError):
        return None


def is_url(s: Any) -> bool:
    """Return True only for real http(s) URL strings."""
    if s is None:
        return False
    # pandas NaN / NA
    try:
        if pd.isna(s):
            return False
    except (TypeError, ValueError):
        pass
    if not isinstance(s, str):
        s = str(s).strip()
    else:
        s = s.strip()
    if not s or s.lower() in ("nan", "none", "null", "<na>", "nat"):
        return False
    return s.startswith("http://") or s.startswith("https://")


def normalize_shape(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = s.upper().strip()
    mapping = {
        "OVAL": "Oval",
        "RBC": "Round",
        "ROUND": "Round",
        "BRILLIANT": "Round",
        "PEAR": "Pear",
        "CUSHION": "Cushion",
        "EMERALD": "Emerald",
        "PRINCESS": "Princess",
        "MARQUISE": "Marquise",
        "RADIANT": "Radiant",
        "HEART": "Heart",
        "ASSCHER": "Asscher",
    }
    return mapping.get(s, s.title())


def normalize_fluorescence(s: Optional[str]) -> Optional[str]:
    if not s:
        return None
    s = s.upper().strip()
    if s in ("N", "NON", "NONE", "NIL"):
        return "None"
    if s in ("F", "FAINT"):
        return "Faint"
    if s in ("M", "MED", "MEDIUM"):
        return "Medium"
    if s in ("S", "STG", "STRONG"):
        return "Strong"
    if s in ("VS", "VST", "VERY STRONG"):
        return "Very Strong"
    return s.title()


# ---------------------------------------------------------------------------
# Parsers for different file layouts
# ---------------------------------------------------------------------------

def detect_header_row(df: pd.DataFrame, max_scan: int = 15) -> int:
    """Find the most likely header row by looking for known column names."""
    known = set(COLUMN_ALIASES.keys())
    best_row, best_score = 0, 0
    for i in range(min(max_scan, len(df))):
        row_vals = [normalize_col(v) for v in df.iloc[i].tolist()]
        score = sum(1 for v in row_vals if v in known)
        if score > best_score:
            best_score = score
            best_row = i
    return best_row if best_score >= 3 else 0


def map_columns(headers: List[str]) -> Dict[int, str]:
    """Map column index -> standard name."""
    mapping = {}
    used = set()
    for idx, h in enumerate(headers):
        norm = normalize_col(h)
        if norm in COLUMN_ALIASES:
            std = COLUMN_ALIASES[norm]
            if std not in used:
                mapping[idx] = std
                used.add(std)
    return mapping


def parse_generic(df_raw: pd.DataFrame, source: str) -> List[Dict]:
    """Generic parser that auto-detects headers and maps columns."""
    header_row = detect_header_row(df_raw)
    headers = [str(h) if not pd.isna(h) else f"col_{i}" for i, h in enumerate(df_raw.iloc[header_row])]
    col_map = map_columns(headers)

    records = []
    for i in range(header_row + 1, len(df_raw)):
        row = df_raw.iloc[i]
        # Skip completely empty rows
        if row.isna().all():
            continue

        rec = {c: None for c in STANDARD_COLUMNS}
        rec["source_file"] = source

        for col_idx, std_name in col_map.items():
            if col_idx >= len(row):
                continue
            val = row.iloc[col_idx]
            if std_name in ("weight", "price_per_ct", "amount"):
                rec[std_name] = to_float(val)
            else:
                rec[std_name] = clean_value(val)

        # Post-process
        rec["shape"] = normalize_shape(rec.get("shape"))
        rec["fluorescence"] = normalize_fluorescence(rec.get("fluorescence"))

        # If lab looks like a certificate number, swap
        if rec.get("lab") and rec["lab"].isdigit() and not rec.get("certificate_no"):
            rec["certificate_no"] = rec["lab"]
            rec["lab"] = "GIA"  # common default

        # Skip rows that have almost no useful data
        if not any([rec.get("weight"), rec.get("certificate_no"), rec.get("stock_no"), rec.get("color")]):
            continue

        records.append(rec)
    return records


def parse_ovals_g_vs2(path: Path) -> List[Dict]:
    """Specific parser for 'Ovals G VS2-SI1.xlsx' style."""
    df = pd.read_excel(path, header=0)
    records = []
    for _, row in df.iterrows():
        rec = {c: None for c in STANDARD_COLUMNS}
        rec["source_file"] = path.name
        rec["stock_no"] = clean_value(row.get("Ref.No"))
        rec["certificate_no"] = clean_value(row.get("CertNo"))
        rec["lab"] = clean_value(row.get("Cert.")) or "GIA"
        rec["shape"] = "Oval"  # filename indicates ovals
        rec["weight"] = to_float(row.get("Cts."))
        rec["mm_size"] = clean_value(row.get("Measurement"))
        rec["color"] = clean_value(row.get("Color"))
        rec["clarity"] = clean_value(row.get("Clarity"))
        rec["polish"] = clean_value(row.get("Pol."))
        rec["symmetry"] = clean_value(row.get("Sym."))
        rec["fluorescence"] = normalize_fluorescence(clean_value(row.get("FL")))
        rec["cut"] = clean_value(row.get("Cut"))
        rec["price_per_ct"] = to_float(row.get("Rate"))
        rec["amount"] = to_float(row.get("Amount"))
        # Image/Video placeholders often "DNA"
        img = clean_value(row.get("Image"))
        if img and is_url(img):
            rec["image_link"] = img
        elif img and img.upper() != "DNA":
            rec["notes"] = f"Image: {img}"
        records.append(rec)
    return records


def parse_simple_customer_list(path: Path) -> List[Dict]:
    """Parser for the common 'Carat Size Range / Certificate / Shape / WEIGHT ...' layout."""
    df = pd.read_excel(path, header=0)
    # Normalize column names
    df.columns = [str(c).strip() for c in df.columns]

    records = []
    for _, row in df.iterrows():
        rec = {c: None for c in STANDARD_COLUMNS}
        rec["source_file"] = path.name

        # Map by known names
        for col in df.columns:
            norm = normalize_col(col)
            if norm in COLUMN_ALIASES:
                std = COLUMN_ALIASES[norm]
                val = row[col]
                if std in ("weight", "price_per_ct", "amount"):
                    rec[std] = to_float(val)
                else:
                    rec[std] = clean_value(val)

        # Special: Certificate column often holds "GIA"
        if rec.get("lab") and not rec.get("certificate_no"):
            # Look for a separate Certificate NO column
            for c in df.columns:
                if "certificate no" in normalize_col(c) or "certno" in normalize_col(c):
                    rec["certificate_no"] = clean_value(row[c])
                    break

        # If "Certificate" was mapped to lab but value is numeric, it's actually cert no
        if rec.get("lab") and rec["lab"].replace(".", "").isdigit():
            rec["certificate_no"] = rec["lab"]
            rec["lab"] = "GIA"

        rec["shape"] = normalize_shape(rec.get("shape"))
        rec["fluorescence"] = normalize_fluorescence(rec.get("fluorescence"))

        # Clean video/certificate links that are placeholders
        PLACEHOLDERS = {
            "DNA", "HD", "PDF", "GIA", "CERTIFICATE", "VIDEO", "VIDEO LINK",
            "IMAGE", "DOWNLOAD", "NON", "NONE", "N/A", "NA", "LINK"
        }
        for link_col in ("video_link", "image_link", "certificate_link"):
            v = rec.get(link_col)
            if v and not is_url(v):
                if v.upper().strip() in PLACEHOLDERS or "link" in v.lower():
                    rec[link_col] = None  # placeholder, not a real link
                    note = f"{link_col}: {v}"
                    rec["notes"] = (rec["notes"] + "; " + note) if rec.get("notes") else note

        if not any([rec.get("weight"), rec.get("certificate_no"), rec.get("color")]):
            continue
        records.append(rec)
    return records


def parse_data_file(path: Path) -> List[Dict]:
    """Parser for the complex 'Data_YYYY-M-D_...xlsx' style with summary rows."""
    df_raw = pd.read_excel(path, header=None)
    header_row = detect_header_row(df_raw)
    # Re-read with correct header
    df = pd.read_excel(path, header=header_row)
    df.columns = [str(c).strip() for c in df.columns]

    records = []
    for _, row in df.iterrows():
        rec = {c: None for c in STANDARD_COLUMNS}
        rec["source_file"] = path.name

        for col in df.columns:
            norm = normalize_col(col)
            if norm in COLUMN_ALIASES:
                std = COLUMN_ALIASES[norm]
                val = row[col]
                if std in ("weight", "price_per_ct", "amount"):
                    rec[std] = to_float(val)
                else:
                    rec[std] = clean_value(val)

        # Lab is usually in "Lab" column
        if "Lab" in df.columns:
            rec["lab"] = clean_value(row["Lab"])
        if "Report No" in df.columns:
            rec["certificate_no"] = clean_value(row["Report No"])

        rec["shape"] = normalize_shape(rec.get("shape"))
        rec["fluorescence"] = normalize_fluorescence(rec.get("fluorescence"))

        # Image/Video are often just the words "Image"/"Video" (placeholders)
        for link_col in ("video_link", "image_link"):
            v = rec.get(link_col)
            if v and not is_url(v):
                rec[link_col] = None

        if not any([rec.get("weight"), rec.get("certificate_no"), rec.get("stock_no")]):
            continue
        records.append(rec)
    return records


# ---------------------------------------------------------------------------
# File type detection & dispatch
# ---------------------------------------------------------------------------

def choose_parser(path: Path) -> List[Dict]:
    """Decide which parser to use based on filename / content."""
    name = path.name.lower()

    if "ovals g vs2" in name or "ovalsg" in name.replace(" ", ""):
        log.info(f"  Using Ovals-G-VS2 parser for {path.name}")
        return parse_ovals_g_vs2(path)

    if name.startswith("data_") or "data_" in name:
        log.info(f"  Using Data-file parser for {path.name}")
        return parse_data_file(path)

    # Default: try the simple customer-list layout first, fall back to generic
    try:
        df = pd.read_excel(path, header=0)
        cols = [normalize_col(c) for c in df.columns]
        if "weight" in cols or "cts." in cols or "carats" in cols:
            log.info(f"  Using simple customer-list parser for {path.name}")
            return parse_simple_customer_list(path)
    except Exception as e:
        log.warning(f"  Simple parser failed for {path.name}: {e}")

    log.info(f"  Using generic parser for {path.name}")
    df_raw = pd.read_excel(path, header=None)
    return parse_generic(df_raw, path.name)


# ---------------------------------------------------------------------------
# Persistence
# ---------------------------------------------------------------------------

def save_to_excel(records: List[Dict], out_path: Path) -> None:
    df = pd.DataFrame(records, columns=STANDARD_COLUMNS)
    # Keep certificate / stock numbers as pure text so Excel doesn't turn them into scientific notation
    for col in ("certificate_no", "stock_no"):
        if col in df.columns:
            df[col] = df[col].astype("string")
    # Order nicely
    df = df.sort_values(by=["shape", "weight", "color", "clarity"], na_position="last")
    df.to_excel(out_path, index=False)
    log.info(f"Saved {len(df)} records to {out_path}")


def save_to_csv(records: List[Dict], out_path: Path) -> None:
    df = pd.DataFrame(records, columns=STANDARD_COLUMNS)
    for col in ("certificate_no", "stock_no"):
        if col in df.columns:
            df[col] = df[col].astype("string")
    df.to_csv(out_path, index=False)
    log.info(f"Saved {len(df)} records to {out_path}")


def save_to_sqlite(records: List[Dict], db_path: Path) -> None:
    try:
        # Ensure parent exists and remove stale lock files if any
        db_path.parent.mkdir(parents=True, exist_ok=True)
        for suffix in ("-journal", "-wal", "-shm"):
            lock = Path(str(db_path) + suffix)
            if lock.exists():
                lock.unlink(missing_ok=True)

        conn = sqlite3.connect(str(db_path), timeout=30)
        df = pd.DataFrame(records, columns=STANDARD_COLUMNS)
        for col in ("certificate_no", "stock_no"):
            if col in df.columns:
                df[col] = df[col].astype("string")
        df.to_sql("diamonds", conn, if_exists="replace", index=False)
        conn.execute("CREATE INDEX IF NOT EXISTS idx_cert ON diamonds(certificate_no)")
        conn.execute("CREATE INDEX IF NOT EXISTS idx_stock ON diamonds(stock_no)")
        conn.commit()
        conn.close()
        log.info(f"Saved {len(df)} records to SQLite database {db_path}")
    except Exception as e:
        log.warning(f"Could not write SQLite DB ({e}). Excel/CSV outputs are still available.")


# ---------------------------------------------------------------------------
# Video download (named by certificate number)
# ---------------------------------------------------------------------------

def _safe_filename(
    cert: Optional[str],
    stock: Optional[str],
    url: str,
    shape: Optional[str] = None,
    weight: Optional[float] = None,
    color: Optional[str] = None,
) -> str:
    """
    Build a filesystem-safe filename.
    Priority: certificate_no → stock_no → ID extracted from URL → shape_weight_color_hash
    """
    base = None
    if cert and str(cert).strip() not in ("", "nan", "None"):
        base = str(cert).strip()
    elif stock and str(stock).strip() not in ("", "nan", "None"):
        base = str(stock).strip()
    else:
        # Try to pull a meaningful ID out of the URL itself
        # e.g. .../VIDEOS/1951695.html  or  ?r=ZSMKN169  or  /u/813286ce-...
        m = re.search(r"/VIDEOS/(\w+)", url, re.I)
        if not m:
            m = re.search(r"[?&](?:r|d|id|cid)=([A-Za-z0-9_-]+)", url)
        if not m:
            m = re.search(r"/u/([a-f0-9-]{8,})", url, re.I)
        if m:
            base = m.group(1)
        else:
            # Last resort: descriptive name + short hash of URL
            parts = []
            if shape:
                parts.append(str(shape))
            if weight is not None:
                parts.append(f"{float(weight):.2f}ct")
            if color:
                parts.append(str(color))
            short = hashlib.md5(url.encode()).hexdigest()[:8]
            parts.append(short)
            base = "_".join(parts) if parts else short

    base = re.sub(r'[<>:"/\\|?*\s]+', "_", base)
    base = base[:80] or "unknown"

    path_ext = Path(urlparse(url).path).suffix.lower()
    if path_ext in (".mp4", ".webm", ".mov", ".avi", ".mkv", ".m4v"):
        ext = path_ext
    elif any(x in url.lower() for x in (".html", "html?", "aspx", "v360")):
        ext = ".html"
    else:
        ext = ".mp4"
    return f"{base}{ext}"


def download_videos(
    records: List[Dict],
    video_dir: Path,
    timeout: int = 45,
    progress_callback=None,
) -> Dict[str, int]:
    """
    Download video (or HTML viewer page) for every stone that has a video_link.

    Files are saved as:  <certificate_no>.mp4   or  <certificate_no>.html
    inside  video_dir.

    Returns a small stats dict: {"downloaded": N, "skipped": N, "failed": N}
    """
    video_dir.mkdir(parents=True, exist_ok=True)
    session = requests.Session()
    session.headers.update({
        "User-Agent": (
            "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
            "AppleWebKit/537.36 (KHTML, like Gecko) "
            "Chrome/120.0.0.0 Safari/537.36"
        )
    })

    stats = {"downloaded": 0, "skipped": 0, "failed": 0, "already": 0}
    candidates = [r for r in records if r.get("video_link") and is_url(r["video_link"])]
    total = len(candidates)

    for idx, rec in enumerate(candidates):
        url = rec["video_link"]
        fname = _safe_filename(
            rec.get("certificate_no"),
            rec.get("stock_no"),
            url,
            shape=rec.get("shape"),
            weight=rec.get("weight"),
            color=rec.get("color"),
        )
        dest = video_dir / fname

        if progress_callback:
            progress_callback(idx + 1, total, fname)

        if dest.exists() and dest.stat().st_size > 500:
            log.info(f"Already have {dest.name}")
            stats["already"] += 1
            continue

        try:
            log.info(f"Downloading [{idx+1}/{total}] {url} -> {dest.name}")
            r = session.get(url, timeout=timeout, stream=True, allow_redirects=True)
            r.raise_for_status()

            content_type = (r.headers.get("Content-Type") or "").lower()
            # If the server returned HTML but we expected media, still save it
            # so the user has a local copy of the viewer page.
            if "text/html" in content_type and not dest.suffix == ".html":
                dest = dest.with_suffix(".html")

            with open(dest, "wb") as f:
                for chunk in r.iter_content(chunk_size=64 * 1024):
                    if chunk:
                        f.write(chunk)

            size = dest.stat().st_size
            if size < 200:
                # Almost certainly an error page
                dest.unlink(missing_ok=True)
                log.warning(f"  File too small ({size} B) – discarded")
                stats["failed"] += 1
            else:
                log.info(f"  Saved {dest.name} ({size:,} bytes)")
                stats["downloaded"] += 1
                # Record the local path back onto the record for convenience
                rec["local_video"] = str(dest)

        except Exception as e:
            log.warning(f"  Failed: {e}")
            stats["failed"] += 1

    log.info(
        f"Video download finished: {stats['downloaded']} downloaded, "
        f"{stats['already']} already present, {stats['failed']} failed, "
        f"{stats['skipped']} skipped"
    )
    return stats


# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------

def collect_files(input_dir: Path) -> List[Path]:
    files = sorted(input_dir.glob("*.xlsx")) + sorted(input_dir.glob("*.xls"))
    # Ignore temporary Excel files
    files = [f for f in files if not f.name.startswith("~$")]
    return files


def main():
    parser = argparse.ArgumentParser(
        description="Compile miscellaneous diamond Excel lists into one standardized inventory."
    )
    parser.add_argument(
        "--input-dir",
        type=Path,
        default=Path("/home/workdir/attachments"),
        help="Folder containing the source Excel files",
    )
    parser.add_argument(
        "--output-dir",
        type=Path,
        default=Path("/home/workdir/artifacts/diamond_compiler/output"),
        help="Where to write compiled results",
    )
    parser.add_argument(
        "--db",
        action="store_true",
        help="Also write results to a SQLite database",
    )
    parser.add_argument(
        "--download-videos",
        action="store_true",
        help="Attempt to download direct video files (many links are interactive viewers)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Verbose logging",
    )
    args = parser.parse_args()

    if args.verbose:
        logging.getLogger().setLevel(logging.DEBUG)

    input_dir = args.input_dir.resolve()
    output_dir = args.output_dir.resolve()
    output_dir.mkdir(parents=True, exist_ok=True)

    files = collect_files(input_dir)
    if not files:
        log.error(f"No Excel files found in {input_dir}")
        sys.exit(1)

    log.info(f"Found {len(files)} Excel file(s) in {input_dir}")

    all_records: List[Dict] = []
    for f in files:
        log.info(f"Processing: {f.name}")
        try:
            recs = choose_parser(f)
            log.info(f"  -> extracted {len(recs)} stones")
            all_records.extend(recs)
        except Exception as e:
            log.exception(f"  Failed to parse {f.name}: {e}")

    if not all_records:
        log.error("No records extracted from any file.")
        sys.exit(1)

    # Deduplicate by certificate_no (prefer first occurrence)
    seen = set()
    unique = []
    for r in all_records:
        key = r.get("certificate_no") or r.get("stock_no") or id(r)
        if key not in seen:
            seen.add(key)
            unique.append(r)
    log.info(f"Total unique stones after dedup: {len(unique)} (from {len(all_records)} raw)")

    # Save outputs
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_path = output_dir / f"compiled_diamonds_{ts}.xlsx"
    csv_path = output_dir / f"compiled_diamonds_{ts}.csv"
    save_to_excel(unique, excel_path)
    save_to_csv(unique, csv_path)

    if args.db:
        db_path = output_dir / "diamonds.db"
        save_to_sqlite(unique, db_path)

    if args.download_videos:
        video_dir = output_dir / "videos"
        download_videos(unique, video_dir)

    # Summary
    print("\n" + "=" * 60)
    print(f"Compiled {len(unique)} unique diamonds")
    print(f"  Excel : {excel_path}")
    print(f"  CSV   : {csv_path}")
    if args.db:
        print(f"  SQLite: {output_dir / 'diamonds.db'}")
    print("=" * 60)

    # Quick stats
    df = pd.DataFrame(unique)
    print("\nBreakdown by shape:")
    print(df["shape"].value_counts(dropna=False).to_string())
    print("\nBreakdown by color:")
    print(df["color"].value_counts(dropna=False).to_string())
    print("\nBreakdown by clarity:")
    print(df["clarity"].value_counts(dropna=False).to_string())


if __name__ == "__main__":
    main()
