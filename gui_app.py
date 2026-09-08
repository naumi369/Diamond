#!/usr/bin/env python3
"""
Diamond Inventory Compiler – GUI (Streamlit)
--------------------------------------------
Run with:
    streamlit run gui_app.py
"""

import io
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import pandas as pd
import streamlit as st

# Make sure we can import the compiler sitting next to this file
sys.path.insert(0, str(Path(__file__).resolve().parent))
from diamond_compiler import (  # noqa: E402
    STANDARD_COLUMNS,
    choose_parser,
    download_videos,
    save_to_csv,
    save_to_excel,
    save_to_sqlite,
)

# ---------------------------------------------------------------------------
# Page config
# ---------------------------------------------------------------------------
st.set_page_config(
    page_title="Diamond Inventory Compiler",
    page_icon="💎",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.title("💎 Diamond Inventory Compiler")
st.caption(
    "Drop miscellaneous customer Excel lists → get one clean, standardized inventory. "
    "Videos are saved as **&lt;certificate_no&gt;.mp4 / .html** in the output folder."
)

# ---------------------------------------------------------------------------
# Sidebar – settings
# ---------------------------------------------------------------------------
with st.sidebar:
    st.header("Settings")
    default_out = str(Path(__file__).resolve().parent / "output")
    output_dir = st.text_input("Output folder", value=default_out)
    save_db = st.checkbox("Also save SQLite database", value=True)
    do_videos = st.checkbox(
        "Download video links",
        value=True,
        help="Saves each video (or HTML viewer page) as <certificate_no>.ext inside a 'videos' sub-folder.",
    )
    st.markdown("---")
    st.markdown(
        "**Supported layouts**\n"
        "- Simple customer lists (WEIGHT / Color / Clarity …)\n"
        "- Ovals G VS2-SI1 style (Ref.No / CertNo …)\n"
        "- Complex market sheets (Data_YYYY-…)\n"
        "- Auto-detect generic Excel"
    )

# ---------------------------------------------------------------------------
# File uploader
# ---------------------------------------------------------------------------
uploaded = st.file_uploader(
    "Upload one or more diamond Excel files (.xlsx)",
    type=["xlsx", "xls"],
    accept_multiple_files=True,
)

if not uploaded:
    st.info("👆 Upload the Excel files you received from customers / suppliers to begin.")
    st.stop()

st.success(f"{len(uploaded)} file(s) ready: " + ", ".join(f.name for f in uploaded))

# ---------------------------------------------------------------------------
# Compile button
# ---------------------------------------------------------------------------
if st.button("🚀 Compile Inventory", type="primary", use_container_width=True):
    out_path = Path(output_dir)
    out_path.mkdir(parents=True, exist_ok=True)
    videos_path = out_path / "videos"

    progress = st.progress(0, text="Starting…")
    log_box = st.empty()
    logs: list[str] = []

    def log(msg: str):
        logs.append(msg)
        log_box.code("\n".join(logs[-30:]), language=None)

    # ---- 1. Parse every uploaded file ----
    all_records = []
    with tempfile.TemporaryDirectory() as tmp:
        tmp_path = Path(tmp)
        for i, uf in enumerate(uploaded):
            progress.progress((i) / max(len(uploaded) + 2, 1), text=f"Parsing {uf.name}…")
            local = tmp_path / uf.name
            local.write_bytes(uf.getvalue())
            try:
                recs = choose_parser(local)
                log(f"✓ {uf.name} → {len(recs)} stones")
                all_records.extend(recs)
            except Exception as e:
                log(f"✗ {uf.name} failed: {e}")

    if not all_records:
        st.error("No stones could be extracted from the uploaded files.")
        st.stop()

    # ---- 2. Deduplicate ----
    progress.progress(0.6, text="Deduplicating…")
    seen = set()
    unique = []
    for r in all_records:
        key = r.get("certificate_no") or r.get("stock_no") or id(r)
        if key not in seen:
            seen.add(key)
            unique.append(r)
    log(f"Total unique stones: {len(unique)} (from {len(all_records)} raw rows)")

    # ---- 3. Save Excel / CSV / DB ----
    progress.progress(0.75, text="Writing Excel & CSV…")
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    excel_file = out_path / f"compiled_diamonds_{ts}.xlsx"
    csv_file = out_path / f"compiled_diamonds_{ts}.csv"
    save_to_excel(unique, excel_file)
    save_to_csv(unique, csv_file)
    log(f"Saved Excel → {excel_file.name}")
    log(f"Saved CSV   → {csv_file.name}")

    if save_db:
        db_file = out_path / "diamonds.db"
        save_to_sqlite(unique, db_file)
        log(f"Saved SQLite → {db_file.name}")

    # ---- 4. Download videos ----
    video_stats = {"downloaded": 0, "already": 0, "failed": 0}
    if do_videos:
        progress.progress(0.85, text="Downloading videos…")
        video_bar = st.progress(0, text="Videos…")

        def video_progress(current, total, fname):
            video_bar.progress(current / max(total, 1), text=f"Video {current}/{total}: {fname}")

        video_stats = download_videos(unique, videos_path, progress_callback=video_progress)
        log(
            f"Videos: {video_stats.get('downloaded', 0)} new, "
            f"{video_stats.get('already', 0)} already present, "
            f"{video_stats.get('failed', 0)} failed"
        )
        video_bar.empty()

    progress.progress(1.0, text="Done!")
    st.balloons()

    # ---- Results summary ----
    df = pd.DataFrame(unique, columns=STANDARD_COLUMNS)

    st.subheader("Summary")
    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Unique stones", len(unique))
    c2.metric("With video link", int(df["video_link"].notna().sum()))
    c3.metric("Videos downloaded", video_stats.get("downloaded", 0) + video_stats.get("already", 0))
    c4.metric("Source files", len(uploaded))

    st.markdown("#### Breakdown")
    col_a, col_b, col_c = st.columns(3)
    with col_a:
        st.write("**Shape**")
        st.dataframe(df["shape"].value_counts(dropna=False).rename("count"), use_container_width=True)
    with col_b:
        st.write("**Color**")
        st.dataframe(df["color"].value_counts(dropna=False).rename("count"), use_container_width=True)
    with col_c:
        st.write("**Clarity**")
        st.dataframe(df["clarity"].value_counts(dropna=False).rename("count"), use_container_width=True)

    st.markdown("#### Compiled data (preview)")
    st.dataframe(df, use_container_width=True, height=400)

    # Download buttons for the generated files
    st.markdown("#### Download results")
    d1, d2 = st.columns(2)
    with d1:
        with open(excel_file, "rb") as f:
            st.download_button(
                "📥 Download Excel",
                data=f,
                file_name=excel_file.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
    with d2:
        with open(csv_file, "rb") as f:
            st.download_button(
                "📥 Download CSV",
                data=f,
                file_name=csv_file.name,
                mime="text/csv",
                use_container_width=True,
            )

    st.info(f"All files also saved on disk under:\n`{out_path}`")
    if do_videos:
        st.info(f"Videos saved under:\n`{videos_path}`  (named by certificate number)")
