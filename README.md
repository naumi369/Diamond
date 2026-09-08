# Diamond Inventory Compiler

A tool that takes miscellaneous Excel diamond lists from different customers / suppliers and merges them into **one clean, standardized inventory**.

## Two ways to use it

### 1. GUI (recommended)

```bash
cd diamond_compiler
pip install -r requirements.txt
streamlit run gui_app.py
```

A browser window opens. Upload your Excel files, click **Compile Inventory**, and optionally download videos. Results appear on screen and are also written to the output folder.

### 2. Command line

```bash
python diamond_compiler.py \
    --input-dir /path/to/excel/folder \
    --output-dir /path/to/save/results \
    --db \
    --download-videos
```

| Flag | Meaning |
|------|---------|
| `--input-dir` | Folder with source `.xlsx` files |
| `--output-dir` | Where compiled files are written |
| `--db` | Also create SQLite database `diamonds.db` |
| `--download-videos` | Download videos named by certificate number |
| `-v` | Verbose logging |

## Standard columns

| Column | Description |
|--------|-------------|
| stock_no | Supplier / internal reference |
| certificate_no | GIA (or other) report number |
| lab | Laboratory (GIA, IGI, …) |
| shape | Oval, Round, Pear, … |
| weight | Carat weight |
| mm_size | Measurements (L × W × D) |
| color / clarity | Grades |
| polish / symmetry / fluorescence / cut | Finish grades |
| price_per_ct / amount | Pricing |
| video_link / image_link / certificate_link | Links |
| source_file | Original Excel the row came from |
| notes | Placeholder text that was cleaned |

## Video downloads

When enabled, every stone that has a `video_link` is downloaded into:

```
output/videos/<certificate_no>.mp4
# or
output/videos/<certificate_no>.html
```

- **Direct media** (`.mp4`, `.webm`, …) → saved as media.
- **Interactive 360° viewers** (v360, diamondview, etc.) → the HTML page is saved so you still have a local copy (you can open it in a browser).
- Filename is always the **certificate number** (falls back to stock number).

## Supported Excel layouts

1. Simple customer lists (`Carat Size Range | Certificate | Shape | WEIGHT | …`)
2. Ovals G VS2-SI1 style (`Ref.No | Cts. | Measurement | CertNo | …`)
3. Complex market sheets (`Data_YYYY-…`) – finds the real header under the summary
4. Generic auto-detect

Just keep dropping new files into the input folder / uploader; the tool adapts.

## Output location

```
diamond_compiler/output/
├── compiled_diamonds_YYYYMMDD_HHMMSS.xlsx
├── compiled_diamonds_YYYYMMDD_HHMMSS.csv
├── diamonds.db
└── videos/
    ├── 1535750439.mp4
    ├── 2547813803.html
    └── …
```
