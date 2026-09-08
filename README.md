# Diamond Inventory Compiler

Compile miscellaneous customer/supplier diamond Excel lists into one clean inventory.

Works **locally** and on **Streamlit Community Cloud**.

## Features

- **Upload & Add** – keep uploading new sheets; stones are merged and de-duplicated
- **Full Inventory** – filterable table + download Master Excel / CSV
- **Summary** – KPIs, charts, shape×color pivot, multi-sheet summary workbook
- Video links stored; actual video files downloadable **locally only**

## Run locally

```bash
cd diamond_compiler
pip install -r requirements.txt
streamlit run gui_app.py
```

## Deploy to Streamlit Cloud

1. Push this folder to a **public GitHub repo** (or private if you have Streamlit team plan)
2. Go to [share.streamlit.io](https://share.streamlit.io) → **New app**
3. Select the repo, branch, and set **Main file path** to `gui_app.py`
4. Deploy

`requirements.txt` is already set for Cloud.

### Important Cloud limitations

| Topic | Reality on Streamlit Cloud |
|-------|----------------------------|
| **Video files** | **Cannot be stored permanently.** Disk is wiped when the app sleeps. Video *links* stay in the table; open them in a browser. For real file downloads, run locally. |
| **Inventory persistence** | Session memory is lost on restart. **Always download the Master Excel** and re-upload it next time together with new sheets. |
| **File size** | Keep individual Excel uploads reasonable (< 200 MB total per session is safer). |

## Recommended workflow (Cloud or local)

1. Upload any previous **Master Excel** (optional first time)
2. Upload new customer sheets
3. Click **Process & Add to Inventory**
4. Review **Full Inventory** and **Summary** tabs
5. Download **Master Excel** and keep it safe
6. Next batch: upload that Master + the new sheets again

Duplicates are removed automatically (by certificate number, then stock number).

## CLI (no GUI)

```bash
python diamond_compiler.py --input-dir ./sheets --output-dir ./output --db --download-videos
```

## Project layout

```
diamond_compiler/
├── gui_app.py              ← Streamlit app (deploy this)
├── diamond_compiler.py     ← parsing / download engine
├── requirements.txt
├── README.md
└── output/                 ← local results & videos/
```
