# Diamond Inventory Compiler

Compile miscellaneous customer/supplier diamond Excel lists into one clean inventory.

Works **locally** and on **Streamlit Community Cloud**, with optional **GitHub persistence** so you do not need to keep files on your computer.

## Features

- Upload sheets → merge & de-duplicate
- **Full Inventory** tab with filters
- **Summary** dashboard + multi-sheet Excel
- **Save master inventory on GitHub** (recommended for Cloud)

## Deploy to Streamlit Cloud

1. Push this repo to GitHub
2. [share.streamlit.io](https://share.streamlit.io) → New app → main file `gui_app.py`
3. Add **Secrets** (App → ⚙️ Settings → Secrets):

```toml
GITHUB_TOKEN = "ghp_xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx"
GITHUB_REPO  = "yourusername/your-repo-name"
GITHUB_BRANCH = "main"
GITHUB_PATH  = "data/master_diamonds.xlsx"
```

### Create the GitHub token

1. GitHub → Settings → Developer settings → Personal access tokens → **Tokens (classic)**
2. Generate new token with the **`repo`** scope
3. Paste it into Streamlit Secrets as `GITHUB_TOKEN`

After that:

- App **loads** `data/master_diamonds.xlsx` from the repo on startup
- After you process new sheets, it **auto-saves** the merged inventory back to GitHub
- You can also click **Save inventory to GitHub** / **Reload from GitHub** in the sidebar

No need to download or re-upload the master file yourself.

## Run locally

```bash
pip install -r requirements.txt
streamlit run gui_app.py
```

(Optional) set the same secrets in `.streamlit/secrets.toml` for local GitHub save.

## Video files

Video **links** are stored in the inventory (and on GitHub).  
Actual video files cannot stay on Streamlit Cloud (temporary disk). Download them only when running locally.

## Project layout

```
├── gui_app.py              # Streamlit UI
├── diamond_compiler.py     # Excel parsers
├── github_storage.py       # Load/save master Excel via GitHub API
├── data/
│   └── master_diamonds.xlsx   # created automatically after first save
├── requirements.txt
└── README.md
```
