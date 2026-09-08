"""
Persist the master diamond inventory on GitHub.

Required Streamlit secrets (Settings → Secrets on share.streamlit.io):

    GITHUB_TOKEN = "ghp_...."          # classic PAT with `repo` scope
    GITHUB_REPO  = "username/diamond"  # owner/repo
    GITHUB_BRANCH = "main"             # optional, default main
    GITHUB_PATH  = "data/master_diamonds.xlsx"  # optional

Or a single TOML block:

    [github]
    token = "ghp_...."
    repo = "username/diamond"
    branch = "main"
    path = "data/master_diamonds.xlsx"
"""

from __future__ import annotations

import base64
import io
from datetime import datetime, timezone
from typing import Any, Dict, Optional, Tuple

import pandas as pd
import requests

API = "https://api.github.com"


def _secrets() -> Dict[str, str]:
    """Read GitHub settings from st.secrets (works only inside Streamlit)."""
    try:
        import streamlit as st
        sec = st.secrets
    except Exception:
        return {}

    # Flat keys
    token = sec.get("GITHUB_TOKEN") or sec.get("github_token")
    repo = sec.get("GITHUB_REPO") or sec.get("github_repo")
    branch = sec.get("GITHUB_BRANCH") or sec.get("github_branch") or "main"
    path = sec.get("GITHUB_PATH") or sec.get("github_path") or "data/master_diamonds.xlsx"

    # Nested [github] table
    if not token and "github" in sec:
        g = sec["github"]
        token = g.get("token") or token
        repo = g.get("repo") or repo
        branch = g.get("branch") or branch
        path = g.get("path") or path

    out = {}
    if token:
        out["token"] = str(token).strip()
    if repo:
        out["repo"] = str(repo).strip()
    if branch:
        out["branch"] = str(branch).strip()
    if path:
        out["path"] = str(path).strip()
    return out


def is_configured() -> bool:
    s = _secrets()
    return bool(s.get("token") and s.get("repo"))


def _headers(token: str) -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
        "User-Agent": "DiamondInventoryCompiler/1.0",
    }


def _get_file(repo: str, path: str, branch: str, token: str) -> Tuple[Optional[bytes], Optional[str]]:
    """
    Return (content_bytes, sha) for an existing file, or (None, None) if missing.
    """
    url = f"{API}/repos/{repo}/contents/{path}"
    r = requests.get(url, headers=_headers(token), params={"ref": branch}, timeout=60)
    if r.status_code == 404:
        return None, None
    r.raise_for_status()
    data = r.json()
    sha = data.get("sha")
    content_b64 = data.get("content", "")
    # GitHub returns base64 with newlines
    raw = base64.b64decode(content_b64.replace("\n", ""))
    return raw, sha


def load_master_excel() -> Tuple[Optional[pd.DataFrame], str]:
    """
    Load master inventory from GitHub.
    Returns (dataframe_or_None, status_message).
    """
    s = _secrets()
    if not s.get("token") or not s.get("repo"):
        return None, "GitHub secrets not configured"

    try:
        raw, sha = _get_file(s["repo"], s["path"], s["branch"], s["token"])
        if raw is None:
            return None, f"No file yet at {s['path']} on {s['branch']}"
        df = pd.read_excel(io.BytesIO(raw))
        return df, f"Loaded {len(df)} stones from GitHub ({s['path']})"
    except Exception as e:
        return None, f"GitHub load failed: {e}"


def save_master_excel(df: pd.DataFrame) -> str:
    """
    Commit the dataframe as Excel to GitHub (create or update).
    Returns a human-readable status message.
    """
    s = _secrets()
    if not s.get("token") or not s.get("repo"):
        return "GitHub secrets not configured – cannot save to cloud"

    if df is None or df.empty:
        return "Nothing to save (inventory is empty)"

    try:
        buf = io.BytesIO()
        out = df.copy()
        for col in ("certificate_no", "stock_no"):
            if col in out.columns:
                out[col] = out[col].astype("string")
        with pd.ExcelWriter(buf, engine="openpyxl") as writer:
            out.to_excel(writer, index=False, sheet_name="Inventory")
        content_b64 = base64.b64encode(buf.getvalue()).decode("ascii")

        # Need current SHA to update an existing file
        _, sha = _get_file(s["repo"], s["path"], s["branch"], s["token"])

        ts = datetime.now(timezone.utc).strftime("%Y-%m-%d %H:%M UTC")
        message = f"Update master diamond inventory ({len(df)} stones) – {ts}"

        payload: Dict[str, Any] = {
            "message": message,
            "content": content_b64,
            "branch": s["branch"],
        }
        if sha:
            payload["sha"] = sha

        url = f"{API}/repos/{s['repo']}/contents/{s['path']}"
        r = requests.put(url, headers=_headers(s["token"]), json=payload, timeout=90)
        if r.status_code in (200, 201):
            commit = r.json().get("commit", {}).get("html_url", "")
            return f"Saved {len(df)} stones to GitHub → {s['path']}" + (
                f"\n{commit}" if commit else ""
            )
        return f"GitHub save failed ({r.status_code}): {r.text[:500]}"
    except Exception as e:
        return f"GitHub save failed: {e}"
