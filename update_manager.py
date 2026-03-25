#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
update_manager.py
GitHub Releases 更新檢查與下載
"""

from __future__ import annotations

import json
import os
import subprocess
import tempfile
from pathlib import Path

import requests

DEFAULT_ASSET_NAME = "CompanyQueryToolSetup.exe"
UPDATE_CONFIG_NAME = "update_config.json"
VERSION_FILE_NAME = "version.txt"

_REQ = {
    "headers": {
        "User-Agent": "CompanyQueryToolUpdater",
        "Accept": "application/vnd.github+json",
    },
    "timeout": 15,
}


def get_version_file_path() -> Path:
    return Path(__file__).with_name(VERSION_FILE_NAME)


def load_app_version() -> str:
    path = get_version_file_path()
    if not path.exists():
        return "0.0.0"
    try:
        return path.read_text(encoding="utf-8").strip() or "0.0.0"
    except Exception:
        return "0.0.0"


APP_VERSION = load_app_version()
_REQ["headers"]["User-Agent"] = f"CompanyQueryToolUpdater/{APP_VERSION}"


def get_update_config_path() -> Path:
    return Path(__file__).with_name(UPDATE_CONFIG_NAME)


def load_update_config() -> dict:
    path = get_update_config_path()
    if not path.exists():
        return {}
    try:
        return json.loads(path.read_text(encoding="utf-8"))
    except Exception:
        return {}


def update_is_configured(config: dict | None = None) -> bool:
    cfg = config or load_update_config()
    return bool(str(cfg.get("github_owner", "")).strip() and str(cfg.get("github_repo", "")).strip())


def normalize_version(version: str) -> tuple[int, ...]:
    raw = str(version or "").strip().lstrip("vV")
    parts = []
    for part in raw.split("."):
        digits = "".join(ch for ch in part if ch.isdigit())
        parts.append(int(digits) if digits else 0)
    while len(parts) < 3:
        parts.append(0)
    return tuple(parts[:3])


def is_newer_version(latest_version: str, current_version: str = APP_VERSION) -> bool:
    return normalize_version(latest_version) > normalize_version(current_version)


def fetch_latest_release(config: dict | None = None) -> dict:
    cfg = config or load_update_config()
    owner = str(cfg.get("github_owner", "")).strip()
    repo = str(cfg.get("github_repo", "")).strip()
    asset_name = str(cfg.get("asset_name", "")).strip() or DEFAULT_ASSET_NAME

    if not owner or not repo:
        return {"configured": False, "error": "尚未設定 GitHub 更新來源"}

    url = f"https://api.github.com/repos/{owner}/{repo}/releases/latest"
    try:
        response = requests.get(url, **_REQ)
        response.raise_for_status()
        data = response.json()
    except Exception as exc:
        hint = ""
        if "404" in str(exc):
            hint = "；若 repo 是 private，請改成 public，或改用公開的 releases repo"
        return {
            "configured": True,
            "error": f"無法取得 GitHub Releases：{exc}{hint}",
        }

    assets = data.get("assets", []) or []
    matched_asset = next(
        (asset for asset in assets if str(asset.get("name", "")).strip() == asset_name),
        None,
    )
    latest_version = str(data.get("tag_name") or data.get("name") or "").strip()
    return {
        "configured": True,
        "owner": owner,
        "repo": repo,
        "latest_version": latest_version,
        "current_version": APP_VERSION,
        "update_available": bool(latest_version and is_newer_version(latest_version, APP_VERSION)),
        "asset_name": asset_name,
        "download_url": (matched_asset or {}).get("browser_download_url", ""),
        "release_page": str(data.get("html_url", "")).strip(),
        "published_at": str(data.get("published_at", "")).strip(),
        "body": str(data.get("body", "")).strip(),
        "asset_found": matched_asset is not None,
    }


def download_and_launch_update(download_url: str, asset_name: str = DEFAULT_ASSET_NAME) -> str:
    if not download_url:
        raise ValueError("缺少更新下載網址")

    response = requests.get(download_url, stream=True, timeout=60)
    response.raise_for_status()

    temp_dir = Path(tempfile.gettempdir()) / "companyquery_updates"
    temp_dir.mkdir(parents=True, exist_ok=True)
    target = temp_dir / asset_name

    with target.open("wb") as fh:
        for chunk in response.iter_content(chunk_size=1024 * 1024):
            if chunk:
                fh.write(chunk)

    try:
        subprocess.Popen(
            ["cmd", "/c", "start", "", str(target)],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
        )
    except Exception:
        os.startfile(str(target))
    return str(target)
