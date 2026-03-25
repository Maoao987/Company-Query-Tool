#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
web_snapshot.py
用 Edge headless 產生較接近原網站畫面的「網頁列印版」PDF。
"""

import asyncio
import base64
import json
import shutil
import socket
import subprocess
import tempfile
import time
import urllib.request
from functools import lru_cache
from pathlib import Path


_EDGE_CANDIDATES = [
    Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"),
    Path(r"C:\Program Files\Google\Chrome\Application\chrome.exe"),
    Path(r"C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"),
]


def _find_edge() -> Path:
    for path in _EDGE_CANDIDATES:
        if path.exists():
            return path
    raise RuntimeError("找不到 Microsoft Edge，無法產生網頁列印版 PDF。")


def _free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind(("127.0.0.1", 0))
        return sock.getsockname()[1]


def _wait_debug_target(port: int, timeout: float = 10.0) -> str:
    deadline = time.time() + timeout
    last_error = None
    while time.time() < deadline:
        try:
            with urllib.request.urlopen(f"http://127.0.0.1:{port}/json/list", timeout=1) as resp:
                targets = json.load(resp)
            for target in targets:
                if target.get("type") == "page" and target.get("webSocketDebuggerUrl"):
                    return target["webSocketDebuggerUrl"]
        except Exception as exc:  # pragma: no cover - best effort on user machine
            last_error = exc
        time.sleep(0.2)
    raise RuntimeError(f"無法連線到 Edge DevTools。{last_error or ''}".strip())


def _run_edge_print(target: str, virtual_time_budget_ms: int = 8000) -> bytes:
    edge = _find_edge()
    with tempfile.TemporaryDirectory(prefix="cqt-edge-print-") as tmpdir:
        out_pdf = Path(tmpdir) / "snapshot.pdf"
        proc = subprocess.run(
            [
                str(edge),
                "--headless",
                "--disable-gpu",
                f"--virtual-time-budget={virtual_time_budget_ms}",
                f"--print-to-pdf={out_pdf}",
                target,
            ],
            check=False,
            capture_output=True,
            text=True,
        )
        if proc.returncode != 0 or not out_pdf.exists():
            stderr = (proc.stderr or proc.stdout or "").strip()
            raise RuntimeError(f"Edge 列印失敗：{stderr or f'code={proc.returncode}'}")
        return out_pdf.read_bytes()


def _ensure_findbiz_snapshot_html(detail_html: str) -> str:
    html = detail_html or ""
    if not html.strip():
        raise RuntimeError("目前沒有可列印的 findbiz 詳細頁內容。")

    head_extra = (
        '<meta charset="utf-8">'
        '<base href="https://findbiz.nat.gov.tw/">'
        "<style>@page{size:A4;margin:10mm;} body{background:#fff !important;}</style>"
    )
    if "<head>" in html:
        html = html.replace("<head>", f"<head>{head_extra}", 1)
    else:
        html = f"<html><head>{head_extra}</head><body>{html}</body></html>"
    return html


@lru_cache(maxsize=16)
def findbiz_web_print_pdf(detail_html: str) -> bytes:
    html = _ensure_findbiz_snapshot_html(detail_html)
    with tempfile.TemporaryDirectory(prefix="cqt-findbiz-") as tmpdir:
        html_path = Path(tmpdir) / "findbiz_snapshot.html"
        html_path.write_text(html, encoding="utf-8")
        return _run_edge_print(html_path.resolve().as_uri(), virtual_time_budget_ms=5000)


def _yahoo_suffix(market: str) -> str:
    return ".TW" if "TWSE" in (market or "") else ".TWO"


@lru_cache(maxsize=32)
def yahoo_quote_web_print_pdf(stock_no: str, market: str) -> bytes:
    if not stock_no:
        raise RuntimeError("缺少股票代號，無法產生 Yahoo Finance 網頁列印版。")
    url = f"https://tw.stock.yahoo.com/quote/{stock_no}{_yahoo_suffix(market)}"
    return _run_edge_print(url, virtual_time_budget_ms=10000)


@lru_cache(maxsize=32)
def yahoo_dividend_web_print_pdf(stock_no: str, market: str) -> bytes:
    if not stock_no:
        raise RuntimeError("缺少股票代號，無法產生 Yahoo 歷史股利網頁列印版。")
    url = f"https://tw.stock.yahoo.com/quote/{stock_no}{_yahoo_suffix(market)}/dividend"
    return _run_edge_print(url, virtual_time_budget_ms=10000)


async def _cdp_send(ws, msg_id: int, method: str, params: dict | None = None) -> tuple[int, dict]:
    current = msg_id + 1
    await ws.send(json.dumps({"id": current, "method": method, "params": params or {}}))
    while True:
        payload = json.loads(await ws.recv())
        if payload.get("id") == current:
            return current, payload


async def _cdp_wait_for_expr(ws, msg_id: int, expr: str, timeout: float) -> int:
    deadline = time.time() + timeout
    while time.time() < deadline:
        msg_id, res = await _cdp_send(
            ws,
            msg_id,
            "Runtime.evaluate",
            {"expression": expr, "returnByValue": True},
        )
        value = res.get("result", {}).get("result", {}).get("value")
        if value:
            return msg_id
        await asyncio.sleep(0.5)
    raise RuntimeError("等待 MOPS 查詢頁面逾時。")


@lru_cache(maxsize=32)
def mops_dividend_web_print_pdf(stock_no: str) -> bytes:
    if not stock_no:
        raise RuntimeError("缺少股票代號，無法產生 MOPS 網頁列印版。")

    edge = _find_edge()
    port = _free_port()
    user_data_dir = tempfile.mkdtemp(prefix="cqt-edge-cdp-")
    proc = subprocess.Popen(
        [
            str(edge),
            f"--remote-debugging-port={port}",
            "--headless=new",
            "--disable-gpu",
            f"--user-data-dir={user_data_dir}",
            "about:blank",
        ],
        stdout=subprocess.DEVNULL,
        stderr=subprocess.DEVNULL,
    )

    try:
        ws_url = _wait_debug_target(port)
        try:
            import websockets
        except ImportError as exc:  # pragma: no cover - depends on runtime env
            raise RuntimeError("缺少 websockets 套件，無法產生 MOPS 網頁列印版。") from exc

        async def _render() -> bytes:
            async with websockets.connect(ws_url, max_size=30_000_000) as ws:
                msg_id = 0
                msg_id, _ = await _cdp_send(ws, msg_id, "Page.enable")
                msg_id, _ = await _cdp_send(ws, msg_id, "Runtime.enable")
                msg_id, _ = await _cdp_send(
                    ws,
                    msg_id,
                    "Page.navigate",
                    {"url": "https://mops.twse.com.tw/mops/#/web/t05st09_2"},
                )
                msg_id = await _cdp_wait_for_expr(
                    ws,
                    msg_id,
                    "(() => !!document.querySelector('#companyId') && !!document.querySelector('#searchBtn'))()",
                    timeout=20,
                )

                fill_expr = f"""
(() => {{
  const input = document.querySelector('#companyId');
  const button = document.querySelector('#searchBtn');
  if (!input || !button) return false;
  input.value = {json.dumps(stock_no)};
  input.dispatchEvent(new Event('input', {{ bubbles: true }}));
  input.dispatchEvent(new Event('change', {{ bubbles: true }}));
  button.click();
  return true;
}})()
"""
                msg_id, res = await _cdp_send(
                    ws,
                    msg_id,
                    "Runtime.evaluate",
                    {"expression": fill_expr, "returnByValue": True},
                )
                if not res.get("result", {}).get("result", {}).get("value"):
                    raise RuntimeError("找不到 MOPS 查詢欄位或查詢按鈕。")

                msg_id = await _cdp_wait_for_expr(
                    ws,
                    msg_id,
                    "(() => document.body.innerText.includes('本資料由') || document.body.innerText.includes('查詢無資料'))()",
                    timeout=20,
                )
                _, pdf = await _cdp_send(
                    ws,
                    msg_id,
                    "Page.printToPDF",
                    {"printBackground": True, "landscape": False},
                )
                return base64.b64decode(pdf["result"]["data"])

        return asyncio.run(_render())
    finally:
        proc.terminate()
        try:
            proc.wait(timeout=5)
        except Exception:
            proc.kill()
        shutil.rmtree(user_data_dir, ignore_errors=True)
