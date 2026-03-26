#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
公司資料查詢工具 - 網頁介面
執行：streamlit run app.py
"""

import html
import io
import time
from datetime import date as date_cls, datetime

import pandas as pd
import streamlit as st

from company_query import (
    RESULT_COLUMNS,
    _flatten_result,
    extract_batch_requests,
    get_trading_days,
    init_caches,
    query_by_stock_no,
    query_by_uid,
    run_batch_request,
    to_excel_bytes,
    to_csv_bytes,
)
from findbiz_scraper import search_companies_by_name
from pdf_report import (
    generate_batch_report_pdf,
    generate_dividend_snapshot_pdf,
    generate_pdf,
    generate_stock_snapshot_pdf,
)
from web_snapshot import (
    findbiz_web_print_pdf,
)
from update_manager import (
    APP_VERSION,
    download_and_launch_update,
    fetch_latest_release,
    load_update_config,
    update_is_configured,
)

# ── 頁面設定 ──────────────────────────────────────────────────────────────────
st.set_page_config(page_title="公司資料查詢", page_icon="🔍", layout="wide")
st.markdown('<div id="page-top"></div>', unsafe_allow_html=True)


def inject_global_styles():
    st.markdown(
        """
        <style>
        :root {
          --ink: #172033;
          --muted: #5d6b81;
          --line: #d8e1ef;
          --soft-line: #e8eef7;
          --paper: rgba(255, 255, 255, 0.92);
          --panel: rgba(248, 251, 255, 0.88);
          --navy: #1f4b84;
          --navy-deep: #14355d;
          --gold: #b6862c;
          --blue-soft: #edf4ff;
          --mint-soft: #eefaf5;
          --shadow: 0 18px 40px rgba(16, 34, 64, 0.08);
          --radius-lg: 24px;
          --radius-md: 18px;
          --radius-sm: 14px;
        }

        .stApp {
          background:
            radial-gradient(circle at top left, rgba(178, 205, 242, 0.28), transparent 28%),
            radial-gradient(circle at top right, rgba(235, 220, 188, 0.20), transparent 22%),
            linear-gradient(180deg, #f4f8fc 0%, #eef3f8 54%, #f8fafc 100%);
          color: var(--ink);
        }

        html, body, [data-testid="stAppViewContainer"], [data-testid="stMain"], [data-testid="stSidebar"] {
          color-scheme: light !important;
          background-color: #f4f8fc !important;
          color: var(--ink) !important;
        }

        [data-testid="stHeader"] {
          background: rgba(244, 248, 252, 0.85) !important;
        }

        [data-testid="stToolbar"] {
          background: transparent !important;
        }

        .stApp button, .stApp input, .stApp textarea, .stApp select {
          color: var(--ink) !important;
        }

        .stApp [data-baseweb="input"] > div,
        .stApp [data-baseweb="select"] > div,
        .stApp [data-baseweb="base-input"] > div,
        .stApp [data-baseweb="textarea"] > div,
        .stApp [data-baseweb="tag"] {
          background: #ffffff !important;
          color: var(--ink) !important;
          border-color: #d8e1ef !important;
        }

        .stApp .stButton > button,
        .stApp .stDownloadButton > button {
          background: linear-gradient(180deg, #ffffff 0%, #eef4ff 100%) !important;
          color: var(--navy-deep) !important;
          border: 1px solid #c7d6ea !important;
        }

        .stApp .stButton > button[kind="primary"],
        .stApp .stDownloadButton > button[kind="primary"],
        .stApp .stFormSubmitButton > button {
          background: linear-gradient(180deg, #2d67ac 0%, #1f4b84 100%) !important;
          color: #ffffff !important;
          border-color: #1f4b84 !important;
          text-shadow: 0 1px 1px rgba(12, 24, 45, 0.18);
        }

        .stApp .stButton > button[kind="primary"] *,
        .stApp .stDownloadButton > button[kind="primary"] *,
        .stApp .stFormSubmitButton > button,
        .stApp .stFormSubmitButton > button *,
        .stApp .stFormSubmitButton > button p,
        .stApp .stFormSubmitButton > button span,
        .stApp .stFormSubmitButton > button div {
          color: #ffffff !important;
          fill: #ffffff !important;
          opacity: 1 !important;
        }

        .stApp [data-baseweb="select"] * ,
        .stApp [data-baseweb="input"] * ,
        .stApp [data-baseweb="textarea"] * {
          color: var(--ink) !important;
        }

        .stApp .stCaption,
        .stApp [data-testid="stCaptionContainer"],
        .stApp .stCaption p,
        .stApp [data-testid="stCaptionContainer"] p {
          color: #4a5d79 !important;
          font-weight: 600 !important;
        }

        .stApp [data-testid="stAppViewContainer"] > .main .block-container {
          max-width: 1220px;
          padding-top: 1.35rem;
          padding-bottom: 3rem;
        }

        h1, h2, h3, .hero-title, .section-title {
          font-family: "PMingLiU", "Noto Serif TC", "Microsoft JhengHei", serif;
          letter-spacing: 0.01em;
        }

        .hero-shell {
          position: relative;
          overflow: hidden;
          padding: 1.45rem 1.6rem;
          border: 1px solid rgba(190, 204, 224, 0.85);
          border-radius: 30px;
          background:
            linear-gradient(140deg, rgba(255,255,255,0.96) 0%, rgba(243,247,252,0.92) 60%, rgba(235,243,251,0.88) 100%);
          box-shadow: var(--shadow);
          margin-bottom: 1.2rem;
        }

        .hero-shell::after {
          content: "";
          position: absolute;
          inset: auto -4rem -5rem auto;
          width: 14rem;
          height: 14rem;
          border-radius: 999px;
          background: radial-gradient(circle, rgba(31, 75, 132, 0.16) 0%, rgba(31, 75, 132, 0) 70%);
        }

        .hero-kicker {
          display: inline-flex;
          align-items: center;
          gap: 0.45rem;
          padding: 0.34rem 0.7rem;
          border-radius: 999px;
          background: rgba(31, 75, 132, 0.08);
          color: var(--navy);
          font-size: 0.84rem;
          font-weight: 700;
          letter-spacing: 0.08em;
          text-transform: uppercase;
        }

        .hero-title {
          margin: 0.8rem 0 0.45rem;
          font-size: clamp(2rem, 4vw, 3rem);
          color: var(--navy-deep);
          line-height: 1.08;
        }

        .hero-lead {
          max-width: 46rem;
          margin: 0;
          color: var(--muted);
          font-size: 1rem;
          line-height: 1.75;
        }

        .hero-source-row {
          display: flex;
          flex-wrap: wrap;
          gap: 0.55rem;
          margin-top: 1rem;
        }

        .source-chip {
          display: inline-flex;
          align-items: center;
          gap: 0.4rem;
          padding: 0.36rem 0.72rem;
          border-radius: 999px;
          background: rgba(255,255,255,0.78);
          border: 1px solid var(--line);
          color: #3c4c63;
          font-size: 0.88rem;
        }

        .panel-title {
          margin: 0.1rem 0 0.35rem;
          font-size: 1.15rem;
          font-weight: 800;
          color: var(--navy-deep);
        }

        .panel-subtitle {
          margin: 0 0 0.9rem;
          color: var(--muted);
          font-size: 0.95rem;
          line-height: 1.7;
        }

        .summary-strip {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
          gap: 0.8rem;
          margin: 0.75rem 0 1rem;
        }

        .summary-card {
          padding: 0.95rem 1rem;
          border-radius: var(--radius-md);
          background: var(--paper);
          border: 1px solid rgba(204, 216, 233, 0.92);
          box-shadow: 0 10px 26px rgba(17, 35, 67, 0.05);
        }

        .summary-label {
          display: block;
          margin-bottom: 0.3rem;
          color: var(--muted);
          font-size: 0.82rem;
          letter-spacing: 0.04em;
        }

        .summary-value {
          display: block;
          color: var(--ink);
          font-size: 1.2rem;
          font-weight: 800;
          line-height: 1.2;
          word-break: break-word;
        }

        .section-card {
          padding: 1.05rem 1.15rem 1.15rem;
          border-radius: var(--radius-lg);
          background: var(--paper);
          border: 1px solid rgba(204, 216, 233, 0.92);
          box-shadow: var(--shadow);
          margin-bottom: 1rem;
        }

        .section-head {
          display: flex;
          align-items: flex-start;
          justify-content: space-between;
          gap: 0.9rem;
          margin-bottom: 0.85rem;
        }

        .section-head-left {
          display: flex;
          align-items: center;
          gap: 0.75rem;
        }

        .section-icon {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          width: 2.5rem;
          height: 2.5rem;
          border-radius: 18px;
          background: linear-gradient(180deg, #eef5ff 0%, #e1ecfb 100%);
          color: var(--navy);
          font-size: 1.2rem;
        }

        .section-title {
          margin: 0;
          font-size: 1.18rem;
          color: var(--navy-deep);
        }

        .section-note {
          margin: 0.2rem 0 0;
          color: var(--muted);
          font-size: 0.88rem;
          line-height: 1.65;
        }

        .info-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          overflow: hidden;
          border: 1px solid var(--soft-line);
          border-radius: 18px;
        }

        .info-table th, .info-table td {
          padding: 0.78rem 0.9rem;
          text-align: left;
          vertical-align: top;
          border-bottom: 1px solid var(--soft-line);
        }

        .info-table tr:last-child th, .info-table tr:last-child td {
          border-bottom: none;
        }

        .info-table th {
          width: 34%;
          background: linear-gradient(180deg, #f8fbff 0%, #f1f6fd 100%);
          color: #38465f;
          font-weight: 700;
        }

        .info-table td {
          background: rgba(255,255,255,0.88);
          color: var(--ink);
          line-height: 1.72;
        }

        .biz-grid {
          display: grid;
          grid-template-columns: repeat(auto-fit, minmax(240px, 1fr));
          gap: 0.7rem;
        }

        .biz-item {
          padding: 0.75rem 0.85rem;
          border-radius: 16px;
          background: linear-gradient(180deg, #fbfdff 0%, #f4f8fd 100%);
          border: 1px solid var(--soft-line);
          color: #304158;
          line-height: 1.7;
        }

        .link-pill-row {
          display: flex;
          flex-wrap: wrap;
          gap: 0.6rem;
        }

        .link-pill-row a {
          display: inline-flex;
          align-items: center;
          gap: 0.45rem;
          padding: 0.55rem 0.9rem;
          border-radius: 999px;
          background: linear-gradient(180deg, #244f88 0%, #173a67 100%);
          border: 1px solid #14355d;
          color: #f8fbff;
          font-weight: 700;
          text-decoration: none;
          box-shadow: 0 10px 22px rgba(20, 53, 93, 0.18);
        }

        .link-pill-row a:hover {
          background: linear-gradient(180deg, #2b5c9d 0%, #1b4377 100%);
          border-color: #173a67;
          color: #ffffff;
        }

        .hint-box {
          padding: 0.85rem 1rem;
          border-radius: 18px;
          background: linear-gradient(180deg, rgba(242, 247, 255, 0.95) 0%, rgba(252, 254, 255, 0.92) 100%);
          border: 1px solid #d8e3f2;
          color: #43536b;
          line-height: 1.72;
        }

        .dividend-table {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          overflow: hidden;
          border: 1px solid var(--soft-line);
          border-radius: 18px;
        }

        .dividend-table thead th {
          background: linear-gradient(180deg, #f4f8ff 0%, #e9f1fb 100%);
          color: #334155;
          font-weight: 700;
          padding: 0.78rem 0.9rem;
          border-bottom: 1px solid var(--soft-line);
        }

        .dividend-table tbody td {
          padding: 0.76rem 0.9rem;
          border-bottom: 1px solid var(--soft-line);
          background: rgba(255,255,255,0.92);
        }

        .dividend-table tbody tr:nth-child(even) td {
          background: #fbfdff;
        }

        .dividend-table tbody tr:last-child td {
          border-bottom: none;
        }

        .soft-caption {
          color: var(--muted);
          font-size: 0.9rem;
          line-height: 1.7;
        }

        .candidate-shell {
          padding: 0.9rem 1rem;
          border-radius: 18px;
          background: rgba(255,255,255,0.82);
          border: 1px solid var(--line);
          margin-bottom: 0.7rem;
        }

        .candidate-name {
          margin: 0;
          color: var(--ink);
          font-weight: 800;
          font-size: 1rem;
        }

        .candidate-meta {
          margin-top: 0.18rem;
          color: var(--muted);
          font-size: 0.88rem;
        }

        .back-top-shell {
          text-align: right;
          margin-top: 0.8rem;
        }

        .back-top-shell a {
          display: inline-flex;
          align-items: center;
          gap: 0.45rem;
          padding: 0.48rem 0.88rem;
          border-radius: 999px;
          border: 1px solid var(--line);
          background: rgba(255,255,255,0.92);
          text-decoration: none;
          color: var(--ink);
          font-weight: 700;
        }

        .stTabs [data-baseweb="tab-list"] {
          gap: 0.55rem;
          background: rgba(255,255,255,0.68);
          border: 1px solid var(--line);
          border-radius: 999px;
          padding: 0.35rem;
          margin-bottom: 1rem;
        }

        .stTabs [data-baseweb="tab"] {
          height: auto;
          padding: 0.52rem 1rem;
          border-radius: 999px;
          color: #44536b;
          font-weight: 800;
        }

        .stTabs [aria-selected="true"] {
          background: linear-gradient(135deg, var(--navy) 0%, #2d6bb0 100%);
          color: white !important;
        }

        div[data-baseweb="select"] > div,
        div[data-testid="stTextInputRootElement"] > div,
        div[data-testid="stFileUploaderDropzone"] {
          border-radius: 18px !important;
          background: rgba(255,255,255,0.86) !important;
          border: 1px solid var(--line) !important;
          box-shadow: none !important;
        }

        div[data-testid="stRadio"] > div {
          background: rgba(255,255,255,0.86);
          border: 1px solid var(--line);
          border-radius: 999px;
          padding: 0.35rem 0.5rem;
        }

        .stButton button,
        .stDownloadButton button {
          min-height: 46px;
          border-radius: 15px;
          font-weight: 800;
          border: 1px solid #c9d7ec;
          background: linear-gradient(180deg, #ffffff 0%, #f2f6fb 100%);
          color: var(--ink);
          box-shadow: 0 10px 22px rgba(18, 34, 58, 0.04);
        }

        .stButton button[kind="primary"],
        .stDownloadButton button[kind="primary"] {
          background: linear-gradient(135deg, var(--navy) 0%, #2d6cb4 100%);
          color: white;
          border-color: rgba(31, 75, 132, 0.75);
        }

        .stAlert {
          border-radius: 18px;
          border: 1px solid var(--line);
        }
        </style>
        """,
        unsafe_allow_html=True,
    )


def render_app_header():
    st.markdown(
        """
        <section class="hero-shell">
          <div class="hero-kicker">Business Registry Intelligence</div>
          <h1 class="hero-title">公司資料查詢工具</h1>
          <p class="hero-lead">
            把公司登記、股價、除權息、來源快照和批量匯出整理在同一個工作台裡。
            你查到的結果可以直接下載成 Excel、CSV、PDF，不用再來回切很多網站。
          </p>
          <div class="hero-source-row">
            <span class="source-chip">工商登記：findbiz.nat.gov.tw</span>
            <span class="source-chip">股價：TWSE / TPEX</span>
            <span class="source-chip">除權息：MOPS / Yahoo Finance</span>
            <span class="source-chip">交付格式：Excel / CSV / PDF</span>
          </div>
        </section>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data(show_spinner=False, ttl=300)
def get_update_status() -> dict:
    return fetch_latest_release()


@st.dialog("發現可安裝的新版本")
def render_update_dialog(status: dict):
    latest_version = status.get("latest_version") or "新版"
    st.markdown(
        f"""
        目前版本：`v{APP_VERSION}`  
        偵測到最新版：`{latest_version}`
        """
    )
    st.write("要不要現在就下載並啟動安裝程式？如果你先不更新，這次開啟期間我就先不再打擾你。")

    release_body = (status.get("body") or "").strip()
    if release_body:
        with st.expander("查看本次更新內容", expanded=False):
            st.markdown(release_body)

    c1, c2 = st.columns(2)
    with c1:
        if st.button("稍後再說", key="dismiss_update_dialog_btn", use_container_width=True):
            st.session_state["dismissed_update_version"] = latest_version
            st.rerun()
    with c2:
        if st.button("立刻下載安裝", key="confirm_update_dialog_btn", type="primary", use_container_width=True):
            try:
                with st.spinner("正在下載更新安裝包並準備啟動，請稍候…"):
                    saved_path = download_and_launch_update(
                        status.get("download_url", ""),
                        status.get("asset_name", "CompanyQueryToolSetup.exe"),
                    )
                st.session_state["dismissed_update_version"] = latest_version
                st.success(f"已開始下載並啟動更新安裝包：{saved_path}")
            except Exception as exc:
                st.error(f"更新下載失敗：{exc}")

    release_page = status.get("release_page", "")
    if release_page:
        st.link_button("查看版本說明", release_page, use_container_width=True)


def maybe_prompt_for_update():
    config = load_update_config()
    if not update_is_configured(config):
        return

    status = get_update_status()
    if status.get("error"):
        return
    if not status.get("asset_found"):
        return
    if not status.get("update_available"):
        return

    latest_version = status.get("latest_version") or ""
    if st.session_state.get("dismissed_update_version") == latest_version:
        return

    render_update_dialog(status)


TERMS_NOTICE_MARKDOWN = """
**使用公告與免責聲明**

1. 本工具整合公開資料來源，僅供內部參考、研究與作業輔助使用，**不構成任何法律、財務、投資或其他專業意見**。  
2. 查詢結果可能因資料來源更新時間、網站調整、網路連線、解析誤差或第三方服務異動而與官方資料不同。  
3. **如有任何差異、疑義或爭議，一律以各官方網站、主管機關或原始公告資料為準。**  
4. 使用者應自行就重要資訊再次查核後再行引用、報告、對外提供或據以決策；因使用本工具資料所生之任何損失、誤判、延誤或責任，需由使用者自行承擔。  
5. 若你不同意以上條件，請勿繼續使用本工具。
"""


@st.dialog("使用公告與免責聲明")
def render_terms_dialog():
    st.markdown(TERMS_NOTICE_MARKDOWN)
    c1, c2 = st.columns(2)
    with c1:
        if st.button("我不同意", key="decline_terms_btn", use_container_width=True):
            st.session_state["terms_declined"] = True
            st.rerun()
    with c2:
        if st.button("我已閱讀並同意", key="accept_terms_btn", type="primary", use_container_width=True):
            st.session_state["terms_accepted"] = True
            st.session_state["terms_declined"] = False
            st.rerun()


def require_terms_acceptance():
    if st.session_state.get("terms_accepted"):
        return
    render_terms_dialog()
    if st.session_state.get("terms_declined"):
        st.error("你尚未同意使用公告與免責聲明，因此目前不能使用本工具。")
        st.info("若要繼續使用，請重新整理或重新開啟頁面後按「我已閱讀並同意」。")
    st.stop()


def _display_value(value) -> str:
    if value is None:
        return "—"
    if isinstance(value, list):
        value = "、".join(str(v).strip() for v in value if str(v).strip())
    text = str(value).strip()
    if not text:
        return "—"
    return html.escape(text).replace("\n", "<br/>")


def render_summary_metrics(items: list[tuple[str, str]]):
    columns = st.columns(len(items))
    for col, (label, value) in zip(columns, items):
        col.markdown(
            f"""
            <div class="summary-card">
              <span class="summary-label">{html.escape(str(label))}</span>
              <span class="summary-value">{_display_value(value)}</span>
            </div>
            """,
            unsafe_allow_html=True,
        )


def render_info_card(title: str, icon: str, rows: list[tuple[str, str]], note: str = ""):
    table_rows = "".join(
        f"<tr><th>{html.escape(str(label))}</th><td>{_display_value(value)}</td></tr>"
        for label, value in rows
    )
    note_html = f'<p class="section-note">{html.escape(note)}</p>' if note else ""
    st.markdown(
        f"""
        <section class="section-card">
          <div class="section-head">
            <div class="section-head-left">
              <div class="section-icon">{icon}</div>
              <div>
                <h3 class="section-title">{html.escape(title)}</h3>
                {note_html}
              </div>
            </div>
          </div>
          <table class="info-table">
            <tbody>{table_rows}</tbody>
          </table>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_business_card(lines: list[str]):
    items = "".join(f'<div class="biz-item">{html.escape(item)}</div>' for item in lines)
    st.markdown(
        f"""
        <section class="section-card">
          <div class="section-head">
            <div class="section-head-left">
              <div class="section-icon">📂</div>
              <div>
                <h3 class="section-title">所營事業</h3>
                <p class="section-note">共 {len(lines)} 項，已整理為容易閱讀的清單。</p>
              </div>
            </div>
          </div>
          <div class="biz-grid">{items}</div>
        </section>
        """,
        unsafe_allow_html=True,
    )


def render_dividend_html_table(divs: list[dict], year: int, period_label: str):
    if not divs:
        st.markdown(
            f"""
            <section class="section-card">
              <div class="section-head">
                <div class="section-head-left">
                  <div class="section-icon">🎯</div>
                  <div>
                    <h3 class="section-title">近兩年除權息資訊</h3>
                    <p class="section-note">目前查無這段查詢區間的除權息紀錄。</p>
                  </div>
                </div>
              </div>
              <div class="hint-box">查詢區間：{html.escape(period_label)}，目前無除權息紀錄。</div>
            </section>
            """,
            unsafe_allow_html=True,
        )
        return

    body_rows = "".join(
        (
            "<tr>"
            f"<td>{_display_value(row.get('日期'))}</td>"
            f"<td>{_display_value(row.get('類別'))}</td>"
            f"<td style=\"text-align:right; font-variant-numeric:tabular-nums;\">{_display_value(row.get('現金股利(元)'))}</td>"
            f"<td style=\"text-align:right; font-variant-numeric:tabular-nums;\">{_display_value(row.get('股票股利(元)'))}</td>"
            "</tr>"
        )
        for row in divs
    )

    st.markdown(
        f"""
        <section class="section-card">
          <div class="section-head">
            <div class="section-head-left">
              <div class="section-icon">🎯</div>
              <div>
                <h3 class="section-title">近兩年除權息資訊</h3>
                <p class="section-note">查詢區間：{html.escape(period_label)}，共 {len(divs)} 筆，已整理為可直接閱讀與輸出的表格。</p>
              </div>
            </div>
          </div>
          <table class="dividend-table">
            <thead>
              <tr>
                <th style="width:22%;">日期</th>
                <th style="width:18%;">類別</th>
                <th style="width:30%; text-align:right;">現金股利(元)</th>
                <th style="width:30%; text-align:right;">股票股利(元)</th>
              </tr>
            </thead>
            <tbody>{body_rows}</tbody>
          </table>
        </section>
        """,
        unsafe_allow_html=True,
    )


inject_global_styles()
render_app_header()
require_terms_acceptance()
maybe_prompt_for_update()

# ── 預熱快取 ──────────────────────────────────────────────────────────────────
@st.cache_resource(show_spinner="載入 ISIN 清單…")
def warm_caches():
    init_caches()

warm_caches()

if "single_result" not in st.session_state:
    st.session_state["single_result"] = None
if "single_result_year" not in st.session_state:
    st.session_state["single_result_year"] = None
if "name_candidates" not in st.session_state:
    st.session_state["name_candidates"] = []
if "name_search_query" not in st.session_state:
    st.session_state["name_search_query"] = ""
if "name_search_year" not in st.session_state:
    st.session_state["name_search_year"] = None
if "name_search_price_date" not in st.session_state:
    st.session_state["name_search_price_date"] = None
if "single_query_submit_requested" not in st.session_state:
    st.session_state["single_query_submit_requested"] = False


def clear_single_result():
    st.session_state["single_result"] = None
    st.session_state["single_result_year"] = None


def reset_name_search_state():
    st.session_state["name_candidates"] = []
    st.session_state["name_search_query"] = ""
    st.session_state["name_search_year"] = None
    st.session_state["name_search_price_date"] = None


def default_price_query_date(year: int) -> date_cls:
    return date_cls(year, 12, 31)


@st.cache_data(show_spinner="載入交易日清單…")
def get_trading_day_dates(year: int) -> list[date_cls]:
    days = get_trading_days(year)
    return days or [default_price_query_date(year)]


def get_previous_trading_day(selected_date: date_cls, year: int) -> date_cls:
    trading_days = get_trading_day_dates(year)
    for trading_day in trading_days:
        if trading_day <= selected_date:
            return trading_day
    return trading_days[-1]


def set_single_result(res: dict, year: int):
    st.session_state["single_result"] = res
    st.session_state["single_result_year"] = year


def request_single_query_submit():
    st.session_state["single_query_submit_requested"] = True


@st.cache_data(show_spinner="生成 findbiz 網頁列印 PDF…")
def get_findbiz_web_print_pdf(detail_html: str) -> bytes:
    return findbiz_web_print_pdf(detail_html)


@st.cache_data(show_spinner="生成股票資訊 PDF…")
def get_stock_snapshot_pdf(res: dict, year: int) -> bytes:
    return generate_stock_snapshot_pdf(res, year)


@st.cache_data(show_spinner="生成除權息資訊 PDF…")
def get_dividend_snapshot_pdf(res: dict, year: int) -> bytes:
    return generate_dividend_snapshot_pdf(res, year)


@st.cache_data(show_spinner="生成批量完整報告 PDF…")
def get_batch_report_pdf(results: list[dict], year: int) -> bytes:
    return generate_batch_report_pdf(results, year)


def render_back_to_top_button():
    st.markdown(
        """
        <div class="back-top-shell">
          <a href="#page-top">回到最上面</a>
        </div>
        """,
        unsafe_allow_html=True,
    )


def _director_role_chip(role: str) -> str:
    palette = {
        "董事長": ("#dbeafe", "#1d4ed8"),
        "副董事長": ("#e0f2fe", "#0369a1"),
        "董事": ("#eef2ff", "#4338ca"),
        "獨立董事": ("#ecfdf5", "#047857"),
        "監察人": ("#fff7ed", "#c2410c"),
    }
    bg, fg = palette.get(role or "", ("#f3f4f6", "#374151"))
    return (
        f'<span style="display:inline-block;padding:0.18rem 0.55rem;'
        f'border-radius:999px;background:{bg};color:{fg};font-weight:600;font-size:0.86rem;">'
        f'{role or "—"}</span>'
    )


def _format_share_count(value: str) -> str:
    text = str(value or "").replace(",", "").strip()
    if not text:
        return "—"
    if text.isdigit():
        return f"{int(text):,}"
    return str(value)


def render_directors_table(directors: list[dict]):
    rows_html = "".join(
        (
            "<tr>"
            f"<td>{html.escape(str(row.get('序號', '—') or '—'))}</td>"
            f"<td>{_director_role_chip(row.get('職稱', ''))}</td>"
            f"<td><strong>{html.escape(str(row.get('姓名', '—') or '—'))}</strong></td>"
            f"<td>{html.escape(str(row.get('所代表法人', '') or '—'))}</td>"
            f"<td style=\"text-align:right; font-variant-numeric: tabular-nums;\">{html.escape(_format_share_count(row.get('持有股份數(股)', '')))}</td>"
            "</tr>"
        )
        for row in directors
    )

    st.markdown(
        f"""
        <style>
        .director-table {{
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          overflow: hidden;
          border: 1px solid #dbe3f0;
          border-radius: 16px;
          background: #ffffff;
          box-shadow: 0 8px 22px rgba(15, 23, 42, 0.04);
        }}
        .director-table thead th {{
          background: linear-gradient(180deg, #f8fbff 0%, #eef5ff 100%);
          color: #334155;
          font-weight: 700;
          text-align: left;
          padding: 0.8rem 0.9rem;
          border-bottom: 1px solid #dbe3f0;
        }}
        .director-table tbody td {{
          padding: 0.8rem 0.9rem;
          border-bottom: 1px solid #edf2f7;
          vertical-align: middle;
        }}
        .director-table tbody tr:nth-child(even) {{
          background: #fbfdff;
        }}
        .director-table tbody tr:last-child td {{
          border-bottom: none;
        }}
        </style>
        <table class="director-table">
          <thead>
            <tr>
              <th style="width: 12%;">序號</th>
              <th style="width: 18%;">職稱</th>
              <th style="width: 18%;">姓名</th>
              <th style="width: 32%;">所代表法人</th>
              <th style="width: 20%; text-align:right;">持有股份數(股)</th>
            </tr>
          </thead>
          <tbody>{rows_html}</tbody>
        </table>
        """,
        unsafe_allow_html=True,
    )


@st.cache_data
def get_batch_template_csv_bytes() -> bytes:
    template_df = pd.DataFrame(
        [
            {"統一編號": "34051920", "股票代號": "", "公司名稱": "台達電子工業股份有限公司", "備註": "範例：用統一編號查詢"},
            {"統一編號": "", "股票代號": "1519", "公司名稱": "華城電機股份有限公司", "備註": "範例：用股票代號查詢"},
            {"統一編號": "", "股票代號": "", "公司名稱": "嘉晶電子股份有限公司", "備註": "範例：用公司名稱查詢"},
        ]
    )
    return template_df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


@st.cache_data
def get_batch_template_excel_bytes() -> bytes:
    template_df = pd.DataFrame(
        [
            {"統一編號": "34051920", "股票代號": "", "公司名稱": "台達電子工業股份有限公司", "備註": "範例：用統一編號查詢"},
            {"統一編號": "", "股票代號": "1519", "公司名稱": "華城電機股份有限公司", "備註": "範例：用股票代號查詢"},
            {"統一編號": "", "股票代號": "", "公司名稱": "嘉晶電子股份有限公司", "備註": "範例：用公司名稱查詢"},
        ]
    )
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        template_df.to_excel(writer, index=False, sheet_name="批量查詢範本")
        ws = writer.sheets["批量查詢範本"]
        ws.column_dimensions["A"].width = 18
        ws.column_dimensions["B"].width = 14
        ws.column_dimensions["C"].width = 28
        ws.column_dimensions["D"].width = 28
        ws.freeze_panes = "A2"
    return buf.getvalue()


def render_result_actions(res: dict, year: int):
    co_name = res.get("公司名稱") or res.get("統一編號") or "company"
    uid = res.get("統一編號") or "unknown"
    detail_html = res.get("_detail_html", "") or ""
    snapshot_pdf_bytes = b""
    if detail_html:
        try:
            snapshot_pdf_bytes = get_findbiz_web_print_pdf(detail_html)
        except Exception as e:
            st.warning(f"findbiz 網頁列印 PDF 產生失敗：{e}")

    st.markdown(
        """
        <section class="section-card" style="padding-bottom:1rem;">
          <div class="section-head" style="margin-bottom:0.5rem;">
            <div class="section-head-left">
              <div class="section-icon">⬇️</div>
              <div>
                <h3 class="section-title">下載與留存</h3>
                <p class="section-note">把本次查詢結果整理成可交付檔案，適合留檔、寄送、覆核或內部彙整。</p>
              </div>
            </div>
          </div>
        </section>
        """,
        unsafe_allow_html=True,
    )

    button_count = 4 if snapshot_pdf_bytes else 3

    columns = st.columns(button_count)
    with columns[0]:
        st.download_button(
            "⬇️ 匯出 Excel",
            data=to_excel_bytes([res]),
            file_name=f"{co_name}_{year}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            help="適合後續整理、篩選與內部追蹤。",
        )
    with columns[1]:
        st.download_button(
            "⬇️ 匯出 CSV",
            data=to_csv_bytes([res]),
            file_name=f"{co_name}_{year}.csv",
            mime="text/csv",
            use_container_width=True,
            help="適合匯入其他系統或做快速資料交換。",
        )
    with columns[2]:
        try:
            pdf_bytes = generate_pdf(res, year)
            st.download_button(
                "⬇️ 完整報告 PDF",
                data=pdf_bytes,
                file_name=f"{co_name}_{year}.pdf",
                mime="application/pdf",
                use_container_width=True,
                help="匯出公司資訊、股票資訊、除權息與來源說明的正式報告版 PDF。",
            )
        except Exception as e:
            st.error(f"PDF 產生失敗：{e}")
    next_col = 3
    if snapshot_pdf_bytes:
        with columns[next_col]:
            st.download_button(
                "⬇️ 官方頁列印 PDF",
                data=snapshot_pdf_bytes,
                file_name=f"{co_name}_{uid}_findbiz_web_print.pdf",
                mime="application/pdf",
                use_container_width=True,
                help="直接下載 findbiz 原網站畫面的網頁列印版 PDF。首次生成會稍慢一些。",
            )


def show_vertical(res: dict, year: int):
    company_name = res.get("公司名稱") or "查無公司名稱"
    stock_no = res.get("股票代號", "")
    market = res.get("市場別", "")
    uid = res.get("統一編號", "")
    price_query_date = res.get("股價查詢日期", "") or f"{year}/12/31"
    findbiz_url = res.get("登記資料來源網址") or f"https://findbiz.nat.gov.tw/fts/query/QueryBar/queryInit.do?banNo={uid}"
    price_source_label = res.get("股價資料來源說明", "")
    price_source_url = res.get("股價資料來源網址", "")
    st.markdown(
        f"""
        <section class="section-card">
          <div class="section-head" style="margin-bottom:0.35rem;">
            <div class="section-head-left">
              <div class="section-icon">🏛️</div>
              <div>
                <h3 class="section-title">{html.escape(company_name)}</h3>
                <p class="section-note">查詢年度 {year}，股價查詢日期 {html.escape(price_query_date)}。以下內容已整合公司登記、股票資訊、除權息與來源頁，適合直接檢視與輸出。</p>
              </div>
            </div>
          </div>
        </section>
        """,
        unsafe_allow_html=True,
    )

    summary_items = [
        ("統一編號", res.get("統一編號", "")),
        ("登記現況", res.get("登記現況", "")),
        ("股票代號", stock_no or "未上市櫃"),
        ("市場別", market or "未上市／未上櫃"),
    ]
    if stock_no:
        summary_items.append(("收盤價(元)", res.get("收盤價(元)") or res.get("年底收盤價(元)", "")))
    render_summary_metrics(summary_items)

    top_left, top_right = st.columns(2)
    with top_left:
        render_info_card(
            "公司基本資料",
            "📋",
            [
                ("統一編號", res.get("統一編號", "")),
                ("登記現況", res.get("登記現況", "")),
                ("股權狀況", res.get("股權狀況", "")),
                ("公司名稱", res.get("公司名稱", "")),
                ("章程所訂外文公司名稱", res.get("章程所訂外文公司名稱", "")),
            ],
            note="核心識別欄位與公司名稱資訊。",
        )
        render_info_card(
            "登記資訊",
            "🏢",
            [
                ("代表人姓名", res.get("代表人姓名", "")),
                ("公司所在地", res.get("公司所在地", "")),
                ("登記機關", res.get("登記機關", "")),
                ("核准設立日期", res.get("核准設立日期", "")),
                ("最後核准變更日期", res.get("最後核准變更日期", "")),
            ],
            note="與公司登記主管機關相關的主要資訊。",
        )
    with top_right:
        render_info_card(
            "資本與股份",
            "💰",
            [
                ("資本總額(元)", res.get("資本總額(元)", "")),
                ("實收資本額(元)", res.get("實收資本額(元)", "")),
                ("每股金額(元)", res.get("每股金額(元)", "")),
                ("已發行股份總數(股)", res.get("已發行股份總數(股)", "")),
            ],
            note="資本結構與已發行股份概況。",
        )
        render_info_card(
            "特別股資訊",
            "📌",
            [
                ("複數表決權特別股", res.get("複數表決權特別股", "")),
                ("對於特定事項具否決權特別股", res.get("對於特定事項具否決權特別股", "")),
                ("特別股董監選任限制", res.get("特別股董監選任限制", "")),
            ],
            note="特別股與董事監察人限制事項。",
        )

    biz = res.get("所營事業", "")
    if isinstance(biz, list):
        biz = "\n".join(biz)
    if biz:
        lines = [l.strip() for l in biz.split("\n") if l.strip()]
        render_business_card(lines)

    directors_term = res.get("董監事任期", "")
    directors = res.get("董監事資料", [])
    if isinstance(directors, list) and directors:
        st.markdown(
            f"""
            <section class="section-card" style="padding-bottom:1rem;">
              <div class="section-head">
                <div class="section-head-left">
                  <div class="section-icon">👥</div>
                  <div>
                    <h3 class="section-title">董監事資料</h3>
                    <p class="section-note">{html.escape(directors_term) if directors_term else "已整理最新抓取到的董監事名單與持股資訊。"}</p>
                  </div>
                </div>
              </div>
            </section>
            """,
            unsafe_allow_html=True,
        )
        render_directors_table(directors)

    if stock_no:
        stock_col, source_col = st.columns([1.15, 0.85])
        with stock_col:
            render_info_card(
                "股票資訊",
                "📈",
                [
                    ("股票代號", stock_no),
                    ("市場別", market),
                    ("股價查詢日期", price_query_date),
                    ("實際收盤日期", res.get("實際收盤日期") or res.get("年底收盤日期", "")),
                    ("收盤價(元)", res.get("收盤價(元)") or res.get("年底收盤價(元)", "")),
                ],
                note="未自訂時預設抓取當年底；若遇休市，會回溯到最近一個有成交資料的日期。",
            )
        with source_col:
            link_items = []
            if price_source_label and price_source_url:
                link_items.append(
                    f'<a href="{html.escape(price_source_url, quote=True)}" target="_blank">股價來源：{html.escape(price_source_label)}</a>'
                )
            link_items.append(
                f'<a href="{html.escape(findbiz_url, quote=True)}" target="_blank">查看 findbiz 官方頁面</a>'
            )
            st.markdown(
                f"""
                <section class="section-card">
                  <div class="section-head">
                    <div class="section-head-left">
                      <div class="section-icon">🔗</div>
                      <div>
                        <h3 class="section-title">來源與快照</h3>
                        <p class="section-note">可直接回到官方來源頁，也能把本次查詢內容另存成 PDF 快照。</p>
                      </div>
                    </div>
                  </div>
                  <div class="link-pill-row">{''.join(link_items)}</div>
                  <div style="height:0.8rem;"></div>
                  <div class="hint-box">目前統一編號：{html.escape(uid)}，若只想留存當次畫面，可直接下載上方的 PDF 快照。</div>
                </section>
                """,
                unsafe_allow_html=True,
            )

        try:
            stock_pdf = get_stock_snapshot_pdf(res, year)
            st.download_button(
                "⬇️ 股票資訊快照 PDF",
                data=stock_pdf,
                file_name=f"{res.get('公司名稱') or stock_no}_{stock_no}_{year}_stock_snapshot.pdf",
                mime="application/pdf",
                use_container_width=True,
                help="直接下載本次查詢結果的股票資訊快照 PDF，包含年底收盤日期與年底收盤價。",
            )
        except Exception as e:
            st.warning(f"股票資訊 PDF 產生失敗：{e}")

        divs = res.get("除權息明細", [])
        period_label = res.get("除權息查詢區間") or f"{year - 1}-{year}"
        render_dividend_html_table(divs if isinstance(divs, list) else [], year, period_label)
        try:
            dividend_pdf = get_dividend_snapshot_pdf(res, year)
            st.download_button(
                "⬇️ 除權息快照 PDF",
                data=dividend_pdf,
                file_name=f"{res.get('公司名稱') or stock_no}_{stock_no}_{year}_dividend_snapshot.pdf",
                mime="application/pdf",
                use_container_width=True,
                help="直接下載本次查詢結果的年度除權息快照 PDF，年度會跟目前查詢結果一致。",
            )
        except Exception as e:
            st.warning(f"除權息資訊 PDF 產生失敗：{e}")
    else:
        st.markdown(
            f"""
            <section class="section-card">
              <div class="section-head">
                <div class="section-head-left">
                  <div class="section-icon">ℹ️</div>
                  <div>
                    <h3 class="section-title">市場資訊</h3>
                    <p class="section-note">目前查無上市／上櫃股票代號，系統已保留公司登記資料與來源頁。</p>
                  </div>
                </div>
              </div>
              <div class="hint-box">市場別：{html.escape(market or '未上市／未上櫃')}</div>
            </section>
            """,
            unsafe_allow_html=True,
        )

    if not stock_no:
        st.markdown(
            f"""
            <section class="section-card">
              <div class="section-head">
                <div class="section-head-left">
                  <div class="section-icon">🔗</div>
                  <div>
                    <h3 class="section-title">來源與快照</h3>
                    <p class="section-note">公司登記資料仍可直接回到官方來源頁，方便再次核對。</p>
                  </div>
                </div>
              </div>
              <div class="link-pill-row">
                <a href="{html.escape(findbiz_url, quote=True)}" target="_blank">查看 findbiz 官方頁面</a>
              </div>
              <div style="height:0.8rem;"></div>
              <div class="hint-box">目前統一編號：{html.escape(uid)}，可直接下載公司基本資料 PDF 或 findbiz 網頁列印 PDF 留存。</div>
            </section>
            """,
            unsafe_allow_html=True,
        )

    if res.get("備註"):
        st.warning(res["備註"])


# ══════════════════════════════════════════════════════════════════════════════
# 分頁：單一查詢 ／ 批量查詢
# ══════════════════════════════════════════════════════════════════════════════

tab_single, tab_batch = st.tabs(["📋 單一查詢", "📂 批量查詢"])

# ── 單一查詢 ──────────────────────────────────────────────────────────────────
with tab_single:
    st.markdown(
        """
        <div class="panel-title">單筆查詢工作台</div>
        <p class="panel-subtitle">
          可以用統一編號直接查，也可以先用公司名稱找出同名候選，或直接輸入股票代號查詢。
        </p>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="hint-box" style="margin-bottom:1rem;">
          建議優先使用統一編號查詢，速度最快也最準；如果只知道公司名稱或股票代號，系統也可以協助對應公司資料。
        </div>
        """,
        unsafe_allow_html=True,
    )

    search_mode = st.radio(
        "搜尋方式",
        ["🔢 統一編號", "🏢 公司名稱", "📈 股票代號"],
        horizontal=True,
        label_visibility="collapsed",
    )

    uid_input = ""
    stock_input = ""
    name_input = ""
    c_query, c_setting = st.columns([1.85, 1.15])
    with c_query:
        if search_mode == "🔢 統一編號":
            uid_input = st.text_input(
                "統一編號（8碼）",
                placeholder="例：34051920",
                max_chars=8,
                key="single_uid_input",
                on_change=request_single_query_submit,
            )
        elif search_mode == "📈 股票代號":
            stock_input = st.text_input(
                "股票代號（4-6碼）",
                placeholder="例：1519",
                max_chars=6,
                key="single_stock_input",
                on_change=request_single_query_submit,
            )
        else:
            name_input = st.text_input(
                "公司名稱（含部分名稱）",
                placeholder="例：台達電子",
                key="single_name_input",
                on_change=request_single_query_submit,
            )
        st.caption("輸入完成後可直接按 Enter 開始查詢。")
    with c_setting:
        year = st.selectbox(
            "查詢年度",
            options=list(range(datetime.now().year - 1, datetime.now().year - 11, -1)),
            key="single_query_year",
        )
        use_custom_price_date = st.checkbox("自訂股價日期", value=False, key="single_use_custom_price_date")
        selected_price_date = st.date_input(
            "股價日期",
            value=default_price_query_date(year),
            min_value=date_cls(year, 1, 1),
            max_value=date_cls(year, 12, 31),
            disabled=not use_custom_price_date,
            key="single_selected_price_date",
        )
        if use_custom_price_date:
            aligned_price_date = get_previous_trading_day(selected_price_date, year)
            if aligned_price_date != selected_price_date:
                st.caption(f"若所選日期不是交易日，系統會自動往前對齊到最近交易日：{aligned_price_date.strftime('%Y/%m/%d')}")
            else:
                st.caption("已選擇交易日，查詢時會直接使用這一天。")
        else:
            st.caption("未指定時預設抓該年度 12/31；若當天非交易日，會自動往前對齊最近交易日。")

    search_btn = st.button("🔍 立即查詢", type="primary", use_container_width=True)
    submit_requested = st.session_state.pop("single_query_submit_requested", False)

    actual_price_date = (
        selected_price_date
        if use_custom_price_date
        else default_price_query_date(year)
    )

    if search_btn or submit_requested:
        clear_single_result()

        if search_mode == "🔢 統一編號":
            reset_name_search_state()
            uid = uid_input.strip()
            if not uid or not uid.isdigit() or len(uid) != 8:
                st.warning("請輸入正確的 8 碼統一編號")
            else:
                with st.spinner("查詢中，請稍候（約 5–15 秒）…"):
                    res = query_by_uid(uid, year, price_date=actual_price_date)

                if res.get("公司名稱"):
                    set_single_result(res, year)
                else:
                    st.error(res.get("備註") or "查無資料，請確認統一編號是否正確")

        elif search_mode == "📈 股票代號":
            reset_name_search_state()
            stock_no = stock_input.strip()
            if not stock_no or not stock_no.isdigit() or len(stock_no) < 4:
                st.warning("請輸入正確的 4 至 6 碼股票代號")
            else:
                with st.spinner("查詢中，請稍候（約 5–15 秒）…"):
                    res = query_by_stock_no(stock_no, year, price_date=actual_price_date)

                if res.get("公司名稱"):
                    set_single_result(res, year)
                else:
                    st.error(res.get("備註") or "查無資料，請確認股票代號是否正確")

        else:  # 公司名稱搜尋
            name = name_input.strip()
            if not name:
                reset_name_search_state()
                st.warning("請輸入公司名稱")
            else:
                with st.spinner("搜尋中，請稍候…"):
                    candidates = search_companies_by_name(name)

                if not candidates:
                    reset_name_search_state()
                    st.error("查無符合公司，請確認名稱或改用統一編號")
                elif len(candidates) == 1:
                    reset_name_search_state()
                    with st.spinner("查詢中，請稍候（約 5–15 秒）…"):
                        res = query_by_uid(candidates[0]["ban"], year, price_date=actual_price_date)
                    if res.get("公司名稱"):
                        set_single_result(res, year)
                    else:
                        st.error(res.get("備註") or "查無資料")
                else:
                    st.session_state["name_candidates"] = candidates
                    st.session_state["name_search_query"] = name
                    st.session_state["name_search_year"] = year
                    st.session_state["name_search_price_date"] = actual_price_date

    if search_mode == "🏢 公司名稱" and st.session_state["name_candidates"]:
        candidates = st.session_state["name_candidates"]
        st.markdown(
            f"""
            <section class="section-card">
              <div class="section-head">
                <div class="section-head-left">
                  <div class="section-icon">🧭</div>
                  <div>
                    <h3 class="section-title">同名公司候選清單</h3>
                    <p class="section-note">找到 {len(candidates)} 筆相符結果，目前搜尋詞：{html.escape(st.session_state['name_search_query'])}</p>
                  </div>
                </div>
              </div>
            </section>
            """,
            unsafe_allow_html=True,
        )

        for idx, candidate in enumerate(candidates, start=1):
            row_cols = st.columns([0.7, 3.2, 1.6, 1.5])
            row_cols[0].markdown(f"**{idx}.**")
            row_cols[1].markdown(
                f"""
                <div class="candidate-shell">
                  <p class="candidate-name">{html.escape(candidate['name'])}</p>
                  <div class="candidate-meta">統一編號：{html.escape(candidate['ban'])}</div>
                </div>
                """,
                unsafe_allow_html=True,
            )
            row_cols[2].markdown(
                f"""
                <div class="candidate-shell" style="text-align:center;">
                  <div class="candidate-meta">可直接查詢</div>
                  <p class="candidate-name" style="margin-top:0.18rem;">{html.escape(candidate['ban'])}</p>
                </div>
                """,
                unsafe_allow_html=True,
            )
            if row_cols[3].button("查詢這家公司", key=f"name_candidate_{candidate['ban']}"):
                with st.spinner("查詢中，請稍候（約 5–15 秒）…"):
                    res = query_by_uid(
                        candidate["ban"],
                        st.session_state["name_search_year"] or year,
                        price_date=st.session_state["name_search_price_date"],
                    )
                if res.get("公司名稱"):
                    set_single_result(res, st.session_state["name_search_year"] or year)
                    st.rerun()
                else:
                    st.error(res.get("備註") or "查無資料")

    current_result = st.session_state["single_result"]
    current_year = st.session_state["single_result_year"]
    if current_result and current_year:
        show_vertical(current_result, current_year)
        render_result_actions(current_result, current_year)
        render_back_to_top_button()


# ── 批量查詢 ──────────────────────────────────────────────────────────────────
with tab_batch:
    st.markdown(
        """
        <div class="panel-title">批量查詢工作台</div>
        <p class="panel-subtitle">
          上傳 CSV 或 Excel，欄位可用「統一編號」、「股票代號」或「公司名稱」。
          也可以先下載標準範本，再把要查的公司名單貼進去。
        </p>
        """,
        unsafe_allow_html=True,
    )
    st.markdown(
        """
        <div class="hint-box" style="margin-bottom:1rem;">
          批量查詢前可以先下載範本。每一列只要填「統一編號 / 股票代號 / 公司名稱」其中一種即可；系統會依序優先使用統編、再用股票代號、最後用公司名稱。
        </div>
        """,
        unsafe_allow_html=True,
    )
    st.caption("目前範本欄位：統一編號、股票代號、公司名稱、備註。若你看到只有統編，通常是開到舊下載檔。")
    template_col1, template_col2, _ = st.columns([1.2, 1.2, 3.6])
    with template_col1:
        st.download_button(
            "⬇️ 下載 CSV 範本",
            data=get_batch_template_csv_bytes(),
            file_name=f"公司查詢批次範本_統編股號名稱_v{APP_VERSION}.csv",
            mime="text/csv",
            use_container_width=True,
        )
    with template_col2:
        st.download_button(
            "⬇️ 下載 Excel 範本",
            data=get_batch_template_excel_bytes(),
            file_name=f"公司查詢批次範本_統編股號名稱_v{APP_VERSION}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
        )

    c1, c2, c3 = st.columns([3, 1, 1.3])
    with c1:
        uploaded = st.file_uploader(
            "選擇 CSV 或 Excel",
            type=["csv", "xlsx", "xls"],
            label_visibility="collapsed",
        )
    with c2:
        year_batch = st.selectbox(
            "查詢年度 ",
            options=list(range(datetime.now().year - 1, datetime.now().year - 11, -1)),
            key="year_batch",
        )
    with c3:
        use_custom_batch_price_date = st.checkbox("自訂批次股價日期", value=False, key="use_custom_batch_price_date")
        selected_batch_price_date = st.date_input(
            "批次股價日期",
            value=default_price_query_date(year_batch),
            min_value=date_cls(year_batch, 1, 1),
            max_value=date_cls(year_batch, 12, 31),
            disabled=not use_custom_batch_price_date,
            key="batch_price_date",
        )
        if use_custom_batch_price_date:
            aligned_batch_date = get_previous_trading_day(selected_batch_price_date, year_batch)
            if aligned_batch_date != selected_batch_price_date:
                st.caption(f"若所選日期不是交易日，系統會自動往前對齊到最近交易日：{aligned_batch_date.strftime('%Y/%m/%d')}")
            else:
                st.caption("已選擇交易日，批次查詢時會直接使用這一天。")
        else:
            st.caption("未指定時預設抓該年度 12/31；若當天非交易日，會自動往前對齊最近交易日。")

    actual_batch_price_date = (
        selected_batch_price_date
        if use_custom_batch_price_date
        else default_price_query_date(year_batch)
    )

    batch_btn = st.button(
        "🚀 立即開始批量查詢", type="primary",
        use_container_width=True, disabled=uploaded is None,
    )

    if batch_btn and uploaded:
        df_in = pd.read_excel(uploaded, dtype=str) \
            if uploaded.name.endswith((".xlsx", ".xls")) \
            else pd.read_csv(uploaded, dtype=str)

        batch_requests, used_columns = extract_batch_requests(df_in)
        used_columns_text = "、".join(used_columns) if used_columns else "未辨識到標準欄位，已改用第一欄自動判斷"

        st.info(f"讀取到 **{len(batch_requests)}** 筆有效查詢，使用欄位：{used_columns_text}")

        if not batch_requests:
            st.warning("找不到可查詢的資料列，請確認檔案內至少填有統一編號、股票代號或公司名稱其中一種。")
            st.stop()

        progress = st.progress(0, text="查詢中…")
        status   = st.empty()
        results: list = []

        label_map = {"uid": "統編", "stock": "股號", "name": "名稱"}
        for idx, request_item in enumerate(batch_requests):
            query_value = request_item["query_value"]
            query_label = label_map.get(request_item["query_type"], "查詢")
            status.text(f"[{idx+1}/{len(batch_requests)}] {query_label}：{query_value}…")
            res = run_batch_request(request_item, year_batch, price_date=actual_batch_price_date)
            results.append(res)
            progress.progress((idx + 1) / len(batch_requests), text=f"{idx+1}/{len(batch_requests)}")
            time.sleep(1.0)

        progress.empty()
        status.empty()
        st.success(f"完成！共 {len(results)} 筆")
        render_summary_metrics(
            [
                ("讀取欄位", used_columns_text),
                ("查詢筆數", len(results)),
                ("查詢年度", year_batch),
                ("成功公司數", sum(1 for r in results if r.get("公司名稱"))),
            ]
        )

        flat = [_flatten_result(r) for r in results]
        st.dataframe(pd.DataFrame(flat, columns=RESULT_COLUMNS),
                     use_container_width=True, hide_index=True)

        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        bcol1, bcol2, bcol3 = st.columns(3)
        with bcol1:
            st.download_button(
                "⬇️ 匯出 Excel",
                data=to_excel_bytes(results),
                file_name=f"批量查詢結果_{ts}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True,
            )
        with bcol2:
            st.download_button(
                "⬇️ 匯出 CSV",
                data=to_csv_bytes(results),
                file_name=f"批量查詢結果_{ts}.csv",
                mime="text/csv",
                use_container_width=True,
            )
        with bcol3:
            st.download_button(
                "⬇️ 批量完整報告 PDF",
                data=get_batch_report_pdf(results, year_batch),
                file_name=f"批量完整報告_{ts}.pdf",
                mime="application/pdf",
                use_container_width=True,
            )
