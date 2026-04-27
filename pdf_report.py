#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
pdf_report.py
生成公司查詢結果 PDF 報告（reportlab + 微軟正黑體）
"""

import io
from datetime import datetime

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (
    HRFlowable, PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table, TableStyle,
)

# ── 字型註冊（微軟正黑體 Regular / Bold）
_FONT_REGULAR = "MSJH"
_FONT_BOLD    = "MSJHBd"

def _register_fonts():
    try:
        pdfmetrics.registerFont(TTFont(_FONT_REGULAR, "C:/Windows/Fonts/msjh.ttc",   subfontIndex=0))
        pdfmetrics.registerFont(TTFont(_FONT_BOLD,    "C:/Windows/Fonts/msjhbd.ttc", subfontIndex=0))
    except Exception:
        # fallback：用 mingliu 或跳過
        try:
            pdfmetrics.registerFont(TTFont(_FONT_REGULAR, "C:/Windows/Fonts/mingliu.ttc", subfontIndex=0))
            pdfmetrics.registerFont(TTFont(_FONT_BOLD,    "C:/Windows/Fonts/mingliu.ttc", subfontIndex=0))
        except Exception:
            pass

_register_fonts()

# ── 顏色
_CLR_HEADER  = colors.HexColor("#1a4f8a")   # 深藍
_CLR_SECTION = colors.HexColor("#2d7dd2")   # 段落標題藍
_CLR_ROW_ALT = colors.HexColor("#f0f5fb")   # 表格交替列
_CLR_LINE    = colors.HexColor("#d0d8e4")
_CLR_LABEL   = colors.HexColor("#eef3f8")
_CLR_TEXT    = colors.HexColor("#2a2f36")
_CLR_BANNER  = colors.HexColor("#e9f2fb")

# ── 樣式
def _styles():
    return {
        "title": ParagraphStyle("title",
            fontName=_FONT_BOLD, fontSize=18,
            textColor=_CLR_HEADER, spaceAfter=4),
        "subtitle": ParagraphStyle("subtitle",
            fontName=_FONT_REGULAR, fontSize=10,
            textColor=colors.grey, spaceAfter=10),
        "section": ParagraphStyle("section",
            fontName=_FONT_BOLD, fontSize=12,
            textColor=_CLR_SECTION, spaceBefore=10, spaceAfter=4),
        "body": ParagraphStyle("body",
            fontName=_FONT_REGULAR, fontSize=10, leading=16),
        "small": ParagraphStyle("small",
            fontName=_FONT_REGULAR, fontSize=8,
            textColor=colors.grey),
        "biz": ParagraphStyle("biz",
            fontName=_FONT_REGULAR, fontSize=9, leading=14),
        "gov": ParagraphStyle("gov",
            fontName=_FONT_BOLD, fontSize=18,
            textColor=_CLR_HEADER, spaceAfter=2),
        "service": ParagraphStyle("service",
            fontName=_FONT_BOLD, fontSize=13,
            textColor=_CLR_TEXT, spaceAfter=2),
        "page_title": ParagraphStyle("page_title",
            fontName=_FONT_BOLD, fontSize=16,
            textColor=_CLR_TEXT, spaceAfter=6),
        "meta": ParagraphStyle("meta",
            fontName=_FONT_REGULAR, fontSize=9, leading=13,
            textColor=colors.HexColor("#516173")),
        "tab": ParagraphStyle("tab",
            fontName=_FONT_REGULAR, fontSize=8.5,
            textColor=colors.HexColor("#667789")),
        "table_label": ParagraphStyle("table_label",
            fontName=_FONT_BOLD, fontSize=9.5, leading=13,
            textColor=_CLR_TEXT),
        "table_value": ParagraphStyle("table_value",
            fontName=_FONT_REGULAR, fontSize=9.5, leading=14,
            textColor=_CLR_TEXT),
    }


def _fmt_value(val) -> str:
    if val is None:
        return "—"
    if isinstance(val, list):
        items = [str(v).strip() for v in val if str(v).strip()]
        return "<br/>".join(items) if items else "—"
    text = str(val).strip()
    return text if text else "—"


def _link_markup(url: str, label: str) -> str:
    safe_url = (url or "").strip()
    safe_label = (label or "").strip() or "查看來源"
    if not safe_url:
        return safe_label
    return f'<link href="{safe_url}" color="#1a4f8a"><u>{safe_label}</u></link>'


def _banner(story: list, page_title: str, meta_lines: list[str]) -> None:
    s = _styles()
    header = Table(
        [[
            Paragraph("經濟部", s["gov"]),
            Paragraph("商工登記公示資料查詢服務", s["service"]),
        ]],
        colWidths=[34*mm, 130*mm],
    )
    header.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), _CLR_BANNER),
        ("LINEBELOW", (0, 0), (-1, -1), 0.8, _CLR_HEADER),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]))
    story.append(header)
    story.append(Spacer(1, 6))
    story.append(Paragraph(page_title, s["page_title"]))
    if meta_lines:
        story.append(Paragraph("<br/>".join(meta_lines), s["meta"]))
    story.append(Spacer(1, 4))
    story.append(HRFlowable(width="100%", thickness=0.8, color=_CLR_LINE, spaceAfter=8))


def _official_info_table(rows: list[tuple], col_w=(42*mm, 132*mm)) -> Table:
    s = _styles()
    data = []
    for label, value in rows:
        data.append([
            Paragraph(str(label), s["table_label"]),
            Paragraph(_fmt_value(value), s["table_value"]),
        ])

    tbl = Table(data, colWidths=col_w, repeatRows=0)
    style = [
        ("FONTNAME", (0, 0), (-1, -1), _FONT_REGULAR),
        ("FONTSIZE", (0, 0), (-1, -1), 9.5),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("GRID", (0, 0), (-1, -1), 0.35, _CLR_LINE),
        ("LEFTPADDING", (0, 0), (-1, -1), 7),
        ("RIGHTPADDING", (0, 0), (-1, -1), 7),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("TEXTCOLOR", (0, 0), (-1, -1), _CLR_TEXT),
    ]
    for idx in range(len(data)):
        style.append(("BACKGROUND", (0, idx), (0, idx), _CLR_LABEL))
        if idx % 2 == 1:
            style.append(("BACKGROUND", (1, idx), (1, idx), _CLR_ROW_ALT))
    tbl.setStyle(TableStyle(style))
    return tbl


def _section_heading(title: str) -> Table:
    s = _styles()
    tbl = Table([[Paragraph(title, s["section"])]], colWidths=[174*mm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("LINEBELOW", (0, 0), (-1, -1), 0.8, _CLR_HEADER),
        ("LEFTPADDING", (0, 0), (-1, -1), 0),
        ("RIGHTPADDING", (0, 0), (-1, -1), 0),
        ("TOPPADDING", (0, 0), (-1, -1), 0),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 2),
    ]))
    return tbl


def _footer(story: list, lines: list[str]) -> None:
    s = _styles()
    story.append(Spacer(1, 8))
    story.append(HRFlowable(width="100%", thickness=0.5, color=_CLR_LINE, spaceAfter=4))
    story.append(Paragraph("<br/>".join(lines), s["small"]))


def _summary_cards(items: list[tuple[str, str]], total_width=174*mm) -> Table:
    s = _styles()
    count = max(len(items), 1)
    col_w = total_width / count
    cells = []
    for label, value in items:
        body = (
            f'<para align="center"><font name="{_FONT_REGULAR}" size="8.5" color="#5b6b7d">{label}</font><br/>'
            f'<font name="{_FONT_BOLD}" size="13" color="#1f2937">{_fmt_value(value)}</font></para>'
        )
        cells.append(Paragraph(body, s["body"]))

    tbl = Table([cells], colWidths=[col_w] * count)
    style = [
        ("BACKGROUND", (0, 0), (-1, -1), colors.white),
        ("BOX", (0, 0), (-1, -1), 0.35, _CLR_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 8),
        ("RIGHTPADDING", (0, 0), (-1, -1), 8),
        ("TOPPADDING", (0, 0), (-1, -1), 8),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
    ]
    for idx in range(count - 1):
        style.append(("LINEAFTER", (idx, 0), (idx, 0), 0.35, _CLR_LINE))
    tbl.setStyle(TableStyle(style))
    return tbl


def _notice_box(text: str) -> Table:
    s = _styles()
    tbl = Table([[Paragraph(_fmt_value(text), s["meta"])]], colWidths=[174*mm])
    tbl.setStyle(TableStyle([
        ("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#f8fbff")),
        ("BOX", (0, 0), (-1, -1), 0.5, colors.HexColor("#cfe0f3")),
        ("LEFTPADDING", (0, 0), (-1, -1), 10),
        ("RIGHTPADDING", (0, 0), (-1, -1), 10),
        ("TOPPADDING", (0, 0), (-1, -1), 7),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
    ]))
    return tbl


def _business_table(items: list[str]) -> Table:
    s = _styles()
    data = [[
        Paragraph("項次", s["table_label"]),
        Paragraph("所營事業資料", s["table_label"]),
    ]]
    for idx, item in enumerate(items, start=1):
        data.append([
            Paragraph(str(idx), s["table_value"]),
            Paragraph(_fmt_value(item), s["table_value"]),
        ])

    tbl = Table(data, colWidths=[18*mm, 156*mm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), _CLR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, _CLR_LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for idx in range(1, len(data)):
        if idx % 2 == 0:
            style.append(("BACKGROUND", (0, idx), (-1, idx), _CLR_ROW_ALT))
    tbl.setStyle(TableStyle(style))
    return tbl


def _source_dividend_table(divs: list[dict]) -> Table:
    s = _styles()
    data = [[
        Paragraph("日期", s["table_label"]),
        Paragraph("類別", s["table_label"]),
        Paragraph("現金股利(元)", s["table_label"]),
        Paragraph("股票股利(元)", s["table_label"]),
    ]]
    for d in divs:
        data.append([
            Paragraph(_fmt_value(d.get("日期")), s["table_value"]),
            Paragraph(_fmt_value(d.get("類別")), s["table_value"]),
            Paragraph(_fmt_value(d.get("現金股利(元)")), s["table_value"]),
            Paragraph(_fmt_value(d.get("股票股利(元)")), s["table_value"]),
        ])

    tbl = Table(data, colWidths=[38*mm, 36*mm, 50*mm, 50*mm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), _CLR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, _CLR_LINE),
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
    ]
    for idx in range(1, len(data)):
        if idx % 2 == 0:
            style.append(("BACKGROUND", (0, idx), (-1, idx), _CLR_ROW_ALT))
    tbl.setStyle(TableStyle(style))
    return tbl


def _director_table(rows: list[dict]) -> Table:
    s = _styles()
    def _share_text(value) -> str:
        text = str(value or "").replace(",", "").strip()
        if text.isdigit():
            return f"{int(text):,}"
        return text or "—"

    data = [[
        Paragraph("序號", s["table_label"]),
        Paragraph("職稱", s["table_label"]),
        Paragraph("姓名", s["table_label"]),
        Paragraph("所代表法人", s["table_label"]),
        Paragraph("持有股份數(股)", s["table_label"]),
    ]]
    for row in rows:
        data.append([
            Paragraph(_fmt_value(row.get("序號")), s["table_value"]),
            Paragraph(f"<b>{_fmt_value(row.get('職稱'))}</b>", s["table_value"]),
            Paragraph(f"<b>{_fmt_value(row.get('姓名'))}</b>", s["table_value"]),
            Paragraph(_fmt_value(row.get("所代表法人")), s["table_value"]),
            Paragraph(_share_text(row.get("持有股份數(股)")), s["table_value"]),
        ])

    tbl = Table(data, colWidths=[18*mm, 28*mm, 28*mm, 60*mm, 40*mm], repeatRows=1)
    style = [
        ("BACKGROUND", (0, 0), (-1, 0), _CLR_HEADER),
        ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
        ("GRID", (0, 0), (-1, -1), 0.35, _CLR_LINE),
        ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ("LEFTPADDING", (0, 0), (-1, -1), 6),
        ("RIGHTPADDING", (0, 0), (-1, -1), 6),
        ("TOPPADDING", (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("ALIGN", (0, 0), (0, -1), "CENTER"),
        ("ALIGN", (4, 1), (4, -1), "RIGHT"),
    ]
    for idx in range(1, len(data)):
        if idx % 2 == 0:
            style.append(("BACKGROUND", (0, idx), (-1, idx), _CLR_ROW_ALT))
    tbl.setStyle(TableStyle(style))
    return tbl


def _info_table(rows: list[tuple], col_w=(55*mm, 105*mm)) -> Table:
    """製作 label | value 直列資訊表格"""
    data = [[Paragraph(f"<b>{lbl}</b>", _styles()["body"]),
             Paragraph(str(val) if val else "—", _styles()["body"])]
            for lbl, val in rows]

    tbl = Table(data, colWidths=col_w)
    style = [
        ("FONTNAME",    (0, 0), (-1, -1), _FONT_REGULAR),
        ("FONTSIZE",    (0, 0), (-1, -1), 10),
        ("VALIGN",      (0, 0), (-1, -1), "TOP"),
        ("TOPPADDING",  (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (0, -1), 6),
        ("LEFTPADDING", (1, 0), (1, -1), 10),
        ("LINEBELOW",   (0, 0), (-1, -1), 0.3, _CLR_LINE),
    ]
    # 交替列底色
    for i in range(0, len(data), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), _CLR_ROW_ALT))

    tbl.setStyle(TableStyle(style))
    return tbl


def _div_table(divs: list[dict]) -> Table:
    """除權息明細表格"""
    headers = ["除息/除權日", "類別", "現金股利(元)", "股票股利(元)"]
    col_w   = [45*mm, 25*mm, 40*mm, 40*mm]

    def _p(txt, bold=False):
        s = _styles()["body"]
        fn = _FONT_BOLD if bold else _FONT_REGULAR
        return Paragraph(str(txt) if txt else "—",
                         ParagraphStyle("_", fontName=fn, fontSize=10))

    data = [[_p(h, bold=True) for h in headers]]
    for d in divs:
        data.append([
            _p(d.get("日期", "")),
            _p(d.get("類別", "")),
            _p(d.get("現金股利(元)", "")),
            _p(d.get("股票股利(元)", "")),
        ])

    tbl = Table(data, colWidths=col_w)
    style = [
        ("BACKGROUND",    (0, 0), (-1, 0), _CLR_HEADER),
        ("TEXTCOLOR",     (0, 0), (-1, 0), colors.white),
        ("FONTNAME",      (0, 0), (-1, -1), _FONT_REGULAR),
        ("FONTSIZE",      (0, 0), (-1, -1), 10),
        ("VALIGN",        (0, 0), (-1, -1), "MIDDLE"),
        ("TOPPADDING",    (0, 0), (-1, -1), 5),
        ("BOTTOMPADDING", (0, 0), (-1, -1), 5),
        ("LINEBELOW",     (0, 0), (-1, -1), 0.3, _CLR_LINE),
        ("GRID",          (0, 0), (-1, -1), 0.3, _CLR_LINE),
    ]
    for i in range(2, len(data), 2):
        style.append(("BACKGROUND", (0, i), (-1, i), _CLR_ROW_ALT))
    tbl.setStyle(TableStyle(style))
    return tbl


def _dividend_period_label(result: dict, year: int) -> str:
    return result.get("除權息查詢區間") or f"{year-1}-{year}"


def _append_company_report_story(story: list, result: dict, year: int) -> None:
    snapshot_at = result.get("_snapshot_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    co_name = result.get("公司名稱") or "—"
    uid = result.get("統一編號") or "—"
    stock_no = result.get("股票代號")
    market = result.get("市場別")
    price_query_date = result.get("股價查詢日期") or f"{year}/12/31"
    findbiz_url = result.get("登記資料來源網址") or "—"
    price_source_label = result.get("股價資料來源說明") or "TWSE／TPEX 公開資料"
    price_source_url = result.get("股價資料來源網址") or ""
    query_page_label = result.get("股價友善查詢說明") or ""
    query_page_url = result.get("股價友善查詢網址") or ""
    period_label = _dividend_period_label(result, year)
    suffix = ".TW" if "TWSE" in (market or "") else ".TWO"
    yahoo_div_url = f"https://tw.stock.yahoo.com/quote/{stock_no}{suffix}/dividend" if stock_no else ""
    mops_url = "https://mops.twse.com.tw/mops/web/t05st09"

    meta_lines = [
        f"統一編號：{uid}　　公司名稱：{co_name}",
        f"查詢年度：{year}　　產製時間：{snapshot_at}",
    ]
    if stock_no:
        meta_lines.append(f"股票代號：{stock_no}　　市場別：{market or '—'}")

    _banner(story, "公司資料查詢報告", meta_lines)
    story.append(_section_heading("公司基本資料"))
    story.append(Spacer(1, 4))
    basic_rows = [
        ("統一編號",                       result.get("統一編號")),
        ("登記現況",                       result.get("登記現況")),
        ("股權狀況",                       result.get("股權狀況")),
        ("公司名稱",                       result.get("公司名稱")),
        ("章程所訂外文公司名稱",           result.get("章程所訂外文公司名稱")),
        ("資本總額(元)",                   result.get("資本總額(元)")),
        ("實收資本額(元)",                 result.get("實收資本額(元)")),
        ("每股金額(元)",                   result.get("每股金額(元)")),
        ("已發行股份總數(股)",             result.get("已發行股份總數(股)")),
        ("代表人姓名",                     result.get("代表人姓名")),
        ("公司所在地",                     result.get("公司所在地")),
        ("登記機關",                       result.get("登記機關")),
        ("核准設立日期",                   result.get("核准設立日期")),
        ("最後核准變更日期",               result.get("最後核准變更日期")),
        ("複數表決權特別股",               result.get("複數表決權特別股")),
        ("對於特定事項具否決權特別股",     result.get("對於特定事項具否決權特別股")),
        ("特別股董監選任限制",             result.get("特別股董監選任限制")),
        ("董監事任期",                     result.get("董監事任期")),
    ]
    story.append(_official_info_table(basic_rows, col_w=(42*mm, 132*mm)))

    directors = result.get("董監事資料", [])
    if isinstance(directors, list) and directors:
        story.append(Spacer(1, 8))
        story.append(_section_heading("董監事資料"))
        story.append(Spacer(1, 4))
        story.append(_director_table(directors))

    biz = result.get("所營事業", [])
    if biz:
        story.append(Spacer(1, 8))
        story.append(_section_heading("所營事業"))
        story.append(Spacer(1, 4))
        biz_list = biz if isinstance(biz, list) else str(biz).split("\n")
        biz_items = [str(b).strip() for b in biz_list if str(b).strip()]
        story.append(_business_table(biz_items))

    if stock_no:
        story.append(Spacer(1, 8))
        story.append(_section_heading("股票資訊"))
        story.append(Spacer(1, 4))
        stock_rows = [
            ("股票代號", stock_no),
            ("市場別", market),
            ("商品類型", result.get("商品類型")),
            ("發行地", result.get("發行地")),
            ("ISIN Code", result.get("ISIN Code")),
            ("股價查詢日期", price_query_date),
            ("實際收盤日期", result.get("實際收盤日期") or result.get("年底收盤日期")),
            ("收盤價", f"{result.get('收盤價(元)') or result.get('年底收盤價(元)', '—')} 元"),
            ("股價資料來源", _link_markup(price_source_url, price_source_label)),
        ]
        story.append(_official_info_table(stock_rows, col_w=(42*mm, 132*mm)))

    story.append(Spacer(1, 8))
    story.append(_section_heading(f"近兩年除權息資訊（{period_label}）"))
    story.append(Spacer(1, 4))
    divs = result.get("除權息明細", [])
    if isinstance(divs, list) and divs:
        story.append(_div_table(divs))
    else:
        story.append(_official_info_table([("狀態", f"近兩年（{period_label}）無除權息紀錄")], col_w=(42*mm, 132*mm)))

    story.append(Spacer(1, 8))
    story.append(_section_heading("來源說明"))
    story.append(Spacer(1, 4))
    source_rows = [
        ("公司登記資料", _link_markup(findbiz_url, "查看 findbiz 官方頁面")),
        ("股價資料來源", _link_markup(price_source_url, price_source_label) if stock_no else "—"),
        ("股價友善查詢頁", _link_markup(query_page_url, query_page_label) if stock_no and query_page_url else "—"),
        (
            "除權息資料來源",
            f"{_link_markup(yahoo_div_url, '查看 Yahoo 股利頁')}<br/>{_link_markup(mops_url, '查看 MOPS 查詢頁')}"
            if stock_no else "—",
        ),
        ("列印來源頁", _link_markup(result.get("列印連結"), "查看來源頁面")),
    ]
    story.append(_official_info_table(source_rows, col_w=(42*mm, 132*mm)))

    note = result.get("備註")
    if note:
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"備註：{note}", _styles()["small"]))

    _footer(story, [
        "資料來源：findbiz.nat.gov.tw（經濟部商工登記公示資料查詢服務）／ TWSE ／ TPEX ／ Yahoo Finance ／ MOPS",
        f"本報告依本次查詢結果產製，對應公司：{co_name}",
    ])


def generate_pdf(result: dict, year: int) -> bytes:
    """
    將 query_by_uid 回傳的 result dict 轉成 PDF bytes。
    用法：
        pdf_bytes = generate_pdf(res, 2024)
        with open("report.pdf","wb") as f: f.write(pdf_bytes)
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=20*mm, rightMargin=20*mm,
        topMargin=18*mm, bottomMargin=18*mm,
    )
    story = []
    _append_company_report_story(story, result, year)
    doc.build(story)
    return buf.getvalue()


def generate_batch_report_pdf(results: list[dict], year: int) -> bytes:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm,
        topMargin=14*mm, bottomMargin=14*mm,
    )
    story = []
    cleaned = [res for res in results if isinstance(res, dict)]
    for idx, result in enumerate(cleaned):
        _append_company_report_story(story, result, year)
        if idx < len(cleaned) - 1:
            story.append(PageBreak())
    doc.build(story)
    return buf.getvalue()


def generate_findbiz_snapshot_pdf(result: dict) -> bytes:
    """
    依本次抓到的 findbiz 公司詳細頁資料，產生較接近來源頁的 PDF 快照。
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm,
        topMargin=14*mm, bottomMargin=14*mm,
    )
    story = []

    source_url = result.get("_detail_url", "")
    snapshot_at = result.get("_snapshot_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    co_name = result.get("公司名稱", "") or "公司基本資料"

    _banner(story, "公司基本資料", [
        f"統一編號：{result.get('統一編號', '—')}　　公司名稱：{co_name}",
        f"快照時間：{snapshot_at}",
    ])
    story.append(_section_heading("公司基本資料"))

    source_rows = [
        ("統一編號", result.get("統一編號")),
        ("登記現況", result.get("登記現況")),
        ("股權狀況", result.get("股權狀況")),
        ("公司名稱", result.get("公司名稱")),
        ("章程所訂外文公司名稱", result.get("章程所訂外文公司名稱")),
        ("資本總額(元)", result.get("資本總額(元)")),
        ("實收資本額(元)", result.get("實收資本額(元)")),
        ("每股金額(元)", result.get("每股金額(元)")),
        ("已發行股份總數(股)", result.get("已發行股份總數(股)")),
        ("代表人姓名", result.get("代表人姓名")),
        ("公司所在地", result.get("公司所在地")),
        ("登記機關", result.get("登記機關")),
        ("核准設立日期", result.get("核准設立日期")),
        ("最後核准變更日期", result.get("最後核准變更日期")),
        ("複數表決權特別股", result.get("複數表決權特別股")),
        ("對於特定事項具否決權特別股", result.get("對於特定事項具否決權特別股")),
        ("特別股董監選任限制", result.get("特別股董監選任限制")),
    ]
    story.append(Spacer(1, 4))
    story.append(_official_info_table(source_rows, col_w=(40*mm, 134*mm)))

    biz = result.get("所營事業", "")
    if biz:
        story.append(Spacer(1, 8))
        story.append(_section_heading("所營事業資料"))
        story.append(Spacer(1, 4))
        biz_list = biz if isinstance(biz, list) else str(biz).split("\n")
        biz_items = [str(b).strip() for b in biz_list if str(b).strip()]
        story.append(_business_table(biz_items))

    note = result.get("備註")
    if note:
        story.append(Spacer(1, 8))
        story.append(Paragraph(f"備註：{note}", _styles()["small"]))

    _footer(story, [
        f"資料來源：findbiz.nat.gov.tw（經濟部商工登記公示資料查詢服務）",
        f"來源網址：{source_url or '—'}",
    ])

    doc.build(story)
    return buf.getvalue()


def generate_stock_snapshot_pdf(result: dict, year: int) -> bytes:
    """
    依本次查詢結果產生股票資訊 PDF 快照。
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm,
        topMargin=14*mm, bottomMargin=14*mm,
    )
    story = []

    stock_no = result.get("股票代號") or "—"
    market = result.get("市場別") or "—"
    co_name = result.get("公司名稱") or "—"
    snapshot_at = result.get("_snapshot_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    price_query_date = result.get("股價查詢日期") or f"{year}/12/31"
    suffix = ".TW" if "TWSE" in market else ".TWO"
    yahoo_url = f"https://tw.stock.yahoo.com/quote/{stock_no}{suffix}" if stock_no != "—" else "—"
    mops_url = "https://mops.twse.com.tw/mops/web/t05st09"
    price_source_label = result.get("股價資料來源說明") or "TWSE／TPEX 公開資料"
    price_source_url = result.get("股價資料來源網址") or "—"
    query_page_label = result.get("股價友善查詢說明") or ""
    query_page_url = result.get("股價友善查詢網址") or "—"

    _banner(story, "股票資訊快照", [
        f"公司名稱：{co_name}　　股票代號：{stock_no}",
        f"查詢年度：{year}　　股價查詢日期：{price_query_date}　　快照時間：{snapshot_at}",
    ])
    story.append(_summary_cards([
        ("股票代號", stock_no),
        ("市場別", market),
        ("發行地", result.get("發行地") or "—"),
        ("股價查詢日期", price_query_date),
        ("收盤價(元)", result.get("收盤價(元)") or result.get("年底收盤價(元)")),
    ]))
    story.append(Spacer(1, 8))
    story.append(_section_heading("股票基本資訊"))
    story.append(Spacer(1, 4))
    story.append(_official_info_table([
        ("公司名稱", co_name),
        ("股票代號", stock_no),
        ("市場別", market),
        ("商品類型", result.get("商品類型")),
        ("發行地", result.get("發行地")),
        ("ISIN Code", result.get("ISIN Code")),
        ("股價查詢日期", price_query_date),
        ("實際收盤日期", result.get("實際收盤日期") or result.get("年底收盤日期")),
        ("收盤價(元)", result.get("收盤價(元)") or result.get("年底收盤價(元)")),
        ("查詢年度", year),
    ]))
    story.append(Spacer(1, 8))
    story.append(_notice_box(
        f"本頁為本次查詢結果快照，若需回原站交叉確認，可使用下方來源連結或至 MOPS 輸入代號 {stock_no}。"
    ))
    story.append(Spacer(1, 8))
    story.append(_section_heading("來源說明"))
    story.append(Spacer(1, 4))
    story.append(_official_info_table([
        ("股價資料來源", _link_markup(price_source_url, price_source_label)),
        ("股價友善查詢頁", _link_markup(query_page_url, query_page_label) if query_page_url != "—" and query_page_label else "—"),
        ("Yahoo Finance", _link_markup(yahoo_url, "查看 Yahoo Finance")),
        ("MOPS 除權息查詢", _link_markup(mops_url, "查看 MOPS 查詢頁")),
        ("查詢提示", f"若需交叉確認 MOPS 原站，可於查詢欄輸入代號 {stock_no}"),
    ]))

    _footer(story, [
        "資料整理來源：TWSE／TPEX 公開資料、Yahoo Finance、MOPS",
        f"本快照依本次查詢結果產製，對應公司：{co_name}",
    ])

    doc.build(story)
    return buf.getvalue()


def generate_dividend_snapshot_pdf(result: dict, year: int) -> bytes:
    """
    依本次查詢結果產生年度除權息 PDF 快照。
    """
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=A4,
        leftMargin=16*mm, rightMargin=16*mm,
        topMargin=14*mm, bottomMargin=14*mm,
    )
    story = []

    stock_no = result.get("股票代號") or "—"
    market = result.get("市場別") or "—"
    co_name = result.get("公司名稱") or "—"
    snapshot_at = result.get("_snapshot_at", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
    suffix = ".TW" if "TWSE" in market else ".TWO"
    yahoo_url = f"https://tw.stock.yahoo.com/quote/{stock_no}{suffix}/dividend" if stock_no != "—" else "—"
    mops_url = "https://mops.twse.com.tw/mops/web/t05st09"
    query_page_label = result.get("股價友善查詢說明") or ""
    query_page_url = result.get("股價友善查詢網址") or "—"
    divs = result.get("除權息明細", [])
    period_label = _dividend_period_label(result, year)

    _banner(story, "近兩年除權息資訊", [
        f"公司名稱：{co_name}　　股票代號：{stock_no}　　市場別：{market}",
        f"查詢區間：{period_label}　　快照時間：{snapshot_at}",
    ])
    div_count = len(divs) if isinstance(divs, list) else 0
    cash_total = "—"
    stock_total = "—"
    if isinstance(divs, list) and divs:
        def _safe_sum(key: str) -> str:
            total = 0.0
            has_value = False
            for item in divs:
                value = str(item.get(key, "")).strip()
                if not value or value == "—":
                    continue
                try:
                    total += float(value)
                    has_value = True
                except Exception:
                    continue
            return f"{total:g}" if has_value else "—"

        cash_total = _safe_sum("現金股利(元)")
        stock_total = _safe_sum("股票股利(元)")

    story.append(_summary_cards([
        ("查詢區間", period_label),
        ("除權息筆數", div_count),
        ("發行地", result.get("發行地") or "—"),
        ("現金股利合計(元)", cash_total),
        ("股票股利合計(元)", stock_total),
    ]))
    story.append(Spacer(1, 8))
    story.append(_section_heading("近兩年除權息明細"))
    story.append(Spacer(1, 4))
    if isinstance(divs, list) and divs:
        story.append(_source_dividend_table(divs))
    else:
        story.append(_official_info_table([
            ("狀態", f"近兩年（{period_label}）無除權息紀錄"),
        ]))

    story.append(Spacer(1, 8))
    story.append(_notice_box(
        f"本頁整理近兩年（{period_label}）除權息資訊；如需再次核對官方內容，可至下方來源頁或於 MOPS 輸入代號 {stock_no} 查詢。"
    ))
    story.append(Spacer(1, 8))
    story.append(_section_heading("來源說明"))
    story.append(Spacer(1, 4))
    story.append(_official_info_table([
        ("股價友善查詢頁", _link_markup(query_page_url, query_page_label) if query_page_url != "—" and query_page_label else "—"),
        ("Yahoo Finance 歷史股利", _link_markup(yahoo_url, "查看 Yahoo 股利頁")),
        ("MOPS 除權息查詢", _link_markup(mops_url, "查看 MOPS 查詢頁")),
        ("查詢提示", f"若需交叉確認 MOPS 原站，可於查詢欄輸入代號 {stock_no}"),
    ]))

    _footer(story, [
        "資料整理來源：Yahoo Finance、MOPS",
        f"本快照依本次查詢結果產製，查詢區間：{period_label}",
    ])

    doc.build(story)
    return buf.getvalue()
