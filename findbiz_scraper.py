"""
findbiz_scraper.py
從 findbiz.nat.gov.tw 抓取公司登記資料

流程：
  1. GET /queryInit  → 取得 session
  2. POST /queryList（infoType=D, qryCond=統一編號）→ 取得 disj token
  3. 組合 detail URL（objectId = base64('HC'+BAN)）
  4. GET detail page → 解析 td.txt_td 欄位
"""

import base64
import re
import time
from datetime import datetime
from typing import Optional

import requests
import urllib3
from bs4 import BeautifulSoup

urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

BASE = "https://findbiz.nat.gov.tw"
_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept-Language": "zh-TW,zh;q=0.9",
}


# ─── 主要公開函式 ────────────────────────────────────────────────

def search_companies_by_name(name: str) -> list:
    """
    以公司名稱搜尋，回傳 list of dict：[{'ban': '12345678', 'name': '公司名稱'}, ...]
    最多回傳前 20 筆。
    """
    results = []
    try:
        s = requests.Session()
        s.headers.update(_HEADERS)

        r0 = s.get(f"{BASE}/fts/query/QueryBar/queryInit.do", verify=False, timeout=12)
        r0.raise_for_status()
        form_action = BASE + BeautifulSoup(r0.content, "html.parser").find("form")["action"]
        time.sleep(0.4)

        r1 = s.post(
            form_action,
            data={"qryCond": name, "fhl": "zh_TW", "infoType": "D", "isAlive": "all"},
            headers={"Referer": r0.url, "Origin": BASE,
                     "Content-Type": "application/x-www-form-urlencoded"},
            verify=False, timeout=12,
        )
        soup = BeautifulSoup(r1.content, "html.parser")

        # 每個公司連結形如 href="...banNo=12345678..."，連結文字是公司名稱
        seen = set()
        for a in soup.find_all("a", href=True):
            href = a.get("href", "")
            m_ban = re.search(r"banNo=(\d{8})", href)
            if not m_ban:
                continue
            ban_val = m_ban.group(1)
            if ban_val in seen:
                continue
            seen.add(ban_val)
            co_name = a.get_text(strip=True)
            if co_name:
                results.append({"ban": ban_val, "name": co_name})
            if len(results) >= 20:
                break
    except Exception:
        pass
    return results


def scrape_company(ban: str) -> dict:
    """
    查詢統一編號 ban 的公司登記資料。

    Returns
    -------
    dict  包含以下欄位（若抓不到則為空字串）：
        統一編號, 公司名稱, 章程所訂外文公司名稱,
        資本總額(元), 實收資本額(元), 每股金額(元), 已發行股份總數(股),
        代表人姓名, 公司所在地, 登記機關,
        核准設立日期, 最後核准變更日期,
        複數表決權特別股, 對於特定事項具否決權特別股,
        特別股董監選任限制, 董監事任期, 董監事資料,
        登記現況, 股權狀況, 所營事業,
        _detail_url, _print_url, _share_url, _error
    """
    result = _empty_result(ban)

    try:
        s = requests.Session()
        s.headers.update(_HEADERS)

        # Step 1: 首頁取得 session
        r0 = s.get(f"{BASE}/fts/query/QueryBar/queryInit.do", verify=False, timeout=12)
        r0.raise_for_status()
        form_action = BASE + BeautifulSoup(r0.content, "html.parser").find("form")["action"]
        time.sleep(0.4)

        # Step 2: 搜尋取得 disj（最多重試 3 次）
        disj = None
        for attempt in range(3):
            r1 = s.post(
                form_action,
                data={"qryCond": ban, "fhl": "zh_TW", "infoType": "D", "isAlive": "all"},
                headers={"Referer": r0.url, "Origin": BASE,
                         "Content-Type": "application/x-www-form-urlencoded"},
                verify=False,
                timeout=12,
            )
            all_hrefs = " ".join(
                a.get("href", "") for a in BeautifulSoup(r1.content, "html.parser").find_all("a", href=True)
            )
            m = re.search(r"disj=([A-Fa-f0-9]{20,})", all_hrefs)
            if m:
                disj = m.group(1)
                break
            time.sleep(1.2)

        if not disj:
            result["_error"] = "無法取得認證 token（disj），請稍後再試"
            return result

        # Step 3: 組 detail URL
        oid = base64.b64encode(("HC" + ban).encode()).decode()
        detail_url = (
            f"{BASE}/fts/query/QueryCmpyDetail/queryCmpyDetail.do"
            f"?objectId={oid}&banNo={ban}&disj={disj}&fhl=zh_TW"
        )
        print_url = detail_url
        share_url = f"{BASE}/fts/query/QueryBar/queryInit.do?banNo={ban}"

        result["_detail_url"] = detail_url
        result["_print_url"] = print_url
        result["_share_url"] = share_url

        # Step 4: 取 detail 頁面
        r2 = s.get(detail_url, headers={"Referer": r1.url}, verify=False, timeout=15)
        r2.encoding = "utf-8"
        result["_detail_html"] = r2.text
        result["_snapshot_at"] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        soup = BeautifulSoup(r2.text, "html.parser")

        # 確認非錯誤頁
        title = soup.title.text.strip() if soup.title else ""
        if "錯誤" in title or not soup.find("td", class_="txt_td"):
            result["_error"] = f"頁面錯誤或找不到資料（title={title}）"
            return result

        # Step 5: 解析 td.txt_td → 下一個 td
        raw = {}
        for label_td in soup.find_all("td", class_="txt_td"):
            key = label_td.get_text(separator=" ", strip=True)
            val_td = label_td.find_next_sibling("td")
            if not key or not val_td:
                continue
            val = val_td.get_text(separator="\n", strip=True)
            if key not in raw:           # 保留第一次出現
                raw[key] = val

        # Step 6: 填入目標欄位
        _fill(result, raw)

        # 所營事業清理
        result["所營事業"] = _clean_business(raw.get("所營事業資料", ""))
        result["董監事任期"], result["董監事資料"] = _parse_directors(soup)

    except Exception as exc:
        result["_error"] = str(exc)

    return result


# ─── 內部輔助函式 ────────────────────────────────────────────────

def _empty_result(ban: str) -> dict:
    return {
        "統一編號": ban,
        "公司名稱": "",
        "章程所訂外文公司名稱": "",
        "資本總額(元)": "",
        "實收資本額(元)": "",
        "每股金額(元)": "",
        "已發行股份總數(股)": "",
        "代表人姓名": "",
        "公司所在地": "",
        "登記機關": "",
        "核准設立日期": "",
        "最後核准變更日期": "",
        "複數表決權特別股": "",
        "對於特定事項具否決權特別股": "",
        "特別股董監選任限制": "",
        "董監事任期": "",
        "董監事資料": [],
        "登記現況": "",
        "股權狀況": "",
        "所營事業": "",
        "_detail_url": "",
        "_print_url": "",
        "_share_url": "",
        "_detail_html": "",
        "_snapshot_at": "",
        "_error": "",
    }


def to_findbiz_snapshot_bytes(result: dict) -> bytes:
    """將當次查詢抓到的 findbiz 詳細頁保存成可下載的 HTML 快照。"""
    html = result.get("_detail_html") or ""
    if not html:
        return b""

    head_extra = (
        '<meta charset="utf-8">'
        '<base href="https://findbiz.nat.gov.tw/">'
    )
    if "<head>" in html:
        html = html.replace("<head>", f"<head>{head_extra}", 1)

    meta_block = (
        "<!-- "
        f"snapshot_at={result.get('_snapshot_at', '')}; "
        f"source_url={result.get('_detail_url', '')}"
        " -->\n"
    )
    return (meta_block + html).encode("utf-8")


def _fill(result: dict, raw: dict) -> None:
    """將 raw html 欄位映射到標準欄位名稱。"""
    mapping = {
        "統一編號":          ["統一編號"],
        "公司名稱":          ["公司名稱"],
        "章程所訂外文公司名稱": ["章程所訂外文公司名稱"],
        "資本總額(元)":      ["資本總額(元)"],
        "實收資本額(元)":    ["實收資本額(元)"],
        "每股金額(元)":      ["每股金額(元)"],
        "已發行股份總數(股)": ["已發行股份總數(股)"],
        "代表人姓名":        ["代表人姓名"],
        "公司所在地":        ["公司所在地"],
        "登記機關":          ["登記機關"],
        "核准設立日期":      ["核准設立日期"],
        "最後核准變更日期":  ["最後核准變更日期"],
        "複數表決權特別股":  ["複數表決權特別股"],
        "對於特定事項具否決權特別股": ["對於特定事項具否決權特別股"],
        "特別股董監選任限制": [
            "特別股股東被選為董事、監察人之 禁止或限制或當選一定名額之權利",
            "特別股股東被選為董事、監察人之禁止或限制或當選一定名額之權利",
        ],
        "登記現況":   ["登記現況"],
        "股權狀況":   ["股權狀況"],
    }
    for dest, sources in mapping.items():
        for src in sources:
            if src in raw and raw[src]:
                # 公司所在地：去掉電子地圖和家數說明
                if dest == "公司所在地":
                    val = raw[src].split("\n")[0].strip()
                else:
                    val = raw[src].split("\n")[0].strip()
                result[dest] = val
                break


def _parse_directors(soup: BeautifulSoup) -> tuple[str, list]:
    directors_term = ""
    directors = []

    container = soup.find(id="tabShareHolderContent")
    if not container:
        return directors_term, directors

    for line in container.get_text("\n", strip=True).splitlines():
        line = line.strip()
        if line.startswith("最近一次登記當屆董監事任期"):
            directors_term = line
            break

    data_table = None
    for table in container.find_all("table"):
        headers = [cell.get_text(" ", strip=True) for cell in table.find_all("th")]
        if "職稱" in headers and "姓名" in headers:
            data_table = table
            break

    if not data_table:
        return directors_term, directors

    for tr in data_table.find_all("tr")[1:]:
        cells = [cell.get_text(" ", strip=True) for cell in tr.find_all(["td", "th"])]
        if len(cells) < 5:
            continue
        directors.append(
            {
                "序號": cells[0],
                "職稱": cells[1],
                "姓名": cells[2],
                "所代表法人": cells[3],
                "持有股份數(股)": cells[4],
            }
        )

    return directors_term, directors


def _clean_business(raw: str) -> str:
    """
    清理所營事業字串：
    - 去掉業別代碼行（如 A101020、CC01010、ZZ99999）
    - 去掉多餘空白/換行/tab
    - 每項一行，去重
    """
    if not raw:
        return ""
    items = []
    for line in re.split(r"[\r\n\t]+", raw):
        line = line.strip()
        if not line:
            continue
        # 跳過純業別代碼行（1-2大寫英文 + 4-6數字，可能有後綴空白）
        if re.match(r"^[A-Z]{0,2}\d{4,6}\s*$", line):
            continue
        if line not in items:
            items.append(line)
    return "\n".join(items)
