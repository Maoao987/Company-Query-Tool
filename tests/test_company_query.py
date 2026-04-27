import unittest
from io import BytesIO
from unittest.mock import patch

from openpyxl import load_workbook
import pandas as pd

import company_query


class _FakeResponse:
    def __init__(self, text):
        self.text = text
        self.encoding = "utf-8"


class CompanyQueryTests(unittest.TestCase):
    def setUp(self):
        company_query._ISIN_BY_STOCK.clear()
        company_query._ISIN_BY_NAME.clear()
        company_query._ISIN_BY_NORMALIZED_NAME.clear()

    def test_load_isin_keeps_etf_metadata(self):
        html = """
        <table class="h4">
          <tr><td>有價證券代號及名稱</td><td>國際證券辨識號碼(ISIN Code)</td><td>上市日</td><td>市場別</td><td>產業別</td><td>CFICode</td><td>備註</td></tr>
          <tr><td>ETF</td></tr>
          <tr><td>00679B　元大美債20年</td><td>TW00000679B0</td><td>2017/01/17</td><td>上櫃</td><td></td><td>CEOIBU</td><td></td></tr>
        </table>
        """

        def fake_get(url, **kwargs):
            if "strMode=4" in url:
                return _FakeResponse(html)
            return _FakeResponse("<html></html>")

        with patch("company_query.requests.get", side_effect=fake_get):
            company_query.load_isin()

        entry = company_query._ISIN_BY_STOCK["00679B"]
        self.assertEqual(entry["security_type"], "ETF")
        self.assertEqual(entry["issue_country"], "台灣")
        self.assertEqual(entry["isin_code"], "TW00000679B0")

        result = {}
        company_query._apply_security_metadata(result, entry)
        self.assertEqual(result["發行地查詢說明"], "ISIN 公開資料查詢")
        self.assertEqual(
            result["發行地查詢網址"],
            "https://isin.twse.com.tw/isin/C_public.jsp?strMode=4",
        )

    def test_batch_infers_alphanumeric_etf_code_as_stock(self):
        df = pd.DataFrame({"query": ["00679B"]})

        requests_list, _ = company_query.extract_batch_requests(df)

        self.assertEqual(requests_list[0]["query_type"], "stock")
        self.assertEqual(requests_list[0]["query_value"], "00679B")

    def test_excel_export_keeps_complete_source_links(self):
        result = {column: "" for column in company_query.RESULT_COLUMNS}
        result.update(
            {
                "公司名稱": "測試公司",
                "股票代號": "2330",
                "股價資料來源說明": "TWSE 個股日成交資訊（官方報表）",
                "股價資料來源網址": "https://example.test/price",
                "股價友善查詢說明": "TWSE 友善查詢頁",
                "股價友善查詢網址": "https://example.test/friendly",
                "公司登記資料說明": "查看 findbiz 官方頁面",
                "登記資料來源網址": "https://example.test/findbiz",
                "除權息資料來源說明": "查看 Yahoo 股利頁；查看 MOPS 查詢頁",
                "Yahoo股利頁網址": "https://example.test/yahoo",
                "MOPS查詢頁網址": "https://example.test/mops",
            }
        )

        workbook = load_workbook(BytesIO(company_query.to_excel_bytes([result])))
        sheet = workbook["查詢結果"]
        headers = [cell.value for cell in sheet[1]]
        row = {header: sheet.cell(row=2, column=index + 1) for index, header in enumerate(headers)}

        self.assertEqual(row["登記資料來源網址"].value, "查看 findbiz 官方頁面")
        self.assertEqual(row["股價資料來源網址"].value, "TWSE 個股日成交資訊（官方報表）")
        self.assertEqual(row["股價友善查詢網址"].value, "TWSE 友善查詢頁")
        self.assertEqual(row["Yahoo股利頁網址"].value, "查看 Yahoo 股利頁")
        self.assertEqual(row["MOPS查詢頁網址"].value, "查看 MOPS 查詢頁")
        self.assertEqual(row["Yahoo股利頁網址"].hyperlink.target, "https://example.test/yahoo")
        self.assertEqual(row["MOPS查詢頁網址"].hyperlink.target, "https://example.test/mops")


if __name__ == "__main__":
    unittest.main()
