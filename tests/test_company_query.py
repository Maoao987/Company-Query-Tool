import unittest
from unittest.mock import patch

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

        with patch("company_query.requests.get", return_value=_FakeResponse(html)):
            company_query.load_isin()

        entry = company_query._ISIN_BY_STOCK["00679B"]
        self.assertEqual(entry["security_type"], "ETF")
        self.assertEqual(entry["issue_country"], "台灣")
        self.assertEqual(entry["isin_code"], "TW00000679B0")

    def test_batch_infers_alphanumeric_etf_code_as_stock(self):
        df = pd.DataFrame({"query": ["00679B"]})

        requests_list, _ = company_query.extract_batch_requests(df)

        self.assertEqual(requests_list[0]["query_type"], "stock")
        self.assertEqual(requests_list[0]["query_value"], "00679B")


if __name__ == "__main__":
    unittest.main()
