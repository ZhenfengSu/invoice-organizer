import tempfile
import unittest
from decimal import Decimal
from pathlib import Path
from unittest.mock import patch

import pymupdf

from processor import Invoice, _llm_classify, classify, detect_currency, extract, package, valid_uscc


class ProcessorTests(unittest.TestCase):
    def test_llm_does_not_send_without_private_configuration(self):
        with patch.dict("os.environ", {"LLM_API_KEY": "", "LLM_BASE_URL": ""}):
            with patch("urllib.request.urlopen") as send:
                self.assertEqual(_llm_classify("private invoice text"), ("", ""))
                send.assert_not_called()

    def test_uscc_check_digit(self):
        self.assertTrue(valid_uscc("91350000MA0000000F"))
        self.assertFalse(valid_uscc("91350000MA0000000X"))

    def test_ambiguous_dollar_is_not_assumed_usd(self):
        self.assertEqual(detect_currency("Total $10.00"), "未标明")
        self.assertEqual(detect_currency("Total €10.00"), "EUR")
        self.assertEqual(detect_currency("Total US$10.00"), "USD")

    def test_duplicate_numbers_are_all_excluded(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            paths = [root / f"invoice-{index}.pdf" for index in (1, 2)]
            for index, path in enumerate(paths):
                document = pymupdf.open()
                page = document.new_page()
                page.insert_text((72, 72), f"test {index}")
                document.save(path)
                document.close()
            records = [Invoice(path=path, source_name=path.name, region="国内", category="餐饮",
                               invoice_number="12345678901234567890", amount=Decimal("10.00"), currency="CNY") for path in paths]
            _archive, summary = package(records, root / "result")
            self.assertEqual(summary["errors"], 2)
            self.assertEqual(summary["domestic"], 0)
            self.assertEqual(summary["domestic_total"], "0.00")

    def test_buyer_tax_pair_mismatch_is_reported(self):
        text = """电子发票（普通发票） 发票号码：26000000000000000001
名称：示例甲科技有限公司
统一社会信用代码/纳税人识别号：91350000MA0000000F
*生产生活服务*餐饮费
价税合计（大写） 肆佰零肆圆整 （小写）¥404.00
"""
        with patch("processor.read_pdf", return_value=(text, text, "PDF 文字")):
            item = extract(Path("example.pdf"), "example.pdf",
                           ["示例甲科技有限公司", "示例乙科技有限公司"],
                           ["91350000MA0000001J", "91350000MA0000000F"], False)
        self.assertIn("购买方名称与税号组合不匹配", item.issues)

    def test_buyer_tax_id_on_next_line(self):
        full = """电子发票（普通发票） 发票号码：26000000000000000002
名称：示例甲科技有限公司        名称：示例乙服务有限公司
统一社会信用代码/纳税人识别号：    统一社会信用代码/纳税人识别号：
91350000MA0000000F    91350000MA0000001J
*交通运输服务*客运服务费
价税合计（大写）柒拾捌元玖角陆分（小写）¥78.96
"""
        left = """名称：示例甲科技有限公司
统一社会信用代码/纳税人识别号：
91350000MA0000000F
*交通运输服务*客运服务费
"""
        with patch("processor.read_pdf", return_value=(full, left, "PDF 文字")):
            item = extract(Path("tax-on-next-line.pdf"), "tax-on-next-line.pdf",
                           ["示例甲科技有限公司"], ["91350000MA0000000F"], False)
        self.assertEqual(item.category, "打车")
        self.assertEqual(item.tax_id, "91350000MA0000000F")
        self.assertEqual(item.issues, [])

    def test_travel_agency_refund_fee_classifies_as_airfare(self):
        region, category = classify("电子发票（普通发票）\n名称：示例旅行社有限公司\n*生产生活服务*退票费")
        self.assertEqual((region, category), ("国内", "机票"))

    def test_overseas_totals_stay_separate_by_currency(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            records = []
            for index, (currency, amount) in enumerate((("EUR", "10.00"), ("USD", "12.50"))):
                path = root / f"receipt-{index}.pdf"
                path.write_bytes(f"file {index}".encode())
                records.append(Invoice(path=path, source_name=path.name, region="海外",
                                       amount=Decimal(amount), currency=currency))
            _archive, summary = package(records, root / "result")
            self.assertEqual(summary["overseas_totals"], {"EUR": "10.00", "USD": "12.50"})

    def test_packaged_pdf_keeps_original_basename(self):
        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            path = root / "00007.pdf"
            path.write_bytes(b"test content")
            record = Invoice(path=path, source_name="上传目录/原始文件.pdf", region="海外",
                             amount=Decimal("1.00"), currency="EUR")
            package([record], root / "result")
            self.assertTrue((root / "result" / "海外收据" / "原始文件.pdf").exists())


if __name__ == "__main__":
    unittest.main()
