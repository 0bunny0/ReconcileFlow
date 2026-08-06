from io import BytesIO
from pathlib import Path
import sys
import unittest

import pandas as pd
from openpyxl import load_workbook

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
from reconciliation import export_excel, reconcile_module_all  # noqa: E402


def book(sheets: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        for name, frame in sheets.items():
            frame.to_excel(writer, sheet_name=name, index=False)
    return output.getvalue()


class ReconciliationTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls) -> None:
        cls.document = book({
            "7月接单明细": pd.DataFrame([
                {"DAY": "2026-07-01", "BIL": "CB001", "JO": "BD001", "PART-DWG": "PART A", "CURR": "USD", "TP-CPO": 10, "VAT PRICE": 70},
                {"DAY": "2026-07-02", "BIL": "CB001", "JO": "-", "PART-DWG": "shipping cost", "CURR": "USD", "TP-CPO": 2, "VAT PRICE": 14},
                {"DAY": "2026-07-03", "BIL": "CB002", "JO": "BD002", "PART-DWG": "PART B", "CURR": "USD", "TP-CPO": 20, "VAT PRICE": 140},
            ]),
            "7月出货明细": pd.DataFrame([
                {"DAY": "2026-07-05", "BIL": "CB001", "JO": "SD001", "PART-DWG": "PART A", "CURR": "USD", "TP-CPO": 8, "VAT PRICE": 56},
                {"DAY": "2026-07-06", "BIL": "CB001", "JO": "-", "PART-DWG": "运费", "CURR": "USD", "TP-CPO": 1, "VAT PRICE": 7},
                {"DAY": "2026-07-07", "BIL": "CB003", "JO": "SD003", "PART-DWG": "PART C", "CURR": "USD", "TP-CPO": 30, "VAT PRICE": 210},
            ]),
        })
        cls.order_amount = book({"Sheet1": pd.DataFrame([
            {"接单时间": "2026-07-01", "客户代码": "CB001", "订单流水号": "BD001", "交易金额": 10, "接单金额(RMB)": 70},
            {"接单时间": "2026-07-03", "客户代码": "CB002", "订单流水号": "BD002", "交易金额": 21, "接单金额(RMB)": 147},
        ])})
        cls.order_freight = book({"Sheet1": pd.DataFrame([
            {"实际出厂日期": "2026-07-02", "客户代码": "CB001", "订单流水号": "BDFEE", "出货运费(原币)": 2, "出货运费(RMB)": 14},
        ])})
        cls.shipment_amount = book({"Shipping List": pd.DataFrame([
            {"出货日期": "2026-07-05", "客户代码": "CB001", "订单流水号": "SD001", "出货金额": 8, "出货金额(RMB)": 56, "交易币种": "USD"},
            {"出货日期": "2026-07-07", "客户代码": "CB003", "订单流水号": "SD003", "出货金额": 32, "出货金额(RMB)": 224, "交易币种": "USD"},
        ])})
        cls.shipment_freight = book({"Shipped Delivery Fee": pd.DataFrame([
            {"实际出厂日期": "2026-07-06", "客户代码": "CB001", "订单流水号": "SDFEE", "出货运费(原币)": 1, "出货运费(RMB)": 7},
        ])})

    def test_order_returns_both_bases_and_hides_zero_differences(self) -> None:
        results = reconcile_module_all(
            self.document, self.order_amount, self.order_freight,
            module="order", tolerance=0.01,
        )
        self.assertEqual(set(results), {"rmb", "original"})
        self.assertNotIn("CB001", set(results["rmb"].customer_differences["客户代码"]))
        self.assertIn("CB002", set(results["rmb"].customer_differences["客户代码"]))
        self.assertAlmostEqual(
            results["original"].customer_all.set_index("客户代码").at["CB002", "差异"],
            1,
        )

    def test_shipment_returns_both_bases_and_matches_freight(self) -> None:
        results = reconcile_module_all(
            self.document, self.shipment_amount, self.shipment_freight,
            module="shipment", tolerance=0.01,
        )
        self.assertEqual(results["rmb"].statistics["自动关联运费行"], 1)
        self.assertAlmostEqual(
            results["rmb"].customer_all.set_index("客户代码").at["CB003", "差异"],
            14,
        )
        self.assertAlmostEqual(
            results["original"].customer_all.set_index("客户代码").at["CB003", "差异"],
            2,
        )

    def test_excel_export_is_valid(self) -> None:
        result = reconcile_module_all(
            self.document, self.shipment_amount, self.shipment_freight,
            module="shipment", tolerance=0.01,
        )["rmb"]
        workbook = load_workbook(BytesIO(export_excel(result)), read_only=True)
        self.assertEqual(
            workbook.sheetnames,
            ["差异汇总", "流水号差异", "全部客户汇总", "系统拼接明细", "文控出货明细", "核对说明"],
        )


if __name__ == "__main__":
    unittest.main()
