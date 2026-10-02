"""Single-file POs must use the uploaded report's reporting period."""
import datetime
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import pandas as pd
import main
from test_po_variants import source_row


class SingleSourceMonthsTests(unittest.TestCase):
    def test_one_source_uses_its_report_months(self):
        for label in ("ASIA", "GREEN"):
            with self.subTest(source=label), tempfile.TemporaryDirectory() as tmp:
                report = Path(tmp) / "report.xlsx"
                report.touch()
                source = pd.DataFrame([source_row("Product", barcode="000123", sales=14)])
                with (
                    patch.object(main, "parse_express_file", return_value=(source, {"months": 3})),
                    patch.object(main, "export_vendor_all_items_excel", return_value="all.xlsx"),
                    patch.object(main, "generate_po_from_combined", return_value="po.xlsx") as generate,
                ):
                    result = main.generate_po_streamlit(
                        express_asia_path=str(report) if label == "ASIA" else "",
                        express_green_path=str(report) if label == "GREEN" else "",
                        catalog_path="catalog.xlsx", vendor_info_path="vendors.xlsx",
                        template_path="template.xlsx", vendor_code="A0029",
                        po_date=datetime.date(2026, 10, 2), rate_thb_per_cny=6,
                        min_factor=4, max_factor=7,
                    )
                row = generate.call_args.kwargs["combined_df"].iloc[0]
                self.assertEqual(row["USE_MONTH"], 5)
                self.assertEqual(row["MIN_NUM"], 20)
                self.assertEqual(row["MAX_NUM"], 35)
                self.assertEqual(result["count_filtered"], 1)


if __name__ == "__main__":
    unittest.main()
