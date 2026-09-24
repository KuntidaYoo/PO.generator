"""Regression checks for distinct product variants in PO downloads."""

from __future__ import annotations

import datetime
import re
import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from unittest.mock import patch

import openpyxl
import pandas as pd
from PIL import Image as PILImage
from openpyxl.drawing.image import Image as XLImage

import main


ROOT = Path(__file__).resolve().parents[1]
TEMPLATE = ROOT / "ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx"


def source_row(
    description: str,
    *,
    barcode: str = "",
    sales: float = 0,
    stock: float = 0,
    on_order: float = 0,
    yuan: float = 0.447,
    code: str = "MC-510",
) -> dict:
    return {
        "buyer": "A0029",
        "barcode": barcode,
        "รหัสสินค้า": code,
        "รายละเอียดสินค้า": description,
        "ยอดขาย": sales,
        "สินค้าคงเหลือ": stock,
        "ON_ORDER": on_order,
        "หยวน": yuan,
    }


def combine(asia: list[dict], green: list[dict], months: int = 3) -> pd.DataFrame:
    return main.build_combined_all(
        pd.DataFrame(asia),
        pd.DataFrame(green),
        months=months,
        min_factor=4,
        max_factor=7,
    )


def parsed_sample_row(barcode: str, description: str) -> dict:
    """Run one screenshot-like Express line through the public parser."""
    line = f"A0029 {barcode} 01-22-1085-E1-1 MC-510 {description} 0.00 0.00 0.00 0.00 0.00 0.447"
    with (
        patch.object(main, "extract_5_6_block_from_money_chunk", return_value=[0.0] * 6),
        patch.object(main, "extract_yuan_after_money_block", return_value=0.447),
    ):
        fields = main.parse_line_to_fields(pd.Series(dtype=object), line)
    assert fields is not None
    code, clean_description = main.split_product_field(fields["สินค้า"])
    return source_row(clean_description, barcode=fields["barcode"], code=code)


def make_catalog(
    path: Path,
    carton_quantities: tuple[int | None, int | None],
    images: bool,
    descriptions: tuple[str, str] = (
        "unrelated catalog option one",
        "unrelated catalog option two",
    ),
) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "A0029"
    ws.append(["item", "picture", "description", "brand", "material", "weight", "qty/carton"])
    ws.append(["MC-510", None, descriptions[0], "Brand", "Steel", 1, carton_quantities[0]])
    ws.append(["MC-510", None, descriptions[1], "Brand", "Steel", 1, carton_quantities[1]])
    if images:
        for row, color in ((2, "red"), (3, "blue")):
            png = path.parent / f"catalog_{row}.png"
            PILImage.new("RGB", (12, 12), color).save(png)
            ws.add_image(XLImage(str(png)), f"B{row}")
    wb.save(path)


class POVariantTests(unittest.TestCase):
    def test_mc510_sample_colors_keep_three_rows(self) -> None:
        # These values are the three A0029 / MC-510 entries in the PO-G sample.
        green = [
            parsed_sample_row("8858778318375", "ก็อกแฟนซีสี-แดง(R)"),
            parsed_sample_row("8858778318498", "ก็อกแฟนซีสี-เขียว(G)"),
            parsed_sample_row("8858778318504", "ก็อกแฟนซีสี-น้ำเงิน(LB)"),
        ]

        rows = combine([], green)

        self.assertEqual(len(rows), 3)
        self.assertEqual(
            set(zip(rows["barcode"], rows["รายละเอียดสินค้า"])),
            {
                ("8858778318375", "ก็อกแฟนซีสี-แดง(R)"),
                ("8858778318498", "ก็อกแฟนซีสี-เขียว(G)"),
                ("8858778318504", "ก็อกแฟนซีสี-น้ำเงิน(LB)"),
            },
        )
        self.assertTrue((rows["ยอดขาย_TOTAL"] == 0).all())
        self.assertTrue((rows["MIN_NUM"] == 0).all())
        self.assertFalse((rows["TOTAL_QTY_NUM"] < rows["MIN_NUM"]).any())

        with tempfile.TemporaryDirectory() as tmp:
            path = main.export_vendor_all_items_excel(rows, "A0029", out_folder=tmp)
            ws = openpyxl.load_workbook(path)["all_items"]
            self.assertEqual(ws.max_row, 4)
            headers = {ws.cell(1, c).value: c for c in range(1, ws.max_column + 1)}
            exported = {
                (ws.cell(r, headers["barcode"]).value, ws.cell(r, headers["รายละเอียดสินค้า"]).value)
                for r in range(2, 5)
            }
            self.assertEqual(exported, set(zip(rows["barcode"], rows["รายละเอียดสินค้า"])))

    def test_exact_match_and_variant_calculations_without_quantity_fanout(self) -> None:
        red = "ก็อกแฟนซีสี-แดง(R)"
        blue = "ก็อกแฟนซีสี-น้ำเงิน(LB)"
        asia = [
            source_row(red, barcode="000123", sales=4, stock=1, on_order=1, yuan=10),
            source_row(red, barcode="", sales=3, stock=0, yuan=10),
        ]
        green = [
            source_row(red, barcode="000123", sales=2, stock=3, yuan=10),
            source_row(blue, barcode="000123", sales=12, stock=20, yuan=10),
            source_row(red, barcode="000123", sales=5, stock=0, yuan=11),
        ]

        rows = combine(asia, green)

        self.assertEqual(len(rows), 4)
        matching_red = rows[
            (rows["barcode"] == "000123")
            & (rows["รายละเอียดสินค้า"] == red)
            & (rows["หยวน"] == 10)
        ].iloc[0]
        self.assertEqual(matching_red["ยอดขาย_ASIA"], 4)
        self.assertEqual(matching_red["ยอดขาย_GREEN"], 2)
        self.assertEqual(matching_red["STOCK_ASIA"], 1)
        self.assertEqual(matching_red["STOCK_GREEN"], 3)
        self.assertEqual(matching_red["ON_ORDER_TOTAL"], 1)
        self.assertEqual(matching_red["USE_MONTH"], 2)
        self.assertEqual(matching_red["TOTAL_QTY_NUM"], 5)
        self.assertEqual(matching_red["MIN_NUM"], 8)
        self.assertEqual(matching_red["MAX_NUM"], 14)

        blue_row = rows[rows["รายละเอียดสินค้า"] == blue].iloc[0]
        self.assertEqual(blue_row["barcode"], "000123")
        self.assertEqual(blue_row["ยอดขาย_ASIA"], 0)
        self.assertEqual(blue_row["USE_MONTH"], 4)
        self.assertEqual(blue_row["MIN_NUM"], 16)
        self.assertEqual(blue_row["TOTAL_QTY_NUM"], 20)

        asia_only = rows[rows["barcode"] == ""].iloc[0]
        self.assertEqual(asia_only["ยอดขาย_ASIA"], 3)
        self.assertEqual(asia_only["ยอดขาย_GREEN"], 0)
        self.assertEqual(asia_only["STOCK_GREEN"], 0)

        different_price = rows[
            (rows["barcode"] == "000123")
            & (rows["รายละเอียดสินค้า"] == red)
            & (rows["หยวน"] == 11)
        ].iloc[0]
        self.assertEqual(different_price["ยอดขาย_ASIA"], 0)
        self.assertEqual(different_price["ยอดขาย_GREEN"], 5)

    def test_asia_only_row_has_no_po_barcode_and_parser_keeps_leading_zero(self) -> None:
        parsed_green = parsed_sample_row("000123", "ก็อกแฟนซีสี-แดง(R)")
        self.assertEqual(parsed_green["barcode"], "000123")

        rows = combine(
            [source_row("ASIA only", barcode="999888", code="MC-512", sales=6)],
            [],
        )
        self.assertEqual(len(rows), 1)
        self.assertEqual(rows.iloc[0]["barcode"], "")
        self.assertEqual(rows.iloc[0]["ยอดขาย_ASIA"], 6)
        self.assertEqual(rows.iloc[0]["ยอดขาย_GREEN"], 0)

    def test_only_identical_source_records_aggregate(self) -> None:
        green = [
            source_row("Red", barcode="0001", sales=2, stock=1, yuan=10),
            source_row("Red", barcode="0001", sales=3, stock=4, yuan=10),
            source_row("Red", barcode="0002", sales=5, stock=6, yuan=10),
            source_row("Red", barcode="0001", sales=7, stock=8, yuan=11),
        ]

        rows = combine([], green)

        self.assertEqual(len(rows), 3)
        identical = rows[(rows["barcode"] == "0001") & (rows["หยวน"] == 10)].iloc[0]
        self.assertEqual(identical["ยอดขาย_GREEN"], 5)
        self.assertEqual(identical["STOCK_GREEN"], 5)
        self.assertEqual(rows[rows["barcode"] == "0002"].iloc[0]["ยอดขาย_GREEN"], 5)
        self.assertEqual(rows[rows["หยวน"] == 11].iloc[0]["ยอดขาย_GREEN"], 7)

    def test_all_items_and_po_keep_leading_zero_barcode_as_text(self) -> None:
        description = "ก็อกแฟนซีสี-แดง(R)"
        rows = combine(
            [source_row(description, barcode="000123", sales=4, stock=1, on_order=1, yuan=10)],
            [source_row(description, barcode="000123", sales=2, stock=3, yuan=10)],
        )
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp)
            all_items_path = main.export_vendor_all_items_excel(rows, "A0029", out_folder=str(output))
            all_items = openpyxl.load_workbook(all_items_path)["all_items"]
            headers = {all_items.cell(1, col).value: col for col in range(1, all_items.max_column + 1)}
            all_barcode = all_items.cell(2, headers["barcode"])
            self.assertEqual(all_barcode.value, "000123")
            self.assertEqual(all_barcode.data_type, "s")

            with patch.object(main, "PO_OUTPUT_FOLDER", str(output)):
                po_path = main.generate_po_from_combined(
                    rows,
                    "A0029",
                    datetime.date(2026, 9, 24),
                    6,
                    str(TEMPLATE),
                    str(output / "missing_catalog.xlsx"),
                    str(output / "missing_vendors.xlsx"),
                    4,
                    7,
                )
            po = openpyxl.load_workbook(po_path)["PO"]
            self.assertEqual(po["X8"].value, "BARCODE")
            self.assertEqual(po["X9"].value, "000123")
            self.assertEqual(po["X9"].data_type, "s")
            self.assertEqual(po["C9"].value, description)
            self.assertEqual(po["R9"].value, 3)
            self.assertEqual(po["S9"].value, 1)
            self.assertEqual(po["T9"].value, 1)
            self.assertEqual(po["O9"].value, 2)
            self.assertIn("$X", str(po.print_area))

    def test_two_po_variants_keep_independent_prices_and_formula_rows(self) -> None:
        red = "ก็อกแฟนซีสี-แดง(R)"
        blue = "ก็อกแฟนซีสี-น้ำเงิน(LB)"
        rows = combine(
            [],
            [
                source_row(red, barcode="000123", sales=6, yuan=10.5),
                source_row(blue, barcode="000456", sales=9, yuan=12.25),
            ],
        )
        self.assertTrue((rows["TOTAL_QTY_NUM"] < rows["MIN_NUM"]).all())

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp)
            catalog_path = output / "catalog.xlsx"
            make_catalog(catalog_path, (10, 10), images=False, descriptions=(red, blue))
            with patch.object(main, "PO_OUTPUT_FOLDER", str(output)):
                po_path = main.generate_po_from_combined(
                    rows,
                    "A0029",
                    datetime.date(2026, 9, 24),
                    6,
                    str(TEMPLATE),
                    str(catalog_path),
                    str(output / "missing_vendors.xlsx"),
                    4,
                    7,
                )
            po = openpyxl.load_workbook(po_path)["PO"]
            by_barcode = {po[f"X{r}"].value: r for r in (9, 10)}
            self.assertEqual(set(by_barcode), {"000123", "000456"})

            for barcode, description, price in (
                ("000123", red, 10.5),
                ("000456", blue, 12.25),
            ):
                row = by_barcode[barcode]
                self.assertEqual(po[f"C{row}"].value, description)
                self.assertEqual(po[f"G{row}"].value, 10)
                self.assertEqual(po[f"L{row}"].value, price)
                self.assertAlmostEqual(po[f"M{row}"].value, price * 6)
                self.assertEqual(po[f"N{row}"].value, f"=L{row}*K{row}")
                self.assertEqual(po[f"K{row}"].value, f"=I{row}+J{row}")
                self.assertEqual(po[f"X{row}"].data_type, "s")

            total_formula = po["N14"].value
            total_range = re.fullmatch(r"=SUM\(N(\d+):N(\d+)\)", total_formula or "")
            self.assertIsNotNone(total_range)
            self.assertLessEqual(int(total_range.group(1)), 9)
            self.assertGreaterEqual(int(total_range.group(2)), 10)
            self.assertIn("$X", str(po.print_area))

    def test_exact_catalog_descriptions_select_matching_images(self) -> None:
        red = "ก็อกแฟนซีสี-แดง(R)"
        blue = "ก็อกแฟนซีสี-น้ำเงิน(LB)"
        rows = combine(
            [],
            [
                source_row(red, barcode="000123", sales=3),
                source_row(blue, barcode="000456", sales=3),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp)
            catalog_path = output / "catalog.xlsx"
            make_catalog(catalog_path, (10, 10), images=True, descriptions=(red, blue))
            with patch.object(main, "PO_OUTPUT_FOLDER", str(output)):
                po_path = main.generate_po_from_combined(
                    rows,
                    "A0029",
                    datetime.date(2026, 9, 24),
                    6,
                    str(TEMPLATE),
                    str(catalog_path),
                    str(output / "missing_vendors.xlsx"),
                    4,
                    7,
                )
            po = openpyxl.load_workbook(po_path)["PO"]
            selected = {
                image.anchor._from.row + 1: PILImage.open(BytesIO(image._data())).convert("RGB").getpixel((0, 0))
                for image in po._images
                if getattr(getattr(image.anchor, "_from", None), "col", None) == 1
                and getattr(getattr(image.anchor, "_from", None), "row", None) in (8, 9)
            }
            self.assertEqual(set(selected), {9, 10})
            for row in (9, 10):
                expected_color = (255, 0, 0) if po[f"C{row}"].value == red else (0, 0, 255)
                self.assertEqual(selected[row], expected_color)

    def test_ambiguous_catalog_image_is_blank(self) -> None:
        rows = combine(
            [],
            [source_row("ก็อกแฟนซีสี-แดง(R)", barcode="000123", sales=3)],
        )
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp)
            catalog_path = output / "catalog.xlsx"
            make_catalog(catalog_path, (10, 10), images=True)
            with patch.object(main, "PO_OUTPUT_FOLDER", str(output)):
                po_path = main.generate_po_from_combined(
                    rows,
                    "A0029",
                    datetime.date(2026, 9, 24),
                    6,
                    str(TEMPLATE),
                    str(catalog_path),
                    str(output / "missing_vendors.xlsx"),
                    4,
                    7,
                    variant_counts_by_code={"MC-510": 2},
                )
            po = openpyxl.load_workbook(po_path)["PO"]
            item_images = [
                image for image in po._images
                if getattr(getattr(image.anchor, "_from", None), "row", None) == 8
                and getattr(getattr(image.anchor, "_from", None), "col", None) == 1
            ]
            self.assertEqual(item_images, [])
            self.assertEqual(po["G9"].value, 10)

    def test_conflicting_catalog_carton_quantity_raises_clear_error(self) -> None:
        rows = combine(
            [],
            [source_row("ก็อกแฟนซีสี-แดง(R)", barcode="000123", sales=3)],
        )
        for quantities in ((10, 20), (None, 10)):
            with self.subTest(carton_quantities=quantities), tempfile.TemporaryDirectory() as tmp:
                output = Path(tmp)
                catalog_path = output / "catalog.xlsx"
                make_catalog(catalog_path, quantities, images=False)
                with patch.object(main, "PO_OUTPUT_FOLDER", str(output)):
                    with self.assertRaisesRegex((ValueError, RuntimeError), r"(?i)carton"):
                        main.generate_po_from_combined(
                            rows,
                            "A0029",
                            datetime.date(2026, 9, 24),
                            6,
                            str(TEMPLATE),
                            str(catalog_path),
                            str(output / "missing_vendors.xlsx"),
                            4,
                            7,
                            variant_counts_by_code={"MC-510": 2},
                        )


if __name__ == "__main__":
    unittest.main()
