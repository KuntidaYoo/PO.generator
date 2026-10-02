"""Company detection prevents swapped Express uploads and stale confirmations."""

from __future__ import annotations

import unittest
from io import BytesIO
from pathlib import Path

import openpyxl

from express_uploads import (
    detect_express_source,
    express_upload_fingerprint,
    inspect_express_uploads,
)


ASIA = "บริษัท เอเซีย โฮม สแตนดาร์ด จำกัด (สำนักงานใหญ่) หน้า : 1"
GREEN = "บริษัท กรีนไลฟ์ เอ็นเตอร์ไพรส์ จำกัด(สำนักงานใหญ่) หน้า : 1"
DOWNLOADS = Path.home() / "Downloads"


def workbook_bytes(rows: list[list[object]], *, second_heading: str | None = None) -> bytes:
    workbook = openpyxl.Workbook()
    worksheet = workbook.active
    for row in rows:
        worksheet.append(row)
    if second_heading is not None:
        workbook.create_sheet("Other company")["A1"] = second_heading
    buffer = BytesIO()
    workbook.save(buffer)
    workbook.close()
    return buffer.getvalue()


class Upload:
    def __init__(self, name: str, contents: bytes):
        self.name = name
        self.contents = contents

    def getvalue(self) -> bytes:
        return self.contents


class ReadFile:
    """File-like input without an UploadedFile/BytesIO getvalue method."""

    def __init__(self, contents: bytes):
        self.buffer = BytesIO(contents)

    def read(self) -> bytes:
        return self.buffer.read()

    def tell(self) -> int:
        return self.buffer.tell()

    def seek(self, position: int) -> int:
        return self.buffer.seek(position)


class ExpressCompanyDetectionTests(unittest.TestCase):
    def test_companies_identified_from_full_header(self):
        for heading, expected in ((ASIA, "ASIA"), (GREEN, "GREEN")):
            with self.subTest(expected=expected):
                self.assertEqual(detect_express_source(workbook_bytes([[heading]])), expected)

    def test_split_cells_nbsp_and_unicode_punctuation(self):
        contents = workbook_bytes([
            [None],
            ["\ufeffบริษัท\u00a0", "กรีนไลฟ์", "เอ็นเตอร์ไพรส์", "—จำกัด—", "หน้า : 1"],
        ])
        self.assertEqual(detect_express_source(contents), "GREEN")

    def test_common_asia_spelling(self):
        self.assertEqual(detect_express_source(workbook_bytes([[ASIA.replace("เอเซีย", "เอเชีย")]])), "ASIA")

    def test_file_position_is_preserved(self):
        stream = ReadFile(workbook_bytes([[ASIA]]))
        stream.seek(27)
        self.assertEqual(detect_express_source(stream), "ASIA")
        self.assertEqual(stream.tell(), 27)

    def test_random_green_product_does_not_identify_company(self):
        contents = workbook_bytes([
            ["บริษัท อื่น จำกัด"],
            ["A0001", "GREEN", GREEN],
        ])
        with self.assertRaisesRegex(ValueError, "หัวรายงาน"):
            detect_express_source(contents)

    def test_partial_company_name_is_not_sufficient(self):
        with self.assertRaisesRegex(ValueError, "หัวรายงาน"):
            detect_express_source(workbook_bytes([["บริษัท กรีนไลฟ์ จำกัด"]]))

    def test_unknown_first_company_does_not_use_later_heading(self):
        with self.assertRaisesRegex(ValueError, "หัวรายงาน"):
            detect_express_source(workbook_bytes([["บริษัท อื่น จำกัด"], [GREEN]]))

    def test_two_companies_in_header_are_ambiguous(self):
        for rows in ([[GREEN + " / " + ASIA]], [[GREEN], [ASIA]]):
            with self.subTest(rows=rows), self.assertRaisesRegex(ValueError, "บริษัทเดียว"):
                detect_express_source(workbook_bytes(rows))

    def test_second_worksheet_does_not_identify_first(self):
        with self.assertRaisesRegex(ValueError, "หัวรายงาน"):
            detect_express_source(workbook_bytes([["บริษัท อื่น จำกัด"]], second_heading=GREEN))

    def test_header_read_is_bounded_by_rows_and_columns(self):
        for rows in ([[None]] * 5 + [[GREEN]], [[None] * 40 + [GREEN]]):
            with self.subTest(rows=rows), self.assertRaisesRegex(ValueError, "หัวรายงาน"):
                detect_express_source(workbook_bytes(rows))

    def test_invalid_workbook_has_actionable_error(self):
        with self.assertRaisesRegex(ValueError, "Express.*xlsx"):
            detect_express_source(b"This is not an Excel workbook")

    @unittest.skipUnless((DOWNLOADS / "PO-G.xlsx").exists() and (DOWNLOADS / "po-a.xlsx").exists(),
                         "Original Express files are not available")
    def test_original_files(self):
        self.assertEqual(detect_express_source((DOWNLOADS / "PO-G.xlsx").read_bytes()), "GREEN")
        self.assertEqual(detect_express_source((DOWNLOADS / "po-a.xlsx").read_bytes()), "ASIA")


class ExpressUploadBindingTests(unittest.TestCase):
    def setUp(self):
        self.asia = Upload("misleading-PO-G.xlsx", workbook_bytes([[ASIA]]))
        self.green = Upload("misleading-po-a.xlsx", workbook_bytes([[GREEN]]))

    def test_any_order_and_misleading_names_bind_original_uploads(self):
        for uploads in ([self.asia, self.green], [self.green, self.asia]):
            with self.subTest(order=[upload.name for upload in uploads]):
                identified = inspect_express_uploads(uploads)
                self.assertIs(identified["ASIA"], self.asia)
                self.assertIs(identified["GREEN"], self.green)

    def test_one_file_reports_only_detected_company(self):
        self.assertEqual(inspect_express_uploads([self.green]), {"GREEN": self.green})

    def test_two_green_files_are_rejected(self):
        duplicate = Upload("other.xlsx", self.green.getvalue())
        with self.assertRaisesRegex(ValueError, "ซ้ำ.*เพียง 1 ไฟล์"):
            inspect_express_uploads([self.green, duplicate])

    def test_two_asia_files_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "เอเซีย.*ซ้ำ"):
            inspect_express_uploads([self.asia, Upload("other.xlsx", self.asia.getvalue())])

    def test_unknown_file_error_identifies_filename(self):
        unknown = Upload("check-this.xlsx", workbook_bytes([["บริษัท อื่น จำกัด"]]))
        with self.assertRaisesRegex(ValueError, "check-this.xlsx.*หัวรายงาน"):
            inspect_express_uploads([self.green, unknown])

    def test_no_files_and_more_than_two_files_are_rejected(self):
        with self.assertRaisesRegex(ValueError, "อย่างน้อย 1"):
            inspect_express_uploads([])
        with self.assertRaisesRegex(ValueError, "ไม่เกิน 2"):
            inspect_express_uploads([self.asia, self.green, self.green])

    def test_fingerprint_ignores_order_but_tracks_contents_names_and_count(self):
        original = express_upload_fingerprint([self.asia, self.green])
        self.assertEqual(original, express_upload_fingerprint([self.green, self.asia]))
        for changed in (
            [Upload("renamed.xlsx", self.asia.getvalue()), self.green],
            [Upload(self.asia.name, workbook_bytes([[ASIA], ["new report"]])), self.green],
            [self.green],
            [self.asia, self.green, self.green],
        ):
            with self.subTest(names=[upload.name for upload in changed]):
                self.assertNotEqual(original, express_upload_fingerprint(changed))

    def test_fingerprint_is_stable_for_repeat_reads(self):
        stream = ReadFile(self.green.getvalue())
        stream.name = self.green.name
        stream.seek(17)
        self.assertEqual(express_upload_fingerprint([stream]), express_upload_fingerprint([self.green]))
        self.assertEqual(stream.tell(), 17)


if __name__ == "__main__":
    unittest.main()
