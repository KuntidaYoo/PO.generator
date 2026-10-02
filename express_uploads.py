"""Identify Express uploads from their company heading, independently of names."""

from __future__ import annotations

import hashlib
import unicodedata
from io import BytesIO
from typing import Any, Iterable

import openpyxl


def _normalize_heading(value: str) -> str:
    text = unicodedata.normalize("NFKC", value).casefold()
    return "".join(
        character for character in text
        if not character.isspace()
        and not unicodedata.category(character).startswith("P")
        and unicodedata.category(character) != "Cf"
    )


_COMPANY_SIGNATURES = {
    "ASIA": tuple(_normalize_heading(name) for name in (
        "บริษัท เอเซีย โฮม สแตนดาร์ด จำกัด",
        "บริษัท เอเชีย โฮม สแตนดาร์ด จำกัด",
    )),
    "GREEN": (_normalize_heading("บริษัท กรีนไลฟ์ เอ็นเตอร์ไพรส์ จำกัด"),),
}
_COMPANY_LABELS = {
    "ASIA": "บริษัทเอเซีย โฮม สแตนดาร์ด",
    "GREEN": "บริษัทกรีนไลฟ์ เอ็นเตอร์ไพรส์",
}


def _file_bytes(data: Any) -> bytes:
    if isinstance(data, (bytes, bytearray, memoryview)):
        return bytes(data)
    if hasattr(data, "getvalue"):
        contents = data.getvalue()
    elif hasattr(data, "read"):
        position = data.tell() if hasattr(data, "tell") else None
        try:
            if hasattr(data, "seek"):
                data.seek(0)
            contents = data.read()
        finally:
            if position is not None and hasattr(data, "seek"):
                data.seek(position)
    else:
        raise ValueError("ไม่สามารถอ่านไฟล์ได้ กรุณาอัปโหลดไฟล์ Express (.xlsx) ใหม่")
    if not isinstance(contents, (bytes, bytearray, memoryview)):
        raise ValueError("ไม่สามารถอ่านไฟล์ได้ กรุณาอัปโหลดไฟล์ Express (.xlsx) ใหม่")
    return bytes(contents)


def detect_express_source(data: Any) -> str:
    """Return ASIA or GREEN from the first worksheet's bounded company header.

    Bytes and uploaded/file-like objects are accepted. File-like positions are
    preserved. Only the first five rows and forty columns are inspected; names
    and product descriptions do not decide which company supplied the report.
    """
    try:
        workbook = openpyxl.load_workbook(
            BytesIO(_file_bytes(data)), read_only=True, data_only=True,
        )
    except Exception as error:
        raise ValueError(
            "อ่านไฟล์ Excel ไม่ได้ กรุณาใช้ไฟล์ Express (.xlsx) ที่เปิดได้ตามปกติ"
        ) from error
    try:
        worksheet = workbook.worksheets[0]
        headings = []
        for row in worksheet.iter_rows(min_row=1, max_row=5, max_col=40, values_only=True):
            heading = _normalize_heading(" ".join(str(value) for value in row if value is not None))
            if heading.startswith("บริษัท"):
                headings.append(heading)
    except Exception as error:
        raise ValueError(
            "อ่านหัวรายงาน Excel ไม่ได้ กรุณาใช้ไฟล์ Express (.xlsx) ที่เปิดได้ตามปกติ"
        ) from error
    finally:
        workbook.close()

    def companies_in(heading: str) -> set[str]:
        return {
            source for source, signatures in _COMPANY_SIGNATURES.items()
            if any(signature in heading for signature in signatures)
        }

    if not headings or not companies_in(headings[0]):
        raise ValueError(
            "ไม่พบชื่อบริษัทเอเซีย โฮม สแตนดาร์ด หรือบริษัทกรีนไลฟ์ เอ็นเตอร์ไพรส์ "
            "ในหัวรายงาน 5 แถวแรก กรุณาตรวจสอบและเลือกไฟล์ Express ของบริษัทที่ถูกต้อง"
        )
    sources = set().union(*(companies_in(heading) for heading in headings))
    if len(sources) != 1:
        raise ValueError(
            "พบชื่อทั้งบริษัทเอเซียและบริษัทกรีนไลฟ์ในหัวไฟล์เดียวกัน "
            "กรุณาใช้ไฟล์ Express ที่มีรายงานของบริษัทเดียวต่อไฟล์"
        )
    return sources.pop()


def inspect_express_uploads(uploaded_files: Iterable[Any] | None) -> dict[str, Any]:
    """Bind one or two uploads to their company, rejecting unknown or duplicate sources."""
    uploads = list(uploaded_files or ())
    if not uploads:
        raise ValueError("กรุณาอัปโหลดไฟล์ Express อย่างน้อย 1 ไฟล์")
    if len(uploads) > 2:
        raise ValueError(
            "อัปโหลด Express ได้ไม่เกิน 2 ไฟล์: บริษัทเอเซีย 1 ไฟล์ "
            "และบริษัทกรีนไลฟ์ 1 ไฟล์"
        )
    sources = {}
    for upload in uploads:
        name = str(getattr(upload, "name", "Express.xlsx"))
        try:
            source = detect_express_source(upload)
        except ValueError as error:
            raise ValueError(f'ไฟล์ "{name}": {error}') from error
        if source in sources:
            previous_name = str(getattr(sources[source], "name", "Express.xlsx"))
            raise ValueError(
                f'พบไฟล์ของ{_COMPANY_LABELS[source]}ซ้ำ: "{previous_name}" และ "{name}" '
                "กรุณาเก็บเพียง 1 ไฟล์ของแต่ละบริษัท"
            )
        sources[source] = upload
    return sources


def express_upload_fingerprint(uploaded_files: Iterable[Any] | None) -> str:
    """Hash names and contents without depending on upload order.

    A changed name, workbook, or file count invalidates confirmation for a
    prior selection. Workbook contents and names are never exposed in the hash.
    """
    items = []
    for upload in uploaded_files or ():
        name = str(getattr(upload, "name", "Express.xlsx"))
        items.append((name, hashlib.sha256(_file_bytes(upload)).digest()))
    digest = hashlib.sha256()
    for name, contents_digest in sorted(items):
        encoded_name = name.encode("utf-8")
        digest.update(len(encoded_name).to_bytes(8, "big"))
        digest.update(encoded_name)
        digest.update(contents_digest)
    return digest.hexdigest()
