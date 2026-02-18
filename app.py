import streamlit as st
import datetime
import tempfile
from pathlib import Path
from openpyxl import Workbook
import main

def make_empty_express_xlsx(path: Path):
    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws["A1"] = "BUYER"  # harmless header so parser finds nothing
    wb.save(path)


st.set_page_config(page_title="PO generator", layout="centered")

st.title("PO generator")
st.write("---")

st.markdown("1. อัปโหลดไฟล์จาก EXPRESS (Greenlife และ AsiaHome)")
up_express_asia = st.file_uploader("อัปโหลดไฟล์ Express ASIA (.xlsx)", type=["xlsx"], key="asia")
up_express_green = st.file_uploader("อัปโหลดไฟล์ Express GREEN (.xlsx)", type=["xlsx"], key="green")

st.markdown("2. อัปโหลดไฟล์ รายละเอียดสินค้า (ที่มีรูป)")
up_catalog = st.file_uploader("อัปโหลดไฟล์ข้อมูลสินค้า (.xlsx)", type=["xlsx"], key="catalog")

st.markdown("3. อัปโหลดไฟล์ รายงานข้อมูลผู้จำหน่าย (ที่อยู่ supplier)")
up_vendorinfo = st.file_uploader("อัปโหลดไฟล์รายงานข้อมูลผู้จำหน่าย (.xlsx)", type=["xlsx"], key="vendorinfo")

st.write("---")

vendor_code_in = st.text_input("รหัส Supplier (เช่น A0001)", value="")
po_date = st.date_input("วันที่", value=datetime.date.today())

min_factor = st.number_input("MIN", min_value=1, max_value=60, value=4, step=1)
max_factor = st.number_input("MAX", min_value=1, max_value=60, value=7, step=1)
rate = st.number_input("Exchange rate (THB/CNY)", min_value=0.01, value=6.0, step=0.1)

btn = st.button("Generate PO")

if btn:
    vendor_code = vendor_code_in.strip().upper()

    if not vendor_code:
        st.error("กรุณาใส่รหัส Supplier")
        st.stop()

    if (up_express_asia is None and up_express_green is None) or up_catalog is None or up_vendorinfo is None:
        st.error("กรุณาอัปโหลด Express อย่างน้อย 1 ไฟล์ (ASIA หรือ GREEN) และไฟล์ข้อ 2-3 ให้ครบ")
        st.stop()

    if max_factor < min_factor:
        st.error("MAX ต้องมากกว่าหรือเท่ากับ MIN")
        st.stop()

    template_repo_path = Path("ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx")
    if not template_repo_path.exists():
        st.error("ไม่พบไฟล์ template: ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx (วางไว้ข้างๆ app.py)")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        p_asia = td / "Express_A.xlsx"
        p_green = td / "Express_G.xlsx"

        if up_express_asia is not None:
            p_asia.write_bytes(up_express_asia.getvalue())

        if up_express_green is not None:
            p_green.write_bytes(up_express_green.getvalue())

        p_catalog = td / "catalog.xlsx"
        p_vendorinfo = td / "vendorinfo.xlsx"
        p_template = td / "template.xlsx"

        p_asia.write_bytes(up_express_asia.getvalue())
        p_green.write_bytes(up_express_green.getvalue())
        p_catalog.write_bytes(up_catalog.getvalue())
        p_vendorinfo.write_bytes(up_vendorinfo.getvalue())
        p_template.write_bytes(template_repo_path.read_bytes())

        result = main.generate_po_streamlit(
            express_asia_path=str(p_asia),
            express_green_path=str(p_green),
            catalog_path=str(p_catalog),
            vendor_info_path=str(p_vendorinfo),
            template_path=str(p_template),
            vendor_code=vendor_code,
            po_date=po_date,
            rate_thb_per_cny=float(rate),
            min_factor=int(min_factor),
            max_factor=int(max_factor),
        )

        st.success(
            f"เสร็จแล้ว ✅ Vendor {vendor_code} | "
            f"All items: {result['count_all']} | Below MIN: {result['count_filtered']}"
        )

        all_path = Path(result["po_all_items"])
        all_bytes = all_path.read_bytes()
        st.download_button(
            "Download ALL items (no MIN filter)",
            data=all_bytes,
            file_name=all_path.name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

        if result["po_filtered"]:
            po_path = Path(result["po_filtered"])
            po_bytes = po_path.read_bytes()
            st.download_button(
                "Download PO (only below MIN)",
                data=po_bytes,
                file_name=po_path.name,
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )
        else:
            st.info("Vendor นี้ไม่มีรายการที่ต่ำกว่า MIN → ไม่มีไฟล์ PO แบบ filtered")
