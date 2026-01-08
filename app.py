import streamlit as st
import datetime
import tempfile
from pathlib import Path

import main  # your main.py

st.set_page_config(page_title="PO generator", layout="centered")

st.title("PO generator")
st.write("---")

# 1) Express uploads
st.markdown("1. อัปโหลดไฟล์จาก EXPRESS (Greenlife และ AsiaHome)")
up_express_asia = st.file_uploader("อัปโหลดไฟล์ Express ASIA (.xlsx)", type=["xlsx"], key="asia")
up_express_green = st.file_uploader("อัปโหลดไฟล์ Express GREEN (.xlsx)", type=["xlsx"], key="green")

# 2) Catalog
st.markdown("2. อัปโหลดไฟล์ รายละเอียดสินค้า (ที่มีรูป)")
up_catalog = st.file_uploader("อัปโหลดไฟล์ข้อมูลสินค้า (.xlsx)", type=["xlsx"], key="catalog")

# 3) Vendor info
st.markdown("3. อัปโหลดไฟล์ รายงานข้อมูลผู้จำหน่าย (ที่อยู่ supplier)")
up_vendorinfo = st.file_uploader("อัปโหลดไฟล์รายงานข้อมูลผู้จำหน่าย (.xlsx)", type=["xlsx"], key="vendorinfo")

st.write("---")

vendor_code = st.text_input("รหัส Supplier (เช่น A0001)", value="")
po_date = st.date_input("วันที่", value=datetime.date.today())

min_factor = st.number_input("MIN", min_value=1, max_value=60, value=4, step=1)
max_factor = st.number_input("MAX", min_value=1, max_value=60, value=7, step=1)

rate = st.number_input("Exchange rate (THB/CNY)", min_value=0.01, value=6.0, step=0.1)

st.write("")

btn = st.button("Generate PO")

if btn:
    # basic validation
    if not vendor_code.strip():
        st.error("กรุณาใส่รหัส Supplier")
        st.stop()

    if up_express_asia is None or up_express_green is None or up_catalog is None or up_vendorinfo is None:
        st.error("กรุณาอัปโหลดไฟล์ให้ครบทุกข้อ (1-3)")
        st.stop()

    if max_factor < min_factor:
        st.error("MAX ต้องมากกว่าหรือเท่ากับ MIN")
        st.stop()

    # save uploads to temp files (so openpyxl/pandas can read)
    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        p_asia = td / "Express_A.xlsx"
        p_green = td / "Express_G.xlsx"
        p_catalog = td / "catalog.xlsx"
        p_vendorinfo = td / "vendorinfo.xlsx"
        p_template = td / "template.xlsx"

        p_asia.write_bytes(up_express_asia.getvalue())
        p_green.write_bytes(up_express_green.getvalue())
        p_catalog.write_bytes(up_catalog.getvalue())
        p_vendorinfo.write_bytes(up_vendorinfo.getvalue())

        # if you keep template in repo, you can just point to it.
        # but if you want user upload template too, add another uploader.
        # here assumes template is in same folder as app.py:
        template_repo_path = Path("ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx")
        if not template_repo_path.exists():
            st.error("ไม่พบไฟล์ template: ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx (วางไว้ข้างๆ app.py)")
            st.stop()
        p_template.write_bytes(template_repo_path.read_bytes())

        # run your pipeline in main.py
        out_path = main.generate_po_streamlit(
            express_asia_path=str(p_asia),
            express_green_path=str(p_green),
            catalog_path=str(p_catalog),
            vendor_info_path=str(p_vendorinfo),
            template_path=str(p_template),
            vendor_code=vendor_code.strip(),
            po_date=po_date,
            rate_thb_per_cny=float(rate),
            min_factor=int(min_factor),
            max_factor=int(max_factor),
        )

        out_bytes = Path(out_path).read_bytes()
        st.success("สร้าง PO สำเร็จ ✅")

        st.download_button(
            "Download PO",
            data=out_bytes,
            file_name=f"PO_{vendor_code.strip().upper()}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )
