import streamlit as st
import datetime
import tempfile
from pathlib import Path
import main
from express_uploads import inspect_express_uploads, express_upload_fingerprint


st.set_page_config(page_title="PO generator", layout="centered")

st.title("PO generator")
st.write("---")

st.markdown("1. อัปโหลดไฟล์จาก EXPRESS (Greenlife และ AsiaHome)")
up_express_files = st.file_uploader(
    "อัปโหลดไฟล์ Express ASIA และ GREEN (.xlsx)",
    type=["xlsx"], accept_multiple_files=True, key="express",
    help="เลือกทั้ง 2 ไฟล์ได้ในช่องเดียว ระบบจะแยก ASIA และ GREEN จากชื่อบริษัทในหัวรายงาน",
)

upload_token = express_upload_fingerprint(up_express_files)
if st.session_state.get("express_upload_token") != upload_token:
    st.session_state["express_upload_token"] = upload_token
    st.session_state["single_express_confirmed"] = None

express_sources = {}
if up_express_files:
    try:
        express_sources = inspect_express_uploads(up_express_files)
    except ValueError as exc:
        st.error(str(exc))

for source in ("ASIA", "GREEN"):
    if source in express_sources:
        st.caption(f"ตรวจพบ {source}: {express_sources[source].name}")

single_file_confirmed = st.session_state.get("single_express_confirmed") == upload_token
if len(express_sources) == 1:
    source = next(iter(express_sources))
    st.warning(
        f"คุณอัปโหลดไฟล์ Express เพียง 1 ไฟล์ ({source}) "
        "กรุณาอัปโหลดทั้ง ASIA และ GREEN หากต้องการใช้เพียงไฟล์เดียว กรุณากดยืนยันก่อนสร้าง PO"
    )
    if not single_file_confirmed:
        if st.button("ยืนยันใช้ไฟล์ Express เพียง 1 ไฟล์", key="confirm_single_express"):
            st.session_state["single_express_confirmed"] = upload_token
            single_file_confirmed = True
    if single_file_confirmed:
        st.success(f"ยืนยันแล้ว: สร้าง PO โดยใช้ไฟล์ {source} เพียงไฟล์เดียว")

express_ready = len(express_sources) == 2 or (len(express_sources) == 1 and single_file_confirmed)

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

btn = st.button("Generate PO", key="generate_po", disabled=not express_ready)

if btn:
    vendor_code = vendor_code_in.strip().upper()

    if not vendor_code:
        st.error("กรุณาใส่รหัส Supplier")
        st.stop()

    if not express_ready:
        st.error("กรุณาอัปโหลด Express ทั้ง ASIA และ GREEN หรือยืนยันใช้เพียงไฟล์เดียว")
        st.stop()

    if up_catalog is None or up_vendorinfo is None:
        st.error("กรุณาอัปโหลดไฟล์รายละเอียดสินค้าและรายงานข้อมูลผู้จำหน่ายให้ครบ")
        st.stop()

    if max_factor < min_factor:
        st.error("MAX ต้องมากกว่าหรือเท่ากับ MIN")
        st.stop()

    template_repo_path = Path(__file__).resolve().with_name("ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx")
    if not template_repo_path.exists():
        st.error("ไม่พบไฟล์ template: ตัวอย่างใบสั่งซื้อต่างประเทศ.xlsx (วางไว้ข้างๆ app.py)")
        st.stop()

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)

        express_paths = {"ASIA": "", "GREEN": ""}
        for source, uploaded_file in express_sources.items():
            path = td / f"Express_{source}.xlsx"
            path.write_bytes(uploaded_file.getvalue())
            express_paths[source] = str(path)

        p_catalog = td / "catalog.xlsx"
        p_vendorinfo = td / "vendorinfo.xlsx"
        p_template = td / "template.xlsx"

        p_catalog.write_bytes(up_catalog.getvalue())
        p_vendorinfo.write_bytes(up_vendorinfo.getvalue())
        p_template.write_bytes(template_repo_path.read_bytes())

        result = main.generate_po_streamlit(
            express_asia_path=express_paths["ASIA"],
            express_green_path=express_paths["GREEN"],
            catalog_path=str(p_catalog),
            catalog_filename=up_catalog.name,
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
