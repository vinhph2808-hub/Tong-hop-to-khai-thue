# -*- coding: utf-8 -*-
"""
web_app.py - Ứng dụng Web xử lý & tổng hợp số liệu tờ khai thuế (Streamlit).
Chạy trực tiếp trên trình duyệt web, hỗ trợ người dùng từ xa không cần cài đặt.
Tương thích 100% với Streamlit Cloud và mạng nội bộ.
Hỗ trợ ưu tiên: XML > PDF > Hình ảnh (PNG, JPG...) chuyển sang PDF/Scan.
"""

import datetime
import io
import os
import shutil
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import google_sheet_logger
import tax_engine

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Tổng Hợp Số Liệu Tờ Khai Thuế",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Khởi tạo trạng thái phiên làm việc (Session State)
if "uploader_key_id" not in st.session_state:
    st.session_state["uploader_key_id"] = 0

if "processed_result" not in st.session_state:
    st.session_state["processed_result"] = None

if "session_device_id" not in st.session_state:
    st.session_state["session_device_id"] = f"TK-WEB-{os.urandom(3).hex().upper()}"

# Lấy thông tin thiết bị và mạng của người dùng
client_info = google_sheet_logger.get_client_info_from_streamlit()

# Tự động ghi nhận máy tính đăng nhập vào Google Sheet (chỉ 1 lần mỗi phiên)
if "session_logged" not in st.session_state:
    st.session_state["session_logged"] = True
    google_sheet_logger.log_device_async(
        action="Đăng nhập hệ thống",
        platform_type="Web Online (Streamlit)",
        device_id=st.session_state["session_device_id"],
        client_name=client_info.get("os_name", "Thiết bị Web"),
        ip_address=client_info.get("ip", "127.0.0.1"),
        user_agent=f"{client_info.get('os_name', 'Web')} | {client_info.get('browser', 'Trình duyệt')}",
        note="Người dùng truy cập vào trang web",
    )

def clear_all_data():
    """Xóa sạch tệp đã tải lên và kết quả hiện tại để nạp đợt dữ liệu mới."""
    st.session_state["uploader_key_id"] += 1
    st.session_state["processed_result"] = None
    st.rerun()

def fmt_vnd(x):
    """Định dạng số theo chuẩn kế toán: dấu chấm '.' phân cách số, không có dấu phẩy ','."""
    if pd.isna(x) or x is None or x == "":
        return "0"
    try:
        val = float(x)
        if val == 0:
            return "0"
        if val.is_integer():
            return f"{int(val):,}".replace(",", ".")
        parts = f"{val:,.2f}".split(".")
        return parts[0].replace(",", ".") + "." + parts[1]
    except Exception:
        return str(x)

def display_styled_dataframe(df: pd.DataFrame):
    """Hiển thị bảng dữ liệu với định dạng số phân cách dấu chấm '.' chuyên nghiệp."""
    num_cols = [c for c in df.columns if c not in ("Tháng", "Mã chỉ tiêu", "Nội dung chỉ tiêu", "Chỉ tiêu", "Tên chỉ tiêu")]
    styler = df.style.format(formatter=fmt_vnd, subset=num_cols)
    st.dataframe(styler, use_container_width=True)

# Tùy biến CSS giao diện hiện đại, chuyên nghiệp
st.markdown(
    """
    <style>
    .main-header {
        font-size: 26px;
        font-weight: 700;
        color: #1E3A8A;
        margin-bottom: 4px;
    }
    .sub-header {
        font-size: 14px;
        color: #64748B;
        margin-bottom: 20px;
    }
    .metric-card {
        background-color: #FFFFFF;
        border: 1px solid #E2E8F0;
        border-radius: 8px;
        padding: 12px 16px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .stDownloadButton button {
        background-color: #16A34A !important;
        color: white !important;
        font-weight: bold !important;
        border-radius: 8px !important;
        padding: 12px 28px !important;
        font-size: 16px !important;
        width: 100%;
        border: none !important;
        box-shadow: 0 2px 4px rgba(22,163,74,0.3);
    }
    .stDownloadButton button:hover {
        background-color: #15803D !important;
    }
    .badge-xml {
        background-color: #DBEAFE;
        color: #1E40AF;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    .badge-pdf {
        background-color: #FEE2E2;
        color: #991B1B;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    .badge-img {
        background-color: #DCFCE7;
        color: #166534;
        padding: 2px 8px;
        border-radius: 4px;
        font-weight: bold;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# Tiêu đề ứng dụng
st.markdown('<div class="main-header">📑 HỆ THỐNG XỬ LÝ & TỔNG HỢP SỐ LIỆU TỜ KHAI THUẾ</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Tự động quét thông minh: <b>Ưu tiên XML eTax ➜ PDF ký số / văn bản ➜ Hình ảnh chuyển PDF / Scan</b> -> Kết chuyển chính xác vào các sheet E343 (GTGT), E350 (TNCN), E370 (NTNN), E382 (TNDN).</div>',
    unsafe_allow_html=True,
)

# Sidebar
with st.sidebar:
    st.header("⚙️ Tùy chọn & Tiện ích")
    
    # Nút Xóa dữ liệu nhanh trên thanh bên
    if st.button("🗑️ Xóa toàn bộ dữ liệu & Làm mới", use_container_width=True, type="secondary"):
        clear_all_data()

    st.markdown("---")
    st.markdown("### 🔍 Bộ lọc kỳ tính thuế:")
    period_mode = st.radio(
        "Chọn chế độ lấy kỳ tờ khai:",
        options=["all", "month", "quarter"],
        format_func=lambda x: "Tất cả các kỳ (Khuyên dùng)" if x == "all" else ("Chỉ lấy tờ khai Tháng" if x == "month" else "Chỉ lấy tờ khai Quý"),
        index=0,
    )

# Khu vực Upload tập tin có gắn key động để hỗ trợ Clear triệt để
current_uploader_key = f"uploader_{st.session_state['uploader_key_id']}"
uploaded_files = st.file_uploader(
    "📤 Chọn hoặc kéo thả các tệp tờ khai thuế (Ưu tiên XML, PDF, Hình ảnh JPG/PNG, Tệp nén ZIP/RAR):",
    type=["xml", "pdf", "zip", "rar", "png", "jpg", "jpeg", "bmp", "tif", "tiff", "webp"],
    accept_multiple_files=True,
    key=current_uploader_key,
    help="Hệ thống hỗ trợ cả XML, PDF, Hình ảnh (tự động chuyển sang PDF), và file nén ZIP/RAR chứa nhiều thư mục con.",
)

# Hiển thị thanh nút bấm thao tác
if uploaded_files or st.session_state["processed_result"] is not None:
    col_run, col_clear = st.columns([3, 2])
    with col_run:
        start_process = st.button(
            "🚀 Bắt đầu quét & bóc tách dữ liệu lên Excel",
            type="primary",
            use_container_width=True,
            disabled=(not uploaded_files),
        )
    with col_clear:
        if st.button("🗑️ Xóa dữ liệu hiện tại để thêm mới", use_container_width=True, type="secondary"):
            clear_all_data()

    if uploaded_files:
        st.success(f"Đã nhận **{len(uploaded_files)}** tệp sẵn sàng bóc tách.")
else:
    start_process = False

# Xử lý khi người dùng bấm Bắt đầu
if start_process and uploaded_files:
    # Ghi nhận hành động bóc tách tờ khai vào Google Sheet ngầm
    google_sheet_logger.log_device_async(
        action="Bóc tách tờ khai thuế",
        platform_type="Web Online (Streamlit)",
        device_id=st.session_state.get("session_device_id", "WEB-CLIENT"),
        client_name=client_info.get("os_name", "Thiết bị Web"),
        ip_address=client_info.get("ip", "127.0.0.1"),
        user_agent=f"{client_info.get('os_name', 'Web')} | {client_info.get('browser', 'Trình duyệt')}",
        file_count=len(uploaded_files),
        note=f"Đã tải lên {len(uploaded_files)} tệp để bóc tách",
    )

    progress_bar = st.progress(0)
    status_text = st.empty()

    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_path = Path(tmp_dir)
        staged_files: list[Path] = []

        status_text.text("Đang nhận dạng và chuẩn bị các tập tin tải lên...")
        for idx, uf in enumerate(uploaded_files):
            dest = tmp_path / uf.name
            with open(dest, "wb") as f:
                f.write(uf.getbuffer())

            ext = dest.suffix.lower()
            if ext in {".zip", ".rar"}:
                sub_extract = tmp_path / f"ext_{idx}"
                sub_extract.mkdir(parents=True, exist_ok=True)
                try:
                    extracted = tax_engine.extract_archive(dest, sub_extract)
                    staged_files.extend(extracted)
                except Exception as exc:
                    st.warning(f"Không thể giải nén {dest.name}: {exc}")
            elif ext in tax_engine.IMAGE_SUFFIXES:
                # Chuyển đổi file hình ảnh sang PDF và đưa vào danh sách xử lý
                converted_pdf = tmp_path / f"{dest.stem}_anh.pdf"
                try:
                    tax_engine.convert_image_to_pdf(dest, converted_pdf)
                    staged_files.append(converted_pdf)
                except Exception:
                    # Nếu lỗi chuyển đổi PDF thì giữ file ảnh gốc để quét OCR
                    staged_files.append(dest)
            elif ext in {".pdf", ".xml"}:
                staged_files.append(dest)

        progress_bar.progress(30)
        status_text.text(f"Đang phân tích và bóc tách dữ liệu từ {len(staged_files)} tài liệu...")

        log_messages = []
        def web_logger(msg: str, msg_type: str = "info"):
            log_messages.append((msg_type, msg))

        # Chọn lọc tờ khai với thứ tự ưu tiên: XML > PDF > Hình ảnh/Scan
        latest_decls, parse_errors, audit_records = tax_engine.select_latest_declarations(
            staged_files,
            logger=web_logger,
            period_mode=period_mode,
        )

        progress_bar.progress(70)
        status_text.text("Đang kết chuyển số liệu vào các sheet mẫu Excel...")

        if not latest_decls:
            st.error("Không tìm thấy tờ khai hợp lệ nào trong các tệp đã tải lên.")
            if parse_errors:
                with st.expander("🔍 Chi tiết nguyên nhân các tệp không đọc được:", expanded=True):
                    for name, err in parse_errors:
                        st.markdown(f"- ❌ **{name}**: {err}")
            st.session_state["processed_result"] = None
        else:
            template_file = Path("MẪU.xlsx")
            if not template_file.exists():
                st.error("Không tìm thấy file mẫu 'MẪU.xlsx' trên hệ thống máy chủ.")
                st.session_state["processed_result"] = None
            else:
                output_file = tmp_path / "KET_QUA_TO_KHAI_THUE.xlsx"
                tax_engine.write_to_excel_openpyxl(
                    template_file,
                    output_file,
                    latest_decls,
                    parse_errors,
                    logger=web_logger,
                )

                progress_bar.progress(100)
                status_text.text("Xử lý hoàn tất!")

                with open(output_file, "rb") as f:
                    excel_data = f.read()

                dfs = tax_engine.generate_summary_dataframes(latest_decls)

                st.session_state["processed_result"] = {
                    "excel_data": excel_data,
                    "latest_decls": latest_decls,
                    "dfs": dfs,
                    "log_messages": log_messages,
                    "audit_records": audit_records,
                    "parse_errors": parse_errors,
                }
                st.rerun()

# Hiển thị kết quả đã xử lý (nếu có trong session_state)
if st.session_state["processed_result"] is not None:
    res = st.session_state["processed_result"]
    latest_decls = res["latest_decls"]
    excel_data = res["excel_data"]
    dfs = res["dfs"]
    log_messages = res["log_messages"]
    audit_records = res.get("audit_records", [])
    parse_errors = res.get("parse_errors", [])

    st.success("🎉 Xử lý và lập bảng tổng hợp số liệu thuế thành công!")

    # Thống kê nhanh
    c1, c2, c3, c4 = st.columns(4)
    vat_count = sum(1 for d in latest_decls if d.tax_type == tax_engine.TAX_VAT)
    pit_count = sum(1 for d in latest_decls if d.tax_type == tax_engine.TAX_PIT)
    cit_count = sum(1 for d in latest_decls if d.tax_type == tax_engine.TAX_CIT)
    fct_count = sum(1 for d in latest_decls if d.tax_type == tax_engine.TAX_FCT)

    with c1:
        st.metric("Thuế GTGT (01/GTGT)", f"{vat_count} kỳ")
    with c2:
        st.metric("Thuế TNCN (05/KK)", f"{pit_count} kỳ")
    with c3:
        st.metric("Quyết toán TNDN", f"{cit_count} kỳ")
    with c4:
        st.metric("Thuế NTNN (01/NTNN)", f"{fct_count} kỳ")

    # Nút tải file Excel
    st.markdown("---")
    col_dl, col_reset = st.columns([3, 1])
    with col_dl:
        st.download_button(
            label="📥 TẢI FILE EXCEL KẾT QUẢ (KẾT QUẢ TỜ KHAI THUẾ.xlsx)",
            data=excel_data,
            file_name="KET_QUA_TO_KHAI_THUE.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            key="btn_download_excel",
        )
    with col_reset:
        if st.button("🔄 Thêm đợt dữ liệu mới", use_container_width=True):
            clear_all_data()

    # Bảng nhật ký quét chi tiết tất cả các file đã nạp
    if audit_records:
        with st.expander("📋 Nhật ký đối chiếu chi tiết tất cả tệp đã nạp (Ưu tiên XML > PDF > Ảnh)", expanded=True):
            df_audit = pd.DataFrame(audit_records)
            st.dataframe(df_audit, use_container_width=True, hide_index=True)

    # Bảng xem trước dữ liệu số liệu
    st.markdown("### 📊 Xem trước bảng số liệu đã bóc tách:")

    tab_vat, tab_pit, tab_cit, tab_fct, tab_log = st.tabs([
        "📊 Thuế GTGT (E343)",
        "👤 Thuế TNCN (E350)",
        "🏢 Quyết toán TNDN (E382)",
        "🌍 Thuế NTNN (E370)",
        "📑 Nhật ký bóc tách",
    ])

    with tab_vat:
        display_styled_dataframe(dfs["VAT"])

    with tab_pit:
        display_styled_dataframe(dfs["PIT"])

    with tab_cit:
        display_styled_dataframe(dfs["CIT"])

    with tab_fct:
        display_styled_dataframe(dfs["FCT"])

    with tab_log:
        st.markdown("**Danh sách các tờ khai đã xử lý:**")
        for t, m in log_messages:
            if t == "success":
                st.success(m)
            elif t == "warning":
                st.warning(m)
            else:
                st.info(m)
