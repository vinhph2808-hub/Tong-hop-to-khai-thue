# -*- coding: utf-8 -*-
"""
web_app.py - Ứng dụng Web xử lý & tổng hợp số liệu tờ khai thuế (Streamlit).
Chạy trực tiếp trên trình duyệt web, hỗ trợ người dùng từ xa không cần cài đặt.
Tương thích 100% với Streamlit Cloud và mạng nội bộ.
"""

import io
import os
import shutil
import tempfile
from pathlib import Path

import pandas as pd
import streamlit as st

import tax_engine

# Cấu hình giao diện Streamlit
st.set_page_config(
    page_title="Tổng Hợp Số Liệu Tờ Khai Thuế",
    page_icon="📑",
    layout="wide",
    initial_sidebar_state="expanded",
)

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
    </style>
    """,
    unsafe_allow_html=True,
)

# Tiêu đề ứng dụng
st.markdown('<div class="main-header">📑 HỆ THỐNG XỬ LÝ & TỔNG HỢP SỐ LIỆU TỜ KHAI THUẾ</div>', unsafe_allow_html=True)
st.markdown(
    '<div class="sub-header">Tự động đọc PDF, XML, ZIP, RAR -> Kết chuyển chính xác vào các sheet E343 (GTGT), E350 (TNCN), E370 (NTNN), E382 (TNDN) theo đúng mẫu Excel chuẩn.</div>',
    unsafe_allow_html=True,
)

# Sidebar
with st.sidebar:
    st.header("⚙️ Tùy chọn & Hướng dẫn")
    st.info(
        "1. Kéo thả hoặc chọn các file **PDF, XML, ZIP, RAR** tờ khai vào khung bên phải.\n"
        "2. Chọn kỳ lọc nếu cần (Tất cả, Theo tháng, Theo quý).\n"
        "3. Bấm **Bắt đầu xử lý** để hệ thống tự động bóc tách số liệu.\n"
        "4. Xem trước bảng số liệu trực tiếp trên web và bấm **Tải file Excel**."
    )

    st.markdown("---")
    st.markdown("### 🔍 Bộ lọc kỳ tính thuế:")
    period_mode = st.radio(
        "Chọn chế độ lấy kỳ tờ khai:",
        options=["all", "month", "quarter"],
        format_func=lambda x: "Tất cả các kỳ (Khuyên dùng)" if x == "all" else ("Chỉ lấy tờ khai Tháng" if x == "month" else "Chỉ lấy tờ khai Quý"),
        index=0,
    )

    st.markdown("---")
    st.markdown("**Các loại tờ khai được hỗ trợ:**")
    st.markdown("- ✅ **01/GTGT:** Thuế Giá trị gia tăng (chỉ tiêu [22], [24], [25], [32a], [43]...)")
    st.markdown("- ✅ **05/KK-TNCN & 02/KK-TNCN:** Thuế Thu nhập cá nhân (chỉ tiêu [29]...)")
    st.markdown("- ✅ **03/TNDN:** Quyết toán thuế Thu nhập doanh nghiệp (chỉ tiêu A1, B1..B13, C3..C10, E)")
    st.markdown("- ✅ **01/NTNN:** Thuế Nhà thầu nước ngoài (chỉ tiêu 09, 13)")

# Khu vực Upload tập tin
uploaded_files = st.file_uploader(
    "📤 Chọn hoặc kéo thả các tệp tờ khai thuế (PDF, XML, ZIP, RAR):",
    type=["pdf", "xml", "zip", "rar"],
    accept_multiple_files=True,
    help="Có thể chọn nhiều tệp cùng một lúc. Hỗ trợ cả file nén ZIP, RAR chứa nhiều thư mục con.",
)

if uploaded_files:
    st.success(f"Đã nhận **{len(uploaded_files)}** tệp sẵn sàng xử lý.")

    col_btn, _ = st.columns([1, 2])
    with col_btn:
        start_process = st.button("🚀 Bắt đầu xử lý dữ liệu & lên số Excel", type="primary", use_container_width=True)

    if start_process:
        progress_bar = st.progress(0)
        status_text = st.empty()

        # Tạo thư mục làm việc tạm thời
        with tempfile.TemporaryDirectory() as tmp_dir:
            tmp_path = Path(tmp_dir)
            staged_files: list[Path] = []

            status_text.text("Đang chuẩn bị và giải nén các tập tin tải lên...")
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
                elif ext in {".pdf", ".xml"}:
                    staged_files.append(dest)

            progress_bar.progress(30)
            status_text.text(f"Đang bóc tách số liệu từ {len(staged_files)} tờ khai...")

            log_messages = []
            def web_logger(msg: str, msg_type: str = "info"):
                log_messages.append((msg_type, msg))

            # Chọn lọc tờ khai mới nhất
            latest_decls, parse_errors = tax_engine.select_latest_declarations(
                staged_files,
                logger=web_logger,
                period_mode=period_mode
            )

            progress_bar.progress(70)
            status_text.text("Đang kết chuyển số liệu vào mẫu Excel...")

            if not latest_decls:
                st.error("Không tìm thấy tờ khai hợp lệ nào trong các tệp đã tải lên.")
            else:
                template_file = Path("MẪU.xlsx")
                if not template_file.exists():
                    st.error("Không tìm thấy file mẫu 'MẪU.xlsx' trên hệ thống máy chủ.")
                else:
                    output_file = tmp_path / "KET_QUA_TO_KHAI_THUE.xlsx"
                    tax_engine.write_to_excel_openpyxl(
                        template_file,
                        output_file,
                        latest_decls,
                        parse_errors,
                        logger=web_logger
                    )

                    progress_bar.progress(100)
                    status_text.text("Xử lý hoàn tất!")
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
                    with open(output_file, "rb") as f:
                        excel_data = f.read()

                    st.markdown("---")
                    st.download_button(
                        label="📥 TẢI FILE EXCEL KẾT QUẢ (KẾT QUẢ TỜ KHAI THUẾ.xlsx)",
                        data=excel_data,
                        file_name="KET_QUA_TO_KHAI_THUE.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key="btn_download_excel",
                    )

                    # Bảng xem trước dữ liệu
                    st.markdown("### 📊 Xem trước bảng số liệu đã bóc tách:")
                    dfs = tax_engine.generate_summary_dataframes(latest_decls)

                    tab_vat, tab_pit, tab_cit, tab_fct, tab_log = st.tabs([
                        "📊 Thuế GTGT (E343)",
                        "👤 Thuế TNCN (E350)",
                        "🏢 Quyết toán TNDN (E382)",
                        "🌍 Thuế NTNN (E370)",
                        "📑 Nhật ký bóc tách",
                    ])

                    with tab_vat:
                        st.dataframe(dfs["VAT"].style.format(precision=0, thousands=","), use_container_width=True)

                    with tab_pit:
                        st.dataframe(dfs["PIT"].style.format(precision=0, thousands=","), use_container_width=True)

                    with tab_cit:
                        st.dataframe(dfs["CIT"].style.format(precision=0, thousands=","), use_container_width=True)

                    with tab_fct:
                        st.dataframe(dfs["FCT"].style.format(precision=0, thousands=","), use_container_width=True)

                    with tab_log:
                        st.markdown("**Danh sách các tờ khai đã xử lý:**")
                        for t, m in log_messages:
                            if t == "success":
                                st.success(m)
                            elif t == "warning":
                                st.warning(m)
                            else:
                                st.info(m)
