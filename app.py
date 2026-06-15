# -*- coding: utf-8 -*-
"""
HỆ THỐNG GẮN MÃ KẾ TOÁN TỰ ĐỘNG - AUTO-ACCOUNTING MAPPER 🚀
Streamlit Web App chạy local hỗ trợ chuẩn hóa danh mục và áp mã kế toán thông minh.
"""

import streamlit as st
import pandas as pd
import io
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

# Import các hàm nghiệp vụ tự thiết lập
from utils.engine import (
    normalize_text,
    normalize_partner_name,
    match_commodity_row,
    match_partner_row,
    analyze_bank_transaction,
    find_next_available_code
)
from utils.sample_generator import (
    get_sample_commodities,
    get_sample_partners,
    get_sample_purchase_ledger,
    get_sample_sales_ledger,
    get_sample_inventory_ledger,
    get_sample_bank_statement
)

# Cấu hình giao diện trang web Streamlit
st.set_page_config(
    page_title="Auto-Accounting Mapper",
    page_icon="🚀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Khởi tạo Session State chứa dữ liệu mặc định hoặc dữ liệu đã upload
if "df_dir_commodities" not in st.session_state:
    st.session_state["df_dir_commodities"] = get_sample_commodities()
if "df_dir_partners" not in st.session_state:
    st.session_state["df_dir_partners"] = get_sample_partners()
if "demo_activated" not in st.session_state:
    st.session_state["demo_activated"] = False

# CSS Tùy biến để tăng mỹ thuật giao diện
st.markdown("""
<style>
    .main-header {
        font-family: 'Space Grotesk', sans-serif;
        font-size: 2.2rem;
        font-weight: 700;
        color: #1e293b;
        margin-bottom: 0.5rem;
        text-align: center;
    }
    .sub-header {
        font-family: 'Inter', sans-serif;
        font-size: 1.1rem;
        color: #64748b;
        margin-bottom: 2rem;
        text-align: center;
    }
    .metric-card {
        background-color: #f8fafc;
        border: 1px solid #e2e8f0;
        padding: 1rem;
        border-radius: 0.5rem;
        text-align: center;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05);
    }
    .metric-value {
        font-size: 1.8rem;
        font-weight: 700;
        color: #2563eb;
    }
    .metric-label {
        font-size: 0.85rem;
        color: #64748b;
        text-transform: uppercase;
        font-weight: 600;
        margin-top: 0.2rem;
    }
</style>
""", unsafe_allow_html=True)

# --- TIÊU ĐỀ CHÍNH ---
st.markdown("<div class='main-header'>🚀 AUTO-ACCOUNTING MAPPER</div>", unsafe_allow_html=True)
st.markdown("<div class='sub-header'>Hệ thống tự động gắn mã hàng hóa, mã đối tác và phân tích sao kê thông minh dựa trên logic chuẩn hóa tiếng Việt & Fuzzy String Match</div>", unsafe_allow_html=True)

# --- SIDEBAR CẤU HÌNH ---
with st.sidebar:
    st.markdown("### 🛠️ CÀI ĐẶT NGƯỠNG TƯƠNG THÍCH")
    st.info("Hãy kéo điều chỉnh các ngưỡng điểm áp dụng cho thuật toán so khớp tự động.")
    
    threshold_auto = st.slider(
        "🟢 Ngưỡng Tự Động Gắn Mã (Score >=)", 
        min_value=50, max_value=100, value=85, step=5,
        help="Dữ liệu đạt điểm từ ngưỡng này trở lên sẽ được tự động gắn mã mà không cần duyệt thủ công."
    )
    
    threshold_check = st.slider(
        "🟡 Ngưỡng Duyệt Thủ Công (Từ x% đến dưới 🟢)", 
        min_value=40, max_value=100, value=70, step=5,
        help="Dữ liệu đạt điểm nằm trong khoảng này sẽ chuyển vào tab 'Cần kiểm tra thủ công'."
    )
    
    st.markdown("---")
    st.markdown("### 🔠 FILE MÃ TỰ ĐỘNG")
    prefix_hh = st.text_input("Tiền tố Mã Hàng Hóa", "HH")
    prefix_kh = st.text_input("Tiền tố Mã Khách Hàng", "KH")
    prefix_ncc = st.text_input("Tiền tố Mã Nhà Cung Cấp", "NCC")
    
    st.markdown("---")
    st.markdown("### 📊 DEMO & TRẢI NGHIỆM")
    
    if st.button("✨ Kích hoạt dữ liệu DEMO mẫu", use_container_width=True):
        st.session_state["demo_activated"] = True
        st.success("Đã kích hoạt dữ liệu khảo sát mẫu đầy đủ các phân hệ!")
        
    if st.session_state["demo_activated"]:
        if st.button("🗑️ Reset về Trống", use_container_width=True):
            st.session_state["demo_activated"] = False
            st.rerun()

    st.markdown("---")
    st.markdown("<div style='font-size: 0.8rem; color: #94a3b8; text-align: center;'>Bản quyền cá nhân © 2026<br>Kế toán kho và Công nợ thông minh</div>", unsafe_allow_html=True)


# --- LUỒNG CHỌN CHẾ ĐỘ XỬ LÝ (MENU CHÍNH) ---
modes = [
    "Chế độ 1: Gắn mã hàng hóa trên bảng kê mua/bán",
    "Chế độ 2: Gắn mã đối tượng trên bảng kê mua/bán",
    "Chế độ 3: Gắn mã đối tác trên sao kê ngân hàng",
    "Chế độ 4: Gắn mã Tổng hợp (Gắn chéo đa phân hệ)"
]
selected_mode = st.selectbox("🎯 HÃY CHỌN PHÂN HỆ XỬ LÝ DỮ LIỆU:", modes)

# Phục vụ hiển thị dữ liệu gốc & danh mục hiện có
st.write("")


# --- HÀM PHỤ TRỢ: PHÂN TÍCH FILE HOẶC CHỌN SHEET ---
def load_and_select_sheet(file_obj, key_suffix=""):
    """
    Hỗ trợ upload Excel hoặc CSV, tự nhận diện sheet và trả về DataFrame
    """
    if file_obj is None:
        return None, None
        
    file_name = file_obj.name
    if file_name.endswith('.csv'):
        # CSV chỉ có 1 bảng
        try:
            df = pd.read_csv(file_obj, encoding='utf-8')
        except UnicodeDecodeError:
            df = pd.read_csv(file_obj, encoding='latin1')
        return df, "Default"
    else:
        # Excel có thể có nhiều sheet
        try:
            xl = pd.ExcelFile(file_obj)
            sheet_names = xl.sheet_names
            if len(sheet_names) > 1:
                selected_sheet = st.selectbox(f"Chọn sheet dữ liệu ({key_suffix}):", sheet_names, key=f"sheet_sel_{key_suffix}")
            else:
                selected_sheet = sheet_names[0]
            df = xl.parse(selected_sheet)
            return df, selected_sheet
        except Exception as e:
            st.error(f"Lỗi đọc file Excel: {e}")
            return None, None

def render_column_mapping(df, expected_cols, key_suffix=""):
    """
    Vẽ bảng cấu hình Mapping cột từ tệp tin upload và cột kỳ vọng
    """
    if df is None:
        return {}
        
    st.markdown(f"**🔗 Khớp Cột Excel gốc sang Cột Chuẩn ({key_suffix}):**")
    cols = ["-- Không sử dụng --"] + list(df.columns)
    mapping = {}
    
    # Grid layout cho mapping cột gọn gàng
    grid_cols = st.columns(3)
    for i, (col_key, label) in enumerate(expected_cols.items()):
        grid_idx = i % 3
        # Tự động gợi ý cột dựa theo từ khóa gần đúng
        default_idx = 0
        normalized_labels = [normalize_text(c) for c in df.columns]
        keywords = [normalize_text(col_key), normalize_text(label)]
        
        # Một số so khớp thông minh:
        if col_key == "ten_hang_hoa":
            keywords += ["hang hoa", "dien giai", "ten hang", "ten vat tu", "ten vat lieu", "dich vu"]
        elif col_key == "don_vi_tinh":
            keywords += ["dvt", "don vi", "tinh"]
        elif col_key == "so_luong":
            keywords += ["qty", "sl", "so luong"]
        elif col_key == "don_gia":
            keywords += ["gia", "don gia", "price"]
        elif col_key == "thanh_tien":
            keywords += ["tc", "thanh tien", "tien", "amount"]
        elif col_key == "ten_doi_tuong":
            keywords += ["nguoi mua", "nguoi ban", "khach hang", "doi tac", "nha cung cap", "ten nguoi", "don vi"]
        elif col_key == "ma_so_thue":
            keywords += ["mst", "ma so thue", "tax"]
        elif col_key == "so_tai_khoan":
            keywords += ["stk", "tai khoan", "so tk"]
        elif col_key == "ngay_giao_dich" or col_key == "ngay_hoa_don" or col_key == "ngay_chung_tu":
            keywords += ["ngay", "date", "thoi gian"]
        elif col_key == "noi_dung_giao_dich":
            keywords += ["diengiai", "noi dung", "transaction", "noi dung giao dich"]
            
        for idx, col_norm in enumerate(normalized_labels):
            if any(kw in col_norm for kw in keywords if len(kw) > 1):
                default_idx = idx + 1
                break
                
        with grid_cols[grid_idx]:
            mapped_col = st.selectbox(
                f"Cột chuẩn: **{label}**",
                cols,
                index=default_idx,
                key=f"map_{col_key}_{key_suffix}"
            )
            if mapped_col != "-- Không sử dụng --":
                mapping[col_key] = mapped_col
                
    return mapping


# ==========================================
# PHƠI BÀY CÁC TAB CHI TIẾT DANH MỤC KHỞI TẠO
# ==========================================
col_d1, col_d2 = st.columns(2)
with col_d1:
    with st.expander("📁 Danh mục Hàng hóa hiện tại (Chuẩn)"):
        st.dataframe(st.session_state["df_dir_commodities"], use_container_width=True, hide_index=True)
with col_d2:
    with st.expander("👥 Danh mục Khách hàng & Nhà cung cấp hiện tại (Chuẩn)"):
        st.dataframe(st.session_state["df_dir_partners"], use_container_width=True, hide_index=True)


# =========================================================================
# CHẾ ĐỘ 1: GẮN MÃ HÀNG HÓA
# =========================================================================
if "Chế độ 1" in selected_mode:
    st.subheader("📦 Gắn mã hàng hóa tự động trên Bảng kê Mua vào / Bán ra")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("#### 1. Tải lên bảng kê kế toán (Excel hoặc CSV)")
        uploaded_file = st.file_uploader("Kéo thả file vào đây", type=["xlsx", "xls", "csv"], key="file_m1")
        
        df_target = None
        if uploaded_file is not None:
            df_target, sheet_name = load_and_select_sheet(uploaded_file, "Bảng kê")
        elif st.session_state["demo_activated"]:
            st.info("💡 Đang sử dụng **Dữ liệu Mua vào Demo** làm đầu vào mẫu.")
            df_target = get_sample_purchase_ledger()
            
        if df_target is not None:
            st.write(f"Đã đọc: **{df_target.shape[0]} hàng**, **{df_target.shape[1]} cột**.")
            st.dataframe(df_target.head(3), use_container_width=True)
            
    with col2:
        st.markdown("#### 2. Cấu hình Cột")
        expected_cols = {
            "ten_hang_hoa": "Tên hàng hóa / diễn giải vụ việc *",
            "don_vi_tinh": "Đơn vị tính",
            "so_luong": "Số lượng",
            "don_gia": "Đơn giá",
            "thanh_tien": "Thành tiền"
        }
        col_map = render_column_mapping(df_target, expected_cols, "bảng_kê_hàng_hóa")

    if df_target is not None:
        st.markdown("---")
        if "ten_hang_hoa" not in col_map:
            st.warning("⚠️ Vui lòng gán tối thiểu cột **Tên hàng hóa / diễn giải vụ việc** để bắt đầu xử lý.")
        else:
            if st.button("⚡ Chạy tự động liên kết hàng hóa", type="primary", use_container_width=True):
                # Khởi chạy thuật toán
                st.info("Đang tự động tính toán độ tương khớp và kiểm tra danh mục...")
                
                output_rows = []
                new_created_items = []
                existing_item_codes = list(st.session_state["df_dir_commodities"]["ma_hang_hoa"].unique())
                
                # Biến đếm tạo mã mới để tránh trùng
                counter_new = 1
                
                for idx, row in df_target.iterrows():
                    desc = row.get(col_map["ten_hang_hoa"], "")
                    uom = row.get(col_map.get("don_vi_tinh", ""), "")
                    
                    # Tiến hành giải thuật
                    matched_code, matched_name, score, reason, is_new = match_commodity_row(
                        row_desc=desc,
                        row_uom=uom,
                        list_commodities=st.session_state["df_dir_commodities"].to_dict('records'),
                        auto_threshold=threshold_auto,
                        check_threshold=threshold_check
                    )
                    
                    action_status = ""
                    # Phân loại trạng thái dựa trên ngưỡng
                    if score >= threshold_auto:
                        action_status = "TỰ ĐỘNG GẮN"
                        is_new_code = False
                    elif score >= threshold_check:
                        action_status = "DUYỆT THỦ CÔNG"
                        is_new_code = False
                    else:
                        action_status = "TẠO MÃ MỚI"
                        is_new_code = True
                        
                    # Tạo mã mới nếu không khớp danh mục dưới ngưỡng
                    if is_new_code:
                        # Trích xuất một số từ khóa ban đầu làm tên chuẩn
                        clean_name = str(desc).strip()
                        new_code, counter_new = find_next_available_code(prefix_hh, existing_item_codes, counter_new)
                        
                        # Ghi nhận vào danh mục tạm cho các dòng sau so khớp
                        new_item = {
                            "ma_hang_hoa": new_code,
                            "ten_hang_hoa_chuan": clean_name,
                            "nhom_hang": "Chờ phân nhóm",
                            "don_vi_tinh": uom if not pd.isna(uom) else "Cái",
                            "quy_cach": "Tự động tạo mới",
                            "tu_khoa_nhan_dien": "",
                            "ghi_chu": f"Tự động tạo dựa trên diễn giải gốc: '{desc}'"
                        }
                        st.session_state["df_dir_commodities"] = pd.concat([
                            st.session_state["df_dir_commodities"],
                            pd.DataFrame([new_item])
                        ], ignore_index=True)
                        existing_item_codes.append(new_code)
                        new_created_items.append(new_item)
                        
                        matched_code = new_code
                        matched_name = clean_name
                        reason = f"Điểm tương hợp quá thấp (<{threshold_check}%). Hệ thống tự sinh mã vật tư mới."
                        score = 100
                    
                    # Ghép vào kết quả đầu ra
                    res_row = row.copy()
                    res_row["Mã hàng hóa"] = matched_code
                    res_row["Tên hàng hóa chuẩn"] = matched_name
                    res_row["Tỷ lệ tương thích (%)"] = score
                    res_row["Mức độ tương thích"] = "Cao" if score >= threshold_auto else ("Cần kiểm tra" if score >= threshold_check else "Mã mới tinh")
                    res_row["Lý do gắn mã"] = reason
                    res_row["Trạng thái xử lý"] = action_status
                    res_row["Ghi chú kiểm tra/Sửa đổi"] = ""
                    
                    output_rows.append(res_row)
                    
                df_res = pd.DataFrame(output_rows)
                st.session_state["df_processed_m1"] = df_res
                st.session_state["new_items_m1"] = pd.DataFrame(new_created_items) if new_created_items else pd.DataFrame()
                st.success("✅ Đã hoàn thành quá trình so khớp tự động!")
                
        # --- BẢNG ĐIỀU KHIỂN & HIỂN THỊ KẾT QUẢ ---
        if "df_processed_m1" in st.session_state:
            df_p = st.session_state["df_processed_m1"]
            
            # Tính toán thống kê nhanh
            total_rows = len(df_p)
            auto_mapped = len(df_p[df_p["Trạng thái xử lý"] == "TỰ ĐỘNG GẮN"])
            manual_check = len(df_p[df_p["Trạng thái xử lý"] == "DUYỆT THỦ CÔNG"])
            new_created = len(df_p[df_p["Trạng thái xử lý"] == "TẠO MÃ MỚI"])
            
            st.markdown("### 📊 DASHBOARD KẾT QUẢ SƠ BỘ")
            db_col1, db_col2, db_col3, db_col4, db_col5 = st.columns(5)
            with db_col1:
                st.markdown(f"<div class='metric-card'><div class='metric-value'>{total_rows}</div><div class='metric-label'>Tổng số dòng</div></div>", unsafe_allow_html=True)
            with db_col2:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#16a34a;'>{auto_mapped}</div><div class='metric-label'>Tự động gắn mã</div></div>", unsafe_allow_html=True)
            with db_col3:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#ca8a04;'>{manual_check}</div><div class='metric-label'>Cần duyệt thủ công</div></div>", unsafe_allow_html=True)
            with db_col4:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#dc2626;'>{new_created}</div><div class='metric-label'>Sinh mã mới</div></div>", unsafe_allow_html=True)
            with db_col5:
                rate = (auto_mapped / total_rows * 100) if total_rows > 0 else 0
                st.markdown(f"<div class='metric-card'><div class='metric-value'>{rate:.1f}%</div><div class='metric-label'>Tỷ lệ tự động tốt</div></div>", unsafe_allow_html=True)
                
            # MÀN HÌNH KIỂM TRA THỦ CÔNG
            st.markdown("---")
            st.markdown("### 🖥️ MÀN HÌNH KIỂM TRA THỦ CÔNG & ĐIỀU CHỈNH KẾT QUẢ")
            st.info("💡 Bạn có thể sửa đổi trực tiếp cột **Mã hàng hóa**, **Tên hàng hóa chuẩn** hoặc nhập **Ghi chú kiểm tra/Sửa đổi** ngay tại bảng dữ liệu dưới đây:")
            
            # Sử dụng st.data_editor của Streamlit cho phép nhập liệu tuyệt vời
            edited_df = st.data_editor(
                df_p,
                column_config={
                    "Mã hàng hóa": st.column_config.SelectboxColumn(
                        "Mã hàng hóa (Có thể chỉnh sửa)",
                        options=list(st.session_state["df_dir_commodities"]["ma_hang_hoa"].unique()),
                        width="medium"
                    ),
                    "Tên hàng hóa chuẩn": st.column_config.TextColumn(
                        "Tên hàng hóa chuẩn"
                    ),
                    "Trạng thái xử lý": st.column_config.SelectboxColumn(
                        "Trạng thái xử lý",
                        options=["TỰ ĐỘNG GẮN", "DUYỆT THỦ CÔNG", "TẠO MÃ MỚI", "BỎ QUA"]
                    )
                },
                use_container_width=True,
                key="editor_m1"
            )
            
            # Nút Lưu lại thay đổi chỉnh sửa
            if st.button("💾 Xác nhận lưu danh mục & chỉnh sửa", use_container_width=True):
                st.session_state["df_processed_m1"] = edited_df
                st.toast("Đã ghi nhận các thay đổi thủ công của bạn!", icon="✅")
                
            # XUẤT EXCEL CHUYÊN NGHIỆP
            st.markdown("---")
            st.markdown("### 📥 XUẤT FILE BÁO CÁO")
            
            # Xuất Excel bằng openpyxl ra bộ nhớ byte
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                # Đưa toàn bộ cột xử lý
                edited_df.to_excel(writer, sheet_name="Dữ liệu gán mã hoàn chỉnh", index=False)
                st.session_state["df_dir_commodities"].to_excel(writer, sheet_name="Danh mục hàng hóa cập nhật", index=False)
                if not st.session_state["new_items_m1"].empty:
                    st.session_state["new_items_m1"].to_excel(writer, sheet_name="Mã hàng tạo mới", index=False)
                    
            st.download_button(
                label="📥 Tải xuống File Excel Kết Quả Nhiều Sheet",
                data=output.getvalue(),
                file_name="auto_mapped_hang_hoa.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================================================================
# CHẾ ĐỘ 2: GẮN MÃ ĐỐI TƯỢNG (KHÁCH HÀNG / NHÀ CUNG CẤP)
# =========================================================================
elif "Chế độ 2" in selected_mode:
    st.subheader("👥 Gắn mã Khách hàng / Nhà cung cấp tự động trên Bảng kê Mua vào / Bán ra")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("#### 1. Tải lên bảng kê kế toán (Excel hoặc CSV)")
        ledger_type = st.radio("Loại bảng kê đầu vào:", ["Mua vào (Gắn Mã Nhà Cung Cấp)", "Bán ra (Gắn Mã Khách Hàng)"], horizontal=True)
        uploaded_file = st.file_uploader("Kéo thả file vào đây", type=["xlsx", "xls", "csv"], key="file_m2")
        
        df_target = None
        if uploaded_file is not None:
            df_target, sheet_name = load_and_select_sheet(uploaded_file, "Bảng kê đối tác")
        elif st.session_state["demo_activated"]:
            if "Mua vào" in ledger_type:
                st.info("💡 Đang sử dụng **Bảng kê Mua vào Demo**.")
                df_target = get_sample_purchase_ledger()
            else:
                st.info("💡 Đang sử dụng **Bảng kê Bán ra Demo**.")
                df_target = get_sample_sales_ledger()
                
        if df_target is not None:
            st.write(f"Đã đọc: **{df_target.shape[0]} hàng**, **{df_target.shape[1]} cột**.")
            st.dataframe(df_target.head(3), use_container_width=True)
            
    with col2:
        st.markdown("#### 2. Cấu hình Cột")
        expected_cols = {
            "ten_doi_tuong": "Tên khách hàng / nhà cung cấp / đơn vị bán *",
            "ma_so_thue": "Mã số thuế đối tác",
            "so_tai_khoan": "Số tài khoản ngân hàng đối tác",
            "ten_hang_hoa": "Diễn giải hóa đơn (hàng hóa)"
        }
        col_map = render_column_mapping(df_target, expected_cols, "bảng_kê_đối_tác")

    if df_target is not None:
        st.markdown("---")
        if "ten_doi_tuong" not in col_map:
            st.warning("⚠️ Vui lòng gán tối thiểu cột **Tên khách hàng / nhà cung cấp / đơn vị bán** để bắt đầu xử lý.")
        else:
            if st.button("⚡ Chạy tự động liên kết đối tác", type="primary", use_container_width=True):
                st.info("Đang tự động xác minh MST, liên kết số tài khoản ngân hàng và so khớp tên...")
                
                output_rows = []
                new_created_partners = []
                existing_partner_codes = list(st.session_state["df_dir_partners"]["ma_doi_tuong"].unique())
                
                # Biến đếm tạo mã mới để tránh trùng dựa trên nhóm nhà cung cấp hoặc khách hàng
                counter_kh = 1
                counter_ncc = 1
                
                is_buyer = "Bán ra" in ledger_type
                
                for idx, row in df_target.iterrows():
                    name_val = row.get(col_map["ten_doi_tuong"], "")
                    mst_val = row.get(col_map.get("ma_so_thue", ""), "")
                    acc_val = row.get(col_map.get("so_tai_khoan", ""), "")
                    invoice_val = row.get(col_map.get("ten_hang_hoa", ""), "")
                    
                    matched_code, matched_name, matched_type, score, reason, is_new = match_partner_row(
                        row_name=name_val,
                        row_mst=mst_val,
                        row_acc=acc_val,
                        invoice_desc=invoice_val,
                        list_partners=st.session_state["df_dir_partners"].to_dict('records'),
                        is_buyer=is_buyer,
                        auto_threshold=threshold_auto,
                        check_threshold=threshold_check
                    )
                    
                    action_status = ""
                    if score >= threshold_auto:
                        action_status = "TỰ ĐỘNG GẮN"
                        is_new_code = False
                    elif score >= threshold_check:
                        action_status = "DUYỆT THỦ CÔNG"
                        is_new_code = False
                    else:
                        action_status = "TẠO MÃ MỚI"
                        is_new_code = True
                        
                    if is_new_code:
                        clean_name = str(name_val).strip()
                        # Xác định loại mã mới
                        if is_buyer:
                            prefix = prefix_kh
                            p_type_new = "Khách hàng"
                            new_code, counter_kh = find_next_available_code(prefix, existing_partner_codes, counter_kh)
                        else:
                            prefix = prefix_ncc
                            p_type_new = "Nhà cung cấp"
                            new_code, counter_ncc = find_next_available_code(prefix, existing_partner_codes, counter_ncc)
                            
                        new_partner = {
                            "ma_doi_tuong": new_code,
                            "ten_doi_tuong": clean_name,
                            "loai_doi_tuong": p_type_new,
                            "ma_so_thue": str(mst_val).strip() if not pd.isna(mst_val) else "",
                            "so_tai_khoan": str(acc_val).strip() if not pd.isna(acc_val) else "",
                            "ngan_hang": "",
                            "dia_chi": "",
                            "tu_khoa_nhan_dien": "",
                            "ghi_chu": f"Tự động tạo từ bảng kê: '{name_val}'"
                        }
                        
                        st.session_state["df_dir_partners"] = pd.concat([
                            st.session_state["df_dir_partners"],
                            pd.DataFrame([new_partner])
                        ], ignore_index=True)
                        existing_partner_codes.append(new_code)
                        new_created_partners.append(new_partner)
                        
                        matched_code = new_code
                        matched_name = clean_name
                        matched_type = p_type_new
                        reason = f"Điểm tương thích quá thấp (<{threshold_check}%). Hệ thống tự sinh mã đối tác mới."
                        score = 100
                        
                    res_row = row.copy()
                    res_row["Mã đối tượng"] = matched_code
                    res_row["Tên đối tượng chuẩn"] = matched_name
                    res_row["Loại đối tượng"] = matched_type
                    res_row["Tỷ lệ tương thích (%)"] = score
                    res_row["Mức độ tương thích"] = "Cao" if score >= threshold_auto else ("Cần kiểm tra" if score >= threshold_check else "Đối tác mới tinh")
                    res_row["Lý do gắn mã"] = reason
                    res_row["Trạng thái xử lý"] = action_status
                    res_row["Ghi chú kiểm tra/Sửa đổi"] = ""
                    
                    output_rows.append(res_row)
                    
                df_res = pd.DataFrame(output_rows)
                st.session_state["df_processed_m2"] = df_res
                st.session_state["new_partners_m2"] = pd.DataFrame(new_created_partners) if new_created_partners else pd.DataFrame()
                st.success("✅ Đã hoàn thành quá trình so khớp đối tác!")

        # --- BẢNG ĐIỀU KHIỂN & HIỂN THỊ KẾT QUẢ ---
        if "df_processed_m2" in st.session_state:
            df_p = st.session_state["df_processed_m2"]
            
            total_rows = len(df_p)
            auto_mapped = len(df_p[df_p["Trạng thái xử lý"] == "TỰ ĐỘNG GẮN"])
            manual_check = len(df_p[df_p["Trạng thái xử lý"] == "DUYỆT THỦ CÔNG"])
            new_created = len(df_p[df_p["Trạng thái xử lý"] == "TẠO MÃ MỚI"])
            
            st.markdown("### 📊 DASHBOARD KẾT QUẢ ĐỐI TÁC SƠ BỘ")
            db_col1, db_col2, db_col3, db_col4, db_col5 = st.columns(5)
            with db_col1:
                st.markdown(f"<div class='metric-card'><div class='metric-value'>{total_rows}</div><div class='metric-label'>Tổng dòng bản kê</div></div>", unsafe_allow_html=True)
            with db_col2:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#16a34a;'>{auto_mapped}</div><div class='metric-label'>Tự động gắn mã</div></div>", unsafe_allow_html=True)
            with db_col3:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#ca8a04;'>{manual_check}</div><div class='metric-label'>Cần duyệt thủ công</div></div>", unsafe_allow_html=True)
            with db_col4:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#dc2626;'>{new_created}</div><div class='metric-label'>Sinh đối tác mới</div></div>", unsafe_allow_html=True)
            with db_col5:
                rate = (auto_mapped / total_rows * 100) if total_rows > 0 else 0
                st.markdown(f"<div class='metric-card'><div class='metric-value'>{rate:.1f}%</div><div class='metric-label'>Khớp đối tác tốt</div></div>", unsafe_allow_html=True)
                
            # MÀN HÌNH KIỂM TRA THỦ CÔNG
            st.markdown("---")
            st.markdown("### 🖥️ MÀN HÌNH KIỂM TRA THỦ CÔNG & ĐIỀU CHỈNH KẾT QUẢ")
            st.info("💡 Điền hoặc chọn các mã đối tượng chuẩn hóa trực tiếp tại bảng:")
            
            edited_df = st.data_editor(
                df_p,
                column_config={
                    "Mã đối tượng": st.column_config.SelectboxColumn(
                        "Mã đối tượng (Có thể chọn sửa)",
                        options=list(st.session_state["df_dir_partners"]["ma_doi_tuong"].unique()),
                        width="medium"
                    ),
                    "Loại đối tượng": st.column_config.SelectboxColumn(
                        "Loại đối tượng",
                        options=["Khách hàng", "Nhà cung cấp"]
                    ),
                    "Trạng thái xử lý": st.column_config.SelectboxColumn(
                        "Trạng thái xử lý",
                        options=["TỰ ĐỘNG GẮN", "DUYỆT THỦ CÔNG", "TẠO MÃ MỚI", "BỎ QUA"]
                    )
                },
                use_container_width=True,
                key="editor_m2"
            )
            
            if st.button("💾 Xác nhận lưu danh mục đối tác & chỉnh sửa", use_container_width=True):
                st.session_state["df_processed_m2"] = edited_df
                st.toast("Đã ghi nhận các thay đổi đối tác thành công!", icon="✅")
                
            # XUẤT EXCEL CHUYÊN NGHIỆP
            st.markdown("---")
            st.markdown("### 📥 XUẤT FILE BÁO CÁO")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, sheet_name="Dữ liệu gán mã đối tác", index=False)
                st.session_state["df_dir_partners"].to_excel(writer, sheet_name="Danh mục đối tác cập nhật", index=False)
                if not st.session_state["new_partners_m2"].empty:
                    st.session_state["new_partners_m2"].to_excel(writer, sheet_name="Mã đối tác mới tạo", index=False)
                    
            st.download_button(
                label="📥 Tải xuống File Excel Đối tác",
                data=output.getvalue(),
                file_name="auto_mapped_doi_tac.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================================================================
# CHẾ ĐỘ 3: GẮN MÃ ĐỐI TÁC TRÊN SAO KÊ NGÂN HÀNG
# =========================================================================
elif "Chế độ 3" in selected_mode:
    st.subheader("🏦 Gắn mã Khách hàng / Nhà cung cấp trên Sao kê Ngân hàng")
    st.info("💡 Thiết lập tự động nhận dạng: thu tiền được gợi ý là Khách hàng, chi tiền được gợi ý là Nhà cung cấp. Đọc nội dung tự phân loạt phí dịch vụ, lương, thuế, nội bộ.")
    
    col1, col2 = st.columns([2, 1])
    
    with col1:
        st.markdown("#### 1. Tải lên sao kê ngân hàng (Excel hoặc CSV)")
        uploaded_file = st.file_uploader("Kéo thả file sao kê vào đây", type=["xlsx", "xls", "csv"], key="file_m3")
        
        df_target = None
        if uploaded_file is not None:
            df_target, sheet_name = load_and_select_sheet(uploaded_file, "Sao kê ngân hàng")
        elif st.session_state["demo_activated"]:
            st.info("💡 Đang sử dụng **Sao kê Ngân hàng Demo**.")
            df_target = get_sample_bank_statement()
            
        if df_target is not None:
            st.write(f"Đã đọc: **{df_target.shape[0]} giao dịch**, **{df_target.shape[1]} cột**.")
            st.dataframe(df_target.head(3), use_container_width=True)
            
    with col2:
        st.markdown("#### 2. Cấu hình Cột")
        expected_cols = {
            "ngay_giao_dich": "Ngày giao dịch *",
            "noi_dung_giao_dich": "Nội dung/Dự toán giao dịch *",
            "so_tien_thu": "Số tiền Thu (Có / Có phát sinh)",
            "so_tien_chi": "Số tiền Chi (Nợ / Nợ phát sinh)",
            "so_tai_khoan_doi_ung": "Số tài khoản đối ứng",
            "ten_doi_tac_sao_ke": "Tên người chuyển/người nhận"
        }
        col_map = render_column_mapping(df_target, expected_cols, "sao_kê_giao_dịch")

    if df_target is not None:
        st.markdown("---")
        if "noi_dung_giao_dich" not in col_map:
            st.warning("⚠️ Vui lòng gán tối thiểu cột **Nội dung/Dự toán giao dịch** để bắt đầu xử lý.")
        else:
            if st.button("⚡ Chạy thuật toán phân tích sao kê", type="primary", use_container_width=True):
                st.info("Hệ thống đang rà soát dòng tiền thu/chi, phân nhóm nghiệp vụ tự động và truy vết đối tác...")
                
                output_rows = []
                for idx, row in df_target.iterrows():
                    desc = row.get(col_map["noi_dung_giao_dich"], "")
                    thu = row.get(col_map.get("so_tien_thu", ""), 0)
                    chi = row.get(col_map.get("so_tien_chi", ""), 0)
                    taikhoan = row.get(col_map.get("so_tai_khoan_doi_ung", ""), "")
                    name_sao_ke = row.get(col_map.get("ten_doi_tac_sao_ke", ""), "")
                    
                    # Gọi giải thuật phân tích sao kê
                    analysis = analyze_bank_transaction(
                        desc=desc,
                        receive_amount=thu,
                        pay_amount=chi,
                        counterpart_acc=taikhoan,
                        counterpart_name=name_sao_ke,
                        list_partners=st.session_state["df_dir_partners"].to_dict('records'),
                        auto_threshold=threshold_auto,
                        check_threshold=threshold_check
                    )
                    
                    res_row = row.copy()
                    res_row["Nhóm giao dịch dự đoán"] = analysis["nhom_giao_dich"]
                    res_row["Mã đối tượng đề xuất"] = analysis["ma_doi_tuong_de_xuat"]
                    res_row["Tên đối tác chuẩn"] = analysis["ten_doi_tuong_goc"]
                    res_row["Điểm tin cậy (%)"] = analysis["diem_tin_cay"]
                    res_row["Nội dung phân tích của máy"] = analysis["ly_do_du_doan"]
                    res_row["Trạng thái"] = "Cần kiểm tra" if analysis["can_kiem_tra"] else "Đã chốt"
                    res_row["Ghi chú kế toán nhập thêm"] = ""
                    
                    output_rows.append(res_row)
                    
                df_res = pd.DataFrame(output_rows)
                st.session_state["df_processed_m3"] = df_res
                st.success("✅ Đã hoàn thành quá trình phân tách và tự động hóa hồ sơ ngân hàng!")

        # --- BẢNG ĐIỀU KHIỂN & HIỂN THỊ KẾT QUẢ ---
        if "df_processed_m3" in st.session_state:
            df_p = st.session_state["df_processed_m3"]
            
            total_rows = len(df_p)
            checked_done = len(df_p[df_p["Trạng thái"] == "Đã chốt"])
            needs_review = len(df_p[df_p["Trạng thái"] == "Cần kiểm tra"])
            
            st.markdown("### 📊 DASHBOARD SAO KÊ NGÂN HÀNG")
            db_col1, db_col2, db_col3, db_col4 = st.columns(4)
            with db_col1:
                st.markdown(f"<div class='metric-card'><div class='metric-value'>{total_rows}</div><div class='metric-label'>Giao dịch sao kê</div></div>", unsafe_allow_html=True)
            with db_col2:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#16a34a;'>{checked_done}</div><div class='metric-label'>Giao dịch tự động khớp</div></div>", unsafe_allow_html=True)
            with db_col3:
                st.markdown(f"<div class='metric-card'><div class='metric-value' style='color:#ca8a04;'>{needs_review}</div><div class='metric-label'>Giao dịch hoài nghi cần duyệt</div></div>", unsafe_allow_html=True)
            with db_col4:
                unmapped_ratio = (needs_review / total_rows * 100) if total_rows > 0 else 0
                st.markdown(f"<div class='metric-card'><div class='metric-value'>{unmapped_ratio:.1f}%</div><div class='metric-label'>Tỷ lệ nghi vấn</div></div>", unsafe_allow_html=True)
                
            # MÀN HÌNH KIỂM TRA THỦ CÔNG
            st.markdown("---")
            st.markdown("### 🖥️ MÀN HÌNH KIỂM TRA THỦ CÔNG & ĐIỀU CHỈNH KẾT QUẢ")
            st.info("💡 Bạn có thể thay đổi dự đoán của hệ thống trực tiếp tại cột **Mã đối tượng đề xuất** và **Nhóm giao dịch dự đoán**:")
            
            edited_df = st.data_editor(
                df_p,
                column_config={
                    "Mã đối tượng đề xuất": st.column_config.SelectboxColumn(
                        "Mã đối tượng (Chuẩn)",
                        options=list(st.session_state["df_dir_partners"]["ma_doi_tuong"].unique()) + ["NHANVIEN", "NGANHANG", "KHO_BAC", "NOI_BO"],
                        width="medium"
                    ),
                    "Nhóm giao dịch dự đoán": st.column_config.SelectboxColumn(
                        "Nghiệp vụ hạch toán",
                        options=[
                            "Thu tiền khách hàng",
                            "Chi thanh toán nhà cung cấp",
                            "Phí ngân hàng / Tiền lãi",
                            "Chi lương / Tạm ứng",
                            "Nộp thuế / Ngân sách",
                            "Giao dịch nội bộ / Quỹ",
                            "Chi phí chưa phân loại (Chưa rõ đối tác)",
                            "Thu tiền lập lờ (Chưa rõ nguồn)"
                        ]
                    ),
                    "Trạng thái": st.column_config.SelectboxColumn(
                        "Trạng thái cột",
                        options=["Đã chốt", "Cần kiểm tra", "Bỏ qua"]
                    )
                },
                use_container_width=True,
                key="editor_m3"
            )
            
            if st.button("💾 Xác nhận lưu kết quả hạch toán ngân hàng", use_container_width=True):
                st.session_state["df_processed_m3"] = edited_df
                st.toast("Đã ghi nhận thay đổi kiểm tra sao kê ngân hàng thành công!", icon="✅")
                
            # XUẤT EXCEL CHUYÊN NGHIỆP
            st.markdown("---")
            st.markdown("### 📥 XUẤT FILE BÁO CÁO")
            
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                edited_df.to_excel(writer, sheet_name="Sao kê đã gán mã", index=False)
                st.session_state["df_dir_partners"].to_excel(writer, sheet_name="Danh mục đối tác", index=False)
                
            st.download_button(
                label="📥 Tải xuống File Excel Sao Kê Ngân Hàng",
                data=output.getvalue(),
                file_name="auto_mapped_sao_ke_ngan_hang.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# =========================================================================
# CHẾ ĐỘ 4: GẮN MÃ TỔNG HỢP (XỬ LÝ ĐỒNG THÌ THỜI ĐA PHÂN HỆ)
# =========================================================================
elif "Chế độ 4" in selected_mode:
    st.subheader("🧩 Gắn mã Tổng hợp (Hóa đơn, Dự án Kho, Đối chiếu Ngân hàng)")
    st.markdown("Liên kết đồng thời Bảng kê mua vào, Bảng kê bán ra, Kho xuất nhập tồn và Sao kê ngân hàng thành một chu trình khép kín, tự đối sánh số dư và dòng tiền chéo.")
    
    # 4 tabs tương ứng với 4 file
    tab_buy, tab_sell, tab_inv, tab_bank = st.tabs([
        "📥 1. Bảng kê Mua vào", 
        "📤 2. Bảng kê Bán ra", 
        "📦 3. Kho Xuất Nhập Tồn", 
        "🏦 4. Sao kê Ngân hàng"
    ])
    
    with tab_buy:
        st.write("Cấu hình tệp tin Bảng kê mua vào")
        buy_file = st.file_uploader("Kéo thả file bảng kê mua vào", type=["xlsx", "xls", "csv"], key="buy_f")
        df_buy = None
        if buy_file is not None:
            df_buy, _ = load_and_select_sheet(buy_file, "Mua vào")
        elif st.session_state["demo_activated"]:
            df_buy = get_sample_purchase_ledger()
            st.info("💡 Sử dụng Bảng kê Mua vào Mẫu")
            
        if df_buy is not None:
            st.dataframe(df_buy.head(2), use_container_width=True)
            col_map_buy = render_column_mapping(df_buy, {
                "ten_hang_hoa": "Tên hàng hóa *",
                "don_vi_tinh": "ĐVT",
                "ten_doi_tuong": "Tên nhà cung cấp *",
                "ma_so_thue": "Mã số thuế",
                "tong_thanh_toan": "Tổng giá trị thanh toán"
            }, "MUA_VAO")
            
    with tab_sell:
        st.write("Cấu hình tệp tin Bảng kê bán ra")
        sell_file = st.file_uploader("Kéo thả file bảng kê bán ra", type=["xlsx", "xls", "csv"], key="sell_f")
        df_sell = None
        if sell_file is not None:
            df_sell, _ = load_and_select_sheet(sell_file, "Bán ra")
        elif st.session_state["demo_activated"]:
            df_sell = get_sample_sales_ledger()
            st.info("💡 Sử dụng Bảng kê Bán ra Mẫu")
            
        if df_sell is not None:
            st.dataframe(df_sell.head(2), use_container_width=True)
            col_map_sell = render_column_mapping(df_sell, {
                "ten_hang_hoa": "Tên hàng hóa *",
                "don_vi_tinh": "ĐVT",
                "ten_doi_tuong": "Tên khách hàng *",
                "ma_so_thue": "Mã số thuế",
                "tong_thanh_toan": "Tổng giá trị thanh toán"
            }, "BAN_RA")
            
    with tab_inv:
        st.write("Cấu hình tệp tin Kho xuất nhập tồn")
        inv_file = st.file_uploader("Kéo thả file thẻ kho/xuất nhập tồn", type=["xlsx", "xls", "csv"], key="inv_f")
        df_inv = None
        if inv_file is not None:
            df_inv, _ = load_and_select_sheet(inv_file, "Kho")
        elif st.session_state["demo_activated"]:
            df_inv = get_sample_inventory_ledger()
            st.info("💡 Sử dụng Kho Xuất Nhập Tồn Mẫu")
            
        if df_inv is not None:
            st.dataframe(df_inv.head(2), use_container_width=True)
            col_map_inv = render_column_mapping(df_inv, {
                "ten_hang_hoa": "Tên hàng hóa *",
                "don_vi_tinh": "ĐVT",
                "ten_doi_tuong": "Tên đối tượng xuất nhập (Giao dịch)",
                "so_luong_nhap": "SL Nhập",
                "so_luong_xuat": "SL Xuất"
            }, "KHO")
            
    with tab_bank:
        st.write("Cấu hình tệp tin Sao kê Ngân hàng")
        bank_file = st.file_uploader("Kéo thả file sao kê tài khoản ngân hàng", type=["xlsx", "xls", "csv"], key="bank_f")
        df_bank = None
        if bank_file is not None:
            df_bank, _ = load_and_select_sheet(bank_file, "Ngân hàng")
        elif st.session_state["demo_activated"]:
            df_bank = get_sample_bank_statement()
            st.info("💡 Sử dụng Sao kê Ngân hàng Mẫu")
            
        if df_bank is not None:
            st.dataframe(df_bank.head(2), use_container_width=True)
            col_map_bank = render_column_mapping(df_bank, {
                "noi_dung_giao_dich": "Nội dung giao diên *",
                "so_tien_thu": "Thu (Có)",
                "so_tien_chi": "Chi (Nợ)",
                "so_tai_khoan_doi_ung": "STK đối ứng",
                "ten_doi_tac_sao_ke": "Tên người chuyển/nhận"
            }, "NGAN_HANG")

    st.markdown("---")
    
    # NÚT CHẠY XỬ LÝ LIÊN PHÂN HỆ
    st.subheader("⚡ CHẠY TỔNG HỢP VÀ ĐỐI CHIẾU")
    st.markdown("<p style='color:#dc2626; font-weight:600;'>Điều kiện: Bạn cần tải cấu hình ít nhất 2 phân hệ để có thể tạo báo cáo chéo hữu hiệu.</p>", unsafe_allow_html=True)
    
    # Nút bấm chạy tổng hợp
    ready_btn = False
    valid_dfs = sum(1 for d in [df_buy, df_sell, df_inv, df_bank] if d is not None)
    
    if valid_dfs >= 2:
        ready_btn = st.button("🔥 BẮT ĐẦU CHẠY PHÂN TÍCH TỔNG HỢP", type="primary", use_container_width=True)
    else:
        st.button("🔥 BẮT ĐẦU CHẠY PHÂN TÍCH TỔNG HỢP (Yêu cầu tải tối thiểu 2 file)", disabled=True, use_container_width=True)
        
    if ready_btn:
        st.toast("Hệ thống đang chạy quy trình so khớp chéo đa phân hệ...", icon="⏳")
        plots_dict = {}
        
        # 1. XỬ LÝ MUA VÀO (GẮN CẢ MÃ HÀNG LẪN MÃ ĐỐI TÁC)
        if df_buy is not None:
            output_buy = []
            for idx, row in df_buy.iterrows():
                # Lấy tên hàng và matching
                h_name = row.get(col_map_buy.get("ten_hang_hoa", ""), "")
                h_uom = row.get(col_map_buy.get("don_vi_tinh", ""), "")
                
                h_code, h_name_chuan, h_score, _, _ = match_commodity_row(
                    row_desc=h_name, row_uom=h_uom,
                    list_commodities=st.session_state["df_dir_commodities"].to_dict('records'),
                    auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                # Lấy nhà cung cấp và matching
                p_name = row.get(col_map_buy.get("ten_doi_tuong", ""), "")
                p_mst = row.get(col_map_buy.get("ma_so_thue", ""), "")
                
                p_code, p_name_chuan, _, p_score, _, _ = match_partner_row(
                    row_name=p_name, row_mst=p_mst, row_acc="", invoice_desc=h_name,
                    list_partners=st.session_state["df_dir_partners"].to_dict('records'),
                    is_buyer=False, auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                res_row = row.copy()
                res_row["Mã hàng hóa"] = h_code if h_score >= threshold_check else "MÃ THỦ CÔNG"
                res_row["Tên hàng chuẩn"] = h_name_chuan if h_score >= threshold_check else h_name
                res_row["Mã đối tượng Nhà cung cấp"] = p_code if p_score >= threshold_check else "MÃ THỦ CÔNG"
                res_row["Tên NCC chuẩn"] = p_name_chuan if p_score >= threshold_check else p_name
                res_row["Độ khớp hàng hóa"] = h_score
                res_row["Độ khớp NCC"] = p_score
                output_buy.append(res_row)
                
            plots_dict["df_buy_res"] = pd.DataFrame(output_buy)
            
        # 2. XỬ LÝ BÁN RA (GẮN CẢ MÃ HÀNG LẪN MÃ ĐỐI TÁC)
        if df_sell is not None:
            output_sell = []
            for idx, row in df_sell.iterrows():
                h_name = row.get(col_map_sell.get("ten_hang_hoa", ""), "")
                h_uom = row.get(col_map_sell.get("don_vi_tinh", ""), "")
                
                h_code, h_name_chuan, h_score, _, _ = match_commodity_row(
                    row_desc=h_name, row_uom=h_uom,
                    list_commodities=st.session_state["df_dir_commodities"].to_dict('records'),
                    auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                p_name = row.get(col_map_sell.get("ten_doi_tuong", ""), "")
                p_mst = row.get(col_map_sell.get("ma_so_thue", ""), "")
                
                p_code, p_name_chuan, _, p_score, _, _ = match_partner_row(
                    row_name=p_name, row_mst=p_mst, row_acc="", invoice_desc=h_name,
                    list_partners=st.session_state["df_dir_partners"].to_dict('records'),
                    is_buyer=True, auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                res_row = row.copy()
                res_row["Mã hàng hóa"] = h_code if h_score >= threshold_check else "MÃ THỦ CÔNG"
                res_row["Tên hàng chuẩn"] = h_name_chuan if h_score >= threshold_check else h_name
                res_row["Mã đối tượng Khách hàng"] = p_code if p_score >= threshold_check else "MÃ THỦ CÔNG"
                res_row["Tên KH chuẩn"] = p_name_chuan if p_score >= threshold_check else p_name
                res_row["Độ khớp hàng hóa"] = h_score
                res_row["Độ khớp KH"] = p_score
                output_sell.append(res_row)
                
            plots_dict["df_sell_res"] = pd.DataFrame(output_sell)

        # 3. XỬ LÝ KHO XUẤT NHẬP TỒN
        if df_inv is not None:
            output_inv = []
            for idx, row in df_inv.iterrows():
                h_name = row.get(col_map_inv.get("ten_hang_hoa", ""), "")
                h_uom = row.get(col_map_inv.get("don_vi_tinh", ""), "")
                
                h_code, h_name_chuan, h_score, _, _ = match_commodity_row(
                    row_desc=h_name, row_uom=h_uom,
                    list_commodities=st.session_state["df_dir_commodities"].to_dict('records'),
                    auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                p_name = row.get(col_map_inv.get("ten_doi_tuong", ""), "")
                p_code, p_name_chuan, p_type, p_score, _, _ = match_partner_row(
                    row_name=p_name, row_mst="", row_acc="", invoice_desc=h_name,
                    list_partners=st.session_state["df_dir_partners"].to_dict('records'),
                    is_buyer=True, # Dự phòng
                    auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                res_row = row.copy()
                res_row["Mã hàng hóa"] = h_code if h_score >= threshold_check else "MÃ THỦ CÔNG"
                res_row["Tên vật tư chuẩn"] = h_name_chuan if h_score >= threshold_check else h_name
                res_row["Mã đối tác đề đề xuất"] = p_code if p_score >= threshold_check else "MÃ KIỂM TRA"
                res_row["Đối tác chuẩn hóa"] = p_name_chuan if p_score >= threshold_check else p_name
                res_row["Loại đối tác"] = p_type if p_score >= threshold_check else "Chưa phân loại"
                output_inv.append(res_row)
                
            plots_dict["df_inv_res"] = pd.DataFrame(output_inv)

        # 4. XỬ LÝ SAO KE NGÂN HÀNG
        if df_bank is not None:
            output_bank = []
            for idx, row in df_bank.iterrows():
                desc = row.get(col_map_bank.get("noi_dung_giao_dich", ""), "")
                thu = row.get(col_map_bank.get("so_tien_thu", ""), 0)
                chi = row.get(col_map_bank.get("so_tien_chi", ""), 0)
                taikhoan = row.get(col_map_bank.get("so_tai_khoan_doi_ung", ""), "")
                name_sao_ke = row.get(col_map_bank.get("ten_doi_tac_sao_ke", ""), "")
                
                analysis = analyze_bank_transaction(
                    desc=desc, receive_amount=thu, pay_amount=chi,
                    counterpart_acc=taikhoan, counterpart_name=name_sao_ke,
                    list_partners=st.session_state["df_dir_partners"].to_dict('records'),
                    auto_threshold=threshold_auto, check_threshold=threshold_check
                )
                
                res_row = row.copy()
                res_row["Nhóm giao dịch dự toán"] = analysis["nhom_giao_dich"]
                res_row["Mã đối tượng gợi đề nghị"] = analysis["ma_doi_tuong_de_xuat"]
                res_row["Tên đối tác ngân hàng"] = analysis["ten_doi_tuong_goc"]
                res_row["Phân tích nghiệp vụ"] = analysis["ly_do_du_doan"]
                output_bank.append(res_row)
                
            plots_dict["df_bank_res"] = pd.DataFrame(output_bank)

        # 5. ĐỐI CHIẾU CHÉO (SỐ LIỆU ĐỐI CHIẾU RECONCILIATION)
        # 5a. Đối chiếu Hóa đơn - Ngân hàng
        df_recon_bank_inv = pd.DataFrame(columns=["Mã đối tượng", "Tên đối tượng", "Tổng tiền bán ra (Hóa đơn)", "Tổng tiền thu (Ngân hàng)", "Chênh lệch thu hồi công nợ"])
        if df_sell is not None and df_bank is not None:
            try:
                # Tính tổng tiền bán theo từng mã đối tượng KH
                df_sell_res = plots_dict["df_sell_res"]
                sell_grouped = df_sell_res.groupby("Mã đối tượng Khách hàng")["tong_thanh_toan"].sum().reset_index()
                
                # Tính tổng tiền thu ngân hàng theo từng mã đối tượng đề nghị
                df_bank_res = plots_dict["df_bank_res"]
                bank_grouped = df_bank_res[df_bank_res["Nhóm giao dịch dự toán"] == "Thu tiền khách hàng"].groupby("Mã đối tượng gợi đề nghị")["so_tien_thu"].sum().reset_index()
                
                # Merge đối chiếu
                recon = pd.merge(
                    sell_grouped, 
                    bank_grouped, 
                    left_on="Mã đối tượng Khách hàng", 
                    right_on="Mã đối tượng gợi đề nghị", 
                    how="outer"
                ).fillna(0)
                
                recon["Mã đối tượng"] = recon["Mã đối tượng Khách hàng"].where(recon["Mã đối tượng Khách hàng"] != 0, recon["Mã đối tượng gợi đề nghị"])
                
                # Lấy tên đối tác tương ứng
                partners_dict = dict(zip(st.session_state["df_dir_partners"]["ma_doi_tuong"], st.session_state["df_dir_partners"]["ten_doi_tuong"]))
                recon["Tên đối tượng"] = recon["Mã đối tượng"].map(partners_dict).fillna("Mã chưa rõ / tạo mới")
                
                recon["Tổng tiền bán ra (Hóa đơn)"] = recon["tong_thanh_toan"]
                recon["Tổng tiền thu (Ngân hàng)"] = recon["so_tien_thu"]
                recon["Chênh lệch thu hồi công nợ"] = recon["Tổng tiền bán ra (Hóa đơn)"] - recon["Tổng tiền thu (Ngân hàng)"]
                
                df_recon_bank_inv = recon[["Mã đối tượng", "Tên đối tượng", "Tổng tiền bán ra (Hóa đơn)", "Tổng tiền thu (Ngân hàng)", "Chênh lệch thu hồi công nợ"]]
            except Exception as e:
                st.warning(f"Không thể lập bảng đối chiếu hóa đơn - ngân hàng: {e}")
                
        plots_dict["df_recon_bank_inv"] = df_recon_bank_inv

        # Lưu vào session để xuất
        st.session_state["m4_all_results"] = plots_dict
        st.success("🎉 Đã hạch toán tổng hợp và tạo đối chiếu thành công!")

    # HIỂN THỊ KẾT QUẢ ĐỐI CHIẾU & SUMMARY
    if "m4_all_results" in st.session_state:
        res_m4 = st.session_state["m4_all_results"]
        
        st.markdown("### 📊 THU HOẠCH ĐỐI CHIẾU NỔI BẬT")
        
        st.write("**Bảng đối chiếu hóa đơn bán ra và luồng tiền thu thực tế từ ngân hàng:**")
        if not res_m4["df_recon_bank_inv"].empty:
            st.dataframe(res_m4["df_recon_bank_inv"].style.format({
                "Tổng tiền bán ra (Hóa đơn)": "{:,.0f} đ",
                "Tổng tiền thu (Ngân hàng)": "{:,.0f} đ",
                "Chênh lệch thu hồi công nợ": "{:,.0f} đ"
            }), use_container_width=True, hide_index=True)
        else:
            st.info("Chưa có đủ số liệu bán ra và ngân hàng để đối chiếu dòng hợp.")
            
        # Preview các sheet đã map
        exp1 = st.expander("👁️ Xem trước kết quả Thẻ Kho xuất nhập vật tư")
        if "df_inv_res" in res_m4:
            with exp1:
                st.dataframe(res_m4["df_inv_res"], use_container_width=True)
                
        # Tải tệp tin tổng hợp nhiều biểu
        st.markdown("---")
        st.markdown("### 📥 TẢI TRỌN BỘ FILE EXCEL ĐỐI CHIẾU LỚN")
        
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='openpyxl') as writer:
            if "df_buy_res" in res_m4:
                res_m4["df_buy_res"].to_excel(writer, sheet_name="7. Bảng kê Mua vào đã gắn mã", index=False)
            if "df_sell_res" in res_m4:
                res_m4["df_sell_res"].to_excel(writer, sheet_name="8. Bảng kê Bán ra đã gắn mã", index=False)
            if "df_inv_res" in res_m4:
                res_m4["df_inv_res"].to_excel(writer, sheet_name="9. Kho xuất tồn đã gắn mã", index=False)
            if "df_bank_res" in res_m4:
                res_m4["df_bank_res"].to_excel(writer, sheet_name="10. Sao kê Ngân hàng đã gắn mã", index=False)
            if "df_recon_bank_inv" in res_m4 and not res_m4["df_recon_bank_inv"].empty:
                res_m4["df_recon_bank_inv"].to_excel(writer, sheet_name="12. Đối chiếu Hóa đơn-Ngân hàng", index=False)
                
            st.session_state["df_dir_commodities"].to_excel(writer, sheet_name="Danh mục hàng hóa", index=False)
            st.session_state["df_dir_partners"].to_excel(writer, sheet_name="Danh mục đối tác", index=False)
            
        st.download_button(
            label="📥 Tải xuống TRỌN BỘ HỒ SƠ QUẢN LÝ TỔNG HỢP (Excel Nhiều Sheet)",
            data=output.getvalue(),
            file_name="ho_so_doi_chieu_ke_toan_tong_hop.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
