# -*- coding: utf-8 -*-
"""
Core Matching and Analytical Engine for Accounting Data Mapping.
Supports string normalization, fuzzy matching, specs extraction,
and transaction categorization.
"""

import re
import pandas as pd
from unidecode import unidecode
from rapidfuzz import fuzz

def normalize_text(text):
    """
    Chuẩn hóa văn bản thường gặp trong hóa đơn và danh mục hàng hóa.
    - Chuyển thành chữ thường.
    - Loại bỏ dấu tiếng Việt (giữ nguyên bảng chữ cái Latin).
    - Giữ lại các ký tự phục vụ kích thước, mẫu mã: số, chữ, dấu chấm, dấu gạch ngang, gạch chéo, ký tự 'x'.
    """
    if pd.isna(text) or not isinstance(text, str):
        return ""
    # Chuyển về chữ thường, loại bỏ dấu tiếng Việt bằng unidecode
    normalized = unidecode(text).lower()
    # Loại bỏ ký tự đặc biệt trừ chữ, số, dấu chấm (.), gạch ngang (-), gạch chéo (/), và chữ 'x' (kích thước)
    normalized = re.sub(r'[^a-z0-9\s\.\-x/]', ' ', normalized)
    # Loại bỏ khoảng trắng thừa
    normalized = re.sub(r'\s+', ' ', normalized).strip()
    return normalized

def normalize_partner_name(name):
    """
    Chuẩn hóa tên đối tượng doanh nghiệp/cá nhân:
    - Loại bỏ dấu và chuẩn hóa text thông thường.
    - Sàng lọc và loại bỏ các định danh loại hình doanh nghiệp phổ biến để tăng chính xác khi so khớp tên riêng.
    """
    text = normalize_text(name)
    if not text:
        return ""
        
    # Danh sách các định danh pháp lý thông dụng (đã viết thường, không dấu)
    redundant_words = [
        r'\bcong ty tnhh mtv\b', r'\bcong ty tnhh\b', r'\bcong ty co phan\b', 
        r'\bcong ty cp\b', r'\bchi nhanh\b', r'\bdoanh nghiep tu nhan\b', 
        r'\bdntn\b', r'\bho kinh doanh\b', r'\bhkd\b', r'\bco ltd\b', 
        r'\bjsc\b', r'\bltd\b', r'\bcong ty\b', r'\bcty\b', r'\bmtv\b', 
        r'\bco phan\b', r'\bcp\b'
    ]
    
    for pattern in redundant_words:
        text = re.sub(pattern, '', text)
        
    text = re.sub(r'\s+', ' ', text).strip()
    return text

def extract_specs(text):
    """
    Trích xuất quy cách, model, dung tích, kích thước chính từ nội dung:
    Ví dụ: '500ml', '1kg', '20kg', 'pcb40', 'a4', 'phi 21', 'd90', '10x20', 'model a-200'
    """
    if not text:
        return []
    
    text_clean = text.lower()
    specs = []
    
    # 1. Trích xuất dung tích / trọng lượng: 500ml, 1.5l, 10kg, 25g...
    weight_volume_matches = re.findall(r'\b\d+(?:\.\d+)?\s*(?:ml|l|g|kg|ton|lit|kg|gam|chai|hop)\b', text_clean)
    if weight_volume_matches:
        specs.extend([w.replace(" ", "") for w in weight_volume_matches])
        
    # 2. Trích xuất kích thước dạng số x số: 10x20, 10 x 20 x 50...
    dimension_matches = re.findall(r'\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?(?:\s*x\s*\d+(?:\.\d+)?)?\b', text_clean)
    if dimension_matches:
        specs.extend([d.replace(" ", "") for d in dimension_matches])
        
    # 3. Các từ khóa model đặc thù: a4, pcb40, d90, phi 21 o90, model ...
    model_patterns = [
        r'\ba\d\b',             # A4, A3
        r'\bpcb\d+\b',          # PCB40, PCB30
        r'\bphi\s*\d+\b',       # Phi 21, phi 90
        r'\bd\d+\b',            # D90, D110
        r'\bmodel\s*[a-z0-9\-]+\b' # Model ABC-123
    ]
    for pattern in model_patterns:
        found = re.findall(pattern, text_clean)
        if found:
            specs.extend([f.strip() for f in found])
            
    return list(set(specs))

def match_commodity_row(row_desc, row_uom, list_commodities, auto_threshold=85, check_threshold=70):
    """
    So khớp một miêu tả hàng hóa với danh mục hàng hóa chuẩn.
    Trả về: (mã_gần_nhất, tên_gần_nhất, điểm_tương_thích, lý_do, là_mã_mới)
    """
    if pd.isna(row_desc) or not str(row_desc).strip():
        return "", "", 0, "Không có nội dung tên hàng hóa", False
        
    desc_norm = normalize_text(str(row_desc))
    desc_specs = extract_specs(str(row_desc))
    uom_norm = normalize_text(str(row_uom)) if not pd.isna(row_uom) else ""

    best_score = 0
    best_match = None
    best_reason = ""

    for item in list_commodities:
        # Cấu trúc của item phải chứa ít nhất:
        # 'ma_hang_hoa', 'ten_hang_hoa_chuan'
        # Có thể có: 'nhom_hang', 'don_vi_tinh', 'quy_cach', 'tu_khoa_nhan_dien'
        item_code = item.get('ma_hang_hoa', '')
        item_name_raw = item.get('ten_hang_hoa_chuan', '')
        item_name_norm = normalize_text(item_name_raw)
        
        item_uom = item.get('don_vi_tinh', '')
        item_uom_norm = normalize_text(item_uom)
        
        item_group = item.get('nhom_hang', '')
        item_keyword = item.get('tu_khoa_nhan_dien', '')
        item_specs = item.get('quy_cach', '')
        
        # Tính điểm thành phần theo đề xuất của bạn
        score = 0
        reasons_list = []
        
        # 1. Tên giống tuyệt đối sau chuẩn hóa
        if desc_norm == item_name_norm and desc_norm != "":
            score += 60
            reasons_list.append("Trùng tên tuyệt đối sau chuẩn hóa (+60)")
        else:
            # Tên giống trên 90% hoặc 75%
            ratio = fuzz.ratio(desc_norm, item_name_norm)
            if ratio >= 90:
                score += 50
                reasons_list.append(f"Tên gần giống {ratio:.1f}% (+50)")
            elif ratio >= 75:
                score += 35
                reasons_list.append(f"Tên gần giống {ratio:.1f}% (+35)")
                
        # 2. Trùng từ khóa nhận diện
        if item_keyword and isinstance(item_keyword, str):
            kw_list = [normalize_text(k) for k in item_keyword.split(",") if k.strip()]
            for kw in kw_list:
                if kw and kw in desc_norm:
                    score += 20
                    reasons_list.append(f"Trùng từ khóa '{kw}' (+20)")
                    break
                    
        # 3. Trùng đơn vị tính
        if uom_norm and item_uom_norm and uom_norm == item_uom_norm:
            score += 10
            reasons_list.append(f"Trùng ĐVT '{item_uom}' (+10)")
            
        # 4. Trùng quy cách/model/dung tích/trọng lượng trích xuất
        if desc_specs:
            target_item_specs_norm = normalize_text(str(item_specs)) + " " + item_name_norm
            matched_specs = []
            for spec in desc_specs:
                if spec in target_item_specs_norm:
                    matched_specs.append(spec)
            if matched_specs:
                score += 20
                reasons_list.append(f"Khớp quy cách/model {[s.upper() for s in matched_specs]} (+20)")
                
        # Giới hạn điểm tích lũy tối đa là 100
        score = min(score, 100)
        
        if score > best_score:
            best_score = score
            best_match = item
            best_reason = "; ".join(reasons_list) if reasons_list else "Độ tương khớp cơ bản"

    if best_match is None:
        return "", "", 0, "Không tìm thấy tương đồng trong danh mục", False

    return (
        best_match.get('ma_hang_hoa'),
        best_match.get('ten_hang_hoa_chuan'),
        best_score,
        best_reason,
        False
    )

def match_partner_row(row_name, row_mst, row_acc, invoice_desc="", list_partners=None, is_buyer=True, auto_threshold=85, check_threshold=70):
    """
    So khớp đối tượng (Khách hàng hoặc Nhà cung cấp) từ dữ liệu kế toán vào danh mục đối tác.
    Trả về: (mã_gần_nhất, tên_gần_nhất, loại_đối_tượng, điểm_tương_thích, lý_do, là_mã_mới)
    """
    if list_partners is None:
        list_partners = []
        
    name_str = str(row_name).strip() if not pd.isna(row_name) else ""
    mst_str = str(row_mst).strip() if not pd.isna(row_mst) else ""
    acc_str = str(row_acc).strip() if not pd.isna(row_acc) else ""
    invoice_str = str(invoice_desc).strip() if not pd.isna(invoice_desc) else ""
    
    if not name_str and not mst_str and not acc_str:
        return "", "", "", 0, "Không có thông tin đối tượng", False

    # Lọc danh sách đối tượng: Nếu là_buyer=True, ưu tiên 'Khách hàng', ngược lại ưu tiên 'Nhà cung cấp'
    # Tuy nhiên vẫn so khớp toàn bộ nếu không tìm được
    partner_norms = []
    for item in list_partners:
        p_type = item.get('loai_doi_tuong', '')
        # Phân loại cho điểm ưu tiên
        type_weight = 5 if ((is_buyer and p_type == "Khách hàng") or (not is_buyer and p_type == "Nhà cung cấp")) else 0
        partner_norms.append((item, type_weight))

    best_score = 0
    best_match = None
    best_reason = ""
    best_type = ""

    name_norm = normalize_partner_name(name_str) if name_str else ""
    invoice_norm = normalize_text(invoice_str) if invoice_str else ""

    for partner, type_weight in partner_norms:
        p_code = partner.get('ma_doi_tuyen', partner.get('ma_doi_tuong', ''))
        p_name_raw = partner.get('ten_doi_tuong', '')
        p_name_norm = normalize_partner_name(p_name_raw)
        
        p_mst = str(partner.get('ma_so_thue', '')).strip()
        p_acc = str(partner.get('so_tai_khoan', '')).strip()
        p_kw = partner.get('tu_khoa_nhan_dien', '')
        p_type = partner.get('loai_doi_tuong', 'Khách hàng')

        score = 0
        reasons_list = []

        # 1. Trùng mã số thuế (Chứng cứ thép)
        if mst_str and p_mst and mst_str.replace("-", "") == p_mst.replace("-", ""):
            score += 60
            reasons_list.append("Trùng Mã số thuế (+60)")

        # 2. Trùng số tài khoản ngân hàng
        if acc_str and p_acc and acc_str == p_acc:
            score += 50
            reasons_list.append("Trùng Số tài khoản (+50)")

        # 3. Tên giống tuyệt đối sau chuẩn hóa
        if name_norm and p_name_norm and name_norm == p_name_norm:
            score += 50
            reasons_list.append("Tên khớp tuyệt đối sau khi loại từ viết tắt (+50)")
        else:
            # Tên tương tự dùng kháp fuzzy
            if name_norm and p_name_norm:
                ratio = fuzz.ratio(name_norm, p_name_norm)
                if ratio >= 90:
                    score += 40
                    reasons_list.append(f"Tên viết tắt gần giống {ratio:.1f}% (+40)")
                elif ratio >= 75:
                    score += 25
                    reasons_list.append(f"Tên viết tắt gần giống {ratio:.1f}% (+25)")

        # 4. Trùng từ khóa nhận diện
        if p_kw and isinstance(p_kw, str) and name_str:
            kws = [normalize_text(k) for k in p_kw.split(",") if k.strip()]
            for kw in kws:
                if kw and kw in name_norm:
                    score += 20
                    reasons_list.append(f"Khớp từ khóa đối tác '{kw}' (+20)")
                    break

        # 5. Nội dung hóa đơn có chứa tên đối tượng
        if invoice_norm and p_name_norm and p_name_norm in invoice_norm:
            score += 20
            reasons_list.append("Nội dung diễn giải có chứa tên khách/nhà cung cấp (+20)")

        # Cộng thêm phần trọng số loại đối tượng ưu tiên để bẻ lái trùng điểm
        score += type_weight
        score = min(score, 100)

        if score > best_score:
            best_score = score
            best_match = partner
            best_reason = "; ".join(reasons_list) if reasons_list else "Khớp thông tin cơ bản"
            best_type = p_type

    if best_match is None:
        return "", "", "", 0, "Không có đối tượng phù hợp", False

    return (
        best_match.get('ma_doi_tuong', best_match.get('ma_doi_tuyen', '')),
        best_match.get('ten_doi_tuong', ''),
        best_type,
        best_score,
        best_reason,
        False
    )

def analyze_bank_transaction(desc, receive_amount, pay_amount, counterpart_acc, counterpart_name, list_partners=None, auto_threshold=85, check_threshold=70):
    """
    Phân tích một dòng giao dịch sao kê ngân hàng.
    Trả về một dict chứa:
    - nhom_giao_dich: Nhóm dự báo (Thu tiền KH, Chi NCC, Lương, Phí ngân hàng, Chuyển nội bộ, v.v.)
    - ma_doi_tuong_de_xuat: Mã đối tượng đề xuất nếu khớp
    - ten_doi_tuong_goc: Tên đối tượng gốc khớp được
    - diem_tin_cay: Từ 0-100
    - ly_do_du_doan: Ghi chú chi tiết thuật toán
    - can_kiem_tra: True/False
    """
    if list_partners is None:
        list_partners = []

    desc_str = str(desc).strip() if not pd.isna(desc) else ""
    desc_norm = normalize_text(desc_str)
    
    amount_in = float(receive_amount) if (not pd.isna(receive_amount) and str(receive_amount).strip() != "") else 0.0
    amount_out = float(pay_amount) if (not pd.isna(pay_amount) and str(pay_amount).strip() != "") else 0.0

    # 1. Kiểm tra các trường hợp đặc biệt không cần đối tác
    # 1a. Phí ngân hàng / Thuế / Lãi tiền gửi
    fee_keywords = ["phi dich vu", "phi quan ly", "sms banking", "phi giao dich", "phi chuyen tien", "tru phi", "phi cuoc", "phi thuong nien", "lai tien gui", "lai nhap goc", "tra lai"]
    if any(k in desc_norm for k in fee_keywords):
        category = "Phí ngân hàng / Tiền lãi"
        reason = "Phát hiện từ khóa liên quan đến dịch vụ và phí ngân hàng"
        if "lai" in desc_norm:
            category = "Lãi tiền gửi"
        return {
            "nhom_giao_dich": category,
            "ma_doi_tuong_de_xuat": "NGANHANG",
            "ten_doi_tuong_goc": "Hệ thống Ngân hàng",
            "diem_tin_cay": 95,
            "ly_do_du_doan": reason,
            "can_kiem_tra": False
        }

    # 1b. Lương / Tạm ứng nhân viên
    salary_keywords = ["luong", "thanh toan luong", "tiet kiem luong", "tam ung", "hoan ung", "khen thuong", "tro cap"]
    if any(k in desc_norm for k in salary_keywords):
        return {
            "nhom_giao_dich": "Chi lương / Tạm ứng",
            "ma_doi_tuong_de_xuat": "NHANVIEN",
            "ten_doi_tuong_goc": "Cán bộ nhân viên công ty",
            "diem_tin_cay": 90,
            "ly_do_du_doan": "Nội dung phản ánh thanh toán lương, thưởng hoặc tạm ứng công tác",
            "can_kiem_tra": True
        }

    # 1c. Thuế / Nộp ngân sách nhà nước
    tax_keywords = ["nop thue", "nop ngan sach", "thue mon bai", "thue gtgt", "thue tndn", "hai quan", "le phi"]
    if any(k in desc_norm for k in tax_keywords):
        return {
            "nhom_giao_dich": "Nộp thuế / Ngân sách",
            "ma_doi_tuong_de_xuat": "KHO_BAC",
            "ten_doi_tuong_goc": "Kho bạc Nhà nước",
            "diem_tin_cay": 95,
            "ly_do_du_doan": "Phát hiện nội dung nộp thuế, lệ phí hải quan ngân sách nhà nước",
            "can_kiem_tra": False
        }

    # 1d. Chuyển tiền nội bộ giữa các tài khoản
    internal_keywords = ["chuyen noi bo", "rut tien mat", "nap tien mat", "nop tien vao tai khoan", "rut tien nhap quy", "nop quy", "rut quy"]
    if any(k in desc_norm for k in internal_keywords):
        return {
            "nhom_giao_dich": "Giao dịch nội bộ / Quỹ",
            "ma_doi_tuong_de_xuat": "NOI_BO",
            "ten_doi_tuong_goc": "Nội bộ doanh nghiệp",
            "diem_tin_cay": 90,
            "ly_do_du_doan": "Nội dung phản ánh nộp/rút tiền mặt hoặc luân chuyển dòng tiền nội bộ",
            "can_kiem_tra": False
        }

    # 2. Định hướng dòng tiền để so khớp đối tác
    # Thu tiền -> ưu tiên Khách hàng (is_buyer=True)
    # Chi tiền -> ưu tiên Nhà cung cấp (is_buyer=False)
    is_buyer = amount_in > 0
    search_name = counterpart_name if (not pd.isna(counterpart_name) and str(counterpart_name).strip() != "") else desc_str
    
    p_code, p_name, p_type, p_score, p_reason, _ = match_partner_row(
        row_name=search_name,
        row_mst="",
        row_acc=counterpart_acc,
        invoice_desc=desc_str,
        list_partners=list_partners,
        is_buyer=is_buyer,
        auto_threshold=auto_threshold,
        check_threshold=check_threshold
    )

    if p_score >= check_threshold:
        # Nếu điểm khớp tốt
        if is_buyer:
            group = "Thu tiền khách hàng"
        else:
            group = "Chi thanh toán nhà cung cấp"
            
        can_check = p_score < auto_threshold
        return {
            "nhom_giao_dich": group,
            "ma_doi_tuong_de_xuat": p_code,
            "ten_doi_tuong_goc": p_name,
            "diem_tin_cay": int(p_score),
            "ly_do_du_doan": f"Khớp đối tác trong danh mục: {p_reason}",
            "can_kiem_tra": can_check
        }
    else:
        # Dưới ngưỡng: Không xác định rõ đối tác nào
        if is_buyer:
            group = "Thu tiền lập lờ (Chưa rõ nguồn)"
            reason = "Dòng tiền thu về nhưng không khớp danh mục đối tác hiện tại."
        elif amount_out > 0:
            group = "Chi phí chưa phân loại (Chưa rõ đối tác)"
            reason = "Dòng tiền chi ra nhưng không khớp danh mục nhà cung cấp hoặc cá nhân."
        else:
            group = "Chưa đủ thông tin phân loại"
            reason = "Giao dịch không có số tiền thu/chi hoặc không có nội dung rõ ràng."

        return {
            "nhom_giao_dich": group,
            "ma_doi_tuong_de_xuat": "",
            "ten_doi_tuong_goc": "",
            "diem_tin_cay": 40,
            "ly_do_du_doan": reason,
            "can_kiem_tra": True
        }

def find_next_available_code(prefix, existing_codes, counter_start=1):
    """
    Tạo mã tự động không trùng lặp dựa trên tiền tố.
    Ví dụ: prefix='HH', existing_codes=['HH001', 'HH002'] -> 'HH003'
    """
    i = counter_start
    while True:
        code = f"{prefix}{i:03d}"
        if code not in existing_codes:
            return code, i
        i += 1
