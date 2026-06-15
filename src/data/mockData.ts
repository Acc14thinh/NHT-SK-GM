import { Commodity, Partner } from "../types";

export const getSampleCommodities = (): Commodity[] => [
  {
    ma_hang_hoa: "HH001",
    ten_hang_hoa_chuan: "Xi măng Vicem Hà Tiên đa dụng",
    nhom_hang: "Vật liệu xây dựng",
    don_vi_tinh: "Bao",
    quy_cach: "PCB40 50kg",
    tu_khoa_nhan_dien: "ha tien, vicem, pcb40, xi mang",
    ghi_chu: "Mặt hàng chính bán sỉ"
  },
  {
    ma_hang_hoa: "HH002",
    ten_hang_hoa_chuan: "Thép cuộn Hòa Phát phi 6",
    nhom_hang: "Thép xây dựng",
    don_vi_tinh: "Kg",
    quy_cach: "Phi 6",
    tu_khoa_nhan_dien: "hoa phat, phi 6, thep cuon",
    ghi_chu: "Hàng nhập từ tổng nhà máy"
  },
  {
    ma_hang_hoa: "HH003",
    ten_hang_hoa_chuan: "Giấy in Double A A4 70gsm",
    nhom_hang: "Văn phòng phẩm",
    don_vi_tinh: "Ram",
    quy_cach: "A4 70gsm",
    tu_khoa_nhan_dien: "double a, a4, 70gsm, giay in",
    ghi_chu: "Đầu mối văn phòng phẩm"
  },
  {
    ma_hang_hoa: "HH004",
    ten_hang_hoa_chuan: "Sơn nội thất Dulux lau chùi hiệu quả 5L",
    nhom_hang: "Sơn & Hóa chất",
    don_vi_tinh: "Thùng",
    quy_cach: "5L",
    tu_khoa_nhan_dien: "dulux, lau chui, son noi that, 5l",
    ghi_chu: "Hàng đại lý cấp 1"
  }
];

export const getSamplePartners = (): Partner[] => [
  {
    ma_doi_tuong: "KH001",
    ten_doi_tuong: "Công ty TNHH Thương mại và Dịch vụ Minh Tâm",
    loai_doi_tuong: "Khách hàng",
    ma_so_thue: "0102030405",
    so_tai_khoan: "1100223344",
    ngan_hang: "Vietcombank",
    dia_chi: "123 Đường Láng, Đống Đa, Hà Nội",
    tu_khoa_nhan_dien: "minh tam, tm dv minh tam",
    ghi_chu: "Khách mua sỉ lâu năm"
  },
  {
    ma_doi_tuong: "NCC001",
    ten_doi_tuong: "Công ty Cổ phần Sắt Thép Hòa Phát",
    loai_doi_tuong: "Nhà cung cấp",
    ma_so_thue: "0304050607",
    so_tai_khoan: "5566778899",
    ngan_hang: "BIDV",
    dia_chi: "Khu nghiệp Phố Nối A, Hưng Yên",
    tu_khoa_nhan_dien: "hoa phat, sat thep hoa phat",
    ghi_chu: "Nhà bán thép chính"
  },
  {
    ma_doi_tuong: "NCC002",
    ten_doi_tuong: "Tổng Công ty Xi măng Việt Nam - Chi nhánh Hà Tiên",
    loai_doi_tuong: "Nhà cung cấp",
    ma_so_thue: "0908070605",
    so_tai_khoan: "9876543210",
    ngan_hang: "Techcombank",
    dia_chi: "360 Xa lộ Hà Nội, Quận 2, TP. Hồ Chí Minh",
    tu_khoa_nhan_dien: "xi mang ha tien, xi mang mien nam",
    ghi_chu: "Nhà bán xi măng bao lớn"
  },
  {
    ma_doi_tuong: "KH002",
    ten_doi_tuong: "Hộ Kinh Doanh Nguyễn Văn Hùng",
    loai_doi_tuong: "Khách hàng",
    ma_so_thue: "8011223344",
    so_tai_khoan: "22334455",
    ngan_hang: "Agribank",
    dia_chi: "Thôn 4, Hoài Đức, Hà Nội",
    tu_khoa_nhan_dien: "nguyen van hung, van hung",
    ghi_chu: "Khách lẻ tự chở"
  }
];

export interface RawPurchase {
  ngay_hoa_don: string;
  so_hoa_don: string;
  ky_hieu_hoa_don: string;
  ten_nguoi_ban: string;
  ma_so_thue_nguoi_ban: string;
  ten_hang_hoa_dich_vu: string;
  don_vi_tinh: string;
  so_luong: number;
  don_gia: number;
  thanh_tien: number;
  thue_gtgt: number;
  tong_thanh_toan: number;
  ghi_chu: string;
}

export const getSamplePurchaseLedger = (): RawPurchase[] => [
  {
    ngay_hoa_don: "2026-06-01",
    so_hoa_don: "0000124",
    ky_hieu_hoa_don: "C26TBA",
    ten_nguoi_ban: "TỔNG CÔNG TY XI MĂNG VIỆT NAM - CN HÀ TIÊN",
    ma_so_thue_nguoi_ban: "0908070605",
    ten_hang_hoa_dich_vu: "Xi măng Vicem Hà Tiên PCB40 đa dụng",
    don_vi_tinh: "Bao",
    so_luong: 200,
    don_gia: 85000,
    thanh_tien: 17000000,
    thue_gtgt: 1700000,
    tong_thanh_toan: 18700000,
    ghi_chu: "Thanh toán ngân hàng"
  },
  {
    ngay_hoa_don: "2026-06-02",
    so_hoa_don: "0000125",
    ky_hieu_hoa_don: "C26TBA",
    ten_nguoi_ban: "Công ty Cổ phần Sắt Thép Hòa Phát",
    ma_so_thue_nguoi_ban: "0304050607",
    ten_hang_hoa_dich_vu: "Thép cuộn phi 6 chính hãng Hòa Phát",
    don_vi_tinh: "Kg",
    so_luong: 1200,
    don_gia: 15000,
    thanh_tien: 18000000,
    thue_gtgt: 1800000,
    tong_thanh_toan: 19800000,
    ghi_chu: "Nhập kho cơ sở 1"
  },
  {
    ngay_hoa_don: "2026-06-03",
    so_hoa_don: "0000219",
    ky_hieu_hoa_don: "B26TBA",
    ten_nguoi_ban: "Công ty TNHH Sơn Nippon Việt Nam",
    ma_so_thue_nguoi_ban: "0405060708",
    ten_hang_hoa_dich_vu: "Sơn bóng cao cấp trong nhà Nippon 5 Lit",
    don_vi_tinh: "Thùng",
    so_luong: 15,
    don_gia: 1100000,
    thanh_tien: 16500000,
    thue_gtgt: 1650000,
    tong_thanh_toan: 18150000,
    ghi_chu: "Hàng nhập thử nghiệm"
  }
];

export interface RawSale {
  ngay_hoa_don: string;
  so_hoa_don: string;
  ky_hieu_hoa_don: string;
  ten_nguoi_mua: string;
  ma_so_thue_nguoi_mua: string;
  ten_hang_hoa_dich_vu: string;
  don_vi_tinh: string;
  so_luong: number;
  don_gia: number;
  thanh_tien: number;
  thue_gtgt: number;
  tong_thanh_toan: number;
  ghi_chu: string;
}

export const getSampleSalesLedger = (): RawSale[] => [
  {
    ngay_hoa_don: "2026-06-05",
    so_hoa_don: "0002001",
    ky_hieu_hoa_don: "K26TBA",
    ten_nguoi_mua: "CÔNG TY TNHH THƯƠNG MẠI MINH TÂM",
    ma_so_thue_nguoi_mua: "0102030405",
    ten_hang_hoa_dich_vu: "Giấy in Double A A4 ĐL 70gsm chất lượng",
    don_vi_tinh: "Ram",
    so_luong: 50,
    don_gia: 65000,
    thanh_tien: 3250000,
    thue_gtgt: 325000,
    tong_thanh_toan: 3575000,
    ghi_chu: "Mua hàng sỉ"
  },
  {
    ngay_hoa_don: "2026-06-06",
    so_hoa_don: "0002002",
    ky_hieu_hoa_don: "K26TBA",
    ten_nguoi_mua: "Ông Nguyễn Văn Hùng",
    ma_so_thue_nguoi_mua: "8011223344",
    ten_hang_hoa_dich_vu: "Xi măng Vicem Hà Tiên PCB40",
    don_vi_tinh: "Bao",
    so_luong: 80,
    don_gia: 95000,
    thanh_tien: 7600000,
    thue_gtgt: 760000,
    tong_thanh_toan: 8360000,
    ghi_chu: "Giao công trình Thôn 4"
  },
  {
    ngay_hoa_don: "2026-06-07",
    so_hoa_don: "0002003",
    ky_hieu_hoa_don: "K26TBA",
    ten_nguoi_mua: "Doanh nghiệp Tư nhân Vận tải Minh Long",
    ma_so_thue_nguoi_mua: "0506070809",
    ten_hang_hoa_dich_vu: "Dịch vụ vận chuyển bốc dỡ VLXD",
    don_vi_tinh: "Chuyến",
    so_luong: 2,
    don_gia: 2500000,
    thanh_tien: 5000000,
    thue_gtgt: 500000,
    tong_thanh_toan: 5500000,
    ghi_chu: "Vận chuyển ngoại thành"
  }
];

export interface RawInventory {
  ngay_chung_tu: string;
  so_chung_tu: string;
  loai_chung_tu: "Nhập" | "Xuất";
  ten_hang_hoa: string;
  don_vi_tinh: string;
  so_luong_nhap: number;
  gia_tri_nhap: number;
  so_luong_xuat: number;
  gia_tri_xuat: number;
  ten_doi_tuong: string;
  ghi_chu: string;
}

export const getSampleInventoryLedger = (): RawInventory[] => [
  {
    ngay_chung_tu: "2026-06-01",
    so_chung_tu: "N001",
    loai_chung_tu: "Nhập",
    ten_hang_hoa: "Xi mang da dung Vicem Ha Tien PCB40",
    don_vi_tinh: "Bao",
    so_luong_nhap: 200,
    gia_tri_nhap: 17000000,
    so_luong_xuat: 0,
    gia_tri_xuat: 0,
    ten_doi_tuong: "CN XI MANG VIET NAM HA TIEN",
    ghi_chu: "Nhập kho theo hóa đơn 124"
  },
  {
    ngay_chung_tu: "2026-06-05",
    so_chung_tu: "X001",
    loai_chung_tu: "Xuất",
    ten_hang_hoa: "Giay in Double A A4 70g",
    don_vi_tinh: "Ram",
    so_luong_nhap: 0,
    gia_tri_nhap: 0,
    so_luong_xuat: 50,
    gia_tri_xuat: 3250000,
    ten_doi_tuong: "CTY TNHH MINH TAM",
    ghi_chu: "Xuất hàng giao theo hóa đơn 001"
  }
];

export interface RawBankStatement {
  ngay_giao_dich: string;
  noi_dung_giao_dich: string;
  so_tien_thu: number;
  so_tien_chi: number;
  so_tai_khoan_doi_ung: string;
  ten_doi_tac_sao_ke: string;
  ghi_chu: string;
}

export const getSampleBankStatement = (): RawBankStatement[] => [
  {
    ngay_giao_dich: "2026-06-01",
    noi_dung_giao_dich: "CONG TY MINH TAM CHUYEN KHOAN MO HONG HOA DON GTGT 2001",
    so_tien_thu: 3575000,
    so_tien_chi: 0,
    so_tai_khoan_doi_ung: "1100223344",
    ten_doi_tac_sao_ke: "CTY TNHH TM DV MINH TAM",
    ghi_chu: "Khớp công nợ hóa đơn 2001"
  },
  {
    ngay_giao_dich: "2026-06-02",
    noi_dung_giao_dich: "VIETCOMBANK TRU PHI DUY TRI TAI KHOAN DOANH NGHIEP THANG 05/2026",
    so_tien_thu: 0,
    so_tien_chi: 22000,
    so_tai_khoan_doi_ung: "VCB-SYSTEM",
    ten_doi_tac_sao_ke: "VIETCOMBANK CO",
    ghi_chu: "Hạch toán phí quản trị"
  },
  {
    ngay_giao_dich: "2026-06-03",
    noi_dung_giao_dich: "UNC THANH TOAN TIÊN THEP CUON CHO KHO HOA PHAT KEM HD 125",
    so_tien_thu: 0,
    so_tien_chi: 19800000,
    so_tai_khoan_doi_ung: "5566778899",
    ten_doi_tac_sao_ke: "CONG TY SAT THEP HOA PHAT",
    ghi_chu: "Thanh toán lấy thép"
  },
  {
    ngay_giao_dich: "2026-06-04",
    noi_dung_giao_dich: "NOP THU THUE GIA TRI GIA TANG KHO BAC NHA NUOC",
    so_tien_thu: 0,
    so_tien_chi: 5500000,
    so_tai_khoan_doi_ung: "7111KOBAC",
    ten_doi_tac_sao_ke: "KHO BAC NHA NUOC QUAN DONG DA",
    ghi_chu: "Thuế giá trị gia tăng"
  },
  {
    ngay_giao_dich: "2026-06-05",
    noi_dung_giao_dich: "UNC CHI LUONG THANG 5/2026 VP CONG TY",
    so_tien_thu: 0,
    so_tien_chi: 45000000,
    so_tai_khoan_doi_ung: "",
    ten_doi_tac_sao_ke: "CAN BO CONG NHAN VIEN",
    ghi_chu: "Chi thanh toán lương nhân công"
  },
  {
    ngay_giao_dich: "2026-06-06",
    noi_dung_giao_dich: "NGUYEN VAN HUNG CK TRA TIEN CHO HOA DON 0002002",
    so_tien_thu: 8360000,
    so_tien_chi: 0,
    so_tai_khoan_doi_ung: "22334455",
    ten_doi_tac_sao_ke: "ONG NGUYEN VAN HUNG",
    ghi_chu: "Thu hồi nợ khách lẻ"
  },
  {
    ngay_giao_dich: "2026-06-07",
    noi_dung_giao_dich: "VIVU CAFE THANH TOAN TIỀN CAP COFFEE CUOI TUAN - KHONG DUNG CO DOANH NGHIEP",
    so_tien_thu: 0,
    so_tien_chi: 350000,
    so_tai_khoan_doi_ung: "998811",
    ten_doi_tac_sao_ke: "LAI VAN TEO CAFE",
    ghi_chu: "Chi tiêu chi nhánh cá nhân"
  }
];
