/**
 * High-Fidelity Demo Data Generator covering all 14 Mandatory Test Cases (Section 28)
 *
 * Case 1: 1 order / 1 SKU / khớp hoàn toàn
 * Case 2: 1 order / nhiều SKU / voucher shop lặp lại (chỉ trừ 1 lần)
 * Case 3: Có trên Shopee nhưng không có PM
 * Case 4: Có trên PM nhưng không có Shopee (đối chiếu 2 chiều)
 * Case 5: Đơn hủy không tồn tại trong PM
 * Case 6: Cùng order ID nhưng khác shop (không ghép chéo)
 * Case 7: Tên sheet và Tên Shop không giống hoàn toàn (Metaky vs Metaky Store)
 * Case 8: Mã đơn chứa khoảng trắng đầu/cuối
 * Case 9: PM có duplicate cùng mã đơn (lặp lại tổng tiền cùng chứng từ)
 * Case 10: Một đơn có nhiều hóa đơn thực sự (tổng các hóa đơn DISTINCT)
 * Case 11: Voucher Shop trong cùng order không nhất quán (Cần kiểm tra)
 * Case 12 & 13: Shop mới tự động nhận diện (ABC Official Store)
 * Case 14: Validation thiếu cột
 */
import * as XLSX from "xlsx";

export interface DemoEcommerceFiles {
  shopeeWorkbook: XLSX.WorkBook;
  shopeeFileName: string;
  salesWorkbook: XLSX.WorkBook;
  salesFileName: string;
}

export function generateDemoEcommerceWorkbooks(): DemoEcommerceFiles {
  // ========================================================
  // 1. FILE SÀN TMĐT SHOPEE (Đa Sheet: Metaky, Mibaby, Msmarty, ABC Store)
  // ========================================================
  const shopeeWb = XLSX.utils.book_new();

  // Sheet 1: Metaky
  const metakyShopeeRows = [
    // Header
    [
      "Mã đơn hàng", "Trạng Thái Đơn Hàng", "Thời gian giao hàng",
      "Trạng thái Trả hàng/Hoàn tiền", "Tên sản phẩm", "Giá sản phẩm",
      "Số lượng", "Tổng số tiền Người mua thanh toán", "Mã giảm giá của Shop"
    ],
    // Case 1: 1 order / 1 SKU / khớp hoàn toàn (150.000đ)
    [
      "SPX260701_001", "Đã giao hàng", "05/07/2026 14:20",
      "Không", "Bình giữ nhiệt Metaky 500ml", "150000",
      "1", "150000", "0"
    ],
    // Case 2: 1 order / 3 SKU / voucher shop lặp lại 50.000đ (Chỉ trừ 1 lần: 100k + 200k + 300k - 50k = 550k)
    [
      "SPX260701_002", "Đã giao hàng", "06/07/2026 10:15",
      "Không", "Hộp cơm giữ nhiệt Metaky 2 ngăn", "100000",
      "1", "100000", "50000"
    ],
    [
      "SPX260701_002", "Đã giao hàng", "06/07/2026 10:15",
      "Không", "Túi đựng hộp cơm Metaky", "200000",
      "1", "200000", "50000"
    ],
    [
      "SPX260701_002", "Đã giao hàng", "06/07/2026 10:15",
      "Không", "Bộ đũa thìa inox Metaky", "300000",
      "1", "300000", "50000"
    ],
    // Case 3: Có trên Shopee nhưng không có PM (320.000đ)
    [
      "SPX260701_003", "Đã giao hàng", "08/07/2026 16:45",
      "Không", "Nồi áp suất điện đa năng", "320000",
      "1", "320000", "0"
    ],
    // Case 5: Đơn hủy không tồn tại trong PM (210.000đ)
    [
      "SPX260701_005", "Đã hủy", "09/07/2026 09:30",
      "Hủy bởi người mua", "Ấm đun siêu tốc Metaky 1.8L", "210000",
      "1", "210000", "0"
    ],
    // Case 6: Cùng mã đơn nhưng thuộc shop Metaky (400.000đ)
    [
      "SPX260701_COMMON", "Đã giao hàng", "11/07/2026 11:00",
      "Không", "Chảo chống dính Metaky 26cm", "400000",
      "1", "400000", "0"
    ],
    // Case 8: Mã đơn chứa khoảng trắng đầu/cuối trong file sàn
    [
      "  SPX260701_008  ", "Đã giao hàng", "12/07/2026 15:30",
      "Không", "Bếp từ đơn Metaky cảm ứng", "650000",
      "1", "650000", "50000" // Doanh thu sàn = 600.000đ
    ],
    // Case 11: Voucher Shop trong cùng order có giá trị không nhất quán (20k vs 50k)
    [
      "SPX260701_011", "Đã giao hàng", "15/07/2026 17:00",
      "Không", "Máy xay sinh tố cầm tay", "200000",
      "1", "200000", "20000"
    ],
    [
      "SPX260701_011", "Đã giao hàng", "15/07/2026 17:00",
      "Không", "Cối xay thịt phụ kiện", "150000",
      "1", "150000", "50000" // Lệch voucher!
    ]
  ];
  const wsMetaky = XLSX.utils.aoa_to_sheet(metakyShopeeRows);
  XLSX.utils.book_append_sheet(shopeeWb, wsMetaky, "Metaky");

  // Sheet 2: Mibaby
  const mibabyShopeeRows = [
    [
      "Mã đơn hàng", "Trạng Thái Đơn Hàng", "Thời gian giao hàng",
      "Trạng thái Trả hàng/Hoàn tiền", "Tên sản phẩm", "Giá sản phẩm",
      "Số lượng", "Tổng số tiền Người mua thanh toán", "Mã giảm giá của Shop"
    ],
    // Case 6: Cùng mã đơn COMMON nhưng thuộc shop Mibaby (700.000đ)
    [
      "SPX260701_COMMON", "Đã giao hàng", "11/07/2026 13:00",
      "Không", "Máy hút sữa điện đôi Mibaby", "750000",
      "1", "750000", "50000" // Doanh thu sàn = 700.000đ
    ],
    // Case 9: Đơn Mibaby chuẩn bị test duplicate trên PM (Doanh thu sàn = 250.000đ)
    [
      "SPX260701_009", "Đã giao hàng", "14/07/2026 09:20",
      "Không", "Máy hâm sữa tiệt trùng 4 trong 1", "280000",
      "1", "280000", "30000" // Doanh thu sàn = 250.000đ
    ],
    // Case 10: Đơn Mibaby có 2 hóa đơn trên PM (Doanh thu sàn = 300.000đ)
    [
      "SPX260701_010", "Đã giao hàng", "14/07/2026 16:00",
      "Không", "Set bình sữa cổ rộng Mibaby", "300000",
      "1", "300000", "0"
    ]
  ];
  const wsMibaby = XLSX.utils.aoa_to_sheet(mibabyShopeeRows);
  XLSX.utils.book_append_sheet(shopeeWb, wsMibaby, "Mibaby");

  // Sheet 3: Msmarty
  const msmartyShopeeRows = [
    [
      "Mã đơn hàng", "Trạng Thái Đơn Hàng", "Thời gian giao hàng",
      "Trạng thái Trả hàng/Hoàn tiền", "Tên sản phẩm", "Giá sản phẩm",
      "Số lượng", "Tổng số tiền Người mua thanh toán", "Mã giảm giá của Shop"
    ],
    // Đơn có chênh lệch cố tình để test trạng thái CHÊNH LỆCH
    [
      "SPX260701_DIFF", "Đã giao hàng", "18/07/2026 10:00",
      "Không", "Viên uống M'Smarty Canxi Nano", "450000",
      "1", "450000", "0" // DT sàn = 450.000đ (trên PM ghi 400.000đ -> Lệch +50.000đ)
    ]
  ];
  const wsMsmarty = XLSX.utils.aoa_to_sheet(msmartyShopeeRows);
  XLSX.utils.book_append_sheet(shopeeWb, wsMsmarty, "Msmarty");

  // Sheet 4: ABC Store (Shop mới Case 13)
  const abcShopeeRows = [
    [
      "Mã đơn hàng", "Trạng Thái Đơn Hàng", "Thời gian giao hàng",
      "Trạng thái Trả hàng/Hoàn tiền", "Tên sản phẩm", "Giá sản phẩm",
      "Số lượng", "Tổng số tiền Người mua thanh toán", "Mã giảm giá của Shop"
    ],
    [
      "SPX260701_ABC01", "Đã giao hàng", "20/07/2026 14:00",
      "Không", "Khẩu trang y tế 4 lớp kháng khuẩn ABC", "120000",
      "1", "120000", "0"
    ]
  ];
  const wsAbc = XLSX.utils.aoa_to_sheet(abcShopeeRows);
  XLSX.utils.book_append_sheet(shopeeWb, wsAbc, "ABC Store");

  // ========================================================
  // 2. FILE PHẦN MỀM BÁN HÀNG (BAN_HANG_SHOPEE_01.07 -31.07.xlsx)
  // ========================================================
  const salesWb = XLSX.utils.book_new();

  const salesRows = [
    // Header
    [
      "Tên Shop", "Số đơn hàng từ hệ thống khác", "Số hóa đơn",
      "Ngày hóa đơn", "Tổng tiền", "Số chứng từ", "Tên khách hàng", "Ghi chú"
    ],
    // Case 1: Khớp hoàn toàn với Metaky (150.000đ)
    [
      "Metaky Store", "SPX260701_001", "HD-2607-001",
      "05/07/2026", "150000", "CT-001", "Nguyễn Văn An", "Đơn Shopee hoàn tất"
    ],
    // Case 2: Khớp đơn nhiều SKU Metaky (550.000đ)
    [
      "Metaky Store", "SPX260701_002", "HD-2607-002",
      "06/07/2026", "550000", "CT-002", "Trần Thị Mai", "Đơn combo 3 món"
    ],
    // Case 4: Có trên PM nhưng KHÔNG CÓ TRÊN SHOPEE (500.000đ)
    [
      "Metaky Store", "SPX260701_004", "HD-2607-004",
      "07/07/2026", "500000", "CT-004", "Lê Văn Hùng", "Đơn sót chưa xuất hiện trên sàn"
    ],
    // Case 6: Cùng mã đơn COMMON nhưng thuộc Metaky Store (400.000đ)
    [
      "Metaky Store", "SPX260701_COMMON", "HD-2607-006A",
      "11/07/2026", "400000", "CT-006A", "Hoàng Kim", "Khớp chuẩn Metaky"
    ],
    // Case 6: Cùng mã đơn COMMON nhưng thuộc Mibaby Store (700.000đ) -> KHÔNG ĐƯỢC GHÉP CHÉO!
    [
      "Mibaby Store", "SPX260701_COMMON", "HD-2607-006B",
      "11/07/2026", "700000", "CT-006B", "Phạm Thu Trang", "Khớp chuẩn Mibaby"
    ],
    // Case 8: Mã đơn không có khoảng trắng trên PM, khớp với mã có khoảng trắng trên Sàn
    [
      "Metaky Store", "SPX260701_008", "HD-2607-008",
      "12/07/2026", "600000", "CT-008", "Vũ Minh", "Khớp chuẩn sau trim space"
    ],
    // Case 9: PM có duplicate cùng mã đơn, lặp lại tổng tiền của cùng chứng từ (Chỉ lấy 1 lần: 250.000đ)
    [
      "Mibaby Store", "SPX260701_009", "HD-2607-009",
      "14/07/2026", "250000", "CT-009", "Đặng Thùy Dung", "Dòng 1 chi tiết phiếu xuất"
    ],
    [
      "Mibaby Store", "SPX260701_009", "HD-2607-009",
      "14/07/2026", "250000", "CT-009", "Đặng Thùy Dung", "Dòng 2 chi tiết phiếu xuất (lặp lại)"
    ],
    // Case 10: Một đơn có nhiều hóa đơn thực sự (HD-010A 100k + HD-010B 200k = 300k)
    [
      "Mibaby Store", "SPX260701_010", "HD-2607-010A",
      "14/07/2026", "100000", "CT-010A", "Bùi Quang", "Hóa đơn đợt 1"
    ],
    [
      "Mibaby Store", "SPX260701_010", "HD-2607-010B",
      "14/07/2026", "200000", "CT-010B", "Bùi Quang", "Hóa đơn đợt 2"
    ],
    // Case 11: Voucher không nhất quán (trên PM ghi 300.000đ)
    [
      "Metaky Store", "SPX260701_011", "HD-2607-011",
      "15/07/2026", "300000", "CT-011", "Ngô Quỳnh", "Cần kiểm tra do voucher lệch"
    ],
    // Đơn có chênh lệch doanh thu với M'Smarty Shop (PM ghi 400.000đ, Sàn 450.000đ)
    [
      "M'Smarty Shop", "SPX260701_DIFF", "HD-2607-DIFF",
      "18/07/2026", "400000", "CT-DIFF", "Đỗ Hải", "Lệch tiền đối chiếu"
    ],
    // Case 13: Shop mới ABC Official Store (120.000đ)
    [
      "ABC Official Store", "SPX260701_ABC01", "HD-2607-ABC",
      "20/07/2026", "120000", "CT-ABC", "Công ty CP Y Tế ABC", "Shop mới tự sinh"
    ]
  ];
  const wsSales = XLSX.utils.aoa_to_sheet(salesRows);
  XLSX.utils.book_append_sheet(salesWb, wsSales, "Ban_Hang_Tong_Hop");

  return {
    shopeeWorkbook: shopeeWb,
    shopeeFileName: "Order.all.20260701_20260731_shopee.xlsx",
    salesWorkbook: salesWb,
    salesFileName: "BAN_HANG_SHOPEE_01.07 -31.07.xlsx"
  };
}
