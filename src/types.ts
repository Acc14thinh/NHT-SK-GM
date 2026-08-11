/**
 * Type declarations for the Auto-Accounting Mapper client-side UI.
 */

export interface Commodity {
  ma_hang_hoa: string;
  ten_hang_hoa_chuan: string;
  nhom_hang: string;
  don_vi_tinh: string;
  quy_cach: string;
  tu_khoa_nhan_dien: string;
  ghi_chu: string;
  don_gia_tham_chieu?: number;
  don_gia_mua_gan_nhat?: number;
  don_gia_ban_gan_nhat?: number;
}

export interface Partner {
  ma_doi_tuong: string;
  ten_doi_tuong: string;
  loai_doi_tuong: "Khách hàng" | "Nhà cung cấp";
  ma_so_thue: string;
  so_tai_khoan: string;
  ngan_hang: string;
  dia_chi: string;
  tu_khoa_nhan_dien: string;
  ghi_chu: string;
}

export interface MatchingConfig {
  autoThreshold: number;  // 85% by default
  checkThreshold: number; // 70% by default
  prefixHH: string;       // "HH"
  prefixKH: string;       // "KH"
  prefixNCC: string;      // "NCC"
  daysBeforeInvoice?: number; // default 7
  daysAfterInvoice?: number;  // default 30
  diffAbsThreshold?: number;  // default 10000
  diffPctThreshold?: number;  // default 0.5
  maxCombinationCount?: number; // default 5

  // Cấu hình bổ sung tiêu chí Đơn giá hàng hóa
  enablePriceMatching?: boolean;
  allowDerivedPrice?: boolean;
  priceRefSource?: "median" | "latest" | "master";
  priceDiffThresholdVeryHigh?: number; // e.g., 2% -> 20 pts
  priceDiffThresholdHigh?: number;     // e.g., 5% -> 15 pts
  priceDiffThresholdMedium?: number;   // e.g., 10% -> 10 pts
  priceDiffThresholdLow?: number;      // e.g., 20% -> 5 pts
  allowUomConversion?: boolean;
}

export interface CommodityCandidate {
  commodity: Commodity;
  totalScore: number;
  scoreName: number;
  scoreSpecs: number;
  scoreUom: number;
  scorePrice: number;
  scoreCategory: number;
  priceDiffPct?: number | null;
  refPrice?: number | null;
}

export interface MappedRow {
  id: string;
  originalText: string;
  originalUom?: string;
  originalMst?: string;
  originalAcc?: string;
  proposedCode: string;
  proposedName: string;
  proposedType?: string;
  score: number;
  reason: string;
  treatment: "TỰ ĐỘNG GẮN" | "DUYỆT THỦ CÔNG" | "TẠO MÃ MỚI" | "BỎ QUA";
  notes: string;
  rawRowData: { [key: string]: any }; // Holds copy of all standard uploaded columns

  // Bổ sung các chỉ số rà soát đơn giá & Top ứng viên
  rawPrice?: number | null;
  normalizedPrice?: number | null;
  priceSource?: "Cột đơn giá gốc" | "Suy ra" | "Không có";
  refPrice?: number | null;
  refPriceType?: "Trung vị" | "Mua mới nhất" | "Bán mới nhất" | "Chung";
  priceDiffAmt?: number | null;
  priceDiffPct?: number | null;
  scoreName?: number;
  scoreSpecs?: number;
  scoreUom?: number;
  scorePrice?: number;
  scoreCategory?: number;
  priceWarning?: string | null;
  top3Candidates?: CommodityCandidate[];
  processingStatus?: string;
}

export interface BankAnalysisResult {
  id: string;
  date: string;
  desc: string;
  amountIn: number;
  amountOut: number;
  counterpartAcc?: string;
  counterpartName?: string;
  predictedGroup: string;
  proposedCode: string;
  proposedName: string;
  score: number;
  reason: string;
  treatment: "Đã chốt" | "Cần kiểm tra" | "Bỏ qua";
  notes: string;
  rawRowData?: { [key: string]: any };

  // Bổ sung các cột phục vụ đối chiếu hóa đơn nâng cao
  proposedType?: string;          // Loại đối tượng ("Khách hàng" | "Nhà cung cấp")
  matchingSource?: string;        // Nguồn gắn mã (ví dụ: "Diễn giải", "Hóa đơn", "Số tiền & Ngày")
  matchedInvoiceNo?: string;      // Số hóa đơn khớp
  matchedInvoiceDate?: string;    // Ngày hóa đơn khớp
  invoiceAmount?: number;         // Giá trị hóa đơn
  differenceAmount?: number;      // Chênh lệch số tiền
  differencePercentage?: number;  // Tỷ lệ chênh lệch (%)
  differenceDays?: number;        // Số ngày chênh lệch
  scoreDesc?: number;             // Điểm khớp nội dung
  scoreName?: number;             // Điểm khớp tên
  scoreAcc?: number;              // Điểm khớp số tài khoản
  scoreMst?: number;              // Điểm khớp mã số thuế
  scoreInvoice?: number;          // Điểm khớp số hóa đơn
  scoreAmount?: number;           // Điểm khớp số tiền
  scoreDate?: number;             // Điểm khớp ngày
  scoreHistory?: number;          // Điểm lịch sử xác nhận
  scorePenalty?: number;          // Điểm phạt mâu thuẫn
  totalScore?: number;            // Tổng điểm tin cậy (0 - 100)
  top3Proposals?: Array<{         // Top 3 phương án đề xuất phù hợp nhất
    code: string;
    name: string;
    invoiceNo?: string;
    invoiceDate?: string;
    invoiceAmount?: number;
    score: number;
    reason: string;
  }>;
  processingStatus?: string;      // Trạng thái xử lý (Tự động gắn theo nội dung, Ghép nhiều hóa đơn, v.v.)
}

export interface ColumnMapping {
  // Mapping of key values
  [key: string]: string;
}
