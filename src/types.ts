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
}

export interface ColumnMapping {
  // Mapping of key values
  [key: string]: string;
}
