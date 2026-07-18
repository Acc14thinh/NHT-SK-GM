import * as XLSX from "xlsx";
import { ColumnMapping } from "../types";

/**
 * Clean and normalize cell content or headers by resolving duplicate or empty names.
 */
export function cleanAndDeduplicateHeaders(headers: string[]): string[] {
  const seen: { [key: string]: number } = {};
  return headers.map((h, index) => {
    let clean = String(h || "").trim()
      .replace(/[\r\n]+/g, " ") // replace newlines with space
      .replace(/\s+/g, " ");    // collapse spaces

    if (!clean) {
      clean = `Cột_Trống_${index + 1}`;
    }

    if (seen[clean] !== undefined) {
      seen[clean]++;
      clean = `${clean}_${seen[clean]}`;
    } else {
      seen[clean] = 0;
    }

    return clean;
  });
}

/**
 * Extract clean, unified headers from raw Excel rows considering single or double row headers.
 * Implements left-to-right forward fill to handle merged title cells correctly (Yêu cầu 8).
 */
export function getHeadersFromRawRows(
  rawRows: any[][],
  headerRowIndex: number,
  headerRowsCount: number
): string[] {
  if (!rawRows || rawRows.length === 0 || headerRowIndex < 0 || headerRowIndex >= rawRows.length) {
    return [];
  }

  // Row 1 (main header row)
  const row1 = rawRows[headerRowIndex] || [];
  
  if (headerRowsCount === 1) {
    // Standard single row header
    const rawHeaders = row1.map(val => {
      if (val === null || val === undefined) return "";
      return String(val).trim();
    });
    return cleanAndDeduplicateHeaders(rawHeaders);
  }

  // Double row header (merged cells)
  const row2 = rawRows[headerRowIndex + 1] || [];
  
  // Forward-fill (left-to-right) on the first row to handle merged columns
  const filledRow1: string[] = [];
  let lastVal = "";
  
  const maxLen = Math.max(row1.length, row2.length);
  for (let i = 0; i < maxLen; i++) {
    const val1 = row1[i];
    const str1 = (val1 !== null && val1 !== undefined) ? String(val1).trim() : "";
    if (str1 !== "") {
      lastVal = str1;
    }
    filledRow1.push(lastVal);
  }

  // Combine row1 and row2
  const combinedHeaders: string[] = [];
  for (let i = 0; i < maxLen; i++) {
    const h1 = filledRow1[i] || "";
    const val2 = row2[i];
    const h2 = (val2 !== null && val2 !== undefined) ? String(val2).trim() : "";
    
    let combined = "";
    if (h1 && h2) {
      if (h1 === h2) {
        combined = h1;
      } else {
        combined = `${h1} - ${h2}`;
      }
    } else if (h1) {
      combined = h1;
    } else if (h2) {
      combined = h2;
    } else {
      combined = "";
    }
    combinedHeaders.push(combined);
  }

  return cleanAndDeduplicateHeaders(combinedHeaders);
}

/**
 * Score a row based on common keywords to detect if it serves as a header.
 */
export function scoreHeaderRow(row: any[]): number {
  if (!row || !Array.isArray(row)) return 0;

  const categories = {
    date: ["ngày", "ngay", "date", "trans", "hạch toán", "hiệu lực", "ngay_gd"],
    content: ["nội dung", "noi dung", "diễn giải", "dien giai", "mô tả", "description", "remark", "narrative", "ghichu", "ghi chú"],
    credit: ["có", "ghi có", "credit", "deposit", "thu", "tiền vào", "phát sinh có", "ps có"],
    debit: ["nợ", "ghi nợ", "debit", "withdrawal", "chi", "tiền ra", "phát sinh nợ", "ps nợ"],
    balance: ["số dư", "so du", "balance", "lũy kế"],
    partner: ["đối tác", "đối tượng", "khách hàng", "nhà cung cấp", "người nhận", "người gửi", "sender", "receiver", "beneficiary", "kh_ncc"],
    counterAccount: ["đối ứng", "doi ung", "tài khoản đối ứng", "tk đối ứng", "stk_doi_ung"]
  };

  let score = 0;

  Object.entries(categories).forEach(([_, keywords]) => {
    const hasMatch = row.some(cell => {
      if (cell === undefined || cell === null) return false;
      const str = String(cell).toLowerCase().trim();
      return keywords.some(kw => str.includes(kw));
    });

    if (hasMatch) {
      score += 15;
    }
  });

  // Count populated text cells to reward typical headers (usually 4 to 15 columns)
  const textCellsCount = row.filter(cell => cell !== undefined && cell !== null && String(cell).trim() !== "").length;
  if (textCellsCount >= 4 && textCellsCount <= 18) {
    score += textCellsCount * 2;
  }

  return score;
}

/**
 * Propose initial bank mapping suggestions based on keyword matching.
 */
export function proposeBankMappings(headers: string[]): ColumnMapping {
  const mapping: ColumnMapping = {
    ngay_giao_dich: "",
    ngay_hieu_luc: "",
    ngay_hach_toan: "",
    noi_dung_giao_dich: "",
    so_tien_thu: "",
    so_tien_chi: "",
    so_du: "",
    ma_giao_dich: "",
    so_chung_tu: "",
    so_tham_chieu: "",
    ten_doi_tac_sao_ke: "",
    so_tai_khoan_doi_ung: "",
    ngan_hang_doi_ung: "",
    ghi_chu: "",
    loai_tien_te: "",
    cot_ngay: "",
    cot_thang: "",
    cot_nam: ""
  };

  headers.forEach(h => {
    const norm = h.toLowerCase().trim();

    // 1. Transaction Date
    if (
      (norm.includes("ngày") && !norm.includes("hiệu lực") && !norm.includes("hạch toán") && !norm.includes("đối ứng") && !norm.includes("hóa đơn") && norm !== "ngày") ||
      norm === "ngày" ||
      norm === "ngay" ||
      norm.includes("ngay_gd") ||
      norm.includes("transaction date") ||
      norm.includes("trans_date")
    ) {
      if (!mapping.ngay_giao_dich) mapping.ngay_giao_dich = h;
    }

    if (norm.includes("hiệu lực") || norm.includes("value date") || norm.includes("value_date")) {
      if (!mapping.ngay_hieu_luc) mapping.ngay_hieu_luc = h;
    }

    if (norm.includes("hạch toán") || norm.includes("posting date") || norm.includes("posting_date") || norm.includes("ngay_ht")) {
      if (!mapping.ngay_hach_toan) mapping.ngay_hach_toan = h;
    }

    // 2. Transaction Content
    if (
      (norm.includes("nội dung") || norm.includes("noi dung") || norm.includes("diễn giải") || norm.includes("dien giai") || norm.includes("mô tả") || norm.includes("description") || norm.includes("remarks") || norm.includes("narrative")) &&
      !norm.includes("tháng") && !norm.includes("ngày") && !norm.includes("số tiền")
    ) {
      if (!mapping.noi_dung_giao_dich) mapping.noi_dung_giao_dich = h;
    }

    // 3. Deposit (Credit)
    if (
      (norm.includes("có") || norm.includes("ghi có") || norm.includes("tiền vào") || norm.includes("credit") || norm.includes("deposit") || norm.includes("thu") || norm.includes("ps có") || norm.includes("phát sinh có")) &&
      !norm.includes("ngày") && !norm.includes("tháng") && !norm.includes("nợ")
    ) {
      if (!mapping.so_tien_thu) mapping.so_tien_thu = h;
    }

    // 4. Withdrawal (Debit)
    if (
      (norm.includes("nợ") || norm.includes("ghi nợ") || norm.includes("tiền ra") || norm.includes("debit") || norm.includes("withdrawal") || norm.includes("chi") || norm.includes("ps nợ") || norm.includes("phát sinh nợ")) &&
      !norm.includes("ngày") && !norm.includes("tháng") && !norm.includes("có")
    ) {
      if (!mapping.so_tien_chi) mapping.so_tien_chi = h;
    }

    // 5. Balance
    if (
      norm.includes("số dư") || norm.includes("so du") || norm.includes("balance") || norm.includes("closing balance") || norm.includes("running balance") || norm.includes("lũy kế")
    ) {
      if (!mapping.so_du) mapping.so_du = h;
    }

    // 6. Counter account
    if (
      norm.includes("tài khoản đối ứng") || norm.includes("tk đối ứng") || norm.includes("so_tai_khoan_doi_ung") || norm.includes("stk đối ứng") || norm.includes("stk_doi_ung") || norm.includes("counter account") || norm.includes("beneficiary account") || norm.includes("tk đối tượng")
    ) {
      if (!mapping.so_tai_khoan_doi_ung) mapping.so_tai_khoan_doi_ung = h;
    }

    // 7. Partner Name
    if (
      norm.includes("tên đối tác") || norm.includes("tên đối tượng") || norm.includes("người chuyển") || norm.includes("người nhận") || norm.includes("tên người chuyển") || norm.includes("beneficiary") || norm.includes("sender") || norm.includes("receiver") || norm.includes("counterparty") || norm.includes("tên khách hàng") || norm.includes("tên ncc")
    ) {
      if (!mapping.ten_doi_tac_sao_ke) mapping.ten_doi_tac_sao_ke = h;
    }

    // 8. Other Fields
    if (norm.includes("mã giao dịch") || norm.includes("ma_gd") || norm.includes("transaction id") || norm.includes("txid") || norm.includes("ref_no") || norm.includes("reference")) {
      if (!mapping.ma_giao_dich) mapping.ma_giao_dich = h;
    }
    if (norm.includes("chứng từ") || norm.includes("document") || norm.includes("số ct")) {
      if (!mapping.so_chung_tu) mapping.so_chung_tu = h;
    }
    if (norm.includes("tham chiếu") || norm.includes("reference") || norm.includes("ref")) {
      if (!mapping.so_tham_chieu) mapping.so_tham_chieu = h;
    }
    if (norm.includes("ngân hàng đối ứng") || norm.includes("ngan_hang_doi_ung") || norm.includes("nh đối ứng") || norm.includes("nh_doi_ung")) {
      if (!mapping.ngan_hang_doi_ung) mapping.ngan_hang_doi_ung = h;
    }
    if (norm.includes("ghi chú") || norm.includes("ghi chu") || norm === "note" || norm === "notes") {
      if (!mapping.ghi_chu) mapping.ghi_chu = h;
    }
    if (norm.includes("tiền tệ") || norm === "currency" || norm === "ccy") {
      if (!mapping.loai_tien_te) mapping.loai_tien_te = h;
    }

    // Split Date columns
    if (norm === "ngày" || norm === "day") {
      if (!mapping.cot_ngay) mapping.cot_ngay = h;
    }
    if (norm === "tháng" || norm === "month") {
      if (!mapping.cot_thang) mapping.cot_thang = h;
    }
    if (norm === "năm" || norm === "year") {
      if (!mapping.cot_nam) mapping.cot_nam = h;
    }
  });

  return mapping;
}

/**
 * Validate mapping selections to ensure required fields are correctly mapped.
 */
export function validateBankMappings(mappings: ColumnMapping): { isValid: boolean; error?: string } {
  if (!mappings.noi_dung_giao_dich) {
    return { isValid: false, error: "Trường 'Nội dung giao dịch' là bắt buộc." };
  }
  if (!mappings.so_tien_thu && !mappings.so_tien_chi) {
    return { isValid: false, error: "Bạn phải cấu hình ít nhất một cột Số tiền thu (Có) hoặc Số tiền chi (Nợ)." };
  }

  const hasMainDate = mappings.ngay_giao_dich || mappings.ngay_hieu_luc || mappings.ngay_hach_toan;
  const hasSplitDate = mappings.cot_ngay && mappings.cot_thang;

  if (!hasMainDate && !hasSplitDate) {
    return { isValid: false, error: "Bạn phải cấu hình cột ngày giao dịch (hoặc hạch toán/hiệu lực) hoặc cấu hình các cột ngày và tháng riêng lẻ." };
  }

  return { isValid: true };
}

/**
 * Parses numeric strings with support for multiple formats: 
 * e.g., 1,000,000; 1.000.000; 1 000 000; 1,000,000.00; 1.000.000,00
 */
export function parseAmount(val: any): number {
  if (val === undefined || val === null) return 0;
  if (typeof val === "number") return val;

  let str = String(val).trim();
  if (!str || str === "-" || str === "—" || str === "–") return 0;

  let isNegative = false;
  if (str.startsWith("(") && str.endsWith(")")) {
    isNegative = true;
    str = str.substring(1, str.length - 1).trim();
  } else if (str.startsWith("-")) {
    isNegative = true;
    str = str.substring(1).trim();
  }

  // Remove currency marks
  str = str.replace(/[$,đ₫]|VND|EUR|USD/gi, "").trim();
  str = str.replace(/\s+/g, "");

  const lastDot = str.lastIndexOf(".");
  const lastComma = str.lastIndexOf(",");

  if (lastDot !== -1 && lastComma !== -1) {
    if (lastDot > lastComma) {
      str = str.replace(/,/g, "");
    } else {
      str = str.replace(/\./g, "").replace(/,/g, ".");
    }
  } else if (lastComma !== -1) {
    const commaCount = (str.match(/,/g) || []).length;
    if (commaCount > 1) {
      str = str.replace(/,/g, "");
    } else {
      const parts = str.split(",");
      if (parts[1].length === 3) {
        str = str.replace(/,/g, "");
      } else {
        str = str.replace(/,/g, ".");
      }
    }
  } else if (lastDot !== -1) {
    const dotCount = (str.match(/\./g) || []).length;
    if (dotCount > 1) {
      str = str.replace(/\./g, "");
    } else {
      const parts = str.split(".");
      if (parts[1].length === 3) {
        str = str.replace(/\./g, "");
      } else {
        // Keep as decimal
      }
    }
  }

  const num = parseFloat(str);
  if (isNaN(num)) return 0;
  return isNegative ? -num : num;
}

export interface MappingWarning {
  field: string;
  label: string;
  column: string;
  warning: string;
}

/**
 * Scan spreadsheet rows to double check data quality against the chosen mapping.
 */
export function checkMappingDataQuality(
  rows: any[],
  mappings: ColumnMapping
): MappingWarning[] {
  const warnings: MappingWarning[] = [];
  if (!rows || rows.length === 0) return warnings;

  const sample = rows.slice(0, 50);

  const isEntirelyEmpty = (col: string) => {
    return sample.every(r => {
      const val = r[col];
      return val === undefined || val === null || String(val).trim() === "";
    });
  };

  const isEntirelyInvalidNumber = (col: string) => {
    return sample.every(r => {
      const val = r[col];
      if (val === undefined || val === null || String(val).trim() === "") return true;
      const num = parseAmount(val);
      return num === 0;
    });
  };

  // 1. Content
  if (mappings.noi_dung_giao_dich && isEntirelyEmpty(mappings.noi_dung_giao_dich)) {
    warnings.push({
      field: "noi_dung_giao_dich",
      label: "Nội dung giao dịch",
      column: mappings.noi_dung_giao_dich,
      warning: "Cột được chọn không có dữ liệu hợp lệ (tất cả dòng đều trống)."
    });
  }

  // 2. Amount Fields
  if (mappings.so_tien_thu && isEntirelyInvalidNumber(mappings.so_tien_thu) && mappings.so_tien_chi && isEntirelyInvalidNumber(mappings.so_tien_chi)) {
    warnings.push({
      field: "so_tien_thu_chi",
      label: "Số tiền thu / chi",
      column: `${mappings.so_tien_thu} / ${mappings.so_tien_chi}`,
      warning: "Cột tiền thu và tiền chi được chọn không có số liệu hợp lệ hoặc toàn bộ bằng 0."
    });
  } else {
    if (mappings.so_tien_thu && isEntirelyInvalidNumber(mappings.so_tien_thu)) {
      if (!mappings.so_tien_chi) {
        warnings.push({
          field: "so_tien_thu",
          label: "Số tiền thu",
          column: mappings.so_tien_thu,
          warning: "Cột được chọn không có dữ liệu số tiền hợp lệ."
        });
      }
    }
    if (mappings.so_tien_chi && isEntirelyInvalidNumber(mappings.so_tien_chi)) {
      if (!mappings.so_tien_thu) {
        warnings.push({
          field: "so_tien_chi",
          label: "Số tiền chi",
          column: mappings.so_tien_chi,
          warning: "Cột được chọn không có dữ liệu số tiền hợp lệ."
        });
      }
    }
  }

  return warnings;
}
