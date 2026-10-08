/**
 * Normalization utilities for E-Commerce Order IDs, Shop Names, and Currency
 * Section 6, 20 & 21
 */

/**
 * Chuẩn hóa Mã đơn hàng:
 * - Luôn giữ dạng chuỗi
 * - Trim khoảng trắng đầu/cuối
 * - Thay thế non-breaking space (\u00A0) và zero-width space
 * - Không chuyển thành số
 * - Không làm mất số 0 đầu
 * - Không chuyển scientific notation
 */
export function normalizeOrderId(val: any): { original: string; normalized: string } {
  if (val === null || val === undefined) {
    return { original: "", normalized: "" };
  }

  let str = String(val);

  // Xử lý các dạng scientific notation nếu lỡ bị chuyển từ floating point Excel
  if (/^\s*[+-]?\d+(\.\d+)?[eE][+-]?\d+\s*$/.test(str)) {
    try {
      // BigInt hoặc format cẩn thận nếu là số nguyên lớn
      const num = Number(str);
      if (!isNaN(num) && Number.isSafeInteger(num)) {
        str = BigInt(Math.round(num)).toString();
      }
    } catch {
      // fallback giữ nguyên str
    }
  }

  const original = str.trim();

  // Chuẩn hóa nội bộ: loại bỏ non-breaking space, zero-width chars, tab, xuống dòng
  const normalized = original
    .replace(/[\u00A0\u1680\u2000-\u200B\u202F\u205F\u3000\uFEFF]/g, " ") // space đặc biệt
    .replace(/[\r\n\t]/g, " ")
    .trim()
    .replace(/\s+/g, " ") // gom nhiều space thành 1 space
    .toUpperCase(); // uppercase để tránh lệch hoa/thường không đáng có ở mã đơn

  return { original, normalized };
}

/**
 * Chuẩn hóa Tên Shop để hỗ trợ đối chiếu và gợi ý ghép Shop:
 * Loại bỏ dấu tiếng Việt, ký tự đặc biệt, chuyển chữ thường để so khớp logic
 */
export function normalizeShopName(name: string): string {
  if (!name) return "";
  return String(name)
    .trim()
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/[^a-z0-9]/g, "")
    .trim();
}

/**
 * Chuẩn hóa số tiền VND:
 * - Xử lý chuỗi có dấu phẩy, dấu chấm phân cách hàng nghìn
 * - Chuyển sang number an toàn
 * - Làm tròn đến đơn vị đồng
 */
export function parseCurrencyVnd(val: any): number {
  if (val === null || val === undefined) return 0;
  if (typeof val === "number") {
    if (isNaN(val) || !isFinite(val)) return 0;
    return Math.round(val);
  }

  let str = String(val).trim();
  if (!str) return 0;

  // Loại bỏ ký hiệu tiền tệ và khoảng trắng
  str = str.replace(/[₫đVNDvnd\s]/g, "");

  // Kiểm tra dấu phân cách thập phân / hàng nghìn
  // Trường hợp format VN: 100.000,00 hoặc 100.000
  // Trường hợp format US: 100,000.00 hoặc 100,000
  if (str.includes(".") && str.includes(",")) {
    if (str.lastIndexOf(",") > str.lastIndexOf(".")) {
      // Dấu phẩy là thập phân: 100.000,00 -> 100000.00
      str = str.replace(/\./g, "").replace(",", ".");
    } else {
      // Dấu chấm là thập phân: 100,000.00 -> 100000.00
      str = str.replace(/,/g, "");
    }
  } else if (str.includes(".")) {
    // Chỉ có dấu chấm: có thể là 100.000 (VND nghìn) hoặc 100.5 (thập phân)
    const parts = str.split(".");
    if (parts.length > 2) {
      // 100.000.000 -> dấu phân cách nghìn
      str = str.replace(/\./g, "");
    } else if (parts[1] && parts[1].length === 3) {
      // 100.000 -> dấu phân cách nghìn trong kế toán VN
      str = str.replace(/\./g, "");
    }
  } else if (str.includes(",")) {
    // Chỉ có dấu phẩy
    const parts = str.split(",");
    if (parts.length > 2) {
      str = str.replace(/,/g, "");
    } else if (parts[1] && parts[1].length === 3) {
      str = str.replace(/,/g, "");
    } else {
      // 100,5 -> thập phân
      str = str.replace(",", ".");
    }
  }

  const num = parseFloat(str);
  return isNaN(num) || !isFinite(num) ? 0 : Math.round(num);
}

/**
 * So sánh 2 số tiền với mức sai số (tolerance)
 */
export function areAmountsEqual(amt1: number, amt2: number, toleranceVnd: number = 1): boolean {
  return Math.abs(amt1 - amt2) <= toleranceVnd;
}

/**
 * Làm sạch tên Sheet Excel:
 * Excel quy định: tối đa 31 ký tự, không chứa \ / ? * [ ] :
 */
export function sanitizeExcelSheetName(name: string, fallback: string = "Shop"): string {
  if (!name || !name.trim()) return fallback.substring(0, 31);
  let clean = name.replace(/[\\/?*\[\]:]/g, "_").trim();
  if (clean.length > 31) {
    clean = clean.substring(0, 31).trim();
  }
  return clean || fallback.substring(0, 31);
}
