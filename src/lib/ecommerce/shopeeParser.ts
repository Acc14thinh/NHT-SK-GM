/**
 * Parser for Marketplace Excel Files (Shopee, with extensibility for TikTok, Lazada)
 * Section 2, 4, 18, 19
 */
import * as XLSX from "xlsx";
import { MarketplaceColumnMapping, MarketplaceRawRow } from "./types";
import { normalizeOrderId } from "./orderNormalizer";
import { getSavedHeaderMappings } from "./shopMapper";

export interface ParsedMarketplaceSheet {
  sheetName: string;
  rowCount: number;
  headers: string[];
  headerRowIndex: number;
  detectedMapping: MarketplaceColumnMapping;
  rows: MarketplaceRawRow[];
  diagnostics: {
    totalRows: number;
    rowsWithOrderId: number;
    rowsMissingOrderId: number;
    uniqueOrderCount: number;
    duplicateOrderCount: number;
  };
}

export interface ParsedMarketplaceResult {
  fileName: string;
  sheets: ParsedMarketplaceSheet[];
  allDetectedShopNames: string[];
  totalRows: number;
  totalUniqueOrders: number;
  warnings: string[];
}

/**
 * Tìm kiếm header khớp nhất theo nguyên tắc:
 * 1. Exact case-sensitive match (ưu tiên cao nhất)
 * 2. Case-insensitive exact match (nếu chỉ có duy nhất 1 cột)
 * 3. Không tự ý fuzzy match nếu có nhiều cột có tên tương tự sau chuẩn hóa (đặc biệt: "Tổng số tiền Người mua thanh toán" vs "Tổng số tiền người mua thanh toán")
 */
export function findStrictMatchingHeader(
  headers: string[],
  exactCaseKeywords: string[],
  fallbackKeywords: string[] = []
): string {
  if (!headers || headers.length === 0) return "";

  // 1. Kiểm tra Exact match có phân biệt HOA/thường (Case-sensitive)
  for (const target of exactCaseKeywords) {
    if (headers.includes(target)) {
      return target;
    }
  }

  // 2. Kiểm tra Exact match không phân biệt hoa thường (Case-insensitive)
  // CHÚ Ý: Nếu có >= 2 cột trong headers cùng khớp khi toLowerCase(), KHÔNG tự chọn bừa!
  for (const target of exactCaseKeywords.concat(fallbackKeywords)) {
    const normTarget = target.toLowerCase().trim();
    const candidateMatches = headers.filter(h => String(h || "").toLowerCase().trim() === normTarget);

    if (candidateMatches.length === 1) {
      return candidateMatches[0];
    } else if (candidateMatches.length > 1) {
      // Có nhiều cột bị trùng sau normalize -> Tìm cột có case-match gần nhất
      const exactCase = candidateMatches.find(h => h.trim() === target.trim());
      if (exactCase) return exactCase;
      // Nếu không có cột nào khớp case chuẩn, để người dùng tự chọn qua dropdown để tránh nhầm lẫn
      return "";
    }
  }

  // 3. Fallback: Contains match nếu danh sách dự phòng có
  for (const kw of fallbackKeywords) {
    const normKw = kw.toLowerCase().trim();
    const matches = headers.filter(h => String(h || "").toLowerCase().includes(normKw));
    if (matches.length === 1) {
      return matches[0];
    }
  }

  return "";
}

/**
 * Backward compatibility alias
 */
export function findBestMatchingHeader(headers: string[], keywords: string[]): string {
  return findStrictMatchingHeader(headers, keywords, []);
}

/**
 * Kiểm tra xem cột có phải được nhận diện chính xác theo chuẩn Shopee không
 */
export function isExactShopeeHeader(field: keyof MarketplaceColumnMapping, header: string): boolean {
  if (!header) return false;
  const hTrim = header.trim();
  const hLower = hTrim.toLowerCase();

  switch (field) {
    case "orderIdCol":
      return hLower === "mã đơn hàng" || hLower === "ma don hang" || hLower === "mã đơn";
    case "orderStatusCol":
      return hLower === "trạng thái đơn hàng" || hLower === "trang thai don hang";
    case "deliveryTimeCol":
      return hLower === "thời gian giao hàng" || hLower === "thoi gian giao hang";
    case "returnRefundCol":
      return hLower === "trạng thái trả hàng/hoàn tiền" || hLower === "trang thai tra hang/hoan tien";
    case "buyerPaidCol":
      // Cực kỳ quan trọng: CHÍNH XÁC "Tổng số tiền Người mua thanh toán" (Chữ N hoa)
      return hTrim === "Tổng số tiền Người mua thanh toán";
    case "shopVoucherCol":
      return hLower === "mã giảm giá của shop" || hLower === "ma giam gia cua shop" || hLower === "voucher của shop";
    default:
      return false;
  }
}

/**
 * Tự động nhận diện cấu hình cột cho file Sàn Shopee theo đúng logic đối chiếu đã kiểm chứng (Section A)
 */
export function autoDetectShopeeColumns(headers: string[]): MarketplaceColumnMapping {
  const saved = getSavedHeaderMappings("marketplace");

  const getCol = (savedKey: string, exactKeywords: string[], fallbackKeywords: string[] = []): string => {
    if (saved[savedKey] && headers.includes(saved[savedKey])) {
      return saved[savedKey];
    }
    return findStrictMatchingHeader(headers, exactKeywords, fallbackKeywords);
  };

  return {
    // KHÓA ĐỐI CHIẾU: Mã đơn hàng (*)
    orderIdCol: getCol("orderIdCol", [
      "Mã đơn hàng", "Mã đơn"
    ], [
      "ma don hang", "order id", "order_sn", "ordersn", "số đơn hàng"
    ]),

    // TÍNH DOANH THU: Tổng số tiền Người mua thanh toán (*)
    // CHÍNH XÁC tới header: "Tổng số tiền Người mua thanh toán" (chữ N hoa). KHÔNG tự map sang "Tổng số tiền người mua thanh toán"
    buyerPaidCol: getCol("buyerPaidCol", [
      "Tổng số tiền Người mua thanh toán"
    ], []),

    // TÍNH DOANH THU: Mã giảm giá của Shop (*)
    shopVoucherCol: getCol("shopVoucherCol", [
      "Mã giảm giá của Shop", "Mã giảm giá của shop"
    ], [
      "ma giam gia cua shop", "voucher của shop", "voucher shop", "shop voucher"
    ]),

    // THÔNG TIN TRẠNG THÁI: Trạng thái đơn hàng (*)
    orderStatusCol: getCol("orderStatusCol", [
      "Trạng Thái Đơn Hàng", "Trạng thái đơn hàng"
    ], [
      "trang thai don hang", "order status"
    ]),

    // THÔNG TIN TRẠNG THÁI: Thời gian giao hàng
    deliveryTimeCol: getCol("deliveryTimeCol", [
      "Thời gian giao hàng"
    ], [
      "thoi gian giao hang", "thời gian hoàn thành", "ngày giao hàng", "delivery time"
    ]),

    // THÔNG TIN TRẠNG THÁI: Trạng thái Trả hàng/Hoàn tiền (*)
    returnRefundCol: getCol("returnRefundCol", [
      "Trạng thái Trả hàng/Hoàn tiền", "Trạng thái trả hàng/hoàn tiền"
    ], [
      "trang thai tra hang/hoan tien", "trả hàng/hoàn tiền", "return / refund status", "hoàn tiền"
    ]),

    skuCol: getCol("skuCol", [
      "Tên sản phẩm", "SKU"
    ], [
      "ten san pham", "mã sản phẩm", "product name"
    ])
  };
}

/**
 * Phân tích workbook của sàn TMĐT
 */
export function parseMarketplaceWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
  customMapping?: Partial<MarketplaceColumnMapping>
): ParsedMarketplaceResult {
  const sheets: ParsedMarketplaceSheet[] = [];
  const warnings: string[] = [];
  let totalRowsAcrossSheets = 0;
  const allUniqueOrdersSet = new Set<string>();

  for (const sheetName of workbook.SheetNames) {
    // Bỏ qua các sheet ẩn hoặc sheet metadata không liên quan
    if (sheetName.toLowerCase().startsWith("__") || sheetName.toLowerCase().includes("readme")) {
      continue;
    }

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    // Đọc dạng mảng các dòng (raw 2D array)
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false // ép chuỗi để không bị scientific notation
    });

    if (!rawData || rawData.length === 0) continue;

    // Tìm dòng header (dòng có chứa "Mã đơn hàng" hoặc điểm số header cao nhất)
    let headerRowIdx = -1;
    for (let r = 0; r < Math.min(10, rawData.length); r++) {
      const rowStr = (rawData[r] || []).map(cell => String(cell || "").toLowerCase()).join(" ");
      if (
        rowStr.includes("mã đơn hàng") ||
        rowStr.includes("ma don hang") ||
        rowStr.includes("mã đơn") ||
        rowStr.includes("order id") ||
        rowStr.includes("order_sn") ||
        rowStr.includes("ordersn") ||
        rowStr.includes("người mua thanh toán") ||
        rowStr.includes("buyer paid") ||
        rowStr.includes("tổng số tiền")
      ) {
        headerRowIdx = r;
        break;
      }
    }

    if (headerRowIdx === -1) {
      // Fallback lấy dòng 0 nếu có dữ liệu
      headerRowIdx = 0;
    }

    const rawHeaders: string[] = (rawData[headerRowIdx] || []).map((h: any, idx: number) => {
      const s = String(h || "").trim();
      return s || `Cột_${idx + 1}`;
    });

    // Detect mapping
    const detected = autoDetectShopeeColumns(rawHeaders);
    const finalMapping: MarketplaceColumnMapping = {
      ...detected,
      ...customMapping
    };

    if (!finalMapping.orderIdCol) {
      warnings.push(`Sheet "${sheetName}": Không tự động nhận diện được cột Mã đơn hàng.`);
    }

    // Đọc các dòng dữ liệu sau header
    const rows: MarketplaceRawRow[] = [];
    const seenOrdersInSheet = new Set<string>();
    let rowsWithOrderId = 0;
    let rowsMissingOrderId = 0;
    let duplicateOrdersInSheet = 0;

    for (let r = headerRowIdx + 1; r < rawData.length; r++) {
      const rowData = rawData[r];
      if (!rowData || rowData.length === 0) continue;

      // Kiểm tra dòng có toàn chuỗi rỗng không
      const hasContent = rowData.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "");
      if (!hasContent) continue;

      const rowObj: MarketplaceRawRow = {};
      rawHeaders.forEach((h, colIdx) => {
        rowObj[h] = rowData[colIdx] !== undefined ? rowData[colIdx] : "";
      });

      const orderIdVal = finalMapping.orderIdCol ? rowObj[finalMapping.orderIdCol] : "";
      const { normalized } = normalizeOrderId(orderIdVal);

      if (normalized) {
        rowsWithOrderId++;
        if (seenOrdersInSheet.has(normalized)) {
          duplicateOrdersInSheet++;
        } else {
          seenOrdersInSheet.add(normalized);
        }
        allUniqueOrdersSet.add(`${sheetName}::${normalized}`);
      } else {
        rowsMissingOrderId++;
      }

      rows.push(rowObj);
    }

    if (workbook.SheetNames.length > 1 && rows.length === 0) {
      continue;
    }

    totalRowsAcrossSheets += rows.length;

    sheets.push({
      sheetName,
      rowCount: rows.length,
      headers: rawHeaders,
      headerRowIndex: headerRowIdx,
      detectedMapping: finalMapping,
      rows,
      diagnostics: {
        totalRows: rows.length,
        rowsWithOrderId,
        rowsMissingOrderId,
        uniqueOrderCount: seenOrdersInSheet.size,
        duplicateOrderCount: duplicateOrdersInSheet
      }
    });
  }

  const allDetectedShopNames = sheets.map(s => s.sheetName);

  return {
    fileName,
    sheets,
    allDetectedShopNames,
    totalRows: totalRowsAcrossSheets,
    totalUniqueOrders: allUniqueOrdersSet.size,
    warnings
  };
}
