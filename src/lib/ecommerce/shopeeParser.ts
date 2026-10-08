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
 * Tìm kiếm header khớp nhất theo danh sách từ khóa
 */
export function findBestMatchingHeader(headers: string[], keywords: string[]): string {
  const normKeywords = keywords.map(k => k.toLowerCase().trim());

  // 1. Exact match
  for (const h of headers) {
    const normH = String(h || "").toLowerCase().trim();
    if (normKeywords.includes(normH)) return h;
  }

  // 2. Contains match
  for (const kw of normKeywords) {
    for (const h of headers) {
      const normH = String(h || "").toLowerCase().trim();
      if (normH.includes(kw)) return h;
    }
  }

  return "";
}

/**
 * Tự động nhận diện cấu hình cột cho file Sàn Shopee (Ưu tiên mapping đã lưu trước đó)
 */
export function autoDetectShopeeColumns(headers: string[]): MarketplaceColumnMapping {
  const saved = getSavedHeaderMappings("marketplace");

  const getCol = (savedKey: string, keywords: string[]): string => {
    if (saved[savedKey] && headers.includes(saved[savedKey])) {
      return saved[savedKey];
    }
    return findBestMatchingHeader(headers, keywords);
  };

  return {
    orderIdCol: getCol("orderIdCol", [
      "mã đơn hàng", "ma don hang", "order id", "mã đơn", "order_sn", "ordersn", "số đơn hàng"
    ]),
    buyerPaidCol: getCol("buyerPaidCol", [
      "tổng số tiền người mua thanh toán", "tong so tien nguoi mua thanh toan",
      "người mua thanh toán", "buyer total", "buyer paid amount", "tiền người mua trả", "tổng giá trị đơn hàng"
    ]),
    shopVoucherCol: getCol("shopVoucherCol", [
      "mã giảm giá của shop", "ma giam gia cua shop",
      "voucher của shop", "voucher shop", "shop voucher", "giảm giá của shop"
    ]),
    orderStatusCol: getCol("orderStatusCol", [
      "trạng thái đơn hàng", "trang thai don hang", "trạng thái đơn", "order status"
    ]),
    deliveryTimeCol: getCol("deliveryTimeCol", [
      "thời gian giao hàng", "thoi gian giao hang",
      "thời gian hoàn thành", "ngày giao hàng", "delivery time"
    ]),
    returnRefundCol: getCol("returnRefundCol", [
      "trạng thái trả hàng/hoàn tiền", "trang thai tra hang/hoan tien",
      "trả hàng/hoàn tiền", "return / refund status", "hoàn tiền"
    ]),
    skuCol: getCol("skuCol", [
      "tên sản phẩm", "ten san pham", "sku", "mã sản phẩm", "product name"
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
        rowStr.includes("order id") ||
        rowStr.includes("người mua thanh toán")
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
