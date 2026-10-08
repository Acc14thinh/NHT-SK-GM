/**
 * Parser for Sales Software Excel Files (Phần mềm bán hàng, ví dụ Misa, KiotViet, Sapo...)
 * Section 2, 4, 9, 18, 19
 */
import * as XLSX from "xlsx";
import { SalesSoftwareColumnMapping, SalesSoftwareRawRow } from "./types";
import { findBestMatchingHeader } from "./shopeeParser";
import { normalizeOrderId } from "./orderNormalizer";
import { getSavedHeaderMappings } from "./shopMapper";

export interface ParsedSalesSoftwareResult {
  fileName: string;
  sheetName: string;
  headers: string[];
  headerRowIndex: number;
  detectedMapping: SalesSoftwareColumnMapping;
  rows: SalesSoftwareRawRow[];
  detectedShops: string[];
  diagnostics: {
    totalRows: number;
    rowsWithShop: number;
    rowsMissingShop: number;
    rowsWithOrderId: number;
    rowsMissingOrderId: number;
    uniqueOrderCount: number;
    duplicateOrderCount: number;
    shopDistribution: { [shopName: string]: number };
  };
  warnings: string[];
}

/**
 * Tự động nhận diện cột cho file Phần Mềm Bán Hàng (Ưu tiên mapping đã lưu trước đó)
 */
export function autoDetectSalesSoftwareColumns(headers: string[]): SalesSoftwareColumnMapping {
  const saved = getSavedHeaderMappings("sales");

  const getCol = (savedKey: string, keywords: string[]): string => {
    if (saved[savedKey] && headers.includes(saved[savedKey])) {
      return saved[savedKey];
    }
    return findBestMatchingHeader(headers, keywords);
  };

  return {
    shopNameCol: getCol("shopNameCol", [
      "tên shop", "ten shop", "cửa hàng", "tên cửa hàng", "gian hàng", "chi nhánh", "shop", "store"
    ]),
    orderIdCol: getCol("orderIdCol", [
      "số đơn hàng từ hệ thống khác", "so don hang tu he thong khac",
      "mã đơn hàng từ hệ thống khác", "mã đơn hàng", "ma don hang", "mã đơn sàn", "mã đơn", "order id", "số đơn hàng"
    ]),
    invoiceNumberCol: getCol("invoiceNumberCol", [
      "số hóa đơn", "so hoa don", "số hđ", "so hd", "hóa đơn số", "invoice no"
    ]),
    invoiceDateCol: getCol("invoiceDateCol", [
      "ngày hóa đơn", "ngay hoa don", "ngày hđ", "ngay hd", "ngày chứng từ", "invoice date"
    ]),
    totalAmountCol: getCol("totalAmountCol", [
      "tổng tiền", "tong tien", "tổng cộng", "thành tiền sau thuế", "thành tiền",
      "thanh tien", "tổng giá trị thanh toán", "tổng thanh toán", "total amount"
    ]),
    voucherNumberCol: getCol("voucherNumberCol", [
      "số chứng từ", "so chung tu", "số ct", "so ct", "mã chứng từ", "voucher no"
    ]),
    customerCol: getCol("customerCol", [
      "tên khách hàng", "khách hàng", "người mua", "customer"
    ]),
    noteCol: getCol("noteCol", [
      "ghi chú", "diễn giải", "note", "nội dung"
    ])
  };
}

/**
 * Phân tích workbook phần mềm bán hàng
 */
export function parseSalesSoftwareWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
  preferredSheet?: string,
  customMapping?: Partial<SalesSoftwareColumnMapping>
): ParsedSalesSoftwareResult {
  const sheetName = preferredSheet && workbook.Sheets[preferredSheet]
    ? preferredSheet
    : workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];
  const warnings: string[] = [];

  if (!worksheet) {
    return {
      fileName,
      sheetName: "",
      headers: [],
      headerRowIndex: 0,
      detectedMapping: {
        shopNameCol: "",
        orderIdCol: "",
        invoiceNumberCol: "",
        invoiceDateCol: "",
        totalAmountCol: ""
      },
      rows: [],
      detectedShops: [],
      diagnostics: {
        totalRows: 0,
        rowsWithShop: 0,
        rowsMissingShop: 0,
        rowsWithOrderId: 0,
        rowsMissingOrderId: 0,
        uniqueOrderCount: 0,
        duplicateOrderCount: 0,
        shopDistribution: {}
      },
      warnings: ["Tệp không chứa sheet dữ liệu hợp lệ."]
    };
  }

  // Đọc raw 2D array
  const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false // ép chuỗi tránh scientific notation
  });

  // Tìm header row
  let headerRowIdx = -1;
  for (let r = 0; r < Math.min(10, rawData.length); r++) {
    const rowStr = (rawData[r] || []).map(cell => String(cell || "").toLowerCase()).join(" ");
    if (
      (rowStr.includes("tên shop") || rowStr.includes("cửa hàng")) &&
      (rowStr.includes("số đơn hàng") || rowStr.includes("mã đơn") || rowStr.includes("hệ thống khác"))
    ) {
      headerRowIdx = r;
      break;
    }
  }

  if (headerRowIdx === -1) {
    headerRowIdx = 0;
  }

  const rawHeaders: string[] = (rawData[headerRowIdx] || []).map((h: any, idx: number) => {
    const s = String(h || "").trim();
    return s || `Cột_${idx + 1}`;
  });

  const detected = autoDetectSalesSoftwareColumns(rawHeaders);
  const finalMapping: SalesSoftwareColumnMapping = {
    ...detected,
    ...customMapping
  };

  if (!finalMapping.orderIdCol) {
    warnings.push("Không tự động nhận diện được cột [Số đơn hàng từ hệ thống khác / Mã đơn]. Vui lòng chọn thủ công.");
  }
  if (!finalMapping.shopNameCol) {
    warnings.push("Không tự động nhận diện được cột [Tên Shop]. Vui lòng chọn thủ công.");
  }

  // Đọc rows
  const rows: SalesSoftwareRawRow[] = [];
  const detectedShopsSet = new Set<string>();
  const shopDistribution: { [shopName: string]: number } = {};
  const seenOrders = new Set<string>();
  let rowsWithShop = 0;
  let rowsMissingShop = 0;
  let rowsWithOrderId = 0;
  let rowsMissingOrderId = 0;
  let duplicateOrderCount = 0;

  for (let r = headerRowIdx + 1; r < rawData.length; r++) {
    const rowData = rawData[r];
    if (!rowData || rowData.length === 0) continue;

    const hasContent = rowData.some(cell => cell !== null && cell !== undefined && String(cell).trim() !== "");
    if (!hasContent) continue;

    const rowObj: SalesSoftwareRawRow = {};
    rawHeaders.forEach((h, colIdx) => {
      rowObj[h] = rowData[colIdx] !== undefined ? rowData[colIdx] : "";
    });

    // Check shop name
    const rawShop = finalMapping.shopNameCol ? String(rowObj[finalMapping.shopNameCol] || "").trim() : "";
    if (rawShop) {
      rowsWithShop++;
      detectedShopsSet.add(rawShop);
      shopDistribution[rawShop] = (shopDistribution[rawShop] || 0) + 1;
    } else {
      rowsMissingShop++;
    }

    // Check order id
    const rawOrderId = finalMapping.orderIdCol ? rowObj[finalMapping.orderIdCol] : "";
    const { normalized } = normalizeOrderId(rawOrderId);

    if (normalized) {
      rowsWithOrderId++;
      const compositeKey = `${rawShop || "UNASSIGNED"}::${normalized}`;
      if (seenOrders.has(compositeKey)) {
        duplicateOrderCount++;
      } else {
        seenOrders.add(compositeKey);
      }
    } else {
      rowsMissingOrderId++;
    }

    rows.push(rowObj);
  }

  return {
    fileName,
    sheetName,
    headers: rawHeaders,
    headerRowIndex: headerRowIdx,
    detectedMapping: finalMapping,
    rows,
    detectedShops: Array.from(detectedShopsSet).sort(),
    diagnostics: {
      totalRows: rows.length,
      rowsWithShop,
      rowsMissingShop,
      rowsWithOrderId,
      rowsMissingOrderId,
      uniqueOrderCount: seenOrders.size,
      duplicateOrderCount,
      shopDistribution
    },
    warnings
  };
}
