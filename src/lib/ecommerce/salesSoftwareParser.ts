/**
 * Parser for Sales Software Excel Files (Phần mềm bán hàng, ví dụ Misa, KiotViet, Sapo...)
 * Section 2, 4, 9, 18, 19
 */
import * as XLSX from "xlsx";
import { SalesSoftwareColumnMapping, SalesSoftwareRawRow } from "./types";
import { findBestMatchingHeader, findStrictMatchingHeader } from "./shopeeParser";
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
 * Kiểm tra xem cột có phải được nhận diện chính xác theo chuẩn Phần Mềm Bán Hàng không (Section B)
 */
export function isExactSalesHeader(field: keyof SalesSoftwareColumnMapping, header: string): boolean {
  if (!header) return false;
  const hTrim = header.trim();
  const hLower = hTrim.toLowerCase();

  switch (field) {
    case "shopNameCol":
      return hLower === "tên shop" || hLower === "ten shop" || hLower === "cửa hàng" || hLower === "tên cửa hàng";
    case "orderIdCol":
      return hLower === "số đơn hàng từ sàn" || hLower === "so don hang tu san" ||
             hLower === "số đơn hàng từ hệ thống khác" || hLower === "so don hang tu he thong khac" ||
             hLower === "mã đơn hàng từ hệ thống khác" || hLower === "mã đơn sàn";
    case "invoiceNumberCol":
      return hLower === "số hóa đơn" || hLower === "số hoá đơn" || hLower === "số chứng từ" ||
             hLower === "so hoa don" || hLower === "so chung tu";
    case "invoiceDateCol":
      return hLower === "ngày hạch toán" || hLower === "ngay hach toan" ||
             hLower === "ngày hóa đơn" || hLower === "ngày hoá đơn" || hLower === "ngày chứng từ" ||
             hLower === "ngay hoa don" || hLower === "ngay chung tu";
    case "totalAmountCol":
      return hLower === "tổng tiền thanh toán" || hLower === "tong tien thanh toan" ||
             hLower === "tổng tiền" || hLower === "tong tien" || hLower === "thành tiền sau thuế" || hLower === "thành tiền";
    case "voucherNumberCol":
      return hLower === "số chứng từ" || hLower === "so chung tu" || hLower === "số hóa đơn" || hLower === "so hoa don";
    case "invoiceDiscountCol":
      return hLower === "giảm giá hóa đơn" || hLower === "giam gia hoa don" || hLower === "voucher" || hLower === "chiết khấu";
    case "marketplaceCol":
      return hLower === "sàn tmđt" || hLower === "san tmdt" || hLower === "kênh bán hàng" || hLower === "kenh ban hang" || hLower === "kênh bán";
    case "cancelledCol":
      return hLower === "hủy" || hLower === "huy" || hLower === "trạng thái hủy" || hLower === "trang thai huy" || hLower === "đã hủy";
    default:
      return false;
  }
}

/**
 * Tự động nhận diện cột cho file Phần Mềm Bán Hàng theo đúng logic đối chiếu đã kiểm chứng (Section B & Section 5)
 */
export function autoDetectSalesSoftwareColumns(headers: string[]): SalesSoftwareColumnMapping {
  const saved = getSavedHeaderMappings("sales");

  const getCol = (savedKey: string, exactKeywords: string[], fallbackKeywords: string[] = []): string => {
    if (saved[savedKey] && headers.includes(saved[savedKey])) {
      return saved[savedKey];
    }
    return findStrictMatchingHeader(headers, exactKeywords, fallbackKeywords);
  };

  return {
    // KHÓA ĐỐI CHIẾU: Tên Shop (*)
    shopNameCol: getCol("shopNameCol", [
      "Tên shop", "Tên Shop", "Tên cửa hàng"
    ], [
      "ten shop", "cửa hàng", "gian hàng", "chi nhánh", "shop", "store"
    ]),

    // KHÓA ĐỐI CHIẾU: Mã đơn từ sàn (*) - Ưu tiên "Số đơn hàng từ hệ thống khác"
    orderIdCol: getCol("orderIdCol", [
      "Số đơn hàng từ hệ thống khác", "Số đơn hàng từ sàn", "Mã đơn hàng từ hệ thống khác", "Mã đơn từ sàn"
    ], [
      "so don hang tu he thong khac", "so don hang tu san", "mã đơn hàng", "ma don hang", "mã đơn sàn", "mã đơn", "order id", "số đơn hàng"
    ]),

    // THÔNG TIN HÓA ĐƠN: Số hóa đơn (*)
    invoiceNumberCol: getCol("invoiceNumberCol", [
      "Số hóa đơn", "Số chứng từ", "Số hoá đơn"
    ], [
      "so hoa don", "so chung tu", "số hđ", "so hd", "hóa đơn số", "invoice no"
    ]),

    // THÔNG TIN HÓA ĐƠN: Ngày hạch toán (*) - Ưu tiên "Ngày hạch toán"
    invoiceDateCol: getCol("invoiceDateCol", [
      "Ngày hạch toán", "Ngày hóa đơn", "Ngày chứng từ", "Ngày hoá đơn"
    ], [
      "ngay hach toan", "ngay hoa don", "ngay chung tu", "ngày hđ", "ngay hd", "invoice date"
    ]),

    // TÍNH DOANH THU: Tổng tiền thanh toán (*) - Ưu tiên "Tổng tiền thanh toán"
    totalAmountCol: getCol("totalAmountCol", [
      "Tổng tiền thanh toán", "Tổng tiền", "Thành tiền sau thuế", "Thành tiền"
    ], [
      "tong tien thanh toan", "tong tien", "tổng cộng", "thanh tien", "tổng giá trị thanh toán", "tổng thanh toán", "total amount"
    ]),

    // TÍNH DOANH THU: Giảm giá hóa đơn / Voucher
    invoiceDiscountCol: getCol("invoiceDiscountCol", [
      "Giảm giá hóa đơn", "Voucher", "Chiết khấu"
    ], [
      "giam gia hoa don", "voucher", "chiet khau", "tiền voucher", "tiền giảm giá"
    ]),

    // VALIDATION & CANCELLED: Kênh bán hàng (*)
    marketplaceCol: getCol("marketplaceCol", [
      "Sàn TMĐT", "Kênh bán hàng", "Kênh bán"
    ], [
      "san tmdt", "kenh ban hang", "kênh bán", "sàn giao dịch", "marketplace", "channel"
    ]),

    // VALIDATION & CANCELLED: Trạng thái hủy
    cancelledCol: getCol("cancelledCol", [
      "Hủy", "Trạng thái hủy", "Đã hủy"
    ], [
      "huy", "trang thai huy", "da huy", "is cancelled", "cancelled"
    ]),

    voucherNumberCol: getCol("voucherNumberCol", [
      "Số chứng từ", "so chung tu", "mã chứng từ"
    ], [
      "số ct", "so ct", "voucher no"
    ]),
    customerCol: getCol("customerCol", [
      "Tên khách hàng", "Khách hàng"
    ], [
      "người mua", "customer"
    ]),
    noteCol: getCol("noteCol", [
      "Ghi chú", "Diễn giải"
    ], [
      "note", "nội dung"
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
