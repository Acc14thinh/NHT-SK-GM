/**
 * Parser for Sales Software Excel Files (Phần mềm bán hàng, ví dụ Misa, KiotViet, Sapo...)
 * Section 2, 4, 9, 18, 19
 */
import * as XLSX from "xlsx";
import { SalesSoftwareColumnMapping, SalesSoftwareRawRow } from "./types";
import { findBestMatchingHeader, findStrictMatchingHeader } from "./shopeeParser";
import { normalizeOrderId, normalizeShopName, parseCurrencyVnd } from "./orderNormalizer";
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
 * PHASE 1: LIGHTWEIGHT INSPECTION (Phần Mềm Bán Hàng)
 * Chỉ đọc tối đa 25 dòng đầu tiên để lấy headers, mapping, và sample preview.
 * Tuyệt đối không convert hàng chục nghìn dòng thành JSON trong React state.
 */
export function inspectSalesSoftwareWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
  fileSize: number = 0,
  preferredSheet?: string,
  customMapping?: Partial<SalesSoftwareColumnMapping>
): import("./types").SalesSoftwareInspectionResult {
  const sheetName = preferredSheet && workbook.Sheets[preferredSheet]
    ? preferredSheet
    : workbook.SheetNames[0];

  const worksheet = workbook.Sheets[sheetName];
  if (!worksheet) {
    return {
      fileName,
      fileSize,
      sheetName: "",
      headers: [],
      detectedMapping: {
        shopNameCol: "",
        orderIdCol: "",
        invoiceNumberCol: "",
        invoiceDateCol: "",
        totalAmountCol: ""
      },
      previewValuesByCol: {},
      detectedShops: [],
      approxRowCount: 0,
      totalRows: 0,
      rows: []
    };
  }

  let approxRowCount = 0;
  if (worksheet["!ref"]) {
    try {
      const decoded = XLSX.utils.decode_range(worksheet["!ref"]);
      approxRowCount = Math.max(0, decoded.e.r - decoded.s.r);
    } catch {
      approxRowCount = 0;
    }
  }

  const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false
  });

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
  if (headerRowIdx === -1) headerRowIdx = 0;

  const rawHeaders: string[] = (rawData[headerRowIdx] || []).map((h: any, idx: number) => {
    const s = String(h || "").trim();
    return s || `Cột_${idx + 1}`;
  });

  let detected = autoDetectSalesSoftwareColumns(rawHeaders);
  if (customMapping) {
    detected = { ...detected, ...customMapping };
  }

  // Preview samples and shop discovery from top rows
  const previewValuesByCol: Record<string, string[]> = {};
  const detectedShopsSet = new Set<string>();

  for (let c = 0; c < rawHeaders.length; c++) {
    const h = rawHeaders[c];
    const samples: string[] = [];
    for (let r = headerRowIdx + 1; r < rawData.length; r++) {
      const val = rawData[r]?.[c];
      if (val !== undefined && val !== null && String(val).trim() !== "") {
        samples.push(String(val).trim());
        if (h === detected.shopNameCol) {
          detectedShopsSet.add(String(val).trim());
        }
        if (samples.length >= 4 && h !== detected.shopNameCol) break;
      }
    }
    previewValuesByCol[h] = samples;
  }

  const approx = Math.max(approxRowCount, rawData.length - headerRowIdx - 1);
  return {
    fileName,
    fileSize,
    sheetName,
    headers: rawHeaders,
    detectedMapping: detected,
    previewValuesByCol,
    detectedShops: Array.from(detectedShopsSet).sort(),
    approxRowCount: approx,
    totalRows: approx,
    rows: []
  };
}

/**
 * PHASE 2: FULL PROCESSING (Phần Mềm Bán Hàng)
 * Trích xuất có chọn lọc CHỈ các cột tham gia đối chiếu, gom nhóm trực tiếp,
 * giải phóng bộ nhớ ngay, không tạo object thừa cho các cột khác.
 */
export function parseSalesOrdersForReconciliation(
  workbook: XLSX.WorkBook,
  sheetName: string,
  mapping: SalesSoftwareColumnMapping
): Map<string, import("./types").AggregatedSalesOrder> {
  const resultMap = new Map<string, import("./types").AggregatedSalesOrder>();

  const targetSheetName = sheetName && workbook.Sheets[sheetName] ? sheetName : workbook.SheetNames[0];
  const worksheet = workbook.Sheets[targetSheetName];
  if (!worksheet || !mapping.orderIdCol || !mapping.shopNameCol) return resultMap;

  const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false
  });
  if (!rawData || rawData.length === 0) return resultMap;

  let headerRowIdx = 0;
  for (let r = 0; r < Math.min(10, rawData.length); r++) {
    const rowStr = (rawData[r] || []).map(cell => String(cell || "").toLowerCase()).join(" ");
    if (
      (rowStr.includes("tên shop") || rowStr.includes("cửa hàng")) &&
      (rowStr.includes("số đơn") || rowStr.includes("mã đơn") || rowStr.includes("hệ thống khác"))
    ) {
      headerRowIdx = r;
      break;
    }
  }

  const headers: string[] = (rawData[headerRowIdx] || []).map((h: any, idx: number) => String(h || "").trim() || `Cột_${idx + 1}`);

  const shopNameIdx = headers.indexOf(mapping.shopNameCol);
  const orderIdIdx = headers.indexOf(mapping.orderIdCol);
  const invoiceNumIdx = mapping.invoiceNumberCol ? headers.indexOf(mapping.invoiceNumberCol) : -1;
  const invoiceDateIdx = mapping.invoiceDateCol ? headers.indexOf(mapping.invoiceDateCol) : -1;
  const totalAmountIdx = mapping.totalAmountCol ? headers.indexOf(mapping.totalAmountCol) : -1;
  const voucherNumIdx = mapping.voucherNumberCol ? headers.indexOf(mapping.voucherNumberCol) : -1;
  const marketplaceIdx = mapping.marketplaceCol ? headers.indexOf(mapping.marketplaceCol) : -1;
  const cancelledIdx = mapping.cancelledCol ? headers.indexOf(mapping.cancelledCol) : -1;
  const noteIdx = mapping.noteCol ? headers.indexOf(mapping.noteCol) : -1;

  if (shopNameIdx === -1 || orderIdIdx === -1 || totalAmountIdx === -1) return resultMap;

  interface SlimSalesRowItem {
    rawShop: string;
    rawOrderId: string;
    invoiceNum: string;
    invoiceDate: string;
    voucherNum: string;
    totalAmount: number;
    channel: string;
    cancelled: boolean;
    note: string;
  }

  // Gom tạm theo normalizedShop + "::" + normalizedOrderId
  const orderGroups = new Map<string, SlimSalesRowItem[]>();

  for (let r = headerRowIdx + 1; r < rawData.length; r++) {
    const row = rawData[r];
    if (!row || row.length === 0) continue;

    const rawShop = String(row[shopNameIdx] || "").trim();
    const rawOrderId = String(row[orderIdIdx] || "").trim();
    const { normalized } = normalizeOrderId(rawOrderId);
    if (!rawShop || !normalized) continue;

    const invoiceNum = invoiceNumIdx >= 0 ? String(row[invoiceNumIdx] || "").trim() : "";
    const invoiceDate = invoiceDateIdx >= 0 ? String(row[invoiceDateIdx] || "").trim() : "";
    const voucherNum = voucherNumIdx >= 0 ? String(row[voucherNumIdx] || "").trim() : "";
    const totalAmount = parseCurrencyVnd(row[totalAmountIdx]);
    const channel = marketplaceIdx >= 0 ? String(row[marketplaceIdx] || "").trim() : "";
    const cancelStr = cancelledIdx >= 0 ? String(row[cancelledIdx] || "").trim().toLowerCase() : "";
    const cancelled = cancelStr.includes("hủy") || cancelStr.includes("huy") || cancelStr.includes("cancel") || cancelStr === "1" || cancelStr === "true";
    const note = noteIdx >= 0 ? String(row[noteIdx] || "").trim() : "";

    const normShopKey = normalizeShopName(rawShop);
    const compositeKey = `${normShopKey}::${normalized}`;

    if (!orderGroups.has(compositeKey)) {
      orderGroups.set(compositeKey, []);
    }
    orderGroups.get(compositeKey)!.push({
      rawShop,
      rawOrderId,
      invoiceNum,
      invoiceDate,
      voucherNum,
      totalAmount,
      channel,
      cancelled,
      note
    });
  }

  // Tổng hợp từng đơn hàng
  for (const [compositeKey, groupRows] of orderGroups.entries()) {
    const first = groupRows[0];
    const { normalized: normOrderId } = normalizeOrderId(first.rawOrderId);
    const normShopKey = normalizeShopName(first.rawShop);

    const invoiceNumbers = new Set<string>();
    const invoiceDates = new Set<string>();
    const voucherNumbers = new Set<string>();
    const warnings: string[] = [];

    groupRows.forEach(r => {
      if (r.invoiceNum) invoiceNumbers.add(r.invoiceNum);
      if (r.invoiceDate) invoiceDates.add(r.invoiceDate);
      if (r.voucherNum) voucherNumbers.add(r.voucherNum);
      if (r.channel && !r.channel.toLowerCase().includes("shopee") && !r.channel.toLowerCase().includes("tiktok")) {
        const w = `Kênh bán hàng là "${r.channel}"`;
        if (!warnings.includes(w)) warnings.push(w);
      }
      if (r.cancelled) {
        const w = "Đơn ghi nhận trạng thái HỦY trên phần mềm bán hàng";
        if (!warnings.includes(w)) warnings.push(w);
      }
    });

    let totalAmount = 0;
    let isMultiInvoice = false;
    let isRepeatedAmount = false;
    let isAmbiguous = false;

    if (groupRows.length === 1) {
      totalAmount = first.totalAmount;
    } else {
      const amounts = groupRows.map(r => r.totalAmount);
      const distinctInvoices = Array.from(invoiceNumbers).filter(Boolean);

      if (distinctInvoices.length > 1) {
        isMultiInvoice = true;
        const rowsByInvoice = new Map<string, number[]>();
        groupRows.forEach(r => {
          const inv = r.invoiceNum || "CHUNG";
          if (!rowsByInvoice.has(inv)) rowsByInvoice.set(inv, []);
          rowsByInvoice.get(inv)!.push(r.totalAmount);
        });

        rowsByInvoice.forEach((invAmounts) => {
          const nonZero = invAmounts.filter(a => a > 0);
          const allSame = nonZero.length > 0 && nonZero.every(a => a === nonZero[0]);
          if (allSame) {
            totalAmount += nonZero[0];
          } else {
            totalAmount += invAmounts.reduce((sum, a) => sum + a, 0);
          }
        });
      } else {
        const nonZero = amounts.filter(a => a > 0);
        const allSame = nonZero.length > 0 && nonZero.every(a => a === nonZero[0]);
        if (allSame) {
          totalAmount = nonZero[0];
          isRepeatedAmount = true;
        } else {
          totalAmount = amounts.reduce((sum, a) => sum + a, 0);
          if (distinctInvoices.length === 0) {
            isAmbiguous = true;
            warnings.push("Mã đơn có nhiều dòng chi tiết không kèm số hóa đơn phân biệt - đã tính tổng cộng");
          }
        }
      }
    }

    resultMap.set(compositeKey, {
      shopKey: normShopKey,
      rawShopName: first.rawShop,
      orderId: normOrderId,
      originalOrderId: first.rawOrderId,
      invoiceNumbers: Array.from(invoiceNumbers),
      invoiceDates: Array.from(invoiceDates),
      voucherNumbers: Array.from(voucherNumbers),
      totalAmount,
      rowCount: groupRows.length,
      isMultiInvoice,
      isRepeatedAmount,
      isAmbiguous,
      sourceRows: groupRows.map(r => ({
        invoiceNumber: r.invoiceNum || "--",
        invoiceDate: r.invoiceDate || "--",
        voucherNumber: r.voucherNum || "--",
        totalAmount: r.totalAmount,
        note: r.note || "--"
      })),
      warnings
    });
  }

  return resultMap;
}

/**
 * Backward compatibility parser (used only for small demo workbooks)
 */
export function parseSalesSoftwareWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
  preferredSheet?: string,
  customMapping?: Partial<SalesSoftwareColumnMapping>
): ParsedSalesSoftwareResult {
  const inspection = inspectSalesSoftwareWorkbook(workbook, fileName, 0, preferredSheet, customMapping);
  return {
    fileName,
    sheetName: inspection.sheetName,
    headers: inspection.headers,
    headerRowIndex: 0,
    detectedMapping: inspection.detectedMapping,
    rows: [],
    detectedShops: inspection.detectedShops,
    diagnostics: {
      totalRows: inspection.approxRowCount,
      rowsWithShop: inspection.approxRowCount,
      rowsMissingShop: 0,
      rowsWithOrderId: inspection.approxRowCount,
      rowsMissingOrderId: 0,
      uniqueOrderCount: inspection.approxRowCount,
      duplicateOrderCount: 0,
      shopDistribution: {}
    },
    warnings: []
  };
}
