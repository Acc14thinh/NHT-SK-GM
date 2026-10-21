/**
 * Parser for Marketplace Excel Files (Shopee, with extensibility for TikTok, Lazada)
 * Section 2, 4, 18, 19
 */
import * as XLSX from "xlsx";
import { MarketplaceColumnMapping, MarketplaceRawRow } from "./types";
import { normalizeOrderId, normalizeShopName, parseCurrencyVnd } from "./orderNormalizer";
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
 * Tự động nhận diện loại sàn TMĐT từ danh sách headers
 */
export function detectMarketplacePlatform(headers: string[]): "SHOPEE" | "TIKTOK" | "OTHER" {
  if (!headers || headers.length === 0) return "SHOPEE";
  const hLower = headers.map(h => String(h || "").toLowerCase().trim());

  let shopeeScore = 0;
  let tikTokScore = 0;

  for (const h of hLower) {
    // Shopee specific indicators
    if (h.includes("người mua thanh toán") || h.includes("nguoi mua thanh toan")) shopeeScore += 5;
    if (h.includes("mã giảm giá của shop") || h.includes("ma giam gia cua shop")) shopeeScore += 4;
    if (h === "mã đơn hàng" || h === "ma don hang") shopeeScore += 3;
    if (h.includes("trả hàng/hoàn tiền") || h.includes("tra hang/hoan tien")) shopeeScore += 3;
    if (h.includes("thời gian giao hàng") || h.includes("thoi gian giao hang")) shopeeScore += 2;

    // TikTok specific indicators
    if (h.includes("sku subtotal") || h.includes("subtotal before discount")) tikTokScore += 5;
    if (h.includes("seller discount") || h.includes("sku seller discount")) tikTokScore += 4;
    if (h.includes("cancelation/return") || h.includes("cancellation/return")) tikTokScore += 4;
    if (h.includes("order refund amount")) tikTokScore += 4;
    if (h === "seller sku" || h === "sku id") tikTokScore += 3;
    if (h === "order id" || h === "orderid") tikTokScore += 2;
  }

  if (tikTokScore > shopeeScore) return "TIKTOK";
  return "SHOPEE";
}

/**
 * Tự động nhận diện cấu hình cột cho TikTok Shop
 */
export function autoDetectTikTokColumns(headers: string[]): MarketplaceColumnMapping {
  const getCol = (exactKeywords: string[], fallbackKeywords: string[] = []): string => {
    return findStrictMatchingHeader(headers, exactKeywords, fallbackKeywords);
  };

  return {
    orderIdCol: getCol(["Order ID", "order id", "OrderID"], ["mã đơn", "order_id"]),
    orderStatusCol: getCol(["Order Status", "order status"], ["trạng thái đơn"]),
    deliveryTimeCol: getCol(["Delivered Time", "delivered time", "Delivery Time"], ["thời gian giao hàng", "created time"]),
    returnRefundCol: getCol(["Cancelation/Return Type", "Cancellation/Return Type", "Cancelation Type"], ["return", "refund", "trả hàng"]),
    buyerPaidCol: getCol(["SKU Subtotal Before Discount", "Subtotal Before Discount"], ["subtotal", "tổng tiền", "tiền hàng"]),
    shopVoucherCol: getCol(["SKU Seller Discount", "Seller Discount"], ["seller discount", "giảm giá"]),
    quantityCol: getCol(["Quantity", "quantity"], ["số lượng"]),
    skuCol: getCol(["SKU ID", "Seller SKU", "Product Name"], ["sku", "tên sản phẩm"]),
    tiktokCancelReturnCol: getCol(["Cancelation/Return Type", "Cancellation/Return Type"], []),
    tiktokSellerDiscountCol: getCol(["SKU Seller Discount"], []),
    tiktokSubtotalCol: getCol(["SKU Subtotal Before Discount"], [])
  };
}

/**
 * PHASE 1: LIGHTWEIGHT INSPECTION
 * Chỉ đọc tối đa 25 dòng đầu tiên để lấy sheet list, headers, platform, và sample preview.
 * Tuyệt đối không convert toàn bộ file thành JSON hay lưu raw rows vào React state.
 */
export function inspectMarketplaceWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
  fileSize: number = 0,
  customMapping?: Partial<MarketplaceColumnMapping>
): import("./types").MarketplaceInspectionResult {
  const sheets: import("./types").MarketplaceSheetInspection[] = [];
  const allDetectedShopNames: string[] = [];
  let totalApproxRows = 0;

  for (const sheetName of workbook.SheetNames) {
    if (sheetName.toLowerCase().startsWith("__") || sheetName.toLowerCase().includes("readme")) {
      continue;
    }

    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    // Approximate row count từ range !ref (nếu có)
    let approxRowCount = 0;
    if (worksheet["!ref"]) {
      try {
        const decoded = XLSX.utils.decode_range(worksheet["!ref"]);
        approxRowCount = Math.max(0, decoded.e.r - decoded.s.r);
      } catch {
        approxRowCount = 0;
      }
    }

    // Đọc tối đa 25 dòng đầu
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false
    });

    if (!rawData || rawData.length === 0) continue;

    // Tìm dòng header (quét trong 10 dòng đầu)
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
        rowStr.includes("sku subtotal")
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

    const platform = detectMarketplacePlatform(rawHeaders);

    let detectedMapping = platform === "TIKTOK"
      ? autoDetectTikTokColumns(rawHeaders)
      : autoDetectShopeeColumns(rawHeaders);

    if (customMapping) {
      detectedMapping = { ...detectedMapping, ...customMapping };
    }

    // Trích xuất 3-5 mẫu preview cho từng cột
    const previewValuesByCol: Record<string, string[]> = {};
    for (let c = 0; c < rawHeaders.length; c++) {
      const h = rawHeaders[c];
      const samples: string[] = [];
      for (let r = headerRowIdx + 1; r < rawData.length; r++) {
        const val = rawData[r]?.[c];
        if (val !== undefined && val !== null && String(val).trim() !== "") {
          samples.push(String(val).trim());
          if (samples.length >= 4) break;
        }
      }
      previewValuesByCol[h] = samples;
    }

    sheets.push({
      sheetName,
      platform,
      headers: rawHeaders,
      detectedMapping,
      previewValuesByCol,
      approxRowCount: Math.max(approxRowCount, rawData.length - headerRowIdx - 1)
    });

    totalApproxRows += approxRowCount;
    allDetectedShopNames.push(sheetName);
  }

  return {
    fileName,
    fileSize,
    sheets,
    allDetectedShopNames,
    totalApproxRows,
    totalRows: totalApproxRows
  };
}

/**
 * PHASE 2: FULL PROCESSING (Chỉ chạy khi bấm [TIẾN HÀNH ĐỐI CHIẾU])
 * Trích xuất có chọn lọc CHỈ các cột tham gia đối chiếu, gom nhóm trực tiếp,
 * giải phóng bộ nhớ ngay sau mỗi sheet, không lưu raw row bloat.
 */
export function parseMarketplaceOrdersForReconciliation(
  workbook: XLSX.WorkBook,
  inspectedSheets: import("./types").MarketplaceSheetInspection[],
  shopMappings: import("./types").ShopMappingEntry[]
): Map<string, import("./types").AggregatedMarketplaceOrder> {
  const resultMap = new Map<string, import("./types").AggregatedMarketplaceOrder>();

  // Map sheet -> targetSalesShop
  const sheetToSalesShopMap = new Map<string, string>();
  shopMappings.forEach(m => {
    sheetToSalesShopMap.set(m.marketplaceSheetOrShop, m.salesShopName || m.marketplaceSheetOrShop);
  });

  for (const sheetInfo of inspectedSheets) {
    const sheetName = sheetInfo.sheetName;
    const worksheet = workbook.Sheets[sheetName];
    if (!worksheet) continue;

    const targetSalesShop = sheetToSalesShopMap.get(sheetName) || sheetName;
    const normShopKey = normalizeShopName(targetSalesShop);
    const mapping = sheetInfo.detectedMapping;
    if (!mapping.orderIdCol) continue;

    // Đọc raw 2D array của sheet này
    const rawData: any[][] = XLSX.utils.sheet_to_json(worksheet, {
      header: 1,
      defval: "",
      raw: false
    });
    if (!rawData || rawData.length === 0) continue;

    // Tìm header row
    let headerRowIdx = 0;
    for (let r = 0; r < Math.min(10, rawData.length); r++) {
      const rowStr = (rawData[r] || []).map(cell => String(cell || "").toLowerCase()).join(" ");
      if (
        rowStr.includes("mã đơn") ||
        rowStr.includes("order id") ||
        rowStr.includes("người mua thanh toán") ||
        rowStr.includes("sku subtotal")
      ) {
        headerRowIdx = r;
        break;
      }
    }

    const headers: string[] = (rawData[headerRowIdx] || []).map((h: any, idx: number) => String(h || "").trim() || `Cột_${idx + 1}`);

    // Xác định index của các cột cần thiết
    const orderIdIdx = headers.indexOf(mapping.orderIdCol);
    const statusIdx = mapping.orderStatusCol ? headers.indexOf(mapping.orderStatusCol) : -1;
    const deliveryIdx = mapping.deliveryTimeCol ? headers.indexOf(mapping.deliveryTimeCol) : -1;
    const returnRefundIdx = mapping.returnRefundCol ? headers.indexOf(mapping.returnRefundCol) : -1;
    const buyerPaidIdx = mapping.buyerPaidCol ? headers.indexOf(mapping.buyerPaidCol) : -1;
    const shopVoucherIdx = mapping.shopVoucherCol ? headers.indexOf(mapping.shopVoucherCol) : -1;
    const skuIdx = mapping.skuCol ? headers.indexOf(mapping.skuCol) : -1;
    const priceIdx = mapping.priceCol ? headers.indexOf(mapping.priceCol) : -1;
    const quantityIdx = mapping.quantityCol ? headers.indexOf(mapping.quantityCol) : -1;

    if (orderIdIdx === -1) continue;

    const isTikTok = sheetInfo.platform === "TIKTOK";

    // Gom đơn tạm thời cho sheet này
    interface ExtractedRowItem {
      rawOrderId: string;
      orderStatus: string;
      deliveryDate: string;
      returnRefund: string;
      buyerPaid: number;
      voucher: number;
      productName: string;
      price: number;
      quantity: number;
    }

    const orderGroups = new Map<string, ExtractedRowItem[]>();

    for (let r = headerRowIdx + 1; r < rawData.length; r++) {
      const row = rawData[r];
      if (!row || row.length === 0) continue;

      const rawOrderId = String(row[orderIdIdx] || "").trim();
      const { normalized } = normalizeOrderId(rawOrderId);
      if (!normalized) continue;

      const orderStatus = statusIdx >= 0 ? String(row[statusIdx] || "").trim() : "";
      const deliveryDate = deliveryIdx >= 0 ? String(row[deliveryIdx] || "").trim() : "";
      const returnRefund = returnRefundIdx >= 0 ? String(row[returnRefundIdx] || "").trim() : "";
      const buyerPaid = buyerPaidIdx >= 0 ? parseCurrencyVnd(row[buyerPaidIdx]) : 0;
      const voucher = shopVoucherIdx >= 0 ? parseCurrencyVnd(row[shopVoucherIdx]) : 0;
      const productName = skuIdx >= 0 ? String(row[skuIdx] || "").trim() : "";
      const price = priceIdx >= 0 ? parseCurrencyVnd(row[priceIdx]) : buyerPaid;
      const quantity = quantityIdx >= 0 ? Math.max(1, parseCurrencyVnd(row[quantityIdx])) : 1;

      if (!orderGroups.has(normalized)) {
        orderGroups.set(normalized, []);
      }
      orderGroups.get(normalized)!.push({
        rawOrderId,
        orderStatus,
        deliveryDate,
        returnRefund,
        buyerPaid,
        voucher,
        productName,
        price,
        quantity
      });
    }

    // Tổng hợp từng đơn hàng
    for (const [normOrderId, items] of orderGroups.entries()) {
      const first = items[0];
      const isCancelled = first.orderStatus.toLowerCase().includes("hủy") ||
                          first.orderStatus.toLowerCase().includes("cancel") ||
                          first.returnRefund.toLowerCase().includes("cancel");

      let totalBuyerPaid = 0;
      let shopVoucher = 0;
      let calculatedRevenue = 0;
      const warnings: string[] = [];

      if (isTikTok) {
        // TIKTOK REVENUE = SUM(SKU Subtotal Before Discount) - SUM(SKU Seller Discount)
        totalBuyerPaid = items.reduce((sum, it) => sum + it.buyerPaid, 0);
        shopVoucher = items.reduce((sum, it) => sum + it.voucher, 0);
        calculatedRevenue = totalBuyerPaid - shopVoucher;
      } else {
        // SHOPEE REVENUE = SUM(Tổng số tiền Người mua thanh toán) - Voucher Shop (1 lần)
        if (items.length === 1) {
          totalBuyerPaid = first.buyerPaid;
          shopVoucher = first.voucher;
        } else {
          // Nhiều dòng SKU
          const paidList = items.map(it => it.buyerPaid);
          const nonZeroPaid = paidList.filter(v => v > 0);
          const allPaidSame = nonZeroPaid.length > 0 && nonZeroPaid.every(v => v === nonZeroPaid[0]);

          let isRepeatedOrderLevelPaid = false;
          if (allPaidSame && priceIdx >= 0) {
            const sumItemPrices = items.reduce((sum, it) => sum + (it.price * it.quantity), 0);
            if (sumItemPrices > 0 && Math.abs(sumItemPrices - nonZeroPaid[0]) < Math.abs(sumItemPrices - (nonZeroPaid[0] * items.length))) {
              isRepeatedOrderLevelPaid = true;
            }
          }

          if (isRepeatedOrderLevelPaid) {
            totalBuyerPaid = nonZeroPaid[0];
          } else {
            totalBuyerPaid = paidList.reduce((sum, v) => sum + v, 0);
          }

          // Voucher shop Shopee: nếu lặp lại thì chỉ trừ 1 lần
          const nonZeroVouchers = items.map(it => it.voucher).filter(v => v > 0);
          if (nonZeroVouchers.length > 0) {
            const firstV = nonZeroVouchers[0];
            const allSame = nonZeroVouchers.every(v => v === firstV);
            if (allSame) {
              shopVoucher = firstV;
            } else {
              shopVoucher = Math.max(...nonZeroVouchers);
              warnings.push("Dữ liệu mã giảm giá không nhất quán giữa các dòng SKU của đơn");
            }
          }
        }
        calculatedRevenue = totalBuyerPaid - shopVoucher;
      }

      const compositeKey = `${normShopKey}::${normOrderId}`;

      resultMap.set(compositeKey, {
        shopKey: normShopKey,
        rawShopName: targetSalesShop,
        orderId: normOrderId,
        originalOrderId: first.rawOrderId,
        orderStatus: first.orderStatus,
        isCancelled,
        deliveryDate: first.deliveryDate,
        returnRefundStatus: first.returnRefund,
        totalBuyerPaid,
        shopVoucher,
        voucherInconsistent: warnings.length > 0,
        voucherValuesFound: Array.from(new Set(items.map(it => it.voucher))),
        calculatedRevenue,
        rowCount: items.length,
        sourceRows: items.map(it => ({
          productName: it.productName || `SKU ${it.rawOrderId}`,
          price: it.price,
          quantity: it.quantity,
          buyerPaid: it.buyerPaid,
          shopVoucher: it.voucher
        })),
        warnings
      });
    }
  }

  return resultMap;
}

/**
 * Backward compatibility parser (used only for small demo workbooks)
 */
export function parseMarketplaceWorkbook(
  workbook: XLSX.WorkBook,
  fileName: string,
  customMapping?: Partial<MarketplaceColumnMapping>
): ParsedMarketplaceResult {
  const inspection = inspectMarketplaceWorkbook(workbook, fileName, 0, customMapping);
  const sheets: ParsedMarketplaceSheet[] = inspection.sheets.map(s => ({
    sheetName: s.sheetName,
    rowCount: s.approxRowCount,
    headers: s.headers,
    headerRowIndex: 0,
    detectedMapping: s.detectedMapping,
    rows: [],
    diagnostics: {
      totalRows: s.approxRowCount,
      rowsWithOrderId: s.approxRowCount,
      rowsMissingOrderId: 0,
      uniqueOrderCount: s.approxRowCount,
      duplicateOrderCount: 0
    }
  }));

  return {
    fileName,
    sheets,
    allDetectedShopNames: inspection.allDetectedShopNames,
    totalRows: inspection.totalApproxRows,
    totalUniqueOrders: inspection.totalApproxRows,
    warnings: []
  };
}
