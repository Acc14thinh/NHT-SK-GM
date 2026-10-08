/**
 * Order Aggregator for Marketplace and Sales Software Files
 * Section 7, 8, 9 & 21
 */
import {
  AggregatedMarketplaceOrder,
  AggregatedSalesOrder,
  MarketplaceColumnMapping,
  MarketplaceRawRow,
  SalesSoftwareColumnMapping,
  SalesSoftwareRawRow,
  ShopMappingEntry
} from "./types";
import { normalizeOrderId, normalizeShopName, parseCurrencyVnd } from "./orderNormalizer";
import { ParsedMarketplaceSheet } from "./shopeeParser";

/**
 * Gom nhóm dữ liệu sàn TMĐT theo Shop + Mã Đơn Hàng
 */
export function aggregateMarketplaceOrders(
  sheets: ParsedMarketplaceSheet[],
  shopMappings: ShopMappingEntry[]
): Map<string, AggregatedMarketplaceOrder> {
  // Key: normalizedShopId + "::" + normalizedOrderId
  const resultMap = new Map<string, AggregatedMarketplaceOrder>();

  // Tạo map tra cứu: marketplaceSheetOrShop -> salesShopName
  const sheetToSalesShopMap = new Map<string, string>();
  shopMappings.forEach(m => {
    sheetToSalesShopMap.set(m.marketplaceSheetOrShop, m.salesShopName || m.marketplaceSheetOrShop);
  });

  for (const sheet of sheets) {
    const rawShopName = sheet.sheetName;
    const targetSalesShop = sheetToSalesShopMap.get(rawShopName) || rawShopName;
    const normShopKey = normalizeShopName(targetSalesShop);

    const mapping = sheet.detectedMapping;
    if (!mapping.orderIdCol) continue;

    // Nhóm tạm thời các dòng theo normalizedOrderId trong sheet này
    const orderGroups = new Map<string, MarketplaceRawRow[]>();

    for (const row of sheet.rows) {
      const orderIdVal = row[mapping.orderIdCol];
      const { normalized } = normalizeOrderId(orderIdVal);
      if (!normalized) continue;

      if (!orderGroups.has(normalized)) {
        orderGroups.set(normalized, []);
      }
      orderGroups.get(normalized)!.push(row);
    }

    // Xử lý từng đơn hàng
    for (const [normOrderId, rows] of orderGroups.entries()) {
      const firstRow = rows[0];
      const originalOrderId = String(firstRow[mapping.orderIdCol] || "").trim();

      // Thông tin trạng thái
      const orderStatus = mapping.orderStatusCol ? String(firstRow[mapping.orderStatusCol] || "").trim() : "";
      const isCancelled = orderStatus.toLowerCase().includes("hủy") ||
                          orderStatus.toLowerCase().includes("cancelled") ||
                          orderStatus.toLowerCase().includes("canceled");

      const deliveryDate = mapping.deliveryTimeCol ? String(firstRow[mapping.deliveryTimeCol] || "").trim() : "";
      const returnRefundStatus = mapping.returnRefundCol ? String(firstRow[mapping.returnRefundCol] || "").trim() : "";

      // 1. Tính tổng "Tổng số tiền Người mua thanh toán"
      let totalBuyerPaid = 0;
      for (const r of rows) {
        if (mapping.buyerPaidCol) {
          totalBuyerPaid += parseCurrencyVnd(r[mapping.buyerPaidCol]);
        }
      }

      // 2. Xử lý "Mã giảm giá của Shop" (Section 7)
      // Nếu có nhiều dòng SKU và mã giảm giá lặp lại: CHỈ TRỪ 1 LẦN.
      // Nếu giá trị mã giảm giá giữa các dòng không nhất quán: CẢNH BÁO.
      const voucherValues: number[] = [];
      const nonZeroVouchers: number[] = [];

      for (const r of rows) {
        if (mapping.shopVoucherCol) {
          const v = parseCurrencyVnd(r[mapping.shopVoucherCol]);
          voucherValues.push(v);
          if (v > 0) nonZeroVouchers.push(v);
        }
      }

      let shopVoucher = 0;
      let voucherInconsistent = false;
      const warnings: string[] = [];

      if (nonZeroVouchers.length > 0) {
        // Kiểm tra xem tất cả các giá trị voucher có giống nhau không
        const firstVal = nonZeroVouchers[0];
        const allSame = nonZeroVouchers.every(v => v === firstVal);

        if (allSame) {
          // Lặp lại cùng 1 giá trị -> CHỈ TRỪ 1 LẦN!
          shopVoucher = firstVal;
        } else {
          // Không nhất quán! (Section 7)
          voucherInconsistent = true;
          shopVoucher = Math.max(...nonZeroVouchers); // tạm lấy max để tính toán nhưng gắn cờ
          warnings.push("Dữ liệu mã giảm giá không nhất quán giữa các dòng SKU của đơn");
        }
      }

      // DOANH THU ĐƠN = Tổng tiền người mua trả - Voucher shop
      const calculatedRevenue = totalBuyerPaid - shopVoucher;

      const compositeKey = `${normShopKey}::${normOrderId}`;

      resultMap.set(compositeKey, {
        shopKey: normShopKey,
        rawShopName: targetSalesShop,
        orderId: normOrderId,
        originalOrderId,
        orderStatus,
        isCancelled,
        deliveryDate,
        returnRefundStatus,
        totalBuyerPaid,
        shopVoucher,
        voucherInconsistent,
        voucherValuesFound: Array.from(new Set(voucherValues)),
        calculatedRevenue,
        rowCount: rows.length,
        sourceRows: rows,
        warnings
      });
    }
  }

  return resultMap;
}

/**
 * Gom nhóm dữ liệu phần mềm bán hàng theo Tên Shop + Số đơn hàng từ hệ thống khác
 * Section 9: Xử lý thông minh trùng lặp, hóa đơn đa dòng, nhiều chứng từ
 */
export function aggregateSalesOrders(
  rows: SalesSoftwareRawRow[],
  mapping: SalesSoftwareColumnMapping
): Map<string, AggregatedSalesOrder> {
  const resultMap = new Map<string, AggregatedSalesOrder>();

  if (!mapping.orderIdCol || !mapping.shopNameCol) return resultMap;

  // Gom tạm theo normalizedShop + "::" + normalizedOrderId
  const orderGroups = new Map<string, { rawShop: string; rawOrderId: string; rows: SalesSoftwareRawRow[] }>();

  for (const row of rows) {
    const rawShop = String(row[mapping.shopNameCol] || "").trim();
    const rawOrderId = row[mapping.orderIdCol];
    const { normalized } = normalizeOrderId(rawOrderId);

    if (!rawShop || !normalized) continue;

    const normShop = normalizeShopName(rawShop);
    const compositeKey = `${normShop}::${normalized}`;

    if (!orderGroups.has(compositeKey)) {
      orderGroups.set(compositeKey, {
        rawShop,
        rawOrderId: String(rawOrderId || "").trim(),
        rows: []
      });
    }
    orderGroups.get(compositeKey)!.rows.push(row);
  }

  // Xử lý từng nhóm đơn hàng bán hàng
  for (const [compositeKey, group] of orderGroups.entries()) {
    const { rawShop, rawOrderId, rows: groupRows } = group;
    const normShopKey = normalizeShopName(rawShop);
    const { normalized: normOrderId } = normalizeOrderId(rawOrderId);

    const invoiceNumbers = new Set<string>();
    const invoiceDates = new Set<string>();
    const voucherNumbers = new Set<string>();
    const warnings: string[] = [];

    // Thu thập các metadata
    groupRows.forEach(r => {
      if (mapping.invoiceNumberCol && r[mapping.invoiceNumberCol]) {
        invoiceNumbers.add(String(r[mapping.invoiceNumberCol]).trim());
      }
      if (mapping.invoiceDateCol && r[mapping.invoiceDateCol]) {
        invoiceDates.add(String(r[mapping.invoiceDateCol]).trim());
      }
      if (mapping.voucherNumberCol && r[mapping.voucherNumberCol]) {
        voucherNumbers.add(String(r[mapping.voucherNumberCol]).trim());
      }
    });

    let totalAmount = 0;
    let isMultiInvoice = false;
    let isRepeatedAmount = false;
    let isAmbiguous = false;

    if (groupRows.length === 1) {
      // Trường hợp 1: Chỉ có 1 dòng duy nhất
      totalAmount = mapping.totalAmountCol ? parseCurrencyVnd(groupRows[0][mapping.totalAmountCol]) : 0;
    } else {
      // Trường hợp 2: Có nhiều dòng cho cùng 1 đơn hàng (Section 9)
      const amounts = groupRows.map(r => mapping.totalAmountCol ? parseCurrencyVnd(r[mapping.totalAmountCol]) : 0);

      // Kiểm tra xem có phân tách theo số hóa đơn hoặc số chứng từ không
      const distinctVouchers = Array.from(voucherNumbers).filter(Boolean);
      const distinctInvoices = Array.from(invoiceNumbers).filter(Boolean);

      // Tình huống A: Cùng 1 số chứng từ/hóa đơn, và tổng tiền bị lặp lại y hệt trên các dòng
      const allAmountsSame = amounts.length > 0 && amounts.every(a => a === amounts[0]);

      if (allAmountsSame && (distinctVouchers.length <= 1 && distinctInvoices.length <= 1)) {
        // Cùng 1 chứng từ và tổng tiền bị lặp lại trên nhiều dòng chi tiết SKU -> Chỉ lấy 1 lần!
        totalAmount = amounts[0];
        isRepeatedAmount = true;
      } else if (distinctInvoices.length > 1) {
        // Tình huống B: Có nhiều số hóa đơn thực sự khác nhau -> Gom tổng theo từng hóa đơn DISTINCT
        isMultiInvoice = true;
        // Group rows by invoice number
        const invoiceToAmount = new Map<string, number>();
        groupRows.forEach(r => {
          const inv = mapping.invoiceNumberCol ? String(r[mapping.invoiceNumberCol] || "").trim() : "CHUNG";
          const amt = mapping.totalAmountCol ? parseCurrencyVnd(r[mapping.totalAmountCol]) : 0;
          // Nếu trong cùng hóa đơn tiền bị lặp, lấy 1 lần
          invoiceToAmount.set(inv, amt);
        });
        totalAmount = Array.from(invoiceToAmount.values()).reduce((sum, a) => sum + a, 0);
      } else if (distinctVouchers.length > 1) {
        // Tình huống C: Nhiều chứng từ khác nhau
        isMultiInvoice = true;
        const voucherToAmount = new Map<string, number>();
        groupRows.forEach(r => {
          const v = mapping.voucherNumberCol ? String(r[mapping.voucherNumberCol] || "").trim() : "CHUNG";
          const amt = mapping.totalAmountCol ? parseCurrencyVnd(r[mapping.totalAmountCol]) : 0;
          voucherToAmount.set(v, amt);
        });
        totalAmount = Array.from(voucherToAmount.values()).reduce((sum, a) => sum + a, 0);
      } else {
        // Tình huống D: Các dòng chi tiết có số tiền khác nhau (mỗi dòng là 1 SKU của cùng 1 đơn) -> SUM các dòng
        // Kiểm tra xem tổng các dòng có hợp lý không
        const sumAll = amounts.reduce((sum, a) => sum + a, 0);
        totalAmount = sumAll;
        if (!allAmountsSame && distinctInvoices.length === 0 && distinctVouchers.length === 0) {
          isAmbiguous = true;
          warnings.push("Mã đơn có nhiều dòng chi tiết không kèm số hóa đơn phân biệt - đã tính tổng cộng");
        }
      }
    }

    resultMap.set(compositeKey, {
      shopKey: normShopKey,
      rawShopName: rawShop,
      orderId: normOrderId,
      originalOrderId: rawOrderId,
      invoiceNumbers: Array.from(invoiceNumbers),
      invoiceDates: Array.from(invoiceDates),
      voucherNumbers: Array.from(voucherNumbers),
      totalAmount,
      rowCount: groupRows.length,
      isMultiInvoice,
      isRepeatedAmount,
      isAmbiguous,
      sourceRows: groupRows,
      warnings
    });
  }

  return resultMap;
}
