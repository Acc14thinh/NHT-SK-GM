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
      if (rows.length === 1) {
        totalBuyerPaid = mapping.buyerPaidCol ? parseCurrencyVnd(rows[0][mapping.buyerPaidCol]) : 0;
      } else {
        // Nhiều dòng SKU cho 1 đơn hàng sàn
        const buyerPaidList = rows.map(r => mapping.buyerPaidCol ? parseCurrencyVnd(r[mapping.buyerPaidCol]) : 0);
        const nonZeroPaid = buyerPaidList.filter(v => v > 0);
        const allPaidSame = nonZeroPaid.length > 0 && nonZeroPaid.every(v => v === nonZeroPaid[0]);

        // Kiểm tra nếu có cột giá sản phẩm để phân biệt giá trị dòng vs giá trị toàn đơn
        let isRepeatedOrderLevelPaid = false;
        if (allPaidSame && mapping.priceCol) {
          const itemPrices = rows.map(r => {
            const p = parseCurrencyVnd(r[mapping.priceCol!]);
            const q = mapping.quantityCol ? Math.max(1, parseCurrencyVnd(r[mapping.quantityCol!])) : 1;
            return p * q;
          });
          const sumItemPrices = itemPrices.reduce((sum, p) => sum + p, 0);
          // Nếu số tiền trên mỗi dòng bằng tổng giá của tất cả các dòng cộng lại
          if (sumItemPrices > 0 && Math.abs(sumItemPrices - nonZeroPaid[0]) < Math.abs(sumItemPrices - (nonZeroPaid[0] * nonZeroPaid.length))) {
            isRepeatedOrderLevelPaid = true;
          }
        }

        if (isRepeatedOrderLevelPaid) {
          totalBuyerPaid = nonZeroPaid[0];
        } else {
          totalBuyerPaid = buyerPaidList.reduce((sum, v) => sum + v, 0);
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
 * Xử lý thông minh trùng lặp, hóa đơn đa dòng, nhiều chứng từ
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
      // Kiểm tra Kênh bán hàng (Section B: Cảnh báo nếu khác Shopee)
      if (mapping.marketplaceCol && r[mapping.marketplaceCol]) {
        const channel = String(r[mapping.marketplaceCol]).trim();
        if (channel && !channel.toLowerCase().includes("shopee")) {
          const w = `Kênh bán hàng là "${channel}" (khác Shopee)`;
          if (!warnings.includes(w)) warnings.push(w);
        }
      }
      // Kiểm tra Trạng thái hủy trên PM
      if (mapping.cancelledCol && r[mapping.cancelledCol]) {
        const cVal = String(r[mapping.cancelledCol]).trim().toLowerCase();
        if (cVal.includes("hủy") || cVal.includes("huy") || cVal.includes("cancel") || cVal === "1" || cVal === "true") {
          const w = "Đơn ghi nhận trạng thái HỦY trên phần mềm bán hàng";
          if (!warnings.includes(w)) warnings.push(w);
        }
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
      // Trường hợp 2: Có nhiều dòng cho cùng 1 đơn hàng
      const amounts = groupRows.map(r => mapping.totalAmountCol ? parseCurrencyVnd(r[mapping.totalAmountCol]) : 0);
      const distinctVouchers = Array.from(voucherNumbers).filter(Boolean);
      const distinctInvoices = Array.from(invoiceNumbers).filter(Boolean);

      if (distinctInvoices.length > 1) {
        // Tình huống B: Có nhiều số hóa đơn thực sự khác nhau -> Gom tổng theo từng hóa đơn DISTINCT
        isMultiInvoice = true;
        const rowsByInvoice = new Map<string, number[]>();
        groupRows.forEach(r => {
          const inv = mapping.invoiceNumberCol ? String(r[mapping.invoiceNumberCol] || "").trim() : "CHUNG";
          const amt = mapping.totalAmountCol ? parseCurrencyVnd(r[mapping.totalAmountCol]) : 0;
          if (!rowsByInvoice.has(inv)) rowsByInvoice.set(inv, []);
          rowsByInvoice.get(inv)!.push(amt);
        });

        totalAmount = 0;
        for (const [, invAmounts] of rowsByInvoice.entries()) {
          const nonZero = invAmounts.filter(a => a > 0);
          if (nonZero.length === 0) continue;
          const allSame = nonZero.every(a => a === nonZero[0]);
          if (allSame) {
            // Tiền lặp lại ở các dòng SKU của cùng 1 hóa đơn -> chỉ lấy 1 lần
            totalAmount += nonZero[0];
            if (invAmounts.length > 1) isRepeatedAmount = true;
          } else {
            // Các dòng chi tiết của cùng hóa đơn có tiền khác nhau -> cộng dồn
            totalAmount += invAmounts.reduce((s, a) => s + a, 0);
          }
        }
      } else if (distinctVouchers.length > 1) {
        // Tình huống C: Nhiều chứng từ khác nhau
        isMultiInvoice = true;
        const rowsByVoucher = new Map<string, number[]>();
        groupRows.forEach(r => {
          const v = mapping.voucherNumberCol ? String(r[mapping.voucherNumberCol] || "").trim() : "CHUNG";
          const amt = mapping.totalAmountCol ? parseCurrencyVnd(r[mapping.totalAmountCol]) : 0;
          if (!rowsByVoucher.has(v)) rowsByVoucher.set(v, []);
          rowsByVoucher.get(v)!.push(amt);
        });

        totalAmount = 0;
        for (const [, vAmounts] of rowsByVoucher.entries()) {
          const nonZero = vAmounts.filter(a => a > 0);
          if (nonZero.length === 0) continue;
          const allSame = nonZero.every(a => a === nonZero[0]);
          if (allSame) {
            totalAmount += nonZero[0];
            if (vAmounts.length > 1) isRepeatedAmount = true;
          } else {
            totalAmount += vAmounts.reduce((s, a) => s + a, 0);
          }
        }
      } else {
        // Cùng 1 hóa đơn/chứng từ (hoặc không có số hóa đơn/chứng từ)
        const nonZero = amounts.filter(a => a > 0);
        const allAmountsSame = nonZero.length > 0 && nonZero.every(a => a === nonZero[0]);

        if (allAmountsSame) {
          // Lặp lại cùng 1 số tiền tổng trên nhiều dòng của đơn -> Chỉ lấy 1 lần!
          totalAmount = nonZero[0];
          isRepeatedAmount = true;
        } else {
          // Các dòng chi tiết có số tiền khác nhau (mỗi dòng 1 sản phẩm) -> SUM
          totalAmount = amounts.reduce((sum, a) => sum + a, 0);
          if (distinctInvoices.length === 0 && distinctVouchers.length === 0) {
            isAmbiguous = true;
            warnings.push("Mã đơn có nhiều dòng chi tiết không kèm số hóa đơn phân biệt - đã tính tổng cộng");
          }
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
