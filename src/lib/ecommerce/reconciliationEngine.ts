/**
 * 2-Way Deterministic Reconciliation Engine for E-Commerce Marketplace Revenue
 * Section 3, 10, 15, 20, 26
 */
import {
  AggregatedMarketplaceOrder,
  AggregatedSalesOrder,
  ReconciliationConfig,
  ReconciliationRecord,
  ReconciliationStatus,
  ReconciliationSummary
} from "./types";
import { areAmountsEqual } from "./orderNormalizer";

export function reconcileEcommerceRevenue(
  marketplaceOrders: Map<string, AggregatedMarketplaceOrder>,
  salesOrders: Map<string, AggregatedSalesOrder>,
  config: ReconciliationConfig
): {
  records: ReconciliationRecord[];
  summary: ReconciliationSummary;
} {
  const records: ReconciliationRecord[] = [];
  const processedKeys = new Set<string>();
  const tolerance = config.toleranceVnd ?? 1;

  // ==========================================
  // CHIỀU 1: SÀN TMĐT -> PHẦN MỀM BÁN HÀNG
  // ==========================================
  for (const [compositeKey, mOrder] of marketplaceOrders.entries()) {
    processedKeys.add(compositeKey);
    const sOrder = salesOrders.get(compositeKey);

    const shopDisplayName = sOrder?.rawShopName || mOrder.rawShopName;
    const warningFlags: string[] = [...mOrder.warnings];
    if (sOrder?.warnings) {
      warningFlags.push(...sOrder.warnings);
    }

    if (sOrder) {
      // Đơn hàng có ở CẢ HAI NGUỒN
      const diff = mOrder.calculatedRevenue - sOrder.totalAmount;
      const isMatch = areAmountsEqual(mOrder.calculatedRevenue, sOrder.totalAmount, tolerance);

      let status: ReconciliationStatus = "KHỚP";
      let note = "";

      if (mOrder.voucherInconsistent || sOrder.isAmbiguous) {
        status = "CẦN KIỂM TRA";
        note = mOrder.voucherInconsistent
          ? "Dữ liệu mã giảm giá của Shop không nhất quán"
          : "Nhiều chứng từ phần mềm cần kiểm tra";
      } else if (isMatch) {
        status = "KHỚP";
        note = "Khớp doanh thu";
      } else {
        status = "CHÊNH LỆCH";
        note = diff > 0
          ? `Sàn cao hơn PM ${Math.abs(diff).toLocaleString("vi-VN")}đ`
          : `PM cao hơn Sàn ${Math.abs(diff).toLocaleString("vi-VN")}đ`;
      }

      records.push({
        id: `rec_${compositeKey}`,
        shopKey: mOrder.shopKey,
        shopDisplayName,
        orderId: mOrder.orderId,
        originalOrderId: mOrder.originalOrderId,

        hasMarketplace: true,
        marketplaceStatus: mOrder.orderStatus,
        marketplaceDeliveryDate: mOrder.deliveryDate,
        marketplaceReturnRefund: mOrder.returnRefundStatus,
        marketplaceRevenue: mOrder.calculatedRevenue,
        marketplaceVoucher: mOrder.shopVoucher,
        marketplaceBuyerPaid: mOrder.totalBuyerPaid,
        marketplaceRowCount: mOrder.rowCount,
        isMarketplaceCancelled: mOrder.isCancelled,

        hasSales: true,
        salesInvoiceNumber: sOrder.invoiceNumbers.join(", "),
        salesInvoiceDate: sOrder.invoiceDates.join(", "),
        salesVoucherNumber: sOrder.voucherNumbers.join(", "),
        salesTotalAmount: sOrder.totalAmount,
        salesRowCount: sOrder.rowCount,

        status,
        revenueDifference: diff,
        toleranceApplied: tolerance,
        note,
        warningFlags,
        marketplaceData: mOrder,
        salesData: sOrder
      });
    } else {
      // Đơn hàng CHỈ CÓ TRÊN SÀN - KHÔNG CÓ TRÊN PHẦN MỀM (Section 10)
      let status: ReconciliationStatus = "SÀN CÓ - PM KHÔNG CÓ";
      let note = "Không tìm thấy trên file bán hàng";

      if (mOrder.isCancelled) {
        status = "ĐƠN HỦY - KHÔNG CÓ TRÊN PM";
        note = "Đơn hủy - không có trên file bán hàng";
      }

      if (mOrder.voucherInconsistent) {
        status = "CẦN KIỂM TRA";
        note = "Dữ liệu mã giảm giá không nhất quán & Chưa có trên PM";
      }

      records.push({
        id: `rec_${compositeKey}`,
        shopKey: mOrder.shopKey,
        shopDisplayName,
        orderId: mOrder.orderId,
        originalOrderId: mOrder.originalOrderId,

        hasMarketplace: true,
        marketplaceStatus: mOrder.orderStatus,
        marketplaceDeliveryDate: mOrder.deliveryDate,
        marketplaceReturnRefund: mOrder.returnRefundStatus,
        marketplaceRevenue: mOrder.calculatedRevenue,
        marketplaceVoucher: mOrder.shopVoucher,
        marketplaceBuyerPaid: mOrder.totalBuyerPaid,
        marketplaceRowCount: mOrder.rowCount,
        isMarketplaceCancelled: mOrder.isCancelled,

        hasSales: false,
        salesInvoiceNumber: "",
        salesInvoiceDate: "",
        salesVoucherNumber: "",
        salesTotalAmount: 0,
        salesRowCount: 0,

        status,
        revenueDifference: mOrder.calculatedRevenue,
        toleranceApplied: tolerance,
        note,
        warningFlags,
        marketplaceData: mOrder
      });
    }
  }

  // ==========================================
  // CHIỀU 2: PHẦN MỀM BÁN HÀNG -> SÀN TMĐT (Section 10, 13, 26)
  // Tìm các đơn có trên PM nhưng chưa được xử lý (không tồn tại trên Sàn)
  // ==========================================
  for (const [compositeKey, sOrder] of salesOrders.entries()) {
    if (processedKeys.has(compositeKey)) continue;

    const shopDisplayName = sOrder.rawShopName;
    const warningFlags: string[] = [...sOrder.warnings];

    let status: ReconciliationStatus = "PM CÓ - SÀN KHÔNG CÓ";
    let note = "Có trên PM nhưng không tìm thấy trên sàn";

    if (sOrder.isAmbiguous) {
      status = "CẦN KIỂM TRA";
      note = "Đơn PM không có trên sàn & nhiều chứng từ chưa rõ";
    }

    records.push({
      id: `rec_${compositeKey}`,
      shopKey: sOrder.shopKey,
      shopDisplayName,
      orderId: sOrder.orderId,
      originalOrderId: sOrder.originalOrderId,

      hasMarketplace: false,
      marketplaceStatus: "",
      marketplaceDeliveryDate: "",
      marketplaceReturnRefund: "",
      marketplaceRevenue: 0,
      marketplaceVoucher: 0,
      marketplaceBuyerPaid: 0,
      marketplaceRowCount: 0,
      isMarketplaceCancelled: false,

      hasSales: true,
      salesInvoiceNumber: sOrder.invoiceNumbers.join(", "),
      salesInvoiceDate: sOrder.invoiceDates.join(", "),
      salesVoucherNumber: sOrder.voucherNumbers.join(", "),
      salesTotalAmount: sOrder.totalAmount,
      salesRowCount: sOrder.rowCount,

      status,
      revenueDifference: -sOrder.totalAmount,
      toleranceApplied: tolerance,
      note,
      warningFlags,
      salesData: sOrder
    });
  }

  // ==========================================
  // TỔNG HỢP SUMMARY & STATS (Section 15)
  // ==========================================
  let totalMarketplaceRevenue = 0;
  let totalSalesRevenue = 0;
  let totalDifferenceAmount = 0;

  let matchedCount = 0;
  let differenceCount = 0;
  let marketplaceOnlyCount = 0;
  let marketplaceOnlyCancelledCount = 0;
  let salesOnlyCount = 0;
  let anomaliesCount = 0;

  const byShopMap: ReconciliationSummary["byShop"] = {};

  for (const rec of records) {
    const shop = rec.shopDisplayName || "Chưa phân loại";
    if (!byShopMap[shop]) {
      byShopMap[shop] = {
        shopName: shop,
        totalOrders: 0,
        matchedCount: 0,
        differenceCount: 0,
        marketplaceOnlyCount: 0,
        salesOnlyCount: 0,
        cancelledCount: 0,
        anomaliesCount: 0,
        marketplaceRevenue: 0,
        salesRevenue: 0,
        differenceAmount: 0
      };
    }

    const s = byShopMap[shop];
    s.totalOrders++;

    if (rec.hasMarketplace) {
      totalMarketplaceRevenue += rec.marketplaceRevenue;
      s.marketplaceRevenue += rec.marketplaceRevenue;
    }
    if (rec.hasSales) {
      totalSalesRevenue += rec.salesTotalAmount;
      s.salesRevenue += rec.salesTotalAmount;
    }
    totalDifferenceAmount += rec.revenueDifference;
    s.differenceAmount += rec.revenueDifference;

    switch (rec.status) {
      case "KHỚP":
        matchedCount++;
        s.matchedCount++;
        break;
      case "CHÊNH LỆCH":
        differenceCount++;
        s.differenceCount++;
        break;
      case "SÀN CÓ - PM KHÔNG CÓ":
        marketplaceOnlyCount++;
        s.marketplaceOnlyCount++;
        break;
      case "ĐƠN HỦY - KHÔNG CÓ TRÊN PM":
        marketplaceOnlyCancelledCount++;
        s.cancelledCount++;
        break;
      case "PM CÓ - SÀN KHÔNG CÓ":
        salesOnlyCount++;
        s.salesOnlyCount++;
        break;
      case "CẦN KIỂM TRA":
        anomaliesCount++;
        s.anomaliesCount++;
        break;
    }
  }

  const summary: ReconciliationSummary = {
    totalMarketplaceOrders: marketplaceOrders.size,
    totalSalesOrders: salesOrders.size,
    totalUniqueOrdersCombined: records.length,

    matchedCount,
    differenceCount,
    marketplaceOnlyCount,
    marketplaceOnlyCancelledCount,
    salesOnlyCount,
    anomaliesCount,

    totalMarketplaceRevenue,
    totalSalesRevenue,
    totalDifferenceAmount,

    byShop: byShopMap
  };

  return { records, summary };
}
