/**
 * Automated Verification Script for E-Commerce Reconciliation Test Cases
 */
import { generateDemoEcommerceWorkbooks } from "./mockEcommerceData";
import { parseMarketplaceWorkbook } from "./shopeeParser";
import { parseSalesSoftwareWorkbook } from "./salesSoftwareParser";
import { buildProposedShopMappings } from "./shopMapper";
import { aggregateMarketplaceOrders, aggregateSalesOrders } from "./aggregator";
import { reconcileEcommerceRevenue } from "./reconciliationEngine";

export function runAllTestCases(): { success: boolean; results: string[]; failedCount: number } {
  const results: string[] = [];
  let failedCount = 0;

  function assert(condition: boolean, testName: string, detail?: string) {
    if (condition) {
      results.push(`✅ [PASS] ${testName}`);
    } else {
      failedCount++;
      results.push(`❌ [FAIL] ${testName} - ${detail || ""}`);
    }
  }

  const { shopeeWorkbook, shopeeFileName, salesWorkbook, salesFileName } = generateDemoEcommerceWorkbooks();

  // 1. Parse both workbooks
  const parsedMarketplace = parseMarketplaceWorkbook(shopeeWorkbook, shopeeFileName);
  const parsedSales = parseSalesSoftwareWorkbook(salesWorkbook, salesFileName);

  assert(parsedMarketplace.sheets.length >= 3, "Parsed marketplace sheets", `Sheets: ${parsedMarketplace.sheets.length}`);
  assert(parsedSales.detectedShops.length >= 3, "Detected sales shops", `Shops: ${parsedSales.detectedShops.join(", ")}`);

  // 2. Shop mapping
  const proposedMappings = buildProposedShopMappings(
    parsedMarketplace.allDetectedShopNames,
    parsedSales.detectedShops
  );

  // Case 7: Tên sheet và Tên shop không giống hoàn toàn
  const metakyMapping = proposedMappings.find(m => m.marketplaceSheetOrShop === "Metaky");
  assert(metakyMapping?.salesShopName === "Metaky Store", "Case 7: Sheet 'Metaky' mapped to 'Metaky Store'");

  const msmartyMapping = proposedMappings.find(m => m.marketplaceSheetOrShop === "Msmarty");
  assert(msmartyMapping?.salesShopName === "M'Smarty Shop", "Case 7: Sheet 'Msmarty' mapped to 'M'Smarty Shop'");

  // 3. Aggregation
  const marketplaceOrders = aggregateMarketplaceOrders(parsedMarketplace.sheets, proposedMappings);
  const salesOrders = aggregateSalesOrders(parsedSales.rows, parsedSales.detectedMapping);

  // Case 2: Multi-SKU voucher shop repeated once
  const order002 = marketplaceOrders.get("metakystore::SPX260701_002");
  assert(order002 !== undefined, "Case 2: Found order SPX260701_002");
  if (order002) {
    assert(order002.totalBuyerPaid === 600000, "Case 2: Total buyer paid = 600.000", `Got ${order002.totalBuyerPaid}`);
    assert(order002.shopVoucher === 50000, "Case 2: Shop voucher deducted ONCE = 50.000", `Got ${order002.shopVoucher}`);
    assert(order002.calculatedRevenue === 550000, "Case 2: Calculated revenue = 550.000", `Got ${order002.calculatedRevenue}`);
  }

  // Case 9: Duplicate rows with repeated total amount on PM
  const order009 = salesOrders.get("mibabystore::SPX260701_009");
  assert(order009 !== undefined && order009.totalAmount === 250000, "Case 9: PM duplicate same voucher taken once (250.000)", `Got ${order009?.totalAmount}`);

  // Case 10: Multi-invoice distinct invoices summed
  const order010 = salesOrders.get("mibabystore::SPX260701_010");
  assert(order010 !== undefined && order010.totalAmount === 300000, "Case 10: PM multiple distinct invoices summed (300.000)", `Got ${order010?.totalAmount}`);

  // 4. Reconciliation
  const { records, summary } = reconcileEcommerceRevenue(marketplaceOrders, salesOrders, {
    marketplaceType: "SHOPEE",
    toleranceVnd: 1,
    periodLabel: "tháng 7",
    autoConfirmExactShopNames: true
  });

  // Case 1: 1 order / 1 SKU / khớp hoàn toàn
  const rec001 = records.find(r => r.orderId === "SPX260701_001");
  assert(rec001?.status === "KHỚP" && rec001.marketplaceRevenue === 150000 && rec001.salesTotalAmount === 150000, "Case 1: Khớp hoàn toàn 150.000");

  // Case 2: Multi-SKU khớp
  const rec002 = records.find(r => r.orderId === "SPX260701_002");
  assert(rec002?.status === "KHỚP" && rec002.marketplaceRevenue === 550000 && rec002.salesTotalAmount === 550000, "Case 2: Multi-SKU khớp doanh thu 550.000");

  // Case 3: Sàn có - PM không có
  const rec003 = records.find(r => r.orderId === "SPX260701_003");
  assert(rec003?.status === "SÀN CÓ - PM KHÔNG CÓ", "Case 3: Sàn có - PM không có");

  // Case 4: PM có - Sàn không có
  const rec004 = records.find(r => r.orderId === "SPX260701_004");
  assert(rec004?.status === "PM CÓ - SÀN KHÔNG CÓ" && rec004.salesTotalAmount === 500000, "Case 4: PM có - Sàn không có");

  // Case 5: Đơn hủy không có trên PM
  const rec005 = records.find(r => r.orderId === "SPX260701_005");
  assert(rec005?.status === "ĐƠN HỦY - KHÔNG CÓ TRÊN PM", "Case 5: Đơn hủy không có trên PM");

  // Case 6: Cùng mã đơn nhưng khác shop không bị ghép chéo
  const commonMetaky = records.find(r => r.orderId === "SPX260701_COMMON" && r.shopDisplayName === "Metaky Store");
  const commonMibaby = records.find(r => r.orderId === "SPX260701_COMMON" && r.shopDisplayName === "Mibaby Store");
  assert(commonMetaky?.status === "KHỚP" && commonMetaky.marketplaceRevenue === 400000, "Case 6: Metaky COMMON matched 400.000");
  assert(commonMibaby?.status === "KHỚP" && commonMibaby.marketplaceRevenue === 700000, "Case 6: Mibaby COMMON matched 700.000 without cross-matching");

  // Case 8: Mã đơn chứa khoảng trắng
  const rec008 = records.find(r => r.orderId === "SPX260701_008");
  assert(rec008?.status === "KHỚP" && rec008.marketplaceRevenue === 600000, "Case 8: Khoảng trắng chuẩn hóa khớp 600.000");

  // Case 11: Voucher không nhất quán
  const rec011 = records.find(r => r.orderId === "SPX260701_011");
  assert(rec011?.status === "CẦN KIỂM TRA", "Case 11: Voucher không nhất quán flagged CẦN KIỂM TRA");

  // Case 12 & 13: Shop mới ABC Store tự sinh
  const recAbc = records.find(r => r.orderId === "SPX260701_ABC01");
  assert(recAbc?.status === "KHỚP" && recAbc.marketplaceRevenue === 120000, "Case 12 & 13: Shop mới tự sinh và khớp 120.000");

  // Difference case
  const recDiff = records.find(r => r.orderId === "SPX260701_DIFF");
  assert(recDiff?.status === "CHÊNH LỆCH" && recDiff.revenueDifference === 50000, "Chênh lệch doanh thu +50.000");

  // Check summary statistics
  assert(summary.matchedCount >= 5, "Summary matched count accurate", `Matched: ${summary.matchedCount}`);
  assert(summary.differenceCount >= 1, "Summary difference count accurate", `Difference: ${summary.differenceCount}`);
  assert(summary.marketplaceOnlyCount >= 1, "Summary marketplace only accurate", `Sàn có: ${summary.marketplaceOnlyCount}`);
  assert(summary.salesOnlyCount >= 1, "Summary sales only accurate", `PM có: ${summary.salesOnlyCount}`);
  assert(summary.anomaliesCount >= 1, "Summary anomalies accurate", `Anomalies: ${summary.anomaliesCount}`);

  return {
    success: failedCount === 0,
    results,
    failedCount
  };
}

// Nếu chạy trực tiếp từ CLI
if (typeof process !== "undefined" && process.argv[1]?.includes("verifyCases")) {
  const result = runAllTestCases();
  console.log(result.results.join("\n"));
  console.log(`\nTổng kết: ${result.failedCount === 0 ? "TẤT CẢ TEST CASES ĐÃ ĐẠT 100%!" : `${result.failedCount} test thất bại`}`);
  if (!result.success) process.exit(1);
}
