/**
 * Multi-Sheet Excel Exporter for E-Commerce Marketplace Revenue Reconciliation
 * Section 11, 12, 13, 25
 */
import * as XLSX from "xlsx";
import { ReconciliationRecord, ReconciliationSummary } from "./types";
import { sanitizeExcelSheetName } from "./orderNormalizer";

export function exportEcommerceReconciliationExcel(
  records: ReconciliationRecord[],
  summary: ReconciliationSummary,
  periodLabel: string = "kỳ này",
  fileNamePrefix: string = "Doi_chieu_DT_Shopee"
): void {
  const wb = XLSX.utils.book_new();

  // Nhóm các records theo shop
  const recordsByShop: { [shopName: string]: ReconciliationRecord[] } = {};
  const pmOnlyRecords: ReconciliationRecord[] = [];
  const anomalyRecords: ReconciliationRecord[] = [];

  for (const rec of records) {
    if (rec.status === "PM CÓ - SÀN KHÔNG CÓ") {
      pmOnlyRecords.push(rec);
    } else {
      const shop = rec.shopDisplayName || "Shop_Khac";
      if (!recordsByShop[shop]) {
        recordsByShop[shop] = [];
      }
      recordsByShop[shop].push(rec);
    }

    if (rec.status === "CẦN KIỂM TRA") {
      anomalyRecords.push(rec);
    }
  }

  // 1. TẠO CÁC SHEET CHO TỪNG SHOP (Bố cục chuẩn theo file mẫu: DATA BÁN HÀNG | PHẦN MỀM BÁN HÀNG | ĐỐI CHIẾU)
  const usedSheetNames = new Set<string>();

  for (const [rawShopName, shopRecords] of Object.entries(recordsByShop)) {
    let sheetName = sanitizeExcelSheetName(rawShopName, "Shop");
    let counter = 1;
    while (usedSheetNames.has(sheetName.toLowerCase())) {
      sheetName = `${sanitizeExcelSheetName(rawShopName).substring(0, 27)}_${counter++}`;
    }
    usedSheetNames.add(sheetName.toLowerCase());

    // Cấu trúc 3 cấp Header giống file mẫu Đối chiếu DT_Shopee_T7.8.xlsx
    const rowsForSheet: any[] = [
      // Dòng 1: Header cấp 1 (Nhóm cột)
      [
        "DATA BÁN HÀNG (SÀN TMĐT)", "", "", "", "",
        "PHẦN MỀM BÁN HÀNG", "", "", "",
        "ĐỐI CHIẾU", "", ""
      ],
      // Dòng 2: Header cấp 2 (Chi tiết từng cột)
      [
        "Mã đơn hàng",
        "Trạng Thái Đơn Hàng",
        "Thời gian giao hàng",
        "Trạng thái Trả hàng/Hoàn tiền",
        "TỔNG DT XUẤT HĐ",
        "Mã đơn",
        "Số hóa đơn",
        "Ngày hóa đơn",
        "Tổng tiền",
        "Chênh lệch DT",
        "Note",
        `Số tiền đã xuất trong ${periodLabel}`
      ]
    ];

    // Dòng dữ liệu
    shopRecords.forEach(r => {
      rowsForSheet.push([
        // SÀN
        String(r.originalOrderId || r.orderId), // giữ nguyên dạng text
        r.marketplaceStatus || "",
        r.marketplaceDeliveryDate || "",
        r.marketplaceReturnRefund || "",
        r.hasMarketplace ? r.marketplaceRevenue : 0,

        // PHẦN MỀM
        r.hasSales ? String(r.originalOrderId || r.orderId) : "",
        r.salesInvoiceNumber || "",
        r.salesInvoiceDate || "",
        r.hasSales ? r.salesTotalAmount : 0,

        // ĐỐI CHIẾU
        r.revenueDifference,
        r.note || r.status,
        r.periodIssuedAmount !== undefined ? r.periodIssuedAmount : ""
      ]);
    });

    const ws = XLSX.utils.aoa_to_sheet(rowsForSheet);

    // Merge header cấp 1
    ws["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 4 } }, // DATA BÁN HÀNG (col 0 - 4)
      { s: { r: 0, c: 5 }, e: { r: 0, c: 8 } }, // PHẦN MỀM BÁN HÀNG (col 5 - 8)
      { s: { r: 0, c: 9 }, e: { r: 0, c: 11 } } // ĐỐI CHIẾU (col 9 - 11)
    ];

    // Cài đặt độ rộng cột
    ws["!cols"] = [
      { wch: 22 }, // Mã đơn hàng
      { wch: 18 }, // Trạng Thái Đơn Hàng
      { wch: 18 }, // Thời gian giao hàng
      { wch: 20 }, // Trả hàng / hoàn tiền
      { wch: 18 }, // TỔNG DT XUẤT HĐ
      { wch: 20 }, // Mã đơn PM
      { wch: 14 }, // Số hóa đơn
      { wch: 14 }, // Ngày hóa đơn
      { wch: 18 }, // Tổng tiền PM
      { wch: 18 }, // Chênh lệch DT
      { wch: 30 }, // Note
      { wch: 22 }  // Số tiền đã xuất
    ];

    XLSX.utils.book_append_sheet(wb, ws, sheetName);
  }

  // 2. SHEET PHỤ ĐỐI CHIẾU NGƯỢC: "PM không có trên Sàn" (Section 13)
  if (pmOnlyRecords.length > 0) {
    let sheetName = "PM khong co tren San";
    const pmRows: any[] = [
      [
        "Tên Shop",
        "Mã đơn hàng",
        "Số hóa đơn",
        "Ngày hóa đơn",
        "Tổng tiền",
        "Số chứng từ",
        "Ghi chú",
        "Trạng thái"
      ]
    ];

    pmOnlyRecords.forEach(r => {
      pmRows.push([
        r.shopDisplayName,
        String(r.originalOrderId || r.orderId),
        r.salesInvoiceNumber,
        r.salesInvoiceDate,
        r.salesTotalAmount,
        r.salesVoucherNumber,
        r.note,
        r.status
      ]);
    });

    const wsPm = XLSX.utils.aoa_to_sheet(pmRows);
    wsPm["!cols"] = [
      { wch: 22 },
      { wch: 24 },
      { wch: 16 },
      { wch: 14 },
      { wch: 18 },
      { wch: 18 },
      { wch: 32 },
      { wch: 20 }
    ];
    XLSX.utils.book_append_sheet(wb, wsPm, sheetName);
  }

  // 3. SHEET DỮ LIỆU CẦN KIỂM TRA (nếu có bất thường)
  if (anomalyRecords.length > 0) {
    const anomalyRows: any[] = [
      [
        "Tên Shop",
        "Mã đơn hàng",
        "Vấn đề phát hiện",
        "DT Sàn",
        "Voucher Sàn",
        "Tổng tiền PM",
        "Chênh lệch",
        "Ghi chú chi tiết"
      ]
    ];

    anomalyRecords.forEach(r => {
      anomalyRows.push([
        r.shopDisplayName,
        String(r.originalOrderId || r.orderId),
        r.warningFlags.join(" | ") || r.note,
        r.marketplaceRevenue,
        r.marketplaceVoucher,
        r.salesTotalAmount,
        r.revenueDifference,
        r.note
      ]);
    });

    const wsAnomaly = XLSX.utils.aoa_to_sheet(anomalyRows);
    wsAnomaly["!cols"] = [
      { wch: 22 },
      { wch: 24 },
      { wch: 40 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 16 },
      { wch: 35 }
    ];
    XLSX.utils.book_append_sheet(wb, wsAnomaly, "Can kiem tra du lieu");
  }

  // 4. SHEET TỔNG QUAN ĐỐI CHIẾU (DASHBOARD SUMMARY)
  const summaryRows: any[] = [
    ["BÁO CÁO TỔNG HỢP ĐỐI CHIẾU DOANH THU SÀN TMĐT VÀ PHẦN MỀM BÁN HÀNG"],
    [`Kỳ đối chiếu: ${periodLabel}`],
    [],
    ["CHỈ SỐ TOÀN HỆ THỐNG", "SỐ LƯỢNG / GIÁ TRỊ"],
    ["Tổng số đơn trên sàn TMĐT", summary.totalMarketplaceOrders],
    ["Tổng số đơn trên phần mềm bán hàng", summary.totalSalesOrders],
    ["Tổng số đơn duy nhất hợp nhất", summary.totalUniqueOrdersCombined],
    ["Số đơn khớp hoàn toàn", summary.matchedCount],
    ["Số đơn có chênh lệch", summary.differenceCount],
    ["Số đơn có trên sàn nhưng thiếu trên PM", summary.marketplaceOnlyCount],
    ["Số đơn hủy (sàn có, PM không có)", summary.marketplaceOnlyCancelledCount],
    ["Số đơn có trên PM nhưng không có trên sàn", summary.salesOnlyCount],
    ["Số đơn có dữ liệu bất thường cần kiểm tra", summary.anomaliesCount],
    ["Tổng doanh thu sàn TMĐT", summary.totalMarketplaceRevenue],
    ["Tổng doanh thu phần mềm bán hàng", summary.totalSalesRevenue],
    ["Tổng chênh lệch doanh thu (Sàn - PM)", summary.totalDifferenceAmount],
    [],
    ["BẢNG TỔNG HỢP THEO TỪNG CỬA HÀNG (SHOP)"],
    [
      "Tên Shop",
      "Tổng đơn",
      "Đơn khớp",
      "Lệch tiền",
      "Thiếu trên PM",
      "Thiếu trên Sàn",
      "Đơn hủy",
      "Cần KT",
      "DT Sàn (VNĐ)",
      "DT Phần Mềm (VNĐ)",
      "Chênh lệch (VNĐ)"
    ]
  ];

  Object.values(summary.byShop).forEach(s => {
    summaryRows.push([
      s.shopName,
      s.totalOrders,
      s.matchedCount,
      s.differenceCount,
      s.marketplaceOnlyCount,
      s.salesOnlyCount,
      s.cancelledCount,
      s.anomaliesCount,
      s.marketplaceRevenue,
      s.salesRevenue,
      s.differenceAmount
    ]);
  });

  const wsSummary = XLSX.utils.aoa_to_sheet(summaryRows);
  wsSummary["!cols"] = [
    { wch: 32 },
    { wch: 14 },
    { wch: 14 },
    { wch: 14 },
    { wch: 16 },
    { wch: 16 },
    { wch: 12 },
    { wch: 12 },
    { wch: 18 },
    { wch: 18 },
    { wch: 18 }
  ];

  // Đặt sheet tổng quan lên đầu
  wb.SheetNames.unshift("Tong quan doi chieu");
  wb.Sheets["Tong quan doi chieu"] = wsSummary;

  // Xuất file
  const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
  XLSX.writeFile(wb, `${fileNamePrefix}_${dateStr}.xlsx`);
}
