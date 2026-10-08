/**
 * TypeScript Data Models for E-Commerce Marketplace Revenue Reconciliation
 * Phân hệ Đối Chiếu Doanh Thu Sàn TMĐT (Shopee, TikTok Shop, Lazada...)
 */

export type MarketplaceType = "SHOPEE" | "TIKTOK" | "LAZADA" | "OTHER";

export type ReconciliationStatus =
  | "KHỚP"
  | "CHÊNH LỆCH"
  | "SÀN CÓ - PM KHÔNG CÓ"
  | "PM CÓ - SÀN KHÔNG CÓ"
  | "ĐƠN HỦY - KHÔNG CÓ TRÊN PM"
  | "CẦN KIỂM TRA";

export interface MarketplaceRawRow {
  [key: string]: any;
}

export interface SalesSoftwareRawRow {
  [key: string]: any;
}

/**
 * Mapping cấu hình cột cho File Sàn TMĐT
 */
export interface MarketplaceColumnMapping {
  orderIdCol: string;         // Mã đơn hàng (Bắt buộc)
  buyerPaidCol: string;       // Tổng số tiền Người mua thanh toán (Shopee)
  shopVoucherCol: string;     // Mã giảm giá của Shop (Shopee)
  orderStatusCol: string;     // Trạng Thái Đơn Hàng
  deliveryTimeCol: string;    // Thời gian giao hàng
  returnRefundCol: string;    // Trạng thái Trả hàng/Hoàn tiền
  skuCol?: string;            // Tên sản phẩm / SKU
  priceCol?: string;          // Giá sản phẩm
  quantityCol?: string;       // Số lượng
}

/**
 * Mapping cấu hình cột cho File Phần Mềm Bán Hàng
 */
export interface SalesSoftwareColumnMapping {
  shopNameCol: string;        // Tên Shop / Cửa hàng (Bắt buộc)
  orderIdCol: string;         // Số đơn hàng từ hệ thống khác / Mã đơn (Bắt buộc)
  invoiceNumberCol: string;   // Số hóa đơn
  invoiceDateCol: string;     // Ngày hóa đơn
  totalAmountCol: string;     // Tổng tiền / Thành tiền sau thuế (Bắt buộc)
  voucherNumberCol?: string;  // Số chứng từ / Mã phiếu
  customerCol?: string;       // Khách hàng
  noteCol?: string;           // Ghi chú
}

/**
 * Cấu hình ghép Shop giữa Sàn và Phần mềm
 */
export interface ShopMappingEntry {
  marketplaceSheetOrShop: string; // Tên sheet trên file sàn hoặc shop sàn
  salesShopName: string;          // Tên Shop trên file phần mềm
  isConfirmed: boolean;           // Người dùng đã xác nhận
}

/**
 * Đơn hàng sau khi đã Group theo Shop + Mã Đơn từ File Sàn
 */
export interface AggregatedMarketplaceOrder {
  shopKey: string;                // Normalized shop key
  rawShopName: string;            // Tên shop / sheet gốc
  orderId: string;                // Normalized order ID
  originalOrderId: string;        // Mã đơn hàng gốc
  orderStatus: string;            // Trạng thái đơn hàng
  isCancelled: boolean;           // Có phải đơn hủy
  deliveryDate: string;           // Thời gian giao hàng
  returnRefundStatus: string;     // Trạng thái trả hàng/hoàn tiền
  totalBuyerPaid: number;         // Tổng các dòng Người mua thanh toán
  shopVoucher: number;            // Voucher của shop (chỉ trừ 1 lần)
  voucherInconsistent: boolean;   // Đánh dấu nếu voucher các dòng SKU không nhất quán
  voucherValuesFound: number[];   // Danh sách các giá trị voucher tìm thấy
  calculatedRevenue: number;      // Doanh thu tính toán = totalBuyerPaid - shopVoucher
  rowCount: number;               // Số dòng SKU
  sourceRows: MarketplaceRawRow[];// Các dòng gốc
  warnings: string[];             // Cảnh báo nếu có
}

/**
 * Đơn hàng sau khi đã Group theo Shop + Mã Đơn từ File Phần Mềm Bán Hàng
 */
export interface AggregatedSalesOrder {
  shopKey: string;                // Normalized shop key
  rawShopName: string;            // Tên Shop gốc
  orderId: string;                // Normalized order ID
  originalOrderId: string;        // Mã đơn hàng gốc
  invoiceNumbers: string[];       // Danh sách số hóa đơn
  invoiceDates: string[];         // Danh sách ngày hóa đơn
  voucherNumbers: string[];       // Danh sách số chứng từ
  totalAmount: number;            // Tổng tiền đối chiếu
  rowCount: number;               // Số dòng chi tiết
  isMultiInvoice: boolean;        // Có nhiều hóa đơn phân biệt
  isRepeatedAmount: boolean;      // Tiền bị lặp lại ở nhiều dòng của cùng chứng từ
  isAmbiguous: boolean;           // Không chắc chắn cách tính gộp
  sourceRows: SalesSoftwareRawRow[];
  warnings: string[];
}

/**
 * Kết quả đối chiếu của 1 đơn hàng (Record)
 */
export interface ReconciliationRecord {
  id: string;
  shopKey: string;
  shopDisplayName: string;        // Tên shop hiển thị (ưu tiên tên shop PM)
  orderId: string;                // Mã đơn hàng
  originalOrderId: string;

  // Dữ liệu sàn
  hasMarketplace: boolean;
  marketplaceStatus: string;
  marketplaceDeliveryDate: string;
  marketplaceReturnRefund: string;
  marketplaceRevenue: number;
  marketplaceVoucher: number;
  marketplaceBuyerPaid: number;
  marketplaceRowCount: number;
  isMarketplaceCancelled: boolean;

  // Dữ liệu phần mềm
  hasSales: boolean;
  salesInvoiceNumber: string;
  salesInvoiceDate: string;
  salesVoucherNumber: string;
  salesTotalAmount: number;
  salesRowCount: number;

  // Kết quả so khớp
  status: ReconciliationStatus;
  revenueDifference: number;      // Doanh thu sàn - Doanh thu PM
  toleranceApplied: number;
  note: string;
  periodIssuedAmount?: number;    // Số tiền đã xuất trong kỳ nếu xác định được
  warningFlags: string[];

  // Dữ liệu chi tiết phục vụ drill-down
  marketplaceData?: AggregatedMarketplaceOrder;
  salesData?: AggregatedSalesOrder;
}

/**
 * Tổng hợp thống kê đối chiếu
 */
export interface ReconciliationSummary {
  totalMarketplaceOrders: number;
  totalSalesOrders: number;
  totalUniqueOrdersCombined: number;

  matchedCount: number;
  differenceCount: number;
  marketplaceOnlyCount: number;
  marketplaceOnlyCancelledCount: number;
  salesOnlyCount: number;
  anomaliesCount: number;

  totalMarketplaceRevenue: number;
  totalSalesRevenue: number;
  totalDifferenceAmount: number;

  byShop: {
    [shopName: string]: {
      shopName: string;
      totalOrders: number;
      matchedCount: number;
      differenceCount: number;
      marketplaceOnlyCount: number;
      salesOnlyCount: number;
      cancelledCount: number;
      anomaliesCount: number;
      marketplaceRevenue: number;
      salesRevenue: number;
      differenceAmount: number;
    };
  };
}

/**
 * Cấu hình đối chiếu
 */
export interface ReconciliationConfig {
  marketplaceType: MarketplaceType;
  toleranceVnd: number;           // Mức chênh lệch cho phép (mặc định 1 đ)
  periodLabel: string;            // Ví dụ "tháng 7", "kỳ này"
  autoConfirmExactShopNames: boolean;
}
