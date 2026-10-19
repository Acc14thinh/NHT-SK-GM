import React, { useState, useMemo, useEffect, useRef } from "react";
import * as XLSX from "xlsx";
import {
  FileSpreadsheet,
  Upload,
  AlertTriangle,
  CheckCircle2,
  XCircle,
  HelpCircle,
  Download,
  Play,
  RotateCcw,
  Store,
  Layers,
  Search,
  Filter,
  Eye,
  SlidersHorizontal,
  ChevronDown,
  ChevronUp,
  TrendingDown,
  TrendingUp,
  Sparkles,
  Check,
  RefreshCw
} from "lucide-react";
import {
  MarketplaceType,
  ReconciliationConfig,
  ReconciliationRecord,
  ReconciliationSummary,
  ShopMappingEntry,
  getColumnLetter,
  MarketplaceColumnMapping,
  SalesSoftwareColumnMapping,
  MarketplaceInspectionResult,
  SalesSoftwareInspectionResult
} from "../../lib/ecommerce/types";
import {
  inspectMarketplaceWorkbook,
  parseMarketplaceOrdersForReconciliation,
  autoDetectShopeeColumns,
  isExactShopeeHeader
} from "../../lib/ecommerce/shopeeParser";
import {
  inspectSalesSoftwareWorkbook,
  parseSalesOrdersForReconciliation,
  autoDetectSalesSoftwareColumns,
  isExactSalesHeader
} from "../../lib/ecommerce/salesSoftwareParser";
import {
  buildProposedShopMappings,
  saveShopMappings,
  saveHeaderMappings
} from "../../lib/ecommerce/shopMapper";
import { reconcileEcommerceRevenue } from "../../lib/ecommerce/reconciliationEngine";
import { exportEcommerceReconciliationExcel } from "../../lib/ecommerce/excelExporter";
import { generateDemoEcommerceWorkbooks } from "../../lib/ecommerce/mockEcommerceData";

export const EcommerceReconciliationView: React.FC = () => {
  // Config state
  const [config, setConfig] = useState<ReconciliationConfig>({
    marketplaceType: "SHOPEE",
    toleranceVnd: 1,
    periodLabel: "tháng 7",
    autoConfirmExactShopNames: true
  });

  // Collapsible Advanced Settings (Mặc định thu gọn theo Section 34)
  const [showAdvancedSettings, setShowAdvancedSettings] = useState<boolean>(false);
  const [showColumnMappingModal, setShowColumnMappingModal] = useState<boolean>(false);
  const [mappingModalTab, setMappingModalTab] = useState<"shopee" | "sales">("shopee");
  const [showShopMappingModal, setShowShopMappingModal] = useState<boolean>(false);

  // Files & Parsed data (Chỉ giữ metadata nhẹ trong state, KHÔNG giữ mảng hàng chục nghìn rows)
  const [marketplaceResult, setMarketplaceResult] = useState<MarketplaceInspectionResult | null>(null);
  const [salesResult, setSalesResult] = useState<SalesSoftwareInspectionResult | null>(null);
  const [shopMappings, setShopMappings] = useState<ShopMappingEntry[]>([]);

  // References giữ file gốc và workbook demo tạm thời để Phase 2 xử lý
  const mktFileRef = useRef<File | null>(null);
  const salesFileRef = useRef<File | null>(null);
  const mktDemoWbRef = useRef<XLSX.WorkBook | null>(null);
  const salesDemoWbRef = useRef<XLSX.WorkBook | null>(null);

  // Reconciliation processing & progress (Section 40)
  const [isProcessing, setIsProcessing] = useState<boolean>(false);
  const [processingProgress, setProcessingProgress] = useState<string>("");
  const [reconciliationRecords, setReconciliationRecords] = useState<ReconciliationRecord[] | null>(null);
  const [reconciliationSummary, setReconciliationSummary] = useState<ReconciliationSummary | null>(null);

  // Filter & Search
  const [selectedShopFilter, setSelectedShopFilter] = useState<string>("ALL");
  const [statusFilter, setStatusFilter] = useState<string>("ALL");
  const [searchQuery, setSearchQuery] = useState<string>("");
  const [selectedRecordDetail, setSelectedRecordDetail] = useState<ReconciliationRecord | null>(null);

  // Notification Toast
  const [toastMessage, setToastMessage] = useState<{ text: string; type: "success" | "warning" | "info" } | null>(null);

  const mktFileInputRef = useRef<HTMLInputElement>(null);
  const salesFileInputRef = useRef<HTMLInputElement>(null);

  const showToast = (text: string, type: "success" | "warning" | "info" = "success") => {
    setToastMessage({ text, type });
    setTimeout(() => setToastMessage(null), 3500);
  };

  // Xem trước giá trị mẫu 3-5 dòng cho cột được chọn trong file Shopee (Lấy trực tiếp từ previewValuesByCol)
  const getShopeePreviewValues = (colName: string): string[] => {
    if (!colName || !marketplaceResult?.sheets[0]?.previewValuesByCol) return [];
    return marketplaceResult.sheets[0].previewValuesByCol[colName] || [];
  };

  // Xem trước giá trị mẫu 3-5 dòng cho cột được chọn trong file Phần Mềm Bán Hàng
  const getSalesPreviewValues = (colName: string): string[] => {
    if (!colName || !salesResult?.previewValuesByCol) return [];
    return salesResult.previewValuesByCol[colName] || [];
  };

  // Cập nhật trường mapping cho file Shopee (đồng bộ trên tất cả các sheet trong workbook)
  const handleUpdateShopeeMapping = (field: keyof MarketplaceColumnMapping, val: string) => {
    if (!marketplaceResult) return;
    const updated = { ...marketplaceResult };
    updated.sheets = updated.sheets.map(s => ({
      ...s,
      detectedMapping: {
        ...s.detectedMapping,
        [field]: val
      }
    }));
    setMarketplaceResult(updated);
    saveHeaderMappings("marketplace", { [field]: val });
  };

  // Cập nhật trường mapping cho file Phần mềm bán hàng
  const handleUpdateSalesMapping = (field: keyof SalesSoftwareColumnMapping, val: string) => {
    if (!salesResult) return;
    const updated = {
      ...salesResult,
      detectedMapping: {
        ...salesResult.detectedMapping,
        [field]: val
      }
    };
    setSalesResult(updated);
    saveHeaderMappings("sales", { [field]: val });
  };

  // Tự động nhận diện lại toàn bộ các cột từ headers thực tế
  const handleResetAutoDetection = () => {
    if (marketplaceResult && marketplaceResult.sheets.length > 0) {
      const updated = { ...marketplaceResult };
      updated.sheets = updated.sheets.map(s => ({
        ...s,
        detectedMapping: autoDetectShopeeColumns(s.headers)
      }));
      setMarketplaceResult(updated);
    }
    if (salesResult) {
      setSalesResult({
        ...salesResult,
        detectedMapping: autoDetectSalesSoftwareColumns(salesResult.headers)
      });
    }
    showToast("Đã tự động nhận diện lại cấu hình cột từ các file!", "info");
  };

  // Helper lấy ký tự cột + tên cột cho Shopee
  const getShopeeColLetterAndName = (colName?: string) => {
    if (!colName || !marketplaceResult?.sheets[0]?.headers) return "";
    const idx = marketplaceResult.sheets[0].headers.indexOf(colName);
    return idx >= 0 ? `${getColumnLetter(idx)} — ${colName}` : colName;
  };

  // Helper lấy ký tự cột + tên cột cho Sales
  const getSalesColLetterAndName = (colName?: string) => {
    if (!colName || !salesResult?.headers) return "";
    const idx = salesResult.headers.indexOf(colName);
    return idx >= 0 ? `${getColumnLetter(idx)} — ${colName}` : colName;
  };

  // Trạng thái kiểm tra 6 cột File Sàn Shopee (Section 1, 3, 4)
  const shopeeChecks = useMemo(() => {
    if (!marketplaceResult || !marketplaceResult.sheets || !marketplaceResult.sheets[0]) return null;
    const m = marketplaceResult.sheets[0].detectedMapping;
    if (!m) return null;
    const items = [
      {
        key: "orderIdCol",
        label: "Mã đơn hàng",
        header: m.orderIdCol,
        isMapped: Boolean(m.orderIdCol)
      },
      {
        key: "orderStatusCol",
        label: "Trạng thái đơn",
        header: m.orderStatusCol,
        isMapped: Boolean(m.orderStatusCol)
      },
      {
        key: "returnRefundCol",
        label: "Trạng thái Trả hàng/Hoàn tiền",
        header: m.returnRefundCol,
        isMapped: Boolean(m.returnRefundCol)
      },
      {
        key: "deliveryTimeCol",
        label: "Thời gian giao hàng",
        header: m.deliveryTimeCol,
        isMapped: Boolean(m.deliveryTimeCol)
      },
      {
        key: "buyerPaidCol",
        label: "Tổng tiền người mua thanh toán",
        header: m.buyerPaidCol,
        isMapped: Boolean(m.buyerPaidCol),
        isExact: m.buyerPaidCol === "Tổng số tiền Người mua thanh toán"
      },
      {
        key: "shopVoucherCol",
        label: "Mã giảm giá Shop",
        header: m.shopVoucherCol,
        isMapped: Boolean(m.shopVoucherCol)
      }
    ];
    const mappedCount = items.filter(i => i.isMapped).length;
    // Kiểm tra riêng cột tiền có bị không chắc chắn không (Section 4)
    const hasUncertainBuyerPaid = !m.buyerPaidCol || m.buyerPaidCol !== "Tổng số tiền Người mua thanh toán";
    return { items, mappedCount, total: 6, hasUncertainBuyerPaid };
  }, [marketplaceResult]);

  // Trạng thái kiểm tra 6 cột File Phần Mềm Bán Hàng (Section 1, 5)
  const salesChecks = useMemo(() => {
    if (!salesResult || !salesResult.detectedMapping) return null;
    const m = salesResult.detectedMapping;
    const items = [
      {
        key: "shopNameCol",
        label: "Tên Shop",
        header: m.shopNameCol,
        isMapped: Boolean(m.shopNameCol)
      },
      {
        key: "orderIdCol",
        label: "Mã đơn từ sàn",
        header: m.orderIdCol,
        isMapped: Boolean(m.orderIdCol)
      },
      {
        key: "invoiceNumberCol",
        label: "Số hóa đơn",
        header: m.invoiceNumberCol,
        isMapped: Boolean(m.invoiceNumberCol)
      },
      {
        key: "invoiceDateCol",
        label: "Ngày hạch toán",
        header: m.invoiceDateCol,
        isMapped: Boolean(m.invoiceDateCol)
      },
      {
        key: "totalAmountCol",
        label: "Tổng tiền thanh toán",
        header: m.totalAmountCol,
        isMapped: Boolean(m.totalAmountCol)
      },
      {
        key: "voucherNumberCol",
        label: "Số chứng từ",
        header: m.voucherNumberCol || m.invoiceNumberCol,
        isMapped: Boolean(m.voucherNumberCol || m.invoiceNumberCol)
      }
    ];
    const mappedCount = items.filter(i => i.isMapped).length;
    return { items, mappedCount, total: 6 };
  }, [salesResult]);

  // Tự động nhận diện và ghép shop khi có đủ 2 nguồn (Section 32, 35)
  useEffect(() => {
    if (marketplaceResult && salesResult) {
      const proposed = buildProposedShopMappings(
        marketplaceResult.allDetectedShopNames || [],
        salesResult.detectedShops || []
      );
      setShopMappings(proposed);
    }
  }, [marketplaceResult, salesResult]);

  // Kiểm tra tính sẵn sàng & lỗi chặn (Section 39)
  const validationStatus = useMemo(() => {
    if (!marketplaceResult || !salesResult) {
      return { isReady: false, blockingError: null, warningText: null };
    }

    // 1. Kiểm tra các cột bắt buộc trên file Shopee
    const shopeeMapping = marketplaceResult.sheets?.[0]?.detectedMapping;
    if (!shopeeMapping?.orderIdCol) {
      return {
        isReady: false,
        blockingError: "Chưa cấu hình cột [Mã đơn hàng] trong file Shopee.",
        actionType: "column_mkt" as const
      };
    }
    if (!shopeeMapping?.buyerPaidCol) {
      return {
        isReady: false,
        blockingError: "Chưa cấu hình cột [Tổng số tiền Người mua thanh toán] trong file Shopee.",
        actionType: "column_mkt" as const
      };
    }
    if (!shopeeMapping?.returnRefundCol) {
      return {
        isReady: false,
        blockingError: "Chưa cấu hình cột [Trạng thái Trả hàng/Hoàn tiền] trong file Shopee.",
        actionType: "column_mkt" as const
      };
    }

    // 2. Kiểm tra các cột bắt buộc trên file Phần mềm bán hàng
    if (!salesResult.detectedMapping?.shopNameCol) {
      return {
        isReady: false,
        blockingError: "Chưa cấu hình cột [Tên Shop] trong file phần mềm bán hàng.",
        actionType: "column_sales" as const
      };
    }
    if (!salesResult.detectedMapping?.orderIdCol) {
      return {
        isReady: false,
        blockingError: "Chưa cấu hình cột [Mã đơn từ sàn] trong file phần mềm bán hàng.",
        actionType: "column_sales" as const
      };
    }
    if (!salesResult.detectedMapping?.totalAmountCol) {
      return {
        isReady: false,
        blockingError: "Chưa cấu hình cột [Tổng tiền] trong file phần mềm bán hàng.",
        actionType: "column_sales" as const
      };
    }

    // 3. Kiểm tra ghép Shop
    if (
      (marketplaceResult.sheets?.length || 0) > 1 &&
      shopMappings.length > 0 &&
      shopMappings.some(m => !m.salesShopName)
    ) {
      return {
        isReady: false,
        blockingError: "Có cửa hàng sàn TMĐT chưa ghép được với Tên Shop trên phần mềm bán hàng.",
        actionType: "shop" as const
      };
    }

    return { isReady: true, blockingError: null, warningText: null };
  }, [marketplaceResult, salesResult, shopMappings]);

  // Xử lý nạp File Sàn Shopee (Tự động phân tích ngầm ngay khi chọn file)
  // Xử lý nạp File Sàn TMĐT (PHASE 1: Lightweight inspection chỉ 25 dòng đầu, cực nhanh, không tốn RAM)
  const handleMarketplaceFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Lưu reference file gốc (0 MB heap overhead)
    mktFileRef.current = file;
    mktDemoWbRef.current = null;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        // Đọc chỉ 25 dòng đầu tiên để lấy sheetNames, headers, platform và preview
        const lightWb = XLSX.read(buffer, { type: "array", sheetRows: 25 });
        const inspected = inspectMarketplaceWorkbook(lightWb, file.name, file.size);
        setMarketplaceResult(inspected);
        setReconciliationRecords(null);
        showToast(`Đã nhận diện file sàn TMĐT: ${inspected.sheets.length} shop/sheet`);
      } catch (err) {
        console.error("[UPLOAD MKT ERROR]", err);
        showToast("Lỗi khi đọc file sàn TMĐT. Vui lòng kiểm tra định dạng Excel.", "warning");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Xử lý nạp File Phần Mềm Bán Hàng (PHASE 1: Lightweight inspection chỉ 25 dòng đầu)
  const handleSalesFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    // Lưu reference file gốc
    salesFileRef.current = file;
    salesDemoWbRef.current = null;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result as ArrayBuffer;
        // Đọc chỉ 25 dòng đầu tiên để lấy headers, mapping và preview
        const lightWb = XLSX.read(buffer, { type: "array", sheetRows: 25 });
        const inspected = inspectSalesSoftwareWorkbook(lightWb, file.name, file.size);
        setSalesResult(inspected);
        setReconciliationRecords(null);
        showToast(`Đã nhận diện file bán hàng: ${inspected.detectedShops.length} shop`);
      } catch (err) {
        console.error("[UPLOAD SALES ERROR]", err);
        showToast("Lỗi khi đọc file phần mềm bán hàng.", "warning");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 1-Click Nạp Dữ Liệu Demo Mẫu (14 trường hợp nghiệp vụ)
  const handleLoadDemoData = () => {
    mktFileRef.current = null;
    salesFileRef.current = null;
    const { shopeeWorkbook, shopeeFileName, salesWorkbook, salesFileName } = generateDemoEcommerceWorkbooks();
    mktDemoWbRef.current = shopeeWorkbook;
    salesDemoWbRef.current = salesWorkbook;

    const parsedM = inspectMarketplaceWorkbook(shopeeWorkbook, shopeeFileName);
    const parsedS = inspectSalesSoftwareWorkbook(salesWorkbook, salesFileName);

    setMarketplaceResult(parsedM);
    setSalesResult(parsedS);

    const proposed = buildProposedShopMappings(
      parsedM.allDetectedShopNames,
      parsedS.detectedShops
    );
    setShopMappings(proposed);
    setReconciliationRecords(null);
    showToast("Đã nạp thành công bộ file mẫu Shopee & Phần mềm bán hàng!", "info");
  };

  // BƯỚC 2: Thực hiện đối chiếu (PHASE 2: FULL PROCESSING CHỌN LỌC)
  const handleExecuteReconciliation = () => {
    if (!marketplaceResult || !salesResult) {
      showToast("Vui lòng tải lên cả 2 file trước khi đối chiếu!", "warning");
      return;
    }

    if (!validationStatus.isReady && validationStatus.blockingError) {
      showToast(validationStatus.blockingError, "warning");
      return;
    }

    setIsProcessing(true);
    setProcessingProgress("Đang chuẩn bị nạp dữ liệu đối chiếu...");

    setTimeout(async () => {
      try {
        saveShopMappings(shopMappings);

        // 1. Nạp và trích xuất chọn lọc dữ liệu Sàn TMĐT
        setProcessingProgress("Đang đọc & trích xuất cột dữ liệu sàn TMĐT...");
        let mktWb: XLSX.WorkBook;
        if (mktFileRef.current) {
          const mktBuf = await mktFileRef.current.arrayBuffer();
          mktWb = XLSX.read(mktBuf, { type: "array" });
        } else if (mktDemoWbRef.current) {
          mktWb = mktDemoWbRef.current;
        } else {
          throw new Error("Không tìm thấy tệp sàn TMĐT.");
        }

        const aggregatedMarketplace = parseMarketplaceOrdersForReconciliation(
          mktWb,
          marketplaceResult.sheets,
          shopMappings
        );

        // 2. Nạp và trích xuất chọn lọc dữ liệu Phần Mềm Bán Hàng
        setProcessingProgress("Đang đọc & trích xuất cột dữ liệu phần mềm bán hàng...");
        let salesWb: XLSX.WorkBook;
        if (salesFileRef.current) {
          const salesBuf = await salesFileRef.current.arrayBuffer();
          salesWb = XLSX.read(salesBuf, { type: "array" });
        } else if (salesDemoWbRef.current) {
          salesWb = salesDemoWbRef.current;
        } else {
          throw new Error("Không tìm thấy tệp phần mềm bán hàng.");
        }

        const aggregatedSales = parseSalesOrdersForReconciliation(
          salesWb,
          salesResult.sheetName,
          salesResult.detectedMapping
        );

        // 3. Thực hiện đối chiếu 2 chiều O(1) Map lookup
        setProcessingProgress("Đang thực hiện đối chiếu doanh thu 2 chiều...");
        const { records, summary } = reconcileEcommerceRevenue(
          aggregatedMarketplace,
          aggregatedSales,
          config
        );

        // 4. Giải phóng bộ nhớ tạm ngay lập tức
        aggregatedMarketplace.clear();
        aggregatedSales.clear();

        setReconciliationRecords(records);
        setReconciliationSummary(summary);
        setIsProcessing(false);
        setProcessingProgress("");
        showToast(`Đối chiếu hoàn tất: ${summary.matchedCount} đơn khớp, ${summary.differenceCount} đơn chênh lệch!`);
      } catch (err: any) {
        console.error("[RECONCILIATION ERROR]", err);
        setIsProcessing(false);
        setProcessingProgress("");
        showToast(`Lỗi trong quá trình đối chiếu: ${err.message || err}`, "warning");
      }
    }, 40);
  };

  // BƯỚC 3: Xuất kết quả Excel đa sheet
  const handleExportExcel = () => {
    if (!reconciliationRecords || !reconciliationSummary) {
      showToast("Chưa có kết quả đối chiếu để xuất Excel!", "warning");
      return;
    }
    exportEcommerceReconciliationExcel(
      reconciliationRecords,
      reconciliationSummary,
      config.periodLabel,
      "Doi_chieu_DT_Shopee"
    );
    showToast("Đã tải xuống file Excel đối chiếu doanh thu đa sheet!");
  };

  // Làm mới (Xóa 2 file & giải phóng memory, giữ nguyên mapping đã lưu trong localStorage - Section 43)
  const handleReset = () => {
    mktFileRef.current = null;
    salesFileRef.current = null;
    mktDemoWbRef.current = null;
    salesDemoWbRef.current = null;
    setMarketplaceResult(null);
    setSalesResult(null);
    setShopMappings([]);
    setReconciliationRecords(null);
    setReconciliationSummary(null);
    setSelectedShopFilter("ALL");
    setStatusFilter("ALL");
    setSearchQuery("");
    if (mktFileInputRef.current) mktFileInputRef.current.value = "";
    if (salesFileInputRef.current) salesFileInputRef.current.value = "";
    showToast("Đã làm mới phiên đối chiếu và giải phóng bộ nhớ.", "info");
  };

  // Danh sách các shop xuất hiện trong kết quả
  const availableShops = useMemo(() => {
    if (!reconciliationSummary) return [];
    return Object.keys(reconciliationSummary.byShop).sort();
  }, [reconciliationSummary]);

  // Bộ lọc kết quả
  const filteredRecords = useMemo(() => {
    if (!reconciliationRecords) return [];

    return reconciliationRecords.filter(rec => {
      // Lọc theo Shop
      if (selectedShopFilter !== "ALL") {
        if (selectedShopFilter === "PM_ONLY") {
          if (rec.status !== "PM CÓ - SÀN KHÔNG CÓ") return false;
        } else if (selectedShopFilter === "ANOMALY") {
          if (rec.status !== "CẦN KIỂM TRA") return false;
        } else if (rec.shopDisplayName !== selectedShopFilter) {
          return false;
        }
      }

      // Lọc theo trạng thái
      if (statusFilter !== "ALL") {
        if (statusFilter === "KHỚP" && rec.status !== "KHỚP") return false;
        if (statusFilter === "CHÊNH LỆCH" && rec.status !== "CHÊNH LỆCH") return false;
        if (statusFilter === "SÀN_THIẾU_PM" && rec.status !== "SÀN CÓ - PM KHÔNG CÓ") return false;
        if (statusFilter === "PM_THIẾU_SÀN" && rec.status !== "PM CÓ - SÀN KHÔNG CÓ") return false;
        if (statusFilter === "ĐƠN_HỦY" && rec.status !== "ĐƠN HỦY - KHÔNG CÓ TRÊN PM") return false;
        if (statusFilter === "CẦN_KT" && rec.status !== "CẦN KIỂM TRA") return false;
      }

      // Lọc theo tìm kiếm từ khóa
      if (searchQuery.trim()) {
        const q = searchQuery.toLowerCase().trim();
        const matchOrder = rec.originalOrderId.toLowerCase().includes(q) || rec.orderId.toLowerCase().includes(q);
        const matchInvoice = rec.salesInvoiceNumber.toLowerCase().includes(q);
        const matchShop = rec.shopDisplayName.toLowerCase().includes(q);
        const matchNote = rec.note.toLowerCase().includes(q);
        if (!matchOrder && !matchInvoice && !matchShop && !matchNote) return false;
      }

      return true;
    });
  }, [reconciliationRecords, selectedShopFilter, statusFilter, searchQuery]);

  return (
    <div className="space-y-6 animate-fade-in pb-16">
      {/* Toast Notification */}
      {toastMessage && (
        <div
          className={`fixed bottom-6 right-6 z-50 px-4 py-3 border-2 border-[#141414] shadow-[4px_4px_0px_#141414] text-xs font-black uppercase tracking-wider flex items-center gap-2 ${
            toastMessage.type === "success"
              ? "bg-[#00ff00] text-black"
              : toastMessage.type === "warning"
              ? "bg-amber-300 text-black"
              : "bg-cyan-200 text-black"
          }`}
        >
          {toastMessage.type === "success" ? <CheckCircle2 size={16} /> : <AlertTriangle size={16} />}
          <span>{toastMessage.text}</span>
        </div>
      )}

      {/* HEADER BANNER */}
      <div className="bg-[#141414] text-white border-2 border-transparent shadow-[6px_6px_0px_#00ff00] p-6 flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4">
        <div>
          <div className="flex items-center gap-2 mb-1">
            <span className="bg-[#ee4d2d] text-white text-[10px] font-black uppercase px-2 py-0.5 border border-black font-mono">
              Shopee
            </span>
            <span className="bg-[#00ff00] text-black text-[10px] font-black uppercase tracking-wider px-2 py-0.5 border border-black">
              ĐỐI CHIẾU 2 CHIỀU TỰ ĐỘNG
            </span>
          </div>
          <h2 className="text-xl sm:text-2xl font-black uppercase tracking-tight text-white">
            Đối Chiếu Doanh Thu Sàn TMĐT
          </h2>
          <p className="text-xs text-slate-300 font-medium mt-1">
            Gom nhóm đa dòng SKU, khấu trừ voucher shop 1 lần duy nhất, đối chiếu chính xác theo Shop + Mã đơn hàng.
          </p>
        </div>

        <div className="flex items-center gap-2 self-end sm:self-center">
          <button
            onClick={handleLoadDemoData}
            className="bg-white hover:bg-yellow-200 text-black text-xs font-black uppercase px-3 py-2 border-2 border-black shadow-[3px_3px_0px_#00ff00] transition cursor-pointer flex items-center gap-1.5"
            title="Nạp nhanh 2 file Shopee & Phần mềm bán hàng mẫu để thử nghiệm tức thì"
          >
            <Sparkles size={14} className="text-amber-600" />
            <span>Nạp dữ liệu mẫu Demo</span>
          </button>

          {(marketplaceResult || salesResult || reconciliationRecords) && (
            <button
              onClick={handleReset}
              className="bg-slate-200 hover:bg-slate-300 text-black text-xs font-bold uppercase px-3 py-2 border-2 border-black transition cursor-pointer flex items-center gap-1"
              title="Xóa 2 file hiện tại và kết quả phiên này (vẫn giữ cấu hình đã nhớ)"
            >
              <RotateCcw size={13} />
              <span>Làm mới</span>
            </button>
          )}
        </div>
      </div>

      {/* BƯỚC 1: TẢI 2 FILE TRÊN CÙNG MỘT MÀN HÌNH (Section 32, 41) */}
      <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-6">
        <div className="flex items-center justify-between border-b-2 border-[#141414] pb-2">
          <h3 className="font-black text-xs uppercase tracking-wider text-black flex items-center gap-2">
            <span className="w-5 h-5 bg-black text-[#00ff00] flex items-center justify-center font-mono text-[11px]">1</span>
            <span>Bước 1: Chọn 2 file dữ liệu</span>
          </h3>
          <span className="text-[10px] text-slate-500 font-bold uppercase">
            Hệ thống tự động đọc &amp; nhận diện cột sau khi chọn file
          </span>
        </div>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* File sàn Shopee */}
          <div className="space-y-2">
            <label className="text-xs uppercase font-black text-black tracking-wider flex items-center justify-between">
              <span className="flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 bg-[#ee4d2d] inline-block border border-black"></span>
                <span>File Dữ Liệu Sàn (Shopee)</span>
              </span>
              {marketplaceResult && (
                <span className="text-[10px] font-mono text-green-700 font-bold flex items-center gap-1">
                  <Check size={12} /> Đã nhận diện
                </span>
              )}
            </label>

            <div
              onClick={() => mktFileInputRef.current?.click()}
              className={`border-2 border-dashed p-4 text-center cursor-pointer transition flex flex-col items-center justify-center min-h-[110px] ${
                marketplaceResult
                  ? "border-green-600 bg-emerald-50/60"
                  : "border-[#141414] hover:bg-slate-50 bg-[#fbfbf9]"
              }`}
            >
              <input
                ref={mktFileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleMarketplaceFileChange}
                className="hidden"
              />
              <Upload size={22} className={marketplaceResult ? "text-green-700 mb-1" : "text-slate-500 mb-1"} />

              {marketplaceResult ? (
                <div className="space-y-1">
                  <div className="text-xs font-black text-black font-mono truncate max-w-xs">
                    {marketplaceResult.fileName}
                  </div>
                  <div className="text-[11px] text-green-800 font-bold font-mono">
                    ✓ {marketplaceResult.totalRows.toLocaleString()} dòng – {marketplaceResult.sheets.length} shop
                  </div>
                </div>
              ) : (
                <div className="space-y-0.5">
                  <div className="text-xs font-black text-black uppercase">
                    Chọn file Shopee (.xlsx)
                  </div>
                  <div className="text-[10px] text-slate-500">
                    Nhấp hoặc kéo thả file đơn hàng Shopee
                  </div>
                </div>
              )}
            </div>

            {/* BLOCK TÓM TẮT TỐI GIẢN: KHỚP CỘT EXCEL – DỮ LIỆU SÀN (Section 1, 3, 4) */}
            {marketplaceResult && shopeeChecks && (
              <div className="border-2 border-black bg-white p-3 space-y-2.5 shadow-[2px_2px_0px_#141414]">
                <div className="flex items-center justify-between border-b-2 border-black pb-1.5">
                  <span className="text-[11px] font-black uppercase text-black tracking-wider flex items-center gap-1.5">
                    <span className="w-2 h-2 bg-[#ee4d2d] inline-block"></span>
                    KHỚP CỘT EXCEL – DỮ LIỆU SÀN
                  </span>
                  <span className={`text-[10px] font-black font-mono px-2 py-0.5 border ${
                    shopeeChecks.mappedCount === 6
                      ? "bg-green-100 text-green-800 border-green-300"
                      : "bg-amber-100 text-amber-800 border-amber-300"
                  }`}>
                    Đã nhận diện: {shopeeChecks.mappedCount}/6 cột
                  </span>
                </div>

                <div className="space-y-1 text-xs">
                  {shopeeChecks.items.map((item) => (
                    <div
                      key={item.key}
                      className="flex items-center justify-between py-0.5 px-1 rounded hover:bg-slate-50 text-[11px]"
                    >
                      <div className="flex items-center gap-1.5 min-w-0">
                        {item.isMapped ? (
                          <span className="text-green-700 font-black">✓</span>
                        ) : (
                          <span className="text-red-600 font-black">✕</span>
                        )}
                        <span className={`font-bold truncate ${item.isMapped ? "text-slate-800" : "text-red-700 font-black"}`}>
                          {item.label}
                        </span>
                      </div>
                      {item.header ? (
                        <span className="text-[10px] text-slate-500 font-mono shrink-0 ml-2">
                          {getShopeeColLetterAndName(item.header)}
                        </span>
                      ) : (
                        <span className="text-[10px] text-red-600 font-bold shrink-0">Chưa chọn</span>
                      )}
                    </div>
                  ))}
                </div>

                {/* Xử lý riêng khi không xác định chắc chắn cột "Tổng tiền người mua thanh toán" (Section 4) */}
                {shopeeChecks.hasUncertainBuyerPaid && (
                  <div className="p-2.5 bg-amber-50 border-2 border-amber-400 space-y-1.5 rounded">
                    <div className="flex items-center gap-1.5 font-black text-amber-900 text-[11px]">
                      <AlertTriangle size={14} className="text-amber-700 shrink-0" />
                      <span>Không xác định chắc chắn cột &quot;Tổng tiền người mua thanh toán&quot;</span>
                    </div>
                    <p className="text-[11px] text-amber-800 font-medium">Vui lòng chọn:</p>
                    <select
                      value={marketplaceResult.sheets[0]?.detectedMapping?.buyerPaidCol || ""}
                      onChange={(e) => handleUpdateShopeeMapping("buyerPaidCol", e.target.value)}
                      className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white cursor-pointer"
                    >
                      <option value="">-- Chọn cột tiền người mua thanh toán --</option>
                      {marketplaceResult.sheets[0]?.headers?.map((h, idx) => (
                        <option key={h} value={h}>
                          {getColumnLetter(idx)} — {h}
                        </option>
                      ))}
                    </select>
                  </div>
                )}

                <div className="pt-1.5 flex items-center justify-between border-t border-slate-200">
                  <span className="text-[10px] text-slate-500 font-mono">
                    {shopeeChecks.mappedCount === 6 ? "Đã nhận diện: 6/6 cột" : `Cần xác nhận ${6 - shopeeChecks.mappedCount} cột`}
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      setMappingModalTab("shopee");
                      setShowColumnMappingModal(true);
                    }}
                    className="px-2.5 py-1 text-[11px] font-black uppercase border border-black bg-slate-100 hover:bg-slate-200 text-black shadow-[1px_1px_0px_#141414] transition cursor-pointer flex items-center gap-1"
                  >
                    <SlidersHorizontal size={11} />
                    <span>[ Chỉnh cột ]</span>
                  </button>
                </div>
              </div>
            )}
          </div>

          {/* File phần mềm bán hàng */}
          <div className="space-y-2">
            <label className="text-xs uppercase font-black text-black tracking-wider flex items-center justify-between">
              <span className="flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 bg-blue-600 inline-block border border-black"></span>
                <span>File Phần Mềm Bán Hàng</span>
              </span>
              {salesResult && (
                <span className="text-[10px] font-mono text-green-700 font-bold flex items-center gap-1">
                  <Check size={12} /> Đã nhận diện
                </span>
              )}
            </label>

            <div
              onClick={() => salesFileInputRef.current?.click()}
              className={`border-2 border-dashed p-4 text-center cursor-pointer transition flex flex-col items-center justify-center min-h-[110px] ${
                salesResult
                  ? "border-green-600 bg-emerald-50/60"
                  : "border-[#141414] hover:bg-slate-50 bg-[#fbfbf9]"
              }`}
            >
              <input
                ref={salesFileInputRef}
                type="file"
                accept=".xlsx,.xls"
                onChange={handleSalesFileChange}
                className="hidden"
              />
              <FileSpreadsheet size={22} className={salesResult ? "text-green-700 mb-1" : "text-slate-500 mb-1"} />

              {salesResult ? (
                <div className="space-y-1">
                  <div className="text-xs font-black text-black font-mono truncate max-w-xs">
                    {salesResult.fileName}
                  </div>
                  <div className="text-[11px] text-green-800 font-bold font-mono">
                    ✓ {(salesResult.totalRows || salesResult.rows?.length || 0).toLocaleString()} dòng – {salesResult.detectedShops.length} shop
                  </div>
                </div>
              ) : (
                <div className="space-y-0.5">
                  <div className="text-xs font-black text-black uppercase">
                    Chọn file bán hàng (.xlsx)
                  </div>
                  <div className="text-[10px] text-slate-500">
                    Nhấp hoặc kéo thả bảng kê bán hàng
                  </div>
                </div>
              )}
            </div>

            {/* BLOCK TÓM TẮT TỐI GIẢN: KHỚP CỘT EXCEL – PHẦN MỀM BÁN HÀNG (Section 1, 5) */}
            {salesResult && salesChecks && (
              <div className="border-2 border-black bg-white p-3 space-y-2.5 shadow-[2px_2px_0px_#141414]">
                <div className="flex items-center justify-between border-b-2 border-black pb-1.5">
                  <span className="text-[11px] font-black uppercase text-black tracking-wider flex items-center gap-1.5">
                    <span className="w-2 h-2 bg-blue-600 inline-block"></span>
                    KHỚP CỘT EXCEL – PHẦN MỀM BÁN HÀNG
                  </span>
                  <span className={`text-[10px] font-black font-mono px-2 py-0.5 border ${
                    salesChecks.mappedCount === 6
                      ? "bg-green-100 text-green-800 border-green-300"
                      : "bg-amber-100 text-amber-800 border-amber-300"
                  }`}>
                    Đã nhận diện: {salesChecks.mappedCount}/6 cột
                  </span>
                </div>

                <div className="space-y-1 text-xs">
                  {salesChecks.items.map((item) => (
                    <div
                      key={item.key}
                      className="flex items-center justify-between py-0.5 px-1 rounded hover:bg-slate-50 text-[11px]"
                    >
                      <div className="flex items-center gap-1.5 min-w-0">
                        {item.isMapped ? (
                          <span className="text-green-700 font-black">✓</span>
                        ) : (
                          <span className="text-red-600 font-black">✕</span>
                        )}
                        <span className={`font-bold truncate ${item.isMapped ? "text-slate-800" : "text-red-700 font-black"}`}>
                          {item.label}
                        </span>
                      </div>
                      {item.header ? (
                        <span className="text-[10px] text-slate-500 font-mono shrink-0 ml-2">
                          {getSalesColLetterAndName(item.header)}
                        </span>
                      ) : (
                        <span className="text-[10px] text-red-600 font-bold shrink-0">Chưa chọn</span>
                      )}
                    </div>
                  ))}
                </div>

                {/* Nếu thiếu cột bắt buộc */}
                {(!salesResult.detectedMapping?.shopNameCol || !salesResult.detectedMapping?.orderIdCol || !salesResult.detectedMapping?.totalAmountCol) && (
                  <div className="p-2 bg-amber-50 border-2 border-amber-400 text-xs space-y-1 rounded">
                    <div className="flex items-center gap-1 font-black text-amber-900 text-[11px]">
                      <AlertTriangle size={13} className="text-amber-700" />
                      <span>Chưa nhận diện đủ các cột bắt buộc trên phần mềm</span>
                    </div>
                    <p className="text-[10px] text-amber-800">
                      Bấm <strong>&quot;[ Chỉnh cột ]&quot;</strong> để chọn cột tương ứng.
                    </p>
                  </div>
                )}

                <div className="pt-1.5 flex items-center justify-between border-t border-slate-200">
                  <span className="text-[10px] text-slate-500 font-mono">
                    {salesChecks.mappedCount === 6 ? "Đã nhận diện: 6/6 cột" : `Cần xác nhận ${6 - salesChecks.mappedCount} cột`}
                  </span>
                  <button
                    type="button"
                    onClick={() => {
                      setMappingModalTab("sales");
                      setShowColumnMappingModal(true);
                    }}
                    className="px-2.5 py-1 text-[11px] font-black uppercase border border-black bg-slate-100 hover:bg-slate-200 text-black shadow-[1px_1px_0px_#141414] transition cursor-pointer flex items-center gap-1"
                  >
                    <SlidersHorizontal size={11} />
                    <span>[ Chỉnh cột ]</span>
                  </button>
                </div>
              </div>
            )}
          </div>
        </div>

        {/* BƯỚC 2: TRẠNG THÁI NHẬN DIỆN & NÚT ĐỐI CHIẾU DUY NHẤT (Section 32, 38, 39) */}
        <div className="pt-2 border-t-2 border-[#141414] space-y-3">
          {marketplaceResult && salesResult ? (
            validationStatus.isReady ? (
              <div className="bg-[#f0f0ed] border-2 border-black p-4 flex flex-col sm:flex-row items-center justify-between gap-4">
                <div className="space-y-1 text-center sm:text-left">
                  <div className="flex items-center gap-1.5 text-xs font-black text-green-800 uppercase">
                    <CheckCircle2 size={16} className="text-green-700" />
                    <span>Sẵn sàng đối chiếu doanh thu</span>
                  </div>
                  <div className="text-[11px] text-slate-700 font-mono">
                    ✓ File sàn TMĐT: {marketplaceResult.totalRows.toLocaleString()} dòng ({marketplaceResult.sheets.length} shop) •
                    ✓ File bán hàng: {(salesResult.totalRows || salesResult.rows?.length || 0).toLocaleString()} dòng ({salesResult.detectedShops.length} shop) •
                    ✓ Đã nhận diện đủ cột bắt buộc
                  </div>
                  <div className="flex flex-wrap items-center gap-3 pt-1">
                    <button
                      type="button"
                      onClick={() => {
                        setMappingModalTab("shopee");
                        setShowColumnMappingModal(true);
                      }}
                      className="text-[11px] font-bold text-blue-700 hover:text-blue-900 underline flex items-center gap-1 cursor-pointer"
                    >
                      <SlidersHorizontal size={12} />
                      Kiểm tra / Tùy chỉnh khớp cột Excel
                    </button>
                    <span className="text-slate-300">|</span>
                    <button
                      type="button"
                      onClick={() => setShowShopMappingModal(true)}
                      className="text-[11px] font-bold text-slate-600 hover:text-black underline flex items-center gap-1 cursor-pointer"
                    >
                      <Store size={12} />
                      Kiểm tra mapping shop ({shopMappings.length} shop)
                    </button>
                  </div>
                </div>

                <button
                  onClick={handleExecuteReconciliation}
                  disabled={isProcessing}
                  className={`px-8 py-3.5 font-black uppercase text-xs border-2 border-black shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer flex items-center gap-2 whitespace-nowrap ${
                    isProcessing
                      ? "bg-yellow-300 text-black cursor-wait"
                      : "bg-[#00ff00] hover:bg-[#05e005] text-black"
                  }`}
                >
                  {isProcessing ? (
                    <>
                      <RefreshCw size={16} className="animate-spin text-black" />
                      <span>{processingProgress || "Đang đối chiếu..."}</span>
                    </>
                  ) : (
                    <>
                      <Play size={16} className="fill-black" />
                      <span>THỰC HIỆN ĐỐI CHIẾU</span>
                    </>
                  )}
                </button>
              </div>
            ) : (
              /* Thông báo lỗi chặn thân thiện (Section 38, 39) */
              <div className="bg-amber-50 border-2 border-amber-500 p-4 flex flex-col sm:flex-row items-center justify-between gap-4">
                <div className="flex items-start gap-2.5">
                  <AlertTriangle size={20} className="text-amber-700 shrink-0 mt-0.5" />
                  <div>
                    <h4 className="text-xs font-black uppercase text-amber-900">
                      Cần chọn lại thông tin cột hoặc shop
                    </h4>
                    <p className="text-[11px] text-amber-800 font-medium mt-0.5">
                      {validationStatus.blockingError}
                    </p>
                  </div>
                </div>

                <button
                  onClick={() => {
                    if (validationStatus.actionType === "shop") {
                      setShowShopMappingModal(true);
                    } else {
                      setShowColumnMappingModal(true);
                    }
                  }}
                  className="bg-amber-400 hover:bg-amber-500 text-black font-black uppercase text-xs px-4 py-2 border-2 border-black shadow-[2px_2px_0px_#141414] transition whitespace-nowrap cursor-pointer"
                >
                  {validationStatus.actionType === "shop" ? "KIỂM TRA SHOP" : "CHỌN CỘT"}
                </button>
              </div>
            )
          ) : (
            <div className="text-center py-2 text-xs text-slate-500 font-medium">
              Vui lòng tải lên cả 2 file để hệ thống tự động kiểm tra và sẵn sàng đối chiếu.
            </div>
          )}

          {/* CÀI ĐẶT NÂNG CAO (Mặc định thu gọn theo Section 34) */}
          <div className="pt-1">
            <button
              onClick={() => setShowAdvancedSettings(!showAdvancedSettings)}
              className="text-[11px] font-bold text-slate-600 hover:text-black flex items-center gap-1.5 transition cursor-pointer"
            >
              <SlidersHorizontal size={13} />
              <span>Cài đặt nâng cao &amp; Tùy chỉnh mapping</span>
              {showAdvancedSettings ? <ChevronUp size={14} /> : <ChevronDown size={14} />}
            </button>

            {showAdvancedSettings && (
              <div className="mt-3 p-4 bg-[#f8f8f6] border-2 border-slate-300 space-y-4 text-xs animate-fade-in">
                <div className="grid grid-cols-1 sm:grid-cols-3 gap-4">
                  <div>
                    <label className="text-[10px] font-black uppercase text-slate-700 block mb-1">
                      Mức dung sai tiền lệch
                    </label>
                    <div className="flex items-center gap-1.5">
                      <input
                        type="number"
                        min="0"
                        max="50000"
                        value={config.toleranceVnd}
                        onChange={(e) => setConfig({ ...config, toleranceVnd: Math.max(0, parseInt(e.target.value) || 0) })}
                        className="w-24 border border-black bg-white p-1.5 font-mono text-xs font-bold text-black"
                      />
                      <span className="text-[10px] text-slate-500 font-bold">VND (Mặc định 1 đ)</span>
                    </div>
                  </div>

                  <div>
                    <label className="text-[10px] font-black uppercase text-slate-700 block mb-1">
                      Kỳ đối chiếu
                    </label>
                    <input
                      type="text"
                      value={config.periodLabel}
                      onChange={(e) => setConfig({ ...config, periodLabel: e.target.value })}
                      placeholder="tháng 7"
                      className="w-36 border border-black bg-white p-1.5 text-xs font-bold text-black"
                    />
                  </div>

                  <div className="flex items-end gap-2">
                    <button
                      onClick={() => setShowColumnMappingModal(true)}
                      className="bg-white hover:bg-slate-100 text-black font-bold uppercase text-[10px] px-3 py-2 border border-black transition cursor-pointer"
                    >
                      Mapping cột thủ công
                    </button>
                    <button
                      onClick={() => setShowShopMappingModal(true)}
                      className="bg-white hover:bg-slate-100 text-black font-bold uppercase text-[10px] px-3 py-2 border border-black transition cursor-pointer"
                    >
                      Mapping shop thủ công
                    </button>
                  </div>
                </div>
              </div>
            )}
          </div>
        </div>
      </div>

      {/* BƯỚC 3: XEM KẾT QUẢ VÀ XUẤT FILE (Section 32, 41) */}
      {reconciliationRecords && reconciliationSummary && (
        <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-6 animate-fade-in">
          {/* Header kết quả & Nút Xuất Excel */}
          <div className="flex flex-col sm:flex-row items-start sm:items-center justify-between gap-4 border-b-2 border-[#141414] pb-4">
            <div>
              <span className="bg-[#00ff00] text-black text-[10px] font-black uppercase px-2 py-0.5 border border-black inline-block">
                KẾT QUẢ ĐỐI CHIẾU
              </span>
              <h3 className="text-base sm:text-lg font-black uppercase tracking-tight text-black mt-1">
                Bảng Tổng Hợp Doanh Thu Chi Tiết
              </h3>
            </div>

            <div className="flex items-center gap-2">
              <button
                onClick={handleExecuteReconciliation}
                disabled={isProcessing}
                className="bg-white hover:bg-yellow-100 text-black text-xs font-bold uppercase px-3 py-2.5 border-2 border-black transition cursor-pointer flex items-center gap-1.5"
                title="Chạy đối chiếu lại với file hiện tại"
              >
                <RefreshCw size={13} className={isProcessing ? "animate-spin" : ""} />
                <span>Đối chiếu lại</span>
              </button>

              <button
                onClick={handleExportExcel}
                className="bg-[#00ff00] hover:bg-[#05e005] text-black font-black uppercase text-xs px-5 py-2.5 border-2 border-black shadow-[3px_3px_0px_#141414] hover:shadow-[5px_5px_0px_#141414] hover:translate-y-[-1px] transition cursor-pointer flex items-center gap-2"
              >
                <Download size={15} />
                <span>XUẤT KẾT QUẢ EXCEL</span>
              </button>
            </div>
          </div>

          {/* DÒNG CHỈ SỐ KPI TỔNG QUAN (Section 32) */}
          <div className="grid grid-cols-2 sm:grid-cols-4 lg:grid-cols-8 gap-2.5 text-center font-mono">
            <div className="p-2.5 border-2 border-black bg-slate-50">
              <div className="text-[10px] font-black uppercase text-slate-500 font-sans">Đơn Sàn</div>
              <div className="text-base font-black text-black mt-0.5">{reconciliationSummary.totalMarketplaceOrders.toLocaleString()}</div>
            </div>
            <div className="p-2.5 border-2 border-black bg-slate-50">
              <div className="text-[10px] font-black uppercase text-slate-500 font-sans">Đơn PM</div>
              <div className="text-base font-black text-black mt-0.5">{reconciliationSummary.totalSalesOrders.toLocaleString()}</div>
            </div>
            <div
              onClick={() => setStatusFilter("KHỚP")}
              className={`p-2.5 border-2 border-black cursor-pointer transition ${
                statusFilter === "KHỚP" ? "bg-[#00ff00] text-black" : "bg-emerald-50 hover:bg-emerald-100"
              }`}
            >
              <div className="text-[10px] font-black uppercase text-emerald-800 font-sans">Khớp</div>
              <div className="text-base font-black text-emerald-900 mt-0.5">{reconciliationSummary.matchedCount.toLocaleString()}</div>
            </div>
            <div
              onClick={() => setStatusFilter("CHÊNH LỆCH")}
              className={`p-2.5 border-2 border-black cursor-pointer transition ${
                statusFilter === "CHÊNH LỆCH" ? "bg-red-500 text-white" : "bg-red-50 hover:bg-red-100"
              }`}
            >
              <div className={`text-[10px] font-black uppercase font-sans ${statusFilter === "CHÊNH LỆCH" ? "text-white" : "text-red-700"}`}>
                Chênh lệch
              </div>
              <div className={`text-base font-black mt-0.5 ${statusFilter === "CHÊNH LỆCH" ? "text-white" : "text-red-600"}`}>
                {reconciliationSummary.differenceCount.toLocaleString()}
              </div>
            </div>
            <div
              onClick={() => setStatusFilter("SÀN_THIẾU_PM")}
              className={`p-2.5 border-2 border-black cursor-pointer transition ${
                statusFilter === "SÀN_THIẾU_PM" ? "bg-amber-300 text-black" : "bg-amber-50 hover:bg-amber-100"
              }`}
            >
              <div className="text-[10px] font-black uppercase text-amber-800 font-sans">Sàn chưa có PM</div>
              <div className="text-base font-black text-amber-900 mt-0.5">{reconciliationSummary.marketplaceOnlyCount.toLocaleString()}</div>
            </div>
            <div
              onClick={() => setStatusFilter("PM_THIẾU_SÀN")}
              className={`p-2.5 border-2 border-black cursor-pointer transition ${
                statusFilter === "PM_THIẾU_SÀN" ? "bg-purple-600 text-white" : "bg-purple-50 hover:bg-purple-100"
              }`}
            >
              <div className={`text-[10px] font-black uppercase font-sans ${statusFilter === "PM_THIẾU_SÀN" ? "text-white" : "text-purple-700"}`}>
                PM chưa có Sàn
              </div>
              <div className={`text-base font-black mt-0.5 ${statusFilter === "PM_THIẾU_SÀN" ? "text-white" : "text-purple-700"}`}>
                {reconciliationSummary.salesOnlyCount.toLocaleString()}
              </div>
            </div>
            <div
              onClick={() => setStatusFilter("ĐƠN_HỦY")}
              className={`p-2.5 border-2 border-black cursor-pointer transition ${
                statusFilter === "ĐƠN_HỦY" ? "bg-slate-700 text-white" : "bg-slate-100 hover:bg-slate-200"
              }`}
            >
              <div className={`text-[10px] font-black uppercase font-sans ${statusFilter === "ĐƠN_HỦY" ? "text-white" : "text-slate-600"}`}>
                Đơn hủy
              </div>
              <div className={`text-base font-black mt-0.5 ${statusFilter === "ĐƠN_HỦY" ? "text-white" : "text-slate-700"}`}>
                {reconciliationSummary.marketplaceOnlyCancelledCount.toLocaleString()}
              </div>
            </div>
            <div
              onClick={() => setStatusFilter("CẦN_KT")}
              className={`p-2.5 border-2 border-black cursor-pointer transition ${
                statusFilter === "CẦN_KT" ? "bg-yellow-400 text-black" : "bg-yellow-50 hover:bg-yellow-100"
              }`}
            >
              <div className="text-[10px] font-black uppercase text-amber-900 font-sans">Cần kiểm tra</div>
              <div className="text-base font-black text-amber-950 mt-0.5">{reconciliationSummary.anomaliesCount.toLocaleString()}</div>
            </div>
          </div>

          {/* DÒNG TỔNG DOANH THU & CHÊNH LỆCH */}
          <div className="bg-[#141414] text-white p-3.5 border-2 border-black flex flex-wrap items-center justify-between gap-4 font-mono text-xs">
            <div className="flex items-center gap-6 flex-wrap">
              <div>
                <span className="text-slate-400 text-[10px] uppercase font-bold block">Tổng DT Sàn:</span>
                <span className="text-sm font-black text-[#00ff00]">
                  {reconciliationSummary.totalMarketplaceRevenue.toLocaleString("vi-VN")} đ
                </span>
              </div>
              <div>
                <span className="text-slate-400 text-[10px] uppercase font-bold block">Tổng DT Phần Mềm:</span>
                <span className="text-sm font-black text-cyan-300">
                  {reconciliationSummary.totalSalesRevenue.toLocaleString("vi-VN")} đ
                </span>
              </div>
              <div>
                <span className="text-slate-400 text-[10px] uppercase font-bold block">Chênh lệch:</span>
                <span className={`text-sm font-black ${
                  reconciliationSummary.totalDifferenceAmount === 0 ? "text-white" : "text-red-400"
                }`}>
                  {reconciliationSummary.totalDifferenceAmount > 0 ? "+" : ""}
                  {reconciliationSummary.totalDifferenceAmount.toLocaleString("vi-VN")} đ
                </span>
              </div>
            </div>

            <div className="text-[11px] text-slate-300 font-sans">
              Hiển thị <strong className="font-mono text-white">{filteredRecords.length}</strong> / {reconciliationRecords.length} đơn
            </div>
          </div>

          {/* THANH BỘ LỌC ĐƠN GIẢN TRÊN CÙNG MỘT MÀN HÌNH (Section 32) */}
          <div className="flex flex-wrap items-center justify-between gap-3 bg-[#f0f0ed] p-2.5 border-2 border-[#141414]">
            {/* Filter pills */}
            <div className="flex items-center gap-1.5 flex-wrap">
              {[
                { key: "ALL", label: "Tất cả" },
                { key: "KHỚP", label: "Khớp" },
                { key: "CHÊNH LỆCH", label: "Chênh lệch" },
                { key: "SÀN_THIẾU_PM", label: "Sàn chưa có PM" },
                { key: "PM_THIẾU_SÀN", label: "PM chưa có Sàn" },
                { key: "ĐƠN_HỦY", label: "Đơn hủy" },
                { key: "CẦN_KT", label: "Cần kiểm tra" }
              ].map(tab => (
                <button
                  key={tab.key}
                  onClick={() => setStatusFilter(tab.key)}
                  className={`px-2.5 py-1 text-xs font-black uppercase border border-black transition ${
                    statusFilter === tab.key
                      ? "bg-[#141414] text-white shadow-[1px_1px_0px_#00ff00]"
                      : "bg-white text-black hover:bg-slate-200"
                  }`}
                >
                  {tab.label}
                </button>
              ))}
            </div>

            {/* Shop select & Search box */}
            <div className="flex items-center gap-2 flex-wrap">
              {availableShops.length > 1 && (
                <div className="flex items-center gap-1 text-xs">
                  <Store size={13} className="text-slate-600" />
                  <select
                    value={selectedShopFilter}
                    onChange={(e) => setSelectedShopFilter(e.target.value)}
                    className="border border-black bg-white p-1 text-xs font-bold text-black focus:outline-none"
                  >
                    <option value="ALL">Tất cả cửa hàng</option>
                    {availableShops.map(s => (
                      <option key={s} value={s}>{s}</option>
                    ))}
                  </select>
                </div>
              )}

              <div className="relative">
                <Search size={12} className="absolute left-2.5 top-2.5 text-slate-400" />
                <input
                  type="text"
                  value={searchQuery}
                  onChange={(e) => setSearchQuery(e.target.value)}
                  placeholder="Tìm mã đơn, số HĐ..."
                  className="border border-black pl-7 pr-2.5 py-1 text-xs font-mono font-bold text-black focus:outline-none focus:bg-yellow-50 w-44"
                />
              </div>
            </div>
          </div>

          {/* BẢNG KẾT QUẢ DUY NHẤT (Section 32) */}
          <div className="border-2 border-[#141414] overflow-x-auto">
            <table className="w-full text-left text-xs border-collapse">
              <thead>
                <tr className="bg-[#141414] text-white font-black uppercase text-[10px] tracking-wider">
                  <th colSpan={5} className="p-2 pl-3 border-r border-slate-700 text-orange-400">
                    DATA BÁN HÀNG (SÀN TMĐT)
                  </th>
                  <th colSpan={4} className="p-2 pl-3 border-r border-slate-700 text-cyan-300 bg-[#1e293b]">
                    PHẦN MỀM BÁN HÀNG
                  </th>
                  <th colSpan={4} className="p-2 pl-3 text-[#00ff00] bg-[#0f172a]">
                    ĐỐI CHIẾU
                  </th>
                </tr>
                <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black text-[11px] font-mono">
                  <th className="p-2.5 pl-3">Shop</th>
                  <th className="p-2.5">Mã đơn hàng</th>
                  <th className="p-2.5">Trạng thái sàn</th>
                  <th className="p-2.5">Thời gian giao</th>
                  <th className="p-2.5 border-r-2 border-[#141414] text-right">Tổng DT xuất HĐ</th>
                  <th className="p-2.5 pl-3">Mã đơn PM</th>
                  <th className="p-2.5">Số Hóa Đơn</th>
                  <th className="p-2.5">Ngày Hóa Đơn</th>
                  <th className="p-2.5 border-r-2 border-[#141414] text-right">Tổng tiền PM</th>
                  <th className="p-2.5 pl-3 text-right">Chênh lệch DT</th>
                  <th className="p-2.5 text-center">Trạng thái</th>
                  <th className="p-2.5">Note</th>
                  <th className="p-2.5 text-center pr-3">Xem</th>
                </tr>
              </thead>
              <tbody className="divide-y divide-slate-200 font-mono text-[11px]">
                {filteredRecords.length > 0 ? (
                  filteredRecords.map((row) => {
                    const isMatch = row.status === "KHỚP";
                    const isDiff = row.status === "CHÊNH LỆCH";
                    const isCancelled = row.status === "ĐƠN HỦY - KHÔNG CÓ TRÊN PM";
                    const isMktOnly = row.status === "SÀN CÓ - PM KHÔNG CÓ";
                    const isSalesOnly = row.status === "PM CÓ - SÀN KHÔNG CÓ";

                    return (
                      <tr
                        key={row.id}
                        className={`hover:bg-[#f0f0ed]/60 transition ${
                          isDiff ? "bg-red-50/50" : isSalesOnly ? "bg-purple-50/30" : ""
                        }`}
                      >
                        <td className="p-2.5 pl-3 font-bold text-black max-w-[120px] truncate" title={row.shopDisplayName}>
                          {row.shopDisplayName}
                        </td>
                        <td className="p-2.5 font-black text-black">
                          {row.hasMarketplace ? row.originalOrderId : <span className="text-slate-400">--</span>}
                        </td>
                        <td className="p-2.5 text-slate-700">
                          {row.hasMarketplace ? (
                            <span className={`px-1.5 py-0.5 border text-[10px] font-bold ${
                              isCancelled ? "bg-slate-100 text-slate-600 border-slate-300" : "bg-emerald-50 text-emerald-800 border-emerald-300"
                            }`}>
                              {row.marketplaceStatus || "Giao thành công"}
                            </span>
                          ) : "--"}
                        </td>
                        <td className="p-2.5 text-slate-600">{row.marketplaceDeliveryDate || "--"}</td>
                        <td className="p-2.5 border-r-2 border-[#141414] text-right font-black text-black">
                          {row.hasMarketplace ? (
                            <div>
                              <span>{row.marketplaceRevenue.toLocaleString("vi-VN")}đ</span>
                              {row.marketplaceVoucher > 0 && (
                                <span className="block text-[9px] text-orange-600 font-normal">
                                  (-{row.marketplaceVoucher.toLocaleString("vi-VN")}đ vc)
                                </span>
                              )}
                            </div>
                          ) : "--"}
                        </td>
                        <td className="p-2.5 pl-3 font-black text-black">
                          {row.hasSales ? row.originalOrderId : <span className="text-slate-400 italic">Chưa có trên PM</span>}
                        </td>
                        <td className="p-2.5 text-slate-800 font-bold">{row.salesInvoiceNumber || "--"}</td>
                        <td className="p-2.5 text-slate-600">{row.salesInvoiceDate || "--"}</td>
                        <td className="p-2.5 border-r-2 border-[#141414] text-right font-black text-blue-900">
                          {row.hasSales ? `${row.salesTotalAmount.toLocaleString("vi-VN")}đ` : "--"}
                        </td>
                        <td className="p-2.5 pl-3 text-right font-black">
                          {row.revenueDifference === 0 ? (
                            <span className="text-green-700">0đ</span>
                          ) : (
                            <span className={row.revenueDifference > 0 ? "text-red-600" : "text-purple-700"}>
                              {row.revenueDifference > 0 ? "+" : ""}{row.revenueDifference.toLocaleString("vi-VN")}đ
                            </span>
                          )}
                        </td>
                        <td className="p-2.5 text-center">
                          <span className={`inline-block px-2 py-0.5 border border-black font-black text-[10px] ${
                            isMatch
                              ? "bg-[#00ff00] text-black shadow-[1px_1px_0px_#141414]"
                              : isDiff
                              ? "bg-red-500 text-white"
                              : isCancelled
                              ? "bg-slate-200 text-slate-700"
                              : isMktOnly
                              ? "bg-amber-300 text-black"
                              : isSalesOnly
                              ? "bg-purple-600 text-white"
                              : "bg-yellow-300 text-black"
                          }`}>
                            {row.status}
                          </span>
                        </td>
                        <td className="p-2.5 text-slate-700 max-w-[170px] truncate" title={row.note}>
                          {row.note}
                        </td>
                        <td className="p-2.5 text-center pr-3">
                          <button
                            onClick={() => setSelectedRecordDetail(row)}
                            className="bg-white hover:bg-yellow-200 text-black px-1.5 py-1 border border-black shadow-[1px_1px_0px_#141414] cursor-pointer inline-flex items-center justify-center"
                            title="Xem chi tiết SKU và hóa đơn"
                          >
                            <Eye size={12} />
                          </button>
                        </td>
                      </tr>
                    );
                  })
                ) : (
                  <tr>
                    <td colSpan={13} className="p-8 text-center text-slate-500 font-bold">
                      Không có đơn hàng nào khớp với bộ lọc hiện tại.
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        </div>
      )}

      {/* MODAL MAPPING CỘT DỮ LIỆU EXCEL THEO ĐÚNG LOGIC ĐÃ KIỂM CHỨNG (Section A, B, C) */}
      {showColumnMappingModal && marketplaceResult && salesResult && (
        <div className="fixed inset-0 bg-[#141414]/85 z-50 flex items-center justify-center p-3 sm:p-4 backdrop-blur-xs">
          <div className="bg-white border-4 border-[#141414] max-w-4xl w-full p-5 sm:p-6 shadow-[10px_10px_0px_#141414] space-y-4 max-h-[92vh] flex flex-col">
            {/* Header Modal */}
            <div className="flex justify-between items-start border-b-2 border-black pb-3 shrink-0">
              <div className="space-y-0.5">
                <div className="flex items-center gap-2">
                  <span className="bg-[#141414] text-white px-2 py-0.5 text-[10px] font-black uppercase">
                    CẤU HÌNH KHỚP CỘT
                  </span>
                  <h4 className="font-black text-sm uppercase tracking-wider text-black">
                    Khớp Cột Dữ Liệu Excel (Shopee &amp; Phần Mềm Bán Hàng)
                  </h4>
                </div>
                <p className="text-[11px] text-slate-600 font-medium">
                  Chỉ hiển thị các trường tham gia tính toán, đối chiếu và xuất kết quả theo quy chuẩn kế toán.
                </p>
              </div>
              <button
                onClick={() => setShowColumnMappingModal(false)}
                className="text-xs font-black px-2.5 py-1 bg-slate-200 hover:bg-slate-300 border border-black cursor-pointer transition"
              >
                ✕ Đóng
              </button>
            </div>

            {/* Switcher Tab 2 File */}
            <div className="flex border-b-2 border-black shrink-0 gap-1 bg-slate-100 p-1">
              <button
                type="button"
                onClick={() => setMappingModalTab("shopee")}
                className={`flex-1 py-2 px-3 text-xs font-black uppercase transition cursor-pointer flex items-center justify-center gap-2 border-2 ${
                  mappingModalTab === "shopee"
                    ? "bg-[#ff5722] text-white border-black shadow-[2px_2px_0px_#141414]"
                    : "bg-white text-slate-700 border-transparent hover:bg-slate-200"
                }`}
              >
                <span>A. File Dữ Liệu Sàn Shopee</span>
                <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono ${
                  mappingModalTab === "shopee" ? "bg-black/30 text-white" : "bg-slate-200 text-black"
                }`}>
                  {marketplaceResult.sheets.length} shop
                </span>
              </button>

              <button
                type="button"
                onClick={() => setMappingModalTab("sales")}
                className={`flex-1 py-2 px-3 text-xs font-black uppercase transition cursor-pointer flex items-center justify-center gap-2 border-2 ${
                  mappingModalTab === "sales"
                    ? "bg-blue-600 text-white border-black shadow-[2px_2px_0px_#141414]"
                    : "bg-white text-slate-700 border-transparent hover:bg-slate-200"
                }`}
              >
                <span>B. File Phần Mềm Bán Hàng</span>
                <span className={`text-[10px] px-1.5 py-0.2 rounded font-mono ${
                  mappingModalTab === "sales" ? "bg-black/30 text-white" : "bg-slate-200 text-black"
                }`}>
                  {salesResult.detectedShops.length} shop
                </span>
              </button>
            </div>

            {/* Nội dung mapping cuộn */}
            <div className="overflow-y-auto flex-1 pr-1 space-y-5">
              {/* TAB A: FILE SÀN SHOPEE */}
              {mappingModalTab === "shopee" && (
                <div className="space-y-5">
                  {/* Banner hướng dẫn Shop Identifier */}
                  <div className="bg-orange-50 border-2 border-orange-300 p-3 text-xs text-orange-950 flex items-start gap-2.5">
                    <CheckCircle2 size={18} className="text-orange-700 shrink-0 mt-0.5" />
                    <div>
                      <span className="font-black uppercase block text-[11px] text-orange-900">
                        Quy tắc Shop Identifier phía sàn Shopee:
                      </span>
                      <p className="text-[11px] text-orange-800 font-medium mt-0.5 leading-relaxed">
                        Mỗi sheet trong file Excel tương ứng với một cửa hàng. Tên sheet được hệ thống tự động sử dụng làm định danh Shop. Do đó, <strong>KHÔNG cần mapping cột Tên Shop phía Shopee</strong>.
                      </p>
                    </div>
                  </div>

                  {/* NHÓM 1: KHÓA ĐỐI CHIẾU */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-3">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-black inline-block"></span>
                        1. Khóa Đối Chiếu
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">So khớp 1-1 đơn hàng</span>
                    </div>

                    <div className="space-y-1.5">
                      <div className="flex flex-wrap items-center justify-between gap-1">
                        <label className="text-xs font-black text-black">
                          Mã đơn hàng <span className="text-red-600 font-black">*</span>
                        </label>
                        <div className="flex items-center gap-1.5">
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            Internal: marketplaceOrderId
                          </span>
                          {isExactShopeeHeader("orderIdCol", marketplaceResult.sheets[0]?.detectedMapping.orderIdCol || "") ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-black text-green-800 bg-green-100 border border-green-300 rounded">
                              <Check size={11} className="text-green-700 stroke-[3]" /> Đã tự động nhận diện chính xác
                            </span>
                          ) : marketplaceResult.sheets[0]?.detectedMapping.orderIdCol ? (
                            <span className="px-2 py-0.5 text-[10px] font-bold text-blue-800 bg-blue-100 border border-blue-300 rounded">
                              Đã chọn
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 text-[10px] font-black text-red-800 bg-red-100 border border-red-300 rounded">
                              Bắt buộc chọn (*)
                            </span>
                          )}
                        </div>
                      </div>
                      <p className="text-[10px] text-slate-500">Auto-map ưu tiên: &quot;Mã đơn hàng&quot;</p>
                      <select
                        value={marketplaceResult.sheets[0]?.detectedMapping.orderIdCol || ""}
                        onChange={(e) => handleUpdateShopeeMapping("orderIdCol", e.target.value)}
                        className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white focus:bg-amber-50"
                      >
                        <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                        {marketplaceResult.sheets[0]?.headers.map((h, idx) => (
                          <option key={h} value={h}>
                            {getColumnLetter(idx)} — {h}
                          </option>
                        ))}
                      </select>

                      {/* Preview Box */}
                      {marketplaceResult.sheets[0]?.detectedMapping.orderIdCol && (
                        <div className="bg-[#f8f8f6] border border-slate-300 p-2 text-[11px] space-y-1">
                          <div className="flex items-center justify-between text-[10px] text-slate-600 font-bold">
                            <span>👁 Xem trước mẫu dữ liệu (3-5 dòng đầu):</span>
                            <span>Cột {getColumnLetter(marketplaceResult.sheets[0].headers.indexOf(marketplaceResult.sheets[0].detectedMapping.orderIdCol))}</span>
                          </div>
                          <div className="flex flex-wrap gap-1.5">
                            {getShopeePreviewValues(marketplaceResult.sheets[0].detectedMapping.orderIdCol).map((v, i) => (
                              <span key={i} className="bg-white px-2 py-0.5 text-[11px] font-mono font-bold text-black border border-slate-300 rounded">
                                {v}
                              </span>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>

                  {/* NHÓM 2: THÔNG TIN TRẠNG THÁI */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-4">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-black inline-block"></span>
                        2. Thông Tin Trạng Thái
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">Trạng thái đơn, giao hàng &amp; hoàn tiền</span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Trạng thái đơn hàng */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Trạng thái đơn hàng <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            orderStatus
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Trạng Thái Đơn Hàng&quot;</span>
                          {isExactShopeeHeader("orderStatusCol", marketplaceResult.sheets[0]?.detectedMapping.orderStatusCol || "") ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={marketplaceResult.sheets[0]?.detectedMapping.orderStatusCol || ""}
                          onChange={(e) => handleUpdateShopeeMapping("orderStatusCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {marketplaceResult.sheets[0]?.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {marketplaceResult.sheets[0]?.detectedMapping.orderStatusCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getShopeePreviewValues(marketplaceResult.sheets[0].detectedMapping.orderStatusCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Thời gian giao hàng */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-bold text-black">
                            Thời gian giao hàng
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            deliveryTime
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Thời gian giao hàng&quot;</span>
                          {isExactShopeeHeader("deliveryTimeCol", marketplaceResult.sheets[0]?.detectedMapping.deliveryTimeCol || "") ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={marketplaceResult.sheets[0]?.detectedMapping.deliveryTimeCol || ""}
                          onChange={(e) => handleUpdateShopeeMapping("deliveryTimeCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Không chọn / Bỏ qua --</option>
                          {marketplaceResult.sheets[0]?.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {marketplaceResult.sheets[0]?.detectedMapping.deliveryTimeCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getShopeePreviewValues(marketplaceResult.sheets[0].detectedMapping.deliveryTimeCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Trạng thái Trả hàng/Hoàn tiền */}
                      <div className="space-y-1.5 md:col-span-2">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Trạng thái Trả hàng/Hoàn tiền <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            returnRefundStatus
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Trạng thái Trả hàng/Hoàn tiền&quot;</span>
                          {isExactShopeeHeader("returnRefundCol", marketplaceResult.sheets[0]?.detectedMapping.returnRefundCol || "") ? (
                            <span className="inline-flex items-center gap-1 px-1.5 py-0.5 text-[10px] font-bold text-green-800 bg-green-100 border border-green-300 rounded">
                              <Check size={10} /> Đã tự động nhận diện chính xác
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={marketplaceResult.sheets[0]?.detectedMapping.returnRefundCol || ""}
                          onChange={(e) => handleUpdateShopeeMapping("returnRefundCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {marketplaceResult.sheets[0]?.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {marketplaceResult.sheets[0]?.detectedMapping.returnRefundCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-2 text-[11px] space-y-1">
                            <div className="flex items-center justify-between text-[10px] text-slate-600 font-bold">
                              <span>👁 Mẫu dữ liệu hoàn tiền:</span>
                            </div>
                            <div className="flex flex-wrap gap-1.5">
                              {getShopeePreviewValues(marketplaceResult.sheets[0].detectedMapping.returnRefundCol).map((v, i) => (
                                <span key={i} className="bg-white px-2 py-0.5 text-[11px] font-mono font-bold text-black border border-slate-300 rounded">
                                  {v}
                                </span>
                              ))}
                            </div>
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* NHÓM 3: TÍNH DOANH THU (CHÚ Ý CẢNH BÁO HAI CỘT KHÁC NHAU) */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-4">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-green-600 inline-block"></span>
                        3. Tính Doanh Thu
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">Cộng dòng tiền &amp; trừ voucher</span>
                    </div>

                    {/* Cột Tổng số tiền Người mua thanh toán */}
                    <div className="space-y-2">
                      <div className="flex flex-wrap items-center justify-between gap-1">
                        <label className="text-xs font-black text-black">
                          Tổng số tiền Người mua thanh toán <span className="text-red-600 font-black">*</span>
                        </label>
                        <div className="flex items-center gap-1.5">
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            Internal: buyerPaidLineAmount
                          </span>
                          {isExactShopeeHeader("buyerPaidCol", marketplaceResult.sheets[0]?.detectedMapping.buyerPaidCol || "") ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-black text-green-800 bg-green-100 border border-green-300 rounded">
                              <Check size={11} className="text-green-700 stroke-[3]" /> Khớp chính xác case-sensitive
                            </span>
                          ) : (
                            <span className="px-2 py-0.5 text-[10px] font-black text-amber-800 bg-amber-100 border border-amber-300 rounded">
                              Cần kiểm tra kỹ
                            </span>
                          )}
                        </div>
                      </div>

                      {/* Khung cảnh báo phân biệt 2 header có thật của Shopee */}
                      <div className="bg-amber-50 border-2 border-amber-400 p-2.5 text-xs space-y-1">
                        <div className="flex items-center gap-1.5 font-black text-amber-900 uppercase text-[11px]">
                          <AlertTriangle size={15} className="text-amber-700 shrink-0" />
                          <span>Cảnh báo kế toán: Phân biệt 2 cột trong file xuất Shopee</span>
                        </div>
                        <p className="text-[11px] text-amber-900 leading-relaxed font-medium">
                          Trong file Shopee thực tế tồn tại đồng thời 2 header: <strong>&quot;Tổng số tiền Người mua thanh toán&quot;</strong> (chữ N hoa) và <strong>&quot;Tổng số tiền người mua thanh toán&quot;</strong> (chữ n thường). Đây là hai cột khác nhau. Reconciliation engine hiện tại bắt buộc phải sử dụng cột chữ N hoa để đối chiếu chính xác.
                        </p>
                      </div>

                      <select
                        value={marketplaceResult.sheets[0]?.detectedMapping.buyerPaidCol || ""}
                        onChange={(e) => handleUpdateShopeeMapping("buyerPaidCol", e.target.value)}
                        className="w-full border-2 border-black p-2.5 text-xs font-mono font-bold bg-white focus:bg-amber-50"
                      >
                        <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                        {marketplaceResult.sheets[0]?.headers.map((h, idx) => (
                          <option key={h} value={h}>
                            {getColumnLetter(idx)} — {h}
                          </option>
                        ))}
                      </select>

                      {marketplaceResult.sheets[0]?.detectedMapping.buyerPaidCol && (
                        <div className="bg-[#f8f8f6] border border-slate-300 p-2 text-[11px] space-y-1">
                          <div className="flex items-center justify-between text-[10px] text-slate-600 font-bold">
                            <span>👁 Mẫu số tiền người mua trả (3-5 dòng đầu):</span>
                            <span>Cột {getColumnLetter(marketplaceResult.sheets[0].headers.indexOf(marketplaceResult.sheets[0].detectedMapping.buyerPaidCol))}</span>
                          </div>
                          <div className="flex flex-wrap gap-1.5">
                            {getShopeePreviewValues(marketplaceResult.sheets[0].detectedMapping.buyerPaidCol).map((v, i) => (
                              <span key={i} className="bg-white px-2 py-0.5 text-[11px] font-mono font-bold text-black border border-slate-300 rounded">
                                {Number(v).toLocaleString("vi-VN")} đ
                              </span>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>

                    {/* Cột Mã giảm giá của Shop */}
                    <div className="space-y-2 pt-2 border-t border-slate-200">
                      <div className="flex flex-wrap items-center justify-between gap-1">
                        <label className="text-xs font-black text-black">
                          Mã giảm giá của Shop <span className="text-red-600 font-black">*</span>
                        </label>
                        <div className="flex items-center gap-1.5">
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            Internal: shopVoucher
                          </span>
                          {isExactShopeeHeader("shopVoucherCol", marketplaceResult.sheets[0]?.detectedMapping.shopVoucherCol || "") ? (
                            <span className="inline-flex items-center gap-1 px-2 py-0.5 text-[10px] font-bold text-green-800 bg-green-100 border border-green-300 rounded">
                              <Check size={11} className="text-green-700" /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                      </div>
                      <p className="text-[10px] text-slate-500">
                        Voucher Shop tài trợ: Chỉ trừ 1 lần duy nhất cho mỗi đơn hàng (kể cả đơn nhiều dòng SKU lặp lại voucher).
                      </p>
                      <select
                        value={marketplaceResult.sheets[0]?.detectedMapping.shopVoucherCol || ""}
                        onChange={(e) => handleUpdateShopeeMapping("shopVoucherCol", e.target.value)}
                        className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                      >
                        <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                        {marketplaceResult.sheets[0]?.headers.map((h, idx) => (
                          <option key={h} value={h}>
                            {getColumnLetter(idx)} — {h}
                          </option>
                        ))}
                      </select>

                      {marketplaceResult.sheets[0]?.detectedMapping.shopVoucherCol && (
                        <div className="bg-[#f8f8f6] border border-slate-300 p-2 text-[11px] space-y-1">
                          <div className="flex items-center justify-between text-[10px] text-slate-600 font-bold">
                            <span>👁 Mẫu voucher shop (3-5 dòng đầu):</span>
                            <span>Cột {getColumnLetter(marketplaceResult.sheets[0].headers.indexOf(marketplaceResult.sheets[0].detectedMapping.shopVoucherCol))}</span>
                          </div>
                          <div className="flex flex-wrap gap-1.5">
                            {getShopeePreviewValues(marketplaceResult.sheets[0].detectedMapping.shopVoucherCol).map((v, i) => (
                              <span key={i} className="bg-white px-2 py-0.5 text-[11px] font-mono font-bold text-black border border-slate-300 rounded">
                                {Number(v).toLocaleString("vi-VN")} đ
                              </span>
                            ))}
                          </div>
                        </div>
                      )}
                    </div>
                  </div>
                </div>
              )}

              {/* TAB B: FILE PHẦN MỀM BÁN HÀNG */}
              {mappingModalTab === "sales" && (
                <div className="space-y-5">
                  {/* NHÓM 1: KHÓA ĐỐI CHIẾU */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-3">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-blue-600 inline-block"></span>
                        1. Khóa Đối Chiếu (Shop + Mã Đơn)
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">Khóa composite ghép với Sàn</span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Tên Shop */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Tên Shop <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesShopName
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Tên shop&quot;</span>
                          {isExactSalesHeader("shopNameCol", salesResult.detectedMapping.shopNameCol) ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={salesResult.detectedMapping.shopNameCol}
                          onChange={(e) => handleUpdateSalesMapping("shopNameCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.shopNameCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.shopNameCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Mã đơn từ sàn */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Mã đơn từ sàn <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesMarketplaceOrderId
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Số đơn hàng từ sàn&quot; / HT khác</span>
                          {isExactSalesHeader("orderIdCol", salesResult.detectedMapping.orderIdCol) ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={salesResult.detectedMapping.orderIdCol}
                          onChange={(e) => handleUpdateSalesMapping("orderIdCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.orderIdCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.orderIdCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* NHÓM 2: THÔNG TIN HÓA ĐƠN */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-3">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-blue-600 inline-block"></span>
                        2. Thông Tin Hóa Đơn &amp; Chứng Từ
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">Đối chiếu số &amp; ngày hóa đơn</span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Số chứng từ / Số hóa đơn */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Số chứng từ / Số hóa đơn <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesInvoiceNumber
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Số hóa đơn&quot; hoặc &quot;Số chứng từ&quot;</span>
                          {isExactSalesHeader("invoiceNumberCol", salesResult.detectedMapping.invoiceNumberCol) ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={salesResult.detectedMapping.invoiceNumberCol}
                          onChange={(e) => handleUpdateSalesMapping("invoiceNumberCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.invoiceNumberCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.invoiceNumberCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Ngày hóa đơn */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Ngày hóa đơn <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesInvoiceDate
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Ngày hóa đơn&quot; hoặc &quot;Ngày chứng từ&quot;</span>
                          {isExactSalesHeader("invoiceDateCol", salesResult.detectedMapping.invoiceDateCol) ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={salesResult.detectedMapping.invoiceDateCol}
                          onChange={(e) => handleUpdateSalesMapping("invoiceDateCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.invoiceDateCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.invoiceDateCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* NHÓM 3: TÍNH DOANH THU */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-3">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-green-600 inline-block"></span>
                        3. Tính Doanh Thu Phần Mềm
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">Tổng tiền bán hàng &amp; Voucher</span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Tổng tiền */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Tổng tiền <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesTotalAmount
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Tổng tiền&quot; / Thành tiền sau thuế</span>
                          {isExactSalesHeader("totalAmountCol", salesResult.detectedMapping.totalAmountCol) ? (
                            <span className="text-green-800 font-bold flex items-center gap-0.5">
                              <Check size={10} /> Đã nhận diện
                            </span>
                          ) : null}
                        </div>
                        <select
                          value={salesResult.detectedMapping.totalAmountCol}
                          onChange={(e) => handleUpdateSalesMapping("totalAmountCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột Excel (*Bắt buộc) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.totalAmountCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.totalAmountCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {Number(v).toLocaleString("vi-VN")} đ
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Giảm giá hóa đơn */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-bold text-black">
                            Giảm giá hóa đơn
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesInvoiceDiscount
                          </span>
                        </div>
                        <div className="flex items-center justify-between text-[10px]">
                          <span className="text-slate-500">Auto-map: &quot;Giảm giá hóa đơn&quot; hoặc &quot;Voucher&quot;</span>
                          {salesResult.detectedMapping.invoiceDiscountCol ? (
                            <span className="text-blue-800 font-bold flex items-center gap-0.5">
                              Đã chọn
                            </span>
                          ) : (
                            <span className="text-slate-400">Không bắt buộc</span>
                          )}
                        </div>
                        <select
                          value={salesResult.detectedMapping.invoiceDiscountCol || ""}
                          onChange={(e) => handleUpdateSalesMapping("invoiceDiscountCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Không có / Bỏ qua (0đ) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.invoiceDiscountCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.invoiceDiscountCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {Number(v).toLocaleString("vi-VN")} đ
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>

                  {/* NHÓM 4: VALIDATION & CANCELLED */}
                  <div className="border-2 border-black p-3.5 bg-white space-y-3">
                    <div className="flex items-center justify-between border-b border-black pb-1.5">
                      <h5 className="font-black text-xs uppercase text-black flex items-center gap-1.5">
                        <span className="w-2.5 h-2.5 bg-purple-600 inline-block"></span>
                        4. Validation Kênh Bán &amp; Trạng Thái Hủy
                      </h5>
                      <span className="text-[10px] font-bold text-slate-500 uppercase">Kiểm soát đơn ngoại sàn &amp; đơn hủy</span>
                    </div>

                    <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                      {/* Kênh bán hàng */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-black text-black">
                            Kênh bán hàng <span className="text-red-600 font-black">*</span>
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            marketplaceName
                          </span>
                        </div>
                        <p className="text-[10px] text-slate-500">
                          Auto-map: &quot;Sàn TMĐT&quot; hoặc &quot;Kênh bán hàng&quot;. Cảnh báo nếu giá trị khác Shopee.
                        </p>
                        <select
                          value={salesResult.detectedMapping.marketplaceCol || ""}
                          onChange={(e) => handleUpdateSalesMapping("marketplaceCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Chọn cột (Mặc định: Tất cả là Shopee) --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.marketplaceCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.marketplaceCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>

                      {/* Trạng thái hủy */}
                      <div className="space-y-1.5">
                        <div className="flex flex-wrap items-center justify-between gap-1">
                          <label className="text-xs font-bold text-black">
                            Trạng thái hủy
                          </label>
                          <span className="text-[10px] text-slate-500 font-mono bg-slate-100 px-1 py-0.5 border border-slate-200">
                            salesIsCancelled
                          </span>
                        </div>
                        <p className="text-[10px] text-slate-500">
                          Auto-map: &quot;Hủy&quot; hoặc &quot;Trạng thái hủy&quot;. Phát hiện đơn đã hủy trên PM.
                        </p>
                        <select
                          value={salesResult.detectedMapping.cancelledCol || ""}
                          onChange={(e) => handleUpdateSalesMapping("cancelledCol", e.target.value)}
                          className="w-full border-2 border-black p-2 text-xs font-mono font-bold bg-white"
                        >
                          <option value="">-- Không có / Bỏ qua --</option>
                          {salesResult.headers.map((h, idx) => (
                            <option key={h} value={h}>
                              {getColumnLetter(idx)} — {h}
                            </option>
                          ))}
                        </select>
                        {salesResult.detectedMapping.cancelledCol && (
                          <div className="bg-[#f8f8f6] border border-slate-300 p-1.5 text-[10px] flex items-center gap-1.5 overflow-x-auto">
                            <span className="text-slate-500 font-bold shrink-0">Mẫu:</span>
                            {getSalesPreviewValues(salesResult.detectedMapping.cancelledCol).map((v, i) => (
                              <span key={i} className="bg-white px-1.5 py-0.5 font-mono font-bold border border-slate-300 rounded shrink-0">
                                {v}
                              </span>
                            ))}
                          </div>
                        )}
                      </div>
                    </div>
                  </div>
                </div>
              )}
            </div>

            {/* Modal Footer Actions */}
            <div className="flex flex-col sm:flex-row items-center justify-between gap-3 pt-3 border-t-2 border-black shrink-0">
              <button
                type="button"
                onClick={handleResetAutoDetection}
                className="w-full sm:w-auto px-4 py-2 text-xs font-bold uppercase border-2 border-black bg-white hover:bg-slate-100 text-black flex items-center justify-center gap-1.5 cursor-pointer transition"
              >
                <RotateCcw size={14} />
                <span>Tự động nhận diện lại</span>
              </button>

              <div className="flex items-center gap-2 w-full sm:w-auto">
                <button
                  type="button"
                  onClick={() => setShowColumnMappingModal(false)}
                  className="flex-1 sm:flex-none px-4 py-2.5 text-xs font-bold uppercase border-2 border-black bg-slate-200 hover:bg-slate-300 text-black cursor-pointer transition"
                >
                  Đóng
                </button>
                <button
                  type="button"
                  onClick={() => {
                    const shopeeM = marketplaceResult.sheets[0]?.detectedMapping;
                    if (!shopeeM?.orderIdCol || !shopeeM?.buyerPaidCol || !shopeeM?.orderStatusCol || !shopeeM?.returnRefundCol) {
                      showToast("Vui lòng cấu hình đủ các trường bắt buộc (*) cho File Shopee!", "warning");
                      setMappingModalTab("shopee");
                      return;
                    }
                    if (!salesResult.detectedMapping.shopNameCol || !salesResult.detectedMapping.orderIdCol || !salesResult.detectedMapping.totalAmountCol) {
                      showToast("Vui lòng cấu hình đủ các trường bắt buộc (*) cho File Phần mềm bán hàng!", "warning");
                      setMappingModalTab("sales");
                      return;
                    }
                    setShowColumnMappingModal(false);
                    showToast("Đã lưu cấu hình mapping cột và sẵn sàng đối chiếu!", "success");
                  }}
                  className="flex-1 sm:flex-none px-6 py-2.5 bg-[#00ff00] hover:bg-[#05e005] text-black font-black uppercase text-xs border-2 border-black shadow-[3px_3px_0px_#141414] hover:shadow-[5px_5px_0px_#141414] transition cursor-pointer flex items-center justify-center gap-1.5"
                >
                  <CheckCircle2 size={16} />
                  <span>Xác nhận &amp; Áp dụng mapping</span>
                </button>
              </div>
            </div>
          </div>
        </div>
      )}

      {/* MODAL MAPPING SHOP THỦ CÔNG (Chỉ hiển thị khi có nghi ngờ nhầm shop hoặc người dùng mở) */}
      {showShopMappingModal && marketplaceResult && salesResult && (
        <div className="fixed inset-0 bg-[#141414]/80 z-50 flex items-center justify-center p-4">
          <div className="bg-white border-4 border-[#141414] max-w-xl w-full p-6 shadow-[8px_8px_0px_#141414] space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex justify-between items-center border-b-2 border-black pb-2">
              <h4 className="font-black text-xs uppercase tracking-wider text-black">
                Kiểm tra ghép cửa hàng (Shop Mapping)
              </h4>
              <button
                onClick={() => setShowShopMappingModal(false)}
                className="text-xs font-bold px-2 py-1 bg-slate-200 hover:bg-slate-300"
              >
                ✕ Đóng
              </button>
            </div>

            <div className="overflow-x-auto">
              <table className="w-full text-left text-xs border-collapse font-mono">
                <thead>
                  <tr className="bg-slate-100 border-b border-black text-[11px] font-sans font-bold">
                    <th className="p-2">Sheet / Shop Sàn</th>
                    <th className="p-2">Tên Shop Bán Hàng</th>
                  </tr>
                </thead>
                <tbody className="divide-y divide-slate-200">
                  {shopMappings.map((m, idx) => (
                    <tr key={m.marketplaceSheetOrShop}>
                      <td className="p-2 font-bold">{m.marketplaceSheetOrShop}</td>
                      <td className="p-2">
                        <select
                          value={m.salesShopName}
                          onChange={(e) => {
                            const updated = [...shopMappings];
                            updated[idx] = { ...updated[idx], salesShopName: e.target.value, isConfirmed: true };
                            setShopMappings(updated);
                          }}
                          className="w-full border border-black p-1 text-xs font-bold bg-white"
                        >
                          {salesResult.detectedShops.map(s => (
                            <option key={s} value={s}>{s}</option>
                          ))}
                        </select>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>

            <div className="text-right pt-2">
              <button
                onClick={() => {
                  saveShopMappings(shopMappings);
                  setShowShopMappingModal(false);
                  showToast("Đã lưu và ghi nhớ mapping cửa hàng!");
                }}
                className="bg-[#141414] text-white hover:bg-black font-black uppercase text-xs px-5 py-2 border-2 border-black"
              >
                Xác nhận &amp; Ghi nhớ
              </button>
            </div>
          </div>
        </div>
      )}

      {/* MODAL DRILL-DOWN CHI TIẾT ĐƠN HÀNG */}
      {selectedRecordDetail && (
        <div className="fixed inset-0 bg-[#141414]/80 z-50 flex items-center justify-center p-4">
          <div className="bg-white border-4 border-[#141414] max-w-3xl w-full p-6 shadow-[8px_8px_0px_#141414] space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex justify-between items-start border-b-2 border-[#141414] pb-3">
              <div>
                <span className="text-[10px] bg-[#141414] text-[#00ff00] px-2 py-0.5 font-mono font-black uppercase inline-block mb-1">
                  Chi tiết đối chiếu đơn hàng
                </span>
                <h4 className="text-base font-black uppercase text-black font-mono">
                  {selectedRecordDetail.originalOrderId} • {selectedRecordDetail.shopDisplayName}
                </h4>
              </div>

              <span className={`px-2.5 py-1 border border-black font-black text-xs font-mono ${
                selectedRecordDetail.status === "KHỚP"
                  ? "bg-[#00ff00] text-black"
                  : selectedRecordDetail.status === "CHÊNH LỆCH"
                  ? "bg-red-500 text-white"
                  : "bg-amber-300 text-black"
              }`}>
                {selectedRecordDetail.status}
              </span>
            </div>

            {/* Chi tiết dữ liệu Sàn Shopee */}
            <div className="bg-slate-50 border-2 border-black p-4 space-y-2">
              <h5 className="font-black text-xs uppercase text-orange-600 flex items-center gap-1.5">
                <Store size={14} />
                <span>Dữ liệu sàn TMĐT ({selectedRecordDetail.marketplaceRowCount} dòng SKU gốc)</span>
              </h5>

              {selectedRecordDetail.hasMarketplace ? (
                <div className="space-y-2 text-xs font-mono">
                  <div className="grid grid-cols-2 sm:grid-cols-4 gap-2 bg-white p-2 border border-slate-300">
                    <div><span className="text-slate-500">Trạng thái:</span> <strong className="text-black">{selectedRecordDetail.marketplaceStatus}</strong></div>
                    <div><span className="text-slate-500">Giao hàng:</span> <strong className="text-black">{selectedRecordDetail.marketplaceDeliveryDate || "--"}</strong></div>
                    <div><span className="text-slate-500">Người mua trả:</span> <strong className="text-black">{selectedRecordDetail.marketplaceBuyerPaid.toLocaleString("vi-VN")}đ</strong></div>
                    <div><span className="text-slate-500">Voucher Shop:</span> <strong className="text-orange-600">-{selectedRecordDetail.marketplaceVoucher.toLocaleString("vi-VN")}đ</strong></div>
                  </div>

                  <div className="text-[11px] text-slate-700 bg-amber-50 p-2 border border-amber-200">
                    <strong>Công thức doanh thu:</strong> Doanh thu sàn = {selectedRecordDetail.marketplaceBuyerPaid.toLocaleString("vi-VN")}đ (Người mua trả) - {selectedRecordDetail.marketplaceVoucher.toLocaleString("vi-VN")}đ (Voucher shop trừ 1 lần) = <strong className="text-black font-black">{selectedRecordDetail.marketplaceRevenue.toLocaleString("vi-VN")}đ</strong>
                  </div>

                  {selectedRecordDetail.marketplaceData?.sourceRows && (
                    <div className="overflow-x-auto border border-slate-300">
                      <table className="w-full text-left text-[10px]">
                        <thead className="bg-slate-200">
                          <tr>
                            <th className="p-1.5">Sản phẩm / SKU</th>
                            <th className="p-1.5 text-right">Giá</th>
                            <th className="p-1.5 text-center">SL</th>
                            <th className="p-1.5 text-right">Người mua trả</th>
                            <th className="p-1.5 text-right">Voucher shop</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {selectedRecordDetail.marketplaceData.sourceRows.map((r, i) => (
                            <tr key={i}>
                              <td className="p-1.5">{r["Tên sản phẩm"] || r["sku"] || `Dòng SKU ${i + 1}`}</td>
                              <td className="p-1.5 text-right">{Number(r["Giá sản phẩm"] || 0).toLocaleString("vi-VN")}đ</td>
                              <td className="p-1.5 text-center">{r["Số lượng"] || 1}</td>
                              <td className="p-1.5 text-right font-bold">{Number(r["Tổng số tiền Người mua thanh toán"] || 0).toLocaleString("vi-VN")}đ</td>
                              <td className="p-1.5 text-right text-orange-600">{Number(r["Mã giảm giá của Shop"] || 0).toLocaleString("vi-VN")}đ</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              ) : (
                <div className="text-xs text-red-600 font-bold">Không tìm thấy mã đơn này trên file sàn TMĐT.</div>
              )}
            </div>

            {/* Chi tiết dữ liệu Phần Mềm Bán Hàng */}
            <div className="bg-slate-50 border-2 border-black p-4 space-y-2">
              <h5 className="font-black text-xs uppercase text-blue-700 flex items-center gap-1.5">
                <FileSpreadsheet size={14} />
                <span>Dữ liệu phần mềm bán hàng ({selectedRecordDetail.salesRowCount} dòng chứng từ)</span>
              </h5>

              {selectedRecordDetail.hasSales ? (
                <div className="space-y-2 text-xs font-mono">
                  <div className="grid grid-cols-2 sm:grid-cols-3 gap-2 bg-white p-2 border border-slate-300">
                    <div><span className="text-slate-500">Số hóa đơn:</span> <strong className="text-black">{selectedRecordDetail.salesInvoiceNumber || "--"}</strong></div>
                    <div><span className="text-slate-500">Ngày hóa đơn:</span> <strong className="text-black">{selectedRecordDetail.salesInvoiceDate || "--"}</strong></div>
                    <div><span className="text-slate-500">Tổng tiền PM:</span> <strong className="text-blue-900 font-black">{selectedRecordDetail.salesTotalAmount.toLocaleString("vi-VN")}đ</strong></div>
                  </div>

                  {selectedRecordDetail.salesData?.sourceRows && (
                    <div className="overflow-x-auto border border-slate-300">
                      <table className="w-full text-left text-[10px]">
                        <thead className="bg-slate-200">
                          <tr>
                            <th className="p-1.5">Số hóa đơn</th>
                            <th className="p-1.5">Ngày hóa đơn</th>
                            <th className="p-1.5">Số chứng từ</th>
                            <th className="p-1.5 text-right">Tổng tiền</th>
                            <th className="p-1.5">Ghi chú</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-slate-200">
                          {selectedRecordDetail.salesData.sourceRows.map((r, i) => (
                            <tr key={i}>
                              <td className="p-1.5 font-bold">{r["Số hóa đơn"] || "--"}</td>
                              <td className="p-1.5">{r["Ngày hóa đơn"] || "--"}</td>
                              <td className="p-1.5">{r["Số chứng từ"] || "--"}</td>
                              <td className="p-1.5 text-right font-black">{Number(r["Tổng tiền"] || 0).toLocaleString("vi-VN")}đ</td>
                              <td className="p-1.5 text-slate-500">{r["Ghi chú"] || "--"}</td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  )}
                </div>
              ) : (
                <div className="text-xs text-red-600 font-bold">Không tìm thấy mã đơn này trên file phần mềm bán hàng.</div>
              )}
            </div>

            <div className="text-right pt-2">
              <button
                onClick={() => setSelectedRecordDetail(null)}
                className="bg-[#141414] text-white hover:bg-black font-black uppercase text-xs px-5 py-2 border-2 border-black shadow-[2px_2px_0px_#00ff00]"
              >
                Đóng
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
  );
};
