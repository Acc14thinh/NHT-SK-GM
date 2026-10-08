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
  ShopMappingEntry
} from "../../lib/ecommerce/types";
import {
  parseMarketplaceWorkbook,
  ParsedMarketplaceResult
} from "../../lib/ecommerce/shopeeParser";
import {
  parseSalesSoftwareWorkbook,
  ParsedSalesSoftwareResult
} from "../../lib/ecommerce/salesSoftwareParser";
import {
  buildProposedShopMappings,
  saveShopMappings,
  saveHeaderMappings
} from "../../lib/ecommerce/shopMapper";
import {
  aggregateMarketplaceOrders,
  aggregateSalesOrders
} from "../../lib/ecommerce/aggregator";
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
  const [showShopMappingModal, setShowShopMappingModal] = useState<boolean>(false);

  // Files & Parsed data
  const [marketplaceResult, setMarketplaceResult] = useState<ParsedMarketplaceResult | null>(null);
  const [salesResult, setSalesResult] = useState<ParsedSalesSoftwareResult | null>(null);
  const [shopMappings, setShopMappings] = useState<ShopMappingEntry[]>([]);

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

  // Tự động nhận diện và ghép shop khi có đủ 2 nguồn (Section 32, 35)
  useEffect(() => {
    if (marketplaceResult && salesResult) {
      const proposed = buildProposedShopMappings(
        marketplaceResult.allDetectedShopNames,
        salesResult.detectedShops
      );
      setShopMappings(proposed);
    }
  }, [marketplaceResult, salesResult]);

  // Kiểm tra tính sẵn sàng & lỗi chặn (Section 39)
  const validationStatus = useMemo(() => {
    if (!marketplaceResult || !salesResult) {
      return { isReady: false, blockingError: null, warningText: null };
    }

    // 1. Kiểm tra cột Mã đơn hàng trên file sàn
    const mMissingOrderId = marketplaceResult.sheets.some(s => !s.detectedMapping.orderIdCol);
    if (mMissingOrderId) {
      return {
        isReady: false,
        blockingError: "Không tìm thấy cột [Mã đơn hàng] trong file Shopee.",
        actionType: "column_mkt" as const
      };
    }

    // 2. Kiểm tra cột Tên Shop trên file phần mềm
    if (!salesResult.detectedMapping.shopNameCol) {
      return {
        isReady: false,
        blockingError: "Không tìm thấy cột [Tên Shop] trong file phần mềm bán hàng.",
        actionType: "column_sales" as const
      };
    }

    // 3. Kiểm tra cột Mã đơn trên file phần mềm
    if (!salesResult.detectedMapping.orderIdCol) {
      return {
        isReady: false,
        blockingError: "Không tìm thấy cột [Số đơn hàng từ hệ thống khác / Mã đơn] trong file phần mềm.",
        actionType: "column_sales" as const
      };
    }

    // 4. Kiểm tra cột Tổng tiền trên file phần mềm
    if (!salesResult.detectedMapping.totalAmountCol) {
      return {
        isReady: false,
        blockingError: "Không tìm thấy cột [Tổng tiền / Thành tiền] trong file phần mềm bán hàng.",
        actionType: "column_sales" as const
      };
    }

    // 5. Kiểm tra nếu có nhiều shop mà chưa ghép được shop nào
    if (
      marketplaceResult.sheets.length > 1 &&
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
  const handleMarketplaceFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "array" });
        const parsed = parseMarketplaceWorkbook(wb, file.name);
        setMarketplaceResult(parsed);
        setReconciliationRecords(null);
        showToast(`Đã nhận diện file Shopee: ${parsed.totalRows.toLocaleString()} dòng – ${parsed.sheets.length} shop`);
      } catch (err) {
        console.error(err);
        showToast("Lỗi khi đọc file sàn TMĐT. Vui lòng kiểm tra định dạng Excel.", "warning");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // Xử lý nạp File Phần Mềm Bán Hàng (Tự động phân tích ngầm ngay khi chọn file)
  const handleSalesFileChange = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const buffer = evt.target?.result;
        const wb = XLSX.read(buffer, { type: "array" });
        const parsed = parseSalesSoftwareWorkbook(wb, file.name);
        setSalesResult(parsed);
        setReconciliationRecords(null);
        showToast(`Đã nhận diện file bán hàng: ${parsed.rows.length.toLocaleString()} dòng – ${parsed.detectedShops.length} shop`);
      } catch (err) {
        console.error(err);
        showToast("Lỗi khi đọc file phần mềm bán hàng.", "warning");
      }
    };
    reader.readAsArrayBuffer(file);
  };

  // 1-Click Nạp Dữ Liệu Demo Mẫu (14 trường hợp nghiệp vụ)
  const handleLoadDemoData = () => {
    const { shopeeWorkbook, shopeeFileName, salesWorkbook, salesFileName } = generateDemoEcommerceWorkbooks();
    const parsedM = parseMarketplaceWorkbook(shopeeWorkbook, shopeeFileName);
    const parsedS = parseSalesSoftwareWorkbook(salesWorkbook, salesFileName);

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

  // BƯỚC 2: Thực hiện đối chiếu (Có hiển thị tiến trình theo Section 40)
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
    setProcessingProgress("Đang khởi tạo index Shop + Mã đơn...");

    setTimeout(() => {
      try {
        saveShopMappings(shopMappings);

        const totalOrdersEstimate = marketplaceResult.totalUniqueOrders + salesResult.diagnostics.uniqueOrderCount;
        setProcessingProgress(`Đang xử lý ${totalOrdersEstimate.toLocaleString()} đơn hàng hai chiều...`);

        // 1. Group dữ liệu sàn theo Shop + Order ID
        const aggregatedMarketplace = aggregateMarketplaceOrders(
          marketplaceResult.sheets,
          shopMappings
        );

        // 2. Group dữ liệu bán hàng theo Shop + Order ID
        const aggregatedSales = aggregateSalesOrders(
          salesResult.rows,
          salesResult.detectedMapping
        );

        // 3. Thực hiện đối chiếu 2 chiều deterministic O(n)
        const { records, summary } = reconcileEcommerceRevenue(
          aggregatedMarketplace,
          aggregatedSales,
          config
        );

        setReconciliationRecords(records);
        setReconciliationSummary(summary);
        setIsProcessing(false);
        setProcessingProgress("");
        showToast(`Đối chiếu hoàn tất: ${summary.matchedCount} đơn khớp, ${summary.differenceCount} đơn chênh lệch!`);
      } catch (err: any) {
        console.error(err);
        setIsProcessing(false);
        setProcessingProgress("");
        showToast(`Lỗi trong quá trình đối chiếu: ${err.message || err}`, "warning");
      }
    }, 400);
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

  // Làm mới (Xóa 2 file & kết quả hiện tại, giữ nguyên mapping đã lưu trong localStorage - Section 43)
  const handleReset = () => {
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
    showToast("Đã làm mới phiên đối chiếu.", "info");
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
                    ✓ {salesResult.rows.length.toLocaleString()} dòng – {salesResult.detectedShops.length} shop
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
                    ✓ File Shopee: {marketplaceResult.totalRows.toLocaleString()} dòng ({marketplaceResult.sheets.length} shop) •
                    ✓ File bán hàng: {salesResult.rows.length.toLocaleString()} dòng ({salesResult.detectedShops.length} shop) •
                    ✓ Đã nhận diện đủ cột bắt buộc
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

      {/* MODAL MAPPING CỘT THỦ CÔNG (Chỉ hiển thị khi người dùng yêu cầu hoặc thiếu cột - Section 32, 38) */}
      {showColumnMappingModal && marketplaceResult && salesResult && (
        <div className="fixed inset-0 bg-[#141414]/80 z-50 flex items-center justify-center p-4">
          <div className="bg-white border-4 border-[#141414] max-w-2xl w-full p-6 shadow-[8px_8px_0px_#141414] space-y-4 max-h-[90vh] overflow-y-auto">
            <div className="flex justify-between items-center border-b-2 border-black pb-2">
              <h4 className="font-black text-xs uppercase tracking-wider text-black">
                Tùy chỉnh chọn cột dữ liệu (Column Mapping)
              </h4>
              <button
                onClick={() => setShowColumnMappingModal(false)}
                className="text-xs font-bold px-2 py-1 bg-slate-200 hover:bg-slate-300"
              >
                ✕ Đóng
              </button>
            </div>

            {/* File Sàn */}
            <div className="space-y-2">
              <span className="text-[10px] font-black uppercase text-orange-600 block">Cột File Sàn Shopee</span>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Mã đơn hàng (*Bắt buộc)</label>
                  <select
                    value={marketplaceResult.sheets[0]?.detectedMapping.orderIdCol || ""}
                    onChange={(e) => {
                      const updated = { ...marketplaceResult };
                      updated.sheets.forEach(s => s.detectedMapping.orderIdCol = e.target.value);
                      setMarketplaceResult(updated);
                      saveHeaderMappings("marketplace", { orderIdCol: e.target.value });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {marketplaceResult.sheets[0]?.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Tiền người mua thanh toán</label>
                  <select
                    value={marketplaceResult.sheets[0]?.detectedMapping.buyerPaidCol || ""}
                    onChange={(e) => {
                      const updated = { ...marketplaceResult };
                      updated.sheets.forEach(s => s.detectedMapping.buyerPaidCol = e.target.value);
                      setMarketplaceResult(updated);
                      saveHeaderMappings("marketplace", { buyerPaidCol: e.target.value });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {marketplaceResult.sheets[0]?.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Mã giảm giá của Shop</label>
                  <select
                    value={marketplaceResult.sheets[0]?.detectedMapping.shopVoucherCol || ""}
                    onChange={(e) => {
                      const updated = { ...marketplaceResult };
                      updated.sheets.forEach(s => s.detectedMapping.shopVoucherCol = e.target.value);
                      setMarketplaceResult(updated);
                      saveHeaderMappings("marketplace", { shopVoucherCol: e.target.value });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Không có (0đ) --</option>
                    {marketplaceResult.sheets[0]?.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Trạng thái đơn hàng</label>
                  <select
                    value={marketplaceResult.sheets[0]?.detectedMapping.orderStatusCol || ""}
                    onChange={(e) => {
                      const updated = { ...marketplaceResult };
                      updated.sheets.forEach(s => s.detectedMapping.orderStatusCol = e.target.value);
                      setMarketplaceResult(updated);
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Tự động --</option>
                    {marketplaceResult.sheets[0]?.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            {/* File Phần mềm bán hàng */}
            <div className="space-y-2 pt-2 border-t border-slate-200">
              <span className="text-[10px] font-black uppercase text-blue-700 block">Cột File Phần Mềm Bán Hàng</span>
              <div className="grid grid-cols-1 sm:grid-cols-2 gap-3 text-xs">
                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Tên Shop (*Bắt buộc)</label>
                  <select
                    value={salesResult.detectedMapping.shopNameCol}
                    onChange={(e) => {
                      setSalesResult({
                        ...salesResult,
                        detectedMapping: { ...salesResult.detectedMapping, shopNameCol: e.target.value }
                      });
                      saveHeaderMappings("sales", { shopNameCol: e.target.value });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {salesResult.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Số đơn hàng từ HT khác / Mã đơn (*Bắt buộc)</label>
                  <select
                    value={salesResult.detectedMapping.orderIdCol}
                    onChange={(e) => {
                      setSalesResult({
                        ...salesResult,
                        detectedMapping: { ...salesResult.detectedMapping, orderIdCol: e.target.value }
                      });
                      saveHeaderMappings("sales", { orderIdCol: e.target.value });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {salesResult.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Tổng tiền đối chiếu (*Bắt buộc)</label>
                  <select
                    value={salesResult.detectedMapping.totalAmountCol}
                    onChange={(e) => {
                      setSalesResult({
                        ...salesResult,
                        detectedMapping: { ...salesResult.detectedMapping, totalAmountCol: e.target.value }
                      });
                      saveHeaderMappings("sales", { totalAmountCol: e.target.value });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {salesResult.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>

                <div>
                  <label className="text-[10px] font-bold block mb-0.5">Số hóa đơn</label>
                  <select
                    value={salesResult.detectedMapping.invoiceNumberCol}
                    onChange={(e) => {
                      setSalesResult({
                        ...salesResult,
                        detectedMapping: { ...salesResult.detectedMapping, invoiceNumberCol: e.target.value }
                      });
                    }}
                    className="w-full border border-black p-1.5 text-xs font-mono font-bold"
                  >
                    <option value="">-- Chọn cột --</option>
                    {salesResult.headers.map(h => (
                      <option key={h} value={h}>{h}</option>
                    ))}
                  </select>
                </div>
              </div>
            </div>

            <div className="text-right pt-2">
              <button
                onClick={() => {
                  setShowColumnMappingModal(false);
                  showToast("Đã lưu cấu hình mapping cột!");
                }}
                className="bg-[#141414] text-white hover:bg-black font-black uppercase text-xs px-5 py-2 border-2 border-black"
              >
                Xác nhận
              </button>
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
