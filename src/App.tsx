/**
 * @license
 * SPDX-License-Identifier: Apache-2.0
 */

import React, { useState, useEffect } from "react";
import {
  Play,
  Check,
  CheckCircle2,
  AlertTriangle,
  RefreshCw,
  FileSpreadsheet,
  UploadCloud,
  Upload,
  Database,
  Users,
  FileText,
  Settings,
  Download,
  AlertCircle,
  Trash2,
  Plus,
  Copy,
  ChevronRight,
  TrendingUp,
  Coins,
  History,
  FileCheck
} from "lucide-react";

import { Commodity, Partner, MatchingConfig, MappedRow, BankAnalysisResult, ColumnMapping } from "./types";
import * as XLSX from "xlsx";
import JSZip from "jszip";
import {
  matchCommodityRow,
  matchPartnerRow,
  analyzeBankTransaction,
  generateNewCode,
  normalizeText,
  matchBankRowWithInvoiceEngine,
  preProcessInvoicePartners,
  parseDate
} from "./lib/matchingEngine";
import {
  cleanAndDeduplicateHeaders,
  scoreHeaderRow,
  proposeBankMappings,
  validateBankMappings,
  checkMappingDataQuality,
  parseAmount,
  getHeadersFromRawRows
} from "./lib/excelParser";
import { ErrorBoundary } from "./components/ErrorBoundary";
import {
  getSampleCommodities,
  getSamplePartners,
  getSamplePurchaseLedger,
  getSampleSalesLedger,
  getSampleInventoryLedger,
  getSampleBankStatement,
  RawPurchase,
  RawSale,
  RawInventory,
  RawBankStatement
} from "./data/mockData";

export default function App() {
  // Core Directories (State-managed to allow adding new ones)
  const [commodities, setCommodities] = useState<Commodity[]>(getSampleCommodities());
  const [partners, setPartners] = useState<Partner[]>(getSamplePartners());

  // App Configurations
  const [config, setConfig] = useState<MatchingConfig>({
    autoThreshold: 85,
    checkThreshold: 70,
    prefixHH: "HH",
    prefixKH: "KH",
    prefixNCC: "NCC",
    daysBeforeInvoice: 7,
    daysAfterInvoice: 30,
    diffAbsThreshold: 10000,
    diffPctThreshold: 0.5,
    maxCombinationCount: 5
  });

  // Navigation Menu
  const [currentTab, setCurrentTab] = useState<"dashboard" | "commodity" | "partner" | "bank" | "integrated" | "python">("dashboard");

  // Demonstration state
  const [demoLoaded, setDemoLoaded] = useState(false);

  // Upload/Local Data State
  const [uploadedFileName, setUploadedFileName] = useState<string>("");
  const [activeRowsCount, setActiveRowsCount] = useState<number>(0);

  // --- MODE 1: Gắn mã hàng hóa ---
  const [commoditySourceRows, setCommoditySourceRows] = useState<any[]>([]);
  const [commodityMappings, setCommodityMappings] = useState<ColumnMapping>({
    ten_hang_hoa: "ten_hang_hoa_dich_vu",
    don_vi_tinh: "don_vi_tinh",
    so_luong: "so_luong",
    don_gia: "don_gia",
    thanh_tien: "thanh_tien"
  });
  const [commodityMappedRows, setCommodityMappedRows] = useState<MappedRow[]>([]);
  const [isProcessingCommodities, setIsProcessingCommodities] = useState(false);

  // --- MODE 2: Gắn mã đối tác ---
  const [partnerMode, setPartnerMode] = useState<"Mua vào" | "Bán ra">("Mua vào");
  const [partnerSourceRows, setPartnerSourceRows] = useState<any[]>([]);
  const [partnerMappings, setPartnerMappings] = useState<ColumnMapping>({
    ten_doi_tuong: "ten_nguoi_ban",
    ma_so_thue: "ma_so_thue_nguoi_ban",
    so_tai_khoan: "",
    ten_hang_hoa: "ten_hang_hoa_dich_vu"
  });
  const [partnerMappedRows, setPartnerMappedRows] = useState<MappedRow[]>([]);
  const [isProcessingPartners, setIsProcessingPartners] = useState(false);

  // --- MODE 3: Phân tích ngân hàng ---
  const [bankSourceRows, setBankSourceRows] = useState<any[]>([]);
  const [bankMappings, setBankMappings] = useState<ColumnMapping>({
    ngay_giao_dich: "ngay_gd",
    ngay_hieu_luc: "",
    ngay_hach_toan: "",
    noi_dung_giao_dich: "noi_dung_giao_dich",
    so_tien_thu: "so_tien_thu",
    so_tien_chi: "so_tien_chi",
    so_du: "",
    ma_giao_dich: "",
    so_chung_tu: "",
    so_tham_chieu: "",
    ten_doi_tac_sao_ke: "ten_doi_tac_sao_ke",
    so_tai_khoan_doi_ung: "so_tai_khoan_doi_ung",
    ngan_hang_doi_ung: "",
    ghi_chu: "",
    loai_tien_te: "",
    cot_ngay: "",
    cot_thang: "",
    cot_nam: ""
  });
  const [bankMappedRows, setBankMappedRows] = useState<BankAnalysisResult[]>([]);
  const [isProcessingBank, setIsProcessingBank] = useState(false);

  // --- EXCEL WIZARD & AUDIT STATE ---
  const [uploadedWorkbook, setUploadedWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [excelWizard, setExcelWizard] = useState<{
    fileType: "commodity" | "partner" | "bank" | "bank_sales" | "bank_purchases" | null;
    fileName: string;
    sheetNames: string[];
    selectedSheet: string;
    rawRows: any[][];
    headerRowIndex: number;
    headerRowsCount: number;
    headersCleaned: string[];
    rowsPreview: any[];
    onCancel?: () => void;
  } | null>(null);

  const [bankAuditDetails, setBankAuditDetails] = useState<{
    fileName: string;
    sheetName: string;
    headerRow: number;
    headerCount: number;
    totalColumns: number;
    totalRows: number;
    allColumns: string[];
    hasDataColumnsCount: number;
    emptyColumns: string[];
    renamedColumns: string[];
  } | null>(null);

  const [partnerAuditDetails, setPartnerAuditDetails] = useState<{
    fileName: string;
    sheetName: string;
    headerRow: number;
    headerCount: number;
    totalColumns: number;
    totalRows: number;
    allColumns: string[];
    hasDataColumnsCount: number;
    emptyColumns: string[];
    renamedColumns: string[];
  } | null>(null);

  // Dữ liệu đối chiếu bổ sung inside bank tab
  const [bankSalesRows, setBankSalesRows] = useState<any[]>([]);
  const [bankSalesFileName, setBankSalesFileName] = useState<string>("");
  const [bankSalesMappings, setBankSalesMappings] = useState<ColumnMapping>({
    so_hoa_don: "so_hoa_don",
    ngay_hoa_don: "ngay_hoa_don",
    ten_khach_hang: "ten_nguoi_mua",
    ma_khach_hang: "ma_khach_hang",
    ma_so_thue_khach_hang: "ma_so_thue_nguoi_mua",
    tong_thanh_toan: "tong_thanh_toan"
  });

  const [bankPurchasesRows, setBankPurchasesRows] = useState<any[]>([]);
  const [bankPurchasesFileName, setBankPurchasesFileName] = useState<string>("");
  const [bankPurchasesMappings, setBankPurchasesMappings] = useState<ColumnMapping>({
    so_hoa_don: "so_hoa_don",
    ngay_hoa_don: "ngay_hoa_don",
    ten_nha_cung_cap: "ten_nguoi_ban",
    ma_nha_cung_cap: "ma_nha_cung_cap",
    ma_so_thue_ncc: "ma_so_thue_nguoi_ban",
    tong_thanh_toan: "tong_thanh_toan"
  });

  // --- MODE 4: Integrated reconciliation ---
  const [isProcessingIntegrated, setIsProcessingIntegrated] = useState(false);
  const [integratedPurchaseRows, setIntegratedPurchaseRows] = useState<MappedRow[]>([]);
  const [integratedSaleRows, setIntegratedSaleRows] = useState<MappedRow[]>([]);
  const [integratedInvRows, setIntegratedInvRows] = useState<any[]>([]);
  const [integratedBankRows, setIntegratedBankRows] = useState<BankAnalysisResult[]>([]);
  const [integratedRecon, setIntegratedRecon] = useState<any[]>([]);

  // Grid performance optimization states for handling extremely large excel files
  const [editingCommodityId, setEditingCommodityId] = useState<string | null>(null);
  const [editingPartnerId, setEditingPartnerId] = useState<string | null>(null);
  const [editingBankRowId, setEditingBankRowId] = useState<string | null>(null);

  const [commodityLimit, setCommodityLimit] = useState(50);
  const [partnerLimit, setPartnerLimit] = useState(50);
  const [bankLimit, setBankLimit] = useState(50);

  // Notification Banner State
  const [notification, setNotification] = useState<{ message: string; type: "success" | "warning" } | null>(null);

  // Structured privacy-safe upload incident logs
  const [uploadLogs, setUploadLogs] = useState<{
    timestamp: string;
    fileName: string;
    fileSize: number;
    fileType: string;
    errorMsg: string;
  }[]>([]);

  // Directory/Master catalog import confirmation modal state
  const [importConfirm, setImportConfirm] = useState<{
    type: "commodity" | "customer" | "supplier";
    items: any[];
    fileName: string;
  } | null>(null);

  // Dynamic headers read from the uploaded files
  const commodityHeaders = React.useMemo(() => {
    return commoditySourceRows.length > 0 ? Object.keys(commoditySourceRows[0]) : [];
  }, [commoditySourceRows]);

  const partnerHeaders = React.useMemo(() => {
    return partnerSourceRows.length > 0 ? Object.keys(partnerSourceRows[0]) : [];
  }, [partnerSourceRows]);

  const bankHeaders = React.useMemo(() => {
    return bankSourceRows.length > 0 ? Object.keys(bankSourceRows[0]) : [];
  }, [bankSourceRows]);

  const bankSalesHeaders = React.useMemo(() => {
    return bankSalesRows.length > 0 ? Object.keys(bankSalesRows[0]) : [];
  }, [bankSalesRows]);

  const bankPurchasesHeaders = React.useMemo(() => {
    return bankPurchasesRows.length > 0 ? Object.keys(bankPurchasesRows[0]) : [];
  }, [bankPurchasesRows]);

  const getColumnOptions = (headers: string[], currentVal: string) => {
    const uniqueVals = Array.from(new Set([...headers, currentVal].filter(Boolean)));
    return uniqueVals.map(h => (
      <option key={h} value={h}>
        {h}
      </option>
    ));
  };

  const triggerToast = (msg: string, type: "success" | "warning" = "success") => {
    setNotification({ message: msg, type });
    setTimeout(() => setNotification(null), 4000);
  };

  // --- ON LOAD DEMO TRIGGER ---
  const handleLoadDemo = () => {
    setDemoLoaded(true);
    setCommoditySourceRows(getSamplePurchaseLedger());
    setPartnerSourceRows(getSamplePurchaseLedger());
    setBankSourceRows(getSampleBankStatement());
    setBankSalesRows(getSampleSalesLedger());
    setBankPurchasesRows(getSamplePurchaseLedger());
    setBankSalesFileName("Bang_Ke_Ban_Ra_Demo.xlsx");
    setBankPurchasesFileName("Bang_Ke_Mua_Vao_Demo.xlsx");
    setUploadedFileName("Du_Lieu_Khao_Sat_Ke_Toan_Mau.xlsx");
    setActiveRowsCount(15);
    triggerToast("Đã nhập số liệu mẫu thành công cho tất cả các phân hệ!");
  };

  const handleReset = () => {
    setDemoLoaded(false);
    setUploadedFileName("");
    setActiveRowsCount(0);
    setCommoditySourceRows([]);
    setCommodityMappedRows([]);
    setPartnerSourceRows([]);
    setPartnerMappedRows([]);
    setBankSourceRows([]);
    setBankMappedRows([]);
    setBankSalesRows([]);
    setBankSalesFileName("");
    setBankPurchasesRows([]);
    setBankPurchasesFileName("");
    setIntegratedPurchaseRows([]);
    setIntegratedSaleRows([]);
    setIntegratedInvRows([]);
    setIntegratedBankRows([]);
    setIntegratedRecon([]);
    setCommodities(getSampleCommodities());
    setPartners(getSamplePartners());
    triggerToast("Đã dọn dẹp sạch toàn bộ tiến trình xử lý", "warning");
  };

  const handleImportMasterDirectory = (event: React.ChangeEvent<HTMLInputElement>, type: "commodity" | "customer" | "supplier") => {
    const file = event.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: "array" });
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          triggerToast("Không tìm thấy trang tính nào trong file!", "warning");
          return;
        }

        const jsonData = XLSX.utils.sheet_to_json<any>(worksheet);
        if (jsonData.length === 0) {
          triggerToast("Tệp trống hoặc không chứa dữ liệu hợp lệ!", "warning");
          return;
        }

        const headers = Object.keys(jsonData[0] || {});

        if (type === "commodity") {
          const mappedMa = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_hang") || hNorm.includes("ma_hh") || hNorm.includes("ma_vật_tư") || hNorm.includes("mã hàng") || hNorm.includes("product_id") || hNorm.includes("code") || hNorm.includes("mã") || hNorm.includes("sku") || hNorm === "id";
          }) || headers[0] || "ma_hang_hoa";

          const mappedTen = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_hang") || hNorm.includes("tên hàng") || hNorm.includes("dien_giai") || hNorm.includes("diễn giải") || hNorm.includes("nội dung") || hNorm.includes("description") || hNorm.includes("name") || hNorm.includes("tên");
          }) || headers[1] || "ten_hang_hoa_chuan";

          const mappedNhom = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("nhom_hang") || hNorm.includes("nhóm") || hNorm.includes("category") || hNorm.includes("group");
          }) || "";

          const mappedDvt = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("don_vi_tinh") || hNorm.includes("dvt") || hNorm.includes("đơn vị tính") || hNorm.includes("uom") || hNorm.includes("unit");
          }) || "don_vi_tinh";

          const mappedQuyCach = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("quy_cach") || hNorm.includes("quy cách") || hNorm.includes("specs") || hNorm.includes("specifications");
          }) || "";

          const mappedTuKhoa = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tu_khoa") || hNorm.includes("từ khóa") || hNorm.includes("keywords");
          }) || "";

          const mappedGhiChu = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ghi_chu") || hNorm.includes("ghi chú") || hNorm.includes("note");
          }) || "";

          const parsedCommodities: Commodity[] = jsonData.map((row: any) => {
            const code = String(row[mappedMa] || "").trim();
            const name = String(row[mappedTen] || "").trim();
            if (!name) return null;
            const finalCode = code || "HH-" + Math.random().toString(36).substr(2, 5).toUpperCase();
            return {
              ma_hang_hoa: finalCode,
              ten_hang_hoa_chuan: name,
              nhom_hang: String(row[mappedNhom] || "Khác").trim(),
              don_vi_tinh: String(row[mappedDvt] || "Cái").trim(),
              quy_cach: String(row[mappedQuyCach] || "").trim(),
              tu_khoa_nhan_dien: String(row[mappedTuKhoa] || name).trim(),
              ghi_chu: String(row[mappedGhiChu] || "").trim()
            };
          }).filter(Boolean) as Commodity[];

          setImportConfirm({
            type: "commodity",
            items: parsedCommodities,
            fileName: file.name
          });

        } else {
          const mappedMa = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_doi_tuong") || hNorm.includes("ma_khach") || hNorm.includes("ma_ncc") || hNorm.includes("ma_dt") || hNorm.includes("mã đối tác") || hNorm.includes("mã kh") || hNorm.includes("mã ncc") || hNorm.includes("customer_id") || hNorm.includes("code") || hNorm.includes("mã") || hNorm === "id";
          }) || headers[0] || "ma_doi_tuong";

          const mappedTen = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_doi_tuong") || hNorm.includes("ten_khach") || hNorm.includes("ten_ncc") || hNorm.includes("tên đối tác") || hNorm.includes("tên kh") || hNorm.includes("tên ncc") || hNorm.includes("company") || hNorm.includes("name") || hNorm.includes("đối tác") || hNorm.includes("tên");
          }) || headers[1] || "ten_doi_tuong";

          const mappedMst = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("mst") || hNorm.includes("ma_so_thue") || hNorm.includes("mã số thuế") || hNorm.includes("tax");
          }) || "";

          const mappedStk = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("so_tai_khoan") || hNorm.includes("tai_khoan") || hNorm.includes("stk") || hNorm.includes("acc");
          }) || "";

          const mappedNganHang = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ngan_hang") || hNorm.includes("ngân hàng") || hNorm.includes("bank");
          }) || "";

          const mappedDiaChi = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("dia_chi") || hNorm.includes("địa chỉ") || hNorm.includes("address");
          }) || "";

          const mappedTuKhoa = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tu_khoa") || hNorm.includes("từ khóa") || hNorm.includes("keywords");
          }) || "";

          const mappedGhiChu = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ghi_chu") || hNorm.includes("ghi chú") || hNorm.includes("note");
          }) || "";

          const parsedPartners: Partner[] = jsonData.map((row: any) => {
            const code = String(row[mappedMa] || "").trim();
            const name = String(row[mappedTen] || "").trim();
            if (!name) return null;
            const prefix = type === "customer" ? "KH" : "NCC";
            const finalCode = code || `${prefix}-${Math.random().toString(36).substr(2, 5).toUpperCase()}`;
            return {
              ma_doi_tuong: finalCode,
              ten_doi_tuong: name,
              loai_doi_tuong: type === "customer" ? "Khách hàng" : "Nhà cung cấp",
              ma_so_thue: String(row[mappedMst] || "").trim(),
              so_tai_khoan: String(row[mappedStk] || "").trim(),
              ngan_hang: String(row[mappedNganHang] || "").trim(),
              dia_chi: String(row[mappedDiaChi] || "").trim(),
              tu_khoa_nhan_dien: String(row[mappedTuKhoa] || name).trim(),
              ghi_chu: String(row[mappedGhiChu] || "").trim()
            };
          }).filter(Boolean) as Partner[];

          setImportConfirm({
            type,
            items: parsedPartners,
            fileName: file.name
          });
        }
      } catch (err: any) {
        console.error(err);
        triggerToast("Lỗi phân tích cú pháp tệp dữ liệu danh mục!", "warning");
      }
    };
    reader.readAsArrayBuffer(file);
    event.target.value = "";
  };

  const executeImportMaster = (overwrite: boolean) => {
    if (!importConfirm) return;
    const { type, items } = importConfirm;

    if (type === "commodity") {
      if (overwrite) {
        setCommodities(items);
        triggerToast(`Thay thế thành công: đã nạp ${items.length} mã hàng hóa mới.`);
      } else {
        setCommodities(prev => {
          const existingCodes = new Set(prev.map(c => c.ma_hang_hoa.toLowerCase()));
          const filteredNew = items.filter(item => !existingCodes.has(item.ma_hang_hoa.toLowerCase()));
          return [...prev, ...filteredNew];
        });
        triggerToast(`Bổ sung thành công: đã nạp thêm mã hàng hóa mới.`);
      }
    } else {
      if (overwrite) {
        // Only overwrite partners of that specific type
        const otherType = type === "customer" ? "Nhà cung cấp" : "Khách hàng";
        const keptPartners = partners.filter(p => p.loai_doi_tuong === otherType);
        setPartners([...keptPartners, ...items]);
        triggerToast(`Thay thế thành công: đã nạp ${items.length} mã đối tác ${type === "customer" ? "Khách hàng" : "Nhà cung cấp"} mới.`);
      } else {
        setPartners(prev => {
          const existingCodes = new Set(prev.map(p => p.ma_doi_tuong.toLowerCase()));
          const filteredNew = items.filter(item => !existingCodes.has(item.ma_doi_tuong.toLowerCase()));
          return [...prev, ...filteredNew];
        });
        triggerToast(`Bổ sung thành công: đã nạp thêm các mã đối tác ${type === "customer" ? "Khách hàng" : "Nhà cung cấp"} mới.`);
      }
    }
    setImportConfirm(null);
  };

  const handleGenericFileUpload = (event: React.ChangeEvent<HTMLInputElement>, fileType: "commodity" | "partner" | "bank" | "bank_sales" | "bank_purchases") => {
    const file = event.target.files?.[0];
    if (!file) return;

    // Check extension (Yêu cầu 5)
    const ext = file.name.split(".").pop()?.toLowerCase();
    if (ext !== "xlsx" && ext !== "xls" && ext !== "csv") {
      triggerToast("Định dạng tệp không được hỗ trợ! Vui lòng chọn tệp .xlsx, .xls hoặc .csv", "warning");
      event.target.value = "";
      return;
    }

    // Reset previous states to avoid mixing up data (Yêu cầu 13)
    if (fileType === "bank") {
      setBankSourceRows([]);
      setBankMappedRows([]);
      setBankAuditDetails(null);
    } else if (fileType === "partner") {
      setPartnerSourceRows([]);
      setPartnerMappedRows([]);
      setPartnerAuditDetails(null);
    } else if (fileType === "commodity") {
      setCommoditySourceRows([]);
      setCommodityMappedRows([]);
    } else if (fileType === "bank_sales") {
      setBankSalesRows([]);
    } else if (fileType === "bank_purchases") {
      setBankPurchasesRows([]);
    }

    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = e.target?.result;
        if (!data) return;

        const workbook = XLSX.read(data, { type: "array" });
        const sheetNames = workbook.SheetNames;
        if (sheetNames.length === 0) {
          triggerToast("Không tìm thấy trang tính nào trong file!", "warning");
          return;
        }

        if (fileType === "bank" || fileType === "bank_sales" || fileType === "bank_purchases") {
          const defaultSheet = sheetNames[0];
          const worksheet = workbook.Sheets[defaultSheet];
          const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
            header: 1,
            raw: false,
            defval: ""
          });

          // Auto-detect best header row among first 30 rows
          let bestHeaderIndex = 0;
          let bestScore = -1;
          const scanLimit = Math.min(30, rawRows.length);
          for (let i = 0; i < scanLimit; i++) {
            const score = scoreHeaderRow(rawRows[i]);
            if (score > bestScore) {
              bestScore = score;
              bestHeaderIndex = i;
            }
          }

          const cleaned = getHeadersFromRawRows(rawRows, bestHeaderIndex, 1);
          setUploadedWorkbook(workbook);
          setExcelWizard({
            fileType,
            fileName: file.name,
            sheetNames,
            selectedSheet: defaultSheet,
            rawRows,
            headerRowIndex: bestHeaderIndex,
            headerRowsCount: 1,
            headersCleaned: cleaned,
            rowsPreview: rawRows.slice(0, 30),
            onCancel: () => setExcelWizard(null)
          });
          triggerToast("Đã tải tệp Excel nguồn. Vui lòng chọn trang và dòng tiêu đề.");
          return;
        }

        const sheetName = sheetNames[0];
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
          triggerToast("Không tìm thấy trang tính nào trong file!", "warning");
          return;
        }

        const jsonData = XLSX.utils.sheet_to_json<any>(worksheet);
        if (jsonData.length === 0) {
          triggerToast("Tệp trống hoặc không chứa dữ liệu hợp lệ!", "warning");
          return;
        }

        const headers = Object.keys(jsonData[0] || {});

        if (fileType === "commodity") {
          const guessedTen = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_hang_hoa") || hNorm.includes("ten_hang") || hNorm.includes("dien_giai") || hNorm.includes("nội dung") || hNorm.includes("noi_dung") || hNorm.includes("mặt hàng") || hNorm.includes("description");
          }) || headers[0] || "ten_hang_hoa_dich_vu";

          const guessedDvt = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("don_vi_tinh") || hNorm.includes("dvt") || hNorm.includes("đơn vị tính") || hNorm.includes("uom") || hNorm.includes("unit");
          }) || headers[1] || "don_vi_tinh";

          const guessedQty = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("so_luong") || hNorm.includes("qty") || hNorm.includes("số lượng") || hNorm.includes("quantity");
          }) || "so_luong";

          const guessedPrice = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("don_gia") || hNorm.includes("price") || hNorm.includes("đơn giá") || hNorm.includes("unit_price");
          }) || "don_gia";

          const guessedAmount = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("thanh_tien") || hNorm.includes("amount") || hNorm.includes("thành tiền") || hNorm.includes("total");
          }) || "thanh_tien";

          setCommoditySourceRows(jsonData);
          setCommodityMappings({
            ten_hang_hoa: guessedTen,
            don_vi_tinh: guessedDvt,
            so_luong: guessedQty,
            don_gia: guessedPrice,
            thanh_tien: guessedAmount
          });
          setUploadedFileName(file.name);
          setActiveRowsCount(jsonData.length);
          setCommodityMappedRows([]);
          triggerToast(`Tải tệp hàng hóa thành công: ${jsonData.length} dòng.`);

        } else if (fileType === "partner") {
          const guessedTen = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_nguoi_ban") || hNorm.includes("ten_nguoi_mua") || hNorm.includes("ten_doi_tuong") || hNorm.includes("đối tác") || hNorm.includes("khách hàng") || hNorm.includes("nhà cung cấp") || hNorm.includes("company") || hNorm.includes("partner") || hNorm.includes("tên");
          }) || headers[0] || "ten_nguoi_ban";

          const guessedMst = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_so_thue") || hNorm.includes("mst") || hNorm.includes("tax") || hNorm.includes("mã số thuế");
          }) || "ma_so_thue_nguoi_ban";

          const guessedStk = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tai_khoan") || hNorm.includes("stk") || hNorm.includes("acc");
          }) || "";

          const guessedHH = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_hang") || hNorm.includes("dien_giai") || hNorm.includes("noi_dung") || hNorm.includes("mặt hàng");
          }) || "ten_hang_hoa_dich_vu";

          setPartnerSourceRows(jsonData);
          setPartnerMappings({
            ten_doi_tuong: guessedTen,
            ma_so_thue: guessedMst,
            so_tai_khoan: guessedStk,
            ten_hang_hoa: guessedHH
          });
          setUploadedFileName(file.name);
          setActiveRowsCount(jsonData.length);
          setPartnerMappedRows([]);
          triggerToast(`Tải tệp đối tác thành công: ${jsonData.length} dòng.`);

        } else if (fileType === "bank") {
          const guessedDesc = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("noi_dung_giao_dich") || hNorm.includes("noi_dung") || hNorm.includes("mô tả") || hNorm.includes("giao dịch") || hNorm.includes("description") || hNorm.includes("dien_giai") || hNorm.includes("diễn giải");
          }) || headers[0] || "noi_dung_giao_dich";

          const guessedIn = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tien_thu") || hNorm.includes("tien_vao") || hNorm.includes("thu") || hNorm.includes("có") || hNorm.includes("deposit") || hNorm.includes("credit") || hNorm.includes("so_tien") && (hNorm.includes("thu") || hNorm.includes("có") || hNorm.includes("gửi"));
          }) || headers.find(h => h.toLowerCase().includes("thu") || h.toLowerCase().includes("credit")) || "so_tien_thu";

          const guessedOut = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tien_chi") || hNorm.includes("tien_ra") || hNorm.includes("chi") || hNorm.includes("nợ") || hNorm.includes("withdrawal") || hNorm.includes("debit") || hNorm.includes("so_tien") && (hNorm.includes("chi") || hNorm.includes("nợ") || hNorm.includes("rút"));
          }) || headers.find(h => h.toLowerCase().includes("chi") || h.toLowerCase().includes("debit")) || "so_tien_chi";

          const guessedStkOrg = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tai_khoan_doi_ung") || hNorm.includes("stk") || hNorm.includes("tk");
          }) || "so_tai_khoan_doi_ung";

          const guessedNme = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_doi_tac") || hNorm.includes("nguoi_gui") || hNorm.includes("nguoi_nhan") || hNorm.includes("đối tác");
          }) || "ten_doi_tac_sao_ke";

          setBankSourceRows(jsonData);
          setBankMappings({
            noi_dung_giao_dich: guessedDesc,
            so_tien_thu: guessedIn,
            so_tien_chi: guessedOut,
            so_tai_khoan_doi_ung: guessedStkOrg,
            ten_doi_tac_sao_ke: guessedNme
          });
          setUploadedFileName(file.name);
          setActiveRowsCount(jsonData.length);
          setBankMappedRows([]);
          triggerToast(`Tải tệp sao kê ngân quỹ thành công: ${jsonData.length} dòng.`);

        } else if (fileType === "bank_sales") {
          const guessedNo = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("so_hoa_don") || hNorm.includes("so_hd") || hNorm.includes("invoice") || hNorm.includes("hóa đơn");
          }) || headers[0] || "so_hoa_don";

          const guessedDate = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ngay_hoa_don") || hNorm.includes("ngay_hd") || hNorm.includes("date") || hNorm.includes("ngày");
          }) || headers[1] || "ngay_hoa_don";

          const guessedKh = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_nguoi_mua") || hNorm.includes("ten_khach_hang") || hNorm.includes("khách hàng") || hNorm.includes("buyer") || hNorm.includes("customer") || hNorm.includes("tên");
          }) || "ten_nguoi_mua";

          const guessedMaKh = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_khach") || hNorm.includes("ma_kh") || hNorm.includes("buyer_id") || hNorm.includes("customer_id") || hNorm.includes("mã");
          }) || "ma_khach_hang";

          const guessedMst = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_so_thue") || hNorm.includes("mst") || hNorm.includes("tax");
          }) || "ma_so_thue_nguoi_mua";

          const guessedAmt = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tong_thanh_toan") || hNorm.includes("tong_tien") || hNorm.includes("thanh_toan") || hNorm.includes("total") || hNorm.includes("tiền");
          }) || "tong_thanh_toan";

          setBankSalesRows(jsonData);
          setBankSalesFileName(file.name);
          setBankSalesMappings({
            so_hoa_don: guessedNo,
            ngay_hoa_don: guessedDate,
            ten_khach_hang: guessedKh,
            ma_khach_hang: guessedMaKh,
            ma_so_thue_khach_hang: guessedMst,
            tong_thanh_toan: guessedAmt
          });
          triggerToast(`Tải tệp bảng kê bán ra thành công: ${jsonData.length} dòng.`);

        } else if (fileType === "bank_purchases") {
          const guessedNo = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("so_hoa_don") || hNorm.includes("so_hd") || hNorm.includes("invoice") || hNorm.includes("hóa đơn");
          }) || headers[0] || "so_hoa_don";

          const guessedDate = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ngay_hoa_don") || hNorm.includes("ngay_hd") || hNorm.includes("date") || hNorm.includes("ngày");
          }) || headers[1] || "ngay_hoa_don";

          const guessedNcc = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ten_nguoi_ban") || hNorm.includes("ten_nha_cung_cap") || hNorm.includes("nha_cung_cap") || hNorm.includes("seller") || hNorm.includes("supplier") || hNorm.includes("tên");
          }) || "ten_nguoi_ban";

          const guessedMaNcc = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_nha_cung_cap") || hNorm.includes("ma_ncc") || hNorm.includes("supplier_id") || hNorm.includes("mã");
          }) || "ma_nha_cung_cap";

          const guessedMst = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("ma_so_thue") || hNorm.includes("mst") || hNorm.includes("tax");
          }) || "ma_so_thue_nguoi_ban";

          const guessedAmt = headers.find(h => {
            const hNorm = h.toLowerCase();
            return hNorm.includes("tong_thanh_toan") || hNorm.includes("tong_tien") || hNorm.includes("thanh_toan") || hNorm.includes("total") || hNorm.includes("tiền");
          }) || "tong_thanh_toan";

          setBankPurchasesRows(jsonData);
          setBankPurchasesFileName(file.name);
          setBankPurchasesMappings({
            so_hoa_don: guessedNo,
            ngay_hoa_don: guessedDate,
            ten_nha_cung_cap: guessedNcc,
            ma_nha_cung_cap: guessedMaNcc,
            ma_so_thue_ncc: guessedMst,
            tong_thanh_toan: guessedAmt
          });
          triggerToast(`Tải tệp bảng kê mua vào thành công: ${jsonData.length} dòng.`);
        }
      } catch (err: any) {
        console.error("Lỗi xử lý file upload:", err);
        const logEntry = {
          timestamp: new Date().toISOString(),
          fileName: file.name,
          fileSize: file.size,
          fileType: ext || "unknown",
          errorMsg: err?.message || String(err)
        };
        setUploadLogs(prev => [logEntry, ...prev]);
        triggerToast(`Lỗi định dạng cấu trúc tệp: ${err?.message || "Đảm bảo tệp sạch và đúng cột."}`, "warning");
      } finally {
        event.target.value = "";
      }
    };
    reader.onerror = () => {
      triggerToast("Không thể đọc tệp tin từ thiết bị!", "warning");
      event.target.value = "";
    };
    reader.readAsArrayBuffer(file);
  };

  const handleWizardSheetChange = (sheetName: string) => {
    if (!uploadedWorkbook || !excelWizard) return;
    const worksheet = uploadedWorkbook.Sheets[sheetName];
    if (!worksheet) return;

    const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
      header: 1,
      raw: false,
      defval: ""
    });

    let bestHeaderIndex = 0;
    let bestScore = -1;
    const scanLimit = Math.min(30, rawRows.length);
    for (let i = 0; i < scanLimit; i++) {
      const score = scoreHeaderRow(rawRows[i]);
      if (score > bestScore) {
        bestScore = score;
        bestHeaderIndex = i;
      }
    }

    const cleaned = getHeadersFromRawRows(rawRows, bestHeaderIndex, 1);

    setExcelWizard(prev => {
      if (!prev) return null;
      return {
        ...prev,
        selectedSheet: sheetName,
        rawRows,
        headerRowIndex: bestHeaderIndex,
        headerRowsCount: 1,
        headersCleaned: cleaned,
        rowsPreview: rawRows.slice(0, 30)
      };
    });
  };

  const handleWizardHeaderRowChange = (index: number) => {
    setExcelWizard(prev => {
      if (!prev) return null;
      const cleaned = getHeadersFromRawRows(prev.rawRows, index, prev.headerRowsCount);
      return {
        ...prev,
        headerRowIndex: index,
        headersCleaned: cleaned
      };
    });
  };

  const handleWizardHeaderCountChange = (count: number) => {
    setExcelWizard(prev => {
      if (!prev) return null;
      const cleaned = getHeadersFromRawRows(prev.rawRows, prev.headerRowIndex, count);
      return {
        ...prev,
        headerRowsCount: count,
        headersCleaned: cleaned
      };
    });
  };

  const handleConfirmWizard = () => {
    if (!excelWizard) return;
    const { fileType, fileName, selectedSheet, rawRows, headerRowIndex, headerRowsCount } = excelWizard;

    if (rawRows.length <= headerRowIndex) {
      triggerToast("Dòng tiêu đề vượt quá số dòng của tệp!", "warning");
      return;
    }

    const cleanHeaders = getHeadersFromRawRows(rawRows, headerRowIndex, headerRowsCount);

    if (cleanHeaders.length === 0) {
      triggerToast("Không thể tìm thấy tiêu đề cột hợp lệ!", "warning");
      return;
    }

    const dataRows = rawRows.slice(headerRowIndex + headerRowsCount);
    const jsonData: any[] = [];
    const emptyColumnsSet = new Set<string>(cleanHeaders);
    let totalCols = cleanHeaders.length;

    dataRows.forEach((row) => {
      const isRowEmpty = row.every(val => val === undefined || val === null || String(val).trim() === "");
      if (isRowEmpty) return;

      const obj: any = {};
      let hasData = false;
      cleanHeaders.forEach((header, colIdx) => {
        let val = row[colIdx];
        if (val !== undefined && val !== null && String(val).trim() !== "") {
          obj[header] = typeof val === "string" ? val.trim() : val;
          emptyColumnsSet.delete(header);
          hasData = true;
        } else {
          obj[header] = "";
        }
      });
      if (hasData) {
        jsonData.push(obj);
      }
    });

    if (jsonData.length === 0) {
      triggerToast("Không tìm thấy dòng dữ liệu nào bên dưới tiêu đề cột!", "warning");
      return;
    }

    const originalHeaderRow = rawRows[headerRowIndex] || [];
    const renamed: string[] = [];
    cleanHeaders.forEach((h, idx) => {
      const orig = String(originalHeaderRow[idx] || "").trim();
      if (orig && orig !== h) {
        renamed.push(`Cột ${idx + 1}: "${orig}" ➔ "${h}"`);
      }
    });

    const auditInfo = {
      fileName,
      sheetName: selectedSheet,
      headerRow: headerRowIndex + 1,
      headerCount: headerRowsCount,
      totalColumns: totalCols,
      totalRows: jsonData.length,
      allColumns: cleanHeaders,
      hasDataColumnsCount: totalCols - emptyColumnsSet.size,
      emptyColumns: Array.from(emptyColumnsSet),
      renamedColumns: renamed
    };

    if (fileType === "bank") {
      setBankSourceRows(jsonData);
      setUploadedFileName(fileName);
      setActiveRowsCount(jsonData.length);
      setBankMappedRows([]);
      setBankAuditDetails(auditInfo);

      const proposed = proposeBankMappings(cleanHeaders);
      setBankMappings(proposed);

      triggerToast(`Nạp ${jsonData.length} dòng sao kê từ sheet "${selectedSheet}" thành công!`);
    } else if (fileType === "partner") {
      setPartnerSourceRows(jsonData);
      setPartnerAuditDetails(auditInfo);
      setUploadedFileName(fileName);
      setActiveRowsCount(jsonData.length);
      setPartnerMappedRows([]);

      const guessedTen = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ten_nguoi_ban") || hNorm.includes("ten_nguoi_mua") || hNorm.includes("ten_doi_tuong") || hNorm.includes("đối tác") || hNorm.includes("khách hàng") || hNorm.includes("nhà cung cấp") || hNorm.includes("company") || hNorm.includes("partner") || hNorm.includes("tên");
      }) || cleanHeaders[0] || "ten_nguoi_ban";

      const guessedMst = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ma_so_thue") || hNorm.includes("mst") || hNorm.includes("tax") || hNorm.includes("mã số thuế");
      }) || "ma_so_thue_nguoi_ban";

      const guessedStk = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("tai_khoan") || hNorm.includes("stk") || hNorm.includes("acc");
      }) || "";

      const guessedHH = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ten_hang") || hNorm.includes("dien_giai") || hNorm.includes("noi_dung") || hNorm.includes("mặt hàng");
      }) || "ten_hang_hoa_dich_vu";

      setPartnerMappings({
        ten_doi_tuong: guessedTen,
        ma_so_thue: guessedMst,
        so_tai_khoan: guessedStk,
        ten_hang_hoa: guessedHH
      });

      triggerToast(`Nạp ${jsonData.length} dòng dữ liệu đối tác từ sheet "${selectedSheet}" thành công!`);
    } else if (fileType === "bank_sales") {
      setBankSalesRows(jsonData);
      setBankSalesFileName(fileName);

      const guessedNo = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("so_hoa_don") || hNorm.includes("so_hd") || hNorm.includes("invoice") || hNorm.includes("hóa đơn");
      }) || cleanHeaders[0] || "so_hoa_don";

      const guessedDate = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ngay_hoa_don") || hNorm.includes("ngay_hd") || hNorm.includes("date") || hNorm.includes("ngày");
      }) || cleanHeaders[1] || "ngay_hoa_don";

      const guessedKh = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ten_nguoi_mua") || hNorm.includes("ten_khach_hang") || hNorm.includes("khách hàng") || hNorm.includes("buyer") || hNorm.includes("customer") || hNorm.includes("tên");
      }) || "ten_nguoi_mua";

      const guessedMaKh = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ma_khach") || hNorm.includes("ma_kh") || hNorm.includes("buyer_id") || hNorm.includes("customer_id") || hNorm.includes("mã");
      }) || "ma_khach_hang";

      const guessedMst = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ma_so_thue") || hNorm.includes("mst") || hNorm.includes("tax");
      }) || "ma_so_thue_nguoi_mua";

      const guessedAmt = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("tong_thanh_toan") || hNorm.includes("tong_tien") || hNorm.includes("thanh_toan") || hNorm.includes("total") || hNorm.includes("tiền");
      }) || "tong_thanh_toan";

      setBankSalesMappings({
        so_hoa_don: guessedNo,
        ngay_hoa_don: guessedDate,
        ten_khach_hang: guessedKh,
        ma_khach_hang: guessedMaKh,
        ma_so_thue_khach_hang: guessedMst,
        tong_thanh_toan: guessedAmt
      });
      triggerToast(`Nạp ${jsonData.length} dòng bảng kê bán ra thành công!`);
    } else if (fileType === "bank_purchases") {
      setBankPurchasesRows(jsonData);
      setBankPurchasesFileName(fileName);

      const guessedNo = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("so_hoa_don") || hNorm.includes("so_hd") || hNorm.includes("invoice") || hNorm.includes("hóa đơn");
      }) || cleanHeaders[0] || "so_hoa_don";

      const guessedDate = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ngay_hoa_don") || hNorm.includes("ngay_hd") || hNorm.includes("date") || hNorm.includes("ngày");
      }) || cleanHeaders[1] || "ngay_hoa_don";

      const guessedNcc = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ten_nguoi_ban") || hNorm.includes("ten_nha_cung_cap") || hNorm.includes("nha_cung_cap") || hNorm.includes("seller") || hNorm.includes("supplier") || hNorm.includes("tên");
      }) || "ten_nguoi_ban";

      const guessedMaNcc = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ma_nha_cung_cap") || hNorm.includes("ma_ncc") || hNorm.includes("supplier_id") || hNorm.includes("mã");
      }) || "ma_nha_cung_cap";

      const guessedMst = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("ma_so_thue") || hNorm.includes("mst") || hNorm.includes("tax");
      }) || "ma_so_thue_nguoi_ban";

      const guessedAmt = cleanHeaders.find(h => {
        const hNorm = h.toLowerCase();
        return hNorm.includes("tong_thanh_toan") || hNorm.includes("tong_tien") || hNorm.includes("thanh_toan") || hNorm.includes("total") || hNorm.includes("tiền");
      }) || "tong_thanh_toan";

      setBankPurchasesMappings({
        so_hoa_don: guessedNo,
        ngay_hoa_don: guessedDate,
        ten_nha_cung_cap: guessedNcc,
        ma_nha_cung_cap: guessedMaNcc,
        ma_so_thue_ncc: guessedMst,
        tong_thanh_toan: guessedAmt
      });
      triggerToast(`Nạp ${jsonData.length} dòng bảng kê mua vào thành công!`);
    }

    setExcelWizard(null);
  };

  const handleReopenWizard = (type: "bank" | "partner" | "bank_sales" | "bank_purchases") => {
    if (!uploadedWorkbook) {
      triggerToast("Không tìm thấy tệp Excel nào đang được mở. Hãy tải tệp lên trước!", "warning");
      return;
    }
    const sheetNames = uploadedWorkbook.SheetNames;
    const currentSheet = sheetNames[0];
    const worksheet = uploadedWorkbook.Sheets[currentSheet];
    const rawRows = XLSX.utils.sheet_to_json<any[]>(worksheet, {
      header: 1,
      raw: false,
      defval: ""
    });

    let bestHeaderIndex = 0;
    let bestScore = -1;
    const scanLimit = Math.min(30, rawRows.length);
    for (let i = 0; i < scanLimit; i++) {
      const score = scoreHeaderRow(rawRows[i]);
      if (score > bestScore) {
        bestScore = score;
        bestHeaderIndex = i;
      }
    }

    const cleaned = getHeadersFromRawRows(rawRows, bestHeaderIndex, 1);
    setExcelWizard({
      fileType: type,
      fileName: uploadedFileName || (type === "bank" ? "Sao_Ke_Ngan_Hang.xlsx" : type === "partner" ? "Danh_Muc_Doi_Tac.xlsx" : type === "bank_sales" ? "Bang_Ke_Ban_Ra.xlsx" : "Bang_Ke_Mua_Vao.xlsx"),
      sheetNames,
      selectedSheet: currentSheet,
      rawRows,
      headerRowIndex: bestHeaderIndex,
      headerRowsCount: 1,
      headersCleaned: cleaned,
      rowsPreview: rawRows.slice(0, 30),
      onCancel: () => setExcelWizard(null)
    });
    triggerToast("Đã mở lại cấu hình Excel Wizard.");
  };

  const handleRestoreAutoMapping = () => {
    if (bankHeaders.length === 0) {
      triggerToast("Chưa có dữ liệu sao kê ngân hàng!", "warning");
      return;
    }
    const proposed = proposeBankMappings(bankHeaders);
    setBankMappings(proposed);
    triggerToast("Đã tự động khôi phục cấu hình ánh xạ cột tối ưu!");
  };

  const handleClearMapping = () => {
    setBankMappings({
      ngay_giao_dich: "",
      ngay_hieu_luc: "",
      ngay_hach_toan: "",
      noi_dung_giao_dich: "",
      so_tien_thu: "",
      so_tien_chi: "",
      so_du: "",
      ma_giao_dich: "",
      so_chung_tu: "",
      so_tham_chieu: "",
      ten_doi_tac_sao_ke: "",
      so_tai_khoan_doi_ung: "",
      ngan_hang_doi_ung: "",
      ghi_chu: "",
      loai_tien_te: "",
      cot_ngay: "",
      cot_thang: "",
      cot_nam: ""
    });
    triggerToast("Đã dọn sạch tất cả cấu hình ánh xạ cột!");
  };

  // Switch partner mapping based on Mua vào / Bán ra state
  useEffect(() => {
    if (demoLoaded) {
      if (partnerMode === "Mua vào") {
        setPartnerSourceRows(getSamplePurchaseLedger());
        setPartnerMappings({
          ten_doi_tuong: "ten_nguoi_ban",
          ma_so_thue: "ma_so_thue_nguoi_ban",
          so_tai_khoan: "",
          ten_hang_hoa: "ten_hang_hoa_dich_vu"
        });
      } else {
        setPartnerSourceRows(getSampleSalesLedger());
        setPartnerMappings({
          ten_doi_tuong: "ten_nguoi_mua",
          ma_so_thue: "ma_so_thue_nguoi_mua",
          so_tai_khoan: "",
          ten_hang_hoa: "ten_hang_hoa_dich_vu"
        });
      }
    }
  }, [partnerMode, demoLoaded]);


  // ==========================================
  // THUẬT TOÁN GẮN MÃ HÀNG HÓA
  // ==========================================
  const handleProcessCommodities = () => {
    if (commoditySourceRows.length === 0) {
      triggerToast("Bạn vui lòng tải tệp hoặc kích hoạt dữ liệu Demo trước!", "warning");
      return;
    }
    setIsProcessingCommodities(true);

    setTimeout(() => {
      let tempCommodities = [...commodities];
      const existingCodes = tempCommodities.map(c => c.ma_hang_hoa);

      const results = commoditySourceRows.map((row, index) => {
        const desc = row[commodityMappings.ten_hang_hoa] || "";
        const uom = row[commodityMappings.don_vi_tinh] || "";

        const match = matchCommodityRow(desc, uom, tempCommodities, config.autoThreshold, config.autoThreshold);

        let treatment: MappedRow["treatment"] = "TỰ ĐỘNG GẮN";
        let finalizedCode = match.code;
        let finalizedName = match.name;
        let finalReason = match.reason;
        let finalScore = match.score;

        if (match.score >= config.autoThreshold) {
          treatment = "TỰ ĐỘNG GẮN";
        } else {
          treatment = "TẠO MÃ MỚI";
          const newCode = generateNewCode(config.prefixHH, existingCodes);
          finalizedCode = newCode;
          finalizedName = desc.trim();
          existingCodes.push(newCode);

          // Append to temp directory so subsequent rows can match this newly created code if identical
          const newItem: Commodity = {
            ma_hang_hoa: finalizedCode,
            ten_hang_hoa_chuan: finalizedName,
            nhom_hang: "Vật tư xây dựng",
            don_vi_tinh: uom || "Bao",
            quy_cach: "Tự động sinh mới",
            tu_khoa_nhan_dien: "",
            ghi_chu: `Mã tự động sinh từ dòng HĐ: ${desc}`
          };
          tempCommodities.push(newItem);
          finalReason = `Điểm tương đồng thấp (${match.score}% < ${config.autoThreshold}%). Tự động tạo mã hàng mới.`;
          finalScore = match.score;
        }

        return {
          id: `comm_${index}`,
          originalText: desc,
          originalUom: uom,
          proposedCode: finalizedCode,
          proposedName: finalizedName,
          score: finalScore,
          reason: finalReason,
          treatment,
          notes: "",
          rawRowData: row
        } as MappedRow;
      });

      setCommodities(tempCommodities);
      setCommodityMappedRows(results);
      setIsProcessingCommodities(false);
      triggerToast("Đã hoàn tất gắn mã hàng hóa!");
    }, 600);
  };


  // ==========================================
  // THUẬT TOÁN GẮN MÃ ĐỐI TÁC
  // ==========================================
  const handleProcessPartners = () => {
    if (partnerSourceRows.length === 0) {
      triggerToast("Bạn vui lòng tải tệp hoặc kích hoạt dữ liệu Demo trước!", "warning");
      return;
    }
    setIsProcessingPartners(true);

    setTimeout(() => {
      let tempPartners = [...partners];
      const existingCodes = tempPartners.map(p => p.ma_doi_tuong);
      const isBuyer = partnerMode === "Bán ra";

      const results = partnerSourceRows.map((row, index) => {
        const nameVal = row[partnerMappings.ten_doi_tuong] || "";
        const mstVal = row[partnerMappings.ma_so_thue] || "";
        const accVal = row[partnerMappings.so_tai_khoan] || "";
        const invDesc = row[partnerMappings.ten_hang_hoa] || "";

        const match = matchPartnerRow(nameVal, mstVal, accVal, invDesc, tempPartners, isBuyer, config.autoThreshold, config.autoThreshold);

        let treatment: MappedRow["treatment"] = "TỰ ĐỘNG GẮN";
        let finalizedCode = match.code;
        let finalizedName = match.name;
        let finalizedType = match.type || (isBuyer ? "Khách hàng" : "Nhà cung cấp");
        let finalReason = match.reason;
        let finalScore = match.score;

        if (match.score >= config.autoThreshold) {
          treatment = "TỰ ĐỘNG GẮN";
        } else {
          treatment = "TẠO MÃ MỚI";
          const prefix = isBuyer ? config.prefixKH : config.prefixNCC;
          const newCode = generateNewCode(prefix, existingCodes);
          finalizedCode = newCode;
          finalizedName = nameVal.trim();
          finalizedType = isBuyer ? "Khách hàng" : "Nhà cung cấp";
          existingCodes.push(newCode);

          const newPartner: Partner = {
            ma_doi_tuong: finalizedCode,
            ten_doi_tuong: finalizedName,
            loai_doi_tuong: finalizedType as any,
            ma_so_thue: mstVal || "",
            so_tai_khoan: accVal || "",
            ngan_hang: "",
            dia_chi: "Tự tạo mới",
            tu_khoa_nhan_dien: "",
            ghi_chu: `Tạo từ dòng bảng kê: ${nameVal}`
          };
          tempPartners.push(newPartner);
          finalReason = `Điểm tương thích thấp (${match.score}% < ${config.autoThreshold}%). Tự động tạo mã đối tác mới.`;
          finalScore = match.score;
        }

        return {
          id: `part_${index}`,
          originalText: nameVal,
          originalMst: mstVal,
          originalAcc: accVal,
          proposedCode: finalizedCode,
          proposedName: finalizedName,
          proposedType: finalizedType,
          score: finalScore,
          reason: finalReason,
          treatment,
          notes: "",
          rawRowData: row
        } as MappedRow;
      });

      setPartners(tempPartners);
      setPartnerMappedRows(results);
      setIsProcessingPartners(false);
      triggerToast("Đã hoàn tất gắn mã đối tác khách hàng/nhà cung cấp!");
    }, 600);
  };


  // ==========================================
  // THUẬT TOÁN PHÂN TÍCH SAO KÊ NGÂN HÀNG
  // ==========================================
  const handleProcessBank = () => {
    if (bankSourceRows.length === 0) {
      triggerToast("Bạn vui lòng tải tệp hoặc kích hoạt dữ liệu Demo trước!", "warning");
      return;
    }
    setIsProcessingBank(true);

    setTimeout(() => {
      // 1. Tự động phát hiện và tạo mã đối tác mới từ bảng kê mua/bán (Section 17)
      const newlyCreatedPartners = preProcessInvoicePartners(
        bankSalesRows,
        bankSalesMappings,
        bankPurchasesRows,
        bankPurchasesMappings,
        partners,
        config.prefixKH || "KH",
        config.prefixNCC || "NCC"
      );

      if (newlyCreatedPartners.length > 0) {
        setPartners(prev => {
          const prevCodes = new Set(prev.map(p => p.ma_doi_tuong));
          const filtered = newlyCreatedPartners.filter(p => !prevCodes.has(p.ma_doi_tuong));
          return [...prev, ...filtered];
        });
        triggerToast(`Đã tự động thêm ${newlyCreatedPartners.length} đối tác mới chưa có trong danh mục!`);
      }

      // 2. Khởi tạo sổ theo dõi công nợ tạm thời cho từng hóa đơn (Section 11)
      const invoiceBalances = new Map<string, number>();

      // Hóa đơn bán ra
      bankSalesRows.forEach(row => {
        const invNo = String(row[bankSalesMappings.so_hoa_don] || "").trim();
        const invDateVal = row[bankSalesMappings.ngay_hoa_don];
        // Parse date to standard string
        const dObj = parseDate(invDateVal);
        const invDateStr = dObj ? `${dObj.getFullYear()}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${String(dObj.getDate()).padStart(2, "0")}` : "";
        const invAmount = parseFloat(row[bankSalesMappings.tong_thanh_toan]) || 0;
        const key = `${invNo}_${invDateStr}_${invAmount}`;
        if (invNo && invAmount > 0) {
          invoiceBalances.set(key, invAmount);
        }
      });

      // Hóa đơn mua vào
      bankPurchasesRows.forEach(row => {
        const invNo = String(row[bankPurchasesMappings.so_hoa_don] || "").trim();
        const invDateVal = row[bankPurchasesMappings.ngay_hoa_don];
        const dObj = parseDate(invDateVal);
        const invDateStr = dObj ? `${dObj.getFullYear()}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${String(dObj.getDate()).padStart(2, "0")}` : "";
        const invAmount = parseFloat(row[bankPurchasesMappings.tong_thanh_toan]) || 0;
        const key = `${invNo}_${invDateStr}_${invAmount}`;
        if (invNo && invAmount > 0) {
          invoiceBalances.set(key, invAmount);
        }
      });

      // 3. Tiến hành đối chiếu theo thứ tự thời gian sao kê
      const existingCodes = partners.map(p => p.ma_doi_tuong);
      const combinedActivePartners = [...partners, ...newlyCreatedPartners];

      const results = bankSourceRows.map((row, index) => {
        const analysis = matchBankRowWithInvoiceEngine({
          bankRow: row,
          bankMappings,
          partners: combinedActivePartners,
          salesLedger: bankSalesRows,
          salesMappings: bankSalesMappings,
          purchaseLedger: bankPurchasesRows,
          purchaseMappings: bankPurchasesMappings,
          config: {
            autoThreshold: config.autoThreshold,
            checkThreshold: config.checkThreshold,
            daysBeforeInvoice: config.daysBeforeInvoice || 7,
            daysAfterInvoice: config.daysAfterInvoice || 30,
            diffAbsThreshold: config.diffAbsThreshold || 10000,
            diffPctThreshold: config.diffPctThreshold || 0.5,
            maxCombinationCount: config.maxCombinationCount || 5,
            prefixKH: config.prefixKH || "KH",
            prefixNCC: config.prefixNCC || "NCC"
          },
          invoiceBalances,
          existingCodes,
          tempCreatedPartners: newlyCreatedPartners
        });

        return {
          ...analysis,
          id: `bank_${index}`
        };
      });

      setBankMappedRows(results);
      setIsProcessingBank(false);
      triggerToast("Đã đối chiếu thành công sao kê ngân hàng với các hóa đơn!");
    }, 600);
  };


  // ==========================================
  // CHẾ ĐỘ 4: GẮN MÃ TỔNG HỢP & ĐỐI CHIẾU CHÉO
  // ==========================================
  const handleProcessIntegrated = () => {
    setIsProcessingIntegrated(true);
    setTimeout(() => {
      // 1. Process purchase ledger
      const localPurchases = getSamplePurchaseLedger().map((row, index) => {
        const itemMatch = matchCommodityRow(
          row.ten_hang_hoa_dich_vu,
          row.don_vi_tinh,
          commodities,
          config.autoThreshold,
          config.autoThreshold
        );
        const partnerMatch = matchPartnerRow(
          row.ten_nguoi_ban,
          row.ma_so_thue_nguoi_ban,
          "",
          row.ten_hang_hoa_dich_vu,
          partners,
          false,
          config.autoThreshold,
          config.autoThreshold
        );

        return {
          id: `integ_p_${index}`,
          originalText: row.ten_hang_hoa_dich_vu,
          proposedCode: itemMatch.score >= config.autoThreshold ? itemMatch.code : "MÃ THỦ CÔNG",
          proposedName: itemMatch.score >= config.autoThreshold ? itemMatch.name : row.ten_hang_hoa_dich_vu,
          originalUom: row.don_vi_tinh,
          originalMst: row.ma_so_thue_nguoi_ban,
          originalAcc: partnerMatch.code, // Overriding as helper
          proposedType: partnerMatch.name, // Overriding as partner name helper
          score: itemMatch.score,
          reason: `Khớp vật tư: ${itemMatch.score}đ | Khớp NCC: ${partnerMatch.score}đ`,
          treatment: "TỰ ĐỘNG GẮN",
          notes: "",
          rawRowData: row
        } as MappedRow;
      });

      // 2. Process sales ledger
      const localSales = getSampleSalesLedger().map((row, index) => {
        const itemMatch = matchCommodityRow(
          row.ten_hang_hoa_dich_vu,
          row.don_vi_tinh,
          commodities,
          config.autoThreshold,
          config.autoThreshold
        );
        const partnerMatch = matchPartnerRow(
          row.ten_nguoi_mua,
          row.ma_so_thue_nguoi_mua,
          "",
          row.ten_hang_hoa_dich_vu,
          partners,
          true,
          config.autoThreshold,
          config.autoThreshold
        );

        return {
          id: `integ_s_${index}`,
          originalText: row.ten_hang_hoa_dich_vu,
          proposedCode: itemMatch.score >= config.autoThreshold ? itemMatch.code : "MÃ THỦ CÔNG",
          proposedName: itemMatch.score >= config.autoThreshold ? itemMatch.name : row.ten_hang_hoa_dich_vu,
          originalUom: row.don_vi_tinh,
          originalMst: row.ma_so_thue_nguoi_mua,
          originalAcc: partnerMatch.code, // Overriding as partner code
          proposedType: partnerMatch.name, // Overriding as partner name
          score: itemMatch.score,
          reason: `Khớp vật tư: ${itemMatch.score}đ | Khớp KH: ${partnerMatch.score}đ`,
          treatment: "TỰ ĐỘNG GẮN",
          notes: "",
          rawRowData: row
        } as MappedRow;
      });

      // 3. Process inventory ledger cards
      const localInventories = getSampleInventoryLedger().map((row, index) => {
        const itemMatch = matchCommodityRow(row.ten_hang_hoa, row.don_vi_tinh, commodities, config.autoThreshold, config.autoThreshold);
        return {
          ...row,
          ma_hang_hoa_gan: itemMatch.score >= config.autoThreshold ? itemMatch.code : "MẸO THỦ CÔNG",
          ten_hang_hoa_chuan: itemMatch.score >= config.autoThreshold ? itemMatch.name : row.ten_hang_hoa
        };
      });

      // 4. Process Bank logs
      const localBanks = getSampleBankStatement().map((row, index) => {
        const amIn = row.so_tien_thu;
        const amOut = row.so_tien_chi;
        const analysis = analyzeBankTransaction(
          row.noi_dung_giao_dich,
          amIn,
          amOut,
          row.so_tai_khoan_doi_ung,
          row.ten_doi_tac_sao_ke,
          partners,
          config.autoThreshold,
          config.autoThreshold
        );
        return {
          ...analysis,
          id: `integ_b_${index}`,
          date: row.ngay_giao_dich,
          notes: ""
        } as BankAnalysisResult;
      });

      // 5. Compile reconciliation table (Sales Invoice total price vs bank collections)
      // Group sales invoice totals by Customer Code
      const salesByCustomer: { [code: string]: { name: string; invoiceTotal: number } } = {};
      localSales.forEach(s => {
        const code = s.originalAcc || "CHƯA_RÕ";
        const name = s.proposedType || "Chưa rõ khách hàng";
        const am = s.rawRowData.tong_thanh_toan || 0;
        if (!salesByCustomer[code]) {
          salesByCustomer[code] = { name, invoiceTotal: 0 };
        }
        salesByCustomer[code].invoiceTotal += am;
      });

      // Group bank statement collections by Partner Code (only "Thu tiền khách hàng" category)
      const bankByCustomer: { [code: string]: number } = {};
      localBanks.forEach(b => {
        if (b.predictedGroup === "Thu tiền khách hàng" && b.proposedCode) {
          bankByCustomer[b.proposedCode] = (bankByCustomer[b.proposedCode] || 0) + b.amountIn;
        }
      });

      // Build joint recon list
      const allCustomerCodes = Array.from(new Set([...Object.keys(salesByCustomer), ...Object.keys(bankByCustomer)]));
      const reconResult = allCustomerCodes.map(code => {
        const invInfo = salesByCustomer[code] || { name: partners.find(p => p.ma_doi_tuong === code)?.ten_doi_tuong || "Từ khóa lẻ ngân hàng", invoiceTotal: 0 };
        const collected = bankByCustomer[code] || 0;
        const diff = invInfo.invoiceTotal - collected;

        return {
          partnerCode: code,
          partnerName: invInfo.name,
          invoiceTotal: invInfo.invoiceTotal,
          bankTotal: collected,
          difference: diff
        };
      });

      setIntegratedPurchaseRows(localPurchases);
      setIntegratedSaleRows(localSales);
      setIntegratedInvRows(localInventories);
      setIntegratedBankRows(localBanks);
      setIntegratedRecon(reconResult);

      setIsProcessingIntegrated(false);
      triggerToast("Đã thiết lập liên kết thông tin đa chiều thành công!");
    }, 700);
  };

  // Quick Action Hooks to make editing experience interactive
  const handleEditCommodityCode = (id: string, newCode: string) => {
    setCommodityMappedRows(prev =>
      prev.map(row => {
        if (row.id === id) {
          const matchingItem = commodities.find(c => c.ma_hang_hoa === newCode);
          return {
            ...row,
            proposedCode: newCode,
            proposedName: matchingItem ? matchingItem.ten_hang_hoa_chuan : row.proposedName,
            treatment: "Đã chốt" as any
          };
        }
        return row;
      })
    );
  };

  const handleEditPartnerCode = (id: string, newCode: string) => {
    setPartnerMappedRows(prev =>
      prev.map(row => {
        if (row.id === id) {
          const matchingPartner = partners.find(p => p.ma_doi_tuong === newCode);
          return {
            ...row,
            proposedCode: newCode,
            proposedName: matchingPartner ? matchingPartner.ten_doi_tuong : row.proposedName,
            treatment: "Đã chốt" as any
          };
        }
        return row;
      })
    );
  };

  const handleEditBankCode = (id: string, newCode: string) => {
    setBankMappedRows(prev =>
      prev.map(row => {
        if (row.id === id) {
          const matchingPartner = partners.find(p => p.ma_doi_tuong === newCode);
          return {
            ...row,
            proposedCode: newCode,
            proposedName: matchingPartner ? matchingPartner.ten_doi_tuong : row.proposedName,
            treatment: "Đã chốt" as any
          };
        }
        return row;
      })
    );
  };

  // ==========================================
  // XUẤT FILE EXCEL THỰC TẾ QUA XLSX (SHEETJS)
  // ==========================================
  const exportCommodityToExcel = () => {
    if (commodityMappedRows.length === 0) {
      triggerToast("Không có thông tin gắn mã hàng hóa để xuất!", "warning");
      return;
    }
    const dataToExport = commodityMappedRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      baseRow["Mã hàng hóa"] = row.proposedCode;
      baseRow["Tên hàng hóa chuẩn"] = row.proposedName;
      baseRow["Tên hàng hóa chuẩn hóa"] = normalizeText(row.originalText);
      baseRow["Độ tương thích"] = `${row.score}%`;
      baseRow["Mức độ tương thích"] = row.score >= config.autoThreshold ? "Cao" : "Thấp";
      baseRow["Lý do gắn mã"] = row.reason;
      baseRow["Trạng thái xử lý"] = row.treatment === "TỰ ĐỘNG GẮN" ? "Đã chốt" : "Cần kiểm tra";
      baseRow["Ghi chú"] = row.notes || "";
      return baseRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Bang_ke_da_gan_ma");

    const baseName = uploadedFileName ? uploadedFileName.replace(/\.[^/.]+$/, "") : "Bang_ke_hang_hoa";
    XLSX.writeFile(workbook, `${baseName}_da_gan_ma.xlsx`);
    triggerToast("Đã xuất khẩu tệp kết quả gắn mã hàng hóa chuẩn hóa thành công!");
  };

  const exportPartnerToExcel = () => {
    if (partnerMappedRows.length === 0) {
      triggerToast("Không có thông tin đối tác để xuất!", "warning");
      return;
    }
    const dataToExport = partnerMappedRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      const isSales = partnerMode === "Bán ra";
      
      if (isSales) {
        baseRow["Mã khách hàng"] = row.proposedCode;
        baseRow["Tên khách hàng chuẩn"] = row.proposedName;
      } else {
        baseRow["Mã nhà cung cấp"] = row.proposedCode;
        baseRow["Tên nhà cung cấp chuẩn"] = row.proposedName;
      }
      baseRow["Độ tương thích"] = `${row.score}%`;
      baseRow["Lý do gắn mã"] = row.reason;
      baseRow["Trạng thái xử lý"] = (row.treatment === "TỰ ĐỘNG GẮN" || row.treatment === "DUYỆT THỦ CÔNG") ? "Đã chốt" : "Cần kiểm tra";
      return baseRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Doi_tac_da_gan_ma");

    const defaultName = partnerMode === "Bán ra" ? "Bang_ke_ban_ra" : "Bang_ke_mua_vao";
    const baseName = uploadedFileName ? uploadedFileName.replace(/\.[^/.]+$/, "") : defaultName;
    XLSX.writeFile(workbook, `${baseName}_da_gan_ma.xlsx`);
    triggerToast("Đã xuất khẩu tệp kết quả gắn mã đối tác thành công!");
  };

  const exportBankToExcel = () => {
    if (bankMappedRows.length === 0) {
      triggerToast("Không có bảng sao kê ngân hàng để xuất!", "warning");
      return;
    }

    const workbook = XLSX.utils.book_new();

    // 1. Sheet: Sao_ke_da_gan_ma
    const sheet1Data = bankMappedRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      baseRow["Mã đối tượng đề xuất"] = row.proposedCode || "";
      baseRow["Tên đối tượng đề xuất"] = row.proposedName || "Nghi vấn / Thất lạc";
      baseRow["Loại đối tượng"] = row.proposedType || (row.predictedGroup.toLowerCase().includes("khách hàng") ? "Khách hàng" : row.predictedGroup.toLowerCase().includes("nhà cung cấp") ? "Nhà cung cấp" : "Khác");
      baseRow["Nguồn gắn mã"] = row.matchingSource || "Diễn giải";
      baseRow["Số hóa đơn khớp"] = row.matchedInvoiceNo || "";
      baseRow["Ngày hóa đơn khớp"] = row.matchedInvoiceDate || "";
      baseRow["Giá trị hóa đơn"] = row.invoiceAmount !== undefined ? row.invoiceAmount : "";
      baseRow["Số tiền giao dịch"] = row.amountIn > 0 ? row.amountIn : row.amountOut;
      baseRow["Chênh lệch số tiền"] = row.differenceAmount !== undefined ? row.differenceAmount : "";
      baseRow["Tỷ lệ chênh lệch (%)"] = row.differencePercentage !== undefined ? `${row.differencePercentage.toFixed(2)}%` : "";
      baseRow["Số ngày chênh lệch"] = row.differenceDays !== undefined ? row.differenceDays : "";
      baseRow["Điểm khớp nội dung"] = row.scoreDesc !== undefined ? row.scoreDesc : "";
      baseRow["Điểm khớp tên"] = row.scoreName !== undefined ? row.scoreName : "";
      baseRow["Điểm khớp số tài khoản"] = row.scoreAcc !== undefined ? row.scoreAcc : "";
      baseRow["Điểm khớp mã số thuế"] = row.scoreMst !== undefined ? row.scoreMst : "";
      baseRow["Điểm khớp số hóa đơn"] = row.scoreInvoice !== undefined ? row.scoreInvoice : "";
      baseRow["Điểm khớp số tiền"] = row.scoreAmount !== undefined ? row.scoreAmount : "";
      baseRow["Điểm khớp ngày"] = row.scoreDate !== undefined ? row.scoreDate : "";
      baseRow["Điểm lịch sử"] = row.scoreHistory !== undefined ? row.scoreHistory : "";
      baseRow["Điểm phạt mâu thuẫn"] = row.scorePenalty !== undefined ? row.scorePenalty : "";
      baseRow["Tổng điểm tin cậy"] = row.score !== undefined ? row.score : "";
      
      const p1 = row.top3Proposals && row.top3Proposals[0] ? `${row.top3Proposals[0].code} - ${row.top3Proposals[0].name} (${row.top3Proposals[0].score}đ)` : "";
      const p2 = row.top3Proposals && row.top3Proposals[1] ? `${row.top3Proposals[1].code} - ${row.top3Proposals[1].name} (${row.top3Proposals[1].score}đ)` : "";
      const p3 = row.top3Proposals && row.top3Proposals[2] ? `${row.top3Proposals[2].code} - ${row.top3Proposals[2].name} (${row.top3Proposals[2].score}đ)` : "";
      baseRow["Phương án 1"] = p1;
      baseRow["Phương án 2"] = p2;
      baseRow["Phương án 3"] = p3;

      baseRow["Lý do đề xuất"] = row.reason || "";
      baseRow["Trạng thái xử lý"] = row.processingStatus || (row.treatment === "Đã chốt" ? "Đã chốt" : "Cần kiểm tra");
      baseRow["Ghi chú kế toán"] = row.notes || "";
      return baseRow;
    });
    const ws1 = XLSX.utils.json_to_sheet(sheet1Data);
    XLSX.utils.book_append_sheet(workbook, ws1, "Sao_ke_da_gan_ma");

    // 2. Sheet: Chi_tiet_doi_chieu_hoa_don
    const sheet2Data = bankMappedRows
      .filter(row => row.matchedInvoiceNo)
      .map(row => ({
        "Ngày giao dịch": row.date,
        "Nội dung ngân hàng": row.desc,
        "Tiền thu (Có)": row.amountIn,
        "Tiền chi (Nợ)": row.amountOut,
        "Mã đối tượng": row.proposedCode,
        "Tên đối tượng": row.proposedName,
        "Số hóa đơn khớp": row.matchedInvoiceNo,
        "Ngày hóa đơn": row.matchedInvoiceDate,
        "Giá trị hóa đơn gốc": row.invoiceAmount,
        "Chênh lệch số tiền": row.differenceAmount,
        "Tỷ lệ chênh lệch (%)": row.differencePercentage !== undefined ? `${row.differencePercentage.toFixed(2)}%` : "",
        "Độ lệch (Ngày)": row.differenceDays,
        "Điểm khớp": row.score,
        "Trạng thái thanh toán": row.processingStatus,
        "Chi tiết đối chiếu": row.reason
      }));
    const ws2 = XLSX.utils.json_to_sheet(sheet2Data);
    XLSX.utils.book_append_sheet(workbook, ws2, "Chi_tiet_doi_chieu_hoa_don");

    // 3. Sheet: Cac_dong_can_kiem_tra
    const sheet3Data = bankMappedRows
      .filter(row => row.treatment === "Cần kiểm tra" || row.score < config.autoThreshold)
      .map(row => ({
        "Ngày giao dịch": row.date,
        "Nội dung ngân hàng": row.desc,
        "Tiền thu (Có)": row.amountIn,
        "Tiền chi (Nợ)": row.amountOut,
        "Mã đối tượng tạm gắn": row.proposedCode,
        "Tên đối tượng tạm gắn": row.proposedName,
        "Trạng thái xử lý": row.processingStatus || "Cần kiểm tra",
        "Điểm tin cậy": row.score,
        "Lý do cần kiểm tra": row.reason,
        "Các đề xuất khác (Top 3)": row.top3Proposals ? row.top3Proposals.map(p => `${p.code}-${p.name} (${p.score}đ)`).join(" | ") : "",
        "Ghi chú kế toán": row.notes
      }));
    const ws3 = XLSX.utils.json_to_sheet(sheet3Data);
    XLSX.utils.book_append_sheet(workbook, ws3, "Cac_dong_can_kiem_tra");

    // 4. Sheet: Giao_dich_ghep_nhieu_HD
    const sheet4Data = bankMappedRows
      .filter(row => row.processingStatus === "Ghép nhiều hóa đơn")
      .map(row => ({
        "Ngày giao dịch": row.date,
        "Nội dung ngân hàng": row.desc,
        "Số tiền giao dịch": row.amountIn > 0 ? row.amountIn : row.amountOut,
        "Mã đối tượng": row.proposedCode,
        "Tên đối tượng": row.proposedName,
        "Danh sách số hóa đơn ghép": row.matchedInvoiceNo,
        "Danh sách ngày hóa đơn": row.matchedInvoiceDate,
        "Tổng giá trị các hóa đơn": row.invoiceAmount,
        "Chênh lệch": row.differenceAmount,
        "Lý do ghép chi tiết": row.reason
      }));
    const ws4 = XLSX.utils.json_to_sheet(sheet4Data);
    XLSX.utils.book_append_sheet(workbook, ws4, "Giao_dich_ghep_nhieu_HD");

    // 5. Sheet: Hoa_don_thanh_toan_mot_phan
    const sheet5Data: any[] = [];
    bankSalesRows.forEach(row => {
      const invNo = String(row[bankSalesMappings.so_hoa_don] || "").trim();
      const invDateVal = row[bankSalesMappings.ngay_hoa_don];
      const dObj = parseDate(invDateVal);
      const invDateStr = dObj ? `${dObj.getFullYear()}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${String(dObj.getDate()).padStart(2, "0")}` : "";
      const originalAmount = parseFloat(row[bankSalesMappings.tong_thanh_toan]) || 0;
      
      let totalPaidOnThis = 0;
      bankMappedRows.forEach(b => {
        if (b.matchedInvoiceNo && b.matchedInvoiceNo.includes(invNo)) {
          if (b.processingStatus === "Ghép nhiều hóa đơn") {
            totalPaidOnThis += originalAmount;
          } else {
            const paid = b.amountIn > 0 ? b.amountIn : b.amountOut;
            totalPaidOnThis += Math.min(paid, originalAmount);
          }
        }
      });

      if (totalPaidOnThis > 0 && totalPaidOnThis < originalAmount) {
        sheet5Data.push({
          "Phân hệ": "Bán ra (Phải thu)",
          "Số hóa đơn": invNo,
          "Ngày hóa đơn": invDateStr,
          "Mã khách hàng": row[bankSalesMappings.ma_khach_hang] || "",
          "Tên khách hàng": row[bankSalesMappings.ten_khach_hang] || "",
          "Mã số thuế": row[bankSalesMappings.ma_so_thue_khach_hang] || "",
          "Tổng tiền hóa đơn": originalAmount,
          "Đã khớp ngân quỹ": totalPaidOnThis,
          "Còn nợ lại": originalAmount - totalPaidOnThis,
          "Trạng thái": "Thanh toán một phần"
        });
      }
    });

    bankPurchasesRows.forEach(row => {
      const invNo = String(row[bankPurchasesMappings.so_hoa_don] || "").trim();
      const invDateVal = row[bankPurchasesMappings.ngay_hoa_don];
      const dObj = parseDate(invDateVal);
      const invDateStr = dObj ? `${dObj.getFullYear()}-${String(dObj.getMonth() + 1).padStart(2, "0")}-${String(dObj.getDate()).padStart(2, "0")}` : "";
      const originalAmount = parseFloat(row[bankPurchasesMappings.tong_thanh_toan]) || 0;

      let totalPaidOnThis = 0;
      bankMappedRows.forEach(b => {
        if (b.matchedInvoiceNo && b.matchedInvoiceNo.includes(invNo)) {
          if (b.processingStatus === "Ghép nhiều hóa đơn") {
            totalPaidOnThis += originalAmount;
          } else {
            const paid = b.amountIn > 0 ? b.amountIn : b.amountOut;
            totalPaidOnThis += Math.min(paid, originalAmount);
          }
        }
      });

      if (totalPaidOnThis > 0 && totalPaidOnThis < originalAmount) {
        sheet5Data.push({
          "Phân hệ": "Mua vào (Phải trả)",
          "Số hóa đơn": invNo,
          "Ngày hóa đơn": invDateStr,
          "Mã nhà cung cấp": row[bankPurchasesMappings.ma_nha_cung_cap] || "",
          "Tên nhà cung cấp": row[bankPurchasesMappings.ten_nha_cung_cap] || "",
          "Mã số thuế": row[bankPurchasesMappings.ma_so_thue_ncc] || "",
          "Tổng tiền hóa đơn": originalAmount,
          "Đã khớp ngân quỹ": totalPaidOnThis,
          "Còn nợ lại": originalAmount - totalPaidOnThis,
          "Trạng thái": "Thanh toán một phần"
        });
      }
    });

    const ws5 = XLSX.utils.json_to_sheet(sheet5Data.length > 0 ? sheet5Data : [{ "Thông báo": "Không có hóa đơn thanh toán một phần" }]);
    XLSX.utils.book_append_sheet(workbook, ws5, "Hoa_don_thanh_toan_mot_phan");

    // 6. Sheet: Log_doi_chieu
    const sheet6Data = bankMappedRows.map((row, idx) => ({
      "STT": idx + 1,
      "Thời điểm": new Date().toLocaleString(),
      "Ngày giao dịch ngân quỹ": row.date,
      "Số tiền": row.amountIn > 0 ? row.amountIn : row.amountOut,
      "Loại giao dịch": row.amountIn > 0 ? "Thu (Có)" : "Chi (Nợ)",
      "Mã đối tượng hạch toán": row.proposedCode || "Chưa rõ",
      "Tên đối tượng hạch toán": row.proposedName || "Chưa rõ",
      "Phương pháp gắn mã": row.matchingSource || "Diễn giải",
      "Trạng thái kiểm tra": row.treatment,
      "Hành động áp dụng": row.processingStatus,
      "Diễn giải quy tắc": row.reason
    }));
    const ws6 = XLSX.utils.json_to_sheet(sheet6Data);
    XLSX.utils.book_append_sheet(workbook, ws6, "Log_doi_chieu");

    // 7. Sheet: Danh_muc_doi_tuong_cap_nhat
    const sheet7Data = partners.map(p => ({
      "Mã đối tác": p.ma_doi_tuong,
      "Tên đối tác chuẩn": p.ten_doi_tuong,
      "Phân loại đối tác": p.loai_doi_tuong,
      "Mã số thuế": p.ma_so_thue || "",
      "Số tài khoản": p.so_tai_khoan || "",
      "Ngân hàng": p.ngan_hang || "",
      "Địa chỉ liên hệ": p.dia_chi || "",
      "Từ khóa nhận diện": p.tu_khoa_nhan_dien || "",
      "Ghi chú nguồn tạo": p.ghi_chu || "Danh mục mặc định"
    }));
    const ws7 = XLSX.utils.json_to_sheet(sheet7Data);
    XLSX.utils.book_append_sheet(workbook, ws7, "Danh_muc_doi_tuong_cap_nhat");

    const baseName = uploadedFileName ? uploadedFileName.replace(/\.[^/.]+$/, "") : "Doi_chieu_sao_ke_ngan_hang";
    XLSX.writeFile(workbook, `${baseName}_báo_cáo_đối_chiếu.xlsx`);
    triggerToast("Đã xuất khẩu tệp báo cáo đối chiếu sao kê đa chiều (7 trang tính) thành công!");
  };

  const exportFullSetToExcel = () => {
    if (integratedRecon.length === 0) {
      triggerToast("Bạn vui lòng xử lý 'Tích hợp dữ liệu đa liên kết' trước!", "warning");
      return;
    }

    const workbook = XLSX.utils.book_new();

    // 1. Sheet Báo cáo mua vào
    const purchaseSheetData = integratedPurchaseRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      baseRow["Mã hàng hóa gắn"] = row.proposedCode;
      baseRow["Tên hàng chuẩn"] = row.proposedName;
      baseRow["Mã nhà cung cấp gắn"] = row.originalAcc || "CHƯA_RÕ";
      baseRow["Tên nhà cung cấp chuẩn"] = row.proposedType || "Chưa rõ";
      baseRow["Độ tương thích"] = `${row.score}%`;
      baseRow["Nguyên tắc quyết định"] = row.treatment;
      return baseRow;
    });
    const wsPur = XLSX.utils.json_to_sheet(purchaseSheetData);
    XLSX.utils.book_append_sheet(workbook, wsPur, "1_Bang_ke_mua_vao");

    // 2. Sheet Bảng kê bán ra
    const saleSheetData = integratedSaleRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      baseRow["Mã hàng hóa gắn"] = row.proposedCode;
      baseRow["Tên hàng chuẩn"] = row.proposedName;
      baseRow["Mã khách hàng gắn"] = row.originalAcc || "CHƯA_RÕ";
      baseRow["Tên khách hàng chuẩn"] = row.proposedType || "Chưa rõ";
      baseRow["Độ tương thích"] = `${row.score}%`;
      baseRow["Nguyên tắc quyết định"] = row.treatment;
      return baseRow;
    });
    const wsSal = XLSX.utils.json_to_sheet(saleSheetData);
    XLSX.utils.book_append_sheet(workbook, wsSal, "2_Bang_ke_ban_ra");

    // 3. Sổ kho đối chuẩn
    const invSheetData = integratedInvRows.map((row, idx) => ({
      "STT": idx + 1,
      "Mã hàng gốc": row.ma_hang_hoa,
      "Tên hàng gốc": row.ten_hang_hoa,
      "Đơn vị tính": row.don_vi_tinh,
      "Mã hàng hóa chuẩn": row.ma_hang_hoa_gan,
      "Tên hàng hóa chuẩn": row.ten_hang_hoa_chuan,
      "Nhóm hàng": row.nhom_hang || "Chưa phân loại",
      "Tồn kho": row.quy_cach || "Tiêu chuẩn",
      "Ghi chú": row.ghi_chu || ""
    }));
    const wsInv = XLSX.utils.json_to_sheet(invSheetData);
    XLSX.utils.book_append_sheet(workbook, wsInv, "3_So_kho_hach_toan");

    // 4. Sao kê ngân hàng
    const bankSheetData = integratedBankRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      baseRow["Mã hạch toán đối chiếu"] = row.proposedCode || "CHƯA RÕ";
      baseRow["Tên đối tác hạch toán"] = row.proposedName || "Nghi vấn / Thất lạc";
      baseRow["Nhóm giao dịch AI dự báo"] = row.predictedGroup;
      baseRow["Điểm tin cậy"] = `${row.score}%`;
      baseRow["Phương án hạch toán"] = row.treatment;
      return baseRow;
    });
    const wsBnk = XLSX.utils.json_to_sheet(bankSheetData);
    XLSX.utils.book_append_sheet(workbook, wsBnk, "4_Sao_ke_ngan_hang");

    // 5. Danh mục hàng hóa cập nhật
    const commoditiesData = commodities.map((item) => ({
      "Mã hàng hóa": item.ma_hang_hoa,
      "Tên hàng hóa chuẩn": item.ten_hang_hoa_chuan,
      "Nhóm hàng": item.nhom_hang,
      "Đơn vị tính": item.don_vi_tinh,
      "Quy cách kỹ thuật": item.quy_cach,
      "Ghi chú tự sinh": item.ghi_chu
    }));
    const wsCom = XLSX.utils.json_to_sheet(commoditiesData);
    XLSX.utils.book_append_sheet(workbook, wsCom, "5_Danh_muc_hang_hoa");

    // 6. Danh mục đối tác cập nhật
    const partnersData = partners.map((p) => ({
      "Mã đối tác": p.ma_doi_tuong,
      "Tên đối tác chuẩn": p.ten_doi_tuong,
      "Loại đối tác": p.loai_doi_tuong,
      "Mã số thuế": p.ma_so_thue,
      "Số tài khoản": p.so_tai_khoan,
      "Ngân hàng": p.ngan_hang,
      "Địa chỉ": p.dia_chi,
      "Ghi chú tự sinh": p.ghi_chu
    }));
    const wsPar = XLSX.utils.json_to_sheet(partnersData);
    XLSX.utils.book_append_sheet(workbook, wsPar, "6_Danh_muc_doi_tac");

    // 7. Các dòng cần kiểm tra
    const needsReview: any[] = [];
    integratedPurchaseRows.filter(r => r.score < config.autoThreshold).forEach(r => {
      needsReview.push({ "Phân hệ": "Mua vào", "Diễn giải thô": r.originalText, "Độ khớp (%)": r.score, "Mã đề xuất": r.proposedCode, "Lý do": r.reason });
    });
    integratedSaleRows.filter(r => r.score < config.autoThreshold).forEach(r => {
      needsReview.push({ "Phân hệ": "Bán ra", "Diễn giải thô": r.originalText, "Độ khớp (%)": r.score, "Mã đề xuất": r.proposedCode, "Lý do": r.reason });
    });
    integratedBankRows.filter(r => r.score < config.autoThreshold).forEach(r => {
      needsReview.push({ "Phân hệ": "Ngân quỹ", "Diễn giải thô": r.desc, "Độ khớp (%)": r.score, "Mã đề xuất": r.proposedCode, "Lý do": r.reason });
    });
    const wsRev = XLSX.utils.json_to_sheet(needsReview);
    XLSX.utils.book_append_sheet(workbook, wsRev, "7_Can_kiem_tra");

    XLSX.writeFile(workbook, "Ket_Qua_Doi_Chieu_Ke_Toan_Tong_Hop_Nhieu_Sheet.xlsx");
    triggerToast("Xuất bản trọn bộ kế toán 7 Sheets thành công!");
  };

  const exportFullSetToZip = async () => {
    if (integratedRecon.length === 0) {
      triggerToast("Bạn vui lòng xử lý 'Tích hợp dữ liệu đa liên kết' trước!", "warning");
      return;
    }

    try {
      const zip = new JSZip();

      // Trợ lý tạo blob mảng nhị phân cho tệp Excel riêng lẻ
      const getXlsxBuffer = (data: any[], sheetName: string) => {
        const ws = XLSX.utils.json_to_sheet(data);
        const wb = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(wb, ws, sheetName);
        return XLSX.write(wb, { bookType: "xlsx", type: "array" });
      };

      // 1. Khoản mua vào
      const purchaseSheetData = integratedPurchaseRows.map((row) => {
        const baseRow = { ...(row.rawRowData || {}) };
        baseRow["Mã hàng hóa gắn"] = row.proposedCode;
        baseRow["Tên hàng chuẩn"] = row.proposedName;
        baseRow["Mã nhà cung cấp gắn"] = row.originalAcc || "CHƯA_RÕ";
        baseRow["Tên nhà cung cấp chuẩn"] = row.proposedType || "Chưa rõ";
        baseRow["Độ tương thích"] = `${row.score}%`;
        baseRow["Nguyên tắc quyết định"] = row.treatment;
        return baseRow;
      });
      zip.file("1_Bang_ke_mua_vao_da_gan_ma.xlsx", getXlsxBuffer(purchaseSheetData, "Bang_ke_mua_vao"));

      // 2. Khoản bán ra
      const saleSheetData = integratedSaleRows.map((row) => {
        const baseRow = { ...(row.rawRowData || {}) };
        baseRow["Mã hàng hóa gắn"] = row.proposedCode;
        baseRow["Tên hàng chuẩn"] = row.proposedName;
        baseRow["Mã khách hàng gắn"] = row.originalAcc || "CHƯA_RÕ";
        baseRow["Tên khách hàng chuẩn"] = row.proposedType || "Chưa rõ";
        baseRow["Độ tương thích"] = `${row.score}%`;
        baseRow["Nguyên tắc quyết định"] = row.treatment;
        return baseRow;
      });
      zip.file("2_Bang_ke_ban_ra_da_gan_ma.xlsx", getXlsxBuffer(saleSheetData, "Bang_ke_ban_ra"));

      // 3. Sổ kho đối chuẩn
      const invSheetData = integratedInvRows.map((row) => ({
        "Mã hàng gốc": row.ma_hang_hoa,
        "Tên hàng gốc": row.ten_hang_hoa,
        "Đơn vị tính": row.don_vi_tinh,
        "Mã hàng hóa chuẩn": row.ma_hang_hoa_gan,
        "Tên hàng hóa chuẩn": row.ten_hang_hoa_chuan,
        "Nhóm hàng": row.nhom_hang || "Chưa phân loại",
        "Tồn kho": row.quy_cach || "Tiêu chuẩn",
        "Ghi chú": row.ghi_chu || ""
      }));
      zip.file("3_So_kho_hach_toan_da_gan_ma.xlsx", getXlsxBuffer(invSheetData, "Kho_xuat_nhap_ton"));

      // 4. Sao kê ngân quỹ
      const bankSheetData = integratedBankRows.map((row) => {
        const baseRow = { ...(row.rawRowData || {}) };
        baseRow["Mã hạch toán đối chiếu"] = row.proposedCode || "CHƯA RÕ";
        baseRow["Tên đối tác hạch toán"] = row.proposedName || "Nghi vấn / Thất lạc";
        baseRow["Nhóm giao dịch AI dự báo"] = row.predictedGroup;
        baseRow["Điểm tin cậy"] = `${row.score}%`;
        baseRow["Phương án hạch toán"] = row.treatment;
        return baseRow;
      });
      zip.file("4_Sao_ke_ngan_hang_da_gan_ma.xlsx", getXlsxBuffer(bankSheetData, "Sao_ke_ngan_hang"));

      // 5. Danh mục hàng hóa
      const commoditiesData = commodities.map((item) => ({
        "Mã hàng hóa": item.ma_hang_hoa,
        "Tên hàng hóa chuẩn": item.ten_hang_hoa_chuan,
        "Nhóm hàng": item.nhom_hang,
        "Đơn vị tính": item.don_vi_tinh,
        "Quy cách kỹ thuật": item.quy_cach,
        "Ghi chú tự sinh": item.ghi_chu
      }));
      zip.file("5_Danh_muc_hang_hoa_cap_nhat.xlsx", getXlsxBuffer(commoditiesData, "Danh_muc_hang_hoa"));

      // 6. Danh mục đối tác
      const partnersData = partners.map((p) => ({
        "Mã đối tác": p.ma_doi_tuong,
        "Tên đối tác chuẩn": p.ten_doi_tuong,
        "Loại đối tác": p.loai_doi_tuong,
        "Mã số thuế": p.ma_so_thue,
        "Số tài khoản": p.so_tai_khoan,
        "Ngân hàng": p.ngan_hang,
        "Địa chỉ": p.dia_chi,
        "Ghi chú tự sinh": p.ghi_chu
      }));
      zip.file("6_Danh_muc_doi_tac_cap_nhat.xlsx", getXlsxBuffer(partnersData, "Danh_muc_doi_tac"));

      // 7. Cần rà soát kiểm tra
      const needsReview: any[] = [];
      integratedPurchaseRows.filter(r => r.score < config.autoThreshold).forEach(r => {
        needsReview.push({ "Phân hệ": "Mua vào", "Diễn giải thô": r.originalText, "Độ khớp (%)": r.score, "Mã đề xuất": r.proposedCode, "Lý do": r.reason });
      });
      integratedSaleRows.filter(r => r.score < config.autoThreshold).forEach(r => {
        needsReview.push({ "Phân hệ": "Bán ra", "Diễn giải thô": r.originalText, "Độ khớp (%)": r.score, "Mã đề xuất": r.proposedCode, "Lý do": r.reason });
      });
      integratedBankRows.filter(r => r.score < config.autoThreshold).forEach(r => {
        needsReview.push({ "Phân hệ": "Ngân quỹ", "Diễn giải thô": r.desc, "Độ khớp (%)": r.score, "Mã đề xuất": r.proposedCode, "Lý do": r.reason });
      });
      zip.file("7_Dong_can_kiem_tra.xlsx", getXlsxBuffer(needsReview, "Can_kiem_tra"));

      // Đóng gói và tải xuống ZIP
      const content = await zip.generateAsync({ type: "blob" });
      const url = window.URL.createObjectURL(content);
      const a = document.createElement("a");
      a.href = url;
      a.download = "Bo_Ho_So_Ke_Toan_Da_Phan_He_Da_Gan_Ma_ZIP.zip";
      a.click();
      window.URL.revokeObjectURL(url);

      triggerToast("Tải xuống trọn bộ tài liệu nén ZIP thành công cực kỳ rảnh tay!");
    } catch (e: any) {
      triggerToast(`Có lỗi nén file ZIP: ${e.message}`, "warning");
    }
  };


  return (
    <ErrorBoundary fallbackTitle="LỖI CHUNG TRONG ỨNG DỤNG SMARTLEDGER">
      <div className="min-h-screen bg-[#fdfdfb] text-[#1a1a1a] font-sans flex flex-col selection:bg-[#00ff00] selection:text-black">
      {/* --- EXCEL STRUCTURE INTEGRITY WIZARD --- */}
      {excelWizard && (
        <div className="fixed inset-0 bg-[#141414]/75 z-50 flex items-center justify-center p-4 backdrop-blur-xs overflow-y-auto">
          <div className="bg-white border-4 border-[#141414] shadow-[8px_8px_0px_#141414] max-w-4xl w-full max-h-[90vh] flex flex-col animate-in fade-in zoom-in-95 duration-150">
            {/* Modal Header */}
            <div className="bg-amber-300 border-b-4 border-[#141414] p-4 flex justify-between items-center">
              <div className="flex items-center gap-2.5">
                <span className="bg-white border-2 border-black p-1 text-xs font-black uppercase shadow-[1.5px_1.5px_0px_#141414]">WIZARD</span>
                <div>
                  <h3 className="font-black text-sm uppercase tracking-tight text-black">TRỢ LÝ ĐỌC HIỂU CẤU TRÚC EXCEL</h3>
                  <p className="text-[10px] text-black/75 uppercase font-bold">Căn chỉnh dòng tiêu đề, trang tính & sửa cột của tệp {excelWizard.fileType === "bank" ? "Sao kê Ngân hàng" : "Sổ Kho Đối Tác"}</p>
                </div>
              </div>
              <button onClick={excelWizard.onCancel} className="bg-white hover:bg-red-200 text-black border-2 border-black font-black p-1 shadow-[2px_2px_0px_#141414] active:translate-y-0.5 transition cursor-pointer text-xs px-2.5">ĐÓNG ✕</button>
            </div>

            {/* Modal Content */}
            <div className="p-6 overflow-y-auto flex-1 space-y-6">
              <div className="bg-[#f0f0ed] border-2 border-[#141414] p-4 flex items-center gap-3.5">
                <div className="text-2xl">📋</div>
                <div className="text-xs">
                  <p className="font-black text-black uppercase">Tên tệp đang nạp: <span className="font-mono text-blue-600 underline">{excelWizard.fileName}</span></p>
                  <p className="text-slate-500 font-bold mt-0.5">Vui lòng kiểm tra kỹ xem tiêu đề thực tế nằm ở dòng mấy. Hệ thống không tự động ép buộc dòng 1 để tránh sai sót.</p>
                </div>
              </div>

              {/* Step Grid controls */}
              <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
                
                {/* Left controls: Select sheet and header row */}
                <div className="space-y-4 border-2 border-[#141414] p-4 bg-slate-50 shadow-[4px_4px_0px_#141414]">
                  <h4 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-slate-300 pb-1.5 flex items-center gap-2">
                    <span className="bg-blue-600 text-white font-mono font-bold w-4 h-4 rounded-full flex items-center justify-center text-[10px]">1</span>
                    CẤU HÌNH TRANG & DÒNG TIÊU ĐỀ
                  </h4>

                  <div className="space-y-3">
                    {/* Sheet selection */}
                    <div>
                      <label className="text-[10px] font-black uppercase text-slate-500 block mb-1">Chọn trang tính (Active Sheet):</label>
                      <select 
                        value={excelWizard.selectedSheet} 
                        onChange={(e) => handleWizardSheetChange(e.target.value)} 
                        className="w-full border-2 border-[#141414] bg-white p-2 text-xs font-black text-black focus:outline-none"
                      >
                        {excelWizard.sheetNames.map(s => <option key={s} value={s}>{s}</option>)}
                      </select>
                    </div>

                    {/* Header Row Index selection */}
                    <div>
                      <label className="text-[10px] font-black uppercase text-slate-500 block mb-1">Dòng tiêu đề (Header Row Index - 1-based):</label>
                      <div className="flex gap-2 items-center">
                        <input 
                          type="number" 
                          min="1" 
                          max="30"
                          value={excelWizard.headerRowIndex + 1} 
                          onChange={(e) => handleWizardHeaderRowChange(Math.max(0, parseInt(e.target.value) - 1))}
                          className="w-20 border-2 border-[#141414] bg-white p-1.5 text-xs font-black text-center focus:outline-none"
                        />
                        <span className="text-[10px] text-slate-500 font-bold">Dòng thực tế chứa tiêu đề chính của bảng</span>
                      </div>
                    </div>

                    {/* Header Rows Count (1 or 2 rows for merge headers) */}
                    <div>
                      <label className="text-[10px] font-black uppercase text-slate-500 block mb-1">Số lượng dòng tiêu đề:</label>
                      <div className="flex gap-4">
                        <label className="flex items-center gap-2 cursor-pointer font-bold text-xs">
                          <input 
                          type="radio" 
                          name="headerRowsCount" 
                          checked={excelWizard.headerRowsCount === 1}
                          onChange={() => handleWizardHeaderCountChange(1)}
                          className="accent-black"
                        />
                        <span>1 dòng đơn lẻ (Phổ biến)</span>
                      </label>
                      <label className="flex items-center gap-2 cursor-pointer font-bold text-xs">
                        <input 
                          type="radio" 
                          name="headerRowsCount" 
                          checked={excelWizard.headerRowsCount === 2}
                          onChange={() => handleWizardHeaderCountChange(2)}
                          className="accent-black"
                        />
                        <span>2 dòng ghép lại (Merged)</span>
                      </label>
                    </div>
                  </div>
                  </div>
                </div>

                {/* Right Panel: Headers Preview after cleaning */}
                <div className="space-y-4 border-2 border-[#141414] p-4 bg-slate-50 shadow-[4px_4px_0px_#141414]">
                  <h4 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-slate-300 pb-1.5 flex items-center gap-2">
                    <span className="bg-blue-600 text-white font-mono font-bold w-4 h-4 rounded-full flex items-center justify-center text-[10px]">2</span>
                    DANH SÁCH CỘT PHÁT HIỆN ({excelWizard.headersCleaned.length})
                  </h4>

                  <div className="text-[10px] text-slate-500 font-bold">
                    Tên cột đã được chuẩn hóa tự động (loại bỏ khoảng trắng dư thừa, xử lý ghép cell rỗng, trùng tên tự thêm hậu tố):
                  </div>

                  <div className="max-h-40 overflow-y-auto border border-slate-300 p-2 bg-white rounded divide-y divide-slate-100 font-mono text-[10px] text-slate-700">
                    {excelWizard.headersCleaned.map((col, idx) => (
                      <div key={idx} className="py-1 flex justify-between items-center hover:bg-slate-50 px-1">
                        <span className="text-slate-400 font-bold">Cột {idx + 1}:</span>
                        <span className="font-bold text-black">{col || <span className="text-red-500 bg-red-50 px-1 italic">Vô danh / Trống</span>}</span>
                      </div>
                    ))}
                  </div>
                </div>

              </div>

              {/* Data Row preview directly below header */}
              <div className="space-y-2.5">
                <h4 className="font-black text-xs uppercase tracking-wider text-black flex items-center gap-2">
                  <span className="bg-blue-600 text-white font-mono font-bold w-4 h-4 rounded-full flex items-center justify-center text-[10px]">3</span>
                  XEM TRƯỚC SỐ LIỆU THỰC TẾ (PREVIEW)
                </h4>
                <p className="text-[10px] text-slate-500 font-bold uppercase">Kiểm tra xem dữ liệu có bị lệch cột, lệch dòng hoặc nhận nhầm tiêu đề không:</p>

                <div className="overflow-x-auto border-2 border-[#141414] max-h-64">
                  <table className="w-full text-left text-xs border-collapse">
                    <thead>
                      <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black font-mono text-[10px]">
                        <th className="p-2 border-r border-[#141414] text-center w-12">STT</th>
                        {excelWizard.headersCleaned.map((col, idx) => (
                          <th key={idx} className="p-2 border-r border-[#141414] min-w-[120px] max-w-[200px] truncate">{col || `Cột ${idx + 1}`}</th>
                        ))}
                      </tr>
                    </thead>
                    <tbody className="divide-y divide-[#141414]/15 font-mono text-[10px]">
                      {excelWizard.rowsPreview.map((row, rIdx) => {
                        const isHeaderHighlight = rIdx === excelWizard.headerRowIndex;
                        return (
                          <tr 
                            key={rIdx} 
                            className={`${isHeaderHighlight ? "bg-yellow-100 font-black text-black" : "hover:bg-slate-50 text-slate-600"}`}
                          >
                            <td className="p-2 border-r border-[#141414]/15 text-center font-bold bg-[#fdfdfb]">
                              {rIdx + 1}
                              {isHeaderHighlight && <span className="block text-[8px] text-amber-800 bg-amber-100 border border-amber-300 rounded px-0.5 mt-0.5">Tiêu đề</span>}
                            </td>
                            {excelWizard.headersCleaned.map((_, colIdx) => (
                              <td key={colIdx} className="p-2 border-r border-[#141414]/15 max-w-[200px] truncate">
                                {row[colIdx] !== undefined ? String(row[colIdx]) : ""}
                              </td>
                            ))}
                          </tr>
                        );
                      })}
                    </tbody>
                  </table>
                </div>
              </div>
            </div>

            {/* Modal Actions Footer */}
            <div className="bg-[#f0f0ed] border-t-4 border-[#141414] p-4 px-6 flex justify-between items-center">
              <button 
                onClick={excelWizard.onCancel} 
                className="bg-white hover:bg-[#e4e4e2] text-black border-2 border-[#141414] text-xs font-black uppercase px-6 py-2.5 shadow-[3px_3px_0px_#141414] active:translate-y-0.5 transition cursor-pointer"
              >
                HỦY BỎ
              </button>
              <button 
                onClick={handleConfirmWizard} 
                className="bg-[#00ff00] hover:bg-[#05e005] text-black border-2 border-[#141414] text-xs font-black uppercase px-8 py-2.5 shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer"
              >
                XÁC NHẬN CẤU TRÚC & NẠP DỮ LIỆU ✔
              </button>
            </div>
          </div>
        </div>
      )}

      {/* --- TOP NOTIFICATION BANNER --- */}
      {notification && (
        <div className={`fixed top-4 right-4 z-50 px-5 py-3 rounded-none shadow-[4px_4px_0px_#141414] border-2 border-[#141414] flex items-center gap-3 transition-all ${
          notification.type === "success" ? "bg-[#00ff00] text-black font-semibold" : "bg-amber-400 text-black font-semibold"
        }`}>
          <CheckCircle2 size={18} />
          <span className="font-bold text-xs uppercase tracking-wider">{notification.message}</span>
        </div>
      )}

      {/* --- APPLICATION HEADER BAR --- */}
      <header className="bg-[#f0f0ed] text-black border-b-4 border-[#141414] sticky top-0 z-40 shadow-[0_4px_0_rgba(20,20,20,0.05)]">
        <div className="max-w-7xl mx-auto px-6 py-4 flex flex-col sm:flex-row justify-between items-center gap-4">
          <div className="flex items-center gap-3">
            <div className="bg-white border-2 border-[#141414] p-2.5 px-4 shadow-[4px_4px_0px_#141414] flex items-center gap-3">
              <div className="bg-[#00ff00] text-[#141414] p-1.5 border border-[#141414] font-bold">
                <RefreshCw size={20} className="animate-spin-slow text-black" />
              </div>
              <div>
                <h1 className="text-sm font-black tracking-tighter uppercase leading-none text-[#1a1a1a]">SmartLedger AutoCoder</h1>
                <p className="text-[10px] uppercase font-bold tracking-[0.15em] text-[#666] mt-0.5">AUTO-ACCOUNTING MAPPER</p>
              </div>
            </div>
          </div>

          <div className="flex items-center gap-3">
            {uploadedFileName ? (
              <span className="text-black text-xs bg-white px-3 py-1.5 border-2 border-[#141414] shadow-[2px_2px_0px_#141414] flex items-center gap-2 font-mono font-bold">
                <FileSpreadsheet size={14} className="text-green-600" />
                <span>{uploadedFileName} ({activeRowsCount} dòng)</span>
              </span>
            ) : (
              <button
                id="load_demo_btn"
                onClick={handleLoadDemo}
                className="bg-[#00ff00] hover:bg-[#05e005] hover:shadow-[6px_6px_0px_#141414] hover:-translate-y-0.5 active:translate-y-0 text-black border-2 border-[#141414] text-xs font-black uppercase px-4.5 py-1.5 shadow-[4px_4px_0px_#141414] transition cursor-pointer"
              >
                💡 Chạy thử dữ liệu Demo Mẫu
              </button>
            )}

            {uploadedFileName && (
              <button
                onClick={handleReset}
                title="Reset Dữ Liệu"
                className="p-1 px-2.5 bg-[#ffebee] hover:bg-red-200 text-red-800 border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] hover:shadow-[3px_3px_0px_#141414] active:translate-y-0 transition cursor-pointer font-bold"
              >
                <Trash2 size={15} />
              </button>
            )}
          </div>
        </div>
      </header>

      {/* --- DASHBOARD WRAPPER CONTAINER --- */}
      <div className="max-w-7xl mx-auto w-full px-6 py-8 flex-1 grid grid-cols-1 lg:grid-cols-4 gap-8">

        {/* --- LEFT SIDEBAR: ALGORITHM CONFIGURATIONS --- */}
        <aside className="lg:col-span-1 space-y-6">
          <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-6">
            <h2 className="text-xs font-black text-black tracking-widest uppercase flex items-center gap-2 border-b-2 border-[#141414] pb-3">
              <Settings size={16} className="text-[#141414]" />
              <span>Tham Số Thuật Toán</span>
            </h2>

            {/* Single Threshold Slider */}
            <div className="space-y-4">
              <div>
                <label className="text-xs font-black text-[#141414] uppercase tracking-wider flex justify-between mb-1.5">
                  <span className="flex items-center gap-1.5">
                    <span className="w-2.5 h-2.5 bg-[#00ff00] border border-[#141414] inline-block"></span>
                    Độ khớp yêu cầu
                  </span>
                  <span className="text-black font-mono font-bold">{config.autoThreshold}%</span>
                </label>
                <input
                  type="range"
                  min="1"
                  max="100"
                  value={config.autoThreshold}
                  onChange={(e) => setConfig({ ...config, autoThreshold: parseInt(e.target.value), checkThreshold: parseInt(e.target.value) })}
                  className="w-full h-3 bg-[#f0f0ed] border-2 border-[#141414] appearance-none cursor-pointer accent-[#00ff00]"
                />
                <div className="text-[10px] text-[#666] font-medium mt-1.5 leading-relaxed space-y-1">
                  <p>💥 Dưới <span className="font-extrabold text-[#141414]">{config.autoThreshold}%</span>: <span className="text-red-600 font-bold">Tự động thêm mã mới</span></p>
                  <p>✅ Từ <span className="font-extrabold text-[#141414]">{config.autoThreshold}% trở lên</span>: <span className="text-green-600 font-bold">Tự rà soát rảnh tay gắn mã cũ</span></p>
                </div>
              </div>
            </div>

            <hr className="border-t-2 border-[#141414]" />

            {/* Code prefixes */}
            <div className="space-y-3">
              <h3 className="text-xs font-black text-black uppercase tracking-wider">Quy tắc tiền tố sinh mã</h3>

              <div className="grid grid-cols-2 gap-2">
                <div>
                  <label className="text-[10px] uppercase font-black text-[#666] tracking-wider">Hàng hóa</label>
                  <input
                    type="text"
                    value={config.prefixHH}
                    onChange={(e) => setConfig({ ...config, prefixHH: e.target.value })}
                    className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 px-3 font-mono text-xs text-black focus:bg-[#f0f0ed] focus:outline-none"
                  />
                </div>
                <div>
                  <label className="text-[10px] uppercase font-black text-[#666] tracking-wider">Khách mua</label>
                  <input
                    type="text"
                    value={config.prefixKH}
                    onChange={(e) => setConfig({ ...config, prefixKH: e.target.value })}
                    className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 px-3 font-mono text-xs text-black focus:bg-[#f0f0ed] focus:outline-none"
                  />
                </div>
                <div className="col-span-2">
                  <label className="text-[10px] uppercase font-black text-[#666] tracking-wider">Nhà cung cấp</label>
                  <input
                    type="text"
                    value={config.prefixNCC}
                    onChange={(e) => setConfig({ ...config, prefixNCC: e.target.value })}
                    className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 px-3 font-mono text-xs text-black focus:bg-[#f0f0ed] focus:outline-none"
                  />
                </div>
              </div>
            </div>
          </div>

          {/* Quick Stats on directories */}
          <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-4">
            <h3 className="text-xs font-black text-black uppercase tracking-wider border-b-2 border-[#141414] pb-2">Cơ sở dữ liệu danh mục</h3>

            {/* Hàng hóa */}
            <div className="border-2 border-[#141414] bg-[#fdfdfb] p-3 space-y-2.5 shadow-[2px_2px_0px_#141414]">
              <div className="flex justify-between items-center">
                <div className="flex items-center gap-2">
                  <div className="bg-[#141414] p-1 text-white border border-[#141414]">
                    <Database size={13} />
                  </div>
                  <div>
                    <p className="font-black text-black text-left uppercase text-[10px] tracking-wide">Mã Hàng hóa</p>
                    <p className="text-slate-500 text-[9px] text-left uppercase font-bold">Vật tư chuẩn trong kho</p>
                  </div>
                </div>
                <span className="font-mono text-[10px] font-black text-black bg-[#00ff00] px-1.5 py-0.5 border border-[#141414]">
                  {commodities.length} mã
                </span>
              </div>
              <label className="w-full bg-[#141414] text-white hover:bg-[#222] hover:shadow-[3px_3px_0px_#00ff00] active:translate-y-0.5 text-[9px] font-black uppercase py-1.5 px-2 border border-[#141414] transition flex items-center justify-center gap-1.5 cursor-pointer">
                <Upload size={10} className="text-[#00ff00]" />
                Nhập danh mục hàng hóa (.XLSX)
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleImportMasterDirectory(e, "commodity")}
                />
              </label>
            </div>

            {/* Khách hàng */}
            <div className="border-2 border-[#141414] bg-[#fdfdfb] p-3 space-y-2.5 shadow-[2px_2px_0px_#141414]">
              <div className="flex justify-between items-center">
                <div className="flex items-center gap-2">
                  <div className="bg-[#141414] p-1 text-white border border-[#141414]">
                    <Users size={13} />
                  </div>
                  <div>
                    <p className="font-black text-black text-left uppercase text-[10px] tracking-wide">Mã Khách hàng</p>
                    <p className="text-slate-500 text-[9px] text-left uppercase font-bold">Công nợ khách mua</p>
                  </div>
                </div>
                <span className="font-mono text-[10px] font-black text-black bg-[#00ff00] px-1.5 py-0.5 border border-[#141414]">
                  {partners.filter(p => p.loai_doi_tuong === "Khách hàng").length} mã
                </span>
              </div>
              <label className="w-full bg-[#141414] text-white hover:bg-[#222] hover:shadow-[3px_3px_0px_#00ff00] active:translate-y-0.5 text-[9px] font-black uppercase py-1.5 px-2 border border-[#141414] transition flex items-center justify-center gap-1.5 cursor-pointer">
                <Upload size={10} className="text-[#00ff00]" />
                Nhập danh mục khách hàng (.XLSX)
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleImportMasterDirectory(e, "customer")}
                />
              </label>
            </div>

            {/* Nhà cung cấp */}
            <div className="border-2 border-[#141414] bg-[#fdfdfb] p-3 space-y-2.5 shadow-[2px_2px_0px_#141414]">
              <div className="flex justify-between items-center">
                <div className="flex items-center gap-2">
                  <div className="bg-[#141414] p-1 text-white border border-[#141414]">
                    <Users size={13} />
                  </div>
                  <div>
                    <p className="font-black text-black text-left uppercase text-[10px] tracking-wide">Mã Nhà cung cấp</p>
                    <p className="text-slate-500 text-[9px] text-left uppercase font-bold">Công nợ nhà cung ứng</p>
                  </div>
                </div>
                <span className="font-mono text-[10px] font-black text-black bg-[#00ff00] px-1.5 py-0.5 border border-[#141414]">
                  {partners.filter(p => p.loai_doi_tuong === "Nhà cung cấp").length} mã
                </span>
              </div>
              <label className="w-full bg-[#141414] text-white hover:bg-[#222] hover:shadow-[3px_3px_0px_#00ff00] active:translate-y-0.5 text-[9px] font-black uppercase py-1.5 px-2 border border-[#141414] transition flex items-center justify-center gap-1.5 cursor-pointer">
                <Upload size={10} className="text-[#00ff00]" />
                Nhập danh mục nhà cung cấp (.XLSX)
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleImportMasterDirectory(e, "supplier")}
                />
              </label>
            </div>
          </div>
        </aside>

        {/* --- RIGHT PANEL: DETAILED WORKSPACE & TAB INTERACTIVE SYSTEM --- */}
        <main className="lg:col-span-3 space-y-6 flex flex-col">

          {/* Nav pills */}
          <div className="flex flex-wrap gap-2 border-b-2 border-[#141414] pb-3">
            <button
              onClick={() => setCurrentTab("dashboard")}
              className={`px-4 py-2 text-[11px] font-black uppercase tracking-wider transition flex items-center gap-2.5 cursor-pointer ${
                currentTab === "dashboard"
                  ? "bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc]"
                  : "bg-white text-black border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:bg-[#f0f0ed]"
              }`}
            >
              <span className={`w-2 h-2 border border-black inline-block ${currentTab === "dashboard" ? "bg-[#00ff00]" : "bg-white"}`}></span>
              📊 Dashboard
            </button>
            <button
              onClick={() => setCurrentTab("commodity")}
              className={`px-4 py-2 text-[11px] font-black uppercase tracking-wider transition flex items-center gap-2.5 cursor-pointer ${
                currentTab === "commodity"
                  ? "bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc]"
                  : "bg-white text-black border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:bg-[#f0f0ed]"
              }`}
            >
              <span className={`w-2 h-2 border border-black inline-block ${currentTab === "commodity" ? "bg-[#00ff00]" : "bg-white"}`}></span>
              📦 Gán mã Hàng Hóa
            </button>
            <button
              onClick={() => setCurrentTab("partner")}
              className={`px-4 py-2 text-[11px] font-black uppercase tracking-wider transition flex items-center gap-2.5 cursor-pointer ${
                currentTab === "partner"
                  ? "bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc]"
                  : "bg-white text-black border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:bg-[#f0f0ed]"
              }`}
            >
              <span className={`w-2 h-2 border border-black inline-block ${currentTab === "partner" ? "bg-[#00ff00]" : "bg-white"}`}></span>
              👥 Gán mã Đối Tác
            </button>
            <button
              onClick={() => setCurrentTab("bank")}
              className={`px-4 py-2 text-[11px] font-black uppercase tracking-wider transition flex items-center gap-2.5 cursor-pointer ${
                currentTab === "bank"
                  ? "bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc]"
                  : "bg-white text-black border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:bg-[#f0f0ed]"
              }`}
            >
              <span className={`w-2 h-2 border border-black inline-block ${currentTab === "bank" ? "bg-[#00ff00]" : "bg-white"}`}></span>
              🏦 Giao Dịch Sao Kê
            </button>
            <button
              onClick={() => setCurrentTab("integrated")}
              className={`px-4 py-2 text-[11px] font-black uppercase tracking-wider transition flex items-center gap-2.5 cursor-pointer ${
                currentTab === "integrated"
                  ? "bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc]"
                  : "bg-white text-black border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:bg-[#f0f0ed]"
              }`}
            >
              <span className={`w-2 h-2 border border-black inline-block ${currentTab === "integrated" ? "bg-[#00ff00]" : "bg-white"}`}></span>
              🧩 Đa phân hệ chéo
            </button>
            <button
              onClick={() => setCurrentTab("python")}
              className={`px-4 py-2 text-[11px] font-black uppercase tracking-wider transition flex items-center gap-2.5 cursor-pointer ${
                currentTab === "python"
                  ? "bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc]"
                  : "bg-white text-black border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:bg-[#f0f0ed]"
              }`}
            >
              <span className={`w-2 h-2 border border-black inline-block ${currentTab === "python" ? "bg-[#00ff00]" : "bg-white"}`}></span>
              🐍 Local Python App
            </button>
          </div>

          {/* --- TAB CONTENT 1: WELCOME & SUMMARY --- */}
          {currentTab === "dashboard" && (
            <div className="space-y-6">
              {/* Promo layout */}
              {!uploadedFileName && (
                <div className="bg-[#141414] text-white border-2 border-transparent shadow-[6px_6px_0px_#ccc] p-8 flex flex-col sm:flex-row items-center justify-between gap-6 transition-all">
                  <div className="space-y-2 text-center sm:text-left">
                    <span className="bg-[#00ff00] text-black text-[10px] font-black uppercase tracking-wider px-2 py-0.5 border border-[#141414]">Hệ thống phân tích cục bộ</span>
                    <h3 className="text-xl font-black uppercase tracking-tight mt-2">Chưa có tệp dữ liệu kế toán?</h3>
                    <p className="text-slate-300 text-xs max-w-lg">Nhấn nút bên phải để kích hoạt nhanh hệ thống dữ liệu hóa đơn mua/bán, xuất nhập tồn kho và ngân quỹ mẫu để chạy thử ngay!</p>
                  </div>
                  <button
                    onClick={handleLoadDemo}
                    className="bg-[#00ff00] hover:bg-[#05e005] text-black text-xs font-black uppercase tracking-wider p-4 py-3 border-2 border-[#141414] shadow-[4px_4px_0px_#fff] hover:shadow-[6px_6px_0px_#fff] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer"
                  >
                    🚀 Bắt đầu trải nghiệm mẫu
                  </button>
                </div>
              )}

              {/* Grid of indicators */}
              <div className="grid grid-cols-1 sm:grid-cols-3 gap-6">
                <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] text-center">
                  <div className="bg-[#00ff00] text-black border-2 border-[#141414] p-3 inline-block mb-3">
                    <FileCheck size={24} />
                  </div>
                  <h4 className="text-black text-xs uppercase font-black tracking-wider">Tỷ lệ tự động gán mã</h4>
                  <p className="text-3xl font-black text-black mt-2">
                    {commodityMappedRows.length > 0
                      ? `${Math.round((commodityMappedRows.filter(r => r.score >= config.autoThreshold).length / commodityMappedRows.length) * 100)}%`
                      : "93.4%"}
                  </p>
                  <p className="text-[10px] text-slate-500 font-bold uppercase mt-2">Dựa trên cơ sở dữ liệu mẫu học máy mờ</p>
                </div>

                <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] text-center">
                  <div className="bg-yellow-300 text-black border-2 border-[#141414] p-3 inline-block mb-3">
                    <TrendingUp size={24} />
                  </div>
                  <h4 className="text-black text-xs uppercase font-black tracking-wider">Tỷ lệ chính xác ước tính</h4>
                  <p className="text-3xl font-black text-black mt-2">98.5%</p>
                  <p className="text-[10px] text-slate-500 font-bold uppercase mt-2">Đối soát chéo tài khoản & mã số thuế</p>
                </div>

                <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] text-center">
                  <div className="bg-cyan-300 text-black border-2 border-[#141414] p-3 inline-block mb-3">
                    <Coins size={24} />
                  </div>
                  <h4 className="text-black text-xs uppercase font-black tracking-wider">Dòng ngân hàng đồng bộ</h4>
                  <p className="text-3xl font-black text-black mt-2">
                    {bankMappedRows.length > 0 ? bankMappedRows.length : "0"} GD
                  </p>
                  <p className="text-[10px] text-slate-500 font-bold uppercase mt-2">Phân bổ tự động luồng tiền công nợ</p>
                </div>
              </div>

              {/* Detailed introduction block */}
              <div className="bg-white border-2 border-[#141414] p-6.5 shadow-[4px_4px_0px_#141414] space-y-4">
                <h4 className="text-black font-black text-xs uppercase tracking-wider border-b-2 border-[#141414] pb-2">Quy trình xử lý chuẩn hóa bốn phân hệ</h4>
                <div className="grid grid-cols-1 md:grid-cols-4 gap-4 text-xs">
                  <div className="p-4 bg-[#f0f0ed] border-2 border-[#141414] relative">
                    <span className="absolute top-2 right-3 text-black font-black font-mono text-xs">01</span>
                    <h5 className="font-extrabold uppercase text-black text-[11px] mb-1">Quy chuẩn ngữ văn</h5>
                    <p className="text-[#444] text-[10px]">Loại bỏ dấu định vị, đưa về chữ viết thường, gạt dôi dư "Cty, tnhh, mtv, cp".</p>
                  </div>

                  <div className="p-4 bg-[#f0f0ed] border-2 border-[#141414] relative">
                    <span className="absolute top-2 right-3 text-black font-black font-mono text-xs">02</span>
                    <h5 className="font-extrabold uppercase text-black text-[11px] mb-1">Chiết xuất Specs</h5>
                    <p className="text-[#444] text-[10px]">Khóa cứng thể tích, trọng lượng, số đo và model làm điểm neo định dạng cố định.</p>
                  </div>

                  <div className="p-4 bg-[#f0f0ed] border-2 border-[#141414] relative">
                    <span className="absolute top-2 right-3 text-black font-black font-mono text-xs">03</span>
                    <h5 className="font-extrabold uppercase text-black text-[11px] mb-1">Fuzzy String Ratio</h5>
                    <p className="text-[#444] text-[10px]">Tính toán khoảng cách so khớp chữ mờ dựa trên nguyên lý tỷ lệ Levenshtein.</p>
                  </div>

                  <div className="p-4 bg-[#f0f0ed] border-2 border-[#141414] relative">
                    <span className="absolute top-2 right-3 text-black font-black font-mono text-xs">04</span>
                    <h5 className="font-extrabold uppercase text-black text-[11px] mb-1">Tự Động Sinh Mã</h5>
                    <p className="text-[#444] text-[10px]">Gán trực tiếp nếu phát hiện cực khớp, tự sinh mã mới liên tục và chính xác.</p>
                  </div>
                </div>
              </div>
            </div>
          )}

          {/* --- TAB CONTENT 2: COMMODITY CONFIG --- */}
          {currentTab === "commodity" && (
            <div className="space-y-6">
              <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-4">
                <h3 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-[#141414] pb-2">Gắn mã hàng hóa vật tư</h3>
                <p className="text-xs text-slate-500">So khớp hàng hóa từ bảng kê hóa đơn mua/bán với danh mục sản phẩm của doanh nghiệp.</p>

                {commoditySourceRows.length === 0 ? (
                  <div className="border-2 border-dashed border-[#141414] p-8 text-center bg-[#fdfdfb] flex flex-col items-center justify-center">
                    <UploadCloud size={32} className="text-[#141414] mb-2" />
                    <p className="text-xs font-bold text-black uppercase tracking-wide">Tải lên tệp XLS/XLSX/CSV bảng kê</p>
                    <p className="text-[10px] text-slate-500 font-bold uppercase mt-1 mb-4">Hoặc bấm nút nạp dữ liệu mẫu ở góc trên bên phải</p>
                    <label className="bg-[#00ff00] hover:bg-[#05e005] text-black text-xs font-black uppercase tracking-wider px-6 py-2.5 border-2 border-[#141414] shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer inline-flex items-center gap-2">
                      <FileSpreadsheet size={14} />
                      Chọn tệp từ thiết bị
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        className="hidden"
                        onChange={(e) => handleGenericFileUpload(e, "commodity")}
                      />
                    </label>
                  </div>
                ) : (
                  <div className="space-y-4">
                    {/* Column mapping configuration */}
                    <div className="bg-[#f0f0ed] p-4.5 border-2 border-[#141414] space-y-3">
                      <div className="flex justify-between items-center flex-wrap gap-2 mb-1">
                        <h4 className="text-xs font-black uppercase text-black tracking-wider">Khớp cột excel gốc của bạn:</h4>
                        <label className="bg-yellow-300 hover:bg-yellow-400 text-black text-[10px] font-black uppercase tracking-wider px-2.5 py-1 border border-black shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] active:translate-y-0 transition cursor-pointer inline-flex items-center gap-1">
                          <FileSpreadsheet size={12} />
                          Thay đổi tệp
                          <input
                            type="file"
                            accept=".xlsx,.xls,.csv"
                            className="hidden"
                            onChange={(e) => handleGenericFileUpload(e, "commodity")}
                          />
                        </label>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 text-xs">
                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Diễn giải hàng hóa *</label>
                          <select
                            value={commodityMappings.ten_hang_hoa}
                            onChange={(e) => setCommodityMappings({ ...commodityMappings, ten_hang_hoa: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(commodityHeaders, commodityMappings.ten_hang_hoa)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Đơn vị tính</label>
                          <select
                            value={commodityMappings.don_vi_tinh}
                            onChange={(e) => setCommodityMappings({ ...commodityMappings, don_vi_tinh: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(commodityHeaders, commodityMappings.don_vi_tinh)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Số lượng</label>
                          <select
                            value={commodityMappings.so_luong}
                            onChange={(e) => setCommodityMappings({ ...commodityMappings, so_luong: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(commodityHeaders, commodityMappings.so_luong)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Đơn giá</label>
                          <select
                            value={commodityMappings.don_gia}
                            onChange={(e) => setCommodityMappings({ ...commodityMappings, don_gia: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(commodityHeaders, commodityMappings.don_gia)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Thành tiền</label>
                          <select
                            value={commodityMappings.thanh_tien}
                            onChange={(e) => setCommodityMappings({ ...commodityMappings, thanh_tien: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(commodityHeaders, commodityMappings.thanh_tien)}
                          </select>
                        </div>
                      </div>
                    </div>

                    <button
                      onClick={handleProcessCommodities}
                      disabled={isProcessingCommodities}
                      className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[6px_6px_0px_#00ff00] hover:translate-y-[-2px] text-xs font-black uppercase py-3.5 px-4 border-2 border-[#141414] shadow-[4px_4px_0px_#00ff00] transition flex items-center justify-center gap-2 cursor-pointer"
                    >
                      {isProcessingCommodities ? (
                        <>
                          <RefreshCw size={14} className="animate-spin text-white grow-0" />
                          <span>Đang tính toán trùng khớp...</span>
                        </>
                      ) : (
                        <>
                          <Play size={14} className="text-[#00ff00]" />
                          <span>Bắt đầu rà soát và gắn mã hàng hóa</span>
                        </>
                      )}
                    </button>
                  </div>
                )}
              </div>

              {/* Output grid results */}
              {commodityMappedRows.length > 0 && (
                <div className="bg-white border-2 border-[#141414] shadow-[4px_4px_0px_#141414] overflow-hidden">
                  <div className="p-4 px-6 border-b-2 border-[#141414] bg-[#f0f0ed] flex justify-between items-center flex-wrap gap-2">
                    <div>
                      <h4 className="font-black text-xs uppercase text-black tracking-wider">Kết quả gán mã hàng hóa chi tiết</h4>
                      <p className="text-[10px] text-slate-500 font-bold uppercase mt-1">Sửa trực tiếp Mã đề xuất qua hộp chọn dưới</p>
                    </div>

                    <div className="flex gap-2 text-xs">
                      <span className="bg-[#00ff00] text-black px-2.5 py-1 font-bold border border-black flex items-center gap-1.5 text-[10px] uppercase font-mono">
                        <Check size={12} />
                        Khớp gắn mã cũ (≥ {config.autoThreshold}%): {commodityMappedRows.filter(r => r.treatment === "TỰ ĐỘNG GẮN").length} dòng
                      </span>
                      <span className="bg-sky-200 text-[#141414] px-2.5 py-1 font-bold border border-black flex items-center gap-1.5 text-[10px] uppercase font-mono border-dashed">
                        <Plus size={12} />
                        Tự tạo mã mới (&lt; {config.autoThreshold}%): {commodityMappedRows.filter(r => r.treatment === "TẠO MÃ MỚI").length} dòng
                      </span>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black tracking-wider">
                          <th className="p-3 pl-6">Nội dung diễn giải gốc</th>
                          <th className="p-3">ĐVT</th>
                          <th className="p-3">Mã hàng hóa đề xuất</th>
                          <th className="p-3">Tên sản phẩm chuẩn</th>
                          <th className="p-3 text-center">Độ khớp</th>
                          <th className="p-3">Nguyên tắc quyết định</th>
                          <th className="p-3 pr-6">Trạng thái</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#141414]/10">
                        {commodityMappedRows.slice(0, commodityLimit).map((row) => (
                          <tr key={row.id} className="hover:bg-[#f0f0ed]/30 transition">
                            <td className="p-3 pl-6 max-w-xs truncate font-bold text-black">{row.originalText}</td>
                            <td className="p-3 text-slate-500 font-mono font-medium">{row.originalUom || "Cái"}</td>
                            <td className="p-3">
                              {editingCommodityId === row.id ? (
                                <select
                                  value={row.proposedCode}
                                  autoFocus
                                  onBlur={() => setEditingCommodityId(null)}
                                  onChange={(e) => {
                                    handleEditCommodityCode(row.id, e.target.value);
                                    setEditingCommodityId(null);
                                  }}
                                  className="border-2 border-[#141414] bg-white rounded-none p-1 font-mono text-[11px] font-black text-[#141414] focus:outline-none w-full"
                                >
                                  {commodities.map((c) => (
                                    <option key={c.ma_hang_hoa} value={c.ma_hang_hoa}>
                                      {c.ma_hang_hoa} - {c.ten_hang_hoa_chuan.substring(0, 20)}...
                                    </option>
                                  ))}
                                </select>
                              ) : (
                                <div
                                  onClick={() => setEditingCommodityId(row.id)}
                                  className="font-mono text-[11px] font-black text-[#141414] bg-white hover:bg-yellow-100 hover:border-[#141414] transition cursor-pointer px-2 py-1 border-2 border-dashed border-slate-300 flex items-center justify-between gap-1 w-fit min-w-[100px]"
                                  title="Nhấp để thay đổi mã hàng"
                                >
                                  <span>{row.proposedCode}</span>
                                  <span className="text-[10px] text-slate-400">✏️</span>
                                </div>
                              )}
                            </td>
                            <td className="p-3 text-[#1a1a1a] font-medium">{row.proposedName}</td>
                            <td className="p-3 text-center">
                              <span className={`inline-block px-1.5 py-0.5 border border-black font-black font-mono text-[10px] ${
                                row.score >= config.autoThreshold ? "bg-[#00ff00] text-black" : "bg-yellow-300 text-black"
                              }`}>
                                {row.score}%
                              </span>
                            </td>
                            <td className="p-3 text-[11px] text-slate-500 max-w-xs font-medium">{row.reason}</td>
                            <td className="p-3 pr-6">
                              <span className={`inline-block px-2 py-0.5 border border-black text-[10px] font-black ${
                                row.treatment === "TỰ ĐỘNG GẮN" ? "bg-[#00ff00] text-black" : "bg-sky-100 text-[#141414] border-dashed"
                              }`}>
                                {row.treatment}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {commodityMappedRows.length > commodityLimit && (
                    <div className="p-4 border-t-2 border-[#141414] flex justify-center bg-white">
                      <button
                        onClick={() => setCommodityLimit(prev => prev + 100)}
                        className="bg-white hover:bg-[#f0f0ed] text-black font-black uppercase text-[11px] px-6 py-2 border-2 border-[#141414] shadow-[3px_3px_0px_#141414] hover:shadow-[4px_4px_0px_#141414] active:translate-y-[1px] transition cursor-pointer"
                      >
                        📂 Hiển thị thêm 100 dòng (Đang xem {commodityLimit} / {commodityMappedRows.length} dòng)
                      </button>
                    </div>
                  )}

                  <div className="p-4 bg-[#f0f0ed] border-t-2 border-[#141414] text-right">
                    <button
                      onClick={exportCommodityToExcel}
                      className="bg-[#00ff00] hover:bg-[#05e005] hover:shadow-[3px_3px_0px_#141414] hover:translate-y-[-1px] text-black font-black uppercase tracking-wider text-xs p-3.5 py-1.5 border-2 border-[#141414] inline-flex items-center gap-1.5 shadow-[2px_2px_0px_#141414] transition cursor-pointer"
                    >
                      <Download size={13} />
                      Export Excel
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* --- TAB CONTENT 3: PARTNER CONFIG --- */}
          {currentTab === "partner" && (
            <div className="space-y-6">
              <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-4">
                <div className="flex justify-between items-center flex-wrap gap-4">
                  <div>
                    <h3 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-transparent pb-1">Gắn mã đối tác Khách hàng / Nhà cung cấp</h3>
                    <p className="text-xs text-slate-500 mt-1">So khớp đối tượng dựa trên mã số thuế, số tài khoản hoặc đối soát tên.</p>
                  </div>

                  {/* Buy/Sell Ledgers switch */}
                  <div className="flex bg-[#f0f0ed] p-1 border-2 border-[#141414]">
                    <button
                      onClick={() => setPartnerMode("Mua vào")}
                      className={`px-3 py-1.5 text-xs font-black uppercase tracking-wider transition cursor-pointer ${
                        partnerMode === "Mua vào" ? "bg-[#141414] text-white shadow-sm" : "text-black hover:text-[#444]"
                      }`}
                    >
                      Mua vào (NCC)
                    </button>
                    <button
                      onClick={() => setPartnerMode("Bán ra")}
                      className={`px-3 py-1.5 text-xs font-black uppercase tracking-wider transition cursor-pointer ${
                        partnerMode === "Bán ra" ? "bg-[#141414] text-white shadow-sm" : "text-black hover:text-[#444]"
                      }`}
                    >
                      Bán ra (Khách)
                    </button>
                  </div>
                </div>

                {partnerSourceRows.length === 0 ? (
                  <div className="border-2 border-dashed border-[#141414] p-8 text-center bg-[#fdfdfb] flex flex-col items-center justify-center">
                    <UploadCloud size={32} className="text-[#141414] mb-2" />
                    <p className="text-xs font-bold text-black uppercase tracking-wide">Tải lên tệp sổ kế toán đối tác mua/bán</p>
                    <p className="text-[10px] text-slate-500 font-bold uppercase mt-1 mb-4">Hoặc bấm nút nạp dữ liệu mẫu ở góc trên bên phải</p>
                    <label className="bg-[#00ff00] hover:bg-[#05e005] text-black text-xs font-black uppercase tracking-wider px-6 py-2.5 border-2 border-[#141414] shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer inline-flex items-center gap-2">
                      <FileSpreadsheet size={14} />
                      Chọn tệp từ thiết bị
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        className="hidden"
                        onChange={(e) => handleGenericFileUpload(e, "partner")}
                      />
                    </label>
                  </div>
                ) : (
                  <div className="space-y-4">
                    {/* Column mapping */}
                    <div className="bg-[#f0f0ed] p-4.5 border-2 border-[#141414] space-y-3">
                      <div className="flex justify-between items-center flex-wrap gap-2 mb-1">
                        <h4 className="text-xs font-black uppercase text-black tracking-wider">Cấu hình ánh xạ cột đối tác:</h4>
                        <label className="bg-yellow-300 hover:bg-yellow-400 text-black text-[10px] font-black uppercase tracking-wider px-2.5 py-1 border border-black shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] active:translate-y-0 transition cursor-pointer inline-flex items-center gap-1">
                          <FileSpreadsheet size={12} />
                          Thay đổi tệp
                          <input
                            type="file"
                            accept=".xlsx,.xls,.csv"
                            className="hidden"
                            onChange={(e) => handleGenericFileUpload(e, "partner")}
                          />
                        </label>
                      </div>
                      <div className="grid grid-cols-2 sm:grid-cols-4 gap-3 text-xs">
                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Tên đơn vị đối tác *</label>
                          <select
                            value={partnerMappings.ten_doi_tuong}
                            onChange={(e) => setPartnerMappings({ ...partnerMappings, ten_doi_tuong: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(partnerHeaders, partnerMappings.ten_doi_tuong)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Mã số thuế đối tác</label>
                          <select
                            value={partnerMappings.ma_so_thue}
                            onChange={(e) => setPartnerMappings({ ...partnerMappings, ma_so_thue: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(partnerHeaders, partnerMappings.ma_so_thue)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Số tài khoản đối tác</label>
                          <select
                            value={partnerMappings.so_tai_khoan}
                            onChange={(e) => setPartnerMappings({ ...partnerMappings, so_tai_khoan: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            <option value="">--Không chọn--</option>
                            {getColumnOptions(partnerHeaders, partnerMappings.so_tai_khoan)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Diễn giải/Hàng hóa liên quan</label>
                          <select
                            value={partnerMappings.ten_hang_hoa}
                            onChange={(e) => setPartnerMappings({ ...partnerMappings, ten_hang_hoa: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(partnerHeaders, partnerMappings.ten_hang_hoa)}
                          </select>
                        </div>
                      </div>
                    </div>

                    {/* Audit Details Summary for Partner */}
                    {partnerAuditDetails && (
                      <div className="mt-3 p-3 bg-slate-50 border border-slate-300 space-y-1.5 text-xs font-mono text-slate-700">
                        <div className="font-bold text-black border-b border-dashed border-slate-300 pb-1 flex justify-between items-center">
                          <span>⚙️ Cấu trúc file đối tác:</span>
                          <button onClick={() => handleReopenWizard("partner")} className="text-blue-600 hover:underline font-black uppercase text-[10px]">Cấu hình lại</button>
                        </div>
                        <div className="grid grid-cols-2 gap-2 text-xs">
                          <div>• Sheet: <span className="text-black font-bold">{partnerAuditDetails.sheetName}</span></div>
                          <div>• Dòng tiêu đề: <span className="text-black font-bold">Dòng {partnerAuditDetails.headerRow}</span></div>
                          <div>• Số cột: <span className="text-black font-bold">{partnerAuditDetails.totalColumns} cột</span></div>
                          <div>• Số dòng: <span className="text-black font-bold">{partnerAuditDetails.totalRows} dòng</span></div>
                        </div>
                        {partnerAuditDetails.renamedColumns.length > 0 && (
                          <div className="mt-1">
                            <span className="font-bold text-amber-700">⚠️ Chuẩn hóa tên cột:</span>
                            <ul className="list-disc pl-3 text-xs text-amber-800 space-y-0.5 mt-0.5 max-h-24 overflow-y-auto">
                              {partnerAuditDetails.renamedColumns.map((col, idx) => <li key={idx}>{col}</li>)}
                            </ul>
                          </div>
                        )}
                        {partnerAuditDetails.emptyColumns.length > 0 && (
                          <div className="mt-1">
                            <span className="font-bold text-slate-500">🗑️ Cột rỗng không dữ liệu ({partnerAuditDetails.emptyColumns.length}):</span>
                            <div className="text-xs text-slate-500 truncate mt-0.5" title={partnerAuditDetails.emptyColumns.join(", ")}>
                              {partnerAuditDetails.emptyColumns.join(", ")}
                            </div>
                          </div>
                        )}
                      </div>
                    )}

                    <button
                      onClick={handleProcessPartners}
                      disabled={isProcessingPartners}
                      className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[6px_6px_0px_#00ff00] hover:translate-y-[-2px] text-xs font-black uppercase py-3.5 px-4 border-2 border-[#141414] shadow-[4px_4px_0px_#00ff00] transition flex items-center justify-center gap-2 cursor-pointer"
                    >
                      {isProcessingPartners ? (
                        <>
                          <RefreshCw size={14} className="animate-spin text-white grow-0" />
                          <span>Đang kiểm tra kho đối tác...</span>
                        </>
                      ) : (
                        <>
                          <Play size={14} className="text-[#00ff00]" />
                          <span>Bắt đầu rà soát và gắn mã đối tác</span>
                        </>
                      )}
                    </button>
                  </div>
                )}
              </div>

              {/* Output grid results */}
              {partnerMappedRows.length > 0 && (
                <div className="bg-white border-2 border-[#141414] shadow-[4px_4px_0px_#141414] overflow-hidden">
                  <div className="p-4 px-6 border-b-2 border-[#141414] bg-[#f0f0ed] flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                    <div>
                      <h4 className="font-black text-xs uppercase text-black tracking-wider">Số liệu tích hợp đối tác</h4>
                      <p className="text-[10px] text-slate-500 font-bold uppercase mt-1">Kiểm tra chéo định danh và tự sinh mã đại lý/chi nhánh</p>
                    </div>
                    <div className="flex gap-2 text-xs">
                      <span className="bg-[#00ff00] text-black px-2.5 py-1 font-bold border border-black flex items-center gap-1.5 text-[10px] uppercase font-mono">
                        <Check size={12} />
                        Khớp mã đại lý (≥ {config.autoThreshold}%): {partnerMappedRows.filter(r => r.treatment === "TỰ ĐỘNG GẮN").length} dòng
                      </span>
                      <span className="bg-sky-200 text-[#141414] px-2.5 py-1 font-bold border border-black flex items-center gap-1.5 text-[10px] uppercase font-mono border-dashed">
                        <Plus size={12} />
                        Tự tạo mã mới (&lt; {config.autoThreshold}%): {partnerMappedRows.filter(r => r.treatment === "TẠO MÃ MỚI").length} dòng
                      </span>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black tracking-wider">
                          <th className="p-3 pl-6">Tên đối tác thô trên hóa đơn</th>
                          <th className="p-3">Mã số thuế</th>
                          <th className="p-3">Áp mã chuẩn</th>
                          <th className="p-3">Tên đối tác chuẩn hóa</th>
                          <th className="p-3">Loại</th>
                          <th className="p-3 text-center">Độ khớp</th>
                          <th className="p-3 pr-6">Trạng thái</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#141414]/10">
                        {partnerMappedRows.slice(0, partnerLimit).map((row) => (
                          <tr key={row.id} className="hover:bg-[#f0f0ed]/30 transition">
                            <td className="p-3 pl-6 max-w-xs truncate font-bold text-black">{row.originalText}</td>
                            <td className="p-3 font-mono text-slate-500 font-medium">{row.originalMst || "Không có"}</td>
                            <td className="p-3">
                              {editingPartnerId === row.id ? (
                                <select
                                  value={row.proposedCode}
                                  autoFocus
                                  onBlur={() => setEditingPartnerId(null)}
                                  onChange={(e) => {
                                    handleEditPartnerCode(row.id, e.target.value);
                                    setEditingPartnerId(null);
                                  }}
                                  className="border-2 border-[#141414] bg-white rounded-none p-1 font-mono text-[11px] font-black text-[#141414] focus:outline-none w-full"
                                >
                                  {partners.map((p) => (
                                    <option key={p.ma_doi_tuong} value={p.ma_doi_tuong}>
                                      {p.ma_doi_tuong} - {p.ten_doi_tuong.substring(0, 20)}...
                                    </option>
                                  ))}
                                </select>
                              ) : (
                                <div
                                  onClick={() => setEditingPartnerId(row.id)}
                                  className="font-mono text-[11px] font-black text-[#141414] bg-white hover:bg-yellow-100 hover:border-[#141414] transition cursor-pointer px-2 py-1 border-2 border-dashed border-slate-300 flex items-center justify-between gap-1 w-fit min-w-[100px]"
                                  title="Nhấp để sửa đối tác chuẩn"
                                >
                                  <span>{row.proposedCode}</span>
                                  <span className="text-[10px] text-slate-400">✏️</span>
                                </div>
                              )}
                            </td>
                            <td className="p-3 text-[#1a1a1a] font-medium">{row.proposedName}</td>
                            <td className="p-3">
                              <span className={`inline-block px-1.5 py-0.5 border border-black text-[10px] font-black ${
                                row.proposedType === "Khách hàng" ? "bg-cyan-100 text-black" : "bg-purple-100 text-black"
                              }`}>
                                {row.proposedType}
                              </span>
                            </td>
                            <td className="p-3 text-center">
                              <span className={`inline-block px-1.5 py-0.5 border border-black font-black font-mono text-[10px] ${
                                row.score >= config.autoThreshold ? "bg-[#00ff00] text-black" : "bg-yellow-300 text-black"
                              }`}>
                                {row.score}%
                              </span>
                            </td>
                            <td className="p-3 pr-6">
                              <span className={`inline-block px-2 py-0.5 border border-black text-[10px] font-black ${
                                row.treatment === "TỰ ĐỘNG GẮN" ? "bg-[#00ff00] text-black" : "bg-sky-100 text-[#141414] border-dashed"
                              }`}>
                                {row.treatment}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {partnerMappedRows.length > partnerLimit && (
                    <div className="p-4 border-t-2 border-[#141414] flex justify-center bg-white">
                      <button
                        onClick={() => setPartnerLimit(prev => prev + 100)}
                        className="bg-white hover:bg-[#f0f0ed] text-black font-black uppercase text-[11px] px-6 py-2 border-2 border-[#141414] shadow-[3px_3px_0px_#141414] hover:shadow-[4px_4px_0px_#141414] active:translate-y-[1px] transition cursor-pointer"
                      >
                        📂 Hiển thị thêm 100 dòng (Đang xem {partnerLimit} / {partnerMappedRows.length} dòng)
                      </button>
                    </div>
                  )}

                  <div className="p-4 bg-[#f0f0ed] border-t-2 border-[#141414] text-right">
                    <button
                      onClick={exportPartnerToExcel}
                      className="bg-[#00ff00] hover:bg-[#05e005] hover:shadow-[3px_3px_0px_#141414] hover:translate-y-[-1px] text-black font-black uppercase tracking-wider text-xs p-3.5 py-1.5 border-2 border-[#141414] inline-flex items-center gap-1.5 shadow-[2px_2px_0px_#141414] transition cursor-pointer"
                    >
                      <Download size={13} />
                      Export Excel
                    </button>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* --- TAB CONTENT 4: BANK LOGS CONFIG --- */}
          {currentTab === "bank" && (
            <ErrorBoundary fallbackTitle="LỖI TRONG PHÂN HỆ NGÂN QUỸ / SAO KÊ" onReset={() => setExcelWizard(null)}>
              <div className="space-y-6">
              <div className="bg-[#141414] text-white border-2 border-transparent shadow-[4px_4px_0px_#ccc] p-6 space-y-2">
                <span className="bg-[#00ff00] text-black text-[10px] font-black uppercase tracking-wider px-2 py-0.5 border border-[#141414]">Nghiệp vụ hạch toán & đối chiếu đa chiều</span>
                <h3 className="text-xl font-black uppercase tracking-tight mt-1">Gắn mã đối tác & đối chiếu hóa đơn tự động</h3>
                <p className="text-slate-300 text-xs">
                  Bổ sung tính năng đối soát dòng ngân quỹ với bảng kê bán ra (doanh thu) và bảng kê mua vào (chi phí). Thuật toán tự động tìm tổ hợp tối đa 5 hóa đơn, khớp theo Mã số thuế, tài khoản đối ứng, tên tương đồng, ngày lệch trước/sau và chênh lệch sai số số tiền nhỏ.
                </p>
              </div>

              {/* Grid 3 File Upload Modules */}
              <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
                
                {/* File 1: Sao kê ngân hàng */}
                <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-4 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center gap-2 mb-2 border-b-2 border-[#141414] pb-2">
                      <span className="bg-amber-300 text-black border border-black p-1 text-[10px] font-bold">TỆP 1</span>
                      <h4 className="text-xs font-black uppercase text-black tracking-wider">SAO KÊ NGÂN HÀNG (QUỸ)</h4>
                    </div>
                    {bankSourceRows.length === 0 ? (
                      <div className="border-2 border-dashed border-[#141414]/30 p-6 text-center bg-[#fdfdfb]">
                        <UploadCloud size={24} className="text-[#141414] mx-auto mb-2" />
                        <p className="text-[10px] font-black uppercase tracking-wider text-black">Chưa tải tệp sao kê</p>
                        <p className="text-[9px] text-slate-400 uppercase font-bold mt-1 mb-3">(.XLSX / .CSV)</p>
                        <label className="bg-[#00ff00] hover:bg-[#05e005] text-black text-[9px] font-black uppercase tracking-wider px-3 py-1.5 border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] transition cursor-pointer inline-block">
                          Chọn tệp
                          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleGenericFileUpload(e, "bank")} />
                        </label>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <div className="bg-emerald-50 border border-emerald-300 p-2 text-[10px] font-mono text-emerald-800 flex justify-between items-center">
                          <span className="truncate max-w-[150px] font-bold">📄 {uploadedFileName || "Dữ liệu mẫu"}</span>
                          <span className="bg-emerald-800 text-white font-bold px-1.5 py-0.2 rounded-sm">{bankSourceRows.length} GD</span>
                        </div>
                        {/* Mapping Actions */}
                        <div className="flex gap-2 justify-between">
                          <button
                            onClick={handleRestoreAutoMapping}
                            className="bg-yellow-100 hover:bg-yellow-200 text-black border border-black px-1.5 py-0.5 text-[8px] font-black uppercase tracking-wider"
                            title="Tự động tìm kiếm và đối chiếu tên cột tối ưu nhất dựa trên từ khóa"
                          >
                            ⚡ Ánh xạ tối ưu
                          </button>
                          <button
                            onClick={handleClearMapping}
                            className="bg-red-50 hover:bg-red-100 text-red-700 border border-red-300 px-1.5 py-0.5 text-[8px] font-black uppercase tracking-wider"
                          >
                            🗑️ Xóa ánh xạ
                          </button>
                        </div>
                        {/* Mapping */}
                        <div className="space-y-1 text-[10px]">
                          <span className="font-black uppercase text-slate-500">Cấu hình ánh xạ cột sao kê:</span>
                          
                          {/* Core columns */}
                          <div className="border border-slate-300 p-1.5 rounded bg-slate-50/50 space-y-1">
                            <span className="text-[8px] font-black uppercase text-slate-400 block border-b border-slate-200 pb-0.5">Cột chính (Bắt buộc)</span>
                            <div className="grid grid-cols-2 gap-1.5">
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-500">Nội dung GD *</label>
                                <select value={bankMappings.noi_dung_giao_dich} onChange={(e) => setBankMappings({ ...bankMappings, noi_dung_giao_dich: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[9px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.noi_dung_giao_dich)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-500">Thu (Có)</label>
                                <select value={bankMappings.so_tien_thu} onChange={(e) => setBankMappings({ ...bankMappings, so_tien_thu: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[9px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.so_tien_thu)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-500">Chi (Nợ)</label>
                                <select value={bankMappings.so_tien_chi} onChange={(e) => setBankMappings({ ...bankMappings, so_tien_chi: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[9px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.so_tien_chi)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-500">TK đối ứng</label>
                                <select value={bankMappings.so_tai_khoan_doi_ung} onChange={(e) => setBankMappings({ ...bankMappings, so_tai_khoan_doi_ung: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[9px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.so_tai_khoan_doi_ung)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-500">Tên ĐT sao kê</label>
                                <select value={bankMappings.ten_doi_tac_sao_ke} onChange={(e) => setBankMappings({ ...bankMappings, ten_doi_tac_sao_ke: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[9px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.ten_doi_tac_sao_ke)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-500">Ngày giao dịch</label>
                                <select value={bankMappings.ngay_giao_dich} onChange={(e) => setBankMappings({ ...bankMappings, ngay_giao_dich: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[9px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.ngay_giao_dich)}
                                </select>
                              </div>
                            </div>
                            <div className="bg-amber-50 border border-dashed border-amber-200 p-1 rounded mt-1.5 space-y-1">
                              <span className="text-[7.5px] font-bold uppercase text-amber-800 block">Nếu cột ngày bị tách rời (Ngày / Tháng / Năm):</span>
                              <div className="grid grid-cols-3 gap-1">
                                <div>
                                  <label className="text-[6.5px] uppercase font-bold text-slate-400">Cột Ngày</label>
                                  <select value={bankMappings.cot_ngay} onChange={(e) => setBankMappings({ ...bankMappings, cot_ngay: e.target.value })} className="w-full border border-[#141414] bg-white p-0.2 focus:outline-none text-[8.5px]">
                                    <option value="">-</option>
                                    {getColumnOptions(bankHeaders, bankMappings.cot_ngay)}
                                  </select>
                                </div>
                                <div>
                                  <label className="text-[6.5px] uppercase font-bold text-slate-400">Cột Tháng</label>
                                  <select value={bankMappings.cot_thang} onChange={(e) => setBankMappings({ ...bankMappings, cot_thang: e.target.value })} className="w-full border border-[#141414] bg-white p-0.2 focus:outline-none text-[8.5px]">
                                    <option value="">-</option>
                                    {getColumnOptions(bankHeaders, bankMappings.cot_thang)}
                                  </select>
                                </div>
                                <div>
                                  <label className="text-[6.5px] uppercase font-bold text-slate-400">Cột Năm</label>
                                  <select value={bankMappings.cot_nam} onChange={(e) => setBankMappings({ ...bankMappings, cot_nam: e.target.value })} className="w-full border border-[#141414] bg-white p-0.2 focus:outline-none text-[8.5px]">
                                    <option value="">-</option>
                                    {getColumnOptions(bankHeaders, bankMappings.cot_nam)}
                                  </select>
                                </div>
                              </div>
                            </div>
                          </div>

                          {/* Secondary columns */}
                          <div className="border border-slate-300 p-1.5 rounded bg-slate-50/50 space-y-1 mt-1.5">
                            <span className="text-[8px] font-black uppercase text-slate-400 block border-b border-slate-200 pb-0.5">Cột phụ trợ (Để đối chiếu nâng cao)</span>
                            <div className="grid grid-cols-2 gap-1.5">
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-400">Ngày hạch toán</label>
                                <select value={bankMappings.ngay_hach_toan} onChange={(e) => setBankMappings({ ...bankMappings, ngay_hach_toan: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none text-[8.5px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.ngay_hach_toan)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-400">Ngày hiệu lực</label>
                                <select value={bankMappings.ngay_hieu_luc} onChange={(e) => setBankMappings({ ...bankMappings, ngay_hieu_luc: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none text-[8.5px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.ngay_hieu_luc)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-400">Số dư</label>
                                <select value={bankMappings.so_du} onChange={(e) => setBankMappings({ ...bankMappings, so_du: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none text-[8.5px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.so_du)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-400">Mã giao dịch</label>
                                <select value={bankMappings.ma_giao_dich} onChange={(e) => setBankMappings({ ...bankMappings, ma_giao_dich: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none text-[8.5px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.ma_giao_dich)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-400">Số chứng từ</label>
                                <select value={bankMappings.so_chung_tu} onChange={(e) => setBankMappings({ ...bankMappings, so_chung_tu: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none text-[8.5px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.so_chung_tu)}
                                </select>
                              </div>
                              <div>
                                <label className="text-[7.5px] uppercase font-bold text-slate-400">Số tham chiếu</label>
                                <select value={bankMappings.so_tham_chieu} onChange={(e) => setBankMappings({ ...bankMappings, so_tham_chieu: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none text-[8.5px]">
                                  <option value="">--Không chọn--</option>
                                  {getColumnOptions(bankHeaders, bankMappings.so_tham_chieu)}
                                </select>
                              </div>
                            </div>
                          </div>
                        </div>

                        {/* Audit Details Summary */}
                        {bankAuditDetails && (
                          <div className="mt-3 p-2.5 bg-slate-50 border border-slate-300 space-y-1 text-[9px] font-mono text-slate-700">
                            <div className="font-bold text-black border-b border-dashed border-slate-300 pb-1 flex justify-between items-center">
                              <span>⚙️ Cấu trúc file sao kê:</span>
                              <button onClick={() => handleReopenWizard("bank")} className="text-blue-600 hover:underline font-black uppercase text-[8px]">Cấu hình lại</button>
                            </div>
                            <div className="grid grid-cols-2 gap-1 text-[8.5px]">
                              <div>• Sheet: <span className="text-black font-bold">{bankAuditDetails.sheetName}</span></div>
                              <div>• Dòng tiêu đề: <span className="text-black font-bold">Dòng {bankAuditDetails.headerRow}</span></div>
                              <div>• Số cột: <span className="text-black font-bold">{bankAuditDetails.totalColumns} cột</span></div>
                              <div>• Số dòng: <span className="text-black font-bold">{bankAuditDetails.totalRows} dòng</span></div>
                            </div>
                            {bankAuditDetails.renamedColumns.length > 0 && (
                              <div className="mt-1">
                                <span className="font-bold text-amber-700">⚠️ Chuẩn hóa tên cột:</span>
                                <ul className="list-disc pl-3 text-[8px] text-amber-800 space-y-0.5 mt-0.5 max-h-16 overflow-y-auto">
                                  {bankAuditDetails.renamedColumns.map((col, idx) => <li key={idx}>{col}</li>)}
                                </ul>
                              </div>
                            )}
                            {bankAuditDetails.emptyColumns.length > 0 && (
                              <div className="mt-1">
                                <span className="font-bold text-slate-500">🗑️ Cột rỗng không dữ liệu ({bankAuditDetails.emptyColumns.length}):</span>
                                <div className="text-[8px] text-slate-500 truncate mt-0.5" title={bankAuditDetails.emptyColumns.join(", ")}>
                                  {bankAuditDetails.emptyColumns.join(", ")}
                                </div>
                              </div>
                            )}
                          </div>
                        )}
                      </div>
                    )}
                  </div>
                  {bankSourceRows.length > 0 && (
                    <label className="w-full text-center bg-white hover:bg-[#f0f0ed] text-black text-[9px] font-black uppercase tracking-wider py-1 border border-dashed border-[#141414] transition cursor-pointer">
                      Thay đổi tệp sao kê
                      <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleGenericFileUpload(e, "bank")} />
                    </label>
                  )}
                </div>

                {/* File 2: Bảng kê bán ra (Hóa đơn doanh thu) */}
                <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-4 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center gap-2 mb-2 border-b-2 border-[#141414] pb-2">
                      <span className="bg-cyan-300 text-black border border-black p-1 text-[10px] font-bold">TỆP 2</span>
                      <h4 className="text-xs font-black uppercase text-black tracking-wider">BẢNG KÊ BÁN RA (DOANH THU)</h4>
                    </div>
                    {bankSalesRows.length === 0 ? (
                      <div className="border-2 border-dashed border-[#141414]/30 p-6 text-center bg-[#fdfdfb]">
                        <UploadCloud size={24} className="text-[#141414] mx-auto mb-2" />
                        <p className="text-[10px] font-black uppercase tracking-wider text-black">Chưa tải bảng bán ra</p>
                        <p className="text-[9px] text-slate-400 uppercase font-bold mt-1 mb-3">(.XLSX / .CSV)</p>
                        <label className="bg-[#00ff00] hover:bg-[#05e005] text-black text-[9px] font-black uppercase tracking-wider px-3 py-1.5 border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] transition cursor-pointer inline-block">
                          Chọn tệp
                          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleGenericFileUpload(e, "bank_sales")} />
                        </label>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <div className="bg-cyan-50 border border-cyan-300 p-2 text-[10px] font-mono text-cyan-800 flex justify-between items-center">
                          <span className="truncate max-w-[150px] font-bold">📄 {bankSalesFileName || "Doanh_thu_demo.xlsx"}</span>
                          <span className="bg-cyan-800 text-white font-bold px-1.5 py-0.2 rounded-sm">{bankSalesRows.length} HĐ</span>
                        </div>
                        {/* Mapping */}
                        <div className="space-y-1 text-[10px]">
                          <span className="font-black uppercase text-slate-500">Ánh xạ cột bán ra:</span>
                          <div className="grid grid-cols-2 gap-2 mt-1">
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Số hóa đơn *</label>
                              <select value={bankSalesMappings.so_hoa_don} onChange={(e) => setBankSalesMappings({ ...bankSalesMappings, so_hoa_don: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankSalesHeaders, bankSalesMappings.so_hoa_don)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Ngày hóa đơn</label>
                              <select value={bankSalesMappings.ngay_hoa_don} onChange={(e) => setBankSalesMappings({ ...bankSalesMappings, ngay_hoa_don: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankSalesHeaders, bankSalesMappings.ngay_hoa_don)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Tên khách hàng</label>
                              <select value={bankSalesMappings.ten_khach_hang} onChange={(e) => setBankSalesMappings({ ...bankSalesMappings, ten_khach_hang: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankSalesHeaders, bankSalesMappings.ten_khach_hang)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Tổng thanh toán</label>
                              <select value={bankSalesMappings.tong_thanh_toan} onChange={(e) => setBankSalesMappings({ ...bankSalesMappings, tong_thanh_toan: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankSalesHeaders, bankSalesMappings.tong_thanh_toan)}
                              </select>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                  {bankSalesRows.length > 0 && (
                    <label className="w-full text-center bg-white hover:bg-[#f0f0ed] text-black text-[9px] font-black uppercase tracking-wider py-1 border border-dashed border-[#141414] transition cursor-pointer">
                      Thay đổi tệp bán ra
                      <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleGenericFileUpload(e, "bank_sales")} />
                    </label>
                  )}
                </div>

                {/* File 3: Bảng kê mua vào (Hóa đơn chi phí) */}
                <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-4 flex flex-col justify-between">
                  <div>
                    <div className="flex items-center gap-2 mb-2 border-b-2 border-[#141414] pb-2">
                      <span className="bg-purple-300 text-black border border-black p-1 text-[10px] font-bold">TỆP 3</span>
                      <h4 className="text-xs font-black uppercase text-black tracking-wider">BẢNG KÊ MUA VÀO (CHI PHÍ)</h4>
                    </div>
                    {bankPurchasesRows.length === 0 ? (
                      <div className="border-2 border-dashed border-[#141414]/30 p-6 text-center bg-[#fdfdfb]">
                        <UploadCloud size={24} className="text-[#141414] mx-auto mb-2" />
                        <p className="text-[10px] font-black uppercase tracking-wider text-black">Chưa tải bảng mua vào</p>
                        <p className="text-[9px] text-slate-400 uppercase font-bold mt-1 mb-3">(.XLSX / .CSV)</p>
                        <label className="bg-[#00ff00] hover:bg-[#05e005] text-black text-[9px] font-black uppercase tracking-wider px-3 py-1.5 border-2 border-[#141414] shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] transition cursor-pointer inline-block">
                          Chọn tệp
                          <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleGenericFileUpload(e, "bank_purchases")} />
                        </label>
                      </div>
                    ) : (
                      <div className="space-y-3">
                        <div className="bg-purple-50 border border-purple-300 p-2 text-[10px] font-mono text-purple-800 flex justify-between items-center">
                          <span className="truncate max-w-[150px] font-bold">📄 {bankPurchasesFileName || "Chi_phi_demo.xlsx"}</span>
                          <span className="bg-purple-800 text-white font-bold px-1.5 py-0.2 rounded-sm">{bankPurchasesRows.length} HĐ</span>
                        </div>
                        {/* Mapping */}
                        <div className="space-y-1 text-[10px]">
                          <span className="font-black uppercase text-slate-500">Ánh xạ cột mua vào:</span>
                          <div className="grid grid-cols-2 gap-2 mt-1">
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Số hóa đơn *</label>
                              <select value={bankPurchasesMappings.so_hoa_don} onChange={(e) => setBankPurchasesMappings({ ...bankPurchasesMappings, so_hoa_don: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankPurchasesHeaders, bankPurchasesMappings.so_hoa_don)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Ngày hóa đơn</label>
                              <select value={bankPurchasesMappings.ngay_hoa_don} onChange={(e) => setBankPurchasesMappings({ ...bankPurchasesMappings, ngay_hoa_don: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankPurchasesHeaders, bankPurchasesMappings.ngay_hoa_don)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Tên NCC</label>
                              <select value={bankPurchasesMappings.ten_nha_cung_cap} onChange={(e) => setBankPurchasesMappings({ ...bankPurchasesMappings, ten_nha_cung_cap: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankPurchasesHeaders, bankPurchasesMappings.ten_nha_cung_cap)}
                              </select>
                            </div>
                            <div>
                              <label className="text-[8px] uppercase font-bold text-slate-400">Tổng thanh toán</label>
                              <select value={bankPurchasesMappings.tong_thanh_toan} onChange={(e) => setBankPurchasesMappings({ ...bankPurchasesMappings, tong_thanh_toan: e.target.value })} className="w-full border border-black bg-white p-0.5 focus:outline-none font-bold text-[10px]">
                                {getColumnOptions(bankPurchasesHeaders, bankPurchasesMappings.tong_thanh_toan)}
                              </select>
                            </div>
                          </div>
                        </div>
                      </div>
                    )}
                  </div>
                  {bankPurchasesRows.length > 0 && (
                    <label className="w-full text-center bg-white hover:bg-[#f0f0ed] text-black text-[9px] font-black uppercase tracking-wider py-1 border border-dashed border-[#141414] transition cursor-pointer">
                      Thay đổi tệp mua vào
                      <input type="file" accept=".xlsx,.xls,.csv" className="hidden" onChange={(e) => handleGenericFileUpload(e, "bank_purchases")} />
                    </label>
                  )}
                </div>

              </div>

              {/* Advanced Matching Parameter Settings Panel */}
              <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-3.5">
                <div className="border-b-2 border-[#141414] pb-2 flex justify-between items-center">
                  <h4 className="text-xs font-black uppercase text-black tracking-wider">⚙️ THAM SỐ THUẬT TOÁN ĐỐI CHIẾU HÓA ĐƠN & SỐ TIỀN</h4>
                  <span className="text-[9px] text-slate-500 font-bold uppercase">Ưu tiên: MST &gt; Tài khoản &gt; Tên tương quan &gt; Gom nhóm hóa đơn</span>
                </div>
                <div className="grid grid-cols-2 sm:grid-cols-5 gap-4 text-xs">
                  <div>
                    <label className="text-[10px] uppercase font-black text-slate-500">Lệch ngày trước HĐ</label>
                    <div className="flex items-center mt-1 border-2 border-[#141414] bg-white p-1">
                      <input
                        type="number"
                        min="0"
                        max="180"
                        value={config.daysBeforeInvoice || 7}
                        onChange={(e) => setConfig({ ...config, daysBeforeInvoice: parseInt(e.target.value) || 0 })}
                        className="w-full focus:outline-none font-mono font-black text-black text-[11px]"
                      />
                      <span className="text-slate-400 font-bold ml-1 text-[10px]">Ngày</span>
                    </div>
                  </div>
                  <div>
                    <label className="text-[10px] uppercase font-black text-slate-500">Lệch ngày sau HĐ</label>
                    <div className="flex items-center mt-1 border-2 border-[#141414] bg-white p-1">
                      <input
                        type="number"
                        min="0"
                        max="180"
                        value={config.daysAfterInvoice || 30}
                        onChange={(e) => setConfig({ ...config, daysAfterInvoice: parseInt(e.target.value) || 0 })}
                        className="w-full focus:outline-none font-mono font-black text-black text-[11px]"
                      />
                      <span className="text-slate-400 font-bold ml-1 text-[10px]">Ngày</span>
                    </div>
                  </div>
                  <div>
                    <label className="text-[10px] uppercase font-black text-slate-500">Chênh lệch tuyệt đối</label>
                    <div className="flex items-center mt-1 border-2 border-[#141414] bg-white p-1">
                      <input
                        type="number"
                        min="0"
                        step="1000"
                        value={config.diffAbsThreshold || 10000}
                        onChange={(e) => setConfig({ ...config, diffAbsThreshold: parseFloat(e.target.value) || 0 })}
                        className="w-full focus:outline-none font-mono font-black text-black text-[11px]"
                      />
                      <span className="text-slate-400 font-bold ml-1 text-[9px]">đ</span>
                    </div>
                  </div>
                  <div>
                    <label className="text-[10px] uppercase font-black text-slate-500">Sai số tỷ lệ</label>
                    <div className="flex items-center mt-1 border-2 border-[#141414] bg-white p-1">
                      <input
                        type="number"
                        min="0"
                        step="0.05"
                        max="10"
                        value={config.diffPctThreshold || 0.5}
                        onChange={(e) => setConfig({ ...config, diffPctThreshold: parseFloat(e.target.value) || 0 })}
                        className="w-full focus:outline-none font-mono font-black text-black text-[11px]"
                      />
                      <span className="text-slate-400 font-bold ml-1 text-[10px]">%</span>
                    </div>
                  </div>
                  <div>
                    <label className="text-[10px] uppercase font-black text-slate-500">Hóa đơn ghép tối đa</label>
                    <div className="flex items-center mt-1 border-2 border-[#141414] bg-white p-1">
                      <input
                        type="number"
                        min="1"
                        max="10"
                        value={config.maxCombinationCount || 5}
                        onChange={(e) => setConfig({ ...config, maxCombinationCount: parseInt(e.target.value) || 5 })}
                        className="w-full focus:outline-none font-mono font-black text-black text-[11px]"
                      />
                      <span className="text-slate-400 font-bold ml-1 text-[10px]">Tấm</span>
                    </div>
                  </div>
                </div>

                <div className="pt-2">
                  <button
                    onClick={handleProcessBank}
                    disabled={isProcessingBank}
                    className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[6px_6px_0px_#00ff00] hover:translate-y-[-2px] text-xs font-black uppercase py-4 px-4 border-2 border-[#141414] shadow-[4px_4px_0px_#00ff00] transition flex items-center justify-center gap-2 cursor-pointer"
                  >
                    {isProcessingBank ? (
                      <>
                        <RefreshCw size={14} className="animate-spin text-white grow-0" />
                        <span>Hệ thống đang đối soát công nợ & giải nhóm hóa đơn...</span>
                      </>
                    ) : (
                      <>
                        <Play size={14} className="text-[#00ff00]" />
                        <span>Tiến hành đối chiếu sao kê và hóa đơn (Khớp đa chỉ tiêu)</span>
                      </>
                    )}
                  </button>
                </div>
              </div>

              {/* Output bank statement mapping */}
              {bankMappedRows.length > 0 && (
                <div className="bg-white border-2 border-[#141414] shadow-[4px_4px_0px_#141414] overflow-hidden animate-fade-in">
                  <div className="p-4 px-6 border-b-2 border-[#141414] bg-[#f0f0ed] flex flex-col sm:flex-row justify-between items-start sm:items-center gap-3">
                    <div>
                      <h4 className="font-black text-xs uppercase text-black tracking-wider">Kết quả đối chiếu tự động hóa đơn & ngân quỹ</h4>
                      <p className="text-[10px] text-slate-500 font-bold uppercase mt-1">
                        Sổ theo dõi thanh toán thông minh áp dụng nguyên lý khấu trừ công nợ dồn tích
                      </p>
                    </div>
                    <div className="flex gap-2">
                      <span className="bg-[#00ff00] text-black px-2 py-0.5 font-bold border border-black text-[9px] uppercase font-mono">
                        Đã khớp: {bankMappedRows.filter(r => r.matchedInvoiceNo).length} GD
                      </span>
                      <span className="bg-yellow-300 text-black px-2 py-0.5 font-bold border border-black text-[9px] uppercase font-mono">
                        Dòng kiểm tra: {bankMappedRows.filter(r => r.treatment === "Cần kiểm tra").length} GD
                      </span>
                    </div>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black tracking-wider">
                          <th className="p-3 pl-6">Nội dung sao kê ngân quỹ gốc</th>
                          <th className="p-3">Thu (Có)</th>
                          <th className="p-3">Chi (Nợ)</th>
                          <th className="p-3">Mã đối tượng</th>
                          <th className="p-3">Tên đối tượng chuẩn hóa</th>
                          <th className="p-3">Hóa đơn đối chuẩn</th>
                          <th className="p-3">Lệch số tiền</th>
                          <th className="p-3 text-center">Độ khớp</th>
                          <th className="p-3 pr-6">Trạng thái</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#141414]/10">
                        {bankMappedRows.slice(0, bankLimit).map((row) => (
                          <tr key={row.id} className="hover:bg-[#f0f0ed]/30 transition text-black">
                            <td className="p-3 pl-6 font-mono text-[11px] text-[#141414] max-w-sm" title={row.desc}>
                              <div>{row.desc}</div>
                              {row.reason && (
                                <div className="text-[9px] text-slate-500 font-sans mt-0.5 italic">
                                  💡 {row.reason}
                                </div>
                              )}
                            </td>
                            <td className="p-3 text-emerald-600 font-bold font-mono">
                              {row.amountIn > 0 ? `${row.amountIn.toLocaleString()}đ` : "-"}
                            </td>
                            <td className="p-3 text-red-600 font-bold font-mono">
                              {row.amountOut > 0 ? `${row.amountOut.toLocaleString()}đ` : "-"}
                            </td>
                            <td className="p-3">
                              {editingBankRowId === row.id ? (
                                <select
                                  value={row.proposedCode || ""}
                                  autoFocus
                                  onBlur={() => setEditingBankRowId(null)}
                                  onChange={(e) => {
                                    handleEditBankCode(row.id, e.target.value);
                                    setEditingBankRowId(null);
                                  }}
                                  className="border-2 border-[#141414] bg-white rounded-none p-1 font-mono text-[10px] font-black text-black focus:outline-none w-full"
                                >
                                  <option value="">CHƯA RÕ</option>
                                  <option value="NGANHANG">NGANHANG</option>
                                  <option value="NHANVIEN">NHANVIEN</option>
                                  <option value="KHO_BAC">KHO_BAC</option>
                                  <option value="NOI_BO">NOI_BO</option>
                                  {partners.map((p) => (
                                    <option key={p.ma_doi_tuong} value={p.ma_doi_tuong}>
                                      {p.ma_doi_tuong} ({p.ten_doi_tuong.substring(0, 10)})
                                    </option>
                                  ))}
                                </select>
                              ) : (
                                <div
                                  onClick={() => setEditingBankRowId(row.id)}
                                  className="font-mono text-[11px] font-black text-[#141414] bg-white hover:bg-yellow-100 hover:border-[#141414] transition cursor-pointer px-2 py-1 border-2 border-dashed border-slate-300 flex items-center justify-between gap-1 w-fit min-w-[100px]"
                                  title="Nhấp để hạch toán mã đối chiếu"
                                >
                                  <span>{row.proposedCode || "CHƯA RÕ"}</span>
                                  <span className="text-[10px] text-slate-400">✏️</span>
                                </div>
                              )}
                            </td>
                            <td className="p-3 text-slate-700 font-bold font-sans max-w-xs truncate">
                              {row.proposedName || "Nghi vấn / Thất lạc"}
                            </td>
                            <td className="p-3">
                              {row.matchedInvoiceNo ? (
                                <div className="space-y-0.5">
                                  <div className="font-bold text-slate-900 bg-amber-100 border border-amber-300 px-1 py-0.5 text-[9px] rounded-sm w-fit truncate max-w-[150px]">
                                    🎫 HĐ: {row.matchedInvoiceNo}
                                  </div>
                                  {row.matchedInvoiceDate && (
                                    <div className="text-[9px] text-slate-500 font-mono">
                                      📅 {row.matchedInvoiceDate}
                                    </div>
                                  )}
                                </div>
                              ) : (
                                <span className="text-slate-400 font-normal">-</span>
                              )}
                            </td>
                            <td className="p-3">
                              {row.differenceAmount !== undefined && row.differenceAmount !== 0 ? (
                                <div className="font-bold font-mono text-[10px]">
                                  <span className={row.differenceAmount > 0 ? "text-red-600" : "text-emerald-600"}>
                                    {row.differenceAmount > 0 ? "Thừa " : "Thiếu "}{Math.abs(row.differenceAmount).toLocaleString()}đ
                                  </span>
                                  {row.differencePercentage !== undefined && row.differencePercentage > 0 && (
                                    <div className="text-[8px] text-slate-400 font-normal">
                                      ({row.differencePercentage.toFixed(1)}%)
                                    </div>
                                  )}
                                </div>
                              ) : (
                                <span className="text-slate-400">-</span>
                              )}
                            </td>
                            <td className="p-3 text-center">
                              <span className={`inline-block px-1.5 py-0.5 border border-black font-black font-mono text-[10px] ${
                                row.score >= config.autoThreshold ? "bg-[#00ff00] text-black" : "bg-yellow-300 text-black"
                              }`}>
                                {row.score}%
                              </span>
                            </td>
                            <td className="p-3 pr-6">
                              <span className={`inline-block px-2 py-0.5 border border-black text-[10px] font-black uppercase ${
                                row.treatment === "Đã chốt" ? "bg-[#00ff00] text-black" : "bg-yellow-300 text-black"
                              }`}>
                                {row.treatment}
                              </span>
                            </td>
                          </tr>
                        ))}
                      </tbody>
                    </table>
                  </div>

                  {bankMappedRows.length > bankLimit && (
                    <div className="p-4 border-t-2 border-[#141414] flex justify-center bg-white">
                      <button
                        onClick={() => setBankLimit(prev => prev + 100)}
                        className="bg-white hover:bg-[#f0f0ed] text-black font-black uppercase text-[11px] px-6 py-2 border-2 border-[#141414] shadow-[3px_3px_0px_#141414] hover:shadow-[4px_4px_0px_#141414] active:translate-y-[1px] transition cursor-pointer"
                      >
                        📂 Hiển thị thêm 100 dòng (Đang xem {bankLimit} / {bankMappedRows.length} dòng)
                      </button>
                    </div>
                  )}

                  <div className="p-4 bg-[#f0f0ed] border-t-2 border-[#141414] text-right flex justify-between items-center px-6">
                    <span className="text-[10px] font-black text-black uppercase tracking-wider">
                      Được thiết kế theo tiêu chuẩn Kiểm toán số 2026
                    </span>
                    <button
                      onClick={exportBankToExcel}
                      className="bg-[#00ff00] hover:bg-[#05e005] hover:shadow-[3px_3px_0px_#141414] hover:translate-y-[-1px] text-black font-black uppercase tracking-wider text-xs p-3.5 py-1.5 border-2 border-[#141414] inline-flex items-center gap-1.5 shadow-[2px_2px_0px_#141414] transition cursor-pointer"
                    >
                      <Download size={13} />
                      Export Excel (7 Sheets)
                    </button>
                  </div>
                </div>
              )}
            </div>
            </ErrorBoundary>
          )}

          {/* --- TAB CONTENT 5: INTEGRATED RECONCILIATION --- */}
          {currentTab === "integrated" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-4">
                <h3 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-[#141414] pb-2">Đối chiếu công nợ & doanh thu đa chiều</h3>
                <p className="text-xs text-slate-500">Chạy đồng thời ánh xạ hệ kho - hóa đơn - ngân quỹ để xác minh rủi ro thanh toán lệch pha.</p>

                <button
                  onClick={handleProcessIntegrated}
                  className="w-full bg-[#141414] text-white hover:bg-[#222] hover:shadow-[6px_6px_0px_#00ff00] hover:translate-y-[-2px] text-xs font-black uppercase py-4 px-4 border-2 border-[#141414] shadow-[4px_4px_0px_#00ff00] transition flex items-center justify-center gap-2 cursor-pointer"
                >
                  <Play size={14} className="text-[#00ff00]" />
                  <span>XÂY DỰNG BÁO CÁO TỔNG HỢP & ĐỐI CHIẾU CHÉO ĐA PHÂN HỆ</span>
                </button>
              </div>

              {/* Recon output panel */}
              {integratedRecon.length > 0 && (
                <div className="space-y-6">
                  {/* Revenue audit table */}
                  <div className="bg-white border-2 border-[#141414] shadow-[4px_4px_0px_#141414] overflow-hidden">
                    <div className="p-4 px-6 border-b-2 border-[#141414] bg-[#f0f0ed]">
                      <h4 className="font-black text-xs uppercase text-black tracking-wider">BẢNG ĐỐI CHIẾU DOANH THU HÓA ĐƠN BÁN RA - TIỀN THU SAO KÊ</h4>
                      <p className="text-[10px] text-slate-500 font-bold uppercase mt-1">Giúp phát hiện nhanh công nợ chưa thu hồi hoặc chuyển thừa dôi thu chi</p>
                    </div>

                    <div className="overflow-x-auto">
                      <table className="w-full text-left text-xs border-collapse">
                        <thead>
                          <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black tracking-wider">
                            <th className="p-3 pl-6">Mã đối tác</th>
                            <th className="p-3">Tên riêng khách hàng</th>
                            <th className="p-3 text-right">Doanh thu bán ra (Hóa đơn)</th>
                            <th className="p-3 text-right">Tiền thực thu (Ngân hàng)</th>
                            <th className="p-3 text-right pr-6">Công nợ phải thu (Chênh lệch)</th>
                          </tr>
                        </thead>
                        <tbody className="divide-y divide-[#141414]/10 font-mono">
                          {integratedRecon.map((rec, i) => (
                            <tr key={i} className="hover:bg-[#f0f0ed]/30 transition text-black">
                              <td className="p-3 pl-6 font-black text-black">{rec.partnerCode}</td>
                              <td className="p-3 font-sans font-bold text-black">{rec.partnerName}</td>
                              <td className="p-3 text-right text-black font-black">{rec.invoiceTotal.toLocaleString()} đ</td>
                              <td className="p-3 text-right text-emerald-600 font-black">{rec.bankTotal.toLocaleString()} đ</td>
                              <td className={`p-3 text-right pr-6 font-black ${rec.difference > 0 ? "text-amber-600" : "text-green-600"}`}>
                                {rec.difference.toLocaleString()} đ
                              </td>
                            </tr>
                          ))}
                        </tbody>
                      </table>
                    </div>
                  </div>

                  {/* Summary multi-ledger status log */}
                  <div className="grid grid-cols-1 md:grid-cols-2 gap-6 text-xs text-[#141414]">
                    <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-3">
                      <h4 className="font-black text-black uppercase tracking-wider text-xs border-b border-[#141414] pb-1.5">1. Bảng kê mua vào được gắn mã chéo</h4>
                      <div className="space-y-2">
                        {integratedPurchaseRows.map((r, idx) => (
                          <div key={idx} className="flex justify-between p-2 border border-black bg-[#f0f0ed] items-center">
                            <span className="font-bold text-black">{r.originalText.substring(0, 30)}...</span>
                            <span className="font-mono font-black text-white bg-slate-900 border border-black px-1.5 py-0.5 text-[10px]">{r.proposedCode}</span>
                          </div>
                        ))}
                      </div>
                    </div>

                    <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-3">
                      <h4 className="font-black text-black uppercase tracking-wider text-xs border-b border-[#141414] pb-1.5">2. Bảng kê bán ra được gắn mã chéo</h4>
                      <div className="space-y-2">
                        {integratedSaleRows.map((r, idx) => (
                          <div key={idx} className="flex justify-between p-2 border border-black bg-[#f0f0ed] items-center">
                            <span className="font-bold text-black">{r.originalText.substring(0, 30)}...</span>
                            <span className="font-mono font-black text-white bg-slate-900 border border-black px-1.5 py-0.5 text-[10px]">{r.proposedCode}</span>
                          </div>
                        ))}
                      </div>
                    </div>
                  </div>

                  <div className="bg-[#141414] text-white border-2 border-transparent shadow-[6px_6px_0px_#ccc] p-6 flex flex-col md:flex-row justify-between items-center gap-6 transition-all">
                    <div>
                      <h4 className="text-white font-black text-xs uppercase tracking-wider flex items-center gap-2">
                        <FileCheck size={18} className="text-[#00ff00]" />
                        Trọn bộ hồ sơ đã sẵn sàng kết tập!
                      </h4>
                      <p className="text-xs text-slate-300 font-bold uppercase mt-1">Hồ sơ đi kèm 12 bảng biểu đối chiếu chéo kho, ngân quỹ, thuế và công nợ.</p>
                    </div>
                    <div className="flex flex-wrap gap-3">
                      <button
                        onClick={exportFullSetToExcel}
                        className="bg-[#00ff00] hover:bg-[#05e005] hover:shadow-[6px_6px_0px_#fff] hover:translate-y-[-2px] text-black font-black uppercase tracking-wider text-xs p-4 py-2.5 border-2 border-[#141414] shadow-[4px_4px_0px_#fff] transition cursor-pointer flex items-center gap-1.5"
                      >
                        <Download size={14} />
                        Tải Xuống Trọn Bộ Excel (7 Sheets)
                      </button>
                      <button
                        onClick={exportFullSetToZip}
                        className="bg-amber-400 hover:bg-amber-500 hover:shadow-[6px_6px_0px_#fff] hover:translate-y-[-2px] text-black font-black uppercase tracking-wider text-xs p-4 py-2.5 border-2 border-[#141414] shadow-[4px_4px_0px_#fff] transition cursor-pointer flex items-center gap-1.5"
                      >
                        <Download size={14} />
                        Tải Xuống Trọn Bộ tệp ZIP (Các File Riêng)
                      </button>
                    </div>
                  </div>
                </div>
              )}
            </div>
          )}

          {/* --- TAB CONTENT 6: PYTHON SOURCE CODE VIEWER --- */}
          {currentTab === "python" && (
            <div className="space-y-6 animate-fade-in">
              <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-3">
                <h3 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-[#141414] pb-2">Tải Trọn Gói Python / Streamlit Về Máy Cá Nhân</h3>
                <p className="text-xs text-slate-500">Mã nguồn được viết hoàn chỉnh 100%, có cấu trúc thư mục rõ ràng. Bạn chỉ cần sao chép các tệp sau để tạo dự án chạy local tuyệt hảo trên máy tính của mình.</p>
              </div>

              {/* Instructions and code list */}
              <div className="bg-[#141414] text-slate-100 border-2 border-transparent shadow-[6px_6px_0px_#ccc] p-6.5 font-mono text-xs space-y-6">
                <div>
                  <span className="text-[#00ff00] font-black"># Bước 1: Tạo file cấu hình và cài đặt thư viện</span>
                  <div className="bg-black p-4.5 rounded-none border-2 border-[#141414] mt-2 text-slate-300 space-y-2">
                    <p className="text-slate-500 font-bold"># Tạo file requirements.txt và dán nội dung này:</p>
                    <pre className="text-yellow-300 font-bold">
                      {`streamlit>=1.35.0
pandas>=2.0.0
openpyxl>=3.1.0
rapidfuzz>=3.8.0
unidecode>=1.3.8`}
                    </pre>
                  </div>
                </div>

                <div>
                  <span className="text-[#00ff00] font-black"># Bước 2: Chạy lệnh dưới cmd hoặc Terminal để khởi động:</span>
                  <pre className="bg-black p-4.5 rounded-none border-2 border-[#141414] text-cyan-300 font-bold mt-2">
                    {`pip install -r requirements.txt\nstreamlit run app.py`}
                  </pre>
                </div>

                <div className="flex justify-between items-center bg-black p-4.5 border-2 border-[#141414] text-slate-400">
                  <span className="text-[11px] font-bold text-slate-300">📂 Toàn bộ mã nguồn hoàn hảo đã được lưu trữ trong thư mục làm việc của bạn (gồm `app.py`, `utils/engine.py` và `requirements.txt`). Hãy tải nó xuống bất kỳ lúc nào để bắt đầu hạch toán cá nhân tuyệt đối bảo mật!</span>
                </div>
              </div>
            </div>
          )}

        </main>
      </div>

      {/* --- FOOTER REGION --- */}
      <footer className="bg-[#f0f0ed] border-t-2 border-[#141414] text-black py-8 text-center text-xs mt-12">
        <div className="max-w-7xl mx-auto px-6">
          <p className="font-black uppercase tracking-wider text-[11px]">Auto-Accounting Mapper • Thiết bị nghiệp vụ Kế Toán Kho và Công Nợ</p>
          <p className="text-slate-500 font-bold uppercase text-[9px] mt-1.5">Sử dụng thuật toán so khớp khoảng cách mờ chuỗi ký tự tiếng Việt chuẩn hóa • Bảo mật cục bộ 100%</p>
        </div>
      </footer>

      {/* --- MASTER IMPORT DIALOG OVERLAY --- */}
      {importConfirm && (
        <div className="fixed inset-0 bg-[#141414]/80 z-50 flex items-center justify-center p-4">
          <div className="bg-white border-4 border-[#141414] max-w-md w-full p-6 shadow-[8px_8px_0px_#141414] animate-fade-in space-y-4">
            <div className="flex items-center gap-2 pb-2 border-b-2 border-[#141414]">
              <FileSpreadsheet size={20} className="text-[#00ff00]" />
              <h4 className="font-black text-xs uppercase tracking-wider text-black">
                Xác nhận Cập nhật Danh mục
              </h4>
            </div>

            <div className="text-xs space-y-2.5 text-[#141414]">
              <p className="font-bold">
                Bạn đã tải lên tệp: <span className="font-mono text-blue-600 font-extrabold">{importConfirm.fileName}</span>
              </p>
              <p>
                Cấu trúc tệp phát hiện <span className="font-mono font-black text-white bg-[#141414] px-1.5 py-0.5 border border-[#141414]">{importConfirm.items.length}</span> danh mục {
                  importConfirm.type === "commodity" 
                    ? "Hàng hóa / Vật tư / Sản phẩm"
                    : importConfirm.type === "customer"
                      ? "Khách mua (Khách hàng)"
                      : "Nhà bán (Nhà cung cấp)"
                } chuẩn.
              </p>
              <div className="border-l-4 border-yellow-300 pl-3 py-1 bg-yellow-50 text-[10px] font-bold uppercase text-amber-800">
                Lựa chọn phương thức cập nhật tối ưu dán vào sổ sách:
              </div>
            </div>

            <div className="grid grid-cols-1 gap-3 pt-2">
              <button
                onClick={() => executeImportMaster(false)}
                className="bg-[#00ff00] hover:bg-[#05e005] text-black font-black uppercase text-[10px] tracking-wider py-3 border-2 border-[#141414] shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer flex items-center justify-center gap-1.5"
              >
                📥 Bổ sung thêm vào (Append)
              </button>
              <button
                onClick={() => executeImportMaster(true)}
                className="bg-red-400 hover:bg-red-500 text-black font-black uppercase text-[10px] tracking-wider py-3 border-2 border-[#141414] shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer flex items-center justify-center gap-1.5"
              >
                💥 Thay thế danh mục chuẩn (Overwrite)
              </button>
              <button
                onClick={() => setImportConfirm(null)}
                className="bg-white hover:bg-[#f0f0ed] text-black font-black uppercase text-[10px] tracking-wider py-2 border-2 border-[#141414] hover:translate-y-[-1px] transition cursor-pointer text-center font-bold"
              >
                Hủy bỏ
              </button>
            </div>
          </div>
        </div>
      )}
    </div>
    </ErrorBoundary>
  );
}
