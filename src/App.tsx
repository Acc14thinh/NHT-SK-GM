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
  normalizeText
} from "./lib/matchingEngine";
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
    prefixNCC: "NCC"
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
    noi_dung_giao_dich: "noi_dung_giao_dich",
    so_tien_thu: "so_tien_thu",
    so_tien_chi: "so_tien_chi",
    so_tai_khoan_doi_ung: "so_tai_khoan_doi_ung",
    ten_doi_tac_sao_ke: "ten_doi_tac_sao_ke"
  });
  const [bankMappedRows, setBankMappedRows] = useState<BankAnalysisResult[]>([]);
  const [isProcessingBank, setIsProcessingBank] = useState(false);

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

  const handleGenericFileUpload = (event: React.ChangeEvent<HTMLInputElement>, fileType: "commodity" | "partner" | "bank") => {
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
        }
      } catch (err: any) {
        console.error(err);
        triggerToast("Lỗi định dạng cấu trúc tệp! Hãy đảm bảo tệp sạch và đúng cột.", "warning");
      }
    };
    reader.readAsArrayBuffer(file);
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
      const results = bankSourceRows.map((row, index) => {
        const desc = row[bankMappings.noi_dung_giao_dich] || "";
        const amIn = parseFloat(row[bankMappings.so_tien_thu]) || 0;
        const amOut = parseFloat(row[bankMappings.so_tien_chi]) || 0;
        const counterpartAcc = row[bankMappings.so_tai_khoan_doi_ung] || "";
        const counterpartName = row[bankMappings.ten_doi_tac_sao_ke] || "";

        const analysis = analyzeBankTransaction(
          desc,
          amIn,
          amOut,
          counterpartAcc,
          counterpartName,
          partners,
          config.autoThreshold,
          config.autoThreshold
        );

        return {
          ...analysis,
          id: `bank_${index}`,
          date: row.ngay_giao_dich || new Date().toISOString().substring(0, 10),
          notes: "",
          rawRowData: row
        } as BankAnalysisResult;
      });

      setBankMappedRows(results);
      setIsProcessingBank(false);
      triggerToast("Đã xếp lớp tự động sao kê ngân hàng!");
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
    const dataToExport = bankMappedRows.map((row) => {
      const baseRow = { ...(row.rawRowData || {}) };
      baseRow["Mã đối tượng"] = row.proposedCode || "";
      baseRow["Tên đối tượng chuẩn"] = row.proposedName || "Nghi vấn / Thất lạc";
      baseRow["Loại đối tượng"] = row.predictedGroup.toLowerCase().includes("khách hàng") ? "Khách hàng" : row.predictedGroup.toLowerCase().includes("nhà cung cấp") ? "Nhà cung cấp" : "Khác";
      baseRow["Điểm tương thích"] = `${row.score}%`;
      baseRow["Nhóm giao dịch AI dự đoán"] = row.predictedGroup;
      baseRow["Lý do dự đoán"] = row.reason;
      baseRow["Trạng thái xử lý"] = row.treatment === "Đã chốt" ? "Đã chốt" : "Cần kiểm tra";
      baseRow["Ghi chú"] = row.notes || "";
      return baseRow;
    });

    const worksheet = XLSX.utils.json_to_sheet(dataToExport);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Sao_ke_da_gan_ma");

    const baseName = uploadedFileName ? uploadedFileName.replace(/\.[^/.]+$/, "") : "Sao_ke_ngan_hang";
    XLSX.writeFile(workbook, `${baseName}_da_gan_ma.xlsx`);
    triggerToast("Đã xuất khẩu tệp hạch toán sao kê ngân quỹ thành công!");
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
    <div className="min-h-screen bg-[#fdfdfb] text-[#1a1a1a] font-sans flex flex-col selection:bg-[#00ff00] selection:text-black">
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
            <div className="space-y-6">
              <div className="bg-white border-2 border-[#141414] p-6 shadow-[4px_4px_0px_#141414] space-y-4">
                <h3 className="font-black text-xs uppercase tracking-wider text-black border-b-2 border-[#141414] pb-2">Hạch toán sao kê ngân hàng</h3>
                <p className="text-xs text-slate-500">Tự lọc và phân bổ tự động nội dung giao dịch ngân hàng theo nghiệp vụ công nợ.</p>

                {bankSourceRows.length === 0 ? (
                  <div className="border-2 border-dashed border-[#141414] p-8 text-center bg-[#fdfdfb] flex flex-col items-center justify-center">
                    <UploadCloud size={32} className="text-[#141414] mb-2" />
                    <p className="text-xs font-bold text-black uppercase tracking-wide">Tải lên tệp sao kê ngân hàng .XLSX/.CSV</p>
                    <p className="text-[10px] text-slate-500 font-bold uppercase mt-1 mb-4">Hoặc nhấn nút tải dữ liệu demo mẫu</p>
                    <label className="bg-[#00ff00] hover:bg-[#05e005] text-black text-xs font-black uppercase tracking-wider px-6 py-2.5 border-2 border-[#141414] shadow-[4px_4px_0px_#141414] hover:shadow-[6px_6px_0px_#141414] hover:translate-y-[-2px] active:translate-y-0 transition cursor-pointer inline-flex items-center gap-2">
                      <FileSpreadsheet size={14} />
                      Chọn tệp từ thiết bị
                      <input
                        type="file"
                        accept=".xlsx,.xls,.csv"
                        className="hidden"
                        onChange={(e) => handleGenericFileUpload(e, "bank")}
                      />
                    </label>
                  </div>
                ) : (
                  <div className="space-y-4">
                    <div className="flex justify-between items-center bg-[#f0f0ed] p-2 border-2 border-[#141414]">
                      <span className="text-[10px] font-black text-black uppercase font-mono">Tệp hiện tại: {uploadedFileName || "Dữ liệu mẫu"}</span>
                      <label className="bg-yellow-300 hover:bg-yellow-400 text-black text-[10px] font-black uppercase tracking-wider px-2.5 py-1 border border-black shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] active:translate-y-0 transition cursor-pointer inline-flex items-center gap-1">
                        <FileSpreadsheet size={12} />
                        Thay đổi tệp
                        <input
                          type="file"
                          accept=".xlsx,.xls,.csv"
                          className="hidden"
                          onChange={(e) => handleGenericFileUpload(e, "bank")}
                        />
                      </label>
                    </div>

                    {/* Cấu hình ánh xạ cột sao kê ngân quỹ */}
                    <div className="bg-[#f0f0ed] p-4.5 border-2 border-[#141414] space-y-3">
                      <h4 className="text-xs font-black uppercase text-black tracking-wider">Cấu hình ánh xạ cột sao kê:</h4>
                      <div className="grid grid-cols-2 sm:grid-cols-5 gap-3 text-xs">
                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Nội dung giao dịch *</label>
                          <select
                            value={bankMappings.noi_dung_giao_dich}
                            onChange={(e) => setBankMappings({ ...bankMappings, noi_dung_giao_dich: e.target.value })}
                            className="w-full mt-1 border-2 border-[#141414] bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(bankHeaders, bankMappings.noi_dung_giao_dich)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Số tiền thu (Có)</label>
                          <select
                            value={bankMappings.so_tien_thu}
                            onChange={(e) => setBankMappings({ ...bankMappings, so_tien_thu: e.target.value })}
                            className="w-full mt-1 border-[#141414] border-2 bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(bankHeaders, bankMappings.so_tien_thu)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Số tiền chi (Nợ)</label>
                          <select
                            value={bankMappings.so_tien_chi}
                            onChange={(e) => setBankMappings({ ...bankMappings, so_tien_chi: e.target.value })}
                            className="w-full mt-1 border-[#141414] border-2 bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(bankHeaders, bankMappings.so_tien_chi)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Tài khoản đối ứng</label>
                          <select
                            value={bankMappings.so_tai_khoan_doi_ung}
                            onChange={(e) => setBankMappings({ ...bankMappings, so_tai_khoan_doi_ung: e.target.value })}
                            className="w-full mt-1 border-[#141414] border-2 bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(bankHeaders, bankMappings.so_tai_khoan_doi_ung)}
                          </select>
                        </div>

                        <div>
                          <label className="text-[10px] uppercase font-black text-slate-500">Đối tác sao kê</label>
                          <select
                            value={bankMappings.ten_doi_tac_sao_ke}
                            onChange={(e) => setBankMappings({ ...bankMappings, ten_doi_tac_sao_ke: e.target.value })}
                            className="w-full mt-1 border-[#141414] border-2 bg-white p-1.5 focus:outline-none font-bold text-black text-[11px]"
                          >
                            {getColumnOptions(bankHeaders, bankMappings.ten_doi_tac_sao_ke)}
                          </select>
                        </div>
                      </div>
                    </div>
                    <button
                      onClick={handleProcessBank}
                      disabled={isProcessingBank}
                      className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[6px_6px_0px_#00ff00] hover:translate-y-[-2px] text-xs font-black uppercase py-3.5 px-4 border-2 border-[#141414] shadow-[4px_4px_0px_#00ff00] transition flex items-center justify-center gap-2 cursor-pointer"
                    >
                      {isProcessingBank ? (
                        <>
                          <RefreshCw size={14} className="animate-spin text-white grow-0" />
                          <span>Đang giải mã và định lớp luồng tiền...</span>
                        </>
                      ) : (
                        <>
                          <Play size={14} className="text-[#00ff00]" />
                          <span>Bắt đầu hạch toán sao kê</span>
                        </>
                      )}
                    </button>
                  </div>
                )}
              </div>

              {/* Output bank statement mapping */}
              {bankMappedRows.length > 0 && (
                <div className="bg-white border-2 border-[#141414] shadow-[4px_4px_0px_#141414] overflow-hidden animate-fade-in">
                  <div className="p-4 px-6 border-b-2 border-[#141414] bg-[#f0f0ed]">
                    <h4 className="font-black text-xs uppercase text-black tracking-wider">Kết quả phân bổ nghiệp vụ ngân quỹ</h4>
                    <p className="text-[10px] text-slate-500 font-bold uppercase mt-1">Các giao dịch dịch vụ phí, nộp ngân sách, chi lương được tách riêng khỏi công nợ mua bán</p>
                  </div>

                  <div className="overflow-x-auto">
                    <table className="w-full text-left text-xs border-collapse">
                      <thead>
                        <tr className="bg-[#f0f0ed] border-b-2 border-[#141414] font-black uppercase text-black tracking-wider">
                          <th className="p-3 pl-6">Nội dung giao dịch ngân hàng gốc</th>
                          <th className="p-3">Thu (Có)</th>
                          <th className="p-3">Chi (Nợ)</th>
                          <th className="p-3">Cột nghiệp vụ dự báo</th>
                          <th className="p-3">Mã đối tác</th>
                          <th className="p-3">Đối tác ngân hàng hợp lịch</th>
                          <th className="p-3 text-center">Độ khớp</th>
                          <th className="p-3 pr-6">Trạng thái</th>
                        </tr>
                      </thead>
                      <tbody className="divide-y divide-[#141414]/10">
                        {bankMappedRows.slice(0, bankLimit).map((row) => (
                          <tr key={row.id} className="hover:bg-[#f0f0ed]/30 transition">
                            <td className="p-3 pl-6 font-mono text-[11px] text-[#141414] max-w-xs">{row.desc}</td>
                            <td className="p-3 text-emerald-600 font-bold font-mono">
                              {row.amountIn > 0 ? `${row.amountIn.toLocaleString()}đ` : "-"}
                            </td>
                            <td className="p-3 text-red-600 font-bold font-mono">
                              {row.amountOut > 0 ? `${row.amountOut.toLocaleString()}đ` : "-"}
                            </td>
                            <td className="p-3">
                              <span className="inline-block px-2 py-0.5 border border-black bg-[#f0f0ed] text-black font-black text-[10px] uppercase font-mono">
                                {row.predictedGroup}
                              </span>
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
                            <td className="p-3 text-slate-600 max-w-xs truncate font-medium">{row.proposedName || "Nghi vấn / Thất lạc"}</td>
                            <td className="p-3 text-center">
                              <span className={`inline-block px-1.5 py-0.5 border border-black font-black font-mono text-[10px] ${
                                row.score >= config.autoThreshold ? "bg-[#00ff00] text-black" : "bg-yellow-300 text-black"
                              }`}>
                                {row.score}%
                              </span>
                            </td>
                            <td className="p-3 pr-6">
                              <span className={`inline-block px-2 py-0.5 border border-black text-[10px] font-black ${
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

                  <div className="p-4 bg-[#f0f0ed] border-t-2 border-[#141414] text-right">
                    <button
                      onClick={exportBankToExcel}
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
  );
}
