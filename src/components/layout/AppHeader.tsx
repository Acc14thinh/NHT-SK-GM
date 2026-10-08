import React from "react";
import {
  Menu,
  FileSpreadsheet,
  Sparkles,
  RotateCcw,
  SlidersHorizontal,
  ChevronRight,
  Database,
  Trash2
} from "lucide-react";
import { NavTabId } from "./AppSidebar";

interface AppHeaderProps {
  currentTab: NavTabId;
  onOpenMobileMenu: () => void;
  onOpenSettings: () => void;
  uploadedFileName: string;
  activeRowsCount: number;
  onLoadDemo: () => void;
  onReset: () => void;
}

const tabMetaMap: Record<NavTabId, { group: string; title: string; subtitle: string }> = {
  dashboard: {
    group: "Tổng quan",
    title: "Dashboard Tổng Quan Hệ Thống",
    subtitle: "Giám sát số liệu, trạng thái bảng kê và tóm tắt tiến trình tự động hóa",
  },
  commodity: {
    group: "Gắn mã dữ liệu",
    title: "Gán Mã Hàng Hóa & Vật Tư Chuẩn",
    subtitle: "So khớp hàng hóa từ bảng kê mua/bán với danh mục sản phẩm doanh nghiệp",
  },
  partner: {
    group: "Gắn mã dữ liệu",
    title: "Gán Mã Đối Tác Khách Hàng / Nhà Cung Cấp",
    subtitle: "Tự động phân giải MST, tên công ty và số tài khoản ngân hàng để gán mã",
  },
  bank: {
    group: "Đối chiếu & Kiểm tra",
    title: "Giao Dịch Sao Kê & Ngân Quỹ Hóa Đơn",
    subtitle: "Phân tích biến động số dư sổ phụ ngân hàng và đối soát hóa đơn 3 bên",
  },
  ecommerce: {
    group: "Đối chiếu & Kiểm tra",
    title: "Đối Chiếu Doanh Thu Sàn TMĐT (Shopee)",
    subtitle: "Gom nhóm SKU đa dòng, xử lý trừ voucher shop 1 lần, đối chiếu 2 chiều với PM bán hàng",
  },
  integrated: {
    group: "Đối chiếu & Kiểm tra",
    title: "Đối Chiếu Công Nợ & Doanh Thu Đa Phân Hệ",
    subtitle: "Chạy đồng thời ánh xạ hệ Kho - Hóa đơn - Ngân quỹ để phát hiện lệch pha thanh toán",
  },
  settings: {
    group: "Hệ thống & Cấu hình",
    title: "Tham Số Thuật Toán & Cơ Sở Dữ Liệu Danh Mục",
    subtitle: "Cấu hình độ khớp yêu cầu, đối chiếu đơn giá, quy tắc sinh mã và import danh mục chuẩn",
  },
  python: {
    group: "Hệ thống & Cấu hình",
    title: "Mã Nguồn Python / Streamlit Cá Nhân",
    subtitle: "Tải toàn bộ mã nguồn trọn gói để chạy offline trên máy trạm nội bộ",
  },
};

export const AppHeader: React.FC<AppHeaderProps> = ({
  currentTab,
  onOpenMobileMenu,
  onOpenSettings,
  uploadedFileName,
  activeRowsCount,
  onLoadDemo,
  onReset,
}) => {
  const currentMeta = tabMetaMap[currentTab] || {
    group: "Hệ thống",
    title: "Bảng Điều Khiển Kế Toán",
    subtitle: "SmartLedger AutoCoder",
  };

  return (
    <header className="bg-white border-b-2 border-[#141414] sticky top-0 z-30 shadow-[0_2px_0_rgba(20,20,20,0.04)]">
      <div className="px-4 sm:px-6 py-3 flex flex-wrap items-center justify-between gap-3">
        {/* Left: Mobile Toggle & Tab Breadcrumb */}
        <div className="flex items-center gap-3 min-w-0">
          {/* Mobile hamburger menu */}
          <button
            onClick={onOpenMobileMenu}
            className="lg:hidden p-2 text-black bg-[#f0f0ed] hover:bg-slate-200 border-2 border-black cursor-pointer shadow-[2px_2px_0px_#141414]"
            title="Mở menu điều hướng"
          >
            <Menu size={18} />
          </button>

          <div className="min-w-0">
            {/* Breadcrumb line */}
            <div className="flex items-center gap-1.5 text-[10px] font-mono text-slate-500 font-bold uppercase truncate">
              <span>SmartLedger</span>
              <ChevronRight size={10} />
              <span>{currentMeta.group}</span>
              <ChevronRight size={10} />
              <span className="text-black font-black">{currentMeta.title}</span>
            </div>

            {/* Title & subtitle */}
            <div className="flex items-center gap-2">
              <h2 className="text-sm sm:text-base font-black uppercase tracking-tight text-black truncate leading-tight">
                {currentMeta.title}
              </h2>
            </div>
          </div>
        </div>

        {/* Right: Quick Action Controls */}
        <div className="flex items-center gap-2 shrink-0">
          {uploadedFileName ? (
            <div className="flex items-center gap-1.5 bg-[#f0f0ed] px-2.5 py-1.5 border-2 border-black text-xs font-mono font-bold text-black shadow-[2px_2px_0px_#141414]">
              <FileSpreadsheet size={13} className="text-green-700" />
              <span className="truncate max-w-[140px] sm:max-w-[200px]" title={uploadedFileName}>
                {uploadedFileName}
              </span>
              <span className="text-[10px] bg-black text-[#00ff00] px-1 py-0.2 ml-1">
                {activeRowsCount}d
              </span>
            </div>
          ) : (
            <button
              onClick={onLoadDemo}
              className="bg-[#00ff00] hover:bg-[#05e005] text-black text-xs font-black uppercase px-3 py-1.5 border-2 border-black shadow-[2px_2px_0px_#141414] hover:shadow-[3px_3px_0px_#141414] hover:translate-y-[-1px] active:translate-y-0 transition cursor-pointer flex items-center gap-1.5"
              title="Nạp nhanh bảng kê hóa đơn mua bán và danh mục mẫu"
            >
              <Sparkles size={13} className="text-black fill-black" />
              <span className="hidden sm:inline">Nạp dữ liệu Demo</span>
              <span className="sm:hidden">Demo</span>
            </button>
          )}

          {/* Quick Settings trigger */}
          <button
            onClick={onOpenSettings}
            className="bg-white hover:bg-[#f0f0ed] text-black text-xs font-bold uppercase px-2.5 py-1.5 border-2 border-black shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] transition cursor-pointer flex items-center gap-1"
            title="Mở bảng cấu hình tham số thuật toán & danh mục"
          >
            <SlidersHorizontal size={13} />
            <span className="hidden md:inline">Tham số</span>
          </button>

          {/* Quick Reset */}
          {uploadedFileName && (
            <button
              onClick={onReset}
              className="bg-red-50 hover:bg-red-100 text-red-700 text-xs font-bold px-2 py-1.5 border-2 border-black shadow-[2px_2px_0px_#141414] hover:translate-y-[-1px] transition cursor-pointer flex items-center gap-1"
              title="Làm mới toàn bộ tiến trình"
            >
              <Trash2 size={13} />
            </button>
          )}
        </div>
      </div>
    </header>
  );
};
