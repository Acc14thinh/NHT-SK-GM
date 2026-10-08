import React from "react";
import {
  SlidersHorizontal,
  Settings,
  Database,
  Users,
  Upload,
  CheckCircle2,
  AlertCircle,
  Coins,
  Tag,
  FileSpreadsheet,
  X
} from "lucide-react";
import { Commodity, Partner, MatchingConfig } from "../../types";

interface AlgorithmSettingsViewProps {
  config: MatchingConfig;
  setConfig: React.Dispatch<React.SetStateAction<MatchingConfig>>;
  commodities: Commodity[];
  partners: Partner[];
  handleImportMasterDirectory: (
    event: React.ChangeEvent<HTMLInputElement>,
    type: "commodity" | "customer" | "supplier"
  ) => void;
  isDrawer?: boolean;
  onCloseDrawer?: () => void;
}

export const AlgorithmSettingsView: React.FC<AlgorithmSettingsViewProps> = ({
  config,
  setConfig,
  commodities,
  partners,
  handleImportMasterDirectory,
  isDrawer = false,
  onCloseDrawer,
}) => {
  const customerCount = partners.filter((p) => p.loai_doi_tuong === "Khách hàng").length;
  const supplierCount = partners.filter((p) => p.loai_doi_tuong === "Nhà cung cấp").length;

  return (
    <div className={`space-y-6 ${isDrawer ? "p-6" : "animate-fade-in"}`}>
      {/* Drawer Header if opened as a slide-over */}
      {isDrawer && (
        <div className="flex items-center justify-between border-b-2 border-black pb-3">
          <div className="flex items-center gap-2">
            <div className="p-1.5 bg-black text-[#00ff00]">
              <SlidersHorizontal size={18} />
            </div>
            <div>
              <h3 className="font-black text-sm uppercase text-black font-sans">
                Tham Số Thuật Toán &amp; Danh Mục
              </h3>
              <p className="text-[10px] text-slate-500 font-mono">
                Cấu hình hệ thống nhận diện tự động
              </p>
            </div>
          </div>

          {onCloseDrawer && (
            <button
              onClick={onCloseDrawer}
              className="p-1.5 bg-slate-100 hover:bg-slate-200 text-black border border-black cursor-pointer"
            >
              <X size={18} />
            </button>
          )}
        </div>
      )}

      {/* Main Grid: Settings & Master Directories */}
      <div className={`grid grid-cols-1 ${isDrawer ? "gap-6" : "lg:grid-cols-2 gap-8"}`}>
        {/* BLOCK 1: THAM SỐ THUẬT TOÁN SO KHỚP */}
        <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-6">
          <div className="flex items-center justify-between border-b-2 border-[#141414] pb-3">
            <h3 className="text-xs font-black text-black tracking-widest uppercase flex items-center gap-2">
              <Settings size={16} className="text-[#141414]" />
              <span>1. Tham Số Ngưỡng So Khớp</span>
            </h3>
            <span className="text-[10px] font-mono text-slate-500 font-bold uppercase">
              Thuật toán tối ưu
            </span>
          </div>

          {/* Slider độ khớp yêu cầu */}
          <div className="space-y-3">
            <div className="flex justify-between items-center">
              <label className="text-xs font-black text-[#141414] uppercase tracking-wider flex items-center gap-1.5">
                <span className="w-2.5 h-2.5 bg-[#00ff00] border border-[#141414] inline-block"></span>
                <span>Độ khớp yêu cầu tối thiểu</span>
              </label>
              <span className="text-xs font-mono font-black bg-black text-[#00ff00] px-2 py-0.5 border border-black shadow-[1px_1px_0px_#141414]">
                {config.autoThreshold}%
              </span>
            </div>

            <input
              type="range"
              min="1"
              max="100"
              value={config.autoThreshold}
              onChange={(e) =>
                setConfig({
                  ...config,
                  autoThreshold: parseInt(e.target.value) || 70,
                  checkThreshold: parseInt(e.target.value) || 70,
                })
              }
              className="w-full h-3 bg-[#f0f0ed] border-2 border-[#141414] appearance-none cursor-pointer accent-[#00ff00]"
            />

            <div className="bg-[#f0f0ed] p-3 border border-slate-300 text-[11px] text-slate-800 space-y-1">
              <p>
                💥 Dưới <strong className="text-black font-extrabold">{config.autoThreshold}%</strong>:{" "}
                <span className="text-red-600 font-bold">Tự động đề xuất sinh mã mới</span>
              </p>
              <p>
                ✅ Từ <strong className="text-black font-extrabold">{config.autoThreshold}% trở lên</strong>:{" "}
                <span className="text-green-700 font-bold">Tự động gắn mã danh mục sẵn có</span>
              </p>
            </div>
          </div>

          {/* Cấu hình đối chiếu Đơn giá hàng hóa */}
          <div className="space-y-3 bg-[#fdfdfb] p-3.5 border-2 border-[#141414]">
            <div className="flex justify-between items-center">
              <h4 className="text-xs font-black text-black uppercase tracking-wider flex items-center gap-1.5">
                <span>💰</span> Đối chiếu Đơn giá hàng hóa
              </h4>
              <label className="inline-flex items-center cursor-pointer gap-1.5">
                <input
                  type="checkbox"
                  checked={config.enablePriceMatching !== false}
                  onChange={(e) =>
                    setConfig({ ...config, enablePriceMatching: e.target.checked })
                  }
                  className="w-4 h-4 border-2 border-[#141414] accent-[#00ff00]"
                />
                <span className="text-[10px] font-black uppercase text-black">Kích hoạt</span>
              </label>
            </div>

            {config.enablePriceMatching !== false && (
              <div className="space-y-3 text-xs pt-2 border-t border-[#141414]/20">
                <div className="flex items-center justify-between gap-2">
                  <label className="text-[10px] uppercase font-black text-[#555]">
                    Cho phép đơn giá suy ra (Thành tiền / SL)
                  </label>
                  <input
                    type="checkbox"
                    checked={config.allowDerivedPrice !== false}
                    onChange={(e) =>
                      setConfig({ ...config, allowDerivedPrice: e.target.checked })
                    }
                    className="w-4 h-4 border-2 border-[#141414] accent-[#00ff00]"
                  />
                </div>

                {/* Thanh kéo ĐỘ KHỚP ĐƠN GIÁ TỐI THIỂU */}
                <div className="space-y-2 pt-2 border-t border-slate-200">
                  <div className="flex justify-between items-center">
                    <label className="text-[10px] uppercase font-black text-[#141414] tracking-wider">
                      ĐỘ KHỚP ĐƠN GIÁ TỐI THIỂU
                    </label>
                    <span className="font-mono text-xs font-black bg-[#141414] text-[#00ff00] px-2 py-0.5 border border-black shadow-[1px_1px_0px_#141414]">
                      {config.priceMatchThreshold ?? 90}%
                    </span>
                  </div>

                  <div className="flex items-center gap-2 text-[10px] font-mono text-slate-500 font-bold">
                    <span>0%</span>
                    <input
                      type="range"
                      min="0"
                      max="100"
                      step="1"
                      value={config.priceMatchThreshold ?? 90}
                      onChange={(e) =>
                        setConfig({
                          ...config,
                          priceMatchThreshold: parseInt(e.target.value) || 0,
                        })
                      }
                      className="w-full h-2.5 bg-[#e5e5e0] border-2 border-[#141414] appearance-none cursor-pointer accent-[#00ff00]"
                    />
                    <span>100%</span>
                  </div>

                  <div className="bg-[#f0f0ed] p-2 border border-slate-300 text-[11px] text-slate-800 font-medium">
                    Đơn giá đạt từ{" "}
                    <span className="text-black font-extrabold underline">
                      {config.priceMatchThreshold ?? 90}%
                    </span>{" "}
                    trở lên sẽ được tính là khớp.
                  </div>
                </div>
              </div>
            )}
          </div>

          {/* Quy tắc tiền tố sinh mã */}
          <div className="space-y-3 pt-2 border-t-2 border-[#141414]">
            <h4 className="text-xs font-black text-black uppercase tracking-wider flex items-center gap-1.5">
              <Tag size={13} />
              <span>Quy tắc tiền tố sinh mã tự động</span>
            </h4>

            <div className="grid grid-cols-3 gap-2">
              <div>
                <label className="text-[10px] uppercase font-black text-[#666] tracking-wider block mb-1">
                  Hàng hóa
                </label>
                <input
                  type="text"
                  value={config.prefixHH}
                  onChange={(e) => setConfig({ ...config, prefixHH: e.target.value })}
                  className="w-full border-2 border-[#141414] bg-white p-1.5 px-2 font-mono text-xs text-black focus:bg-yellow-50 focus:outline-none text-center font-bold"
                />
              </div>
              <div>
                <label className="text-[10px] uppercase font-black text-[#666] tracking-wider block mb-1">
                  Khách mua
                </label>
                <input
                  type="text"
                  value={config.prefixKH}
                  onChange={(e) => setConfig({ ...config, prefixKH: e.target.value })}
                  className="w-full border-2 border-[#141414] bg-white p-1.5 px-2 font-mono text-xs text-black focus:bg-yellow-50 focus:outline-none text-center font-bold"
                />
              </div>
              <div>
                <label className="text-[10px] uppercase font-black text-[#666] tracking-wider block mb-1">
                  Nhà cung cấp
                </label>
                <input
                  type="text"
                  value={config.prefixNCC}
                  onChange={(e) => setConfig({ ...config, prefixNCC: e.target.value })}
                  className="w-full border-2 border-[#141414] bg-white p-1.5 px-2 font-mono text-xs text-black focus:bg-yellow-50 focus:outline-none text-center font-bold"
                />
              </div>
            </div>
          </div>
        </div>

        {/* BLOCK 2: CƠ SỞ DỮ LIỆU DANH MỤC CHUẨN */}
        <div className="bg-white border-2 border-[#141414] p-5 shadow-[4px_4px_0px_#141414] space-y-6">
          <div className="flex items-center justify-between border-b-2 border-[#141414] pb-3">
            <h3 className="text-xs font-black text-black tracking-widest uppercase flex items-center gap-2">
              <Database size={16} className="text-[#141414]" />
              <span>2. Cơ Sở Dữ Liệu Danh Mục Chuẩn</span>
            </h3>
            <span className="text-[10px] font-mono text-slate-500 font-bold uppercase">
              Tải &amp; Cập nhật
            </span>
          </div>

          <p className="text-xs text-slate-600">
            Hệ thống dùng các danh mục này làm căn cứ chuẩn để đối chiếu, so khớp và gắn mã cho các bảng kê mua bán.
          </p>

          <div className="space-y-4">
            {/* Card Danh mục Hàng hóa */}
            <div className="border-2 border-[#141414] bg-[#fdfdfb] p-4 space-y-3 shadow-[2px_2px_0px_#141414]">
              <div className="flex justify-between items-center">
                <div className="flex items-center gap-2.5">
                  <div className="bg-[#141414] p-1.5 text-white border border-[#141414]">
                    <Database size={15} />
                  </div>
                  <div>
                    <h4 className="font-black text-black uppercase text-xs tracking-wide">
                      Danh Mục Hàng Hóa &amp; Vật Tư
                    </h4>
                    <p className="text-slate-500 text-[10px] uppercase font-bold">
                      Sản phẩm chuẩn trong kho &amp; giá tham chiếu
                    </p>
                  </div>
                </div>
                <span className="font-mono text-xs font-black text-black bg-[#00ff00] px-2 py-0.5 border border-[#141414]">
                  {commodities.length} mã
                </span>
              </div>

              <label className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[3px_3px_0px_#00ff00] active:translate-y-0.5 text-xs font-black uppercase py-2 px-3 border border-[#141414] transition flex items-center justify-center gap-2 cursor-pointer">
                <Upload size={12} className="text-[#00ff00]" />
                <span>Nhập danh mục hàng hóa (.XLSX)</span>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleImportMasterDirectory(e, "commodity")}
                />
              </label>
            </div>

            {/* Card Danh mục Khách hàng */}
            <div className="border-2 border-[#141414] bg-[#fdfdfb] p-4 space-y-3 shadow-[2px_2px_0px_#141414]">
              <div className="flex justify-between items-center">
                <div className="flex items-center gap-2.5">
                  <div className="bg-[#141414] p-1.5 text-white border border-[#141414]">
                    <Users size={15} />
                  </div>
                  <div>
                    <h4 className="font-black text-black uppercase text-xs tracking-wide">
                      Danh Mục Khách Hàng (Người mua)
                    </h4>
                    <p className="text-slate-500 text-[10px] uppercase font-bold">
                      Mã đối tác công nợ đầu ra
                    </p>
                  </div>
                </div>
                <span className="font-mono text-xs font-black text-black bg-[#00ff00] px-2 py-0.5 border border-[#141414]">
                  {customerCount} mã
                </span>
              </div>

              <label className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[3px_3px_0px_#00ff00] active:translate-y-0.5 text-xs font-black uppercase py-2 px-3 border border-[#141414] transition flex items-center justify-center gap-2 cursor-pointer">
                <Upload size={12} className="text-[#00ff00]" />
                <span>Nhập danh mục khách hàng (.XLSX)</span>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleImportMasterDirectory(e, "customer")}
                />
              </label>
            </div>

            {/* Card Danh mục Nhà cung cấp */}
            <div className="border-2 border-[#141414] bg-[#fdfdfb] p-4 space-y-3 shadow-[2px_2px_0px_#141414]">
              <div className="flex justify-between items-center">
                <div className="flex items-center gap-2.5">
                  <div className="bg-[#141414] p-1.5 text-white border border-[#141414]">
                    <Users size={15} />
                  </div>
                  <div>
                    <h4 className="font-black text-black uppercase text-xs tracking-wide">
                      Danh Mục Nhà Cung Cấp (Người bán)
                    </h4>
                    <p className="text-slate-500 text-[10px] uppercase font-bold">
                      Mã đối tác mua vào &amp; chi phí
                    </p>
                  </div>
                </div>
                <span className="font-mono text-xs font-black text-black bg-[#00ff00] px-2 py-0.5 border border-[#141414]">
                  {supplierCount} mã
                </span>
              </div>

              <label className="w-full bg-[#141414] text-white hover:bg-black hover:shadow-[3px_3px_0px_#00ff00] active:translate-y-0.5 text-xs font-black uppercase py-2 px-3 border border-[#141414] transition flex items-center justify-center gap-2 cursor-pointer">
                <Upload size={12} className="text-[#00ff00]" />
                <span>Nhập danh mục nhà cung cấp (.XLSX)</span>
                <input
                  type="file"
                  accept=".xlsx,.xls,.csv"
                  className="hidden"
                  onChange={(e) => handleImportMasterDirectory(e, "supplier")}
                />
              </label>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
};
