import React, { useState } from "react";
import {
  LayoutDashboard,
  Package,
  Users,
  Landmark,
  ShoppingCart,
  Layers,
  SlidersHorizontal,
  Code2,
  ChevronLeft,
  ChevronRight,
  ChevronDown,
  ChevronUp,
  Search,
  Sparkles,
  RefreshCw,
  ExternalLink,
  Store,
  CheckCircle2,
  X
} from "lucide-react";

export type NavTabId =
  | "dashboard"
  | "commodity"
  | "partner"
  | "bank"
  | "integrated"
  | "ecommerce"
  | "python"
  | "settings";

interface NavItem {
  id: NavTabId;
  label: string;
  shortLabel: string;
  icon: React.ComponentType<{ size?: number; className?: string }>;
  description: string;
  badge?: string | number;
}

interface NavGroup {
  id: string;
  title: string;
  items: NavItem[];
}

interface AppSidebarProps {
  currentTab: NavTabId;
  onSelectTab: (tab: NavTabId) => void;
  isCollapsed: boolean;
  onToggleCollapse: () => void;
  isMobileOpen: boolean;
  onCloseMobile: () => void;
  commoditiesCount: number;
  partnersCount: number;
}

export const AppSidebar: React.FC<AppSidebarProps> = ({
  currentTab,
  onSelectTab,
  isCollapsed,
  onToggleCollapse,
  isMobileOpen,
  onCloseMobile,
  commoditiesCount,
  partnersCount,
}) => {
  const [searchQuery, setSearchQuery] = useState("");
  const [collapsedGroups, setCollapsedGroups] = useState<Record<string, boolean>>({});

  const navGroups: NavGroup[] = [
    {
      id: "overview",
      title: "Tổng quan",
      items: [
        {
          id: "dashboard",
          label: "Dashboard Tổng Quan",
          shortLabel: "Dashboard",
          icon: LayoutDashboard,
          description: "Báo cáo tổng hợp số liệu & nạp bảng kê",
        },
      ],
    },
    {
      id: "autocode",
      title: "Gắn mã dữ liệu",
      items: [
        {
          id: "commodity",
          label: "Gán mã Hàng Hóa",
          shortLabel: "Hàng Hóa",
          icon: Package,
          description: "So khớp vật tư, quy cách & đơn giá",
          badge: commoditiesCount > 0 ? `${commoditiesCount} mã` : undefined,
        },
        {
          id: "partner",
          label: "Gán mã Đối Tác",
          shortLabel: "Đối Tác",
          icon: Users,
          description: "Khách hàng & Nhà cung cấp",
          badge: partnersCount > 0 ? `${partnersCount} mã` : undefined,
        },
      ],
    },
    {
      id: "reconciliation",
      title: "Đối chiếu & Kiểm tra",
      items: [
        {
          id: "bank",
          label: "Giao Dịch Sao Kê",
          shortLabel: "Sao Kê",
          icon: Landmark,
          description: "Ngân quỹ, sổ phụ & hóa đơn",
        },
        {
          id: "ecommerce",
          label: "Doanh Thu Sàn TMĐT",
          shortLabel: "Sàn TMĐT",
          icon: ShoppingCart,
          description: "Đối chiếu 2 chiều Shopee & PM bán hàng",
          badge: "Mới",
        },
        {
          id: "integrated",
          label: "Đa Phân Hệ Chéo",
          shortLabel: "Đa Phân Hệ",
          icon: Layers,
          description: "Kiểm tra chéo Kho - Hóa đơn - Ngân quỹ",
        },
      ],
    },
    {
      id: "tools",
      title: "Hệ thống & Cấu hình",
      items: [
        {
          id: "settings",
          label: "Tham Số & Danh Mục",
          shortLabel: "Cài Đặt",
          icon: SlidersHorizontal,
          description: "Độ khớp, đơn giá, tiền tố & tải danh mục",
        },
        {
          id: "python",
          label: "Local Python App",
          shortLabel: "Python App",
          icon: Code2,
          description: "Tải mã nguồn chạy trạm nội bộ",
        },
      ],
    },
  ];

  const toggleGroup = (groupId: string) => {
    setCollapsedGroups((prev) => ({ ...prev, [groupId]: !prev[groupId] }));
  };

  const filteredGroups = navGroups.map((group) => {
    if (!searchQuery.trim()) return group;
    const query = searchQuery.toLowerCase().trim();
    const items = group.items.filter(
      (item) =>
        item.label.toLowerCase().includes(query) ||
        item.description.toLowerCase().includes(query)
    );
    return { ...group, items };
  }).filter((group) => group.items.length > 0);

  const handleItemClick = (id: NavTabId) => {
    onSelectTab(id);
    if (isMobileOpen) {
      onCloseMobile();
    }
  };

  return (
    <>
      {/* Mobile Backdrop */}
      {isMobileOpen && (
        <div
          onClick={onCloseMobile}
          className="fixed inset-0 bg-black/60 z-40 lg:hidden backdrop-blur-xs transition-opacity"
        />
      )}

      {/* Sidebar Container */}
      <aside
        className={`fixed top-0 bottom-0 left-0 z-50 bg-[#141414] text-white flex flex-col border-r-2 border-black transition-all duration-200 ease-in-out ${
          isMobileOpen
            ? "translate-x-0 w-72"
            : "-translate-x-full lg:translate-x-0"
        } ${isCollapsed ? "lg:w-18" : "lg:w-68"}`}
      >
        {/* Brand / Logo Header */}
        <div className="h-16 px-4 border-b-2 border-neutral-800 flex items-center justify-between shrink-0 bg-black">
          <div className="flex items-center gap-3 overflow-hidden">
            <div className="w-9 h-9 bg-[#00ff00] text-black font-black flex items-center justify-center shrink-0 border border-black shadow-[2px_2px_0px_#fff]">
              <RefreshCw size={18} className="text-black font-black" />
            </div>

            {(!isCollapsed || isMobileOpen) && (
              <div className="overflow-hidden">
                <h1 className="text-sm font-black tracking-wider uppercase text-white truncate leading-tight font-sans">
                  SmartLedger
                </h1>
                <p className="text-[10px] font-mono tracking-widest text-[#00ff00] font-bold uppercase truncate">
                  Auto-Accounting
                </p>
              </div>
            )}
          </div>

          {/* Close on mobile */}
          <button
            onClick={onCloseMobile}
            className="lg:hidden p-1.5 text-slate-400 hover:text-white border border-neutral-800"
          >
            <X size={18} />
          </button>
        </div>

        {/* Search input in sidebar (when expanded) */}
        {!isCollapsed && (
          <div className="p-3 border-b border-neutral-800 bg-[#1a1a1a]">
            <div className="relative">
              <Search
                size={13}
                className="absolute left-2.5 top-2.5 text-slate-400"
              />
              <input
                type="text"
                value={searchQuery}
                onChange={(e) => setSearchQuery(e.target.value)}
                placeholder="Tìm phân hệ..."
                className="w-full bg-[#111] border border-neutral-700 pl-8 pr-2.5 py-1 text-xs text-white placeholder:text-neutral-500 font-sans focus:outline-none focus:border-[#00ff00]"
              />
            </div>
          </div>
        )}

        {/* Navigation Menu Items */}
        <div className="flex-1 overflow-y-auto px-2 py-3 space-y-4 select-none scrollbar-thin">
          {filteredGroups.map((group) => {
            const isGroupCollapsed = collapsedGroups[group.id] && !isCollapsed;

            return (
              <div key={group.id} className="space-y-1">
                {/* Group Title Header */}
                {!isCollapsed ? (
                  <div
                    onClick={() => toggleGroup(group.id)}
                    className="px-2.5 py-1 flex items-center justify-between text-[10px] font-black uppercase tracking-wider text-neutral-400 hover:text-white cursor-pointer"
                  >
                    <span>{group.title}</span>
                    <span className="text-neutral-500">
                      {isGroupCollapsed ? (
                        <ChevronDown size={12} />
                      ) : (
                        <ChevronUp size={12} />
                      )}
                    </span>
                  </div>
                ) : (
                  <div className="h-2" />
                )}

                {/* Items in Group */}
                {!isGroupCollapsed && (
                  <div className="space-y-0.5">
                    {group.items.map((item) => {
                      const Icon = item.icon;
                      const isActive = currentTab === item.id;

                      return (
                        <button
                          key={item.id}
                          onClick={() => handleItemClick(item.id)}
                          title={isCollapsed ? `${item.label} - ${item.description}` : undefined}
                          className={`w-full text-left flex items-center gap-3 px-3 py-2.5 text-xs font-bold transition-all relative cursor-pointer border ${
                            isActive
                              ? "bg-white text-black border-white shadow-[2px_2px_0px_#00ff00]"
                              : "text-neutral-300 hover:text-white hover:bg-neutral-900 border-transparent"
                          } ${isCollapsed ? "justify-center px-0" : ""}`}
                        >
                          {/* Active Indicator bar */}
                          {isActive && (
                            <span className="absolute left-0 top-1 bottom-1 w-1 bg-[#00ff00]" />
                          )}

                          <div
                            className={`shrink-0 p-1 border ${
                              isActive
                                ? "bg-black text-[#00ff00] border-black"
                                : "bg-neutral-800 text-neutral-300 border-neutral-700"
                            }`}
                          >
                            <Icon size={16} />
                          </div>

                          {!isCollapsed && (
                            <div className="flex-1 min-w-0 pr-1">
                              <div className="flex items-center justify-between gap-1">
                                <span className={`truncate text-xs ${isActive ? "font-black" : "font-bold"}`}>
                                  {item.label}
                                </span>
                                {item.badge && (
                                  <span
                                    className={`text-[9px] font-mono px-1.5 py-0.2 border shrink-0 font-bold ${
                                      isActive
                                        ? "bg-black text-[#00ff00] border-black"
                                        : "bg-neutral-800 text-neutral-300 border-neutral-700"
                                    }`}
                                  >
                                    {item.badge}
                                  </span>
                                )}
                              </div>
                              <p className={`text-[10px] truncate ${isActive ? "text-neutral-600 font-medium" : "text-neutral-500 font-normal"}`}>
                                {item.description}
                              </p>
                            </div>
                          )}
                        </button>
                      );
                    })}
                  </div>
                )}
              </div>
            );
          })}
        </div>

        {/* Footer info & Collapse Toggle */}
        <div className="p-3 border-t-2 border-neutral-800 bg-black shrink-0 flex items-center justify-between">
          {!isCollapsed && (
            <div className="overflow-hidden">
              <span className="text-[10px] font-mono font-bold text-neutral-400 block truncate">
                NHT Studio • v2.8 PRO
              </span>
              <span className="text-[9px] font-mono text-[#00ff00] block truncate">
                ● Hệ thống sẵn sàng
              </span>
            </div>
          )}

          {/* Desktop collapse button */}
          <button
            onClick={onToggleCollapse}
            title={isCollapsed ? "Mở rộng thanh menu" : "Thu gọn thanh menu"}
            className="hidden lg:flex p-1.5 bg-neutral-900 hover:bg-neutral-800 text-neutral-300 hover:text-white border border-neutral-700 cursor-pointer items-center justify-center mx-auto"
          >
            {isCollapsed ? <ChevronRight size={16} /> : <ChevronLeft size={16} />}
          </button>
        </div>
      </aside>
    </>
  );
};
