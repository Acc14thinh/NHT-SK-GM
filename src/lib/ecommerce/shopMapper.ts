/**
 * Shop Mapping Module: Maps Marketplace Sheets / Shops to Sales Software Shops
 * Section 4 & 5
 */
import { ShopMappingEntry } from "./types";
import { normalizeShopName } from "./orderNormalizer";

const LOCAL_STORAGE_KEY = "nht_ecommerce_shop_mappings_v1";
const LOCAL_STORAGE_HEADER_MKT = "nht_ecommerce_mkt_headers_v1";
const LOCAL_STORAGE_HEADER_SALES = "nht_ecommerce_sales_headers_v1";

/**
 * Đọc mapping header đã lưu từ localStorage
 */
export function getSavedHeaderMappings(type: "marketplace" | "sales"): Record<string, string> {
  try {
    const key = type === "marketplace" ? LOCAL_STORAGE_HEADER_MKT : LOCAL_STORAGE_HEADER_SALES;
    const raw = localStorage.getItem(key);
    if (!raw) return {};
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

/**
 * Lưu mapping header vào localStorage
 */
export function saveHeaderMappings(type: "marketplace" | "sales", mapping: Record<string, string>): void {
  try {
    const key = type === "marketplace" ? LOCAL_STORAGE_HEADER_MKT : LOCAL_STORAGE_HEADER_SALES;
    localStorage.setItem(key, JSON.stringify(mapping));
  } catch (e) {
    console.warn("Không thể lưu header mapping vào localStorage", e);
  }
}

/**
 * Đọc mapping đã lưu từ localStorage
 */
export function getSavedShopMappings(): Record<string, string> {
  try {
    const raw = localStorage.getItem(LOCAL_STORAGE_KEY);
    if (!raw) return {};
    return JSON.parse(raw);
  } catch {
    return {};
  }
}

/**
 * Lưu mapping vào localStorage
 */
export function saveShopMappings(mappings: ShopMappingEntry[]): void {
  try {
    const existing = getSavedShopMappings();
    mappings.forEach(m => {
      if (m.marketplaceSheetOrShop && m.salesShopName) {
        existing[m.marketplaceSheetOrShop] = m.salesShopName;
      }
    });
    localStorage.setItem(LOCAL_STORAGE_KEY, JSON.stringify(existing));
  } catch (e) {
    console.warn("Không thể lưu shop mapping vào localStorage", e);
  }
}

/**
 * Tạo danh sách mapping đề xuất giữa Sàn và Phần Mềm:
 * - Trường hợp 1-1: 1 sheet sàn, 1 shop PM -> tự động ghép
 * - Dùng mapping đã lưu trước đó nếu có
 * - Gợi ý dựa trên normalized text similarity
 */
export function buildProposedShopMappings(
  marketplaceSheets: string[],
  salesShops: string[]
): ShopMappingEntry[] {
  const saved = getSavedShopMappings();
  const results: ShopMappingEntry[] = [];

  // Trường hợp 1-1
  if (marketplaceSheets.length === 1 && salesShops.length === 1) {
    return [{
      marketplaceSheetOrShop: marketplaceSheets[0],
      salesShopName: salesShops[0],
      isConfirmed: true
    }];
  }

  // Chuẩn bị normalized sets
  const availableSalesShops = [...salesShops];

  for (const sheet of marketplaceSheets) {
    // 1. Kiểm tra đã lưu trước đó
    if (saved[sheet] && salesShops.includes(saved[sheet])) {
      results.push({
        marketplaceSheetOrShop: sheet,
        salesShopName: saved[sheet],
        isConfirmed: true
      });
      continue;
    }

    // 2. Exact match không phân biệt hoa thường
    const exactMatch = salesShops.find(s => s.toLowerCase().trim() === sheet.toLowerCase().trim());
    if (exactMatch) {
      results.push({
        marketplaceSheetOrShop: sheet,
        salesShopName: exactMatch,
        isConfirmed: true
      });
      continue;
    }

    // 3. Normalized similarity (ví dụ "Metaky" vs "Metaky Store", "Msmarty" vs "M'Smarty Shop")
    const normSheet = normalizeShopName(sheet);
    let bestCandidate = "";
    let isCandidateCertain = false;

    for (const shop of availableSalesShops) {
      const normShop = normalizeShopName(shop);
      if (normShop === normSheet) {
        bestCandidate = shop;
        isCandidateCertain = true;
        break;
      }
      if (normShop.includes(normSheet) || normSheet.includes(normShop)) {
        bestCandidate = shop;
        // Nếu bắt đầu bằng tiền tố hoặc chứa toàn bộ -> xác suất cao
        if (normShop.startsWith(normSheet) || normSheet.startsWith(normShop)) {
          isCandidateCertain = true;
        }
        break;
      }
    }

    results.push({
      marketplaceSheetOrShop: sheet,
      salesShopName: bestCandidate || (salesShops[0] || ""),
      isConfirmed: isCandidateCertain
    });
  }

  return results;
}
