/**
 * TypeScript implementation of the Accounting Matching Engine.
 * Optimized for handling extremely large datasets (10,000+ items) efficiently
 * using O(1) Pre-Normalization Caches, 1D space Levenshtein DP, and Early Break Bounds.
 */

import { Commodity, Partner, MappedRow, BankAnalysisResult, ColumnMapping } from "../types";
import { parseAmount } from "./excelParser";

// Maximum size in memory for primitive key caches to avoid leaks
const MAX_CACHE_SIZE = 50000;

// Primitive Key Caches
const normTextCache = new Map<string, string>();
const normPartnerCache = new Map<string, string>();
const fuzzyRatioCache = new Map<string, number>();

// Dynamic Caches bounded clearing
function checkAndClearCaches() {
  if (normTextCache.size > MAX_CACHE_SIZE) normTextCache.clear();
  if (normPartnerCache.size > MAX_CACHE_SIZE) normPartnerCache.clear();
  if (fuzzyRatioCache.size > MAX_CACHE_SIZE) fuzzyRatioCache.clear();
}

// Pre-normalization caches for database objects (WeakMap automatically handles GC when arrays change)
interface CachedCommodityNorm {
  nameNorm: string;
  uomNorm: string;
  keywords: string[];
  specsNorm: string;
}

interface CachedPartnerNorm {
  nameNorm: string;
  mstClean: string;
  accClean: string;
  keywords: string[];
}

const commodityNormCache = new WeakMap<Commodity, CachedCommodityNorm>();
const partnerNormCache = new WeakMap<Partner, CachedPartnerNorm>();

// Removes Vietnamese accents / diacritics
export function removeVietnameseAccents(str: any): string {
  if (str === null || str === undefined) return "";
  const safeStr = typeof str === "string" ? str : String(str);
  if (!safeStr) return "";
  return safeStr
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");
}

export function normalizeText(text: any): string {
  if (text === null || text === undefined) return "";
  const safeText = typeof text === "string" ? text : String(text);
  if (!safeText) return "";
  const cached = normTextCache.get(safeText);
  if (cached !== undefined) return cached;

  let norm = removeVietnameseAccents(safeText).toLowerCase();
  // Keep only alphanumeric characters, spaces, and specs helpers . - x /
  norm = norm.replace(/[^a-z0-9\s\.\-x/]/g, " ");
  // Remove multiple spaces
  norm = norm.replace(/\s+/g, " ").trim();

  normTextCache.set(safeText, norm);
  return norm;
}

export function normalizePartnerName(name: any): string {
  if (name === null || name === undefined) return "";
  const safeName = typeof name === "string" ? name : String(name);
  if (!safeName) return "";
  const cached = normPartnerCache.get(safeName);
  if (cached !== undefined) return cached;

  let text = normalizeText(safeName);

  const redundantWords = [
    /\bcong ty tnhh mtv\b/g,
    /\bcong ty tnhh\b/g,
    /\bcong ty co phan\b/g,
    /\bcong ty cp\b/g,
    /\bchi nhanh\b/g,
    /\bdoanh nghiep tu nhan\b/g,
    /\bdntn\b/g,
    /\bho kinh doanh\b/g,
    /\bhkd\b/g,
    /\bco ltd\b/g,
    /\bjsc\b/g,
    /\bltd\b/g,
    /\bcong ty\b/g,
    /\bcty\b/g,
    /\bmtv\b/g,
    /\bco phan\b/g,
    /\bcp\b/g
  ];

  for (const regex of redundantWords) {
    text = text.replace(regex, "");
  }

  const result = text.replace(/\s+/g, " ").trim();
  normPartnerCache.set(name, result);
  return result;
}

export function extractSpecs(text: string): string[] {
  if (!text) return [];
  const textClean = text.toLowerCase();
  const specs: string[] = [];

  // Extract weight/volume (e.g., 500ml, 1kg, 20kg, 5l)
  const volMatches = textClean.match(/\b\d+(?:\.\d+)?\s*(?:ml|l|g|kg|ton|lit|gam|chai|hop)\b/g);
  if (volMatches) {
    volMatches.forEach(v => specs.push(v.replace(/\s+/g, "")));
  }

  // Extract dimensions like 10x20
  const dimMatches = textClean.match(/\b\d+(?:\.\d+)?\s*x\s*\d+(?:\.\d+)?(?:\s*x\s*\d+(?:\.\d+)?)?\b/g);
  if (dimMatches) {
    dimMatches.forEach(d => specs.push(d.replace(/\s+/g, "")));
  }

  // Common types / models like A4, PCB40, Phi 6
  const models = [
    /\ba\d\b/g,            // A4, A3
    /\bpcb\d+\b/g,         // PCB40
    /\bphi\s*\d+\b/g,      // Phi 6
    /\bd\d+\b/g,           // D90
    /\bmodel\s*[a-z0-9\-]+\b/g
  ];

  for (const regex of models) {
    const matches = textClean.match(regex);
    if (matches) {
      matches.forEach(m => specs.push(m.trim()));
    }
  }

  return Array.from(new Set(specs));
}

/**
 * Optimized Levenshtein Distance using a 1D DP Array and Cache Lookup.
 * Dramatically reduces garbage collection (GC) pauses and memory overhead
 * on huge Excel datasets (runs in under 1ms).
 */
export function getFuzzyRatio(s1: string, s2: string): number {
  if (s1 === s2) return 100;

  // Stable cache lookup
  const cacheKey = s1 < s2 ? `${s1}|${s2}` : `${s2}|${s1}`;
  const cached = fuzzyRatioCache.get(cacheKey);
  if (cached !== undefined) return cached;

  const len1 = s1.length;
  const len2 = s2.length;
  if (len1 === 0 && len2 === 0) {
    fuzzyRatioCache.set(cacheKey, 100);
    return 100;
  }
  if (len1 === 0 || len2 === 0) {
    fuzzyRatioCache.set(cacheKey, 0);
    return 0;
  }

  // We want to reduce allocations by working on the shorter string
  const shortStr = len1 < len2 ? s1 : s2;
  const longStr = len1 < len2 ? s2 : s1;
  const shortLen = shortStr.length;
  const longLen = longStr.length;

  // Single flat 1D array instead of a full matrix allocation
  const dp = new Array(shortLen + 1);
  for (let i = 0; i <= shortLen; i++) {
    dp[i] = i;
  }

  for (let i = 1; i <= longLen; i++) {
    let prevDiag = dp[0];
    dp[0] = i;
    const charLong = longStr[i - 1];

    for (let j = 1; j <= shortLen; j++) {
      const temp = dp[j];
      if (charLong === shortStr[j - 1]) {
        dp[j] = prevDiag;
      } else {
        dp[j] = Math.min(
          prevDiag + 1,  // substitution
          dp[j - 1] + 1, // insertion
          dp[j] + 1      // deletion
        );
      }
      prevDiag = temp;
    }
  }

  const distance = dp[shortLen];
  const ratio = Math.round((1 - distance / longLen) * 100);
  fuzzyRatioCache.set(cacheKey, ratio);
  return ratio;
}

export function matchCommodityRow(
  rowDesc: any,
  rowUom: any,
  commodities: Commodity[],
  autoThreshold: number = 85,
  checkThreshold: number = 70
): { code: string; name: string; score: number; reason: string } {
  const safeDesc = typeof rowDesc === "string" ? rowDesc : String(rowDesc || "");
  if (!safeDesc.trim()) {
    return { code: "", name: "", score: 0, reason: "Không có mô tả tên hàng" };
  }

  checkAndClearCaches();

  const descNorm = normalizeText(safeDesc);

  let bestScore = 0;
  let bestMatch: Commodity | null = null;
  let bestReason = "Không tìm thấy trong danh mục";

  for (const item of commodities) {
    // Retrieve pre-normalized representations in O(1) via WeakMap
    let cached = commodityNormCache.get(item);
    if (!cached) {
      cached = {
        nameNorm: normalizeText(item.ten_hang_hoa_chuan),
        uomNorm: normalizeText(item.don_vi_tinh),
        keywords: (item.tu_khoa_nhan_dien || "").split(",").map(k => normalizeText(k)).filter(Boolean),
        specsNorm: normalizeText(item.quy_cach || "")
      };
      commodityNormCache.set(item, cached);
    }

    const { nameNorm: itemNameNorm } = cached;

    let score = 0;
    if (descNorm === itemNameNorm && descNorm !== "") {
      score = 100;
    } else {
      score = getFuzzyRatio(descNorm, itemNameNorm);
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatch = item;
      bestReason = score === 100
        ? "Trùng khớp chính xác 100% cột tên hàng"
        : `Trùng khớp ${score}% cột tên hàng`;
    }

    // Early break if perfect match is found
    if (bestScore >= 100) {
      break;
    }
  }

  if (!bestMatch) {
    return { code: "", name: "", score: 0, reason: bestReason };
  }

  return {
    code: bestMatch.ma_hang_hoa,
    name: bestMatch.ten_hang_hoa_chuan,
    score: bestScore,
    reason: bestReason
  };
}

export function matchPartnerRow(
  rowName: string,
  rowMst: string,
  rowAcc: string,
  invoiceDesc: string,
  partners: Partner[],
  isBuyer: boolean,
  autoThreshold: number = 85,
  checkThreshold: number = 70
): { code: string; name: string; type: string; score: number; reason: string } {
  if (!rowName && !rowMst && !rowAcc) {
    return { code: "", name: "", type: "", score: 0, reason: "Không có thông tin đối tác" };
  }

  checkAndClearCaches();

  const nameNorm = normalizePartnerName(rowName);
  const mstClean = rowMst ? rowMst.trim().replace(/-/g, "") : "";
  const accClean = rowAcc ? rowAcc.trim() : "";
  const invoiceNorm = normalizeText(invoiceDesc);

  let bestScore = 0;
  let bestMatch: Partner | null = null;
  let bestReason = "Không có đối tượng phù hợp";

  for (const item of partners) {
    // Retrieve pre-normalized representations in O(1) via WeakMap
    let cached = partnerNormCache.get(item);
    if (!cached) {
      cached = {
        nameNorm: normalizePartnerName(item.ten_doi_tuong),
        mstClean: item.ma_so_thue ? item.ma_so_thue.trim().replace(/-/g, "") : "",
        accClean: item.so_tai_khoan ? item.so_tai_khoan.trim() : "",
        keywords: (item.tu_khoa_nhan_dien || "").split(",").map(k => normalizeText(k)).filter(Boolean)
      };
      partnerNormCache.set(item, cached);
    }

    const { nameNorm: pNameNorm, mstClean: pMst, accClean: pAcc } = cached;

    let score = 0;
    const reasons: string[] = [];

    // 1. Trùng mã số thuế (Ưu tiên tuyệt đối)
    if (mstClean && pMst && mstClean === pMst) {
      score = 100;
      reasons.push("Trùng khớp MST tuyệt đối 100%");
    }
    // 2. Trùng tài khoản ngân hàng (Ưu tiên tuyệt đối)
    else if (accClean && pAcc && accClean === pAcc) {
      score = 100;
      reasons.push("Trùng khớp Số tài khoản tuyệt đối 100%");
    }
    // 3. Tên trùng tuyệt đối
    else if (nameNorm && pNameNorm && nameNorm === pNameNorm) {
      score = 100;
      reasons.push("Tên khớp tuyệt đối 100%");
    }
    // 4. Khớp fuzzy trực tiếp
    else if (nameNorm && pNameNorm) {
      score = getFuzzyRatio(nameNorm, pNameNorm);
      reasons.push(`Độ trùng khớp tên: ${score}%`);
    }

    if (score > bestScore) {
      bestScore = score;
      bestMatch = item;
      bestReason = reasons.join("; ");
    }

    // Early break if perfect match is found
    if (bestScore >= 100) {
      break;
    }
  }

  if (!bestMatch) {
    return { code: "", name: "", type: "", score: 0, reason: bestReason };
  }

  return {
    code: bestMatch.ma_doi_tuong,
    name: bestMatch.ten_doi_tuong,
    type: bestMatch.loai_doi_tuong,
    score: bestScore,
    reason: bestReason
  };
}

export interface BankReconConfig {
  autoThreshold: number;
  checkThreshold: number;
  daysBeforeInvoice: number; // default 7
  daysAfterInvoice: number;  // default 30
  diffAbsThreshold: number;  // default 10000
  diffPctThreshold: number;  // default 0.5
  maxCombinationCount: number; // default 5
  prefixKH: string;
  prefixNCC: string;
}

export interface MatchingScores {
  scoreDesc: number;
  scoreName: number;
  scoreAmount: number;
  scoreDate: number;
  scoreAcc: number;
  scoreMst: number;
  scoreInvoice: number;
  scoreHistory: number;
  scorePenalty: number;
  totalScore: number;
  reasons: string[];
}

// Robust helper to parse diverse date types (standard string, Vietnamese DD/MM/YYYY, Excel serial)
export function parseDate(dateStr: any): Date | null {
  if (!dateStr) return null;
  if (dateStr instanceof Date) return dateStr;
  const str = String(dateStr).trim();
  if (!str) return null;

  // 1. Try YYYY-MM-DD
  let match = str.match(/^(\d{4})[-/](\d{1,2})[-/](\d{1,2})/);
  if (match) {
    return new Date(parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]));
  }

  // 2. Try DD/MM/YYYY or DD-MM-YYYY
  match = str.match(/^(\d{1,2})[-/](\d{1,2})[-/](\d{4})/);
  if (match) {
    return new Date(parseInt(match[3]), parseInt(match[2]) - 1, parseInt(match[1]));
  }

  // 3. Try Excel numeric serial values
  const num = parseFloat(str);
  if (!isNaN(num) && num > 30000 && num < 60000) {
    const excelEpoch = new Date(1899, 11, 30);
    const msInDay = 24 * 60 * 60 * 1000;
    return new Date(excelEpoch.getTime() + num * msInDay);
  }

  const d = new Date(str);
  return isNaN(d.getTime()) ? null : d;
}

export function formatDateString(dateStr: any): string {
  const d = parseDate(dateStr);
  if (!d) return "";
  const yyyy = d.getFullYear();
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const dd = String(d.getDate()).padStart(2, "0");
  return `${yyyy}-${mm}-${dd}`;
}

export function getDaysDifference(d1: any, d2: any): number {
  const date1 = parseDate(d1);
  const date2 = parseDate(d2);
  if (!date1 || !date2) return 9999;
  const diffTime = Math.abs(date1.getTime() - date2.getTime());
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24));
}

export function getDateScore(absDiffDays: number): number {
  if (absDiffDays === 0) return 20;
  if (absDiffDays <= 3) return 15;
  if (absDiffDays <= 7) return 10;
  if (absDiffDays <= 15) return 7;
  if (absDiffDays <= 30) return 5;
  return 0;
}

export function checkAmountWithinThreshold(
  bankAmount: number,
  invoiceAmount: number,
  diffAbsThreshold: number = 10000,
  diffPctThreshold: number = 0.5
): { isMatch: boolean; isExact: boolean; diffAmt: number; diffPct: number } {
  const diffAmt = Math.abs(bankAmount - invoiceAmount);
  if (diffAmt === 0) {
    return { isMatch: true, isExact: true, diffAmt: 0, diffPct: 0 };
  }
  const diffPct = (diffAmt / invoiceAmount) * 100;
  const isWithinAbs = diffAmt <= diffAbsThreshold;
  const isWithinPct = diffPct <= diffPctThreshold;
  const isMatch = isWithinAbs || isWithinPct;
  return { isMatch, isExact: false, diffAmt, diffPct };
}

// Find partner record related to an invoice row
export function getInvoicePartner(
  invoiceRow: any,
  isSales: boolean,
  partners: Partner[],
  mappings: ColumnMapping
): Partner | null {
  const nameVal = isSales ? invoiceRow[mappings.ten_khach_hang] : invoiceRow[mappings.ten_nha_cung_cap];
  const mstVal = isSales ? invoiceRow[mappings.ma_so_thue_khach_hang] : invoiceRow[mappings.ma_so_thue_ncc];
  const codeVal = isSales ? invoiceRow[mappings.ma_khach_hang] : invoiceRow[mappings.ma_nha_cung_cap];

  if (codeVal) {
    const p = partners.find(item => String(item.ma_doi_tuong).trim().toLowerCase() === String(codeVal).trim().toLowerCase());
    if (p) return p;
  }

  if (!nameVal && !mstVal) return null;

  // Fallback to partner match
  const match = matchPartnerRow(
    String(nameVal || ""),
    String(mstVal || ""),
    "",
    "",
    partners,
    isSales,
    80,
    70
  );

  if (match.code) {
    const p = partners.find(item => item.ma_doi_tuong === match.code);
    if (p) return p;
  }

  return null;
}

// Find a subset of invoices belonging to the same partner that sum to bankAmount (pruned subset sum)
export function findInvoiceCombinations(
  invoices: any[],
  bankAmount: number,
  maxCombo: number = 5,
  absThreshold: number = 10000,
  pctThreshold: number = 0.5,
  amountKey: string
): any[][] {
  const results: any[][] = [];

  function search(startIndex: number, currentCombo: any[], currentSum: number) {
    const diff = Math.abs(currentSum - bankAmount);
    const pct = currentSum > 0 ? (diff / currentSum) * 100 : 999;

    if (diff === 0 || diff <= absThreshold || pct <= pctThreshold) {
      if (currentCombo.length > 0) {
        results.push([...currentCombo]);
      }
    }

    if (currentCombo.length >= maxCombo) {
      return;
    }

    for (let i = startIndex; i < sortedInvoices.length; i++) {
      const invAmt = parseFloat(sortedInvoices[i][amountKey]) || 0;
      if (invAmt <= 0) continue;

      if (currentSum + invAmt > bankAmount + Math.max(absThreshold, bankAmount * pctThreshold / 100)) {
        continue; // Pruning path for positive values
      }

      currentCombo.push(sortedInvoices[i]);
      search(i + 1, currentCombo, currentSum + invAmt);
      currentCombo.pop();
    }
  }

  const sortedInvoices = [...invoices].sort((a, b) => {
    const amtA = parseFloat(a[amountKey]) || 0;
    const amtB = parseFloat(b[amountKey]) || 0;
    return amtA - amtB;
  });

  search(0, [], 0);
  return results;
}

// Groups ledger rows by invoice number to sum up multi-line invoices before matching (Section 4)
export function groupLedgerByInvoice(ledger: any[], mappings: ColumnMapping): any[] {
  if (!ledger || ledger.length === 0) return [];
  const groups = new Map<string, { key: string; rows: any[] }>();

  ledger.forEach(row => {
    const rawNo = String(row[mappings.so_hoa_don] || "").trim();
    const cleanNo = rawNo.replace(/^0+/, "");
    if (!rawNo) return; // ignore rows without invoice number

    if (!groups.has(cleanNo)) {
      groups.set(cleanNo, { key: cleanNo, rows: [] });
    }
    groups.get(cleanNo)!.rows.push(row);
  });

  const groupedRows: any[] = [];
  for (const [cleanNo, group] of groups.entries()) {
    const firstRow = group.rows[0];
    const totalAmount = group.rows.reduce((sum, r) => {
      const val = parseFloat(r[mappings.tong_thanh_toan]) || 0;
      return sum + val;
    }, 0);

    const mergedRow = { ...firstRow };
    mergedRow[mappings.tong_thanh_toan] = totalAmount;
    mergedRow._subRows = group.rows;

    groupedRows.push(mergedRow);
  }

  return groupedRows;
}

// Calculate the multi-criteria matching scores between bank row and specific invoice
export function computeMatchingScores(params: {
  bankDesc: string;
  bankAmount: number;
  bankDateStr: string;
  bankCounterpartAcc?: string;
  bankCounterpartName?: string;
  invoiceNo?: string;
  invoiceDateStr?: string;
  invoiceAmount?: number;
  invoiceMst?: string;
  partner: Partner;
  config: {
    daysBeforeInvoice: number;
    daysAfterInvoice: number;
    diffAbsThreshold: number;
    diffPctThreshold: number;
  };
  isSales: boolean;
  isInvoiceAlreadyPaidFully: boolean;
  hasMultipleClientsWithSameAmount: boolean;
  confirmedAccountRules?: Record<string, string>;
  allEligiblePartners?: Partner[];
}): MatchingScores {
  const normDesc = normalizeText(params.bankDesc);
  const partnerNameNorm = normalizePartnerName(params.partner.ten_doi_tuong);
  const counterpartNameNorm = params.bankCounterpartName ? normalizePartnerName(params.bankCounterpartName) : "";
  const counterpartAccClean = params.bankCounterpartAcc ? params.bankCounterpartAcc.trim() : "";

  let scoreDesc = 0;
  let scoreAmount = 0;
  let scoreDate = 0;

  // Detailed indicator scores
  let scoreMst = 0;
  let scoreAcc = 0;
  let scoreInvoice = 0;
  let scoreName = 0;
  let scoreHistory = 0;
  let scorePenalty = 0;
  const reasons: string[] = [];

  // Evaluate description match (Max 50 points):
  // 1. Trùng số hóa đơn
  if (params.invoiceNo) {
    const cleanInvoiceNo = String(params.invoiceNo).trim().replace(/^0+/, "");
    if (cleanInvoiceNo && cleanInvoiceNo.length >= 2 && normDesc.includes(cleanInvoiceNo)) {
      scoreInvoice = 50;
      if (50 > scoreDesc) {
        scoreDesc = 50;
        reasons.push("Trùng khớp số hóa đơn (+50đ)");
      }
    }
  }

  // 2. Trùng mã số thuế
  if (params.invoiceMst && params.partner.ma_so_thue) {
    const mstClean1 = String(params.invoiceMst).trim().replace(/-/g, "");
    const mstClean2 = String(params.partner.ma_so_thue).trim().replace(/-/g, "");
    if (mstClean1 && mstClean2 && mstClean1 === mstClean2) {
      scoreMst = 50;
      if (50 > scoreDesc) {
        scoreDesc = 50;
        reasons.push("Trùng khớp mã số thuế (+50đ)");
      }
    }
  }

  // 3. Trùng số tài khoản
  if (counterpartAccClean && params.partner.so_tai_khoan) {
    const accClean2 = String(params.partner.so_tai_khoan).trim();
    if (counterpartAccClean === accClean2 || normDesc.includes(accClean2)) {
      scoreAcc = 50;
      if (50 > scoreDesc) {
        scoreDesc = 50;
        reasons.push("Trùng khớp số tài khoản (+50đ)");
      }
    }
  }

  // 4. Trùng mã đối tượng
  const codeClean = String(params.partner.ma_doi_tuong).trim().toLowerCase();
  if (codeClean && normDesc.includes(codeClean)) {
    if (45 > scoreDesc) {
      scoreDesc = 45;
      reasons.push("Trùng mã đối tượng trong diễn giải (+45đ)");
    }
  }

  // 5. Trùng chính xác tên công ty chuẩn hóa
  let isExactNameMatch = false;
  if (partnerNameNorm && partnerNameNorm.length > 3) {
    if (counterpartNameNorm === partnerNameNorm || normDesc.includes(partnerNameNorm)) {
      scoreName = 40;
      isExactNameMatch = true;
      if (40 > scoreDesc) {
        scoreDesc = 40;
        reasons.push("Trùng chính xác tên đối tác chuẩn hóa (+40đ)");
      }
    }
  }

  // 6. Tên gần giống fuzzy
  if (!isExactNameMatch && partnerNameNorm && partnerNameNorm.length > 3) {
    let maxRatio = 0;
    if (counterpartNameNorm) {
      maxRatio = Math.max(maxRatio, getFuzzyRatio(partnerNameNorm, counterpartNameNorm));
    }
    maxRatio = Math.max(maxRatio, getFuzzyRatio(partnerNameNorm, normDesc));

    if (maxRatio >= 90) {
      scoreName = 30;
      if (30 > scoreDesc) {
        scoreDesc = 30;
        reasons.push(`Tên đối tác gần giống trên 90% (Độ khớp: ${maxRatio}%) (+30đ)`);
      }
    } else if (maxRatio >= 75) {
      scoreName = 20;
      if (20 > scoreDesc) {
        scoreDesc = 20;
        reasons.push(`Tên đối tác gần giống từ 75% đến dưới 90% (Độ khớp: ${maxRatio}%) (+20đ)`);
      }
    }
  }

  // 7. Từ khóa nhận diện / viết tắt
  const keywords = (params.partner.tu_khoa_nhan_dien || "").split(",").map(k => normalizeText(k)).filter(Boolean);
  if (keywords.some(kw => normDesc.includes(kw))) {
    if (25 > scoreDesc) {
      scoreDesc = 25;
      reasons.push("Khớp từ khóa nhận diện của đối tác (+25đ)");
    }
  }

  // Evaluate Amount (Max 35 points):
  if (params.invoiceAmount !== undefined) {
    const amtCheck = checkAmountWithinThreshold(
      params.bankAmount,
      params.invoiceAmount,
      params.config.diffAbsThreshold,
      params.config.diffPctThreshold
    );
    if (amtCheck.isExact) {
      scoreAmount = 35;
      reasons.push("Số tiền khớp chính xác 100% (+35đ)");
    } else if (amtCheck.isMatch) {
      scoreAmount = 20;
      reasons.push(`Số tiền lệch trong ngưỡng (+20đ, Chênh: ${Math.round(amtCheck.diffAmt).toLocaleString()}đ)`);
    } else {
      scoreAmount = 0;
      reasons.push("Số tiền ngoài ngưỡng chênh lệch (0đ)");
    }
  }

  // Evaluate Date (Max 15 points):
  if (params.invoiceDateStr && params.bankDateStr) {
    const t_bank = parseDate(params.bankDateStr);
    const t_invoice = parseDate(params.invoiceDateStr);
    if (t_bank && t_invoice) {
      const diffMs = t_bank.getTime() - t_invoice.getTime();
      const diffDays = diffMs / (1000 * 60 * 60 * 24);

      if (diffDays >= -params.config.daysBeforeInvoice && diffDays <= params.config.daysAfterInvoice) {
        const absDays = Math.abs(Math.round(diffDays));
        if (absDays === 0) {
          scoreDate = 15;
          reasons.push("Giao dịch trùng ngày hóa đơn (+15đ)");
        } else if (absDays <= 3) {
          scoreDate = 12;
          reasons.push(`Chênh lệch ngày rất nhỏ (${absDays} ngày) (+12đ)`);
        } else if (absDays <= 7) {
          scoreDate = 8;
          reasons.push(`Chênh lệch ngày nhỏ (${absDays} ngày) (+8đ)`);
        } else if (absDays <= 15) {
          scoreDate = 5;
          reasons.push(`Chênh lệch ngày trung bình (${absDays} ngày) (+5đ)`);
        } else if (absDays <= 30) {
          scoreDate = 3;
          reasons.push(`Chênh lệch ngày lớn (${absDays} ngày) (+3đ)`);
        }
      } else {
        scoreDate = 0;
        reasons.push(`Ngày nằm ngoài khoảng cài đặt (Lệch ${Math.round(diffDays)} ngày) (0đ)`);
      }
    }
  }

  // Confirm historical matches if any
  if (counterpartAccClean && params.confirmedAccountRules) {
    const confirmedCode = params.confirmedAccountRules[counterpartAccClean];
    if (confirmedCode && confirmedCode === params.partner.ma_doi_tuong) {
      scoreHistory = 30;
      reasons.push("Lịch sử số tài khoản đã được xác nhận (+30đ)");
    }
  }

  // Apply penalties (e.g. cash flow direction contradiction)
  if (!params.isSales && params.partner.loai_doi_tuong === "Khách hàng") {
    scorePenalty += 100;
    reasons.push("Phạt: Dòng tiền chi nhưng đề xuất Khách hàng (-100đ)");
  }
  if (params.isSales && params.partner.loai_doi_tuong === "Nhà cung cấp") {
    scorePenalty += 100;
    reasons.push("Phạt: Dòng tiền thu nhưng đề xuất Nhà cung cấp (-100đ)");
  }
  if (params.isInvoiceAlreadyPaidFully) {
    scorePenalty += 40;
    reasons.push("Phạt: Hóa đơn đã được thanh toán đủ từ trước (-40đ)");
  }

  const rawTotal = scoreDesc + scoreAmount + scoreDate + scoreHistory - scorePenalty;
  const totalScore = Math.max(0, Math.min(100, rawTotal));

  return {
    scoreDesc,
    scoreName,
    scoreAmount,
    scoreDate,
    scoreAcc,
    scoreMst,
    scoreInvoice,
    scoreHistory,
    scorePenalty,
    totalScore,
    reasons
  };
}

// Pre-processes invoices to automatically generate newly detected partners (Section 17)
export function preProcessInvoicePartners(
  salesLedger: any[],
  salesMappings: ColumnMapping,
  purchaseLedger: any[],
  purchaseMappings: ColumnMapping,
  existingPartners: Partner[],
  prefixKH: string,
  prefixNCC: string
): Partner[] {
  const newPartners: Partner[] = [];
  const existingCodes = [...existingPartners.map(p => p.ma_doi_tuong)];
  const activePartners = [...existingPartners];

  function findPartner(name: string, mst: string): Partner | null {
    const nameNorm = normalizePartnerName(name);
    const mstClean = mst ? String(mst).trim().replace(/-/g, "") : "";

    return activePartners.find(p => {
      if (mstClean && p.ma_so_thue && String(p.ma_so_thue).trim().replace(/-/g, "") === mstClean) {
        return true;
      }
      if (nameNorm && normalizePartnerName(p.ten_doi_tuong) === nameNorm) {
        return true;
      }
      return false;
    }) || null;
  }

  // Sales
  salesLedger.forEach(row => {
    const name = row[salesMappings.ten_khach_hang] || "";
    const mst = row[salesMappings.ma_so_thue_khach_hang] || "";
    if (!name && !mst) return;

    const existing = findPartner(name, mst);
    if (!existing) {
      const nextCode = generateNewCode(prefixKH, existingCodes);
      const newP: Partner = {
        ma_doi_tuong: nextCode,
        ten_doi_tuong: String(name || "Khách hàng mới").trim(),
        loai_doi_tuong: "Khách hàng",
        ma_so_thue: String(mst || "").trim(),
        so_tai_khoan: "",
        ngan_hang: "",
        dia_chi: "Tự tạo từ bảng kê bán ra",
        tu_khoa_nhan_dien: "",
        ghi_chu: "Nguồn tạo mã: Bảng kê bán ra"
      };
      newPartners.push(newP);
      activePartners.push(newP);
      existingCodes.push(nextCode);
    }
  });

  // Purchase
  purchaseLedger.forEach(row => {
    const name = row[purchaseMappings.ten_nha_cung_cap] || "";
    const mst = row[purchaseMappings.ma_so_thue_ncc] || "";
    if (!name && !mst) return;

    const existing = findPartner(name, mst);
    if (!existing) {
      const nextCode = generateNewCode(prefixNCC, existingCodes);
      const newP: Partner = {
        ma_doi_tuong: nextCode,
        ten_doi_tuong: String(name || "Nhà cung cấp mới").trim(),
        loai_doi_tuong: "Nhà cung cấp",
        ma_so_thue: String(mst || "").trim(),
        so_tai_khoan: "",
        ngan_hang: "",
        dia_chi: "Tự tạo từ bảng kê mua vào",
        tu_khoa_nhan_dien: "",
        ghi_chu: "Nguồn tạo mã: Bảng kê mua vào"
      };
      newPartners.push(newP);
      activePartners.push(newP);
      existingCodes.push(nextCode);
    }
  });

  return newPartners;
}

interface InvoiceCandidate {
  invoiceRow: any;
  partner: Partner;
  scores: MatchingScores;
  invoiceNo: string;
  invoiceDateStr: string;
  invoiceAmount: number;
  unpaidAmount: number;
  differenceAmount: number;
  differencePercentage: number;
  differenceDays: number;
}

interface ComboCandidate {
  partner: Partner;
  invoices: any[];
  combinedAmount: number;
  differenceAmount: number;
  differencePercentage: number;
  combinedScore: number;
  reason: string;
  scoreName?: number;
  scoreAcc?: number;
  scoreMst?: number;
  scoreInvoice?: number;
  scoreAmount?: number;
  scoreDate?: number;
  scoreHistory?: number;
  scorePenalty?: number;
  reasons?: string[];
}

// Master execution engine to analyze a single bank statement row with multi-criteria rules
export function matchBankRowWithInvoiceEngine(params: {
  bankRow: any;
  bankMappings: ColumnMapping;
  partners: Partner[];
  salesLedger: any[];
  salesMappings: ColumnMapping;
  purchaseLedger: any[];
  purchaseMappings: ColumnMapping;
  config: BankReconConfig;
  invoiceBalances: Map<string, number>; // remaining amount tracker
  existingCodes: string[];
  tempCreatedPartners: Partner[];
  confirmedAccountRules?: Record<string, string>;
}): BankAnalysisResult {
  const cleanId = params.bankRow._id || `bank_${Math.random().toString(36).substr(2, 9)}`;
  const desc = String(params.bankRow[params.bankMappings.noi_dung_giao_dich] || "");
  const normDesc = normalizeText(desc);

  const amIn = parseAmount(params.bankRow[params.bankMappings.so_tien_thu]);
  const amOut = parseAmount(params.bankRow[params.bankMappings.so_tien_chi]);
  const bankAmount = amIn > 0 ? amIn : amOut;
  const isSales = amIn > 0;

  const bankDateVal = params.bankRow[params.bankMappings.ngay_giao_dich] || "";
  const bObj = parseDate(bankDateVal);
  const bankDateStr = bObj ? `${bObj.getFullYear()}-${String(bObj.getMonth() + 1).padStart(2, "0")}-${String(bObj.getDate()).padStart(2, "0")}` : "";

  const counterpartAcc = String(params.bankRow[params.bankMappings.so_tai_khoan_doi_ung] || "").trim();
  const counterpartName = String(params.bankRow[params.bankMappings.ten_doi_tac_sao_ke] || "").trim();

  // Rule 12: Exclusions checking (Interest, Fees, Salary, Tax, Internal, Loan, Personal)
  const isInterest = normDesc.includes("lai") || normDesc.includes("tien gui") || normDesc.includes("tiet kiem") || normDesc.includes("interest");
  const isFee = normDesc.includes("phi dich vu") || normDesc.includes("phi duy tri") || normDesc.includes("phi chuyen tien") || normDesc.includes("sms") || normDesc.includes("fee") || normDesc.includes("cuoc thue bao");
  const isSalary = normDesc.includes("luong") || normDesc.includes("tam ung") || normDesc.includes("thuong") || normDesc.includes("salary") || normDesc.includes("phu cap");
  const isTax = normDesc.includes("thue") || normDesc.includes("nop ngan sach") || normDesc.includes("tax") || normDesc.includes("hai quan");
  const isInternal = normDesc.includes("noi bo") || normDesc.includes("chuyen quy") || normDesc.includes("rut tien") || normDesc.includes("nop tien mat") || normDesc.includes("nap tien");
  const isLoan = normDesc.includes("vay") || normDesc.includes("tra no vay") || normDesc.includes("tat toan") || normDesc.includes("dao han") || normDesc.includes("credit") || normDesc.includes("loan");
  const isPersonal = normDesc.includes("ca nhan") || normDesc.includes("chi tieu") || normDesc.includes("mua sam") || normDesc.includes("rut mat") || normDesc.includes("an uong");

  if (isInterest || isFee || isSalary || isTax || isInternal || isLoan || isPersonal) {
    let reasonExclude = "";
    if (isInterest) reasonExclude = "Giao dịch lãi tiền gửi / lãi tiết kiệm";
    else if (isFee) reasonExclude = "Giao dịch phí ngân hàng / dịch vụ";
    else if (isSalary) reasonExclude = "Giao dịch thanh toán lương / thưởng / tạm ứng";
    else if (isTax) reasonExclude = "Giao dịch nộp thuế / nghĩa vụ ngân sách nhà nước";
    else if (isInternal) reasonExclude = "Giao dịch chuyển tiền nội bộ / rút nộp quỹ";
    else if (isLoan) reasonExclude = "Giao dịch giải ngân / tất toán khoản vay";
    else if (isPersonal) reasonExclude = "Giao dịch chi tiêu cá nhân / không liên quan sản xuất kinh doanh";

    return {
      id: cleanId,
      date: bankDateStr,
      desc: desc,
      amountIn: amIn,
      amountOut: amOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: reasonExclude,
      proposedCode: "",
      proposedName: "",
      proposedType: "",
      score: 0,
      reason: `Loại trừ tự động: ${reasonExclude}`,
      treatment: "Cần kiểm tra",
      notes: reasonExclude,
      rawRowData: params.bankRow,
      matchingSource: "Loại trừ",
      processingStatus: "Giao dịch không liên quan công nợ",
      scoreName: 0,
      totalScore: 0
    };
  }

  // Group ledger rows by invoice number first (Section 4)
  const salesLedgerGrouped = groupLedgerByInvoice(params.salesLedger, params.salesMappings);
  const purchaseLedgerGrouped = groupLedgerByInvoice(params.purchaseLedger, params.purchaseMappings);

  const ledger = isSales ? salesLedgerGrouped : purchaseLedgerGrouped;
  const ledgerMappings = isSales ? params.salesMappings : params.purchaseMappings;

  // Filter active partners based on type
  const activePartners = [...params.partners, ...params.tempCreatedPartners];

  // Find if multiple clients with the same amount exist
  const matchingAmountPartners = new Set<string>();
  for (const inv of ledger) {
    const invoiceAmount = parseAmount(inv[ledgerMappings.tong_thanh_toan]);
    if (invoiceAmount <= 0) continue;
    
    const amtCheck = checkAmountWithinThreshold(
      bankAmount,
      invoiceAmount,
      params.config.diffAbsThreshold,
      params.config.diffPctThreshold
    );
    if (amtCheck.isMatch) {
      const partner = getInvoicePartner(inv, isSales, activePartners, ledgerMappings);
      if (partner) {
        matchingAmountPartners.add(partner.ma_doi_tuong);
      }
    }
  }
  const hasMultipleClientsWithSameAmount = matchingAmountPartners.size >= 2;

  // Single candidate evaluation loop
  const candidates: InvoiceCandidate[] = [];

  for (const inv of ledger) {
    const invoiceNo = String(inv[ledgerMappings.so_hoa_don] || "").trim();
    const invoiceDateStr = formatDateString(inv[ledgerMappings.ngay_hoa_don]);
    const originalInvoiceAmount = parseAmount(inv[ledgerMappings.tong_thanh_toan]);

    if (!invoiceNo || originalInvoiceAmount <= 0) continue;

    const partner = getInvoicePartner(inv, isSales, activePartners, ledgerMappings);
    if (!partner) continue;

    // Prune incompatible partner types: Thu only Khách hàng, Chi only Nhà cung cấp
    if (isSales && partner.loai_doi_tuong !== "Khách hàng") continue;
    if (!isSales && partner.loai_doi_tuong !== "Nhà cung cấp") continue;

    const invoiceKey = `${invoiceNo}_${invoiceDateStr}_${originalInvoiceAmount}`;
    const unpaidAmount = params.invoiceBalances.has(invoiceKey)
      ? params.invoiceBalances.get(invoiceKey)!
      : originalInvoiceAmount;

    const isInvoiceAlreadyPaidFully = unpaidAmount <= 0;
    const invoiceMst = isSales ? inv[ledgerMappings.ma_so_thue_khach_hang] : inv[ledgerMappings.ma_so_thue_ncc];

    // Calculate score
    const scores = computeMatchingScores({
      bankDesc: desc,
      bankAmount: bankAmount,
      bankDateStr: bankDateStr,
      bankCounterpartAcc: counterpartAcc,
      bankCounterpartName: counterpartName,
      invoiceNo,
      invoiceDateStr,
      invoiceAmount: unpaidAmount > 0 ? unpaidAmount : originalInvoiceAmount,
      invoiceMst,
      partner,
      config: params.config,
      isSales,
      isInvoiceAlreadyPaidFully,
      hasMultipleClientsWithSameAmount,
      confirmedAccountRules: params.confirmedAccountRules,
      allEligiblePartners: activePartners
    });

    const diffDays = getDaysDifference(bankDateStr, invoiceDateStr);
    const amtCheck = checkAmountWithinThreshold(
      bankAmount,
      unpaidAmount > 0 ? unpaidAmount : originalInvoiceAmount,
      params.config.diffAbsThreshold,
      params.config.diffPctThreshold
    );

    candidates.push({
      invoiceRow: inv,
      partner,
      scores,
      invoiceNo,
      invoiceDateStr,
      invoiceAmount: originalInvoiceAmount,
      unpaidAmount,
      differenceAmount: amtCheck.diffAmt,
      differencePercentage: amtCheck.diffPct,
      differenceDays: diffDays
    });
  }

  candidates.sort((a, b) => b.scores.totalScore - a.scores.totalScore);

  // Combination evaluations (Section 10: multiple invoices for one partner)
  const partnerInvoicesMap = new Map<string, { partner: Partner; invoices: any[] }>();
  for (const inv of ledger) {
    const partner = getInvoicePartner(inv, isSales, activePartners, ledgerMappings);
    if (!partner) continue;
    if (isSales && partner.loai_doi_tuong !== "Khách hàng") continue;
    if (!isSales && partner.loai_doi_tuong !== "Nhà cung cấp") continue;

    const code = partner.ma_doi_tuong;
    if (!partnerInvoicesMap.has(code)) {
      partnerInvoicesMap.set(code, { partner, invoices: [] });
    }

    const invoiceNo = String(inv[ledgerMappings.so_hoa_don] || "").trim();
    const invoiceDateStr = formatDateString(inv[ledgerMappings.ngay_hoa_don]);
    const originalInvoiceAmount = parseAmount(inv[ledgerMappings.tong_thanh_toan]);
    const invoiceKey = `${invoiceNo}_${invoiceDateStr}_${originalInvoiceAmount}`;
    const unpaidAmount = params.invoiceBalances.has(invoiceKey) ? params.invoiceBalances.get(invoiceKey)! : originalInvoiceAmount;

    if (unpaidAmount > 0) {
      partnerInvoicesMap.get(code)!.invoices.push({
        ...inv,
        _unpaidAmount: unpaidAmount,
        _invoiceNo: invoiceNo,
        _invoiceDateStr: invoiceDateStr,
        _originalInvoiceAmount: originalInvoiceAmount
      });
    }
  }

  const comboCandidates: ComboCandidate[] = [];

  for (const [code, entry] of partnerInvoicesMap.entries()) {
    if (entry.invoices.length < 2) continue;
    
    // Safety check: limit combinations input size
    const slicedInvoices = entry.invoices.slice(0, 15);

    const combos = findInvoiceCombinations(
      slicedInvoices,
      bankAmount,
      params.config.maxCombinationCount || 5,
      params.config.diffAbsThreshold,
      params.config.diffPctThreshold,
      "_unpaidAmount"
    );

    for (const combo of combos) {
      const combinedAmount = combo.reduce((s, item) => s + item._unpaidAmount, 0);
      const diffAmt = Math.abs(bankAmount - combinedAmount);
      const diffPct = combinedAmount > 0 ? (diffAmt / combinedAmount) * 100 : 0;

      const p = entry.partner;
      const invoiceMst = isSales ? combo[0][ledgerMappings.ma_so_thue_khach_hang] : combo[0][ledgerMappings.ma_so_thue_ncc];

      // Score combo using unified logic (Section 11)
      const scores = computeMatchingScores({
        bankDesc: desc,
        bankAmount: bankAmount,
        bankDateStr: bankDateStr,
        bankCounterpartAcc: counterpartAcc,
        bankCounterpartName: counterpartName,
        invoiceNo: combo.map(item => item._invoiceNo).join("+"),
        invoiceDateStr: combo[0]._invoiceDateStr,
        invoiceAmount: combinedAmount,
        invoiceMst,
        partner: p,
        config: params.config,
        isSales,
        isInvoiceAlreadyPaidFully: false,
        hasMultipleClientsWithSameAmount,
        confirmedAccountRules: params.confirmedAccountRules,
        allEligiblePartners: activePartners
      });

      const invoiceNumbers = combo.map(item => item._invoiceNo).join(", ");

      comboCandidates.push({
        partner: p,
        invoices: combo,
        combinedAmount,
        differenceAmount: diffAmt,
        differencePercentage: diffPct,
        combinedScore: scores.totalScore,
        reason: `Khớp tổ hợp ${combo.length} hóa đơn: ${invoiceNumbers}. Tổng tiền: ${combinedAmount.toLocaleString()}đ.`,
        scoreName: scores.scoreName,
        scoreAcc: scores.scoreAcc,
        scoreMst: scores.scoreMst,
        scoreInvoice: scores.scoreInvoice,
        scoreAmount: scores.scoreAmount,
        scoreDate: scores.scoreDate,
        scoreHistory: scores.scoreHistory,
        scorePenalty: scores.scorePenalty,
        reasons: scores.reasons
      });
    }
  }

  comboCandidates.sort((a, b) => b.combinedScore - a.combinedScore);

  // Compile top 3 proposals
  const top3: BankAnalysisResult["top3Proposals"] = [];
  candidates.forEach(c => {
    top3.push({
      code: c.partner.ma_doi_tuong,
      name: c.partner.ten_doi_tuong,
      invoiceNo: c.invoiceNo,
      invoiceDate: c.invoiceDateStr,
      invoiceAmount: c.unpaidAmount,
      score: c.scores.totalScore,
      reason: c.scores.reasons.join("; ")
    });
  });

  comboCandidates.forEach(c => {
    top3.push({
      code: c.partner.ma_doi_tuong,
      name: c.partner.ten_doi_tuong,
      invoiceNo: c.invoices.map(item => item._invoiceNo).join("+"),
      invoiceDate: c.invoices.map(item => item._invoiceDateStr).join("+"),
      invoiceAmount: c.combinedAmount,
      score: c.combinedScore,
      reason: c.reason
    });
  });

  top3.sort((a, b) => b.score - a.score);
  const finalTop3 = top3.slice(0, 3);

  const bestSingle = candidates[0];
  const bestCombo = comboCandidates[0];

  const hasSingle = !!bestSingle;
  const hasCombo = !!bestCombo;

  let isComboSelected = false;
  let finalMatch: any = null;

  if (hasSingle && hasCombo) {
    if (bestCombo.combinedScore > bestSingle.scores.totalScore && bestCombo.combinedAmount > 0) {
      isComboSelected = true;
      finalMatch = bestCombo;
    } else {
      isComboSelected = false;
      finalMatch = bestSingle;
    }
  } else if (hasSingle) {
    isComboSelected = false;
    finalMatch = bestSingle;
  } else if (hasCombo) {
    isComboSelected = true;
    finalMatch = bestCombo;
  }

  // FALLBACK: NO MATCH FOUND
  if (!finalMatch) {
    const searchName = counterpartName || desc;
    let fallbackPartner: Partner | null = null;
    if (counterpartAcc) {
      const cleanAcc = counterpartAcc.trim();
      fallbackPartner = activePartners.find(p => p.so_tai_khoan && p.so_tai_khoan.trim() === cleanAcc) || null;
    }
    
    if (!fallbackPartner && searchName) {
      const partnerMatch = matchPartnerRow(searchName, "", counterpartAcc || "", desc, activePartners, isSales, params.config.autoThreshold, params.config.autoThreshold);
      if (partnerMatch.score >= params.config.autoThreshold) {
        fallbackPartner = activePartners.find(p => p.ma_doi_tuong === partnerMatch.code) || null;
      }
    }

    if (fallbackPartner && (isSales ? fallbackPartner.loai_doi_tuong === "Khách hàng" : fallbackPartner.loai_doi_tuong === "Nhà cung cấp")) {
      const scores = computeMatchingScores({
        bankDesc: desc,
        bankAmount: bankAmount,
        bankDateStr: bankDateStr,
        bankCounterpartAcc: counterpartAcc,
        bankCounterpartName: counterpartName,
        partner: fallbackPartner,
        config: params.config,
        isSales,
        isInvoiceAlreadyPaidFully: false,
        hasMultipleClientsWithSameAmount: false,
        confirmedAccountRules: params.confirmedAccountRules,
        allEligiblePartners: activePartners
      });

      // Fallback is only automatically matched if highly confident (e.g. exact name, account, or tax code match)
      const isConfidentDesc = scores.scoreAcc >= 50 || scores.scoreMst >= 50 || scores.scoreHistory >= 30 || scores.scoreName >= 40;
      const finalTreatment = isConfidentDesc ? "Đã chốt" : "Cần kiểm tra";
      const finalStatus = isConfidentDesc ? "Tự động gắn" : "Cần kiểm tra lại";

      return {
        id: cleanId,
        date: bankDateStr,
        desc: desc,
        amountIn: amIn,
        amountOut: amOut,
        counterpartAcc,
        counterpartName,
        predictedGroup: isSales ? "Thu tiền khách hàng" : "Chi thanh toán nhà cung cấp",
        proposedCode: fallbackPartner.ma_doi_tuong,
        proposedName: fallbackPartner.ten_doi_tuong,
        proposedType: fallbackPartner.loai_doi_tuong,
        score: scores.totalScore,
        reason: "Khớp đối tác theo tài khoản ngân hàng hoặc danh mục chuẩn hóa không có hóa đơn",
        treatment: finalTreatment,
        notes: "Gắn mã dựa trên thông tin danh mục đối tác",
        rawRowData: params.bankRow,
        matchingSource: "Diễn giải",
        processingStatus: finalStatus,
        scoreName: scores.scoreName,
        scoreAcc: scores.scoreAcc,
        scoreMst: scores.scoreMst,
        scoreInvoice: scores.scoreInvoice,
        scoreAmount: scores.scoreAmount,
        scoreDate: scores.scoreDate,
        scoreHistory: scores.scoreHistory,
        scorePenalty: scores.scorePenalty,
        totalScore: scores.totalScore
      };
    } else {
      return {
        id: cleanId,
        date: bankDateStr,
        desc: desc,
        amountIn: amIn,
        amountOut: amOut,
        counterpartAcc,
        counterpartName,
        predictedGroup: isSales ? "Thu tiền khách hàng (Chưa rõ)" : "Chi thanh toán nhà cung cấp (Chưa rõ)",
        proposedCode: "",
        proposedName: "",
        proposedType: isSales ? "Khách hàng" : "Nhà cung cấp",
        score: 0,
        reason: "Chưa đủ căn cứ tạo mã khách hàng mới",
        treatment: "Cần kiểm tra",
        notes: "Chưa đủ căn cứ tạo mã khách hàng mới",
        rawRowData: params.bankRow,
        matchingSource: "Không khớp",
        processingStatus: "Cần kiểm tra lại",
        scoreName: 0,
        totalScore: 0
      };
    }
  }

  // SCENARIO A: COMBINATION SELECTED
  if (isComboSelected) {
    const combo = finalMatch as ComboCandidate;
    const secondScore = comboCandidates[1] ? comboCandidates[1].combinedScore : 0;
    const diffScore = combo.combinedScore - secondScore;

    // Check Case A & Case B (Section 8)
    const hasStrongId =
      combo.scoreMst >= 50 ||
      combo.scoreAcc >= 50 ||
      combo.scoreInvoice >= 45 ||
      combo.scoreHistory >= 30 ||
      combo.scoreName >= 40;

    const isCaseA = hasStrongId && (combo.combinedScore >= 80);
    const isCaseB =
      !hasStrongId &&
      combo.scoreAmount === 35 &&
      combo.scoreDate > 0 &&
      !hasMultipleClientsWithSameAmount &&
      (diffScore >= 15);

    const canAutoAssignCombo =
      combo.combinedScore >= 85 &&
      combo.scorePenalty === 0 &&
      (isCaseA || isCaseB);

    const finalTreatment = canAutoAssignCombo ? "Đã chốt" : "Cần kiểm tra";
    let finalStatus = "";
    if (canAutoAssignCombo) {
      finalStatus = "Tự động gắn";
    } else if (hasMultipleClientsWithSameAmount || (diffScore >= 0 && diffScore < 15)) {
      finalStatus = "Nhiều phương án, cần kiểm tra lại";
    } else if (combo.combinedScore >= 70) {
      finalStatus = "Đề xuất, cần kiểm tra lại";
    } else {
      finalStatus = "Cần kiểm tra lại";
    }

    // Allocate payment
    combo.invoices.forEach(item => {
      const invoiceKey = `${item._invoiceNo}_${item._invoiceDateStr}_${item._originalInvoiceAmount}`;
      params.invoiceBalances.set(invoiceKey, 0);
    });

    const invoiceNumbers = combo.invoices.map(item => item._invoiceNo).join("+");
    const invoiceDates = combo.invoices.map(item => item._invoiceDateStr).join("+");

    return {
      id: cleanId,
      date: bankDateStr,
      desc: desc,
      amountIn: amIn,
      amountOut: amOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: isSales ? "Thu tiền khách hàng" : "Chi thanh toán nhà cung cấp",
      proposedCode: combo.partner.ma_doi_tuong,
      proposedName: combo.partner.ten_doi_tuong,
      proposedType: combo.partner.loai_doi_tuong,
      score: combo.combinedScore,
      reason: combo.reason,
      treatment: finalTreatment,
      notes: `Khớp tổ hợp ${combo.invoices.length} hóa đơn: ${invoiceNumbers}. Tổng tiền: ${combo.combinedAmount.toLocaleString()}đ. Tất cả đều đã tất toán.`,
      rawRowData: params.bankRow,
      matchingSource: "Số tiền & Ngày",
      processingStatus: finalStatus,
      matchedInvoiceNo: invoiceNumbers,
      matchedInvoiceDate: invoiceDates,
      invoiceAmount: combo.combinedAmount,
      differenceAmount: combo.differenceAmount,
      differencePercentage: combo.differencePercentage,
      differenceDays: 0,
      scoreName: combo.scoreName,
      scoreAcc: combo.scoreAcc,
      scoreMst: combo.scoreMst,
      scoreInvoice: combo.scoreInvoice,
      scoreAmount: combo.scoreAmount,
      scoreDate: combo.scoreDate,
      scoreHistory: combo.scoreHistory,
      scorePenalty: combo.scorePenalty,
      totalScore: combo.combinedScore,
      top3Proposals: finalTop3
    };
  }

  // SCENARIO B: SINGLE INVOICE CANDIDATE SELECTED
  else {
    const cand = finalMatch as InvoiceCandidate;
    const secondScore = candidates[1] ? candidates[1].scores.totalScore : 0;
    const diffScore = cand.scores.totalScore - secondScore;

    // Check Case A & Case B (Section 8)
    const hasStrongId =
      cand.scores.scoreMst >= 50 ||
      cand.scores.scoreAcc >= 50 ||
      cand.scores.scoreInvoice >= 45 ||
      cand.scores.scoreHistory >= 30 ||
      cand.scores.scoreName >= 40;

    const isCaseA = hasStrongId && (cand.scores.totalScore >= 80);
    const isCaseB =
      !hasStrongId &&
      cand.scores.scoreAmount === 35 &&
      cand.scores.scoreDate > 0 &&
      !hasMultipleClientsWithSameAmount &&
      (diffScore >= 15);

    const canAutoAssign =
      cand.scores.totalScore >= 85 &&
      cand.scores.scorePenalty === 0 &&
      (isCaseA || isCaseB);

    const finalTreatment = canAutoAssign ? "Đã chốt" : "Cần kiểm tra";
    let finalStatus = "";
    if (canAutoAssign) {
      finalStatus = "Tự động gắn";
    } else if (hasMultipleClientsWithSameAmount || (diffScore >= 0 && diffScore < 15)) {
      finalStatus = "Nhiều phương án, cần kiểm tra lại";
    } else if (cand.scores.totalScore >= 70) {
      finalStatus = "Đề xuất, cần kiểm tra lại";
    } else {
      finalStatus = "Cần kiểm tra lại";
    }

    // Allocate payment
    const originalInvoiceAmount = cand.invoiceAmount;
    const invoiceKey = `${cand.invoiceNo}_${cand.invoiceDateStr}_${originalInvoiceAmount}`;
    const unpaidAmountBefore = cand.unpaidAmount;

    let allocatedPayment = 0;
    let remainingDebt = 0;

    if (bankAmount < unpaidAmountBefore) {
      allocatedPayment = bankAmount;
      remainingDebt = unpaidAmountBefore - bankAmount;
      params.invoiceBalances.set(invoiceKey, remainingDebt);
    } else {
      allocatedPayment = unpaidAmountBefore;
      remainingDebt = 0;
      params.invoiceBalances.set(invoiceKey, 0);
    }

    const payStatusStr = remainingDebt === 0 ? "Thanh toán đủ" : "Thanh toán một phần";

    return {
      id: cleanId,
      date: bankDateStr,
      desc: desc,
      amountIn: amIn,
      amountOut: amOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: isSales ? "Thu tiền khách hàng" : "Chi thanh toán nhà cung cấp",
      proposedCode: cand.partner.ma_doi_tuong,
      proposedName: cand.partner.ten_doi_tuong,
      proposedType: cand.partner.loai_doi_tuong,
      score: cand.scores.totalScore,
      reason: cand.scores.reasons.join("; "),
      treatment: finalTreatment,
      notes: `Hóa đơn ${cand.invoiceNo} (${cand.invoiceDateStr}) - ${payStatusStr}. Thu lần này: ${allocatedPayment.toLocaleString()}đ. Nợ còn lại: ${remainingDebt.toLocaleString()}đ`,
      rawRowData: params.bankRow,
      matchingSource: cand.scores.scoreInvoice >= 45 || cand.scores.scoreMst >= 50 ? "Hóa đơn" : (cand.scores.scoreName >= 20 ? "Diễn giải" : "Số tiền & Ngày"),
      processingStatus: finalStatus,
      matchedInvoiceNo: cand.invoiceNo,
      matchedInvoiceDate: cand.invoiceDateStr,
      invoiceAmount: originalInvoiceAmount,
      differenceAmount: cand.differenceAmount,
      differencePercentage: cand.differencePercentage,
      differenceDays: cand.differenceDays,
      scoreName: cand.scores.scoreName,
      scoreAcc: cand.scores.scoreAcc,
      scoreMst: cand.scores.scoreMst,
      scoreInvoice: cand.scores.scoreInvoice,
      scoreAmount: cand.scores.scoreAmount,
      scoreDate: cand.scores.scoreDate,
      scoreHistory: cand.scores.scoreHistory,
      scorePenalty: cand.scores.scorePenalty,
      totalScore: cand.scores.totalScore,
      top3Proposals: finalTop3
    };
  }
}

export function analyzeBankTransaction(
  desc: string,
  amountIn: number,
  amountOut: number,
  counterpartAcc: string,
  counterpartName: string,
  partners: Partner[],
  autoThreshold: number = 85,
  checkThreshold: number = 70
): BankAnalysisResult {
  const cleanId = `bank_simple_${Math.random().toString(36).substr(2, 9)}`;
  const descNorm = normalizeText(desc);

  // 1. Exclusions check (standard fallback)
  const isInterest = descNorm.includes("lai") || descNorm.includes("nhap goc") || descNorm.includes("tra lai");
  const isFee = ["phi dich vu", "phi quan ly", "sms banking", "phi giao dich", "phi chuyen tien", "tru phi", "phi cuoc", "phi thuong nien", "tru cuoc"].some(kw => descNorm.includes(kw));
  if (isFee || isInterest) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: isInterest ? "Lãi tiền gửi" : "Phí ngân hàng / Tiền lãi",
      proposedCode: "NGANHANG",
      proposedName: "Hệ thống Ngân hàng",
      score: 100,
      reason: "Hạch toán phí dịch vụ hoặc tiền lãi suất gửi tiết kiệm ngân hàng",
      treatment: "Đã chốt",
      notes: "Loại trừ khỏi công nợ",
      matchingSource: "Không liên quan công nợ",
      processingStatus: "Giao dịch không liên quan công nợ"
    };
  }

  const salaryKeywords = ["luong", "thanh toan luong", "tiet kiem luong", "tam ung", "hoan ung", "khen thuong", "tro cap"];
  if (salaryKeywords.some(keyword => descNorm.includes(keyword))) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: "Chi lương / Tạm ứng",
      proposedCode: "NHANVIEN",
      proposedName: "Cán bộ nhân viên công ty",
      score: 90,
      reason: "Hạch toán chi lương thưởng, bảo hiểm hoặc tạm ứng nhân viên",
      treatment: "Cần kiểm tra",
      notes: "Loại trừ khỏi công nợ",
      matchingSource: "Không liên quan công nợ",
      processingStatus: "Giao dịch không liên quan công nợ"
    };
  }

  const taxKeywords = ["nop thue", "nop ngan sach", "thue mon bai", "thue gtgt", "thue tndn", "hai quan", "le phi"];
  if (taxKeywords.some(keyword => descNorm.includes(keyword))) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: "Nộp thuế / Ngân sách",
      proposedCode: "KHO_BAC",
      proposedName: "Kho bạc Nhà nước",
      score: 100,
      reason: "Thuế giá trị gia tăng, thuế môn bài, thuế TNDN hoặc hải quan",
      treatment: "Đã chốt",
      notes: "Loại trừ khỏi công nợ",
      matchingSource: "Không liên quan công nợ",
      processingStatus: "Giao dịch không liên quan công nợ"
    };
  }

  const internalKeywords = ["chuyen noi bo", "rut tien mat", "nap tien mat", "nop tien vao tai khoan", "rut tien nhap quy", "nop quy", "rut quy", "rut nhap quy"];
  if (internalKeywords.some(keyword => descNorm.includes(keyword))) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: "Giao dịch nội bộ / Quỹ",
      proposedCode: "NOI_BO",
      proposedName: "Chuyển khoản nội bộ công ty",
      score: 100,
      reason: "Giao dịch quỹ nội bộ của doanh nghiệp",
      treatment: "Đã chốt",
      notes: "Loại trừ khỏi công nợ",
      matchingSource: "Không liên quan công nợ",
      processingStatus: "Giao dịch không liên quan công nợ"
    };
  }

  const isBuyer = amountIn > 0;
  const searchName = counterpartName || desc;

  const match = matchPartnerRow(searchName, "", counterpartAcc, desc, partners, isBuyer, autoThreshold, autoThreshold);

  if (match.score >= autoThreshold) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc: desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: isBuyer ? "Thu tiền khách hàng" : "Chi thanh toán nhà cung cấp",
      proposedCode: match.code,
      proposedName: match.name,
      score: match.score,
      reason: `Khớp đối tác tự động: ${match.reason}`,
      treatment: "Đã chốt",
      notes: "",
      matchingSource: "Diễn giải",
      processingStatus: "Tự động gắn theo nội dung"
    };
  }

  return {
    id: cleanId,
    date: new Date().toISOString().substring(0, 10),
    desc: desc,
    amountIn,
    amountOut,
    counterpartAcc,
    counterpartName,
    predictedGroup: isBuyer ? "Thu nguồn chưa rõ (Cần đối chiếu)" : "Chi chưa rõ đối tác (Cần đối chiếu)",
    proposedCode: match.code || "",
    proposedName: match.name || "Nghi vấn / Thất lạc",
    score: match.score,
    reason: match.code
      ? `Điểm tương thích thấp (${match.score}% < ${autoThreshold}%). Đề xuất: ${match.reason}`
      : "Không tìm thấy thông tin đối tác phù hợp",
    treatment: "Cần kiểm tra",
    notes: "",
    matchingSource: "Diễn giải",
    processingStatus: "Không đủ căn cứ xác định"
  };
}

export function generateNewCode(prefix: string, existingCodes: string[]): string {
  let counter = 1;
  const codesSet = new Set(existingCodes.map(c => String(c).trim().toUpperCase()));
  while (true) {
    const code = `${prefix}${String(counter).padStart(3, "0")}`;
    if (!codesSet.has(code.toUpperCase())) {
      return code;
    }
    counter++;
  }
}
