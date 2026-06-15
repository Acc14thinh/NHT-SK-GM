/**
 * TypeScript implementation of the Accounting Matching Engine.
 * Optimized for handling extremely large datasets (10,000+ items) efficiently
 * using O(1) Pre-Normalization Caches, 1D space Levenshtein DP, and Early Break Bounds.
 */

import { Commodity, Partner, MappedRow, BankAnalysisResult } from "../types";

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
export function removeVietnameseAccents(str: string): string {
  if (!str) return "";
  return str
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/đ/g, "d")
    .replace(/Đ/g, "D");
}

export function normalizeText(text: string): string {
  if (!text) return "";
  const cached = normTextCache.get(text);
  if (cached !== undefined) return cached;

  let norm = removeVietnameseAccents(text).toLowerCase();
  // Keep only alphanumeric characters, spaces, and specs helpers . - x /
  norm = norm.replace(/[^a-z0-9\s\.\-x/]/g, " ");
  // Remove multiple spaces
  norm = norm.replace(/\s+/g, " ").trim();

  normTextCache.set(text, norm);
  return norm;
}

export function normalizePartnerName(name: string): string {
  if (!name) return "";
  const cached = normPartnerCache.get(name);
  if (cached !== undefined) return cached;

  let text = normalizeText(name);

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
  rowDesc: string,
  rowUom: string,
  commodities: Commodity[],
  autoThreshold: number = 85,
  checkThreshold: number = 70
): { code: string; name: string; score: number; reason: string } {
  if (!rowDesc || !rowDesc.trim()) {
    return { code: "", name: "", score: 0, reason: "Không có mô tả tên hàng" };
  }

  checkAndClearCaches();

  const descNorm = normalizeText(rowDesc);

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
  const cleanId = Math.random().toString(36).substr(2, 9);
  const descNorm = normalizeText(desc);

  // 1a. Fee or Interest
  const feeKeywords = ["phi dich vu", "phi quan ly", "sms banking", "phi giao dich", "phi chuyen tien", "tru phi", "phi cuoc", "phi thuong nien", "lai tien gui", "lai nhap goc", "tra lai"];
  if (feeKeywords.some(keyword => descNorm.includes(keyword))) {
    const isInterest = descNorm.includes("lai");
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc: desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: isInterest ? "Lãi tiền gửi" : "Phí ngân hàng / Tiền lãi",
      proposedCode: "NGANHANG",
      proposedName: "Hệ thống Ngân hàng",
      score: 95,
      reason: "Phát hiện từ khóa liên quan đến thanh toán phí hoặc nhận lãi ngân hàng",
      treatment: "Đã chốt",
      notes: ""
    };
  }

  // 1b. Employee Salary / Advance
  const salaryKeywords = ["luong", "thanh toan luong", "tiet kiem luong", "tam ung", "hoan ung", "khen thuong", "tro cap"];
  if (salaryKeywords.some(keyword => descNorm.includes(keyword))) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc: desc,
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
      notes: ""
    };
  }

  // 1c. Tax / Budget payment
  const taxKeywords = ["nop thue", "nop ngan sach", "thue mon bai", "thue gtgt", "thue tndn", "hai quan", "le phi"];
  if (taxKeywords.some(keyword => descNorm.includes(keyword))) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc: desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: "Nộp thuế / Ngân sách",
      proposedCode: "KHO_BAC",
      proposedName: "Kho bạc Nhà nước",
      score: 95,
      reason: "Phát hiện từ khóa nộp thuế lệ phí ngân sách nhà nước",
      treatment: "Đã chốt",
      notes: ""
    };
  }

  // 1d. Internal Transfer
  const internalKeywords = ["chuyen noi bo", "rut tien mat", "nap tien mat", "nop tien vao tai khoan", "rut tien nhap quy", "nop quy", "rut quy"];
  if (internalKeywords.some(keyword => descNorm.includes(keyword))) {
    return {
      id: cleanId,
      date: new Date().toISOString().substring(0, 10),
      desc: desc,
      amountIn,
      amountOut,
      counterpartAcc,
      counterpartName,
      predictedGroup: "Giao dịch nội bộ / Quỹ",
      proposedCode: "NOI_BO",
      proposedName: "Chuyển khoản nội bộ công ty",
      score: 90,
      reason: "Nội dung thể hiện luân chuyển nguồn vốn trong nội bộ doanh nghiệp",
      treatment: "Đã chốt",
      notes: ""
    } as BankAnalysisResult;
  }

  // 2. Map partner
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
      notes: ""
    };
  }

  // 3. Unclassified or Under threshold but has best match
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
      ? `Điểm tương thích thấp (${match.score}% < ${autoThreshold}%). Đề xuất đối chuẩn thử với: ${match.reason}`
      : "Không tìm thấy thông tin đối tác tương thích nào trong danh mục",
    treatment: "Cần kiểm tra",
    notes: ""
  };
}

export function generateNewCode(prefix: string, existingCodes: string[]): string {
  let counter = 1;
  while (true) {
    const code = `${prefix}${String(counter).padStart(3, "0")}`;
    if (!existingCodes.includes(code)) {
      return code;
    }
    counter++;
  }
}
