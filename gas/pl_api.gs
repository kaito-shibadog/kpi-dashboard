/**
 * KPI Dashboard — Google Apps Script
 * シート: <統合>PL
 *
 * 【シート構造】
 *   Row 1: タイトル行 (PL：<統合>, フィルタ等)
 *   Row 2: ヘッダー行 (連結勘定科目, 11月, 12月, 1月 ... 合計)
 *   Row 3〜: データ行
 *   A列: 大分類カテゴリ (結合セル)
 *   B列: 連結勘定科目
 *   C列〜: 月次数値
 */

const SHEET_NAME = "<統合>PL";
const HEADER_ROW = 2;   // 月ラベルが入っている行 (1始まり)
const DATA_ROW   = 3;   // データ開始行
const CAT_COL    = 0;   // A列: 大分類
const ITEM_COL   = 1;   // B列: 勘定科目
const DATA_COL   = 2;   // C列〜: 月次数値

// サマリ行キーワード
const SUMMARY_KEYS = {
  revenue:     "売上高合計",
  cogs:        "売上原価合計",
  grossProfit: "売上総利益",
  personnel:   "人件費合計",
  sgaTotal:    "販売費及び一般管理費合計",
  opProfit:    "営業利益",
  ordProfit:   "経常利益",
  netProfit:   "当期純利益",
};

// 売上部門カテゴリ (A列の値)
const REVENUE_DEPT_PATTERNS = ["直接雇用", "派遣", "TSUNAGITE", "TSUNAGITEプラットフォーム"];

// コスト系アイテムキーワード (これらを含む行は費用として除外)
const COST_ITEM_PATTERNS = [
  "原価", "人件費", "役員報酬", "給料賃金", "賞与", "法定福利費",
  "福利厚生費", "研修採用費", "業務委託料", "荷造運賃", "広告宣伝費",
  "接待交際費", "旅費交通費", "通信費", "修繕費", "水道光熱費",
  "備品", "消耗品", "車両費", "リース料", "地代家賃", "保険料",
  "租税公課", "支払手数料", "支払報酬", "会議費", "新聞図書費",
  "減価償却費", "諸会費", "長期前払費用", "雑費", "受取利息", "雑収入",
  "支払利息", "雑損失", "法人税",
];

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  const result = getPLData();
  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// MAIN
// ============================================================
function getPLData() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return { error: `シート "${SHEET_NAME}" が見つかりません` };

  const values = sheet.getDataRange().getValues();

  // ── 月ヘッダーを抽出 (月ラベルのみに絞る) ──────────────
  const headerRow = values[HEADER_ROW - 1];
  const months = [];
  for (let c = DATA_COL; c < headerRow.length; c++) {
    const label = normalizeMonthLabel(headerRow[c]);
    if (label) months.push({ col: c, label });
  }
  const monthLabels = months.map(m => m.label);
  const n = months.length;

  // ── サマリ行を検索 ───────────────────────────────────────
  const summary = {};
  for (const [key, keyword] of Object.entries(SUMMARY_KEYS)) {
    const row = findRowByKeyword(values, keyword, [CAT_COL, ITEM_COL]);
    summary[key] = row
      ? months.map(m => toNum(row[m.col]))
      : new Array(n).fill(null);
  }

  // ── 部門別売上集計 (キーワードで判定) ────────────────────
  const revenueByDept = {};
  let lastCat = "";
  for (let r = DATA_ROW - 1; r < values.length; r++) {
    const row  = values[r];
    const cat  = row[CAT_COL] ? String(row[CAT_COL]).trim() : lastCat;
    const item = row[ITEM_COL] ? String(row[ITEM_COL]).trim() : "";
    if (cat) lastCat = cat;
    if (!item) continue;

    // サマリ行はスキップ
    const isSummaryRow = Object.values(SUMMARY_KEYS).some(kw => item.includes(kw) || cat.includes(kw));
    if (isSummaryRow) continue;

    // 売上部門 かつ コストアイテムでない行
    if (!isRevenueDept(cat)) continue;
    if (isCostItem(item))    continue;

    if (!revenueByDept[cat]) revenueByDept[cat] = new Array(n).fill(0);
    months.forEach((m, i) => revenueByDept[cat][i] += toNum(row[m.col]));
  }

  // ── 費用内訳 ─────────────────────────────────────────────
  const costBreakdown = {
    cogs:      summary.cogs,
    personnel: summary.personnel,
    otherSga:  summary.sgaTotal && summary.personnel
      ? summary.sgaTotal.map((v, i) => (v || 0) - (summary.personnel[i] || 0))
      : new Array(n).fill(null),
  };

  return {
    updatedAt:    new Date().toISOString(),
    months:       monthLabels,
    summary,
    revenueByDept,
    costBreakdown,
    _debug: {
      monthsFound: monthLabels,
      rowCount:    values.length,
      deptKeys:    Object.keys(revenueByDept),
    },
  };
}

// ============================================================
// UTILS
// ============================================================
function isRevenueDept(cat) {
  return REVENUE_DEPT_PATTERNS.some(p => cat.includes(p)) || cat === "その他";
}

function isCostItem(item) {
  return COST_ITEM_PATTERNS.some(p => item.includes(p));
}

function findRowByKeyword(values, keyword, cols) {
  for (const row of values) {
    for (const c of cols) {
      if (String(row[c] || "").includes(keyword)) return row;
    }
  }
  return null;
}

function toNum(v) {
  if (typeof v === "number") return v;
  return parseFloat(String(v).replace(/[,¥\s]/g, "")) || 0;
}

// 月ラベルのみを返す。月でない値（合計、名前など）はnullを返す
function normalizeMonthLabel(v) {
  if (v instanceof Date) {
    const y = v.getFullYear();
    const m = String(v.getMonth() + 1).padStart(2, "0");
    return `${y}-${m}`;
  }
  const s = String(v).trim();
  // "2025/11" or "2025-11"
  const full = s.match(/^(\d{4})[\/\-](\d{1,2})$/);
  if (full) return `${full[1]}-${full[2].padStart(2, "0")}`;
  // "11月" のような短い月表記
  const short = s.match(/^(\d{1,2})月$/);
  if (short) return short[0];
  // 上記以外（"合計", "その他", 人名など）は除外
  return null;
}
