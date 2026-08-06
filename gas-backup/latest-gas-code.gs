/**
 * KPI Dashboard — Google Apps Script v2
 * 仕様確定版：全シート・行・列マッピング完全実装
 *
 * ✅ 修正内容：
 * - 業種別支援人数：B4:M19 → B4:N19 (10月データ追加)
 * - 業種別紹介人数：B24:M39 → B24:N39 (10月データ追加)
 *
 * ✅ デプロイ URL:
 * https://script.google.com/macros/s/AKfycbw5Oqel437ID97nr5JZGVzE0dYvxpfDYqwVNkcmT7yKGyftqhL6bSxFEnFp-3RIDQ70mA/exec
 */

// ============================================================
// ENTRY POINT
// ============================================================
function doGet(e) {
  try {
    const result = getMultiSheetData();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        error: err.toString(),
        stack: err.stack,
        timestamp: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// MAIN DATA ORCHESTRATOR
// ============================================================
function getMultiSheetData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const result = {
    updatedAt: new Date().toISOString(),
    sheets: {}
  };

  // ========== 会計タブ ==========
  // PL：実績+見込
  const plActualSheet = ss.getSheetByName("実績+見込PL");
  if (plActualSheet) {
    result.sheets["PL_actual"] = extractPLSheet(plActualSheet, "実績+見込PL");
  }

  // PL：予算
  const plBudgetSheet = ss.getSheetByName("修正後予算PL");
  if (plBudgetSheet) {
    result.sheets["PL_budget"] = extractPLSheet(plBudgetSheet, "修正後予算PL");
  }

  // CF
  const cfSheet = ss.getSheetByName("実績+見込CF");
  if (cfSheet) {
    result.sheets["CF"] = extractCFSheet(cfSheet);
  }

  // セグメント PL：直接雇用
  const directSheet = ss.getSheetByName("【経営管理表IP】PL：直接雇用");
  if (directSheet) {
    result.sheets["PL_direct"] = extractSegmentPLSheet(directSheet, "【経営管理表IP】PL：直接雇用");
  }

  // セグメント PL：派遣
  const dispatchSheet = ss.getSheetByName("【経営管理表IP】PL：派遣");
  if (dispatchSheet) {
    result.sheets["PL_dispatch"] = extractSegmentPLSheet(dispatchSheet, "【経営管理表IP】PL：派遣");
  }

  // ========== 事業指標タブ ==========
  const indicatorSheet = ss.getSheetByName("事業指標サマリ");
  if (indicatorSheet) {
    result.sheets["genre_support"] = extractGenreSupportData(indicatorSheet);
    result.sheets["genre_intro"] = extractGenreIntroData(indicatorSheet);
    result.sheets["dispatch_variables"] = extractDispatchVariables(indicatorSheet);
  }

  // ========== 昨対比タブ ==========
  if (indicatorSheet) {
    result.sheets["yoy_comparison"] = extractYoYComparisonData(indicatorSheet);
  }

  // ========== パイプラインタブ ==========
  result.sheets["pipeline_summary"] = extractPipelineSummary(ss);
  result.sheets["pipeline_02_solution"] = extractPipelineStageData(ss, "02 解決策合意", 5); // 上位5社
  result.sheets["pipeline_03_pricing"] = extractPipelineStageData(ss, "03 価格合意", null); // 全案件

  return result;
}

// ============================================================
// 会計タブ：PL シート抽出（実績+見込PL / 修正後予算PL）
// ============================================================
// 列マッピング：C=11月, D=12月, ..., M=10月
function extractPLSheet(sheet, sheetName) {
  const values = sheet.getDataRange().getValues();

  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];
  const monthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]; // C=2, D=3, ..., M=13

  const extractRow = (rowIdx) => {
    if (rowIdx >= values.length) return null;
    const row = values[rowIdx];
    return monthCols.map(col => toNum(row[col]));
  };

  const rows = [
    { item: "売上高合計", values: extractRow(28) },      // 行29（0-indexed: 28）
    { item: "売上総利益", values: extractRow(55) },      // 行56（0-indexed: 55）
    { item: "営業利益", values: extractRow(88) },        // 行89（0-indexed: 88）
  ].filter(r => r.values !== null);

  return {
    sheetName: sheetName,
    months: months,
    rows: rows,
    debug: {
      totalRows: values.length,
      extractedCount: rows.length
    }
  };
}

// ============================================================
// 会計タブ：セグメント PL シート抽出（直接雇用 / 派遣）
// ============================================================
// 列マッピング：C=11月, D=12月, ..., M=10月
function extractSegmentPLSheet(sheet, sheetName) {
  const values = sheet.getDataRange().getValues();

  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];
  const monthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]; // C=2, D=3, ..., M=13

  const extractRow = (rowIdx) => {
    if (rowIdx >= values.length) return null;
    const row = values[rowIdx];
    return monthCols.map(col => toNum(row[col]));
  };

  const rows = [
    { item: "売上高合計", values: extractRow(28) },           // 行29
    { item: "売上総利益", values: extractRow(55) },           // 行56
    { item: "営業利益", values: extractRow(88) },             // 行89
    { item: "営業利益(本社費用割賦後)", values: extractRow(91) }, // 行92
  ].filter(r => r.values !== null);

  return {
    sheetName: sheetName,
    months: months,
    rows: rows,
    debug: {
      totalRows: values.length,
      extractedCount: rows.length
    }
  };
}

// ============================================================
// 会計タブ：CF シート抽出
// ============================================================
// 列マッピング：B=11月, C=12月, ..., M=10月
function extractCFSheet(sheet) {
  const values = sheet.getDataRange().getValues();

  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];
  const monthCols = [1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]; // B=1, C=2, ..., M=12

  const extractRow = (rowIdx) => {
    if (rowIdx >= values.length) return null;
    const row = values[rowIdx];
    return monthCols.map(col => toNum(row[col]));
  };

  const rows = [
    { item: "現預金残高", values: extractRow(32) },      // 行33（0-indexed: 32）
    { item: "バーンレート", values: extractRow(40) },    // 行41
    { item: "月末時点ランウェイ", values: extractRow(41) }, // 行42
  ].filter(r => r.values !== null);

  return {
    sheetName: "実績+見込CF",
    months: months,
    rows: rows,
    debug: {
      totalRows: values.length,
      extractedCount: rows.length
    }
  };
}

// ============================================================
// 事業指標タブ：業種別支援人数
// ============================================================
// 業種ラベル：B4-B19
// データ：C4-N19（11月～10月） ✅ FIXED: M19 → N19
function extractGenreSupportData(sheet) {
  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];

  try {
    const data = sheet.getRange('B4:N19').getValues();

    const rows = data.map((row, idx) => ({
      genre: String(row[0] || ''),
      values: row.slice(1).map(v => toNum(v))
    }));

    return {
      sheetName: "事業指標サマリ",
      dataType: "genre_support",
      months: months,
      rows: rows,
      debug: {
        rowCount: rows.length
      }
    };
  } catch (err) {
    return { error: 'Failed to extract genre support data: ' + err.toString() };
  }
}

// ============================================================
// 事業指標タブ：業種別紹介人数
// ============================================================
// 業種ラベル：B24-B39
// データ：C24-N39（11月～10月） ✅ FIXED: M39 → N39
function extractGenreIntroData(sheet) {
  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];

  try {
    const data = sheet.getRange('B24:N39').getValues();

    const rows = data.map((row, idx) => ({
      genre: String(row[0] || ''),
      values: row.slice(1).map(v => toNum(v))
    }));

    return {
      sheetName: "事業指標サマリ",
      dataType: "genre_intro",
      months: months,
      rows: rows,
      debug: {
        rowCount: rows.length
      }
    };
  } catch (err) {
    return { error: 'Failed to extract genre intro data: ' + err.toString() };
  }
}

// ============================================================
// 事業指標タブ：派遣重要変数
// ============================================================
// 総時間：行105, 平均時給：行107, 総人数：行108, 平均時間：行109
// 列マッピング：C=11月, D=12月, ..., M=10月
function extractDispatchVariables(sheet) {
  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];
  const monthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]; // C=2, D=3, ..., M=13

  try {
    const values = sheet.getDataRange().getValues();

    const extractRow = (rowIdx) => {
      if (rowIdx >= values.length) return null;
      const row = values[rowIdx];
      return monthCols.map(col => toNum(row[col]));
    };

    const rows = [
      { item: "総時間", values: extractRow(104) },      // 行105（0-indexed: 104）
      { item: "平均時給", values: extractRow(106) },    // 行107
      { item: "総人数", values: extractRow(107) },      // 行108
      { item: "平均時間", values: extractRow(108) },    // 行109
    ].filter(r => r.values !== null);

    return {
      sheetName: "事業指標サマリ",
      dataType: "dispatch_variables",
      months: months,
      rows: rows,
      debug: {
        extractedCount: rows.length
      }
    };
  } catch (err) {
    return { error: 'Failed to extract dispatch variables: ' + err.toString() };
  }
}

// ============================================================
// 昨対比タブ：年間比較データ
// ============================================================
// 紹介人数（前年）：行115, 支援人数（前年）：行116
// 派遣総人数（前年）：行117, 派遣総時間（前年）：行118
// 列マッピング：C=11月, D=12月, ..., M=10月
function extractYoYComparisonData(sheet) {
  const months = ['11月', '12月', '1月', '2月', '3月', '4月', '5月', '6月', '7月', '8月', '9月', '10月'];
  const monthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]; // C=2, D=3, ..., M=13

  try {
    const values = sheet.getDataRange().getValues();

    const extractRow = (rowIdx) => {
      if (rowIdx >= values.length) return null;
      const row = values[rowIdx];
      return monthCols.map(col => toNum(row[col]));
    };

    const result = {
      sheetName: "事業指標サマリ",
      dataType: "yoy_comparison",
      months: months,
      previousYear: {
        intro: extractRow(114),           // 行115（紹介人数前年）
        support: extractRow(115),         // 行116（支援人数前年）
        dispatch_people: extractRow(116), // 行117（派遣総人数前年）
        dispatch_hours: extractRow(117)   // 行118（派遣総時間前年）
      },
      debug: {
        rowsExtracted: [115, 116, 117, 118]
      }
    };

    return result;
  } catch (err) {
    return { error: 'Failed to extract YoY comparison data: ' + err.toString() };
  }
}

// ============================================================
// パイプラインタブ：全体ファネル別金額（サマリ）
// ============================================================
// シート：【Hubspot IP】CRM集計
// 9行目：01 課題合意, 10行目：02 解決策合意, 11行目：03 価格合意
// 列：C-I（2026年4月～12月）
function extractPipelineSummary(ss) {
  const sheet = ss.getSheetByName("【Hubspot IP】CRM集計");
  if (!sheet) {
    return { error: "【Hubspot IP】CRM集計シートが見つかりません" };
  }

  try {
    const months = ["2026-04", "2026-05", "2026-06", "2026-07", "2026-08", "2026-09", "2026-10", "2026-11", "2026-12"];
    const monthCols = [2, 3, 4, 5, 6, 7, 8, 9, 10]; // C=2, D=3, ..., J=10

    const values = sheet.getDataRange().getValues();

    const extractRow = (rowIdx) => {
      if (rowIdx >= values.length) return null;
      const row = values[rowIdx];
      return monthCols.map(col => toNum(row[col]));
    };

    const result = {
      sheetName: "【Hubspot IP】CRM集計",
      dataType: "pipeline_summary",
      months: months,
      stages: {
        "01_awareness": extractRow(8),      // 行9（0-indexed: 8）
        "02_solution": extractRow(9),       // 行10
        "03_pricing": extractRow(10)        // 行11
      }
    };

    return result;
  } catch (err) {
    return { error: 'Failed to extract pipeline summary: ' + err.toString() };
  }
}

// ============================================================
// パイプラインタブ：ステージ別案件一覧（金額・人数）
// ============================================================
// シート：【Hubspot IP】CRM一覧(金額) と 【Hubspot IP】CRM一覧(見込み人数)
// B列：取引ステージ, C列：取引先名, D列～：月毎の金額/人数
// D列=2026年7月, E列=2026年8月, ..., I列=2026年12月
function extractPipelineStageData(ss, stageName, topN) {
  const amountSheet = ss.getSheetByName("【Hubspot IP】CRM一覧(金額)");
  const peopleSheet = ss.getSheetByName("【Hubspot IP】CRM一覧(見込み人数)");

  if (!amountSheet || !peopleSheet) {
    return { error: "パイプラインシートが見つかりません" };
  }

  try {
    const months = ["2026-07", "2026-08", "2026-09", "2026-10", "2026-11", "2026-12"];
    const dataStartCol = 3; // D列（0-indexed: 3）
    const monthCount = 6; // D-I列

    const amountValues = amountSheet.getDataRange().getValues();
    const peopleValues = peopleSheet.getDataRange().getValues();

    // ステージでフィルタしたデータを集計
    const monthlyData = {};
    for (let m = 0; m < monthCount; m++) {
      monthlyData[months[m]] = [];
    }

    // 金額シートから該当ステージのデータを抽出
    for (let i = 1; i < amountValues.length; i++) {
      const row = amountValues[i];
      const stage = String(row[1] || '').trim(); // B列
      const company = String(row[2] || '').trim(); // C列

      if (stage === stageName && company) {
        // 同じ行のデータを人数シートから取得
        const peopleRow = i < peopleValues.length ? peopleValues[i] : [];

        // 月毎のデータを集計
        for (let m = 0; m < monthCount; m++) {
          const col = dataStartCol + m;
          const amount = toNum(row[col]);
          const people = toNum(peopleRow[col]);

          if (amount > 0 || people > 0) {
            monthlyData[months[m]].push({
              company: company,
              amount: amount,
              people: people
            });
          }
        }
      }
    }

    // 月毎に金額でソートして、topN でカット
    for (let month in monthlyData) {
      monthlyData[month].sort((a, b) => b.amount - a.amount);
      if (topN) {
        monthlyData[month] = monthlyData[month].slice(0, topN);
      }
    }

    return {
      sheetName: "【Hubspot IP】CRM一覧(金額/人数)",
      dataType: "pipeline_stage",
      stageName: stageName,
      months: months,
      data: monthlyData,
      topN: topN
    };
  } catch (err) {
    return { error: 'Failed to extract pipeline stage data: ' + err.toString() };
  }
}

// ============================================================
// UTILITIES
// ============================================================

function toNum(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  const parsed = parseFloat(String(v).replace(/[,¥\s]/g, ""));
  return isNaN(parsed) ? 0 : parsed;
}

// スタブ関数（必要に応じて実装）
function sendDailyPLAlert() {
  Logger.log('sendDailyPLAlert: スタブ関数');
}

function setupDailyAlert() {
  Logger.log('setupDailyAlert: スタブ関数');
}
