/**
 * KPI ダッシュボード — Google Apps Script
 * データ定義シートを読んで、動的にデータを取得・返す
 */

const CONFIG = {
  SPREADSHEET_ID: '17YWeqS08Deu3HVmmJSjRLTHXHLyQL-ZtnIVqLEkiEn8',
  DATA_DEFINITION_SHEET: 'データ定義'
};

// ============================================================
// ENTRY POINT（ウェブアプリ用）
// ============================================================
function doGet(e) {
  try {
    const result = getAllDashboardData();
    return ContentService
      .createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ error: err.toString(), stack: err.stack }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// MAIN：データ定義から全データを取得
// ============================================================
function getAllDashboardData() {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const defSheet = ss.getSheetByName(CONFIG.DATA_DEFINITION_SHEET);

  if (!defSheet) {
    return { error: 'データ定義シートが見つかりません' };
  }

  const result = {
    updatedAt: new Date().toISOString(),
    sheets: {}
  };

  try {
    const defValues = defSheet.getDataRange().getValues();

    // ヘッダー行をスキップ（1行目）
    for (let i = 1; i < defValues.length; i++) {
      const row = defValues[i];
      const itemName = String(row[0] || '').trim();
      const sheetName = String(row[1] || '').trim();
      const cellRange = String(row[2] || '').trim();
      const dataType = String(row[3] || '').trim();

      if (!itemName || !sheetName || !cellRange) {
        continue; // 空行をスキップ
      }

      try {
        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
          result.sheets[itemName] = { error: `シート「${sheetName}」が見つかりません` };
          continue;
        }

        const rangeData = sheet.getRange(cellRange).getValues();
        result.sheets[itemName] = {
          sheetName: sheetName,
          cellRange: cellRange,
          dataType: dataType,
          data: rangeData
        };

      } catch (err) {
        result.sheets[itemName] = { error: `データ取得エラー: ${err.toString()}` };
      }
    }

  } catch (err) {
    return { error: 'データ定義読み込みエラー: ' + err.toString() };
  }

  return result;
}

// ============================================================
// ユーティリティ
// ============================================================

function toNum(v) {
  if (typeof v === "number") return v;
  if (!v) return 0;
  return parseFloat(String(v).replace(/[,¥\s]/g, "")) || 0;
}
