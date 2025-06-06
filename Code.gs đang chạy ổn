// Code.gs
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Dashboard Khách Hàng');
}

const HIDDEN_SHEETS = ['Hợp Đồng', 'TEMPLATE', 'TRANG_CHU', 'LOG_DOI_TEN'];

function getCustomerList() {
  return SpreadsheetApp.getActiveSpreadsheet()
    .getSheets()
    .map(sheet => sheet.getName())
    .filter(name => !HIDDEN_SHEETS.includes(name));
}

function getCustomerDataJSON(sheetName) {
  if (HIDDEN_SHEETS.includes(sheetName)) {
    return {error: 'Sheet bị ẩn hoặc không hợp lệ'};
  }

  const cache = CacheService.getScriptCache();
  const cacheKey = 'customerData_' + sheetName;
  let cached = cache.get(cacheKey);

  if (cached) {
    try {
      return JSON.parse(cached);
    } catch (e) {
      // lỗi parse thì bỏ qua
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {error: 'Sheet không tồn tại'};

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  let lastTextRow = 0;
  const values = sheet.getRange(1, 2, lastRow, lastColumn - 1).getDisplayValues();
  for (let r = values.length - 1; r >= 0; r--) {
    if (values[r].some(cell => cell !== '')) {
      lastTextRow = r + 1;
      break;
    }
  }
  if (lastTextRow === 0) lastTextRow = 1;

  const dataRange = sheet.getRange(1, 1, lastTextRow, lastColumn);
  const dataValues = dataRange.getDisplayValues();
  const backgrounds = dataRange.getBackgrounds();
  const fontColors = dataRange.getFontColors();
  const alignments = dataRange.getHorizontalAlignments();

  let colWidths = [];
  for (let c = 2; c <= lastColumn; c++) {
    colWidths.push(sheet.getColumnWidth(c));
  }

  let data = [];
  for (let i = 0; i < dataValues.length; i++) {
    let row = [];
    for (let j = 1; j < dataValues[i].length; j++) {
      row.push({
        value: dataValues[i][j],
        bg: backgrounds[i][j],
        color: fontColors[i][j],
        align: alignments[i][j]
      });
    }
    data.push(row);
  }

  const result = {
    data: data,
    colWidths: colWidths,
  };

  const now = new Date();
  const nextMidnight = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1);
  const secondsToMidnight = Math.floor((nextMidnight - now) / 1000);

  cache.put(cacheKey, JSON.stringify(result), secondsToMidnight);
  return result;
}

function clearCacheForCustomer(sheetName) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'customerData_' + sheetName;
  cache.remove(cacheKey);
  return true;
}
