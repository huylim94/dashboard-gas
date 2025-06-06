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
      console.error("Lỗi khi parse cache, bỏ qua cache:", e);
    }
  }

  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  if (!sheet) return {error: 'Sheet không tồn tại'};

  const lastColumn = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();

  // Tìm hàng cuối cùng có dữ liệu để tối ưu phạm vi lấy
  let lastTextRow = 0;
  // Lấy toàn bộ dữ liệu từ cột A trở đi để tìm hàng cuối cùng có nội dung
  const tempValuesForLastRowCheck = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  for (let r = tempValuesForLastRowCheck.length - 1; r >= 0; r--) {
    if (tempValuesForLastRowCheck[r].some(cell => cell !== '')) {
      lastTextRow = r + 1;
      break;
    }
  }
  if (lastTextRow === 0) lastTextRow = 1; // Mặc định là 1 nếu không có văn bản

  // Lấy tất cả dữ liệu, style, gộp ô trong cùng một dải ô chính xác
  const dataRange = sheet.getRange(1, 1, lastTextRow, lastColumn);
  const dataValues = dataRange.getDisplayValues();
  const backgrounds = dataRange.getBackgrounds();
  const fontColors = dataRange.getFontColors();
  const alignments = dataRange.getHorizontalAlignments();
  const fontWeights = dataRange.getFontWeights(); // Thêm để lấy fontWeight
  const fontStyles = dataRange.getFontStyles();   // Thêm để lấy fontStyle
  const textDecorations = dataRange.getTextDecorations(); // Thêm để lấy textDecoration
  const fontSizes = dataRange.getFontSizes();     // Thêm để lấy fontSize
  const wraps = dataRange.getWraps();             // Thêm để lấy wrap (tự động xuống dòng)

  const mergedRanges = sheet.getMergedRanges();

  // --- Xử lý thông tin gộp ô ---
  const mergedCellsInfo = new Map(); // Key: "row,col" (0-indexed), Value: {rowspan, colspan, isStart}
  mergedRanges.forEach(range => {
    const startRow = range.getRow() - 1; // Convert to 0-indexed
    const startCol = range.getColumn() - 1; // Convert to 0-indexed
    const numRows = range.getNumRows();
    const numCols = range.getNumColumns();

    for (let r = startRow; r < startRow + numRows; r++) {
      for (let c = startCol; c < startCol + numCols; c++) {
        mergedCellsInfo.set(`${r},${c}`, {
          isStart: (r === startRow && c === startCol), // Chỉ ô đầu tiên của vùng gộp mới là start
          rowspan: numRows,
          colspan: numCols
        });
      }
    }
  });
  // --- Kết thúc xử lý gộp ô ---

  let colWidths = [];
  for (let c = 1; c <= lastColumn; c++) { // Lấy độ rộng cho tất cả các cột
    colWidths.push(sheet.getColumnWidth(c));
  }

  let data = [];
  for (let i = 0; i < dataValues.length; i++) { // i là chỉ số hàng (0-indexed)
    let row = [];
    for (let j = 0; j < dataValues[i].length; j++) { // j là chỉ số cột (0-indexed)
      const cellInfo = mergedCellsInfo.get(`${i},${j}`); // Kiểm tra xem ô này có thuộc vùng gộp không

      const cellData = {
        value: dataValues[i][j],
        bg: backgrounds[i][j],
        color: fontColors[i][j],
        align: alignments[i][j],
        fontWeight: fontWeights[i][j],     // Thêm
        fontStyle: fontStyles[i][j],       // Thêm
        textDecoration: textDecorations[i][j], // Thêm
        fontSize: fontSizes[i][j],         // Thêm
        wrap: wraps[i][j]                  // Thêm
      };

      if (cellInfo) { // Nếu ô này thuộc một vùng gộp
        if (cellInfo.isStart) { // Nếu đây là ô bắt đầu của vùng gộp
          cellData.rowspan = cellInfo.rowspan;
          cellData.colspan = cellInfo.colspan;
        } else { // Nếu đây là ô con bị che bởi vùng gộp
          cellData.skip = true; // Đánh dấu để bỏ qua render ở HTML
        }
      }
      row.push(cellData);
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
