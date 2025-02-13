function getColumnIndex(column) {
  let columnIndex = 0;
  for (let i = 0; i < column.length; i++) {
    columnIndex = columnIndex * 26 + (column.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
  }
  return columnIndex;
}

function getColumnCount(range) {
  // コロンで区切って範囲の部分を取得
  var parts = range.split(":");

  // 開始セルと終了セルを取得
  var startCell = parts[0];
  var endCell = parts[1];

  // 開始セルの列と終了セルの列を取得
  var startColumn = startCell.match(/[A-Z]+/)[0];
  var endColumn = endCell.match(/[A-Z]+/)[0];

  // 列をインデックスに変換
  var startColumnIndex = getColumnIndex(startColumn);
  var endColumnIndex = getColumnIndex(endColumn);

  // カラム数を計算
  var columnCount = endColumnIndex - startColumnIndex + 1;

  return columnCount;
}

function findNonEmptyCell(cell) {
  
  // アクティブなセルが属するシートを取得
  var sheet = cell.getSheet();

  // 現在のセルの行と列を取得
  var row = cell.getRow();
  var column = cell.getColumn();
  
  // 空でないセルが見つかるまでループ
  while (row > 1) { // 1行目までに制限
    var value = sheet.getRange(row, column).getValue();
    
    if (value !== "") {
      return value;
    }
    
    // 一つ上の行に移動
    row--;
  }
  
  Logger.log("★★★★★★★★★★★★★★★★★No non-empty cell found above the current cell.");
  return null;
}


function addSheetToActiveSpreadsheet(spreadsheet,sheetName) {
  var existingSheet = spreadsheet.getSheetByName(sheetName);

  if (existingSheet) {
    var ui = SpreadsheetApp.getUi();
    // var response = ui.alert(
    //   'シートの削除',
    //   'シート "NewSheet" は既に存在します。削除しますか？',
    //   ui.ButtonSet.YES_NO
    // );

    // if (response == ui.Button.YES) {
      // シートを削除
      spreadsheet.deleteSheet(existingSheet);
      // 新しいシートを作成
      var newSheet = spreadsheet.insertSheet(sheetName);
      return newSheet;
    // } else {
    //   return null;
    // }
  } else {
    // 新しいシートを作成
    var newSheet = spreadsheet.insertSheet(sheetName);
    return newSheet;
  }
}

