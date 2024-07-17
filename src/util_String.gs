function convertLoopRange(range) {
  // コロンで区切って範囲の部分を取得
  var parts = range.split(":");

  // 開始セルと終了セルを取得
  var startCell = parts[0];
  var endCell = parts[1];

  // 開始セルの列を取得
  var startColumn = startCell.match(/[A-Z]+/)[0];
  // 終了セルの行番号を取得
  var endRow = endCell.match(/[0-9]+/)[0];

  // 開始セルの行番号を取得
  var startRow = startCell.match(/[0-9]+/)[0];

  // 結果の範囲を作成
  var newRange = startColumn + startRow + ":" + startColumn + endRow;

  return newRange;
}

function capitalizeFirstLetter(input) {
  return input.charAt(0).toLowerCase() + input.slice(1);
}

function countNonEmptyElements(array) {
  var count = 0;
  for (var i = 0; i < array.length; i++) {
    if (array[i] !== "") {
      count++;
    }
  }
  return count;
}

function countNonNullElements(array) {
  var count = 0;
  for (var i = 0; i < array.length; i++) {
    if (array[i].toString().toLowerCase() != "null"){
      count++;
    }
  }
  return count;
}
