// ヘッダークラスの定義
class specManager {
  constructor() {
    Logger.log("pearent");
    // 現在のアクティブなスプレッドシートを取得
    this.spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    this.sheet = this.spreadsheet.getActiveSheet();

    
    // アクティブなシートから現在の選択範囲（アクティブなセルの範囲）を取得
    this.activeRange = this.spreadsheet.getActiveRange();

    this.superHeaderRange = this.sheet.getRange("c2");

    // 選択範囲の行と列の開始位置とサイズを取得
    this.startRow = this.activeRange.getRow();
    this.startColumn = this.activeRange.getColumn();
    this.numRows = this.activeRange.getNumRows();
    this.numColumns = this.activeRange.getNumColumns();
    
    // 開始セル
    this.startRange = this.sheet.getRange(this.startRow, this.startColumn)
    var address = this.startRange.getA1Notation();

    // 選択範囲の値を取得（必要であれば）
    this.values = this.activeRange.getValues();
    this.targetRange = this.sheet.getRange(this.startRow,this.startColumn,this.numRows, COLUMNS_Enum - this.startColumn + 1)

    // アノテーション範囲
    this.anotationRng = this.sheet.getRange("AA9:AC1000")
  }

  getStartRange(){
    return this.startRange
  }

  getPearentRange(){
    return this.startRange.offset(-1,-1)
  }

  gerTargetRange() {
    return this.targetRange
  }


  getPosPhysics() {
    return COLUMNS_Physics - this.startColumn
  }

  getPosType() {
    return COLUMNS_Type - this.startColumn
  }

  getPosRequired() {
    return COLUMNS_Required - this.startColumn
  }

  getPosMin() {
    return COLUMNS_Min - this.startColumn
  }

  getPosMax() {
    return COLUMNS_Max - this.startColumn
  }
  
  getPosEnum() {
    return COLUMNS_Enum - this.startColumn
  }

  getAnotationRng(){
    return this.anotationRng
  }

  getClassName(){
      return this.getPearentRange().offset(0,this.getPosType() + 1 ).getValue();

  }
}
