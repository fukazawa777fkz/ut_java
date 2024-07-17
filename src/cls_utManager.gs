// ヘッダークラスの定義
class utManager {
  constructor() {
    // Logger.log("utManager:constructor------>")
    // 現在のアクティブなスプレッドシートを取得（excelで言うところのWorkbook）
    this.ss = SpreadsheetApp.getActiveSpreadsheet();
    // Logger.log("this.ss.getName() :" + this.ss.getName())

    // 現在のアクティブなシートを取得
    this.sheet = this.ss.getActiveSheet();
    // Logger.log("this.sheet.getName() :" + this.sheet.getName())
    
    // アクティブなシートから現在の選択範囲（アクティブなセルの範囲）を取得
    this.activeRange = this.ss.getActiveRange();
    this.startAdderss = this.activeRange.getA1Notation()
    // Logger.log("this.startAdderss :" + this.startAdderss)

    // 実行範囲
    this.loopAddress = convertLoopRange(this.startAdderss)
    // Logger.log("this.loopAddress :" + this.loopAddress)

    // 入力カラムMAX
    this.fieldCount = getColumnCount(this.startAdderss)
    // Logger.log("this.fieldCount :" + this.fieldCount)

    // 期待値カラム_HTTPステータス
    this.offsetHttpStatus =  this.fieldCount
    // Logger.log("this.offsetHttpStatus :" + this.offsetHttpStatus)

    // 期待値カラム_HTML名
    this.offsetHtml =  this.fieldCount + 1
    // Logger.log("this.offsetHtml :" + this.offsetHtml)

    // 期待値カラム_errorCode
    this.offsetErrorCode =  this.fieldCount + 2
    // Logger.log("this.offsetErrorCode :" + this.offsetErrorCode)

    // 開始セル
    this.startRow = this.activeRange.getRow();
    this.startColumn = this.activeRange.getColumn();
    this.startRange = this.sheet.getRange(this.startRow, this.startColumn)
    // var address = this.startRange.getA1Notation();

    // エンドポイント
    this.endpoint = this.startRange.offset(-7, -1)
    // Logger.log("endpoint:" + this.endpoint.getValue())

    // メソッド
    this.method = this.startRange.offset(-6, -1)
    // Logger.log("method:" + this.method.getValue())

    // 機能名
    this.functionName = this.startRange.offset(-5, -1)
    // Logger.log("function:" + this.functionName.getValue())

    // 入力フォーム
    this.inputForm = this.startRange.offset(-4, -1)
    // Logger.log("form:" + this.inputForm.getValue())

    // Logger.log("<------utManager:constructor")

  }

  gerTargetRange() {
    return this.sheet.getRange(this.loopAddress)
  }

  geEndpoint() {
    return this.endpoint.getValue()
  }

  getFunctionName(){
    return this.functionName.getValue()
  }

  getTetsFuntionName(rowIndex) {
    var testItem = this.startRange.offset(rowIndex, -1).getValue()
    var testcase = findNonEmptyCell(this.startRange.offset(rowIndex, -2))
    return testcase + "_" + testItem

  }

  getMesod(){
    return this.method.getValue()
  }

  gerCurrentFieldValues(rowIndex){
    var r1 = this.startRange.offset(rowIndex, 0).getA1Notation()
    var r2 = this.startRange.offset(rowIndex, this.fieldCount -1).getA1Notation() 
    return this.sheet.getRange(r1 + ":" + r2).getValues()
  }

  gerCurrentErrors(rowIndex){
    var r1 = this.startRange.offset(rowIndex, this.offsetErrorCode)
    var r2 = this.startRange.offset(rowIndex, (this.offsetErrorCode + this.fieldCount) -1)
    var currentErrorFeildRange = this.sheet.getRange(r1.getA1Notation() + ":" + r2.getA1Notation())
    var values = currentErrorFeildRange.getValues()
    var ret = ""
    for (var i = 0; i < values[0].length; i++) {
      if (values[0][i] != "") {
        if (ret != "") {
          ret = ret + ","
        }
        ret = ret + '"' + r1.offset(-rowIndex -1,i).getValue() + '"'
      }
    }
    return ret
  }

  gerCurrentErrorFieldRange(rowIndex){
    var r1 = this.startRange.offset(rowIndex, this.offsetErrorCode)
    return r1
  }

  getFieldName(columnIndex){
    return this.startRange.offset(-1,columnIndex).getValue()
  }

  // 設定値を取得する
  //  戻り値は２つ。 第２戻り値… 0:値を返却、1:コードとして返却
  getDisplayValue(value) {

    if (value == "昨日") {
      return ["LocalDate.now().minusDays(1).toString()",0]
    }
    if (value == "今日") {
      return [ "LocalDate.now().toString()",0]
    }
    if (value == "明日") {
      return [ "LocalDate.now().plusDays(1).toString()",0]
    }
    if (value == "明後日") {
      return [ "LocalDate.now().plusDays(2).toString()",0]
    }
    if (value == "明々後日") {
      return [ "LocalDate.now().plusDays(3).toString()",0]
    }
    if (value.toString().toLowerCase() === "null") {
      return [ "null",2]
    }
    if (value instanceof Date && !isNaN(value.getTime())) {
      return [ Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd"), 1]
    }

    return [value, 1]
  }

  getReturnHtml(rowIndex){
    return this.startRange.offset(rowIndex,this.offsetHtml).getValue()
  }  

  isNormalTermination(rowIndex){
    var testcase = findNonEmptyCell(this.startRange.offset(rowIndex, -2))
    if (testcase == "正常系") {
      return true
    }
    return false
  }

  getFormTag(){
    // 最初の文字を小文字に変換して返す
    return capitalizeFirstLetter(this.inputForm.getValue())
  }

  getFieldErrors(rowIndex) {
    var formTag = this.getFormTag()
    var fields = this.gerCurrentErrors(rowIndex)
    var ret = '"' + formTag + '",' + fields
    return ret 
  }
  getFieldErrorValues(rowIndex) {
    var r1 = this.startRange.offset(rowIndex, this.offsetErrorCode)
    var r2 = this.startRange.offset(rowIndex, (this.offsetErrorCode + this.fieldCount) -1)
    var currentErrorFeildRange = this.sheet.getRange(r1.getA1Notation() + ":" + r2.getA1Notation())
    return currentErrorFeildRange.getValues()
  }
}
