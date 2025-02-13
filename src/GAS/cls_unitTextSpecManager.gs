class ItemInfo{
  constructor(array,specManager, fileId, value) {
    this.value = value;
    this.name = array[0];
    this.physics = array[specManager.getPosPhysics()];
    this.type = array[specManager.getPosType()];
    this.required = array[specManager.getPosRequired()];
    this.min = array[specManager.getPosMin()];
    this.max = array[specManager.getPosMax()];
    this.enum = array[specManager.getPosEnum()];
    Logger.log(fileId);
    this.fileId = fileId;
  }
}


class unitTextSpecManager  extends specManager{
    constructor(){
      super();
      this.current = 1;

      this.newSheet = addSheetToActiveSpreadsheet(this.spreadsheet, 'NewSheet')
      // this.newSheet = this.spreadsheet.insertSheet('NewSheet');
      this.stdRange = this.newSheet.getRange("e8");
      this.headerRange = this.newSheet.getRange("c2");
      // this.minMaxList = [];  // this.maxListを初期化
      this.minList = [];  // this.maxListを初期化
      this.maxList = [];  // this.maxListを初期化
      this.requiredList = [];  // this.maxListを初期化
      this.enumList = [];  // this.maxListを初期化
      this.defaultList = [];
      var values = this.gerTargetRange().getValues();
      for (var i = 0; i < values.length; i++) {

        if (values[i][this.getPosRequired()] != ""){
          this.addList(this.requiredList, values[i], i, this.getRequiredValue(values[i]));
        }

        if (values[i][this.getPosEnum()] != ""){
          this.addList(this.enumList, values[i], this.getPosEnum());
        }

        this.addList(this.defaultList, values[i],i);
        if ((values[i][this.getPosMin()] != "") && (values[i][this.getPosMax()] != "")){
          this.addList(this.minList, values[i],i, values[i][this.getPosMin()]);
          this.addList(this.maxList, values[i],i, values[i][this.getPosMax()]);
        }

        this.stdRange.offset(-1,i).setValue(values[i][0]);
        this.stdRange.offset(0,i).setValue(values[i][this.getPosPhysics()]);
        this.stdRange.offset(0,i + values.length +2 ).setValue(values[i][this.getPosPhysics()]);
      }

      this.stdRange.offset(-2,+0 ).setValue('入力');
      this.stdRange.offset(-2,values.length +0 ).setValue('期待値');
      this.stdRange.offset(-1,values.length +0 ).setValue('HTTPステータス');
      this.stdRange.offset(-1,values.length +1 ).setValue('HTML名');
      this.stdRange.offset(-1,values.length +2 ).setValue('errorCode');
      
    }

    addList(list, array, fileId, value = "") {

      var itemInof = new ItemInfo(array, this, fileId, value);
      list.push(itemInof);
    }


    writeNormal(){
      this.stdRange.offset(this.current, -2).setValue("正常系");
      this.stdRange.offset(this.current, -1).setValue("最小");
      this.setFeildsValues(TestType.NORMAL,this.minList, 0);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("最大");
      this.setFeildsValues(TestType.NORMAL,this.maxList, 0);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("必須のみ");
      this.setFeildsValues(TestType.NORMAL,this.requiredList, 0, true);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("空文字");
      this.setFeildsDilect("");

      this.current++;

    }

    writeAbnormal(){
      this.stdRange.offset(this.current, -2).setValue("異常系");
      this.stdRange.offset(this.current, -1).setValue("最小");
      this.setFeildsValues(TestType.ABNORMAL,this.minList, -1);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("最大");
      this.setFeildsValues(TestType.ABNORMAL,this.maxList, +1);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("null値");
      this.setFeildsValues(TestType.ABNORMAL,this.requiredList, 0, true, true);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("空文字");
      this.setFeildsDilect("", true);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("半角スペース");
      this.setFeildsDilect(" ", false);

      this.current++;
      this.stdRange.offset(this.current, -1).setValue("全角スペース");
      this.setFeildsDilect("　", false);

    }

    // * 概要：入力値を設定する
    // * param：list 対象リスト
    // * param：valueOffset 異常系の設定をする場合は値を設定しておく
    // * param：isRequired 必須入力項目かどうか（その他の設定をするとき、必須でないものはnullに設定される）
    // * param：isToggle 必須入力項目でないものをnullを設定するようになる

    setFeildsValues(testType, list, valueOffset, isRequired = false, isToggle = false) {
      // nullの指定は、isToggleを元に設定する

      for(var index =0; index < list.length; index++ ){
        var itemInfo = list[index];
        var rngObj = this.stdRange.offset(this.current, itemInfo.fileId);
        if (isToggle == false) {
          if ((itemInfo.type == 'Integer') || (itemInfo.type == 'Int')){
            rngObj.setValue(parseFloat(itemInfo.value + valueOffset));
            var errorCodeType = ErrorCodeType.Max
            if (valueOffset < 0) {
              errorCodeType = ErrorCodeType.Min;
            }
            this.setErrorCode(testType,errorCodeType, rngObj,itemInfo);
          } else if ((itemInfo.type == 'String') ){
            rngObj.setFormula('=REPT("a",' + parseFloat(itemInfo.value + valueOffset) + ')');
            this.setErrorCode(testType,ErrorCodeType.Size, rngObj,itemInfo);
          } else {
            rngObj.setValue(parseFloat(itemInfo.value + valueOffset));
            this.setErrorCode(testType,ErrorCodeType.Size, rngObj,itemInfo);
          }
          if (rngObj.getValue() == "") {
            Logger.log(rngObj.getSheet().getName());
            Logger.log(rngObj.getA1Notation());
            rngObj.setBackground('yellow');
          }
        } else {
          if (itemInfo.type == 'Integer') {
            rngObj.setValue('null');
            this.setErrorCode(testType,ErrorCodeType.NotNull, rngObj,itemInfo);
          }
          else if (itemInfo.type == 'Int') {
            rngObj.setValue('0');
            this.setErrorCode(testType,ErrorCodeType.Invalid, rngObj,itemInfo);
          }
          else if (itemInfo.type == 'String') {
            rngObj.setValue('null');
            this.setErrorCode(testType,ErrorCodeType.NotNull, rngObj,itemInfo);
          } else {
            rngObj.setValue('null');
            this.setErrorCode(testType,ErrorCodeType.NotNull, rngObj,itemInfo);
          }
        }
      }

      // 対象リストになかったフィールドを設定する。
      this.setOtherFeildsValue(list,isRequired,isToggle,testType,valueOffset);
    }


    setErrorCode(testType,errorCodeType, rngObj,itemInfo){
      if (testType == TestType.NORMAL){
        return;
      }

      rngObj.offset(0,this.numRows + FixedFiledNum).setValue(errorCodeType)

    }

    contain(list, physics) {
      for(var index =0; index < list.length; index++ ){
        var itemInfo = list[index];
        if (itemInfo.physics == physics) {
          return true;
        }
      }
      return false;
    }

    // 対象リストになかったフィールドを設定する。
    setOtherFeildsValue(list,isRequired,isToggle, testType,valueOffset){
      for(var index =0; index < this.defaultList.length; index++ ){
        if (this.contain(list, this.defaultList[index].physics) == false) {

          var itemInfo = this.defaultList[index];
          var rngObj = this.stdRange.offset(this.current, itemInfo.fileId);
          if ((isRequired == true) && (isToggle == false)) {
            rngObj.setValue("null");
          } else {
            var value = itemInfo.value;
            var errorFlg = false;
            if (itemInfo.value === "" && itemInfo.min != "") {
              value = itemInfo.min + valueOffset;
              if (value < itemInfo.min) {
                errorFlg = true;
              } else {
                value = itemInfo.min;
              }
            }
            if (itemInfo.value === "" && itemInfo.max != "") {
              value = itemInfo.max + valueOffset;
              if (value > itemInfo.max) {
                errorFlg = true;
              } else {
                value = itemInfo.max;
              }
            }

            if (value === "") {
              value = 2;
            }


            if ((itemInfo.type == 'Integer') || (itemInfo.type == 'Int')){
              rngObj.setValue(value);
              var errorCodeType = ErrorCodeType.Max
              if (valueOffset < 0) {
                errorCodeType = ErrorCodeType.Min;
              }
              if (errorFlg) {
                this.setErrorCode(testType,errorCodeType, rngObj,itemInfo);
              }
            } else if ((itemInfo.type == 'String') ){
              rngObj.setFormula('=REPT("z",' + value + ')');
              // this.setErrorCode(testType,ErrorCodeType.Size, rngObj,itemInfo);
            } else if ((itemInfo.type == 'Date') ){
              // 今日の日付を取得
              var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
              rngObj.setValue(formattedDate);
              // this.setErrorCode(testType,ErrorCodeType.Past, rngObj,itemInfo);
            } else {
              rngObj.setValue(value );
              // this.setErrorCode(testType,ErrorCodeType.Invalid, rngObj,itemInfo);
            }
            if (rngObj.getValue() == "") {
              Logger.log(rngObj.getA1Notation());
              rngObj.setBackground('yellow');
            }
          }
        }
      }      
    }


    setFeildsDilect(paramValue,isStringForced = false){
      for(var index =0; index < this.defaultList.length; index++ ){
        var itemInfo = this.defaultList[index];
        var rngObj = this.stdRange.offset(this.current, itemInfo.fileId);
        var value = itemInfo.value;
        if (value == "" && itemInfo.min != "") {
          value = itemInfo.min;
        }
        if (value == "" && itemInfo.max != "") {
          value = itemInfo.max;
        }
        if ((itemInfo.type == 'Integer') || (itemInfo.type == 'Int')){
          if (value == "") {
            value = 0;
          }
          rngObj.setValue(value);
        } else if ((itemInfo.type == 'String') ){
          if (isStringForced == true) {
              rngObj.setValue(paramValue);
          } else {
            if (value != "") {
              rngObj.setFormula('=REPT("' + paramValue + '",' + value + ')');
            } else {
              rngObj.setValue(paramValue);
            }
          }
        } else if ((itemInfo.type == 'Date') ){
          // 今日の日付を取得
          var formattedDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy/MM/dd');
          rngObj.setValue(formattedDate);
        } else {
          rngObj.setValue(paramValue);
        }
        if (rngObj.getValue() == "") {
          Logger.log(rngObj.getA1Notation());
          rngObj.setBackground('yellow');
        }
      }
    }


    getRequiredValue(line) {
      var type = line[this.getPosType()];
      var max = line[this.getPosMax()];
      var min = line[this.getPosMin()];

      if ((type == 'String') ){

      } else{
        // 最小があるなら最小値
        if (min != "") {
          return min;
        }
        // 最大があるなら最大値
        if (max != "") {
          return max;
        }
      }
      return 6;
    }

    writeHeader(){
      this.headerRange.offset(0,0).setValue("エンドポイント");
      Logger.log(this.activeRange);

      this.headerRange.offset(0,1).setValue(this.superHeaderRange.offset(0,5).getValue());

      this.headerRange.offset(1,0).setValue("メソッド");
      this.headerRange.offset(1,1).setValue(this.superHeaderRange.offset(1,5).getValue());

      this.headerRange.offset(2,0).setValue("機能名");
      this.headerRange.offset(2,1).setValue(this.superHeaderRange.offset(2,5).getValue());

      this.headerRange.offset(3,0).setValue("入力フォーム");
      this.headerRange.offset(3,1).setValue(this.getClassName());
      
    }

}


// function setYellowBackgroundDebug() {
//   // アクティブなスプレッドシートを取得
//   var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
//   var sheet = spreadsheet.getActiveSheet();

//   // セル A1 を取得
//   var range = sheet.getRange('A1');

//   // ログ出力でセルの確認
//   Logger.log('Sheet Name: ' + sheet.getName());
//   Logger.log('Range: ' + range.getA1Notation());
  
//   // 背景色を黄色に設定
//   range.setBackground('yellow');

//   // 背景色の確認
//   var bgColor = range.getBackground();
//   Logger.log('Background Color: ' + bgColor);
// }


