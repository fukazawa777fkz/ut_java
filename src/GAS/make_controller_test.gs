// google drive保存先


function make_controller_test() {
  var utMng = new utManager();
  var fileName = utMng.getFunctionName() + "_Test.java";
  var google_forderId = "12U9NOQpXjMIxUta3fQ2Y5BWCEhR8bqUo";
  var fileMng = new fileManager(google_forderId, fileName, "controller", MimeType.PLAIN_TEXT);

  // テストコード作成
  createTestCode(fileMng,utMng)

  // 結果をファイルに保存
  fileMng.witeFToTextFile();

  var dialogMng = new dialogManager(fileMng.getUrl(), fileMng.getFolderUrl(), fileName, "controller") ;
  dialogMng.showDialog();
}


function createTestCode(fileMng, utMng) {

  // クラス定義
  createClassDefine(fileMng,utMng);

  var testCaseValues = utMng.gerTargetRange().getValues()
  for (var i = 0; i < testCaseValues.length; i++) {

    // 最初（テスト関数作成～リクエスト）
    createFuncRequest(fileMng,i,utMng);

    // リクエストのパラメータ設定を作成
    createRequestParam(fileMng,i,utMng);

    // HTTPステータス
    createHttpStatus(fileMng)

    // HTML名
    createReturnHtmlName(fileMng,i,utMng)

    // エラー情報
    createErrorInfo(fileMng,i,utMng)

    // テスト関数定義終了
    fileMng.write("        }")
    fileMng.write("")

  }

  // クラス定義終了
  fileMng.write("    }")

}

function createClassDefine(fileMng, utMng){
  fileMng.write(" ".repeat(4) + "@Nested");
  fileMng.write(" ".repeat(4) + "@DisplayName(" + '"' + utMng.geEndpoint() + '")');
  fileMng.write(" ".repeat(4) + "class " + utMng.getFunctionName() + " {");
}

function createFuncRequest(fileMng,i,utMng){
    fileMng.write("        " + "@Test");
    fileMng.write("        " + "public void " + utMng.getTetsFuntionName(i) + "() throws Exception {");
    fileMng.write("            " + "mockMvc.perform(" + utMng.getMesod() + "(" + '"' + utMng.geEndpoint() + '")' );
}

function createRequestParam(fileMng, i,utMng){
  var currentFieldValues = utMng.gerCurrentFieldValues(i)


  var paramMaxCount =  countNonNullElements(currentFieldValues[0])
  var paramCount = 0;
  for (var j =0; j < currentFieldValues[0].length; j++) {
    var paramLine ="";
    var ret = utMng.getDisplayValue(currentFieldValues[0][j])
    if (ret[0] == "null") {
      continue;
    }
    if (ret[1] == 0) {
      // ダブルコーテーションで囲む
      paramLine = "                " + '.param("' + utMng.getFieldName(j) + '", ' + ret[0]  +')';
    } else {
      // ダブルコーテーションで囲まない
      paramLine = "                " + '.param("' + utMng.getFieldName(j) + '", "' + ret[0]  +'")';
    }
    if (++paramCount == paramMaxCount) {
      // 最後のパラメータ設定は")"を付与
      paramLine = paramLine + ")";
    }
    fileMng.write(paramLine);
  }
}

function createHttpStatus(fileMng){
  fileMng.write("                " + ".andExpect(status().isOk())")
}

function createReturnHtmlName(fileMng,i,utMng) {
  fileMng.write("                " + ".andExpect(view().name(" + '"' + utMng.getReturnHtml(i)+ '"))')
}

function createErrorInfo(fileMng, i,utMng){
  if (utMng.isNormalTermination(i)){
    fileMng.write("                " + ".andExpect(model().hasNoErrors());")
  } else {
    var fieldErrors = utMng.getFieldErrors(i)
    fileMng.write("                " + ".andExpect(model().attributeHasFieldErrors(" + fieldErrors  + "))")
    var errorRng = utMng.gerCurrentErrorFieldRange(i)
    var errorValues =  utMng.getFieldErrorValues(i)
    var errorNum = countNonEmptyElements(errorValues[0]);
    var errorCount = 0;
    for (var errorIndex =0; errorIndex < errorValues[0].length; errorIndex++) {
      var errorCode  = errorValues[0][errorIndex]
      if (errorCode != "") {
        var errorFeild = errorRng.offset(-i -1,errorIndex).getValue()
        var fieldErrorCode =  '"' + utMng.getFormTag() +  '","' + errorFeild + '","' + errorCode + '"'
        var andExpect = "                " + ".andExpect(model().attributeHasFieldErrorCode(" + fieldErrorCode  + "))";
        if (++errorCount == errorNum) {
          // 最後のエラーはセミコロンをつける
          andExpect = andExpect + ";";
        }
        fileMng.write(andExpect);
      }
    }
  }
}

