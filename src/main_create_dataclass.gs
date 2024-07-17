// google drive保存先
function create_data_class() {

  // 初期化
  var specMng = new specManager();
  var google_forderId = "1en5k_tzPlL7_P6GEIdrUcJWlH4u2Vvez";
  var fileName = specMng.getPearentRange().offset(0,specMng.getPosType() + 1 ).getValue() + ".java";
  var fileMng = new fileManager(google_forderId, fileName, "model", MimeType.PLAIN_TEXT);

  // 関数ヘッダ作成
  createFuncHeder(specMng,fileMng);

  // 関数本体作成
  createFuncBody(specMng,fileMng);

  // 結果をファイルに保存
  fileMng.witeFToTextFile();

  var dialogMng = new dialogManager(fileMng.getUrl(), fileMng.getFolderUrl(),fileName, "model") ;
  dialogMng.showDialog();

}

// 関数ヘッダ作成
function createFuncHeder(specMng,fileMng){
  fileMng.write('/**')
  fileMng.write(' * ' + specMng.getPearentRange().getValue())
  fileMng.write(' *')
  var values = specMng.gerTargetRange().getValues()
  for (var i = 0; i < values.length; i++) {
    fileMng.write(' * @property ' + values[i][0])
  }
  fileMng.write(' */')
}

// 関数本体作成
function createFuncBody(specMng,fileMng){
  var className = specMng.getPearentRange().offset(0,specMng.getPosType() + 1 ).getValue()
  fileMng.write('@Data')
  fileMng.write('public class ' + className + ' {')

  var values = specMng.gerTargetRange().getValues()
  for (var i = 0; i < values.length; i++) {
    // createAnotationForItemCheck(specMng, values[i], fileMng)
    createAnotationForList(specMng, values[i],fileMng)
    fileMng.write('    private ' + values[i][specMng.getPosType()] + ' ' + values[i][specMng.getPosPhysics()] + ';')
    fileMng.write('')
  }
  fileMng.write('}')

}

// アノテーション作成
function createAnotationForList(specMng, line, fileMng) {

  var physics = line[specMng.getPosPhysics()]
  var required = line[specMng.getPosRequired()]
  var min = line[specMng.getPosMin()]
  var max = line[specMng.getPosMax()]
  var type = line[specMng.getPosType()]

  // 論理名
  fileMng.write("    // " +  line[0]);

  // 必須項目
  if (required == "有") {
    fileMng.write("    " +  "@NotNull" + "(message=" + '"' + "入力してください。" + '")' );
  }

  if (type == "String") {
    if ((min != "") && (max != "")){
      fileMng.write("    " +  "@Size(min=" + min + ",max=" + max + ",message=" + '"' + min + "文字から" + max + "文字で" + "指定してください。" + '")' );
    }
    if ((min != "") && (max == "")){
      fileMng.write("    " +  "@Size(min=" + min + ",message=" + '"' + min + "文字以上で" + "指定してください。" + '")' );
    }
    if ((min == "") && (max != "")){
      fileMng.write("    " +  "@Size(max=" + max + ",message=" + '"' + max + "文字で以下で" + "指定してください。" + '")' );
    }
  }

  if (type == "Integer") {
    if ((min != "") && (max != "")){
      fileMng.write("    " +  "@Min(value=" + min + ",message=" + '"' + min + "-" + max + "で" + "指定してください。" + '")' );
      fileMng.write("    " +  "@Max(value=" + max + ",message=" + '"' + min + "-" + max + "で" + "指定してください。" + '")' );
    }
    if ((min != "") && (max == "")){
      if (min == 1) {
        fileMng.write("    " +  "@Min(value=" + min + ",message=" + '"' + "正の整数を入力してください。" + '")' );
      } else {
        fileMng.write("    " +  "@Min(value=" + min + ",message=" + '"' + min + "以上で" + "指定してください。" + '")' );
      }
    }
    if ((min == "") && (max != "")){
      fileMng.write("    " +  "@Max(value=" + max + ",message=" + '"' + max + "以下で" + "指定してください。" + '")' );
    }
  }

}

// アノテーション作成
function createAnotationForItemCheck(specMng, line, fileMng) {
  var ano_values = specMng.getAnotationRng().getValues();
  for (var i = 0; i < ano_values.length; i++) {
    var baseFieldName = line[0];
    var fieldName = ano_values[i][ANO_COLUMNS_logicalName];
    var anotation = ano_values[i][ANO_COLUMNS_Anotation];
    var message = ano_values[i][ANO_COLUMNS_ErrorMessage];


    if (anotation = "@size"){

    }
    if ((anotation = "@min") || (anotation = "@max")) {

    }
    if (baseFieldName == fieldName ){
      fileMng.write("    " +  anotation + "(message=" + '"' + message + '")' );
    }
    if (fieldName =='') {
      break;
    }
  }
}
