
function main_create_unit_test() {
  // 初期化
  var utSpecManager = new unitTextSpecManager();

  var cnt = utSpecManager.minList.length;

  Logger.log(utSpecManager.requiredList.length);
  Logger.log(utSpecManager.minList.length);
  Logger.log(utSpecManager.maxList.length);
  Logger.log(utSpecManager.enumList.length);

  utSpecManager.writeHeader();
  utSpecManager.writeNormal();
  utSpecManager.writeAbnormal();
  
}


