function  onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Custom Menu')
    .addItem('class作成', 'create_data_class')
    .addItem('controller_test作成', 'make_controller_test')
    .addItem('単体試験書作成', 'main_create_unit_test')
    .addToUi();
}