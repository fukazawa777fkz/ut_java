class dialogManager{
  
  constructor(fileUrl,folderUrl,fileName,folderName){
    this.fileUrl = fileUrl;
    this.folderUrl = folderUrl;
    this.fileName = fileName;
    this.folderName = folderName;

  }


  showDialog(){
    // カスタムHTMLダイアログを表示
    var htmlOutput = HtmlService.createHtmlOutputFromFile('Dialog')
        .setWidth(400)
        .setHeight(200)
        .setSandboxMode(HtmlService.SandboxMode.IFRAME);

    htmlOutput.append(`
      <script>
        updateLinks("${this.fileUrl}", "${this.folderUrl}", "${this.fileName}", "${this.folderName}");
      </script>
    `);
    SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'File Created');

  }
}

function deleteFileFromDrive(fileUrl) {
  try {
    // ファイルIDをURLから抽出
    var fileId = extractFileId(fileUrl);
    var file = DriveApp.getFileById(fileId);
    file.setTrashed(true);
    return 'File deleted successfully.';
  } catch (e) {
    throw new Error('Error deleting file: ' + e.toString());
  }
}

function extractFileId(fileUrl) {
  var parts = fileUrl.split('/');
  var fileId = parts[parts.length - 2]; // URL形式に依存しているので確認が必要
  return fileId;
}
