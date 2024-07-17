class fileManager {

  constructor(parentFolderId, fileName, folderName, mimeType){

    this.fileName = fileName;
    this.newFolderName = folderName;
    this.parentFolderId = parentFolderId;
    this.mimeType = mimeType;
    this.fileContent = "";

  }

  write(value) {
    this.fileContent += "\n" + value;
  }


  witeFToTextFile(){
    var file;
    if (this.newFolderName == "") {
      // フォルダ名を指定しなかった時
      file = DriveApp.createFile(this.fileName, this.fileContent, this.mimeType);

    } else {
      // フォルダ名を指定したとき
      var parentFolder = DriveApp.getFolderById(this.parentFolderId);
      // 特定のフォルダ内で同名のフォルダが存在するか確認
      var folders = parentFolder.getFoldersByName(this.newFolderName);
      var newFolder;
      if (folders.hasNext()) {
        // フォルダが存在する場合、そのフォルダを取得
        newFolder = folders.next();
      } else {
        // フォルダが存在しない場合、新しいフォルダを作成
        newFolder = parentFolder.createFolder(this.newFolderName);
      }
      file = newFolder.createFile(this.fileName, this.fileContent, MimeType.PLAIN_TEXT);
    }
    this.url = file.getUrl();
    this.folderUrl = newFolder.getUrl();
    return file.getUrl();
  }

  getUrl(){
    Logger.log(this.url)
    return this.url;
  }

  getFolderUrl(){
    return this.folderUrl;
  }
}


