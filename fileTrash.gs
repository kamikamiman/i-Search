// ------------------------------------------------------ //
//  結合前の不要なファイルを削除                                //
// ------------------------------------------------------ //

function FileTrash() {
  
  console.log("FileTrash 実行!");

  const delFolder = DriveApp.getFolderById(mergeman)  // 削除したいファイルのフォルダ
  const delfiles = delFolder.getFiles();              // 削除したいファイル
  
  // 指定したフォルダ内にアップロード関連のファイルがあれば削除
  while(delfiles.hasNext()) {
    
    const file = delfiles.next();    // ファイルを取得
    const fileName = file.getName(); // ファイル名を取得
    let fileNum;
    
    // 画像 ・ 動画 の場合
    if ( image || video ) {
      fileNum = fileName.indexOf(fileName);
      
    // 画像 ・ 動画以外 の場合
    } else {
      fileNum = fileName.indexOf(rename);
    };
      
    /* アップロードファイル名(rename)がファイル名(fileName)に
       含まれていれば、そのファイルを削除 */
    if ( fileNum != -1 ) {
      file.setTrashed(true);
    };
    
  };
  
};

