// ------------------------------------------------------ //
//  結合されたpdfファイルを指定のフォルダに移動                    //
// ------------------------------------------------------ //

async function FileMove() {

  let copyId; // コピーファイルのID
  
  /* アップロードファイルの内容を取り出し、回答フォームの内容と比較。
     一致した内容のIDを [folderId] に格納する。 */
  contentsIds.forEach( el => {
    if ( C === el.contents ) {
      copyId = el.id;
    };
  });
 
  let copyName; // コピーファイルのファイル名

  // 画像 ・ 動画 以外の場合
  if ( !image && !video ) {
    copyName = renameFile;
  
  // 画像 ・ 動画 の場合 
  } else {
    copyName = fileName;
  };

  const copyFolder = DriveApp.getFolderById(copyId);             // コピーファイルの保存先フォルダID
  const originalFile = DriveApp.getFilesByName(copyName).next(); // コピー元のオブジェクトを取得
  const newFile = originalFile.makeCopy(copyName, copyFolder);   // 指定したフォルダにファイルをコピー

//  if ( !pdf && !excel && !wordLeg && !word && !sheets && !docs ) {
//    const rootFolder = DriveApp.getRootFolder();                   // ルートフォルダ
//    const delFile = rootFolder.getFilesByName(renameExt).next();   // ルートフォルダーのrenameExtを含むファイル名を格納     
//    delFile.setTrashed(true); // ルートフォルダのファイルをゴミ箱に入れる。
//  };

await FileTrash(); // アップロードフォルダ内の不要なファイルをゴミ箱に入れる。

};