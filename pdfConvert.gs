// ------------------------------------------------------- //
//          アップロードファイルをPDFに変換                        //
// ------------------------------------------------------- //
function PdfConvert() {
  
  console.log("PdfConvert 実行!");
  
  const folder = DriveApp.getFolderById(mergeman);   // フォルダを取得
    
  // ドキュメント ・ スプレットシート の場合
  if ( docs || sheets ) {
    ConvertGFile(); // pdf変換
    
  // ドキュメント ・ スプレットシート 以外の場合
  } else {
    ConvertApi(id); // pdf変換
  };
  
  
  // ドキュメント ・ スプレットシート をpdf変換
  async function ConvertGFile() {
    
    console.log("ConvertGFile 実行!");
    
    if ( docs ) {
      
      DocsPdfConvert(); // ドキュメントにキーワードフォーム追加・pdf変換

    } else {
      
    const _pdfFile = uploadFile.getAs(MimeType.PDF); // コンバートファイルをpdf変換
    let pdfFile = folder.createFile(_pdfFile);       // pdfを作成
    const pdfFileName = `${renameExt}.pdf`;          // pdfファイル名
    await pdfFile.setName(pdfFileName);              // pdfファイル名を変更

    }

  };

  
  // ------------------------------------------------------- //
  // エクセル ・ ワード ・ テキスト をpdf変換する (Drive API)           //
  // ------------------------------------------------------- //
  
  function ConvertApi(id) {
    
    console.log("ConvertApi 実行!");

    const upFileBlob = uploadFile.getBlob();    // ブロブを取得
    const upFileName = uploadFile.getName();    // ファイル名を取得
    const upFolder = mergeman;  // 指定したIDファイルがあるフォルダID
    const conFileId = GetConvertFileId(upFileBlob, upFileName, upFolder);  // コンバートファイルIDを取得
    const conFile = DriveApp.getFileById(conFileId); // コンバートファイルを取得
    
    // ワード ・ テキスト の場合
    if ( wordLeg || word || text ) {

      DocsPdfConvert(); // 変換したドキュメントファイルにキーワードフォーム追加 ・ pdf変換
    
    // ワード ・ テキスト 以外の場合
    } else {

      const _pdfFile = conFile.getAs(MimeType.PDF); // コンバートファイルをpdf変換
      let pdfFile = folder.createFile(_pdfFile);    // pdfファイルを作成

    };
    
  };
  
  
  // ------------------------------------------------------- //
  //         Drive.Files.insetを使った変換方法                  //
  // ------------------------------------------------------- //
  
  function GetConvertFileId (upFileBlob, upFileName, upFolder) {
    
    console.log("GetConvertFileId 実行!");
    
    let mimeType; // ファイルの種類
    
    // エクセルの場合
    if ( excel || excelLeg ) {

      mimeType = MimeType.GOOGLE_SHEETS; // スプレットシートに変換
    
    // エクセル以外の場合
    } else {

      mimeType = MimeType.GOOGLE_DOCS;   // ドキュメントに変換

    };
    
    //変換情報
    const files = {
      title: upFileName,  // ファイル名
      mimeType: mimeType, // ファイルの種類
      parents: [{id: upFolder}], // ファイルの保存先
    };
    
    const res = Drive.Files.insert(files, upFileBlob); //Drive APIで変換
    return res.id; //変換シートのIDを返す
    
  };

};
