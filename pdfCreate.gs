  // ------------------------------------------------------- //
  // キーワードフォーム(スプレットシート)をPDF化して指定フォルダに保存 //
  // ------------------------------------------------------- //
  function PdfCreate() {
    SpreadsheetApp.flush();
    const sheetId = keyForm.getSheetId(); // スプレットシートのIDを取得
    const url = 'https://docs.google.com/spreadsheets/d/1gwfUf30jWIMy-emXiJLN4ZwmXU2cXwtp4EUfD4slWxE/export?exportFormat=pdf&gid=SID'.replace('SID', sheetId);
    const token = ScriptApp.getOAuthToken();
    const response = UrlFetchApp.fetch(url, {
      headers:{
        'Authorization': 'Bearer '+token
      }
    });
    
    let fileName; // キーワードフォームのファイル名    
    let folderId; // 保存先フォルダのID
    
  
    // 豆知識・プチ情報の場合のファイル名と保存先フォルダID
    if ( AD !== '' ) {
      fileName = `${AD}.pdf`; // ファイル名を指定
      
      /* アップロードファイルの内容を取り出し、回答フォームの内容と比較。
         一致した内容のIDを [folderId] に格納する。 */
      contentsIds.forEach( el => {
        if ( C === el.contents ) folderId = el.id;
      });
    
    // 豆知識・プチ情報の場合のファイル名と保存先フォルダID
    } else {
      fileName = `${rename}_キーワード.pdf`; // キーワードフォームのファイル名を格納
      folderId = mergeman;
    };
    
    const blob = response.getBlob().setName(fileName); // pdfの名前
    const folder = DriveApp.getFolderById(folderId);   // pdfの保存先フォルダを指定
    const requestForm = folder.createFile(blob);       // フォルダ内にpdfを作成

  };
  