/*****************************************************/
/***   変換したドキュメントにキーワードを書込み、pdf変換   ***/
/*****************************************************/

async function DocsPdfConvert () {
  
  let datas = getDatas;       // フォームの取得データ配列getDatasを格納
  const notDatas = [ B, S ];  // キーワードフォームに記載不要なデータはここに記入
  let setDatas = [];          // キーワードフォームに書き込むデータ
  
  // [datas] と [notDatas] 内のデータを比較。
  // 重複していないデータを [setDatas] に追加する。
  datas.concat(notDatas).forEach( data => {
    if ( datas.includes(data) && !notDatas.includes(data)) {
      setDatas.push(data);
    }
  });

  WhiteDocs();        // ドキュメントにキーワードを書き込む
  await DocsPdf();    // ドキュメントをpdf変換





  /******************************************************/
  /***     Word >> 変換したドキュメントにキーワードを書込む      ***/
  /******************************************************/

  function WhiteDocs () {
    const folder = DriveApp.getFolderById(mergeman); // フォルダを取得
    const fileList = folder.getFiles();              // フォルダ内のファイルを取得
  
    // フォルダ内にファイルが存在する場合
    while (fileList.hasNext()) {
    
      const file = fileList.next();                  // ファイルを取得
      const fileName = file.getName();               // ファイル名を取得
    
      // ファイル名が一致する場合
      if ( fileName === renameExt2 ) {
      
        let fileId = file.getId();                   // ファイルIDを取得
        let docFile = DriveApp.getFileById(fileId);  // ファイルを取得
        let docUrl = docFile.getUrl();               // ファイルURLを取得
        let doc = DocumentApp.openByUrl(docUrl);     // ファイルを開く
        let body = doc.getBody();                    // ファイル内データを取得
        body.appendPageBreak();                      // 新たにページを作成
      
        // 配列setDatas の中身を順番に取り出し、キーワードを書き込んでいく
        setDatas.forEach( setData => {
                       
           // setData にデータが入っている場合
          if ( setData != "" ) {
        
            if ( setData === AD ) '\n\n';  // setDataが AD(豆知識タイトル)だった場合は改行
            if ( setData === Z  ) '\n';    // setDataが Z(豆知識・プチ情報)だった場合は改行
        
            if ( setData === A ) {
              body.appendParagraph('登録日時 ： {setData}');
              body.replaceText('{setData}', setData);
            } else if ( setData === Y ) {
              body.appendParagraph('登録者 ： {setData}');
              body.replaceText('{setData}', setData);
            } else if ( setData === W ) {
              body.appendParagraph('作成者 ： {setData}');
              body.replaceText('{setData}', setData);
            } else if ( setData === X ) {
              body.appendParagraph('キーワード ： {setData}');
              body.replaceText('{setData}', setData);
            } else {
              body.replaceText('{setData}', setData);
            };
         
          };
       
        });
    
      doc.saveAndClose(); // 強制的にファイルを更新
    
      };
  
    };

  };

  
  /******************************************************/
  /***       キーワードを書き込んだファイルをpdf変換する     ***/
  /******************************************************/

  function DocsPdf () {
  
    const folder = DriveApp.getFolderById(mergeman);  // フォルダを取得
    const fileList = folder.getFiles();               // フォルダ内のファイルを取得

    // フォルダ内のファイルが存在すれば実行
    while (fileList.hasNext()){
     
       const file = fileList.next();    // ファイルを取得
       const fileName = file.getName(); // ファイル名を取得
     
       // ファイル名が一致した場合
       if ( fileName === renameExt2 ) {
          let fileId = file.getId();                     // ファイルIDを取得
          const convFile = DriveApp.getFileById(fileId); // ファイルを取得
          const pdfFile  = convFile.getAs(MimeType.PDF); // pdf変換
          const createPdf = folder.createFile(pdfFile);  // pdfファイル作成
          const pdfId = createPdf.getId();               // pdfファイルID取得
          const pdfName = `${rename}.pdf`;               // pdfファイル名を指定
          const pdfRename = DriveApp.getFileById(pdfId).setName(pdfName); // ファイル名を変更
       };
    
    };
  
  };



};