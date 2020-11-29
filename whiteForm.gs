// ------------------------------------------------------ //
//  スプレットシートの情報を別のシートに書込                     //
// ------------------------------------------------------ //

function WhiteForm() {
  
  console.log("WhiteForm 実行!");
  
  const datas = getDatas;         // フォームの取得データ配列getDatasを格納
  const setSubDatas = dataTitle   // キーワードフォームに記載するデータのタイトル（A列）
  const notDatas = getNotDatas;   // キーワードフォームに記載不要なデータ
  let setDatas = [];              // キーワードフォームに書き込むデータ（B列）

  
  // [getDatas] と [notDatas] 内のデータを比較し、重複していないデータを [setDatas] に追加する。
  datas.concat(notDatas).forEach( data => {
     if ( datas.includes(data) && !notDatas.includes(data)) {
         setDatas.push(data);
     };
  });


  // 補足説明文を書込む (A列)
  let i = 1; // 書込み行
  // 配列setSubDatas に入っているデータを順番に取り出し、補足説明文を書き込んでいく
  setSubDatas.forEach( setSubData => {
     keyForm.getRange(`A${i}`).setValue(setSubData); // setSubDataを書込む 
     i++;                                            // 改行
  });


  // 取得情報を書込む (B列)  
  i = 1; // 書込み行
  // 配列setDatas に入っているデータを順番に取り出し、キーワードを書き込んでいく
  setDatas.forEach( setData => {

     if ( setData != "" ) {
        if ( setData === I ) i += 2;  // setDataが I(豆知識タイトル)だった場合は改行
        if ( setData === J ) i++;     // setDataが J(豆知識・プチ情報)だった場合は改行
     }

     keyForm.getRange(`B${i}`).setValue(setData); // setDataを書込む
     if ( setData !== "" ) i++;                   // setDataの中身が空でない場合は改行

  });


  // 問い合わせ・要望の項目に回答があったら実行
  if ( AM !== "" ) {
  
    // スプレットシート情報
    const ssForm = SpreadsheetApp.openById('1gctMJ1s7HJ51XQojAJDITSD9cab9hSB8f4CHBlYo94g');  // スプレットシート情報（書込み先）
    const reqForm = ssForm.getSheetByName('問い合わせ・要望');                                  // スプレットシート（問い合わせ・要望）情報
    const lastRow = reqForm.getLastRow();                                                    // スプレットシートの最終行目を取得する。
  
    // スプレットシートに記入
    const reqA  = reqForm.getRange(lastRow+1,  1).setValue(A);   // タイムスタンプ
    const reqB  = reqForm.getRange(lastRow+1,  2).setValue(AL);  // メールアドレス
    const reqC  = reqForm.getRange(lastRow+1,  3).setValue(AM);  // 問い合わせ・要望
  
  };




};
