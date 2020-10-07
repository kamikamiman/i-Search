// ------------------------------------------------------ //
//  スプレットシートの情報を別のシートに書込                     //
// ------------------------------------------------------ //

function WhiteForm() {
  
  let datas = getDatas;

  // キーワードフォームに記載不要なデータはここに記入
  const notDatas = [ AM, S ];
  
  // キーワードフォームに書き込むデータ
  let setDatas = [];
  
  

  // [getDatas] と [notDatas] 内のデータを比較し、重複していないデータを [setDatas] に追加する。
  datas.concat(notDatas).forEach( data => {
     if ( datas.includes(data) && !notDatas.includes(data)) {
         setDatas.push(data);
     };
  });

  
  // 取得情報を書込む (B列)  
  let i = 1; // 書込み行

  // 配列setDatas に入っているデータを順番に取り出し、キーワードを書き込んでいく
  setDatas.forEach( setData => {

     if ( setData != "" ) {
        if ( setData === AD ) i += 2;  // setDataが AD(豆知識タイトル)だった場合は改行
        if ( setData === Z  ) i++;     // setDataが Z(豆知識・プチ情報)だった場合は改行
     }

     keyForm.getRange(`B${i}`).setValue(setData); // setDataを書込む
     if ( setData !== "" ) i++;                   // setDataの中身が空でない場合は改行

  });



  // 補足説明文を書込む (A列)
  const AAA = "登録日時";
  const YYY = "登録者";
  const WWW = "作成者";
  const XXX = "キーワード";
  const setSubDatas = [ AAA, YYY, WWW, XXX ];     // 補足説明文を配列に格納

  i = 1; // 書込み行
  // 配列setSubDatas に入っているデータを順番に取り出し、補足説明文を書き込んでいく
  setSubDatas.forEach( setSubData => {
     keyForm.getRange(`A${i}`).setValue(setSubData); // setSubDataを書込む 
     i++;                                            // 改行
  });

};
