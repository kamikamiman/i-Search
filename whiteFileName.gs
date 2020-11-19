// ------------------------------------------------------------ //
//  スプレットシートの情報をファイル名に書込 （ 画像, 動画の場合のみ ）   //
// ------------------------------------------------------------ //

function WhiteFileName() {
  
  console.log("WhiteFileName 実行!");
  
  let datas = getDatas;

  // Xのデータが空だった場合
  if ( F === "-------" ) {
    const delNum = datas.indexOf("-------");
    datas.splice(delNum,1);
  };
  
  
  // キーワードフォームに記載不要なデータはここに記入
  const notDatas = [ A, B, C, D, G, J ];
  
  // キーワードフォームに書き込むデータ
  let setDatas = [];


  // [getDatas] と [notDatas] 内のデータを比較し、重複していないデータを [setDatas] に追加
  datas.concat(notDatas).forEach( data => {
     if ( datas.includes(data) && !notDatas.includes(data)) setDatas.push(data);
  });

  
  // 配列setDatas に入っているデータを順番に取り出し、キーワードを書き込んでいく
  setDatas.forEach( setData => {

      // fileNameが空の場合
      if ( fileName == null ) {
        fileName = G;              // fileName が空の場合
      } else {
        fileName += `_${setData}`; // fileName が空でない場合 
      };

  });

  fileName += `.${extension}`; 
  uploadFile.setName(fileName);

};

