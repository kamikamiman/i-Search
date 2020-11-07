// ------------------------------------------------------- //  
// スプレットシートのセル情報を削除                              //
// ------------------------------------------------------- //
function SSRemove() {
  const keyFormLastRow = keyForm.getLastRow();               // スプレットシート（キーワード書込先）の最終行目を取得
  keyForm.getRange(1, 1, keyFormLastRow, 2).clearContent();  // 指定したセルのコンテンツを削除
}