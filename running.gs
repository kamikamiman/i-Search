// トリガー実行時
function Running() {

  let msg = ""; // エラー時のメッセージ
  const lock = LockService.getDocumentLock(); // ドキュメントロックを使用

  // 30秒間のロックを取得
  try {
    lock.waitLock(30000);       // ロックを実施
    console.log("Main 実行！");  // 実行確認用
    Main();                     // 関数を実行
  } catch(e) {
    console.log("e:" + e);
    console.log("e.message:" + e.message);
    // ロックが取得できなかった時の処理を記述
    const checkword = "ロックのタイムアウト: 別のプロセスがロックを保持している時間が長すぎました。"
    if ( e.message == checkword ) {
      // ロックエラーの場合
      msg = "他の人が実行中です。";
      ErrMail(msg);
      console.log(msg);
    } else {
      // ロックエラー以外の場合
      msg = e.message;
      ErrMail(msg);
      console.log(msg);
    }
  } finally {
    // ロックを開放
    lock.releaseLock();
    console.log("正常に実行されました！");
  }

};

