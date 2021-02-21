// エラーとなった場合に内容をメールで通知
function ErrMail(msg) {

  console.log("ErrMail 実行!");
        
  // メールの送信情報
  const to = "k.kamikura@isowa.co.jp";  // 送信先
  const subject = "正常にスクリプトが完了できませんでした。";    // タイトル
        body = '\
    iサーチ運営チームのみなさま\n\
    \n\
    いつもお仕事お疲れ様です。\n\
    iサーチへのファイルを以下のエラーによりアップロードできませんでした。\n\
    \n\
    ${msg}\n\
    \n\
    \n\
    よろしくお願いします。\n\
    '
    .replace('${msg}', msg)
  const options = { name: 'iサーチ',               // 送信元の名前
                     bcc: 'k.kamikura@isowa.co.jp' // bcc 送信先
                  }

  // 送信実行    
  GmailApp.sendEmail(
    to,
    subject,
    body,
    options
  )


};  // ErrMail(msg)_END
