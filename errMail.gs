function ErrMail() {
    
    console.log("ErrMail 実行!");
    
    const to      = "k.kamikura@isowa.co.jp";                    // 送信先
    const subject = "【iサーチ】ファイルのアップロードに失敗しました。"; // タイトル
    const body    = '\
${D}さん\n\
\n\
いつもお仕事お疲れ様です。\n\
iサーチにファイルのアップロードが出来ませんでした。\n\
\n\
申し訳ございませんが、ファイル内容を再度ご確認お願いいたします。\n\
\n\
改善・不明点等ありましたら、お気軽に問い合わせフォーム\n\
よりご連絡下さい。'
.replace('${D}', D)
;  // 本文
    
    // 送信元の名前,bcc送信先
    const options = { name: 'iサーチ',
                      bcc: 'k.kamikura@isowa.co.jp'
                    };

    // メール送信実行    
    GmailApp.sendEmail(
              to,
              subject,
              body,
              options
    );
}
