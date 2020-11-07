// ------------------------------------------------------- //
//       フォーム回答者のアドレスを取得してファイル保存完了を通知する     //
// ------------------------------------------------------- //

function SendMail() {
  
  // アドレスが見つからない場合の送信先
  const errAddress = 'k.kamikura@isowa.co.jp';
  
  // isowaビトのアドレスが記載されたスプレットシートを取得
  const ssId = SpreadsheetApp.openById('1r9Ok3NF0_lwNa2fzCcpttteEim_Kb79xOBLB8GsJjnA');
  const addressSS = ssId.getSheetByName('メールアドレス一覧（ISOWA）');
  const arrData   = addressSS.getDataRange().getValues();
  const addressUrl = 'https://docs.google.com/spreadsheets/d/1r9Ok3NF0_lwNa2fzCcpttteEim_Kb79xOBLB8GsJjnA/edit#gid=0';

  // アドレスリストの配列の行と列を入替
  const _ = Underscore.load();               // アンダースコアを使用
  const arrTrans = _.zip.apply(_, arrData);  // 配列の行と列を入替
  const resNum = arrTrans[1].indexOf(Y);     // 回答者名と一致した行番号(開始No:0)

  // メール送信用
  let reply;   // メール送信先
  let title;   // メールタイトル
  let content; // メール本文
  const fileUrl = newFile.getUrl(); // アップロード完了後のファイルURL

  
  // 回答者名(Y)とスプレットシートの名前が一致したアドレスを取得
  if ( resNum !== -1 ) {
    
    reply   = arrTrans[2][resNum];
    title   = '【iサーチ】ファイルのアップロード完了通知';
    content = '\
${D}さん\n\
\n\
いつもお仕事お疲れ様です。\n\
iサーチに貴重な資料の登録をありがとうございます。\n\
${fileUrl}\n\
\n\
おかげ様でISOWAビトが安心して業務を行う事ができます。\n\
\n\
改善・不明点等ありましたら、お気軽に問い合わせフォーム\n\
よりご連絡下さい。'
.replace('${D}', D)
.replace('${fileUrl}', fileUrl);

    
    
    MailContents( reply, title, content );

  // 回答者名(Y)のアドレスが不明な場合
  } else {
  
    reply   = errAddress;
    title   = '【iサーチ】アップロードファイル登録者のアドレスが見つかりませんでした。';
    content = '\
iサーチ運営チームのみなさま\n\
\n\
いつもお仕事お疲れ様です。\n\
iサーチにファイルをアップロードした${D}さんの\n\
アドレスが見つかりませんでした。\n\
アップロードファイル：${fileUrl}\n\
アドレスリスト：${addressUrl}\n\
\n\
登録者名とアドレスリスト(スプレットシート)を\n\
確認お願いします。'
.replace('${D}', D)
.replace('${fileUrl}', fileUrl)
.replace('${addressUrl}', addressUrl);
               
    MailContents( reply, title, content );
  
  };
  
  
  
  function MailContents( reply, title, content ) {
    
    const to = reply;      // 送信先
    const subject = title; // タイトル
    const body = content;  // 本文
    
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
    
}

