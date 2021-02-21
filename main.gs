/*
　・ 登録者名にスペースがあった場合は、スペースを削除して格納する。
　　 スプレットの情報と完全一致した場合のみアドレスを取得する為。
　　　>>> 完了

　・ アップロードファイルを複数登録する場合に対応する。

　・ アップロードされない時がある。正常にアップロードできなかった後に
　　 何か影響があるのか要確認。→ scriptの実行する順番を確実に守るようにする。
     >>> ファイル名が一致していない可能性が高い。

*/

// デバック時はここから直接実行
function Main() {

  // スプレットシートを取得
  const getSS = SpreadsheetApp.openById('11Y5gKhXQnXspVgih58s3_DDT5orEz-WAZ4sCwSXpAOs'); // スプレットシート情報（読出し）
  const setSS = SpreadsheetApp.openById('1zzq1OQxTZJOFo2eFaTGkNLyipGWDjl9i-LijcgIU3Hw'); // スプレットシート情報（書込み）
  const form  = getSS.getSheetByName('フォームの回答 1');                                   // スプレットシート（フォーム回答）情報
  let keyForm = setSS.getSheetByName('キーワードフォーム');                                  // スプレットシート（キーワードフォーム）情報

  // シートの最終行・最終列を取得
  const formLastRow = form.getLastRow();    // 最終行目を取得（フォーム回答1）
  const formLastCol = form.getLastColumn(); // 最終列目を取得（フォーム回答1）

  // アップロードした元ファイルの保存先フォルダID
  const upfolder = '0BxQbdvn-LT2ufmlKRk5SY3k2VFg0MzQxak1uYVZpNXdDYURKVkk2RHgtSUk1WDI3UExMWlk';

  // スプレットシート(フォーム回答1)の最終行の回答全てを取得
  const getAnswers = form.getRange(formLastRow, 1, 1, formLastCol).getValues().flat();

  // フォーム回答を個別に取得
  const __A = getAnswers[0];  
  const _A  = new Date(__A);
  const A  = Utilities.formatDate(_A, 'JST', 'yyyy/M/d');  // タイムスタンプ
  const B  = getAnswers[1];     // ファイル選択           >>> メールアドレス 
  const C  = getAnswers[2].replace(/　/g,"");  // 作成者（フルネーム）    >>> 登録者(フルネーム)
  const D  = getAnswers[3];     // 登録者（フルネーム）    >>> 作成者(フルネーム)
  const E  = getAnswers[4];     // お客様名              >>> ファイル内容
  const F  = getAnswers[5];     // 関連付けしたいキーワード >>> 豆知識・プチ情報(アップロードファイルの有無)
  const G  = getAnswers[6];     // ファイル内容           >>> 豆知識・プチ情報(タイトル)
  const H  = getAnswers[7];     // ファイル内容（その他）   >>> 豆知識・プチ情報(内容)
  const I  = getAnswers[8];     // 豆知識・プチ情報のタイトル >>> 機械の種類
  const J  = getAnswers[9];     // 豆知識・プチ情報         >>> コルゲータの種類
  const K  = getAnswers[10];    // 機械の種類              >>> ミルロールスタンドの機種
  const L  = getAnswers[11];    // コルゲータの種類         >>> キーワード(部位)
  const M  = getAnswers[12];    // ミルロールスタンドの機種  >>> キーワード(動作)
  const N  = getAnswers[13];    // スプライサの機種        >>> キーワード(症状・状態)
  const O  = getAnswers[14];    // シングルフェーサの機種   >>> キーワード(部品)
  const P  = getAnswers[15];    // ブレーキスタンドの機種   >>> スプライサの機種
  const Q  = getAnswers[16];    // プレヒータの機種        >>> シングルフェーサの機種
  const R  = getAnswers[17];    // グルーマシンの機種      >>> キーワード(部位)
  const S  = getAnswers[18];    // ダブルフェーサの機種     >>> キーワード(動作)
  const T  = getAnswers[19];    // スリッタスコアラの機種   >>> キーワード(症状・状態)
  const U  = getAnswers[20];    // カッターの機種          >>> キーワード(部品)
  const V  = getAnswers[21];    // スタッカーの機種        >>> ブレーキスタンドの機種
  const W  = getAnswers[22];    // 管理装置の機種（コルゲータ） >>> キーワード(部位)
  const X  = getAnswers[23];    // 製函機の種類              >>> キーワード(動作)
  const Y  = getAnswers[24];    // アイビス・ファルコンの機種   >>> キーワード(症状・状態)
  const Z  = getAnswers[25];    // フレキソの機種            >>> キーワード(部品)
  const AA = getAnswers[26];    // プリスロの機種            >>> プレヒータの機種
  const AB = getAnswers[27];    // フォルダーグルアの機種      >>> キーワード(部位)
  const AC = getAnswers[28];    // スタッカーの機種（製函機）  >>> キーワード(動作)
  const AD = getAnswers[29];    // 管理装置（製函機）         >>> キーワード(症状・状態)
  const AE = getAnswers[30];    // カウンターエジェクタの機種  >>> キーワード(部品)
  const AF = getAnswers[31];    // 他社製・取売機            >>> グルーマシンの機種
  const AG = getAnswers[32];    // その他（製函機）          >>> キーワード(部位)
  const AH = getAnswers[33];    // 機器・部品の種類          >>> キーワード(動作)
  const AI = getAnswers[34];    // 機器・部品名（機械）       >>> キーワード(症状・状態)
  const AJ = getAnswers[35];    // 機器・部品名（電気）       >>> キーワード(部品)
  const AK = getAnswers[36];    // 社内業務マニュアル         >>> ダブルフェーサの機種
  const AL = getAnswers[37];    // メールアドレス            >>> キーワード(部位)
  const AM = getAnswers[38];    // 問合せ・要望              >>> キーワード(動作)
  const AN = getAnswers[39];    // 関連付けしたいキーワード    >>> キーワード(症状・状態)
  const AO = getAnswers[40];    // ファイル選択              >>> キーワード(部品)
  const AP = getAnswers[41];    // 関連付けしたいキーワード    >>> スリッタスコアラの機種
  const AQ = getAnswers[42];    // 関連付けしたいキーワード    >>> キーワード(部位)
  const AR = getAnswers[43];    // 関連付けしたいキーワード    >>> キーワード(動作)
  const AS = getAnswers[44];    // 関連付けしたいキーワード    >>> キーワード(症状・状態)
  const AT = getAnswers[45];    // 関連付けしたいキーワード    >>> キーワード(部品)
  const AU = getAnswers[46];    // 関連付けしたいキーワード    >>> カッターの機種
  const AV = getAnswers[47];    // 関連付けしたいキーワード    >>> キーワード(部位)
  const AW = getAnswers[48];    // 関連付けしたいキーワード    >>> キーワード(動作)
  const AX = getAnswers[49];    // 関連付けしたいキーワード    >>> キーワード(症状・状態)
  const AY = getAnswers[50];    // 関連付けしたいキーワード    >>> キーワード(部品)
  const AZ = getAnswers[51];    // 関連付けしたいキーワード    >>> スタッカーの機種
  const BA = getAnswers[52];    // 関連付けしたいキーワード    >>> キーワード(部位)
  const BB = getAnswers[53];    // 関連付けしたいキーワード    >>> キーワード(動作)
  const BC = getAnswers[54];    // 関連付けしたいキーワード    >>> キーワード(症状・状態)
  const BD = getAnswers[55];    // 関連付けしたいキーワード    >>> キーワード(部品)
  const BE = getAnswers[56];    // 関連付けしたいキーワード    >>> 管理装置の機種
  const BF = getAnswers[57];    // 関連付けしたいキーワード    >>> キーワード(部位)
  const BG = getAnswers[58];    // 関連付けしたいキーワード    >>> キーワード(動作)
  const BH = getAnswers[59];    // 関連付けしたいキーワード    >>> キーワード(症状・状態)
  const BI = getAnswers[60];    // 関連付けしたいキーワード    >>> キーワード(部品)
  const BJ = getAnswers[61];    // 関連付けしたいキーワード    >>> ファイル選択
  const BK = getAnswers[62];    // 関連付けしたいキーワード    >>> お客様名
  const BL = getAnswers[63];    // 関連付けしたいキーワード    >>> 何か改善してほしい点・追加要望等あれば記入して下さい。
  const BM = getAnswers[64];    // 関連付けしたいキーワード    >>> キーワード(部位) スプライサ
  const BN = getAnswers[65];    // 関連付けしたいキーワード    >>> キーワード(動作) スプライサ
  const BO = getAnswers[66];    // 関連付けしたいキーワード    >>> キーワード(症状・状態) スプライサ
  const BP = getAnswers[67];    // 関連付けしたいキーワード    >>> キーワード(部品) スプライサ


  // googleフォームからの取得データ(googleフォームに追加したら追加必要)
  // G, H は、配列の最後に記入。プチ情報タイトル・内容の為、最後に書込。
  const getDatas = [ A,  B,  C,  D,  E,  F,  I,  J,  K,  L,  M,  N,
                     O,  P,  Q,  R,  S,  T,  U,  V,  W,  X,  Y,  Z,
                    AA, AB, AC, AD, AE, AF, AG, AH, AI, AJ, AK, AL,
                    AM, AN, AO, AP, AQ, AR, AS, AT, AU, AV, AW, AX,
                    AY, AZ, BA, BB, BC, BD, BE, BF, BG, BH, BI, BJ,
                    BK, BL, BM, BL, BO, BP, G, H ];

  // フォーム回答の設定(関数内でフォーム回答を使用する場合は、ここに追加した変数を使用して！)
  const timeStamp       = A;   // タイムスタンプ
  const mailAddress     = B;   // メールアドレス
  const registPerson    = C;   // 登録者
  const fileContents    = E;   // ファイル内容
  const notUploadFile   = F;   // 豆知識・プチ情報のアップロードファイル有無
  const littleInfoTitle = G;   // 豆知識・プチ情報のタイトル
  const littleInfoConts = H;   // 豆知識・プチ情報の内容
  const uploadFileUrl   = BJ;  // アップロードした元ファイルのURL
  const requests        = BL;  // 改善点・要望

  // キーワードフォームに記載不要なデータ
  const getNotDatas = [ mailAddress, notUploadFile, uploadFileUrl, requests ];

  // キーワードフォームに記載する説明文
  const AAA = "登録日時";
  const CCC = "登録者";
  const DDD = "作成者";
  const KEY = "キーワード";
  const dataTitle = [ AAA, DDD, CCC, KEY ];     // 説明文を配列に格納

  // アップロードファイル内容と保存先フォルダID
  const a = { contents:"取説・マニュアル（ISOWAオリジナル）",      id:"1a3JnMp6ZfToZ4lthL4bGDdZOkT1IIpTR" };
  const b = { contents:"機器・部品マニュアル（メーカー）",         id:"1OxoFmxgU-hS07Klcqa7S7I12rpt1tlJa" };
  const c = { contents:"手順書",                               id:"1c7TqF0wHPrFCdfIpx7xfWdTE6-MTbvRW" };
  const d = { contents:"調整要領書",                            id:"1kJRlkkqtn3EFAlpdYJzCgRWLEtWBrcib" };
  const e = { contents:"トラブルシューティング・定数表・パラメータ", id:"1KkrkvTAx0oEtyaTZTYYYVzq35I055lz5" };
  const f = { contents:"仕様書",                               id:"1far2FpiadbUZ_XHQD9BzwKlJVySGEKx9" };
  const g = { contents:"報告書",                               id:"1oXXtqrc96Ze4mykkfvI6K9bkzTvflNpE" };
  const h = { contents:"見解書",                               id:"16sky4P09sT2KdW1ngcPdeH5f0AjAaA-u" };
  const i = { contents:"点検・検査表",                          id:"1cL_4s38of31LA_xJKxLyplD_lu_5nIWl" };
  const j = { contents:"設計変更",                             id:"1DLPbnYjJQmkib4AmmfKdV2YNk9eHU42N" };
  const k = { contents:"アイレポ（修理・組立）",                 id:"1-cSlKa2xEy1raZqf5ySf4lLOWQ4A8hm-" };
  const l = { contents:"豆知識・プチ情報",                      id:"1W34glQWjt-spp7Hl5IN0TzczUnW8dyXF" };
  const m = { contents:"社内資料（業務マニュアル）",              id:"1tQsfk5n3MME2zJopDE2QShXN8v5jtYsK" };
  const n = { contents:"社内資料（フォーム）",                   id:"1GQTozvYrRCdq7c24J8Nk3K8YpJ1HHi8h" };
  const o = { contents:"その他",                               id:"1RxdO0E4rNmEM1Yg9zwQrBh4PufuIDvRn" };
  const p = { contents:"画像",                                id:"1y5_S_pZxmPAFyY9P59uL9ggtNTJq5bud" };
  const q = { contents:"動画",                                id:"1NMIkUvyFwJUqHLQaIu4jWbb5AX5HeqIs" };
  const r = { contents:"", id:"" };
  const s = { contents:"", id:"" };
  const t = { contents:"", id:"" };
  const u = { contents:"", id:"" };
  const v = { contents:"", id:"" };
  const w = { contents:"", id:"" };
  const x = { contents:"", id:"" };
  const y = { contents:"", id:"" };
  const z = { contents:"", id:"" };

  // ファイル内容と保存先フォルダIDを格納する配列
  const contentsIds = [ a, b, c, d, e, f, g, h, i, j, k,l, m, n, o, p, q,
                        r, s, t, u, v, w, x, y, z ];

  // アップロードファイル有無 判定
  const uploadFileExists = uploadFileUrl !== "";


  // それぞれ処理を実行
  class Obj {

    constructor() {};


    // ========== [ メソッド ] アップロードファイル・キーワードファイル名を変更 ========== //

    GetFileData() {

      console.log("GetFileData 実行!");

      // 変数の定義・初期化
      let id;                  // アップロードID
      let uploadFile;          // アップロードファイル
      let upLoadName;          // アップロードされたファイル名を取得(名前有)
      let upLoadNames = [];    // 配列 アップロードされたファイル名
      let upLoadUrls  = [];    // 配列 アップロードされたファイルURL

      console.log("uploadFileExists(アップロードファイルの存在判定):" + uploadFileExists);  // ログ確認用

      // アップロードファイルが存在する場合に実行
      if ( uploadFileExists ) {
        const multiple = uploadFileUrl.indexOf(", ") !== -1; // アップロードファイルが複数あるか判定

        // アップロードファイルのURLを配列で取得
        if ( multiple ) {
          upLoadUrls = uploadFileUrl.split(', ');            // 複数ファイルの場合
        } else {
          upLoadUrls = [ uploadFileUrl ];                    // 1ファイルのみの場合
        }

        // アップロードファイル名を取得、配列に追加
        upLoadUrls.forEach( el => {
          id = el.split('=')[1]                     // 取得したアップロードファイルのurlからID部分のみ抽出
          uploadFile = DriveApp.getFileById(id);    // IDよりファイルを取得
          upLoadName = uploadFile.getName();        // アップロードされたファイル名を取得(名前有)
          upLoadNames.push(upLoadName);             // アップロードされたファイル名(名前有)を配列に追加
        })
      }

      console.log("upLoadNames:" + upLoadNames);
      console.log("upLoadUrls:" + upLoadUrls);

      // オブジェクトに追加
      this.upLoadUrls  = upLoadUrls;    // アップロードファイルURL(生データ)
      this.upLoadNames = upLoadNames;   // アップロードファイル名(生データ)

    };  // GetFileData()_END



    // ========== [ メソッド ] 指定フォルダにアップロードファイルのコピーを追加 ========== //

    FileMove() {

      console.log("FileMove 実行!");
    
      // 変数の定義・初期化
      let copyId;               // コピーファイルのID
      let fileName;             // コピー元のファイル名
      let fileName0, fileName1; // コピー元のファイル名(分割)
      let fileUrl;              // コピーファイルURL
      let fileRename;           // 名前が付けられる前の元々のファイル名
      let keywordName;          // キーワードファイルのファイル名
      let fileNames    = [];    // 配列 [コピー元のファイル名]
      let fileUrls     = [];    // 配列 [コピーファイルURL]
      let fileRenames  = [];    // 名前が付けられる前の元々のファイル名の配列
      let keywordNames = [];    // 配列 [キーワードファイルのファイル名]

      let folder = DriveApp.getFolderById(upfolder);         // コピー元ファイルのフォルダ
      let files  = folder.getFiles();                        // コピー元ファイル;
    
      // 登録先フォルダ名と回答フォームの内容が一致したフォルダIDを[copyId]に格納する。
      contentsIds.forEach( el => {
        if ( fileContents === el.contents ) copyId = el.id;
      })


      for ( let i = 0; i < 50; i++ ) {
        if ( !files.hasNext() ) {
          console.log("フォルダ内にファイルが存在しませんでした！");
          SleepTimer();
        }
      }


      // 指定フォルダ内にファイルが存在しない場合は、5秒間スリープ   *** 対策ソフト_20210219 ***
      function SleepTimer() {
        let now = Utilities.formatDate(new Date(), "Asia/Tokyo", "HH:mm:ss");   // 現在の時間
        console.log("now(現在時間):" + now);

        folder = DriveApp.getFolderById(upfolder);                // コピー元ファイルのフォルダ
        files  = folder.getFiles();                               // コピー元ファイル
        let staySecond = 5;                                     // 遅延時間の設定(秒)
        Utilities.sleep(staySecond * 1000);                     // 遅延実行
      };

      // アップロード関連のファイルが存在した場合、コピーして指定フォルダに保存(URLを取得)
      while(files.hasNext()) {
        let file  = files.next();                             // ファイルを取得
        fileName  = file.getName();                           // ファイル名を取得
        fileName0 = fileName.split(' - ')[0];                 // アップロードファイル名(名前無 ・ 拡張子無)
        fileName1 = fileName.split(' - ')[1];                 // アップロードファイルの登録者名以降の文字列を抽出

        // 登録者名を削除したファイル名を取得
        if ( fileName.indexOf(".") !== -1 ) {
          const extension  = fileName1.split('.')[1];         // 拡張子を抽出
          fileRename = `${fileName0}.${extension}`;           // ファイルに拡張子がある場合
        } else {
          fileRename = fileName0;                             // ファイルに拡張子がない場合
        }

        keywordName = `${fileName0}_キーワード.pdf`;            // キーワードファイル名
        const fileNum = this.upLoadNames.indexOf(fileName);   // ファイル名と一致した配列番号
        const nameJudge = ( fileNum !== -1 );                 // ファイル名の有無判定

        // ログ確認用
        console.log("fileName(ファイル名 名前有・拡張子含):" + fileName);
        console.log("fileName0(ファイル名 名前無・拡張子無):" + fileName0);
        console.log("fileName1(ファイル名 名前のみ・拡張子含):" + fileName1);
        console.log("fileRename(登録者名を削除したファイル名):" + fileRename);
        console.log("keywordName(キーワードファイル名):" + keywordName);
        console.log("fileNum(アップロードファイルの格納された配列と取得したファイル名が一致した配列番号):" + fileNum);
        console.log("upLoadNames(アップロードファイル名):" + this.upLoadNames);
        console.log("nameJudge(ファイル名一致判定):" + nameJudge);
          
        // ファイル名が一致したら実行
        if ( nameJudge ) {
          console.log("ファイル名が一致しました！");
          const copyFolder = DriveApp.getFolderById(copyId);                   // コピーファイルの保存先フォルダID
          const originalFile = DriveApp.getFilesByName(fileName).next();       // コピー元のオブジェクトを取得
          const newFile = originalFile.makeCopy(fileRename, copyFolder);       // 指定したフォルダにファイルをコピー
          fileUrl = newFile.getUrl();                                          // ファイルのURLを取得

          // 配列に追加
          fileNames.push(fileName);                                            // ファイル名(元データ)を配列に追加
          fileRenames.push(fileRename);                                        // ファイル名(コピーデータ)を配列に追加
          fileUrls.push(fileUrl);                                              // ファイルURL(コピーデータ)を配列に追加
          keywordNames.push(keywordName);                                      // キーワードファイル名を配列に追加
        }
      }

      // console.log("fileUrls:" + fileUrls);

     // オブジェクトに追加　
     this.fileNames   = fileNames;        // ファイル名(元データ)
     this.fileRenames = fileRenames;      // ファイル名(コピーデータ)
     this.fileUrls    = fileUrls;         // ファイルURL(コピーデータ)
     this.keywordNames = keywordNames;    // キーワードファイル名

    };    // FileMove()_END





    // ========== [ メソッド ] フォーム回答をキーワードファイル(スプレットシート)に書込 ========== //
     
    WhiteForm() {

      console.log("WhiteForm 実行!");
  
      // 変数の定義・初期化
      const datas       = getDatas;      // フォームの取得データ配列getDatasを格納
      const setSubDatas = dataTitle;     // キーワードファイルに記載するデータのタイトル（A列）
      const notDatas    = getNotDatas;   // キーワードファイルに記載不要なデータ
      let   setDatas    = [];            // キーワードファイルに書き込むデータ（B列）

      // [getDatas] と [notDatas] 内のデータを比較し、重複していないデータを [setDatas] に追加
      datas.concat(notDatas).forEach( data => {
        if ( datas.includes(data) && !notDatas.includes(data)) setDatas.push(data);
      });

      // キーワード説明文を書込 (スプレットシートのA列)
      let num = 1; // 書込み行(初期設定)

      // 配列[setSubDatas]に入っている説明文を書込
      setSubDatas.forEach( setSubData => {
        keyForm.getRange(`A${num}`).setValue(setSubData); // setSubDataを書込む 
        num++;                                            // 改行
      });

      // キーワードを書込 (スプレットシートのB列)  
      num = 1; // 書込み行(初期設定)

      // 配列[setDatas]に入っているキーワードを書込
      setDatas.forEach( setData => {
        if ( setData != "" ) {
          if ( setData === littleInfoTitle ) num += 2;  // setDataが 豆知識タイトルだった場合は改行
          if ( setData === littleInfoConts ) num++;     // setDataが 豆知識・プチ情報内容だった場合は改行
        }

        keyForm.getRange(`B${num}`).setValue(setData);  // setDataを書込む
        if ( setData !== "" ) num++;                    // setDataの中身が空でない場合は改行
      });

      // ファイル名・URLを書込
      if ( uploadFileUrl !== "" ) {
        let i = 0;
        this.fileUrls.forEach( fileUrl => {
          keyForm.getRange(`B${num}`).setValue(this.fileRenames[i]);
          num++;          
          keyForm.getRange(`B${num}`).setValue(fileUrl);
          num++;
          i++;
        })
      }

      // 問い合わせ・要望の項目に回答があったら、問い合わせ・要望のスプレットシートに内容を書込
      if ( requests !== "" ) {

        console.log("問い合わせ・要望あり！");

        // スプレットシート情報
        const ssForm = SpreadsheetApp.openById('1gctMJ1s7HJ51XQojAJDITSD9cab9hSB8f4CHBlYo94g');  // スプレットシート情報（書込み先）
        const reqForm = ssForm.getSheetByName('問い合わせ・要望');                                  // スプレットシート（問い合わせ・要望）情報
        const lastRow = reqForm.getLastRow();                                                    // スプレットシートの最終行目を取得する。
  
        // スプレットシートに記入
        const reqA  = reqForm.getRange(lastRow+1,  1).setValue(timeStamp);    // タイムスタンプ
        const reqB  = reqForm.getRange(lastRow+1,  2).setValue(mailAddress);  // メールアドレス
        const reqC  = reqForm.getRange(lastRow+1,  3).setValue(requests);     // 問い合わせ・要望 

        // オブジェクトに追加
        this.timeStamp   = timeStamp;    // 日時
        this.mailAddress = mailAddress;  // メールアドレス
        this.requests    = requests;     // 問い合わせ内容
      } else {
        console.log("問い合わせ・要望なし！");
      }

    };  // WhiteForm()_END



    // ========== [ メソッド ] キーワードフォーム(スプレットシート)をPDF変換して指定フォルダに追加 ========== //
     
    PdfCreate() {

      console.log("PdfCreate 実行!");
  
      // キーワードフォーム(スプレットシート)をpdf変換
      SpreadsheetApp.flush();
      const sheetId = keyForm.getSheetId(); // スプレットシートのIDを取得
      const url = 'https://docs.google.com/spreadsheets/d/1zzq1OQxTZJOFo2eFaTGkNLyipGWDjl9i-LijcgIU3Hw/export?exportFormat=pdf&gid=SID'.replace('SID', sheetId);
      const token = ScriptApp.getOAuthToken();
      const response = UrlFetchApp.fetch(url, {
        headers:{
          'Authorization': 'Bearer '+token
        }
      });
  
      // 変数の定義
      let fileName;    // キーワードフォームのファイル名
      let folderId;    // アップロードファイルの保存先フォルダのID
      let blob;        // PDF
      let folder;      // PDFの保存先フォルダ
      let requestForm; // フォルダに保存したPDF
      let folderIdKey = '14YTMjpzTq77UDN-bPT_zAGELWJudJuli';  // キーワードフォームの保存先フォルダID
  
      /* ファイル内容が豆知識・プチ情報(アップロードファイルなし)の場合 //
      // 回答フォームからファイルのタイトルを作成・内容を書込。        //
      // 豆知識・プチ情報のフォルダに保存                          */
      if ( notUploadFile === 'なし' ) {
        fileName = `${littleInfoTitle}.pdf`; // ファイル名を指定
    
        // 保存先フォルダと回答フォームの内容が一致した内容のフォルダIDを [folderId] に格納
        contentsIds.forEach( el => {
          if ( fileContents === el.contents ) folderId = el.id;
        });
          
        // pdfを指定フォルダに作成
        blob = response.getBlob().setName(fileName); // pdfの名前
        folder = DriveApp.getFolderById(folderId);   // pdfの保存先フォルダを指定
        requestForm = folder.createFile(blob);       // フォルダ内にpdfを作成

      // ファイル内容が豆知識・プチ情報以外の場合
      } else {

        // キーワードファイルを指定フォルダに作成
        folderId = folderIdKey;                                  // ファイルの保存先フォルダID
        
        // キーワードファイルを作成
        this.keywordNames.forEach( keywordName => {
        blob = response.getBlob().setName(keywordName);          // pdfの名前
        folder = DriveApp.getFolderById(folderId);               // pdfの保存先フォルダを指定
        requestForm = folder.createFile(blob);                   // フォルダ内にpdfを作成
        });
      }

      const keywordUrl = requestForm.getUrl();                   // キーワードファイルのURLを取得

      // オブジェクトに追加
      this.keywordUrl = keywordUrl;

    };  // PdfCreate()_END



    // ========== [ メソッド ] キーワードファイル(スプレットシート)の内容をクリア ========== //

    SSRemove() {

      console.log("SSRemove 実行!");
  
      const keyFormLastRow = keyForm.getLastRow();               // スプレットシート（キーワード書込先）の最終行目を取得
      keyForm.getRange(1, 1, keyFormLastRow, 2).clearContent();  // 指定したセルのコンテンツを削除

    };  // SSRemove()_END



    // ========== [ メソッド ] アップロードした元ファイルをアップロードフォルダ内から削除 ========== //

    FileTrash() {

      console.log("FileTrash 実行!");

      const delFolder = DriveApp.getFolderById(upfolder)             // 削除したいファイルのフォルダ
      const delfiles = delFolder.getFiles();                         // 削除したいファイル
  
      // 指定したフォルダ内にアップロード関連のファイルがあれば削除
      while(delfiles.hasNext()) {
        const file = delfiles.next();                                // ファイルを取得
        const fileName = file.getName();                             // ファイル名を取得
        const nameJudge = this.upLoadNames.indexOf(fileName) !== -1; // ファイル名の有無判定
        if ( nameJudge ) {
          file.setTrashed(true);                                     // ファイルを削除
          console.log(`${fileName}を削除しました！`);
        }
      }

    };  // FileTrash()_END


    // ========== [ メソッド ] フォーム回答者にファイル保存完了通知を送信 ========== //

    SendMail() {
      
      console.log("SendMail 実行!");
      
      // アドレスが見つからない場合の送信先
      const errAddress = 'k.kamikura@isowa.co.jp';
      
      // 問合せ・要望に回答があった場合の送信先
      const contactAddress = 'k.kamikura@isowa.co.jp';
      
      // isowaビトのアドレスが記載されたスプレットシートを取得
      const ssId = SpreadsheetApp.openById('1r9Ok3NF0_lwNa2fzCcpttteEim_Kb79xOBLB8GsJjnA');
      const addressSS = ssId.getSheetByName('メールアドレス一覧（ISOWA）');
      const arrData   = addressSS.getDataRange().getValues();
      const addressUrl = 'https://docs.google.com/spreadsheets/d/1r9Ok3NF0_lwNa2fzCcpttteEim_Kb79xOBLB8GsJjnA/edit#gid=0';

      // アドレスリストの配列の行と列を入替
      const _ = Underscore.load();                          // アンダースコアを使用
      const arrTrans = _.zip.apply(_, arrData);             // 配列の行と列を入替
      const resNum = arrTrans[1].indexOf(registPerson);     // 回答者名と一致した行番号(開始No:0)

      // ファイルURL
      let setKeywordUrl, setUrls;
      if ( this.keywordUrl !== undefined ) setKeywordUrl = this.keywordUrl;
      if ( this.fileUrls !== undefined ) setUrls = this.fileUrls;

      // 問い合わせ・要望フォームのURL
      const requestsUrl = "https://forms.gle/GkYxu4ENEmq2RzvX9";

      // メール送信用
      let reply;     // メール送信先
      let title;     // メールタイトル
      let content;   // メール本文

      // アップロードファイル名とURLの文言
      let setNameUrl = SetNamesUrls(this.fileRenames) 


  // 登録者(C)にアップロード完了通知メールを送信
      if ( resNum !== -1 ) {    
        reply   = arrTrans[2][resNum];
        title   = '【iサーチ】ファイルのアップロード完了通知';
        content = '\
    ${registPerson}さん\n\
    \n\
    いつもお仕事お疲れ様です。\n\
    iサーチに貴重な資料の登録をありがとうございます。\n\
    キーワードファイル：\n\
    ${setKeywordUrl}\n\n\
    アップロードファイル：\n\
    ${setNameUrl}\n\
    \n\
    おかげ様でISOWAビトが安心して業務を行う事ができます。\n\
    \n\
    改善・不明点等ありましたら、お気軽に問い合わせフォーム\n\
    よりご連絡下さい。\n\
    ${requestsUrl}'
    .replace('${registPerson}', registPerson)
    .replace('${setKeywordUrl}', setKeywordUrl)
    .replace('${setNameUrl}', setNameUrl)
    .replace('${requestsUrl}', requestsUrl);

        // メールを送信
        MailContents( reply, title, content );

  // 登録者のアドレスが不明な場合、errAddressにメール送信
      } else {     
        reply   = errAddress;
        title   = '【iサーチ】アップロードファイル登録者のアドレスが見つかりませんでした。';
        content = '\
    iサーチ運営チームのみなさま\n\
    \n\
    いつもお仕事お疲れ様です。\n\
    iサーチにファイルをアップロードした${registPerson}さんの\n\
    アドレスが見つかりませんでした。\n\
    キーワードファイル：\n\
    ${setKeywordUrl}\n\n\
    アップロードファイル：\n\
    ${setNameUrl}\n\
    アドレスリスト：${addressUrl}\n\
    \n\
    登録者名とアドレスリスト(スプレットシート)を\n\
    確認お願いします。'
    .replace('${registPerson}', registPerson)
    .replace('${setNameUrl}', setNameUrl)
    .replace('${setKeywordUrl}', setKeywordUrl)
    .replace('${addressUrl}', addressUrl);
                  
        // メールを送信
        MailContents( reply, title, content );
      };
      
      
  // 問合せ・要望に回答があった場合、メールを送信
      if ( requests !== "" ) {
        reply   = contactAddress;
        title   = '【iサーチ】フォームから問い合わせ・要望がありました。';
        content = '\
    iサーチ運営チームのみなさま\n\
    \n\
    いつもお仕事お疲れ様です。\n\
    iサーチにファイルをアップロードした${registPerson}さんから\n\
    以下の問い合わせ・要望がありました。\n\
    \n\
    ${requests}\n\
    よろしくお願いします。\n\
    '
    .replace('${requests}', requests)
    .replace('${registPerson}', registPerson)
                  
        // メールを送信
        MailContents( reply, title, content );
      }
      
      // [関数]アップロードファイル名・URLの文言
      function SetNamesUrls(fileRenames) {

        // 変数の定義・初期化
        let nameAndUrl;
        let i = 0;
        
        // ファイル名・URLを書込
        setUrls.forEach( setUrl => {
          const name = fileRenames[i]; 
          const url = setUrl;

          // 初回実行
          if ( i === 0) {
            nameAndUrl = `${name}\n\ 　${url}\n\
  `
          // 2回目以降
          } else {
            nameAndUrl += `  ${name}\n\ 　${url}\n\
  `
          }
          i++;  // 次のファイルを取得
        })

        return nameAndUrl;  // setNameUrlへ返却
      };  // SetNamesUrls()_END


      // [関数]メール送信
      function MailContents( reply, title, content ) {
        
        console.log("MailContents 実行!");
        
        // メールの送信情報
        const to = reply;                               // 送信先
        const subject = title;                          // タイトル
        const body = content;                           // 本文
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
      
      };  // MailContents()_END
    };    // SendMail()_END
  };      // Class Obj()_END


  // === オブジェクトを作成(アップロードファイル・キーワードファイルを指定フォルダに作成) === //

  const objArr = [];     // 配列を初期化
  const obj = new Obj(); // オブジェクト{obj}作成
  obj.GetFileData();     // アップロードファイル・キーワードファイル名を変更
  obj.FileMove();        // 指定フォルダにアップロードファイルのコピーを追加
  obj.WhiteForm();       // フォーム回答をキーワードファイル(スプレットシート)に書込
  obj.PdfCreate();       // キーワードファイル(スプレットシート)をPDF変換して指定フォルダに追加
  obj.SSRemove();        // キーワードファイル(スプレットシート)の内容をクリア
  obj.FileTrash();       // アップロードした元ファイルをアップロードフォルダ内から削除
  obj.SendMail();        // フォーム回答者にファイル保存の完了通知を送信

  objArr.push(obj);      // 配列[objArr]にオブジェクトを追加
  console.log(objArr);   // ログ確認用



};   // Main()_END


/* 
・ アップロードファイルのファイル名取得(拡張子無し)                   >>> 完了
・ キーワードファイル(スプレットシート)にキーワードを書込              >>> 完了  
・  〃  をファイル名_キーワードにrename・pdf変換し、指定フォルダに保存 >>> 完了
・ アップロードファイルを指定のフォルダに移動                        >>> 完了
・ アップロードファイルのURLをスプレットシートに書込                  >>> 完了
・ キーワードファイル(スプレットシート)の内容をクリア                 >>> 完了
・ アップロードファイルを削除                                     >>> 完了
・ 配列[keyword], [getData]の内容確認                           >>> 完了
・ 取得した情報が変更した場合でもすぐに修正できる様にする。            >>> 完了
・ 同時実行時の処理対策                                          >>> 完了
・ アップロードファイル名を変更(登録者名を削除)                      >>> 完了
・ メールを送信                                                 >>> 完了
・ エラーとなった場合にエラー内容をメールで通知。                     >>> 完了
・ 登録者名にスペースがあった場合は、スペースを削除して格納する。        >>> 完了
・ アップロードファイルを複数登録する場合に対応する。                  >>> 完了
・ アップロードファイルがGOOGLEドライブのフォルダに保存されるまで
　 タイムラグがある。～約40秒
   >>> フォルダ内にファイルが無かった場合、5秒間待って再度ファイル
   　　 の存在確認を行う処理を追加。

メール通知ソフトについて
・ 問い合わせ・要望のリンクを貼った方が親切。  >>> 完了


*/