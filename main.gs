// ------------------------------------------------------- //
// スプレットシートを取得                                        //
// ------------------------------------------------------- //
const getSS = SpreadsheetApp.openById('1b10DCcyucIn7v2qa1SOIPID_K8DC9KU4hRybq6feHbI'); // スプレットシート情報（読出し）
const form  = getSS.getSheetByName('フォームの回答');                                     // スプレットシート（フォーム回答）情報
const UrlSS = getSS.getSheetByName('ファイルURL');                                       // スプレットシート（ファイルURL）情報
const setSS = SpreadsheetApp.openById('1gwfUf30jWIMy-emXiJLN4ZwmXU2cXwtp4EUfD4slWxE'); // スプレットシート情報（書込み）
let keyForm = setSS.getSheetByName('キーワードフォーム');                                   // スプレットシート（キーワードフォーム）情報

const mergeman = '0B2IK-M_HFyeqfkpuOVpVdFA4aUgtU2EtTGtxNXJlWWJOajVlcG9PczIyM3BUZWhsbERiOFU'; //結合するPDFが入ってるフォルダID
const formLastRow = form.getLastRow();   // スプレットシート（フォーム回答）の最終行目の情報を取得する。

// ------------------------------------------------------- //
// Googleフォームからスプレットシートに書き込まれた各セル情報を取得       //
// ------------------------------------------------------- //

const __A = form.getRange(formLastRow,  1).getValue();  
const _A = new Date(__A);
const A = Utilities.formatDate(_A, 'JST', 'yyyy/M/d');  // タイムスタンプ
const B  = form.getRange(formLastRow,  2).getValue();   // ファイル選択
const C  = form.getRange(formLastRow,  3).getValue();   // ファイル内容
const D  = form.getRange(formLastRow,  4).getValue();   // 製函機の種類
const E  = form.getRange(formLastRow,  5).getValue();   // 機械の種類
const F  = form.getRange(formLastRow,  6).getValue();   // コルゲータの種類
const G  = form.getRange(formLastRow,  7).getValue();   // 資料内容（その他）
const H  = form.getRange(formLastRow,  8).getValue();   // シングルフェーサの機種
const I  = form.getRange(formLastRow,  9).getValue();   // スプライサの機種
const J  = form.getRange(formLastRow, 10).getValue();   // ミルロールスタンドの機種
const K  = form.getRange(formLastRow, 11).getValue();   // ブレーキスタンドの機種
const L  = form.getRange(formLastRow, 12).getValue();   // プレヒータの機種
const M  = form.getRange(formLastRow, 13).getValue();   // グルーマシンの機種
const N  = form.getRange(formLastRow, 14).getValue();   // ダブルフェーサの機種
const O  = form.getRange(formLastRow, 15).getValue();   // スリッタスコアラの機種
const P  = form.getRange(formLastRow, 16).getValue();   // カッターの機種
const Q  = form.getRange(formLastRow, 17).getValue();   // スタッカーの機種
const R  = form.getRange(formLastRow, 18).getValue();   // アイビス・ファルコンの機種
const S  = form.getRange(formLastRow, 19).getValue();   // 機器・部品の種類
const T  = form.getRange(formLastRow, 20).getValue();   // 機器・部品名（機械）
const U  = form.getRange(formLastRow, 21).getValue();   // 機器・部品名（電気）
const V  = form.getRange(formLastRow, 22).getValue();   // 業務マニュアル
const W  = form.getRange(formLastRow, 23).getValue();   // 作成者（フルネーム）
  let X  = form.getRange(formLastRow, 24).getValue();   // 関連付けしたいキーワード
const Y  = form.getRange(formLastRow, 25).getValue();   // 登録者（フルネーム）
const Z  = form.getRange(formLastRow, 26).getValue();   // 豆知識・プチ情報
const AA = form.getRange(formLastRow, 27).getValue();   // 社内業務マニュアル
const AB = form.getRange(formLastRow, 28).getValue();   // お客様名
const AC = form.getRange(formLastRow, 29).getValue();   // 管理装置の機種
const AD = form.getRange(formLastRow, 30).getValue();   // 豆知識・プチ情報のタイトル
const AE = form.getRange(formLastRow, 31).getValue();   // フレキソの機種
const AF = form.getRange(formLastRow, 32).getValue();   // プリスロの機種
const AG = form.getRange(formLastRow, 33).getValue();   // その他（製函機）
const AH = form.getRange(formLastRow, 34).getValue();   // 他社製・取売り機（製函機）
const AI = form.getRange(formLastRow, 35).getValue();   // フォルダーグルアの機種
const AJ = form.getRange(formLastRow, 36).getValue();   // カウンターエジェクタの機種
const AK = form.getRange(formLastRow, 37).getValue();   // スタッカーの機種（製函機）
const AL = form.getRange(formLastRow, 38).getValue();   // 管理装置の機種（製函機）
const AM = form.getRange(formLastRow, 39).getValue();   // 予備
const AN = form.getRange(formLastRow, 40).getValue();   // 予備
const AO = form.getRange(formLastRow, 41).getValue();   // 予備
const AP = form.getRange(formLastRow, 42).getValue();   // 予備
const AQ = form.getRange(formLastRow, 43).getValue();   // 予備
const AR = form.getRange(formLastRow, 44).getValue();   // 予備
const AS = form.getRange(formLastRow, 45).getValue();   // 予備
const AT = form.getRange(formLastRow, 46).getValue();   // 予備
const AU = form.getRange(formLastRow, 47).getValue();   // 予備
const AV = form.getRange(formLastRow, 48).getValue();   // 予備
const AW = form.getRange(formLastRow, 49).getValue();   // 予備
const AX = form.getRange(formLastRow, 50).getValue();   // 予備
const AY = form.getRange(formLastRow, 51).getValue();   // 予備
const AZ = form.getRange(formLastRow, 52).getValue();   // 予備


if ( X === "" ) X = "-------"; 

// googleフォームからの取得データを配列に格納。
const getDatas = [ A, Y, W, X, B, C, D, E, F, G, H, I, J, K,
                   L, M, N, O, P, Q, R, S, T, U, V, AA, AB,
                  AC, AE, AF, AG, AH, AI, AJ, AK, AD, AE, AF,
                  AG, AH, AI, AJ, AK, AL, Z ];
 

// アップロードファイル内容と保存先フォルダID
const a = { contents:"取説・マニュアル（ISOWAオリジナル）", id:"1X2zHJweXS8n-rLvE8medP86GeordrxYB" };
const b = { contents:"機器・部品マニュアル（メーカー）",    id:"10YmfTyhQjhwoydZojvcWuBoBH3B2ET7b" };
const c = { contents:"手順書",                     id:"1yDxGP9qByBMKxyXJ5SuYqXZMwRnGsFR3" };
const d = { contents:"調整要領書",                  id:"1_bbbPwpDckPiCULxuuaBDOJPS_m6rV7X" };
const e = { contents:"トラブルシューティング",          id:"11z8g_6fx0_HctB38n3rXcJSqIye9Ned8" };
const f = { contents:"仕様書",                     id:"1_aXRNswdLe6FBWn1hz4NGA0DBah_i9Fd" };
const g = { contents:"報告書",                     id:"1YYEHNVZGgCrQodGSVlwAMjHfamT77ibt" };
const h = { contents:"見解書",                     id:"1OYsfa463huf9_6ke6EuAUp_yyLz4D1aH" };
const i = { contents:"点検・検査表",                 id:"1ngXN4An07NkWzMeI-qZp2NOp5NVYdnya" };
const j = { contents:"設計変更",                   id:"14RvghAV4G8TP0juj-cRqefspohUvcWly" };
const k = { contents:"アイレポ（修理・組立）",           id:"1qwm-LfZnhzdFqpwfjCWdEREYY_2YRVZZ" };
const l = { contents:"豆知識・プチ情報",              id:"1BGfaGgm_JkEa-VvQW3iuOJO3cjDLrsil" };
const m = { contents:"社内資料（業務マニュアル）",       id:"1iQU8jhGuMnS0jc7VHUQoPc6q8ol_wD7E" };
const n = { contents:"社内資料（フォーム）",            id:"1YEllBfGu5k7rPaX8LGsr8L4i3Drx1Fpt" };
const o = { contents:"その他",                     id:"1NjwPcpJfPqewkoWYMQUs3Cy9LbFt3cPC" };
const p = { contents:"画像",                      id:"1WP77ffzOalWNoaGj7nPSv2aVkfsHLImG" };
const q = { contents:"動画",                      id:"14HfgOm2nRKgnRzFFOv84QWSPMxqOqrc4" };
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
const contentsIds = [ a, b, c, d, e, f, g, h, i, j, k,l, m, n, o, p, q, r, s, t, u, v,
                      w, x, y, z ];


let id;         // アップロードID
let uploadFile; // アップロードファイル
let upLoadName; // アップロードされたファイル名を取得(名前有)
let rename;     // アップロードファイル名(名前無 ・ 拡張子無)
let renameFile; // アップロードファイル名.pdf
let rename2;    // アップロード時に付いた名前以降の文字列
let renameExt;  // アップロードファイル名
let extension;  // 拡張子
let fileName;   // アップロードファイル名

let wordLeg;
let word;
let excel;
let sheets;
let docs;
let text;
let pdf;


// ファイル内容の真偽（条件式で使用）
const notImage = C !== "画像"; // 画像でないファイル 
const notVideo = C !== "動画"; // 動画でないファイル
const uploadFileExists = B !== ""; // アップロードファイル有
const image = C == "画像"; // 画像ファイル
const video = C == "動画"; // 動画ファイル
const memo  = AD !== "";   // 豆知識・プチ情報 有 


// -------------------------------------------------- //
//            アップロードされたファイル名を変更する            //
// -------------------------------------------------- //

// アップロードファイルが存在する場合
if ( uploadFileExists ) {

  id = B.split('=')[1]                   // 取得したアップロードファイルのurlからID部分のみ抽出
  uploadFile = DriveApp.getFileById(id); // IDよりファイルを取得
  upLoadName = uploadFile.getName();     // アップロードされたファイル名を取得(名前有)  
  rename = upLoadName.split(' - ')[0];   // アップロードファイル名(名前無 ・ 拡張子無)
  renameFile = `${rename}.pdf`;          // アップロードファイル名.pdf
  rename2 = upLoadName.split(' - ')[1];  // アップロード時に付いた名前以降の文字列を抽出
  
  // 拡張子がある場合
  if ( rename2.indexOf(".") !== -1 ) {
    extension  = rename2.split('.')[1];        // 拡張子を抽出
    personName = rename2.split('.')[0];        // 登録者名を抽出
    renameExt  = `${rename}.${extension}`;     // アップロードファイル名(名前無 ・ 拡張子有) ??
    renameExt2 = `${rename} - ${personName}`;  // アップロードファイル名(名前有 ・ 拡張子無) ??
  
  // 拡張子がない場合
  } else {
    renameExt  = rename;      // アップロードファイル名(名前無 ・ 拡張子無)
    renameExt2 = upLoadName;  // アップロードファイル名(名前有 ・ 拡張子無)
  };
  
  // アップデートファイル内容の真偽（条件式で使用）
  wordLeg = uploadFile.getMimeType() == MimeType.MICROSOFT_WORD_LEGACY; // ワードファイル
  word    = uploadFile.getMimeType() == MimeType.MICROSOFT_WORD;        // ワードファイル
  excel   = uploadFile.getMimeType() == MimeType.MICROSOFT_EXCEL;       // エクセル
  sheets  = uploadFile.getMimeType() == MimeType.GOOGLE_SHEETS;         // スプレットシート
  docs    = uploadFile.getMimeType() == MimeType.GOOGLE_DOCS;           // ドキュメント
  text    = uploadFile.getMimeType() == MimeType.PLAIN_TEXT;            // テキストファイル
  pdf     = uploadFile.getMimeType() == MimeType.PDF;                   // pdf
    
};




// ------------------------------------------------------ //
//  関数を実行                                              //
// ------------------------------------------------------ //
function Main() {
  
  const lock = LockService.getDocumentLock(); //ドキュメントロックを使用する
  
  //360秒間のロックを実施 ( 複数実行された時の対策 )
  if (lock.tryLock(360000)) {
    try {
      
      WhiteForm(); // 回答フォーム情報をキーワードフォームに書込む。
      
      // アップロードファイルが存在する場合
      if ( uploadFileExists ) {
        
        // 画像 ・ 動画 ・ pdf 以外の場合
        if ( !image && !video && !pdf ) {
          PdfConvert();    // アップロードされたファイルをPDFに変換
          
        // 画像 ・ 動画 の場合
        } else if ( image || video ) {
          WhiteFileName(); // アップロードファイル名を変更
        };
        
      };
      
      
      // エクセル ・ スプレットシート ・ pdf ・ 豆知識 の場合
      if ( excel || sheets || pdf || memo ) {
        PdfCreate(); // キーワードフォームをPDFに変換
      }
      SSRemove(); // WhiteForm()で書き込んだ情報をシートから削除
      
     
      // アップデートファイルが存在する場合
      if ( uploadFileExists ) {
        
        // エクセル ・ スプレットシート ・ pdf の場合
        if ( excel || sheets || pdf ) {
          PdfMerge(); // pdf結合を実行 (アップロードファイル + キーワードフォーム)
          
        // エクセル ・ スプレットシート ・ pdf 以外の場合
        } else {
          FileMove(); // pdfを指定フォルダに移動
        };
        
      };
      
    } catch(e) {
      Logger.log("エラーが発生しました！"); // エラー時の処理を記述
    } finally {
      lock.releaseLock(); // ロックを開放
    };
    
  };
  
};