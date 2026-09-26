/**
 * ===================================================================
 * アーム機種別出荷明細PDF 自動生成（Google Apps Script 単体版 / 日またぎ防止）
 * ------------------------------------------------------------
 * 方式: HTML→PDF ＋ 自前ページ制御（バケット方式）
 *   ・1ページに載せる行数を固定（ROWS_PER_PAGE=52）し、入るだけ複数日を詰めて区切りで強制改ページ。
 *     → 平均的な2日分（約24行×2）は1枚にまだまる。入り切る限り2日・3日と同一ページに載る。
 *     → 同一出荷日がページの変わり目で分断されない（reportlab の KeepTogether 相当）。
 *     → 余りは下部の余白になる（日またぎ防止を優先）。
 *   ・1日だけで ROWS_PER_PAGE を超える超大量日は、その日だけ自然に複数ページへ（不可避）。
 *
 * 処理: マスター「アーム出荷明細」フォルダの最新「日程表変更」.xlsm を
 *   スプレッドシートに一時変換して読み取り → 実行日から直近31日を抽出 →
 *   体裁付きHTMLを生成 → PDF化 → 「①出荷作業用」に保存。元Excelは変更しない。
 *
 * 列: 検査完了日 / 機器 / 機種 / 図番 / 号機 / 仕様 / 最新出荷日 / 情報① / 出荷先（9列）
 * 体裁: 13ton行=黄色 / 出荷日グループ先頭に赤い区切り線 / 見出しは各ページ上部に繰り返し /
 *       最新出荷日はグループ内の全行に表示（先頭=大きく黒／継続=小さくグレー）/ ブームブラケット除外 / 号機は整数表示
 *
 * 実行タイミング: 元の.xlsmは別の自動化がメール添付から本フォルダへ保存する。
 *   このスクリプトは POLL_INTERVAL_MINUTES おきに対象ファイルをチェックし、
 *   「前回PDF化した時と同じファイル(ID+更新日時)」なら何もしない。
 *   実際にファイルが更新された時だけPDFを再生成する（毎日律儀に作り直さない）。
 *
 * セットアップ（初回のみ）:
 *   1. script.google.com に貼付
 *   2. 「サービス ＋」→ Drive API(v3・識別子 Drive) を追加
 *   3. プロジェクト設定でタイムゾーン=Asia/Tokyo
 *   4. installPdfTrigger を1回実行（権限許可）→ 以後 POLL_INTERVAL_MINUTES おきに自動チェック
 *   5. すぐ作るときは generateArmPDF を手動実行
 *
 * 調整ポイント:
 *   ・ページの変わり目で日付が割れる/はみ出す場合は ROWS_PER_PAGE を1〜2下げる（例:52→50）。
 *   ・もっと詰めたい場合は少し上げる（実データ検証では56行まで日またぎ無し）。
 *   ・行高/文字サイズを詰めた密度版。仕様欄が長い日が多いと1ページの実容量は変動するため、52は安全側の既定値。
 * ==========================================================
 */

// ==== 設定 ====
var SRC_FOLDER_ID = '1NS4WoClO0xlGWSxvFimFcqFGUT0jOKQL'; // 読み込み元：アーム出荷明細フォルダ(.xlsm)
var OUT_FOLDER_ID = '1-RrriADJSmBt2a7HsYpb_c8TUHHJpxkz'; // 保存先：①出荷作業用
var NAME_KEYWORD  = '日程表変更';
var SHEET_NAME    = '出荷明細';
var DAYS          = 31;
var ROWS_PER_PAGE = 52;          // 1ページの最大データ行数。全列8ptに統一した密度での既定値。
                                 //   ・日付が割れる/はみ出す → 1〜2下げる
                                 //   ・もっと詰めたい → 少し上げる
var TZ            = 'Asia/Tokyo';
var POLL_INTERVAL_MINUTES = 15;  // 対象ファイルの更新チェック間隔
var NOTIFY_EMAIL  = Session.getEffectiveUser().getEmail(); // 対象ファイルが見つからない時の通知先

// 出荷明細シートの列(1-based)
var COL = { insp:5, kk:6, zu:9, go:12, spec:13, ship:33, info:40, dest:43 };

var LAST_SOURCE_KEY   = 'ARM_LAST_SOURCE_SIGNATURE';
var LAST_NOTIFY_KEY   = 'ARM_LAST_NOTIFY_DATE';

var SNAPSHOT_FILE_NAME = 'arm_pdf_snapshot.json'; // 前回生成時点の内容（変更点の赤字判定用）

// 同一出荷日内の出荷先の並び順。無い出荷先はスキップされるだけなので問題ない。
// リストに無い出荷先は末尾に回す。
var DEST_ORDER = ['あゆみ', '本間', '東条', '正和', '正和(13ton)'];


/** 手動実行用エントリポイント。見つかったファイルで無条件に生成する。 */
function generateArmPDF(){
  var res = resolveSource_();
  if(!res.file){
    Logger.log('対象ファイル（日程表変更 .xlsm）が見つかりません');
    notifySourceNotFound_();
    return;
  }
  var ok = generateArmPDF_core_(res.file);
  if(ok) saveSourceSignature_(res.file);
}

/**
 * ポーリング用エントリポイント。トリガーはこちらを呼ぶ。
 * 前回PDF化した時と同じファイル(ID+更新日時)なら何もしない。
 */
function checkAndGenerateArmPDF(){
  var res = resolveSource_();
  if(!res.file){
    Logger.log('対象ファイル（日程表変更 .xlsm）が見つかりません');
    notifySourceNotFound_();
    return;
  }

  var sig = sourceSignature_(res.file);
  var props = PropertiesService.getScriptProperties();
  if(props.getProperty(LAST_SOURCE_KEY) === sig){
    Logger.log('変更なし（' + res.file.getName() + '）のためスキップ');
    return;
  }

  Logger.log('更新を検知: ' + res.file.getName());
  var ok = generateArmPDF_core_(res.file);
  if(ok) props.setProperty(LAST_SOURCE_KEY, sig);
}


/** 実際の生成処理。成功時true、対象データ0件などでスキップした場合false。 */
function generateArmPDF_core_(src){
  Logger.log('元ファイル: '+src.getName());
  cleanupStaleTempFiles_();

  // xlsm → スプレッドシート（読み取り用の一時コピー）
  var conv = Drive.Files.create(
    { name:'tmp_arm_src', mimeType: MimeType.GOOGLE_SHEETS },
    src.getBlob(), { supportsAllDrives:true });
  var convId = conv.id;

  try{
    var ss = SpreadsheetApp.openById(convId);
    var sh = ss.getSheetByName(SHEET_NAME);
    if(!sh) throw new Error('シート「'+SHEET_NAME+'」が見つかりません');
    var lastRow = sh.getLastRow();
    if(lastRow < 6){ Logger.log('データ行なし'); return false; }
    var n = lastRow - 5;

    var insp = sh.getRange(6,COL.insp,n,1).getValues();
    var kk   = sh.getRange(6,COL.kk,  n,1).getValues();
    var zu   = sh.getRange(6,COL.zu,  n,1).getValues();
    var go   = sh.getRange(6,COL.go,  n,1).getValues();
    var spec = sh.getRange(6,COL.spec,n,1).getValues();
    var ship = sh.getRange(6,COL.ship,n,1).getValues();
    var info = sh.getRange(6,COL.info,n,1).getValues();
    var dest = sh.getRange(6,COL.dest,n,1).getValues();

    var start = midnight_(new Date());
    var end = new Date(start.getTime()); end.setDate(end.getDate()+DAYS);

    var rows=[];
    for(var i=0;i<n;i++){
      var s = ship[i][0];
      if(!(s instanceof Date)) continue;
      var sd = midnight_(s);
      if(!(sd.getTime()>=start.getTime() && sd.getTime()<end.getTime())) continue;
      var kkv=String(kk[i][0]||''), specv=String(spec[i][0]||'');
      if(kkv.indexOf('ブームブラケット')>=0 || specv.indexOf('ブームブラケット')>=0) continue;
      var kiki=kkv, kishu='', p=kkv.indexOf('　');
      if(p>=0){ kiki=kkv.substring(0,p); kishu=kkv.substring(p+1); }
      rows.push({ insp:mdOrStr_(insp[i][0]), kiki:kiki, kishu:kishu, zu:strv_(zu[i][0]),
                  go:numi_(go[i][0]), spec:specv, ship:sd, info:strv_(info[i][0]),
                  dest:strv_(dest[i][0]), is13:(kkv.indexOf('13')===0) });
    }
    if(rows.length===0){ Logger.log('該当データ0件のためPDFは作成しません'); return false; }

    // 前回生成時からの変更行を検出（赤字表示用）。図番＋号機で同一出荷物とみなす。
    var prevSnapshot = loadSnapshot_();
    markChanges_(rows, prevSnapshot);

    rows.sort(function(a,b){ return (a.ship.getTime()-b.ship.getTime()) || (destRank_(a.dest)-destRank_(b.dest)); });

    // 出荷日ごとにグループ化
    var groups=[], gi=0;
    while(gi<rows.length){
      var gj=gi;
      while(gj<rows.length && rows[gj].ship.getTime()===rows[gi].ship.getTime()) gj++;
      groups.push(rows.slice(gi,gj)); gi=gj;
    }

    // ページへバケット（グループは分割しない。単独で超過する日だけ自然フロー）
    // 各日の末尾に合計行（通常の行より背が高い）が付くので、グループの行数は +2 で数える。
    var pages=[], cur=[], curcnt=0;
    for(var gk=0; gk<groups.length; gk++){
      var g=groups[gk], gn=g.length+2;
      if(gn>ROWS_PER_PAGE){ if(cur.length){pages.push(cur);cur=[];curcnt=0;} pages.push([g]); continue; }
      if(curcnt+gn>ROWS_PER_PAGE){ pages.push(cur); cur=[g]; curcnt=gn; }
      else { cur.push(g); curcnt+=gn; }
    }
    if(cur.length) pages.push(cur);

    var rng = fmtJ_(start)+' 〜 '+fmtJ_(new Date(end.getTime()-86400000));
    var html = buildHtml_(pages, rng);

    var srcTag = srcDateTag_(src.getName()); // 日程表変更ファイル名に埋め込まれた日付（例:9月16日）
    var pdfName='アーム出荷明細'+(srcTag?'('+srcTag+')':'')+'_'+Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd')+'.pdf';
    var blob=Utilities.newBlob(html,'text/html','arm.html').getAs('application/pdf').setName(pdfName);

    var outFolder=DriveApp.getFolderById(OUT_FOLDER_ID);
    var ex=outFolder.getFilesByName(pdfName);
    while(ex.hasNext()) ex.next().setTrashed(true);   // 同名の古い版はゴミ箱へ
    var saved=outFolder.createFile(blob);
    Logger.log('保存完了: '+pdfName+' / 件数='+rows.length+' / ページ='+pages.length+' / '+saved.getUrl());

    saveSnapshot_(rows, prevSnapshot); // 次回比較用に今回の内容を保存
    return true;

  } finally {
    safeRemove_(convId);
  }
}


// ==== HTML生成 ====
function buildHtml_(pages, rng){
  var css =
    '@page{size:A4 portrait;margin:6mm;}'
  // 印刷エンジンは既定でbackground(塗り)を「インク節約」のため印刷しない。
  // 枠線(border)は対象外なので赤い区切り線だけ出て黄色い塗りが消える、という症状になる。
  // print-color-adjust:exactで塗りも忠実に出すよう強制する。
  + '*{font-family:"IPAGothic","IPAPGothic","Noto Sans JP",sans-serif;box-sizing:border-box;'
  + '-webkit-print-color-adjust:exact;print-color-adjust:exact;color-adjust:exact;}'
  + 'body{margin:0;}'
  + '.page.brk{page-break-after:always;break-after:page;}'
  + 'tbody.grp{page-break-inside:avoid;break-inside:avoid;}'    // 出荷日グループ単位で分断防止
  + '.title{text-align:center;font-size:13pt;font-weight:bold;margin:0 0 1px;}'
  + '.sub{text-align:center;font-size:9pt;margin:0 0 3px;}'
  + 'table{border-collapse:collapse;width:100%;table-layout:fixed;}'
  + 'thead th{background:#d2d2d2;border:0.4pt solid #000;font-size:8pt;text-align:center;vertical-align:middle;padding:1.5px;line-height:1.15;}'
  + 'td{border:0.4pt solid #000;font-size:8pt;vertical-align:middle;padding:1px 2.5px;overflow-wrap:anywhere;line-height:1.15;}'
  + 'tr.gtop td{border-top:1.6pt solid #d90000;}'               // グループ先頭の赤線
  + 'tr.r13 td{background:#ffff8c;}'                            // 13ton 黄色
  + '.c-insp{width:15mm}.c-kiki{width:21mm}.c-kishu{width:17mm}.c-zu{width:27mm}'
  + '.c-go{width:9mm;text-align:center}.c-spec{width:44mm}'
  + '.c-ship{width:18mm;text-align:center}.c-info{width:19mm}.c-dest{width:22mm}'
  + '.d1{font-size:8pt;font-weight:bold}.dc{font-size:8pt;color:#666}'
  + '.dprev{font-size:6.5pt;white-space:nowrap;}'                // 出荷日変更時の「旧日付」注記（改行位置が乱れないよう別行・小さめに）
  + 'tr.gsum td{background:#333;color:#fff;font-size:10pt;font-weight:bold;text-align:right;padding:3px 8px;border:1pt solid #000;}' // 1日ごとの出荷合計行（濃い帯）
  + '.gtot{color:#ffe14d;font-size:12pt;margin-left:12px;}'      // 合計本数だけさらに強調
  + 'tr.chg td, tr.chg td *{color:#d90000 !important;}';        // 前回から変更/新規の行は赤字

  var thead='<thead><tr>'
    +'<th class="c-insp">検査<br>完了日</th><th class="c-kiki">機器</th><th class="c-kishu">機種</th>'
    +'<th class="c-zu">図番</th><th class="c-go">号<br>機</th><th class="c-spec">仕様</th>'
    +'<th class="c-ship">最新<br>出荷日</th><th class="c-info">情報①</th><th class="c-dest">出荷先</th>'
    +'</tr></thead>';

  var body='';
  for(var pi=0; pi<pages.length; pi++){
    var pg=pages[pi];
    var last=(pi===pages.length-1);
    var cls='page'+(last?'':' brk');
    var tbodies='';
    for(var g=0; g<pg.length; g++){
      var grp=pg[g];
      var trs='';
      for(var k=0;k<grp.length;k++){
        var x=grp[k];
        var rowcls=[]; if(x.is13) rowcls.push('r13'); if(k===0) rowcls.push('gtop'); if(x.chg) rowcls.push('chg');
        // 出荷日が変わった行は「旧日付」を小さめの別行で注記する（1行に→でつなぐと折返し位置が乱れて読みにくいため）
        var prevLine = x.prevShip ? '<div class="dprev">(旧'+fmtJ_(x.prevShip)+')</div>' : '';
        var dcell = (k===0)
          ? '<div class="d1">'+fmtJ_(x.ship)+'</div>'+prevLine
          : '<div class="dc">'+fmtJ_(x.ship)+'</div>'+prevLine;
        trs+='<tr class="'+rowcls.join(' ')+'">'
          +'<td class="c-insp">'+esc_(x.insp)+'</td>'
          +'<td class="c-kiki">'+esc_(x.kiki)+'</td>'
          +'<td class="c-kishu">'+esc_(x.kishu)+'</td>'
          +'<td class="c-zu">'+esc_(x.zu)+'</td>'
          +'<td class="c-go">'+esc_(x.go)+'</td>'
          +'<td class="c-spec">'+esc_(x.spec)+'</td>'
          +'<td class="c-ship">'+dcell+'</td>'
          +'<td class="c-info">'+esc_(x.info)+'</td>'
          +'<td class="c-dest">'+esc_(x.dest)+'</td>'
          +'</tr>';
      }
      // 1日ごとの出荷本数の内訳と合計（1行=1本）。グループと同じ<tbody>に入れて一緒に改ページさせる。
      var sum = daySummary_(grp);
      trs+='<tr class="gsum"><td colspan="9">'+fmtJ_(grp[0].ship)+'　'+esc_(sum.items)
          +'<span class="gtot">合計 '+sum.total+'本</span></td></tr>';
      // 出荷日グループごとに<tbody>を分け、そのグループだけpage-break-inside:avoidする。
      // ページ全体をavoid指定すると、行の折り返しで見積もり行数(ROWS_PER_PAGE)を実際の高さが
      // わずかに超えた場合に、印刷エンジンが同一出荷日の途中で強制的にページを割ってしまう
      // （最後の1行だけ次ページに孤立する等）不具合があったため、分断禁止の単位をグループへ
      // 縮小し、はみ出しても頁またぎはグループの境界でのみ起きるようにする。
      tbodies+='<tbody class="grp">'+trs+'</tbody>';
    }
    body+='<div class="'+cls+'"><div class="title">アーム機種別出荷明細</div>'
        +'<div class="sub">出荷日 '+esc_(rng)+'（直近1ヶ月）</div>'
        +'<table>'+thead+tbodies+'</table></div>';
  }
  return '<!doctype html><html lang="ja"><head><meta charset="utf-8"><style>'+css+'</style></head><body>'+body+'</body></html>';
}


// ==== 対象ファイル解決 ====

/**
 * SRC_FOLDER_ID配下からNAME_KEYWORDを含む最新の.xlsmを探す。
 * ファイル名から日付が読み取れる場合はその日付が最新のものを、
 * 1件も読み取れない場合は更新日時が最新のものにフォールバックする
 * (命名規則がずれても「対象ファイルなし」で無言停止しないようにするため)。
 */
function resolveSource_(){
  var folder=DriveApp.getFolderById(SRC_FOLDER_ID);
  var it=folder.getFiles(), best=null, bestKey=-1, candidates=[];
  while(it.hasNext()){
    var f=it.next(), name=f.getName();
    if(name.indexOf(NAME_KEYWORD)<0) continue;
    if(name.toLowerCase().slice(-5)!=='.xlsm') continue;
    candidates.push(f);
    var k=dateKeyFromName_(name);
    if(k>bestKey){ bestKey=k; best=f; }
  }
  if(best) return { file: best, viaFallback: false };
  if(candidates.length>0){
    candidates.sort(function(a,b){ return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
    Logger.log('警告: ファイル名から日付を判定できないため、更新日時で最新の「'+candidates[0].getName()+'」を使用します');
    return { file: candidates[0], viaFallback: true };
  }
  return { file: null, viaFallback: false };
}

function sourceSignature_(f){
  return f.getId() + '::' + f.getLastUpdated().getTime();
}

function saveSourceSignature_(f){
  PropertiesService.getScriptProperties().setProperty(LAST_SOURCE_KEY, sourceSignature_(f));
}

/** 対象ファイルが見つからない時、1日1回までメール通知する(トリガーが連続で無言失敗しないように) */
function notifySourceNotFound_(){
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd');
  if(props.getProperty(LAST_NOTIFY_KEY) === today) return;
  try{
    MailApp.sendEmail(NOTIFY_EMAIL,
      '【アーム出荷明細PDF】対象ファイルが見つかりません',
      '「'+NAME_KEYWORD+'」を含む.xlsmファイルが下記フォルダに見つからないため、\n' +
      'アーム機種別出荷明細PDFを生成できませんでした。\n\n' +
      'フォルダ: https://drive.google.com/drive/folders/'+SRC_FOLDER_ID+'\n\n' +
      '(このメールは1日1回までの通知です)');
    props.setProperty(LAST_NOTIFY_KEY, today);
  }catch(e){
    Logger.log('通知メール送信に失敗: '+e.message);
  }
}


// ==== ヘルパ ====
/**
 * 1日分の出荷本数の内訳。13ton(黄色の行)は「13トン」、白い行は出荷先で
 * 正和→「ライン」、あゆみ・本間・東条はそれぞれ、それ以外は「その他」に数える。
 * 0本の項目は出さない。{items: 内訳の文字列, total: 合計本数} を返す。
 */
function daySummary_(grp){
  var c = { line:0, ayumi:0, honma:0, tojo:0, t13:0, other:0 };
  for(var i=0;i<grp.length;i++){
    var x = grp[i], d = String(x.dest).trim();
    if(x.is13) c.t13++;
    else if(d==='正和') c.line++;
    else if(d==='あゆみ') c.ayumi++;
    else if(d==='本間') c.honma++;
    else if(d==='東条') c.tojo++;
    else c.other++;
  }
  var items = [['あゆみ',c.ayumi],['本間',c.honma],['東条',c.tojo],['その他',c.other],['ライン',c.line],['13トン',c.t13]];
  var parts = [];
  for(var k=0;k<items.length;k++){ if(items[k][1]>0) parts.push(items[k][0]+' '+items[k][1]+'本'); }
  return { items: parts.join('　'), total: grp.length };
}
/** DEST_ORDERでの並び順。無い出荷先は末尾に回す。 */
function destRank_(dest){
  var i = DEST_ORDER.indexOf(dest);
  return (i>=0) ? i : DEST_ORDER.length;
}
function dateKeyFromName_(name){
  var m=name.match(/\((\d{2})年(\d{1,2})月(\d{1,2})日\)/);
  if(!m) return -1;
  return (2000+parseInt(m[1],10))*10000+parseInt(m[2],10)*100+parseInt(m[3],10);
}
/** 日程表変更ファイル名の「(25年9月16日)」部分から「9/16」を取り出す。無ければnull。 */
function srcDateTag_(name){
  var m=name.match(/\((\d{2})年(\d{1,2})月(\d{1,2})日\)/);
  if(!m) return null;
  return parseInt(m[2],10)+'/'+parseInt(m[3],10);
}
function midnight_(d){ return new Date(d.getFullYear(),d.getMonth(),d.getDate()); }
function fmtJ_(d){ return (d.getMonth()+1)+'月'+d.getDate()+'日'; }
function mdOrStr_(v){ if(v instanceof Date) return (v.getMonth()+1)+'/'+v.getDate(); return v==null?'':String(v); }
function strv_(v){ return v==null?'':String(v); }
function numi_(v){
  if(v==null||v==='') return '';
  if(typeof v==='number') return (v===Math.floor(v))?String(Math.floor(v)):String(v);
  return String(v);
}
function esc_(t){
  return String(t==null?'':t)
    .replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/>/g,'&gt;').replace(/"/g,'&quot;');
}
function safeRemove_(id){ try{ Drive.Files.remove(id); }catch(e){ try{ DriveApp.getFileById(id).setTrashed(true);}catch(e2){} } }

// ==== 前回との差分（変更行の赤字表示） ====

/** 図番＋号機で同一出荷物とみなすためのキー。 */
function snapshotKey_(zu, go){ return zu+'||'+go; }

/**
 * 前回生成時の内容と比較し、変更があった行・新規行に x.chg=true を立てる（赤字用、1回だけ光る）。
 * 出荷日については別に、最初に見つけた時点の日付(origShip)とも比較し、ズレていれば
 * x.prevShip に元の日付(Date)を入れる（表示用）。こちらはorigShipに一致するまでは
 * 赤字かどうかに関わらずずっと表示され続ける（黒字の行にも旧日付の注記が出る）。
 * prevMapがnull（＝比較対象となる前回スナップショットがまだ存在しない）の場合は、
 * 比較のしようがないので今回分は基準点として扱い、どの行も赤字にしない。
 *
 * 検査完了日(insp)は比較対象に含めない。検査が進むたびに空欄→日付へと
 * ほぼ毎日大量の行で自然に埋まっていく項目なので、これを含めると「本当に
 * 注意すべき変更」が埋もれてページの大半が赤字になってしまうため。
 */
function markChanges_(rows, prevMap){
  var noBaseline = (prevMap===null);
  for(var i=0;i<rows.length;i++){
    var x = rows[i];
    if(noBaseline){ x.chg=false; x.prevShip=null; continue; }
    var prev = prevMap[snapshotKey_(x.zu, x.go)];
    if(!prev){ x.chg=true; x.prevShip=null; continue; }
    var shipMs = x.ship.getTime();
    x.chg = (prev.kiki!==x.kiki || prev.kishu!==x.kishu ||
             prev.spec!==x.spec || prev.ship!==shipMs ||
             prev.info!==x.info || prev.dest!==x.dest || prev.is13!==x.is13);
    x.prevShip = (prev.origShip!==shipMs) ? new Date(prev.origShip) : null;
  }
}

/**
 * 今回生成分をスナップショットとして保存し、次回の比較に使う。
 * origShip（最初に見つけた時点の出荷日）は前回のスナップショットにあればそのまま
 * 引き継ぎ、今回上書きしない。新規の出荷物だけ今回の出荷日を基準にする。
 */
function saveSnapshot_(rows, prevMap){
  try{
    var map={};
    for(var i=0;i<rows.length;i++){
      var x=rows[i];
      var key = snapshotKey_(x.zu,x.go);
      var prev = prevMap && prevMap[key];
      var origShip = (prev && prev.origShip!=null) ? prev.origShip : x.ship.getTime();
      map[key] = { origShip:origShip, kiki:x.kiki, kishu:x.kishu,
        spec:x.spec, ship:x.ship.getTime(), info:x.info, dest:x.dest, is13:x.is13 };
    }
    var folder = DriveApp.getFolderById(OUT_FOLDER_ID);
    var ex = folder.getFilesByName(SNAPSHOT_FILE_NAME);
    while(ex.hasNext()) ex.next().setTrashed(true);
    folder.createFile(SNAPSHOT_FILE_NAME, JSON.stringify(map), MimeType.PLAIN_TEXT);
  }catch(e){
    Logger.log('スナップショットの保存に失敗: '+e.message);
  }
}

/**
 * 前回生成時のスナップショットを読み込む。
 * ファイルが無い（初回実行など）・読込に失敗した場合はnullを返す
 * ＝比較対象が無いので今回は赤字を出さず基準点にする、とmarkChanges_側で扱う。
 */
function loadSnapshot_(){
  try{
    var folder = DriveApp.getFolderById(OUT_FOLDER_ID);
    var it = folder.getFilesByName(SNAPSHOT_FILE_NAME);
    if(!it.hasNext()) return null;
    return JSON.parse(it.next().getBlob().getDataAsString('UTF-8'));
  }catch(e){
    Logger.log('スナップショットの読込に失敗: '+e.message);
    return null;
  }
}

/**
 * クラッシュ(GASのINTERNALエンジンエラー等)でfinallyが実行されず
 * tmp_arm_srcが残ってしまった場合の掃除。実行中の他インスタンスを誤って
 * 消さないよう、作成から10分以上経過したものだけ対象にする。
 */
function cleanupStaleTempFiles_(){
  try{
    var it = DriveApp.searchFiles("title = 'tmp_arm_src' and trashed = false");
    var cutoff = Date.now() - 10*60*1000;
    while(it.hasNext()){
      var f = it.next();
      if(f.getDateCreated().getTime() < cutoff){
        try{ Drive.Files.remove(f.getId()); }catch(e){ try{ f.setTrashed(true); }catch(e2){} }
      }
    }
  }catch(e){
    Logger.log('一時ファイルの掃除に失敗: '+e.message);
  }
}

function installPdfTrigger(){
  ScriptApp.getProjectTriggers().forEach(function(t){
    var fn = t.getHandlerFunction();
    if(fn==='generateArmPDF' || fn==='checkAndGenerateArmPDF') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('checkAndGenerateArmPDF').timeBased().everyMinutes(POLL_INTERVAL_MINUTES).create();
  Logger.log(POLL_INTERVAL_MINUTES+'分おきに更新チェックする自動実行トリガーを設置しました。'+
             '対象ファイルが実際に更新された時だけPDFを再生成します。');
}