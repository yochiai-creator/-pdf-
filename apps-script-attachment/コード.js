/************************************************************************
 * アーム工場メール 添付ファイル自動保存スクリプト
 * ---------------------------------------------------------------------
 * Gmailのラベル「アーム工場」に付いたメールの添付ファイルのうち、
 * ファイル名に「日程表変更」を含むExcel（出荷予定 日程表変更A(◯月◯日).xlsm）を
 * Googleドライブのフォルダ「アーム出荷明細」に自動で保存します。
 *
 * ・同じ名前のファイルが既にフォルダにあれば保存しません（重複防止）
 * ・処理したメールには「ドライブ保存済み」ラベルを付けて、次回以降スキップします
 *   （ただし保存に失敗した添付が1つでもあるスレッドにはラベルを付けません）
 *
 * 作成: 野田組データ管理AI プロジェクト
 ************************************************************************/

// ============================= 設定 ============================
// ★ ここだけ確認すればOK。基本は変更不要です。

// 対象のGmailラベル名
const LABEL_NAME = 'アーム工場';

// 保存先のGoogleドライブ フォルダID（フォルダ「アーム出荷明細」）
const FOLDER_ID = '1NS4WoClO0xlGWSxvFimFcqFGUT0jOKQL';

// 添付ファイル名にこの文字が含まれるものだけを保存する
// （「日程表変更」を含むファイル = 出荷予定 日程表変更A(◯月◯日).xlsm）
const NAME_MUST_CONTAIN = '日程表変更';

// 処理済みメールに付けるラベル（重複処理を防ぐため）
const PROCESSED_LABEL = 'ドライブ保存済み';

// 1回の実行で処理する最大スレッド数
const MAX_THREADS = 100;
// ====================================================


/**
 * メイン処理（日程表変更の保存）。
 * saveArmOtherAttachments のトリガーからも一緒に呼ばれます。
 * 手動で動作確認するときも、この関数を実行してください。
 */
function saveArmAttachmentsToDrive() {
  const folder = DriveApp.getFolderById(FOLDER_ID);

  // 処理済みラベルを準備（無ければ作成）
  let processedLabel = GmailApp.getUserLabelByName(PROCESSED_LABEL);
  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(PROCESSED_LABEL);
  }

  // 「アーム工場」ラベルが付いていて、まだ「ドライブ保存済み」でなく、添付があるスレッドを検索
  const query = 'label:"' + LABEL_NAME + '" -label:"' + PROCESSED_LABEL + '" has:attachment';
  const threads = GmailApp.search(query, 0, MAX_THREADS);

  let savedCount = 0;
  let skippedExisting = 0;
  let failedCount = 0;
  let scannedThreads = 0;

  threads.forEach(function (thread) {
    scannedThreads++;
    let threadHadFailure = false;
    const messages = thread.getMessages();

    messages.forEach(function (message) {
      const attachments = message.getAttachments();

      attachments.forEach(function (att) {
        const name = att.getName();

        // ファイル名に「日程表変更」を含むものだけ対象
        if (name.indexOf(NAME_MUST_CONTAIN) === -1) {
          Logger.log('対象外（スキップ）: ' + name);
          return;
        }

        // Drive側の一時的な不調で1件失敗しても、ここで実行全体を止めない。
        // 止めてしまうと、後ろに並んでいるスレッドが毎回処理されないまま
        // 同じ場所で詰まり続けることになる。
        try {
          // 同名ファイルが既にフォルダにあれば保存しない
          if (existsInFolder_(folder, name)) {
            skippedExisting++;
            Logger.log('既に存在（スキップ）: ' + name);
            return;
          }

          // フォルダに保存
          folder.createFile(att.copyBlob()).setName(name);
          savedCount++;
          Logger.log('★ 保存しました: ' + name);
        } catch (e) {
          failedCount++;
          threadHadFailure = true;
          Logger.log('!! 保存失敗（次回再試行します）: ' + name + ' / ' + e.message);
        }
      });
    });

    // 失敗した添付が1つでもあるスレッドには処理済みラベルを付けない。
    // 付けてしまうと次回の検索対象から外れ、そのファイルが永久に保存されなくなる。
    if (threadHadFailure) {
      Logger.log('  → 未保存が残るためラベル付与を見送り: ' + thread.getFirstMessageSubject());
      return;
    }
    thread.addLabel(processedLabel);
  });

  Logger.log(
    '=== 完了 === 対象スレッド:' + scannedThreads +
    ' / 新規保存:' + savedCount +
    ' / 既存のためスキップ:' + skippedExisting +
    ' / 保存失敗:' + failedCount
  );
}


/**
 * 【初回に1回だけ実行】自動実行トリガーを作成します。
 * 既に同じトリガーがある場合は作り直します。
 * ※ saveArmOtherAttachments のトリガーが日程表変更の保存も呼ぶので、通常は不要です。
 */
function createTimeTrigger() {
  // 同じ関数の既存トリガーを削除
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'saveArmAttachmentsToDrive') {
      ScriptApp.deleteTrigger(t);
    }
  });

  ScriptApp.newTrigger('saveArmAttachmentsToDrive')
    .timeBased()
    .everyHours(4)
    .create();

  Logger.log('4時間ごとの自動実行トリガーを作成しました。');
}


/**
 * 【任意】自動実行を止めたいときに実行します。
 */
function deleteTimeTrigger() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'saveArmAttachmentsToDrive') {
      ScriptApp.deleteTrigger(t);
    }
  });
  Logger.log('自動実行トリガーを削除しました。');
}


function saveArmOtherAttachments() {
  // 日程表変更ファイルの保存もこのトリガーで一緒に行う
  try { saveArmAttachmentsToDrive(); } catch (e) { Logger.log('日程表変更の保存でエラー: ' + e.message); }

  // 添付ファイル名にキーワードを含むものを、指定フォルダへ自動保存
  var A = '14m8LUAQh7Bnn6HHxb2H9k6euN18MGXjg'; // ④部品関係
  var B = '1htLk5GWp1Uw0oRPP5W70Y-RSmfifiC2g'; // ③コベルコ大日程
  var OILBATH = '1rJhRvTggAmMmv9HgiJzQmuQma4f-4O80'; // ④部品関係/オイルバス
  var BOSS = '1bopDyK1X4LmlRMy0TBduOADDFJtoe6RZ';    // ④部品関係/ボス加工

  // 上から順に照合し、最初に一致したルールの保存先へ入れる。
  // 「ボス内製」「廻り止め」は「内製手配」にも一致してしまうため、必ず先に書く。
  var RULES = [
    { key: 'ボス内製',     folderId: BOSS },
    { key: '廻り止め',     folderId: BOSS },
    { key: '回り止め',     folderId: BOSS },
    { key: '単品出荷',     folderId: BOSS },
    { key: 'オイルバス',   folderId: OILBATH },
    { key: '内製手配',     folderId: A },
    { key: '大日程',       folderId: B }
  ];

  var cache = {};
  function getFolder(id) { if (!cache[id]) cache[id] = DriveApp.getFolderById(id); return cache[id]; }

  var threads = GmailApp.search('label:"アーム工場" has:attachment', 0, 100);
  var saved = 0;
  var skipped = 0;
  var failed = 0;
  for (var i = 0; i < threads.length; i++) {
    var msgs = threads[i].getMessages();
    for (var m = 0; m < msgs.length; m++) {
      var atts = msgs[m].getAttachments();
      for (var a = 0; a < atts.length; a++) {
        var name = atts[a].getName();
        for (var r = 0; r < RULES.length; r++) {
          if (name.indexOf(RULES[r].key) === -1) continue;
          // 1件の失敗で実行全体を止めない（止めると後ろのファイルが毎回処理されない）
          try {
            var folder = getFolder(RULES[r].folderId);
            if (existsInFolder_(folder, name)) { skipped++; break; }
            folder.createFile(atts[a].copyBlob()).setName(name);
            saved++;
            Logger.log('saved: ' + name + ' -> ' + folder.getName());
          } catch (e) {
            failed++;
            Logger.log('!! 保存失敗（次回再試行します）: ' + name + ' / ' + e.message);
          }
          break;
        }
      }
    }
  }
  Logger.log('done saved=' + saved + ' skipped=' + skipped + ' failed=' + failed);
}

function createOtherTrigger() {
  var ts = ScriptApp.getProjectTriggers();
  for (var i = 0; i < ts.length; i++) {
    if (ts[i].getHandlerFunction() === 'saveArmOtherAttachments') ScriptApp.deleteTrigger(ts[i]);
  }
  ScriptApp.newTrigger('saveArmOtherAttachments').timeBased().everyHours(4).create();
  Logger.log('trigger created');
}

/**
 * フォルダ内に同名ファイルが「生きた状態で」存在するか。
 * getFilesByName はゴミ箱に入れたファイルも返すため、そのまま重複判定に使うと
 * 「手動で消したのに、もうあると判定されて二度と保存されない」状態になる。
 */
function existsInFolder_(folder, name) {
  var it = folder.getFilesByName(name);
  while (it.hasNext()) {
    if (!it.next().isTrashed()) return true;
  }
  return false;
}

/** アーム工場: 保存先フォルダIDの疎通確認 */
function checkArmFolderIds() {
  var ids = {
    'メイン(アーム出荷明細)': FOLDER_ID,
    'A(④部品関係)': '14m8LUAQh7Bnn6HHxb2H9k6euN18MGXjg',
    'B(③コベルコ大日程)': '1htLk5GWp1Uw0oRPP5W70Y-RSmfifiC2g',
    'オイルバス(④部品関係/オイルバス)': '1rJhRvTggAmMmv9HgiJzQmuQma4f-4O80',
    'ボス加工(④部品関係/ボス加工)': '1bopDyK1X4LmlRMy0TBduOADDFJtoe6RZ'
  };
  for (var k in ids) {
    try {
      var f = DriveApp.getFolderById(ids[k]);
      Logger.log(k + ': OK 「' + f.getName() + '」');
    } catch (e) {
      Logger.log(k + ': NG (' + ids[k] + ') - ' + e.message);
    }
  }
}
