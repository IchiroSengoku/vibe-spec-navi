/**
 * Vibeコーディング 仕様づくりナビ - スプレッドシート記録スクリプト
 */

/**
 * 接続確認用（ブラウザのアドレスバーにWebアプリURLを直接貼り付けて開いたときに動作します）
 */
function doGet(e) {
  return ContentService.createTextOutput("OK: Vibe仕様づくりナビ連携APIは正常に稼働しています。")
    .setMimeType(ContentService.MimeType.TEXT);
}

/**
 * 動作テスト用
 * Apps Script エディタの上部で「testSave」を選択し、「実行」をクリックするとテスト行が追加されます。
 * ※初回のみ「承認が必要です」という画面が出るので、画面の案内に従ってアクセスを許可してください。
 */
function testSave() {
  var dummyEvent = {
    postData: {
      contents: JSON.stringify({
        timestamp: Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"),
        sessionId: "test-001",
        submissionCount: 1,
        appName: "テスト用アプリ（動作確認）",
        targetUser: "テスト利用者",
        problem: "テストの困りごと",
        goal: "テストの解決目標",
        prompt: "これは動作テスト用のプロンプトです。"
      })
    }
  };
  var result = doPost(dummyEvent);
  Logger.log("テスト実行結果: " + result.getContent());
}

/**
 * 本番データ受信処理
 */
function doPost(e) {
  var lock = LockService.getScriptLock();
  // 30秒間ロックを試みる（多人数が一斉に送信した場合のデータ競合対策）
  if (!lock.tryLock(30000)) {
    return ContentService.createTextOutput(JSON.stringify({ status: "busy" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var raw = e && e.postData ? e.postData.contents : "";
    var data = {};

    if (raw) {
      try {
        data = JSON.parse(raw);
      } catch (jsonErr) {
        // URLエンコードされている場合のフォールバック (data=%7B...%7D)
        if (raw.indexOf("data=") === 0) {
          data = JSON.parse(decodeURIComponent(raw.substring(5).replace(/\+/g, " ")));
        } else if (e.parameter && e.parameter.data) {
          data = JSON.parse(e.parameter.data);
        }
      }
    } else if (e && e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss) {
      throw new Error("スプレッドシートが見つかりません。必ずスプレッドシートのメニュー「拡張機能」>「Apps Script」から開いてください。");
    }
    // 常に1番目のシート（左端のタブ）に書き込みます
    var sheet = ss.getSheets()[0];

    // シートが空（新規作成時）の場合は1行目にヘッダーを自動生成
    if (sheet.getLastRow() === 0) {
      var headers = [
        "送信日時",
        "匿名ID",
        "送信回数",
        "アプリ名",
        "誰のため",
        "困っていること",
        "どうなってほしいか",
        "使う場面",
        "使用機器",
        "想定時間",
        "場面の理由",
        "機能1",
        "機能1優先度",
        "機能1仕様",
        "機能2",
        "機能2優先度",
        "機能2仕様",
        "機能3",
        "機能3優先度",
        "機能3仕様",
        "機能で大切にしたこと",
        "保存の必要",
        "保存するデータ",
        "保存方法",
        "保存データ表示",
        "データ削除",
        "データ出力",
        "使いやすさ",
        "その他の条件",
        "完成条件",
        "問題の改善予想",
        "一番大切にしたこと",
        "AIプロンプト本文"
      ];
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length)
        .setFontWeight("bold")
        .setBackground("#eaf2ff")
        .setHorizontalAlignment("center");
      sheet.setFrozenRows(1);
    }

    // データの追記
    var row = [
      data.timestamp || Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy/MM/dd HH:mm:ss"),
      data.sessionId || "",
      data.submissionCount ? (data.submissionCount + "回目") : "",
      data.appName || "",
      data.targetUser || "",
      data.problem || "",
      data.goal || "",
      data.scene || "",
      data.device || "",
      data.useTime || "",
      data.sceneReason || "",
      data.feature1Name || "",
      data.feature1Priority || "",
      data.feature1Detail || "",
      data.feature2Name || "",
      data.feature2Priority || "",
      data.feature2Detail || "",
      data.feature3Name || "",
      data.feature3Priority || "",
      data.feature3Detail || "",
      data.featureReason || "",
      data.saveNeed || "",
      data.savedData || "",
      data.saveMethod || "",
      data.dataDisplay || "",
      data.deleteData || "",
      data.exportData || "",
      data.usability || "",
      data.otherConditions || "",
      data.completion || "",
      data.evaluation || "",
      data.mostImportant || "",
      data.prompt || ""
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({ status: "success" }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    console.error("エラー詳細: " + err.toString());
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
