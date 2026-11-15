/**
 * Vibeコーディング 仕様づくりナビ - スプレッドシート記録スクリプト
 * 
 * 【設定手順】
 * 1. 記録先とするGoogleスプレッドシートを開きます。
 * 2. 画面上のメニュー「拡張機能」>「Apps Script」をクリックします。
 * 3. 表示されたエディタの内容をすべて消し、このコードをそのまま貼り付けます。
 * 4. 保存アイコン（または Ctrl+S / Cmd+S）を押して保存します。
 * 5. 画面右上の「デプロイ」青色ボタン >「新しいデプロイ」をクリックします。
 * 6. 歯車アイコン（種類の選択）をクリックし、「ウェブアプリ」を選択します。
 * 7. 次のように設定します：
 *    - 説明: 仕様ナビ記録
 *    - 次のユーザーとして実行: 自分（あなたのGoogleアカウント）
 *    - アクセスできるユーザー: 全員 (※重要：生徒がGoogleログイン不要で送信できるようにするため)
 * 8. 「デプロイ」ボタンをクリックします。（初回はアクセスの承認画面が出るので許可してください）
 * 9. 発行された「ウェブアプリのURL」（https://script.google.com/macros/s/.../exec）をコピーします。
 * 10. index.html 内の GAS_WEBAPP_URL = "ここにURL"; に貼り付ければ連携完了です！
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  // 30秒間ロックを試みる（多人数が一斉に送信した場合のデータ競合対策）
  if (!lock.tryLock(30000)) {
    return ContentService.createTextOutput(JSON.stringify({ status: "busy" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    var raw = e.postData ? e.postData.contents : "";
    var data = {};

    if (e.parameter && e.parameter.data) {
      data = JSON.parse(e.parameter.data);
    } else if (raw) {
      data = JSON.parse(raw);
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getActiveSheet();

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
    return ContentService.createTextOutput(JSON.stringify({ status: "error", message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
