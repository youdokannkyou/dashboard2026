/**
 * 用度環境課共有シート ― プロジェクト保存用 Google Apps Script
 *
 * 【設置方法】
 * 1. 対象のスプレッドシートを開き、「拡張機能」→「Apps Script」を選択
 * 2. デフォルトで開かれる .gs ファイルの中身をすべて削除し、このコードを貼り付け
 * 3. 保存(Ctrl+S / Cmd+S)
 * 4. 画面右上の「デプロイ」→「デプロイを管理」を開く
 *    - 既存のWebアプリのデプロイがあれば、鉛筆アイコンから「編集」
 *    - 「バージョン」を「新しいバージョン」にして「デプロイ」をクリック
 *    ※ URLは変わらないので、サイト側の設定(GAS_URL)はそのままでOK
 * 5. 初回のみ、権限の承認(自分のGoogleアカウントでアクセスを許可)が求められます
 *
 * 【前提】
 * このスプレッドシートに「プロジェクト」という名前のシートがあり、
 * 1行目が A:プロジェクトID / B:プロジェクト名 / C:作成日 / D:状態 になっていること。
 */

const SHEET_NAME = 'プロジェクト';

/**
 * GET リクエスト: プロジェクト一覧を返す
 * 例: {GAS_URL}?action=list
 */
function doGet(e) {
  try {
    const sheet = getProjectSheet_();
    const lastRow = sheet.getLastRow();
    let projects = [];

    if (lastRow >= 2) {
      const values = sheet.getRange(2, 1, lastRow - 1, 4).getValues(); // A〜D列、2行目以降
      projects = values
        .filter(row => row[0] !== '' && row[0] !== null) // プロジェクトIDが空の行は無視
        .map(row => ({
          id: String(row[0]),
          name: row[1],
          createdAt: formatDate_(row[2]),
          status: row[3]
        }));
    }

    return jsonOutput_({ success: true, projects: projects });
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err) });
  }
}

/**
 * POST リクエスト: プロジェクトを1件追加する
 * body(JSON文字列): { action: 'createProject', id, name, createdAt, status }
 */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'createProject') {
      if (!body.name) {
        return jsonOutput_({ success: false, error: 'プロジェクト名が空です' });
      }
      const sheet = getProjectSheet_();
      const id = body.id || Utilities.getUuid();
      const createdAt = body.createdAt || formatDate_(new Date());
      const status = body.status || '進行中';

      sheet.appendRow([id, body.name, createdAt, status]);

      return jsonOutput_({ success: true, id: id });
    }

    return jsonOutput_({ success: false, error: '不明なactionです: ' + body.action });
  } catch (err) {
    return jsonOutput_({ success: false, error: String(err) });
  }
}

/* ---------- 内部ヘルパー ---------- */

function getProjectSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    throw new Error('「' + SHEET_NAME + '」シートが見つかりません');
  }
  return sheet;
}

function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return value ? String(value) : '';
}

function jsonOutput_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
