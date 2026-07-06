/**
 * ========================================
 * DAILY CORE - Sheet Sync
 * 作成日: 2026-07-06
 * 作成者: teramachi@terarbeit.com
 * 概要: DAILY CORE(PWA)から日次習慣記録をPOST受信し、スプレッドシートに1日1行でupsert蓄積する。
 *       将来ヘルスケア指標とのbefore/after分析の土台。トークンはScriptPropertiesに保存(コード直書き禁止=鉄則1)。
 * ----------------------------------------
 * 更新履歴:
 * 2026-07-06 teramachi@terarbeit.com 初版作成
 * ========================================
 */

const SHEET_NAME = 'DAILY CORE Log';
const HEADERS = ['date', 'weekday', 'done', 'total', 'coreDone', 'streak', 'maxStreak', 'updatedAt'];
const TOKEN_PROP = 'SYNC_TOKEN'; // ScriptPropertiesのキー。値はメニュー「同期トークンを設定」で保存する

/**
 * PWAからのPOSTを受信して1行upsertする。
 * @param {Object} e - postData.contentsにJSON文字列({token,date,weekday,done,total,coreDone,streak,maxStreak})
 * @return {TextOutput} JSON結果
 */
function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) {
      return json_({ ok: false, error: 'no body' });
    }
    const body = JSON.parse(e.postData.contents);
    const expected = PropertiesService.getScriptProperties().getProperty(TOKEN_PROP);
    if (!expected || body.token !== expected) {
      return json_({ ok: false, error: 'unauthorized' });
    }
    if (!body.date) {
      return json_({ ok: false, error: 'missing date' });
    }
    upsertRow_(body);
    return json_({ ok: true, date: body.date });
  } catch (err) {
    Logger.log('doPost error: ' + err.message);
    return json_({ ok: false, error: String(err.message) });
  }
}

/**
 * dateをキーに既存行を探して更新、無ければ追記する。
 * @param {Object} rec - 記録オブジェクト
 */
function upsertRow_(rec) {
  const sheet = getSheet_();
  const now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy-MM-dd HH:mm:ss');
  const doneStr = Array.isArray(rec.done) ? rec.done.join(',') : (rec.done || '');
  const row = [
    String(rec.date), rec.weekday, doneStr, rec.total,
    rec.coreDone, rec.streak, rec.maxStreak, now
  ];
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    const dates = sheet.getRange(2, 1, lastRow - 1, 1).getValues(); // 一括取得(ループ内getValue禁止=鉄則)
    for (let i = 0; i < dates.length; i++) {
      if (normDate_(dates[i][0]) === normDate_(rec.date)) {
        sheet.getRange(i + 2, 1, 1, row.length).setValues([row]);
        return;
      }
    }
  }
  sheet.appendRow(row);
}

/**
 * 日付セルがDate型に自動変換されても文字列比較できるよう正規化する。
 * @param {*} v - セル値(文字列 or Date)
 * @return {string} 'yyyy-MM-dd'
 */
function normDate_(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, 'Asia/Tokyo', 'yyyy-MM-dd');
  }
  return String(v).trim();
}

/**
 * ログ用シートを取得。無ければ作成しヘッダ・書式を設定する。
 * @return {Sheet}
 */
function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
    sheet.setFrozenRows(1);
    sheet.getRange('A:A').setNumberFormat('@'); // date列を文字列固定(Date自動変換で重複行を防ぐ)
  }
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===== 非エンジニア運用UI(鉄則2,3) ===== */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DAILY CORE')
    .addItem('同期トークンを設定', 'setTokenDialog_')
    .addItem('シートを初期化', 'initSheet_')
    .addToUi();
}

/**
 * DAILY CORE側と同じトークンをScriptPropertiesに保存する。
 */
function setTokenDialog_() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('同期トークン設定', 'DAILY CORE側(index.html)と同じトークンを貼り付けてください:', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() === ui.Button.OK) {
    const t = res.getResponseText().trim();
    if (t) {
      PropertiesService.getScriptProperties().setProperty(TOKEN_PROP, t);
      ui.alert('トークンを保存しました。');
    } else {
      ui.alert('空欄のため保存しませんでした。');
    }
  }
}

function initSheet_() {
  getSheet_();
  SpreadsheetApp.getUi().alert('シート「' + SHEET_NAME + '」を用意しました。');
}
