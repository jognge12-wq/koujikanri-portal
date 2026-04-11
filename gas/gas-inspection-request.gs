/**
 * ============================================================
 * 建設現場 性能評価検査 依頼自動化システム
 * Google Apps Script (GAS) - バックエンド
 * ============================================================
 *
 * 【セットアップ手順】
 * 1. Google スプレッドシートを新規作成
 * 2. ツール > スクリプトエディタ でこのコードを貼り付け
 * 3. 下記の定数を自分の環境に合わせて変更
 * 4. デプロイ > 新しいデプロイ > ウェブアプリ
 *    - 実行ユーザー: 自分
 *    - アクセス: 全員（匿名ユーザーを含む）※URLが外部に漏れないよう管理
 * 5. デプロイURLをHTMLファイルの GAS_URL に貼り付け
 */

// ============================================================
// ★ 設定値（必ず自分の環境に合わせて変更してください）
// ============================================================

const CONFIG = {
  // スプレッドシートID（URLの /d/〇〇〇/edit の〇〇〇部分）
  SPREADSHEET_ID: '1--Uq_DhPNaNKHJbEWQaaKe_saM0yjSs4fp_XdpNtNu8',

  // シート名
  SHEET_INSPECTIONS: '検査依頼',    // 依頼記録シート
  SHEET_PROPERTIES:  '物件マスタ',  // 物件一覧シート

  // カレンダー設定
  CALENDAR_COLOR_ID: 1,   // 1 = ラベンダー色
  INSPECTION_DURATION_MINUTES: 60,  // 検査の所要時間（分）

  // メール設定
  DEFAULT_CC: '',  // 常にCC送信したいアドレス（不要なら空文字）
  DEFAULT_SENDER_NAME: '',  // 送信者名（空の場合はGoogleアカウント名）
};

// ============================================================
// メインエントリポイント
// ============================================================

/**
 * POSTリクエストを受け取るエントリポイント
 */
function doPost(e) {
  // CORS対応ヘッダー
  const headers = {
    'Access-Control-Allow-Origin': '*',
    'Access-Control-Allow-Methods': 'POST, GET, OPTIONS',
    'Access-Control-Allow-Headers': 'Content-Type',
  };

  try {
    const data = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    switch (action) {
      case 'getProjects':
  　　　　return handleGetProjects();
　　　　case 'saveProject':
  　　　　return handleSaveProject(data);
　　　　case 'updateProject':
  　　　　return handleUpdateProject(data);
　　　　case 'deleteProject':
  　　　　return handleDeleteProject(data);
      case 'submitInspection':
        result = submitInspection(data);
        break;
      case 'updateStatus':
        result = updateStatus(data);
        break;
      case 'getInspections':
        result = getInspections(data);
        break;
      default:
        result = { success: false, error: `不明なアクション: ${action}` };
    }

    return buildResponse(result);

  } catch (err) {
    console.error('doPost エラー:', err.toString());
    return buildResponse({ success: false, error: err.toString() });
  }
}

/**
 * GETリクエストを受け取るエントリポイント（物件マスタ取得など）
 */
function doGet(e) {
  try {
    const action = e.parameter.action;
    let result;

    switch (action) {
      case 'getProperties':
        result = getProperties();
        break;
      case 'getInspections':
        result = getInspections(e.parameter);
      break;
      default:
        result = { success: false, error: `不明なアクション: ${action}` };
    }

    return buildResponse(result);

  } catch (err) {
    console.error('doGet エラー:', err.toString());
    return buildResponse({ success: false, error: err.toString() });
  }
}

/** レスポンス生成ヘルパー */
function buildResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// 1. 検査依頼メイン処理
// ============================================================

/**
 * 検査依頼の総合処理
 * - スプレッドシート保存
 * - メール送信（skipEmail: true の場合はスキップ）
 * - カレンダー登録（isUpdate: true の場合は既存削除→再作成）
 */
function submitInspection(data) {
  const results = {};

  // バリデーション
  const validation = validateInspectionData(data);
  if (!validation.valid) {
    return { success: false, error: validation.error };
  }

  // 1. スプレッドシートに保存
  const saveResult = saveToSpreadsheet(data);
  results.saved = saveResult;

  // 保存失敗時はここで止める
  if (!saveResult.success) {
    return {
      success: false,
      error: `スプレッドシート保存エラー: ${saveResult.error}`,
      results: results
    };
  }

  // 2. メール送信（skipEmail の場合はスキップ）
  if (data.skipEmail) {
    results.email = { success: true, skipped: true };
  } else {
    try {
      results.email = sendInspectionEmail(data);
    } catch (err) {
      results.email = { success: false, error: err.toString() };
      console.error('メール送信エラー:', err);
    }
  }

  // 3. カレンダー登録（エラーでも続行）
  try {
    // 変更時は既存イベントを削除してから再作成
    if (data.isUpdate && data.previousDate) {
      deleteExistingCalendarEvent(data);
    }
    results.calendar = addToCalendar(data);
  } catch (err) {
    results.calendar = { success: false, error: err.toString() };
    console.error('カレンダー登録エラー:', err);
  }

  // 全体の成功判定（スプレッドシート保存成功 = 全体成功とする）
  const overallSuccess = saveResult.success;

  return {
    success: overallSuccess,
    rowId: saveResult.rowId,
    results: results,
    message: buildResultMessage(results)
  };
}

/** バリデーション */
function validateInspectionData(data) {
  if (!data.propertyName)  return { valid: false, error: '物件名が未入力です' };
  if (!data.process)       return { valid: false, error: '工程が未選択です' };
  if (!data.inspectionDate) return { valid: false, error: '検査日が未入力です' };
  if (!data.inspectionTime) return { valid: false, error: '検査時間が未入力です' };
  // skipEmail の場合はメールアドレス不要
  if (!data.skipEmail && !data.recipientEmail) {
    return { valid: false, error: '送信先メールアドレスが未設定です' };
  }
  return { valid: true };
}

/** 結果サマリーメッセージ生成 */
function buildResultMessage(results) {
  const parts = [];
  if (results.saved?.success)     parts.push('記録✅');
  if (results.email?.skipped)     parts.push('メール送信スキップ');
  else if (results.email?.success)    parts.push('メール送信✅');
  else if (results.email && !results.email.success) parts.push('メール送信⚠️');
  if (results.calendar?.success)  parts.push('カレンダー登録✅');
  else if (results.calendar && !results.calendar.success) parts.push('カレンダー登録⚠️');
  return parts.join(' / ');
}

// ============================================================
// 2. スプレッドシート操作
// ============================================================

/** 検査依頼をスプレッドシートに保存 */
function saveToSpreadsheet(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);

    // シートが存在しない場合は作成
    if (!sheet) {
      sheet = ss.insertSheet(CONFIG.SHEET_INSPECTIONS);
      initializeInspectionSheet(sheet);
    }

    const rowId = Utilities.getUuid();
    const now = new Date();

    sheet.appendRow([
      Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss'),  // 登録日時
      data.propertyName,      // 物件名
      data.propertyAddress || '',  // 物件住所
      data.process,           // 工程
      data.inspectionDate,    // 検査日
      data.inspectionTime,    // 検査時間
      data.isUpdate ? '変更' : '予定',  // ステータス
      data.notes || '',       // 備考
      data.recipientEmail || '',  // 送信先
      rowId                   // 行ID（更新時のキー）
    ]);

    return { success: true, rowId: rowId };

  } catch (err) {
    console.error('saveToSpreadsheet エラー:', err);
    return { success: false, error: err.toString() };
  }
}

/** シートの初期ヘッダー設定 */
function initializeInspectionSheet(sheet) {
  const headers = [
    '登録日時', '物件名', '物件住所', '工程',
    '検査日', '検査時間', 'ステータス', '備考',
    '送信先メール', '行ID'
  ];
  sheet.appendRow(headers);

  // ヘッダー行の書式設定
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#4285F4');
  headerRange.setFontColor('#FFFFFF');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
}

/** ステータス更新 */
function updateStatus(data) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);
    if (!sheet) return { success: false, error: 'シートが見つかりません' };

    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const rowIdColIdx = headers.indexOf('行ID');
    const statusColIdx = headers.indexOf('ステータス');

    if (rowIdColIdx === -1 || statusColIdx === -1) {
      return { success: false, error: '列が見つかりません' };
    }

    for (let i = 1; i < allData.length; i++) {
      if (allData[i][rowIdColIdx] === data.rowId) {
        sheet.getRange(i + 1, statusColIdx + 1).setValue(data.status);
        return { success: true };
      }
    }

    return { success: false, error: '対象レコードが見つかりません' };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** 検査依頼一覧を取得 */
function getInspections(params) {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_INSPECTIONS);
    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, inspections: [] };
    }

    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const inspections = allData.slice(1).map(row => {
      const obj = {};
      headers.forEach((h, i) => { obj[h] = row[i]; });
      return obj;
    });

    return { success: true, inspections: inspections };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

/** 物件マスタを取得 */
function getProperties() {
  try {
    const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_PROPERTIES);

    if (!sheet || sheet.getLastRow() <= 1) {
      return { success: true, properties: [] };
    }

    const allData = sheet.getDataRange().getValues();
    const headers = allData[0];
    const properties = allData.slice(1)
      .filter(row => row[0])  // 空行除外
      .map(row => {
        const obj = {};
        headers.forEach((h, i) => { obj[h] = row[i]; });
        return obj;
      });

    return { success: true, properties: properties };

  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// ============================================================
// 3. メール送信
// ============================================================

/**
 * 検査依頼メールを送信
 */
function sendInspectionEmail(data) {
  try {
    const subject = buildEmailSubject(data);
    const body    = buildEmailBody(data);

    const options = {};

    // CC設定
    const ccList = [CONFIG.DEFAULT_CC, data.ccEmail]
      .filter(addr => addr && addr.trim())
      .join(',');
    if (ccList) options.cc = ccList;

    // 送信者名設定
    const senderName = data.senderName || CONFIG.DEFAULT_SENDER_NAME;
    if (senderName) options.name = senderName;

    // BCCで自分にも送る場合（任意）
    if (data.bccEmail) options.bcc = data.bccEmail;

    GmailApp.sendEmail(data.recipientEmail, subject, body, options);

    console.log(`メール送信完了: ${data.recipientEmail} / 件名: ${subject}`);
    return { success: true, to: data.recipientEmail };

  } catch (err) {
    console.error('sendInspectionEmail エラー:', err);
    return { success: false, error: err.toString() };
  }
}

/** メール件名を生成（フロントから送られた件名を優先使用） */
function buildEmailSubject(data) {
  if (data.emailSubject) return data.emailSubject;
  return `【検査依頼】${data.propertyName}　${data.process}`;
}

/** メール本文を生成 */
function buildEmailBody(data) {
  const dateStr = formatDateJP(data.inspectionDate);
  const dayStr  = getDayOfWeek(data.inspectionDate);
  const timeStr = data.inspectionTime;

  // 登録番号を notes から抽出
  const regMatch = data.notes ? data.notes.match(/登録番号: (.+?)(\n|$)/) : null;
  const reg = regMatch ? regMatch[1] : '';

  // 変更通知かどうか
  const isChange = data.isUpdate;

  // 変更通知用の旧→新日程テキスト
  let changeLine = '';
  if (isChange && data.previousDate) {
    const prevDateStr = formatDateJP(data.previousDate);
    const prevDayStr  = getDayOfWeek(data.previousDate);
    changeLine = `\n【変更内容】\n　変更前：${prevDateStr}（${prevDayStr}）\n　変更後：${dateStr}（${dayStr}） ${timeStr}〜\n`;
  }

  // 署名ブロック
  let signature = '';
  if (data.senderName || data.senderCompany || data.senderPhone) {
    signature = '\n──────────────────\n';
    if (data.senderName)    signature += `${data.senderName}\n`;
    if (data.senderCompany) signature += `${data.senderCompany}\n`;
    if (data.senderPhone)   signature += `TEL: ${data.senderPhone}\n`;
    signature += '──────────────────';
  }

  if (isChange) {
    // ===== 日程変更通知メール =====
    return `
【転送先】nagoya-yoyaku@whec.co.jp

株式会社西日本住宅評価センター
名古屋支店　検査部　御中

お世話になっております。

先日ご依頼いたしました検査日程の変更をお願いしたく、
ご連絡申し上げます。

━━━━━━━━━━━━━━━━━━━━
■ 検査日程変更のお願い
━━━━━━━━━━━━━━━━━━━━
物件名　　　：${data.propertyName}
登録番号　　：${reg}
検査種別　　：${data.process}
${changeLine}
検査希望日　：${dateStr}（${dayStr}）
開始希望時間：${timeStr}〜
━━━━━━━━━━━━━━━━━━━━

お手数をおかけいたしますが、
ご確認・ご調整のほど、よろしくお願いいたします。
${signature}`.trim();

  } else {
    // ===== 新規検査依頼メール =====
    return `
【転送先】nagoya-yoyaku@whec.co.jp

株式会社西日本住宅評価センター
名古屋支店　検査部　御中

お世話になっております。

下記の通り、検査のご依頼を申し上げます。
ご確認・ご調整のほど、よろしくお願いいたします。

━━━━━━━━━━━━━━━━━━━━
■ 検査依頼内容
━━━━━━━━━━━━━━━━━━━━
物件名　　　：${data.propertyName}
登録番号　　：${reg}
検査種別　　：${data.process}
検査希望日　：${dateStr}（${dayStr}）
開始希望時間：${timeStr}〜
━━━━━━━━━━━━━━━━━━━━

何かご不明な点がございましたら、
お気軽にご連絡くださいますようお願いいたします。
${signature}`.trim();
  }
}

/** 日付を日本語形式に変換（例: 2025-06-15 → 2025年6月15日）*/
function formatDateJP(dateStr) {
  try {
    const d = new Date(dateStr + 'T00:00:00');
    return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy年M月d日');
  } catch (e) {
    return dateStr;
  }
}

/** 曜日を返す */
function getDayOfWeek(dateStr) {
  const days = ['日', '月', '火', '水', '木', '金', '土'];
  try {
    const d = new Date(dateStr + 'T00:00:00');
    return days[d.getDay()];
  } catch (e) {
    return '';
  }
}

// ============================================================
// 4. Googleカレンダー登録
// ============================================================

/**
 * 既存のカレンダーイベントを検索して削除（日程変更時に使用）
 */
function deleteExistingCalendarEvent(data) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();
    const searchDate = new Date(data.previousDate + 'T00:00:00');
    const nextDay    = new Date(searchDate.getTime() + 24 * 60 * 60 * 1000);
    const title      = `【${data.process}】${data.propertyName}`;

    const events = calendar.getEvents(searchDate, nextDay);
    for (const ev of events) {
      if (ev.getTitle() === title) {
        ev.deleteEvent();
        console.log(`既存イベント削除: ${title} / ${data.previousDate}`);
        break;
      }
    }
  } catch (err) {
    console.warn('既存イベント削除エラー（続行）:', err);
  }
}

/**
 * Googleカレンダーに検査予定を登録
 */
function addToCalendar(data) {
  try {
    const calendar = CalendarApp.getDefaultCalendar();

    // 日時解析
    const startTime = new Date(`${data.inspectionDate}T${data.inspectionTime}:00`);
    if (isNaN(startTime.getTime())) {
      return { success: false, error: '日時の解析に失敗しました' };
    }

    const endTime = new Date(startTime.getTime() + CONFIG.INSPECTION_DURATION_MINUTES * 60 * 1000);

    const title       = `【${data.process}】${data.propertyName}`;
    const description = buildCalendarDescription(data);
    // location フィールドを優先、なければ propertyAddress
    const location    = data.location || data.propertyAddress || '';

    const event = calendar.createEvent(title, startTime, endTime, {
      description: description,
      location:    location,
      colorId:     CONFIG.CALENDAR_COLOR_ID,
    });

    console.log(`カレンダー登録完了: ${title} / eventId: ${event.getId()}`);
    return { success: true, eventId: event.getId() };

  } catch (err) {
    console.error('addToCalendar エラー:', err);
    return { success: false, error: err.toString() };
  }
}

/** カレンダー説明文を生成 */
function buildCalendarDescription(data) {
  const requestDate = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd');
  const parts = [`依頼日：${requestDate}`];
  if (data.notes) parts.push(data.notes);
  if (data.isUpdate) parts.push('※ 日程変更');
  return parts.join('\n');
}

// ============================================================
// 5. プロジェクト管理（Notion連携 + スプレッドシート）
// ============================================================

// プロジェクト一覧を取得
function fetchNotionProjects() {
  var props = PropertiesService.getScriptProperties();
  var token = props.getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN がスクリプトプロパティに設定されていません');

  var DB_ID = '2f56ad84-6221-80a9-891b-ef7e5514fa78';
  var url = 'https://api.notion.com/v1/databases/' + DB_ID + '/query';

  var payload = {
    filter: {
      property: '本体着工',
      date: { on_or_after: '2026-01-08' }
    },
    sorts: [{ property: '本体着工', direction: 'ascending' }],
    page_size: 100
  };

  var options = {
    method: 'post',
    contentType: 'application/json',
    headers: {
      'Authorization': 'Bearer ' + token,
      'Notion-Version': '2022-06-28'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  var response = UrlFetchApp.fetch(url, options);
  var json = JSON.parse(response.getContentText());

  if (!json.results) {
    throw new Error('Notion API エラー: ' + response.getContentText());
  }

  return json.results.map(function(page) {
    var titleProp = page.properties['物件名'];
    var name = (titleProp && titleProp.title && titleProp.title[0])
      ? titleProp.title[0].plain_text
      : '（名称未設定）';
    return { id: page.id, name: name };
  });
}

function getSheetData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_PROPERTIES) ? CONFIG.SHEET_PROPERTIES : '物件マスタ';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(['id', 'reg', 'history', 'delegate', 'times']);
    return {};
  }
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var timesColIdx = headers.indexOf('times');
  var map = {};
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    try {
      var entry = {
        reg:      row[1].toString(),
        history:  JSON.parse(row[2] || '{}'),
        delegate: JSON.parse(row[3] || '{}')
      };
      // times列がある場合は読み取り
      if (timesColIdx >= 0 && row[timesColIdx]) {
        entry.times = JSON.parse(row[timesColIdx]);
      } else {
        entry.times = {};
      }
      map[row[0].toString()] = entry;
    } catch (e) {}
  }
  return map;
}

function handleGetProjects() {
  try {
    var notionProjects = fetchNotionProjects();
    var sheetData = getSheetData();
    var projects = notionProjects.map(function(p) {
      var s = sheetData[p.id] || {};
      return {
        id:       p.id,
        name:     p.name,
        reg:      s.reg      || '',
        history:  s.history  || {},
        delegate: s.delegate || {},
        times:    s.times    || {}
      };
    });
    return jsonResponse({ success: true, projects: projects });
  } catch (e) {
    return jsonResponse({ success: false, error: e.toString() });
  }
}

function handleSaveProject(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_PROPERTIES) ? CONFIG.SHEET_PROPERTIES : '物件マスタ';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(['id', 'reg', 'history', 'delegate', 'times']);
    }
    // times列がなければ追加
    ensureTimesColumn_(sheet);
    var rows = sheet.getDataRange().getValues();
    var timesCol = rows[0].indexOf('times') + 1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) {
        sheet.getRange(i + 1, 2).setValue(data.reg || '');
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(data.history  || {}));
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(data.delegate || {}));
        if (timesCol > 0) sheet.getRange(i + 1, timesCol).setValue(JSON.stringify(data.times || {}));
        return jsonResponse({ success: true });
      }
    }
    var newRow = [
      data.id,
      data.reg || '',
      JSON.stringify(data.history  || {}),
      JSON.stringify(data.delegate || {})
    ];
    if (timesCol > 0) {
      // timesCol が5列目なら5個の要素が必要
      while (newRow.length < timesCol - 1) newRow.push('');
      newRow.push(JSON.stringify(data.times || {}));
    }
    sheet.appendRow(newRow);
    return jsonResponse({ success: true });
  } catch (e) {
    return jsonResponse({ success: false, error: e.toString() });
  }
}

function handleUpdateProject(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_PROPERTIES) ? CONFIG.SHEET_PROPERTIES : '物件マスタ';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonResponse({ success: false, error: '物件マスタ シートが見つかりません' });
    // times列がなければ追加
    ensureTimesColumn_(sheet);
    var rows = sheet.getDataRange().getValues();
    var timesCol = rows[0].indexOf('times') + 1;
    for (var i = 1; i < rows.length; i++) {
      if (rows[i][0].toString() === data.id.toString()) {
        sheet.getRange(i + 1, 2).setValue(data.reg || '');
        sheet.getRange(i + 1, 3).setValue(JSON.stringify(data.history  || {}));
        sheet.getRange(i + 1, 4).setValue(JSON.stringify(data.delegate || {}));
        if (timesCol > 0) sheet.getRange(i + 1, timesCol).setValue(JSON.stringify(data.times || {}));
        return jsonResponse({ success: true });
      }
    }
    return handleSaveProject(data);
  } catch (e) {
    return jsonResponse({ success: false, error: e.toString() });
  }
}

/**
 * 物件マスタに times 列がなければ追加するヘルパー
 */
function ensureTimesColumn_(sheet) {
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headers.indexOf('times') === -1) {
    var nextCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, nextCol).setValue('times');
  }
}

function handleDeleteProject(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEET_PROPERTIES) ? CONFIG.SHEET_PROPERTIES : '物件マスタ';
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) return jsonResponse({ success: true });
    var rows = sheet.getDataRange().getValues();
    for (var i = rows.length - 1; i >= 1; i--) {
      if (rows[i][0].toString() === data.id.toString()) {
        sheet.deleteRow(i + 1);
        return jsonResponse({ success: true });
      }
    }
    return jsonResponse({ success: true });
  } catch (e) {
    return jsonResponse({ success: false, error: e.toString() });
  }
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
