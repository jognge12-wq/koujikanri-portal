// ============================================================
// 施工管理スケジュール自動登録 - Google Apps Script v1.5
// ============================================================
// 【v1.5 修正内容】
// - カレンダー色定数名を正しいGAS APIの定数名に再修正
//   TANGERINE → ORANGE（ミカン）、PEACOCK → CYAN（ピーコック）、GRAPE → MAUVE（グレープ）
//   ※GAS CalendarApp.EventColor の正式定数名はUI表示名と異なる
// 【v1.4 修正内容（一部誤り）】
// - カレンダー色定数名を変更（TANGERINE/PEACOCK/GRAPEはGASに存在しないため無効だった）
// - 施工計画説明の場所(location)を登録するよう修正（空白→locに変更）
// 【v1.3 修正内容】
// - ScriptApp.getService().getUrl() が古いURLを返す問題を修正
//   → WEBAPP_URL に現在のデプロイURLを直接設定する方式に変更
// 【v1.2 修正内容】
// - 「基礎検査・竣工検査」などの「〇日前」計算で祝日に当たった場合、
//   前倒し方向（より前の日付）にスキップするよう修正
// - 手動テスト関数 testManual() を追加
// ============================================================

// ▼ 設定エリア（ここだけ変更すればOK）
const CALENDAR_ID   = 'jognge12@gmail.com'; // ← 自分のGoogleカレンダーIDに変更
const CONFIRM_EMAIL = 'jognge12@gmail.com'; // ← 確認メール送信先（通常は同じアドレス）

// ★ デプロイURLを直接記入（デプロイを管理 → ウェブアプリのURLをコピーして貼り付け）
// 「デプロイを管理」→ アクティブなデプロイを選択 → ウェブアプリのURLをここに貼る
const WEBAPP_URL = 'https://script.google.com/macros/s/AKfycbwCVoCBuecRtBNdmY_HJ3-a8xjMxs2AjxGTgIKxMQo7AnuIqS8RNgWpLxVfSFZnVQAD5A/exec';

// カレンダーイベントの色設定
// ※GASの正式定数名はUI表示名（ミカン/ピーコック/グレープ）と異なる
const COLOR_CHAKOU  = CalendarApp.EventColor.ORANGE; // 本体着工 → ミカン(Tangerine=colorId:6)
const COLOR_TATEMAE = CalendarApp.EventColor.CYAN;   // 建て方   → ピーコック(Peacock=colorId:7)
const COLOR_DEFAULT = CalendarApp.EventColor.MAUVE;  // それ以外 → グレープ(Grape=colorId:3)

// ============================================================
// 祝日リスト 2025〜2027年
// ============================================================
const HOLIDAYS = [
  '2025-01-01','2025-01-13','2025-02-11','2025-02-23','2025-02-24',
  '2025-03-20','2025-04-29','2025-05-03','2025-05-04','2025-05-05',
  '2025-05-06','2025-07-21','2025-08-11','2025-09-15','2025-09-22','2025-09-23',
  '2025-10-13','2025-11-03','2025-11-23','2025-11-24',
  '2026-01-01','2026-01-12','2026-02-11','2026-02-23','2026-03-20',
  '2026-04-29','2026-05-03','2026-05-04','2026-05-05','2026-05-06',
  '2026-07-20','2026-08-11','2026-09-21','2026-09-22','2026-09-23',
  '2026-10-12','2026-11-03','2026-11-23',
  '2027-01-01','2027-01-11','2027-02-11','2027-02-23','2027-03-21',
  '2027-04-29','2027-05-03','2027-05-04','2027-05-05',
  '2027-07-19','2027-08-11','2027-09-20','2027-09-23',
  '2027-10-11','2027-11-03','2027-11-23',
];

// ============================================================
// 三隣亡リスト 2025〜2027年
// ============================================================
const BAD_DAYS_SANRINBO = [
  '2025-01-03','2025-01-15','2025-01-27','2025-02-08','2025-02-20',
  '2025-03-04','2025-03-16','2025-03-28','2025-04-09','2025-04-21',
  '2025-05-03','2025-05-15','2025-05-27','2025-06-08','2025-06-20',
  '2025-07-02','2025-07-14','2025-07-26','2025-08-07','2025-08-19',
  '2025-08-31','2025-09-12','2025-09-24','2025-10-06','2025-10-18',
  '2025-10-30','2025-11-11','2025-11-23','2025-12-05','2025-12-17','2025-12-29',
  '2026-01-10','2026-01-22','2026-02-03','2026-02-15','2026-02-27',
  '2026-03-11','2026-03-23','2026-04-04','2026-04-16','2026-04-28',
  '2026-05-10','2026-05-22','2026-06-03','2026-06-15','2026-06-27',
  '2026-07-09','2026-07-21','2026-08-02','2026-08-14','2026-08-26',
  '2026-09-07','2026-09-19','2026-10-01','2026-10-13','2026-10-25',
  '2026-11-06','2026-11-18','2026-11-30','2026-12-12','2026-12-24',
  '2027-01-05','2027-01-17','2027-01-29','2027-02-10','2027-02-22',
  '2027-03-06','2027-03-18','2027-03-30','2027-04-11','2027-04-23',
  '2027-05-05','2027-05-17','2027-05-29','2027-06-10','2027-06-22',
  '2027-07-04','2027-07-16','2027-07-28','2027-08-09','2027-08-21',
  '2027-09-02','2027-09-14','2027-09-26','2027-10-08','2027-10-20',
  '2027-11-01','2027-11-13','2027-11-25','2027-12-07','2027-12-19','2027-12-31',
];

// ============================================================
// メイン処理①: フォーム送信時 → 確認メールを送信
// ============================================================
function onFormSubmit(e) {
  try {
    const itemResponses = e.response.getItemResponses();
    const data = {};
    itemResponses.forEach(function(item) {
      data[item.getItem().getTitle()] = item.getResponse();
    });

    const bukkenName     = (data['物件名'] || '').trim();
    const location       = (data['場所（市町村）'] || '').trim();
    const chakouDateStr  = (data['本体着工日'] || '').trim();
    const tatemaeStr     = (data['建て方'] || '').trim();
    const hikiwatashiStr = (data['引渡し日'] || '').trim();

    if (!bukkenName || !chakouDateStr || !tatemaeStr || !hikiwatashiStr) {
      throw new Error('必須項目が入力されていません');
    }

    const chakouDate      = parseDate(chakouDateStr);
    const tatemaeDate     = parseDate(tatemaeStr);
    const hikiwatashiDate = parseDate(hikiwatashiStr);

    const schedules = calcSchedules(bukkenName, location, chakouDate, tatemaeDate, hikiwatashiDate);
    const token = saveToSheet(bukkenName, location, schedules);
    sendConfirmEmail(bukkenName, schedules, token);

    Logger.log('確認メール送信完了: ' + bukkenName);

  } catch (err) {
    Logger.log('エラー: ' + err.toString());
    MailApp.sendEmail(
      CONFIRM_EMAIL,
      '【エラー】スケジュール登録に失敗しました',
      'エラー内容:\n' + err.toString() + '\n\nフォームを再送信してください。'
    );
  }
}

// ============================================================
// ★ 手動テスト用関数（フォームなしでテスト可能）
// Apps Script エディタ上でこの関数を選択して実行してください
// ============================================================
function testManual() {
  // ↓ テストしたい値に書き換えてください
  const bukkenName     = 'テスト様邸';
  const location       = '浜松市';
  const chakouDateStr  = '2026-05-11'; // 本体着工日
  const tatemaeStr     = '2026-06-08'; // 建て方
  const hikiwatashiStr = '2026-09-25'; // 引渡し日

  const chakou      = parseDate(chakouDateStr);
  const tatemae     = parseDate(tatemaeStr);
  const hikiwatashi = parseDate(hikiwatashiStr);

  const schedules = calcSchedules(bukkenName, location, chakou, tatemae, hikiwatashi);

  Logger.log('===== スケジュール計算結果 =====');
  schedules.forEach(function(s) {
    const start   = new Date(s.start);
    const end     = new Date(s.end);
    const timeStr = s.allDay ? '終日' : pad(start.getHours()) + ':' + pad(start.getMinutes()) + '〜' + pad(end.getHours()) + ':' + pad(end.getMinutes());
    Logger.log('【工程' + String(s.step).padStart(2, '0') + '】' + s.title + ' / ' + formatDate(start) + ' ' + timeStr);
  });
  Logger.log('================================');
  Logger.log('合計: ' + schedules.length + '件');
}

// ============================================================
// メイン処理②: メールのリンクをクリック → カレンダーに登録
// ============================================================
function doGet(e) {
  const token  = e.parameter.token;
  const action = e.parameter.action;

  // ポータルフォームからの送信
  if (e.parameter.mode === 'submitSchedule') {
    return handlePortalSubmit(e.parameter);
  }

  if (action === 'cancel') {
    deleteFromSheet(token);
    return HtmlService.createHtmlOutput('<h2>登録をキャンセルしました。</h2>');
  }

  if (action === 'register') {
    const result = registerFromSheet(token);
    if (result.success) {
      return HtmlService.createHtmlOutput(
        '<h2>✅ カレンダーへの登録が完了しました！</h2>' +
        '<p><b>' + result.bukkenName + '</b> の ' + result.count + '件のスケジュールを登録しました。</p>' +
        '<p><a href="https://calendar.google.com">Googleカレンダーを開く</a></p>'
      );
    } else {
      return HtmlService.createHtmlOutput('<h2>❌ エラーが発生しました</h2><p>' + result.error + '</p>');
    }
  }

  // パラメータなし = ポータルからのiframe表示 → 入力フォームを返す
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle('工程作成')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
// フォーム画面HTML
// ============================================================
function getFormHtml() {
  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>工程作成</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Kaku Gothic ProN', sans-serif;
    font-size: 16px;
    background: #f0f4f2;
    min-height: 100vh;
    padding: 0 0 32px;
  }
  header { display: none; }
  .card {
    background: #fff;
    border-radius: 14px;
    margin: 12px 16px;
    padding: 16px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  .card h2 {
    font-size: 13px;
    font-weight: 600;
    color: #1D6B40;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    margin-bottom: 16px;
    padding-bottom: 10px;
    border-bottom: 1px solid #e8f0eb;
  }
  .field { margin-bottom: 18px; }
  .field:last-child { margin-bottom: 0; }
  label {
    display: block;
    font-size: 14px;
    font-weight: 600;
    color: #333;
    margin-bottom: 8px;
  }
  label .required { color: #e05555; font-size: 12px; margin-left: 4px; }
  input[type="text"], select {
    display: block; width: 100%;
    padding: 10px 12px;
    font-size: 15px; font-family: inherit;
    border: 1.5px solid #dde8e2; border-radius: 12px;
    background: #fff; color: #222;
    -webkit-appearance: none; appearance: none;
    transition: border-color 0.2s;
  }
  input[type="text"]:focus, select:focus { outline: none; border-color: #2D8A4E; box-shadow: 0 0 0 3px rgba(45,138,78,0.15); }
  .date-row { position: relative; }
  .date-display {
    display: flex; align-items: center; justify-content: space-between;
    width: 100%; padding: 8px 10px; font-size: 14px;
    border: 1.5px solid #dde8e2; border-radius: 12px;
    background: #fff; cursor: pointer; min-height: 42px;
    transition: border-color 0.2s;
  }
  .date-display.filled { border-color: #2D8A4E; }
  .date-placeholder { color: #aab8c2; font-size: 14px; }
  .date-hidden {
    position: absolute; top: 0; left: 0;
    width: 100%; height: 100%; opacity: 0; cursor: pointer; font-size: 16px;
    -webkit-appearance: none; appearance: none;
  }
  .cal-icon { flex-shrink: 0; color: #aab8c2; }
  select {
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23999' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
    background-repeat: no-repeat; background-position: right 16px center; padding-right: 40px;
  }
  .btn-submit {
    display: block; width: calc(100% - 32px); margin: 16px;
    padding: 14px; background: #1D6B40; color: #fff;
    font-size: 15px; font-weight: 700; font-family: inherit;
    border: none; border-radius: 14px; cursor: pointer;
    box-shadow: 0 4px 14px rgba(29,107,64,0.3);
    transition: opacity 0.15s;
  }
  .btn-submit:active { opacity: 0.8; }
  .btn-submit:disabled { opacity: 0.6; cursor: not-allowed; }
  .result-card {
    background: #fff; border-radius: 14px; margin: 12px 16px; padding: 20px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08); text-align: center;
  }
  .result-card.success { border-left: 4px solid #1D6B40; }
  .result-card.error   { border-left: 4px solid #e05555; }
  .result-card h3 { font-size: 15px; font-weight: 700; margin-bottom: 8px; }
  .result-card p  { font-size: 13px; color: #555; line-height: 1.6; }
</style>
</head>
<body>
<form id="scheduleForm" onsubmit="handleSubmit(event)">
  <div class="card">
    <h2>基本情報</h2>
    <div class="field">
      <label>物件名 <span class="required">*</span></label>
      <input type="text" id="bukkenName" placeholder="〇〇様邸" required>
    </div>
    <div class="field">
      <label>場所（市町村）</label>
      <select id="location">
        <option value="">（選択してください）</option>
        <option>岐阜市</option><option>大垣市</option><option>高山市</option>
        <option>多治見市</option><option>関市</option><option>中津川市</option>
        <option>美濃市</option><option>瑞浪市</option><option>羽島市</option>
        <option>恵那市</option><option>美濃加茂市</option><option>土岐市</option>
        <option>各務原市</option><option>可児市</option><option>山県市</option>
        <option>瑞穂市</option><option>飛騨市</option><option>本巣市</option>
        <option>郡上市</option>
      </select>
    </div>
  </div>
  <div class="card">
    <h2>\u5de5\u7a0b</h2>
    <div class="field">
      <label>\u672c\u4f53\u7740\u5de5\u65e5 <span class="required">*</span></label>
      <div class="date-row">
        <div class="date-display" id="dd_chakou">
          <span id="dw_chakou" class="date-placeholder">\u65e5\u4ed8\u3092\u9078\u629e</span>
          <svg class="cal-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
        </div>
        <input type="date" id="chakouDate" class="date-hidden" oninput="updateDay('chakouDate','dw_chakou','dd_chakou')" required>
      </div>
    </div>
    <div class="field">
      <label>\u5efa\u3066\u65b9 <span class="required">*</span></label>
      <div class="date-row">
        <div class="date-display" id="dd_tatemae">
          <span id="dw_tatemae" class="date-placeholder">\u65e5\u4ed8\u3092\u9078\u629e</span>
          <svg class="cal-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
        </div>
        <input type="date" id="tatemaeDate" class="date-hidden" oninput="updateDay('tatemaeDate','dw_tatemae','dd_tatemae')" required>
      </div>
    </div>
    <div class="field">
      <label>\u5f15\u6e21\u3057\u65e5 <span class="required">*</span></label>
      <div class="date-row">
        <div class="date-display" id="dd_hikiwatashi">
          <span id="dw_hikiwatashi" class="date-placeholder">\u65e5\u4ed8\u3092\u9078\u629e</span>
          <svg class="cal-icon" xmlns="http://www.w3.org/2000/svg" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round"><rect x="3" y="4" width="18" height="18" rx="2" ry="2"/><line x1="16" y1="2" x2="16" y2="6"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="3" y1="10" x2="21" y2="10"/></svg>
        </div>
        <input type="date" id="hikiwatashiDate" class="date-hidden" oninput="updateDay('hikiwatashiDate','dw_hikiwatashi','dd_hikiwatashi')" required>
      </div>
    </div>
  </div>
  <button type="submit" class="btn-submit" id="submitBtn">\u5de5\u7a0b\u3092Google\u30ab\u30ec\u30f3\u30c0\u30fc\u306b\u767b\u9332\u3059\u308b</button>
</form>
<div id="resultArea"></div>

<script>
const WEBAPP_URL = '${WEBAPP_URL}';

var DAY_NAMES = ['\u65e5','\u6708','\u706b','\u6c34','\u6728','\u91d1','\u571f'];
function updateDay(inputId, spanId, boxId) {
  var val  = document.getElementById(inputId).value;
  var span = document.getElementById(spanId);
  var box  = document.getElementById(boxId);
  if (!val) {
    span.className = 'date-placeholder';
    span.innerHTML = '\u65e5\u4ed8\u3092\u9078\u629e';
    box.classList.remove('filled');
    return;
  }
  var d   = new Date(val + 'T00:00:00');
  var dow = d.getDay();
  var dayColor = dow === 0 ? '#ef4444' : dow === 6 ? '#3b82f6' : '#1e293b';
  span.className = '';
  span.innerHTML =
    '<span style="color:#1e293b;font-weight:600">' + (d.getMonth()+1) + '\u6708' + d.getDate() + '\u65e5</span>' +
    '<span style="color:' + dayColor + ';font-weight:700;margin-left:2px">(' + DAY_NAMES[dow] + ')</span>';
  box.classList.add('filled');
}

function handleSubmit(e) {
  e.preventDefault();
  const btn = document.getElementById('submitBtn');
  btn.disabled = true;
  btn.textContent = '送信中…';

  const params = new URLSearchParams({
    mode:      'submitSchedule',
    bukkenName:     document.getElementById('bukkenName').value.trim(),
    location:       document.getElementById('location').value,
    chakouDate:     document.getElementById('chakouDate').value,
    tatemaeDate:    document.getElementById('tatemaeDate').value,
    hikiwatashiDate: document.getElementById('hikiwatashiDate').value,
  });

  fetch(WEBAPP_URL + '?' + params.toString())
    .then(function(r) { return r.json(); })
    .then(function(json) {
      var area = document.getElementById('resultArea');
      if (json.success) {
        document.getElementById('scheduleForm').style.display = 'none';
        area.innerHTML =
          '<div class="result-card success">' +
          '<h3>✅ 送信完了</h3>' +
          '<p>' + json.bukkenName + ' の工程スケジュールを計算しました。<br>' +
          '確認メールを送信しましたので、メール内のリンクをタップしてGoogleカレンダーへ登録してください。</p>' +
          '</div>';
      } else {
        area.innerHTML = '<div class="result-card error"><h3>❌ エラー</h3><p>' + (json.error || '送信に失敗しました') + '</p></div>';
        btn.disabled = false;
        btn.textContent = '\u5de5\u7a0b\u3092Google\u30ab\u30ec\u30f3\u30c0\u30fc\u306b\u767b\u9332\u3059\u308b';
      }
    })
    .catch(function(err) {
      document.getElementById('resultArea').innerHTML =
        '<div class="result-card error"><h3>❌ 通信エラー</h3><p>' + err.toString() + '</p></div>';
      btn.disabled = false;
      btn.textContent = '工程をカレンダーに登録する';
    });
}
</script>
</body>
</html>`;
}

// ============================================================
// ポータルフォームからの送信処理
// ============================================================
function handlePortalSubmit(params) {
  try {
    var bukkenName      = (params.bukkenName || '').trim();
    var location        = (params.location || '').trim();
    var chakouDateStr   = (params.chakouDate || '').trim();
    var tatemaeDateStr  = (params.tatemaeDate || '').trim();
    var hikiwatashiStr  = (params.hikiwatashiDate || '').trim();

    if (!bukkenName || !chakouDateStr || !tatemaeDateStr || !hikiwatashiStr) {
      return ContentService.createTextOutput(JSON.stringify({ success: false, error: '必須項目が未入力です' }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var chakou      = parseDate(chakouDateStr);
    var tatemae     = parseDate(tatemaeDateStr);
    var hikiwatashi = parseDate(hikiwatashiStr);

    var schedules = calcSchedules(bukkenName, location, chakou, tatemae, hikiwatashi);
    var token     = saveToSheet(bukkenName, location, schedules);
    sendConfirmEmail(bukkenName, schedules, token);

    return ContentService.createTextOutput(JSON.stringify({ success: true, bukkenName: bukkenName }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// スプレッドシートへの一時保存
// ============================================================
function saveToSheet(bukkenName, location, schedules) {
  const ss    = getOrCreateSheet();
  const sheet = ss.getSheetByName('pending') || ss.insertSheet('pending');
  const token = Utilities.getUuid();
  const serialized = schedules.map(function(s) {
    return Object.assign({}, s, { start: s.start.getTime(), end: s.end.getTime() });
  });
  const json = JSON.stringify({ bukkenName: bukkenName, location: location, schedules: serialized });
  sheet.appendRow([token, new Date(), json]);
  return token;
}

function deleteFromSheet(token) {
  const ss    = getOrCreateSheet();
  const sheet = ss.getSheetByName('pending');
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  for (var i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === token) { sheet.deleteRow(i + 1); break; }
  }
}

function registerFromSheet(token) {
  try {
    const ss    = getOrCreateSheet();
    const sheet = ss.getSheetByName('pending');
    if (!sheet) throw new Error('保存データが見つかりません');

    const data = sheet.getDataRange().getValues();
    var found  = null;
    var rowIdx = -1;
    for (var i = 0; i < data.length; i++) {
      if (data[i][0] === token) { found = JSON.parse(data[i][2]); rowIdx = i + 1; break; }
    }
    if (!found) throw new Error('すでに登録済みか、トークンが見つかりません');

    const calendar  = CalendarApp.getCalendarById(CALENDAR_ID);
    if (!calendar) throw new Error('カレンダーが見つかりません。CALENDAR_IDを確認してください: ' + CALENDAR_ID);

    const schedules = found.schedules.map(function(s) {
      return Object.assign({}, s, { start: new Date(s.start), end: new Date(s.end) });
    });

    const results = [];
    for (var j = 0; j < schedules.length; j++) {
      results.push(registerWithConflictCheck(calendar, schedules[j]));
    }

    // ToDo: 引継ぎ(工程1)の前日に「〇〇様邸質疑まとめ」を終日タスクで作成
    var step1 = null;
    for (var k = 0; k < results.length; k++) {
      if (results[k].step === 1) { step1 = results[k]; break; }
    }
    if (step1) {
      const todoDate = new Date(step1.start);
      todoDate.setDate(todoDate.getDate() - 1);
      calendar.createAllDayEvent(found.bukkenName + '様邸質疑まとめ', todoDate, {
        description: '施工計画説明前の質疑事項まとめ（ToDo）'
      });
    }

    sheet.deleteRow(rowIdx);
    return { success: true, bukkenName: found.bukkenName, count: results.length };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

function getOrCreateSheet() {
  const files = DriveApp.getFilesByName('スケジュール登録_一時保存');
  if (files.hasNext()) return SpreadsheetApp.open(files.next());
  return SpreadsheetApp.create('スケジュール登録_一時保存');
}

// ============================================================
// 確認メール送信
// ============================================================
function sendConfirmEmail(bukkenName, schedules, token) {
  // WEBAPP_URL を優先使用。空の場合のみ ScriptApp から取得（フォールバック）
  const scriptUrl   = WEBAPP_URL || ScriptApp.getService().getUrl();
  const registerUrl = scriptUrl + '?action=register&token=' + token;
  const cancelUrl   = scriptUrl + '?action=cancel&token=' + token;

  var lines = schedules.map(function(s) {
    const start   = new Date(s.start);
    const end     = new Date(s.end);
    const timeStr = s.allDay ? '終日' : pad(start.getHours()) + ':' + pad(start.getMinutes()) + '〜' + pad(end.getHours()) + ':' + pad(end.getMinutes());
    const loc     = s.location ? '　📍' + s.location : '';
    return '【工程' + String(s.step).padStart(2, '0') + '】' + s.title + '\n　📅 ' + formatDate(start) + '　' + timeStr + loc;
  });

  const body =
    bukkenName + ' のスケジュール登録内容をご確認ください。\n\n' +
    '━━━━━━━━━━━━━━━━━━━━\n' +
    lines.join('\n\n') + '\n' +
    '━━━━━━━━━━━━━━━━━━━━\n\n' +
    '▼ 内容が正しければ「登録する」をクリックしてください\n' +
    registerUrl + '\n\n' +
    '▼ キャンセルする場合はこちら\n' +
    cancelUrl + '\n\n' +
    '※このメールは自動送信されています。\n' +
    '※登録後はリンクが無効になります。';

  MailApp.sendEmail({
    to: CONFIRM_EMAIL,
    subject: '【確認】' + bukkenName + ' スケジュール登録内容',
    body: body,
  });
}

function pad(n) { return String(n).padStart(2, '0'); }

// ============================================================
// 16工程スケジュール算出
// ============================================================
function calcSchedules(bukkenName, location, chakou, tatemae, hikiwatashi) {
  const n   = bukkenName;
  const loc = location;
  const schedules = [];

  // --- 工程2: 施工計画説明（着工前の直近土曜、不可なら日曜）---
  var step2Date = prevSaturdayOrSunday(addDays(chakou, -1));
  schedules.push({ step: 2, title: n + '施工計画説明', start: setTime(step2Date, 9, 0), end: setTime(step2Date, 10, 30), location: loc, description: '所要1.5h', notification: true, allDay: false, color: COLOR_DEFAULT });

  // --- 工程1: 引継ぎ（施工計画説明の前日、火・水・第1木曜は前倒し）---
  var step1Date = avoidTueWedFirstThu(addDays(step2Date, -1));
  schedules.push({ step: 1, title: n + '引継ぎ', start: setTime(step1Date, 17, 0), end: setTime(step1Date, 18, 0), location: '', description: '所要1.0h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程3: 本体着工（基準日①）---
  schedules.push({ step: 3, title: '🚜' + n + '本体着工', start: chakou, end: chakou, location: loc, description: '', notification: false, allDay: true, color: COLOR_CHAKOU });

  // --- 工程4: 遣り方検査（着工翌日、水・日・祝除く）---
  var step4Date = skipWedSunHoliday(addDays(chakou, 1));
  schedules.push({ step: 4, title: n + '遣り方検査', start: setTime(step4Date, 8, 30), end: setTime(step4Date, 10, 0), location: loc, description: '所要1.5h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程5: 配筋検査（着工8日後、水・日・祝除く）---
  var step5Date = skipWedSunHoliday(addDays(chakou, 8));
  schedules.push({ step: 5, title: n + '配筋検査', start: setTime(step5Date, 8, 30), end: setTime(step5Date, 10, 30), location: loc, description: '所要2.0h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程6: 型枠検査（配筋検査8日後、水・日・祝除く）---
  var step6Date = skipWedSunHoliday(addDays(step5Date, 8));
  schedules.push({ step: 6, title: n + '型枠検査', start: setTime(step6Date, 8, 30), end: setTime(step6Date, 10, 0), location: loc, description: '所要1.5h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程7: 基礎検査（建て方6日前、水・日・祝除く）---
  // ★ 修正: 「〇日前」は前倒し方向（より前の日付）にスキップ
  var step7Date = skipWedSunHolidayBackward(addDays(tatemae, -6));
  schedules.push({ step: 7, title: n + '基礎検査', start: setTime(step7Date, 8, 30), end: setTime(step7Date, 9, 30), location: loc, description: '所要1.0h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程8: 建て方（基準日②、三隣亡・仏滅・赤口は警告+吉日提案）---
  var rokuyoInfo   = checkRokuyo(tatemae);
  var step8Warning = '';
  if (rokuyoInfo.isBad || isSanrinbo(tatemae)) {
    var suggested = findGoodDay(tatemae);
    step8Warning  = '⚠️ 建て方日(' + formatDate(tatemae) + ')は' + (rokuyoInfo.name || '三隣亡') + 'です。近隣吉日(' + formatDate(suggested) + ')をご提案します。';
  }
  schedules.push({ step: 8, title: '⚒️' + n + '建て方', start: setTime(tatemae, 8, 0), end: setTime(tatemae, 9, 0), location: loc, description: step8Warning || '所要1.0h', notification: false, allDay: false, color: COLOR_TATEMAE });

  // --- 工程9: 構造検査（建て方5日後、水・日・祝除く）---
  var step9Date = skipWedSunHoliday(addDays(tatemae, 5));
  schedules.push({ step: 9, title: n + '構造検査', start: setTime(step9Date, 8, 30), end: setTime(step9Date, 10, 30), location: loc, description: '所要2.0h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程10: 構造立会い（建て方以降の直近土曜、不可なら日曜）---
  var step10Date = nextSaturdayOrSunday(addDays(tatemae, 1));
  schedules.push({ step: 10, title: n + '構造立会い', start: setTime(step10Date, 9, 0), end: setTime(step10Date, 10, 30), location: loc, description: '所要1.5h', notification: true, allDay: false, color: COLOR_DEFAULT });

  // --- 工程11: 雨仕舞い検査（建て方16日後、水・日・祝除く）---
  var step11Date = skipWedSunHoliday(addDays(tatemae, 16));
  schedules.push({ step: 11, title: n + '雨仕舞い検査', start: setTime(step11Date, 8, 30), end: setTime(step11Date, 10, 0), location: loc, description: '所要1.5h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程14: 竣工検査（引渡し14日前、水・日・祝除く）---
  // ★ 修正: 「〇日前」は前倒し方向（より前の日付）にスキップ
  var step14Date = skipWedSunHolidayBackward(addDays(hikiwatashi, -14));
  schedules.push({ step: 14, title: n + '竣工検査', start: setTime(step14Date, 8, 30), end: setTime(step14Date, 14, 0), location: loc, description: '所要5.5h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程12: 木完検査（竣工検査14日前・同曜日、水・日・祝除く）---
  // 14日＝2週間のため同曜日が保たれる
  var step12Date = skipWedSunHolidayBackward(addDays(step14Date, -14));
  schedules.push({ step: 12, title: n + '木完検査', start: setTime(step12Date, 8, 30), end: setTime(step12Date, 10, 0), location: loc, description: '所要1.5h（竣工検査14日前）', notification: false, allDay: false, color: COLOR_DEFAULT });

  // --- 工程13: 木完立会い（木完検査前後の直近土日）---
  var step13Date = nearestSaturdaySunday(step12Date);
  schedules.push({ step: 13, title: n + '木完立会い', start: setTime(step13Date, 9, 0), end: setTime(step13Date, 10, 30), location: loc, description: '所要1.5h', notification: true, allDay: false, color: COLOR_DEFAULT });

  // --- 工程15: 竣工立会い（引渡し7日前の直近土曜、不可なら日曜）---
  var step15Date = prevSaturdayOrSunday(addDays(hikiwatashi, -7));
  schedules.push({ step: 15, title: n + '竣工立会い', start: setTime(step15Date, 9, 0), end: setTime(step15Date, 11, 0), location: loc, description: '所要2.0h', notification: true, allDay: false, color: COLOR_DEFAULT });

  // --- 工程16: 引渡し（基準日③）---
  schedules.push({ step: 16, title: '🔑' + n + '引渡し', start: setTime(hikiwatashi, 10, 0), end: setTime(hikiwatashi, 12, 0), location: loc, description: '所要2.0h', notification: false, allDay: false, color: COLOR_DEFAULT });

  // 工程番号順に並び替えて返す
  schedules.sort(function(a, b) { return a.step - b.step; });
  return schedules;
}

// ============================================================
// カレンダー登録（重複チェック付き）
// ============================================================
function registerWithConflictCheck(calendar, s) {
  var startTime = new Date(s.start);
  var endTime   = new Date(s.end);
  var adjusted  = false;

  if (!s.allDay) {
    var attempts = 0;
    while (hasConflict(calendar, startTime, endTime) && attempts < 14) {
      startTime = addDays(startTime, 1);
      endTime   = addDays(endTime, 1);
      while (isWedSunHoliday(startTime)) {
        startTime = addDays(startTime, 1);
        endTime   = addDays(endTime, 1);
      }
      adjusted = true;
      attempts++;
    }
    var event = calendar.createEvent(s.title, startTime, endTime, {
      location: s.location || '',
      description: s.description || ''
    });
    if (s.color) event.setColor(s.color);
    if (s.notification) {
      event.removeAllReminders();
      event.addPopupReminder(calcNotifyMinutes(startTime));
    }
  } else {
    var allDayEvent = calendar.createAllDayEvent(s.title, startTime, {
      location: s.location || '',
      description: s.description || ''
    });
    if (s.color) allDayEvent.setColor(s.color);
  }

  return Object.assign({}, s, { start: startTime, end: endTime, adjusted: adjusted });
}

// ============================================================
// ユーティリティ関数
// ============================================================

// 日付文字列をDateに変換（YYYY/MM/DD, YYYY-MM-DD, MM/DD 形式に対応）
function parseDate(str) {
  str = str.trim();
  var m;
  if ((m = str.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/))) {
    return new Date(parseInt(m[1]), parseInt(m[2]) - 1, parseInt(m[3]));
  }
  if ((m = str.match(/^(\d{1,2})[\/](\d{1,2})$/))) {
    return new Date(new Date().getFullYear(), parseInt(m[1]) - 1, parseInt(m[2]));
  }
  throw new Error('日付の形式が不正です: ' + str);
}

function addDays(date, n) { var d = new Date(date); d.setDate(d.getDate() + n); return d; }
function setTime(date, h, m) { var d = new Date(date); d.setHours(h, m, 0, 0); return d; }
function formatDate(date) {
  var d = new Date(date);
  return d.getFullYear() + '/' + (d.getMonth() + 1) + '/' + d.getDate() + '(' + ['日','月','火','水','木','金','土'][d.getDay()] + ')';
}

// 水・日・祝日チェック
function isWedSunHoliday(date) {
  var d = new Date(date);
  var dow = d.getDay();
  if (dow === 0 || dow === 3) return true; // 日曜(0) または 水曜(3)
  var key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  return HOLIDAYS.indexOf(key) >= 0;
}

// 水・日・祝なら翌日方向にスキップ（着工後の検査などに使用）
function skipWedSunHoliday(date) {
  var d = new Date(date);
  while (isWedSunHoliday(d)) { d = addDays(d, 1); }
  return d;
}

// ★ v1.2 追加: 水・日・祝なら前日方向にスキップ（「〇日前」の検査に使用）
function skipWedSunHolidayBackward(date) {
  var d = new Date(date);
  while (isWedSunHoliday(d)) { d = addDays(d, -1); }
  return d;
}

// 指定日以前の直近土曜を返す（なければ直近日曜）
function prevSaturdayOrSunday(date) {
  var d = new Date(date);
  for (var i = 0; i < 14; i++) { if (d.getDay() === 6) return d; d = addDays(d, -1); }
  d = new Date(date);
  for (var j = 0; j < 14; j++) { if (d.getDay() === 0) return d; d = addDays(d, -1); }
  return date;
}

// 指定日以降の直近土曜を返す（なければ直近日曜）
function nextSaturdayOrSunday(date) {
  var d = new Date(date);
  for (var i = 0; i < 14; i++) { if (d.getDay() === 6) return d; d = addDays(d, 1); }
  d = new Date(date);
  for (var j = 0; j < 14; j++) { if (d.getDay() === 0) return d; d = addDays(d, 1); }
  return date;
}

// 指定日前後で最も近い土日を返す
function nearestSaturdaySunday(date) {
  var d = new Date(date);
  for (var i = 1; i <= 7; i++) {
    var fwd = addDays(d, i); if (fwd.getDay() === 6 || fwd.getDay() === 0) return fwd;
    var bwd = addDays(d, -i); if (bwd.getDay() === 6 || bwd.getDay() === 0) return bwd;
  }
  return date;
}

// 火・水・第1木曜を前倒し回避（引継ぎに使用）
function avoidTueWedFirstThu(date) {
  var d = new Date(date);
  for (var i = 0; i < 7; i++) {
    var dow = d.getDay();
    if (dow === 2 || dow === 3) { d = addDays(d, -1); continue; } // 火(2)・水(3)
    if (dow === 4 && d.getDate() <= 7) { d = addDays(d, -1); continue; } // 第1木曜(4)
    break;
  }
  return d;
}

// 三隣亡チェック
function isSanrinbo(date) {
  var d = new Date(date);
  var key = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  return BAD_DAYS_SANRINBO.indexOf(key) >= 0;
}

// 六曜チェック（仏滅・赤口）※将来実装用のスタブ
function checkRokuyo(date) {
  // TODO: 六曜計算を実装する場合はここに追加
  return { isBad: false, name: '' };
}

// 吉日（三隣亡以外・日曜以外）を最大14日の範囲で探す
function findGoodDay(date) {
  var d = new Date(date);
  for (var i = 1; i <= 14; i++) {
    var fwd = addDays(d, i); if (!isSanrinbo(fwd) && fwd.getDay() !== 0) return fwd;
    var bwd = addDays(d, -i); if (!isSanrinbo(bwd) && bwd.getDay() !== 0) return bwd;
  }
  return addDays(d, 1);
}

// 同時間帯に既存予定があるかチェック
function hasConflict(calendar, start, end) {
  return calendar.getEvents(start, end).length > 0;
}

// 「前日16:00」通知のための分数を計算
function calcNotifyMinutes(eventStart) {
  var d = new Date(eventStart);
  var p = new Date(d); p.setDate(p.getDate() - 1); p.setHours(16, 0, 0, 0);
  return Math.round((d - p) / 60000);
}
