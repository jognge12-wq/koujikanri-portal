/**
 * ============================================================
 * 現場タスク自動登録システム  -  Google Apps Script (Web App版 v3)
 * ============================================================
 *
 * v3変更点:
 *   doGet に以下を追加
 *   - mode=getProperties  → 担当物件一覧DBから物件と日付を取得してJSONで返す
 *   - mode=updateDates    → 指定ページIDの日付3点をPATCHで更新してJSONで返す
 *   - notionPatch() ヘルパー追加
 *
 * ============================================================
 * 【Web アプリとしてデプロイする手順】
 * ============================================================
 *   1. Apps Script エディタを開く
 *   2. このコードを貼り付けて保存（Ctrl+S）
 *   3. 右上「デプロイ」→「既存のデプロイを管理」
 *   4. 鉛筆アイコン（編集）→ バージョン「新バージョン」
 *   5. 「デプロイ」→ URLは変わらずそのまま使えます
 *
 * ============================================================
 */

// ============================================================
// 定数（変更不要）
// ============================================================
const PROPERTY_DB_ID     = '2f56ad84622180a9891bef7e5514fa78'; // 担当物件一覧
const TASK_DB_ID         = '2f66ad84622181b5be55f5da4df2dba3'; // 全現場共通タスク
const NOTION_API_BASE    = 'https://api.notion.com/v1';
const NOTION_API_VERSION = '2022-06-28';


// ============================================================
// Web アプリ: フォーム画面を返す (GET)
// ★ v2: mode=submit でポータルからの登録に対応
// ★ v3: mode=getProperties / mode=updateDates を追加
// ============================================================
function doGet(e) {
  // アイコン配信
  if (e && e.parameter && e.parameter.mode === 'icon') {
    var svg = '<svg xmlns="http://www.w3.org/2000/svg" width="180" height="180" viewBox="0 0 180 180"><rect width="180" height="180" rx="40" fill="%231D6B40"/><path d="M90 35L30 90h15v55h35v-35h20v35h35V90h15L90 35z" fill="white"/></svg>';
    return ContentService.createTextOutput(svg).setMimeType(ContentService.MimeType.XML);
  }

  // ★ v3追加: 物件一覧取得
  if (e && e.parameter && e.parameter.mode === 'getProperties') {
    return handleGetProperties();
  }

  // ★ v3追加: 日付更新
  if (e && e.parameter && e.parameter.mode === 'updateDates') {
    return handleUpdateDates(e.parameter);
  }

  // ★ ポータルサイトからのフォーム送信
  if (e && e.parameter && e.parameter.mode === 'submit') {
    return handlePortalSubmit(e.parameter);
  }

  // ★ v4追加: ダッシュボードデータ一括取得（優先タスク + カレンダー）
  if (e && e.parameter && e.parameter.mode === 'getDashboardData') {
    return handleGetDashboardData();
  }

  // ★ v5追加: Googleタスクを完了にする
  if (e && e.parameter && e.parameter.mode === 'completeGoogleTask') {
    return handleCompleteGoogleTask(e.parameter);
  }

  // 通常: フォーム画面を返す（iPhoneから直接アクセスした場合）
  return HtmlService.createHtmlOutput(getFormHtml())
    .setTitle('新規物件登録')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// ============================================================
// ★ v3追加: 担当物件一覧DB から物件と日付を取得してJSONで返す
// ============================================================
function handleGetProperties() {
  try {
    var token = getToken();

    // 工程作成対象の物件のみ取得:
    //   - 「進捗」に「引渡し」を含む物件を除外（引渡し済み）
    //   - 「本体着工」が今日以前の物件を除外（着工済み）
    //   → 着工前の物件のみ、本体着工日の昇順で表示
    var now = new Date();
    var todayStr = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');

    var result = notionPost('/databases/' + PROPERTY_DB_ID + '/query', {
      filter: {
        and: [
          { property: '進捗', multi_select: { does_not_contain: '引渡し' } },
          {
            or: [
              { property: '本体着工', date: { is_empty: true } },
              { property: '本体着工', date: { after: todayStr } }
            ]
          }
        ]
      },
      sorts: [{ property: '本体着工', direction: 'ascending' }],
      page_size: 100
    }, token);

    if (result.object === 'error') throw new Error(result.message);

    var properties = (result.results || []).map(function(page) {
      var props = page.properties || {};
      var titleArr = props['物件名'] && props['物件名'].title ? props['物件名'].title : [];
      var name = titleArr.length > 0 ? titleArr[0].plain_text : '';
      var chakou      = props['本体着工'] && props['本体着工'].date ? props['本体着工'].date.start : null;
      var tatemae     = props['建て方']   && props['建て方'].date   ? props['建て方'].date.start   : null;
      var hikiwatashi = props['引渡し']   && props['引渡し'].date   ? props['引渡し'].date.start   : null;
      var cityArr     = props['市町村']   && props['市町村'].multi_select ? props['市町村'].multi_select : [];
      var city        = cityArr.length > 0 ? cityArr[0].name : '';
      return { id: page.id, name: name, chakou: chakou, tatemae: tatemae, hikiwatashi: hikiwatashi, city: city };
    }).filter(function(p) { return p.name !== ''; });

    return ContentService.createTextOutput(JSON.stringify({
      success: true, properties: properties
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ getProperties エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// ★ v3追加: 指定ページIDの日付プロパティをPATCHで更新
// パラメータ: pageId, 本体着工, 建て方, 引渡し（日付文字列 YYYY-MM-DD）
// ============================================================
function handleUpdateDates(p) {
  try {
    var token  = getToken();
    var pageId = p['pageId'];
    if (!pageId) throw new Error('pageId が指定されていません');

    var updates = {};
    var chakou      = p['本体着工'] || '';
    var tatemae     = p['建て方']   || '';
    var hikiwatashi = p['引渡し']   || '';

    if (chakou)      updates['本体着工'] = { date: { start: chakou } };
    if (tatemae)     updates['建て方']   = { date: { start: tatemae } };
    if (hikiwatashi) updates['引渡し']   = { date: { start: hikiwatashi } };

    if (Object.keys(updates).length === 0) {
      return ContentService.createTextOutput(JSON.stringify({
        success: true, message: '更新対象なし'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var result = notionPatch('/pages/' + pageId, { properties: updates }, token);
    if (result.object === 'error') throw new Error(result.message);

    Logger.log('✓ 日付更新完了: ' + pageId);
    return ContentService.createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ updateDates エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// ★ v2追加: ポータルからのfetch送信を処理してJSONを返す
// ============================================================
function handlePortalSubmit(p) {
  try {
    var propertyName = (p['物件名'] || '').trim();
    if (!propertyName) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, message: '物件名が入力されていません'
      })).setMimeType(ContentService.MimeType.JSON);
    }

    var city         = p['市町村'] || '';
    var startDate    = p['本体着工'] || null;
    var frameDate    = p['建て方']   || null;
    var deliveryDate = p['引渡し']   || null;
    var sakiSoto  = p['先外']  === 'on';
    var kairo     = p['改良']  === 'on';
    var gaiko     = p['外構']  === 'on';
    var shizumono = p['鎮物']  === 'on';
    var munefuda  = p['棟札']  === 'on';

    Logger.log('▶ ポータル登録開始: ' + propertyName);

    var propertyPageId = createPropertyPage(
      propertyName, city,
      startDate || null, frameDate || null, deliveryDate || null,
      sakiSoto, kairo, gaiko, shizumono, munefuda
    );
    Logger.log('✓ 物件作成完了 (ID: ' + propertyPageId + ')');

    var tasks = getMasterTasks();
    Logger.log('✓ タスク取得: ' + tasks.length + '件');

    var BATCH_SIZE = 20;
    for (var i = 0; i < tasks.length; i += BATCH_SIZE) {
      createTasksBatch(tasks.slice(i, i + BATCH_SIZE), propertyPageId);
      if (i + BATCH_SIZE < tasks.length) Utilities.sleep(400);
    }

    Logger.log('✅ 完了: 「' + propertyName + '」に ' + tasks.length + '件');
    return ContentService.createTextOutput(JSON.stringify({
      success: true, name: propertyName, count: tasks.length
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}


// ============================================================
// Web アプリ: フォーム送信を受け取る (POST) ※既存フォーム用
// ============================================================
function doPost(e) {
  try {
    const p = e.parameter;

    const propertyName = (p['物件名'] || '').trim();
    if (!propertyName) {
      return HtmlService.createHtmlOutput(getResultHtml(false, '物件名が入力されていません'));
    }

    const city         = p['市町村'] || '';
    const startDate    = p['本体着工'] || null;
    const frameDate    = p['建て方']   || null;
    const deliveryDate = p['引渡し']   || null;
    const sakiSoto  = p['先外']  === 'on';
    const kairo     = p['改良']  === 'on';
    const gaiko     = p['外構']  === 'on';
    const shizumono = p['鎮物']  === 'on';
    const munefuda  = p['棟札']  === 'on';

    Logger.log('▶ Web登録開始: ' + propertyName);

    const propertyPageId = createPropertyPage(
      propertyName, city, startDate, frameDate, deliveryDate,
      sakiSoto, kairo, gaiko, shizumono, munefuda
    );
    Logger.log('✓ 物件作成完了 (ID: ' + propertyPageId + ')');

    const tasks = getMasterTasks();
    Logger.log('✓ タスク取得: ' + tasks.length + '件');

    const BATCH_SIZE = 20;
    for (let i = 0; i < tasks.length; i += BATCH_SIZE) {
      const batch = tasks.slice(i, i + BATCH_SIZE);
      createTasksBatch(batch, propertyPageId);
      if (i + BATCH_SIZE < tasks.length) Utilities.sleep(400);
    }

    Logger.log('✅ 完了: 「' + propertyName + '」に ' + tasks.length + '件');
    return HtmlService.createHtmlOutput(
      getResultHtml(true, propertyName, tasks.length)
    );

  } catch (err) {
    Logger.log('❌ エラー: ' + err.message);
    return HtmlService.createHtmlOutput(getResultHtml(false, err.message));
  }
}


// ============================================================
// サーバー関数: google.script.run から呼ばれる
// ============================================================
function processForm(formData) {
  var propertyName = (formData['物件名'] || '').trim();
  if (!propertyName) return { success: false, message: '物件名が入力されていません' };
  var city = formData['市町村'] || '';
  var startDate = formData['本体着工'] || null;
  var frameDate = formData['建て方'] || null;
  var deliveryDate = formData['引渡し'] || null;
  var sakiSoto = formData['先外'] === true;
  var kairo = formData['改良'] === true;
  var gaiko = formData['外構'] === true;
  var shizumono = formData['鎮物'] === true;
  var munefuda = formData['棟札'] === true;
  Logger.log('▶ Web登録開始: ' + propertyName);
  try {
    var propertyPageId = createPropertyPage(propertyName, city, startDate, frameDate, deliveryDate, sakiSoto, kairo, gaiko, shizumono, munefuda);
    var tasks = getMasterTasks();
    var BATCH_SIZE = 20;
    for (var i = 0; i < tasks.length; i += BATCH_SIZE) {
      createTasksBatch(tasks.slice(i, i + BATCH_SIZE), propertyPageId);
      if (i + BATCH_SIZE < tasks.length) Utilities.sleep(400);
    }
    return { success: true, name: propertyName, count: tasks.length };
  } catch (err) {
    Logger.log('✘ エラー: ' + err.message);
    return { success: false, message: err.message };
  }
}


// ============================================================
// HTML: 登録フォーム（iPhoneから直接アクセスした場合に表示）
// ============================================================
function getFormHtml() {
  return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<meta name="apple-mobile-web-app-capable" content="yes">
<meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
<meta name="apple-mobile-web-app-title" content="物件登録">
<title>新規物件登録</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, BlinkMacSystemFont, 'Hiragino Kaku Gothic ProN', sans-serif;
    font-size: 16px;
    background: #f0f4f2;
    min-height: 100vh;
    padding: 0 0 20px;
  }
  header { display: none; }
  .card {
    background: #fff;
    border-radius: 14px;
    margin: 8px 16px;
    padding: 12px 14px;
    box-shadow: 0 1px 4px rgba(0,0,0,0.08);
  }
  .card h2 {
    font-size: 11px;
    font-weight: 700;
    color: #1D6B40;
    text-transform: uppercase;
    letter-spacing: 0.06em;
    margin-bottom: 10px;
    padding-bottom: 8px;
    border-bottom: 1px solid #e8f0eb;
  }
  .field { margin-bottom: 10px; }
  .field:last-child { margin-bottom: 0; }
  label {
    display: block;
    font-size: 13px;
    font-weight: 600;
    color: #333;
    margin-bottom: 5px;
  }
  label .required {
    color: #e05555;
    font-size: 11px;
    margin-left: 3px;
    font-weight: 500;
  }
  input[type="text"],
  input[type="date"],
  select {
    display: block;
    width: 100%;
    padding: 8px 10px;
    font-size: 14px;
    font-family: inherit;
    border: 1.5px solid #dde8e2;
    border-radius: 10px;
    background: #fff;
    color: #222;
    -webkit-appearance: none;
    appearance: none;
    transition: border-color 0.2s;
  }
  input[type="text"]:focus,
  input[type="date"]:focus,
  select:focus {
    outline: none;
    border-color: #2D8A4E;
    box-shadow: 0 0 0 3px rgba(45,138,78,0.12);
  }
  select {
    background-image: url("data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' width='12' height='8' viewBox='0 0 12 8'%3E%3Cpath d='M1 1l5 5 5-5' stroke='%23999' stroke-width='1.5' fill='none' stroke-linecap='round'/%3E%3C/svg%3E");
    background-repeat: no-repeat;
    background-position: right 12px center;
    padding-right: 32px;
  }
  .date-row {
    position: relative;
  }
  .date-display {
    display: flex;
    align-items: center;
    justify-content: space-between;
    width: 100%;
    padding: 8px 10px;
    font-size: 14px;
    border: 1.5px solid #dde8e2;
    border-radius: 10px;
    background: #fff;
    cursor: pointer;
    min-height: 40px;
  }
  .date-display.filled { border-color: #2D8A4E; }
  .date-placeholder { color: #aab8c2; font-size: 14px; }
  .date-hidden {
    position: absolute;
    top: 0; left: 0;
    width: 100%; height: 100%;
    opacity: 0;
    cursor: pointer;
    font-size: 16px;
  }
  .cal-icon { flex-shrink: 0; opacity: 0.4; }
  .date-fields { display: grid; grid-template-columns: 1fr; gap: 10px; }
  .check-grid { display: grid; grid-template-columns: 1fr; gap: 7px; }
  .check-item {
    display: flex;
    align-items: center;
    gap: 8px;
    padding: 9px 12px;
    border: 1.5px solid #dde8e2;
    border-radius: 10px;
    cursor: pointer;
    transition: background 0.15s;
    -webkit-tap-highlight-color: transparent;
  }
  .check-item:active { background: #f0f8f4; }
  .check-item input[type="checkbox"] {
    width: 20px; height: 20px;
    border-radius: 6px;
    border: 2px solid #bbb;
    appearance: none; -webkit-appearance: none;
    background: #fff;
    cursor: pointer;
    flex-shrink: 0;
    transition: all 0.15s;
    position: relative;
  }
  .check-item input[type="checkbox"]:checked {
    background: #2D8A4E;
    border-color: #2D8A4E;
  }
  .check-item input[type="checkbox"]:checked::after {
    content: '';
    position: absolute;
    left: 5px; top: 2px;
    width: 6px; height: 10px;
    border: 2.5px solid #fff;
    border-top: none; border-left: none;
    transform: rotate(45deg);
  }
  .check-item span {
    font-size: 13px;
    font-weight: 500;
    color: #333;
  }
  .btn-submit {
    display: block;
    width: calc(100% - 32px);
    margin: 12px 16px;
    padding: 12px;
    background: #1D6B40;
    color: #fff;
    font-size: 14px;
    font-weight: 700;
    font-family: inherit;
    border: none;
    border-radius: 12px;
    cursor: pointer;
    box-shadow: 0 4px 14px rgba(29,107,64,0.3);
    -webkit-tap-highlight-color: transparent;
    transition: opacity 0.15s;
  }
  .btn-submit:active { opacity: 0.8; }
  .btn-submit:disabled { opacity: 0.6; cursor: not-allowed; }
</style>
</head>
<body>
<header>
  <div>
    <h1>新規物件登録</h1>
    <p>Notion 担当物件一覧へ登録</p>
  </div>
</header>

<form method="POST" id="regForm" onsubmit="handleSubmit(event)">

  <div class="card">
    <h2>基本情報</h2>
    <div class="field">
      <label>物件名 <span class="required">*</span></label>
      <input type="text" name="物件名" id="propName" placeholder="〇〇様邸" required>
    </div>
    <div class="field">
      <label>市町村</label>
      <select name="市町村">
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
    <h2>工程</h2>
    <div class="date-fields">
      <div class="field">
        <label>本体着工</label>
        <div class="date-row">
          <div class="date-display" id="dd_chakou">
            <span id="dw_chakou" class="date-placeholder">\u65e5\u4ed8\u3092\u9078\u629e</span>
            <svg class="cal-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#333" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="16" y1="2" x2="16" y2="6"/></svg>
          </div>
          <input type="date" name="本体着工" id="d_chakou" class="date-hidden" oninput="updateDay('d_chakou','dw_chakou','dd_chakou')">
        </div>
      </div>
      <div class="field">
        <label>建て方</label>
        <div class="date-row">
          <div class="date-display" id="dd_tatemae">
            <span id="dw_tatemae" class="date-placeholder">\u65e5\u4ed8\u3092\u9078\u629e</span>
            <svg class="cal-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#333" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="16" y1="2" x2="16" y2="6"/></svg>
          </div>
          <input type="date" name="建て方" id="d_tatemae" class="date-hidden" oninput="updateDay('d_tatemae','dw_tatemae','dd_tatemae')">
        </div>
      </div>
      <div class="field">
        <label>引渡し</label>
        <div class="date-row">
          <div class="date-display" id="dd_hiki">
            <span id="dw_hiki" class="date-placeholder">\u65e5\u4ed8\u3092\u9078\u629e</span>
            <svg class="cal-icon" width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="#333" stroke-width="2"><rect x="3" y="4" width="18" height="18" rx="2"/><line x1="3" y1="10" x2="21" y2="10"/><line x1="8" y1="2" x2="8" y2="6"/><line x1="16" y1="2" x2="16" y2="6"/></svg>
          </div>
          <input type="date" name="引渡し" id="d_hiki" class="date-hidden" oninput="updateDay('d_hiki','dw_hiki','dd_hiki')">
        </div>
      </div>
    </div>
  </div>

  <div class="card">
    <h2>チェック項目</h2>
    <div class="check-grid">
      <label class="check-item">
        <input type="checkbox" name="棟札" value="on"> <span>棟札</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="鎮物" value="on"> <span>鎮め物</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="先外" value="on"> <span>先行外構</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="改良" value="on"> <span>地盤改良</span>
      </label>
      <label class="check-item">
        <input type="checkbox" name="外構" value="on"> <span>外構工事</span>
      </label>
    </div>
  </div>

  <button type="submit" class="btn-submit" id="submitBtn">Notion に登録する</button>

</form>

<script>
var DAY_NAMES = ['\u65e5','\u6708','\u706b','\u6c34','\u6728','\u91d1','\u571f'];

function updateDay(inputId, spanId, boxId) {
  var val = document.getElementById(inputId).value;
  var span = document.getElementById(spanId);
  var box  = document.getElementById(boxId);
  if (!val) {
    span.className = 'date-placeholder';
    span.innerHTML = '\u65e5\u4ed8\u3092\u9078\u629e';
    box.classList.remove('filled');
    return;
  }
  var d   = new Date(val);
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
  btn.textContent = '登録中...';
  e.target.submit();
}
</script>
</body>
</html>`;
}


// ============================================================
// HTML: 結果画面
// ============================================================
function getResultHtml(success, messageOrName, taskCount) {
  if (success) {
    return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>登録完了</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, sans-serif; background: #f0f4f2; min-height: 100vh;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    padding: 40px 24px; text-align: center;
  }
  .icon { font-size: 64px; margin-bottom: 20px; }
  h1 { font-size: 22px; font-weight: 700; color: #1D6B40; margin-bottom: 8px; }
  .name { font-size: 18px; font-weight: 600; color: #333; margin-bottom: 6px; }
  .count { font-size: 14px; color: #666; margin-bottom: 32px; }
  a.btn {
    display: block; padding: 14px 32px;
    background: #1D6B40; color: #fff;
    font-size: 16px; font-weight: 700;
    border-radius: 12px; text-decoration: none;
    box-shadow: 0 3px 10px rgba(29,107,64,0.3);
  }
</style>
</head>
<body>
  <div class="icon">✅</div>
  <h1>登録完了！</h1>
  <div class="name">「${messageOrName}」</div>
  <div class="count">${taskCount}件のタスクを作成しました</div>
  <a class="btn" href="javascript:history.back()">続けて登録する</a>
</body>
</html>`;
  } else {
    return `<!DOCTYPE html>
<html lang="ja">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1">
<title>エラー</title>
<style>
  * { box-sizing: border-box; margin: 0; padding: 0; }
  body {
    font-family: -apple-system, sans-serif; background: #fff5f5; min-height: 100vh;
    display: flex; flex-direction: column;
    align-items: center; justify-content: center;
    padding: 40px 24px; text-align: center;
  }
  .icon { font-size: 64px; margin-bottom: 20px; }
  h1 { font-size: 20px; font-weight: 700; color: #c0392b; margin-bottom: 12px; }
  p { font-size: 14px; color: #666; margin-bottom: 32px; line-height: 1.6; }
  a.btn {
    display: block; padding: 14px 32px;
    background: #555; color: #fff;
    font-size: 16px; font-weight: 700;
    border-radius: 12px; text-decoration: none;
  }
</style>
</head>
<body>
  <div class="icon">❌</div>
  <h1>登録エラー</h1>
  <p>${messageOrName}</p>
  <a class="btn" href="javascript:history.back()">戻る</a>
</body>
</html>`;
  }
}


// ============================================================
// Notion: 物件ページを作成
// ============================================================
function createPropertyPage(name, city, startDate, frameDate, deliveryDate,
                             sakiSoto, kairo, gaiko, shizumono, munefuda) {
  const token = getToken();

  const properties = {
    '物件名': { title: [{ text: { content: name } }] },
    '進捗':   { multi_select: [{ name: '着工前' }] },
    '先外':   { checkbox: sakiSoto  },
    '改良':   { checkbox: kairo     },
    '外構':   { checkbox: gaiko     },
    '鎮物':   { checkbox: shizumono },
    '棟札':   { checkbox: munefuda  }
  };

  if (city && city !== '（選択してください）') {
    properties['市町村'] = { multi_select: [{ name: city }] };
  }
  if (startDate)    properties['本体着工'] = { date: { start: startDate } };
  if (frameDate)    properties['建て方']   = { date: { start: frameDate } };
  if (deliveryDate) properties['引渡し']   = { date: { start: deliveryDate } };

  const result = notionPost('/pages', {
    parent:     { database_id: PROPERTY_DB_ID },
    properties: properties
  }, token);

  if (result.object === 'error') {
    throw new Error('物件作成エラー: ' + result.message);
  }

  return result.id;
}


// ============================================================
// Notion: タスクをまとめて作成
// ============================================================
function createTasksBatch(tasks, propertyPageId) {
  const token = getToken();

  for (const task of tasks) {
    const result = notionPost('/pages', {
      parent: { database_id: TASK_DB_ID },
      properties: {
        'タスク名': { title: [{ text: { content: task.taskName } }] },
        '工事進捗': { status: { name: task.stage } },
        '並び順':   { number: task.order },
        '物件名':   { relation: [{ id: propertyPageId }] }
      }
    }, token);

    if (result.object === 'error') {
      Logger.log('タスク作成エラー「' + task.taskName + '」: ' + result.message);
    }
  }
}


// ============================================================
// Googleスプレッドシート: マスタータスクを取得
// ============================================================
function getMasterTasks() {
  const sheetId = PropertiesService.getScriptProperties().getProperty('MASTER_SHEET_ID');
  if (!sheetId) throw new Error('スクリプトプロパティ「MASTER_SHEET_ID」が設定されていません');

  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName('マスタータスク一覧');
  if (!sheet) throw new Error('シート「マスタータスク一覧」が見つかりません');

  const data  = sheet.getDataRange().getValues();
  const tasks = [];

  for (let i = 1; i < data.length; i++) {
    const order    = data[i][0];
    const stage    = data[i][1];
    const taskName = data[i][3];

    if (order && taskName) {
      tasks.push({
        order:    Number(order),
        stage:    String(stage).trim(),
        taskName: String(taskName).trim()
      });
    }
  }

  return tasks.sort((a, b) => a.order - b.order);
}


// ============================================================
// 既存ページへのタスクバックフィル（任意）
// ============================================================
function backfillTasks() {
  const propertyPageId = '3356ad84622181609481c1c5c07e7217'; // ← 対象ページID

  Logger.log('▶ バックフィル開始: ' + propertyPageId);
  const tasks = getMasterTasks();
  Logger.log('✓ タスク: ' + tasks.length + '件');

  const BATCH_SIZE = 20;
  for (let i = 0; i < tasks.length; i += BATCH_SIZE) {
    createTasksBatch(tasks.slice(i, i + BATCH_SIZE), propertyPageId);
    Logger.log('  登録中: ' + Math.min(i + BATCH_SIZE, tasks.length) + '/' + tasks.length);
    if (i + BATCH_SIZE < tasks.length) Utilities.sleep(400);
  }
  Logger.log('✅ バックフィル完了！ ' + tasks.length + '件');
}


// ============================================================
// Google フォームからのトリガー（既存フォームを使う場合）
// ============================================================
function onFormSubmit(e) {
  try {
    const r = {};
    if (e && e.namedValues) {
      Object.assign(r, e.namedValues);
    } else if (e && e.response) {
      e.response.getItemResponses().forEach(function(ir) {
        r[ir.getItem().getTitle()] = [ir.getResponse()];
      });
    } else {
      throw new Error('フォームデータが取得できませんでした');
    }

    const propertyName = (r['物件名'] || [''])[0].trim();
    if (!propertyName) throw new Error('物件名が入力されていません');

    const city         = (r['市町村']  || [''])[0];
    const startDate    = parseDate((r['本体着工'] || [''])[0]);
    const frameDate    = parseDate((r['建て方']   || [''])[0]);
    const deliveryDate = parseDate((r['引渡し']   || [''])[0]);
    const sakiSoto  = (r['先外'] || ['いいえ'])[0] === 'はい';
    const kairo     = (r['改良'] || ['いいえ'])[0] === 'はい';
    const gaiko     = (r['外構'] || ['いいえ'])[0] === 'はい';
    const shizumono = (r['鎮物'] || ['いいえ'])[0] === 'はい';
    const munefuda  = (r['棟札'] || ['いいえ'])[0] === 'はい';

    Logger.log('▶ 登録開始: ' + propertyName);

    const propertyPageId = createPropertyPage(
      propertyName, city, startDate, frameDate, deliveryDate,
      sakiSoto, kairo, gaiko, shizumono, munefuda
    );

    const tasks = getMasterTasks();
    const BATCH_SIZE = 20;
    for (let i = 0; i < tasks.length; i += BATCH_SIZE) {
      createTasksBatch(tasks.slice(i, i + BATCH_SIZE), propertyPageId);
      Logger.log('  タスク: ' + Math.min(i + BATCH_SIZE, tasks.length) + '/' + tasks.length);
      if (i + BATCH_SIZE < tasks.length) Utilities.sleep(400);
    }

    Logger.log('✅ 完了: 「' + propertyName + '」に ' + tasks.length + '件');
  } catch (err) {
    Logger.log('❌ エラー: ' + err.message);
    throw err;
  }
}


// ============================================================
// Notion API ヘルパー: POST
// ============================================================
function notionPost(endpoint, payload, token) {
  var MAX_RETRIES = 3;

  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    var res = UrlFetchApp.fetch(NOTION_API_BASE + endpoint, {
      method:  'post',
      headers: {
        'Authorization':  'Bearer ' + token,
        'Notion-Version': NOTION_API_VERSION,
        'Content-Type':   'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var body = res.getContentText();

    if (code >= 200 && code < 500) {
      try {
        return JSON.parse(body);
      } catch (e) {
        Logger.log('⚠ JSONパース失敗 (HTTP ' + code + '): ' + body.substring(0, 200));
        return { object: 'error', message: 'レスポンスがJSONではありません (HTTP ' + code + ')' };
      }
    }

    var wait = (code === 429) ? 2000 : 1000 * attempt;
    Logger.log('⚠ HTTP ' + code + ' - ' + wait + 'ms後にリトライ (' + attempt + '/' + MAX_RETRIES + ')');
    Utilities.sleep(wait);
  }

  return { object: 'error', message: 'Notion API が応答しません (HTTP ' + res.getResponseCode() + ')' };
}


// ============================================================
// ★ v3追加: Notion API ヘルパー: PATCH
// ============================================================
function notionPatch(endpoint, payload, token) {
  var MAX_RETRIES = 3;

  for (var attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    var res = UrlFetchApp.fetch(NOTION_API_BASE + endpoint, {
      method:  'patch',
      headers: {
        'Authorization':  'Bearer ' + token,
        'Notion-Version': NOTION_API_VERSION,
        'Content-Type':   'application/json'
      },
      payload:            JSON.stringify(payload),
      muteHttpExceptions: true
    });

    var code = res.getResponseCode();
    var body = res.getContentText();

    if (code >= 200 && code < 500) {
      try {
        return JSON.parse(body);
      } catch (e) {
        return { object: 'error', message: 'レスポンスがJSONではありません (HTTP ' + code + ')' };
      }
    }

    var wait = (code === 429) ? 2000 : 1000 * attempt;
    Logger.log('⚠ PATCH HTTP ' + code + ' - リトライ (' + attempt + '/' + MAX_RETRIES + ')');
    Utilities.sleep(wait);
  }

  return { object: 'error', message: 'Notion API が応答しません' };
}


function getToken() {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('スクリプトプロパティ「NOTION_TOKEN」が設定されていません');
  return token;
}

function parseDate(dateStr) {
  if (!dateStr || dateStr.trim() === '') return null;
  const s = dateStr.trim();
  const m1 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
  if (m1) return m1[1] + '-' + m1[2].padStart(2,'0') + '-' + m1[3].padStart(2,'0');
  const m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m2) return m2[3] + '-' + m2[1].padStart(2,'0') + '-' + m2[2].padStart(2,'0');
  try {
    const d = new Date(s);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, 'Asia/Tokyo', 'yyyy-MM-dd');
  } catch(e) {}
  return null;
}


// ============================================================
// ★ v5追加: Googleタスクを完了にする
// mode=completeGoogleTask&taskId=xxx&listId=yyy
// ============================================================
function handleCompleteGoogleTask(params) {
  try {
    var taskId = params.taskId;
    var listId = params.listId;
    if (!taskId || !listId) {
      return ContentService.createTextOutput(JSON.stringify({
        success: false, error: 'パラメータ不足（taskId・listIdが必要）'
      })).setMimeType(ContentService.MimeType.JSON);
    }
    Tasks.Tasks.patch({ status: 'completed' }, listId, taskId);
    return ContentService.createTextOutput(JSON.stringify({
      success: true
    })).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    Logger.log('❌ completeGoogleTask エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, error: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}

// ============================================================
// ★ v4追加: ダッシュボードデータ一括取得
// mode=getDashboardData
//   - Notion優先タスク（優先=YES, 完了=NO）をリアルタイム取得
//   - Googleカレンダーの着工・引渡しイベントをリアルタイム取得
//   - 返値: JSON { success, timestamp, priorityTasks, calendarData }
// ============================================================
function handleGetDashboardData() {
  try {
    var token = getToken();
    var now = new Date();
    var timestamp = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd HH:mm');
    var today     = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    var currentMonth = parseInt(Utilities.formatDate(now, 'Asia/Tokyo', 'M')) - 1; // 0-based

    // ── 1. 物件マップ（引渡し日・建て方日キャッシュ）──────────────────
    var propsResult = notionPost('/databases/' + PROPERTY_DB_ID + '/query', { page_size: 100 }, token);
    var propertyMap = {}; // cleanId → { name, hikiwatashi, tatemae }
    (propsResult.results || []).forEach(function(page) {
      var pr    = page.properties || {};
      var tArr  = (pr['物件名'] && pr['物件名'].title) ? pr['物件名'].title : [];
      var name  = tArr.length > 0 ? tArr[0].plain_text : '';
      var hiki  = (pr['引渡し'] && pr['引渡し'].date)  ? pr['引渡し'].date.start  : null;
      var tate  = (pr['建て方'] && pr['建て方'].date)   ? pr['建て方'].date.start  : null;
      propertyMap[page.id.replace(/-/g, '')] = { name: name, hikiwatashi: hiki, tatemae: tate };
    });

    // ── 2. 優先タスク（優先=YES, 完了=NO）────────────────────────────
    var tasksResult = notionPost('/databases/' + TASK_DB_ID + '/query', {
      filter: {
        and: [
          { property: '優先',  checkbox: { equals: true  } },
          { property: '完了',  checkbox: { equals: false } }
        ]
      },
      sorts: [{ property: '物件名', direction: 'ascending' }],
      page_size: 100
    }, token);

    var priorityTasks = [];
    var taskId = 1;
    (tasksResult.results || []).forEach(function(page) {
      var pr    = page.properties || {};
      var tArr  = (pr['タスク名'] && pr['タスク名'].title) ? pr['タスク名'].title : [];
      var title = tArr.length > 0 ? tArr[0].plain_text : '';
      if (!title) return;

      var relArr           = (pr['物件名'] && pr['物件名'].relation) ? pr['物件名'].relation : [];
      var propertyNotionId = relArr.length > 0 ? relArr[0].id.replace(/-/g, '') : '';
      var propertyName     = (propertyNotionId && propertyMap[propertyNotionId]) ? propertyMap[propertyNotionId].name : '';
      var stage            = (pr['工事進捗'] && pr['工事進捗'].select) ? pr['工事進捗'].select.name : '';

      // ステータス計算（次マイルストーンまで何日か）
      var status = 'green';
      if (propertyNotionId && propertyMap[propertyNotionId]) {
        var dates      = propertyMap[propertyNotionId];
        var refDateStr = dates.hikiwatashi || dates.tatemae;
        if (refDateStr) {
          var refDate   = new Date(refDateStr + 'T00:00:00+09:00');
          var daysUntil = Math.floor((refDate.getTime() - now.getTime()) / 86400000);
          if      (daysUntil <= 7)  status = 'red';
          else if (daysUntil <= 14) status = 'amber';
        }
      }

      priorityTasks.push({
        id: taskId++,
        notionPageId:    page.id.replace(/-/g, ''),
        title:           title,
        property:        propertyName,
        propertyNotionId: propertyNotionId,
        stage:           stage,
        status:          status,
        completed:       false
      });
    });

    // ── 3. Googleカレンダー（着工/引渡し統計 + 週間スケジュール）─────
    //    ※ getAllCalendars() は1回のみ呼び出し（タイムアウト防止）
    var calErrorMsg = '';
    var startedProperties  = [];
    var monthlyStarts      = [];
    var deliveredProperties = [];
    var monthlyDeliveries  = [];
    var yearlyStartCounts    = [0,0,0,0,0,0,0,0,0,0,0,0];
    var yearlyDeliveryCounts = [0,0,0,0,0,0,0,0,0,0,0,0];
    var scheduleEvents = [];
    var todayStr     = Utilities.formatDate(now, 'Asia/Tokyo', 'yyyy-MM-dd');
    var tomorrowDate = new Date(now.getTime() + 24 * 60 * 60 * 1000);
    var tomorrowStr  = Utilities.formatDate(tomorrowDate, 'Asia/Tokyo', 'yyyy-MM-dd');

    // 3月1日〜12月31日のスケジュールを収集（年間カバー）
    var dayOfWeek = now.getDay();
    var weekSunday = new Date(now.getFullYear(), now.getMonth(), now.getDate() - dayOfWeek);
    var schedMap = {};
    var weekDateStrs = [];
    // 過去2ヶ月〜未来12ヶ月（今日基準で動的に算出）
    var schedStart = new Date(now.getFullYear(), now.getMonth() - 2, 1); // 2ヶ月前の1日
    var schedEnd   = new Date(now.getFullYear(), now.getMonth() + 13, 0); // 12ヶ月後の月末
    for (var wd = new Date(schedStart); wd <= schedEnd; wd.setDate(wd.getDate() + 1)) {
      var wdStr = Utilities.formatDate(wd, 'Asia/Tokyo', 'yyyy-MM-dd');
      schedMap[wdStr] = [];
      weekDateStrs.push(wdStr);
    }

    // イベント個別色 → Hex変換マップ（Googleカレンダー EventColor ID）
    var evColorMap = {
      '1':'#a4bdfc','2':'#7ae7bf','3':'#dbadff','4':'#ff887c',
      '5':'#fbd75b','6':'#ffb878','7':'#46d6db','8':'#e1e1e1',
      '9':'#5484ed','10':'#51b749','11':'#dc2127'
    };

    try {
      var year       = now.getFullYear();
      // タイムゾーン明示（toISOString()はUTC変換するため+09:00で指定）
      var rangeStartISO = year + '-01-01T00:00:00+09:00';
      var rangeEndISO   = (year + 1) + '-01-01T00:00:00+09:00';

      // Calendar Advanced Service でカレンダー一覧取得
      var calListResult = Calendar.CalendarList.list();
      var calendars = calListResult.items || [];

      calendars.forEach(function(cal) {
        var calId    = cal.id;
        var calName  = cal.summary || '';
        var calColor = cal.backgroundColor || '#4285f4';
        var isSys    = calName.indexOf('祝日') !== -1 || calName.toLowerCase().indexOf('birthday') !== -1;

        // イベント取得（ページネーション対応）
        var pageToken = null;
        do {
          var params = {
            timeMin: rangeStartISO,
            timeMax: rangeEndISO,
            singleEvents: true,
            orderBy: 'startTime',
            maxResults: 2500
          };
          if (pageToken) params.pageToken = pageToken;
          var evResult = Calendar.Events.list(calId, params);
          var items = evResult.items || [];

          items.forEach(function(ev) {
            var evTitle   = ev.summary || '';
            var isAllDay  = !!ev.start.date;
            var evStartRaw = isAllDay ? ev.start.date : ev.start.dateTime;
            // 全日イベントはJST明示（"2026-04-09"→UTC解釈を防止）
            var evStart   = isAllDay ? new Date(evStartRaw + 'T00:00:00+09:00') : new Date(evStartRaw);
            var evDateStr = Utilities.formatDate(evStart, 'Asia/Tokyo', 'yyyy-MM-dd');

            // イベント個別色があればそれを使用、なければカレンダー色
            var evColorId = ev.colorId || '';
            var color = (evColorId && evColorMap[evColorId]) ? evColorMap[evColorId] : calColor;

            // 対象期間の全イベントを収集（祝日カレンダーも含む）
            if (schedMap[evDateStr] !== undefined) {
              var startTime = null, endTime = null;
              if (!isAllDay) {
                startTime = Utilities.formatDate(evStart, 'Asia/Tokyo', 'HH:mm');
                if (ev.end && ev.end.dateTime) {
                  endTime = Utilities.formatDate(new Date(ev.end.dateTime), 'Asia/Tokyo', 'HH:mm');
                }
              }
              var evObj = {
                title:     evTitle,
                isAllDay:  isAllDay,
                startTime: startTime,
                endTime:   endTime,
                color:     color
              };
              // Googleカレンダーの「場所」フィールドがあれば追加
              if (ev.location) {
                evObj.location = ev.location;
              }
              schedMap[evDateStr].push(evObj);
            }

            // 着工・引渡しはシステムカレンダーをスキップ
            if (isSys) return;
            var isStart    = evTitle.indexOf('本体着工') !== -1;
            var isDelivery = evTitle.indexOf('引渡') !== -1;
            if (!isStart && !isDelivery) return;

            var evMonth  = parseInt(Utilities.formatDate(evStart, 'Asia/Tokyo', 'M')) - 1;
            var m        = parseInt(Utilities.formatDate(evStart, 'Asia/Tokyo', 'M'));
            var d        = parseInt(Utilities.formatDate(evStart, 'Asia/Tokyo', 'd'));
            var dateMMDD = m + '/' + d;

            // 物件名抽出（「様邸」の前の部分）
            var nameMatch = evTitle.match(/([^\s　🚜🔑・]+様)邸/);
            if (!nameMatch) return;
            var name = nameMatch[1];

            if (isStart) {
              yearlyStartCounts[evMonth]++;
              if (evDateStr <= today) {
                startedProperties.push({ name: name, date: dateMMDD, _sort: evMonth * 100 + d });
              } else if (evMonth === currentMonth) {
                monthlyStarts.push({ name: name, date: dateMMDD, _sort: evMonth * 100 + d });
              }
            } else {
              yearlyDeliveryCounts[evMonth]++;
              if (evDateStr <= today) {
                deliveredProperties.push({ name: name, date: dateMMDD, _sort: evMonth * 100 + d });
              } else if (evMonth === currentMonth) {
                monthlyDeliveries.push({ name: name, date: dateMMDD, _sort: evMonth * 100 + d });
              }
            }
          });

          pageToken = evResult.nextPageToken;
        } while (pageToken);
      });

      // 日付順ソート後、_sort フィールドを除去
      function sortAndClean(arr) {
        arr.sort(function(a, b) { return a._sort - b._sort; });
        arr.forEach(function(item) { delete item._sort; });
        return arr;
      }
      sortAndClean(startedProperties);
      sortAndClean(monthlyStarts);
      sortAndClean(deliveredProperties);
      sortAndClean(monthlyDeliveries);

      // scheduleEvents を組み立て（前4週〜後5週の63日分）
      scheduleEvents = weekDateStrs.map(function(ds) {
        return { dateStr: ds, events: schedMap[ds] || [] };
      });

    } catch(calErr) {
      Logger.log('⚠ カレンダー取得エラー: ' + calErr.message);
      calErrorMsg = calErr.message || '不明なエラー';
      scheduleEvents = weekDateStrs.map(function(ds) {
        return { dateStr: ds, events: [] };
      });
    }

    // ── 4. Google Tasks（未完了タスク）──────────────────────────────
    var gTasks = [];
    try {
      var taskLists = Tasks.Tasklists.list();
      if (taskLists.items) {
        taskLists.items.forEach(function(tl) {
          var result = Tasks.Tasks.list(tl.id, {
            showCompleted: false,
            showHidden: false,
            maxResults: 100
          });
          if (result.items) {
            result.items.forEach(function(task) {
              if (!task.title) return;
              var obj = { title: task.title, id: task.id, listId: tl.id };
              if (task.due) {
                // GAS Tasks APIはdue を RFC3339 で返す（例: 2026-04-10T00:00:00.000Z）
                var dueDate = new Date(task.due);
                obj.due = Utilities.formatDate(dueDate, 'Asia/Tokyo', 'yyyy-MM-dd');
              }
              // task.notes にdeadlineが書いてある場合（オプション）
              if (task.notes) {
                var dlMatch = task.notes.match(/deadline[:\s]*(\d{4}-\d{2}-\d{2})/i);
                if (dlMatch) obj.deadline = dlMatch[1];
              }
              gTasks.push(obj);
            });
          }
        });
      }
      // 期日順ソート（期日なしは末尾）
      gTasks.sort(function(a, b) {
        if (!a.due && !b.due) return 0;
        if (!a.due) return 1;
        if (!b.due) return -1;
        return a.due < b.due ? -1 : a.due > b.due ? 1 : 0;
      });
    } catch(taskErr) {
      Logger.log('⚠ Google Tasks取得エラー: ' + taskErr.message);
    }

    return ContentService.createTextOutput(JSON.stringify({
      success:        true,
      timestamp:      timestamp,
      calError:       calErrorMsg,
      priorityTasks:  priorityTasks,
      scheduleEvents: scheduleEvents,
      googleTasks:    gTasks,
      calendarData: {
        startedProperties:   startedProperties,
        monthlyStarts:       monthlyStarts,
        deliveredProperties: deliveredProperties,
        monthlyDeliveries:   monthlyDeliveries,
        yearlyStartCounts:   yearlyStartCounts,
        yearlyDeliveryCounts: yearlyDeliveryCounts
      }
    })).setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    Logger.log('❌ getDashboardData エラー: ' + err.message);
    return ContentService.createTextOutput(JSON.stringify({
      success: false, message: err.message
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
