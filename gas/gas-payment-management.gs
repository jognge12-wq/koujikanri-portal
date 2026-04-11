// ============================================================
// 支払い管理 入力フォーム — Google Apps Script (サーバー側)
// v3: 物件フィルターを「建て方〜引渡し1ヶ月」範囲に変更・引渡し順ソート
// ============================================================
// 【設定】スクリプトプロパティに以下を登録してください
//   NOTION_TOKEN : Notion インテグレーショントークン (secret_xxx...)
// ============================================================

// ---------- 定数 ----------
const NOTION_API = 'https://api.notion.com/v1';
const NOTION_VERSION = '2022-06-28';

// Notion データベースID
const DB_ID = {
  物件:       '2f56ad84622180a9891bef7e5514fa78',
  業者:       '93c54c72efdd402e972c3c3bfea1f583',
  担当者:     '46cee566ac534c73a1187f93c4a7887f',
  支払い管理: '134e9a121b8a4b999c47bb70142a8cfb'
};

// ---------- Webアプリ エントリーポイント ----------
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('支払い登録')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ---------- Notion API ヘルパー ----------
function notionHeaders_() {
  const token = PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_TOKEN が設定されていません。スクリプトプロパティに登録してください。');
  return {
    'Authorization': 'Bearer ' + token,
    'Content-Type': 'application/json',
    'Notion-Version': NOTION_VERSION
  };
}

function notionPost_(endpoint, payload) {
  const url = NOTION_API + endpoint;
  const options = {
    method: 'post',
    headers: notionHeaders_(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (res.getResponseCode() >= 400) {
    throw new Error('Notion API Error: ' + JSON.stringify(json));
  }
  return json;
}

function notionPatch_(endpoint, payload) {
  const url = NOTION_API + endpoint;
  const options = {
    method: 'patch',
    headers: notionHeaders_(),
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  const res = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(res.getContentText());
  if (res.getResponseCode() >= 400) {
    throw new Error('Notion API Error: ' + JSON.stringify(json));
  }
  return json;
}

// ---------- データ取得 ----------

/**
 * 物件一覧を取得
 * 抽出範囲: 建て方が設定済み（着工済み） かつ 引渡しから1ヶ月以内
 * ソート: 引渡し日 昇順（未設定は先頭）
 */
function getProperties() {
  const withDate = [];    // 引渡し日あり → 引渡し昇順（近い順）
  const withoutDate = []; // 引渡し日なし → 末尾に表示
  let hasMore = true;
  let startCursor = undefined;

  // 今日と1ヶ月前の日付を計算
  const today = new Date();
  const todayStr = Utilities.formatDate(today, 'Asia/Tokyo', 'yyyy-MM-dd');
  const oneMonthAgo = new Date();
  oneMonthAgo.setMonth(oneMonthAgo.getMonth() - 1);
  const oneMonthAgoStr = Utilities.formatDate(oneMonthAgo, 'Asia/Tokyo', 'yyyy-MM-dd');

  while (hasMore) {
    const payload = {
      filter: {
        and: [
          // 建て方が設定済みかつ今日以前（着工済み）
          { property: '建て方', date: { is_not_empty: true } },
          { property: '建て方', date: { on_or_before: todayStr } },
          // 引渡しが未設定 OR 引渡しが1ヶ月前以降
          {
            or: [
              { property: '引渡し', date: { is_empty: true } },
              { property: '引渡し', date: { on_or_after: oneMonthAgoStr } }
            ]
          }
        ]
      },
      sorts: [{ property: '引渡し', direction: 'ascending' }],
      page_size: 100
    };
    if (startCursor) payload.start_cursor = startCursor;

    const res = notionPost_('/databases/' + DB_ID.物件 + '/query', payload);
    res.results.forEach(function(page) {
      const title = page.properties['物件名']?.title?.[0]?.plain_text || '';
      if (title && title !== '原本') {
        const hikiwatashi = page.properties['引渡し']?.date?.start || '';
        if (hikiwatashi) {
          withDate.push({ id: page.id, name: title });
        } else {
          withoutDate.push({ id: page.id, name: title });
        }
      }
    });
    hasMore = res.has_more;
    startCursor = res.next_cursor;
  }
  // 引渡し日ありを先頭（引渡し昇順）、なしを末尾
  return withDate.concat(withoutDate);
}

/**
 * 業者一覧を取得
 */
function getVendors() {
  const allResults = [];
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const payload = {
      sorts: [{ property: '業者名', direction: 'ascending' }],
      page_size: 100
    };
    if (startCursor) payload.start_cursor = startCursor;

    const res = notionPost_('/databases/' + DB_ID.業者 + '/query', payload);
    res.results.forEach(function(page) {
      const name = page.properties['業者名']?.title?.[0]?.plain_text || '';
      const type = page.properties['業種']?.select?.name || '';
      if (name) {
        allResults.push({ id: page.id, name: name, type: type });
      }
    });
    hasMore = res.has_more;
    startCursor = res.next_cursor;
  }
  return allResults;
}

/**
 * 担当者一覧を取得（所属業者IDも含む）
 */
function getContacts() {
  const allResults = [];
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const payload = {
      sorts: [{ property: '担当者名', direction: 'ascending' }],
      page_size: 100
    };
    if (startCursor) payload.start_cursor = startCursor;

    const res = notionPost_('/databases/' + DB_ID.担当者 + '/query', payload);
    res.results.forEach(function(page) {
      const name = page.properties['担当者名']?.title?.[0]?.plain_text || '';
      const vendorIds = (page.properties['所属業者']?.relation || []).map(function(r) { return r.id; });
      if (name) {
        allResults.push({ id: page.id, name: name, vendorIds: vendorIds });
      }
    });
    hasMore = res.has_more;
    startCursor = res.next_cursor;
  }
  return allResults;
}

/**
 * フォーム表示用の全データをまとめて取得
 */
function getAllFormData() {
  return {
    properties: getProperties(),
    vendors: getVendors(),
    contacts: getContacts()
  };
}

// ---------- 新規登録 ----------

/**
 * 新しい業者を登録
 */
function addNewVendor(name, type) {
  const properties = {
    '業者名': { title: [{ text: { content: name } }] }
  };
  if (type) {
    properties['業種'] = { select: { name: type } };
  }
  const res = notionPost_('/pages', {
    parent: { database_id: DB_ID.業者 },
    properties: properties
  });
  return { id: res.id, name: name, type: type || '' };
}

/**
 * 新しい担当者を登録
 */
function addNewContact(name, vendorId) {
  const properties = {
    '担当者名': { title: [{ text: { content: name } }] }
  };
  if (vendorId) {
    properties['所属業者'] = { relation: [{ id: vendorId }] };
  }
  const res = notionPost_('/pages', {
    parent: { database_id: DB_ID.担当者 },
    properties: properties
  });
  return { id: res.id, name: name, vendorIds: vendorId ? [vendorId] : [] };
}

/**
 * 支払い管理の既存データを検索（物件・業者で絞り込み）
 */
function searchPayments(propertyId, vendorId) {
  const allResults = [];
  let hasMore = true;
  let startCursor = undefined;

  const andFilters = [];
  if (propertyId) {
    andFilters.push({ property: '物件名', relation: { contains: propertyId } });
  }
  if (vendorId) {
    andFilters.push({ property: '業者', relation: { contains: vendorId } });
  }

  while (hasMore) {
    const payload = {
      sorts: [{ property: '登録日', direction: 'descending' }],
      page_size: 100
    };
    if (andFilters.length > 0) {
      payload.filter = andFilters.length === 1 ? andFilters[0] : { and: andFilters };
    }
    if (startCursor) payload.start_cursor = startCursor;

    const res = notionPost_('/databases/' + DB_ID.支払い管理 + '/query', payload);
    res.results.forEach(function(page) {
      const title = page.properties['タイトル']?.title?.[0]?.plain_text || '';
      const estimate = page.properties['見積もり額']?.number || 0;
      const payment = page.properties['支払い額']?.number || 0;
      const memo = page.properties['備考']?.rich_text?.[0]?.plain_text || '';
      const date = page.properties['登録日']?.date?.start || '';
      allResults.push({
        id: page.id,
        title: title,
        estimate: estimate,
        payment: payment,
        memo: memo,
        date: date
      });
    });
    hasMore = res.has_more;
    startCursor = res.next_cursor;
  }
  return allResults;
}

/**
 * 既存の支払い情報を更新
 */
function updatePayment(pageId, data) {
  const properties = {
    '見積もり額': { number: Number(data.estimate) || 0 },
    '支払い額': { number: Number(data.payment) || 0 }
  };
  if (data.memo !== undefined) {
    properties['備考'] = { rich_text: [{ text: { content: data.memo || '' } }] };
  }
  notionPatch_('/pages/' + pageId, { properties: properties });
  return { success: true, pageId: pageId };
}

/**
 * 一覧表示用: 全支払いデータを取得（物件名・業者名・担当者名を解決済み）
 */
function getAllPaymentsForList() {
  const formData = getAllFormData();
  const propMap = {};
  formData.properties.forEach(function(p) { propMap[p.id] = p.name; });
  const vendorMap = {};
  formData.vendors.forEach(function(v) { vendorMap[v.id] = v.name; });
  const contactMap = {};
  formData.contacts.forEach(function(c) { contactMap[c.id] = c.name; });

  const allResults = [];
  let hasMore = true;
  let startCursor = undefined;

  while (hasMore) {
    const payload = {
      sorts: [{ property: '登録日', direction: 'descending' }],
      page_size: 100
    };
    if (startCursor) payload.start_cursor = startCursor;

    const res = notionPost_('/databases/' + DB_ID.支払い管理 + '/query', payload);
    res.results.forEach(function(page) {
      const title = page.properties['タイトル']?.title?.[0]?.plain_text || '';
      const estimate = page.properties['見積もり額']?.number || 0;
      const payment = page.properties['支払い額']?.number || 0;
      const date = page.properties['登録日']?.date?.start || '';
      const propId = (page.properties['物件名']?.relation || [])[0]?.id || '';
      const vendorId = (page.properties['業者']?.relation || [])[0]?.id || '';
      const contactId = (page.properties['担当者']?.relation || [])[0]?.id || '';

      allResults.push({
        id: page.id,
        title: title,
        estimate: estimate,
        payment: payment,
        date: date,
        propertyId: propId,
        propertyName: propMap[propId] || '',
        vendorId: vendorId,
        vendorName: vendorMap[vendorId] || '',
        contactId: contactId,
        contactName: contactMap[contactId] || ''
      });
    });
    hasMore = res.has_more;
    startCursor = res.next_cursor;
  }
  return allResults;
}

/**
 * 支払い情報を登録
 */
function submitPayment(data) {
  const titleText = (data.propertyName || '') + '_' + (data.vendorName || '');

  const properties = {
    'タイトル': { title: [{ text: { content: titleText } }] },
    '物件名': { relation: [{ id: data.propertyId }] },
    '業者': { relation: [{ id: data.vendorId }] },
    '見積もり額': { number: Number(data.estimate) || 0 },
    '支払い額': { number: Number(data.payment) || 0 },
    '登録日': { date: { start: new Date().toISOString().split('T')[0] } }
  };

  if (data.contactId) {
    properties['担当者'] = { relation: [{ id: data.contactId }] };
  }
  if (data.memo) {
    properties['備考'] = { rich_text: [{ text: { content: data.memo } }] };
  }

  const res = notionPost_('/pages', {
    parent: { database_id: DB_ID.支払い管理 },
    properties: properties
  });

  return { success: true, pageId: res.id };
}
