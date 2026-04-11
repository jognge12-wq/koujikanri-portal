// ═══════════════════════════════════════════════════════════════
//  施工管理ポータル — タスク管理 GAS
//  【データの流れ】
//    Notion（担当物件一覧）→ 物件名リストだけ取得（読み取りのみ）
//    タスク管理はスプレッドシートで完結・Notionとのタスク連携なし
// ═══════════════════════════════════════════════════════════════

// ── ① Notion API キー（デプロイ済みと同じ値）──────────────────
const NOTION_API_KEY      = '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';
const NOTION_PROPS_DB_ID  = '2f56ad84622180a9891bef7e5514fa78'; // 担当物件一覧DB

// ── ② Notion「担当物件一覧」DBのプロパティ名 ─────────────────
const NF = {
  NAME     : '物件名',   // タイトルプロパティ（必須）
  PROGRESS : '進捗',     // 進行状況（select または multi_select）
  CITY     : '市町村',   // 市町村（テキスト or select）
  START    : '本体着工', // 本体着工日（date）
  FRAME    : '建て方',   // 建て方日（date）
  HANDOVER : '引渡し',   // 引渡し日（date）
};

// ── ③ 除外する物件名（テンプレート・原本など） ────────────────
const EXCLUDE_NAMES = ['原本', '原本(コピー)'];
// ────────────────────────────────────────────────────────────

// ── スプレッドシートID（setup()実行後に自動保存）─────────────
const PROP_SS_ID = 'SPREADSHEET_ID';

// ── シート名 ─────────────────────────────────────────────────
const SH = {
  TASKS   : 'タスク一覧',
  MASTER  : 'マスタータスク',
  HISTORY : '変更履歴',
  CONFIG  : '設定',
};

// ── タスク一覧 列番号 ─────────────────────────────────────────
const C = {
  ID       : 1,  // A タスクID (T-0001)
  PROPERTY : 2,  // B 物件名
  NAME     : 3,  // C タスク名
  PHASE    : 4,  // D 工事進捗
  PRIORITY : 5,  // E 優先 (TRUE/FALSE)
  DONE     : 6,  // F 完了 (TRUE/FALSE)
  DUE      : 7,  // G 期日
  ORDER    : 8,  // H 並び順
  UPDATED  : 9,  // I 更新日時
};

// ── マスタータスク 列番号 ─────────────────────────────────────
const MC = {
  ID      : 1,  // A マスターID (M-001)
  NAME    : 2,  // B タスク名
  PHASE   : 3,  // C 工事進捗
  PRIORITY: 4,  // D 優先 (TRUE/FALSE)
  ORDER   : 5,  // E 並び順
  ACTIVE  : 6,  // F 有効 (TRUE/FALSE)
};


// ═══════════════════════════════════════════════════════════════
//  【STEP 1】 初回セットアップ — GASエディタで1回だけ実行
// ═══════════════════════════════════════════════════════════════
function setup() {
  Logger.log('=== セットアップ開始 ===');

  const ss = SpreadsheetApp.create('施工管理_タスク管理');
  const ssId = ss.getId();
  PropertiesService.getScriptProperties().setProperty(PROP_SS_ID, ssId);

  _setupMasterSheet(ss);
  _setupTaskSheet(ss);
  _setupHistorySheet(ss);
  _setupConfigSheet(ss);

  // デフォルトシートを削除
  const def = ss.getSheetByName('シート1');
  if (def) ss.deleteSheet(def);

  // アクティブシートをマスタータスクに
  ss.setActiveSheet(ss.getSheetByName(SH.MASTER));

  Logger.log('=== セットアップ完了 ===');
  Logger.log('URL: ' + ss.getUrl());
  Logger.log('次の手順: マスタータスクシートにタスクを登録してください');
}

function _setupMasterSheet(ss) {
  const sh = ss.insertSheet(SH.MASTER);
  const headers = ['マスターID', 'タスク名', '工事進捗', '優先', '並び順', '有効'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#006837').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  sh.setColumnWidth(1, 80);   // ID
  sh.setColumnWidth(2, 220);  // タスク名
  sh.setColumnWidth(3, 110);  // 工事進捗
  sh.setColumnWidth(4, 60);   // 優先
  sh.setColumnWidth(5, 70);   // 並び順
  sh.setColumnWidth(6, 60);   // 有効

  // マスタータスク177件（工事進捗フェーズ順・工程内順番順）
  const samples = [
    ['M-001', '配置基準の境界確認：丁張りが可能か', '現場FB', false, 1, true],
    ['M-002', '引込位置の確認：電柱・弱電位置の確認', '現場FB', false, 2, true],
    ['M-003', '電線保護カバー・敷鉄板・道路使用の有無の確認', '現場FB', false, 3, true],
    ['M-004', '仮設計画図・生産補正シートの作成', '現場FB', false, 4, true],
    ['M-005', '資料提出', '現場FB', false, 5, true],
    ['M-006', '図面チェック※Documents使用', '図面FB', false, 1, true],
    ['M-007', '図面チェック完了をLINEで報告', '図面FB', false, 2, true],
    ['M-008', 'テント設営の依頼', '地鎮祭', false, 1, true],
    ['M-009', '奉献酒の用意', '地鎮祭', false, 2, true],
    ['M-010', '鎮め物を業者へ渡す', '地鎮祭', false, 3, true],
    ['M-011', '仮設計画図の作成・手配', '地鎮祭', false, 4, true],
    ['M-012', '地縄張りの依頼', '地縄立会い', false, 1, true],
    ['M-013', '近隣挨拶分の作成', '地縄立会い', false, 2, true],
    ['M-014', '近隣挨拶タオルの用意', '地縄立会い', false, 3, true],
    ['M-015', '地縄張り用配置図を設計担当へLINE確認', '地縄立会い', false, 4, true],
    ['M-016', '長期優良住宅の申請日の確認※地盤改良工事前の申請が必須', '地縄立会い', false, 5, true],
    ['M-017', '立会いノートの作成', '地縄立会い', false, 6, true],
    ['M-018', '施工計画説明の日時決定', '地縄立会い', false, 7, true],
    ['M-019', 'ショールームの予約', '地縄立会い', false, 8, true],
    ['M-020', '引継ぎ会日時の依頼', '地縄立会い', false, 9, true],
    ['M-021', '立会い完了の報告、施工計画説明日時のカレンダー共有', '地縄立会い', false, 10, true],
    ['M-022', '実績入力：地縄立会い', '地縄立会い', false, 11, true],
    ['M-023', '立会いノート・遣り方検査日時依頼の伝言板アップ', '地縄立会い', false, 12, true],
    ['M-024', '工程表の作成', '生産移管', false, 1, true],
    ['M-025', '引継書の印刷', '生産移管', false, 2, true],
    ['M-026', '電子発注明細の印刷', '生産移管', false, 3, true],
    ['M-027', '仮設計画図兼作業指示書の作成', '生産移管', false, 4, true],
    ['M-028', 'NACCS業者登録・メンテ', '生産移管', false, 5, true],
    ['M-029', '杭ナビデータ有：業者登録※栃井建設のみ', '生産移管', false, 6, true],
    ['M-030', '電子先行発注：プレカット・軒天', '生産移管', false, 7, true],
    ['M-031', '長期優良住宅認可予定日の確認', '生産移管', false, 8, true],
    ['M-032', '生産補正シート作成・予算入力依頼', '生産移管', false, 9, true],
    ['M-033', '図面マーキング', '生産移管', false, 10, true],
    ['M-034', '引継ぎ前チェック：引継ぎ書・提案工事内容・エプコ配管経路図・スミテン図', '生産移管', false, 11, true],
    ['M-035', '設計指示価格の見積もり期日確認', '生産移管', false, 12, true],
    ['M-036', '引継ぎ会資料の提出', '生産移管', false, 13, true],
    ['M-037', '実績入力：引継ぎ', '生産移管', false, 14, true],
    ['M-038', '立会いノートの作成', '施工計画説明', false, 1, true],
    ['M-039', 'お客様への質疑まとめ', '施工計画説明', false, 2, true],
    ['M-040', '着工合意書の内容確認', '施工計画説明', false, 3, true],
    ['M-041', 'お客様配布物：工程表・製本図面・工事案内ファイル', '施工計画説明', false, 4, true],
    ['M-042', '構造立会い日の決定', '施工計画説明', false, 5, true],
    ['M-043', '手形式の有無の確認', '施工計画説明', false, 6, true],
    ['M-044', '棟札ご持参の案内', '施工計画説明', false, 7, true],
    ['M-045', 'TV・ネットの早期申込みの説明', '施工計画説明', false, 8, true],
    ['M-046', '実績入力：施工計画説明', '施工計画説明', false, 9, true],
    ['M-047', '立会いノートを伝言板にアップ：契約電気容量・メーター名義名を共有', '施工計画説明', false, 10, true],
    ['M-048', '朱書き図面を伝言板にアップ', '施工計画説明', false, 11, true],
    ['M-049', '基礎工事計画書の提出', '遣り方検査', false, 1, true],
    ['M-050', '遣り方シールの確認・記載', '遣り方検査', false, 2, true],
    ['M-051', '引継ぎ会資料の持参：工務店担当にサインもらう', '遣り方検査', false, 3, true],
    ['M-052', '安全日誌の記載', '遣り方検査', false, 4, true],
    ['M-053', 'NACCSへ着工写真をアップ', '遣り方検査', false, 5, true],
    ['M-054', '引継ぎ会資料の提出', '遣り方検査', false, 6, true],
    ['M-055', '実績入力：本体着工', '遣り方検査', false, 7, true],
    ['M-056', '実績入力：着工前ミーティング', '遣り方検査', false, 8, true],
    ['M-057', '建性①配筋検査の申込', '遣り方検査', false, 9, true],
    ['M-058', 'CON打設の近隣挨拶分の作成・持参', '配筋検査', false, 1, true],
    ['M-059', '長期優良住宅が認可済かの確認※ベースCON打設までが必須', '配筋検査', false, 2, true],
    ['M-060', '島基礎の計測', '配筋検査', false, 3, true],
    ['M-061', 'スリーブ位置の計測・記録', '配筋検査', false, 4, true],
    ['M-062', 'スペーサーブロック・シート重ね長さの計測・記録', '配筋検査', false, 5, true],
    ['M-063', 'コーナー・隅角部補強金・主筋継手位置の記録', '配筋検査', false, 6, true],
    ['M-064', '性能評価シールを確認看板に貼る', '配筋検査', false, 7, true],
    ['M-065', '近隣挨拶：CON打設', '配筋検査', false, 8, true],
    ['M-066', '安全日誌の記載', '配筋検査', false, 9, true],
    ['M-067', 'iPadで性能評価書類の記載・報告※生産補助へ', '配筋検査', false, 10, true],
    ['M-068', '図面の提出', '配筋検査', false, 11, true],
    ['M-069', '実績入力：配筋検査', '配筋検査', false, 12, true],
    ['M-070', '島基礎の計測・記録', '型枠検査', false, 1, true],
    ['M-071', 'ボルト類の位置・レベルの計測・記録', '型枠検査', false, 2, true],
    ['M-072', 'スラブ厚の計測・記録', '型枠検査', false, 3, true],
    ['M-073', '被り厚検査棒による被り厚の確認', '型枠検査', false, 4, true],
    ['M-074', '安全日誌の記載', '型枠検査', false, 5, true],
    ['M-075', '図面提出', '型枠検査', false, 6, true],
    ['M-076', '建て方の近隣挨拶分の作成・持参', '基礎検査', false, 1, true],
    ['M-077', '防蟻パイプの本数計測・記録', '基礎検査', false, 2, true],
    ['M-078', '近隣挨拶：建て方', '基礎検査', false, 3, true],
    ['M-079', '安全日誌の記載', '基礎検査', false, 4, true],
    ['M-080', '図面提出', '基礎検査', false, 5, true],
    ['M-081', '実績入力：基礎検査', '基礎検査', false, 6, true],
    ['M-082', '伝言板で構造検査・構造立会い日時の共有', '基礎検査', false, 7, true],
    ['M-083', '基礎精算の依頼', '基礎検査', false, 8, true],
    ['M-084', '建て方の予定の社内報告・カレンダー登録', '基礎検査', false, 9, true],
    ['M-085', '建て方前の入金確認※支払い条件による', '建て方', false, 1, true],
    ['M-086', '建て方人数の確認、施主報告', '建て方', false, 2, true],
    ['M-087', '建て方計画書の確認・承認', '建て方', false, 3, true],
    ['M-088', '全景写真撮影、次長へLINE報告', '建て方', false, 4, true],
    ['M-089', '安全日誌の確認・記載', '建て方', false, 5, true],
    ['M-090', '足場点検実施→是正があればLINEグループで指示', '建て方', false, 6, true],
    ['M-091', '実績入力：建て方開始', '建て方', false, 7, true],
    ['M-092', '実績入力：野地板完了', '建て方', false, 8, true],
    ['M-093', '建て方完了の報告→施主・社内', '建て方', false, 9, true],
    ['M-094', '建性②：構造検査日時の確認※申込みは生産事務が行う', '建て方', false, 10, true],
    ['M-095', '階高の計測・記録', '構造検査', false, 1, true],
    ['M-096', '構造材種の確認', '構造検査', false, 2, true],
    ['M-097', 'センサーライトを玄関先に取付', '構造検査', false, 3, true],
    ['M-098', '防蟻剤容器に危険物シールが貼ってあるか確認', '構造検査', false, 4, true],
    ['M-099', '各階消火器設置の確認', '構造検査', false, 5, true],
    ['M-100', '安全日誌の記載', '構造検査', false, 6, true],
    ['M-101', '図面提出', '構造検査', false, 7, true],
    ['M-102', '基礎精算・発注', '構造検査', false, 8, true],
    ['M-103', 'お客様に立会い日程・棟札・手形の確認', '構造立会い', false, 1, true],
    ['M-104', '立会いノートの作成', '構造立会い', false, 2, true],
    ['M-105', '木完立会い日時の決定', '構造立会い', false, 3, true],
    ['M-106', '支給品を木完立会い持参の案内', '構造立会い', false, 4, true],
    ['M-107', '社内：立会い完了の報告', '構造立会い', false, 5, true],
    ['M-108', '伝言板で木完検査・木完立会い日時の共有', '構造立会い', false, 6, true],
    ['M-109', '「木完検査」NACCS送信ボタンのクリック', '構造立会い', false, 7, true],
    ['M-110', '実績入力：構造検査※生産事務代理', '雨仕舞い検査', false, 1, true],
    ['M-111', '電気配線・BOXの確認', '雨仕舞い検査', false, 2, true],
    ['M-112', '防火区画テープ貼りの確認・記録※平屋は該当無し', '雨仕舞い検査', false, 3, true],
    ['M-113', '安全日誌の記入', '雨仕舞い検査', false, 4, true],
    ['M-114', '図面提出', '雨仕舞い検査', false, 5, true],
    ['M-115', '実績入力：構造雨仕舞い', '雨仕舞い検査', false, 6, true],
    ['M-116', '建性③：断熱検査日時の確認※申込みは生産事務が行う', '雨仕舞い検査', false, 7, true],
    ['M-117', '吹付の近隣挨拶分の作成', '雨仕舞い検査', false, 8, true],
    ['M-118', '近隣挨拶：吹付作業', '雨仕舞い検査', false, 9, true],
    ['M-119', '実績入力：断熱検査※生産事務代理', '木完検査', false, 1, true],
    ['M-120', '実績入力：左官防水検査', '木完検査', false, 2, true],
    ['M-121', '実績入力：足場解体', '木完検査', false, 3, true],
    ['M-122', '社内：追加変更の覚書の有無の確認', '木完検査', false, 4, true],
    ['M-123', 'BCのPBビスピッチ確認・記録', '木完検査', false, 5, true],
    ['M-124', '天井高さの測定', '木完検査', false, 6, true],
    ['M-125', 'クロスサンプルの持参・貼付け', '木完検査', false, 7, true],
    ['M-126', '安全日誌の記載', '木完検査', false, 8, true],
    ['M-127', '図面提出', '木完検査', false, 9, true],
    ['M-128', '実績入力：木完検査', '木完検査', false, 10, true],
    ['M-129', '建性④：完成検査の申込み', '木完検査', false, 11, true],
    ['M-130', '完了検査の申込み→申込み用紙・省令準耐火チェックシートを生産事務へ提出', '木完検査', false, 12, true],
    ['M-131', '伝言板で竣工検査日時の共有', '木完検査', false, 13, true],
    ['M-132', '仮設撤去の依頼', '木完検査', false, 14, true],
    ['M-133', 'お客様に立会い日時の確認', '木完立会い', false, 1, true],
    ['M-134', '立会いノートの作成', '木完立会い', false, 2, true],
    ['M-135', '施主支給品の受取り', '木完立会い', false, 3, true],
    ['M-136', 'ライフライン名義変更の案内', '木完立会い', false, 4, true],
    ['M-137', '竣工立会い・引渡しスケジュールの確定', '木完立会い', false, 5, true],
    ['M-138', '最終金額確認書についての案内', '木完立会い', false, 6, true],
    ['M-139', '社内：立会い完了の報告・引渡し日のカレンダー登録', '木完立会い', false, 7, true],
    ['M-140', '引渡し申請をLINEで依頼', '木完立会い', false, 8, true],
    ['M-141', '竣工立会い・引渡しの工程メンテ', '木完立会い', false, 9, true],
    ['M-142', '「引渡し」NACCS送信ボタンのクリック', '木完立会い', false, 10, true],
    ['M-143', '伝言板で引渡しまでのスケジュール共有', '木完立会い', false, 11, true],
    ['M-144', '取説ファイリング', '竣工検査', false, 1, true],
    ['M-145', '玄関土間の段差の測定・記録', '竣工検査', false, 2, true],
    ['M-146', 'センサー照明の設定番号確認', '竣工検査', false, 3, true],
    ['M-147', '防犯カメラの登録確認・センサー消音', '竣工検査', false, 4, true],
    ['M-148', '給気フィルターが入っているか、内部清掃', '竣工検査', false, 5, true],
    ['M-149', 'UB立上りスリーブの断熱施工・点検口テープ貼りの確認', '竣工検査', false, 6, true],
    ['M-150', '熱源機の2重ナット固定・アース接続の確認', '竣工検査', false, 7, true],
    ['M-151', 'エアコンスリーブ断熱材が有るか、内部清掃状況の確認', '竣工検査', false, 8, true],
    ['M-152', '給水・給湯ヘッダーの2箇所以上の固定確認', '竣工検査', false, 9, true],
    ['M-153', '小屋裏の断熱施工確認', '竣工検査', false, 10, true],
    ['M-154', '小屋裏の太陽光配線確認、LINEグループへ報告', '竣工検査', false, 11, true],
    ['M-155', '安全日誌の記載', '竣工検査', false, 12, true],
    ['M-156', '手直し資料のまとめ、伝言板アップ', '竣工検査', false, 13, true],
    ['M-157', '図面の提出', '竣工検査', false, 14, true],
    ['M-158', '実績入力：竣工検査・4回目検査', '竣工検査', false, 15, true],
    ['M-159', '社内：完了検査・性能評価検査の完了報告', '竣工検査', false, 16, true],
    ['M-160', 'お客様に立会い日時の確認', '竣工立会い', false, 1, true],
    ['M-161', '立会いノートの作成', '竣工立会い', false, 2, true],
    ['M-162', '災害用スリーブの説明', '竣工立会い', false, 3, true],
    ['M-163', '汚水桝（地域による）・雨水桝の清掃の説明', '竣工立会い', false, 4, true],
    ['M-164', '三協立山アルミの電池錠：取説を送る、ユーザー登録の依頼', '竣工立会い', false, 5, true],
    ['M-165', '引渡しの実印持参の案内', '竣工立会い', false, 6, true],
    ['M-166', '手直し資料のまとめ、伝言板アップ', '竣工立会い', false, 7, true],
    ['M-167', '実績入力：竣工立会い', '竣工立会い', false, 8, true],
    ['M-168', '立会い完了の社内報告', '竣工立会い', false, 9, true],
    ['M-169', 'お客様に引渡し日時の確認', '引渡し', false, 1, true],
    ['M-170', '立会いノートの作成', '引渡し', false, 2, true],
    ['M-171', '引渡し書類・記念時計品・鍵の持参', '引渡し', false, 3, true],
    ['M-172', '電気錠・電池錠の設定・登録', '引渡し', false, 4, true],
    ['M-173', '入居後訪問（LINE聞き取り）の案内', '引渡し', false, 5, true],
    ['M-174', '引渡受書・外観写真をLINEで送る', '引渡し', false, 6, true],
    ['M-175', '完成写真をLINEグループに送る', '引渡し', false, 7, true],
    ['M-176', '引渡し書類の提出（受書・AM引継書・長期優良認定書・図面）', '引渡し', false, 8, true],
    ['M-177', '実績入力：入居予定日', '引渡し', false, 9, true],
  ];
  sh.getRange(2, 1, samples.length, 6).setValues(samples);
}

function _setupTaskSheet(ss) {
  const sh = ss.insertSheet(SH.TASKS);
  const headers = ['タスクID', '物件名', 'タスク名', '工事進捗', '優先', '完了', '期日', '並び順', '更新日時'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#006837').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);

  sh.setColumnWidth(1, 80);   // ID
  sh.setColumnWidth(2, 130);  // 物件名
  sh.setColumnWidth(3, 220);  // タスク名
  sh.setColumnWidth(4, 110);  // 工事進捗
  sh.setColumnWidth(5, 60);   // 優先
  sh.setColumnWidth(6, 60);   // 完了
  sh.setColumnWidth(7, 90);   // 期日
  sh.setColumnWidth(8, 70);   // 並び順
  sh.setColumnWidth(9, 140);  // 更新日時
}

function _setupHistorySheet(ss) {
  const sh = ss.insertSheet(SH.HISTORY);
  const headers = ['変更日時', 'タスクID', 'タスク名', '物件名', '変更項目', '変更前', '変更後'];
  sh.getRange(1, 1, 1, headers.length).setValues([headers])
    .setBackground('#2c3038').setFontColor('#ffffff').setFontWeight('bold');
  sh.setFrozenRows(1);
}

function _setupConfigSheet(ss) {
  const sh = ss.insertSheet(SH.CONFIG);
  sh.getRange('A1:B1').setValues([['設定項目', '値']])
    .setBackground('#2c3038').setFontColor('#ffffff').setFontWeight('bold');
  sh.getRange('A2').setValue('最終物件同期');
  sh.getRange('A3').setValue('タスク総数');

  // 工事進捗マスタ
  sh.getRange('D1').setValue('工事進捗マスタ').setFontWeight('bold');
  const phases = [
    '現場FB','図面FB','地鎮祭','地縄立会い','生産移管','施工計画説明',
    '遣り方検査','配筋検査','型枠検査','基礎検査','建て方','構造検査',
    '構造立会い','雨仕舞い検査','木完検査','木完立会い','竣工検査','竣工立会い','引渡し',
  ];
  phases.forEach((p, i) => sh.getRange(i + 2, 4).setValue(p));
}


// ═══════════════════════════════════════════════════════════════
//  【STEP 2】 物件にタスクを一括生成
//  新しい物件を担当するとき、マスタータスクから一括でタスクを作成
//  GASエディタで実行 or APIで呼び出し
// ═══════════════════════════════════════════════════════════════
function initPropertyTasks(propertyName) {
  if (!propertyName) throw new Error('物件名を指定してください');

  const ss = _getSS();
  const masterSh = ss.getSheetByName(SH.MASTER);
  const taskSh   = ss.getSheetByName(SH.TASKS);

  // マスタータスクを取得（有効なもののみ）
  const masterData = masterSh.getDataRange().getValues().slice(1)
    .filter(r => r[MC.ID - 1] && r[MC.ACTIVE - 1] === true);

  if (masterData.length === 0) {
    throw new Error('有効なマスタータスクがありません');
  }

  // 既存タスクの最大番号を取得
  const taskData = taskSh.getDataRange().getValues().slice(1).filter(r => r[C.ID - 1]);
  const maxNum = taskData.reduce((max, r) => {
    const n = parseInt(String(r[C.ID - 1]).replace('T-', ''), 10);
    return isNaN(n) ? max : Math.max(max, n);
  }, 0);

  // マスターからタスクを生成
  const newRows = masterData.map((m, i) => {
    const taskId = 'T-' + String(maxNum + i + 1).padStart(4, '0');
    return [
      taskId,
      propertyName,
      m[MC.NAME - 1],
      m[MC.PHASE - 1],
      m[MC.PRIORITY - 1],
      false,            // 完了: 初期値false
      '',               // 期日: 空
      m[MC.ORDER - 1],
      new Date(),
    ];
  });

  const startRow = taskSh.getLastRow() + 1;
  taskSh.getRange(startRow, 1, newRows.length, 9).setValues(newRows);

  // 設定シートに記録
  const configSh = ss.getSheetByName(SH.CONFIG);
  configSh.getRange('B3').setValue(taskSh.getLastRow() - 1);

  Logger.log(`${propertyName}: ${newRows.length}件のタスクを生成しました`);
  _writeHistory('', '', propertyName, '物件タスク初期化', '', `${newRows.length}件生成`);

  return newRows.length;
}


// ═══════════════════════════════════════════════════════════════
//  WebApp — ダッシュボードとのAPI連携
//  デプロイ → ウェブアプリとして公開 してURLを取得
// ═══════════════════════════════════════════════════════════════
function doGet(e) {
  const mode = e.parameter.mode || '';
  const output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    let result;
    switch (mode) {
      case 'getTasks'         : result = _apiGetTasks(e.parameter);       break;
      case 'getProperties'    : result = _apiGetProperties();              break;
      case 'getMasterTasks'   : result = _apiGetMasterTasks(e.parameter); break;
      case 'updateTask'       : result = _apiUpdateTask(e.parameter);     break;
      case 'addTask'          : result = _apiAddTask(e.parameter);        break;
      case 'addMasterTask'    : result = _apiAddMasterTask(e.parameter);  break;
      case 'updateMasterTask' : result = _apiUpdateMasterTask(e.parameter); break;
      case 'moveTask'         : result = _apiMoveTask(e.parameter);       break;
      case 'reorderTask'      : result = _apiReorderTask(e.parameter);       break;
      case 'initProperty'     : result = _apiInitProperty(e.parameter);     break;
      case 'getHistory'       : result = _apiGetHistory(e.parameter);       break;
      case 'deleteMasterTask'     : result = _apiDeleteMasterTask(e.parameter);     break;
      case 'getPhases'            : result = _apiGetPhases();                       break;
      case 'addPhase'             : result = _apiAddPhase(e.parameter);             break;
      case 'renamePhase'          : result = _apiRenamePhase(e.parameter);          break;
      case 'updatePropertyCheck'  : result = _apiUpdatePropertyCheck(e.parameter); break;
      default: result = { error: 'unknown mode: ' + mode };
    }
    output.setContent(JSON.stringify({ ok: true, data: result }));
  } catch (err) {
    output.setContent(JSON.stringify({ ok: false, error: err.message }));
  }

  return output;
}


// ── API実装 ──────────────────────────────────────────────────

// タスク一覧取得
// ?mode=getTasks
// ?mode=getTasks&property=大山S・M様
// ?mode=getTasks&done=false
function _apiGetTasks(params) {
  const sh = _getSS().getSheetByName(SH.TASKS);
  const data = sh.getDataRange().getValues();

  let tasks = data.slice(1).map(row => ({
    id       : row[C.ID - 1],
    property : row[C.PROPERTY - 1],
    name     : row[C.NAME - 1],
    phase    : row[C.PHASE - 1],
    priority : row[C.PRIORITY - 1] === true || row[C.PRIORITY - 1] === 'TRUE',
    done     : row[C.DONE - 1] === true || row[C.DONE - 1] === 'TRUE',
    due      : row[C.DUE - 1]
                 ? Utilities.formatDate(new Date(row[C.DUE - 1]), 'Asia/Tokyo', 'yyyy-MM-dd')
                 : '',
    order    : Number(row[C.ORDER - 1]) || 0,
    updated  : row[C.UPDATED - 1] ? String(row[C.UPDATED - 1]) : '',
  })).filter(t => t.id);

  if (params.property) tasks = tasks.filter(t => t.property === params.property);
  if (params.done === 'false') tasks = tasks.filter(t => !t.done);
  if (params.done === 'true')  tasks = tasks.filter(t => t.done);
  if (params.priority === 'true') tasks = tasks.filter(t => t.priority);

  // 並び順でソート
  tasks.sort((a, b) => a.order - b.order);
  return tasks;
}

// 物件一覧取得（Notionから）
// ?mode=getProperties
function _apiGetProperties() {
  return _getNotionProperties();
}

// マスタータスク一覧取得
// ?mode=getMasterTasks
// ?mode=getMasterTasks&phase=地鎮祭
function _apiGetMasterTasks(params) {
  const sh = _getSS().getSheetByName(SH.MASTER);
  const data = sh.getDataRange().getValues();

  let masters = data.slice(1).map(row => ({
    id      : row[MC.ID - 1],
    name    : row[MC.NAME - 1],
    phase   : row[MC.PHASE - 1],
    priority: row[MC.PRIORITY - 1] === true || row[MC.PRIORITY - 1] === 'TRUE',
    order   : Number(row[MC.ORDER - 1]) || 0,
    active  : row[MC.ACTIVE - 1] === true || row[MC.ACTIVE - 1] === 'TRUE',
  })).filter(m => m.id);

  if (params.phase) masters = masters.filter(m => m.phase === params.phase);
  if (params.active !== 'false') masters = masters.filter(m => m.active);

  masters.sort((a, b) => a.order - b.order);
  return masters;
}

// タスク更新
// ?mode=updateTask&id=T-0001&field=done&value=true
// ?mode=updateTask&id=T-0001&field=name&value=新しいタスク名
// ?mode=updateTask&id=T-0001&field=due&value=2026-05-01
// ?mode=updateTask&id=T-0001&field=priority&value=true
function _apiUpdateTask(params) {
  const { id, field, value } = params;
  if (!id || !field) throw new Error('id と field は必須です');

  const sh = _getSS().getSheetByName(SH.TASKS);
  const data = sh.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[C.ID - 1] === id);
  if (rowIdx < 0) throw new Error('タスクが見つかりません: ' + id);

  const shRow = rowIdx + 1;
  let colNum, newValue, oldValue;

  switch (field) {
    case 'done':
      colNum   = C.DONE;
      oldValue = data[rowIdx][C.DONE - 1];
      newValue = value === 'true' || value === true;
      break;
    case 'priority':
      colNum   = C.PRIORITY;
      oldValue = data[rowIdx][C.PRIORITY - 1];
      newValue = value === 'true' || value === true;
      break;
    case 'name':
      colNum   = C.NAME;
      oldValue = data[rowIdx][C.NAME - 1];
      newValue = String(value);
      break;
    case 'phase':
      colNum   = C.PHASE;
      oldValue = data[rowIdx][C.PHASE - 1];
      newValue = String(value);
      break;
    case 'due':
      colNum   = C.DUE;
      oldValue = data[rowIdx][C.DUE - 1];
      newValue = value ? new Date(value) : '';
      break;
    case 'order':
      colNum   = C.ORDER;
      oldValue = data[rowIdx][C.ORDER - 1];
      newValue = Number(value);
      break;
    default:
      throw new Error('不明なフィールド: ' + field);
  }

  sh.getRange(shRow, colNum).setValue(newValue);
  sh.getRange(shRow, C.UPDATED).setValue(new Date());

  _writeHistory(
    id,
    data[rowIdx][C.NAME - 1],
    data[rowIdx][C.PROPERTY - 1],
    field,
    oldValue,
    newValue
  );

  return { updated: true, id, field, newValue };
}

// タスク追加（物件への個別追加）
// ?mode=addTask&property=大山S・M様&name=タスク名&phase=地鎮祭&priority=true
function _apiAddTask(params) {
  const { property, name, phase } = params;
  if (!property || !name || !phase) throw new Error('property, name, phase は必須');

  const sh = _getSS().getSheetByName(SH.TASKS);
  const taskData = sh.getDataRange().getValues().slice(1).filter(r => r[C.ID - 1]);

  const maxNum = taskData.reduce((max, r) => {
    const n = parseInt(String(r[C.ID - 1]).replace('T-', ''), 10);
    return isNaN(n) ? max : Math.max(max, n);
  }, 0);
  const newId = 'T-' + String(maxNum + 1).padStart(4, '0');

  const samePhase = taskData.filter(r => r[C.PROPERTY-1] === property && r[C.PHASE-1] === phase);
  const maxOrder = samePhase.reduce((max, r) => Math.max(max, Number(r[C.ORDER-1]) || 0), 0);

  sh.appendRow([
    newId, property, name, phase,
    params.priority === 'true',
    false, '', maxOrder + 1, new Date(),
  ]);

  _writeHistory(newId, name, property, 'タスク追加', '', name);
  return { added: true, id: newId };
}

// マスタータスク追加
// ?mode=addMasterTask&name=タスク名&phase=地鎮祭&priority=true&order=4
function _apiAddMasterTask(params) {
  const { name, phase } = params;
  if (!name || !phase) throw new Error('name と phase は必須');

  const sh = _getSS().getSheetByName(SH.MASTER);
  const data = sh.getDataRange().getValues().slice(1).filter(r => r[MC.ID - 1]);
  const maxNum = data.reduce((max, r) => {
    const n = parseInt(String(r[MC.ID - 1]).replace('M-', ''), 10);
    return isNaN(n) ? max : Math.max(max, n);
  }, 0);
  const newId = 'M-' + String(maxNum + 1).padStart(3, '0');

  const order = params.order ? Number(params.order) : data.length + 1;
  sh.appendRow([newId, name, phase, params.priority === 'true', order, true]);

  _writeHistory('', name, 'マスター', 'マスタータスク追加', '', name);
  return { added: true, id: newId };
}

// マスタータスク更新
// ?mode=updateMasterTask&id=M-001&field=name&value=新しいタスク名
function _apiUpdateMasterTask(params) {
  const { id, field, value } = params;
  if (!id || !field) throw new Error('id と field は必須');

  const sh = _getSS().getSheetByName(SH.MASTER);
  const data = sh.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[MC.ID - 1] === id);
  if (rowIdx < 0) throw new Error('マスタータスクが見つかりません: ' + id);

  const colMap = { name: MC.NAME, phase: MC.PHASE, priority: MC.PRIORITY, order: MC.ORDER, active: MC.ACTIVE };
  const colNum = colMap[field];
  if (!colNum) throw new Error('不明なフィールド: ' + field);

  const oldValue = data[rowIdx][colNum - 1];
  let newValue = value;
  if (field === 'priority' || field === 'active') newValue = value === 'true' || value === true;
  if (field === 'order') newValue = Number(value);

  sh.getRange(rowIdx + 1, colNum).setValue(newValue);
  _writeHistory('', data[rowIdx][MC.NAME - 1], 'マスター', 'マスター.' + field, oldValue, newValue);

  return { updated: true, id, field, newValue };
}

// 工程間タスク移動
// ?mode=moveTask&id=T-0001&phase=基礎検査
function _apiMoveTask(params) {
  return _apiUpdateTask({ id: params.id, field: 'phase', value: params.phase });
}

// 並び順変更
// ?mode=reorderTask&id=T-0001&order=3
function _apiReorderTask(params) {
  return _apiUpdateTask({ id: params.id, field: 'order', value: params.order });
}

// 物件タスク一括生成
// ?mode=initProperty&property=新しい物件名
function _apiInitProperty(params) {
  const { property } = params;
  if (!property) throw new Error('property は必須');
  const count = initPropertyTasks(property);
  return { initialized: true, property, count };
}

// 変更履歴取得
// ?mode=getHistory&limit=50&property=大山S・M様
function _apiGetHistory(params) {
  const limit = parseInt(params.limit || '50', 10);
  const sh = _getSS().getSheetByName(SH.HISTORY);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  let history = sh.getRange(2, 1, lastRow - 1, 7).getValues()
    .filter(r => r[0])
    .map(r => ({
      timestamp : String(r[0]),
      taskId    : r[1],
      taskName  : r[2],
      property  : r[3],
      field     : r[4],
      before    : String(r[5]),
      after     : String(r[6]),
    }))
    .reverse();

  if (params.property)    history = history.filter(h => h.property === params.property);
  if (params.masterOnly === 'true') history = history.filter(h => h.property === 'マスター');
  return history.slice(0, limit);
}


// マスタータスク削除（完全削除）
// ?mode=deleteMasterTask&id=M-010
function _apiDeleteMasterTask(params) {
  const { id } = params;
  if (!id) throw new Error('id は必須');
  const sh = _getSS().getSheetByName(SH.MASTER);
  const data = sh.getDataRange().getValues();
  const rowIdx = data.findIndex((r, i) => i > 0 && r[MC.ID - 1] === id);
  if (rowIdx < 0) throw new Error('マスタータスクが見つかりません: ' + id);
  sh.deleteRow(rowIdx + 1);
  _writeHistory('', data[rowIdx][MC.NAME - 1], 'マスター', 'マスタータスク削除', data[rowIdx][MC.NAME - 1], '');
  return { deleted: true, id };
}

// 工程一覧取得（設定シートのD列から）
// ?mode=getPhases
function _apiGetPhases() {
  const sh = _getSS().getSheetByName(SH.CONFIG);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const phases = [];
  for (let i = 2; i <= lastRow; i++) {
    const v = sh.getRange(i, 4).getValue();
    if (v) phases.push(String(v));
  }
  return phases;
}

// 工程追加（設定シートのD列に追加）
// ?mode=addPhase&name=新しい工程名
function _apiAddPhase(params) {
  const { name } = params;
  if (!name) throw new Error('name は必須');
  const sh = _getSS().getSheetByName(SH.CONFIG);
  const lastRow = sh.getLastRow();
  sh.getRange(lastRow + 1, 4).setValue(name);
  return { added: true, name };
}

// 工程名変更（設定シート + マスタータスク + タスク一覧 を一括更新）
// ?mode=renamePhase&oldName=旧工程名&newName=新工程名
function _apiRenamePhase(params) {
  const { oldName, newName } = params;
  if (!oldName || !newName) throw new Error('oldName と newName は必須');
  const ss = _getSS();

  // 設定シート
  const configSh = ss.getSheetByName(SH.CONFIG);
  const configData = configSh.getDataRange().getValues();
  configData.forEach((row, i) => {
    if (row[3] === oldName) configSh.getRange(i + 1, 4).setValue(newName);
  });

  // マスタータスク
  const masterSh = ss.getSheetByName(SH.MASTER);
  const masterData = masterSh.getDataRange().getValues();
  let masterCount = 0;
  masterData.forEach((row, i) => {
    if (i > 0 && row[MC.PHASE - 1] === oldName) {
      masterSh.getRange(i + 1, MC.PHASE).setValue(newName);
      masterCount++;
    }
  });

  // タスク一覧
  const taskSh = ss.getSheetByName(SH.TASKS);
  const taskData = taskSh.getDataRange().getValues();
  let taskCount = 0;
  taskData.forEach((row, i) => {
    if (i > 0 && row[C.PHASE - 1] === oldName) {
      taskSh.getRange(i + 1, C.PHASE).setValue(newName);
      taskCount++;
    }
  });

  _writeHistory('', '', 'マスター', '工程名変更', oldName, newName);
  return { renamed: true, oldName, newName, masterCount, taskCount };
}


// ═══════════════════════════════════════════════════════════════
//  Notion API — 物件名リストのみ取得
// ═══════════════════════════════════════════════════════════════
function _getNotionProperties() {
  const results = [];
  let cursor = null;

  do {
    const body = { page_size: 100 };
    if (cursor) body.start_cursor = cursor;

    const res = _notionPost(
      'https://api.notion.com/v1/databases/' + NOTION_PROPS_DB_ID + '/query',
      body
    );

    res.results.forEach(page => {
      const p = page.properties;

      // 物件名（必須）
      const name = p[NF.NAME]?.title?.[0]?.plain_text || '';
      if (!name) return;

      // 除外リストに含まれるものはスキップ
      if (EXCLUDE_NAMES.includes(name)) return;

      // 進捗（select / multi_select 両対応）
      const progress = p[NF.PROGRESS]?.select?.name
        || p[NF.PROGRESS]?.multi_select?.[0]?.name
        || '';

      // 市町村（rich_text / select 両対応）
      const city = p[NF.CITY]?.rich_text?.[0]?.plain_text
        || p[NF.CITY]?.select?.name
        || '';

      // 各日付（date）
      const start    = p[NF.START]?.date?.start    || '';
      const frame    = p[NF.FRAME]?.date?.start    || '';
      const handover = p[NF.HANDOVER]?.date?.start || '';

      // チェックボックスフィールド
      const checks = {
        '棟札': p['棟札']?.checkbox === true,
        '鎮物': p['鎮物']?.checkbox === true,
        '先外': p['先外']?.checkbox === true,
        '改良': p['改良']?.checkbox === true,
        '外構': p['外構']?.checkbox === true,
      };

      // Notion page ID（ダッシュなし形式）
      const notionPageId = page.id.replace(/-/g, '');

      results.push({ name, progress, city, start, frame, handover, notionPageId, checks });
    });

    cursor = res.has_more ? res.next_cursor : null;
  } while (cursor);

  // 設定シートに最終同期日時を記録
  try {
    _getSS().getSheetByName(SH.CONFIG).getRange('B2').setValue(new Date());
  } catch(e) {}

  return results;
}

// ── Notion 物件チェックボックス更新 ──────────────────────────
// ?mode=updatePropertyCheck&pageId=NOTION_PAGE_ID&field=棟札&value=true
function _apiUpdatePropertyCheck(params) {
  const { pageId, field, value } = params;
  if (!pageId || !field) throw new Error('pageId と field は必須です');

  // セキュリティ: 許可するフィールドのみ
  const ALLOWED = ['棟札', '鎮物', '先外', '改良', '外構'];
  if (!ALLOWED.includes(field)) throw new Error('不正なフィールド名: ' + field);

  const boolValue = value === 'true' || value === true;

  _notionPatch(
    'https://api.notion.com/v1/pages/' + pageId,
    { properties: { [field]: { checkbox: boolValue } } }
  );

  return { pageId, field, value: boolValue };
}

function _notionPatch(url, body) {
  const token = (typeof NOTION_API_KEY === 'string' && NOTION_API_KEY.indexOf('xxxx') === -1)
    ? NOTION_API_KEY
    : PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_API_KEY を設定してください（ファイル先頭の①）');
  const res = UrlFetchApp.fetch(url, {
    method: 'patch',
    headers: {
      'Authorization'  : 'Bearer ' + token,
      'Notion-Version' : '2022-06-28',
      'Content-Type'   : 'application/json',
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const json = JSON.parse(res.getContentText());
  if (json.object === 'error') throw new Error('Notion API Error: ' + json.message);
  return json;
}

function _notionPost(url, body) {
  // APIキーは①のNOTION_API_KEY定数から読み取る
  // （セキュリティ強化したい場合はスクリプトプロパティ NOTION_TOKEN も自動で使える）
  const token = (typeof NOTION_API_KEY === 'string' && NOTION_API_KEY.indexOf('xxxx') === -1)
    ? NOTION_API_KEY
    : PropertiesService.getScriptProperties().getProperty('NOTION_TOKEN');
  if (!token) throw new Error('NOTION_API_KEY を設定してください（ファイル先頭の①）');
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    headers: {
      'Authorization'  : 'Bearer ' + token,
      'Notion-Version' : '2022-06-28',
      'Content-Type'   : 'application/json',
    },
    payload: JSON.stringify(body),
    muteHttpExceptions: true,
  });
  const json = JSON.parse(res.getContentText());
  if (json.object === 'error') throw new Error('Notion API Error: ' + json.message);
  return json;
}


// ═══════════════════════════════════════════════════════════════
//  ヘルパー
// ═══════════════════════════════════════════════════════════════
function _getSS() {
  const id = PropertiesService.getScriptProperties().getProperty(PROP_SS_ID);
  if (!id) throw new Error('setup() を先に実行してください');
  return SpreadsheetApp.openById(id);
}

function _writeHistory(taskId, taskName, property, field, before, after) {
  _getSS().getSheetByName(SH.HISTORY)
    .appendRow([new Date(), taskId, taskName, property, field, String(before), String(after)]);
}
