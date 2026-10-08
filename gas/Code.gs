/**
 * ご近所マルシェ＆助け合い広場 — GASバックエンド (Code.gs)
 *
 * 初回手順:
 *  1. スプレッドシートを新規作成 → 拡張機能 > Apps Script を開き、このコードを貼り付け
 *  2. プロジェクトの設定 > スクリプトプロパティに GEMINI_API_KEY を登録
 *     （任意: GEMINI_MODEL / ADMIN_EMAIL）
 *  3. エディタで setup() を1回実行（シート作成・Driveフォルダ作成・PEPPER生成・権限承認）
 *  4. デプロイ > 新しいデプロイ > ウェブアプリ
 *     実行ユーザー: 自分 / アクセスできるユーザー: 全員
 *  5. 発行されたURLを index.html の GAS_URL に貼り付け
 *  ※ コードを変更したら「デプロイを管理」> 編集 > バージョン「新バージョン」で再デプロイ
 */

// ===================== 設定 =====================
const APP_NAME = 'ご近所マルシェ＆助け合い広場';
const APP_URL = 'https://kenken6291.github.io/gokinjo-marche/';
const SESSION_TTL = 21600;         // 6時間（CacheService上限）
const MAX_LOGIN_FAIL = 5;          // 連続失敗でロック
const LOCK_MINUTES = 15;
const DEFAULT_GEMINI_MODEL = 'gemini-3.8-flash';

const SHEETS = {
  Users:   ['userId', 'email', 'passwordHash', 'salt', 'isTemp', 'area', 'nickname', 'createdAt', 'failCount', 'lockedUntil', 'lat', 'lng'],
  Posts:   ['postId', 'userId', 'type', 'category', 'title', 'detail', 'imageUrl', 'status', 'area', 'contacts', 'createdAt', 'place', 'lat', 'lng', 'station'],
  Events:  ['eventId', 'organizerId', 'category', 'title', 'datetime', 'place', 'detail', 'imageUrl', 'joinCount', 'likeCount', 'slots', 'contacts', 'area', 'createdAt', 'lat', 'lng', 'station'],
  Matches: ['matchId', 'targetId', 'targetType', 'applicantId', 'kind', 'note', 'createdAt']
};

const EVENT_CATEGORIES = ['marche', 'flea', 'offkai', 'matsuri', 'help'];
const POST_TYPES = ['exchange', 'help'];
const POST_CATEGORIES = ['vegetable', 'lend', 'barter', 'help'];
const POST_STATUSES = ['open', 'talking', 'done'];

function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.openById(props_().getProperty('SPREADSHEET_ID')); }
function sheet_(name) { return ss_().getSheetByName(name); }

// ===================== 初期セットアップ =====================
function setup() {
  const p = props_();
  let ss;
  try { ss = SpreadsheetApp.getActiveSpreadsheet(); } catch (e) { ss = null; }
  if (!ss) ss = SpreadsheetApp.create(APP_NAME + ' DB');
  p.setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(SHEETS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, SHEETS[name].length).setValues([SHEETS[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  const def = ss.getSheetByName('シート1') || ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 1) ss.deleteSheet(def);

  if (!p.getProperty('DRIVE_FOLDER_ID')) {
    const folder = DriveApp.createFolder(APP_NAME + '_画像');
    p.setProperty('DRIVE_FOLDER_ID', folder.getId());
  }
  if (!p.getProperty('PEPPER')) p.setProperty('PEPPER', Utilities.getUuid() + Utilities.getUuid());
  if (!p.getProperty('GEMINI_MODEL')) p.setProperty('GEMINI_MODEL', DEFAULT_GEMINI_MODEL);

  // 権限承認を促すためのダミー呼び出し
  GmailApp.getAliases();
  Logger.log('セットアップ完了: ' + ss.getUrl());
}

// ===================== ルーティング =====================
function doGet() {
  return json_({ ok: true, app: APP_NAME, time: new Date().toISOString() });
}

function doPost(e) {
  let req;
  try {
    req = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, error: 'リクエストの形式が正しくありません' });
  }
  const action = req.action;
  const routes = {
    // 認証（ログイン不要）
    register: register_,
    login: login_,
    resetPassword: resetPassword_,
    // 閲覧（ログイン不要）
    listEvents: listEvents_,
    listPosts: listPosts_,
    stations: stations_,
    // 以下ログイン必須
    me: me_,
    logout: logout_,
    changePassword: changePassword_,
    updateProfile: updateProfile_,
    uploadImage: uploadImage_,
    createEvent: createEvent_,
    updateEvent: updateEvent_,
    deleteEvent: deleteEvent_,
    toggleJoin: toggleJoin_,
    toggleLike: toggleLike_,
    toggleSlot: toggleSlot_,
    eventMembers: eventMembers_,
    createPost: createPost_,
    updatePost: updatePost_,
    deletePost: deletePost_,
    requestPost: requestPost_,
    aiWrite: aiWrite_,
    geocode: geocodeApi_
  };
  const publicActions = ['register', 'login', 'resetPassword', 'listEvents', 'listPosts', 'stations'];
  const fn = routes[action];
  if (!fn) return json_({ ok: false, error: '不明な操作です: ' + action });

  try {
    let user = null;
    if (req.token) user = userFromToken_(req.token);
    if (publicActions.indexOf(action) === -1) {
      if (!user) return json_({ ok: false, error: 'ログインの有効期限が切れました。もう一度ログインしてください', code: 'AUTH' });
      if (String(user.isTemp) === 'true' && action !== 'changePassword' && action !== 'me' && action !== 'logout') {
        return json_({ ok: false, error: '先に本パスワードを設定してください', code: 'TEMP' });
      }
    }
    const result = fn(req, user);
    return json_(Object.assign({ ok: true }, result));
  } catch (err) {
    console.error(err);
    return json_({ ok: false, error: err.message || String(err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// ===================== シート操作ヘルパー =====================
function readAll_(name) {
  const sh = sheet_(name);
  const values = sh.getDataRange().getValues();
  const head = values.shift();
  return values.map((row, i) => {
    const o = { _row: i + 2 };
    head.forEach((h, j) => { o[h] = row[j]; });
    return o;
  });
}

function append_(name, obj) {
  const head = SHEETS[name];
  sheet_(name).appendRow(head.map(h => obj[h] === undefined ? '' : obj[h]));
}

function updateRow_(name, rowNum, obj) {
  const head = SHEETS[name];
  const sh = sheet_(name);
  const current = sh.getRange(rowNum, 1, 1, head.length).getValues()[0];
  const next = head.map((h, j) => obj[h] === undefined ? current[j] : obj[h]);
  sh.getRange(rowNum, 1, 1, head.length).setValues([next]);
}

function findBy_(name, key, value) {
  return readAll_(name).find(r => String(r[key]) === String(value)) || null;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function newId_(prefix) {
  return prefix + '_' + Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyMMddHHmmss') + '_' + Utilities.getUuid().slice(0, 6);
}

function now_() { return new Date().toISOString(); }

function clean_(s, max) {
  s = String(s === undefined || s === null ? '' : s).trim();
  if (max) s = s.slice(0, max);
  // 数式インジェクション対策（= + - @ で始まる文字列は文字列として保存）
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function parseJson_(s, fallback) {
  try { return s ? JSON.parse(s) : fallback; } catch (e) { return fallback; }
}

// ===================== パスワード・セッション =====================
function hash_(password, salt) {
  const pepper = props_().getProperty('PEPPER');
  let digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + password + pepper, Utilities.Charset.UTF_8);
  for (let i = 0; i < 500; i++) {
    digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, digest.concat(Utilities.newBlob(salt).getBytes()));
  }
  // 16進文字列で保存（Base64の先頭「+」がシートで数式扱いされるのを防ぐ）
  return digest.map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
}

function genTempPassword_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZabcdefghijkmnpqrstuvwxyz23456789'; // 紛らわしい文字を除外
  let s = '';
  for (let i = 0; i < 8; i++) s += chars.charAt(Math.floor(Math.random() * chars.length));
  return s;
}

function createSession_(userId) {
  const token = Utilities.getUuid() + Utilities.getUuid().replace(/-/g, '');
  CacheService.getScriptCache().put('s_' + token, userId, SESSION_TTL);
  return token;
}

function userFromToken_(token) {
  const cache = CacheService.getScriptCache();
  const userId = cache.get('s_' + token);
  if (!userId) return null;
  cache.put('s_' + token, userId, SESSION_TTL); // 利用のたびに延長
  return findBy_('Users', 'userId', userId);
}

function publicUser_(u) {
  return {
    userId: u.userId, email: u.email, nickname: u.nickname, area: u.area, isTemp: String(u.isTemp) === 'true',
    lat: num_(u.lat), lng: num_(u.lng)
  };
}

function validEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validPassword_(pw) {
  return typeof pw === 'string' && pw.length >= 8 && /[A-Za-z]/.test(pw) && /[0-9]/.test(pw);
}

function sendMail_(to, subject, body) {
  const remaining = MailApp.getRemainingDailyQuota();
  if (remaining < 1) throw new Error('本日のメール送信上限に達しました。明日もう一度お試しください');
  try {
    GmailApp.sendEmail(to, '【' + APP_NAME + '】' + subject,
      body + '\n\n――――――――――\n' + APP_NAME + '\n' + APP_URL + '\n※このメールは送信専用です。',
      { name: APP_NAME });
    console.log('メール送信OK: ' + to + ' / 残り' + (remaining - 1) + '通');
  } catch (e) {
    console.error('メール送信失敗: ' + to + ' / ' + e);
    throw new Error('メールを送れませんでした（' + e.message + '）。管理者にお知らせください');
  }
}

// ===================== 診断（エディタから実行） =====================
/**
 * メールが届かないときにエディタで実行してください。
 * 設定の確認と、自分あてのテストメール送信を行います。結果は「実行ログ」に出ます。
 */
function diagnose() {
  const p = props_();
  const log = (ok, msg) => Logger.log((ok ? '✅ ' : '❌ ') + msg);
  log(!!p.getProperty('SPREADSHEET_ID'), 'SPREADSHEET_ID ' + (p.getProperty('SPREADSHEET_ID') || '未設定 → setup() を実行'));
  log(!!p.getProperty('DRIVE_FOLDER_ID'), 'DRIVE_FOLDER_ID ' + (p.getProperty('DRIVE_FOLDER_ID') || '未設定 → setup() を実行'));
  log(!!p.getProperty('PEPPER'), 'PEPPER ' + (p.getProperty('PEPPER') ? '設定済み' : '未設定 → setup() を実行'));
  log(!!p.getProperty('GEMINI_API_KEY'), 'GEMINI_API_KEY ' + (p.getProperty('GEMINI_API_KEY') ? '設定済み' : '未設定'));
  try {
    const n = readAll_('Users').length;
    log(true, 'Usersシート読み込みOK（' + n + '人）');
  } catch (e) { log(false, 'Usersシートを読めません: ' + e.message); }

  const me = Session.getEffectiveUser().getEmail();
  Logger.log('送信元アカウント: ' + me);
  Logger.log('本日の残り送信可能数: ' + MailApp.getRemainingDailyQuota());
  try {
    sendMail_(me, 'テストメール', 'このメールが届いていれば、メール送信は正常です。');
    log(true, 'テストメールを ' + me + ' に送りました。受信箱を確認してください');
  } catch (e) { log(false, e.message); }
}

/** 特定アドレスへテスト送信（TEST_TO を書き換えて実行） */
function testMailTo() {
  const TEST_TO = 'example@example.com';
  sendMail_(TEST_TO, 'テストメール', 'このメールが届いていれば、このアドレスへの送信は正常です。');
  Logger.log('送信しました: ' + TEST_TO);
}

// ===================== 認証API =====================
function register_(req) {
  const email = clean_(req.email, 200).toLowerCase();
  const nickname = clean_(req.nickname, 30);
  const area = clean_(req.area, 50);
  if (!validEmail_(email)) throw new Error('メールアドレスの形式を確認してください');
  if (!nickname) throw new Error('ニックネームを入力してください');
  if (!area) throw new Error('お住まいの地域を入力してください');

  return withLock_(() => {
    if (findBy_('Users', 'email', email)) throw new Error('このメールアドレスは登録済みです。「パスワードを忘れた」から仮パスワードを再発行できます');
    const temp = genTempPassword_();
    const salt = Utilities.getUuid();
    // 先にメールを送り、送れなかった場合は登録しない（「登録済み」で詰まるのを防ぐ）
    sendMail_(email, '仮パスワードのお知らせ',
      nickname + ' さん\n\nご登録ありがとうございます。\n\n仮パスワード： ' + temp +
      '\n\nこの仮パスワードでログインすると、本パスワードの設定画面が表示されます。\n本パスワードを設定するまで、ほかの機能は使えません。');
    const g = safeGeocode_(area);
    append_('Users', {
      userId: newId_('U'), email: email, passwordHash: hash_(temp, salt), salt: salt,
      isTemp: 'true', area: area, nickname: nickname, createdAt: now_(), failCount: 0, lockedUntil: '',
      lat: g ? round_(g.lat, 3) : '', lng: g ? round_(g.lng, 3) : ''
    });
    return { message: '仮パスワードをメールで送りました。届かない場合は迷惑メールフォルダを確認してください' };
  });
}

function login_(req) {
  const email = clean_(req.email, 200).toLowerCase();
  const password = String(req.password || '');
  const u = findBy_('Users', 'email', email);
  if (!u) throw new Error('メールアドレスかパスワードが違います');

  if (u.lockedUntil && new Date(u.lockedUntil) > new Date()) {
    throw new Error('ログインに続けて失敗したため、' + LOCK_MINUTES + '分間ロック中です。時間をおくか、パスワードを再発行してください');
  }
  if (hash_(password, u.salt) !== u.passwordHash) {
    const fail = Number(u.failCount || 0) + 1;
    const lockedUntil = fail >= MAX_LOGIN_FAIL ? new Date(Date.now() + LOCK_MINUTES * 60000).toISOString() : '';
    updateRow_('Users', u._row, { failCount: fail >= MAX_LOGIN_FAIL ? 0 : fail, lockedUntil: lockedUntil });
    throw new Error('メールアドレスかパスワードが違います（あと' + Math.max(0, MAX_LOGIN_FAIL - fail) + '回でロック）');
  }
  updateRow_('Users', u._row, { failCount: 0, lockedUntil: '' });
  return { token: createSession_(u.userId), user: publicUser_(u) };
}

function resetPassword_(req) {
  const email = clean_(req.email, 200).toLowerCase();
  if (!validEmail_(email)) throw new Error('メールアドレスの形式を確認してください');
  const u = findBy_('Users', 'email', email);
  // 登録の有無を外部に漏らさないため、常に同じ応答を返す
  if (u) {
    const temp = genTempPassword_();
    const salt = Utilities.getUuid();
    // 先にメールを送り、送れた場合だけパスワードを差し替える
    sendMail_(email, '仮パスワードの再発行',
      u.nickname + ' さん\n\n新しい仮パスワード： ' + temp +
      '\n\nログイン後、本パスワードを設定してください。\n心当たりがない場合はこのメールを破棄してください。');
    updateRow_('Users', u._row, { passwordHash: hash_(temp, salt), salt: salt, isTemp: 'true', failCount: 0, lockedUntil: '' });
  } else {
    console.log('resetPassword: 未登録のアドレス ' + email);
  }
  return { message: '登録済みのアドレスであれば、新しい仮パスワードを送りました。届かない場合は迷惑メールフォルダを確認してください' };
}

function me_(req, user) {
  if (num_(user.lat) === null && user.area) {
    const g = safeGeocode_(user.area);
    if (g) {
      updateRow_('Users', user._row, { lat: round_(g.lat, 3), lng: round_(g.lng, 3) });
      user = findBy_('Users', 'userId', user.userId);
    }
  }
  return { user: publicUser_(user) };
}

function logout_(req) {
  CacheService.getScriptCache().remove('s_' + req.token);
  return {};
}

function changePassword_(req, user) {
  const pw = String(req.newPassword || '');
  if (!validPassword_(pw)) throw new Error('パスワードは英字と数字を含む8文字以上にしてください');
  if (req.currentPassword !== undefined && String(user.isTemp) !== 'true') {
    if (hash_(String(req.currentPassword), user.salt) !== user.passwordHash) throw new Error('現在のパスワードが違います');
  }
  const salt = Utilities.getUuid();
  updateRow_('Users', user._row, { passwordHash: hash_(pw, salt), salt: salt, isTemp: 'false' });
  const u = findBy_('Users', 'userId', user.userId);
  return { user: publicUser_(u), message: 'パスワードを設定しました' };
}

function updateProfile_(req, user) {
  const upd = {};
  if (req.nickname !== undefined || req.area !== undefined) {
    const nickname = clean_(req.nickname, 30);
    const area = clean_(req.area, 50);
    if (!nickname || !area) throw new Error('ニックネームと地域を入力してください');
    upd.nickname = nickname;
    upd.area = area;
    if (area !== String(user.area) && !validLatLng_(req.lat, req.lng)) {
      const g = safeGeocode_(area);
      if (g) { upd.lat = round_(g.lat, 3); upd.lng = round_(g.lng, 3); }
    }
  }
  if (validLatLng_(req.lat, req.lng)) {
    // 自宅を特定されないよう約100m単位に丸めて保存（本人以外には返さない）
    upd.lat = round_(Number(req.lat), 3);
    upd.lng = round_(Number(req.lng), 3);
  }
  updateRow_('Users', user._row, upd);
  return { user: publicUser_(findBy_('Users', 'userId', user.userId)) };
}

// ===================== 画像アップロード =====================
function uploadImage_(req, user) {
  const dataUrl = String(req.dataUrl || '');
  const m = dataUrl.match(/^data:(image\/(png|jpeg|jpg|webp|gif));base64,(.+)$/);
  if (!m) throw new Error('画像の形式が正しくありません（JPEG/PNG/WebP/GIF）');
  const bytes = Utilities.base64Decode(m[3]);
  if (bytes.length > 5 * 1024 * 1024) throw new Error('画像が大きすぎます（5MBまで）');
  const ext = m[2] === 'jpeg' ? 'jpg' : m[2];
  const blob = Utilities.newBlob(bytes, m[1], user.userId + '_' + Date.now() + '.' + ext);
  const folder = DriveApp.getFolderById(props_().getProperty('DRIVE_FOLDER_ID'));
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { imageUrl: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1200', fileId: file.getId() };
}

// ===================== 連絡先（外部ツール） =====================
function sanitizeContacts_(c) {
  c = c || {};
  const url = v => {
    v = clean_(v, 300);
    return /^https?:\/\//i.test(v) ? v : '';
  };
  return {
    zoom: url(c.zoom),
    meet: url(c.meet),
    line: url(c.line),
    facebook: url(c.facebook),
    email: validEmail_(clean_(c.email, 200)) ? clean_(c.email, 200) : '',
    phone: clean_(c.phone, 20).replace(/[^\d\-+]/g, '')
  };
}

// ===================== イベント =====================
function sanitizeSlots_(slots) {
  if (!Array.isArray(slots)) return [];
  return slots.slice(0, 10).map((s, i) => ({
    id: s.id || ('slot' + (i + 1)),
    label: clean_(s.label, 40),
    capacity: Math.max(1, Math.min(99, Number(s.capacity) || 1)),
    members: Array.isArray(s.members) ? s.members : []
  })).filter(s => s.label);
}

function listEvents_(req, user) {
  const users = userMap_();
  const myId = user ? user.userId : null;
  const mine = myId ? myMatches_(myId) : {};
  const showPast = !!req.includePast;
  const today = new Date(); today.setHours(0, 0, 0, 0);

  const events = readAll_('Events')
    .filter(ev => showPast || !ev.datetime || new Date(ev.datetime) >= today)
    .map(ev => {
      const slots = parseJson_(ev.slots, []).map(s => ({
        id: s.id, label: s.label, capacity: s.capacity, count: s.members.length,
        mine: myId ? s.members.indexOf(myId) !== -1 : false,
        memberNames: myId ? s.members.map(id => (users[id] || {}).nickname || '退会済み') : []
      }));
      return {
        eventId: ev.eventId, category: ev.category, title: ev.title, datetime: ev.datetime,
        place: ev.place, detail: ev.detail, imageUrl: ev.imageUrl, area: ev.area,
        lat: num_(ev.lat), lng: num_(ev.lng), station: parseJson_(ev.station, null),
        joinCount: Number(ev.joinCount || 0), likeCount: Number(ev.likeCount || 0),
        slots: slots,
        organizer: { userId: ev.organizerId, nickname: (users[ev.organizerId] || {}).nickname || '退会済み' },
        // 連絡先はログイン会員のみに公開
        contacts: myId ? parseJson_(ev.contacts, {}) : null,
        joined: !!mine['join:' + ev.eventId],
        liked: !!mine['like:' + ev.eventId],
        isOwner: myId === ev.organizerId,
        createdAt: ev.createdAt
      };
    })
    .sort((a, b) => new Date(a.datetime) - new Date(b.datetime));
  return { events: events };
}

function createEvent_(req, user) {
  const d = req.data || {};
  if (EVENT_CATEGORIES.indexOf(d.category) === -1) throw new Error('種類を選んでください');
  if (!clean_(d.title)) throw new Error('タイトルを入力してください');
  if (!d.datetime || isNaN(new Date(d.datetime))) throw new Error('日時を入力してください');
  if (!clean_(d.place)) throw new Error('開催場所を入力してください');
  const loc = resolveLocation_(d, clean_(d.place, 100) + ' ' + (clean_(d.area, 50) || user.area), null, 6);
  const ev = {
    eventId: newId_('E'), organizerId: user.userId, category: d.category,
    title: clean_(d.title, 60), datetime: new Date(d.datetime).toISOString(), place: clean_(d.place, 100),
    detail: clean_(d.detail, 3000), imageUrl: clean_(d.imageUrl, 300),
    joinCount: 0, likeCount: 0,
    slots: JSON.stringify(sanitizeSlots_(d.slots)),
    contacts: JSON.stringify(sanitizeContacts_(d.contacts)),
    area: clean_(d.area, 50) || user.area, createdAt: now_(),
    lat: loc ? loc.lat : '', lng: loc ? loc.lng : '', station: loc && loc.station ? JSON.stringify(loc.station) : ''
  };
  withLock_(() => append_('Events', ev));
  return { eventId: ev.eventId };
}

function updateEvent_(req, user) {
  const d = req.data || {};
  return withLock_(() => {
    const ev = findBy_('Events', 'eventId', req.eventId);
    if (!ev) throw new Error('イベントが見つかりません');
    if (ev.organizerId !== user.userId) throw new Error('主催者のみ編集できます');
    // 既存の出店枠の申込者を引き継ぐ
    const oldSlots = parseJson_(ev.slots, []);
    const newSlots = sanitizeSlots_(d.slots).map(s => {
      const old = oldSlots.find(o => o.id === s.id);
      s.members = old ? old.members : [];
      return s;
    });
    const place = clean_(d.place, 100) || ev.place;
    const area = clean_(d.area, 50) || ev.area;
    const placeChanged = place !== String(ev.place) || area !== String(ev.area);
    let loc = null;
    if (validLatLng_(d.lat, d.lng) || placeChanged || num_(ev.lat) === null) {
      loc = resolveLocation_(d, place + ' ' + area, null, 6);
    }
    updateRow_('Events', ev._row, {
      lat: loc ? loc.lat : undefined, lng: loc ? loc.lng : undefined,
      station: loc ? (loc.station ? JSON.stringify(loc.station) : '') : undefined,
      category: EVENT_CATEGORIES.indexOf(d.category) !== -1 ? d.category : ev.category,
      title: clean_(d.title, 60) || ev.title,
      datetime: d.datetime && !isNaN(new Date(d.datetime)) ? new Date(d.datetime).toISOString() : ev.datetime,
      place: clean_(d.place, 100) || ev.place,
      detail: clean_(d.detail, 3000),
      imageUrl: d.imageUrl !== undefined ? clean_(d.imageUrl, 300) : ev.imageUrl,
      slots: JSON.stringify(newSlots),
      contacts: JSON.stringify(sanitizeContacts_(d.contacts)),
      area: clean_(d.area, 50) || ev.area
    });
    return {};
  });
}

function deleteEvent_(req, user) {
  return withLock_(() => {
    const ev = findBy_('Events', 'eventId', req.eventId);
    if (!ev) throw new Error('イベントが見つかりません');
    if (ev.organizerId !== user.userId) throw new Error('主催者のみ削除できます');
    sheet_('Events').deleteRow(ev._row);
    deleteMatchesFor_(ev.eventId);
    return {};
  });
}

function toggleJoin_(req, user) {
  return toggleReaction_(req.eventId, user, 'join', 'joinCount');
}

function toggleLike_(req, user) {
  return toggleReaction_(req.eventId, user, 'like', 'likeCount');
}

function toggleReaction_(eventId, user, kind, countCol) {
  return withLock_(() => {
    const ev = findBy_('Events', 'eventId', eventId);
    if (!ev) throw new Error('イベントが見つかりません');
    const existing = readAll_('Matches').find(m => m.targetId === eventId && m.applicantId === user.userId && m.kind === kind);
    let active;
    if (existing) {
      sheet_('Matches').deleteRow(existing._row);
      active = false;
    } else {
      append_('Matches', { matchId: newId_('M'), targetId: eventId, targetType: 'event', applicantId: user.userId, kind: kind, note: '', createdAt: now_() });
      active = true;
    }
    const count = Math.max(0, Number(ev[countCol] || 0) + (active ? 1 : -1));
    const upd = {}; upd[countCol] = count;
    updateRow_('Events', ev._row, upd);

    if (kind === 'join' && active && ev.organizerId !== user.userId) {
      const org = findBy_('Users', 'userId', ev.organizerId);
      if (org) {
        try {
          sendMail_(org.email, '「' + ev.title + '」に参加表明がありました',
            org.nickname + ' さん\n\n' + user.nickname + ' さん（' + user.area + '）が「' + ev.title + '」に参加表明しました。\n現在の参加予定：' + count + '人\n\n参加者一覧はサイトのイベント詳細から確認できます。');
        } catch (e) { console.warn('通知メール失敗: ' + e); }
      }
    }
    return { active: active, count: count };
  });
}

function toggleSlot_(req, user) {
  return withLock_(() => {
    const ev = findBy_('Events', 'eventId', req.eventId);
    if (!ev) throw new Error('イベントが見つかりません');
    const slots = parseJson_(ev.slots, []);
    const slot = slots.find(s => s.id === req.slotId);
    if (!slot) throw new Error('募集枠が見つかりません');
    const idx = slot.members.indexOf(user.userId);
    let active;
    if (idx !== -1) {
      slot.members.splice(idx, 1);
      active = false;
    } else {
      if (slot.members.length >= slot.capacity) throw new Error('この枠は定員に達しました');
      slot.members.push(user.userId);
      active = true;
    }
    updateRow_('Events', ev._row, { slots: JSON.stringify(slots) });

    if (active && ev.organizerId !== user.userId) {
      const org = findBy_('Users', 'userId', ev.organizerId);
      if (org) {
        try {
          sendMail_(org.email, '「' + ev.title + '」の募集枠に申込がありました',
            org.nickname + ' さん\n\n' + user.nickname + ' さんが「' + slot.label + '」に申し込みました。（' + slot.members.length + '/' + slot.capacity + '）');
        } catch (e) { console.warn('通知メール失敗: ' + e); }
      }
    }
    return { active: active, count: slot.members.length };
  });
}

// 主催者と参加者だけが参加者一覧を見られる
function eventMembers_(req, user) {
  const ev = findBy_('Events', 'eventId', req.eventId);
  if (!ev) throw new Error('イベントが見つかりません');
  const joins = readAll_('Matches').filter(m => m.targetId === ev.eventId && m.kind === 'join');
  const isMember = ev.organizerId === user.userId || joins.some(m => m.applicantId === user.userId);
  if (!isMember) throw new Error('参加表明すると参加者一覧を見られます');
  const users = userMap_();
  return {
    members: joins.map(m => ({ nickname: (users[m.applicantId] || {}).nickname || '退会済み', area: (users[m.applicantId] || {}).area || '' }))
  };
}

// ===================== おすそ分け・助け合い =====================
function listPosts_(req, user) {
  const users = userMap_();
  const myId = user ? user.userId : null;
  const reqs = readAll_('Matches').filter(m => m.targetType === 'post' && m.kind === 'request');
  const posts = readAll_('Posts').map(p => {
    const rs = reqs.filter(m => m.targetId === p.postId);
    return {
      postId: p.postId, type: p.type, category: p.category, title: p.title, detail: p.detail,
      imageUrl: p.imageUrl, status: p.status, area: p.area, createdAt: p.createdAt,
      place: p.place, lat: num_(p.lat), lng: num_(p.lng), station: parseJson_(p.station, null),
      owner: { userId: p.userId, nickname: (users[p.userId] || {}).nickname || '退会済み' },
      contacts: myId ? parseJson_(p.contacts, {}) : null,
      requestCount: rs.length,
      requested: myId ? rs.some(m => m.applicantId === myId) : false,
      requesters: myId === p.userId ? rs.map(m => ({ nickname: (users[m.applicantId] || {}).nickname || '退会済み', note: m.note, createdAt: m.createdAt })) : [],
      isOwner: myId === p.userId
    };
  }).sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return { posts: posts };
}

function createPost_(req, user) {
  const d = req.data || {};
  if (POST_CATEGORIES.indexOf(d.category) === -1) throw new Error('カテゴリを選んでください');
  if (!clean_(d.title)) throw new Error('タイトルを入力してください');
  const type = d.category === 'help' ? 'help' : 'exchange';
  const place = clean_(d.place, 100);
  const area = clean_(d.area, 50) || user.area;
  const fallback = num_(user.lat) !== null ? { lat: num_(user.lat), lng: num_(user.lng) } : null;
  const loc = resolveLocation_(d, place ? place + ' ' + area : area, fallback, 3);
  const post = {
    postId: newId_('P'), userId: user.userId, type: type, category: d.category,
    title: clean_(d.title, 60), detail: clean_(d.detail, 2000), imageUrl: clean_(d.imageUrl, 300),
    status: 'open', area: clean_(d.area, 50) || user.area,
    contacts: JSON.stringify(sanitizeContacts_(d.contacts)), createdAt: now_(),
    place: place, lat: loc ? loc.lat : '', lng: loc ? loc.lng : '',
    station: loc && loc.station ? JSON.stringify(loc.station) : ''
  };
  withLock_(() => append_('Posts', post));
  return { postId: post.postId };
}

function updatePost_(req, user) {
  const d = req.data || {};
  return withLock_(() => {
    const p = findBy_('Posts', 'postId', req.postId);
    if (!p) throw new Error('投稿が見つかりません');
    if (p.userId !== user.userId) throw new Error('投稿者のみ編集できます');
    const upd = {};
    if (d.status !== undefined) {
      if (POST_STATUSES.indexOf(d.status) === -1) throw new Error('状態が正しくありません');
      upd.status = d.status;
    }
    if (d.title !== undefined) upd.title = clean_(d.title, 60) || p.title;
    if (d.detail !== undefined) upd.detail = clean_(d.detail, 2000);
    if (d.category !== undefined && POST_CATEGORIES.indexOf(d.category) !== -1) {
      upd.category = d.category;
      upd.type = d.category === 'help' ? 'help' : 'exchange';
    }
    if (d.imageUrl !== undefined) upd.imageUrl = clean_(d.imageUrl, 300);
    if (d.area !== undefined) upd.area = clean_(d.area, 50) || p.area;
    if (d.contacts !== undefined) upd.contacts = JSON.stringify(sanitizeContacts_(d.contacts));
    if (d.place !== undefined || validLatLng_(d.lat, d.lng)) {
      const place = d.place !== undefined ? clean_(d.place, 100) : String(p.place || '');
      const area = upd.area || p.area;
      const fallback = num_(user.lat) !== null ? { lat: num_(user.lat), lng: num_(user.lng) } : null;
      const loc = resolveLocation_(d, place ? place + ' ' + area : area, fallback, 3);
      upd.place = place;
      if (loc) {
        upd.lat = loc.lat; upd.lng = loc.lng;
        upd.station = loc.station ? JSON.stringify(loc.station) : '';
      }
    }
    updateRow_('Posts', p._row, upd);
    return {};
  });
}

function deletePost_(req, user) {
  return withLock_(() => {
    const p = findBy_('Posts', 'postId', req.postId);
    if (!p) throw new Error('投稿が見つかりません');
    if (p.userId !== user.userId) throw new Error('投稿者のみ削除できます');
    sheet_('Posts').deleteRow(p._row);
    deleteMatchesFor_(p.postId);
    return {};
  });
}

// 「ほしい・手伝いたい」申込（取り消しも可）
function requestPost_(req, user) {
  return withLock_(() => {
    const p = findBy_('Posts', 'postId', req.postId);
    if (!p) throw new Error('投稿が見つかりません');
    if (p.userId === user.userId) throw new Error('自分の投稿には申し込めません');
    const existing = readAll_('Matches').find(m => m.targetId === p.postId && m.applicantId === user.userId && m.kind === 'request');
    if (existing) {
      sheet_('Matches').deleteRow(existing._row);
      return { active: false };
    }
    if (p.status === 'done') throw new Error('この投稿は受付を終了しました');
    const note = clean_(req.note, 300);
    append_('Matches', { matchId: newId_('M'), targetId: p.postId, targetType: 'post', applicantId: user.userId, kind: 'request', note: note, createdAt: now_() });
    const owner = findBy_('Users', 'userId', p.userId);
    if (owner) {
      try {
        sendMail_(owner.email, '「' + p.title + '」に申込がありました',
          owner.nickname + ' さん\n\n' + user.nickname + ' さん（' + user.area + '）から申込がありました。' +
          (note ? '\n\nひとこと：' + note : '') +
          '\n\nやりとりは投稿に載せた連絡先（LINE・メール等）で行われます。サイトの投稿詳細で申込者を確認できます。');
      } catch (e) { console.warn('通知メール失敗: ' + e); }
    }
    return { active: true };
  });
}

// ===================== 位置情報 =====================
function num_(v) {
  if (v === '' || v === null || v === undefined) return null;
  const n = Number(v);
  return isNaN(n) ? null : n;
}

function round_(n, digits) {
  const f = Math.pow(10, digits);
  return Math.round(Number(n) * f) / f;
}

function validLatLng_(lat, lng) {
  const a = num_(lat), b = num_(lng);
  return a !== null && b !== null && Math.abs(a) <= 90 && Math.abs(b) <= 180 && !(a === 0 && b === 0);
}

function distanceM_(lat1, lng1, lat2, lng2) {
  const R = 6371000, toR = Math.PI / 180;
  const dLat = (lat2 - lat1) * toR, dLng = (lng2 - lng1) * toR;
  const a = Math.sin(dLat / 2) * Math.sin(dLat / 2) +
    Math.cos(lat1 * toR) * Math.cos(lat2 * toR) * Math.sin(dLng / 2) * Math.sin(dLng / 2);
  return Math.round(2 * R * Math.asin(Math.sqrt(a)));
}

/** 住所・場所名 → 緯度経度（Apps Script標準のMapsサービス、結果は6時間キャッシュ） */
function geocode_(address) {
  address = String(address || '').replace(/^'/, '').trim();
  if (!address) return null;
  const cache = CacheService.getScriptCache();
  const key = 'g_' + Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, address, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 0xff).toString(16)).slice(-2)).join('');
  const hit = cache.get(key);
  if (hit) return hit === 'null' ? null : JSON.parse(hit);
  const res = Maps.newGeocoder().setLanguage('ja').setRegion('jp').geocode(address);
  let out = null;
  if (res.status === 'OK' && res.results && res.results.length) {
    const r = res.results[0];
    out = { lat: r.geometry.location.lat, lng: r.geometry.location.lng, formatted: r.formatted_address };
  }
  cache.put(key, out ? JSON.stringify(out) : 'null', 21600);
  return out;
}

function safeGeocode_(address) {
  try { return geocode_(address); } catch (e) { console.warn('geocode失敗: ' + e); return null; }
}

/** 最寄り駅（HeartRails Express API）。近い順・駅名で重複をまとめて返す */
function nearestStations_(lat, lng, n) {
  const cache = CacheService.getScriptCache();
  const key = 'st_' + round_(lat, 3) + '_' + round_(lng, 3);
  const hit = cache.get(key);
  let list;
  if (hit) {
    list = JSON.parse(hit);
  } else {
    const url = 'https://express.heartrails.com/api/json?method=getStations&x=' + lng + '&y=' + lat;
    const res = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('駅の情報を取得できませんでした');
    const raw = ((JSON.parse(res.getContentText()).response || {}).station) || [];
    const byName = {};
    raw.forEach(st => {
      const name = st.name;
      const d = distanceM_(lat, lng, Number(st.y), Number(st.x));
      if (!byName[name]) byName[name] = { name: name, lines: [], distance: d, lat: Number(st.y), lng: Number(st.x) };
      if (byName[name].lines.indexOf(st.line) === -1) byName[name].lines.push(st.line);
      if (d < byName[name].distance) byName[name].distance = d;
    });
    list = Object.keys(byName).map(k => byName[k]).sort((a, b) => a.distance - b.distance);
    cache.put(key, JSON.stringify(list), 21600);
  }
  return list.slice(0, n || 3);
}

function safeNearestStation_(lat, lng) {
  try {
    const st = nearestStations_(lat, lng, 1)[0];
    return st ? { name: st.name, distance: st.distance, lines: st.lines } : null;
  } catch (e) { console.warn('駅取得失敗: ' + e); return null; }
}

/**
 * 位置を決める：地図で指定された緯度経度 → 住所から検索 → 予備の位置 の順。
 * digits: 保存する小数桁（6=正確、3=約100mぼかし）
 */
function resolveLocation_(d, address, fallback, digits) {
  let lat = null, lng = null;
  if (validLatLng_(d.lat, d.lng)) {
    lat = Number(d.lat); lng = Number(d.lng);
  } else {
    const g = safeGeocode_(address);
    if (g) { lat = g.lat; lng = g.lng; }
    else if (fallback) { lat = fallback.lat; lng = fallback.lng; }
  }
  if (lat === null) return null;
  lat = round_(lat, digits); lng = round_(lng, digits);
  return { lat: lat, lng: lng, station: safeNearestStation_(lat, lng) };
}

// API: 住所・場所名から位置を探す（会員のみ）
function geocodeApi_(req) {
  const q = clean_(req.address, 150);
  if (!q) throw new Error('住所か場所の名前を入力してください');
  const g = geocode_(q);
  if (!g) throw new Error('場所が見つかりませんでした。市区町村名から入れるか、地図をタップして選んでください');
  return { lat: g.lat, lng: g.lng, formatted: g.formatted };
}

// API: 最寄り駅（ログイン不要）
function stations_(req) {
  if (!validLatLng_(req.lat, req.lng)) throw new Error('位置が正しくありません');
  return { stations: nearestStations_(Number(req.lat), Number(req.lng), 3) };
}

/**
 * 位置情報を追加する前に作られた会員・イベント・投稿に、位置と最寄り駅を付けます。
 * setup() のあとにエディタで1回実行してください。
 */
function backfillLocations() {
  let n = 0;
  readAll_('Users').forEach(u => {
    if (num_(u.lat) === null && u.area) {
      const g = safeGeocode_(u.area);
      if (g) { updateRow_('Users', u._row, { lat: round_(g.lat, 3), lng: round_(g.lng, 3) }); n++; }
    }
  });
  const users = {};
  readAll_('Users').forEach(u => { users[u.userId] = u; });
  readAll_('Events').forEach(ev => {
    if (num_(ev.lat) === null) {
      const loc = resolveLocation_({}, ev.place + ' ' + ev.area, null, 6);
      if (loc) { updateRow_('Events', ev._row, { lat: loc.lat, lng: loc.lng, station: loc.station ? JSON.stringify(loc.station) : '' }); n++; }
    }
  });
  readAll_('Posts').forEach(p => {
    if (num_(p.lat) === null) {
      const u = users[p.userId] || {};
      const fb = num_(u.lat) !== null ? { lat: num_(u.lat), lng: num_(u.lng) } : null;
      const loc = resolveLocation_({}, p.place ? p.place + ' ' + p.area : p.area, fb, 3);
      if (loc) { updateRow_('Posts', p._row, { lat: loc.lat, lng: loc.lng, station: loc.station ? JSON.stringify(loc.station) : '' }); n++; }
    }
  });
  Logger.log('位置を付けた件数: ' + n);
}

// ===================== 共通 =====================
function userMap_() {
  const map = {};
  readAll_('Users').forEach(u => { map[u.userId] = { nickname: u.nickname, area: u.area }; });
  return map;
}

function myMatches_(userId) {
  const map = {};
  readAll_('Matches').forEach(m => { if (m.applicantId === userId) map[m.kind + ':' + m.targetId] = true; });
  return map;
}

function deleteMatchesFor_(targetId) {
  const sh = sheet_('Matches');
  const rows = readAll_('Matches').filter(m => m.targetId === targetId).map(m => m._row).sort((a, b) => b - a);
  rows.forEach(r => sh.deleteRow(r));
}

// ===================== Gemini（告知文アシスト） =====================
function aiWrite_(req, user) {
  const apiKey = props_().getProperty('GEMINI_API_KEY');
  if (!apiKey) throw new Error('AIアシストは準備中です（GEMINI_API_KEY未設定）');
  const model = props_().getProperty('GEMINI_MODEL') || DEFAULT_GEMINI_MODEL;
  const d = req.data || {};
  const isEvent = req.kind === 'event';

  const info = isEvent
    ? '種類: ' + clean_(d.categoryLabel, 30) + '\nタイトル(仮): ' + clean_(d.title, 60) + '\n日時: ' + clean_(d.datetimeText, 60) +
      '\n場所: ' + clean_(d.place, 100) + '\n地域: ' + (clean_(d.area, 50) || user.area) + '\nメモ: ' + clean_(d.detail, 800) +
      '\n募集枠: ' + clean_(d.slotsText, 300)
    : 'カテゴリ: ' + clean_(d.categoryLabel, 30) + '\nタイトル(仮): ' + clean_(d.title, 60) + '\n地域: ' + (clean_(d.area, 50) || user.area) +
      '\nメモ: ' + clean_(d.detail, 800);

  const prompt =
    'あなたは地域のマルシェや縁日のチラシを書くのが得意なコピーライターです。\n' +
    (isEvent ? '次の地域イベントの告知文' : '次のおすそ分け・物々交換・お手伝いの紹介文') + 'を書いてください。\n\n' + info + '\n\n' +
    '条件:\n' +
    '- 読み手はご近所の若者から高齢者まで。むずかしい言葉・カタカナ語は避け、やさしい日本語で\n' +
    '- 楽しさが伝わる、あたたかい口調。絵文字は2〜4個まで\n' +
    '- 入力にない日時・料金・持ち物などの事実を作らない\n' +
    '- catchcopy は20文字以内の見出し、title は30文字以内、body は200〜350文字程度。' + (isEvent ? '日時と場所を本文に必ず入れる' : '') + '\n' +
    '- 次のJSONだけを返す: {"catchcopy":"...","title":"...","body":"..."}';

  const url = 'https://generativelanguage.googleapis.com/v1beta/models/' + encodeURIComponent(model) + ':generateContent';
  const res = UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: { 'x-goog-api-key': apiKey },
    muteHttpExceptions: true,
    payload: JSON.stringify({
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.8, responseMimeType: 'application/json' }
    })
  });
  if (res.getResponseCode() !== 200) {
    console.error(res.getContentText());
    throw new Error('AIの文章作成に失敗しました（' + res.getResponseCode() + '）。少し時間をおいて試してください');
  }
  const body = JSON.parse(res.getContentText());
  const text = (((body.candidates || [])[0] || {}).content || {}).parts ? body.candidates[0].content.parts.map(p => p.text || '').join('') : '';
  const out = parseJson_(text.replace(/```json|```/g, '').trim(), null);
  if (!out) throw new Error('AIの返答を読み取れませんでした。もう一度お試しください');
  return { catchcopy: clean_(out.catchcopy, 40), title: clean_(out.title, 60), body: clean_(out.body, 1500) };
}
