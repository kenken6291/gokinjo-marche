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
  Users:   ['userId', 'email', 'passwordHash', 'salt', 'isTemp', 'area', 'nickname', 'createdAt', 'failCount', 'lockedUntil'],
  Posts:   ['postId', 'userId', 'type', 'category', 'title', 'detail', 'imageUrl', 'status', 'area', 'contacts', 'createdAt'],
  Events:  ['eventId', 'organizerId', 'category', 'title', 'datetime', 'place', 'detail', 'imageUrl', 'joinCount', 'likeCount', 'slots', 'contacts', 'area', 'createdAt'],
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
    aiWrite: aiWrite_
  };
  const publicActions = ['register', 'login', 'resetPassword', 'listEvents', 'listPosts'];
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
  return { userId: u.userId, email: u.email, nickname: u.nickname, area: u.area, isTemp: String(u.isTemp) === 'true' };
}

function validEmail_(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function validPassword_(pw) {
  return typeof pw === 'string' && pw.length >= 8 && /[A-Za-z]/.test(pw) && /[0-9]/.test(pw);
}

function sendMail_(to, subject, body) {
  GmailApp.sendEmail(to, '【' + APP_NAME + '】' + subject, body + '\n\n――――――――――\n' + APP_NAME + '\n' + APP_URL + '\n※このメールは送信専用です。', { name: APP_NAME });
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
    if (findBy_('Users', 'email', email)) throw new Error('このメールアドレスは登録済みです。「パスワードを忘れた」から再発行できます');
    const temp = genTempPassword_();
    const salt = Utilities.getUuid();
    append_('Users', {
      userId: newId_('U'), email: email, passwordHash: hash_(temp, salt), salt: salt,
      isTemp: 'true', area: area, nickname: nickname, createdAt: now_(), failCount: 0, lockedUntil: ''
    });
    sendMail_(email, '仮パスワードのお知らせ',
      nickname + ' さん\n\nご登録ありがとうございます。\n\n仮パスワード： ' + temp +
      '\n\nこの仮パスワードでログインすると、本パスワードの設定画面が表示されます。\n本パスワードを設定するまで、ほかの機能は使えません。');
    return { message: '仮パスワードをメールで送りました' };
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
  const u = findBy_('Users', 'email', email);
  // 登録の有無を外部に漏らさないため、常に同じ応答を返す
  if (u) {
    const temp = genTempPassword_();
    const salt = Utilities.getUuid();
    updateRow_('Users', u._row, { passwordHash: hash_(temp, salt), salt: salt, isTemp: 'true', failCount: 0, lockedUntil: '' });
    sendMail_(email, '仮パスワードの再発行',
      u.nickname + ' さん\n\n新しい仮パスワード： ' + temp +
      '\n\nログイン後、本パスワードを設定してください。\n心当たりがない場合はこのメールを破棄してください。');
  }
  return { message: '登録済みのアドレスであれば、新しい仮パスワードを送りました' };
}

function me_(req, user) {
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
  const nickname = clean_(req.nickname, 30);
  const area = clean_(req.area, 50);
  if (!nickname || !area) throw new Error('ニックネームと地域を入力してください');
  updateRow_('Users', user._row, { nickname: nickname, area: area });
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
  const ev = {
    eventId: newId_('E'), organizerId: user.userId, category: d.category,
    title: clean_(d.title, 60), datetime: new Date(d.datetime).toISOString(), place: clean_(d.place, 100),
    detail: clean_(d.detail, 3000), imageUrl: clean_(d.imageUrl, 300),
    joinCount: 0, likeCount: 0,
    slots: JSON.stringify(sanitizeSlots_(d.slots)),
    contacts: JSON.stringify(sanitizeContacts_(d.contacts)),
    area: clean_(d.area, 50) || user.area, createdAt: now_()
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
    updateRow_('Events', ev._row, {
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
  const post = {
    postId: newId_('P'), userId: user.userId, type: type, category: d.category,
    title: clean_(d.title, 60), detail: clean_(d.detail, 2000), imageUrl: clean_(d.imageUrl, 300),
    status: 'open', area: clean_(d.area, 50) || user.area,
    contacts: JSON.stringify(sanitizeContacts_(d.contacts)), createdAt: now_()
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
