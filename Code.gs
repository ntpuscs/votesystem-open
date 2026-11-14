/**
 * 議會表決系統 — Google Apps Script 後端
 * 本系統由國立臺北大學三峽校區學生議會提供並協助架設
 *
 * 使用方式：在 Google 試算表中「擴充功能 > Apps Script」貼上本檔，
 * 先執行一次 setup() 建立工作表，再「部署 > 新增部署作業 > 網頁應用程式」。
 * 詳細步驟請見 README.md。
 */

const SHEET_COUNCIL = '議員';
const SHEET_SETTINGS = '設定';
const SHEET_LOG = '表決紀錄';
const SHEET_DETAIL = '記名明細';
const STATE_KEY = 'VOTE_STATE';
const RULES = ['簡單多數決', '出席過半數', '出席三分之二', '全體過半數', '全體三分之二'];
const ADMIN_TTL_MS = 12 * 60 * 60 * 1000; // 秘書處登入有效 12 小時

/* ───────────── 初始化 ───────────── */

function setup() {
  const ss = SpreadsheetApp.getActive();
  ensureSheet_(ss, SHEET_COUNCIL,
    ['帳號', '密碼', '選區（選填）', '姓名', '排序'],
    [['d01', '請改成密碼', '', '範例議員', 1]]);
  ensureSheet_(ss, SHEET_SETTINGS,
    ['項目', '值'],
    [['秘書處密碼', '請改成密碼'], ['會議名稱', '（會議名稱）']]);
  ensureSheet_(ss, SHEET_LOG,
    ['表決編號', '開始時間', '截止時間', '議案編號', '表決事項', '案別', '方式', '門檻',
     '出席人數', '全體人數', '贊成/同意', '反對/不同意', '棄權', '未投票', '結果', '結果說明']);
  ensureSheet_(ss, SHEET_DETAIL,
    ['表決編號', '議案編號', '表決事項', '選區', '姓名', '投票']);
}

function ensureSheet_(ss, name, header, samples) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold');
    sh.setFrozenRows(1);
    if (samples && samples.length) {
      sh.getRange(2, 1, samples.length, samples[0].length).setValues(samples);
    }
  }
}

/** 緊急用：清除所有進行中狀態（簽到、登入、目前表決）。不影響試算表紀錄。 */
function resetAllState() {
  PropertiesService.getScriptProperties().deleteProperty(STATE_KEY);
  CacheService.getScriptCache().removeAll(['council', 'settings', 'history']);
}

/* ───────────── HTTP 入口 ───────────── */

function doPost(e) {
  let out;
  try {
    const req = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = handle_(req);
  } catch (err) {
    const msg = String((err && err.message) || err);
    out = msg === 'AUTH'
      ? { ok: false, relogin: true, error: '登入已失效，請重新登入' }
      : { ok: false, error: msg };
  }
  out.now = Date.now();
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  if (!params.action) params.action = 'screen';
  return doPost({ postData: { contents: JSON.stringify(params) } });
}

function handle_(req) {
  let st = loadState_();
  // 時間到自動截止（由任何一個請求觸發）
  if (st.vote && st.vote.status === 'voting' && Date.now() >= st.vote.endsAt) {
    withLock_(function () { return null; });
    st = loadState_();
  }

  switch (req.action) {
    /* 公開 */
    case 'screen':
      return screen_(st);

    /* 議員 */
    case 'login':
      return withLock_(function (s) { return login_(s, req); });
    case 'me':
      return me_(st, req.token);
    case 'vote':
      return withLock_(function (s) { return vote_(s, req); });
    case 'logout':
      return withLock_(function (s) { delete s.tokens[req.token]; return { ok: true }; });

    /* 秘書處 */
    case 'adminLogin':
      return withLock_(function (s) { return adminLogin_(s, req); });
    case 'adminLogout':
      return withLock_(function (s) { delete s.admins[req.adminToken]; return { ok: true }; });
    case 'admin':
      requireAdmin_(st, req.adminToken);
      return admin_(st);
    case 'setPresence':
      return withLock_(function (s) {
        requireAdmin_(s, req.adminToken);
        if (!council_().some(function (m) { return m.account === req.account; })) throw new Error('查無此議員');
        s.attendance[req.account] = !!req.present;
        return { ok: true };
      });
    case 'resetAttendance':
      return withLock_(function (s) {
        requireAdmin_(s, req.adminToken);
        if (s.vote && s.vote.status === 'voting') throw new Error('表決進行中，無法清除簽到');
        s.attendance = {};
        s.tokens = {};
        s.vote = null;
        return { ok: true };
      });
    case 'start':
      return withLock_(function (s) { requireAdmin_(s, req.adminToken); return start_(s, req); });
    case 'extend':
      return withLock_(function (s) {
        requireAdmin_(s, req.adminToken);
        if (!s.vote || s.vote.status !== 'voting') throw new Error('目前沒有進行中的表決');
        const sec = Math.min(300, Math.max(5, Number(req.seconds) || 30));
        s.vote.endsAt += sec * 1000;
        return { ok: true };
      });
    case 'stop':
      return withLock_(function (s) {
        requireAdmin_(s, req.adminToken);
        if (!s.vote || s.vote.status !== 'voting') throw new Error('目前沒有進行中的表決');
        closeVote_(s);
        return { ok: true };
      });
    case 'reloadCouncil':
      requireAdmin_(st, req.adminToken);
      CacheService.getScriptCache().removeAll(['council', 'settings', 'history']);
      return { ok: true };

    default:
      throw new Error('未知的操作');
  }
}

/* ───────────── 狀態存取 ───────────── */

function loadState_() {
  const raw = PropertiesService.getScriptProperties().getProperty(STATE_KEY);
  const st = raw ? JSON.parse(raw) : {};
  st.attendance = st.attendance || {}; // 帳號 → true/false
  st.tokens = st.tokens || {};         // 議員登入 token → 帳號
  st.admins = st.admins || {};         // 秘書處 token → 到期時間
  st.vote = st.vote || null;
  return st;
}

function saveState_(st) {
  PropertiesService.getScriptProperties().setProperty(STATE_KEY, JSON.stringify(st));
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const st = loadState_();
    if (st.vote && st.vote.status === 'voting' && Date.now() >= st.vote.endsAt) closeVote_(st);
    const out = fn(st);
    saveState_(st);
    return out;
  } finally {
    lock.releaseLock();
  }
}

/* ───────────── 試算表資料 ───────────── */

function council_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('council');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_COUNCIL);
  if (!sh) throw new Error('找不到「議員」工作表，請先執行 setup()');
  const rows = sh.getDataRange().getDisplayValues().slice(1);
  const list = rows
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r, i) {
      return {
        account: String(r[0]).trim(),
        password: String(r[1]),
        district: String(r[2]).trim(),
        name: String(r[3]).trim(),
        order: Number(r[4]) || i + 1
      };
    });
  list.sort(function (a, b) { return a.order - b.order; });
  cache.put('council', JSON.stringify(list), 120);
  return list;
}

function setting_(key) {
  const cache = CacheService.getScriptCache();
  let map = cache.get('settings');
  if (map) {
    map = JSON.parse(map);
  } else {
    map = {};
    const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_SETTINGS);
    if (sh) {
      sh.getDataRange().getDisplayValues().slice(1).forEach(function (r) {
        if (r[0]) map[String(r[0]).trim()] = String(r[1]);
      });
    }
    cache.put('settings', JSON.stringify(map), 120);
  }
  return map[key] || '';
}

function history_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('history');
  if (hit) return JSON.parse(hit);
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_LOG);
  let out = [];
  if (sh && sh.getLastRow() > 1) {
    const n = Math.min(30, sh.getLastRow() - 1);
    out = sh.getRange(sh.getLastRow() - n + 1, 1, n, 16).getDisplayValues().reverse().map(function (r) {
      return {
        id: r[0], start: r[1], billId: r[3], title: r[4], kind: r[5], mode: r[6], rule: r[7],
        present: r[8], yes: r[10], no: r[11], abs: r[12], notVoted: r[13], result: r[14]
      };
    });
  }
  cache.put('history', JSON.stringify(out), 300);
  return out;
}

/* ───────────── 議員 ───────────── */

function login_(st, req) {
  const account = String(req.account || '').trim();
  const password = String(req.password || '');
  const m = council_().find(function (x) { return x.account === account; });
  if (!m || !password || m.password !== password) throw new Error('帳號或密碼錯誤');
  // 同一帳號只保留最新一台裝置
  Object.keys(st.tokens).forEach(function (t) { if (st.tokens[t] === account) delete st.tokens[t]; });
  const token = Utilities.getUuid();
  st.tokens[token] = account;
  // 登入只代表「已登入」；出席須由秘書處在控制台勾選確認，避免未到場者遠端投票
  return { ok: true, token: token, district: m.district, name: m.name };
}

function me_(st, token) {
  const account = st.tokens[token];
  if (!account) throw new Error('AUTH');
  const m = council_().find(function (x) { return x.account === account; });
  if (!m) throw new Error('AUTH');
  const v = st.vote;
  let vote = null;
  if (v) {
    const c = tally_(v);
    vote = {
      id: v.id, billId: v.billId, title: v.title, kind: v.kind, named: v.named, rule: v.rule,
      needText: needText_(v), endsAt: v.endsAt, status: v.status,
      eligible: v.eligible.indexOf(account) >= 0,
      myChoice: v.votes ? (v.votes[account] || null) : null,
      voted: c.voted, eligibleCount: v.eligible.length
    };
  }
  return {
    ok: true,
    me: { district: m.district, name: m.name },
    present: !!st.attendance[account],
    meeting: setting_('會議名稱'),
    vote: vote
  };
}

function vote_(st, req) {
  const account = st.tokens[req.token];
  if (!account) throw new Error('AUTH');
  const v = st.vote;
  if (!v) throw new Error('目前沒有進行中的表決');
  if (v.status !== 'voting' || Date.now() >= v.endsAt) throw new Error('表決已截止');
  if (v.eligible.indexOf(account) < 0) throw new Error('表決開始時秘書處尚未確認你出席，無法參與本次表決');
  const allowed = v.kind === '人事案' ? ['yes', 'no'] : ['yes', 'no', 'abs'];
  if (allowed.indexOf(req.choice) < 0) throw new Error('選項無效');
  // 確認送出後不可更改（重送相同選項視為成功，避免網路重試誤判）
  if (v.votes[account]) {
    if (v.votes[account] === req.choice) return { ok: true, myChoice: req.choice };
    return { ok: false, error: '你已完成投票，無法更改', myChoice: v.votes[account] };
  }
  v.votes[account] = req.choice;
  return { ok: true, myChoice: req.choice };
}

/* ───────────── 秘書處 ───────────── */

function adminLogin_(st, req) {
  const pw = setting_('秘書處密碼');
  if (!pw || String(req.password || '') !== pw) throw new Error('密碼錯誤');
  const now = Date.now();
  Object.keys(st.admins).forEach(function (t) { if (st.admins[t] < now) delete st.admins[t]; });
  const token = Utilities.getUuid();
  st.admins[token] = now + ADMIN_TTL_MS;
  return { ok: true, adminToken: token };
}

function requireAdmin_(st, token) {
  if (!token || !st.admins[token] || st.admins[token] < Date.now()) throw new Error('AUTH');
}

function admin_(st) {
  const base = screen_(st);
  const online = {};
  Object.keys(st.tokens).forEach(function (t) { online[st.tokens[t]] = true; });
  base.members = council_().map(function (m) {
    return {
      account: m.account, district: m.district, name: m.name,
      present: !!st.attendance[m.account], online: !!online[m.account]
    };
  });
  base.history = history_();
  return base;
}

function start_(st, req) {
  if (st.vote && st.vote.status === 'voting') throw new Error('已有進行中的表決');
  const members = council_();
  const eligible = members
    .filter(function (m) { return st.attendance[m.account]; })
    .map(function (m) { return m.account; });
  if (!eligible.length) throw new Error('目前沒有經確認出席的議員');
  const title = String(req.title || '').trim();
  if (!title) throw new Error('請填寫表決事項');
  const seconds = Math.min(600, Math.max(10, Number(req.seconds) || 60));
  const t = Date.now();
  st.vote = {
    id: Utilities.formatDate(new Date(t), Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss'),
    billId: normBill_(req.billId),
    title: title.slice(0, 500),
    kind: req.kind === '人事案' ? '人事案' : '一般議案',
    named: req.named !== false,
    rule: RULES.indexOf(req.rule) >= 0 ? req.rule : '簡單多數決',
    seconds: seconds,
    startedAt: t,
    endsAt: t + seconds * 1000,
    status: 'voting',
    eligible: eligible,        // 開始表決時經秘書處確認出席者才可投票
    total: members.length,
    votes: {}
  };
  return { ok: true };
}

function normBill_(s) {
  // 議案編號為選填，依各會格式原樣保存
  return String(s || '').trim().slice(0, 60);
}

/* ───────────── 計票 ───────────── */

function labels_(kind) {
  return kind === '人事案'
    ? { yes: '同意', no: '不同意', abs: '棄權' }
    : { yes: '贊成', no: '反對', abs: '棄權' };
}

function need_(v) {
  const p = v.eligible.length;
  const t = v.total;
  switch (v.rule) {
    case '出席過半數': return Math.floor(p / 2) + 1;
    case '出席三分之二': return Math.ceil(p * 2 / 3);
    case '全體過半數': return Math.floor(t / 2) + 1;
    case '全體三分之二': return Math.ceil(t * 2 / 3);
    default: return null; // 簡單多數決
  }
}

function needText_(v) {
  const L = labels_(v.kind);
  const need = need_(v);
  if (need === null) return L.yes + '多於' + L.no + '即通過' + (v.kind === '人事案' ? '' : '，棄權不計');
  return L.yes + ' ≥ ' + need + ' 票';
}

function tally_(v) {
  if (!v.votes && v.counts) return v.counts;
  const c = { yes: 0, no: 0, abs: 0 };
  Object.keys(v.votes || {}).forEach(function (acct) {
    if (v.eligible.indexOf(acct) >= 0 && c[v.votes[acct]] !== undefined) c[v.votes[acct]]++;
  });
  c.voted = c.yes + c.no + c.abs;
  c.notVoted = v.eligible.length - c.voted;
  c.eligible = v.eligible.length;
  return c;
}

function result_(v, c) {
  const L = labels_(v.kind);
  const need = need_(v);
  const simple = need === null;
  const pass = simple ? c.yes > c.no : c.yes >= need;
  const cmp = L.yes + ' ' + c.yes + ' 票、' + L.no + ' ' + c.no + ' 票';
  if (pass) {
    return { text: '通過', tone: 'pass', note: simple ? cmp : L.yes + ' ' + c.yes + ' 票，達門檻 ' + need + ' 票' };
  }
  if (c.yes === c.no) return { text: '可否同數', tone: 'tie', note: '由主席裁決' };
  return { text: '未通過', tone: 'fail', note: simple ? cmp : L.yes + ' ' + c.yes + ' 票，未達門檻 ' + need + ' 票' };
}

function closeVote_(st) {
  const v = st.vote;
  if (!v || v.status !== 'voting') return;
  const c = tally_(v);
  v.status = 'closed';
  v.closedAt = Date.now();
  v.counts = c;
  v.result = result_(v, c);
  writeRecord_(v, c);
  if (!v.named) delete v.votes; // 不記名：截止後立即刪除個別投票
}

function writeRecord_(v, c) {
  const ss = SpreadsheetApp.getActive();
  const tz = Session.getScriptTimeZone();
  const f = function (t) { return Utilities.formatDate(new Date(t), tz, 'yyyy/MM/dd HH:mm:ss'); };
  const L = labels_(v.kind);
  const log = ss.getSheetByName(SHEET_LOG);
  log.appendRow([
    safe_(v.id), f(v.startedAt), f(v.closedAt), safe_(v.billId), safe_(v.title), v.kind,
    v.named ? '記名' : '不記名', v.rule, v.eligible.length, v.total,
    c.yes, c.no, v.kind === '人事案' ? '—' : c.abs, c.notVoted, v.result.text, v.result.note
  ]);
  if (v.named) {
    const detail = ss.getSheetByName(SHEET_DETAIL);
    const rows = council_().map(function (m) {
      let s;
      if (v.eligible.indexOf(m.account) < 0) s = '未出席';
      else s = v.votes[m.account] ? L[v.votes[m.account]] : '未投票';
      return [safe_(v.id), safe_(v.billId), safe_(v.title), m.district, m.name, s];
    });
    if (rows.length) detail.getRange(detail.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  }
  CacheService.getScriptCache().remove('history');
}

/** 避免文字被試算表當成公式或日期 */
function safe_(s) {
  s = String(s == null ? '' : s);
  return /^[=+\-@]/.test(s) || /^\d/.test(s) ? "'" + s : s;
}

/* ───────────── 公開畫面 ───────────── */

function screen_(st) {
  const members = council_();
  const out = {
    ok: true,
    meeting: setting_('會議名稱'),
    total: members.length,
    presentCount: members.filter(function (m) { return st.attendance[m.account]; }).length,
    members: members.map(function (m) {
      return { district: m.district, name: m.name, present: !!st.attendance[m.account] };
    }),
    vote: null
  };
  const v = st.vote;
  if (v) {
    out.vote = {
      id: v.id, billId: v.billId, title: v.title, kind: v.kind, named: v.named, rule: v.rule,
      needText: needText_(v), endsAt: v.endsAt, status: v.status, total: v.total,
      tallies: tally_(v),
      result: v.status === 'closed' ? v.result : null
    };
    if (v.named) {
      out.vote.tiles = members.map(function (m) {
        let s = 'away';
        if (v.eligible.indexOf(m.account) >= 0) s = (v.votes && v.votes[m.account]) || 'none';
        return { district: m.district, name: m.name, s: s };
      });
    }
  }
  return out;
}
