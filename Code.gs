/**
 * సిరిగినేడి ఫ్యామిలీ — Google Apps Script backend
 * - సభ్యుల వివరాలు, పోస్టులు: ఈ Google Sheet లో (Members, Posts, Creds ట్యాబ్‌లు)
 * - ఫోటోలు, వీడియోలు: Google Drive ఫోల్డర్ "Siriginedi Family Media" లో
 * ఈ స్క్రిప్ట్ Google Sheet నుంచే (Extensions > Apps Script) తెరవాలి.
 */
const SN = { members: 'Members', posts: 'Posts', creds: 'Creds' };
const NUM = ['age', 'ord', 'joined', 'ts'];
const BOOL = ['admin', 'added', 'sample'];
const JSN = ['media'];
const FOLDER = 'Siriginedi Family Media';

function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh(c) {
  const name = SN[c] || c;
  let s = ss().getSheetByName(name);
  if (!s) {
    s = ss().insertSheet(name);
    s.getRange(1, 1, s.getMaxRows(), s.getMaxColumns()).setNumberFormat('@');
    s.getRange(1, 1).setValue('id');
    s.setFrozenRows(1);
  }
  return s;
}
function folder() {
  const it = DriveApp.getFoldersByName(FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER);
}
function store(name, mime, b64) {
  const blob = Utilities.newBlob(Utilities.base64Decode(b64), mime, name + '_' + Date.now());
  const f = folder().createFile(blob);
  f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = f.getId();
  return mime.indexOf('video') === 0
    ? 'https://drive.google.com/file/d/' + id + '/preview'
    : 'https://lh3.googleusercontent.com/d/' + id;
}
function saveData(url, name) {
  const m = /^data:([^;]+);base64,([\s\S]*)$/.exec(url);
  return m ? store(name, m[1], m[2]) : url;
}
function prep(d) {
  d = Object.assign({}, d);
  if (typeof d.photo === 'string' && d.photo.indexOf('data:') === 0) d.photo = saveData(d.photo, 'photo');
  if (Array.isArray(d.media)) {
    d.media = d.media.map(function (m) {
      return (m.u && String(m.u).indexOf('data:') === 0) ? { t: m.t, u: saveData(m.u, 'post') } : m;
    });
  }
  return d;
}
function ser(v) {
  if (v === true) return 'true';
  if (v === false || v === null || v === undefined) return '';
  if (typeof v === 'object') return JSON.stringify(v);
  return String(v);
}
function rowToDoc(h, r) {
  const d = {};
  h.forEach(function (k, i) {
    if (i === 0 || r[i] === '' || r[i] === undefined || !k) return;
    let x = r[i];
    if (BOOL.indexOf(k) >= 0) x = (x === 'true' || x === true);
    else if (NUM.indexOf(k) >= 0) x = Number(x);
    else if (JSN.indexOf(k) >= 0) { try { x = JSON.parse(x); } catch (e) { x = []; } }
    else x = String(x);
    d[k] = x;
  });
  return d;
}
function put(c, id, d, merge) {
  const s = sh(c);
  d = prep(d);
  const h = s.getRange(1, 1, 1, s.getLastColumn()).getValues()[0];
  Object.keys(d).forEach(function (k) {
    if (h.indexOf(k) < 0) {
      h.push(k);
      if (h.length > s.getMaxColumns()) s.insertColumnsAfter(s.getMaxColumns(), 10);
      s.getRange(1, h.length).setValue(k);
      s.getRange(1, h.length, s.getMaxRows(), 1).setNumberFormat('@');
    }
  });
  const last = s.getLastRow();
  const ids = s.getRange(1, 1, last, 1).getValues().map(function (r) { return String(r[0]); });
  const ri = ids.indexOf(String(id));
  let row = new Array(h.length).fill('');
  if (ri > 0 && merge) {
    const old = s.getRange(ri + 1, 1, 1, h.length).getValues()[0];
    row = old.map(function (x) { return x === undefined ? '' : x; });
  }
  row[0] = id;
  Object.keys(d).forEach(function (k) { row[h.indexOf(k)] = ser(d[k]); });
  if (ri <= 0 && last + 1 > s.getMaxRows()) s.insertRowsAfter(s.getMaxRows(), 50);
  const rng = s.getRange(ri > 0 ? ri + 1 : last + 1, 1, 1, h.length);
  rng.setNumberFormat('@').setValues([row]);
}
function del(c, id) {
  const s = sh(c);
  const ids = s.getRange(1, 1, s.getLastRow(), 1).getValues().map(function (r) { return String(r[0]); });
  const ri = ids.indexOf(String(id));
  if (ri > 0) s.deleteRow(ri + 1);
}

// ---------- పాస్‌వర్డ్‌లు (Creds ట్యాబ్; క్లయింట్‌కు ఎప్పుడూ పంపము) ----------
function credFind(phone) {
  const s = sh('creds');
  if (s.getLastColumn() < 4) { s.getRange(1, 1, 1, 4).setValues([['phone', 'pw', 'q', 'a']]); }
  const v = s.getDataRange().getValues();
  for (let i = 1; i < v.length; i++) if (String(v[i][0]) === String(phone)) return { s: s, i: i + 1, r: v[i] };
  return { s: s, i: 0, r: null };
}
function credSet(phone, pw, q, a) {
  const f = credFind(phone);
  const row = [phone, pw, q, a];
  if (!f.i && f.s.getLastRow() + 1 > f.s.getMaxRows()) f.s.insertRowsAfter(f.s.getMaxRows(), 50);
  const rng = f.s.getRange(f.i || f.s.getLastRow() + 1, 1, 1, 4);
  rng.setNumberFormat('@').setValues([row]);
}

// ---------- API ----------
function act(p) {
  const a = p.a;
  if (a === 'list') {
    if (!SN[p.c] || p.c === 'creds') throw new Error('bad collection');
    const v = sh(p.c).getDataRange().getValues(), h = v[0];
    let docs = v.slice(1).filter(function (r) { return r[0] !== ''; })
      .map(function (r) { return { id: String(r[0]), d: rowToDoc(h, r) }; });
    if (p.c === 'posts') {
      docs.sort(function (x, y) { return (y.d.ts || 0) - (x.d.ts || 0); });
      if (p.limit) docs = docs.slice(0, p.limit);
    }
    return { docs: docs };
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try {
    if (a === 'upload') return { url: store(p.name || 'file', p.mime || 'application/octet-stream', p.data) };
    if (a === 'set' || a === 'update' || a === 'add' || a === 'delete') {
      if (p.c !== 'members' && p.c !== 'posts') throw new Error('bad collection');
      if (a === 'set') put(p.c, p.id, p.d, false);
      else if (a === 'update') put(p.c, p.id, p.d, true);
      else if (a === 'add') { const id = Utilities.getUuid(); put(p.c, id, p.d, false); return { id: id }; }
      else del(p.c, p.id);
      return { ok: true };
    }
    if (a === 'cred_set') {
      if (credFind(p.phone).i) throw new Error('exists');
      credSet(p.phone, p.pw, p.q, p.a);
      return { ok: true };
    }
    if (a === 'login') { const f = credFind(p.phone); return { ok: !!f.i && String(f.r[1]) === String(p.pw) }; }
    if (a === 'q') { const f = credFind(p.phone); return { q: f.i ? String(f.r[2]) : '' }; }
    if (a === 'reset') {
      const f = credFind(p.phone);
      if (!f.i || String(f.r[3]) !== String(p.a)) return { ok: false };
      credSet(p.phone, p.pw, f.r[2], f.r[3]);
      return { ok: true };
    }
    throw new Error('unknown action');
  } finally {
    lock.releaseLock();
  }
}
function doPost(e) {
  let r;
  try { r = act(JSON.parse(e.postData.contents)); } catch (x) { r = { err: String(x) }; }
  return ContentService.createTextOutput(JSON.stringify(r)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return ContentService.createTextOutput('OK'); }

// ---------- ఒక్కసారి రన్ చేయండి: షీట్లు + అడ్మిన్ మాత్రమే (నమూనా సభ్యులు లేరు) ----------
function setup() {
  sh('members'); sh('posts'); credFind('x');
  const J = Date.now();
  put('members', 'SFM01', {
    name: 'సిరిగినేడి అడ్మిన్', surname: 'సిరిగినేడి', gender: 'M', age: 40, phone: '9100308833',
    father: 'సిరిగినేడి సుబ్బారావు', grandfather: 'సిరిగినేడి కృష్ణమూర్తి',
    village: 'నమూనా గ్రామం', mandal: 'నమూనా మండలం', admin: true, joined: J
  }, false);
  // పాస్‌వర్డ్ 123456 యొక్క హ్యాష్ ("sfm|" + పాస్‌వర్డ్ యొక్క SHA-256)
  credSet('9100308833', 'b0d57c19427665b4002a2b5727f830ee3d77e366f61824b4f912fd287a5282ea', 'మీ ఇంటి పేరు?',
    hashOf('సిరిగినేడి'));
}
// ఇప్పటికే షీట్‌లో ఉన్న నమూనా సభ్యులను (sample=true) ఒక్కసారి రన్ చేసి తొలగించండి
function clearSamples() {
  const s = sh('members');
  const v = s.getDataRange().getValues(), h = v[0], si = h.indexOf('sample');
  if (si < 0) return;
  for (let i = v.length - 1; i >= 1; i--) if (String(v[i][si]) === 'true') s.deleteRow(i + 1);
}
function hashOf(t) {
  const b = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, 'sfm|' + t.trim().replace(/\s+/g, ' ').toLowerCase(), Utilities.Charset.UTF_8);
  return b.map(function (x) { return ('0' + (x < 0 ? x + 256 : x).toString(16)).slice(-2); }).join('');
}
