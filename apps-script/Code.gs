/**
 * Urusan Keluarga WANZA: sambungan Google Sheet untuk app di GitHub Pages.
 * Pasang dalam Sheet keluarga: Extensions > Apps Script, tampal fail ini, Simpan,
 * jalankan setup() sekali, kemudian Deploy > New deployment > Web app
 * (Execute as: Me, Who has access: Anyone).
 *
 * Data kekal dalam Sheet ini. App hanya dapat apa yang dibenarkan oleh PIN setiap ahli.
 */

var CONFIG = {
  ADMIN: 'Faizat',
  DOCS_FOLDER: '',   // ID folder Drive untuk dokumen keluarga
  RESIT_FOLDER: '',  // ID folder Drive untuk resit bayaran
  MEMBERS: ['Sanusi', 'Mahaiza', 'Faris', 'Faizat', 'Fahmi', 'Putri Auni'],
  TOKEN_DAYS: 60
};

var SS = function () { return SpreadsheetApp.getActive(); };

/* ---------- pemasangan ---------- */
function setup() {
  var p = PropertiesService.getScriptProperties();
  if (!p.getProperty('SECRET')) p.setProperty('SECRET', Utilities.getUuid() + Utilities.getUuid());
  var ss = SS();
  var ak = ss.getSheetByName('Akses');
  if (!ak) {
    ak = ss.insertSheet('Akses');
    ak.appendRow(['Ahli', 'Peranan', 'PIN sementara', 'PIN hash', 'Emel untuk lupa PIN', 'Kemaskini']);
    CONFIG.MEMBERS.forEach(function (m) {
      ak.appendRow([m, m === CONFIG.ADMIN ? 'Admin' : 'Ahli', "'" + randPin(), '', '', new Date()]);
    });
    ak.setFrozenRows(1);
    ak.getRange('A1:F1').setFontWeight('bold').setBackground('#2b2d42').setFontColor('#ffffff');
    ak.setColumnWidths(1, 6, 160);
  }
  if (!ss.getSheetByName('Aset')) {
    var a = ss.insertSheet('Aset');
    a.appendRow(['ID', 'Jenis', 'Pemilik', 'Nama / butiran', 'No. plat / rujukan', 'Tarikh tamat / perlu baharu',
      'Servis seterusnya (tarikh)', 'Servis seterusnya (km)', 'Odometer semasa (km)', 'Simpanan dokumen', 'Pautan fail', 'Nota', 'Kemaskini']);
    a.setFrozenRows(1);
    a.getRange('A1:M1').setFontWeight('bold').setBackground('#3d6dff').setFontColor('#ffffff');
  }
  if (!ss.getSheetByName('Emas & Zakat')) {
    var e = ss.insertSheet('Emas & Zakat');
    e.appendRow(['Ahli', 'Simpanan emas (gram)', 'Harga emas semasa (RM/gram)', 'Tarikh mula haul', 'Zakat dibayar (RM)', 'Tarikh bayar zakat', 'Tempat simpan', 'Nota', 'Kemaskini']);
    CONFIG.MEMBERS.forEach(function (m) { e.appendRow([m, '', '', '', '', '', '', '', '']); });
    e.setFrozenRows(1);
    e.getRange('A1:I1').setFontWeight('bold').setBackground('#f0b400').setFontColor('#2b2d42');
  }
  Logger.log('Siap. PIN sementara setiap ahli ada dalam tab Akses.');
}

function randPin() { return String(Math.floor(100000 + Math.random() * 900000)); }

/* ---------- pintu masuk ---------- */
function doGet() {
  return out({ ok: true, app: 'Urusan Keluarga WANZA' });
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents || '{}');
    if (b.act === 'login') return out(login(b.who, String(b.pin || '')));
    if (b.act === 'forgot') return out(forgot(b.who));
    var u = verify(b.tok);
    if (!u) return out({ ok: false, err: 'auth' });
    if (b.act === 'setpin') return out(setPin(u, String(b.oldPin || ''), String(b.pin || '')));
    if (b.act === 'resetpin') return out(adminReset(u, b.who));
    if (b.act === 'call') return out({ ok: true, r: callTool(u, b.server, b.tool, b.input || {}) });
    if (b.act === 'me') return out({ ok: true, who: u.who, role: u.role });
    return out({ ok: false, err: 'act' });
  } catch (x) {
    return out({ ok: false, err: String(x && x.message || x) });
  }
}

function out(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

/* ---------- PIN ---------- */
function akRows() {
  var sh = SS().getSheetByName('Akses');
  var v = sh.getRange(1, 1, sh.getLastRow(), 6).getDisplayValues();
  return { sh: sh, v: v };
}
function findMember(who) {
  var a = akRows();
  for (var i = 1; i < a.v.length; i++) if (a.v[i][0].toLowerCase() === String(who || '').toLowerCase()) return { sh: a.sh, row: i + 1, r: a.v[i] };
  return null;
}
function hash(who, pin) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret + '|' + who + '|' + pin);
  return Utilities.base64Encode(raw);
}
function login(who, pin) {
  var c = CacheService.getScriptCache(), k = 'f_' + who, n = Number(c.get(k) || 0);
  if (n >= 5) return { ok: false, err: 'locked' };
  var m = findMember(who);
  if (!m || !/^\d{4,8}$/.test(pin)) { c.put(k, n + 1, 900); return { ok: false, err: 'pin' }; }
  var temp = m.r[2].replace(/^'/, ''), h = m.r[3];
  var good = (h && h === hash(m.r[0], pin)) || (!h && temp && temp === pin) || (h && temp && temp === pin);
  if (!good) { c.put(k, n + 1, 900); return { ok: false, err: 'pin' }; }
  c.remove(k);
  var role = m.r[1] === 'Admin' ? 'admin' : 'ahli';
  return { ok: true, tok: sign(m.r[0], role), who: m.r[0], role: role, mustChange: !!(temp && temp === pin) };
}
function setPin(u, oldPin, pin) {
  if (!/^\d{6}$/.test(pin)) return { ok: false, err: 'format' };
  var m = findMember(u.who);
  if (!m) return { ok: false, err: 'who' };
  var temp = m.r[2].replace(/^'/, ''), h = m.r[3];
  var good = (h && h === hash(m.r[0], oldPin)) || (temp && temp === oldPin);
  if (!good) return { ok: false, err: 'pin' };
  m.sh.getRange(m.row, 3, 1, 2).setValues([['', hash(m.r[0], pin)]]);
  m.sh.getRange(m.row, 6).setValue(new Date());
  return { ok: true };
}
function forgot(who) {
  var m = findMember(who);
  if (!m) return { ok: true };
  var c = CacheService.getScriptCache(), k = 'r_' + m.r[0];
  if (c.get(k)) return { ok: true, sent: 'wait' };
  c.put(k, '1', 600);
  var pin = randPin();
  m.sh.getRange(m.row, 3).setValue("'" + pin);
  m.sh.getRange(m.row, 6).setValue(new Date());
  var mail = m.r[4], adminMail = Session.getEffectiveUser().getEmail();
  var to = /@/.test(mail) ? mail : adminMail;
  MailApp.sendEmail(to, 'PIN sementara Urusan Keluarga WANZA',
    'PIN sementara untuk ' + m.r[0] + ': ' + pin + '\n\nMasuk dengan PIN ini, kemudian tetapkan PIN baru.\n' +
    (to === adminMail && m.r[0] !== CONFIG.ADMIN ? '\n(' + m.r[0] + ' tiada emel dalam tab Akses. Sila sampaikan PIN ini secara peribadi.)' : ''));
  return { ok: true, sent: to === mail ? 'self' : 'admin' };
}
function adminReset(u, who) {
  if (u.role !== 'admin') return { ok: false, err: 'role' };
  var m = findMember(who);
  if (!m) return { ok: false, err: 'who' };
  var pin = randPin();
  m.sh.getRange(m.row, 3, 1, 2).setValues([["'" + pin, '']]);
  m.sh.getRange(m.row, 6).setValue(new Date());
  return { ok: true, pin: pin };
}

/* ---------- token ---------- */
function sign(who, role) {
  var exp = Date.now() + CONFIG.TOKEN_DAYS * 864e5, body = who + '|' + role + '|' + exp;
  return Utilities.base64EncodeWebSafe(body) + '.' + mac(body);
}
function mac(body) {
  var secret = PropertiesService.getScriptProperties().getProperty('SECRET');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(body, secret));
}
function verify(tok) {
  if (!tok || tok.indexOf('.') < 0) return null;
  var p = tok.split('.'), body = Utilities.newBlob(Utilities.base64DecodeWebSafe(p[0])).getDataAsString();
  if (mac(body) !== p[1]) return null;
  var f = body.split('|');
  if (Number(f[2]) < Date.now()) return null;
  var m = findMember(f[0]);
  if (!m) return null;
  return { who: f[0], role: (m.r[1] === 'Admin' ? 'admin' : 'ahli') };
}

/* ---------- Sheet & Drive ---------- */
var HUTANG_HIDE = [9, 10, 17, 18, 19, 20, 21, 22, 23, 24, 25, 33, 34, 35, 36, 37, 38, 39];
var BAYAR_HIDE = [5, 10, 11, 12];

function callTool(u, server, tool, input) {
  var admin = u.role === 'admin';
  if (server === 'Google Sheets') {
    var range = String(input.range || '');
    var tab = range.split('!')[0].replace(/'/g, '');
    if (tab === 'Akses') throw new Error('denied');
    if (tool === 'get_values') return { values: readFor(u, range, tab) };
    if (tool === 'update_values') {
      if (!admin && !(tab === 'Emas & Zakat' && ownEmasRow(u, range))) throw new Error('denied');
      writeRange(range, input.values);
      return { ok: true };
    }
    if (tool === 'append_values') {
      if (!admin && !(tab === 'Semakan' && input.values.every(function (r) { return String(r[0]) === u.who; }))) throw new Error('denied');
      appendRows(tab, input.values);
      return { ok: true };
    }
    throw new Error('tool');
  }
  if (server === 'Google Drive' && tool === 'create_file') {
    var folder = input.parentId === CONFIG.RESIT_FOLDER && admin ? CONFIG.RESIT_FOLDER : CONFIG.DOCS_FOLDER;
    if (!folder) throw new Error('folder');
    var blob = Utilities.newBlob(Utilities.base64Decode(input.base64Content), input.contentMimeType, input.title);
    var f = DriveApp.getFolderById(folder).createFile(blob);
    return { id: f.getId(), viewUrl: f.getUrl() };
  }
  throw new Error('server');
}

function readFor(u, range, tab) {
  var sh = SS().getSheetByName(tab);
  if (!sh) return [];
  var v = sh.getRange(range.split('!')[1] ? range : tab + '!A1:Z').getDisplayValues();
  while (v.length && v[v.length - 1].join('') === '') v.pop();
  if (u.role === 'admin') return v;
  var me = u.who.toLowerCase();
  var mine = function (s) { return String(s || '').toLowerCase().split(/&|,|\|/).map(function (x) { return x.trim(); }).indexOf(me) >= 0; };
  if (tab === 'Hutang') return v.map(function (r, i) { return i ? r.map(function (c, k) { return HUTANG_HIDE.indexOf(k) >= 0 ? '' : c; }) : r; });
  if (tab === 'Bayaran') return v.map(function (r, i) { return i ? r.map(function (c, k) { return BAYAR_HIDE.indexOf(k) >= 0 ? '' : c; }) : r; });
  if (tab === 'Keluarga') return v.map(function (r, i) {
    if (!i || String(r[1]).toLowerCase() === me) return r;
    return r.map(function (c, k) { return [0, 1, 5, 6, 7, 19].indexOf(k) >= 0 ? c : ''; });
  });
  if (tab === 'Semakan') return v.filter(function (r, i) { return !i || mine(r[0]); });
  if (tab === 'Aset') return v.filter(function (r, i) { return !i || mine(r[2]) || /keluarga|semua/i.test(r[2]); });
  if (tab === 'Emas & Zakat') return v.filter(function (r, i) { return !i || mine(r[0]); });
  return [];
}

function ownEmasRow(u, range) {
  var m = /![A-Z]+(\d+)(?::[A-Z]+(\d+))?$/.exec(range);
  if (!m || (m[2] && m[2] !== m[1])) return false;
  var row = Number(m[1]), sh = SS().getSheetByName('Emas & Zakat');
  return row > 1 && String(sh.getRange(row, 1).getDisplayValue()).toLowerCase() === u.who.toLowerCase();
}

function writeRange(range, values) {
  var tab = range.split('!')[0].replace(/'/g, ''), sh = SS().getSheetByName(tab);
  var start = sh.getRange(range.split('!')[1].split(':')[0]);
  var r0 = start.getRow(), c0 = start.getColumn();
  values.forEach(function (row, i) {
    var seg = null;
    for (var k = 0; k <= row.length; k++) {
      var val = k < row.length ? row[k] : null;
      if (val !== null && val !== undefined) {
        if (!seg) seg = { c: k, vals: [] };
        seg.vals.push(val);
      } else if (seg) {
        sh.getRange(r0 + i, c0 + seg.c, 1, seg.vals.length).setValues([seg.vals]);
        seg = null;
      }
    }
  });
}

function appendRows(tab, values) {
  var sh = SS().getSheetByName(tab);
  var col = sh.getRange('A1:A').getDisplayValues(), last = 0;
  for (var i = 0; i < col.length; i++) if (col[i][0] !== '') last = i + 1;
  writeRange(tab + '!A' + (last + 1), values);
}
