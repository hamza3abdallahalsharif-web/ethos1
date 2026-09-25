/* ==========================================================================
   Ethos International School – Student Council Vote – BACKEND (Google Sheets)
   Paste this whole file into:  Google Sheet > Extensions > Apps Script
   Then Deploy > New deployment > Web app > Execute as: Me > Access: Anyone
   (If you already deployed an older version: Deploy > Manage deployments >
    pencil > Version: New version > Deploy.)

   Sheets used (created automatically):
     Votes      – Student Council ballots
     Devices    – one row per phone per election ("scope" column)
     HeadVotes  – Head Boy / Head Girl ballots (level 1, 2, 3)
     Config     – stores the Head election state (current level + finalists)
   ========================================================================== */

var TEACHER_PASSWORD = 'Eis@Ethos@123';   // <-- change the results password here

function sheet_(name, header) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var s = ss.getSheetByName(name);
  if (!s) { s = ss.insertSheet(name); s.appendRow(header); }
  return s;
}
var DEV_HEAD_ = ['time_cairo', 'device_type', 'device_model', 'os', 'browser', 'screen'];
function withDevHeader_(s, from) {            // adds the device columns to old sheets too
  if (s.getRange(1, from).getValue() === '') s.getRange(1, from, 1, 6).setValues([DEV_HEAD_]);
  return s;
}
function votes_()     { return withDevHeader_(sheet_('Votes',     ['time', 'mode', 'year', 'candidate1', 'candidate2']), 6); }
function headVotes_() { return withDevHeader_(sheet_('HeadVotes', ['time', 'mode', 'round', 'year', 'boy', 'girl']), 7); }
function clip_(v) { return String(v === undefined || v === null ? '' : v).replace(/[\r\n=+@]/g, ' ').slice(0, 80); }
/* the columns saved next to every vote: readable Cairo time + what device/browser was used */
function devCols_(body, now) {
  var d = body.dev || {};
  return [Utilities.formatDate(now, 'Africa/Cairo', 'yyyy-MM-dd HH:mm:ss'),
          clip_(d.type), clip_(d.model), clip_(d.os), clip_(d.browser), clip_(d.screen)];
}
function config_()    { return sheet_('Config',    ['key', 'value']); }
function devices_() {
  var s = sheet_('Devices', ['token', 'year', 'time', 'scope']);
  if (s.getRange(1, 4).getValue() === '') s.getRange(1, 4).setValue('scope');
  return s;
}

/* ---------- Candidates added live by the teacher (admin) ----------
   Stored separately from the built-in roster baked into index.html, so every
   phone/computer just fetches this list and merges it in - no redeploy needed.
   id scheme: "<year><class>-<n>" - same shape as the built-in ids, so every
   existing id regex on the client and here still matches them unchanged.
   n starts at 50 per class and only ever increases (tracked in Config), so a
   newly added candidate can never collide with a built-in id (built-in
   rosters never have more than a handful of names per class) or with an
   older id that was later removed. */
function candidates_() { return sheet_('Candidates', ['id', 'year', 'cls', 'name', 'sex', 'time']); }
function candidateList_() {
  var s = candidates_(), n = s.getLastRow(), out = [];
  if (n > 1) {
    var v = s.getRange(2, 1, n - 1, 5).getValues();
    for (var i = 0; i < v.length; i++) {
      out.push({ id: String(v[i][0]), year: v[i][1], cls: String(v[i][2]), name: String(v[i][3]), sex: String(v[i][4] || 'n') });
    }
  }
  return out;
}
function nextCandNum_(year, cls) {           // collision-safe, never-reused counter per class
  var s = config_(), n = s.getLastRow(), key = 'nid_' + year + cls;
  if (n > 1) {
    var v = s.getRange(2, 1, n - 1, 2).getValues();
    for (var i = 0; i < v.length; i++) {
      if (v[i][0] === key) {
        var num = parseInt(v[i][1], 10) + 1;
        s.getRange(i + 2, 2).setValue(num);
        return num;
      }
    }
  }
  s.appendRow([key, 50]);                    // first dynamic candidate for this class
  return 50;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function doGet() { return json_({ ok: true, app: 'ethos-vote' }); }

function doPost(e) {
  var out;
  try {
    var body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = handle_(String(body.action || ''), body);
  } catch (err) {
    out = { ok: false, error: String(err) };
  }
  return json_(out);
}

/* ---------- device lock (one vote per phone, per election/level) ---------- */
function scopeOf_(v) { return /^(council|head1|head2|head3)$/.test(String(v)) ? String(v) : 'council'; }
function cleanTokens_(body) {
  var res = [], t = body.tokens || [];
  for (var i = 0; i < t.length; i++) {
    var x = String(t[i]);
    if (/^[a-f0-9]{32}$/.test(x) && res.indexOf(x) < 0) res.push(x);
  }
  return res;
}
function deviceSet_(scope) {
  var s = devices_(), n = s.getLastRow(), set = {};
  if (n > 1) {
    var v = s.getRange(2, 1, n - 1, 4).getValues();
    for (var i = 0; i < v.length; i++) {
      if (String(v[i][3] || 'council') === scope) set[String(v[i][0])] = true;
    }
  }
  return set;
}
function lockedToken_(tokens, scope) {
  var set = deviceSet_(scope);
  for (var i = 0; i < tokens.length; i++) if (set[tokens[i]]) return tokens[i];
  return null;
}
function newToken_() { return Utilities.getUuid().replace(/-/g, ''); }
function needPassword_(body) {
  if (String(body.password || '') !== TEACHER_PASSWORD) throw new Error('auth');
}
/* remembers this phone for this scope; returns {locked:true,token} if it already voted */
function takeDevice_(body, year, scope, now) {
  var tokens = cleanTokens_(body), lt = lockedToken_(tokens, scope);
  if (lt) return { locked: true, token: lt };
  var tok = tokens[0] || newToken_(), all = tokens.slice();
  if (all.indexOf(tok) < 0) all.push(tok);
  var d = devices_();
  for (var k = 0; k < all.length; k++) d.appendRow([all[k], year, now, scope]);
  return { locked: false, token: tok };
}
function clearDevices_(keepFn) {
  var s = devices_(), n = s.getLastRow();
  if (n < 2) return;
  var v = s.getRange(2, 1, n - 1, 4).getValues(), keep = [];
  for (var i = 0; i < v.length; i++) if (keepFn(String(v[i][3] || 'council'))) keep.push(v[i]);
  s.deleteRows(2, n - 1);
  if (keep.length) s.getRange(2, 1, keep.length, 4).setValues(keep);
}

/* ---------- Head Boy / Head Girl state ---------- */
function emptyHead_() { return { round: 0, enabled: true, finalists: { b: [], g: [] }, winners: { b: [], g: [] } }; }
function getHead_() {
  var s = config_(), n = s.getLastRow();
  if (n > 1) {
    var v = s.getRange(2, 1, n - 1, 2).getValues();
    for (var i = 0; i < v.length; i++) {
      if (v[i][0] === 'head') { try { return JSON.parse(v[i][1]); } catch (e) {} }
    }
  }
  return emptyHead_();
}
function setHead_(h) {
  var s = config_(), n = s.getLastRow();
  if (n > 1) {
    var v = s.getRange(2, 1, n - 1, 2).getValues();
    for (var i = 0; i < v.length; i++) {
      if (v[i][0] === 'head') { s.getRange(i + 2, 2).setValue(JSON.stringify(h)); return; }
    }
  }
  s.appendRow(['head', JSON.stringify(h)]);
}
/* Head Boy/Girl candidate ids only ever come from Year 10 or 11 */
function cleanIds_(a) {
  var r = [];
  if (!a || !a.length) return r;
  for (var i = 0; i < a.length; i++) {
    var x = String(a[i]);
    if (/^(10|11)[A-D]-\d{1,2}$/.test(x) && r.indexOf(x) < 0) r.push(x);
  }
  return r;
}
function cleanHead_(h) {
  var r = parseInt(h.round, 10);
  if (!(r >= 0 && r <= 4)) return null;
  var f = h.finalists || {}, w = h.winners || {};
  return {
    round: r,
    enabled: h.enabled !== false,
    finalists: { b: cleanIds_(f.b), g: cleanIds_(f.g) },
    winners: { b: cleanIds_(w.b), g: cleanIds_(w.g) }
  };
}

/* is this Head Boy + Head Girl pair legal for this round? */
function headAllowed_(head, round, boy, girl) {
  var re = /^(10|11)[A-D]-\d{1,2}$/;
  if (!re.test(boy) || !re.test(girl) || boy === girl) return false;
  if (round === 1) return true;                                    // open nomination: any Y10/11 student
  if (round === 3) return head.finalists.b.indexOf(boy) >= 0 && head.finalists.g.indexOf(girl) >= 0; // interview finalists only
  return false;
}

function handle_(action, body) {
  if (action === 'ping') return { ok: true, app: 'ethos-vote' };

  if (action === 'status') {
    var tokens = cleanTokens_(body);
    var lt = lockedToken_(tokens, scopeOf_(body.scope));
    return { ok: true, locked: lt !== null, token: lt || tokens[0] || newToken_() };
  }

  /* ------------------------- Student Council vote (Years 4-9) ------------------------- */
  if (action === 'vote') {
    var mode = String(body.mode || ''), year = parseInt(body.year, 10), picks = body.picks || [];
    if (mode !== 'phone' && mode !== 'pc') return { ok: false, error: 'mode' };
    if (mode === 'phone' && (year < 7 || year > 9)) return { ok: false, error: 'year' };
    if (mode === 'pc' && (year < 4 || year > 6)) return { ok: false, error: 'year' };
    if (picks.length < 1 || picks.length > 2) return { ok: false, error: 'picks' };
    if (picks.length === 2 && picks[0] === picks[1]) return { ok: false, error: 'picks' };
    for (var i = 0; i < picks.length; i++) {
      var m = /^(\d{1,2})[A-D]-\d{1,2}$/.exec(String(picks[i]));
      if (!m || parseInt(m[1], 10) !== year) return { ok: false, error: 'candidate' };
    }
    var now = new Date();
    var lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      var tok = null;
      if (mode === 'phone') {
        var dv = takeDevice_(body, year, 'council', now);
        if (dv.locked) return { ok: false, locked: true, token: dv.token };
        tok = dv.token;
      }
      votes_().appendRow([now, mode, year, picks[0], picks[1] || ''].concat(devCols_(body, now)));
      return tok ? { ok: true, token: tok } : { ok: true };
    } finally {
      lock.releaseLock();
    }
  }

  /* ------------------------- Head Boy / Head Girl (Years 10 & 11 only) -------------------------
     round 1 = Year Group nomination (Y10/11 vote, any candidate)
     round 2 = Interview            (teacher-only, no voting)
     round 3 = Students Vote        (whole school votes among interview finalists)
     round 4 = Winners announced                                                          */
  if (action === 'head_state') return { ok: true, head: getHead_() };

  if (action === 'head_vote') {
    var hmode = String(body.mode || ''), round = parseInt(body.round, 10);
    var boy = String(body.boy || ''), girl = String(body.girl || '');
    if (hmode !== 'phone' && hmode !== 'pc') return { ok: false, error: 'mode' };
    var head = getHead_();
    if (head.enabled === false) return { ok: false, error: 'disabled' };
    if (!(round === 1 || round === 3) || round !== head.round) return { ok: false, error: 'round', round: head.round };
    if (round === 1 && hmode !== 'phone') return { ok: false, error: 'mode' };   // nominations are phone-only (Y10/11 students)
    if (!headAllowed_(head, round, boy, girl)) return { ok: false, error: 'candidate' };
    var hyear = parseInt(body.year, 10) || 0;
    var hnow = new Date();
    var hlock = LockService.getScriptLock();
    hlock.waitLock(25000);
    try {
      var htok = null;
      if (hmode === 'phone') {
        var hd = takeDevice_(body, hyear, 'head' + round, hnow);
        if (hd.locked) return { ok: false, locked: true, token: hd.token };
        htok = hd.token;
      }
      headVotes_().appendRow([hnow, hmode, round, hyear, boy, girl].concat(devCols_(body, hnow)));
      return htok ? { ok: true, token: htok } : { ok: true };
    } finally {
      hlock.releaseLock();
    }
  }

  if (action === 'head_set') {
    needPassword_(body);
    var clean = cleanHead_(body.head || {});
    if (!clean) return { ok: false, error: 'head' };
    setHead_(clean);
    return { ok: true };
  }

  if (action === 'head_results') {
    needPassword_(body);
    var hs = headVotes_(), hn = hs.getLastRow();
    var counts = { 1: {}, 2: {}, 3: {} };
    var ballots = { 1: { total: 0, phone: 0, pc: 0 }, 2: { total: 0, phone: 0, pc: 0 }, 3: { total: 0, phone: 0, pc: 0 } };
    if (hn > 1) {
      var hv = hs.getRange(2, 1, hn - 1, 6).getValues();
      for (var q = 0; q < hv.length; q++) {
        var rd = parseInt(hv[q][2], 10);
        if (!counts[rd]) continue;
        var bb = String(hv[q][4]), gg = String(hv[q][5]), md2 = String(hv[q][1]);
        counts[rd][bb] = (counts[rd][bb] || 0) + 1;
        counts[rd][gg] = (counts[rd][gg] || 0) + 1;
        ballots[rd].total++;
        ballots[rd][md2] = (ballots[rd][md2] || 0) + 1;
      }
    }
    return { ok: true, head: getHead_(), counts: counts, ballots: ballots };
  }

  if (action === 'head_reset') {
    needPassword_(body);
    var h1 = headVotes_();
    if (h1.getLastRow() > 1) h1.deleteRows(2, h1.getLastRow() - 1);
    clearDevices_(function (sc) { return sc.indexOf('head') !== 0; });
    setHead_(emptyHead_());
    return { ok: true };
  }

  /* ------------------------- Council results / reset ------------------------- */
  if (action === 'results') {
    needPassword_(body);
    var s = votes_(), n = s.getLastRow();
    var counts2 = {}, b = { total: 0, phone: 0, pc: 0 }, years = {};
    if (n > 1) {
      var v = s.getRange(2, 1, n - 1, 5).getValues();
      for (var r = 0; r < v.length; r++) {
        var md = String(v[r][1]), yr = v[r][2];
        counts2[v[r][3]] = (counts2[v[r][3]] || 0) + 1;
        if (v[r][4]) counts2[v[r][4]] = (counts2[v[r][4]] || 0) + 1;
        b.total++; b[md] = (b[md] || 0) + 1;
        years[yr] = (years[yr] || 0) + 1;
      }
    }
    b.years = years;
    return { ok: true, counts: counts2, ballots: b };
  }

  if (action === 'reset') {           // Student Council only (Head election has its own reset)
    needPassword_(body);
    var s1 = votes_();
    if (s1.getLastRow() > 1) s1.deleteRows(2, s1.getLastRow() - 1);
    clearDevices_(function (sc) { return sc.indexOf('head') === 0; });
    return { ok: true };
  }

  /* ------------------------- Candidates (teacher admin: add/remove live) -------------------------
     "candidates_list" is unauthenticated on purpose - every voting device needs it just to draw
     the ballot, the same way it already needs the built-in roster baked into index.html.          */
  if (action === 'candidates_list') {
    return { ok: true, candidates: candidateList_() };
  }

  if (action === 'candidate_add') {
    needPassword_(body);
    var cYear = parseInt(body.year, 10);
    var cCls = String(body.cls || '').toUpperCase();
    var cName = clip_(body.name).trim();
    if (!(cYear >= 4 && cYear <= 11)) return { ok: false, error: 'year' };
    if (['A', 'B', 'C', 'D'].indexOf(cCls) < 0) return { ok: false, error: 'cls' };
    if (!cName) return { ok: false, error: 'name' };
    var cSex = 'n';
    if (cYear === 10 || cYear === 11) {           // Head Boy / Head Girl years need a gender lane
      cSex = (body.sex === 'b' || body.sex === 'g') ? body.sex : '';
      if (!cSex) return { ok: false, error: 'sex' };
    }
    var cLock = LockService.getScriptLock();
    cLock.waitLock(25000);
    var cId;
    try {
      cId = cYear + cCls + '-' + nextCandNum_(cYear, cCls);
      candidates_().appendRow([cId, cYear, cCls, cName, cSex, new Date()]);
    } finally {
      cLock.releaseLock();
    }
    return { ok: true, id: cId, year: cYear, cls: cCls, name: cName, sex: cSex };
  }

  if (action === 'candidate_remove') {
    needPassword_(body);
    var rId = String(body.id || '');
    var cs = candidates_(), cn = cs.getLastRow();
    if (cn > 1) {
      var cv = cs.getRange(2, 1, cn - 1, 1).getValues();
      for (var ci = 0; ci < cv.length; ci++) {
        if (String(cv[ci][0]) === rId) { cs.deleteRow(ci + 2); return { ok: true }; }
      }
    }
    return { ok: false, error: 'not_found' };
  }

  return { ok: false, error: 'unknown_action' };
}
