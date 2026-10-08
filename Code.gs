/*************************************************************
 * STUDENT MANAGEMENT SYSTEM  -  Google Apps Script Backend
 * Dr. Punjabrao Deshmukh Polytechnic, Amravati
 *
 * Backed entirely by one Google Sheet (multiple tabs).
 * Web based - works on mobile and laptop.
 *************************************************************/

// ====== CONFIG ======
// You can paste EITHER the bare Sheet ID, OR the full Sheet URL — both work.
const SPREADSHEET_ID = '1w2q_9AjVCacdx9IS_K-Ul9o0m721HKwv34GCSkPZ3QE'; // <-- paste the Sheet ID or URL
const APP_TITLE = 'Student Management System - Dr. P. D. Polytechnic, Amravati';

// Logo (base64) used in the report-card PDF header.
const LOGO_DATA_URI = 'data:image/jpeg;base64,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';

// Fixed signatures shown on every report card.
const SIG_LEFT_NAME = 'R.H.Rathod,';
const SIG_LEFT_ROLE = 'Convener CO';
const SIG_RIGHT_NAME = 'Dr.G.R.Gosavi,';
const SIG_RIGHT_ROLE = 'Principal';

// Extracts the ID if a full URL was pasted; otherwise returns the value as-is.
function _sheetId(){
  const v = String(SPREADSHEET_ID).trim();
  const m = v.match(/\/d\/([a-zA-Z0-9-_]+)/);
  return m ? m[1] : v;
}
function _openSS(){ return SpreadsheetApp.openById(_sheetId()); }


// Sheet (tab) names
const SH_USERS    = 'Users';     // login accounts
const SH_STUDENTS = 'Students';
const SH_CLASSES  = 'Classes';
const SH_SUBJECTS = 'Subjects';  // subject + teacher mapping
const SH_RESULTS  = 'Results';   // marks
const SH_SESSIONS = 'Sessions';   // attendance sessions
const SH_ATTMARKS = 'AttMarks';   // attendance P/A/L per student per session
const SH_PLAN = 'PlannedSessions'; // pre-planned lecture dates per timetable (filled in advance)
const SH_TEACHER_SUBJ = 'TeacherSubjects'; // which subjects each teacher account is allotted to
const SH_ATTSUMMARY = 'AttSummary'; // uploaded per-subject attendance summaries (any class)

// ====== Attendance module config ======
const ATT_TOKEN_TTL_MS = 12*60*60*1000;


const INSTITUTE_NAME = 'Dr.Panjabrao Deshmukh Polytechnic, Amravati';
const DEPT_NAME      = 'Computer Engineering Dept';
const CLASS_NAME     = 'Fifth Semester Computer Engineering (CO-5-K)';
const PASS_THRESHOLD = 75;
const SH_ATT_SUBJECTS = 'AttendanceSubjects'; // admin-managed attendance subject list (semester-wise)

// Existing sheets created before this feature won't have this tab yet — add it once, seeded with
// the subjects this app originally shipped with, so existing attendance data keeps working.
// Creates the subject sheet if missing. No subjects are auto-added — the admin fully owns this
// list via the Subjects tab, starting empty.
function _ensureAttSubjectsSheet(){
  const ss=_openSS();
  let sh=ss.getSheetByName(SH_ATT_SUBJECTS);
  if(!sh){
    sh=ss.insertSheet(SH_ATT_SUBJECTS);
    sh.appendRow(['subjectKey','subjectName','subjectCode','type','semester']);
  }
  return sh;
}

// Loads the live, admin-editable subject list in the same shape the old hardcoded SUBJECTS
// constant used: { key: {name, code, type, semester} }. Called fresh wherever subjects are needed,
// so admin edits take effect immediately with no redeploy.
function _subjects(){
  _ensureAttSubjectsSheet();
  const out={};
  readTable(SH_ATT_SUBJECTS).forEach(r=>{
    const key=String(r.subjectKey||'').trim();
    if(!key) return;
    out[key] = { name:String(r.subjectName||'').trim(), code:String(r.subjectCode||'').trim(),
                 type:(String(r.type||'TH').trim().toUpperCase()==='PR'?'PR':'TH'),
                 semester:String(r.semester||'').trim() };
  });
  return out;
}
const ATT_LOGO = 'data:image/jpeg;base64,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';

// ---------- Entry point ----------
// Adds a "Setup" menu to the Google Sheet so you can run setup with one click
// (no need to use the Apps Script function dropdown).
function onOpen(){
  try{
    SpreadsheetApp.getUi()
      .createMenu('⚙ App Setup')
      .addItem('1. Create tabs + admin (setupSheets)', 'setupSheets')
      .addSeparator()
      .addItem('Fix admin login (resetAdmin)', 'resetAdmin')
      .addItem('Clear PA Test Marks data (keeps attendance)', 'clearPatMarksData')
      .addToUi();
  }catch(e){}
}

function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle(APP_TITLE)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ---------- One time setup: creates tabs + headers + default admin ----------
function setupSheets() {
  const ss = _openSS();
  const need = {
    [SH_USERS]:    ['username','password','role','name'],
    [SH_STUDENTS]: ['enrollment','name','className','rollNo','batch','email','mobile','spi'],
    [SH_CLASSES]:  ['className','description','academicYear','season'],
    [SH_SUBJECTS]: ['className','subjectCode','subjectName','teacher','maxMarks'],
    [SH_RESULTS]:  ['enrollment','className','subjectCode','marks','examName'],
    [SH_SESSIONS]: ['sessionId','subject','date','time','batch','topic','takenBy','createdAt'],
    [SH_ATTMARKS]: ['sessionId','roll','status'],
    [SH_PLAN]: ['planId','subject','date','time','batch','status','skipReason','sessionId','createdAt'],
    [SH_TEACHER_SUBJ]: ['username','subject'],
    [SH_ATT_SUBJECTS]: ['subjectKey','subjectName','subjectCode','type','semester'],
    [SH_ATTSUMMARY]: ['enrollment','className','subject','present','total','pct']
  };
  Object.keys(need).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) sh.appendRow(need[name]);
  });
  // migrate: add any missing header columns to existing sheets
  Object.keys(need).forEach(name => {
    const sh = ss.getSheetByName(name);
    const head = sh.getRange(1, 1, 1, sh.getLastColumn() || 1).getValues()[0];
    need[name].forEach((col, i) => {
      if (head[i] !== col){
        if (i >= head.length) sh.getRange(1, i + 1).setValue(col);
      }
    });
  });
  // default admin if Users empty
  const u = ss.getSheetByName(SH_USERS);
  if (u.getLastRow() === 1) {
    u.appendRow(['admin', 'admin123', 'admin', 'Administrator']);
  }
  // remove default Sheet1 if present and empty-ish
  const s1 = ss.getSheetByName('Sheet1');
  if (s1 && ss.getSheets().length > 1) { try { ss.deleteSheet(s1); } catch (err) {} }
  return 'Setup complete. Default admin -> username: admin , password: admin123. Next: log in and upload your class list from Attendance -> Students & Batches.';
}

/**
 * Clears ONLY the PA Test Marks data (Results, Subjects, Classes tabs).
 * - Header rows are preserved so the tabs remain valid.
 * - The Students, Users, Sessions and AttMarks tabs are NOT touched, so the
 *   student list, logins and all attendance data remain fully intact.
 * When run from the "App Setup" menu it asks for confirmation first.
 */
function clearPatMarksData() {
  const ss = _openSS();
  const targets = [SH_RESULTS, SH_SUBJECTS, SH_CLASSES];

  // Confirmation prompt when run from the Sheet menu (UI available).
  let ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) { ui = null; }
  if (ui) {
    const resp = ui.alert(
      'Clear PA Test Marks data?',
      'This will delete all data in the Results, Subjects and Classes tabs (marks / report-card data).\n\n' +
      'Your student list, logins and ALL attendance data will NOT be affected.\n\n' +
      'This cannot be undone. Continue?',
      ui.ButtonSet.YES_NO);
    if (resp !== ui.Button.YES) return 'Cancelled. Nothing was deleted.';
  }

  const cleared = [];
  targets.forEach(function (name) {
    const sh = ss.getSheetByName(name);
    if (!sh) { cleared.push(name + ': tab not found (skipped)'); return; }
    const last = sh.getLastRow();
    if (last > 1) {
      sh.getRange(2, 1, last - 1, sh.getLastColumn()).clearContent();
      cleared.push(name + ': ' + (last - 1) + ' rows cleared');
    } else {
      cleared.push(name + ': already empty');
    }
  });

  const msg = 'PA Test Marks data cleared (headers kept). ' + cleared.join('; ') +
              '. Attendance and student list are untouched.';
  if (ui) ui.alert('Done', msg, ui.ButtonSet.OK);
  return msg;
}

// ---------- Run this if admin login fails. Shows what the server actually sees. ----------
function diagnoseLogin() {
  const ss = _openSS();
  const sh = ss.getSheetByName(SH_USERS);
  if (!sh) return 'PROBLEM: Users tab does not exist. Run setupSheets first.';
  const rows = sh.getDataRange().getValues();
  let out = 'Users tab has ' + rows.length + ' row(s) (including header).\n';
  rows.forEach((r, i) => {
    out += 'Row ' + (i+1) + ': username="' + r[0] + '" password="' + r[1]
         + '" role="' + r[2] + '"\n';
  });
  out += '\nTest login admin/admin123 -> ' + JSON.stringify(login('admin','admin123'));
  Logger.log(out);
  return out;
}

// ---------- Run this to (re)create or reset the admin account ----------
function resetAdmin() {
  const ss = _openSS();
  let sh = ss.getSheetByName(SH_USERS);
  if (!sh) { sh = ss.insertSheet(SH_USERS); sh.appendRow(['username','password','role','name']); }
  const data = sh.getDataRange().getValues();
  // remove any existing admin rows
  for (let i = data.length - 1; i >= 1; i--) {
    if (String(data[i][0]).trim() === 'admin') sh.deleteRow(i+1);
  }
  // store password as plain text (force text format so it is never coerced)
  const r = sh.getLastRow() + 1;
  sh.getRange(r, 1, 1, 4).setNumberFormat('@').setValues([['admin','admin123','admin','Administrator']]);
  return 'Admin reset. username: admin , password: admin123';
}

// ---------- Helpers ----------
function getSheet(name){
  const sh=_openSS().getSheetByName(name);
  if(!sh) throw new Error('Sheet "'+name+'" was not found in the spreadsheet. Check the tab name, or run setupSheets from the Apps Script editor.');
  return sh;
}

function readTable(name){
  const sh = getSheet(name);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0];
  return values.slice(1).map(r => {
    const o = {};
    head.forEach((h,i)=> o[h]=r[i]);
    return o;
  });
}

// ====== AUTH ======
// Token format: base64(username|role|expiryMillis) + "." + signature
// Signature = HMAC-SHA256 of the payload using a server-only secret.
// The client cannot forge or alter it, so the server trusts only the token.

const TOKEN_TTL_MS = 8 * 60 * 60 * 1000; // 8 hours

function _secret(){
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty('AUTH_SECRET');
  if (!s){ s = Utilities.getUuid() + Utilities.getUuid(); props.setProperty('AUTH_SECRET', s); }
  return s;
}

function _sign(payload){
  const raw = Utilities.computeHmacSha256Signature(payload, _secret());
  return Utilities.base64EncodeWebSafe(raw);
}

function _makeToken(username, role){
  const payload = Utilities.base64EncodeWebSafe(
    username + '|' + role + '|' + (Date.now() + TOKEN_TTL_MS));
  return payload + '.' + _sign(payload);
}

// Returns {username, role} if valid, else null.
function _verifyToken(token){
  if (!token || token.indexOf('.') < 0) return null;
  const parts = token.split('.');
  const payload = parts[0], sig = parts[1];
  if (_sign(payload) !== sig) return null;                 // tampered or forged
  const decoded = Utilities.newBlob(Utilities.base64DecodeWebSafe(payload)).getDataAsString();
  const bits = decoded.split('|');
  if (bits.length !== 3) return null;
  if (Date.now() > Number(bits[2])) return null;           // expired
  return { username: bits[0], role: bits[1] };
}

function _requireAdmin(token){
  const s = _verifyToken(token);
  if (!s || s.role !== 'admin') throw new Error('Not authorized. Please log in as admin.');
  return s;
}

// Generic auth (any valid login) and staff-only (non-student) guards for the attendance module.
function _require(token){ const s=_verifyToken(token); if(!s) throw new Error('Session expired. Please log in again.'); return s; }
function _requireStaff(token){ const s=_require(token); if(s.role==='student') throw new Error('Not authorized.'); return s; }

// Existing sheets created before multi-teacher support won't have this tab yet — add it once.
function _ensureTeacherSheet(){
  const ss=_openSS();
  let sh=ss.getSheetByName(SH_TEACHER_SUBJ);
  if(!sh){ sh=ss.insertSheet(SH_TEACHER_SUBJ); sh.appendRow(['username','subject']); }
  return sh;
}

// The subject keys this logged-in user may take attendance / see reports for.
// Admins can always act on every subject; teachers are limited to what's been allotted to them.
function _mySubjects(token){
  const s=_requireStaff(token);
  if(s.role==='admin') return Object.keys(_subjects());
  _ensureTeacherSheet();
  return readTable(SH_TEACHER_SUBJ)
    .filter(r=>String(r.username).trim()===String(s.username).trim())
    .map(r=>String(r.subject).trim());
}

// Throws if the logged-in user isn't allotted to this subject (admins always pass).
function _assertSubjectAllowed(token, subject){
  const allowed=_mySubjects(token);
  if(allowed.indexOf(subject)<0) throw new Error('You are not allotted to this subject.');
}
function _s(v){ return (v===null||v===undefined)?'':String(v); }
function _t(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'HH:mm'); } return _s(v); }
function _d(v){ if(v instanceof Date){ return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd'); } return String(v); }


function login(username, password){
  const users = readTable(SH_USERS);
  const u = users.find(x => String(x.username).trim() === String(username).trim()
                        && String(x.password) === String(password));
  if (!u) return { ok:false, msg:'Invalid username or password' };
  return { ok:true, role:u.role, name:u.name, username:u.username,
           token:_makeToken(u.username, u.role) };
}

// ====== OTP LOGIN (Mobile number) ======
// Sends a one-time code to the student's registered mobile via Fast2SMS.
// Requires a Script Property named FAST2SMS_API_KEY (Project Settings ->
// Script Properties). Get the key from your Fast2SMS dashboard. Never hardcode
// it in this file.

const OTP_TTL_SECONDS = 300;             // OTP valid for 5 minutes
const OTP_RESEND_COOLDOWN_SECONDS = 45;  // min gap between two requests for the same number
const OTP_MAX_ATTEMPTS = 5;              // wrong tries allowed before the OTP is killed

function _otpCache(){ return CacheService.getScriptCache(); }

function _findStudentByMobile(mobile){
  const clean = String(mobile).trim();
  return readTable(SH_STUDENTS).find(x => String(x.mobile).trim() === clean);
}

// Step 1: student submits their mobile number.
function requestOtp(mobile){
  const clean = String(mobile).trim();
  if (!/^\d{10}$/.test(clean)) return { ok:false, msg:'Enter a valid 10-digit mobile number' };

  const cache = _otpCache();
  const cooldownKey = 'otp_cd_' + clean;
  if (cache.get(cooldownKey)) return { ok:false, msg:'Please wait a moment before requesting another OTP' };

  const stu = _findStudentByMobile(clean);
  // Same message whether or not the number is registered, so this endpoint
  // can't be used to check which numbers belong to enrolled students.
  if (!stu) return { ok:true, msg:'If this number is registered, an OTP has been sent' };

  const otp = String(Math.floor(100000 + Math.random() * 900000)); // 6-digit code
  cache.put('otp_' + clean, otp, OTP_TTL_SECONDS);
  cache.put(cooldownKey, '1', OTP_RESEND_COOLDOWN_SECONDS);
  cache.remove('otp_attempts_' + clean);

  _sendOtpSms(clean, otp);
  return { ok:true, msg:'If this number is registered, an OTP has been sent' };
}

// Step 2: student submits the code they received.
function verifyOtp(mobile, otp){
  const clean = String(mobile).trim();
  const cache = _otpCache();
  const key = 'otp_' + clean;
  const attemptsKey = 'otp_attempts_' + clean;

  const stored = cache.get(key);
  if (!stored) return { ok:false, msg:'OTP expired or not requested. Please request a new one.' };

  const attempts = Number(cache.get(attemptsKey) || 0);
  if (attempts >= OTP_MAX_ATTEMPTS){
    cache.remove(key);
    return { ok:false, msg:'Too many incorrect attempts. Please request a new OTP.' };
  }

  if (String(otp).trim() !== stored){
    cache.put(attemptsKey, String(attempts + 1), OTP_TTL_SECONDS);
    return { ok:false, msg:'Incorrect OTP' };
  }

  cache.remove(key);
  cache.remove(attemptsKey);

  const stu = _findStudentByMobile(clean);
  if (!stu) return { ok:false, msg:'Student not found' };

  // Student accounts are keyed by enrollment number in the Users sheet, so we
  // reuse that identity - the rest of the app (dashboards, tokens) needs no changes.
  const users = readTable(SH_USERS);
  const u = users.find(x => String(x.username).trim() === String(stu.enrollment).trim());
  if (!u) return { ok:false, msg:'No login account found for this student. Contact admin.' };

  return { ok:true, role:u.role, name:u.name, username:u.username,
           token:_makeToken(u.username, u.role) };
}

function _sendOtpSms(mobile, otp){
  const apiKey = PropertiesService.getScriptProperties().getProperty('FAST2SMS_API_KEY');
  if (!apiKey) throw new Error('FAST2SMS_API_KEY is not set. Add it under Project Settings -> Script Properties.');

  const url = 'https://www.fast2sms.com/dev/bulkV2'
    + '?authorization=' + encodeURIComponent(apiKey)
    + '&route=otp'
    + '&variables_values=' + otp
    + '&numbers=' + clean_(mobile);

  const res = UrlFetchApp.fetch(url, { method: 'get', muteHttpExceptions: true });
  const body = JSON.parse(res.getContentText());
  if (!body.return) throw new Error('SMS send failed: ' + res.getContentText());
}
function clean_(v){ return String(v).trim(); }

// ====== STUDENT (read) ======
// Token-based: a student can ONLY see their own data. The enrollment is taken
// from the verified token, never from a client-supplied value. Admins may pass
// an explicit enrollment to view any student.
function getStudentDashboard(token, requestedEnrollment){
  const s = _verifyToken(token);
  if (!s) return { ok:false, msg:'Session expired. Please log in again.' };
  const enrollment = (s.role === 'admin' && requestedEnrollment)
                     ? requestedEnrollment : s.username;
  const stu = readTable(SH_STUDENTS).find(x => String(x.enrollment).trim() === String(enrollment).trim());
  if (!stu) return { ok:false, msg:'Student not found' };
  const results  = readTable(SH_RESULTS).filter(r => String(r.enrollment).trim() === String(enrollment).trim());
  const subjects = readTable(SH_SUBJECTS).filter(sub => sub.className === stu.className);
  const rows = results.map(r => {
    const sub = subjects.find(x => x.subjectCode === r.subjectCode) || {};
    return {
      subjectCode: r.subjectCode,
      subjectName: sub.subjectName || r.subjectCode,
      teacher: sub.teacher || '',
      marks: r.marks,
      maxMarks: sub.maxMarks || 100,
      examName: r.examName
    };
  });
  return { ok:true, student:stu, results:rows };
}

// ====== ADMIN: Students ======
function listStudents(token){ _requireAdmin(token); return readTable(SH_STUDENTS); }


// ---- Flexible batch cutoffs (admin-editable, stored in Script Properties) ----
// Two cutoffs split the roll range into A / B / C:
//   A = rolls 1..cutAB , B = (cutAB+1)..cutBC , C = (cutBC+1)..
// Defaults keep the original 22/22/22 split.
function _getCutoffs(){
  const p=PropertiesService.getScriptProperties();
  const ab=parseInt(p.getProperty('BATCH_CUT_AB')||'22',10);
  const bc=parseInt(p.getProperty('BATCH_CUT_BC')||'44',10);
  return { ab:ab, bc:bc };
}
function getBatchCutoffs(token){ _requireStaff(token); return _getCutoffs(); }
function setBatchCutoffs(token, cutAB, cutBC){
  _requireAdmin(token);
  cutAB=parseInt(cutAB,10); cutBC=parseInt(cutBC,10);
  if(!cutAB || !cutBC || cutBC<=cutAB) return {ok:false,msg:'Cutoffs must be numbers with B/C cutoff greater than A/B cutoff.'};
  const p=PropertiesService.getScriptProperties();
  p.setProperty('BATCH_CUT_AB',String(cutAB));
  p.setProperty('BATCH_CUT_BC',String(cutBC));
  return {ok:true, ab:cutAB, bc:cutBC};
}

// Derive A/B/C batch from a roll number using the current cutoffs.
function _batchOf(roll){
  const n=parseInt(String(roll).replace(/\D/g,''),10);
  if(!n) return '';
  const c=_getCutoffs();
  if(n<=c.ab) return 'A';
  if(n<=c.bc) return 'B';
  return 'C';
}

// Re-apply the current cutoffs to every student (auto-assign A/B/C by roll).
// Does NOT touch students whose batch was manually set if keepManual is true
// (manual overrides are marked by a trailing '*' — see setStudentBatch).
function autoAssignBatches(token, keepManual){
  _requireAdmin(token);
  const sh=getSheet(SH_STUDENTS);
  const data=sh.getDataRange().getValues();
  const head=data[0].map(x=>String(x).trim());
  const rollCol=head.indexOf('rollNo'), batchCol=head.indexOf('batch');
  if(rollCol<0||batchCol<0) return {ok:false,msg:'Students tab missing rollNo/batch columns'};
  let changed=0;
  for(let i=1;i<data.length;i++){
    const cur=String(data[i][batchCol]||'');
    if(keepManual && cur.indexOf('*')>=0) continue; // leave manual overrides alone
    const b=_batchOf(data[i][rollCol]);
    if(b && b!==cur.replace('*','')){ sh.getRange(i+1,batchCol+1).setValue(b); changed++; }
    else if(cur.indexOf('*')>=0){ sh.getRange(i+1,batchCol+1).setValue(cur.replace('*','')); }
  }
  return {ok:true, changed:changed};
}

// Move one student to a specific batch (manual override, marked so auto-assign can skip it).
// Batch names are free-form — not limited to A/B/C — so admins can use whatever grouping suits
// a given subject/class (e.g. "A", "Group 1", "Morning", "Batch-X").
function setStudentBatch(token, enrollment, batch){
  _requireAdmin(token);
  batch=String(batch||'').trim();
  if(!batch) return {ok:false,msg:'Enter a batch name'};
  if(batch.length>40) return {ok:false,msg:'Batch name is too long'};
  const sh=getSheet(SH_STUDENTS);
  const data=sh.getDataRange().getValues();
  const head=data[0].map(x=>String(x).trim());
  const enrCol=head.indexOf('enrollment'), batchCol=head.indexOf('batch');
  for(let i=1;i<data.length;i++){
    if(String(data[i][enrCol]).trim()===String(enrollment).trim()){
      sh.getRange(i+1,batchCol+1).setValue(batch+'*'); // '*' marks a manual override
      return {ok:true};
    }
  }
  return {ok:false,msg:'Student not found'};
}

// Existing sheets created before the SPI feature won't have a 'spi' column yet — add it once, in place.
function _ensureSpiColumn(){
  const sh=getSheet(SH_STUDENTS);
  const lastCol=Math.max(sh.getLastColumn(),1);
  const head=sh.getRange(1,1,1,lastCol).getValues()[0].map(x=>String(x).trim());
  if(head.indexOf('spi')<0){ sh.getRange(1,lastCol+1).setValue('spi'); }
}

// List CO-5-K students with their current SPI-eligibility flag (for the admin selection screen).
function listSpiEligibility(token){
  _requireAdmin(token);
  _ensureSpiColumn();
  const norm=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
  const TARGET=norm(CLASS_NAME_SHORT);
  let list=readTable(SH_STUDENTS)
    .filter(s=> norm(s.className)===TARGET || String(s.className||'').trim()==='')
    .map(s=>({roll:String(s.rollNo), enrollment:String(s.enrollment), name:_s(s.name),
               spi: String(s.spi||'').trim().toUpperCase()==='Y'}));
  list.sort((a,b)=>Number(a.roll)-Number(b.roll));
  return list;
}

// Save the fixed set of students who take SPI (checked = Y, unchecked = blank). One sheet pass.
function saveSpiSelection(token, enrollments){
  _requireAdmin(token);
  _ensureSpiColumn();
  const want=new Set((enrollments||[]).map(e=>String(e).trim()));
  const sh=getSheet(SH_STUDENTS);
  const data=sh.getDataRange().getValues();
  const head=data[0].map(x=>String(x).trim());
  const enrCol=head.indexOf('enrollment'), spiCol=head.indexOf('spi');
  if(enrCol<0||spiCol<0) return {ok:false,msg:'Students tab missing enrollment/spi columns'};
  let changed=0;
  for(let i=1;i<data.length;i++){
    const enr=String(data[i][enrCol]).trim();
    const cur=String(data[i][spiCol]||'').trim().toUpperCase();
    const wantVal = want.has(enr) ? 'Y' : '';
    if(cur!==wantVal){ sh.getRange(i+1,spiCol+1).setValue(wantVal); changed++; }
  }
  return {ok:true, changed:changed, count:want.size};
}


function listStudentsWithBatch(token, classFilter){
  _requireStaff(token);
  const norm=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
  const want = classFilter ? norm(classFilter) : norm(CLASS_NAME_SHORT);
  const isCO5K = want===norm(CLASS_NAME_SHORT);
  const list=readTable(SH_STUDENTS)
    .filter(s=>{
      const c=norm(s.className);
      // For CO-5-K also include blank-class rows (older data uploaded before classes were tagged).
      return c===want || (isCO5K && String(s.className||'').trim()==='');
    })
    .map(s=>({
      roll:String(s.rollNo), enrollment:String(s.enrollment), name:_s(s.name),
      className:String(s.className||''),
      batch:String(s.batch||'').replace('*',''), manual:String(s.batch||'').indexOf('*')>=0
    }));
  list.sort((a,b)=>Number(a.roll)-Number(b.roll));
  return list;
}

// Distinct batch names actually in use (optionally scoped to one class) — batches are free-form,
// not limited to A/B/C, so dropdowns elsewhere populate from whatever names admins have actually used.
function listBatchesInUse(token, classFilter){
  _requireStaff(token);
  const norm=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
  const want = classFilter ? norm(classFilter) : '';
  const set=new Set();
  readTable(SH_STUDENTS).forEach(s=>{
    if(want && norm(s.className)!==want) return;
    const b=String(s.batch||'').replace('*','').trim();
    if(b) set.add(b);
  });
  return Array.from(set).sort();
}

// Add ONE student manually (schema-safe; writes by column name; creates a login).
// s = {roll, enrollment, name, batch?}. If batch omitted/invalid, it is auto-assigned by cutoffs.
function attAddStudent(token, s){
  _requireAdmin(token);
  const roll=String(s.roll||'').trim();
  const enr=String(s.enrollment||'').trim();
  const nm=String(s.name||'').trim();
  if(!enr||!nm||!roll) return {ok:false,msg:'Roll No, Enrollment and Name are all required.'};
  if(readTable(SH_STUDENTS).some(x=>String(x.enrollment).trim()===enr))
    return {ok:false,msg:'A student with this enrollment already exists.'};

  // Class: must be one of the classes the app manages. Defaults to CO-5-K.
  let cls=String(s.className||'').trim().toUpperCase();
  const match=ALL_CLASSES.filter(c=>_normCls(c)===_normCls(cls))[0];
  cls = match || CLASS_NAME_SHORT;

  // Batch is free-form (not limited to A/B/C) and can apply to any class with batch-wise practicals.
  // If none was given, fall back to the roll-number auto-assign convenience (A/B/C by cutoff).
  let batch=String(s.batch||'').trim();
  if(!batch) batch=_batchOf(roll);

  const sh=getSheet(SH_STUDENTS);
  const head=sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const col=n=>head.indexOf(n);
  const row=new Array(head.length).fill('');
  if(col('enrollment')>=0) row[col('enrollment')]=enr;
  if(col('name')>=0) row[col('name')]=nm;
  if(col('className')>=0) row[col('className')]=cls;
  if(col('rollNo')>=0) row[col('rollNo')]=roll;
  if(col('batch')>=0) row[col('batch')]=batch;
  sh.getRange(sh.getLastRow()+1,1,1,head.length).setValues([row]);

  // login (username = password = enrollment)
  const usr=getSheet(SH_USERS);
  if(!readTable(SH_USERS).some(u=>String(u.username).trim()===enr))
    usr.appendRow([enr, enr, 'student', nm]);
  return {ok:true, className:cls};
}

// Edit an existing student's roll / name / batch (found by original enrollment).
// Enrollment itself is the key and is not changed here.
function attUpdateStudent(token, enrollment, fields){
  _requireAdmin(token);
  const sh=getSheet(SH_STUDENTS);
  const data=sh.getDataRange().getValues();
  const head=data[0].map(x=>String(x).trim());
  const col=n=>head.indexOf(n);
  const enrCol=col('enrollment');
  for(let i=1;i<data.length;i++){
    if(String(data[i][enrCol]).trim()===String(enrollment).trim()){
      if(fields.roll!==undefined && col('rollNo')>=0) sh.getRange(i+1,col('rollNo')+1).setValue(String(fields.roll).trim());
      if(fields.name!==undefined && col('name')>=0){
        const nm=String(fields.name).trim();
        sh.getRange(i+1,col('name')+1).setValue(nm);
        // keep the login display name in sync
        const usr=getSheet(SH_USERS); const ud=usr.getDataRange().getValues();
        for(let k=1;k<ud.length;k++){ if(String(ud[k][0]).trim()===String(enrollment).trim()){ usr.getRange(k+1,4).setValue(nm); break; } }
      }
      if(fields.batch!==undefined && col('batch')>=0){
        let b=String(fields.batch).trim();
        if(b) sh.getRange(i+1,col('batch')+1).setValue(b+'*'); // '*' marks a manual override
      }
      return {ok:true};
    }
  }
  return {ok:false,msg:'Student not found.'};
}

// Upload a student list (Excel/CSV) for the attendance/shared roster.
// Header must contain Roll No, Enrollment, Name (+ optional Batch).
// If Batch column present, it is used as-is; otherwise batches are auto-assigned by cutoffs.
function uploadStudentList(token, payload){
  _requireAdmin(token);
  if(!payload || !payload.base64) return {ok:false,msg:'No file received'};
  let tempId=null;
  try{
    const bytes=Utilities.base64Decode(payload.base64);
    const mime=payload.mimeType||MimeType.MICROSOFT_EXCEL;
    const blob=Utilities.newBlob(bytes,mime,payload.fileName||'upload');
    const name=String(payload.fileName||'').toLowerCase();
    let rows;
    if(name.endsWith('.csv')||mime==='text/csv'){
      rows=Utilities.parseCsv(blob.getDataAsString());
    } else {
      if(typeof Drive==='undefined'||!Drive.Files) return {ok:false,msg:'Enable the Drive service (Services → Drive API), then retry.'};
      let created;
      if(typeof Drive.Files.create==='function') created=Drive.Files.create({name:'TMP_'+Date.now(),mimeType:MimeType.GOOGLE_SHEETS},blob);
      else created=Drive.Files.insert({title:'TMP_'+Date.now(),mimeType:MimeType.GOOGLE_SHEETS},blob,{convert:true});
      tempId=created.id;
      rows=SpreadsheetApp.openById(tempId).getSheets()[0].getDataRange().getDisplayValues();
    }
    if(!rows||rows.length<2) return {ok:false,msg:'File has no data rows'};
    const header=rows[0].map(h=>String(h).trim().toLowerCase());
    const find=cands=>{ for(let i=0;i<header.length;i++) for(const c of cands) if(header[i].indexOf(c)>=0) return i; return -1; };
    const cRoll=find(['roll']), cEnr=find(['enroll','enrol']), cName=find(['name']), cBatch=find(['batch']);
    if(cRoll<0||cEnr<0||cName<0) return {ok:false,msg:'Header must contain Roll No, Enrollment and Name columns'};

    function normId(v){ let s=String(v==null?'':v).trim(); if(/[eE]\+?\d|\.\d+/.test(s)&&!isNaN(Number(s))){const n=Number(s); if(isFinite(n)) s=n.toLocaleString('en-US',{useGrouping:false,maximumFractionDigits:0});} return s; }

    const stu=getSheet(SH_STUDENTS);
    const head=stu.getRange(1,1,1,stu.getLastColumn()).getValues()[0].map(x=>String(x).trim());
    const col=n=>head.indexOf(n);
    const clsCol=col('className'), enrCol=col('enrollment');

    // The class these uploaded students belong to (this upload replaces ONLY this class).
    // The class these uploaded students belong to (this upload replaces ONLY this class).
    // Chosen by the admin on the upload screen; defaults to CO-5-K.
    let uploadClass = CLASS_NAME_SHORT;
    if(payload.className){
      const want=_normCls(payload.className);
      const m=ALL_CLASSES.filter(c=>_normCls(c)===want)[0];
      if(m) uploadClass=m;
    }
    const _norm=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
    const uploadClassNorm=_norm(uploadClass);

    // Preserve all EXISTING students that are NOT in the uploaded class, and remember
    // which of their enrollments already have logins (so we don't disturb them).
    const allData = stu.getLastRow()>1 ? stu.getRange(2,1,stu.getLastRow()-1,head.length).getValues() : [];
    const keptRows=[];            // rows of other classes, kept as-is
    const removedEnr={};          // enrollments of the old uploaded-class rows being replaced
    allData.forEach(r=>{
      const cls = clsCol>=0 ? _norm(r[clsCol]) : '';
      if(cls===uploadClassNorm){ removedEnr[String(r[enrCol]).trim()]=true; } // old CO-5-K -> drop
      else if(String(r[enrCol]).trim()!==''){ keptRows.push(r); }             // other class -> keep
    });

    const usr=getSheet(SH_USERS);
    const existingUser={}; readTable(SH_USERS).forEach(u=>existingUser[String(u.username).trim()]=true);
    const newLogins=[];
    const out=[];
    for(let r=1;r<rows.length;r++){
      const enr=normId(rows[r][cEnr]); const nm=String(rows[r][cName]||'').trim();
      if(!enr||!nm) continue;
      const roll=normId(rows[r][cRoll]).replace(/\.0$/,'');
      // Batch is free-form, not limited to A/B/C, and can apply to any class. If none was given,
      // fall back to the roll-number auto-assign convenience (A/B/C by cutoff) as a default.
      let batch = cBatch>=0 ? String(rows[r][cBatch]||'').trim() : '';
      if(!batch) batch=_batchOf(roll);
      const row=new Array(head.length).fill('');
      if(col('enrollment')>=0) row[col('enrollment')]=enr;
      if(col('name')>=0) row[col('name')]=nm;
      if(col('className')>=0) row[col('className')]=uploadClass;
      if(col('rollNo')>=0) row[col('rollNo')]=roll;
      if(col('batch')>=0) row[col('batch')]=batch;
      out.push(row);
      if(!existingUser[enr]){ newLogins.push([enr,enr,'student',nm]); existingUser[enr]=true; }
    }
    if(!out.length) return {ok:false,msg:'No valid rows found'};

    // Rewrite the Students sheet = kept (other classes) + newly uploaded class.
    if(stu.getLastRow()>1) stu.getRange(2,1,stu.getLastRow()-1,head.length).clearContent();
    const finalRows=keptRows.concat(out);
    stu.getRange(2,1,finalRows.length,head.length).setValues(finalRows);
    if(newLogins.length) usr.getRange(usr.getLastRow()+1,1,newLogins.length,4).setValues(newLogins);
    return {ok:true, count:out.length, logins:newLogins.length, kept:keptRows.length};
  }catch(err){
    return {ok:false,msg:'Could not read file: '+err.message};
  }finally{
    if(tempId){ try{ DriveApp.getFileById(tempId).setTrashed(true);}catch(e){} }
  }
}

function addStudent(token, s){
  _requireAdmin(token);
  const sh = getSheet(SH_STUDENTS);
  const exists = readTable(SH_STUDENTS).some(x => String(x.enrollment).trim() === String(s.enrollment).trim());
  if (exists) return { ok:false, msg:'Enrollment already exists' };
  // Batches (A/B/C) apply only to CO-5-K, whose practicals are batch-wise.
  const batch = _normCls(s.className)===_normCls(CLASS_NAME_SHORT) ? _batchOf(s.rollNo) : '';
  sh.appendRow([s.enrollment, s.name, s.className, s.rollNo, batch, s.email || '', s.mobile || '']);
  // student login (username = password = enrollment, consistent across the app)
  const u = getSheet(SH_USERS);
  const hasUser = readTable(SH_USERS).some(x => String(x.username).trim() === String(s.enrollment).trim());
  if (!hasUser) u.appendRow([s.enrollment, s.enrollment, 'student', s.name]);
  return { ok:true };
}

function deleteStudent(token, enrollment){
  _requireAdmin(token);
  _deleteRowByKey(SH_STUDENTS, 'enrollment', enrollment);
  _deleteRowByKey(SH_USERS, 'username', enrollment);
  return { ok:true };
}

// Delete several students at once (by array of enrollment numbers).
// Also removes their logins and their result rows. One pass per sheet = fast.
function deleteStudentsBulk(token, enrollments){
  _requireAdmin(token);
  if (!enrollments || !enrollments.length) return { ok:false, msg:'Nothing selected' };
  const set = {};
  enrollments.forEach(e => set[String(e).trim()] = true);

  const removed = _deleteRowsWhere(SH_STUDENTS, 'enrollment', en => set[en]);
  _deleteRowsWhere(SH_USERS, 'username', un => set[un]);
  _deleteRowsWhere(SH_RESULTS, 'enrollment', en => set[en]);
  return { ok:true, removed:removed };
}

// Delete every student in a class (plus their logins and results).
function deleteClassStudents(token, className){
  _requireAdmin(token);
  if (!className) return { ok:false, msg:'No class given' };
  const enrolls = readTable(SH_STUDENTS)
    .filter(s => String(s.className).trim() === String(className).trim())
    .map(s => String(s.enrollment).trim());
  if (!enrolls.length) return { ok:false, msg:'That class has no students.' };
  return deleteStudentsBulk(token, enrolls);
}

// Generic: delete all rows where the value in `keyCol` passes `predicate`.
// Returns the number of rows removed. Single rewrite of the sheet = efficient.
function _deleteRowsWhere(sheetName, keyCol, predicate){
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  if (data.length < 2) return 0;
  const idx = data[0].indexOf(keyCol);
  if (idx < 0) return 0;
  const keep = [data[0]];
  let removed = 0;
  for (let i = 1; i < data.length; i++){
    if (predicate(String(data[i][idx]).trim())) { removed++; }
    else { keep.push(data[i]); }
  }
  if (removed > 0){
    sh.clearContents();
    sh.getRange(1, 1, keep.length, data[0].length).setValues(keep);
  }
  return removed;
}

// ====== ADMIN: Classes ======
function listClasses(token){ _requireAdmin(token); return readTable(SH_CLASSES); }
function addClass(token, c){
  _requireAdmin(token);
  const exists = readTable(SH_CLASSES).some(x => x.className === c.className);
  if (exists) return { ok:false, msg:'Class already exists' };
  getSheet(SH_CLASSES).appendRow([c.className, c.description || '']);
  return { ok:true };
}
function deleteClass(token, className){ _requireAdmin(token); _deleteRowByKey(SH_CLASSES,'className',className); return {ok:true}; }

// ====== ADMIN: Subjects + Teacher ======
function listSubjects(token){ _requireAdmin(token); return readTable(SH_SUBJECTS); }
function addSubject(token, s){
  _requireAdmin(token);
  getSheet(SH_SUBJECTS).appendRow([s.className, s.subjectCode, s.subjectName, s.teacher, s.maxMarks || 100]);
  return { ok:true };
}
function deleteSubject(token, className, subjectCode){
  _requireAdmin(token);
  const sh = getSheet(SH_SUBJECTS);
  const data = sh.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--){
    if (data[i][0] === className && data[i][1] === subjectCode){ sh.deleteRow(i+1); }
  }
  return { ok:true };
}

// ====== ADMIN: Marks entry ======
function saveMark(token, m){
  _requireAdmin(token);
  const sh = getSheet(SH_RESULTS);
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++){
    if (String(data[i][0])===String(m.enrollment) && data[i][2]===m.subjectCode && data[i][4]===m.examName){
      sh.getRange(i+1, 4).setValue(m.marks); // update marks
      return { ok:true, updated:true };
    }
  }
  sh.appendRow([m.enrollment, m.className, m.subjectCode, m.marks, m.examName]);
  return { ok:true, updated:false };
}

// ====== ADMIN: Import directly from an uploaded file (server-side parse) ======
// The browser sends the raw file as base64 — no external JS library needed.
// Apps Script converts xlsx/csv to a temporary Google Sheet, reads the rows,
// then hands them to bulkImport(). Works regardless of CDN / CSP / admin policy.
function bulkImportFile(token, payload){
  _requireAdmin(token);
  if (!payload || !payload.base64) return { ok:false, msg:'No file received' };
  if (!payload.className) return { ok:false, msg:'Select a class first' };

  let tempId = null;
  try {
    const bytes = Utilities.base64Decode(payload.base64);
    const mime  = payload.mimeType || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    const blob  = Utilities.newBlob(bytes, mime, payload.fileName || 'upload');

    let rows;
    const name = String(payload.fileName || '').toLowerCase();
    if (name.endsWith('.csv') || mime === 'text/csv'){
      // CSV: parse text directly
      const text = blob.getDataAsString();
      rows = Utilities.parseCsv(text);
    } else {
      // xlsx/xls: convert to a temporary Google Sheet via the Drive service,
      // then read it. Works with whichever Drive API version is enabled:
      //   v3 -> Drive.Files.create(resource, blob)  (resource uses `name`)
      //   v2 -> Drive.Files.insert(resource, blob, {convert:true})  (uses `title`)
      let createdId;
      if (typeof Drive === 'undefined' || !Drive.Files){
        return { ok:false, msg:'The Drive service is not enabled. In the Apps Script '
               + 'editor, click + next to "Services", add "Drive API", save, then redeploy.' };
      }
      if (typeof Drive.Files.create === 'function'){
        const resource = { name: 'TMP_IMPORT_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
        const file = Drive.Files.create(resource, blob);
        createdId = file.id;
      } else if (typeof Drive.Files.insert === 'function'){
        const resource = { title: 'TMP_IMPORT_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
        const file = Drive.Files.insert(resource, blob, { convert: true });
        createdId = file.id;
      } else {
        return { ok:false, msg:'Drive service is present but has no create/insert method. '
               + 'Remove and re-add the Drive API service in the editor, then redeploy.' };
      }
      tempId = createdId;
      const ss = SpreadsheetApp.openById(tempId);
      const sh = ss.getSheets()[0];
      // read as display strings so big enrollment numbers don't become 2.5E10
      const rng = sh.getDataRange();
      rows = rng.getDisplayValues();
    }

    if (!rows || rows.length < 2) return { ok:false, msg:'File has no data rows' };

    const result = bulkImport(token, { rows: rows, className: payload.className,
                                       examName: payload.examName || 'PAT-1' });
    return result;
  } catch (err) {
    return { ok:false, msg:'Could not read the file on the server: ' + err.message
                          + '. If this mentions "Drive", enable the Drive advanced service (see setup guide).' };
  } finally {
    if (tempId){ try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) {} }
  }
}

// ====== ADMIN: Bulk import of a class marks sheet ======
// Expects: rows = array of arrays (first row = header), where the header is
//   Roll No | Enrollment no | Name | <SUBJECT_CODE> | <SUBJECT_CODE> ...
// className, examName supplied by admin. Creates students + logins, auto-creates
// any missing subjects for the class, and stores any marks present in the cells.
function bulkImport(token, payload){
  _requireAdmin(token);
  const rows = payload.rows || [];
  const className = payload.className;
  const examName = payload.examName || 'PAT-1';
  if (!className) return { ok:false, msg:'Select a class first' };
  if (rows.length < 2) return { ok:false, msg:'No data rows found' };

  // --- locate header columns ---
  const header = rows[0].map(h => String(h).trim());
  const lower = header.map(h => h.toLowerCase());
  const findCol = function(cands){
    for (let i = 0; i < lower.length; i++){
      for (let c = 0; c < cands.length; c++){ if (lower[i].indexOf(cands[c]) >= 0) return i; }
    }
    return -1;
  };
  const cRoll = findCol(['roll']);
  const cEnr  = findCol(['enroll','enrol']);
  const cName = findCol(['name']);
  if (cEnr < 0 || cName < 0)
    return { ok:false, msg:'Header must contain "Enrollment" and "Name" columns' };

  // subject columns = every column that is not roll/enroll/name and has a header
  const subjectCols = [];
  header.forEach((h, i) => {
    if (i !== cRoll && i !== cEnr && i !== cName && h) subjectCols.push({ idx:i, code:h });
  });

  // --- ensure subjects exist for this class (auto-create missing) ---
  const subSheet = getSheet(SH_SUBJECTS);
  const existingSubs = readTable(SH_SUBJECTS).filter(s => s.className === className);
  subjectCols.forEach(sc => {
    const found = existingSubs.some(s => String(s.subjectCode).trim().toLowerCase() === sc.code.toLowerCase());
    if (!found){
      // subjectName/teacher left for admin to complete; default max 30 (PAT max)
      subSheet.appendRow([className, sc.code, sc.code, '', 30]);
    }
  });

  // --- preload existing students / users / results for speed ---
  const stuSheet = getSheet(SH_STUDENTS);
  const usrSheet = getSheet(SH_USERS);
  const resSheet = getSheet(SH_RESULTS);
  const existingStu = readTable(SH_STUDENTS);
  const existingUsr = readTable(SH_USERS);
  const resData = resSheet.getDataRange().getValues(); // for update-in-place

  let added = 0, updated = 0, marksSaved = 0;

  // Convert values like "2.5410360117E10" or 25410360117 to a plain integer string.
  function normId(v){
    let s = String(v == null ? '' : v).trim();
    if (s === '') return '';
    if (/[eE]\+?\d|\.\d+/.test(s) && !isNaN(Number(s))){
      const n = Number(s);
      if (isFinite(n)) s = n.toLocaleString('en-US', { useGrouping:false, maximumFractionDigits:0 });
    }
    return s;
  }

  for (let r = 1; r < rows.length; r++){
    const row = rows[r];
    const enr = normId(row[cEnr]);
    const nm  = String(row[cName] == null ? '' : row[cName]).trim();
    if (!enr || !nm) continue; // skip blank lines
    const roll = cRoll >= 0 ? normId(row[cRoll]).replace(/\.0$/, '') : '';

    // student
    const stuExists = existingStu.some(s => String(s.enrollment).trim() === enr);
    if (stuExists){
      // update name/class/roll
      _updateStudentRow(stuSheet, enr, nm, className, roll);
      updated++;
    } else {
      stuSheet.appendRow([enr, nm, className, roll, _batchOf(roll), '', '']);
      existingStu.push({ enrollment:enr });
      added++;
    }
    // login
    const usrExists = existingUsr.some(u => String(u.username).trim() === enr);
    if (!usrExists){ usrSheet.appendRow([enr, enr, 'student', nm]); existingUsr.push({ username:enr }); }

    // marks
    subjectCols.forEach(sc => {
      const cell = row[sc.idx];
      if (cell === '' || cell === null || cell === undefined) return;
      const mark = String(cell).trim();
      if (mark === '') return;
      // update if same enrollment+subject+exam exists, else append
      let done = false;
      for (let i = 1; i < resData.length; i++){
        if (String(resData[i][0]).trim() === enr && String(resData[i][2]).trim() === sc.code
            && String(resData[i][4]).trim() === examName){
          resSheet.getRange(i+1, 4).setValue(mark); done = true; break;
        }
      }
      if (!done){
        resSheet.appendRow([enr, className, sc.code, mark, examName]);
        resData.push([enr, className, sc.code, mark, examName]);
      }
      marksSaved++;
    });
  }

  return { ok:true, added:added, updated:updated, marksSaved:marksSaved,
           subjects: subjectCols.map(s => s.code) };
}

function _updateStudentRow(sh, enrollment, name, className, roll){
  const data = sh.getDataRange().getValues();
  for (let i = 1; i < data.length; i++){
    if (String(data[i][0]).trim() === String(enrollment).trim()){
      sh.getRange(i+1, 2).setValue(name);
      sh.getRange(i+1, 3).setValue(className);
      if (roll){ sh.getRange(i+1, 4).setValue(roll); sh.getRange(i+1, 5).setValue(_batchOf(roll)); }
      return;
    }
  }
}

// ====== Report Card rendering (shared by admin + student) ======
// Maps an exam name to the printed title.
//   PAT-1 / PAT 1 / PAT1 -> "Progressive Theory Test -1"
//   anything else        -> the exam name as-is (e.g. "Unit Test 1")
function _testTitleFromExam(examName){
  const m = String(examName || '').match(/^\s*PAT[\s\-]?(\d+)\s*$/i);
  if (m) return 'Progressive Theory Test -' + m[1];
  return String(examName || '');
}

// Normalises a season label. Accepts an explicit value ("Winter"/"Summer"),
// or derives it from an academic-year string that starts with W/S (e.g. "W-2026").
function _normSeason(season, academicYear){
  const s = String(season || '').trim().toLowerCase();
  if (s === 'winter' || s === 'w') return 'Winter';
  if (s === 'summer' || s === 's') return 'Summer';
  const ay = String(academicYear || '').trim().toUpperCase();
  if (/^W[\s\-]/.test(ay) || ay === 'W') return 'Winter';
  if (/^S[\s\-]/.test(ay) || ay === 'S') return 'Summer';
  return '';
}

// Returns the HTML for ONE student's report card.
// Always shows both Progressive Assessment Tests side by side, per subject.
function _renderCardHtml(stu, subjects, results, semesterLabel, academicYear, withBreak, season){
  const brk = withBreak ? 'page-break-after:always;' : '';
  let html = '<div style="padding:26px 36px;' + brk + '">';

  html += '<table style="width:100%;border-collapse:collapse;"><tr>'
        + '<td style="text-align:center;border:none;">'
        + '<img src="' + LOGO_DATA_URI + '" style="height:62px;"><br>'
        + '<span style="font-size:21px;font-weight:bold;">Dr.Panjabrao Deshmukh Polytechnic, Amravati</span><br>'
        + '<span style="font-size:15px;font-weight:bold;">Computer Engg. dept.</span>'
        + '</td></tr></table>';

  html += '<div style="font-size:10px;margin-top:6px;">'
        + '<span style="color:#b00;font-weight:bold;">Vision -</span> To educate students, to be computer '
        + 'technology expert &amp; technocrat, with brainstorming and learning hardware, software, co-curricular '
        + 'and extracurricular concepts.<br>'
        + '<span style="color:#b00;font-weight:bold;">Mission -</span> To enhance and pacify knowledge thirst '
        + 'in order to strengthen IT and Computer skills in building blocks of Nation.</div>';

  // Report card now always covers both Progressive Assessment Tests together, not one at a time.
  html += '<div style="text-align:center;font-size:15px;margin:14px 0 8px;">'
        + 'Progressive Assessment Test&nbsp;&nbsp;Report card</div>';
  // Semester + season + academic year, centered together on one line.
  // If the year already carries a W-/S- prefix, drop it so we don't repeat the season.
  const yearClean = String(academicYear || '').replace(/^[WS][\s\-]+/i, '').trim();
  const seasonTxt = season ? season + ' ' : '';
  const ayTxt = yearClean ? '&nbsp;&nbsp;&nbsp;&nbsp;Academic Year: ' + seasonTxt + yearClean : '';
  html += '<div style="text-align:center;font-size:13px;margin:6px 0;">'
        + 'Semester : ' + semesterLabel + ayTxt + '</div>';
  html += '<div style="border-top:3px solid #1a4d8f;margin:8px 0 18px;"></div>';

  const field = function(lbl, val){
    return '<table style="width:100%;border-collapse:collapse;font-size:13px;margin:8px 0;"><tr>'
         + '<td style="border:none;width:175px;">' + lbl + '</td>'
         + '<td style="border:none;font-weight:bold;border-bottom:1px solid #000;">&nbsp;' + val + '</td>'
         + '</tr></table>';
  };
  html += field('Roll Number:', stu.rollNo || '');
  html += field('Enrollment Number:', stu.enrollment);
  html += field('Name of Student:', stu.name);

  html += '<table style="width:100%;border-collapse:collapse;margin-top:16px;font-size:13px;">'
        + '<tr>'
        + '<th style="border:1px solid #000;padding:8px;">Subject</th>'
        + '<th style="border:1px solid #000;padding:8px;width:70px;">Max</th>'
        + '<th style="border:1px solid #000;padding:8px;width:80px;">PAT-1</th>'
        + '<th style="border:1px solid #000;padding:8px;width:80px;">PAT-2</th>'
        + '</tr>';
  subjects.forEach(s => {
    const find = (exam) => results.find(x => String(x.enrollment).trim() === String(stu.enrollment).trim()
                              && String(x.subjectCode).trim() === String(s.subjectCode).trim()
                              && String(x.examName).trim().toUpperCase() === exam);
    const r1 = find('PAT-1');
    const r2 = find('PAT-2');
    const m1 = r1 && r1.marks !== '' ? Number(r1.marks) : null;
    const m2 = r2 && r2.marks !== '' ? Number(r2.marks) : null;
    html += '<tr>'
          + '<td style="border:1px solid #000;padding:8px;text-align:left;">' + (s.subjectName || s.subjectCode) + '</td>'
          + '<td style="border:1px solid #000;padding:8px;text-align:center;">' + (s.maxMarks || 30) + '</td>'
          + '<td style="border:1px solid #000;padding:8px;text-align:center;">' + (m1 === null ? '' : m1) + '</td>'
          + '<td style="border:1px solid #000;padding:8px;text-align:center;">' + (m2 === null ? '' : m2) + '</td>'
          + '</tr>';
  });
  html += '</table>';

  html += '<table style="width:100%;border-collapse:collapse;margin-top:56px;font-size:13px;text-align:center;"><tr>'
        + '<td style="border:none;width:50%;">' + SIG_LEFT_NAME + '<br>' + SIG_LEFT_ROLE + '</td>'
        + '<td style="border:none;width:50%;">' + SIG_RIGHT_NAME + '<br>' + SIG_RIGHT_ROLE + '</td>'
        + '</tr></table>';

  html += '<div style="font-style:italic;font-size:11px;margin-top:28px;">If above marks are not properly '
        + 'displayed, then contact to concern Subject Teacher and make necessary corrections</div>';
  html += '</div>';
  return html;
}

function _cardPageWrap(inner){
  return '<html><head><meta charset="utf-8"></head>'
       + '<body style="font-family:\'Times New Roman\',serif;color:#000;margin:0;">'
       + inner + '</body></html>';
}

// Picks the most recent exam name present for a class (fallback PAT-1).
function _latestExamForClass(className){
  const rows = readTable(SH_RESULTS).filter(r => r.className === className && r.examName);
  if (!rows.length) return 'PAT-1';
  // last appended wins as "latest"
  return String(rows[rows.length - 1].examName);
}

// Store academicYear / season / semester label on the class row (Classes tab).
function _saveClassMeta(className, academicYear, season, semesterLabel){
  const sh = getSheet(SH_CLASSES);
  const data = sh.getDataRange().getValues();
  const head = data[0];
  const cAY  = head.indexOf('academicYear');
  const cSE  = head.indexOf('season');
  const cDesc= head.indexOf('description');
  for (let i = 1; i < data.length; i++){
    if (String(data[i][0]).trim() === String(className).trim()){
      if (cAY  >= 0 && academicYear)   sh.getRange(i+1, cAY+1).setValue(academicYear);
      if (cSE  >= 0 && season)         sh.getRange(i+1, cSE+1).setValue(season);
      if (cDesc>= 0 && semesterLabel)  sh.getRange(i+1, cDesc+1).setValue(semesterLabel);
      return;
    }
  }
}

// Read back the stored meta for a class (used by the student's own card).
function _getClassMeta(className){
  const row = readTable(SH_CLASSES).find(c => String(c.className).trim() === String(className).trim());
  return {
    academicYear:  row && row.academicYear ? String(row.academicYear) : '',
    season:        row && row.season ? String(row.season) : '',
    semesterLabel: row && row.description ? String(row.description) : className
  };
}

// ====== ADMIN: Report Card PDF — whole class, one card per page ======
function generateResultPdf(token, className, semesterLabel, academicYear, season){
  _requireAdmin(token);
  semesterLabel  = semesterLabel || className;
  academicYear   = academicYear || '';
  season         = _normSeason(season, academicYear);

  const students = readTable(SH_STUDENTS).filter(s => s.className === className);
  const subjects = readTable(SH_SUBJECTS).filter(s => s.className === className);
  // Pull BOTH PAT-1 and PAT-2 so the card shows both tests.
  const results  = readTable(SH_RESULTS).filter(r => r.className === className
                     && (String(r.examName).trim().toUpperCase()==='PAT-1' || String(r.examName).trim().toUpperCase()==='PAT-2'));

  if (!students.length) return { ok:false, msg:'No students in this class' };

  // Remember these on the class row so each student's own card shows the same.
  _saveClassMeta(className, academicYear, season, semesterLabel);

  let inner = '';
  students.forEach((stu, idx) => {
    inner += _renderCardHtml(stu, subjects, results, semesterLabel,
                             academicYear, idx < students.length - 1, season);
  });

  const blob = HtmlService.createHtmlOutput(_cardPageWrap(inner)).getBlob()
                 .getAs('application/pdf')
                 .setName('ReportCards_' + className + '.pdf');
  const folder = _getReportFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return { ok:true, url:file.getUrl(), name:file.getName(), count:students.length };
}

// ====== STUDENT: own report card PDF (token identity only) ======
function getMyReportCard(token){
  const s = _verifyToken(token);
  if (!s) return { ok:false, msg:'Session expired. Please log in again.' };

  const enrollment = s.username; // students locked to their own enrollment
  const stu = readTable(SH_STUDENTS).find(x => String(x.enrollment).trim() === String(enrollment).trim());
  if (!stu) return { ok:false, msg:'Your student record was not found. Please contact the office.' };

  const className = stu.className;
  const subjects  = readTable(SH_SUBJECTS).filter(x => x.className === className);
  // Pull BOTH PAT-1 and PAT-2 so the card shows both tests.
  const results   = readTable(SH_RESULTS).filter(r => r.className === className
                      && (String(r.examName).trim().toUpperCase()==='PAT-1' || String(r.examName).trim().toUpperCase()==='PAT-2'));

  // Reuse the academic year / season / semester the admin set for this class.
  const meta = _getClassMeta(className);
  const season = _normSeason(meta.season, meta.academicYear);
  const inner = _renderCardHtml(stu, subjects, results,
                                meta.semesterLabel, meta.academicYear, false, season);

  const blob = HtmlService.createHtmlOutput(_cardPageWrap(inner)).getBlob()
                 .getAs('application/pdf')
                 .setName('ReportCard_' + enrollment + '.pdf');
  const folder = _getReportFolder();
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id = file.getId();
  return {
    ok: true,
    url: 'https://drive.google.com/uc?export=download&id=' + id, // direct download
    previewUrl: 'https://drive.google.com/file/d/' + id + '/preview', // embeddable
    name: file.getName()
  };
}

function _getReportFolder(){
  const name = 'SMS_Reports';
  const it = DriveApp.getFoldersByName(name);
  return it.hasNext() ? it.next() : DriveApp.createFolder(name);
}

function _deleteRowByKey(sheetName, keyCol, keyVal){
  const sh = getSheet(sheetName);
  const data = sh.getDataRange().getValues();
  const idx = data[0].indexOf(keyCol);
  for (let i = data.length - 1; i >= 1; i--){
    if (String(data[i][idx]).trim() === String(keyVal).trim()){ sh.deleteRow(i+1); }
  }
}


/*************************************************************
 * ATTENDANCE MODULE  (shares Users + Students with marks)
 *************************************************************/
// ====== ADMIN: manage lecturer (teacher) accounts and their subject allotment ======
function listTeachers(token){
  _requireAdmin(token);
  _ensureTeacherSheet();
  const users=readTable(SH_USERS).filter(u=>u.role==='teacher');
  const allot=readTable(SH_TEACHER_SUBJ);
  return users.map(u=>({
    username:String(u.username), name:_s(u.name),
    subjects: allot.filter(a=>String(a.username).trim()===String(u.username).trim()).map(a=>String(a.subject).trim())
  }));
}

function createTeacher(token, username, password, name){
  _requireAdmin(token);
  username=String(username||'').trim();
  password=String(password||'');
  name=String(name||'').trim();
  if(!username || !password || !name) return {ok:false,msg:'Username, password and name are required'};
  const users=readTable(SH_USERS);
  if(users.some(u=>String(u.username).trim()===username)) return {ok:false,msg:'That username already exists'};
  getSheet(SH_USERS).appendRow([username, password, 'teacher', name]);
  return {ok:true};
}

function deleteTeacher(token, username){
  _requireAdmin(token);
  _deleteRowByKey(SH_USERS, 'username', username);
  _ensureTeacherSheet();
  _deleteRowsWhere(SH_TEACHER_SUBJ, 'username', v=>v===String(username).trim());
  return {ok:true};
}

// Overwrites the full set of subjects a teacher is allotted to (checkbox-style save: pass every
// subject that should remain checked, in one call, rather than adding/removing one at a time).
function setTeacherSubjects(token, username, subjects){
  _requireAdmin(token);
  const SUBJECTS=_subjects();
  _ensureTeacherSheet();
  _deleteRowsWhere(SH_TEACHER_SUBJ, 'username', v=>v===String(username).trim());
  const sh=getSheet(SH_TEACHER_SUBJ);
  const rows=(subjects||[]).filter(subj=>SUBJECTS[subj]).map(subj=>[String(username).trim(), subj]);
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,2).setValues(rows);
  return {ok:true, count:rows.length};
}

// ====== ADMIN: manage the attendance subject list (semester-wise, no hardcoding) ======
function listAttSubjects(token){
  _requireStaff(token); // admins see all; teachers can view the list too (read-only) when allotting isn't relevant
  _ensureAttSubjectsSheet();
  return readTable(SH_ATT_SUBJECTS).map(r=>({
    subjectKey:String(r.subjectKey||'').trim(), subjectName:String(r.subjectName||'').trim(),
    subjectCode:String(r.subjectCode||'').trim(),
    type:(String(r.type||'TH').trim().toUpperCase()==='PR'?'PR':'TH'),
    semester:String(r.semester||'').trim()
  })).filter(r=>r.subjectKey);
}

function addAttSubject(token, subjectKey, subjectName, subjectCode, type, semester){
  _requireAdmin(token);
  subjectKey=String(subjectKey||'').trim();
  subjectName=String(subjectName||'').trim();
  if(!subjectKey || !subjectName) return {ok:false,msg:'Subject key and name are required'};
  _ensureAttSubjectsSheet();
  const sh=getSheet(SH_ATT_SUBJECTS);
  const existing=readTable(SH_ATT_SUBJECTS);
  if(existing.some(r=>String(r.subjectKey).trim()===subjectKey)) return {ok:false,msg:'That subject key already exists'};
  sh.appendRow([subjectKey, subjectName, String(subjectCode||'').trim(),
                (String(type||'TH').trim().toUpperCase()==='PR'?'PR':'TH'), String(semester||'').trim()]);
  return {ok:true};
}

function deleteAttSubject(token, subjectKey){
  _requireAdmin(token);
  _ensureAttSubjectsSheet();
  _deleteRowsWhere(SH_ATT_SUBJECTS, 'subjectKey', v=>v===String(subjectKey).trim());
  // Also clear any teacher allotments pointing at the removed subject, so stale entries don't linger.
  _ensureTeacherSheet();
  _deleteRowsWhere(SH_TEACHER_SUBJ, 'subject', v=>v===String(subjectKey).trim());
  return {ok:true};
}

function getAttMeta(token){
  const SUBJECTS=_subjects();
  const mine = _mySubjects(token); // admins get every subject, always; teachers only what's been allotted to them
  const mk=(k)=>({key:k,label:k+'  —  '+SUBJECTS[k].name,type:SUBJECTS[k].type,semester:SUBJECTS[k].semester});
  const subs=Object.keys(SUBJECTS).filter(k=>mine.indexOf(k)>=0).map(mk);
  return { subjects:subs, takeSubjects:subs, institute:INSTITUTE_NAME, dept:DEPT_NAME, className:CLASS_NAME, threshold:PASS_THRESHOLD };
}

function getStudentsFor(token, subject, batch){
  _requireStaff(token);
  _assertSubjectAllowed(token, subject);
  const SUBJECTS=_subjects();
  // Students belong to whichever class the subject itself is tagged with (its "semester" field),
  // not a single hardcoded class — this is what lets different subjects serve different classes
  // (e.g. a CO-1-K subject pulls CO-1-K students, a CO-5-K subject pulls CO-5-K students).
  const norm=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
  const subjClass = (SUBJECTS[subject] && SUBJECTS[subject].semester) ? SUBJECTS[subject].semester : CLASS_NAME_SHORT;
  const TARGET=norm(subjClass);
  const isDefaultClass = TARGET===norm(CLASS_NAME_SHORT);
  let list=readTable(SH_STUDENTS)
    .filter(s=> norm(s.className)===TARGET || (isDefaultClass && String(s.className||'').trim()===''))
    .map(s=>({roll:String(s.rollNo),enrollment:String(s.enrollment),name:_s(s.name),batch:String(s.batch||'').replace('*',''),
               spi:String(s.spi||'').trim().toUpperCase()==='Y'}));
  list.sort((a,b)=>Number(a.roll)-Number(b.roll));
  const isPR = SUBJECTS[subject] && SUBJECTS[subject].type==='PR';
  // SPI is taken by only a fixed, pre-selected subset of the class (set on the Students &
  // Batches screen) — it is NOT the whole class, and NOT filtered by A/B/C batch.
  const isSPI = subject==='SPI-PR';
  if(isSPI){ list=list.filter(s=>s.spi); }
  else if(isPR && batch && batch!=='ALL'){ list=list.filter(s=>s.batch===batch); }
  return list;
}

// ====== LECTURE PLANNING (fill in timetable dates in advance) ======
const SKIP_REASONS = ['Leave','Exchanged','Holiday','Other'];

// Self-heals: creates the PlannedSessions sheet if it doesn't exist yet (older deployments).
function _planSheet(){
  const ss=_openSS();
  let sh=ss.getSheetByName(SH_PLAN);
  if(!sh){ sh=ss.insertSheet(SH_PLAN); sh.appendRow(['planId','subject','date','time','batch','status','skipReason','sessionId','createdAt']); }
  return sh;
}

// Add ONE planned lecture date (manual entry).
function addPlannedSession(token, subject, date, time, batch){
  _requireStaff(token);
  if(!_subjects()[subject]) return {ok:false,msg:'Unknown subject'};
  _assertSubjectAllowed(token, subject);
  if(!date) return {ok:false,msg:'Date is required'};
  _planSheet();
  const dup=readTable(SH_PLAN).some(p=>p.subject===subject && _d(p.date)===date && _s(p.time)===(time||'') &&
                                        _s(p.batch)===(batch||'') && String(p.status||'PENDING')==='PENDING');
  if(dup) return {ok:false,msg:'That date is already planned for this subject.'};
  const sh=getSheet(SH_PLAN);
  const id='PL'+Date.now()+Math.floor(Math.random()*1000);
  sh.appendRow([id, subject, date, time||'', batch||'', 'PENDING', '', '', new Date()]);
  return {ok:true, planId:id};
}

// Bulk-generate planned dates across a date range for chosen weekdays (0=Sun..6=Sat), matching the timetable.
function bulkPlanSessions(token, subject, dayTimes, batch, fromDate, toDate){
  _requireStaff(token);
  if(!_subjects()[subject]) return {ok:false,msg:'Unknown subject'};
  _assertSubjectAllowed(token, subject);
  if(!fromDate||!toDate) return {ok:false,msg:'From and To dates are required'};
  if(!dayTimes || !dayTimes.length) return {ok:false,msg:'Select at least one weekday and its time'};
  _planSheet();
  const existing=readTable(SH_PLAN).filter(p=>p.subject===subject && String(p.status||'PENDING')==='PENDING')
                    .map(p=>_d(p.date)+'|'+_s(p.time)+'|'+_s(p.batch));
  const sh=getSheet(SH_PLAN);
  // dayTimes: [{day:'1', time:'09:00'}, {day:'3', time:'11:00'}, ...] — each weekday can have its own lecture time.
  const dayMap={};
  dayTimes.forEach(dt=>{ dayMap[Number(dt.day)] = dt.time||''; });
  const rows=[];
  let d=new Date(fromDate+'T00:00:00');
  const end=new Date(toDate+'T00:00:00');
  while(d<=end){
    const dow=d.getDay();
    if(dayMap.hasOwnProperty(dow)){
      const ds=Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      const time=dayMap[dow];
      const key=ds+'|'+time+'|'+(batch||'');
      if(existing.indexOf(key)<0){
        const id='PL'+Date.now()+Math.floor(Math.random()*100000);
        rows.push([id, subject, ds, time, batch||'', 'PENDING', '', '', new Date()]);
        existing.push(key);
      }
    }
    d.setDate(d.getDate()+1);
  }
  if(rows.length) sh.getRange(sh.getLastRow()+1,1,rows.length,9).setValues(rows);
  return {ok:true, added:rows.length};
}

// PENDING planned dates for a subject (any date, past or future) — feeds the Take Attendance picker.
function listUpcomingPlans(token, subject){
  _requireStaff(token);
  _assertSubjectAllowed(token, subject);
  _planSheet();
  let list=readTable(SH_PLAN).filter(p=>p.subject===subject && String(p.status||'PENDING')==='PENDING')
              .map(p=>({planId:String(p.planId), subject:_s(p.subject), date:_d(p.date), time:_t(p.time), batch:_s(p.batch)}));
  list.sort((a,b)=> a.date<b.date?-1:(a.date>b.date?1:0));
  return list;
}

// ALL planned entries for a subject (any status) — feeds the admin planning/management list.
function listAllPlans(token, subject){
  _requireStaff(token);
  _assertSubjectAllowed(token, subject);
  _planSheet();
  let list=readTable(SH_PLAN).filter(p=>p.subject===subject)
              .map(p=>({planId:String(p.planId), subject:_s(p.subject), date:_d(p.date), time:_t(p.time), batch:_s(p.batch),
                        status:_s(p.status)||'PENDING', skipReason:_s(p.skipReason), sessionId:_s(p.sessionId)}));
  list.sort((a,b)=> a.date<b.date?-1:(a.date>b.date?1:0));
  return list;
}

function deletePlannedSession(token, planId){
  _requireStaff(token);
  _planSheet();
  const plan=readTable(SH_PLAN).find(p=>String(p.planId)===String(planId));
  if(plan) _assertSubjectAllowed(token, plan.subject);
  _deleteRowsWhere(SH_PLAN,'planId',v=>v===String(planId));
  return {ok:true};
}

// Mark a planned lecture as NOT conducted (Leave/Exchanged/Holiday/Other). No attendance is written —
// this date is simply excluded, same as a student's own Leave mark never counts toward totals.
function markPlanNotConducted(token, planId, reason){
  _requireStaff(token);
  if(SKIP_REASONS.indexOf(reason)<0) return {ok:false,msg:'Invalid reason'};
  _planSheet();
  const planRow=readTable(SH_PLAN).find(p=>String(p.planId)===String(planId));
  if(planRow) _assertSubjectAllowed(token, planRow.subject);
  const sh=getSheet(SH_PLAN);
  const data=sh.getDataRange().getValues();
  const head=data[0].map(x=>String(x).trim());
  const idCol=head.indexOf('planId'), stCol=head.indexOf('status'), rsCol=head.indexOf('skipReason');
  for(let i=1;i<data.length;i++){
    if(String(data[i][idCol])===String(planId)){
      sh.getRange(i+1,stCol+1).setValue('SKIPPED');
      sh.getRange(i+1,rsCol+1).setValue(reason);
      return {ok:true};
    }
  }
  return {ok:false,msg:'Planned session not found'};
}

function saveAttendance(token, payload){
  const s=_requireStaff(token);
  const {subject,date,time,batch,topic,marks,sessionId,planId}=payload;
  if(!subject||!date) return {ok:false,msg:'Subject and date are required'};
  if(!_subjects()[subject]) return {ok:false,msg:'Unknown subject'};
  _assertSubjectAllowed(token, subject);

  const sessSheet=getSheet(SH_SESSIONS);
  const markSheet=getSheet(SH_ATTMARKS);
  let sid=sessionId;

  // map header -> column index (1-based) so writes never depend on column order
  const head=sessSheet.getRange(1,1,1,sessSheet.getLastColumn()).getValues()[0].map(x=>String(x).trim());
  const col=name=>head.indexOf(name)+1;  // 0 if missing
  const setCell=(rowIdx,name,val)=>{ const c=col(name); if(c>0) sessSheet.getRange(rowIdx,c).setValue(val); };

  if(sid){
    // editing existing: remove old marks for this session
    _deleteRowsWhere(SH_ATTMARKS,'sessionId',v=>v===String(sid));
    const data=sessSheet.getDataRange().getValues();
    for(let i=1;i<data.length;i++){
      if(String(data[i][0])===String(sid)){
        const rowIdx=i+1;
        setCell(rowIdx,'date',date);
        setCell(rowIdx,'time',time||'');
        setCell(rowIdx,'batch',batch||'');
        setCell(rowIdx,'topic',topic||'');
        break;
      }
    }
  } else {
    sid='S'+Date.now();
    const rowIdx=sessSheet.getLastRow()+1;
    // build a full-width row using header positions
    const rowVals=new Array(head.length).fill('');
    const put=(name,val)=>{ const c=col(name); if(c>0) rowVals[c-1]=val; };
    put('sessionId',sid); put('subject',subject); put('date',date); put('time',time||'');
    put('batch',batch||''); put('topic',topic||''); put('takenBy',s.username); put('createdAt',new Date());
    sessSheet.getRange(rowIdx,1,1,head.length).setValues([rowVals]);
  }
  // write marks
  const rows=(marks||[]).map(m=>[sid,String(m.roll),(m.status==='A'||m.status==='L')?m.status:'P']);
  if(rows.length) markSheet.getRange(markSheet.getLastRow()+1,1,rows.length,3).setValues(rows);

  // If this attendance came from a planned date, mark that plan entry DONE and link it to the session.
  if(planId){
    _planSheet();
    const psh=getSheet(SH_PLAN);
    const pdata=psh.getDataRange().getValues();
    const phead=pdata[0].map(x=>String(x).trim());
    const pIdCol=phead.indexOf('planId'), pStCol=phead.indexOf('status'), pSidCol=phead.indexOf('sessionId');
    if(pIdCol>=0 && pStCol>=0){
      for(let i=1;i<pdata.length;i++){
        if(String(pdata[i][pIdCol])===String(planId)){
          psh.getRange(i+1,pStCol+1).setValue('DONE');
          if(pSidCol>=0) psh.getRange(i+1,pSidCol+1).setValue(sid);
          break;
        }
      }
    }
  }
  return {ok:true, sessionId:sid, count:rows.length};
}

function listSessions(token, subject){
  _requireStaff(token);
  const mine=_mySubjects(token);
  let s=readTable(SH_SESSIONS).map(x=>({
    sessionId:String(x.sessionId), subject:_s(x.subject), date:_d(x.date), time:_t(x.time), batch:_s(x.batch), topic:_s(x.topic), takenBy:_s(x.takenBy)
  }));
  if(subject && subject!=='ALL'){
    _assertSubjectAllowed(token, subject);
    s=s.filter(x=>x.subject===subject);
  } else {
    s=s.filter(x=>mine.indexOf(x.subject)>=0); // 'ALL' means all of MY subjects, not every subject in the system
  }
  s.sort((a,b)=> (a.date<b.date?1:-1));
  return s;
}

function getSession(token, sessionId){
  _requireStaff(token);
  const sess=readTable(SH_SESSIONS).find(x=>String(x.sessionId)===String(sessionId));
  if(!sess) return {ok:false,msg:'Session not found'};
  _assertSubjectAllowed(token, sess.subject);
  const marks=readTable(SH_ATTMARKS).filter(m=>String(m.sessionId)===String(sessionId))
                .map(m=>({roll:String(m.roll),status:m.status}));
  const markMap={}; marks.forEach(m=>markMap[m.roll]=m.status);
  const students=getStudentsFor(token, sess.subject, sess.batch);
  students.forEach(st=> st.status = markMap[st.roll] || 'P');
  return {ok:true, session:{sessionId:String(sess.sessionId),subject:_s(sess.subject),date:_d(sess.date),time:_t(sess.time),batch:_s(sess.batch),topic:_s(sess.topic)}, students:students};
}

function deleteSession(token, sessionId){
  _requireStaff(token);
  const sess=readTable(SH_SESSIONS).find(x=>String(x.sessionId)===String(sessionId));
  if(sess) _assertSubjectAllowed(token, sess.subject);
  _deleteRowsWhere(SH_SESSIONS,'sessionId',v=>v===String(sessionId));
  _deleteRowsWhere(SH_ATTMARKS,'sessionId',v=>v===String(sessionId));
  return {ok:true};
}

function buildRegister(token, subject, fromDate, toDate, batch){
  _requireStaff(token);
  _assertSubjectAllowed(token, subject);
  const SUBJECTS=_subjects();
  let sessions=readTable(SH_SESSIONS).filter(x=>x.subject===subject).map(x=>({
    sessionId:String(x.sessionId),date:_d(x.date),time:_t(x.time),batch:_s(x.batch),topic:_s(x.topic) }));
  if(fromDate) sessions=sessions.filter(s=>s.date>=fromDate);
  if(toDate)   sessions=sessions.filter(s=>s.date<=toDate);

  // which students: PR subject uses union of batches that appear in sessions; TH = whole class
  const isPR = SUBJECTS[subject] && SUBJECTS[subject].type==='PR';
  // SPI is taken as one session for the whole class (session.batch stores 'SPI', not A/B/C),
  // so it must NOT be treated as batch-restricted like OSY/ACN practicals — otherwise every
  // student gets filtered out (their batch is A/B/C, never 'SPI') and the report looks empty.
  const isSPI = subject==='SPI-PR';
  // Batch-wise report filter: only meaningful for non-SPI PR subjects. 'ALL' or blank = combined (old behaviour).
  const batchFilter = (isPR && !isSPI && batch && String(batch).toUpperCase()!=='ALL') ? String(batch).trim() : '';
  if(batchFilter) sessions=sessions.filter(s=>s.batch===batchFilter);

  sessions.sort((a,b)=> a.date<b.date?-1:(a.date>b.date?1: (String(a.time)<String(b.time)?-1:1)));

  // Also pull in PLANNED dates that haven't become a real session yet (still Pending, or marked
  // Not Conducted) so the report shows the full picture — what's planned, what's done, what's
  // skipped and why — not just dates attendance was actually taken for.
  _planSheet();
  let plans=readTable(SH_PLAN).filter(p=>p.subject===subject && String(p.status||'PENDING')!=='DONE')
              .map(p=>({planId:String(p.planId), date:_d(p.date), time:_t(p.time), batch:_s(p.batch),
                        status:String(p.status||'PENDING'), skipReason:_s(p.skipReason)}));
  if(fromDate) plans=plans.filter(p=>p.date>=fromDate);
  if(toDate)   plans=plans.filter(p=>p.date<=toDate);
  if(batchFilter) plans=plans.filter(p=>p.batch===batchFilter);

  const allMarks=readTable(SH_ATTMARKS);
  const markBy={}; // sessionId -> {roll:status}
  allMarks.forEach(m=>{ const sid=String(m.sessionId); (markBy[sid]=markBy[sid]||{})[String(m.roll)]=m.status; });

  const _normCls=v=>String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,'');
  const _subjClass = (SUBJECTS[subject] && SUBJECTS[subject].semester) ? SUBJECTS[subject].semester : CLASS_NAME_SHORT;
  const _TARGET=_normCls(_subjClass);
  // Reports show only students of the class this subject belongs to.
  let students=readTable(SH_STUDENTS)
    .filter(s=> _normCls(s.className)===_TARGET)
    .map(s=>({roll:String(s.rollNo),name:_s(s.name),batch:String(s.batch||'').replace('*',''),
               spi:String(s.spi||'').trim().toUpperCase()==='Y'}));
  students.sort((a,b)=>Number(a.roll)-Number(b.roll));

  if(isSPI){
    // SPI is taken by only the pre-selected subset of the class, not everyone.
    students=students.filter(s=>s.spi);
  } else if(isPR){
    if(batchFilter){
      // Explicit batch chosen for the report: show only that batch's students.
      students=students.filter(s=>s.batch===batchFilter);
    } else {
      // Combined view: restrict to students whose batch was actually taken OR planned (avoids empty rows)
      const batches=new Set(sessions.map(s=>s.batch).filter(Boolean).concat(plans.map(p=>p.batch).filter(Boolean)));
      if(batches.size) students=students.filter(s=>batches.has(s.batch));
    }
  }

  // Real (taken) sessions first, then planned/not-conducted dates as extra columns with no marks yet.
  // A synthetic 'PLAN_...' id never matches any real AttMarks row, so those cells come out blank
  // automatically via the same cell logic below — no attendance is implied for them.
  const cols=sessions.map(s=>({sessionId:s.sessionId,date:s.date,time:s.time,batch:s.batch,topic:s.topic||''}))
    .concat(plans.map(p=>({sessionId:'PLAN_'+p.planId,date:p.date,time:p.time,batch:p.batch,
       topic: p.status==='SKIPPED' ? ('Not Conducted — '+(p.skipReason||'Other')) : 'Planned — not yet taken'})));
  cols.sort((a,b)=> a.date<b.date?-1:(a.date>b.date?1: (String(a.time)<String(b.time)?-1:1)));

  const rows=students.map(st=>{
    let present=0, total=0;
    const cells=cols.map(c=>{
      const status=(markBy[c.sessionId]||{})[st.roll];
      if(status===undefined){ return ''; }      // student not part of that session / date not yet taken
      if(status!=='L'){ total++; if(status==='P') present++; }  // Leave excluded from total entirely
      return status;
    });
    const pct = total? Math.round(present*1000/total)/10 : 0;
    return {roll:st.roll,name:st.name,batch:st.batch,cells:cells,present:present,total:total,pct:pct};
  });
  return { subject:subject, subjectName:SUBJECTS[subject].name, code:SUBJECTS[subject].code,
           type:SUBJECTS[subject].type, className:_subjClass, batch:batchFilter, columns:cols, rows:rows, threshold:PASS_THRESHOLD };
}

function _registerHtml(reg, opts){
  opts=opts||{};
  const dateLabel = (d)=>{ const p=String(d).split('-'); return p.length===3? (p[2]+'/'+p[1]) : d; };
  let h='<html><head><meta charset="utf-8">'
   + '<style>@page{ size: A4 portrait; margin: 10mm; } body{ font-family:\'Times New Roman\',Times,serif; color:#000; font-size:10pt; } thead{ display: table-header-group; }</style>'
   + '</head><body>';
  h+='<div style="text-align:center;">'
   + '<img src="'+LOGO_DATA_URI+'" style="height:60px;"><br>'
   + '<div style="font-size:16pt;font-weight:bold;">'+INSTITUTE_NAME+'</div>'
   + '<div style="font-size:12pt;font-weight:600;">'+DEPT_NAME+'</div>'
   + '<div style="font-size:10pt;margin-top:3px;">Attendance Register &mdash; '+reg.subjectName+' ('+reg.code+') ['+reg.subject+']'+(reg.batch?(' &mdash; Batch '+reg.batch):'')+'</div>'
   + '<div style="font-size:10pt;color:#444;">'+(reg.className||CLASS_NAME)+'</div>'
   + '</div>';
  if(opts.defaulters){
    h+='<div style="font-size:12pt;font-weight:bold;color:#a31515;margin-top:8px;">Defaulter List (Below '+reg.threshold+'%)</div>';
  }

  const thStyle='border:1px solid #888;padding:4px;font-size:10pt;';
  const tdStyle='border:1px solid #aaa;padding:4px;font-size:10pt;';
  const colsPerPageUsed = Number(opts.colsPerPage)>0 ? Number(opts.colsPerPage) : 9;

  if(opts.defaulters){
    // Never wide (no date columns) — a single table is fine.
    let rows=reg.rows.filter(r=>r.total>0 && r.pct<reg.threshold);
    h+='<table style="border-collapse:collapse;margin-top:10px;">';
    h+='<tr style="background:#1A4D8F;color:#fff;">'
     + '<th style="'+thStyle+'">Roll</th><th style="'+thStyle+'text-align:left;white-space:nowrap;">Name</th>'
     + '<th style="'+thStyle+'">P</th><th style="'+thStyle+'">Tot</th><th style="'+thStyle+'">%</th></tr>';
    rows.forEach(r=>{
      h+='<tr style="background:#fde8e8;">'
       + '<td style="'+tdStyle+'text-align:center;">'+r.roll+'</td>'
       + '<td style="'+tdStyle+'white-space:nowrap;">'+r.name+'</td>'
       + '<td style="'+tdStyle+'text-align:center;">'+r.present+'</td>'
       + '<td style="'+tdStyle+'text-align:center;">'+r.total+'</td>'
       + '<td style="'+tdStyle+'text-align:center;color:#c0392b;font-weight:bold;">'+r.pct+'</td></tr>';
    });
    h+='</table>';
  } else {
    // A wide register (many date columns) can't fit one page width at a readable font size.
    // Split the date columns into page-sized chunks. Roll repeats on every chunk so each block is
    // still identifiable, but Name is only shown on the first chunk — it's already established
    // there, so repeating it on every continuation block just wastes width. Totals appear once,
    // on the final chunk, so nothing gets clipped. No forced page breaks between chunks — they
    // flow naturally one after another, so short registers don't waste blank space with an
    // artificial break; the PDF renderer only starts a new page when a block actually doesn't fit.
    const COLS_PER_PAGE = colsPerPageUsed;
    const chunks=[];
    for(let i=0;i<reg.columns.length;i+=COLS_PER_PAGE) chunks.push(reg.columns.slice(i,i+COLS_PER_PAGE));
    if(!chunks.length) chunks.push([]); // no sessions at all — still show Roll/Name/P/Tot/%

    chunks.forEach((chunkCols,ci)=>{
      const isLast=ci===chunks.length-1;
      const showName=true; // Name repeats alongside Roll on every page
      // Each date-block starts fresh at the top of a full page — never partway down leftover
      // space from the previous block. If a block itself has more students than fit on one
      // physical page, the header row (thead) automatically repeats on the continuation page.
      // Each date-block is its own table with its own repeating header (thead). Within a block,
      // rows continue naturally across pages if the roster is long (with the header repeating).
      // But a NEW block always starts on a fresh page — otherwise its roll 1 would land directly
      // under the previous block's roll 68 on the same page, which reads as one broken table.
      h+='<table style="border-collapse:collapse;margin-top:6px;'+(ci>0?'page-break-before:always;':'')+'">';
      h+='<thead><tr style="background:#1A4D8F;color:#fff;">'
       + '<th style="'+thStyle+'">Roll</th>'
       + (showName? '<th style="'+thStyle+'text-align:left;white-space:nowrap;">Name</th>' : '');
      chunkCols.forEach(c=>{ h+='<th style="'+thStyle+'">'+dateLabel(c.date)+(c.batch?('<br>'+c.batch):'')+'</th>'; });
      if(isLast){
        h+='<th style="'+thStyle+'">P</th><th style="'+thStyle+'">Tot</th><th style="'+thStyle+'">%</th>';
      }
      h+='</tr></thead><tbody>';
      reg.rows.forEach(r=>{
        const low = r.total>0 && r.pct<reg.threshold;
        h+='<tr'+(low?' style="background:#fde8e8;"':'')+'>'
         + '<td style="'+tdStyle+'text-align:center;">'+r.roll+'</td>'
         + (showName? '<td style="'+tdStyle+'white-space:nowrap;">'+r.name+'</td>' : '');
        const startIdx=ci*COLS_PER_PAGE;
        chunkCols.forEach((c,j)=>{
          const cell=r.cells[startIdx+j];
          const col = cell==='A'? 'color:#c0392b;font-weight:bold;' : (cell==='P'?'color:#1c6b34;':(cell==='L'?'color:#b8860b;font-weight:bold;':''));
          h+='<td style="border:1px solid #ccc;padding:4px;font-size:10pt;text-align:center;'+col+'">'+(cell||'-')+'</td>';
        });
        if(isLast){
          h+='<td style="'+tdStyle+'text-align:center;">'+r.present+'</td>'
           + '<td style="'+tdStyle+'text-align:center;">'+r.total+'</td>'
           + '<td style="'+tdStyle+'text-align:center;'+(low?'color:#c0392b;font-weight:bold;':'')+'">'+r.pct+'</td>';
        }
        h+='</tr>';
      });
      h+='</tbody></table>';
    });
  }

  h+='<p style="font-size:10pt;color:#666;margin-top:8px;">P = Present, A = Absent, L = Leave (excluded from Total/%), - = not part of that session'
   +(reg.columns.length>colsPerPageUsed && !opts.defaulters?' &mdash; wide registers continue below/on the next page as needed; Roll and Name repeat on every page, totals are with the last block.':'')
   +'. Generated '+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'dd/MM/yyyy HH:mm')+'.</p>';

  // Topics Conducted (date-wise) — not shown on the defaulters-only sheet
  if(!opts.defaulters){
    const dateLabel2=(d)=>{ const p=String(d).split('-'); return p.length===3? (p[2]+'/'+p[1]+'/'+p[0]) : d; };
    h+='<div style="font-size:12pt;font-weight:bold;margin-top:16px;page-break-before:always;">Topic Conducted</div>';
    h+='<table style="border-collapse:collapse;width:100%;margin-top:5px;">';
    h+='<tr style="background:#eef2f7;">'
     + '<th style="border:1px solid #aaa;padding:4px;font-size:10pt;width:90px;">Date</th>'
     + '<th style="border:1px solid #aaa;padding:4px;font-size:10pt;width:60px;">Time</th>'
     + '<th style="border:1px solid #aaa;padding:4px;font-size:10pt;width:60px;">Batch</th>'
     + '<th style="border:1px solid #aaa;padding:4px;font-size:10pt;text-align:left;">Topic Conducted</th></tr>';
    reg.columns.forEach(c=>{
      h+='<tr>'
       + '<td style="border:1px solid #ccc;padding:4px;font-size:10pt;text-align:center;">'+dateLabel2(c.date)+'</td>'
       + '<td style="border:1px solid #ccc;padding:4px;font-size:10pt;text-align:center;">'+(c.time||'')+'</td>'
       + '<td style="border:1px solid #ccc;padding:4px;font-size:10pt;text-align:center;">'+(c.batch||'-')+'</td>'
       + '<td style="border:1px solid #ccc;padding:4px;font-size:10pt;">'+(c.topic||'')+'</td></tr>';
    });
    h+='</table>';
  }

  h+='</body></html>';
  return h;
}

function generateReportPdf(token, subject, fromDate, toDate, defaultersOnly, batch, colsPerPage){
  _requireStaff(token);
  const reg=buildRegister(token, subject, fromDate, toDate, batch);
  if(!reg.columns.length) return {ok:false,msg:'No attendance sessions found for this subject/date range'+(reg.batch?(' and Batch '+reg.batch):'')+'.'};
  const html=_registerHtml(reg, {defaulters:!!defaultersOnly, colsPerPage:colsPerPage});
  const out=HtmlService.createHtmlOutput(html).getBlob().getAs('application/pdf')
              .setName('Attendance_'+subject+(reg.batch?('_Batch'+reg.batch):'')+(defaultersOnly?'_Defaulters':'')+'.pdf');
  const file=_reportFolder().createFile(out);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  const id=file.getId();
  return {ok:true, url:'https://drive.google.com/uc?export=download&id='+id,
          previewUrl:'https://drive.google.com/file/d/'+id+'/preview', name:file.getName()};
}

function generateReportExcel(token, subject, fromDate, toDate, batch){
  _requireStaff(token);
  const reg=buildRegister(token, subject, fromDate, toDate, batch);
  if(!reg.columns.length) return {ok:false,msg:'No attendance sessions found for this subject/date range'+(reg.batch?(' and Batch '+reg.batch):'')+'.'};
  // Build a temporary Google Sheet, then export as xlsx
  const tmp=SpreadsheetApp.create('ATT_'+subject+(reg.batch?('_'+reg.batch):'')+'_'+Date.now());
  const sh=tmp.getActiveSheet();
  const header=['Roll','Name'].concat(reg.columns.map(c=>c.date+(c.batch?(' ['+c.batch+']'):''))).concat(['Present','Total','%']);
  const data=[header];
  reg.rows.forEach(r=> data.push([r.roll,r.name].concat(r.cells).concat([r.present,r.total,r.pct])));
  sh.getRange(1,1,data.length,header.length).setValues(data);
  sh.getRange(1,1,1,header.length).setFontWeight('bold');

  // Topics Conducted block, a couple of rows below the register
  let tr=data.length+3;
  sh.getRange(tr,1).setValue('Topic Conducted').setFontWeight('bold');
  tr++;
  sh.getRange(tr,1,1,4).setValues([['Date','Time','Batch','Topic Conducted']]).setFontWeight('bold');
  tr++;
  const topicRows=reg.columns.map(c=>[c.date,(c.time||''),(c.batch||''),(c.topic||'')]);
  if(topicRows.length){ sh.getRange(tr,1,topicRows.length,4).setValues(topicRows); }
  SpreadsheetApp.flush();
  const url='https://docs.google.com/spreadsheets/d/'+tmp.getId()+'/export?format=xlsx';
  const blob=UrlFetchApp.fetch(url,{headers:{Authorization:'Bearer '+ScriptApp.getOAuthToken()}}).getBlob()
              .setName('Attendance_'+subject+(reg.batch?('_Batch'+reg.batch):'')+'.xlsx');
  const file=_reportFolder().createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  try{ DriveApp.getFileById(tmp.getId()).setTrashed(true); }catch(e){}
  const id=file.getId();
  return {ok:true, url:'https://drive.google.com/uc?export=download&id='+id, name:file.getName()};
}

function getRegister(token, subject, fromDate, toDate, batch){
  _requireStaff(token);
  return buildRegister(token, subject, fromDate, toDate, batch);
}

// Ensure the AttSummary tab exists (self-heals without re-running setupSheets).
function _attSummarySheet(){
  const ss=_openSS();
  let sh=ss.getSheetByName(SH_ATTSUMMARY);
  if(!sh){ sh=ss.insertSheet(SH_ATTSUMMARY); sh.appendRow(['enrollment','className','subject','present','total','pct']); }
  return sh;
}

// small class-normaliser used below and in getMyAttendance
function _normCls(v){ return String(v||'').toUpperCase().replace(/[^A-Z0-9]/g,''); }

// Upload an attendance SUMMARY (already-calculated) for any class.
// Wide file: Roll No, Enrollment, Name, then per subject two columns:
//   "<Subject> Present" and "<Subject> %"   (Total is derived from Present and %).
// Stored per student-subject and shown in the student's login, separately from any
// live app-tracked attendance. Uploading REPLACES the summary for the classes present
// in the file (other classes' summaries are left untouched).
function uploadAttendanceSummary(token, payload){
  _requireAdmin(token);
  if(!payload || !payload.base64) return {ok:false,msg:'No file received'};
  let tempId=null;
  try{
    const bytes=Utilities.base64Decode(payload.base64);
    const mime=payload.mimeType||MimeType.MICROSOFT_EXCEL;
    const blob=Utilities.newBlob(bytes,mime,payload.fileName||'upload');
    const name=String(payload.fileName||'').toLowerCase();
    const isCsv = name.endsWith('.csv') || mime==='text/csv' || mime==='application/csv';
    let rows;
    if(isCsv){
      rows=Utilities.parseCsv(blob.getDataAsString());
    } else {
      // Excel needs conversion to a temporary Google Sheet via the advanced Drive service.
      // If it is not enabled, return a clear instruction (CSV always works without it).
      if(typeof Drive==='undefined' || !Drive.Files){
        return {ok:false, msg:'To upload an Excel file, enable the Drive service once: in the Apps Script editor, click Services (+) in the left sidebar, add "Drive API", Save, then redeploy. Or simply save your file as CSV (File → Save As → CSV) and upload that — CSV needs no setup.'};
      }
      let created;
      if(typeof Drive.Files.insert==='function'){
        created=Drive.Files.insert({title:'TMP_'+Date.now(), mimeType:MimeType.GOOGLE_SHEETS}, blob, {convert:true});
      } else {
        created=Drive.Files.create({name:'TMP_'+Date.now(), mimeType:MimeType.GOOGLE_SHEETS}, blob);
      }
      tempId = created.id || created.getId && created.getId();
      if(!tempId) return {ok:false, msg:'Excel conversion failed. Please save the file as CSV and upload again.'};
      rows=SpreadsheetApp.openById(tempId).getSheets()[0].getDataRange().getDisplayValues();
    }
    if(!rows||rows.length<2) return {ok:false,msg:'File has no data rows'};

    const header=rows[0].map(h=>String(h).trim());
    const lower=header.map(h=>h.toLowerCase());
    const find=cands=>{ for(let i=0;i<lower.length;i++) for(const c of cands) if(lower[i].indexOf(c)>=0) return i; return -1; };
    const cEnr=find(['enroll','enrol']); const cClass=find(['class']);
    if(cEnr<0) return {ok:false,msg:'Header must contain an Enrollment column.'};

    // Identify subject column groups: "<Subject> Present", "<Subject> Total", "<Subject> %".
    const subjCols={};
    header.forEach((h,i)=>{
      const hl=h.toLowerCase();
      let kind=null, label=null;
      if(hl.indexOf('present')>=0){ kind='present'; label=h.replace(/present/ig,'').trim(); }
      else if(hl.indexOf('total')>=0){ kind='total'; label=h.replace(/total/ig,'').trim(); }
      else if(hl.indexOf('%')>=0 || hl.indexOf('percent')>=0){ kind='pct'; label=h.replace(/percentage|percent|%/ig,'').trim(); }
      if(kind && label){ label=label.replace(/[-_:]+$/,'').trim(); (subjCols[label]=subjCols[label]||{})[kind]=i; }
    });
    const subjects=Object.keys(subjCols);
    if(!subjects.length) return {ok:false,msg:'No subject columns found. Use headers like "OSY Present", "OSY Total" and "OSY %".'};

    function num(v){ const n=parseFloat(String(v==null?'':v).replace(/[^0-9.\-]/g,'')); return isNaN(n)?null:n; }
    function normId(v){ let s=String(v==null?'':v).trim(); if(/[eE]\+?\d|\.\d+/.test(s)&&!isNaN(Number(s))){const n=Number(s); if(isFinite(n)) s=n.toLocaleString('en-US',{useGrouping:false,maximumFractionDigits:0});} return s; }

    const classesInFile={};
    const outRows=[];
    for(let r=1;r<rows.length;r++){
      const enr=normId(rows[r][cEnr]); if(!enr) continue;
      const cls=cClass>=0 ? String(rows[r][cClass]||'').trim() : '';
      if(cls) classesInFile[_normCls(cls)]=true;
      subjects.forEach(label=>{
        const pc=subjCols[label];
        const present = pc.present!==undefined ? num(rows[r][pc.present]) : null;
        let total     = pc.total!==undefined   ? num(rows[r][pc.total])   : null;
        let pct       = pc.pct!==undefined     ? num(rows[r][pc.pct])     : null;
        // Prefer the Total given in the file. Compute % from Present/Total when possible.
        if(total!==null && total>0 && present!==null){ pct=Math.round(present*1000/total)/10; }
        // If no Total column but Present and % given, derive Total (fallback, may round).
        else if(total===null && present!==null && pct && pct>0){ total=Math.round(present*100/pct); }
        if(present===null && total===null && pct===null) return; // nothing for this subject/student
        outRows.push([enr, cls, label, present===null?'':present, total===null?'':total, pct===null?'':pct]);
      });
    }
    if(!outRows.length) return {ok:false,msg:'No valid data rows found.'};

    const sh=_attSummarySheet();
    const existing = sh.getLastRow()>1 ? sh.getRange(2,1,sh.getLastRow()-1,6).getValues() : [];
    const enrInFile={}; outRows.forEach(r=> enrInFile[String(r[0]).trim()]=true);
    const kept = existing.filter(r=>{
      const rc=_normCls(r[1]);
      if(Object.keys(classesInFile).length){ return !classesInFile[rc]; }
      return !enrInFile[String(r[0]).trim()];
    });
    if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,6).clearContent();
    const finalRows=kept.concat(outRows);
    if(finalRows.length) sh.getRange(2,1,finalRows.length,6).setValues(finalRows);

    return {ok:true, students:Object.keys(enrInFile).length, subjects:subjects.length, rows:outRows.length};
  }catch(err){
    return {ok:false,msg:'Could not read file: '+err.message};
  }finally{
    if(tempId){ try{ DriveApp.getFileById(tempId).setTrashed(true);}catch(e){} }
  }
}

function clearAttendanceSummary(token){
  _requireAdmin(token);
  const sh=_attSummarySheet();
  if(sh.getLastRow()>1) sh.getRange(2,1,sh.getLastRow()-1,6).clearContent();
  return {ok:true};
}

function getMyAttendance(token){
  const s=_require(token);
  const SUBJECTS=_subjects();
  const en=String(s.username).trim();
  const me=readTable(SH_STUDENTS).find(x=>String(x.enrollment).trim()===en);
  if(!me) return {ok:false,msg:'Your record was not found. Please contact the office.'};

  // (1) LIVE app-tracked subjects — for any class, but ONLY subjects that belong to this student's
  // own class. Marks are stored by roll number, so without this a CO-5-K student could pick up a
  // CO-1-K subject's marks that happen to share their roll number.
  const myCls=_normCls(me.className);
  const myKeys=Object.keys(SUBJECTS).filter(k=> _normCls(SUBJECTS[k].semester||CLASS_NAME_SHORT)===myCls);
  const liveSubjects=[];
  if(myKeys.length){
    const myRoll=String(me.rollNo).trim();
    const marks=readTable(SH_ATTMARKS).filter(m=>String(m.roll).trim()===myRoll);
    const statusBySession={};
    marks.forEach(m=>{ statusBySession[String(m.sessionId).trim()]=m.status; });
    const sessions=readTable(SH_SESSIONS);
    const per={};
    myKeys.forEach(k=> per[k]={present:0,total:0});
    sessions.forEach(se=>{
      const sub=String(se.subject).trim();
      if(!per[sub]) return;
      const st=statusBySession[String(se.sessionId).trim()];
      if(st===undefined) return;
      if(st==='L') return; // Leave is excluded entirely — doesn't count toward total lectures
      per[sub].total++;
      if(st==='P') per[sub].present++;
    });
    myKeys.forEach(k=>{
      const p=per[k].present, t=per[k].total;
      if(t>0) liveSubjects.push({ key:k, name:SUBJECTS[k].name, code:SUBJECTS[k].code, type:SUBJECTS[k].type,
             present:p, total:t, pct: Math.round(p*1000/t)/10 });
    });
  }

  // (2) UPLOADED summary subjects — for any class (CO-1-K, CO-3-K, or CO-5-K extra subjects).
  const uploaded=readTable(SH_ATTSUMMARY).filter(r=>String(r.enrollment).trim()===en);
  const uploadedSubjects=uploaded.map(r=>{
    const p=r.present===''||r.present===undefined?null:Number(r.present);
    const t=r.total===''||r.total===undefined?null:Number(r.total);
    let pct=r.pct===''||r.pct===undefined?null:Number(r.pct);
    if(pct===null && p!==null && t){ pct=Math.round(p*1000/t)/10; }
    return { name:String(r.subject), present:(p===null?'':p), total:(t===null?'':t), pct:(pct===null?0:pct) };
  });

  if(!liveSubjects.length && !uploadedSubjects.length){
    return {ok:false, msg:'No attendance has been recorded for you yet. Please contact the office.'};
  }

  // Overall combines whatever numeric present/total we have (live + uploaded with totals).
  let P=0,T=0;
  liveSubjects.forEach(x=>{P+=x.present;T+=x.total;});
  uploadedSubjects.forEach(x=>{ if(x.total!=='' && x.present!==''){ P+=Number(x.present); T+=Number(x.total); } });

  return { ok:true, name:me.name, roll:String(me.rollNo), enrollment:en,
           className:String(me.className||''), batch:String(me.batch||'').replace('*',''),
           liveSubjects:liveSubjects, uploadedSubjects:uploadedSubjects,
           overall:{present:P,total:T,pct: T? Math.round(P*1000/T)/10:0},
           threshold:PASS_THRESHOLD };
}

function _reportFolder(){
  const name='Attendance_Reports';
  const it=DriveApp.getFoldersByName(name);
  return it.hasNext()?it.next():DriveApp.createFolder(name);
}


// Class name applied to uploaded students (Fifth Sem Computer Engg).
const CLASS_NAME_SHORT = 'CO-5-K';

// All classes the app can hold students for. CO-5-K is the class whose attendance is
// taken inside the app; the others are maintained via the uploaded attendance summary.
const ALL_CLASSES = ['CO-1-K', 'CO-3-K', 'CO-5-K'];
function getClassList(token){ _requireStaff(token); return ALL_CLASSES.slice(); }