/**
 * ส่งข้อสอบจาก Google Sheets ให้เว็บ "จ่าพัน สาย 1"
 *
 * วิธีใช้ (ทำครั้งเดียว)
 *  1) สร้าง Google Sheets ใหม่ → เมนู ส่วนขยาย → Apps Script → วางโค้ดนี้ทั้งไฟล์ → บันทึก
 *  2) รีเฟรชชีต แล้วเลือกเมนู "ข้อสอบ" → "1) สร้างชีตข้อสอบ"
 *  3) Deploy → New deployment → ประเภท Web app
 *       Execute as: Me   |   Who has access: Anyone
 *     คัดลอก URL ที่ลงท้าย /exec ไปใส่ examApiUrl ใน SITE_CONFIG ของ index.html
 *
 * หลังจากนั้นเพิ่ม/แก้/ลบข้อสอบในชีตได้เลย ไม่ต้อง Deploy ใหม่ (เว็บอัปเดตภายใน ~2 นาที
 * หรือสั่งเมนู "ล้างแคช" เพื่ออัปเดตทันที) ต้อง Deploy ใหม่ก็ต่อเมื่อแก้โค้ดในไฟล์นี้
 */

const SHEET_NAME = 'ข้อสอบ';
const HEADERS = ['subject', 'question', 'A', 'B', 'C', 'D', 'answer', 'explanation', 'topic', 'difficulty', 'lawRef', 'active'];

// ชื่อวิชา (ต้องตรงกับในเว็บ) — คอลัมน์ subject ใส่ชื่อไทยหรือคีย์ภาษาอังกฤษก็ได้
const SUBJECTS = [
  ['military_service_law', 'กฎหมาย & พ.ร.บ. ทหาร 2497'],
  ['sassadee_duties', 'การปฏิบัติหน้าที่สายงานสัสดี'],
  ['religion_morality', 'การศาสนาและศีลธรรม'],
  ['general_knowledge_current_affairs', 'ความรู้รอบตัว & สังคม'],
  ['thai', 'ภาษาไทย'],
  ['english', 'ภาษาอังกฤษ'],
  ['saraban', 'ระเบียบงานสารบรรณ'],
  ['general_aptitude', 'ความรู้ความสามารถทั่วไป'],
];

const CACHE_KEY = 'exam_payload_v1';
const CACHE_TTL_SEC = 120;
const CACHE_CHUNK = 30000; // อักขระต่อก้อน (ไทย 1 ตัว = 3 ไบต์ จำกัดก้อนละ 100KB)

/* ---------------------------------------------------------------- Web API */

function doGet() {
  let json = cacheGet_();
  if (!json) {
    json = JSON.stringify(buildPayload_());
    cachePut_(json);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function buildPayload_() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  const questions = [];
  if (sheet && sheet.getLastRow() > 1) {
    const values = sheet.getDataRange().getValues();
    const head = values[0].map(h => String(h).trim());
    for (let r = 1; r < values.length; r++) {
      const row = {};
      head.forEach((h, c) => { if (h) row[h] = String(values[r][c] == null ? '' : values[r][c]).trim(); });
      if (!row.question) continue;                 // แถวว่าง
      if (isOff_(row.active)) continue;            // ปิดการใช้งานข้อนี้ (active = FALSE)
      delete row.active;
      questions.push(row);
    }
  }
  return { updated: new Date().toISOString(), count: questions.length, questions: questions };
}

function isOff_(v) {
  return ['false', '0', 'no', 'n', 'ไม่', 'ปิด'].indexOf(String(v).trim().toLowerCase()) !== -1;
}

/* ------------------------------------------------------------------ Cache */

function cacheGet_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(CACHE_KEY + '_n'));
    if (!n) return null;
    const keys = [];
    for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    const got = c.getAll(keys);
    let out = '';
    for (let i = 0; i < n; i++) {
      const part = got[CACHE_KEY + '_' + i];
      if (part == null) return null;
      out += part;
    }
    return out;
  } catch (e) { return null; }
}

function cachePut_(json) {
  try {
    const c = CacheService.getScriptCache();
    const n = Math.ceil(json.length / CACHE_CHUNK);
    const obj = {};
    for (let i = 0; i < n; i++) obj[CACHE_KEY + '_' + i] = json.substr(i * CACHE_CHUNK, CACHE_CHUNK);
    obj[CACHE_KEY + '_n'] = String(n);
    c.putAll(obj, CACHE_TTL_SEC);
  } catch (e) { /* ข้อสอบเยอะมากจนแคชไม่ได้ ก็ใช้แบบไม่แคช */ }
}

function clearCache_() {
  try {
    const c = CacheService.getScriptCache();
    const n = Number(c.get(CACHE_KEY + '_n')) || 0;
    const keys = [CACHE_KEY + '_n'];
    for (let i = 0; i < n; i++) keys.push(CACHE_KEY + '_' + i);
    c.removeAll(keys);
  } catch (e) {}
}

function onEdit() { clearCache_(); }

/* ------------------------------------------------------------------- Menu */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('ข้อสอบ')
    .addItem('1) สร้างชีตข้อสอบ (ครั้งแรก)', 'setupSheet')
    .addItem('2) ตรวจสอบความถูกต้อง', 'validateSheet')
    .addItem('3) ล้างแคช (อัปเดตเว็บทันที)', 'clearCacheMenu')
    .addToUi();
}

function clearCacheMenu() {
  clearCache_();
  SpreadsheetApp.getActive().toast('ล้างแคชแล้ว เว็บจะโหลดข้อสอบล่าสุดในครั้งถัดไป', 'ข้อสอบ', 5);
}

function setupSheet() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (sh && sh.getLastRow() > 1) {
    SpreadsheetApp.getUi().alert('มีชีต "' + SHEET_NAME + '" และมีข้อมูลอยู่แล้ว ไม่สร้างทับ');
    return;
  }
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  sh.clear();
  sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS])
    .setFontWeight('bold').setBackground('#141E46').setFontColor('#ffffff');
  sh.setFrozenRows(1);

  // แถวตัวอย่าง 3 แถว: active = FALSE จึงยังไม่แสดงบนเว็บ แก้เป็น TRUE หรือเขียนทับด้วยข้อจริง
  const labels = SUBJECTS.map(s => s[1]);
  sh.getRange(2, 1, 3, HEADERS.length).setValues([
    [labels[0], 'ตัวอย่าง: ข้อใดถูกต้อง?', 'ตัวเลือกที่ 1', 'ตัวเลือกที่ 2', 'ตัวเลือกที่ 3', 'ตัวเลือกที่ 4', 'B', 'คำอธิบายเฉลย', 'การตรวจเลือก', 'medium', 'พ.ร.บ. รับราชการทหาร มาตรา 18', false],
    [labels[1], 'ตัวอย่างข้อที่ 2', 'ก', 'ข', 'ค', 'ง', 'A', 'คำอธิบายเฉลย', 'บัญชีเรียก', 'easy', '', false],
    [labels[4], 'ตัวอย่างข้อที่ 3', 'ก', 'ข', 'ค', 'ง', 'C', 'คำอธิบายเฉลย', 'หลักภาษา', 'hard', '', false],
  ]);

  const rows = sh.getMaxRows() - 1;   // ทุกแถวที่มีในชีต
  sh.getRange(2, 1, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(labels, true).setAllowInvalid(false).build());
  sh.getRange(2, 7, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['A', 'B', 'C', 'D'], true).setAllowInvalid(false).build());
  sh.getRange(2, 10, rows, 1).setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['easy', 'medium', 'hard'], true).setAllowInvalid(true).build());
  sh.getRange(2, 12, rows, 1).insertCheckboxes();
  sh.getRange(5, 12, rows - 3, 1).setValue(true);   // แถวใหม่ติ๊ก "เปิดใช้" ไว้ให้เลย

  [200, 380, 160, 160, 160, 160, 70, 320, 140, 90, 200, 70].forEach((w, i) => sh.setColumnWidth(i + 1, w));
  sh.getRange(2, 1, rows, HEADERS.length).setWrap(true).setVerticalAlignment('top');
  sh.getRange(1, 1, 1, HEADERS.length).setNote('คำอธิบายคอลัมน์\nsubject = วิชา (เลือกจากรายการ)\nquestion = โจทย์\nA–D = ตัวเลือก\nanswer = เฉลย A–D\nexplanation = คำอธิบาย\ntopic = หัวข้อ\ndifficulty = easy/medium/hard\nlawRef = อ้างอิงมาตรา (ไม่บังคับ)\nactive = ติ๊กเพื่อแสดงบนเว็บ');
  SpreadsheetApp.getActive().toast('สร้างชีตแล้ว ลองแก้ 3 แถวตัวอย่างหรือเริ่มกรอกข้อสอบได้เลย', 'ข้อสอบ', 8);
}

/** ตรวจแถวที่กรอกไม่ครบ/ผิด แล้วระบายสีแดงอ่อน พร้อมสรุปจำนวนข้อต่อวิชา */
function validateSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh || sh.getLastRow() < 2) { SpreadsheetApp.getUi().alert('ยังไม่มีข้อมูลในชีต "' + SHEET_NAME + '"'); return; }
  const values = sh.getDataRange().getValues();
  const head = values[0].map(h => String(h).trim());
  const idx = {}; head.forEach((h, i) => idx[h] = i);
  const norm = t => String(t).toLowerCase().replace(/[\s_\-]/g, '');
  const subjOK = {};
  SUBJECTS.forEach(s => { subjOK[norm(s[0])] = s[1]; subjOK[norm(s[1])] = s[1]; });
  const ans = ['A', 'B', 'C', 'D', 'ก', 'ข', 'ค', 'ง', '1', '2', '3', '4'];
  const problems = []; const count = {}; const seen = {};
  sh.getRange(2, 1, Math.max(1, values.length - 1), head.length).setBackground(null);

  for (let r = 1; r < values.length; r++) {
    const v = k => String(values[r][idx[k]] == null ? '' : values[r][idx[k]]).trim();
    if (!v('question')) continue;
    const bad = [];
    const subj = subjOK[norm(v('subject'))];
    if (!subj) bad.push('วิชาไม่ถูกต้อง');
    ['A', 'B', 'C', 'D'].forEach(k => { if (!v(k)) bad.push('ไม่มีตัวเลือก ' + k); });
    if (ans.indexOf(v('answer').toUpperCase().replace(/[.)\s]/g, '')) === -1) bad.push('เฉลยไม่ถูกต้อง');
    const key = norm(v('subject')) + '|' + v('question');
    if (seen[key]) bad.push('ซ้ำกับแถว ' + seen[key]); else seen[key] = r + 1;
    if (bad.length) {
      problems.push('แถว ' + (r + 1) + ': ' + bad.join(', '));
      sh.getRange(r + 1, 1, 1, head.length).setBackground('#fde2e2');
    } else if (!(idx.active != null && isOff_(v('active')))) {
      count[subj] = (count[subj] || 0) + 1;
    }
  }
  const total = Object.keys(count).reduce((a, k) => a + count[k], 0);
  let msg = 'ข้อสอบที่พร้อมแสดงบนเว็บ: ' + total + ' ข้อ\n' + Object.keys(count).map(k => '• ' + k + ': ' + count[k]).join('\n');
  msg += problems.length ? '\n\nพบปัญหา ' + problems.length + ' แถว (ระบายสีแดงในชีตแล้ว เว็บจะข้ามแถวเหล่านี้)\n' + problems.slice(0, 15).join('\n') : '\n\nไม่พบแถวที่มีปัญหา';
  SpreadsheetApp.getUi().alert(msg);
}
