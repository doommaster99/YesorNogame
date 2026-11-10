/**
 * เกม "ใช่หรือไม่?" — ฝั่งเซิร์ฟเวอร์ (Google Apps Script)
 * - แสดงเกมเป็น Web App (ไฟล์ index.html, ไม่บังคับ)
 * - เปิด API (JSON) ให้เกมที่โฮสต์ที่อื่น เช่น GitHub Pages เรียกใช้
 * - อ่านคำถามจากชีต "Questions"
 * - บันทึกคะแนนลงชีต "Scores"
 * - ส่งกระดานคะแนน (Leaderboard)
 */
const SHEET_ID = '1FIx4LCj-s_cUxyRcK2h32BJooXXf95VJF2_9UiSOK_0'; // ID ของ Google Sheet (ว่าง = ใช้สเปรดชีตที่สคริปต์ผูกอยู่)
const SHEET_Q = 'Questions';
const SHEET_S = 'Scores';

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * GET ?action=questions            -> คำถามทั้งหมด
 * GET ?action=leaderboard&limit=10 -> กระดานคะแนน
 * GET (ไม่มี action)               -> หน้าเกม (ถ้ามีไฟล์ index.html ในโปรเจกต์) หรือข้อความยืนยันว่า API ทำงาน
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.action) {
    try {
      if (p.action === 'questions') return json_({ ok: true, data: getQuestions() });
      if (p.action === 'leaderboard') return json_({ ok: true, data: getLeaderboard(p.limit) });
      return json_({ error: 'unknown action' });
    } catch (err) {
      return json_({ error: String(err.message || err) });
    }
  }
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('ใช่หรือไม่? ภาพถ่ายทางอากาศและดาวเทียม')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput('Yes/No Geography API is running.');
  }
}

/** POST (text/plain, เนื้อหาเป็น JSON ของผลการเล่น) -> บันทึกคะแนน */
function doPost(e) {
  try {
    return json_({ ok: true, data: saveResult(JSON.parse(e.postData.contents)) });
  } catch (err) {
    return json_({ error: String(err.message || err) });
  }
}

function ss_() {
  return SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

/** รันครั้งเดียวจากตัวแก้ไข เพื่อสร้างชีตและตัวอย่างข้อมูล */
function setup() {
  const ss = ss_();

  let q = ss.getSheetByName(SHEET_Q);
  if (!q) {
    q = ss.insertSheet(SHEET_Q);
    q.appendRow(['ข้อความคำถาม', 'คำตอบ (ใช่/ไม่)', 'คำอธิบาย', 'ระดับ (1-3)', 'ใช้งาน (TRUE/FALSE)']);
    q.appendRow(['ภาพถ่ายทางอากาศถ่ายจากอุปกรณ์ที่อยู่ในชั้นบรรยากาศ เช่น เครื่องบิน บอลลูน หรือโดรน', 'ใช่', 'ภาพถ่ายดาวเทียมถ่ายจากอวกาศ', 1, true]);
    q.appendRow(['ภาพถ่ายดาวเทียมถ่ายได้เฉพาะตอนกลางวันเท่านั้น', 'ไม่', 'ดาวเทียมเรดาร์ถ่ายได้ทั้งกลางคืนและทะลุเมฆ', 2, true]);
    q.appendRow(['GPS เป็นดาวเทียมที่ใช้ถ่ายภาพพื้นผิวโลก', 'ไม่', 'GPS ใช้บอกตำแหน่ง ไม่ได้ถ่ายภาพ', 1, true]);
    q.appendRow(['ค่า NDVI ที่สูงใกล้ 1 มักหมายถึงพืชพรรณหนาแน่นและสมบูรณ์', 'ใช่', 'พืชสุขภาพดีสะท้อนอินฟราเรดใกล้มาก', 3, true]);
    q.setFrozenRows(1);
    q.getRange(1, 1, 1, 5).setFontWeight('bold');
    q.setColumnWidth(1, 520);
    q.setColumnWidth(3, 320);
  }

  let s = ss.getSheetByName(SHEET_S);
  if (!s) {
    s = ss.insertSheet(SHEET_S);
    s.appendRow(['เวลา', 'จำนวนผู้เล่น', 'ลำดับที่นั่ง', 'ชื่อ', 'คะแนน', 'ตอบถูก', 'จำนวนข้อ', 'เวลาต่อข้อ (วินาที)']);
    s.setFrozenRows(1);
    s.getRange(1, 1, 1, 8).setFontWeight('bold');
    s.getRange('D:D').setNumberFormat('@'); // เก็บชื่อเป็นข้อความล้วน ป้องกันการถูกตีความเป็นสูตร
  }
}

/** ส่งคำถามให้หน้าเกม: [[ข้อความ, คำตอบ(true/false), คำอธิบาย, ระดับ]] */
function getQuestions() {
  const sh = ss_().getSheetByName(SHEET_Q);
  if (!sh) throw new Error('ไม่พบชีต ' + SHEET_Q + ' กรุณารันฟังก์ชัน setup() ก่อน');

  const T = ['ใช่', 'true', 't', 'yes', 'y', '1', 'จริง', 'ถูก'];
  const F = ['ไม่', 'ไม่ใช่', 'false', 'f', 'no', 'n', '0', 'เท็จ', 'ผิด'];
  const out = [];

  sh.getDataRange().getValues().slice(1).forEach(function (r) {
    const text = String(r[0]).trim();
    const a = String(r[1]).trim().toLowerCase();
    const enabled = String(r[4]).trim().toLowerCase() !== 'false';
    if (!text || !enabled) return;

    const ans = T.indexOf(a) > -1 ? true : F.indexOf(a) > -1 ? false : null;
    if (ans === null) return;

    const lv = parseInt(r[3], 10);
    out.push([
      text,
      ans,
      String(r[2]).trim() || (ans ? 'คำตอบคือ ใช่' : 'คำตอบคือ ไม่ใช่'),
      [1, 2, 3].indexOf(lv) > -1 ? lv : 0
    ]);
  });
  return out;
}

function clamp_(v, min, max) {
  v = Number(v);
  if (isNaN(v)) v = min;
  return Math.max(min, Math.min(max, Math.round(v)));
}

function clean_(s) {
  return String(s || '').replace(/[\u0000-\u001f]/g, ' ').trim().slice(0, 30);
}

/** บันทึกผลหนึ่งเกม: { n, time, total, players: [{ name, score, correct }] } */
function saveResult(p) {
  if (!p || !Array.isArray(p.players) || !p.players.length) throw new Error('ข้อมูลไม่ถูกต้อง');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sh = ss_().getSheetByName(SHEET_S);
    if (!sh) throw new Error('ไม่พบชีต ' + SHEET_S + ' กรุณารันฟังก์ชัน setup() ก่อน');

    const now = new Date();
    const n = clamp_(p.n, 1, 4);
    const total = clamp_(p.total, 1, 100);
    const time = clamp_(p.time, 3, 120);
    const rows = p.players.slice(0, 4).map(function (x, i) {
      return [now, n, i + 1, clean_(x.name) || 'ผู้เล่น ' + (i + 1),
              clamp_(x.score, 0, 100000), clamp_(x.correct, 0, total), total, time];
    });

    const start = sh.getLastRow() + 1;
    sh.getRange(start, 4, rows.length, 1).setNumberFormat('@');
    sh.getRange(start, 1, rows.length, 8).setValues(rows);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/** กระดานคะแนน: ส่งกลับเป็น JSON ล้วน (วันที่เป็นข้อความ) */
function getLeaderboard(limit) {
  const sh = ss_().getSheetByName(SHEET_S);
  if (!sh) throw new Error('ไม่พบชีต ' + SHEET_S);

  const tz = Session.getScriptTimeZone();
  return sh.getDataRange().getValues().slice(1)
    .filter(function (r) { return r[0]; })
    .map(function (r) {
      return {
        date: Utilities.formatDate(new Date(r[0]), tz, 'dd/MM/yyyy HH:mm'),
        name: String(r[3]),
        score: Number(r[4]),
        correct: Number(r[5]),
        total: Number(r[6])
      };
    })
    .sort(function (a, b) { return b.score - a.score; })
    .slice(0, Math.min(Number(limit) || 10, 50));
}
