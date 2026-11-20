/**
 * Insect Morphology 3D — บันทึกคะแนนแบบทดสอบลง Google Sheet
 *
 * วิธีติดตั้ง (ทำครั้งเดียว)
 * 1. เปิด Google Sheet ที่จะเก็บคะแนน → ส่วนขยาย (Extensions) → Apps Script
 * 2. ลบโค้ดเดิมใน Code.gs แล้ววางโค้ดทั้งไฟล์นี้ → บันทึก
 * 3. เลือกฟังก์ชัน setup แล้วกด Run (เรียกใช้) ครั้งแรกจะขออนุญาตเข้าถึง Sheet ให้กดอนุญาต
 *    setup จะสร้างแผ่นงาน Attempts, Answers, Roster, Summary, Item stats พร้อมหัวตารางและสูตร
 * 4. ทำให้ใช้งานได้ (Deploy) → การทำให้ใช้งานได้รายการใหม่ (New deployment) → ประเภท: เว็บแอป (Web app)
 *    - Execute as: Me (บัญชีของอาจารย์)
 *    - Who has access: Anyone (ทุกคน)  ← นักศึกษาไม่ต้องล็อกอิน
 *    กด Deploy แล้วคัดลอก URL ที่ลงท้ายด้วย /exec ส่งให้ผู้ดูแลเว็บ ไปใส่ใน SHEET_URL ของ index.html
 * 5. ใส่รหัสนักศึกษาที่ลงทะเบียนในแผ่นงาน Roster คอลัมน์ A (ไม่ใส่ก็ได้ แต่จะตรวจรหัสพิมพ์ผิดไม่ได้)
 *
 * แก้โค้ดแล้วต้อง Deploy → Manage deployments → แก้ไข (ดินสอ) → Version: New version → Deploy
 * จึงจะใช้โค้ดใหม่ (URL เดิม)
 */

const TZ = 'Asia/Bangkok';
const ATT = 'Attempts', ANS = 'Answers', ROS = 'Roster', SUM = 'Summary', ITEM = 'Item stats';
const ATT_HEAD = ['เวลาเริ่มทำ', 'attemptId', 'รหัสนักศึกษา', 'ระดับ', 'หมวด', 'รูปแบบ', 'จับเวลา (วินาที)', 'สถานะ',
  'เวลาส่ง', 'จำนวนข้อ', 'ทำแล้ว', 'ถูกครั้งแรก', 'ร้อยละ', 'ใช้เวลา (วินาที)', 'ข้อที่ควรทบทวน', 'อยู่ในรายชื่อ', 'เวอร์ชัน'];
const ANS_HEAD = ['เวลา', 'attemptId', 'รหัสนักศึกษา', 'ระดับ', 'ข้อที่', 'ประเภท', 'key', 'ชื่อ', 'ถูกครั้งแรก (1/0)', 'จำนวนครั้งที่ตอบ'];

function setup() {
  const ss = SpreadsheetApp.getActive();
  ss.setSpreadsheetTimeZone(TZ);
  const sheet = (name, head) => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    if (head && sh.getLastRow() === 0) {
      sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
      sh.setFrozenRows(1);
    }
    return sh;
  };
  const att = sheet(ATT, ATT_HEAD);
  att.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  att.getRange('I:I').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  att.getRange('C:C').setNumberFormat('@');
  const ans = sheet(ANS, ANS_HEAD);
  ans.getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  ans.getRange('C:C').setNumberFormat('@');
  const ros = sheet(ROS, ['รหัสนักศึกษา', 'ชื่อ-สกุล (ไม่บังคับ)', 'กลุ่ม (ไม่บังคับ)']);
  ros.getRange('A:A').setNumberFormat('@');

  const sum = sheet(SUM);
  sum.clear();
  sum.getRange('A1').setValue('สรุปรายคน: เฉพาะรอบที่ทำครบหรือหมดเวลา (สถานะ = ส่งแล้ว) · เกณฑ์ผ่านแก้ได้ที่ช่อง I1');
  sum.getRange('H1').setValue('เกณฑ์ผ่าน (%)');
  sum.getRange('I1').setValue(80);
  sum.getRange('A3').setFormula(
    "=IFERROR(QUERY(Attempts!A:Q,\"select C, D, count(B), max(M), avg(M), min(A), max(I) " +
    "where H = 'ส่งแล้ว' group by C, D order by C, D " +
    "label C 'รหัสนักศึกษา', D 'ระดับ', count(B) 'ทำ (ครั้ง)', max(M) 'คะแนนสูงสุด %', avg(M) 'เฉลี่ย %', " +
    "min(A) 'ทำครั้งแรก', max(I) 'ส่งล่าสุด'\",1),\"ยังไม่มีข้อมูล\")");
  sum.getRange('H3').setValue('ผ่านเกณฑ์');
  sum.getRange('H4').setFormula('=ARRAYFORMULA(IF(A4:A="",,IF(D4:D>=$I$1,"ผ่าน","ยังไม่ผ่าน")))');
  sum.getRange('F:G').setNumberFormat('yyyy-mm-dd hh:mm');
  sum.getRange('E:E').setNumberFormat('0.0');
  sum.setFrozenRows(3);

  const item = sheet(ITEM);
  item.clear();
  item.getRange('A1').setValue('ร้อยละที่ตอบถูกในครั้งแรกของแต่ละข้อ (ทุกคน ทุกรอบ) เรียงจากข้อที่ยากที่สุด');
  item.getRange('A3').setFormula(
    "=IFERROR(QUERY(Answers!A:J,\"select D, H, count(I), avg(I) where I is not null group by D, H order by avg(I) asc " +
    "label D 'ระดับ', H 'ชื่อ', count(I) 'ถูกถาม (ครั้ง)', avg(I) 'สัดส่วนถูกครั้งแรก'\",1),\"ยังไม่มีข้อมูล\")");
  item.getRange('D:D').setNumberFormat('0%');

  const s1 = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (s1 && s1.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(s1);
}

function doGet() {
  return json({ok: true, app: 'insect-morphology-3d'});
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const p = JSON.parse(e.postData.contents);
    const id = String(p.studentId || '').replace(/\D/g, '');
    if (!/^\d{10}$/.test(id)) return json({ok: false, error: 'studentId must be 10 digits'});
    if (!p.attemptId || String(p.attemptId).length > 64) return json({ok: false, error: 'bad attemptId'});
    const ss = SpreadsheetApp.getActive();
    const att = ss.getSheetByName(ATT), ans = ss.getSheetByName(ANS);
    if (!att || !ans) return json({ok: false, error: 'run setup() first'});
    const now = new Date();
    const known = inRoster(ss, id);
    const lvl = Number(p.level) || '', total = Number(p.total) || 0;
    const row = findRow(att, String(p.attemptId));

    if (p.event === 'start') {
      if (!row) att.appendRow([now, String(p.attemptId), id, lvl, String(p.group || ''), String(p.mode || ''),
        Number(p.timeLimit) || 0, 'กำลังทำ', '', total, 0, 0, '', '', '', known === null ? '' : (known ? 'ใช่' : 'ไม่พบ'), String(p.ver || '')]);
      return json({ok: true, known: known, time: now.toISOString()});
    }

    if (p.event === 'finish') {
      const answers = Array.isArray(p.answers) ? p.answers.slice(0, 200) : [];
      const right = answers.filter(a => a.ok).length;
      const pct = total ? Math.round(100 * right / total) : 0;
      const wrong = answers.filter(a => !a.ok).map(a => String(a.name)).join(', ');
      const values = [String(p.group || ''), String(p.mode || ''), Number(p.timeLimit) || 0,
        'ส่งแล้ว', now, total, Number(p.answered) || 0, right, pct, Number(p.durationSec) || '', wrong.slice(0, 5000),
        known === null ? '' : (known ? 'ใช่' : 'ไม่พบ'), String(p.ver || '')];
      if (row) {
        att.getRange(row, 5, 1, values.length).setValues([values]);
      } else {
        att.appendRow([now, String(p.attemptId), id, lvl].concat(values));
      }
      if (answers.length) {
        ans.getRange(ans.getLastRow() + 1, 1, answers.length, ANS_HEAD.length).setValues(answers.map(a => [
          now, String(p.attemptId), id, lvl, Number(a.n) || '', String(a.kind || ''), String(a.key || ''),
          String(a.name || ''), a.ok ? 1 : 0, Number(a.tries) || 0]));
      }
      return json({ok: true, known: known, time: now.toISOString(), right: right, pct: pct});
    }
    return json({ok: false, error: 'unknown event'});
  } catch (err) {
    return json({ok: false, error: String(err)});
  } finally {
    lock.releaseLock();
  }
}

/* true / false when the Roster has IDs, null when the Roster is empty */
function inRoster(ss, id) {
  const sh = ss.getSheetByName(ROS);
  if (!sh || sh.getLastRow() < 2) return null;
  const ids = sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues().map(r => String(r[0]).replace(/\D/g, ''));
  return ids.indexOf(id) !== -1;
}

function findRow(sh, attemptId) {
  const n = sh.getLastRow();
  if (n < 2) return 0;
  const hit = sh.getRange(2, 2, n - 1, 1).createTextFinder(attemptId).matchEntireCell(true).findNext();
  return hit ? hit.getRow() : 0;
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
