/**
 * ระบบลงทะเบียนผ่อนผันการเข้ารับราชการทหารกองประจำการ ปีการศึกษา 2569
 * Backend: Google Apps Script + Google Sheets
 *
 * วิธีใช้ (ย่อ): ดู README.md
 *  1) สร้าง Google Sheet ใหม่ > ส่วนขยาย > Apps Script
 *  2) วางไฟล์นี้เป็น Code.gs และสร้างไฟล์ HTML ชื่อ index แล้ววางเนื้อหา index.html
 *  3) รันฟังก์ชัน setup() หนึ่งครั้ง (ตั้งรหัสผ่านผู้ดูแล + สร้างชีต)
 *  4) ทำให้ใช้งานได้ > การทำให้ใช้งานได้รายการใหม่ > เว็บแอป
 *     - ดำเนินการในฐานะ: ฉัน   - ผู้ที่มีสิทธิ์เข้าถึง: ทุกคน
 */

const SHEET_NAME = 'ลงทะเบียน';
const DEADLINE = new Date('2026-12-18T23:59:59+07:00');
const REQUIRED_BIRTH_YEAR = 2549;
const STATUSES = ['รอส่งเอกสาร', 'ได้รับเอกสารแล้ว', 'เอกสารไม่ครบ', 'ส่งเรื่องผ่อนผันแล้ว'];

// ลำดับคอลัมน์ในชีต (key, หัวคอลัมน์)
const COLUMNS = [
  ['ref', 'เลขอ้างอิง'], ['createdAt', 'วันที่ลงทะเบียน'], ['status', 'สถานะ'], ['note', 'หมายเหตุ'],
  ['late', 'หลังกำหนด'], ['deferredBefore', 'เคยผ่อนผัน'],
  ['prefix', 'คำนำหน้า'], ['firstName', 'ชื่อ'], ['lastName', 'นามสกุล'], ['nationalId', 'เลขประจำตัวประชาชน'],
  ['birthDate', 'วันเกิด (ว/ด/ปพ.ศ.)'], ['phone', 'โทรศัพท์'], ['contact', 'อีเมล/LINE'],
  ['studentId', 'รหัสนักศึกษา'], ['level', 'ระดับชั้น'], ['major', 'สาขาวิชา'], ['room', 'ห้อง'], ['advisor', 'ครูที่ปรึกษา'],
  ['sd9No', 'เลข สด.9'], ['district', 'อำเภอ (ภูมิลำเนาทหาร)'], ['province', 'จังหวัด'],
  ['changedName', 'เปลี่ยนชื่อ-สกุล'], ['changedDomicile', 'ย้ายภูมิลำเนาทหาร'], ['docs', 'เอกสารที่เตรียม (หมายเลข)']
];

/* ---------- Setup (รันครั้งเดียว) ---------- */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('ADMIN_PASSWORD')) props.setProperty('ADMIN_PASSWORD', 'เปลี่ยนรหัสนี้2569');
  getSheet_();
  Logger.log('พร้อมใช้งาน รหัสผ่านผู้ดูแลปัจจุบัน: ' + props.getProperty('ADMIN_PASSWORD') +
             ' (เปลี่ยนได้ที่ การตั้งค่าโปรเจ็กต์ > คุณสมบัติสคริปต์ > ADMIN_PASSWORD)');
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, COLUMNS.length).setValues([COLUMNS.map(c => c[1])])
      .setFontWeight('bold').setBackground('#0f2a5c').setFontColor('#ffffff');
    sh.setFrozenRows(1);
    // เก็บเลขบัตร/โทร/รหัสเป็นข้อความ ไม่ให้ศูนย์หน้าหาย
    ['nationalId', 'phone', 'studentId', 'sd9No'].forEach(k => sh.getRange(2, colIndex_(k), sh.getMaxRows() - 1).setNumberFormat('@'));
  }
  return sh;
}
function colIndex_(key) { return COLUMNS.findIndex(c => c[0] === key) + 1; }

/* ---------- Web entry points ---------- */
// เปิดเว็บแอปโดยตรง / ฝังใน Google Sites
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ลงทะเบียนผ่อนผันทหาร 2569')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// เรียกจากหน้าเว็บที่โฮสต์ภายนอก (เช่น GitHub Pages)
function doPost(e) {
  let out;
  try {
    const body = JSON.parse(e.postData.contents || '{}');
    out = api(body.action, body.payload || {});
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

// เรียกผ่าน google.script.run (เมื่อเปิดผ่าน Apps Script / Google Sites)
function api(action, p) {
  try {
    switch (action) {
      case 'register': return register_(p);
      case 'status': return status_(p);
      case 'adminList': checkAdmin_(p.password); return { ok: true, records: readAll_() };
      case 'adminUpdate': checkAdmin_(p.password); return update_(p);
      default: return { ok: false, error: 'unknown action' };
    }
  } catch (err) {
    return { ok: false, error: String(err.message || err) };
  }
}

/* ---------- Actions ---------- */
function register_(p) {
  const nid = String(p.nationalId || '').replace(/\D/g, '');
  if (!validThaiId_(nid)) throw new Error('เลขประจำตัวประชาชนไม่ถูกต้อง');
  if (p.deferredBefore === 'เคย') throw new Error('ผู้ที่เคยผ่อนผันแล้วไม่ต้องลงทะเบียนซ้ำ');
  const year = Number(String(p.birthDate || '').split('/')[2]);
  if (year !== REQUIRED_BIRTH_YEAR) throw new Error('ปีนี้รับเฉพาะผู้ที่เกิดปี พ.ศ. ' + REQUIRED_BIRTH_YEAR);
  ['prefix', 'firstName', 'lastName', 'phone', 'studentId', 'level', 'major', 'district', 'province'].forEach(k => {
    if (!String(p[k] || '').trim()) throw new Error('ข้อมูลไม่ครบ: ' + k);
  });

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sh = getSheet_();
    const last = sh.getLastRow();
    if (last > 1) {
      const ids = sh.getRange(2, colIndex_('nationalId'), last - 1).getDisplayValues().flat();
      if (ids.indexOf(nid) !== -1) throw new Error('เลขประจำตัวประชาชนนี้ลงทะเบียนไว้แล้ว — ตรวจสอบได้ที่แท็บ ตรวจสอบสถานะ');
    }
    const ref = 'MD69-' + String(last).padStart(4, '0');
    const rec = Object.assign({}, p, {
      ref: ref, nationalId: nid, createdAt: new Date(), status: STATUSES[0], note: '',
      late: new Date() > DEADLINE ? 'ใช่ (ค่าปรับ 300 บาท)' : '',
      changedName: p.changedName ? 'ใช่' : '', changedDomicile: p.changedDomicile ? 'ใช่' : ''
    });
    const row = COLUMNS.map(c => {
      const v = rec[c[0]];
      return v === undefined || v === null ? '' : (v instanceof Date ? v : String(v).slice(0, 200));
    });
    sh.getRange(last + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
    sh.getRange(last + 1, colIndex_('createdAt')).setNumberFormat('dd/MM/yyyy HH:mm');
    return { ok: true, ref: ref };
  } finally {
    lock.releaseLock();
  }
}

function status_(p) {
  const nid = String(p.nationalId || '').replace(/\D/g, '');
  const sid = String(p.studentId || '').trim();
  const r = readAll_().find(x => x.nationalId === nid && String(x.studentId).trim() === sid);
  if (!r) throw new Error('ไม่พบข้อมูล กรุณาตรวจสอบเลขประจำตัวประชาชนและรหัสนักศึกษา');
  return { ok: true, record: {
    ref: r.ref, name: r.prefix + r.firstName + ' ' + r.lastName, status: r.status,
    note: r.note, createdAt: r.createdAt, late: !!r.late
  }};
}

function update_(p) {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) throw new Error('ไม่พบรายการ');
  const refs = sh.getRange(2, 1, last - 1).getDisplayValues().flat();
  const i = refs.indexOf(p.ref);
  if (i === -1) throw new Error('ไม่พบรายการ');
  if (p.status !== undefined) {
    if (STATUSES.indexOf(p.status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    sh.getRange(i + 2, colIndex_('status')).setValue(p.status);
  }
  if (p.note !== undefined) sh.getRange(i + 2, colIndex_('note')).setValue(String(p.note).slice(0, 300));
  return { ok: true };
}

/* ---------- Helpers ---------- */
function readAll_() {
  const sh = getSheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  const vals = sh.getRange(2, 1, last - 1, COLUMNS.length).getValues();
  const disp = sh.getRange(2, 1, last - 1, COLUMNS.length).getDisplayValues();
  return vals.map((row, r) => {
    const o = {};
    COLUMNS.forEach((c, j) => {
      const v = row[j];
      o[c[0]] = v instanceof Date ? v.toISOString() : disp[r][j];
    });
    return o;
  });
}

function checkAdmin_(pw) {
  const real = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!real || pw !== real) {
    Utilities.sleep(800); // ชะลอการเดารหัสผ่าน
    throw new Error('รหัสผ่านไม่ถูกต้อง');
  }
}

function validThaiId_(id) {
  if (!/^\d{13}$/.test(id)) return false;
  let sum = 0;
  for (let i = 0; i < 12; i++) sum += Number(id[i]) * (13 - i);
  return (11 - (sum % 11)) % 10 === Number(id[12]);
}
