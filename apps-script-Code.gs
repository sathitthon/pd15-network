/**
 * บ้านเน็ต PD15 — รับข้อมูลจาก router (MikroTik script dash-push) ทุก 5 นาที แล้วเปิดให้หน้า 3D อ่าน
 * - router ส่ง GET ?t=รหัสลับ&dn=..&dm=.. → บันทึก (รหัสผิด = ทิ้ง)
 * - ไม่มี t → ส่ง JSON (ค่าล่าสุด + กราฟ 24 ชม. ราย 30 นาที) สำหรับหน้าเว็บ อ่านอย่างเดียว
 * - ข้อมูลที่เก็บเป็นยอดนับเท่านั้น ไม่มี IP / MAC / ข้อมูลรายห้อง
 * ติดตั้ง: รัน setup() ครั้งเดียว (สร้างชีต เน็ต_PD15_live + แท็บ + รหัสลับ + trigger ล้างข้อมูลเก่า) → Deploy เป็น Web app (Execute as: Me, Who has access: Anyone)
 */
function ss_() { return SpreadsheetApp.openById(PropertiesService.getScriptProperties().getProperty('SHEET_ID')); }
const TAB_NOW = 'ล่าสุด';
const TAB_HIS = 'ประวัติ';
const KEYS = ['dn', 'dm', 'up', 'cpu', 'ut', 'ppp', 'png', 'bad', 'l2', 'l3', 'l4', 'l5', 'l6', 'l7', 'l8', 'h', 'of'];
const HEAD = ['เวลา'].concat(KEYS);
const KEEP_DAYS = 30;

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.t !== undefined) return save_(p);
  return json_(read_());
}

function save_(p) {
  const tok = PropertiesService.getScriptProperties().getProperty('TOKEN');
  if (!tok || p.t !== tok) return ContentService.createTextOutput('no');
  const ss = ss_();
  const row = [new Date()].concat(KEYS.map(k => k === 'ut' ? "'" + clean_(p[k]) : clean_(p[k])));
  ss.getSheetByName(TAB_NOW).getRange(2, 1, 1, row.length).setValues([row]);
  ss.getSheetByName(TAB_HIS).appendRow(row);
  CacheService.getScriptCache().remove('out');
  return ContentService.createTextOutput('ok');
}

function clean_(v) {
  if (v === undefined || v === null) return '';
  const s = String(v).slice(0, 80);
  return /^-?\d+(\.\d+)?$/.test(s) ? Number(s) : s.replace(/[^\w.\-:]/g, '');
}

function read_() {
  const cache = CacheService.getScriptCache();
  const hit = cache.get('out');
  if (hit) return JSON.parse(hit);
  const ss = ss_();
  const nowRow = ss.getSheetByName(TAB_NOW).getRange(2, 1, 1, HEAD.length).getValues()[0];
  const now = {};
  if (nowRow[0]) {
    now.ts = new Date(nowRow[0]).getTime();
    KEYS.forEach((k, i) => { now[k] = nowRow[i + 1]; });
  }
  const sh = ss.getSheetByName(TAB_HIS);
  const last = sh.getLastRow();
  const hist = [];
  let peak = 0;
  if (last > 1) {
    const n = Math.min(last - 1, 300);
    const rows = sh.getRange(last - n + 1, 1, n, HEAD.length).getValues();
    const since = Date.now() - 24 * 3600 * 1000;
    const buckets = {};
    rows.forEach(r => {
      const t = new Date(r[0]).getTime();
      if (!t || t < since) return;
      const b = Math.floor(t / 1800000) * 1800000;
      const x = buckets[b] || (buckets[b] = { s: 0, n: 0, m: 0, a: 0 });
      x.s += Number(r[1]) || 0; x.n++; x.m = Math.max(x.m, Number(r[2]) || 0);
      x.a = [9, 10, 11, 12, 13, 14, 15].reduce((s, i) => s + (Number(r[i]) || 0), 0);
      peak = Math.max(peak, Number(r[2]) || 0);
    });
    Object.keys(buckets).sort().forEach(b => {
      const x = buckets[b];
      const label = Utilities.formatDate(new Date(Number(b)), 'Asia/Bangkok', 'HH:mm');
      hist.push([label, Math.round(x.s / x.n), x.m, x.a]);
    });
  }
  const out = { now: now, hist: hist, peak: peak };
  cache.put('out', JSON.stringify(out), 30);
  return out;
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

/** ลบแถวประวัติที่เก่ากว่า KEEP_DAYS วัน — trigger วันละครั้ง */
function cleanup() {
  const sh = ss_().getSheetByName(TAB_HIS);
  const last = sh.getLastRow();
  if (last < 2) return;
  const cut = Date.now() - KEEP_DAYS * 86400000;
  const times = sh.getRange(2, 1, last - 1, 1).getValues();
  let n = 0;
  while (n < times.length && new Date(times[n][0]).getTime() < cut) n++;
  if (n > 0) sh.deleteRows(2, n);
}

/** รันครั้งเดียว: สร้างแท็บ หัวตาราง รหัสลับ และ trigger ล้างข้อมูล */
function setup() {
  const props = PropertiesService.getScriptProperties();
  if (!props.getProperty('SHEET_ID')) props.setProperty('SHEET_ID', SpreadsheetApp.create('เน็ต_PD15_live').getId());
  const ss = ss_();
  [TAB_NOW, TAB_HIS].forEach(name => {
    const sh = ss.getSheetByName(name) || ss.insertSheet(name);
    sh.getRange(1, 1, 1, HEAD.length).setValues([HEAD]).setFontWeight('bold');
    sh.setFrozenRows(1);
  });
  if (!props.getProperty('TOKEN')) props.setProperty('TOKEN', Utilities.getUuid().replace(/-/g, ''));
  ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'cleanup').forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('cleanup').timeBased().everyDays(1).atHour(4).create();
  const def = ss.getSheetByName('Sheet1') || ss.getSheetByName('ชีต1');
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
  Logger.log('SHEET = ' + ss.getUrl());
  Logger.log('TOKEN = ' + props.getProperty('TOKEN'));
}
