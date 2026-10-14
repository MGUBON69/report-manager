const CONFIG = {
  BOOKING_SHEET: 'BK_M2_BOOKINGS',
  SOURCE_SHEET: 'DEALER BK INPUT V2',
  MASTER_SHEET: 'BK_M2_MASTER',
  HEADERS: ['BookingID','BookingDate','CustomerName','Model','SubModel','Color','Finance','BookingStatus','CarStatus','Note','SalesConsultant','CreatedAt','UpdatedAt'],
  MODELS: ['MG3 HYBRID+','MG ZS','MG5 PRO','MG HS PHEV','MG URBAN','MG4 MCE','MG4 X Power','MG S5 EV','MG Maxus7','MG Maxus9','MG IM6','MG IM5','MG CYBERSTER'],
  SUBMODELS: ['D','D+','X','X LR','X+','V+','STD','MAX','ULTRA','Premium','Performance','Long Range (LR)'],
  COLORS: ['WHITE','GREY','BLACK','GREEN','BEIGE','TITANIUM','ORANGE','RED','BLUE','SILVER'],
  FINANCE: ['เงินสด','KL กสิกร','AY กรุงศรี','TTB ธนชาต','TISCO','ธนาคารอื่น','รอระบุ'],
  BOOKING_STATUS: ['จองใหม่','เตรียมเอกสาร รอเซ็นไฟแนนซ์ (ดาวน์15%ขึ้นไป)','ยื่นไฟแนนซ์แล้ว','ไฟแนนซ์อนุมัติแล้ว/เงินสด รอส่งมอบ (เดือนนี้)','ไฟแนนซ์อนุมัติแล้ว/เงินสด รอส่งมอบ (เดือนอื่น)','ส่งมอบแล้ว','ยกเลิก (ระบุเหตุผลในหมายเหตุ)'],
  CAR_STATUS: ['มีรถพร้อม','ไม่มีรถ','รอจัดสรร','รอเข้า','อื่นๆ']
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('BK M2 | MG Booking Control')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function setupBkM2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('กรุณาเปิด Apps Script จาก Google Sheet ที่ต้องการใช้เป็นฐานข้อมูล');
  let sh = ss.getSheetByName(CONFIG.BOOKING_SHEET);
  if (!sh) sh = ss.insertSheet(CONFIG.BOOKING_SHEET);
  if (sh.getLastRow() === 0) {
    sh.getRange(1,1,1,CONFIG.HEADERS.length).setValues([CONFIG.HEADERS]);
    sh.setFrozenRows(1);
    sh.getRange(1,1,1,CONFIG.HEADERS.length).setFontWeight('bold').setBackground('#12233f').setFontColor('#ffffff');
  } else {
    const current = sh.getRange(1,1,1,Math.max(sh.getLastColumn(),CONFIG.HEADERS.length)).getValues()[0];
    if (String(current[0]) !== 'BookingID') throw new Error('ชีต BK_M2_BOOKINGS มีข้อมูล/หัวตารางเดิม กรุณาตรวจสอบก่อน ไม่เขียนทับข้อมูล');
  }
  let master = ss.getSheetByName(CONFIG.MASTER_SHEET);
  if (!master) master = ss.insertSheet(CONFIG.MASTER_SHEET);
  if (master.getLastRow() === 0) {
    const rows = [['Type','Value']];
    [['Model',CONFIG.MODELS],['SubModel',CONFIG.SUBMODELS],['Color',CONFIG.COLORS],['Finance',CONFIG.FINANCE],['BookingStatus',CONFIG.BOOKING_STATUS],['CarStatus',CONFIG.CAR_STATUS]]
      .forEach(([type, values]) => values.forEach(v => rows.push([type,v])));
    master.getRange(1,1,rows.length,2).setValues(rows);
    master.setFrozenRows(1);
    master.getRange(1,1,1,2).setFontWeight('bold').setBackground('#12233f').setFontColor('#ffffff');
  }
  return {ok:true, message:'เตรียมชีตฐานข้อมูลเรียบร้อย', bookingSheet:CONFIG.BOOKING_SHEET, masterSheet:CONFIG.MASTER_SHEET};
}

function getInitialData() {
  setupBkM2();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(CONFIG.BOOKING_SHEET);
  const values = sh.getDataRange().getValues();
  const bookings = values.slice(1).filter(r => r[0]).map(rowToObject_);
  const masterSh = ss.getSheetByName(CONFIG.MASTER_SHEET);
  const master = {};
  if (masterSh.getLastRow() > 1) {
    masterSh.getRange(2,1,masterSh.getLastRow()-1,2).getDisplayValues().forEach(r => {
      if (!master[r[0]]) master[r[0]] = [];
      master[r[0]].push(r[1]);
    });
  }
  Object.keys(master).forEach(k => master[k] = [...new Set(master[k])]);
  return {bookings, master, serverTime: new Date().toISOString()};
}

function saveBooking(payload) {
  validatePayload_(payload);
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    setupBkM2();
    const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.BOOKING_SHEET);
    const values = sh.getDataRange().getValues();
    const id = String(payload.BookingID || '').trim();
    const now = new Date();
    let rowIndex = -1;
    if (id) {
      for (let i=1;i<values.length;i++) if (String(values[i][0]) === id) { rowIndex = i+1; break; }
      if (rowIndex < 0) throw new Error('ไม่พบ Booking ID ที่ต้องการแก้ไข อาจมีผู้อื่นลบข้อมูลไปแล้ว');
    }
    // Prevent accidental duplicate customer+date+model bookings, while allowing explicit edits.
    const dateKey = String(payload.BookingDate);
    for (let i=1;i<values.length;i++) {
      if (rowIndex === i+1) continue;
      const r = values[i];
      if (String(r[2]).trim().toLowerCase() === String(payload.CustomerName).trim().toLowerCase()
        && dateCellKey_(r[1]) === dateKey
        && String(r[3]).trim() === String(payload.Model).trim()
        && String(r[7]) !== 'ยกเลิก (ระบุเหตุผลในหมายเหตุ)') {
        throw new Error('พบรายการชื่อลูกค้า วันที่ และรุ่นรถซ้ำกัน กรุณาตรวจสอบรายการเดิมก่อนบันทึก');
      }
    }
    const record = [
      id || Utilities.getUuid(),
      new Date(payload.BookingDate + 'T00:00:00'),
      clean_(payload.CustomerName), clean_(payload.Model), clean_(payload.SubModel), clean_(payload.Color),
      clean_(payload.Finance), clean_(payload.BookingStatus), clean_(payload.CarStatus), clean_(payload.Note),
      clean_(payload.SalesConsultant), rowIndex > 0 ? values[rowIndex-1][11] : now, now
    ];
    if (rowIndex > 0) sh.getRange(rowIndex,1,1,CONFIG.HEADERS.length).setValues([record]);
    else sh.appendRow(record);
    return {ok:true, booking:rowToObject_(record), message: rowIndex > 0 ? 'แก้ไขข้อมูลแล้ว' : 'บันทึก Booking แล้ว'};
  } finally { lock.releaseLock(); }
}

function deleteBooking(id) {
  if (!id) throw new Error('ไม่พบ Booking ID');
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CONFIG.BOOKING_SHEET);
  if (!sh || sh.getLastRow() < 2) throw new Error('ยังไม่มีข้อมูล');
  const values = sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();
  for (let i=0;i<values.length;i++) {
    if (values[i][0] === String(id)) { sh.deleteRow(i+2); return {ok:true,message:'ลบรายการแล้ว'}; }
  }
  throw new Error('ไม่พบรายการที่ต้องการลบ');
}

function importExistingBookings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const src = ss.getSheetByName(CONFIG.SOURCE_SHEET);
  if (!src) throw new Error('ไม่พบชีตต้นทาง ' + CONFIG.SOURCE_SHEET);
  setupBkM2();
  const target = ss.getSheetByName(CONFIG.BOOKING_SHEET);
  if (target.getLastRow() > 1) throw new Error('ชีตปลายทางมีข้อมูลแล้ว เพื่อป้องกันการนำเข้าซ้ำ ระบบยกเลิกการนำเข้า');
  const data = src.getDataRange().getValues();
  const out = [];
  for (let i=3;i<data.length;i++) {
    const r = data[i];
    const date = r[1], name = r[3], model = r[4];
    if (!date || !name || !model) continue;
    const status = clean_(r[8]) || 'จองใหม่';
    // Scope is MG only; exclude any non-MG rows.
    if (!/^MG/i.test(String(model).trim())) continue;
    out.push([
      Utilities.getUuid(), date instanceof Date ? date : new Date(date),
      clean_(name), clean_(model), clean_(r[5]), clean_(r[6]), clean_(r[7]),
      status, clean_(r[9]), clean_(r[10]), clean_(r[11]), new Date(), new Date()
    ]);
  }
  if (out.length) target.getRange(2,1,out.length,CONFIG.HEADERS.length).setValues(out);
  return {ok:true, imported:out.length, message:'นำเข้าข้อมูล MG จากชีตต้นทางแล้ว'};
}

function getDashboard() {
  const data = getInitialData().bookings;
  const active = data.filter(b => b.BookingStatus !== 'ยกเลิก (ระบุเหตุผลในหมายเหตุ)');
  const delivered = active.filter(b => b.BookingStatus === 'ส่งมอบแล้ว').length;
  const pending = active.filter(b => b.BookingStatus.indexOf('รอส่งมอบ') >= 0).length;
  const now = new Date();
  const month = Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyy-MM');
  const newThisMonth = active.filter(b => String(b.BookingDate).slice(0,7) === month).length;
  const byModel = {};
  const byStatus = {};
  const bySales = {};
  active.forEach(b => {
    byModel[b.Model] = (byModel[b.Model] || 0) + 1;
    byStatus[b.BookingStatus] = (byStatus[b.BookingStatus] || 0) + 1;
    const name = b.SalesConsultant || 'ไม่ระบุ';
    bySales[name] = (bySales[name] || 0) + 1;
  });
  return {total:active.length, delivered, pending, newThisMonth, cancelled:data.length-active.length, byModel, byStatus, bySales};
}

function validatePayload_(p) {
  if (!p) throw new Error('ไม่มีข้อมูลที่ส่งมา');
  ['BookingDate','CustomerName','Model','SubModel','Color','Finance','BookingStatus','CarStatus','SalesConsultant'].forEach(k => {
    if (!String(p[k] || '').trim()) throw new Error('กรุณากรอกข้อมูล: ' + k);
  });
  const date = new Date(p.BookingDate + 'T00:00:00');
  if (isNaN(date.getTime())) throw new Error('วันที่จองไม่ถูกต้อง');
  if (!/^MG/i.test(p.Model)) throw new Error('ระบบนี้รองรับเฉพาะแบรนด์ MG');
  if (String(p.CustomerName).length > 150) throw new Error('ชื่อลูกค้ายาวเกินไป');
}
function rowToObject_(r) {
  return {
    BookingID:String(r[0] || ''),
    BookingDate:dateCellKey_(r[1]),
    CustomerName:String(r[2] || ''), Model:String(r[3] || ''), SubModel:String(r[4] || ''),
    Color:String(r[5] || ''), Finance:String(r[6] || ''), BookingStatus:String(r[7] || ''),
    CarStatus:String(r[8] || ''), Note:String(r[9] || ''), SalesConsultant:String(r[10] || ''),
    CreatedAt:r[11] instanceof Date ? r[11].toISOString() : String(r[11] || ''),
    UpdatedAt:r[12] instanceof Date ? r[12].toISOString() : String(r[12] || '')
  };
}
function dateCellKey_(v) {
  if (v instanceof Date && !isNaN(v.getTime())) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const s = String(v || '');
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
  const d = new Date(s);
  return isNaN(d.getTime()) ? s : Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}
function clean_(v) { return String(v == null ? '' : v).trim(); }
