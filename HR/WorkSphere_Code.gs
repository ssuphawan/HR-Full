// =====================================================
// WorkSphere HR Enterprise — Google Apps Script
// Version: 2.0  |  2568
// =====================================================
//
// ✅ คำแนะนำการตั้งค่า (ทำครั้งเดียว):
//
// 1. สร้าง Google Sheet ใหม่
// 2. Extensions → Apps Script → วางโค้ดนี้ใน Code.gs
// 3. ใส่ Spreadsheet ID ใน SPREADSHEET_ID ด้านล่าง
//    (copy จาก URL: /spreadsheets/d/[ID_ตรงนี้]/edit)
// 4. บันทึกไฟล์ (Ctrl+S)
// 5. รัน setupSheets() ครั้งแรก (เพื่อสร้าง Sheet + ข้อมูลตัวอย่าง)
// 6. Deploy → New deployment
//      Type        : Web app
//      Execute as  : Me
//      Who has access: Anyone
// 7. Copy "Web app URL" → ไปใส่ใน HTML ที่ var GAS_URL = '...'
// =====================================================

const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE';

// ── ชื่อ Sheet ทั้งหมด ─────────────────────────────
const SHEET = {
  users:       'Users',
  employees:   'Employees',
  departments: 'Departments',
  positions:   'Positions',
  attendance:  'Attendance',
  leave:       'Leave',
  shift:       'Shift',
  ot:          'OT',
  assets:      'Assets',
  expense:     'Expense',
  meeting:     'Meeting',
  announce:    'Announce',
  kpi:         'KPI',
  docs:        'Docs',
};

// ── Schema: columns ของแต่ละ Sheet ─────────────────
const SCHEMA = {
  Users:       ['ID','Name','Email','Password','Role','Ini','Grad','Status','CreatedAt'],
  Employees:   ['ID','FirstName','LastName','Email','Phone','Department','Position','StartDate','Salary','Type','Status','CreatedAt'],
  Departments: ['ID','Name','Head','Description','Status','EmployeeCount','CreatedAt'],
  Positions:   ['ID','Name','Department','Grade','MinSalary','MaxSalary','EmployeeCount','CreatedAt'],
  Attendance:  ['ID','Employee','Date','TimeIn','TimeOut','TotalHours','Status','Note','CreatedAt'],
  Leave:       ['ID','Employee','Type','StartDate','EndDate','Days','Reason','Delegate','Status','Note','CreatedAt'],
  Shift:       ['ID','Employee','ShiftName','StartDate','EndDate','TimeIn','TimeOut','Department','CreatedAt'],
  OT:          ['ID','Employee','Date','Hours','Rate','Reason','Amount','Status','Note','CreatedAt'],
  Assets:      ['ID','Name','Type','SerialNo','Value','Owner','ReceiveDate','Note','Status','CreatedAt'],
  Expense:     ['ID','Employee','Type','Amount','Description','Date','Status','Note','CreatedAt'],
  Meeting:     ['ID','Room','Title','Booker','Date','StartTime','EndTime','Attendees','Equipment','Note','Status','CreatedAt'],
  Announce:    ['ID','Title','Category','Level','Content','Pinned','Author','Status','CreatedAt'],
  KPI:         ['ID','Name','Department','Quarter','Target','Current','Progress','Weight','Status','CreatedAt'],
  Docs:        ['ID','Name','Type','Date','Uploader','FileUrl','Status','CreatedAt'],
};

// ─────────────────────────────────────────────────────
// HELPERS
// ─────────────────────────────────────────────────────

function resp_(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
function ok_(extra)   { return resp_(Object.assign({ ok: true  }, extra || {})); }
function fail_(msg)   { return resp_({ ok: false, error: msg }); }

function ss_()        { return SpreadsheetApp.openById(SPREADSHEET_ID); }
function ws_(name)    {
  var w = ss_().getSheetByName(name);
  if (!w) throw new Error('ไม่พบ Sheet: ' + name);
  return w;
}

function genId_(prefix) {
  return (prefix || 'ROW').substring(0, 3).toUpperCase() + '-' +
    Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyMMddHHmm') + '-' +
    Math.floor(Math.random() * 90 + 10);
}

function now_() {
  return Utilities.formatDate(new Date(), 'Asia/Bangkok', 'dd/MM/yy HH:mm');
}

function toJson_(ws) {
  var raw = ws.getDataRange().getValues();
  if (raw.length < 2) return [];
  var headers = raw[0];
  return raw.slice(1).map(function(row) {
    var obj = {};
    headers.forEach(function(h, i) { obj[h] = row[i] !== undefined ? String(row[i]) : ''; });
    return obj;
  });
}

// ─────────────────────────────────────────────────────
// doGet — อ่านข้อมูล
// ─────────────────────────────────────────────────────

function doGet(e) {
  try {
    var sheetKey = e.parameter.sheet;
    var sheetName = SHEET[sheetKey];
    if (!sheetName) return fail_('ไม่รู้จัก sheet: ' + sheetKey);
    var rows = toJson_(ws_(sheetName));
    return ok_({ data: rows });
  } catch(err) {
    return fail_(err.message);
  }
}

// ─────────────────────────────────────────────────────
// doPost — เขียน / แก้ไข / ลบ / login
// ─────────────────────────────────────────────────────

function doPost(e) {
  try {
    var body   = JSON.parse(e.postData.contents);
    var action = body.action;
    var sKey   = body.sheet;
    var data   = body.data   || {};
    var id     = body.id     || '';
    var status = body.status || '';
    var note   = body.note   || '';

    if (action === 'login') return handleLogin_(data);

    var sheetName = SHEET[sKey];
    if (!sheetName) return fail_('ไม่รู้จัก sheet: ' + sKey);
    var sheet = ws_(sheetName);

    switch (action) {
      case 'insert':  return doInsert_(sheet, sheetName, data);
      case 'update':  return doUpdate_(sheet, id, data);
      case 'delete':  return doDelete_(sheet, id);
      case 'approve': return doApprove_(sheet, id, status, note);
      default:        return fail_('ไม่รู้จัก action: ' + action);
    }
  } catch(err) {
    return fail_(err.message);
  }
}

// ─────────────────────────────────────────────────────
// LOGIN
// ─────────────────────────────────────────────────────

function handleLogin_(data) {
  var email = (data.email || '').toLowerCase().trim();
  var pass  = (data.password || '').trim();
  var rows  = toJson_(ws_(SHEET.users));
  var user  = rows.filter(function(r) {
    return r.Email.toLowerCase() === email && r.Password === pass && r.Status === 'ใช้งาน';
  })[0];
  if (!user) return fail_('อีเมลหรือรหัสผ่านไม่ถูกต้อง');
  // ไม่ส่ง Password กลับ
  return ok_({ user: { Name: user.Name, Role: user.Role, Ini: user.Ini, Grad: user.Grad, Email: user.Email } });
}

// ─────────────────────────────────────────────────────
// CRUD
// ─────────────────────────────────────────────────────

function doInsert_(sheet, sheetName, data) {
  var schema = SCHEMA[sheetName];
  if (!schema) return fail_('ไม่มี schema สำหรับ: ' + sheetName);

  if (!data.ID)        data.ID        = genId_(sheetName);
  if (!data.CreatedAt) data.CreatedAt = now_();
  if (data.Status === undefined || data.Status === '') data.Status = 'ใช้งาน';

  var row = schema.map(function(h) { return data[h] !== undefined ? data[h] : ''; });
  sheet.appendRow(row);
  SpreadsheetApp.flush();
  return ok_({ row: data });
}

function doUpdate_(sheet, id, data) {
  var all     = sheet.getDataRange().getValues();
  var headers = all[0];
  var idIdx   = headers.indexOf('ID');
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][idIdx]) === String(id)) {
      headers.forEach(function(h, j) {
        if (data[h] !== undefined) sheet.getRange(i + 1, j + 1).setValue(data[h]);
      });
      SpreadsheetApp.flush();
      return ok_({});
    }
  }
  return fail_('ไม่พบ ID: ' + id);
}

function doDelete_(sheet, id) {
  var all   = sheet.getDataRange().getValues();
  var idIdx = all[0].indexOf('ID');
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][idIdx]) === String(id)) {
      sheet.deleteRow(i + 1);
      SpreadsheetApp.flush();
      return ok_({});
    }
  }
  return fail_('ไม่พบ ID: ' + id);
}

function doApprove_(sheet, id, status, note) {
  var all       = sheet.getDataRange().getValues();
  var headers   = all[0];
  var idIdx     = headers.indexOf('ID');
  var statusIdx = headers.indexOf('Status');
  var noteIdx   = headers.indexOf('Note');
  for (var i = 1; i < all.length; i++) {
    if (String(all[i][idIdx]) === String(id)) {
      if (statusIdx >= 0) sheet.getRange(i + 1, statusIdx + 1).setValue(status);
      if (noteIdx   >= 0 && note) sheet.getRange(i + 1, noteIdx + 1).setValue(note);
      SpreadsheetApp.flush();
      return ok_({});
    }
  }
  return fail_('ไม่พบ ID: ' + id);
}

// ─────────────────────────────────────────────────────
// SETUP — รันครั้งเดียวเพื่อสร้าง Sheet ทั้งหมด
// ─────────────────────────────────────────────────────

function setupSheets() {
  var ss  = ss_();
  var now = now_();

  // สร้างทุก Sheet ตาม Schema
  Object.keys(SCHEMA).forEach(function(name) {
    var headers = SCHEMA[name];
    var sheet   = ss.getSheetByName(name) || ss.insertSheet(name);
    sheet.clearContents();

    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setValues([headers])
      .setFontWeight('bold')
      .setBackground('#7c3aed')
      .setFontColor('#ffffff')
      .setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.setColumnWidths(1, headers.length, 160);
  });

  // ── ข้อมูลตั้งต้น: Users ──────────────────────────
  var usersWs = ss.getSheetByName('Users');
  usersWs.getRange(2, 1, 3, 9).setValues([
    ['USR-001','อภิชาต วงศ์สุวรรณ','admin@techstar.co.th','admin1234','Super Admin','SA',
     'linear-gradient(135deg,#7c3aed,#a855f7)','ใช้งาน', now],
    ['USR-002','สุภาวณี ดีมาก','s.suphawansr@gmail.com','hr1234','HR Manager','สภ',
     'linear-gradient(135deg,#ec4899,#f472b6)','ใช้งาน', now],
    ['USR-003','พนักงาน ทดสอบ','emp@techstar.co.th','emp1234','Employee','EM',
     'linear-gradient(135deg,#3b82f6,#14b8a6)','ใช้งาน', now],
  ]);

  // ── ข้อมูลตั้งต้น: Departments ───────────────────
  var deptWs = ss.getSheetByName('Departments');
  deptWs.getRange(2, 1, 6, 7).setValues([
    ['DEP-001','เทคโนโลยี',  'วสันต์ มั่นคง',   'ทีมพัฒนาซอฟต์แวร์และ IT',   'ดำเนินการ','248',now],
    ['DEP-002','การตลาด',    'นวลจันทร์ แสนดี', 'ทีมการตลาดและสื่อสารองค์กร','ดำเนินการ','95', now],
    ['DEP-003','การขาย',     'สมพงษ์ ใจดี',     'ทีมขายและพัฒนาธุรกิจ',       'ดำเนินการ','187',now],
    ['DEP-004','บัญชี',      'กานดา นาคแก้ว',   'ทีมบัญชีและการเงิน',          'ดำเนินการ','62', now],
    ['DEP-005','ปฏิบัติการ', 'อรรถพล สุขใส',    'ทีมปฏิบัติการและโลจิสติกส์', 'ดำเนินการ','412',now],
    ['DEP-006','HR',          'สุภาวณี ดีมาก',   'ทีมทรัพยากรบุคคล',            'ดำเนินการ','38', now],
  ]);

  // ── ข้อมูลตั้งต้น: Positions ──────────────────────
  var posWs = ss.getSheetByName('Positions');
  posWs.getRange(2, 1, 4, 8).setValues([
    ['POS-001','Senior Developer','เทคโนโลยี','Grade 5','70000','120000','42',now],
    ['POS-002','Marketing Manager','การตลาด','Grade 6','80000','140000','8', now],
    ['POS-003','Sales Executive','การขาย','Grade 3','30000','60000','95', now],
    ['POS-004','Accountant','บัญชี','Grade 4','45000','80000','20',now],
  ]);

  // ── ข้อมูลตั้งต้น: Employees ──────────────────────
  var empWs = ss.getSheetByName('Employees');
  empWs.getRange(2, 1, 4, 12).setValues([
    ['EMP-001','สมพงษ์','ใจดี','somphong@techstar.co.th','081-234-5678','เทคโนโลยี','Senior Developer','01/01/66','85000','พนักงานประจำ','ทำงาน',now],
    ['EMP-002','นวลจันทร์','แสนดี','nualjan@techstar.co.th','082-345-6789','การตลาด','Marketing Manager','01/03/66','78000','พนักงานประจำ','ทำงาน',now],
    ['EMP-003','วสันต์','มั่นคง','wasan@techstar.co.th','083-456-7890','เทคโนโลยี','Tech Lead','15/06/65','95000','พนักงานประจำ','ทำงาน',now],
    ['EMP-004','กานดา','นาคแก้ว','kanda@techstar.co.th','084-567-8901','บัญชี','Accountant','01/09/66','55000','พนักงานประจำ','ทำงาน',now],
  ]);

  // ── ข้อมูลตั้งต้น: Assets ─────────────────────────
  var astWs = ss.getSheetByName('Assets');
  astWs.getRange(2, 1, 3, 10).setValues([
    ['AST-001','MacBook Pro 14','คอมพิวเตอร์','C02XXXXXX','85000','สมพงษ์ ใจดี','01/01/66','','ใช้งาน',now],
    ['AST-002','iPhone 15 Pro','มือถือ','F2XXXXXXX','45000','นวลจันทร์ แสนดี','01/03/66','','ใช้งาน',now],
    ['AST-003','Dell 27" Monitor','จอแสดงผล','CN-XXXXXX','18000','','01/06/66','','ว่าง',now],
  ]);

  // ── ข้อมูลตั้งต้น: Announce ───────────────────────
  var annWs = ss.getSheetByName('Announce');
  annWs.getRange(2, 1, 2, 9).setValues([
    ['ANN-001','ประกาศหยุดวันหยุดนักขัตฤกษ์','HR','สำคัญ','บริษัทจะหยุดในวันที่ 5 มิถุนายน 2568 วันวิสาขบูชา','true','สุภาวณี ดีมาก','เผยแพร่',now],
    ['ANN-002','อบรม Safety ประจำปี 2568','อบรม','ทั่วไป','ขอเชิญพนักงานทุกท่านเข้าร่วมอบรม Safety วันที่ 15 มิ.ย. 2568','false','สุภาวณี ดีมาก','เผยแพร่',now],
  ]);

  // ── ข้อมูลตั้งต้น: KPI ────────────────────────────
  var kpiWs = ss.getSheetByName('KPI');
  kpiWs.getRange(2, 1, 3, 10).setValues([
    ['KPI-001','ยอดขายรายเดือน','การขาย','Q2/2568','฿50M','฿43M','86','30','ใช้งาน',now],
    ['KPI-002','อัตราการรักษาพนักงาน','HR','Q2/2568','92%','94%','100','20','ใช้งาน',now],
    ['KPI-003','ลูกค้าใหม่ต่อเดือน','การตลาด','Q2/2568','500 ราย','312 ราย','62','25','ใช้งาน',now],
  ]);

  SpreadsheetApp.flush();

  Browser.msgBox(
    '✅ WorkSphere HR — Setup เสร็จสมบูรณ์!\n\n' +
    'สร้าง Sheet ทั้งหมด ' + Object.keys(SCHEMA).length + ' Sheet\n' +
    'พร้อมข้อมูลตัวอย่างแล้ว\n\n' +
    '➤ ขั้นตอนต่อไป:\n' +
    '  Deploy → Manage deployments → New deployment\n' +
    '  Type: Web app | Execute as: Me | Access: Anyone\n\n' +
    '  แล้ว Copy URL ไปใส่ใน HTML\n  ที่บรรทัด: var GAS_URL = "URL_ที่นี่"'
  );
}
