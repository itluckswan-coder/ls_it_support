// =========================
// CONFIG
// =========================
const SPREADSHEET_ID = '1pCQTfoK96qvBrAZxy_nnVLGT2EZXlcupGQtkGk3FRjM';
const SHEET_NAME = 'แจ้งปัญหา';

// =========================
// HTML VIEW (Web App)
// =========================
function doGet(e) {
  try {
    const template = HtmlService.createTemplateFromFile('index');
    return template
      .evaluate()
      .setTitle('ระบบแจ้งปัญหา IT')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return HtmlService.createHtmlOutput("❌ เกิดข้อผิดพลาดในการโหลดหน้าเว็บ: " + err.message);
  }
}

// =========================
// ส่งเรื่องแจ้งปัญหา
// =========================
function submitIssue(payload) {
  try {
    if (!payload) throw new Error("❌ ไม่พบข้อมูลที่ส่งมา");
    if (!payload.name) throw new Error("กรุณากรอกชื่อผู้แจ้ง");
    if (!payload.dept) throw new Error("กรุณาเลือกแผนก");
    if (!payload.device) throw new Error("กรุณาเลือกอุปกรณ์/ระบบ");
    if (!payload.issue) throw new Error("กรุณากรอกรายละเอียดอาการ");

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error("❌ ไม่พบชีตชื่อ: " + SHEET_NAME);

    const ticketId = "T" + new Date().getTime();

    sh.appendRow([
      new Date(),
      ticketId,
      payload.name,
      payload.dept,
      payload.device,
      payload.issue,
      'กำลังดำเนินการ',
      'Thinnathep',
      ''
    ]);

    // ล้าง cache ทิ้งเพื่อให้ค้นหาข้อมูลล่าสุดได้
    CacheService.getScriptCache().remove("ticket_data");

    return { ok: true, ticketId };

  } catch (err) {
    return { ok: false, error: err.message };
  }
}

// =========================
// ค้นหาสถานะ (ล่างขึ้นบน)
// =========================
function getStatus(query) {
  try {
    const cache = CacheService.getScriptCache();
    let cached = cache.get("ticket_data");

    let values, header;

    if (cached) {
      const parsed = JSON.parse(cached);
      header = parsed.header;
      values = parsed.values;
    } else {
      const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
      const sh = ss.getSheetByName(SHEET_NAME);
      if (!sh) throw new Error("❌ ไม่พบชีต");

      header = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

      // อ่านเฉพาะ 100 แถวล่าสุด (ปรับได้)
      const lastRow = sh.getLastRow();
      const startRow = Math.max(2, lastRow - 100);
      values = sh.getRange(startRow, 1, lastRow - startRow + 1, header.length).getValues();

      cache.put(
        "ticket_data",
        JSON.stringify({ header, values }),
        60 * 2 // cache 2 นาที
      );
    }

    const idx = {
      Timestamp: header.indexOf('Timestamp'),
      TicketID: header.indexOf('TicketID'),
      Name: header.indexOf('Name'),
      Department: header.indexOf('Department'),
      Device: header.indexOf('Device'),
      Issue: header.indexOf('Issue'),
      Status: header.indexOf('Status'),
      Assignee: header.indexOf('Assignee'),
      Note: header.indexOf('Note')
    };

    const normalize = t => (t || '').toString().trim().toLowerCase().normalize('NFC');
    const q = normalize(query);

    let latest = null;

    for (let i = values.length - 1; i >= 0; i--) { // 🔥 loop from bottom
      if (normalize(values[i][idx.Name]).includes(q)) {
        latest = convertRow(values[i], idx);
        break;
      }
    }

    return { ok: true, items: latest ? [latest] : [] };

  } catch (err) {
    return { ok: false, error: err.message, items: [] };
  }
}

// =========================
// Dashboard: ดึงรายการทั้งหมด
// =========================
function getAllTickets() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sh = ss.getSheetByName(SHEET_NAME);
    if (!sh) throw new Error("❌ ไม่พบชีตชื่อ: " + SHEET_NAME);

    const values = sh.getDataRange().getValues();
    if (values.length < 2) throw new Error("❌ ยังไม่มีข้อมูลในชีต");

    const header = values.shift();

    const idx = {
      Timestamp: header.indexOf('Timestamp') !== -1 ? header.indexOf('Timestamp') : header.indexOf('วันที่เวลา'),
      TicketID: header.indexOf('TicketID') !== -1 ? header.indexOf('TicketID') : header.indexOf('รหัส Ticket'),
      Name: header.indexOf('Name') !== -1 ? header.indexOf('Name') : header.indexOf('ผู้แจ้ง'),
      Department: header.indexOf('Department') !== -1 ? header.indexOf('Department') : header.indexOf('แผนก'),
      Device: header.indexOf('Device') !== -1 ? header.indexOf('Device') : header.indexOf('อุปกรณ์/ระบบ'),
      Issue: header.indexOf('Issue') !== -1 ? header.indexOf('Issue') : header.indexOf('อาการ'),
      Status: header.indexOf('Status') !== -1 ? header.indexOf('Status') : header.indexOf('สถานะ'),
      Assignee: header.indexOf('Assignee') !== -1 ? header.indexOf('Assignee') : header.indexOf('ผู้รับผิดชอบ'),
      Note: header.indexOf('Note') !== -1 ? header.indexOf('Note') : header.indexOf('หมายเหตุ')
    };

    // Reverse เพื่อให้ล่าสุดขึ้นบน
    const items = values.map(r => convertRow(r, idx)).reverse();
    return { ok: true, items };

  } catch (err) {
    return { ok: false, error: err.message, items: [] };
  }
}

// =========================
// Helper: แปลงแถวเป็น Object
// =========================
function safe(r, i) {
  return i >= 0 ? r[i] : '';
}

function convertRow(r, idx) {
  return {
    ticketId: safe(r, idx.TicketID),
    name: safe(r, idx.Name),
    department: safe(r, idx.Department),
    device: safe(r, idx.Device),
    issue: safe(r, idx.Issue),
    status: safe(r, idx.Status),
    assignee: safe(r, idx.Assignee),
    timestamp: safe(r, idx.Timestamp),
    note: safe(r, idx.Note)
  };
}
