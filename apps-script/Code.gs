/**
 * GR Production Dashboard — Google Apps Script Web App
 *
 * ต้นทาง: Google Sheet "Incentive GR" ชีต 'GR'
 * ปลายทาง: GR_Production_Dashboard*.html (fetchLive)
 *
 * -------------------------------------------------------------------------
 * สำคัญ: การกรอกข้อมูลใน Sheet ไม่ต้อง deploy ใหม่
 * doGet() รันสดทุกครั้งที่มีการเรียก และอ่านค่าจาก Sheet ณ วินาทีนั้น
 * จะต้อง deploy ใหม่เฉพาะตอนที่แก้ไฟล์นี้เท่านั้น
 *
 * ตอน deploy: Deploy → Manage deployments → ✏️ → Version: New version
 * อย่ากด "New deployment" เพราะจะได้ URL ใหม่ แล้วเครื่องที่เคยเชื่อมต่อ
 * สำเร็จจะยังยิงไป URL เดิมที่ค้างใน localStorage
 * -------------------------------------------------------------------------
 */

var SHEET_NAME = 'GR';

// ต้องตรงกับค่า AUTH_HASH ในไฟล์ HTML (sha256 ของรหัสผ่านเข้าใช้งาน)
var EXPECTED_KEY_HASH = 'fb8da96a6ac06bda69cddfc23e875dbad850043989033f66e42acbb5c6ce91c9';

/**
 * พารามิเตอร์ที่รองรับ (ทุกตัวเป็น optional — ไม่ใส่ = ได้ข้อมูลทั้งหมดเหมือนเดิม)
 *
 *   key    = AUTH_HASH                (บังคับ)
 *   meta   = 1                        คืนแค่สรุป {rows, last_date, ...} ไม่ต้องโหลดข้อมูลทั้งก้อน
 *   since  = YYYY-MM-DD               คืนเฉพาะแถวที่ date_iso >= ค่านี้
 *   slim   = 1                        ตัดฟิลด์ที่ dashboard ไม่ได้ใช้ออก (payload เล็กลง ~35%)
 */
function doGet(e) {
  try {
    var p = (e && e.parameter) ? e.parameter : {};

    if ((p.key || '') !== EXPECTED_KEY_HASH) {
      return jsonResponse({ error: 'Unauthorized: missing or invalid key' });
    }

    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet) {
      return jsonResponse({ error: 'ไม่พบชีตชื่อ "' + SHEET_NAME + '" ในไฟล์นี้' });
    }

    var lastCol = sheet.getLastColumn();
    if (lastCol < 1 || sheet.getLastRow() < 2) return jsonResponse([]);

    // ---- 1) อ่านหัวตารางแถวเดียว เพื่อหาตำแหน่งคอลัมน์ ----
    var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function (h) {
      return (h || '').toString().trim();
    });
    var idx = mapColumns(headers);

    if (idx.date < 0) {
      return jsonResponse({ error: 'ไม่พบคอลัมน์ "วันที่ทำงาน" ในแถวหัวตาราง' });
    }

    // ---- 2) หาแถวสุดท้ายที่มี "วันที่ทำงาน" จริง ----
    //
    // ห้ามใช้ getLastRow() ตรง ๆ เป็นขอบเขตการอ่านข้อมูล เพราะคอลัมน์ "ลำดับ"
    // ถูกลากเลขไว้ล่วงหน้าเกินกว่าแถวที่มีข้อมูลจริงมาก Google จึงนับแถวเปล่า
    // เหล่านั้นว่ามีเนื้อหาด้วย (ณ ส.ค. 2026: getLastRow() = 3,125 แต่มีข้อมูล
    // จริง 1,213 แถว) การอ่านคอลัมน์วันที่คอลัมน์เดียวก่อน ถูกกว่าการดึงทั้ง
    // ตารางประมาณ 25 เท่า แล้วค่อยอ่านเฉพาะบล็อกที่มีข้อมูลจริง
    var scanRows = sheet.getLastRow() - 1;
    var dateCol = sheet.getRange(2, idx.date + 1, scanRows, 1).getDisplayValues();

    var lastDataRow = 0;
    for (var i = dateCol.length - 1; i >= 0; i--) {
      if ((dateCol[i][0] || '').toString().trim() !== '') { lastDataRow = i + 2; break; }
    }
    if (!lastDataRow) return jsonResponse(p.meta ? emptyMeta() : []);

    // ---- 3) โหมด meta: ตอบสรุปอย่างเดียว ไม่ต้องอ่านทั้งตาราง ----
    if (p.meta) {
      var lastIso = null, count = 0;
      for (var m = 0; m < dateCol.length; m++) {
        var iso = parseDate(dateCol[m][0]);
        if (!iso) continue;
        count++;
        if (!lastIso || iso > lastIso) lastIso = iso;
      }
      return jsonResponse({
        rows: count,
        last_date: lastIso,
        last_data_row: lastDataRow,
        sheet_last_row: sheet.getLastRow(),
        generated_at: new Date().toISOString()
      });
    }

    // ---- 4) อ่านเฉพาะบล็อกที่มีข้อมูลจริง ----
    var values = sheet.getRange(2, 1, lastDataRow - 1, lastCol).getDisplayValues();

    var since = (p.since || '').toString().trim();   // 'YYYY-MM-DD'
    var slim = !!p.slim;
    var out = [];

    for (var r = 0; r < values.length; r++) {
      var row = values[r];
      var dateRaw = cell(row, idx.date);
      var dateIso = parseDate(dateRaw);
      if (!dateIso) continue;                        // ข้ามแถวว่าง / แถวที่ไม่มีวันที่
      if (since && dateIso < since) continue;

      var actual = toNum(cell(row, idx.actual_pcs));
      var cost = toNum(cell(row, idx.cost_per_pc));

      var rec = {
        date_iso: dateIso,
        month: dateIso.substring(0, 7),
        customer: cell(row, idx.customer).trim(),
        po: cell(row, idx.po),
        item_code: cell(row, idx.item_code),
        dept: cell(row, idx.dept).trim(),
        material: cell(row, idx.material),
        machine_hours: toNum(cell(row, idx.machine_hours)),
        target_100: toNum(cell(row, idx.target_100)),
        target_95: toNum(cell(row, idx.target_95)),
        required_pcs: toNum(cell(row, idx.required_pcs)),
        actual_pcs: actual,
        defects: toNum(cell(row, idx.defects)),
        employee: cell(row, idx.employee).trim(),
        excess_baht: toNum(cell(row, idx.excess_baht)),
        value_baht: (actual != null && cost != null) ? actual * cost : null
      };

      // ฟิลด์ที่ dashboard ปัจจุบันไม่ได้ใช้ — ส่งต่อไว้เพื่อความเข้ากันได้ย้อนหลัง
      // เรียก ?slim=1 เพื่อตัดออกเมื่อข้อมูลโตจนขนาด payload เริ่มเป็นปัญหา
      //
      // หมายเหตุ achv_pct: dashboard คำนวณเองจาก actual_pcs / required_pcs
      // (แก้ไว้ใน f531d24) ค่าที่ส่งจากที่นี่ใช้ target_100 เป็นตัวหาร ซึ่งเป็น
      // คนละสูตรและถูกทิ้งทั้งหมดที่ฝั่ง client — คงไว้เพื่อไม่ให้ผู้ใช้เดิมพัง
      if (!slim) {
        rec.no = cell(row, idx.no);
        rec.date = dateRaw;
        rec.work_hours = toNum(cell(row, idx.work_hours));
        rec.num_workers = toNum(cell(row, idx.num_workers));
        rec.cost_per_pc = cost;
        rec.excess_pcs = toNum(cell(row, idx.excess_pcs));
        rec.excess_pct = toPct(cell(row, idx.excess_pct));
        var t100 = rec.target_100;
        rec.achv_pct = (actual != null && t100) ? Math.round((actual / t100) * 1000) / 10 : null;
      }

      out.push(rec);
    }

    return jsonResponse(out);
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

/** จับคู่ชื่อหัวตารางกับฟิลด์ — ทนต่อการสลับ/แทรกคอลัมน์ใน Sheet */
function mapColumns(headers) {
  function colIndex(patterns) {
    for (var i = 0; i < headers.length; i++) {
      var h = headers[i];
      var ok = patterns.every(function (p) { return h.indexOf(p) !== -1; });
      if (ok) return i;
    }
    return -1;
  }
  return {
    no: colIndex(['ลำดับ']),
    date: colIndex(['วันที่ทำงาน']),
    customer: colIndex(['ลูกค้า']),
    po: colIndex(['เลขที่ใบสั่งผลิต']),
    item_code: colIndex(['รหัส', 'Item']),
    dept: colIndex(['หน่วยงาน']),
    material: colIndex(['ชื่อวัตถุดิบ']),
    work_hours: colIndex(['เวลาทำงานทั้งหมด']),
    machine_hours: colIndex(['เครื่องจักรทำงาน']),
    target_100: colIndex(['เป้าหมาย', '100']),
    target_95: colIndex(['เป้าหมาย', '95']),
    required_pcs: colIndex(['ต้องทำ']),
    actual_pcs: colIndex(['ได้จริง']),
    defects: colIndex(['ของเสีย']),
    num_workers: colIndex(['จำนวนคนทำงาน']),
    employee: colIndex(['ชื่อพนักงาน']),
    cost_per_pc: colIndex(['ต้นทุน']),
    excess_baht: colIndex(['ส่วนต่างที่เกินจากเป้า']),
    excess_pcs: colIndex(['จำนวนชิ้นที่เกินจากเป้า']),
    excess_pct: colIndex(['ที่เกินจากเป้า', '%'])
  };
}

function cell(row, i) {
  if (i < 0 || i >= row.length) return '';
  return (row[i] === null || row[i] === undefined) ? '' : row[i].toString();
}

function toNum(s) {
  s = (s || '').toString().trim().replace(/,/g, '');
  if (s === '') return null;
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function toPct(s) {
  s = (s || '').toString().trim().replace(/%/g, '');
  if (s === '') return null;
  var n = parseFloat(s);
  return isNaN(n) ? null : n;
}

/** 'D/M/YYYY' (หรือ 'D/M/YY') → 'YYYY-MM-DD' */
function parseDate(s) {
  s = (s || '').toString().trim();
  var parts = s.split('/');
  if (parts.length !== 3) return null;
  var day = parseInt(parts[0], 10), mon = parseInt(parts[1], 10), yr = parseInt(parts[2], 10);
  if (!day || !mon || !yr) return null;
  if (mon > 12 || day > 31) return null;
  if (yr < 100) yr += 2000;
  var mm = mon < 10 ? '0' + mon : '' + mon;
  var dd = day < 10 ? '0' + day : '' + day;
  return yr + '-' + mm + '-' + dd;
}

function emptyMeta() {
  return { rows: 0, last_date: null, generated_at: new Date().toISOString() };
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
