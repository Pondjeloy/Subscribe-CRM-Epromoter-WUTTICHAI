// ════════════════════════════════════════════════════════
//  LG Subscribe CRM — Google Apps Script API  (MASTER)
//  Deploy: Web App | Execute as: Me | Anyone
//  แก้แล้ว Redeploy: Manage Deployments → Edit → New version → Deploy
//
//  [แก้ไขล่าสุด] 'Lead LG Success': statusCol/picCol/notesCol เลื่อน
//  จาก 8/9/10 → 7/8/9 เพราะโครงสร้างชีตมีคอลัมน์ระหว่าง
//  Model Code กับ Status ลดลง (เหลือ Order No. / Total Rental
//  Amount / Price Policy Name 3 คอลัมน์) ทำให้ Status ที่เคย
//  อยู่คอลัมน์ I เลื่อนมาอยู่คอลัมน์ H, PIC จาก J → I, Remark
//  จาก K → J ตามภาพชีตที่ส่งมา (name/phone/productType ที่
//  column B/C/D ไม่เปลี่ยน จึงไม่ต้องแก้ใน parse())
//
//  หมายเหตุ: 'Lead Subscribe Lg.com' คืนกลับเป็นค่าเดิมแล้ว
//  (รอบที่แล้วแก้ผิดชีตเพราะเข้าใจผิดว่าภาพที่ส่งมาคือชีตนี้)
//
//  [Sync ส.ค. 2569] Spreadsheet ใหม่ (Aug 2026)
//  https://docs.google.com/spreadsheets/d/1aiyNPYJHy3TA0NRj7vNUBgZs_htLiPJ3VALGeL_nLA0
//  (เดิม ก.ค.: 1wEiFHLZKq9ZKEEeuiNEvap-dCzzgrQl0t0nFtt7ZfOI)
//
//  [Aug ส.ค. 2569] Meta Densu Aug — ชื่อชีตเปลี่ยนจาก Meta Densu July
//  โครงคอลัมน์เดียวกับ July: N=Status O=PIC(Epromoter) P=Remark  (picCol=14)
//  อ่านทุกแถวในชีต Meta Densu Aug (alias: Meta Densu July ยัง resolve ได้)
//
//  [Hybrid] ชีตรายเดือน + หลีดอื่น mapping เดียวกับมิถุนายน
//  Meta → logic Meta Densu Aug / Meta Densu / Meta ITAX
//  หลีดอื่น (LG.com, LG Success, Consult, POP UP Braner) → column map เดิม
//
//  [POP UP ก.ค. 2569 — ดึง 0 รายชื่อ]
//  โครง Braner (ยืนยันจากชีตจริง):
//  A=วันที่ B=ยินยอม C=ชื่อ D=นามสกุล E=อีเมล F=เบอร์ G=Line H=จังหวัด
//  I=ประเภทสินค้า J=รหัสตัวแทน(ถ้ามี) K=Status L=PIC M=Remark
//  หาแท็บแบบ alias/fuzzy + อ่านหัวตาราง PIC/Status/Remark อัตโนมัติ
//  ถ้า picCol เดิม match 0 แถว → สแกนหาคอลัมน์ที่มีชื่อ promoter มากสุด
//
//  [Meta ITAX] โครงสร้างคอลัมน์เดียวกับ Meta Densu Aug
//  A=ชำระ C=จังหวัด E=สินค้า F=วันที่สะดวก G=เวลาติดต่อ
//  H=ชื่อ I=อายุ J=เบอร์ K=Email N=สถานะ O=Epromoter (P=หมายเหตุ ถ้ามี)
// ════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1aiyNPYJHy3TA0NRj7vNUBgZs_htLiPJ3VALGeL_nLA0'; // ส.ค. 2569
var PROMOTER       = 'POND';

// ชื่อ canonical ใน CRM (1 การ์ด POP UP) — resolveSheet หาแท็บจริงให้
var SHEET_NAMES = [
  'Meta Densu Aug','Meta Densu','Meta ITAX',
  'Lead Subscribe Lg.com','Lead LG Success','Lead Consult',
  'Lead Subscribe POP UP Braner'
];

// ชื่อแท็บทางเลือก (สะกดผิด / เปลี่ยนชื่อ)
var SHEET_ALIASES = {
  'Meta Densu Aug': [
    'Meta Densu Aug',
    'Meta Densu August',
    'Meta Densu Aug 2026',
    'Meta Densu ส.ค.',
    'Meta Densu July' // fallback ชีตเก่า
  ],
  'Lead Subscribe POP UP Braner': [
    'Lead Subscribe POP UP Braner',
    'Lead Subscribe POP UP Banner',
    'Lead POP UP Braner',
    'Lead POP UP Banner',
    'POP UP Braner',
    'POP UP Banner',
    'POP UP Bannar',
    'POP UP',
    'Lead Subscribe POP UP'
  ]
};

// ── ทดสอบก่อน Deploy ครั้งแรก ──────────────────────────
function testConnection() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('OK: ' + ss.getSheets().map(function(s){ return s.getName(); }).join(', '));
}

// ── Entry point ─────────────────────────────────────────
function doGet(e) {
  var p        = (e && e.parameter) ? e.parameter : {};
  var action   = p.action || 'ping';
  var promoter = p.promoter || PROMOTER;
  var result;

  try {
    switch (action) {
      case 'ping':          result = { success:true, message:'API ready', promoter:promoter }; break;
      case 'getCustomers':  result = getCustomers(promoter); break;
      case 'debugSheets':   result = debugSheets(promoter); break;
      case 'getHeaders':    result = getHeaders(); break;
      case 'checkRows':     result = checkRows(p.sheet||'Lead LG Success', parseInt(p.n||'5',10), promoter); break;
      case 'updateStatus':  result = updateStatus(p.sheet, parseInt(p.row,10), p.status||''); break;
      case 'getNotes':      result = getNotes(p.sheet, parseInt(p.row,10)); break;
      case 'appendNote':    result = appendNote(p.sheet, parseInt(p.row,10), p.note||''); break;
      case 'updateNotes':   result = updateNotes(p.sheet, parseInt(p.row,10), p.notes||''); break;
      case 'setNoteHighlight': result = setNoteHighlight(p.sheet, parseInt(p.row,10), parseInt(p.level||'0',10)); break;
      default:              result = { success:false, error:'Unknown action: '+action };
    }
  } catch(err) {
    result = { success:false, error:err.toString() };
  }

  var cb   = p.callback;
  var json = JSON.stringify(result);
  var mime = cb ? ContentService.MimeType.JAVASCRIPT : ContentService.MimeType.JSON;
  return ContentService.createTextOutput(cb ? cb+'('+json+')' : json).setMimeType(mime);
}

// ── Column config (0-based: A=0, B=1, … M=12 … W=22) ───
function getSheetConfig(name) {
  var cfg = {

    // Meta Densu Aug: A=ชำระ C=จังหวัด E=สินค้า F=วันที่สะดวก G=ช่วงเวลาติดต่อ
    // H=ชื่อ I=อายุ J=เบอร์ K=email N=สถานะ O=Epromoter P=หมายเหตุ
    'Meta Densu Aug': {
      picCol:14, statusCol:13, notesCol:15,
      parse: function(row, disp) { return parseMetaDensuJulyRow(row, disp); }
    },
    // alias ชื่อเก่า (ชีต July ยังอยู่) — โครงคอลัมน์เดียวกัน
    'Meta Densu July': {
      picCol:14, statusCol:13, notesCol:15,
      parse: function(row, disp) { return parseMetaDensuJulyRow(row, disp); }
    },

    // Meta Densu (เก่า): โครงสร้างมิถุนายน — M=Status N=Epromoter O=Remark
    'Meta Densu': {
      picCol:13, statusCol:12, notesCol:14,
      parse: function(row, disp) { return parseMetaDensuLegacyRow(row, disp); }
    },

    // Meta ITAX: คอลัมน์เดียวกับ Meta Densu Aug
    // A=ชำระ C=จังหวัด E=สินค้า F=วันที่สะดวก G=เวลาติดต่อ
    // H=ชื่อ I=อายุ J=เบอร์ K=Email N=สถานะ O=Epromoter P=หมายเหตุ
    'Meta ITAX': {
      picCol:14, statusCol:13, notesCol:15,
      parse: function(row, disp) { return parseMetaDensuJulyRow(row, disp); }
    },

    // ── หลีดอื่น: mapping มิถุนายน 2569 (บนชีตกรกฎาคม) ──────────────
    'Lead Subscribe Lg.com': {
      picCol:8, statusCol:7, notesCol:9,
      parse: function(row) { return {
        name:           clean(row[1]),
        phone:          clean(row[6]),
        email:          '',
        age:            '',
        paymentChannel: '',
        province:       '',
        productType:    clean(row[3]),
        lineId:         ''
      };}
    },

    'Lead LG Success': {
      picCol:8, statusCol:7, notesCol:9,
      parse: function(row) { return {
        name:           clean(row[1]),   // B - Customer Name
        phone:          clean(row[2]),   // C - Mobile No.
        email:          '',
        age:            '',
        paymentChannel: '',
        province:       '',
        productType:    clean(row[3]),   // D - Model Code
        lineId:         ''
      };}
    },

    'Lead Consult': {
      picCol:22, statusCol:21, notesCol:23,
      parse: function(row) { return {
        name:           clean(row[4]),
        phone:          clean(row[6]),
        email:          clean(row[5]),
        age:            '',
        paymentChannel: '',
        province:       clean(row[10]),
        productType:    clean(row[9]),
        lineId:         clean(row[7])
      };}
    },

    // POP UP Braner: A วันที่ B ยินยอม C ชื่อ D นามสกุล E อีเมล F เบอร์
    // G Line H จังหวัด I สินค้า J รหัสตัวแทน K Status L PIC M Remark
    'Lead Subscribe POP UP Braner': {
      picCol:11, statusCol:10, notesCol:12,
      isPopup: true,
      parse: function(row) {
        var first = clean(row[2]), last = clean(row[3]);
        return {
          name:           (first+' '+last).trim(),
          phone:          clean(row[5]),
          email:          clean(row[4]),
          age:            '',
          paymentChannel: '',
          province:       clean(row[7]),
          productType:    clean(row[8]),
          lineId:         clean(row[6])
        };
      }
    },

    // โครงสั้น (ถ้ามีแท็บนี้): B ชื่อเต็ม C เบอร์ F Sale Consultant G สถานะ H หมายเหตุ
    'POP UP Bannar': {
      picCol:5, statusCol:6, notesCol:7,
      isPopup: true,
      parse: function(row) {
        var type = clean(row[3]), prod = clean(row[4]);
        var productType = type && prod ? (type + ' · ' + prod) : (type || prod);
        return {
          name:           clean(row[1]),
          phone:          clean(row[2]),
          email:          '',
          age:            '',
          paymentChannel: '',
          province:       '',
          productType:    productType,
          lineId:         ''
        };
      }
    }
  };
  // alias ชื่อแท็บ → ใช้ config เดียวกับ Braner
  if (!cfg[name] && isPopupSheetName(name)) {
    return cfg['Lead Subscribe POP UP Braner'];
  }
  return cfg[name] || null;
}

// ── หาชีตจริง (alias + fuzzy POP UP) ────────────────────
function isPopupSheetName(name) {
  return /pop\s*up|braner|bannar|banner/i.test(String(name || ''));
}

function resolveSheet(ss, preferredName) {
  var tried = {};
  function tryName(n) {
    if (!n || tried[n]) return null;
    tried[n] = true;
    var sh = ss.getSheetByName(n);
    return sh ? { sheet: sh, actualName: n } : null;
  }
  var hit = tryName(preferredName);
  if (hit) return hit;
  var aliases = SHEET_ALIASES[preferredName] || [];
  for (var i = 0; i < aliases.length; i++) {
    hit = tryName(aliases[i]);
    if (hit) return hit;
  }
  if (isPopupSheetName(preferredName)) {
    var all = ss.getSheets();
    for (var j = 0; j < all.length; j++) {
      var n = all[j].getName();
      if (isPopupSheetName(n)) return { sheet: all[j], actualName: n };
    }
  }
  return null;
}

function findHeaderCol(headers, keywords) {
  for (var i = 0; i < headers.length; i++) {
    var h = normalizeKey(headers[i]);
    if (!h) continue;
    for (var k = 0; k < keywords.length; k++) {
      if (h.indexOf(normalizeKey(keywords[k])) !== -1) return i;
    }
  }
  return -1;
}

// สแกนหาคอลัมน์ที่มี promoter มากสุด (เมื่อ map เดิม match 0)
function autoFindPicCol(data, promoter) {
  var maxCols = 0, r, c, n, bestCol = -1, bestCount = 0;
  for (r = 1; r < data.length; r++) {
    if (data[r].length > maxCols) maxCols = data[r].length;
  }
  for (c = 0; c < maxCols; c++) {
    n = 0;
    for (r = 1; r < data.length; r++) {
      if (data[r].length > c && isPromoter(data[r][c], promoter)) n++;
    }
    if (n > bestCount) { bestCount = n; bestCol = c; }
  }
  return bestCount > 0 ? bestCol : -1;
}

function countPromoterInCol(data, col, promoter) {
  var n = 0, i;
  if (col < 0) return 0;
  for (i = 1; i < data.length; i++) {
    if (data[i].length > col && isPromoter(data[i][col], promoter)) n++;
  }
  return n;
}

// ปรับ config POP UP จากหัวตาราง + auto PIC
function refinePopupConfig(baseCfg, headers, data, promoter) {
  var cfg = {
    picCol: baseCfg.picCol,
    statusCol: baseCfg.statusCol,
    notesCol: baseCfg.notesCol,
    isPopup: true,
    parse: baseCfg.parse,
    detect: {}
  };
  var picH = findHeaderCol(headers, ['PIC', 'Epromoter', 'Sale Consultant', 'Consultant']);
  var stH  = findHeaderCol(headers, ['Status', 'สถานะ', 'ดำเนินการ']);
  var noH  = findHeaderCol(headers, ['Remark', 'หมายเหตุ', 'Notes', 'Note']);
  if (picH >= 0) { cfg.picCol = picH; cfg.detect.picFromHeader = colLetter(picH); }
  if (stH  >= 0) { cfg.statusCol = stH; cfg.detect.statusFromHeader = colLetter(stH); }
  if (noH  >= 0) { cfg.notesCol = noH; cfg.detect.notesFromHeader = colLetter(noH); }

  // โครงสั้น: มี Sale Consultant / คอลัมน์น้อย → parse แบบ Bannar
  var hJoin = headers.map(function(h){ return clean(h).toLowerCase(); }).join(' ');
  if (hJoin.indexOf('sale consultant') !== -1 || (headers.length <= 10 && picH >= 0 && picH <= 6)) {
    var shortCfg = getSheetConfig('POP UP Bannar');
    if (shortCfg) {
      cfg.parse = shortCfg.parse;
      if (picH < 0) cfg.picCol = shortCfg.picCol;
      if (stH < 0)  cfg.statusCol = shortCfg.statusCol;
      if (noH < 0)  cfg.notesCol = shortCfg.notesCol;
      cfg.detect.layout = 'short';
    }
  } else {
    cfg.detect.layout = 'braner';
  }

  var matched = countPromoterInCol(data, cfg.picCol, promoter);
  cfg.detect.matchedWithPicCol = matched;
  if (matched === 0 && data.length > 1) {
    var auto = autoFindPicCol(data, promoter);
    if (auto >= 0) {
      cfg.detect.autoPicCol = colLetter(auto);
      cfg.picCol = auto;
      // status/notes มักอยู่ข้างๆ PIC: Status ก่อน PIC, Remark หลัง PIC
      if (stH < 0) cfg.statusCol = Math.max(0, auto - 1);
      if (noH < 0) cfg.notesCol = auto + 1;
    }
  }
  return cfg;
}

function getRuntimeConfig(canonicalName, sheet, data, promoter) {
  var cfg = getSheetConfig(canonicalName) || getSheetConfig(sheet.getName());
  if (!cfg) return null;
  if (!cfg.isPopup && !isPopupSheetName(canonicalName) && !isPopupSheetName(sheet.getName())) {
    return cfg;
  }
  var headers = (data && data.length) ? data[0] : sheet.getRange(1,1,1,Math.max(sheet.getLastColumn(),1)).getValues()[0];
  return refinePopupConfig(cfg, headers, data || [headers], promoter);
}

// ── getCustomers ────────────────────────────────────────
function getCustomers(promoter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetNames = SHEET_NAMES;
  var all = [], idNum = 1, sourceCounts = {}, sheetsFound = {};
  var usedSheetIds = {}; // กันอ่านแท็บ POP UP ซ้ำ

  for (var s = 0; s < sheetNames.length; s++) {
    var sName = sheetNames[s];
    var baseCfg = getSheetConfig(sName);
    if (!baseCfg) continue;
    var resolved = resolveSheet(ss, sName);
    if (!resolved) { Logger.log('Not found: '+sName); sheetsFound[sName] = false; continue; }
    var sheet = resolved.sheet;
    var actualName = resolved.actualName;
    var sheetId = String(sheet.getSheetId());
    if (usedSheetIds[sheetId]) {
      sheetsFound[sName] = true;
      sourceCounts[sName] = sourceCounts[sName] || 0;
      continue;
    }
    usedSheetIds[sheetId] = true;
    sheetsFound[sName] = true;
    // รายงานทั้งชื่อ canonical และชื่อแท็บจริง
    if (actualName !== sName) sheetsFound[actualName] = true;
    sourceCounts[sName] = 0;

    var range = sheet.getDataRange();
    var data  = range.getValues();
    var disp  = range.getDisplayValues();
    var cfg   = getRuntimeConfig(sName, sheet, data, promoter);
    if (!cfg) continue;

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.length <= cfg.picCol) continue;
      if (!isPromoter(row[cfg.picCol], promoter)) continue;

      var fields = cfg.parse(row, disp[i]);
      if (!fields.name && !fields.phone) continue;

      var notes = '';
      if (cfg.notesCol !== undefined && row.length > cfg.notesCol)
        notes = clean(row[cfg.notesCol]);

      // source = ชื่อแท็บจริง (เพื่อ updateStatus/notes เขียนถูกชีต)
      var cust = { id:idNum++, source:actualName, row:i+1,
                   status:clean(row[cfg.statusCol]), notes:notes };
      for (var k in fields) cust[k] = fields[k];
      all.push(cust);
      sourceCounts[sName]++;
      if (actualName !== sName) {
        sourceCounts[actualName] = (sourceCounts[actualName] || 0) + 1;
      }
    }
  }
  var deduped = dedupeCustomers(all);
  return { success:true, count:deduped.length, rawCount:all.length,
           sourceCounts:sourceCounts, sheetsFound:sheetsFound, data:deduped };
}

// เปิดชีต + config รันไทม์ (รองรับ alias / หัวตาราง POP UP)
function openSheetWithConfig(sheetName, promoter) {
  if (!sheetName) return null;
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  var actual = sheetName;
  if (!sheet) {
    var resolved = resolveSheet(ss, sheetName);
    if (!resolved) return null;
    sheet = resolved.sheet;
    actual = resolved.actualName;
  }
  var data = sheet.getDataRange().getValues();
  var cfg = getRuntimeConfig(actual, sheet, data, promoter || PROMOTER)
         || getSheetConfig(actual)
         || getSheetConfig(sheetName);
  if (!cfg) return null;
  return { sheet: sheet, cfg: cfg, name: actual };
}

// ── updateStatus ────────────────────────────────────────
function updateStatus(sheetName, rowNum, newStatus) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var opened = openSheetWithConfig(sheetName);
  if (!opened) return { success:false, error:'Sheet not found: '+sheetName };
  opened.sheet.getRange(rowNum, opened.cfg.statusCol+1).setValue(newStatus);
  SpreadsheetApp.flush();
  return { success:true, sheet:opened.name, row:rowNum, status:newStatus };
}

// ── getNotes ────────────────────────────────────────────
function getNotes(sheetName, rowNum) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var opened = openSheetWithConfig(sheetName);
  if (!opened || opened.cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  return { success:true, notes:clean(opened.sheet.getRange(rowNum, opened.cfg.notesCol+1).getValue()) };
}

// ── appendNote ──────────────────────────────────────────
function appendNote(sheetName, rowNum, note) {
  if (!sheetName || !rowNum || !note) return { success:false, error:'Missing params' };
  var opened = openSheetWithConfig(sheetName);
  if (!opened || opened.cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  var cell = opened.sheet.getRange(rowNum, opened.cfg.notesCol+1);
  var cur  = clean(cell.getValue());
  cell.setValue(cur ? cur+'\n'+note : note);
  SpreadsheetApp.flush();
  return { success:true };
}

// ── setNoteHighlight — ใส่สีพื้นหลัง + ตัวหนา + ตัวอักษรแดง ช่อง Remark ──
// ไม่แก้ข้อความในเซลล์ — แค่จัดรูปแบบ
// level: 0=ล้างสไตล์, 1=เหลือง, 2=ส้ม, 3=เขียว
function setNoteHighlight(sheetName, rowNum, level) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var opened = openSheetWithConfig(sheetName);
  if (!opened || opened.cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  var colors = {
    0: null,        // ล้างสีพื้น
    1: '#FEF08A',   // เหลือง
    2: '#FDBA74',   // ส้ม
    3: '#86EFAC'    // เขียว
  };
  var lv = (level === 1 || level === 2 || level === 3) ? level : 0;
  var cell = opened.sheet.getRange(rowNum, opened.cfg.notesCol+1);
  cell.setBackground(colors[lv]);
  if (lv === 0) {
    cell.setFontWeight('normal');
    cell.setFontColor(null);
  } else {
    cell.setFontWeight('bold');
    cell.setFontColor('#C8102E');
  }
  SpreadsheetApp.flush();
  return { success:true, sheet:opened.name, row:rowNum, level:lv, color:colors[lv] };
}

// ── updateNotes (overwrite — used for delete) ───────────
function updateNotes(sheetName, rowNum, notes) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var opened = openSheetWithConfig(sheetName);
  if (!opened || opened.cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  opened.sheet.getRange(rowNum, opened.cfg.notesCol+1).setValue(notes);
  SpreadsheetApp.flush();
  return { success:true };
}

// ── debugSheets ─────────────────────────────────────────
function debugSheets(promoter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetNames = SHEET_NAMES;
  var report = [];
  sheetNames.forEach(function(sName) {
    var resolved = resolveSheet(ss, sName);
    if (!resolved) { report.push({sheet:sName, found:false}); return; }
    var data = resolved.sheet.getDataRange().getValues();
    var cfg  = getRuntimeConfig(sName, resolved.sheet, data, promoter);
    if (!cfg) { report.push({sheet:sName, found:true, actualName:resolved.actualName, hint:'no config'}); return; }
    var matched = 0, counts = {};
    for (var i = 1; i < data.length; i++) {
      var v = data[i].length > cfg.picCol ? clean(data[i][cfg.picCol]) : '(short)';
      counts[v] = (counts[v]||0)+1;
      if (isPromoter(v, promoter)) matched++;
    }
    report.push({
      sheet:sName, actualName:resolved.actualName, found:true,
      totalRows:data.length-1, picColumn:colLetter(cfg.picCol),
      matchedPOND:matched, allPICValues:counts, detect:cfg.detect||null
    });
  });
  return { success:true, allTabs:ss.getSheets().map(function(s){return s.getName();}), report:report };
}

// ── getHeaders ──────────────────────────────────────────
function getHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {};
  SHEET_NAMES.forEach(function(n) {
    var resolved = resolveSheet(ss, n);
    if (!resolved) { result[n]='ไม่พบ Sheet'; return; }
    result[n] = resolved.sheet.getRange(1,1,1,resolved.sheet.getLastColumn()).getValues()[0]
                     .map(function(h,i){ return colLetter(i)+': '+(h||'(ว่าง)'); });
    if (resolved.actualName !== n) result[n+' (แท็บจริง: '+resolved.actualName+')'] = result[n];
  });
  return { success:true, headers:result, allTabs:ss.getSheets().map(function(s){return s.getName();}) };
}

// ── checkRows ───────────────────────────────────────────
function checkRows(sheetName, n, promoter) {
  var opened = openSheetWithConfig(sheetName, promoter);
  if (!opened) return { success:false, error:'ไม่พบ Sheet: '+sheetName };
  var cfg = opened.cfg, data = opened.sheet.getDataRange().getValues();
  var start = Math.max(1, data.length-n), rows = [];
  for (var i = start; i < data.length; i++) {
    var row    = data[i];
    var picVal = row.length > cfg.picCol ? String(row[cfg.picCol]) : '(short)';
    var fields = cfg.parse(row);
    var skip   = '';
    if (row.length <= cfg.picCol)          skip = 'แถวสั้น';
    else if (!isPromoter(picVal, promoter)) skip = colLetter(cfg.picCol)+'="'+picVal+'" ไม่ใช่ '+(promoter||PROMOTER);
    else if (!fields.name && !fields.phone) skip = 'ชื่อ+เบอร์ว่าง';
    rows.push({ sheetRow:i+1, passed:skip==='', skipReason:skip||'-',
                picVal:picVal, name:fields.name||'', phone:fields.phone||'' });
  }
  return { success:true, sheet:opened.name, totalRows:data.length-1, checkedLast:n,
           picColumn:colLetter(cfg.picCol), detect:cfg.detect||null, rows:rows };
}

// ── Helpers ─────────────────────────────────────────────
function clean(v) {
  if (v===null||v===undefined) return '';
  return String(v).trim();
}

function cleanDisplay(v, displayVal) {
  if (v instanceof Date) return String(displayVal||'').trim();
  return clean(v);
}

function normalizeKey(v) {
  return clean(v).replace(/[​-‍﻿]/g,'').replace(/\s+/g,'').toUpperCase();
}

function isPromoter(v, promoter) {
  return normalizeKey(v).indexOf(normalizeKey(promoter || PROMOTER)) !== -1;
}

function colLetter(index) {
  var n=index+1, s='';
  while (n>0) { var r=(n-1)%26; s=String.fromCharCode(65+r)+s; n=Math.floor((n-1)/26); }
  return s;
}

function calcAge(val) {
  if (!val) return '';
  var d;
  if (val instanceof Date)                             d = val;
  else if (typeof val==='number' && val>1000)          d = new Date(Math.round((val-25569)*86400*1000));
  else                                                  d = new Date(val);
  if (!isNaN(d.getTime())) {
    var age = Math.floor((new Date()-d)/(365.25*24*60*60*1000));
    if (age>0 && age<120) return age;
  }
  return clean(val);
}

// ── Meta Densu เก่า (มิ.ย.) — อ่านรายชื่อ legacy ในชีตกรกฎาคม ──
// A=จังหวัด B=สินค้า D=ชำระ F=ชื่อ G=เพศ H/I=เบอร์ J=email · M=Status N=Epromoter O=Remark
function parseMetaDensuLegacyRow(row, disp) {
  var hDisp = cleanDisplay(row[7], disp&&disp[7]);
  var iVal  = clean(row[8]);
  var jVal  = clean(row[9]);
  var email = '';
  if (iVal.indexOf('@') !== -1)      email = iVal;
  else if (jVal.indexOf('@') !== -1) email = jVal;
  function isPhone(v) {
    return v.replace(/\D/g,'').length >= 9 && v.indexOf('@') === -1;
  }
  var phone = isPhone(iVal) ? iVal : (isPhone(hDisp) ? hDisp : '');
  var ageRaw = calcAge(row[7]);
  var age = (String(ageRaw).replace(/\D/g,'').length >= 9) ? '' : ageRaw;
  return {
    name:           clean(row[5]),
    phone:          phone,
    email:          email,
    age:            age,
    contactTime:    '',
    convenientDate: '',
    paymentChannel: clean(row[3]),
    province:       clean(row[0]),
    productType:    clean(row[1]),
    lineId:         ''
  };
}

// ── Meta Densu July / Meta ITAX parse ─────────────────
// A=ช่องทางชำระ C=จังหวัด E=สินค้า F=วันที่สะดวก G=ช่วงเวลาติดต่อ
// H=ชื่อ I=อายุ J=เบอร์ K=email · N=สถานะ O=Epromoter P=หมายเหตุ
function parseMetaDensuJulyRow(row, disp) {
  var ageRaw = calcAge(row[8]);
  var age = (String(ageRaw).replace(/\D/g,'').length >= 9) ? '' : ageRaw;
  return {
    name:           clean(row[7]),
    phone:          cleanDisplay(row[9], disp&&disp[9]),
    email:          clean(row[10]),
    age:            age,
    contactTime:    clean(row[6]),
    convenientDate: cleanDisplay(row[5], disp&&disp[5]),
    paymentChannel: clean(row[0]),
    province:       clean(row[2]),
    productType:    clean(row[4]),
    lineId:         ''
  };
}

function isJuly2026Row(row, dateCol, dispRow) {
  var v = row[dateCol];
  var dispVal = dispRow && dispRow[dateCol];
  if (v === null || v === undefined || v === '') return false;
  var d;
  if (v instanceof Date) d = v;
  else if (typeof v === 'number' && v > 1000) d = new Date(Math.round((v - 25569) * 86400 * 1000));
  else {
    var s = clean(dispVal || v);
    var m = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
    if (m) {
      var y = parseInt(m[3], 10);
      if (y >= 2500) y -= 543;
      d = new Date(y, parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }
  }
  if (!d || isNaN(d.getTime())) return false;
  return d.getFullYear() === 2026 && d.getMonth() === 6;
}

// ── Dedup Meta Densu Aug + Meta Densu เก่า (Aug ชนะถ้าเบอร์ซ้ำ) ──
var META_PAIR = {'Meta Densu Aug':true, 'Meta Densu July':true, 'Meta Densu':true};
var SOURCE_PRIORITY = {
  'Meta Densu Aug': 1,
  'Meta Densu July': 2,
  'Meta Densu': 3
};

function phoneKey(phone) {
  var d = clean(phone).replace(/\D/g, '');
  if (d.length >= 11 && d.indexOf('66') === 0) d = '0' + d.slice(2);
  if (d.length === 9) d = '0' + d;
  return d.length >= 9 ? d : '';
}

function pickPreferredCustomer(a, b) {
  var pa = SOURCE_PRIORITY[a.source] || 50;
  var pb = SOURCE_PRIORITY[b.source] || 50;
  if (pa !== pb) return pa < pb ? a : b;
  var sa = clean(a.status).length, sb = clean(b.status).length;
  if (sa !== sb) return sa > sb ? a : b;
  return (a.row || 0) >= (b.row || 0) ? a : b;
}

function dedupeMetaOnly(list) {
  var byPhone = {}, noPhone = [], out = [], k, i, a;
  for (i = 0; i < list.length; i++) {
    var c = list[i];
    var key = phoneKey(c.phone);
    if (!key) { noPhone.push(c); continue; }
    if (!byPhone[key]) { byPhone[key] = {p: c, alts: []}; continue; }
    var bucket = byPhone[key];
    var winner = pickPreferredCustomer(bucket.p, c);
    var loser  = winner === bucket.p ? c : bucket.p;
    bucket.p = winner;
    bucket.alts.push({source: loser.source, row: loser.row, status: loser.status});
  }
  for (k in byPhone) {
    var b = byPhone[k], p = b.p;
    if (b.alts.length) {
      p.altRows = b.alts;
      p.sourceMerged = p.source;
      for (a = 0; a < b.alts.length; a++) {
        if (p.sourceMerged.indexOf(b.alts[a].source) === -1)
          p.sourceMerged += ' · ' + b.alts[a].source;
      }
    }
    out.push(p);
  }
  return out.concat(noPhone);
}

function dedupeCustomers(list) {
  var meta = [], rest = [], i, id = 1;
  for (i = 0; i < list.length; i++) {
    if (META_PAIR[list[i].source]) meta.push(list[i]);
    else rest.push(list[i]);
  }
  var merged = dedupeMetaOnly(meta).concat(rest);
  for (i = 0; i < merged.length; i++) merged[i].id = id++;
  return merged;
}