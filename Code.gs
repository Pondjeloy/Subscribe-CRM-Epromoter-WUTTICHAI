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
//  [Sync ก.ค. 2569] Spreadsheet ใหม่
//  https://docs.google.com/spreadsheets/d/1wEiFHLZKq9ZKEEeuiNEvap-dCzzgrQl0t0nFtt7ZfOI
//
//  [July ก.ค. 2569] Meta Densu July — เลื่อนคอลัมน์ท้าย +1 จาก Meta Densu เดิม
//  Meta Densu เดิม: M=Status N=PIC(Epromoter) O=Remark
//  Meta Densu July:  N=Status O=PIC(Epromoter) P=Remark  (picCol=14)
//  อ่านทุกแถวในชีต Meta Densu July (ไม่กรองวันที่ D — ชื่อชีตคือ July แล้ว)
//  รวมเบอร์ซ้ำเฉพาะ Meta Densu July + Meta Credit July
//
//  [Hybrid ก.ค. 2569] ชีตใหม่กรกฎาคม + หลีดอื่น mapping เดียวกับมิถุนายน
//  Meta → logic กรกฎาคม (Meta Densu/Credit July)
//  หลีดอื่น (LG.com, LG Success, Consult, POP UP*) → column map เดิมมิ.ย. 2569
// ════════════════════════════════════════════════════════

var SPREADSHEET_ID = '1wEiFHLZKq9ZKEEeuiNEvap-dCzzgrQl0t0nFtt7ZfOI';
var PROMOTER       = 'POND';

var SHEET_NAMES = [
  'Meta Densu July','Meta Densu',
  'Lead Subscribe Lg.com','Lead LG Success','Lead Consult',
  'Lead Subscribe POP UP Braner'
];

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

    // Meta Densu July: A=ชำระ C=จังหวัด E=สินค้า F=วันที่สะดวก G=ช่วงเวลาติดต่อ
    // H=ชื่อ I=อายุ J=เบอร์ K=email N=สถานะ O=Epromoter P=หมายเหตุ
    'Meta Densu July': {
      picCol:14, statusCol:13, notesCol:15,
      parse: function(row, disp) { return parseMetaDensuJulyRow(row, disp); }
    },

    // Meta Densu (เก่า): โครงสร้างมิถุนายน — M=Status N=Epromoter O=Remark
    'Meta Densu': {
      picCol:13, statusCol:12, notesCol:14,
      parse: function(row, disp) { return parseMetaDensuLegacyRow(row, disp); }
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

    'Lead Subscribe POP UP Braner': {
      picCol:11, statusCol:10, notesCol:12,
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
    }
  };
  return cfg[name] || null;
}

// ── getCustomers ────────────────────────────────────────
function getCustomers(promoter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetNames = SHEET_NAMES;
  var all = [], idNum = 1, sourceCounts = {}, sheetsFound = {};

  for (var s = 0; s < sheetNames.length; s++) {
    var sName = sheetNames[s];
    var cfg   = getSheetConfig(sName);
    if (!cfg) continue;
    var sheet = ss.getSheetByName(sName);
    if (!sheet) { Logger.log('Not found: '+sName); sheetsFound[sName] = false; continue; }
    sheetsFound[sName] = true;
    sourceCounts[sName] = 0;

    var range = sheet.getDataRange();
    var data  = range.getValues();
    var disp  = range.getDisplayValues();

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.length <= cfg.picCol) continue;
      if (!isPromoter(row[cfg.picCol], promoter)) continue;

      var fields = cfg.parse(row, disp[i]);
      if (!fields.name && !fields.phone) continue;

      var notes = '';
      if (cfg.notesCol !== undefined && row.length > cfg.notesCol)
        notes = clean(row[cfg.notesCol]);

      var cust = { id:idNum++, source:sName, row:i+1,
                   status:clean(row[cfg.statusCol]), notes:notes };
      for (var k in fields) cust[k] = fields[k];
      all.push(cust);
      sourceCounts[sName]++;
    }
  }
  var deduped = dedupeCustomers(all);
  return { success:true, count:deduped.length, rawCount:all.length,
           sourceCounts:sourceCounts, sheetsFound:sheetsFound, data:deduped };
}

// ── updateStatus ────────────────────────────────────────
function updateStatus(sheetName, rowNum, newStatus) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var cfg = getSheetConfig(sheetName);
  if (!cfg) return { success:false, error:'Unknown sheet: '+sheetName };
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success:false, error:'Sheet not found' };
  sheet.getRange(rowNum, cfg.statusCol+1).setValue(newStatus);
  SpreadsheetApp.flush();
  return { success:true, sheet:sheetName, row:rowNum, status:newStatus };
}

// ── getNotes ────────────────────────────────────────────
function getNotes(sheetName, rowNum) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var cfg = getSheetConfig(sheetName);
  if (!cfg || cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success:false, error:'Sheet not found' };
  return { success:true, notes:clean(sheet.getRange(rowNum, cfg.notesCol+1).getValue()) };
}

// ── appendNote ──────────────────────────────────────────
function appendNote(sheetName, rowNum, note) {
  if (!sheetName || !rowNum || !note) return { success:false, error:'Missing params' };
  var cfg = getSheetConfig(sheetName);
  if (!cfg || cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success:false, error:'Sheet not found' };
  var cell = sheet.getRange(rowNum, cfg.notesCol+1);
  var cur  = clean(cell.getValue());
  cell.setValue(cur ? cur+'\n'+note : note);
  SpreadsheetApp.flush();
  return { success:true };
}

// ── setNoteHighlight — ใส่สีพื้นหลังช่อง Remark (ไม่แตะข้อความ) ──
// level: 0=ล้างสี, 1=เหลือง(สนใจ), 2=ส้ม(โอกาสสูง), 3=เขียว(ใกล้ปิด)
function setNoteHighlight(sheetName, rowNum, level) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var cfg = getSheetConfig(sheetName);
  if (!cfg || cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success:false, error:'Sheet not found' };
  var colors = {
    0: null,        // ล้างสี
    1: '#FEF08A',   // เหลือง — สนใจ
    2: '#FDBA74',   // ส้ม — มีโอกาสสูง
    3: '#86EFAC'    // เขียว — ใกล้ปิด
  };
  var lv = (level === 1 || level === 2 || level === 3) ? level : 0;
  var cell = sheet.getRange(rowNum, cfg.notesCol+1);
  cell.setBackground(colors[lv]);
  SpreadsheetApp.flush();
  return { success:true, sheet:sheetName, row:rowNum, level:lv, color:colors[lv] };
}

// ── updateNotes (overwrite — used for delete) ───────────
function updateNotes(sheetName, rowNum, notes) {
  if (!sheetName || !rowNum) return { success:false, error:'Missing params' };
  var cfg = getSheetConfig(sheetName);
  if (!cfg || cfg.notesCol===undefined) return { success:false, error:'No config/notesCol' };
  var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
  if (!sheet) return { success:false, error:'Sheet not found' };
  sheet.getRange(rowNum, cfg.notesCol+1).setValue(notes);
  SpreadsheetApp.flush();
  return { success:true };
}

// ── debugSheets ─────────────────────────────────────────
function debugSheets(promoter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheetNames = SHEET_NAMES;
  var report = [];
  sheetNames.forEach(function(sName) {
    var cfg   = getSheetConfig(sName);
    var sheet = ss.getSheetByName(sName);
    if (!sheet) { report.push({sheet:sName, found:false}); return; }
    if (!cfg)   { report.push({sheet:sName, found:true, hint:'no config'}); return; }
    var data = sheet.getDataRange().getValues();
    var matched = 0, counts = {};
    for (var i = 1; i < data.length; i++) {
      var v = data[i].length > cfg.picCol ? clean(data[i][cfg.picCol]) : '(short)';
      counts[v] = (counts[v]||0)+1;
      if (isPromoter(v, promoter)) matched++;
    }
    report.push({ sheet:sName, found:true, totalRows:data.length-1,
                  picColumn:colLetter(cfg.picCol), matchedPOND:matched, allPICValues:counts });
  });
  return { success:true, allTabs:ss.getSheets().map(function(s){return s.getName();}), report:report };
}

// ── getHeaders ──────────────────────────────────────────
function getHeaders() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var result = {};
  SHEET_NAMES.forEach(function(n) {
    var sheet = ss.getSheetByName(n);
    if (!sheet) { result[n]='ไม่พบ Sheet'; return; }
    result[n] = sheet.getRange(1,1,1,sheet.getLastColumn()).getValues()[0]
                     .map(function(h,i){ return colLetter(i)+': '+(h||'(ว่าง)'); });
  });
  return { success:true, headers:result };
}

// ── checkRows ───────────────────────────────────────────
function checkRows(sheetName, n, promoter) {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success:false, error:'ไม่พบ Sheet: '+sheetName };
  var cfg = getSheetConfig(sheetName);
  if (!cfg)  return { success:false, error:'ไม่มี config: '+sheetName };
  var data = sheet.getDataRange().getValues();
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
                picVal:picVal, name:String(row[1]||''), phone:String(row[2]||'') });
  }
  return { success:true, sheet:sheetName, totalRows:data.length-1, checkedLast:n, rows:rows };
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

// ── Meta Densu July parse ─────────────────────────────
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

// ── Dedup Meta Densu July + Meta Densu เก่า (July ชนะถ้าเบอร์ซ้ำ) ──
var META_PAIR = {'Meta Densu July':true, 'Meta Densu':true};
var SOURCE_PRIORITY = {
  'Meta Densu July': 1,
  'Meta Densu': 2
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