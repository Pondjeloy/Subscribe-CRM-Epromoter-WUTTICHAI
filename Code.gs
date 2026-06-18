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
// ════════════════════════════════════════════════════════

var SPREADSHEET_ID = '13rWPSzpyfoEFVyMLgcIHRG9f_ODpo26mYQdvKsEJWns';
var PROMOTER       = 'POND';

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

    'Meta Densu': {
      picCol:13, statusCol:12, notesCol:14,
      parse: function(row, disp) {
        // H(7)=age/phone-fallback  I(8)=phone-or-email  J(9)=email-fallback
        var hDisp = cleanDisplay(row[7], disp&&disp[7]);
        var iVal  = clean(row[8]);
        var jVal  = clean(row[9]);

        // Email: first of I, J that contains @
        var email = '';
        if (iVal.indexOf('@') !== -1)      email = iVal;
        else if (jVal.indexOf('@') !== -1) email = jVal;

        // Phone: I first (if ≥9 digits, no @), else try H
        function isPhone(v) {
          return v.replace(/\D/g,'').length >= 9 && v.indexOf('@') === -1;
        }
        var phone = isPhone(iVal) ? iVal : (isPhone(hDisp) ? hDisp : '');

        // Age: H(7) — calcAge returns '' if H turns out to be phone
        var ageRaw = calcAge(row[7]);
        var age = (String(ageRaw).replace(/\D/g,'').length >= 9) ? '' : ageRaw;

        return {
          name:           clean(row[5]),      // F
          phone:          phone,
          email:          email,
          age:            age,
          gender:         clean(row[6]),      // G
          paymentChannel: clean(row[3]),
          province:       clean(row[0]),
          productType:    clean(row[1]),
          lineId:         ''
        };
      }
    },

    'Meta Credit': {
      picCol:12, statusCol:11, notesCol:13,
      parse: function(row, disp) { return {
        name:           clean(row[5]),                          // F
        phone:          cleanDisplay(row[7], disp&&disp[7]),   // H
        email:          clean(row[8]),                          // I
        age:            calcAge(row[6]),                        // G
        paymentChannel: clean(row[3]),
        province:       clean(row[0]),
        productType:    clean(row[1]),
        lineId:         ''
      };}
    },

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

    // ── [แก้ไข] Status/PIC/Remark เลื่อนซ้าย 1 คอลัมน์ ──
    // เดิม picCol:9 statusCol:8 notesCol:10 (J/I/K)
    // ตอนนี้ตามภาพชีตจริง: Status=H(7) PIC=I(8) Remark=J(9)
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
  var sheetNames = [
    'Meta Densu','Meta Credit','Lead Subscribe Lg.com',
    'Lead LG Success','Lead Consult','Lead Subscribe POP UP Braner'
  ];
  var all = [], idNum = 1;

  for (var s = 0; s < sheetNames.length; s++) {
    var sName = sheetNames[s];
    var cfg   = getSheetConfig(sName);
    if (!cfg) continue;
    var sheet = ss.getSheetByName(sName);
    if (!sheet) { Logger.log('Not found: '+sName); continue; }

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
    }
  }
  return { success:true, count:all.length, data:all };
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
  var sheetNames = [
    'Meta Densu','Meta Credit','Lead Subscribe Lg.com',
    'Lead LG Success','Lead Consult','Lead Subscribe POP UP Braner'
  ];
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
  ['Meta Densu','Meta Credit','Lead Subscribe Lg.com',
   'Lead LG Success','Lead Consult','Lead Subscribe POP UP Braner'].forEach(function(n) {
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