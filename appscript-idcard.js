// ── بطاقات الهوية – معسكر المعيصم 1447 ──────────────────────────────────────
// انشر هذا الكود في Google Apps Script كـ Web App (تنفيذ: أنا، وصول: الجميع)

var SHEET_NAME = 'id_cards';

function getSheet() {
  var prop = PropertiesService.getScriptProperties().getProperty('IDCARD_SHEET_ID');
  var ss;
  if (prop) {
    try { ss = SpreadsheetApp.openById(prop); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create('بطاقات الهوية - معسكر المعيصم 1447');
    PropertiesService.getScriptProperties().setProperty('IDCARD_SHEET_ID', ss.getId());
  }
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['رقم الهوية', 'الاسم الكامل', 'القائد', 'الجناح', 'تاريخ التسجيل']);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    if (data.action === 'register') {
      var id     = String(data.id     || '').trim();
      var name   = String(data.name   || '').trim();
      var leader = String(data.leader || '').trim();
      var wing   = String(data.wing   || '').trim();
      if (!id || !name) { return out({ok: false, msg: 'بيانات ناقصة'}); }

      var sheet = getSheet();
      var rows  = sheet.getDataRange().getValues();
      var existingRow = -1;
      rows.forEach(function(row, i) {
        if (i > 0 && String(row[0]).trim() === id) { existingRow = i + 1; }
      });

      var ts = Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm');
      if (existingRow > 0) {
        sheet.getRange(existingRow, 1, 1, 5).setValues([[id, name, leader, wing, ts]]);
      } else {
        sheet.appendRow([id, name, leader, wing, ts]);
      }
      return out({ok: true});
    }
    return out({ok: false, msg: 'action غير معروف'});
  } catch (err) {
    return out({ok: false, msg: err.toString()});
  }
}

function doGet(e) {
  var action = e.parameter.action;

  // التحقق من هوية فرد واحد بالرقم
  if (action === 'verify') {
    var id    = String(e.parameter.id || '').trim();
    var sheet = getSheet();
    var rows  = sheet.getDataRange().getValues();
    var found = null;
    rows.forEach(function(row, i) {
      if (i > 0 && String(row[0]).trim() === id) {
        found = { id: row[0], name: row[1], leader: row[2], wing: row[3], ts: row[4] };
      }
    });
    if (found) { return out({ok: true, data: found}); }
    return out({ok: false, msg: 'غير مسجّل'});
  }

  // استعراض كل المسجلين (للمشرف)
  if (action === 'getAll') {
    var sheet  = getSheet();
    var rows   = sheet.getDataRange().getValues();
    var result = [];
    rows.forEach(function(row, i) {
      if (i > 0 && row[0]) {
        result.push({ id: row[0], name: row[1], leader: row[2], wing: row[3], ts: row[4] });
      }
    });
    return out({ok: true, data: result});
  }

  return out({ok: false, msg: 'action غير معروف'});
}
