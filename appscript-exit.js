var ADMIN_PASS = 'admin1447';

function getSheets() {
  var props = PropertiesService.getScriptProperties();
  var sid = props.getProperty('EXIT_SHEET_ID');
  if (!sid) {
    var ss = SpreadsheetApp.create('استئذان الخروج - معسكر المعيصم 1447');
    sid = ss.getId();
    props.setProperty('EXIT_SHEET_ID', sid);
    var ls = ss.getActiveSheet();
    ls.setName('leaders');
    ls.appendRow(['id','key','name']);
    ss.insertSheet('requests').appendRow([
      'id','name','leaderId','leaderName',
      'reason','returnDate','returnTime',
      'status','timestamp','note'
    ]);
  }
  var ss = SpreadsheetApp.openById(sid);
  return {
    leaders:  ss.getSheetByName('leaders'),
    requests: ss.getSheetByName('requests')
  };
}

function verifyLeader(sheets, leaderId, leaderKey) {
  var rows = sheets.leaders.getDataRange().getValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][0]) === String(leaderId) &&
        String(rows[i][1]) === String(leaderKey)) {
      return true;
    }
  }
  return false;
}

function doPost(e) {
  var sheets = getSheets();
  var d = JSON.parse(e.postData.contents);
  var action = d.action || '';

  if (action === 'addRequest') {
    sheets.requests.appendRow([
      d.id, d.name||'', d.leaderId||'', d.leaderName||'',
      d.reason||'', d.returnDate||'', d.returnTime||'',
      'pending', new Date().toISOString(), ''
    ]);
    return ContentService.createTextOutput('ok');
  }

  return ContentService.createTextOutput('error:unknown');
}

function doGet(e) {
  var action = e.parameter.action || '';
  var sheets = getSheets();

  if (action === 'getLeaders') {
    var rows = sheets.leaders.getDataRange().getValues();
    var list = [];
    rows.slice(1).forEach(function(r) { list.push({id:r[0], name:r[2]}); });
    return out(list);
  }

  if (action === 'getStatus') {
    var rid = e.parameter.id;
    var rows = sheets.requests.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(rid)) {
        return out({name:rows[i][1], leaderName:rows[i][3], reason:rows[i][4],
                    returnDate:rows[i][5], returnTime:rows[i][6],
                    status:rows[i][7], note:rows[i][9]});
      }
    }
    return out({error:'not found'});
  }

  if (action === 'getRequests') {
    var lid = e.parameter.leaderId;
    var lkey = e.parameter.key;
    if (!verifyLeader(sheets, lid, lkey)) return out({error:'unauthorized'});
    var rows = sheets.requests.getDataRange().getValues();
    var list = [];
    rows.slice(1).forEach(function(r) {
      if (String(r[2]) === String(lid)) {
        list.push({id:r[0], name:r[1], reason:r[4],
                   returnDate:r[5], returnTime:r[6],
                   status:r[7], timestamp:r[8], note:r[9]});
      }
    });
    list.reverse();
    return out(list);
  }

  if (action === 'updateStatus') {
    var lid = e.parameter.leaderId;
    var lkey = e.parameter.key;
    if (!verifyLeader(sheets, lid, lkey)) return out({error:'unauthorized'});
    var rid = e.parameter.requestId;
    var rows = sheets.requests.getDataRange().getValues();
    for (var j = 1; j < rows.length; j++) {
      if (String(rows[j][0]) === String(rid) && String(rows[j][2]) === String(lid)) {
        sheets.requests.getRange(j+1, 8).setValue(e.parameter.status);
        sheets.requests.getRange(j+1, 10).setValue(e.parameter.note || '');
        return out({status:'ok'});
      }
    }
    return out({error:'not found'});
  }

  var pass = e.parameter.pass || '';
  if (pass !== ADMIN_PASS) return out({error:'unauthorized'});

  if (action === 'getAllLeaders') {
    var rows = sheets.leaders.getDataRange().getValues();
    var list = [];
    rows.slice(1).forEach(function(r) { list.push({id:r[0], key:r[1], name:r[2]}); });
    return out(list);
  }

  if (action === 'addLeader') {
    var name = e.parameter.name;
    var id   = String(new Date().getTime());
    var key  = Math.random().toString(36).substr(2, 12);
    sheets.leaders.appendRow([id, key, name]);
    return out({status:'ok', id:id, key:key});
  }

  if (action === 'deleteLeader') {
    var lid  = e.parameter.id;
    var rows = sheets.leaders.getDataRange().getValues();
    for (var i = 1; i < rows.length; i++) {
      if (String(rows[i][0]) === String(lid)) {
        sheets.leaders.deleteRow(i + 1);
        return out({status:'ok'});
      }
    }
    return out({error:'not found'});
  }

  if (action === 'getAllRequests') {
    var rows = sheets.requests.getDataRange().getValues();
    var list = [];
    rows.slice(1).forEach(function(r) {
      list.push({id:r[0], name:r[1], leaderId:r[2], leaderName:r[3],
                 reason:r[4], returnDate:r[5], returnTime:r[6],
                 status:r[7], timestamp:r[8], note:r[9]});
    });
    list.reverse();
    return out(list);
  }

  return out({error:'unknown'});
}

function out(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
