// IV Report — Google Apps Script
// Deploy as Web App: Execute as Me, Access: Anyone
// Copy the deployment URL into the PWA Settings screen

const SHEET = "IV Report";

function doPost(e) {
  try {
    const m = JSON.parse(e.postData.contents);
    write(m);
    return ContentService.createTextOutput(JSON.stringify({status:"ok"})).setMimeType(ContentService.MimeType.JSON);
  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({status:"error",msg:err.message})).setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet() {
  return ContentService.createTextOutput(JSON.stringify({status:"ok",msg:"IV Report endpoint live"})).setMimeType(ContentService.MimeType.JSON);
}

function write(m) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET);
    const h = ["#","Type","Serial Number","Model","MVPS","DCB","String","Inverter",
      "Rated Pmax","T1 Perf%","T1 Pmax Meas","T1 Pmax Pred","T1 Pmax STC",
      "T2 Perf%","T2 Pmax Meas","T2 Pmax Pred","T2 Pmax STC",
      "Bifacial W","Bifacial %","Alerts","Assessment","Recorded"];
    sheet.getRange(1,1,1,h.length).setValues([h]).setBackground("#0d1b2a").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  const t1 = m.t1||{}, t2 = m.t2||{}, np = m.nameplate||{};
  const dW = (t1.pmaxSTC && t2.pmaxSTC) ? +(t1.pmaxSTC-t2.pmaxSTC).toFixed(2) : '';
  const dP = (t1.pmaxSTC && t2.pmaxSTC) ? +(dW/t1.pmaxSTC*100).toFixed(1) : '';
  const row = sheet.getLastRow()+1;
  sheet.appendRow([
    row-1, (m.type||'normal').toUpperCase(), m.serialNumber||'', m.model||'',
    m.mvps||'', m.dcb||'', m.string||'', m.inverter||'',
    np.pmax||'',
    t1.performance||'', t1.pmaxMeasured||'', t1.pmaxPredicted||'', t1.pmaxSTC||'',
    t2.performance!=null?t2.performance:'N/A',
    t2.pmaxMeasured!=null?t2.pmaxMeasured:'N/A',
    t2.pmaxPredicted!=null?t2.pmaxPredicted:'N/A',
    t2.pmaxSTC!=null?t2.pmaxSTC:'N/A',
    dW, dP,
    [t1.alerts,t2.alerts].filter(Boolean).join(' | ')||'None',
    m.assessment||'', new Date().toLocaleString()
  ]);
  const color = m.type==='damaged'?'#fce8e6':m.type==='spare'?'#e8f0fe':'#ffffff';
  sheet.getRange(row,1,1,22).setBackground(color);
}
