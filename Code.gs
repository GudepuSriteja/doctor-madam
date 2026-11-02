/** Run setup() once, then deploy this project as a web app. */
function setup() {
  const props = PropertiesService.getScriptProperties();
  let id = props.getProperty('RESPONSE_SHEET_ID');
  let ss;
  if (id) { ss = SpreadsheetApp.openById(id); }
  else {
    ss = SpreadsheetApp.create('Doctor Madam - Form Responses');
    id = ss.getId();
    props.setProperty('RESPONSE_SHEET_ID', id);
    const sheet = ss.getSheets()[0];
    sheet.setName('Responses');
    sheet.appendRow(['Submitted at','How she is feeling','Food','Meet today','Preferred time','Place suggestion','Stay longer']);
    sheet.setFrozenRows(1);
  }
  Logger.log('Response spreadsheet: ' + ss.getUrl());
  return ss.getUrl();
}

function doPost(e) {
  try {
    const raw = e && e.parameter && e.parameter.payload;
    if (!raw) throw new Error('Missing payload');
    const data = JSON.parse(raw);
    const id = PropertiesService.getScriptProperties().getProperty('RESPONSE_SHEET_ID');
    if (!id) throw new Error('Run setup() first');
    const sheet = SpreadsheetApp.openById(id).getSheetByName('Responses');
    sheet.appendRow([new Date(), data.mood || '', data.food || '', data.invite || '', data.time || '', data.place || '', data.escape || '']);
    return ContentService.createTextOutput('OK');
  } catch (err) {
    return ContentService.createTextOutput('ERROR: ' + err.message);
  }
}