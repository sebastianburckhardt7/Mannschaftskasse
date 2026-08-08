// =============================================
// MANNSCHAFTSKASSE – Google Apps Script Backend
// Einmalig in Google Sheets einfügen:
// Erweiterungen → Apps Script → Code.gs ersetzen
// Danach: Bereitstellen → Als Web-App bereitstellen
// Zugriff: Jeder (auch anonym)
// =============================================

const SHEET_ID = "1rcrEYpv0ln_NJG87ZBaidVqJWq8K_POVRa6XI6kkuTU";
const SHEET_TRANSACTIONS = "Transaktionen";
const SHEET_MEMBERS = "Mitglieder";

function doGet(e) {
  const action = e.parameter.action;

  if (action === "getTransactions") {
    return getTransactions();
  } else if (action === "getMembers") {
    return getMembers();
  } else if (action === "getSummary") {
    return getSummary();
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "Unknown action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const action = data.action;

  if (action === "addTransaction") {
    return addTransaction(data);
  } else if (action === "addMember") {
    return addMember(data);
  } else if (action === "deleteMember") {
    return deleteMember(data);
  }

  return ContentService.createTextOutput(JSON.stringify({ error: "Unknown action" }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === SHEET_TRANSACTIONS) {
      sheet.appendRow(["ID", "Datum", "Typ", "Mitglied", "Beschreibung", "Betrag", "Erfasst von"]);
    } else if (name === SHEET_MEMBERS) {
      sheet.appendRow(["ID", "Name", "Position", "Aktiv"]);
    }
  }
  return sheet;
}

function getTransactions() {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify({ success: true, data: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getMembers() {
  const sheet = getSheet(SHEET_MEMBERS);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
  return ContentService.createTextOutput(JSON.stringify({ success: true, data: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function addTransaction(data) {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const id = Utilities.getUuid();
  const now = new Date().toISOString();
  sheet.appendRow([
    id,
    data.datum || now,
    data.typ,       // "Strafe" | "Einzahlung" | "Ausgabe"
    data.mitglied,
    data.beschreibung,
    parseFloat(data.betrag),
    data.erfasstVon
  ]);
  return ContentService.createTextOutput(JSON.stringify({ success: true, id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function addMember(data) {
  const sheet = getSheet(SHEET_MEMBERS);
  const id = Utilities.getUuid();
  sheet.appendRow([id, data.name, data.position || "Spieler", true]);
  return ContentService.createTextOutput(JSON.stringify({ success: true, id }))
    .setMimeType(ContentService.MimeType.JSON);
}

function deleteMember(data) {
  const sheet = getSheet(SHEET_MEMBERS);
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === data.id) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
  return ContentService.createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSummary() {
  const sheet = getSheet(SHEET_TRANSACTIONS);
  const data = sheet.getDataRange().getValues().slice(1);
  let kassenstand = 0;
  let totalStrafen = 0;
  let totalEinzahlungen = 0;
  let totalAusgaben = 0;

  data.forEach(row => {
    const typ = row[2];
    const betrag = parseFloat(row[5]) || 0;
    if (typ === "Strafe") { totalStrafen += betrag; kassenstand += betrag; }
    else if (typ === "Einzahlung") { totalEinzahlungen += betrag; kassenstand += betrag; }
    else if (typ === "Ausgabe") { totalAusgaben += betrag; kassenstand -= betrag; }
  });

  return ContentService.createTextOutput(JSON.stringify({
    success: true,
    data: { kassenstand, totalStrafen, totalEinzahlungen, totalAusgaben }
  })).setMimeType(ContentService.MimeType.JSON);
}
