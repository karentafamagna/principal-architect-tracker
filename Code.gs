const SHEETS = [
  "Dashboard_Metrics",
  "Trainees",
  "Trainee_Timeline",
  "Prompt_Portfolio",
  "Prompt_Usage",
  "AI_Timeline",
  "Architecture_Assets",
  "Evidence_Log",
  "Goals"
];

function doGet(e) {
  const out = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  SHEETS.forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) return;

    const rows = sh.getDataRange().getValues();
    const head = rows.shift();

    out[name] = rows
      .filter(r => r.join("") !== "")
      .map(r => {
        const o = {};
        head.forEach((h, i) => o[h] = r[i]);
        return o;
      });
  });

  return ContentService
    .createTextOutput(JSON.stringify({ data: out }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents); // { action, sheet, row }
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(body.sheet);

  if (!sh) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: "no sheet" }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (body.action === "add" || body.action === "append") {
    const head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const newRow = head.map(h => body.row[h] !== undefined ? body.row[h] : "");
    sh.appendRow(newRow);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: false, error: "unknown action" }))
    .setMimeType(ContentService.MimeType.JSON);
}
