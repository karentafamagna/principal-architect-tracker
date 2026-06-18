# Principal Architect Tracker — Backend Setup

The dashboard works immediately with built-in sample data (demo mode). To make it read live data **and save the data-entry forms back to your Google Sheet**, point it at an Apps Script Web App.

## 1. Paste your Web App URL

In `index_live.html`, line ~ `const API_URL = "..."` — replace the placeholder with your deployed Apps Script URL.

## 2. Apps Script the dashboard expects

The dashboard makes two calls:

- **Read:**  `GET  {API_URL}?action=getAll` → returns `{ data: { SheetName: [ {col:val,...}, ... ] } }`
- **Write:** `POST {API_URL}` with JSON body `{ action:"add", sheet:"Trainees", row:{...} }`

Paste this into your Apps Script project (Extensions → Apps Script) and re-deploy as a Web App (Execute as: Me, Access: Anyone):

```javascript
const SHEETS = ["Dashboard_Metrics","Trainees","Prompt_Portfolio","Architecture_Assets","Evidence_Log","Goals"];

function doGet(e){
  const out = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SHEETS.forEach(name=>{
    const sh = ss.getSheetByName(name);
    if(!sh) return;
    const rows = sh.getDataRange().getValues();
    const head = rows.shift();
    out[name] = rows.filter(r=>r.join("")!=="").map(r=>{
      const o={}; head.forEach((h,i)=>o[h]=r[i]); return o;
    });
  });
  return ContentService.createTextOutput(JSON.stringify({data:out}))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e){
  const body = JSON.parse(e.postData.contents);   // {action,sheet,row}
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(body.sheet);
  if(!sh) return ContentService.createTextOutput(JSON.stringify({ok:false,error:"no sheet"}))
    .setMimeType(ContentService.MimeType.JSON);

  const head = sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0];
  const newRow = head.map(h => body.row[h] !== undefined ? body.row[h] : "");
  sh.appendRow(newRow);

  return ContentService.createTextOutput(JSON.stringify({ok:true}))
    .setMimeType(ContentService.MimeType.JSON);
}
```

## 3. Column headers the forms write

Make sure row 1 of each tab uses these header names. Extra columns are fine.

- **Trainees:** Name, Program, Status, Current Utilization, Revenue, Active Projects
- **Prompt_Portfolio:** Prompt Name, Platform, Status, Est Hours Saved / Use, Notes
- **Architecture_Assets:** Asset Name, Category, Status, Projects Using, Dollar Value
- **Evidence_Log:** Date, Category, Evidence Title, Metric Impact
- **Goals:** Goal Name, Goal Type, Current, Target, Progress

## Notes

- The POST is sent with `Content-Type: text/plain` on purpose — this avoids a CORS preflight that Apps Script doesn't answer.
- Until the URL is set, every form still works in **demo mode**. Entries appear in the view but aren't persisted.
- The toast confirms each save. On a sync failure the entry stays visible locally and you'll see a warning.
