// Backend for the "Stay updated on baby news" form on tanazandvince.com/three.
//
// Setup (one time):
// 1. Create a Google Sheet, then Extensions → Apps Script, and paste this file in.
// 2. Deploy → New deployment → Web app. Execute as: Me. Who has access: Anyone.
// 3. Copy the web app URL into SIGNUPS_URL in three/index.html.
// Sign-ups appear as rows (name, timestamp, phone) in the sheet.

function doPost(e) {
  const data = JSON.parse(e.postData.contents);
  const name = String(data.name || '').trim().slice(0, 200);
  const phone = String(data.phone || '').trim().slice(0, 50);
  if (!name) return json({ ok: false });
  // Leading apostrophe keeps Sheets from treating the number as a formula or math
  SpreadsheetApp.getActiveSpreadsheet().getSheets()[0].appendRow([name, new Date(), phone ? "'" + phone : '']);
  return json({ ok: true });
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
