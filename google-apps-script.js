function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "gallery") {
    return ContentService
      .createTextOutput(JSON.stringify({ items: getOfficeGallery_(ss) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "results") {
    return ContentService
      .createTextOutput(JSON.stringify({ items: getSmileResults_(ss) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      message: "Use ?action=gallery or ?action=results"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : "";
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (action === "gallery") {
    return ContentService
      .createTextOutput(JSON.stringify({ items: getOfficeGallery_(ss) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  if (action === "results") {
    return ContentService
      .createTextOutput(JSON.stringify({ items: getSmileResults_(ss) }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  return ContentService
    .createTextOutput(JSON.stringify({
      ok: true,
      message: "Use ?action=gallery or ?action=results"
    }))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOfficeGallery_(ss) {
  return getSheetRows_(ss, "office_gallery");
}

function getSmileResults_(ss) {
  return getSheetRows_(ss, "smile_results");
}

function getSheetRows_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const values = sheet.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  return mapRows_(values);
}

function mapRows_(values) {
  const headers = values[0];
  return values
    .slice(1)
    .filter(row => row.join("") !== "")
    .map(row => {
      const obj = {};
      headers.forEach((h, i) => obj[h] = row[i]);
      return obj;
    });
}

function getSmileResults_(ss) {
  const sheet = ss.getSheetByName("smile_results");
  const values = sheet.getDataRange().getValues();
  return mapRows_(values);
}

function mapRows_(values) {
  const headers = values[0];
  return values.slice(1).filter(r => r.join("") !== "").map(row => {
    const obj = {};
    headers.forEach((h, i) => obj[h] = row[i]);
    return obj;
  });
}
