/**
 * Egyedi kiadás azonosító (ID) generálása
 * Formátum: ÉÉÉÉHHNN-Költséghely-Sorszám (pl. 20260723-Rendezvény-005)
 */
function generateExpenseId(costCenter) {
  const now = new Date();

  const date = Utilities.formatDate(
    now,
    APP.TIMEZONE,
    "yyyyMMdd"
  );

  const sheet = SpreadsheetApp
    .openById(CONFIG.SHEET_ID)
    .getSheetByName(APP.SHEETS.EXPENSES);

  const lastRow = sheet.getLastRow();

  const sequence = String(lastRow).padStart(3, "0");

  return date + "-" + costCenter + "-" + sequence;
}

function KERES_LINKKELER(keresettErtek, keresesiTartomany, eredmenyTartomany) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var searchValues = keresesiTartomany.getValues();
  var resultRichText = eredmenyTartomany.getRichTextValues();
  
  for (var i = 0; i < searchValues.length; i++) {
    if (searchValues[i][0] == keresettErtek) {
      var url = resultRichText[i][0].getLinkUrl();
      var text = resultRichText[i][0].getText();
      return url ? text + " (URL: " + url + ")" : text;
    }
  }
  return "Nincs találat";
}