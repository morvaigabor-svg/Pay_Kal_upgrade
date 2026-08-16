function writeLog(level, module, message) {
  try {
    const auth = getUserAuth();
    const csoportId = auth.csoportId || "ISMERETLEN";
    const email = auth.email || Session.getActiveUser().getEmail() || "ISMERETLEN";

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const sheet = ss.getSheetByName(APP.SHEETS.LOG);

    if (sheet) {
      sheet.appendRow([
        new Date(), // Idő
        level,      // Szint
        module,     // Modul
        csoportId,  // Csoport_ID
        email,      // Felhasználó Email
        message     // Üzenet
      ]);
    }
  } catch (e) {
    Logger.log("Log írási hiba: " + e.toString());
  }
}