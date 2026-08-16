/**
 * Kikeresi az aktív csoporthoz tartozó Google Drive mappa ID-ját
 */
function getExpenseFolder() {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
  const groupsSheet = ss.getSheetByName(APP.SHEETS.GROUPS);
  
  if (groupsSheet && groupsSheet.getLastRow() >= 2) {
    const data = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, 3).getValues();
    // A: Csoport_ID, B: Város, C: Mappa_ID
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][0]).trim() === String(auth.csoportId).trim() && data[i][2]) {
        return DriveApp.getFolderById(data[i][2]);
      }
    }
  }

  throw new Error("Nem található kijelölt Drive mappa a megadott csoportnak: " + auth.csoportId);
}

function uploadExpenseImage(imageData, expenseId, imageNumber) {
  const folder = getExpenseFolder();
  
  const contentType = imageData.match(/data:(.*);base64/)[1];
  const base64 = imageData.split(",")[1];
  const bytes = Utilities.base64Decode(base64);
  const extension = contentType.split("/")[1] || "jpg";

  const fileName = expenseId + "_" + imageNumber + "." + extension;
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = folder.createFile(blob);

  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  return {
    id: file.getId(),
    url: file.getUrl(),
    name: file.getName()
  };
}