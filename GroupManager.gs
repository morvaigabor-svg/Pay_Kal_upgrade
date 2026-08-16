/**
 * Megnyitja az űrlapot (Megnövelt magassággal: 465px)
 */
/**
 * Megnyitja az űrlapot (Szélesség megnövelve 50%-kal: 780px)
 */
/**
 * Megnyitja az űrlapot (Méret a hatalmas FolderPickerhez igazítva)
 */
function openNewGroupModal() {
  const template = HtmlService.createTemplateFromFile('NewGroupDialog');
  template.cities = getCitiesFromSheet();
  
  const html = template.evaluate()
    .setWidth(1150) // 👈 Keret szélessége
    .setHeight(700) // 👈 Keret magassága
    .setTitle('Új Csoport És Mappa Létrehozása');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Új Csoport Regisztrációja');
}

/**
 * Kiolvassa a városokat a 'Városok' munkalap A2:A tartományából
 */
function getCitiesFromSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Városok");
  if (!sheet) return [];
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  
  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  return values.map(row => row[0]).filter(Boolean);
}

/**
 * Átadja a biztonsági tokent a Google Picker API-nak
 */
function getOAuthToken() {
  DriveApp.getRootFolder(); // Kikényszeríti a Drive scope engedélyt
  return ScriptApp.getOAuthToken();
}

/**
 * Új csoport és mappa létrehozása
 */
function createGroupAndFolder(data) {
  try {
    let parentFolder;
    if (data.parentFolderId && data.parentFolderId !== 'root') {
      parentFolder = DriveApp.getFolderById(data.parentFolderId);
    } else {
      parentFolder = DriveApp.getRootFolder();
    }

    const newFolder = parentFolder.createFolder(data.folderName);
    const newFolderId = newFolder.getId();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName("Csoportok");
    
    if (!sheet) {
      sheet = ss.insertSheet("Csoportok");
      sheet.appendRow(["Csoport_ID", "Város", "Mappa_ID", "Létrehozva"]);
    }

    sheet.appendRow([
      data.csoportId,
      data.varos,
      newFolderId,
      new Date()
    ]);

    return { success: true, folderUrl: newFolder.getUrl() };

  } catch (error) {
    return { success: false, message: error.toString() };
  }
}