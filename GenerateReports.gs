/**
 * GenerateReports.gs
 * Automatizált olvasási kimutatások, Drive parancsikonok és blokk-linkek generálása
 */

const REPORT_SYSTEM_CONFIG = {
  // IDE MÁSOLD A FŐ GYŰJTŐMAPPA ID-JÁT, AHOL A SHEETS FÁJLOK LÉTREJÖNNEK:
  MASTER_REPORTS_FOLDER_ID: "19yARIVLE08f8ZO1ZYxsdqhZyfdC27s0Q"
};

/**
 * Megnyitja a megemelt méretű párbeszédablakot a Picker miatt
 */
function openReportGeneratorDialog() {
  const html = HtmlService.createTemplateFromFile('ReportDialog')
    .evaluate()
    .setWidth(1150)
    .setHeight(720)
    .setTitle('Olvasási Kimutatások Generálása');
    
  SpreadsheetApp.getUi().showModalDialog(html, 'Olvasási Kimutatások Generálása');
}

/**
 * Biztonsági OAuth Token átadása a Google Picker API-nak
 */
function getOAuthToken() {
  DriveApp.getRootFolder(); // Kikényszeríti a Drive jogosultságot
  return ScriptApp.getOAuthToken();
}

/**
 * Adatok beolvasása az aktív táblázatból (Csoportok és Városok)
 */
function getReportOptionsData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. Csoportok beolvasása ('Csoportok' lap A2:A)
  const csSheet = ss.getSheetByName('Csoportok') || ss.getSheetByName('csoportok');
  const csoportok = [];
  if (csSheet && csSheet.getLastRow() >= 2) {
    const csVals = csSheet.getRange(2, 1, csSheet.getLastRow() - 1, 1).getValues();
    csVals.forEach(row => { if (row[0]) csoportok.push(String(row[0]).trim()); });
  }

  // 2. Városok beolvasása ('Városok' lap A2:A)
  const vSheet = ss.getSheetByName('Városok') || ss.getSheetByName('városok');
  const varosok = [];
  if (vSheet && vSheet.getLastRow() >= 2) {
    const vVals = vSheet.getRange(2, 1, vSheet.getLastRow() - 1, 1).getValues();
    vVals.forEach(row => { if (row[0]) varosok.push(String(row[0]).trim()); });
  }

  return { csoportok: csoportok, varosok: varosok };
}

/**
 * Fő táblázat Számolótábla blokk-linkjeinek frissítése (L oszlopba = 12. oszlop)
 */
function frissitsBlokkLinkeket() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const forrasLap = ss.getSheetByName('Költségek');
  const celLap = ss.getSheetByName('Számolótábla');
  
  if (!forrasLap || !celLap) return;

  const forrasUtolsoSor = forrasLap.getLastRow();
  const szotar = {};
  
  if (forrasUtolsoSor >= 2) {
    const forrasIdek = forrasLap.getRange(2, 11, forrasUtolsoSor - 1, 1).getValues(); // K oszlop = Költési azonosító (11)
    const forrasRichTextek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getRichTextValues(); // F oszlop = Blokk linkek (6)
    
    for (let i = 0; i < forrasIdek.length; i++) {
      const id = String(forrasIdek[i][0]).trim();
      if (id) {
        szotar[id] = forrasRichTextek[i][0];
      }
    }
  }
  
  const celUtolsoSor = celLap.getLastRow();
  if (celUtolsoSor < 2) return;
  
  const celIdek = celLap.getRange(2, 2, celUtolsoSor - 1, 1).getValues(); // B oszlop = Költési azonosító (2)
  const kimenet = [];
  const uresRichText = SpreadsheetApp.newRichTextValue().setText("").build();
  
  for (let j = 0; j < celIdek.length; j++) {
    const celId = String(celIdek[j][0]).trim();
    if (celId && szotar[celId]) {
      kimenet.push([szotar[celId]]);
    } else {
      kimenet.push([uresRichText]);
    }
  }
  
  // 12. oszlop = L oszlop!
  celLap.getRange(2, 12, kimenet.length, 1).setRichTextValues(kimenet);
}

/**
 * Kimutatások, parancsikonok ÉS blokk-linkek generálása
 */
function processReportGeneration(payload) {
  try {
    // 1. ELŐ-LÉPÉS: Blokk-linkek frissítése a Fő Táblázat L oszlopában
    frissitsBlokkLinkeket();

    const masterSS = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheetId = masterSS.getId();
    const masterFolder = DriveApp.getFolderById(REPORT_SYSTEM_CONFIG.MASTER_REPORTS_FOLDER_ID);
    
    // 2. Költségek blokk-linkjeinek beolvasása szótárba (ID -> RichText)
    const koltsegekLap = masterSS.getSheetByName('Költségek');
    const blokkSzotar = {};
    if (koltsegekLap && koltsegekLap.getLastRow() >= 2) {
      const kIdek = koltsegekLap.getRange(2, 11, koltsegekLap.getLastRow() - 1, 1).getValues();
      const kRich = koltsegekLap.getRange(2, 6, koltsegekLap.getLastRow() - 1, 1).getRichTextValues();
      for (let i = 0; i < kIdek.length; i++) {
        const id = String(kIdek[i][0]).trim();
        if (id) blokkSzotar[id] = kRich[i][0];
      }
    }

    // 3. Számolótábla adatainak beolvasása a sorrend tartásához (Col1..Col10)
    const szamoloLap = masterSS.getSheetByName('Számolótábla');
    let szamoloAdatok = [];
    if (szamoloLap && szamoloLap.getLastRow() >= 2) {
      szamoloAdatok = szamoloLap.getRange(2, 1, szamoloLap.getLastRow() - 1, 10).getValues();
    }

    const results = [];
    const uresRichText = SpreadsheetApp.newRichTextValue().setText("").build();

    // 4. Riportok generálása ciklusban
    payload.items.forEach(item => {
      try {
        const isGroup = (payload.type === 'GROUP');
        const titleName = isGroup ? `Kimutatás - ${item.id}` : `Városi Kimutatás - ${item.id}`;
        
        // A) Új Google Sheet létrehozása
        const newSS = SpreadsheetApp.create(titleName);
        const newFile = DriveApp.getFileById(newSS.getId());
        newFile.moveTo(masterFolder);
        newFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

        // B) Munkalap és Fejlécek (12 oszlop: A-L)
        const sheet = newSS.getActiveSheet();
        sheet.setName("Számolótábla");

        const headers = [[
          "Csoport_ID", "Költési azonosító", "Dátum", "Tranzakció célja", 
          "Tranzakció típusa", "Fizetési mód", "Egyenleg változás (Ft)", 
          "Készpénz változás", "Számla változás", "Város", "Egyéb / K oszlop", "Blokkok"
        ]];
        sheet.getRange(1, 1, 1, 12).setValues(headers).setFontWeight("bold").setBackground("#eef2f7");

        // C) IMPORTRANGE + QUERY képlet beírása (A2:K tartományra)
        const filterCol = isGroup ? "Col1" : "Col10";
        const formulaLocal = `=QUERY(IMPORTRANGE("${masterSheetId}"; "Számolótábla!A2:K"); "SELECT * WHERE ${filterCol} = '${item.id}'"; 0)`;
        sheet.getRange("A2").setFormulaLocal(formulaLocal);

        // D) Blokk-linkek másolása pontosan az L oszlopba (12. oszlop)
        const riportBlokkok = [];
        const itemIdClean = String(item.id).trim();

        szamoloAdatok.forEach(row => {
          // Csoport esetén Col1 (index 0), Város esetén Col10 (index 9)
          const matchVal = isGroup ? String(row[0]).trim() : String(row[9]).trim();
          if (matchVal === itemIdClean) {
            const tranzakcioId = String(row[1]).trim(); // Col2 = Költési azonosító
            if (tranzakcioId && blokkSzotar[tranzakcioId]) {
              riportBlokkok.push([blokkSzotar[tranzakcioId]]);
            } else {
              riportBlokkok.push([uresRichText]);
            }
          }
        });

        if (riportBlokkok.length > 0) {
          sheet.getRange(2, 12, riportBlokkok.length, 1).setRichTextValues(riportBlokkok); // 12. oszlop (L oszlop)
        }

        // E) Drive Parancsikon létrehozása
        if (item.targetFolderId && item.targetFolderId !== 'root') {
          const targetFolder = DriveApp.getFolderById(item.targetFolderId);
          targetFolder.createShortcut(newSS.getId());
        }

        results.push({ name: item.id, status: 'OK' });
      } catch (err) {
        results.push({ name: item.id, status: 'HIBA: ' + err.message });
      }
    });

    return { success: true, results: results };
  } catch (globalErr) {
    return { success: false, message: globalErr.toString() };
  }
}