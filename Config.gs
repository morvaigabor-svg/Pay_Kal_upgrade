/**
 * @file Config.gs
 * @description Az alkalmazás globális konfigurációs beállításai.
 */

const APP = Object.freeze({
  NAME: "PayKal - Wallet of Calasanz",
  VERSION: "0.2.0",
  TIMEZONE: "Europe/Budapest",
  SHEETS: Object.freeze({
    USERS: "Felhasználók",
    GROUPS: "Csoportok",
    MEMBERS: "KM Tagok", // Frissítve a pontos munkalap névre!
    CITIES: "Városok",
    SETTINGS: "Beállítások",
    BALANCE: "Egyenleg",
    EXPENSES: "Költségek",
    INCOME: "Bevételek",
    TRANSFERS: "Pénzmozgások",
    SUMMARY: "Összesítő",
    LOG: "Log"
  })
});

/*EXCEL AZONOSÍTÓJA*/
const CONFIG = {
  SHEET_ID: "1Jhn7WKONe5UmzUdN4Vnd9lNcJrlxq6IK7fPMrJWh3Nc"
};

/* Egyetlen központi menü az összes adminisztrációs funkcióhoz*/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('PayKal Admin')
    .addItem('➕ Új csoport létrehozása', 'openNewGroupModal')
    .addSeparator()
    .addItem('📊 Kimutatások generálása', 'openReportGeneratorDialog')
    .addSeparator()
    .addItem('Blokkok frissítése', 'frissitsBlokkLinkeket')
    .addToUi();
}

function frissitsBlokkLinkeket() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var forrasLap = ss.getSheetByName('Költségek');
  var celLap = ss.getSheetByName('Számolótábla');
  
  if (!forrasLap || !celLap) return;

  // 1. Költségek blokk-linkjeinek beolvasása (K oszlop = ID [11], F oszlop = Blokk linkek [6])
  var forrasUtolsoSor = forrasLap.getLastRow();
  var szotar = {};
  
  if (forrasUtolsoSor >= 2) {
    var forrasIdek = forrasLap.getRange(2, 11, forrasUtolsoSor - 1, 1).getValues();
    var forrasRichTextek = forrasLap.getRange(2, 6, forrasUtolsoSor - 1, 1).getRichTextValues();
    
    for (var i = 0; i < forrasIdek.length; i++) {
      var id = forrasIdek[i][0];
      if (id) {
        szotar[id] = forrasRichTextek[i][0];
      }
    }
  }
  
  // 2. Számolótábla feldolgozása (B oszlop = ID [2])
  var celUtolsoSor = celLap.getLastRow();
  if (celUtolsoSor < 2) return;
  
  var celIdek = celLap.getRange(2, 2, celUtolsoSor - 1, 1).getValues();
  var kimenet = [];
  var uresRichText = SpreadsheetApp.newRichTextValue().setText("").build();
  
  for (var j = 0; j < celIdek.length; j++) {
    var celId = celIdek[j][0];
    
    // Ha az azonosító megtalálható a Költségek között, áthozza a linkeket
    if (celId && szotar[celId]) {
      kimenet.push([szotar[celId]]);
    } else {
      // Ha ez Bevétel / Átvezetés (vagy üres sor), simán üresen hagyja a cellát és folytatja
      kimenet.push([uresRichText]);
    }
  }
  
  // 3. Beírás a Számolótábla L oszlopába (12. oszlop)
  // (Ha más oszlopba szeretnéd tenni a linkeket, a 12-est írd át)
  var celOszlopIndex = 12; 
  celLap.getRange(2, celOszlopIndex, kimenet.length, 1).setRichTextValues(kimenet);
}