/**
 * Service.Settings.gs
 * Beállítások, törzsadatok és dinamikus opciók beolvasása
 */

function getSettingsData() {
  const auth = getUserAuth();
  if (!auth.authorized) throw new Error("Jogosulatlan hozzáférés!");

  const csoportId = String(auth.csoportId || "").trim();
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  let costCentersBase = [];
  let paymentMethods = [];
  let incomePurposesBase = [];
  let incomePaymentMethods = [];
  let settingF = [];
  let projectTypes = [];

  // 1. Beállítások munkalap beolvasása (A-G oszlopok)
  const settingsSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.SETTINGS) ? APP.SHEETS.SETTINGS : "Beállítások";
  const settingsSheet = ss.getSheetByName(settingsSheetName);

  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const lastRow = settingsSheet.getLastRow();
    const data = settingsSheet.getRange(2, 1, lastRow - 1, 7).getValues();

    data.forEach(row => {
      if (row[0]) costCentersBase.push(String(row[0]).trim());       // A2:A - Kiadási jogcímek
      if (row[1]) paymentMethods.push(String(row[1]).trim());        // B2:B - Kiadási fizetési módok
      if (row[2]) incomePaymentMethods.push(String(row[2]).trim());  // C2:C - Bevételei fizetési módok
      if (row[3]) incomePurposesBase.push(String(row[3]).trim());    // D2:D - Bevételei jogcímek
      if (row[5]) settingF.push(String(row[5]).trim());              // F2:F - Egyéb beállítások

      // E2:E (Index 4: Csoportvezető) vs G2:G (Index 6: Koordinátor)
      const pType = auth.isCoordinator ? row[6] : row[4];
      if (pType) projectTypes.push(String(pType).trim());
    });
  }

  // 2. Tagok beolvasása (A oszlop: Név | B oszlop: Csoport ID)
  let payers = [];
  const membersSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.MEMBERS) ? APP.SHEETS.MEMBERS : "Tagok";
  const membersSheet = ss.getSheetByName(membersSheetName);

  if (membersSheet && membersSheet.getLastRow() >= 2) {
    const mData = membersSheet.getRange(2, 1, membersSheet.getLastRow() - 1, 2).getValues();

    mData.forEach(row => {
      const memberName = row[0] ? String(row[0]).trim() : "";
      const groupVal = row[1] ? String(row[1]).trim() : "";

      if (memberName !== "" && groupVal.toLowerCase() === csoportId.toLowerCase()) {
        payers.push(memberName);
      }
    });
  }

  // Koordinátorok esetén a "KM Egyesület" automatikus hozzáadása
  if (auth.isCoordinator) {
    payers.unshift("KM Egyesület");
  }

  // 3. Aktív projektek kinyerése
  const activeProjects = getActiveProjectsData(csoportId, ss);
  const activeProjectNames = activeProjects.map(p => p.name);

  // Kiadások jogcímeinek összekészítése (Projekt nevek beszúrása)
  const costCenters = [
    ...costCentersBase.filter(item => item.toLowerCase() !== "egyéb" && item.toLowerCase() !== "egyéb kiadások"),
    ...activeProjectNames,
    "Egyéb kiadások"
  ];

  // Bevételek jogcímeinek összekészítése (Projekt nevek beszúrása)
  const incomePurposes = [
    ...incomePurposesBase.filter(item => item.toLowerCase() !== "egyéb" && item.toLowerCase() !== "egyéb bevételek"),
    ...activeProjectNames,
    "Egyéb bevételek"
  ];

  return {
    costCenters: [...new Set(costCenters)],
    paymentMethods: [...new Set(paymentMethods)],
    incomePurposes: [...new Set(incomePurposes)],
    incomePaymentMethods: [...new Set(incomePaymentMethods)],
    payers: [...new Set(payers)],
    settingF: [...new Set(settingF)],
    projectTypes: [...new Set(projectTypes)],
    activeProjects: activeProjects
  };
}