/**
 * Web App belépési pont
 */
function doGet() {
  return HtmlService
    .createTemplateFromFile("index")
    .evaluate()
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=1.0, user-scalable=no');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Felhasználó azonosítása és pontos koordinátori jogosultságok felépítése
 */
function getUserAuth() {
  try {
    const activeEmail = Session.getActiveUser().getEmail();
    const scriptUrl = ScriptApp.getService().getUrl();
    const switchUrl = "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl);

    if (!activeEmail || activeEmail.trim() === "") {
      return { authorized: false, needsAuth: true, switchAccountUrl: switchUrl };
    }

    const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);
    const usersSheet = ss.getSheetByName(APP.SHEETS.USERS);
    
    let isAuthorized = false;
    let userGroup = null; // C oszlop: csoport_ID
    let userRole = null;  // D oszlop: SZEREPKOR (VEZETO / KOORDINATOR)
    let userName = null;  // B oszlop: Név

    if (usersSheet && usersSheet.getLastRow() >= 2) {
      const usersData = usersSheet.getRange(2, 1, usersSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < usersData.length; i++) {
        const sheetEmail = String(usersData[i][0]).trim().toLowerCase();
        if (sheetEmail === activeEmail.trim().toLowerCase()) {
          isAuthorized = true;
          userName = String(usersData[i][1] || "").trim();
          userGroup = String(usersData[i][2] || "").trim();
          userRole = String(usersData[i][3] || "").trim();
          break;
        }
      }
    }

    if (!isAuthorized) {
      return { authorized: false, needsAuth: false, email: activeEmail, switchAccountUrl: switchUrl };
    }

    const isCoordinator = (userRole.toUpperCase() === "KOORDINATOR" || userRole.toLowerCase().includes("koordiná"));

    // Csoportok és városok feltérképezése a Csoportok munkalapról
    const groupsSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.GROUPS) ? APP.SHEETS.GROUPS : "Csoportok";
    const groupsSheet = ss.getSheetByName(groupsSheetName);
    
    let userCity = "";
    const cityGroups = []; // A koordinátor városához tartozó csoportok listája

    if (groupsSheet && groupsSheet.getLastRow() >= 2) {
      const groupsData = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, 2).getValues();
      const groupToCityMap = {};

      groupsData.forEach(row => {
        const gId = String(row[0] || "").trim();
        const gCity = String(row[1] || "").trim();
        if (gId) groupToCityMap[gId] = gCity;
      });

      // Koordinátor városának meghatározása a saját csoport_ID-ja alapján
      if (userGroup && groupToCityMap[userGroup]) {
        userCity = groupToCityMap[userGroup];
      }

      // Városhoz tartozó csoportok kigyűjtése
      if (userCity) {
        Object.keys(groupToCityMap).forEach(gId => {
          if (groupToCityMap[gId].toLowerCase() === userCity.toLowerCase()) {
            cityGroups.push(gId);
          }
        });
      }
    }

    let coordinatorOptions = [];
    if (isCoordinator) {
      // 1. Saját vezetői csoport (Alapértelmezett)
      coordinatorOptions.push({
        id: userGroup,
        viewType: "LEADERSHIP_GROUP",
        name: "👑 Saját vezetői csoport (" + userGroup + ")"
      });

      // 2. Városi összesítő nézet
      if (userCity) {
        coordinatorOptions.push({
          id: "CITY_SUMMARY",
          viewType: "CITY_SUMMARY",
          name: "📊 " + userCity + " - Városi összesítő nézet"
        });
      }

      // 3. Városi egyes csoportok nézetei (a saját vezetői csoporton kívüliek)
      cityGroups.forEach(gId => {
        if (gId !== userGroup) {
          coordinatorOptions.push({
            id: gId,
            viewType: "GROUP_READONLY",
            name: "📁 " + gId + " (Csoportnézet)"
          });
        }
      });
    }

    return {
      authorized: true,
      email: activeEmail,
      name: userName,
      csoportId: userGroup,
      role: userRole,
      isCoordinator: isCoordinator,
      userCity: userCity,
      coordinatorOptions: coordinatorOptions
    };

  } catch (error) {
    const scriptUrl = ScriptApp.getService().getUrl();
    return {
      authorized: false,
      needsAuth: false,
      email: "Azonosítási hiba történt",
      switchAccountUrl: "https://accounts.google.com/AccountChooser?continue=" + encodeURIComponent(scriptUrl)
    };
  }
}

/**
 * 1 db villámgyors összefogó betöltés az indításkor (Batch App State)
 * Ez azonnal átadja a jogosultságot, beállításokat és az első dashboard nézetet
 */
function getInitialAppState() {
  var auth = getUserAuth();
  if (!auth.authorized) {
    return { auth: auth };
  }

  var targetGroup = auth.csoportId;
  var settings = getSettingsData();
  var dashboard = getPayKalDashboardDataImpl("1H", "ALL", targetGroup, "LEADERSHIP_GROUP");

  return {
    auth: auth,
    settings: settings,
    dashboard: dashboard
  };
}

/* --- API WRAPPER FÜGGVÉNYEK --- */
function getSettings() { return getSettingsData(); }
function getExpenseId(costCenter) { return generateExpenseId(costCenter); }
function saveExpense(expense, imageUrls, expenseId, gpsCoords) { return saveExpenseData(expense, imageUrls, expenseId, gpsCoords); }
function saveIncome(incomeData) { return saveIncomeData(incomeData); }
function saveTransfer(transferData) { return saveTransferData(transferData); }
function getPayKalDashboardData(timeFilter, selectedProject, targetCsoportId, viewType) { 
  return getPayKalDashboardDataImpl(timeFilter, selectedProject, targetCsoportId, viewType); 
}