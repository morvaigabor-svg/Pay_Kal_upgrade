/**
 * Service.Dashboard.gs
 * PayKal Dashboard Adatkezelő - 3 koordinátori nézettel és Sávdiagram támogatással
 */

function getPayKalDashboardDataImpl(timeFilter, selectedProject, targetCsoportId, viewType) {
  const auth = getUserAuth();
  if (!auth.authorized) {
    throw new Error("Jogosulatlan hozzáférés!");
  }

  // Alapértelmezett beállítások
  let activeTarget = targetCsoportId || auth.csoportId;
  let activeViewType = viewType || "LEADERSHIP_GROUP";

  // Ellenőrzés: ha nem koordinátor, csak a saját csoportját érheti el
  if (!auth.isCoordinator) {
    activeTarget = auth.csoportId;
    activeViewType = "STANDARD";
  }

  const currentProject = selectedProject ? String(selectedProject).trim() : "ALL";
  const ss = SpreadsheetApp.openById(CONFIG.SHEET_ID);

  // Város csoportjainak lekérése
  let cityGroupIds = [];
  const userCity = auth.userCity || "";
  
  if (userCity) {
    const groupsSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.GROUPS) ? APP.SHEETS.GROUPS : "Csoportok";
    const groupsSheet = ss.getSheetByName(groupsSheetName);
    if (groupsSheet && groupsSheet.getLastRow() >= 2) {
      const groupsData = groupsSheet.getRange(2, 1, groupsSheet.getLastRow() - 1, 2).getValues();
      for (let i = 0; i < groupsData.length; i++) {
        const gId = String(groupsData[i][0] || "").trim();
        const gCity = String(groupsData[i][1] || "").trim();
        if (gCity.toLowerCase() === userCity.toLowerCase()) {
          cityGroupIds.push(gId);
        }
      }
    }
  }

  // --- 1. NÉZET: VÁROSI ÖSSZESÍTŐ (SÁVDIAGRAMMAL) ---
  if (activeViewType === "CITY_SUMMARY") {
    let cityTotalCash = 0;
    let cityTotalBank = 0;
    let cityTotal = 0;

    const groupBalances = {};
    cityGroupIds.forEach(id => { groupBalances[id] = { cash: 0, bank: 0, total: 0 }; });

    const balanceSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.BALANCE) ? APP.SHEETS.BALANCE : "Egyenleg";
    const balanceSheet = ss.getSheetByName(balanceSheetName);

    if (balanceSheet && balanceSheet.getLastRow() >= 2) {
      const balanceData = balanceSheet.getRange(2, 1, balanceSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < balanceData.length; i++) {
        const rowGId = String(balanceData[i][0]).trim();
        if (cityGroupIds.includes(rowGId)) {
          const c = parseAmount(balanceData[i][1]);
          const b = parseAmount(balanceData[i][2]);
          const t = parseAmount(balanceData[i][3]) || (c + b);

          groupBalances[rowGId] = { cash: c, bank: b, total: t };
          cityTotalCash += c;
          cityTotalBank += b;
          cityTotal += t;
        }
      }
    }

    // Sávdiagram adatstruktúrájának felépítése
    const barLabels = [];
    const barCashData = [];
    const barBankData = [];

    Object.keys(groupBalances).forEach(gId => {
      barLabels.push(gId);
      barCashData.push(groupBalances[gId].cash);
      barBankData.push(groupBalances[gId].bank);
    });

    return {
      viewType: "CITY_SUMMARY",
      csoportId: activeTarget || "",
      selectedProject: "ALL",
      displayTitle: `${userCity.toUpperCase()} - ÖSSZVAGYON`,
      totalBalance: cityTotal,
      cashBalance: cityTotalCash,
      bankBalance: cityTotalBank,
      chartType: "BAR",
      barData: {
        labels: barLabels,
        cash: barCashData,
        bank: barBankData
      },
      projects: [],
      labels: [],
      values: [],
      readOnly: true
    };
  }

  // --- 2. És 3. NÉZET: EGYEDI CSOPORTNÉZET (Saját vagy kiválasztott) ---
  let totalBalance = 0;
  let cashBalance = 0;
  let bankBalance = 0;

  if (currentProject === "ALL") {
    const balanceSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.BALANCE) ? APP.SHEETS.BALANCE : "Egyenleg";
    const balanceSheet = ss.getSheetByName(balanceSheetName);
    
    if (balanceSheet && balanceSheet.getLastRow() >= 2) {
      const balanceData = balanceSheet.getRange(2, 1, balanceSheet.getLastRow() - 1, 4).getValues();
      for (let i = 0; i < balanceData.length; i++) {
        if (String(balanceData[i][0]).trim() === activeTarget) {
          cashBalance = parseAmount(balanceData[i][1]);
          bankBalance = parseAmount(balanceData[i][2]);
          totalBalance = parseAmount(balanceData[i][3]) || (cashBalance + bankBalance);
          break;
        }
      }
    }
  }

  // Számolótábla idősor adatai
  const calcSheetName = (typeof APP !== 'undefined' && APP.SHEETS && APP.SHEETS.CALC) ? APP.SHEETS.CALC : "Számolótábla";
  const calcSheet = ss.getSheetByName(calcSheetName);

  const dailyDeltas = {};
  let totalNetDelta = 0;
  let projCashDelta = 0;
  let projBankDelta = 0;
  let projTotalDelta = 0;

  if (calcSheet && calcSheet.getLastRow() >= 2) {
    const calcData = calcSheet.getRange(2, 1, calcSheet.getLastRow() - 1, 11).getValues();
    
    for (let i = 0; i < calcData.length; i++) {
      const rowCsoport = String(calcData[i][0] || "").trim();
      const rawDate = calcData[i][2];
      const rowProjId = String(calcData[i][3] || "").trim();
      const rowProjName = String(calcData[i][4] || "").trim();

      const netDelta = parseAmount(calcData[i][8]);
      const cashDelta = parseAmount(calcData[i][9]);
      const bankDelta = parseAmount(calcData[i][10]);

      if (rowCsoport.toLowerCase() === activeTarget.toLowerCase()) {
        const matchesProject = (currentProject === "ALL") || 
                               (rowProjId.toLowerCase() === currentProject.toLowerCase()) || 
                               (rowProjName.toLowerCase() === currentProject.toLowerCase());

        if (matchesProject) {
          projCashDelta += cashDelta;
          projBankDelta += bankDelta;
          projTotalDelta += netDelta;

          if (rawDate) {
            const dateKey = formatDateKey(rawDate);
            if (dateKey) {
              dailyDeltas[dateKey] = (dailyDeltas[dateKey] || 0) + netDelta;
              totalNetDelta += netDelta;
            }
          }
        }
      }
    }
  }

  if (currentProject !== "ALL") {
    cashBalance = projCashDelta;
    bankBalance = projBankDelta;
    totalBalance = projTotalDelta;
  }

  // Idősoros grafikon felépítése
  const sortedDates = Object.keys(dailyDeltas).sort();
  const initialBalance = totalBalance - totalNetDelta;

  let runningBalance = initialBalance;
  const timeSeries = sortedDates.map(dateStr => {
    runningBalance += dailyDeltas[dateStr];
    const parts = dateStr.split('-');
    return {
      dateObj: new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10)),
      balance: runningBalance
    };
  });

  const now = new Date();
  let startDate = new Date();
  switch (timeFilter) {
    case '1H': startDate.setMonth(now.getMonth() - 1); break;
    case '3H': startDate.setMonth(now.getMonth() - 3); break;
    case '6H': startDate.setMonth(now.getMonth() - 6); break;
    case '1E': startDate.setFullYear(now.getFullYear() - 1); break;
    case '2E': startDate.setFullYear(now.getFullYear() - 2); break;
    case 'O':  startDate = new Date(2000, 0, 1); break;
    default:   startDate.setMonth(now.getMonth() - 1); break;
  }

  const filteredPoints = timeSeries.filter(pt => pt.dateObj >= startDate && pt.dateObj <= now);
  const labels = [];
  const values = [];

  if (filteredPoints.length > 0) {
    filteredPoints.forEach(pt => {
      labels.push(`${String(pt.dateObj.getMonth() + 1).padStart(2, '0')}.${String(pt.dateObj.getDate()).padStart(2, '0')}.`);
      values.push(pt.balance);
    });
  } else {
    labels.push(`${String(now.getMonth() + 1).padStart(2, '0')}.${String(now.getDate()).padStart(2, '0')}.`);
    values.push(totalBalance);
  }

  const isOwnGroup = (activeTarget === auth.csoportId);

  // Kijelölt csoport (activeTarget) aktív projektjeinek kinyerése
  let groupProjects = [];
  try {
    if (typeof getActiveProjectsData === "function") {
      groupProjects = getActiveProjectsData(activeTarget, ss);
    }
  } catch (e) {
    groupProjects = [];
  }

  return {
    viewType: activeViewType,
    csoportId: activeTarget,
    displayTitle: isOwnGroup ? `${activeTarget.toUpperCase()} CSOPORT VAGYONA` : `${activeTarget.toUpperCase()} (CSOPORTNÉZET)`,
    selectedProject: currentProject,
    totalBalance: totalBalance,
    cashBalance: cashBalance,
    bankBalance: bankBalance,
    chartType: "LINE",
    labels: labels,
    values: values,
    projects: groupProjects || [],
    readOnly: !isOwnGroup
  };
}

function formatDateKey(rawDate) {
  if (!rawDate) return null;
  const tz = (typeof APP !== 'undefined' && APP.TIMEZONE) ? APP.TIMEZONE : Session.getScriptTimeZone();
  if (rawDate instanceof Date && !isNaN(rawDate.getTime())) return Utilities.formatDate(rawDate, tz, "yyyy-MM-dd");
  if (typeof rawDate === 'string' && rawDate.trim() !== '') {
    let clean = rawDate.trim().replace(/\./g, '-').replace(/\s+/g, '');
    if (clean.endsWith('-')) clean = clean.slice(0, -1);
    const d = new Date(clean);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }
  return null;
}

function parseAmount(val) {
  if (typeof val === 'number') return isNaN(val) ? 0 : val;
  if (!val) return 0;
  let str = String(val).trim().replace(/[\s\u200B-\u200D\uFEFF]/g, '').replace(',', '.').replace(/[−–]/g, '-').replace(/[^0-9.-]/g, '');
  const num = parseFloat(str);
  return isNaN(num) ? 0 : num;
}