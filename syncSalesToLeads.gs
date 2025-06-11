function loadLeadsFromAPI() {
  const apiUrl = "https://netzwerk.uppr.de/api//6115e2ebc15bf7cffcf39c56dfce109acc702fe1/admin/5/get-orders.json?condition[period][from]=2025-05-05&condition[period][to]=2030-12-31&condition[paymentstatus]=all&condition[l:status]=open,confirmed,canceled,paidout&condition[l:campaigns]=168";

  const projectNames = {
    "3926759": "Performance Channel - Retargeting",
    "200543": "Adcell",
    "3035121": "iGraal DE",
    "133609": "GP One",
    "200542": "weltderrabatte.de",
    "4775371": "Tradetracker.de",
    "50008": "Shoop.de",
    "13513771": "Website",
    "50525": "advancedStore - Retargeting",
    "51761": "COUPONS.DE",
    "5931943": "Pepper Voucher DE",
    "5184902": "www.iftra.de",
    "51904": "www.sparwelt.de",
    "9886663": "TopCashback DE",
    "3202636": "Shopmate",
    "717776": "Kupona Display Performance und Retargeting",
    "3222797": "Gutscheine.codes",
    "9748711": "Kupona Rebounce",
    "4429980": "DISCOUNTO",
    "9742982": "buswelt.de"
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: { "Content-Type": "application/json" }
    });

    const jsonData = JSON.parse(response.getContentText());
    if (!Array.isArray(jsonData)) return [["Keine Daten"]];

    const headers = [
      "Zeitraum",
      "Timestamp",
      "Anzahl Tage seit Bestellung",
      "Ordertoken",
      "ISO Buchungsnummer",
      "Gesamtpreis übergeben trendtours",
      "Publisher-Provi",
      "uppr Fee",
      "Projekt",
      "Status",
      "Freigabedatum",
      "Begründung"
    ];

    const heute = new Date();
    const startDate = new Date("2025-05-05T00:00:00Z");

    const rows = jsonData
      .filter(order => {
        const orderDate = new Date(order.timestamp);
        return (
          orderDate >= startDate &&
          order.publisher_id !== 1 &&
          order.publisher_id !== 1001 &&
          order.publisher_id !== 1002 &&
          order.trigger_id !== 100 &&
          order.trigger_id !== 1 &&
          order.trigger_id !== 3
        );
      })
      .map(order => {
        const orderDate = new Date(order.timestamp);
        const tageSeitBestellung = Math.floor((heute - orderDate) / (1000 * 60 * 60 * 24));
        const kw = calculateCalendarWeek(orderDate);
        const timestamp = formatDateWithTime(orderDate);
        const projectName = projectNames[order.project_id] || order.project_id || "";
        const turnover = Number(order.turnover || 0);

        return [
          kw,
          timestamp,
          tageSeitBestellung >= 60 ? "60 Tage erreicht" : "60 Tage noch nicht erreicht",
          order.ordertoken || "",
          "",
          formatEuro(turnover),
          formatEuro(turnover * 0.07),
          formatEuro(turnover * 0.015),
          projectName,
          "",
          "",
          "",
          orderDate // → Hilfsspalte nur zum Sortieren
        ];
      });

    return [headers, ...rows];
  } catch (error) {
    Logger.log("Fehler bei API-Abfrage: " + error);
    return [["Fehler beim Abrufen der Daten: " + error]];
  }
}

function refreshLeadsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName("Leads");
  const data = loadLeadsFromAPI();

  if (!Array.isArray(data) || data.length < 2) {
    Logger.log("Daten ungültig oder leer – kein Update durchgeführt.");
    return;
  }

  const headers = data[0];
  const newRows = data.slice(1);

  const sheetData = leadsSheet.getDataRange().getValues();
  const sheetHeaders = sheetData[0].map(h => h.trim().toLowerCase());
  const sheetRows = sheetData.slice(1);

  const ordertokenCol = headers.findIndex(h => h.trim().toLowerCase() === "ordertoken");
  const sheetOrdertokenCol = sheetHeaders.indexOf("ordertoken");

  if (ordertokenCol === -1 || sheetOrdertokenCol === -1) {
    Logger.log("❌ Ordertoken-Spalte nicht gefunden.");
    return;
  }

  const existingTokens = new Set(
    sheetRows.map(row => String(row[sheetOrdertokenCol]).trim())
  );

  const incoming = newRows
    .filter(row => !existingTokens.has(String(row[ordertokenCol]).trim()))
    .map(row => ({ row, sortDate: row[12] })) // ⬅ index 12 = unsichtbares Date

  if (incoming.length > 0) {
    // Alte Daten neu laden & sortieren
  const allRows = sheetRows.map(r => {
  let sortDate;

  if (r[1] instanceof Date) {
    sortDate = r[1];
  } else if (typeof r[1] === "string" && r[1].includes(".")) {
    const [day, month, year] = r[1].split(" ")[0].split(".");
    sortDate = new Date(`${year}-${month}-${day}`);
  } else {
    sortDate = new Date("2100-01-01"); // fallback weit in Zukunft
  }

  return { row: r, sortDate };
}).concat(incoming);

    allRows.sort((a, b) => a.sortDate - b.sortDate);
    const sorted = allRows.map(obj => obj.row);

    // Setze zurück & schreibe alles neu
    leadsSheet.clearContents();
    leadsSheet.getRange(1, 1, 1, headers.length).setValues([headers]); // Kopf
    leadsSheet.getRange(2, 1, sorted.length, headers.length).setValues(
      sorted.map(r => r.slice(0, headers.length)) // schneidet Timestamp-Hilfsspalte ab
    );

    Logger.log(`✅ ${incoming.length} neue Leads ergänzt und sortiert.`);
  } else {
    Logger.log("Keine neuen Leads.");
  }

  syncStatusFromLeadsToSales();
}

function syncStatusFromLeadsToSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const leadsSheet = ss.getSheetByName("Leads");
  const salesSheet = ss.getSheetByName("Sales");

  const leadsData = leadsSheet.getDataRange().getValues();
  const salesData = salesSheet.getDataRange().getValues();

  const leadsHeaders = leadsData[0].map(h => h.trim().toLowerCase());
  const salesHeaders = salesData[0].map(h => h.trim().toLowerCase());

  const leadsOrdertokenCol = leadsHeaders.indexOf("ordertoken");
  const leadsStatusCol = leadsHeaders.indexOf("status");
  const leadsFreigabeCol = leadsHeaders.indexOf("freigabedatum");
  const leadsBegruendungCol = leadsHeaders.indexOf("begründung");

  const salesOrdertokenCol = salesHeaders.indexOf("ordertoken");
  const salesStatusCol = salesHeaders.indexOf("status");
  const salesFreigabeCol = salesHeaders.indexOf("freigabedatum");
  const salesBegruendungCol = salesHeaders.indexOf("begründung");

  const salesMap = new Map();
  for (let i = 1; i < salesData.length; i++) {
    const token = String(salesData[i][salesOrdertokenCol]).trim();
    if (token) salesMap.set(token, i + 1);
  }

  let synced = 0;
  for (let i = 1; i < leadsData.length; i++) {
    const row = leadsData[i];
    const token = String(row[leadsOrdertokenCol]).trim();
    const status = String(row[leadsStatusCol]).trim();

    if (status === "2" && salesMap.has(token)) {
      const salesRow = salesMap.get(token);
      salesSheet.getRange(salesRow, salesStatusCol + 1).setValue(status);
      salesSheet.getRange(salesRow, salesFreigabeCol + 1).setValue(row[leadsFreigabeCol]);
      salesSheet.getRange(salesRow, salesBegruendungCol + 1).setValue(row[leadsBegruendungCol]);
      synced++;
    }
  }

  Logger.log(`🔄 ${synced} Einträge von Leads ➝ Sales synchronisiert.`);
}

function calculateCalendarWeek(date) {
  const tempDate = new Date(date.getTime());
  tempDate.setHours(0, 0, 0, 0);
  tempDate.setDate(tempDate.getDate() + 3 - ((tempDate.getDay() + 6) % 7));
  const week1 = new Date(tempDate.getFullYear(), 0, 4);
  const weekNumber = 1 + Math.round(((tempDate - week1) / 86400000 - 3 + ((week1.getDay() + 6) % 7)) / 7);
  return `KW ${weekNumber}/${tempDate.getFullYear()}`;
}

function formatDateWithTime(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${day}.${month}.${year} ${hours}:${minutes}`;
}

function formatEuro(amount) {
  return Number(amount).toLocaleString("de-DE", {
    style: "currency",
    currency: "EUR"
  });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("🔁 Sync")
    .addItem("Status sync Leads ➝ Sales", "syncStatusFromLeadsToSales")
    .addItem("Leads aktualisieren", "refreshLeadsSheet")
    .addItem("Sales aktualisieren", "refreshSalesSheet")
    .addItem("ReportingLeads aktualisieren", "refreshReportingLeads")
    .addItem("ReportingSales aktualisieren", "refreshReportingSales")
    .addToUi();
}
 
