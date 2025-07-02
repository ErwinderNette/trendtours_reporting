function refreshReportingSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reporting Sales");

  if (!sheet) {
    Logger.log("❌ Tabellenblatt 'Reporting Sales' nicht gefunden.");
    return;
  }

  const apiUrl = "https://netzwerk.uppr.de/api//6115e2ebc15bf7cffcf39c56dfce109acc702fe1/admin/5/get-orders.json?condition[period][from]=2025-05-05&condition[period][to]=2030-12-31&condition[paymentstatus]=all&condition[l:status]=open,confirmed,canceled,paidout&condition[l:campaigns]=168";

  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: {
        "Content-Type": "application/json"
      }
    });

    const jsonData = JSON.parse(response.getContentText());
    if (!Array.isArray(jsonData)) {
      Logger.log("❌ Ungültige API-Antwort.");
      return;
    }

    const startDate = new Date("2025-05-05T00:00:00Z");
    const filtered = jsonData.filter(order => {
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
    });

    const grouped = {};
    filtered.forEach(order => {
      const date = new Date(order.timestamp);
      const kwKey = getWeekKey(date);
      if (!grouped[kwKey]) grouped[kwKey] = { count: 0, sum: 0 };
      grouped[kwKey].count++;
      grouped[kwKey].sum += Number(order.turnover) || 0;
    });

    // Bestehende Daten auslesen
    const existingRows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    const kwToRowMap = new Map();
    for (let i = 0; i < existingRows.length; i++) {
      const kw = existingRows[i][0];
      if (kw) {
        kwToRowMap.set(kw, i + 2); // Zeile im Sheet
      }
    }

    // Neue Daten sortieren
    const sortedEntries = Object.entries(grouped).sort(([a], [b]) => {
      const dateA = parseDateFromKW(a);
      const dateB = parseDateFromKW(b);
      return dateA - dateB;
    });

    let updated = 0;
    let inserted = 0;

    sortedEntries.forEach(([kw, values]) => {
      const rowData = [[kw, values.count, formatEuro(values.sum)]];
      const rowIndex = kwToRowMap.get(kw);

      if (rowIndex) {
        sheet.getRange(rowIndex, 1, 1, 3).setValues(rowData);
        updated++;
      } else {
        const insertRow = sheet.getLastRow() + 1;
        sheet.getRange(insertRow, 1, 1, 3).setValues(rowData);
        inserted++;
      }
    });

    Logger.log(`📊 Reporting Sales: ${updated} Zeilen aktualisiert, ${inserted} ergänzt.`);

  } catch (error) {
    Logger.log("❌ Fehler bei der API-Abfrage: " + error);
  }
  function getWeekKey(date) {
  const week = getCalendarWeek(date);

  const day = date.getDay();
  const diffToMonday = (day === 0 ? -6 : 1) - day;
  const monday = new Date(date);
  monday.setDate(date.getDate() + diffToMonday);

  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);

  const fmt = d => `${String(d.getDate()).padStart(2, '0')}.${String(d.getMonth() + 1).padStart(2, '0')}.${d.getFullYear()}`;
  return `KW${week}//${fmt(monday)}-${fmt(sunday)}`;
}

function getCalendarWeek(date) {
  const target = new Date(date.valueOf());
  const dayNr = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(target.getFullYear(), 0, 4);
  const diff = target - firstThursday;
  return 1 + Math.round(diff / (7 * 24 * 60 * 60 * 1000));
}

function formatEuro(amount) {
  return Number(amount).toLocaleString("de-DE", {
    style: "currency",
    currency: "EUR"
  });
}

function parseDateFromKW(kwString) {
  const parts = kwString.split("//");
  if (parts.length !== 2) return new Date("2100-01-01");
  const startDateStr = parts[1].split("-")[0];
  const [dd, mm, yyyy] = startDateStr.split(".");
  return new Date(`${yyyy}-${mm}-${dd}`);
}

}
