function refreshReportingLeads() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reporting Leads");

  if (!sheet) {
    Logger.log("Tabellenblatt 'ReportingLeads' nicht gefunden.");
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
      Logger.log("Keine gültigen Daten aus der API.");
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

    // Gruppieren nach Kalenderwoche
    const grouped = {};
    filtered.forEach(order => {
      const date = new Date(order.timestamp);
      const kwKey = getWeekKey(date); // z.B. "KW21//06.05.2025–12.05.2025"
      if (!grouped[kwKey]) {
        grouped[kwKey] = { count: 0, sum: 0 };
      }
      grouped[kwKey].count++;
      grouped[kwKey].sum += Number(order.turnover) || 0;
    });

    // Vorhandene Wochen (Spalte A) prüfen
    const existingWeeks = new Set(
      sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
    );

    const newRows = Object.entries(grouped)
      .filter(([kw]) => !existingWeeks.has(kw))
      .sort()
      .map(([kw, values]) => [
        kw,
        values.count,
        formatEuro(values.sum)
      ]);

    if (newRows.length > 0) {
      const firstFreeRow = sheet.getLastRow() + 1;
      sheet.getRange(firstFreeRow, 1, newRows.length, 3).setValues(newRows);
      Logger.log(`${newRows.length} neue Wochen in ReportingLeads ergänzt.`);
    } else {
      Logger.log("Keine neuen Wochen für ReportingLeads.");
    }

  } catch (error) {
    Logger.log("Fehler bei der API-Abfrage: " + error);
  }
}
