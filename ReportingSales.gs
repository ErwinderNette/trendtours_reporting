function refreshReportingSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Reporting Sales");

  const apiUrl = "https://netzwerk.uppr.de/api//6115e2ebc15bf7cffcf39c56dfce109acc702fe1/admin/5/get-orders.json?condition[period][from]=2025-05-05&condition[period][to]=2030-12-31&condition[paymentstatus]=all&condition[l:status]=open,confirmed,canceled,paidout&condition[l:campaigns]=168";

  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: {
        "Content-Type": "application/json"
      }
    });

    const jsonData = JSON.parse(response.getContentText());
    if (!Array.isArray(jsonData)) return;

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
      if (!grouped[kwKey]) {
        grouped[kwKey] = { count: 0, sum: 0 };
      }
      grouped[kwKey].count++;
      grouped[kwKey].sum += Number(order.turnover) || 0;
    });

    // Lese vorhandene KW-Einträge aus Spalte A (KW)
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
      Logger.log(`${newRows.length} neue Wochen ergänzt.`);
    } else {
      Logger.log("Keine neuen Wochen zu ergänzen.");
    }

  } catch (error) {
    Logger.log("Fehler bei der API-Abfrage: " + error);
  }
}

function getWeekKey(date) {
  const start = getMonday(date);
  const end = new Date(start);
  end.setDate(start.getDate() + 6);
  const weekNumber = getCalendarWeek(date);
  const startStr = formatShortDate(start);
  const endStr = formatShortDate(end);
  return `KW${weekNumber}//${startStr}-${endStr}`;
}

function getMonday(d) {
  const date = new Date(d);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1);
  return new Date(date.setDate(diff));
}

function getCalendarWeek(date) {
  const target = new Date(date.valueOf());
  const dayNr = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(target.getFullYear(), 0, 4);
  const diff = target - firstThursday;
  return 1 + Math.round(diff / (7 * 24 * 60 * 60 * 1000));
}

function formatShortDate(date) {
  const dd = String(date.getDate()).padStart(2, "0");
  const mm = String(date.getMonth() + 1).padStart(2, "0");
  const yyyy = date.getFullYear();
  return `${dd}.${mm}.${yyyy}`;
}

function formatEuro(amount) {
  return Number(amount).toLocaleString("de-DE", {
    style: "currency",
    currency: "EUR"
  });
}
