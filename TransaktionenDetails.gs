function TransaktionenDetails() {
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
    "9742982": "buswelt.de",
    "9748711":"Kupona Rebounce"
  };

  try {
    const response = UrlFetchApp.fetch(apiUrl, {
      method: "get",
      headers: { "Content-Type": "application/json" }
    });

    const jsonData = JSON.parse(response.getContentText());
    if (!Array.isArray(jsonData)) return [["Keine Daten"]];

    const headers = [
      "Kalenderwoche",
      "Datum/Zeit (DE)",
      "Reisestart (Description)",
      "Anzahl Tage seit Bestellung",
      "Ordertoken",
      "ISO Buchungsnummer",
      "Gesamtpreis Trendtours (Turnover)",
      "Publisher-Provi",
      "uppr Fee",
      "Project",
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
        const kw = getCalendarWeek(orderDate);
        const timestamp = formatDateWithTime(orderDate);
        const tageSeitBestellung = Math.floor((heute - orderDate) / (1000 * 60 * 60 * 24));
        const statusText = tageSeitBestellung >= 30 ? "30 Tage erreicht" : "30 Tage noch nicht erreicht";
        const reisestart = order.description?.trim() || "";
        const parsedReisestart = parseAnyDate(reisestart);
        const reisestartFormatted = parsedReisestart ? formatShortDate(parsedReisestart) : reisestart;
        const turnover = Number(order.turnover || 0);
        const projectName = projectNames[order.project_id] || order.project_id || "";

        // Wir hängen am Ende das reine Date-Objekt als Hilfsspalte an (Index 13), um später sortieren zu können.
        return [
          `KW${kw}`,
          timestamp,
          reisestartFormatted,
          statusText,
          order.ordertoken || "",
          "",
          formatEuro(turnover),
          formatEuro(turnover * 0.07),
          formatEuro(turnover * 0.015),
          projectName,
          "",
          "",
          "",
          orderDate // ➕ Hilfsspalte zum Sortieren
        ];
      });

    return [headers, ...rows];
  } catch (error) {
    Logger.log("Fehler bei API-Abfrage: " + error);
    return [["Fehler beim Abrufen der Daten: " + error]];
  }
}

function refreshSalesSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const salesSheet = ss.getSheetByName("Sales");
  const leadsSheet = ss.getSheetByName("Leads");
  const data = TransaktionenDetails();

  if (!Array.isArray(data) || data.length < 2) {
    Logger.log("Keine Daten vorhanden.");
    return;
  }

  const headers = data[0];
  const newRows = data.slice(1);

  // Bestehende Daten aus dem Sheet einlesen
  const sheetData = salesSheet.getDataRange().getValues();
  const sheetHeaders = sheetData[0].map(h => h.trim().toLowerCase());
  const sheetRows = sheetData.slice(1);

  // Spaltenindex für Ordertoken bestimmen
  const ordertokenCol = headers.findIndex(h => h.trim().toLowerCase() === "ordertoken");
  const sheetOrdertokenCol = sheetHeaders.indexOf("ordertoken");

  // Vorhandene Ordertokens sammeln
  const existingTokens = new Set(
    sheetRows.map(row => String(row[sheetOrdertokenCol]).trim())
  );

  // Nur die neuen Zeilen auswählen (nach Ordertoken)
  const incoming = newRows
    .filter(row => !existingTokens.has(String(row[ordertokenCol]).trim()))
    .map(row => ({ row, sortDate: row[13] })); // Index 13 = orderDate (Date-Objekt)

  if (incoming.length > 0) {
    // Alle Zeilen (bestehende + neue) zusammenführen und nach Datum sortieren
    const allRows = sheetRows.map(r => {
      // r[1] kann entweder String (z.B. "05.06.2025 14:30") oder ein Date-Objekt sein
      let sortDate;
      if (r[1] instanceof Date) {
        // Ist bereits ein Date-Objekt
        sortDate = r[1];
      } else {
        // Erwartet String-Format "DD.MM.YYYY HH:MM"
        const teile = String(r[1]).split(" ");
        const datumsteil = teile[0] || "";
        const zeitsteil = teile[1] || "00:00";
        const [tag, monat, jahr] = datumsteil.split(".");
        const [stunden, minuten] = zeitsteil.split(":");
        sortDate = new Date(
          parseInt(jahr, 10),
          parseInt(monat, 10) - 1,
          parseInt(tag, 10),
          parseInt(stunden, 10),
          parseInt(minuten, 10)
        );
      }
      return {
        row: r.concat(""), // Dummy-Spalte für Einheitlichkeit (insg. 14 Spalten)
        sortDate: sortDate
      };
    }).concat(incoming);

    // Sortieren nach dem sortDate-Feld
    allRows.sort((a, b) => a.sortDate - b.sortDate);

    // Nur die Zeilen ohne Hilfsspalte (letzte Spalte) zurückgeben
    const sorted = allRows.map(obj => obj.row);

    // Sheet leeren und neu beschreiben
    salesSheet.clearContents();
    salesSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    salesSheet.getRange(2, 1, sorted.length, headers.length).setValues(
      sorted.map(r => r.slice(0, headers.length)) // Hilfsspalte (Index 13) entfernen
    );

    Logger.log(`✅ ${incoming.length} neue Sales ergänzt und sortiert.`);
  } else {
    Logger.log("Keine neuen Sales.");
  }
   Logger.log(`✅ ${incoming.length} neue Sales ergänzt und sortiert.`);
  updateSalesByReisestart(); // ⬅️ hier
  // Falls gewünscht: Status von Leads ➝ Sales synchronisieren
  if (typeof syncStatusFromLeadsToSales === "function") {
    syncStatusFromLeadsToSales();
  }
}

function updateSalesByReisestart() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sales");
  const data = sheet.getDataRange().getValues();
  const headers = data[0];

  const reisestartCol = headers.indexOf("Reisestart (Description)");
  const tageCol = headers.indexOf("Anzahl Tage seit Bestellung");

  if (reisestartCol === -1 || tageCol === -1) {
    Logger.log("❌ Spalten nicht gefunden.");
    return;
  }

  const heute = new Date();
  let updated = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const reisestartStr = row[reisestartCol];

    let datum = null;

    if (reisestartStr instanceof Date) {
      datum = reisestartStr;
    } else if (typeof reisestartStr === "string" && reisestartStr.includes(".")) {
      const [dd, mm, yyyy] = reisestartStr.split(".");
      datum = new Date(`${yyyy}-${mm}-${dd}`);
    }

   if (datum instanceof Date && !isNaN(datum)) {
  const diffTage = Math.max(0, Math.floor((heute - datum) / (1000 * 60 * 60 * 24)));
  const neuerWert = diffTage >= 30 ? "30 Tage erreicht" : "30 Tage noch nicht erreicht";

     if (row[tageCol] !== neuerWert) {
    sheet.getRange(i + 1, tageCol + 1).setValue(neuerWert);
    updated++;
  }
}

  }

  Logger.log(`📆 ${updated} Sales-Zeilen basierend auf Reisestart aktualisiert.`);
}

// ➕ Ergänzende Hilfsfunktionen

function getCalendarWeek(date) {
  const target = new Date(date.valueOf());
  const dayNr = (date.getDay() + 6) % 7;
  target.setDate(target.getDate() - dayNr + 3);
  const firstThursday = new Date(target.getFullYear(), 0, 4);
  const diff = target - firstThursday;
  return 1 + Math.round(diff / (7 * 24 * 60 * 60 * 1000));
}

function parseAnyDate(input) {
  const dmy = input.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (dmy) return new Date(parseInt(dmy[3], 10), parseInt(dmy[2], 10) - 1, parseInt(dmy[1], 10));
  const ydm = input.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (ydm) return new Date(parseInt(ydm[1], 10), parseInt(ydm[2], 10) - 1, parseInt(ydm[3], 10));
  return null;
}

function formatDateWithTime(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  const hours = String(date.getHours()).padStart(2, "0");
  const minutes = String(date.getMinutes()).padStart(2, "0");
  return `${day}.${month}.${year} ${hours}:${minutes}`;
}

function formatShortDate(date) {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

function formatEuro(amount) {
  return Number(amount).toLocaleString("de-DE", {
    style: "currency",
    currency: "EUR"
  });
}
