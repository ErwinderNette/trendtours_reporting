# trendtours_reporting
# 📊 Automatisiertes Reporting für Affiliate-Kampagnen mit Google Sheets + Apps Script

Dieses Projekt automatisiert die Datenverarbeitung und das Reporting von Affiliate-Transaktionen über das Netzwerk [uppr.de](https://netzwerk.uppr.de). Es basiert auf **Google Sheets** und der integrierten **Google Apps Script**-Plattform. Transaktionen werden aus einer REST-API geladen, gefiltert, formatiert, dedupliziert, chronologisch einsortiert und ausgewertet.

---

## 🚀 Features

- 🔁 **Automatischer Import** von Leads & Sales aus API
- 📆 **Chronologische Einsortierung** nach Timestamp
- 🧮 Automatische Berechnung: Publisher-Provision, Gebühren
- ✅ **Sync-Funktion**: Status-Übertragung Leads ➝ Sales
- 📅 **Reporting-Tabellen** nach Kalenderwochen (Anzahl & Umsatz)
- 🧠 **Intelligente Deduplizierung** per Ordertoken
- 🖱️ **Benutzerdefiniertes Menü** zur manuellen Steuerung

---

## 📁 Struktur des Google Sheets

| Tabellenblatt         | Funktion                                                  |
|-----------------------|-----------------------------------------------------------|
| `Leads`               | Unverifizierte Transaktionen (mit Status = 2 Freigabe)    |
| `Sales`               | Alle bestätigten Transaktionen mit Umsatz                 |
| `Reporting Leads`     | Aggregierte Übersicht Leads pro Kalenderwoche             |
| `Reporting Sales`     | Aggregierte Sales inkl. Umsatz pro Kalenderwoche          |

---

## 🛠️ Setup-Anleitung

### 1. 📄 Google Sheet vorbereiten
- Erstelle ein neues Google Spreadsheet
- Benenne folgende Sheets: `Leads`, `Sales`, `Reporting Leads`, `Reporting Sales`

### 2. ⚙️ Apps Script öffnen
- In Google Sheets → **Erweiterungen → Apps Script**
- Alle `.gs`-Dateien einfügen (z. B. `leads.gs`, `sales.gs`, `reporting.gs` etc.)

### 3. 🔑 API-URL & Key
- Die `apiUrl` enthält bereits den Token für den Zugriff auf die Daten (z. B. `...get-orders.json?...`)
- Stelle sicher, dass dieser Schlüssel gültig ist

### 4. ✅ Trigger setzen
- Öffne **Apps Script → Auslöser (Trigger)**
- Erstelle z. B. einen Zeit-Trigger für:
  - `refreshLeadsSheet` (z. B. alle 6 Stunden)
  - `refreshSalesSheet`
  - `refreshReportingLeads`
  - `refreshReportingSales`

### 5. 📋 Menü aktivieren
- Füge im Apps Script `onOpen()` hinzu:
```js
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
