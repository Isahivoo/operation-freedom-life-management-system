function saveHistorySnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  const history = ss.getSheetByName("History");

  // Aktuelle Werte aus dem Dashboard
  const datum = new Date();
  const gesamteSchulden = dashboard.getRange("H9").getValue();
  const einnahmen = dashboard.getRange("H6").getValue();
  const ausgaben = dashboard.getRange("H7").getValue();
  const cashflow = dashboard.getRange("H8").getValue();
  const verfuegbaresGeld = dashboard.getRange("H12").getValue();
  const schuldenfortschritt = dashboard.getRange("H19").getValue();

  // Anzahl der abgeschlossenen Ziele
  const goals = ss.getSheetByName("Goals");
  const statusValues = goals
    .getRange("F2:F" + goals.getLastRow())
    .getValues()
    .flat();

  const zieleAbgeschlossen = statusValues.filter(
    status => status === "Abgeschlossen"
  ).length;

  // Neuen Snapshot in History hinzufügen
  history.appendRow([
    datum,
    gesamteSchulden,
    einnahmen,
    ausgaben,
    cashflow,
    verfuegbaresGeld,
    schuldenfortschritt,
    zieleAbgeschlossen
  ]);
}

function saveWeeklyReviewSnapshot() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weeklyReview = ss.getSheetByName("Weekly Review");
  const weeklyHistory = ss.getSheetByName("Weekly History");

  const values = [
    weeklyReview.getRange("B1").getValue(),  // Wochenbeginn
    weeklyReview.getRange("B2").getValue(),  // Wochenziel
    weeklyReview.getRange("B5").getValue(),  // Einnahmen
    weeklyReview.getRange("B6").getValue(),  // Ausgaben
    weeklyReview.getRange("B7").getValue(),  // Cashflow
    weeklyReview.getRange("B8").getValue(),  // Arbeitstage
    weeklyReview.getRange("B9").getValue(),  // Ziele abgeschlossen
    weeklyReview.getRange("B10").getValue(), // Gesamtbewertung
    weeklyReview.getRange("B13").getValue(), // Erfolge
    weeklyReview.getRange("B14").getValue(), // Probleme
    weeklyReview.getRange("B15").getValue(), // Lösungen
    weeklyReview.getRange("B16").getValue(), // Was habe ich gelernt?
    weeklyReview.getRange("B19").getValue()  // Fokus nächste Woche
  ];

  weeklyHistory.appendRow(values);
}

function runOperationFreedomAutomation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dashboard = ss.getSheetByName("Dashboard");
  const logSheet = ss.getSheetByName("Automation Log");

  try {
    // Daily Snapshot speichern
    saveHistorySnapshot();

    // Formeln aktualisieren
    SpreadsheetApp.flush();

    // Letztes Update im Dashboard speichern
    dashboard.getRange("L2").setValue(new Date());

    // Erfolgreichen Lauf protokollieren
    logSheet.appendRow([
      new Date(),
      "Erfolgreich",
      "Daily Automation",
      "Mission: " + dashboard.getRange("B6").getDisplayValue() +
      " | System Status: " + dashboard.getRange("I32").getDisplayValue()
    ]);

  } catch (error) {

    // Fehler protokollieren
    logSheet.appendRow([
      new Date(),
      "Fehler",
      "Daily Automation",
      error.toString()
    ]);

    // Fehler an Apps Script weitergeben
    throw error;
  }
}
