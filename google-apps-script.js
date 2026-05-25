// ═══════════════════════════════════════════════════════════════
//  HOCHZEIT RSVP – Google Apps Script
// ═══════════════════════════════════════════════════════════════

const MAX_COMPANIONS = 5;   // max. Begleitungen pro Person
const SUMMARY_MARKER = '📊 GESAMT';

// ── Hauptfunktion ────────────────────────────────────────────
function doPost(e) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    const data  = JSON.parse(e.postData.contents);

    // Begleitungen als Array
    const companions = data.companions
      ? data.companions.split(',').map(s => s.trim()).filter(Boolean)
      : [];

    // Spaltenüberschriften beim ersten Aufruf anlegen
    if (sheet.getLastRow() === 0) {
      initSheet(sheet);
    }

    // Zusammenfassungszeile temporär entfernen (wird am Ende neu angelegt)
    removeSummaryRow(sheet);

    // Anzahl Personen: nur bei "Ja" zählen
    const isAttending = data.attendance === 'Ja';
    const count = isAttending ? 1 + companions.length : 0;

    // Zeile aufbauen
    const row = [
      new Date().toLocaleString('de-DE', { timeZone: 'Europe/Berlin' }),
      data.name || '—',
    ];
    for (let i = 0; i < MAX_COMPANIONS; i++) {
      row.push(companions[i] || '');   // je eine Spalte pro Begleitung
    }
    row.push(
      data.attendance || '—',
      data.food       || '—',
      data.babies     || '—',
      data.message    || '',
      count                            // ganz rechts: Anzahl Personen
    );

    sheet.appendRow(row);

    // Zusammenfassungszeile am Ende neu anlegen
    addSummaryRow(sheet);

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Tabelle initialisieren ───────────────────────────────────
function initSheet(sheet) {
  const headers = ['Zeitstempel', 'Name'];
  for (let i = 1; i <= MAX_COMPANIONS; i++) headers.push('Begleitung ' + i);
  headers.push('Anwesenheit', 'Menüwahl', 'Baby / Kleinkind', 'Nachricht', 'Anzahl Personen');

  sheet.appendRow(headers);

  const hRange = sheet.getRange(1, 1, 1, headers.length);
  hRange.setFontWeight('bold');
  hRange.setBackground('#F5E6D0');
  hRange.setFontColor('#5C3B1E');
  sheet.setFrozenRows(1);

  // Spaltenbreiten
  sheet.setColumnWidth(1, 155);                              // Zeitstempel
  sheet.setColumnWidth(2, 170);                              // Name
  for (let i = 3; i <= 2 + MAX_COMPANIONS; i++)
    sheet.setColumnWidth(i, 155);                            // Begleitungen
  sheet.setColumnWidth(3 + MAX_COMPANIONS, 120);            // Anwesenheit
  sheet.setColumnWidth(4 + MAX_COMPANIONS, 125);            // Menüwahl
  sheet.setColumnWidth(5 + MAX_COMPANIONS, 155);            // Baby
  sheet.setColumnWidth(6 + MAX_COMPANIONS, 220);            // Nachricht
  sheet.setColumnWidth(7 + MAX_COMPANIONS, 130);            // Anzahl
}

// ── Zusammenfassungszeile entfernen ─────────────────────────
function removeSummaryRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const val = String(sheet.getRange(lastRow, 2).getValue());
  if (val.startsWith(SUMMARY_MARKER)) sheet.deleteRow(lastRow);
}

// ── Zusammenfassungszeile anlegen ───────────────────────────
function addSummaryRow(sheet) {
  const totalCols  = 2 + MAX_COMPANIONS + 5;  // alle Spalten
  const anzahlCol  = totalCols;               // letzte Spalte = Anzahl (1-basiert)
  const lastData   = sheet.getLastRow();
  const letter     = colLetter(anzahlCol);

  // Leere Zeile als Trenner
  const gap = new Array(totalCols).fill('');
  sheet.appendRow(gap);

  // Zusammenfassungszeile
  const summaryRow = new Array(totalCols).fill('');
  summaryRow[1] = SUMMARY_MARKER + ' — Personen mit Zusage "Ja"';   // Spalte B
  sheet.appendRow(summaryRow);

  const summaryRowNum = sheet.getLastRow();

  // Formel: Summe der Anzahl-Spalte (nur Datenzeilen, nicht Leerzeile)
  sheet.getRange(summaryRowNum, anzahlCol)
    .setFormula(`=SUM(${letter}2:${letter}${lastData})`);

  // Styling
  const summaryRange = sheet.getRange(summaryRowNum, 1, 1, totalCols);
  summaryRange.setFontWeight('bold');
  summaryRange.setBackground('#FFF3CD');
  summaryRange.setFontColor('#7B4F00');
  sheet.getRange(summaryRowNum, anzahlCol).setFontSize(13);
}

// ── Hilfsfunktion: Spaltennummer → Buchstabe ────────────────
function colLetter(n) {
  let s = '';
  while (n > 0) {
    s = String.fromCharCode(64 + ((n - 1) % 26 + 1)) + s;
    n = Math.floor((n - 1) / 26);
  }
  return s;
}
