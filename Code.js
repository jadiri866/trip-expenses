// ============================================================
// UNIVERSAL TRIP EXPENSE TRACKER — Master Apps Script
// ============================================================
// HOW TO INSTALL:
// 1. Open your MASTER Google Sheet (Trip Expense Tracker - Master)
// 2. Click Extensions → Apps Script
// 3. Delete any existing code, paste this entire script, Save
// 4. Click Deploy → New deployment
// 5. Type: Web app
// 6. Execute as: Me
// 7. Who has access: Anyone
// 8. Click Deploy, authorize, copy the Web App URL
// 9. Paste that URL into the HTML file where it says MASTER_SCRIPT_URL
// ============================================================

const MASTER_SHEET_ID = "1GfEo_5rTPZfeJvrR9oMEFFyFSRE_kxySiSNYdUSq3Zk";
const MASTER_TAB      = "Trips";

// ── Router ──────────────────────────────────────────────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action;
    let result;

    if      (action === "createTrip")  result = createTrip(data);
    else if (action === "logExpense")  result = logExpense(data);
    else if (action === "deleteRow")   result = deleteRow(data);
    else throw new Error("Unknown action: " + action);

    return jsonResponse({ status: "success", ...result });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.message });
  }
}

function doGet(e) {
  // Handle payload-based GET requests (all actions route through GET)
  if (e.parameter && e.parameter.payload) {
    try {
      const data   = JSON.parse(e.parameter.payload);
      const action = data.action;
      let result;
      if      (action === "getTrips")                      result = getTrips(data.email);
      else if (action === "createTrip")                    result = createTrip(data);
      else if (action === "logExpense")                    result = logExpense(data);
      else if (action === "deleteRow")                     result = deleteRow(data);
      else if (action === "addParticipant")                result = addParticipant(data);
      else if (action === "updateExpensesForNewParticipant") result = updateExpensesForNewParticipant(data);
      else if (action === "previewImport")                 result = previewImport(data);
      else if (action === "importTrip")                    result = importTrip(data);
      else if (action === "getExpenses")                   result = getExpenses(data);
      else if (action === "archiveTrip")                   result = archiveTrip(data);
      else if (action === "updateExpense")                 result = updateExpense(data);
      else throw new Error("Unknown action: " + action);
      return jsonResponse({ status: "success", ...result });
    } catch(err) {
      return jsonResponse({ status: "error", message: err.message });
    }
  }
  return jsonResponse({ status: "ok", message: "Universal Trip Expense Tracker is running." });
}

// ── Get Trips for a user ─────────────────────────────────────
function getTrips(email) {
  if (!email) throw new Error("Email required");

  ensureMasterSheet();
  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(MASTER_TAB);
  const rows  = sheet.getDataRange().getValues();

  const trips = [];
  for (let i = 1; i < rows.length; i++) {
    const row          = rows[i];
    const tripId       = row[0];
    const tripName     = row[1];
    const sheetId      = row[2];
    const sheetUrl     = row[3];
    const createdAt    = row[4];
    const participants = JSON.parse(row[5] || "[]");

    const isParticipant = participants.some(
      p => p.email && p.email.toLowerCase() === email.toLowerCase()
    );

    if (isParticipant) {
      const archived = row[6] === true || row[6] === "TRUE" || row[6] === "true";
    trips.push({ tripId, tripName, sheetId, sheetUrl, createdAt, participants, archived });
    }
  }

  return { trips };
}

// ── Create a new trip ────────────────────────────────────────
function createTrip(data) {
  const { tripName, participants } = data;
  if (!tripName)    throw new Error("Trip name required");
  if (!participants || participants.length < 1) throw new Error("At least one participant required");

  // 1. Create a new Google Sheet
  const newSS   = SpreadsheetApp.create(tripName + " — Expenses");
  const newSSId = newSS.getId();
  const newSSUrl = newSS.getUrl();
  const sheet   = newSS.getActiveSheet();
  sheet.setName("Expenses");

  // 2. Build header row
  const headers = ["What for", "Who Paid", "Total Amount"];
  participants.forEach(p => headers.push(p.name + " owes"));
  headers.push("Paid?", "Notes");
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setValues([headers]);
  headerRange.setFontWeight("bold");
  headerRange.setBackground("#1a3c1f");
  headerRange.setFontColor("#a5d6a7");
  sheet.setFrozenRows(1);

  // 3. Build summary tab
  buildSummaryTab(newSS, participants);

  // 4. Share with all participants
  participants.forEach(p => {
    if (p.email) {
      try { newSS.addEditor(p.email); } catch(e) {}
    }
  });

  // 5. Register in master sheet
  ensureMasterSheet();
  const masterSS    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_TAB);
  const tripId      = "trip_" + Date.now();
  const createdAt   = new Date().toISOString().split("T")[0];

  masterSheet.appendRow([
    tripId,
    tripName,
    newSSId,
    newSSUrl,
    createdAt,
    JSON.stringify(participants)
  ]);

  return { tripId, tripName, sheetId: newSSId, sheetUrl: newSSUrl, createdAt, participants };
}

// ── Build Summary Tab ────────────────────────────────────────
function buildSummaryTab(ss, participants) {
  const s = ss.insertSheet("Summary");

  // Headers
  s.getRange("A1").setValue("Who Paid");
  s.getRange("B1").setValue("Total Spent");
  s.getRange("C1").setValue("Amount Owed by Others");
  s.getRange("A1:C1").setFontWeight("bold").setBackground("#1a3c1f").setFontColor("#a5d6a7");

  participants.forEach((p, i) => {
    const row           = i + 2;
    const owesColIndex  = 4 + i; // D=4, E=5, F=6...
    const owesColLetter = columnToLetter(owesColIndex);
    s.getRange(row, 1).setValue(p.name);
    s.getRange(row, 2).setFormula(`=SUMIF(Expenses!B:B,"${p.name}",Expenses!C:C)`);
    s.getRange(row, 3).setFormula(`=SUMIF(Expenses!B:B,"${p.name}",Expenses!${owesColLetter}:${owesColLetter})`);
  });

  // Net settlement
  const netRow = participants.length + 3;
  s.getRange(netRow, 1).setValue("NET SETTLEMENT").setFontWeight("bold");

  participants.forEach((p, i) => {
    const row       = netRow + 1 + i;
    const spentCell = "B" + (i + 2);
    const owedCell  = "C" + (i + 2);
    s.getRange(row, 1).setValue(p.name);
    s.getRange(row, 2).setFormula(`=MAX(0,${spentCell}-${owedCell})`);
    s.getRange(row, 3).setValue("← net owed to them (0 = they owe someone)");
    s.getRange(row, 3).setFontColor("#888888").setFontStyle("italic");
    s.getRange(row, 4).setFormula(`=MAX(0,${owedCell}-${spentCell})`);
    s.getRange(row, 4).setFontColor("#888888");
    s.getRange(row, 5).setValue("← they owe this much total");
    s.getRange(row, 5).setFontColor("#888888").setFontStyle("italic");
  });

  s.autoResizeColumns(1, 5);
}

// ── Log an Expense ───────────────────────────────────────────
function logExpense(data) {
  const { sheetId, participants, expense } = data;
  if (!sheetId) throw new Error("Sheet ID required");

  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Expenses");
  if (!sheet) throw new Error("Expenses tab not found");

  const row = [expense.whatFor, expense.paidBy, parseFloat(expense.total)];
  participants.forEach(p => {
    row.push(parseFloat(expense.splits[p.name]) || 0);
  });
  row.push(""); // Paid?
  row.push(expense.notes || "");

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const newRowIndex = lastRow + 1;
  sheet.getRange(newRowIndex, 1, 1, row.length).setValues([row]);

  return { rowIndex: newRowIndex };
}

// ── Delete a Row (Undo) ──────────────────────────────────────
function deleteRow(data) {
  const { sheetId, rowIndex } = data;
  if (!sheetId)  throw new Error("Sheet ID required");
  if (!rowIndex) throw new Error("Row index required");

  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Expenses");
  if (!sheet) throw new Error("Expenses tab not found");

  // Safety check: never delete row 1 (header)
  if (rowIndex <= 1) throw new Error("Cannot delete header row");

  sheet.deleteRow(rowIndex);
  return { deleted: rowIndex };
}

// ── Add Participant to Existing Trip ─────────────────────────
function addParticipant(data) {
  const { tripId, sheetId, newParticipant, participants } = data;
  if (!tripId)         throw new Error("Trip ID required");
  if (!sheetId)        throw new Error("Sheet ID required");
  if (!newParticipant) throw new Error("New participant required");

  // 1. Add new "owes" column to the Expenses sheet
  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Expenses");
  if (!sheet) throw new Error("Expenses tab not found");

  // Find where to insert: after last "owes" column, before "Paid?"
  // Header row: What for | Who Paid | Total Amount | [P owes]... | Paid? | Notes
  const headers     = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const paidColIdx  = headers.indexOf("Paid?"); // 0-based
  if (paidColIdx < 0) throw new Error("Could not find 'Paid?' column in sheet");

  // Insert a new column before "Paid?" (1-based for Sheets API)
  const newColPos = paidColIdx + 1; // 1-based
  sheet.insertColumnBefore(newColPos);
  sheet.getRange(1, newColPos).setValue(newParticipant.name + " owes");
  sheet.getRange(1, newColPos).setFontWeight("bold").setBackground("#1a3c1f").setFontColor("#a5d6a7");

  // Fill existing rows with 0
  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    const zeros = Array(lastRow - 1).fill([0]);
    sheet.getRange(2, newColPos, lastRow - 1, 1).setValues(zeros);
  }

  // 2. Share sheet with new participant
  if (newParticipant.email) {
    try { ss.addEditor(newParticipant.email); } catch(e) {}
  }

  // 3. Update master sheet participant list
  ensureMasterSheet();
  const masterSS    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_TAB);
  const rows        = masterSheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === tripId) {
      const existing = JSON.parse(rows[i][5] || "[]");
      existing.push(newParticipant);
      masterSheet.getRange(i + 1, 6).setValue(JSON.stringify(existing));
      break;
    }
  }

  // 4. Return existing expense rows for the review screen
  const existingRows = [];
  if (lastRow > 1) {
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    data.forEach((row, i) => {
      if (row[0]) { // has a "What for" value
        existingRows.push({
          rowIndex: i + 2, // 1-based, +1 for header
          whatFor:  row[0],
          paidBy:   row[1],
          total:    row[2],
        });
      }
    });
  }

  return { existingRows };
}

// ── Update Expenses for New Participant (after review) ────────
function updateExpensesForNewParticipant(data) {
  const { sheetId, participants, newParticipant, rowsToUpdate } = data;
  if (!sheetId)     throw new Error("Sheet ID required");
  if (!rowsToUpdate || rowsToUpdate.length === 0) return { updated: 0 };

  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Expenses");
  if (!sheet) throw new Error("Expenses tab not found");

  const headers        = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const totalTravelers = participants.reduce((s, p) => s + p.travelers, 0);
  const newParticipantCol = headers.indexOf(newParticipant.name + " owes") + 1; // 1-based
  if (newParticipantCol < 1) throw new Error("Could not find column for " + newParticipant.name);

  // For each row being updated:
  // - recalculate everyone's share based on full participant list
  // - update the new participant's column with their share
  // - update all OTHER participant columns proportionally
  rowsToUpdate.forEach(rowIndex => {
    const rowData   = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    const total     = parseFloat(rowData[2]) || 0;

    // Recalculate each participant's share
    let assigned = 0;
    participants.forEach((p, i) => {
      const colIdx = headers.indexOf(p.name + " owes") + 1; // 1-based
      if (colIdx < 1) return;
      let share;
      if (i === participants.length - 1) {
        share = Math.round((total - assigned) * 100) / 100;
      } else {
        share = Math.round(total * p.travelers / totalTravelers * 100) / 100;
        assigned += share;
      }
      sheet.getRange(rowIndex, colIdx).setValue(share);
    });
  });

  return { updated: rowsToUpdate.length };
}



// ── Preview an Existing Trip (read sheet before importing) ───
function previewImport(data) {
  const { sheetId, userEmail } = data;
  if (!sheetId) throw new Error("Sheet ID required");

  // Check not already registered
  ensureMasterSheet();
  const masterSS    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_TAB);
  const rows        = masterSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === sheetId) {
      throw new Error("This sheet is already registered as \"" + rows[i][1] + "\"");
    }
  }

  // Open the sheet and read it
  const ss       = SpreadsheetApp.openById(sheetId);
  const tripName = ss.getName().replace(" — Expenses", "").replace(" - Expenses", "").trim();

  // Find the Expenses tab (try common names)
  const sheet = ss.getSheetByName("Expenses") || ss.getSheetByName("Sheet1") || ss.getActiveSheet();

  // Find the header row — scan first 20 rows for "Who Paid"
  const numRows     = Math.min(20, sheet.getLastRow());
  const numCols     = sheet.getLastColumn();
  if (numRows === 0 || numCols === 0) throw new Error("Sheet appears to be empty.");

  const searchRange = sheet.getRange(1, 1, numRows, numCols).getValues();
  let headerRowIdx  = -1;
  for (let i = 0; i < searchRange.length; i++) {
    const rowStr = searchRange[i].map(c => (c||"").toString()).join(",");
    if (rowStr.indexOf("Who Paid") >= 0) { headerRowIdx = i; break; }
  }
  if (headerRowIdx < 0) throw new Error("Could not find a header row with 'Who Paid'. Make sure row 13 (or similar) has your column headers.");

  // Extract participant names from columns ending in " owes" OR " owe"
  const headers      = searchRange[headerRowIdx];
  const participants = [];
  headers.forEach(h => {
    const hStr = (h || "").toString().trim();
    // Match "X owes" or "X owe" (case-insensitive)
    const match = hStr.match(/^(.+?)\s+owe[s]?$/i);
    if (match) {
      const name = match[1].trim();
      if (name && name.toLowerCase() !== "jace owes") { // avoid false matches
        participants.push({ name, email: "", travelers: 1 });
      }
    }
  });

  if (participants.length === 0) {
    throw new Error("Could not detect participants. Header row was: " + headers.join(" | "));
  }

  // Pre-fill userEmail on whichever participant name matches
  if (userEmail) {
    const emailLower = userEmail.toLowerCase();
    const prefix     = emailLower.split("@")[0];
    participants.forEach(p => {
      const nameLower = p.name.toLowerCase();
      // Direct name match or email prefix contains first name
      if (prefix === nameLower ||
          prefix.replace(/[^a-z]/g,"").indexOf(nameLower.split(" ")[0].replace(/[^a-z]/g,"")) >= 0 ||
          nameLower.split(" ")[0].indexOf(prefix.replace(/[^a-z]/g,"")) >= 0) {
        p.email = userEmail;
      }
    });
  }

  return { tripName, sheetId, participants, headerRowIdx };
}

// ── Import an Existing Trip ──────────────────────────────────
function importTrip(data) {
  const { tripName, sheetId, sheetUrl, participants } = data;
  if (!tripName)    throw new Error("Trip name required");
  if (!sheetId)     throw new Error("Sheet ID required");
  if (!participants || participants.length < 1) throw new Error("At least one participant required");

  ensureMasterSheet();
  const masterSS    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const masterSheet = masterSS.getSheetByName(MASTER_TAB);

  // Check not already registered
  const rows = masterSheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][2] === sheetId) {
      throw new Error("Already registered as: " + rows[i][1]);
    }
  }

  const tripId      = "trip_" + Date.now();
  const createdAt   = new Date().toISOString().split("T")[0];
  const resolvedUrl = sheetUrl || ("https://docs.google.com/spreadsheets/d/" + sheetId + "/edit");

  masterSheet.appendRow([tripId, tripName, sheetId, resolvedUrl, createdAt, JSON.stringify(participants)]);

  // Share with all participants who have emails
  const ss = SpreadsheetApp.openById(sheetId);
  participants.forEach(p => { if (p.email) { try { ss.addEditor(p.email); } catch(e) {} } });

  return { tripId, tripName, sheetId, sheetUrl: resolvedUrl, createdAt, participants };
}


// ── Get Expenses (live read from trip sheet) ─────────────────
function getExpenses(data) {
  const { sheetId } = data;
  if (!sheetId) throw new Error("Sheet ID required");

  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Expenses");
  if (!sheet) throw new Error("Expenses tab not found");

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { expenses: [] };

  // Find header row — scan first 20 rows for "Who Paid"
  const allData   = sheet.getDataRange().getValues();
  let headerRowIdx = -1;
  for (let i = 0; i < Math.min(20, allData.length); i++) {
    if (allData[i].join(",").indexOf("Who Paid") >= 0) { headerRowIdx = i; break; }
  }
  if (headerRowIdx < 0) return { expenses: [] };

  const headers  = allData[headerRowIdx];
  const expenses = [];

  // Find notes column index from headers (last column)
  const notesColIdx = headers.length - 1;

  for (let i = headerRowIdx + 1; i < allData.length; i++) {
    const row = allData[i];
    if (!row[0] && !row[1]) continue; // skip empty rows

    // Pad row to header length so column lookups don't return undefined
    while (row.length < headers.length) row.push("");

    const expense = {
      whatFor:  row[0] ? row[0].toString() : "",
      paidBy:   row[1] ? row[1].toString() : "",
      total:    row[2] ? parseFloat(row[2].toString().replace(/[$,]/g,"")) || 0 : 0,
      splits:   {},
      notes:    row[notesColIdx] ? row[notesColIdx].toString() : "",
      rowIndex: headerRowIdx + 1 + i, // headerRowIdx is 0-based; row i is at sheet row (headerRowIdx+1) + (i - headerRowIdx) = i+1 (1-based)
    };

    // Extract splits from "X owes" OR "X owe" columns (handle both variants)
    headers.forEach((h, colIdx) => {
      const hStr = (h || "").toString().trim();
      if (hStr.match(/ owe[s]?$/i)) {
        const name = hStr.replace(/ owe[s]?$/i, "").trim();
        const val  = row[colIdx];
        expense.splits[name] = val !== "" && val !== null && val !== undefined
          ? parseFloat(val.toString().replace(/[$,]/g,"")) || 0
          : 0;
      }
    });

    if (expense.whatFor || expense.paidBy) expenses.push(expense);
  }

  return { expenses };
}

// ── Archive a Trip ────────────────────────────────────────────
function archiveTrip(data) {
  const { tripId } = data;
  if (!tripId) throw new Error("Trip ID required");

  ensureMasterSheet();
  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(MASTER_TAB);
  const rows  = sheet.getDataRange().getValues();

  for (let i = 1; i < rows.length; i++) {
    if (rows[i][0] === tripId) {
      sheet.getRange(i + 1, 7).setValue(true); // Column G = Archived
      return { tripId, archived: true };
    }
  }
  throw new Error("Trip not found: " + tripId);
}

// ── Debt Simplification (minimize Venmo payments) ────────────
function simplifyDebts(participants, expenses) {
  // Calculate net balance for each participant
  // Positive = owed money, Negative = owes money
  const balances = {};
  participants.forEach(p => { balances[p.name] = 0; });

  expenses.forEach(exp => {
    const payer = exp.paidBy;
    if (balances[payer] !== undefined) {
      balances[payer] += parseFloat(exp.total) || 0;
    }
    Object.entries(exp.splits || {}).forEach(([name, amt]) => {
      if (balances[name] !== undefined) {
        balances[name] -= parseFloat(amt) || 0;
      }
    });
  });

  // Build creditors (owed money) and debtors (owe money)
  const creditors = [];
  const debtors   = [];
  Object.entries(balances).forEach(([name, bal]) => {
    const rounded = Math.round(bal * 100) / 100;
    if (rounded > 0.01)  creditors.push({ name, amount: rounded });
    if (rounded < -0.01) debtors.push({ name, amount: -rounded });
  });

  creditors.sort((a,b) => b.amount - a.amount);
  debtors.sort((a,b) => b.amount - a.amount);

  // Greedy matching algorithm
  const payments = [];
  let ci = 0, di = 0;
  while (ci < creditors.length && di < debtors.length) {
    const credit = creditors[ci];
    const debt   = debtors[di];
    const amount = Math.min(credit.amount, debt.amount);
    const rounded = Math.round(amount * 100) / 100;
    if (rounded > 0.01) {
      payments.push({ from: debt.name, to: credit.name, amount: rounded });
    }
    credit.amount -= amount;
    debt.amount   -= amount;
    if (credit.amount < 0.01) ci++;
    if (debt.amount < 0.01)   di++;
  }

  return { balances, payments };
}


// ── Update an Existing Expense Row ───────────────────────────
function updateExpense(data) {
  const { sheetId, rowIndex, participants, expense } = data;
  if (!sheetId)   throw new Error("Sheet ID required");
  if (!rowIndex)  throw new Error("Row index required");
  if (rowIndex <= 1) throw new Error("Cannot overwrite header row");

  const ss    = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName("Expenses");
  if (!sheet) throw new Error("Expenses tab not found");

  // Build updated row same structure as logExpense
  const row = [expense.whatFor, expense.paidBy, parseFloat(expense.total)];
  participants.forEach(p => {
    row.push(parseFloat(expense.splits[p.name]) || 0);
  });
  row.push(""); // Paid?
  row.push(expense.notes || "");

  sheet.getRange(rowIndex, 1, 1, row.length).setValues([row]);
  return { updated: rowIndex };
}

function ensureMasterSheet() {
  const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);
  let sheet = ss.getSheetByName(MASTER_TAB);
  if (!sheet) {
    sheet = ss.insertSheet(MASTER_TAB);
    sheet.getRange("A1:G1").setValues([["Trip ID", "Trip Name", "Sheet ID", "Sheet URL", "Created At", "Participants (JSON)", "Archived"]]);
    sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#1a3c1f").setFontColor("#a5d6a7");
  }
  return sheet;
}

function columnToLetter(col) {
  let letter = "";
  while (col > 0) {
    const rem = (col - 1) % 26;
    letter    = String.fromCharCode(65 + rem) + letter;
    col       = Math.floor((col - 1) / 26);
  }
  return letter;
}

function jsonResponse(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}