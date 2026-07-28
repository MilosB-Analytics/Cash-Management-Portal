const SHEET_NAME = "Main Safe";
const SETTINGS_SHEET_NAME = "Settings";
const SMALL_SAFE_SHEET_NAME = "Small Safe";
const TRANSACTION_LOG_SHEET_NAME = "Transaction Log";

function doGet() {
  const template = HtmlService.createTemplateFromFile("index");
  template.destinations = getDestinations();
  template.activityShops = getActivityShops();
  return template.evaluate();
}

function getActivityShops() {
  return getDestinations().filter(function(destination) {
    return (
      
      !destination.endsWith(" Drop Safe")
    );
  });
}

function submitShopActivity(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  const enteredDate = new Date(data.date);
  const today = new Date();

  today.setHours(0,0,0,0);
  enteredDate.setHours(0,0,0,0);

  if (enteredDate > today) {
  throw new Error("Future dates are not allowed.");
  }

  try {
    const clean = validateShopActivityData(data);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const transactionId = getOrCreateTransactionId(data);
    assertTransactionNotUsed(transactionId);

    const ledger = resolveLedger(ss, clean.shop);

    const type = clean.action === "Income" ? "Deposit" : "Expense";
    const description = clean.description || (
      clean.action === "Income" ? "Daily income" : "Daily expense"
    );

    const newBalance = postToLedger(ledger, {
      transactionId: transactionId,
      date: clean.date,
      type: type,
      amount: clean.amount,
      description: description,
      counterparty: ""
    });

    logTransaction({
      transactionId: transactionId,
      action: clean.action,
      source: clean.shop,
      target: "",
      amount: clean.amount,
      description: description
    });

    return {
      status: "success",
      transactionId: transactionId,
      action: clean.action,
      shop: clean.shop,
      amount: clean.amount,
      type: type,
      description: description,
      newBalance: newBalance
    };

  } finally {
    lock.releaseLock();
  }
}

function submitCorrection(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  const enteredDate = new Date(data.date);
  const today = new Date();

  today.setHours(0,0,0,0);
  enteredDate.setHours(0,0,0,0);

  if (enteredDate > today) {
  throw new Error("Future dates are not allowed.");
  }
  

  try {
    const clean = validateCorrectionData(data);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const transactionId = getOrCreateTransactionId(data);
    assertTransactionNotUsed(transactionId);

    const ledger = resolveLedger(ss, clean.destination);

    const type = clean.correctionType === "Increase" ? "Deposit" : "Withdrawal";
    const description = "Correction: " + clean.description;

    const newBalance = postToLedger(ledger, {
      transactionId: transactionId,
      date: clean.date,
      type: type,
      amount: clean.amount,
      description: description,
      counterparty: ""
    });

    logTransaction({
      transactionId: transactionId,
      action: "Correction - " + clean.correctionType,
      source: clean.destination,
      target: "",
      amount: clean.amount,
      description: description
    });

    return {
      status: "success",
      transactionId: transactionId,
      destination: clean.destination,
      correctionType: clean.correctionType,
      type: type,
      amount: clean.amount,
      description: description,
      newBalance: newBalance
    };

  } finally {
    lock.releaseLock();
  }
}

function validateCorrectionData(data) {
  const amount = Number(data.amount);
  const validationAmount = Number(data.validationAmount);

  if (!data.date || !data.destination || !data.correctionType || !data.description) {
    throw new Error("Date, destination, correction type, and reason are required.");
  }

  if (data.correctionType !== "Increase" && data.correctionType !== "Decrease") {
    throw new Error("Unknown correction type.");
  }

  if (!amount || isNaN(amount) || amount <= 0) {
    throw new Error("Amount must be greater than 0.");
  }

  if (amount !== validationAmount) {
    throw new Error("Amount and validation amount do not match.");
  }

  return {
    date: data.date,
    destination: String(data.destination).trim(),
    correctionType: String(data.correctionType).trim(),
    amount: amount,
    description: String(data.description || "").trim()
  };
}

function validateShopActivityData(data) {
  const amount = Number(data.amount);
  const validationAmount = Number(data.validationAmount);

  if (!data.date || !data.action || !data.shop || !data.description) {
    throw new Error("Date, action, shop and description are required.");
  }

  

  if (data.action !== "Income" && data.action !== "Expense") {
    throw new Error("Unknown shop activity type.");
  }

  if (!amount || isNaN(amount) || amount <= 0) {
    throw new Error("Amount must be greater than 0.");
  }

 

  if (amount !== validationAmount) {
    throw new Error("Amount and validation amount do not match.");
  }

  return {
    date: data.date,
    action: String(data.action).trim(),
    shop: String(data.shop).trim(),
    amount: amount,
    description: String(data.description || "").trim()
  };
}

function getShops() {
  function getDestinations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    throw new Error("Settings sheet not found.");
  }

  return sheet
    .getRange("A:A")
    .getValues()
    .flat()
    .filter(String);
  }
}

function getDestinations() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SETTINGS_SHEET_NAME);

  if (!sheet) {
    throw new Error("Settings sheet not found.");
  }

  return sheet
    .getRange("A:A")
    .getValues()
    .flat()
    .filter(String);
}

function submitCashMovement(data) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  const enteredDate = new Date(data.date);
  const today = new Date();

  today.setHours(0,0,0,0);
  enteredDate.setHours(0,0,0,0);

  if (enteredDate > today) {
  throw new Error("Future dates are not allowed.");
  }

  try {
    const clean = validateMovementData(data);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const transactionId = getOrCreateTransactionId(data);
    assertTransactionNotUsed(transactionId);

    if (clean.currentDestination === clean.endDestination) {
      throw new Error("Current destination and end destination cannot be the same.");
    }

    const currentLedger = resolveLedger(ss, clean.currentDestination);
    const endLedger = resolveLedger(ss, clean.endDestination);

    const sourceLedger =
      clean.movementType === "Withdrawal" ? currentLedger : endLedger;

    const targetLedger =
      clean.movementType === "Withdrawal" ? endLedger : currentLedger;

    const description =
      clean.description ||
      "Cash moved from " + sourceLedger.label + " to " + targetLedger.label;

    const sourceBalance = postToLedger(sourceLedger, {
      transactionId: transactionId,
      date: clean.date,
      type: "Withdrawal",
      amount: clean.amount,
      description: description,
      counterparty: targetLedger.label
    });

    const targetBalance = postToLedger(targetLedger, {
      transactionId: transactionId,
      date: clean.date,
      type: "Deposit",
      amount: clean.amount,
      description: description,
      counterparty: sourceLedger.label
    });

    logTransaction({
      transactionId: transactionId,
      action: "Cash Movement",
      source: sourceLedger.label,
      target: targetLedger.label,
      amount: clean.amount,
      description: description
    });

    return {
      status: "success",
      transactionId: transactionId,
      source: sourceLedger.label,
      target: targetLedger.label,
      amount: clean.amount,
      sourceBalance: sourceBalance,
      targetBalance: targetBalance
    };

  } finally {
    lock.releaseLock();
  }
}

function validateMovementData(data) {
  const amount = Number(data.amount);
  const validationAmount = Number(data.validationAmount);

  if (!data.date || !data.currentDestination || !data.movementType || !data.endDestination) {
    throw new Error("Date, shop, current destination, movement type, and end destination are required.");
  }

  if (!amount || isNaN(amount) || amount <= 0) {
    throw new Error("Amount must be greater than 0.");
  }

  if (amount !== validationAmount) {
    throw new Error("Amount and validation amount do not match.");
  }

  return {
  date: data.date,
  currentDestination: String(data.currentDestination).trim(),
  movementType: String(data.movementType).trim(),
  endDestination: String(data.endDestination).trim(),
  amount,
  description: String(data.description || "").trim()
};
}

function resolveLedger(ss, destination) {
  if (destination === "Main Safe") {
    return {
      key: "main-safe",
      label: "Main Safe",
      sheet: getRequiredSheet(ss, SHEET_NAME),
      startRow: 9,
      startCol: 2,
      columns: 8,
      balanceOffset: 5,
      layout: "SAFE"
    };
  }

  if (destination === "Small Safe") {
    return {
      key: "small-safe",
      label: "Small Safe",
      sheet: getRequiredSheet(ss, SMALL_SAFE_SHEET_NAME),
      startRow: 9,
      startCol: 2,
      columns: 8,
      balanceOffset: 5,
      layout: "SAFE"
    };
  }

  return {
    key: destination,
    label: destination,
    sheet: getRequiredSheet(ss, destination),
    startRow: 8,
    startCol: 2,
    columns: 7,
    balanceOffset: 4,
    layout: "SHOP"
  };
}

function getRequiredSheet(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) throw new Error('Sheet "' + sheetName + '" not found.');
  return sheet;
}

function postToLedger(ledger, entry) {
  const row = findNextLedgerRow(ledger.sheet, ledger.startRow, ledger.startCol);
  const previousBalance = getPreviousBalance(ledger, row);



  const newBalance = previousBalance + (entry.type === "Deposit" ? entry.amount : -entry.amount);
  const userEmail = getUserEmail();

  if (ledger.layout === "SAFE") {
    ledger.sheet.getRange(row, ledger.startCol, 1, 8).setValues([[
      entry.date,
      entry.counterparty,
      entry.description,
      entry.type,
      entry.amount,
      newBalance,
      userEmail,
      entry.transactionId
    ]]);
  } else {
    ledger.sheet.getRange(row, ledger.startCol, 1, 7).setValues([[
      entry.date,
      entry.type,
      entry.description,
      entry.amount,
      newBalance,
      userEmail,
      entry.transactionId
    ]]);
  }

  return newBalance;
}

function findNextLedgerRow(sheet, startRow, startCol) {

  const rowCount = sheet.getMaxRows() - startRow + 1;
  const values = sheet.getRange(startRow, startCol, rowCount, 1).getValues();

  let lastRow = startRow - 1;

  for (let i = 0; i < values.length; i++) {
    if (values[i][0] !== "" && values[i][0] !== null) {
      lastRow = startRow + i;
    }
  }

  const nextRow = lastRow + 1;

  ensureRows(sheet, nextRow);

  return nextRow;

}

function getPreviousBalance(ledger, nextRow) {
  if (nextRow === ledger.startRow) return 0;

  const balanceCol = ledger.startCol + ledger.balanceOffset;
  return Number(ledger.sheet.getRange(nextRow - 1, balanceCol).getValue()) || 0;
}

function getBalanceForDestination(destination) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ledger = resolveLedger(ss, destination);
  const nextRow = findNextLedgerRow(ledger.sheet, ledger.startRow, ledger.startCol);
  return getPreviousBalance(ledger, nextRow);
}

function getUserEmail() {
  return Session.getActiveUser().getEmail() || "Unknown user";
}

function getTransactionLogSheet() {
  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(TRANSACTION_LOG_SHEET_NAME);

  if (!sheet) {
    throw new Error('Sheet "' + TRANSACTION_LOG_SHEET_NAME + '" not found.');
  }

  return sheet;
}

function logTransaction(entry) {
  const sheet = getTransactionLogSheet();

  sheet.appendRow([
    entry.transactionId,
    new Date(),
    getUserEmail(),
    entry.action,
    entry.source || "",
    entry.target || "",
    entry.amount,
    entry.description || ""
  ]);
}

function getOrCreateTransactionId(data) {
  if (data && data.transactionId) {
    return String(data.transactionId).trim();
  }

  return "TX-" + Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyyMMdd-HHmmss"
  ) + "-" + Math.floor(Math.random() * 100000);
}

function assertTransactionNotUsed(transactionId) {
  const sheet = getTransactionLogSheet();
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) return;

  const ids = sheet
    .getRange(2, 1, lastRow - 1, 1)
    .getValues()
    .flat();

  if (ids.indexOf(transactionId) !== -1) {
    throw new Error("Duplicate submission blocked. Transaction ID already exists: " + transactionId);
  }
}

function ensureRows(sheet, requiredRow) {

  const maxRows = sheet.getMaxRows();

  if (requiredRow >= maxRows - 10) {

    sheet.insertRowsAfter(maxRows, 100);

    console.log(
      "Added 100 rows to " +
      sheet.getName() +
      ". New total: " +
      sheet.getMaxRows()
    );
  }

}

function getDestinationHistory(destination) {

  if (!destination) {
    return "<p>No destination selected.</p>";
  }

  const sheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(destination);

  if (!sheet) {
    return "<p>Sheet not found.</p>";
  }

  const START_ROW = 8;
  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) {
    return "<p>No transactions.</p>";
  }

  const rowsToRead = Math.min(10, lastRow - START_ROW + 1);

  const firstRow = lastRow - rowsToRead + 1;

  const data = sheet
    .getRange(firstRow, 2, rowsToRead, 7)
    .getDisplayValues()
    .reverse();

  let html = `
      <table class="historyTable">
          <thead>
              <tr>
                  <th>Date</th>
                  <th>Shop Cash Flow</th>
                  <th>Description</th>
                  <th>Type</th>
                  <th>Amount</th>
              </tr>
          </thead>
          <tbody>
  `;

  data.forEach(function(row){

      html += `
          <tr>
              <td>${row[0]}</td>
              <td>${row[1]}</td>
              <td>${row[2]}</td>
              <td>${row[3]}</td>
              <td>${row[4]}</td>
          </tr>
      `;

  });

  html += `
          </tbody>
      </table>
  `;

  return html;

}
