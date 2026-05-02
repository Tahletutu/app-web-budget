const CONFIG = {
  SPREADSHEET_ID: 'PASTE_SPREADSHEET_ID_HERE',
  TRANSACTIONS_SHEET: 'Transactions',
  DASHBOARD_SHEET: 'Dashboard',
  DEFAULT_CATEGORIES: ['Courses', 'Loyer', 'Transport', 'Loisirs', 'Salaire', 'Factures', 'Santé', 'Autre'],
  MAX_RECENT_ROWS: 12,
};

function doGet(e) {
  try {
    ensureSetup_();
    refreshDashboard_();

    const action = String((e && e.parameter && e.parameter.action) || 'bootstrap').toLowerCase();

    switch (action) {
      case 'bootstrap':
        return jsonResponse_({ ok: true, data: getBootstrapData_() });
      case 'summary':
        return jsonResponse_({ ok: true, data: { summary: buildSummary_() } });
      case 'transactions':
        return jsonResponse_({ ok: true, data: { recentTransactions: getRecentTransactions_() } });
      case 'health':
        return jsonResponse_({ ok: true, data: { status: 'ok' } });
      default:
        return jsonResponse_({ ok: false, error: 'Action GET inconnue.' });
    }
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function doPost(e) {
  try {
    ensureSetup_();

    const payload = parseBody_(e);
    const action = String(payload.action || '').toLowerCase();

    switch (action) {
      case 'addtransaction': {
        const transaction = validateTransaction_(payload.transaction || payload);
        appendTransaction_(transaction);
        refreshDashboard_();
        return jsonResponse_({
          ok: true,
          message: 'Transaction ajoutée avec succès.',
          data: getBootstrapData_(),
        });
      }
      default:
        return jsonResponse_({ ok: false, error: 'Action POST inconnue.' });
    }
  } catch (error) {
    return jsonResponse_({ ok: false, error: error.message || String(error) });
  }
}

function onEdit(e) {
  const sheet = e && e.range ? e.range.getSheet() : null;
  if (sheet && sheet.getName() === CONFIG.TRANSACTIONS_SHEET) {
    refreshDashboard_();
  }
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Budget App')
    .addItem('Initialiser / réparer la feuille', 'setupBudgetSheet')
    .addItem('Recalculer le dashboard', 'refreshDashboard_')
    .addToUi();
}

function setupBudgetSheet() {
  ensureSetup_();
  refreshDashboard_();
}

function ensureSetup_() {
  const ss = openSpreadsheet_();
  const txSheet = getOrCreateSheet_(ss, CONFIG.TRANSACTIONS_SHEET);
  const dashboardSheet = getOrCreateSheet_(ss, CONFIG.DASHBOARD_SHEET);

  const header = [['Date', 'Type', 'Catégorie', 'Libellé', 'Montant', 'Notes', 'Créé le']];
  if (txSheet.getLastRow() === 0) {
    txSheet.getRange(1, 1, 1, header[0].length).setValues(header);
  } else {
    txSheet.getRange(1, 1, 1, header[0].length).setValues(header);
  }

  txSheet.setFrozenRows(1);
  txSheet.getRange('A:A').setNumberFormat('dd/MM/yyyy');
  txSheet.getRange('E:E').setNumberFormat('#,##0.00 €');
  txSheet.autoResizeColumns(1, 7);

  if (dashboardSheet.getLastRow() === 0) {
    dashboardSheet.getRange('A1').setValue('Dashboard Budget');
  }
}

function openSpreadsheet_() {
  if (!CONFIG.SPREADSHEET_ID || CONFIG.SPREADSHEET_ID === 'PASTE_SPREADSHEET_ID_HERE') {
    throw new Error('Remplace CONFIG.SPREADSHEET_ID dans Code.gs par l\'ID de ton Google Sheet.');
  }
  return SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
}

function getOrCreateSheet_(ss, name) {
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function parseBody_(e) {
  if (!e || !e.postData || !e.postData.contents) {
    return {};
  }

  const raw = e.postData.contents;

  try {
    return JSON.parse(raw);
  } catch (error) {
    return { raw };
  }
}

function validateTransaction_(input) {
  const timezone = Session.getScriptTimeZone();
  const typeRaw = String(input.type || '').trim().toLowerCase();
  const amount = Number(String(input.amount || '').replace(',', '.'));

  if (!typeRaw) {
    throw new Error('Le type est obligatoire.');
  }
  if (!Number.isFinite(amount) || amount <= 0) {
    throw new Error('Le montant doit être un nombre supérieur à 0.');
  }

  const type = normalizeType_(typeRaw);
  const date = input.date ? new Date(input.date) : new Date();

  if (isNaN(date.getTime())) {
    throw new Error('La date est invalide.');
  }

  return {
    date,
    dateLabel: Utilities.formatDate(date, timezone, 'yyyy-MM-dd'),
    type,
    category: String(input.category || 'Autre').trim() || 'Autre',
    label: String(input.label || '').trim(),
    amount,
    notes: String(input.notes || '').trim(),
    createdAt: new Date(),
  };
}

function normalizeType_(typeRaw) {
  if (['revenu', 'income', 'recette'].includes(typeRaw)) return 'revenu';
  if (['depense', 'dépense', 'expense', 'sortie'].includes(typeRaw)) return 'depense';
  throw new Error('Le type doit être revenu ou depense.');
}

function appendTransaction_(transaction) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = openSpreadsheet_().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
    sheet.appendRow([
      transaction.date,
      transaction.type,
      transaction.category,
      transaction.label,
      transaction.amount,
      transaction.notes,
      transaction.createdAt,
    ]);
    sheet.getRange(sheet.getLastRow(), 1).setNumberFormat('dd/MM/yyyy');
    sheet.getRange(sheet.getLastRow(), 5).setNumberFormat('#,##0.00 €');
  } finally {
    lock.releaseLock();
  }
}

function getAllTransactions_() {
  const sheet = openSpreadsheet_().getSheetByName(CONFIG.TRANSACTIONS_SHEET);
  const values = sheet.getDataRange().getValues();

  if (values.length <= 1) {
    return [];
  }

  return values.slice(1).filter(row => row[0] && row[1] && row[4] !== '').map(row => ({
    date: row[0] instanceof Date ? row[0] : new Date(row[0]),
    type: String(row[1] || ''),
    category: String(row[2] || ''),
    label: String(row[3] || ''),
    amount: Number(row[4] || 0),
    notes: String(row[5] || ''),
    createdAt: row[6] ? new Date(row[6]) : null,
  }));
}

function buildSummary_() {
  const timezone = Session.getScriptTimeZone();
  const transactions = getAllTransactions_();
  const now = new Date();
  const currentMonth = Utilities.formatDate(now, timezone, 'yyyy-MM');

  let totalIncome = 0;
  let totalExpense = 0;
  let monthIncome = 0;
  let monthExpense = 0;
  let expenseCount = 0;

  const monthsMap = {};
  const categoryMap = {};

  transactions.forEach(tx => {
    const monthKey = Utilities.formatDate(tx.date, timezone, 'yyyy-MM');
    const amount = Math.abs(Number(tx.amount) || 0);

    if (!monthsMap[monthKey]) {
      monthsMap[monthKey] = { income: 0, expense: 0 };
    }

    if (tx.type === 'revenu') {
      totalIncome += amount;
      monthsMap[monthKey].income += amount;
      if (monthKey === currentMonth) monthIncome += amount;
    } else {
      totalExpense += amount;
      monthsMap[monthKey].expense += amount;
      if (monthKey === currentMonth) monthExpense += amount;
      expenseCount += 1;
      categoryMap[tx.category] = (categoryMap[tx.category] || 0) + amount;
    }
  });

  const monthlyBreakdown = Object.keys(monthsMap)
    .sort()
    .reverse()
    .map(month => ({
      month,
      income: round2_(monthsMap[month].income),
      expense: round2_(monthsMap[month].expense),
      balance: round2_(monthsMap[month].income - monthsMap[month].expense),
    }));

  const categoryBreakdown = Object.keys(categoryMap)
    .map(category => ({ category, total: round2_(categoryMap[category]) }))
    .sort((a, b) => b.total - a.total);

  return {
    totalIncome: round2_(totalIncome),
    totalExpense: round2_(totalExpense),
    totalBalance: round2_(totalIncome - totalExpense),
    currentMonthIncome: round2_(monthIncome),
    currentMonthExpense: round2_(monthExpense),
    currentMonthBalance: round2_(monthIncome - monthExpense),
    transactionCount: transactions.length,
    averageExpense: round2_(expenseCount ? totalExpense / expenseCount : 0),
    monthlyBreakdown,
    categoryBreakdown,
  };
}

function getRecentTransactions_() {
  const timezone = Session.getScriptTimeZone();
  return getAllTransactions_()
    .sort((a, b) => b.date.getTime() - a.date.getTime())
    .slice(0, CONFIG.MAX_RECENT_ROWS)
    .map(tx => ({
      date: Utilities.formatDate(tx.date, timezone, 'yyyy-MM-dd'),
      type: tx.type,
      category: tx.category,
      label: tx.label,
      amount: round2_(tx.amount),
      notes: tx.notes,
    }));
}

function listCategories_() {
  const categories = new Set(CONFIG.DEFAULT_CATEGORIES);
  getAllTransactions_().forEach(tx => {
    if (tx.category) categories.add(tx.category);
  });
  return Array.from(categories).sort((a, b) => a.localeCompare(b, 'fr'));
}

function getBootstrapData_() {
  return {
    summary: buildSummary_(),
    recentTransactions: getRecentTransactions_(),
    categories: listCategories_(),
  };
}

function refreshDashboard_() {
  const sheet = openSpreadsheet_().getSheetByName(CONFIG.DASHBOARD_SHEET);
  const summary = buildSummary_();
  const recent = getRecentTransactions_();

  sheet.clear();
  sheet.setFrozenRows(1);

  sheet.getRange('A1').setValue('Dashboard Budget');
  sheet.getRange('A2:B7').setValues([
    ['Solde total', summary.totalBalance],
    ['Revenus total', summary.totalIncome],
    ['Dépenses total', summary.totalExpense],
    ['Revenus mois en cours', summary.currentMonthIncome],
    ['Dépenses mois en cours', summary.currentMonthExpense],
    ['Solde mois en cours', summary.currentMonthBalance],
  ]);

  sheet.getRange('D2:E4').setValues([
    ['Nb opérations', summary.transactionCount],
    ['Dépense moyenne', summary.averageExpense],
    ['Catégories suivies', summary.categoryBreakdown.length],
  ]);

  sheet.getRange('A9:D9').setValues([['Mois', 'Revenus', 'Dépenses', 'Solde']]);
  const monthRows = summary.monthlyBreakdown.length
    ? summary.monthlyBreakdown.map(item => [item.month, item.income, item.expense, item.balance])
    : [['Aucune donnée', '', '', '']];
  sheet.getRange(10, 1, monthRows.length, 4).setValues(monthRows);

  sheet.getRange('F9:G9').setValues([['Catégorie', 'Total dépenses']]);
  const categoryRows = summary.categoryBreakdown.length
    ? summary.categoryBreakdown.map(item => [item.category, item.total])
    : [['Aucune donnée', '']];
  sheet.getRange(10, 6, categoryRows.length, 2).setValues(categoryRows);

  sheet.getRange('I9:N9').setValues([['Date', 'Type', 'Catégorie', 'Libellé', 'Montant', 'Notes']]);
  const recentRows = recent.length
    ? recent.map(item => [item.date, item.type, item.category, item.label, item.amount, item.notes])
    : [['', '', '', '', '', '']];
  sheet.getRange(10, 9, recentRows.length, 6).setValues(recentRows);

  sheet.getRange('B2:B7').setNumberFormat('#,##0.00 €');
  sheet.getRange('E3:E3').setNumberFormat('#,##0.00 €');
  sheet.getRange('B10:D' + (9 + monthRows.length)).setNumberFormat('#,##0.00 €');
  sheet.getRange('G10:G' + (9 + categoryRows.length)).setNumberFormat('#,##0.00 €');
  sheet.getRange('M10:M' + (9 + recentRows.length)).setNumberFormat('#,##0.00 €');

  sheet.getRange('A1:N1').setFontWeight('bold');
  sheet.getRange('A9:D9').setFontWeight('bold');
  sheet.getRange('F9:G9').setFontWeight('bold');
  sheet.getRange('I9:N9').setFontWeight('bold');
  sheet.autoResizeColumns(1, 14);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

function round2_(value) {
  return Math.round((Number(value) + Number.EPSILON) * 100) / 100;
}
