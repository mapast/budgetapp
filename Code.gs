/**
 * Mój budżet – backend (Google Apps Script, Web App)
 * Arkusz: „Budget app”
 *
 * GET  ?action=all               → wszystko, czego potrzebuje aplikacja (miesiąc, kategorie, cele, wpłaty, plan, kurs)
 * POST {action:'addExpense', category, amount, description?, month?}   month: 'yyyy-MM' (domyślnie bieżący)
 * POST {action:'deleteExpense', id}
 * POST {action:'addDeposit', goal, amount, currency}        currency: 'zł' | '€'
 * POST {action:'editDeposit', id, amount, currency}
 * POST {action:'deleteDeposit', id}
 *
 * POST wysyłaj jako text/plain (bez nagłówka JSON), żeby przeglądarka nie robiła zapytania preflight.
 */

const TZ = 'Europe/Warsaw';
const SHEET = {
  expenses: 'Wydatki',
  categories: 'Kategorie',
  goals: 'Cele',
  plan: 'Oszczędności – plan',
  deposits: 'Oszczędności – rzeczywiste',
  payments: 'Płatności',
  poduszka: 'Poduszka',
  edo: 'Poduszka – EDO'
};

// Typy wpłat na Poduszkę (kolumna G w „Oszczędności – rzeczywiste”)
const POD_TYPES = { bos: 'BOŚ', eur: 'EUR', edo: 'EDO' };

// Od tego dnia zakładka „Wydatki” pokazuje płatności na następny miesiąc (wypłata z 26. finansuje kolejny miesiąc)
const NEXT_MONTH_FROM_DAY = 26;

// Kolumny w zakładce „Oszczędności – plan” (indeks od 0, A = 0), wiersze z miesiącami: 11–21
const PLAN_FIRST_ROW = 11;
const PLAN_LAST_ROW = 21;
const PLAN_COLS = {
  'IKZE':     { stan: 17, zl: 24, eur: 25 },
  'Korea':    { stan: 18, zl: 26, eur: 27 },
  'Poduszka': { stan: 19, zl: 28, eur: 29 },
  'Operacja': { stan: 20, zl: 31, eur: 30 }
};

// ── Wejście ──────────────────────────────────────────────────────────

function doGet(e) {
  try {
    const action = (e && e.parameter && e.parameter.action) || 'all';
    if (action === 'all') return json_(readAll_());
    return json_({ ok: false, error: 'Nieznana akcja: ' + action });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    let result;
    switch (body.action) {
      case 'addExpense':    result = addExpense_(body); break;
      case 'deleteExpense': result = deleteRowById_(SHEET.expenses, body.id); break;
      case 'addDeposit':    result = addDeposit_(body); break;
      case 'editDeposit':   result = editDeposit_(body); break;
      case 'deleteDeposit': result = deleteDeposit_(body); break;
      default: return json_({ ok: false, error: 'Nieznana akcja: ' + body.action });
    }
    SpreadsheetApp.flush();
    return json_(Object.assign({ ok: true }, result, { data: readAll_() }));
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

// ── Zapis ────────────────────────────────────────────────────────────

function addExpense_(b) {
  const amount = toNumber_(b.amount);
  if (!b.category) throw new Error('Brak kategorii');
  if (!(amount > 0)) throw new Error('Kwota musi być większa od zera');
  const sh = sheet_(SHEET.expenses);
  const id = newId_();
  const now = new Date();
  const month = /^\d{4}-\d{2}$/.test(String(b.month || '')) ? String(b.month) : monthKey_(now);
  sh.appendRow([id, now, "'" + month, b.category, amount, b.description || '']);
  return { id: id };
}

function addDeposit_(b) {
  const amount = toNumber_(b.amount);
  if (!b.goal) throw new Error('Brak celu');
  if (!(amount > 0)) throw new Error('Kwota musi być większa od zera');
  const type = b.goal === 'Poduszka' ? podType_(b.type) : '';
  const currency = type === POD_TYPES.eur ? '€' : type ? 'zł' : (b.currency === '€' ? '€' : 'zł');
  const sh = sheet_(SHEET.deposits);
  const id = newId_();
  const now = new Date();
  sh.appendRow([id, now, "'" + monthKey_(now), b.goal, amount, currency, type]);
  if (type === POD_TYPES.edo) addTranche_(id, now, amount);
  return { id: id };
}

function podType_(t) {
  if (t === 'edo' || t === POD_TYPES.edo) return POD_TYPES.edo;
  if (t === 'eur' || t === POD_TYPES.eur) return POD_TYPES.eur;
  return POD_TYPES.bos;
}

// Rejestr transz EDO: ID | Data zakupu | Kwota | (liczba obligacji – formuła) | Oprocentowanie 1. roku | (rocznica – formuła) | Uwagi
function addTranche_(id, date, amount) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.edo);
  if (!sh) return;
  const rate = ss.getSheetByName(SHEET.poduszka).getRange('B13').getValue();
  const ids = sh.getRange(2, 1, Math.max(1, sh.getMaxRows() - 1), 1).getValues();
  let row = 2;
  while (row - 2 < ids.length && ids[row - 2][0] !== '') row++;
  sh.getRange(row, 1, 1, 3).setValues([[id, date, amount]]);
  sh.getRange(row, 5).setValue(rate);
}

function removeTranche_(id) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.edo);
  if (!sh) return;
  const row = findRow_(sh, id);
  if (!row) return;
  sh.getRange(row, 1, 1, 3).clearContent();
  sh.getRange(row, 5).clearContent();
  sh.getRange(row, 7).clearContent();
}

function deleteDeposit_(b) {
  removeTranche_(b.id);
  return deleteRowById_(SHEET.deposits, b.id);
}

function editDeposit_(b) {
  const amount = toNumber_(b.amount);
  if (!(amount > 0)) throw new Error('Kwota musi być większa od zera');
  const sh = sheet_(SHEET.deposits);
  const row = findRow_(sh, b.id);
  if (!row) throw new Error('Nie znaleziono wpłaty ' + b.id);
  const isPod = sh.getRange(row, 4).getValue() === 'Poduszka' && b.type;
  const newType = isPod ? podType_(b.type) : '';
  const cur = newType === POD_TYPES.eur ? '€' : newType ? 'zł' : (b.currency === '€' ? '€' : 'zł');
  sh.getRange(row, 5, 1, 2).setValues([[amount, cur]]);
  if (isPod) {
    const type = newType, was = sh.getRange(row, 7).getValue();
    sh.getRange(row, 7).setValue(type);
    if (was === POD_TYPES.edo && type !== POD_TYPES.edo) removeTranche_(b.id);
    if (type === POD_TYPES.edo && was !== POD_TYPES.edo) addTranche_(b.id, sh.getRange(row, 2).getValue(), amount);
  }
  const edoSh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.edo);
  const tr = edoSh ? findRow_(edoSh, b.id) : 0;
  if (tr) edoSh.getRange(tr, 3).setValue(amount);
  return { id: b.id };
}

function deleteRowById_(name, id) {
  const sh = sheet_(name);
  const row = findRow_(sh, id);
  if (!row) throw new Error('Nie znaleziono wiersza ' + id);
  sh.deleteRow(row);
  return { id: id };
}

// ── Odczyt ───────────────────────────────────────────────────────────

function readAll_() {
  const now = new Date();
  const month = monthKey_(now);
  const day = Number(Utilities.formatDate(now, TZ, 'd'));
  const prevMonth = monthKey_(new Date(now.getFullYear(), now.getMonth() - 1, 1));
  const nextMonth = monthKey_(new Date(now.getFullYear(), now.getMonth() + 1, 1));
  const paymentsMonth = day >= NEXT_MONTH_FROM_DAY ? nextMonth : month;
  return {
    ok: true,
    month: month,
    paymentsMonth: paymentsMonth,
    months: [prevMonth, month, nextMonth],
    budgets: readBudgets_([prevMonth, month, nextMonth]),
    payments: readPayments_(),
    poduszka: readPoduszka_(),
    expenses: readExpenses_([prevMonth, month, nextMonth]),
    rate: num_(sheet_(SHEET.goals).getRange('B11').getValue()),
    budget: readBudget_(),
    goals: readGoals_(),
    deposits: readDeposits_(),
    plan: readPlan_()
  };
}

function readBudget_() {
  const sh = sheet_(SHEET.categories);
  const total = num_(sh.getRange('B3').getValue());
  const rows = sh.getRange('A6:F15').getValues();
  const pool = sh.getRange('A16:F16').getValues()[0];
  const categories = [];
  rows.forEach(function (r) {
    const name = r[0];
    if (!name) return;
    const hasLimit = r[1] !== '' && r[1] !== null;
    const limitThisMonth = hasLimit ? num_(r[3]) : null;
    // kategoria tylko na wybrany miesiąc (np. Święta) – pomijamy w pozostałych
    if (hasLimit && limitThisMonth === 0 && r[2] !== '') return;
    categories.push({ name: name, limit: limitThisMonth, spent: num_(r[4]) });
  });
  const spent = categories.reduce(function (s, c) { return s + c.spent; }, 0);
  return {
    total: total,
    spent: spent,
    left: total - spent,
    pool: { limit: num_(pool[3]), spent: num_(pool[4]) },
    categories: categories
  };
}

// Budżet dla dowolnego miesiąca, liczony w skrypcie (arkusz liczy tylko bieżący).
function readBudgets_(months) {
  const sh = sheet_(SHEET.categories);
  const base = num_(sh.getRange('B2').getValue());
  const rows = sh.getRange('A6:C15').getValues().filter(function (r) { return r[0]; });
  const exp = readExpenses_(months);
  const out = {};
  months.forEach(function (m) {
    const mm = m.slice(5, 7);
    const spentBy = {};
    exp.filter(function (e) { return e.month === m; }).forEach(function (e) { spentBy[e.category] = (spentBy[e.category] || 0) + e.amount; });
    let total = base, fixed = 0, poolSpent = 0;
    const categories = [];
    rows.forEach(function (r) {
      const name = String(r[0]);
      const hasLimit = r[1] !== '' && r[1] !== null;
      const only = String(r[2] || '').trim();
      const onlyMm = only ? ('0' + only).slice(-2) : '';
      if (only && onlyMm !== mm) return;              // np. Święta tylko w grudniu
      const limit = hasLimit ? num_(r[1]) : null;
      if (only) total += limit || 0;
      if (limit !== null) fixed += limit; else poolSpent += spentBy[name] || 0;
      categories.push({ name: name, limit: limit, spent: spentBy[name] || 0 });
    });
    const spent = categories.reduce(function (s, c) { return s + c.spent; }, 0);
    out[m] = { total: total, spent: spent, left: total - spent, pool: { limit: total - fixed, spent: poolSpent }, categories: categories };
  });
  return out;
}

function readGoals_() {
  const sh = sheet_(SHEET.goals);
  const rate = num_(sh.getRange('B11').getValue());
  const rows = sh.getRange('A2:L5').getValues();
  const deposits = readDeposits_();
  return rows.filter(function (r) { return r[0]; }).map(function (r) {
    const name = r[0], cur = r[1] === '€' ? '€' : 'zł', start = num_(r[3]);
    const mine = deposits.filter(function (d) { return d.goal === name; });
    const zl = mine.filter(function (d) { return d.currency === 'zł'; }).reduce(function (s, d) { return s + d.amount; }, 0);
    const eur = mine.filter(function (d) { return d.currency === '€'; }).reduce(function (s, d) { return s + d.amount; }, 0);
    return {
      name: name,
      currency: cur,
      goal: num_(r[2]),
      goalZl: num_(r[11]),
      start: start,
      saved: num_(r[5]),               // w walucie celu, po dzisiejszym kursie
      savedZl: (cur === 'zł' ? start : 0) + zl,
      savedEur: (cur === '€' ? start : 0) + eur,
      progress: num_(r[6]),            // 0–1
      deadline: r[7] ? String(r[7]) : null,
      expectedEnd: r[9] ? String(r[9]) : null,
      plannedThisMonth: num_(r[10]),
      rate: rate
    };
  });
}

function readDeposits_() {
  const sh = sheet_(SHEET.deposits);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues()
    .filter(function (r) { return r[0] && r[3]; })
    .map(function (r) {
      return {
        id: String(r[0]),
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]),
        month: String(r[2]),
        goal: r[3],
        amount: num_(r[4]),
        currency: r[5] === '€' ? '€' : 'zł',
        type: r[3] === 'Poduszka' ? (r[6] === POD_TYPES.edo ? 'edo' : r[6] === POD_TYPES.eur ? 'eur' : 'bos') : ''
      };
    });
}

function readPoduszka_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(SHEET.poduszka);
  if (!sh) return null;
  const v = sh.getRange('A1:G26').getValues();
  const c = function (a1) { const m = /^([A-G])(\d+)$/.exec(a1); return v[Number(m[2]) - 1]['ABCDEFG'.indexOf(m[1])]; };
  const fmtD = function (d) { return d instanceof Date ? Utilities.formatDate(d, TZ, 'yyyy-MM-dd') : String(d || ''); };
  const schedule = [];
  for (let r = 16; r <= 26; r++) {
    const row = v[r - 1];
    if (!row[0]) continue;
    schedule.push({ month: String(row[0]).replace(/^'/, ''), bos: num_(row[1]), eur: num_(row[2]), edo: num_(row[3]), tranche: String(row[4] || ''), totalZl: num_(row[5]) });
  }
  const tranches = [];
  const ed = ss.getSheetByName(SHEET.edo);
  if (ed && ed.getLastRow() > 1) {
    ed.getRange(2, 1, ed.getLastRow() - 1, 6).getValues().forEach(function (r) {
      if (!r[0] && !r[2]) return;
      tranches.push({ id: String(r[0]), date: fmtD(r[1]), amount: num_(r[2]), bonds: num_(r[3]), rate: num_(r[4]), anniversary: fmtD(r[5]) });
    });
  }
  return {
    bos: { goal: num_(c('B5')), saved: num_(c('C5')), where: String(c('F5') || '') },
    eur: { goal: num_(c('B6')), saved: num_(c('C6')), where: String(c('F6') || '') },
    edo: { goal: num_(c('B8')), saved: num_(c('C8')), where: String(c('F8') || '') },
    goalZl: num_(c('B9')), totalZl: num_(c('C9')),
    expectedEnd: String(c('B10') || ''),
    edoRate: num_(c('B13')),
    schedule: schedule,
    tranches: tranches
  };
}

function readPayments_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET.payments);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues()
    .filter(function (r) { return r[0] && r[1] && num_(r[2]) > 0; })
    .map(function (r) {
      return { id: String(r[0]), name: String(r[1]), amount: num_(r[2]), category: String(r[3] || 'Rachunki'), day: r[4] === '' ? null : num_(r[4]) || null };
    });
}

function readExpenses_(months) {
  const sh = sheet_(SHEET.expenses);
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 6).getValues()
    .filter(function (r) { return r[0] && months.indexOf(String(r[2])) !== -1; })
    .map(function (r) {
      return {
        id: String(r[0]),
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]),
        month: String(r[2]),
        category: String(r[3]),
        amount: num_(r[4]),
        description: String(r[5] || '')
      };
    });
}

function readPlan_() {
  const sh = sheet_(SHEET.plan);
  const rows = sh.getRange(PLAN_FIRST_ROW, 1, PLAN_LAST_ROW - PLAN_FIRST_ROW + 1, 32).getValues();
  const goals = sheet_(SHEET.goals).getRange('A2:C5').getValues();
  const out = {};
  goals.forEach(function (g) {
    const name = g[0], target = num_(g[2]), c = PLAN_COLS[name];
    if (!c) return;
    const list = [];
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const zl = num_(r[c.zl]), eur = num_(r[c.eur]), stan = num_(r[c.stan]);
      if (zl > 0 || eur > 0) list.push({ month: String(r[0]), zl: round2_(zl), eur: round2_(eur), total: round2_(stan) });
      if (stan >= target - 0.5) break;
    }
    out[name] = list;
  });
  return out;
}

// ── Pomocnicze ───────────────────────────────────────────────────────

function sheet_(name) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
  if (!sh) throw new Error('Brak zakładki: ' + name);
  return sh;
}

function findRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2 || !id) return 0;
  const ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  return 0;
}

function monthKey_(d) { return Utilities.formatDate(d, TZ, 'yyyy-MM'); }
function newId_() { return Utilities.getUuid().slice(0, 8); }
function num_(v) { const n = Number(v); return isFinite(n) ? n : 0; }
function toNumber_(v) { return Number(String(v).replace(',', '.').replace(/\s/g, '')); }
function round2_(n) { return Math.round(n * 100) / 100; }

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

// Do testów z edytora: Uruchom → testRead, wynik w „Dzienniku wykonania”
function testRead() { Logger.log(JSON.stringify(readAll_(), null, 2)); }
