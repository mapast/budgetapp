/**
 * Mój budżet – backend (Google Apps Script, Web App)
 * Arkusz: „Budget app”
 *
 * GET  ?action=all               → wszystko, czego potrzebuje aplikacja (miesiąc, kategorie, cele, wpłaty, plan, kurs)
 * POST {action:'addExpense', category, amount, description?}
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
  deposits: 'Oszczędności – rzeczywiste'
};

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
      case 'deleteDeposit': result = deleteRowById_(SHEET.deposits, body.id); break;
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
  sh.appendRow([id, now, "'" + monthKey_(now), b.category, amount, b.description || '']);
  return { id: id };
}

function addDeposit_(b) {
  const amount = toNumber_(b.amount);
  if (!b.goal) throw new Error('Brak celu');
  if (!(amount > 0)) throw new Error('Kwota musi być większa od zera');
  const currency = b.currency === '€' ? '€' : 'zł';
  const sh = sheet_(SHEET.deposits);
  const id = newId_();
  const now = new Date();
  sh.appendRow([id, now, "'" + monthKey_(now), b.goal, amount, currency]);
  return { id: id };
}

function editDeposit_(b) {
  const amount = toNumber_(b.amount);
  if (!(amount > 0)) throw new Error('Kwota musi być większa od zera');
  const sh = sheet_(SHEET.deposits);
  const row = findRow_(sh, b.id);
  if (!row) throw new Error('Nie znaleziono wpłaty ' + b.id);
  sh.getRange(row, 5, 1, 2).setValues([[amount, b.currency === '€' ? '€' : 'zł']]);
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
  const month = monthKey_(new Date());
  return {
    ok: true,
    month: month,
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
  return sh.getRange(2, 1, last - 1, 6).getValues()
    .filter(function (r) { return r[0] && r[3]; })
    .map(function (r) {
      return {
        id: String(r[0]),
        date: r[1] instanceof Date ? Utilities.formatDate(r[1], TZ, 'yyyy-MM-dd') : String(r[1]),
        month: String(r[2]),
        goal: r[3],
        amount: num_(r[4]),
        currency: r[5] === '€' ? '€' : 'zł'
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
