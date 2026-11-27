/**
 * Автожурнал — мобильное приложение + боковая панель
 * Лист данных: "Журнал"
 * A=дата, B=место, C=операция, D=цена, E=кол-во, F=сумма, G=пробег, H=комментарий
 */

const SHEET_NAME = 'Журнал';
const HEADER_ROW = 1;
const PURCHASE_MILEAGE = 202909;          // пробег при покупке
const PURCHASE_DATE = '2025-09-20';       // дата покупки

// ==================== МЕНЮ ====================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Автожурнал')
    .addItem('➕ Добавить запись (панель)', 'showSidebar')
    .addItem('📱 Открыть приложение', 'openWebApp')
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutput(getAppHtml(true))
    .setTitle('Автожурнал')
    .setWidth(340);
  SpreadsheetApp.getUi().showSidebar(html);
}

function openWebApp() {
  const url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('Сначала сделай Deploy → New deployment → Web app');
    return;
  }
  const html = HtmlService.createHtmlOutput(
    `<script>window.open('${url}','_blank');google.script.host.close();</script>`
  );
  SpreadsheetApp.getUi().showModalDialog(html, 'Opening...');
}

// ==================== WEB APP ====================
function doGet() {
  return HtmlService.createHtmlOutput(getAppHtml(false))
    .setTitle('Автожурнал')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ==================== ДАННЫЕ ====================
function getLastMileage() {
  const sheet = getSheet_();
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return 0;
  const values = sheet.getRange(HEADER_ROW + 1, 7, lastRow - HEADER_ROW, 1).getValues();
  for (let i = values.length - 1; i >= 0; i--) {
    const v = values[i][0];
    if (v !== '' && v !== null && !isNaN(v)) return Number(v);
  }
  return 0;
}

function getLastFuelPrice() {
  const sheet = getSheet_();
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow <= HEADER_ROW) return 0;
  const data = sheet.getRange(HEADER_ROW + 1, 3, lastRow - HEADER_ROW, 2).getValues(); // C=op, D=price
  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === 'топливо' && data[i][1]) return Number(data[i][1]);
  }
  return 0;
}

function getDropdownData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let operations = ['топливо', 'мойка', 'запчасти', 'расходники', 'услуги', 'ремонт', 'покупка авто', '-'];
  let places = [];

  const ref = ss.getSheetByName('Справочники');
  if (ref) {
    const ops = ref.getRange('A2:A').getValues().flat().filter(String);
    const pls = ref.getRange('B2:B').getValues().flat().filter(String);
    if (ops.length) operations = ops;
    if (pls.length) places = pls;
  }

  if (!places.length) {
    const sheet = getSheet_();
    if (sheet && sheet.getLastRow() > HEADER_ROW) {
      const vals = sheet.getRange(HEADER_ROW + 1, 2, sheet.getLastRow() - HEADER_ROW, 1).getValues().flat().filter(String);
      places = [...new Set(vals)].sort();
    }
  }

  return {
    operations,
    places,
    lastMileage: getLastMileage(),
    lastFuelPrice: getLastFuelPrice(),
    today: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd')
  };
}

function addRecord(data) {
  const sheet = getSheet_();
  if (!sheet) throw new Error('Лист "Журнал" не найден');

  const date = data.date ? new Date(data.date) : new Date();
  const price = parseFloat(String(data.price).replace(',', '.')) || 0;
  const qty = parseFloat(String(data.qty).replace(',', '.')) || 0;
  const sum = Math.round(price * qty * 100) / 100;
  const mileage = parseFloat(String(data.mileage).replace(',', '.')) || 0;
  const operation = (data.operation || '').trim();

  // Если операции ещё нет в справочнике — добавляем
  if (operation) {
    addToReference_('operations', operation);
  }

  sheet.appendRow([
    date,
    data.place || '',
    operation,
    price,
    qty,
    sum,
    mileage,
    data.comment || ''
  ]);

  const r = sheet.getLastRow();
  sheet.getRange(r, 1).setNumberFormat('yyyy-mm-dd');
  sheet.getRange(r, 4).setNumberFormat('0.00');
  sheet.getRange(r, 5).setNumberFormat('0.###');
  sheet.getRange(r, 6).setNumberFormat('0.00');

  return {
    success: true,
    message: 'Добавлено',
    sum,
    lastMileage: mileage,
    lastFuelPrice: operation === 'топливо' ? price : getLastFuelPrice(),
    newOperation: operation
  };
}

function getStats() {
  const sheet = getSheet_();
  if (!sheet || sheet.getLastRow() <= HEADER_ROW) {
    return emptyStats_();
  }

  const data = sheet.getRange(HEADER_ROW + 1, 1, sheet.getLastRow() - HEADER_ROW, 8).getValues();
  let totalFuelMoney = 0, totalFuelLiters = 0;
  let byCat = { топливо: 0, мойка: 0, запчасти: 0, ремонт: 0, услуги: 0, расходники: 0 };
  let maxMileage = 0;
  let lastOilChangeMileage = null;

  data.forEach(row => {
    const op = row[2];
    const sum = Number(row[5]) || 0;
    const qty = Number(row[4]) || 0;
    const mileage = Number(row[6]) || 0;
    if (mileage > maxMileage) maxMileage = mileage;

    if (byCat.hasOwnProperty(op)) byCat[op] += sum;
    if (op === 'топливо') {
      totalFuelMoney += sum;
      totalFuelLiters += qty;
    }
    // грубый поиск последней замены масла (по комментарию или типу)
    if (op === 'расходники' && row[7] && String(row[7]).toLowerCase().includes('масл')) {
      lastOilChangeMileage = mileage;
    }
  });

  const myMileage = Math.max(0, maxMileage - PURCHASE_MILEAGE);
  const days = Math.max(1, Math.floor((new Date() - new Date(PURCHASE_DATE)) / 86400000));
  const totalSpend = Object.values(byCat).reduce((a, b) => a + b, 0);
  const avgL100 = myMileage > 0 ? (totalFuelLiters / myMileage * 100) : 0;
  const costKmFuel = myMileage > 0 ? totalFuelMoney / myMileage : 0;
  const costKmAll = myMileage > 0 ? totalSpend / myMileage : 0;
  const perDay = totalSpend / days;
  const kmPerDay = myMileage / days;

  return {
    totalSpend: round2_(totalSpend),
    byCat: {
      топливо: round2_(byCat.топливо),
      мойка: round2_(byCat.мойка),
      запчасти: round2_(byCat.запчасти),
      ремонт: round2_(byCat.ремонт),
      услуги: round2_(byCat.услуги),
      расходники: round2_(byCat.расходники)
    },
    myMileage: Math.round(myMileage),
    avgL100: round2_(avgL100),
    kmPerDay: round1_(kmPerDay),
    kmPerMonth: Math.round(kmPerDay * 30),
    litersPerMonth: round1_(kmPerDay * 30 * avgL100 / 100),
    fuelMoneyPerMonth: Math.round(kmPerDay * 30 * costKmFuel),
    costKmFuel: round2_(costKmFuel),
    costKmAll: round2_(costKmAll),
    perDay: round2_(perDay),
    days,
    maxMileage: Math.round(maxMileage),
    purchaseDate: PURCHASE_DATE,
    afterOil: lastOilChangeMileage ? Math.round(maxMileage - lastOilChangeMileage) : null
  };
}

function getJournal(limit) {
  limit = limit || 40;
  const sheet = getSheet_();
  if (!sheet || sheet.getLastRow() <= HEADER_ROW) return [];

  const lastRow = sheet.getLastRow();
  const start = Math.max(HEADER_ROW + 1, lastRow - limit + 1);
  const rows = sheet.getRange(start, 1, lastRow - start + 1, 8).getValues();

  return rows.reverse().map(r => ({
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], Session.getScriptTimeZone(), 'dd.MM.yy') : r[0],
    place: r[1] || '',
    op: r[2] || '',
    price: r[3],
    qty: r[4],
    sum: r[5],
    mileage: r[6],
    comment: r[7] || ''
  }));
}

function getReglament() {
  // Можно потом читать из листа "Регламент". Пока — актуальные данные из твоего скрина.
  const current = getLastMileage();
  const items = [
    { name: 'Масло ДВС', every: 5000, last: 214410, date: '23.08.2026', volume: '4.3л', type: '5W40 Лукойл' },
    { name: 'Масло АКПП', every: 25000, last: 210834, date: '17.05.2026', volume: '2,5–3,2л', type: 'ATF Z1' },
    { name: 'Масло ГУР', every: 25000, last: null, date: null, volume: '1л', type: 'PSF' },
    { name: 'Топливный фильтр', every: 25000, last: null, date: null, volume: '', type: '' },
    { name: 'Тормозная жидкость', every: null, last: null, date: null, volume: '1л', type: 'DOT 3/4', note: '1 раз в 2 года' },
    { name: 'Антифриз', every: null, last: 200000, date: null, volume: '5л', type: 'G12+', note: '1 раз в 2 года' },
    { name: 'Колодки передние', every: 35000, last: 214410, date: '23.08.2026', volume: '', type: '8445' },
    { name: 'Колодки задние', every: 40000, last: null, date: null, volume: '', type: '8255' },
    { name: 'Свечи', every: 20000, last: null, date: null, volume: '', type: '' },
    { name: 'Регулировка клапанов', every: 40000, last: null, date: null, volume: '', type: '' }
  ];

  return items.map(it => {
    const nextMileage = it.every && it.last ? it.last + it.every : (it.every ? current + it.every : null);
    const left = nextMileage ? nextMileage - current : null;
    return {
      ...it,
      nextMileage,
      left,
      status: left === null ? 'unknown' : left <= 0 ? 'overdue' : left < 1000 ? 'soon' : 'ok'
    };
  });
}

// ==================== HELPERS ====================
function addToReference_(type, value) {
  if (!value) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let ref = ss.getSheetByName('Справочники');
  if (!ref) {
    ref = ss.insertSheet('Справочники');
    ref.getRange('A1').setValue('операция');
    ref.getRange('B1').setValue('место');
  }

  const col = type === 'operations' ? 1 : 2; // A = операции, B = места
  const last = ref.getLastRow();
  const existing = last > 1
    ? ref.getRange(2, col, last - 1, 1).getValues().flat().map(String)
    : [];

  if (!existing.includes(value)) {
    ref.getRange(last + 1, col).setValue(value);
  }
}

function getSheet_() {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
}
function round2_(n) { return Math.round(n * 100) / 100; }
function round1_(n) { return Math.round(n * 10) / 10; }
function emptyStats_() {
  return {
    totalSpend: 0, byCat: {}, myMileage: 0, avgL100: 0, kmPerDay: 0, kmPerMonth: 0,
    litersPerMonth: 0, fuelMoneyPerMonth: 0, costKmFuel: 0, costKmAll: 0, perDay: 0,
    days: 0, maxMileage: 0, purchaseDate: PURCHASE_DATE, afterOil: null
  };
}

// ==================== HTML ====================
function getAppHtml(isSidebar) {
  return `
<!DOCTYPE html>
<html lang="ru">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover">
  <meta name="apple-mobile-web-app-capable" content="yes">
  <meta name="apple-mobile-web-app-status-bar-style" content="black-translucent">
  <meta name="theme-color" content="#1a73e8">
  <title>Автожурнал</title>
  <style>
    :root {
      --primary: #1a73e8;
      --bg: #f0f2f5;
      --card: #fff;
      --text: #202124;
      --muted: #5f6368;
      --border: #dadce0;
      --ok: #137333;
      --warn: #e37400;
      --bad: #c5221f;
      --nav-h: 56px;
      --safe-b: env(safe-area-inset-bottom, 0px);
    }
    * { box-sizing: border-box; margin: 0; padding: 0; -webkit-tap-highlight-color: transparent; }
    html, body { height: 100%; font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif; background: var(--bg); color: var(--text); font-size: 15px; }
    body { display: flex; flex-direction: column; max-width: 480px; margin: 0 auto; }

    .screens { flex: 1; overflow-y: auto; -webkit-overflow-scrolling: touch; padding-bottom: calc(var(--nav-h) + var(--safe-b) + 8px); }
    .screen { display: none; padding: 12px; }
    .screen.active { display: block; }

    /* ===== ADD ===== */
    .sum-mini {
      background: linear-gradient(135deg, #e8f0fe, #d2e3fc);
      border-radius: 10px;
      padding: 8px 12px;
      text-align: center;
      margin-bottom: 10px;
      font-size: 1.25rem;
      font-weight: 700;
      color: var(--primary);
    }
    label { display: block; font-size: 0.72rem; font-weight: 600; color: var(--muted); margin: 8px 0 3px; }
    input, select, textarea {
      width: 100%; padding: 10px 12px; border: 1.5px solid var(--border); border-radius: 10px;
      font-size: 16px; background: #fff; appearance: none;
    }
    input:focus, select:focus, textarea:focus { outline: none; border-color: var(--primary); }
    textarea { min-height: 48px; resize: none; }
    .row { display: flex; gap: 8px; }
    .row > div { flex: 1; }
    .hint { font-size: 0.7rem; color: #9aa0a6; margin-top: 2px; }

    .quick-ops { display: flex; flex-wrap: wrap; gap: 6px; margin-bottom: 4px; }
    .quick-ops button {
      flex: 1 1 28%; padding: 8px 4px; border: 1.5px solid var(--border); border-radius: 8px;
      background: #fff; font-size: 0.8rem; font-weight: 500;
    }
    .quick-ops button.active { background: var(--primary); color: #fff; border-color: var(--primary); }

    .btn {
      width: 100%; padding: 12px; border: none; border-radius: 10px; font-size: 0.95rem; font-weight: 600;
      margin-top: 10px; cursor: pointer;
    }
    .btn-primary { background: var(--primary); color: #fff; }
    .btn-primary:disabled { background: #dadce0; color: #9aa0a6; }
    .btn-ghost { background: #f1f3f4; color: var(--text); margin-top: 6px; }

    .status { margin-top: 10px; padding: 10px; border-radius: 10px; text-align: center; font-weight: 500; display: none; font-size: 0.9rem; }
    .status.ok { display: block; background: #e6f4ea; color: var(--ok); }
    .status.err { display: block; background: #fce8e6; color: var(--bad); }

    /* ===== CARDS / LISTS ===== */
    .card {
      background: var(--card); border-radius: 12px; padding: 12px 14px;
      box-shadow: 0 1px 4px rgba(0,0,0,0.06); margin-bottom: 10px;
    }
    .card-title { font-size: 0.75rem; font-weight: 600; color: var(--muted); margin-bottom: 8px; text-transform: uppercase; letter-spacing: 0.03em; }
    .stat-row { display: flex; justify-content: space-between; align-items: baseline; padding: 5px 0; border-bottom: 1px solid #f0f0f0; }
    .stat-row:last-child { border-bottom: none; }
    .stat-label { color: var(--muted); font-size: 0.88rem; }
    .stat-value { font-weight: 600; font-size: 0.95rem; }
    .big-num { font-size: 1.5rem; font-weight: 700; color: var(--primary); text-align: center; padding: 6px 0; }

    .cat-row { display: flex; justify-content: space-between; padding: 7px 0; border-bottom: 1px solid #f0f0f0; }
    .cat-row:last-child { border-bottom: none; }
    .cat-name { font-size: 0.9rem; }
    .cat-sum { font-weight: 600; }

    .reg-item { padding: 10px 0; border-bottom: 1px solid #f0f0f0; }
    .reg-item:last-child { border-bottom: none; }
    .reg-name { font-weight: 600; font-size: 0.92rem; }
    .reg-meta { font-size: 0.78rem; color: var(--muted); margin-top: 2px; }
    .reg-left { font-size: 0.82rem; font-weight: 600; margin-top: 3px; }
    .reg-left.ok { color: var(--ok); }
    .reg-left.soon { color: var(--warn); }
    .reg-left.overdue { color: var(--bad); }

    .j-item { padding: 10px 0; border-bottom: 1px solid #f0f0f0; }
    .j-item:last-child { border-bottom: none; }
    .j-top { display: flex; justify-content: space-between; font-weight: 600; font-size: 0.9rem; }
    .j-sub { font-size: 0.78rem; color: var(--muted); margin-top: 2px; }
    .j-sum { color: var(--primary); }

    /* ===== BOTTOM NAV ===== */
    .nav {
      position: fixed; bottom: 0; left: 0; right: 0; max-width: 480px; margin: 0 auto;
      height: calc(var(--nav-h) + var(--safe-b));
      padding-bottom: var(--safe-b);
      background: #fff; border-top: 1px solid var(--border);
      display: flex; z-index: 100;
      box-shadow: 0 -2px 10px rgba(0,0,0,0.05);
    }
    .nav button {
      flex: 1; border: none; background: none; padding: 6px 0 4px;
      font-size: 0.65rem; color: var(--muted); display: flex; flex-direction: column; align-items: center; gap: 2px;
    }
    .nav button.active { color: var(--primary); }
    .nav .icon { font-size: 1.25rem; line-height: 1; }
  </style>
</head>
<body>
  <div class="screens">
    <!-- 1. ADD -->
    <div class="screen active" id="screen-add">
      <div class="sum-mini" id="sumDisplay">0.00 ₽</div>

      <label>Дата</label>
      <input type="date" id="date">

      <label>Операция</label>
      <div class="quick-ops" id="quickOps"></div>
      <input type="text" id="operation" list="operationsList" placeholder="Выбери или впиши свою">
      <datalist id="operationsList"></datalist>

      <label>Место</label>
      <input type="text" id="place" list="placesList" placeholder="АЗС, сервис...">
      <datalist id="placesList"></datalist>

      <div class="row">
        <div>
          <label>Цена</label>
          <input type="number" id="price" inputmode="decimal" step="0.01" min="0" placeholder="0">
        </div>
        <div>
          <label>Кол-во</label>
          <input type="number" id="qty" inputmode="decimal" step="0.001" min="0" placeholder="1">
        </div>
      </div>

      <label>Пробег</label>
      <input type="number" id="mileage" inputmode="numeric" step="1" min="0">
      <div class="hint">Последний пробег — можно поправить</div>

      <label>Комментарий</label>
      <textarea id="comment" placeholder="Необязательно"></textarea>

      <button class="btn btn-primary" id="btnAdd" onclick="submit()">Добавить</button>
      <button class="btn btn-ghost" onclick="resetForm()">Очистить</button>
      <div id="status" class="status"></div>
    </div>

    <!-- 2. STATS -->
    <div class="screen" id="screen-stats">
      <div class="card">
        <div class="card-title">Владение</div>
        <div class="big-num" id="s-days">—</div>
        <div style="text-align:center;color:var(--muted);font-size:0.8rem;margin-bottom:8px">дней</div>
        <div class="stat-row"><span class="stat-label">Мой пробег</span><span class="stat-value" id="s-mileage">—</span></div>
        <div class="stat-row"><span class="stat-label">Средний пробег</span><span class="stat-value" id="s-kmday">—</span></div>
        <div class="stat-row"><span class="stat-label">Общий пробег</span><span class="stat-value" id="s-totalkm">—</span></div>
      </div>
      <div class="card">
        <div class="card-title">Расход</div>
        <div class="stat-row"><span class="stat-label">л / 100 км</span><span class="stat-value" id="s-l100">—</span></div>
        <div class="stat-row"><span class="stat-label">₽ / км (бензин)</span><span class="stat-value" id="s-costfuel">—</span></div>
        <div class="stat-row"><span class="stat-label">₽ / км (всё)</span><span class="stat-value" id="s-costall">—</span></div>
        <div class="stat-row"><span class="stat-label">В месяц (л / ₽)</span><span class="stat-value" id="s-month">—</span></div>
        <div class="stat-row"><span class="stat-label">Владение в день</span><span class="stat-value" id="s-perday">—</span></div>
        <div class="stat-row"><span class="stat-label">После замены масла</span><span class="stat-value" id="s-oil">—</span></div>
      </div>
    </div>

    <!-- 3. EXPENSES -->
    <div class="screen" id="screen-exp">
      <div class="card">
        <div class="card-title">Всего расходов</div>
        <div class="big-num" id="e-total">—</div>
      </div>
      <div class="card" id="e-cats"></div>
    </div>

    <!-- 4. REGLAMENT -->
    <div class="screen" id="screen-reg">
      <div class="card" id="r-list"></div>
    </div>

    <!-- 5. JOURNAL -->
    <div class="screen" id="screen-journal">
      <div class="card" id="j-list"><div style="color:var(--muted);text-align:center;padding:20px">Загрузка...</div></div>
    </div>
  </div>

  <!-- BOTTOM NAV -->
  <nav class="nav">
    <button class="active" data-screen="add" onclick="go('add')">
      <span class="icon">＋</span>Запись
    </button>
    <button data-screen="stats" onclick="go('stats')">
      <span class="icon">📊</span>Стат
    </button>
    <button data-screen="exp" onclick="go('exp')">
      <span class="icon">💰</span>Расх
    </button>
    <button data-screen="reg" onclick="go('reg')">
      <span class="icon">🔧</span>ТО
    </button>
    <button data-screen="journal" onclick="go('journal')">
      <span class="icon">📋</span>Журн
    </button>
  </nav>

  <script>
    let lastMileage = 0, lastFuelPrice = 0, currentOp = 'топливо';
    let cache = {};

    // ---- init ----
    google.script.run.withSuccessHandler(initForm).withFailureHandler(e => showStatus(e.message, false)).getDropdownData();

    function initForm(data) {
      lastMileage = data.lastMileage || 0;
      lastFuelPrice = data.lastFuelPrice || 0;
      document.getElementById('date').value = data.today;
      document.getElementById('mileage').value = lastMileage;

      const opSelect = document.getElementById('operation');
      const quick = document.getElementById('quickOps');
      data.operations.forEach(op => {
        const o = document.createElement('option');
        o.value = op; o.textContent = op; opSelect.appendChild(o);
        if (['топливо','мойка','запчасти','услуги','расходники'].includes(op)) {
          const b = document.createElement('button');
          b.type = 'button'; b.textContent = op;
          b.onclick = () => selectOp(op);
          if (op === 'топливо') b.classList.add('active');
          quick.appendChild(b);
        }
      });
      opSelect.value = 'топливо';
      applyFuelPrice();

      data.places.forEach(p => {
        const o = document.createElement('option'); o.value = p;
        document.getElementById('placesList').appendChild(o);
      });

      document.getElementById('price').addEventListener('input', calcSum);
      document.getElementById('qty').addEventListener('input', calcSum);
      opSelect.addEventListener('change', () => { currentOp = opSelect.value; updateQuick(); onOpChange(); });
    }

    function selectOp(op) {
      currentOp = op;
      document.getElementById('operation').value = op;
      updateQuick();
      onOpChange();
    }
    function updateQuick() {
      document.querySelectorAll('.quick-ops button').forEach(b => b.classList.toggle('active', b.textContent === currentOp));
    }
    function onOpChange() {
      if (currentOp === 'топливо') applyFuelPrice();
      const qty = document.getElementById('qty');
      if (!qty.value && ['мойка','услуги','запчасти','расходники','ремонт'].includes(currentOp)) {
        qty.value = 1; calcSum();
      }
    }
    function applyFuelPrice() {
      if (lastFuelPrice > 0) {
        document.getElementById('price').value = lastFuelPrice;
        calcSum();
      }
    }
    function calcSum() {
      const p = parseFloat(document.getElementById('price').value) || 0;
      const q = parseFloat(document.getElementById('qty').value) || 0;
      const s = Math.round(p * q * 100) / 100;
      document.getElementById('sumDisplay').textContent = s.toLocaleString('ru-RU', {minimumFractionDigits:2, maximumFractionDigits:2}) + ' ₽';
    }

    function submit() {
      const btn = document.getElementById('btnAdd');
      btn.disabled = true; btn.textContent = 'Сохраняю...';
      const data = {
        date: document.getElementById('date').value,
        place: document.getElementById('place').value.trim(),
        operation: document.getElementById('operation').value,
        price: document.getElementById('price').value,
        qty: document.getElementById('qty').value,
        mileage: document.getElementById('mileage').value,
        comment: document.getElementById('comment').value.trim()
      };
      google.script.run.withSuccessHandler(res => {
        showStatus('✓ ' + res.message + ' · ' + res.sum + ' ₽', true);
        lastMileage = res.lastMileage;
        if (res.lastFuelPrice) lastFuelPrice = res.lastFuelPrice;
        document.getElementById('place').value = '';
        document.getElementById('qty').value = '';
        document.getElementById('comment').value = '';
        document.getElementById('sumDisplay').textContent = '0.00 ₽';
        document.getElementById('mileage').value = lastMileage;
        if (currentOp === 'топливо') document.getElementById('price').value = lastFuelPrice;
        else document.getElementById('price').value = '';
        btn.disabled = false; btn.textContent = 'Добавить';
        cache = {}; // сброс кэша статистики
        if (navigator.vibrate) navigator.vibrate(25);
      }).withFailureHandler(e => {
        showStatus(e.message || 'Ошибка', false);
        btn.disabled = false; btn.textContent = 'Добавить';
      }).addRecord(data);
    }

    function resetForm() {
      document.getElementById('place').value = '';
      document.getElementById('price').value = currentOp === 'топливо' ? lastFuelPrice : '';
      document.getElementById('qty').value = '';
      document.getElementById('comment').value = '';
      document.getElementById('mileage').value = lastMileage;
      document.getElementById('sumDisplay').textContent = '0.00 ₽';
      document.getElementById('status').style.display = 'none';
      calcSum();
    }

    function showStatus(msg, ok) {
      const el = document.getElementById('status');
      el.textContent = msg;
      el.className = 'status ' + (ok ? 'ok' : 'err');
    }

    // ---- navigation ----
    function go(name) {
      document.querySelectorAll('.screen').forEach(s => s.classList.remove('active'));
      document.getElementById('screen-' + name).classList.add('active');
      document.querySelectorAll('.nav button').forEach(b => b.classList.toggle('active', b.dataset.screen === name));
      if (name === 'stats') loadStats();
      if (name === 'exp') loadExp();
      if (name === 'reg') loadReg();
      if (name === 'journal') loadJournal();
    }

    function loadStats() {
      if (cache.stats) return renderStats(cache.stats);
      google.script.run.withSuccessHandler(d => { cache.stats = d; renderStats(d); }).getStats();
    }
    function renderStats(d) {
      document.getElementById('s-days').textContent = d.days;
      document.getElementById('s-mileage').textContent = d.myMileage.toLocaleString('ru-RU') + ' км';
      document.getElementById('s-kmday').textContent = d.kmPerDay + ' / ' + d.kmPerMonth + ' км';
      document.getElementById('s-totalkm').textContent = d.maxMileage.toLocaleString('ru-RU') + ' км';
      document.getElementById('s-l100').textContent = d.avgL100;
      document.getElementById('s-costfuel').textContent = d.costKmFuel;
      document.getElementById('s-costall').textContent = d.costKmAll;
      document.getElementById('s-month').textContent = d.litersPerMonth + ' л / ' + d.fuelMoneyPerMonth + ' ₽';
      document.getElementById('s-perday').textContent = d.perDay + ' ₽';
      document.getElementById('s-oil').textContent = d.afterOil != null ? d.afterOil + ' км' : '—';
    }

    function loadExp() {
      if (cache.stats) return renderExp(cache.stats);
      google.script.run.withSuccessHandler(d => { cache.stats = d; renderExp(d); }).getStats();
    }
    function renderExp(d) {
      document.getElementById('e-total').textContent = d.totalSpend.toLocaleString('ru-RU', {minimumFractionDigits:2}) + ' ₽';
      const order = ['топливо','мойка','запчасти','ремонт','услуги','расходники'];
      document.getElementById('e-cats').innerHTML = order.map(k => {
        const v = d.byCat[k] || 0;
        return \`<div class="cat-row"><span class="cat-name">\${k}</span><span class="cat-sum">\${v.toLocaleString('ru-RU',{minimumFractionDigits:2})} ₽</span></div>\`;
      }).join('');
    }

    function loadReg() {
      if (cache.reg) return renderReg(cache.reg);
      google.script.run.withSuccessHandler(d => { cache.reg = d; renderReg(d); }).getReglament();
    }
    function renderReg(items) {
      document.getElementById('r-list').innerHTML = items.map(it => {
        let leftTxt = '—', cls = '';
        if (it.left != null) {
          if (it.left <= 0) { leftTxt = 'Просрочено на ' + Math.abs(it.left) + ' км'; cls = 'overdue'; }
          else { leftTxt = 'Осталось ' + it.left + ' км'; cls = it.status; }
        } else if (it.note) leftTxt = it.note;
        return \`<div class="reg-item">
          <div class="reg-name">\${it.name}</div>
          <div class="reg-meta">\${it.every ? 'каждые ' + it.every + ' км' : it.note || ''} \${it.type ? '· ' + it.type : ''} \${it.volume ? '· ' + it.volume : ''}</div>
          <div class="reg-meta">Посл.: \${it.last ? it.last + (it.date ? ' (' + it.date + ')' : '') : '—'}</div>
          <div class="reg-left \${cls}">\${leftTxt}\${it.nextMileage ? ' → ' + it.nextMileage : ''}</div>
        </div>\`;
      }).join('');
    }

    function loadJournal() {
      google.script.run.withSuccessHandler(renderJournal).getJournal(50);
    }
    function renderJournal(rows) {
      if (!rows.length) {
        document.getElementById('j-list').innerHTML = '<div style="text-align:center;color:var(--muted);padding:20px">Пусто</div>';
        return;
      }
      document.getElementById('j-list').innerHTML = rows.map(r => \`
        <div class="j-item">
          <div class="j-top">
            <span>\${r.op}</span>
            <span class="j-sum">\${Number(r.sum).toLocaleString('ru-RU',{minimumFractionDigits:2})} ₽</span>
          </div>
          <div class="j-sub">\${r.date} · \${r.place || '—'} · \${r.mileage || '—'} км\${r.comment ? ' · ' + r.comment : ''}</div>
        </div>
      \`).join('');
    }
  </script>
</body>
</html>
  `;
}
