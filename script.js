// ====================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ======================
let allData = [];
let uniqueTariffs = [];
let uniqueUValues = [];
let kcrEnabled = true;
let COL = {};
let lastExportData = null;

const EXCLUDED_UC = 'Тестовый УЦ АО "КАЛУГА АСТРАЛ"';

// ====================== ОПРЕДЕЛЕНИЕ КОЛОНОК ======================
function detectColumns(headerRow) {
  const MAP = {
    "дата начала":             "V",
    "дата окончания":          "END",
    "удостоверяющий центр":    "AA",
    "спец.предложение":        "U",
    "специальное предложение": "U",
    "тариф":                   "Q",
    "снилс":                   "J",
    "инн":                     "F",
    "удаленная схема":         "REMOTE"
  };

  const result = {};
  headerRow.forEach((cell, idx) => {
    const key = String(cell || "").trim().toLowerCase();
    if (MAP[key]) result[MAP[key]] = idx;
  });

  const required = ["V", "END", "AA", "U", "Q", "J", "F"];
  const missing = required.filter(k => result[k] === undefined);
  if (missing.length) console.warn("Не найдены колонки:", missing);

  return result;
}

// ====================== ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ======================
function excelToDate(serial) {
  if (typeof serial !== "number" || serial < 1) return null;
  return new Date((serial - 25569) * 86400000);
}

function parseDate(val) {
  if (typeof val === "number") return excelToDate(val);
  if (val instanceof Date) return val;
  return null;
}

async function readFile(file) {
  if (file.name.toLowerCase().endsWith('.zip')) {
    const zip = await JSZip.loadAsync(file);
    const excelEntry = Object.keys(zip.files).find(name =>
      name.toLowerCase().endsWith('.xlsx') || name.toLowerCase().endsWith('.xls')
    );
    if (!excelEntry) throw new Error("В ZIP-архиве не найден Excel-файл");
    const arrayBuffer = await zip.files[excelEntry].async("arraybuffer");
    return parseExcel(new Uint8Array(arrayBuffer));
  } else {
    const arrayBuffer = await file.arrayBuffer();
    return parseExcel(new Uint8Array(arrayBuffer));
  }
}

function parseExcel(data) {
  const workbook = XLSX.read(data, { type: "array" });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  return XLSX.utils.sheet_to_json(worksheet, { header: 1 });
}

// ====================== ЗАГРУЗКА ======================
async function handleUpload() {
  const file = document.getElementById("file1").files[0];
  if (!file) return;
  const statusEl = document.getElementById("status1");
  statusEl.innerHTML = "⏳ Обработка файла...";

  try {
    allData = await readFile(file);
    if (!allData.length) throw new Error("Файл пуст");

    COL = detectColumns(allData[0]);

    const tempQ = new Set(), tempU = new Set(), tempAA = new Set();

    for (let i = 1; i < allData.length; i++) {
      const row = allData[i];
      if (COL.Q !== undefined && row.length > COL.Q) {
        const v = String(row[COL.Q] || "").trim();
        if (v) tempQ.add(v);
      }
      if (COL.U !== undefined && row.length > COL.U) {
        const v = String(row[COL.U] || "").trim();
        if (v) tempU.add(v);
      }
      if (COL.AA !== undefined && row.length > COL.AA) {
        const v = String(row[COL.AA] || "").trim();
        if (v && v !== EXCLUDED_UC) tempAA.add(v);
      }
    }

    uniqueTariffs = Array.from(tempQ).sort((a, b) => a.localeCompare(b, 'ru'));
    uniqueUValues = Array.from(tempU).sort((a, b) => a.localeCompare(b, 'ru'));

    populateFilters();
    populateGeneralStats();
    document.getElementById("loaded-config").style.display = "block";
    document.getElementById("analyzeBtn").disabled = false;

    statusEl.innerHTML = `✅ <strong>${file.name}</strong><br><small>Строк данных: ${(allData.length - 1).toLocaleString('ru-RU')}</small>`;
  } catch (e) {
    statusEl.innerHTML = `❌ Ошибка: ${e.message}`;
    console.error(e);
  }
}

// ====================== ФИЛЬТРЫ ======================
function populateFilters() {
  // Тарифы
  document.getElementById("q-filters").innerHTML = uniqueTariffs.map(val =>
    `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
  ).join("");

  // Специальные предложения
  const main = [], kcr = [];
  uniqueUValues.forEach(val => {
    val.toUpperCase().includes("KCR") || val.toUpperCase().includes("КЦР") ? kcr.push(val) : main.push(val);
  });

  document.getElementById("u-main").innerHTML = main.map(val =>
    `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
  ).join("");

  document.getElementById("u-kcr").innerHTML = kcr.map(val =>
    `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
  ).join("");

  // Удостоверяющие центры
  const uniqueAA = Array.from(new Set(
    allData.slice(1)
      .map(row => COL.AA !== undefined && row.length > COL.AA ? String(row[COL.AA] || "").trim() : null)
      .filter(v => v && v !== EXCLUDED_UC)
  )).sort((a, b) => a.localeCompare(b, 'ru'));

  document.getElementById("aa-filter-container").innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
      <h4 style="margin:0;">Удостоверяющий центр <small style="font-weight:400; color:#64748b;">(% считается только по отмеченным)</small></h4>
      <div style="display:flex; gap:8px;">
        <button type="button" onclick="selectAllAA(true)" style="font-size:0.8em; padding:4px 10px; background:#3b82f6; color:white; border:none; border-radius:6px; cursor:pointer;">Все</button>
        <button type="button" onclick="selectAllAA(false)" style="font-size:0.8em; padding:4px 10px; background:#94a3b8; color:white; border:none; border-radius:6px; cursor:pointer;">Ни одного</button>
      </div>
    </div>
    <div id="aa-filters" class="checkbox-group">
      ${uniqueAA.map(val => `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`).join("")}
    </div>`;

  // Удаленная схема
  populateRemoteSchemeFilter();
}

function populateRemoteSchemeFilter() {
  const remoteCol = COL.REMOTE;
  const container = document.getElementById("remote-filter-container");

  if (remoteCol === undefined) {
    container.innerHTML = `<p style="color:#ef4444; font-size:0.9em;">Колонка «Удаленная схема» не найдена</p>`;
    return;
  }

  const values = new Set();
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row.length > remoteCol) {
      const v = String(row[remoteCol] || "").trim();
      if (v) values.add(v);
    }
  }

  const uniqueRemote = Array.from(values).sort();

  container.innerHTML = `
    <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
      <h4 style="margin:0;">Удаленная схема <small style="font-weight:400; color:#64748b;">(фильтр применяется к сравнению)</small></h4>
      <div style="display:flex; gap:8px;">
        <button type="button" onclick="selectAllRemote(true)" style="font-size:0.8em; padding:4px 10px; background:#3b82f6; color:white; border:none; border-radius:6px; cursor:pointer;">Все</button>
        <button type="button" onclick="selectAllRemote(false)" style="font-size:0.8em; padding:4px 10px; background:#94a3b8; color:white; border:none; border-radius:6px; cursor:pointer;">Ни одного</button>
      </div>
    </div>
    <div id="remote-filters" class="checkbox-group">
      ${uniqueRemote.map(val => 
        `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
      ).join("")}
    </div>`;
}

// ====================== КНОПКИ "ВСЕ / НИ ОДНОГО" ======================
function selectAllQ(checked) {
  document.querySelectorAll("#q-filters input[type='checkbox']").forEach(cb => cb.checked = checked);
}

function selectAllU(checked) {
  document.querySelectorAll("#u-main input[type='checkbox'], #u-kcr input[type='checkbox']").forEach(cb => cb.checked = checked);
}

function selectAllAA(checked) {
  document.querySelectorAll("#aa-filters input[type='checkbox']").forEach(cb => cb.checked = checked);
}

function selectAllRemote(checked) {
  document.querySelectorAll("#remote-filters input[type='checkbox']").forEach(cb => cb.checked = checked);
}

// ====================== KCR ======================
function toggleKCR() {
  kcrEnabled = !kcrEnabled;
  const btn = document.getElementById("kcr-button");
  btn.textContent = kcrEnabled ? "Убрать КЦР" : "Вернуть КЦР";
  btn.style.background = kcrEnabled ? "#ef4444" : "#22c55e";

  document.querySelectorAll("#q-filters input[type='checkbox']").forEach(cb => {
    if (cb.value.toUpperCase().includes("КЦР") || cb.value.toUpperCase().includes("KCR")) 
      cb.checked = kcrEnabled;
  });
  document.querySelectorAll("#u-kcr input[type='checkbox']").forEach(cb => cb.checked = kcrEnabled);
}

function filterKCROnly() {
  document.querySelectorAll("#q-filters input[type='checkbox']").forEach(cb => {
    cb.checked = cb.value.toUpperCase().includes("КЦР") || cb.value.toUpperCase().includes("KCR");
  });
  document.querySelectorAll("#u-main input[type='checkbox']").forEach(cb => cb.checked = false);
  document.querySelectorAll("#u-kcr input[type='checkbox']").forEach(cb => cb.checked = true);
  kcrEnabled = true;
  const btn = document.getElementById("kcr-button");
  btn.textContent = "Убрать КЦР";
  btn.style.background = "#ef4444";
}

// ====================== ПЕРИОД ВОЗВРАТА ======================
function handleGracePeriodChange() {
  const sel = document.getElementById("grace-period");
  const wrap = document.getElementById("grace-custom-wrap");
  wrap.style.display = sel.value === "custom" ? "flex" : "none";
}

function getGracePeriodDays() {
  const sel = document.getElementById("grace-period");
  if (sel.value === "custom") {
    const custom = document.getElementById("grace-custom").value.trim();
    const val = parseInt(custom);
    if (!isNaN(val) && val >= 0) return val;
    return 90;
  }
  return parseInt(sel.value) || 90;
}

// ====================== СТАТИСТИКА ПО УЦ ======================
function populateGeneralStats() {
  const html = `
    <h3 style="margin-bottom:20px; text-align:center;">Статистика по удостоверяющим центрам</h3>
    <div style="margin-bottom:20px; text-align:center;">
      <select id="aa-year-filter" onchange="filterAAStats()" style="padding:10px 16px; border-radius:12px; border:2px solid #e2e8f0; margin-right:10px;">
        <option value="">Все годы</option>
      </select>
      <select id="aa-month-filter" onchange="filterAAStats()" style="padding:10px 16px; border-radius:12px; border:2px solid #e2e8f0;">
        <option value="">Все месяцы</option>
      </select>
    </div>
    <table class="result-table">
      <thead>
        <tr>
          <th>Удостоверяющий центр</th>
          <th style="text-align:right">Кол-во</th>
          <th style="text-align:right">%</th>
        </tr>
      </thead>
      <tbody id="aa-tbody"></tbody>
    </table>`;

  document.getElementById("general-stats").innerHTML = html;
  document.getElementById("general-stats").style.display = "block";
  populateAAFilters();
  filterAAStats();
}

function populateAAFilters() {
  const yearSelect = document.getElementById("aa-year-filter");
  const monthSelect = document.getElementById("aa-month-filter");
  const years = new Set(), months = new Set();

  if (COL.V === undefined) return;

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row.length > COL.V) {
      const d = parseDate(row[COL.V]);
      if (d) {
        years.add(d.getFullYear());
        months.add(d.getMonth() + 1);
      }
    }
  }

  Array.from(years).sort((a,b) => b - a).forEach(y => yearSelect.appendChild(new Option(y, y)));
  const monthNames = ["","Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
  Array.from(months).sort((a,b) => a - b).forEach(m => monthSelect.appendChild(new Option(monthNames[m], m)));
}

function filterAAStats() {
  const yearF = document.getElementById("aa-year-filter").value;
  const monthF = document.getElementById("aa-month-filter").value;
  const tbody = document.getElementById("aa-tbody");

  if (COL.AA === undefined || COL.V === undefined) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;color:#ef4444;">Колонки не определены</td></tr>`;
    return;
  }

  const countMap = {};
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row.length <= COL.AA) continue;
    const center = String(row[COL.AA] || "").trim();
    if (!center || center === EXCLUDED_UC) continue;

    const date = parseDate(row[COL.V]);
    if (!date) continue;
    if (yearF && date.getFullYear() != yearF) continue;
    if (monthF && (date.getMonth() + 1) != monthF) continue;

    countMap[center] = (countMap[center] || 0) + 1;
  }

  const sorted = Object.entries(countMap).sort((a, b) => b[1] - a[1]);
  const total = sorted.reduce((s, [, c]) => s + c, 0);

  tbody.innerHTML = sorted.length
    ? sorted.map(([center, cnt]) => {
        const pct = total > 0 ? (cnt / total * 100).toFixed(1) : "0";
        return `<tr>
          <td>${center}</td>
          <td style="text-align:right">${cnt.toLocaleString('ru-RU')}</td>
          <td style="text-align:right; color:#1e40af; font-weight:600;">${pct}%</td>
        </tr>`;
      }).join("") +
      `<tr style="background:#f1f5f9; font-weight:700;">
        <td>Итого</td>
        <td style="text-align:right">${total.toLocaleString('ru-RU')}</td>
        <td style="text-align:right">100%</td>
      </tr>`
    : `<tr><td colspan="3" style="text-align:center;color:#64748b;">Нет данных</td></tr>`;
}

// ====================== ВЫЧИСЛЕНИЕ ПРОДЛЕНИЯ ======================
function calcRenewal(rows, p1s, p1e, p2s, p2e, includedQ, includedU, gracePeriodDays, includedRemote) {
  if (COL.END === undefined) return null;
  const GRACE_MS = gracePeriodDays * 24 * 60 * 60 * 1000;

  const setINN_P1 = new Set();
  const innP1LastEnd = {};
  const rowsP1 = [];

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const qval = String(row[COL.Q] || "").trim();
    if (qval && !includedQ.has(qval)) continue;
    const uval = String(row[COL.U] || "").trim();
    if (uval && !includedU.has(uval)) continue;
    if (includedRemote && includedRemote.size > 0 && COL.REMOTE !== undefined) {
      const remoteVal = String(row[COL.REMOTE] || "").trim();
      if (remoteVal && !includedRemote.has(remoteVal)) continue;
    }

    const inn = String(row[COL.F] || "").trim().toUpperCase();
    if (!inn) continue;
    const date = parseDate(row[COL.V]);
    if (!date || date < p1s || date > p1e) continue;

    setINN_P1.add(inn);
    rowsP1.push({ inn, rowIndex: i });

    const endDate = parseDate(row[COL.END]);
    if (endDate) {
      if (!innP1LastEnd[inn] || endDate > innP1LastEnd[inn]) {
        innP1LastEnd[inn] = endDate;
      }
    }
  }

  const innNextStart = {};
  const innNextRowIndex = {};

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const inn = String(row[COL.F] || "").trim().toUpperCase();
    if (!inn || !setINN_P1.has(inn)) continue;
    const date = parseDate(row[COL.V]);
    if (!date || date <= p1e) continue;

    if (!innNextStart[inn] || date < innNextStart[inn]) {
      innNextStart[inn] = date;
      innNextRowIndex[inn] = i;
    }
  }

  const renewedINN = new Set();
  const retainedINN = new Set();
  const lapsedINN = new Set();

  for (const inn of setINN_P1) {
    const prevEnd = innP1LastEnd[inn];
    const nextStart = innNextStart[inn];

    if (!nextStart) {
      lapsedINN.add(inn);
      continue;
    }
    if (!prevEnd) {
      retainedINN.add(inn);
      continue;
    }

    if (nextStart <= prevEnd) {
      renewedINN.add(inn);
    } else if (nextStart.getTime() <= prevEnd.getTime() + GRACE_MS) {
      retainedINN.add(inn);
    } else {
      lapsedINN.add(inn);
    }
  }

  const renewedRowIndexes = new Set();
  const retainedRowIndexes = new Set();
  const lapsedRowIndexes = new Set();

  for (const inn of renewedINN) if (innNextRowIndex[inn] !== undefined) renewedRowIndexes.add(innNextRowIndex[inn]);
  for (const inn of retainedINN) if (innNextRowIndex[inn] !== undefined) retainedRowIndexes.add(innNextRowIndex[inn]);

  const seenLapsed = new Set();
  for (let k = rowsP1.length - 1; k >= 0; k--) {
    const { inn, rowIndex } = rowsP1[k];
    if (lapsedINN.has(inn) && !seenLapsed.has(inn)) {
      lapsedRowIndexes.add(rowIndex);
      seenLapsed.add(inn);
    }
  }

  const denominator = setINN_P1.size;
  const renewalCount = renewedINN.size;
  const retainedCount = retainedINN.size;
  const lapsedCount = lapsedINN.size;

  const renewalPct = denominator > 0 ? (renewalCount / denominator * 100).toFixed(2) : "0";
  const retainedPct = denominator > 0 ? (retainedCount / denominator * 100).toFixed(2) : "0";
  const lapsedPct = denominator > 0 ? (lapsedCount / denominator * 100).toFixed(2) : "0";

  return {
    renewalCount, retainedCount, lapsedCount, denominator,
    renewalPct, retainedPct, lapsedPct,
    renewedRowIndexes, retainedRowIndexes, lapsedRowIndexes
  };
}

// ====================== СКАЧИВАНИЕ ======================
function downloadExport(type) {
  if (!lastExportData) return;
  const { renewal, header } = lastExportData;
  let wsData = [header];
  let filename = "";

  if (type === "renewed" && renewal) {
    for (let i = 1; i < allData.length; i++) {
      if (renewal.renewedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "продлились.xlsx";
  } else if (type === "retained" && renewal) {
    for (let i = 1; i < allData.length; i++) {
      if (renewal.retainedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "удержались.xlsx";
  } else if (type === "lapsed" && renewal) {
    for (let i = 1; i < allData.length; i++) {
      if (renewal.lapsedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "отвалились.xlsx";
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), "Лист1");
  XLSX.writeFile(wb, filename);
}

// ====================== ОСНОВНОЙ АНАЛИЗ ======================
async function analyzeFiles() {
  if (!allData.length) return alert("Сначала загрузите файл!");

  const p1s = new Date(document.getElementById("p1-start").value);
  const p1e = new Date(document.getElementById("p1-end").value);
  const p2s = new Date(document.getElementById("p2-start").value);
  const p2e = new Date(document.getElementById("p2-end").value);
  p1e.setHours(23, 59, 59, 999);
  p2e.setHours(23, 59, 59, 999);

  if (isNaN(p1s) || isNaN(p1e) || isNaN(p2s) || isNaN(p2e)) {
    return alert("Укажите корректные даты для обоих периодов!");
  }

  const gracePeriodDays = getGracePeriodDays();

  const includedQ = new Set();
  document.querySelectorAll("#q-filters input:checked").forEach(cb => includedQ.add(cb.value));

  const includedU = new Set();
  document.querySelectorAll("#u-main input:checked, #u-kcr input:checked").forEach(cb => includedU.add(cb.value));

  const includedRemote = new Set();
  document.querySelectorAll("#remote-filters input:checked").forEach(cb => includedRemote.add(cb.value));

  // Определяем все возможные значения удалённой схемы для разбивки
  const allRemoteValues = new Set();
  if (COL.REMOTE !== undefined) {
    for (let i = 1; i < allData.length; i++) {
      const v = String(allData[i][COL.REMOTE] || "").trim();
      if (v) allRemoteValues.add(v);
    }
  }
  const showRemoteBreakdown = includedRemote.size >= 2 && allRemoteValues.size >= 2;

  // Расчёт по СНИЛС
  let setJ1 = new Set(), setJ2 = new Set();
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const qval = COL.Q !== undefined ? String(row[COL.Q] || "").trim() : "";
    const uval = COL.U !== undefined ? String(row[COL.U] || "").trim() : "";
    const remoteVal = COL.REMOTE !== undefined ? String(row[COL.REMOTE] || "").trim() : "";

    if ((qval && !includedQ.has(qval)) || 
        (uval && !includedU.has(uval)) ||
        (includedRemote.size > 0 && remoteVal && !includedRemote.has(remoteVal))) continue;

    const date = parseDate(row[COL.V]);
    if (!date) continue;

    const jval = COL.J !== undefined ? String(row[COL.J] || "").trim().toUpperCase() : "";
    if (!jval) continue;

    if (date >= p1s && date <= p1e) setJ1.add(jval);
    else if (date >= p2s && date <= p2e) setJ2.add(jval);
  }

  const matchJ = [...setJ1].filter(snils => setJ2.has(snils)).length;
  const convJ = setJ1.size ? (matchJ / setJ1.size * 100).toFixed(2) : "0";

  // Расчёт по ИНН (основной — с применением фильтра удалённой схемы)
  const renewal = calcRenewal(allData, p1s, p1e, p2s, p2e, includedQ, includedU, gracePeriodDays, includedRemote);
  lastExportData = { renewal, header: allData[0] };

  // Разбивка по значениям удалённой схемы (если выбрано несколько)
  let remoteBreakdownHTML = "";
  if (showRemoteBreakdown) {
    const breakdownItems = [];
    for (const remVal of Array.from(includedRemote).sort()) {
      const singleSet = new Set([remVal]);
      const r = calcRenewal(allData, p1s, p1e, p2s, p2e, includedQ, includedU, gracePeriodDays, singleSet);
      if (r) breakdownItems.push({ label: remVal, r });
    }
    if (breakdownItems.length > 0) {
      remoteBreakdownHTML = `
        <div style="margin-top:40px; border-top:2px solid #e2e8f0; padding-top:30px;">
          <h3 style="text-align:center; margin-bottom:20px;">📊 Разбивка по «Удаленная схема»</h3>
          <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:30px;">
            ${breakdownItems.map(({ label, r }) => `
              <div style="min-width:360px;">
                <h4 style="text-align:center; margin-bottom:10px; color:#0f766e; background:#f0fdfa; padding:8px 16px; border-radius:10px;">
                  Удаленная схема: <strong>${label}</strong>
                </h4>
                <table class="result-table">
                  <tr><th>Метрика</th><th style="text-align:right">Кол-во</th><th style="text-align:right">%</th></tr>
                  <tr><td>Уникальных ИНН в П1</td><td style="text-align:right">${r.denominator.toLocaleString('ru-RU')}</td><td>—</td></tr>
                  <tr style="background:#faf5ff;color:#6d28d9"><td>🔄 Продлились</td><td style="text-align:right">${r.renewalCount}</td><td>${r.renewalPct}%</td></tr>
                  <tr style="background:#f0fdf4;color:#166534"><td>✅ Удержались</td><td style="text-align:right">${r.retainedCount}</td><td>${r.retainedPct}%</td></tr>
                  <tr style="background:#fff1f2;color:#b91c1c"><td>❌ Отвалились</td><td style="text-align:right">${r.lapsedCount}</td><td>${r.lapsedPct}%</td></tr>
                </table>
              </div>`).join("")}
          </div>
        </div>`;
    }
  }

  const graceLabel = gracePeriodDays === 90 ? "квартал (90 дней)" :
                     gracePeriodDays === 365 ? "год (365 дней)" : 
                     `${gracePeriodDays} дней`;

  const renewalBlock = renewal ? `
    <div style="min-width:460px;">
      <h3 style="text-align:center; color:#7c3aed;">По ИНН — Продление / Удержание / Отвал</h3>
      <table class="result-table">
        <tr><th>Метрика</th><th style="text-align:right">Кол-во</th><th style="text-align:right">%</th></tr>
        <tr><td>Уникальных ИНН в П1</td><td style="text-align:right">${renewal.denominator.toLocaleString('ru-RU')}</td><td>—</td></tr>
        <tr style="background:#faf5ff;color:#6d28d9"><td>🔄 Продлились</td><td style="text-align:right">${renewal.renewalCount}</td><td>${renewal.renewalPct}%</td></tr>
        <tr style="background:#f0fdf4;color:#166534"><td>✅ Удержались</td><td style="text-align:right">${renewal.retainedCount}</td><td>${renewal.retainedPct}%</td></tr>
        <tr style="background:#fff1f2;color:#b91c1c"><td>❌ Отвалились</td><td style="text-align:right">${renewal.lapsedCount}</td><td>${renewal.lapsedPct}%</td></tr>
      </table>
    </div>` : `<p style="color:#ef4444;">Колонка «Дата окончания» не найдена</p>`;

  const exportButtons = renewal ? `
    <div style="margin-top:40px; border-top:2px solid #e2e8f0; padding-top:30px;">
      <h3 style="margin-bottom:16px;">📥 Скачать выгрузки</h3>
      <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:14px;">
        <button onclick="downloadExport('renewed')" style="background:linear-gradient(90deg,#7c3aed,#a78bfa); padding:12px 24px;">🔄 Продлились (${renewal.renewalCount})</button>
        <button onclick="downloadExport('retained')" style="background:linear-gradient(90deg,#166534,#22c55e); padding:12px 24px;">✅ Удержались (${renewal.retainedCount})</button>
        <button onclick="downloadExport('lapsed')" style="background:linear-gradient(90deg,#b91c1c,#ef4444); padding:12px 24px;">❌ Отвалились (${renewal.lapsedCount})</button>
      </div>
    </div>` : "";

  const resultDiv = document.getElementById("result");
  resultDiv.innerHTML = `
    <div class="center">
      <h2>✅ Результат сравнения</h2>
      <p style="font-size:1.25em; margin:25px 0;">
        <strong>Период 1:</strong> ${p1s.toLocaleDateString('ru-RU')} — ${p1e.toLocaleDateString('ru-RU')}<br>
        <strong>Период 2:</strong> ${p2s.toLocaleDateString('ru-RU')} — ${p2e.toLocaleDateString('ru-RU')}
      </p>
      <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:40px; margin-top:30px;">
        <div style="min-width:420px;">
          <h3 style="text-align:center; color:#1e40af;">По СНИЛС — Удержание</h3>
          <table class="result-table">
            <tr><th>Метрика</th><th style="text-align:right">Значение</th></tr>
            <tr><td>Уникальных СНИЛС в П1</td><td>${setJ1.size.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Совпадений в П2</td><td>${matchJ.toLocaleString('ru-RU')}</td></tr>
            <tr style="background:#f0fdf4;color:#166534"><td><strong>% Удержания</strong></td><td><strong>${convJ}%</strong></td></tr>
          </table>
        </div>
        ${renewalBlock}
      </div>
      ${remoteBreakdownHTML}
      ${exportButtons}
    </div>`;

  resultDiv.style.display = "block";
  resultDiv.scrollIntoView({ behavior: "smooth" });
}