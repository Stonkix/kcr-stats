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
  if (typeof val === "string") {
    const s = val.trim();
    if (!s) return null;

    // dd.mm.yyyy (or dd/mm/yyyy), optionally with time
    const ruMatch = s.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{2,4})(?:\s+.*)?$/);
    if (ruMatch) {
      const day = Number(ruMatch[1]);
      const month = Number(ruMatch[2]) - 1;
      let year = Number(ruMatch[3]);
      if (year < 100) year += 2000;
      const d = new Date(year, month, day);
      if (!isNaN(d.getTime())) return d;
    }

    // Fallback for ISO-like strings
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) return parsed;
  }
  return null;
}

function normalizeText(val) {
  return String(val ?? "")
    .replace(/[\u00A0\u2007\u202F]/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeInn(val) {
  const raw = normalizeText(val).replace(/\s+/g, "");
  if (!raw) return "";

  const noDecimalTail = raw.replace(/[.,]0+$/, "");
  const digits = noDecimalTail.replace(/\D/g, "");

  if (digits.length === 10 || digits.length === 12) return digits;
  return digits || noDecimalTail.toUpperCase();
}

function rowPassesSelectedFilters(row, includedQ, includedU, includedRemote) {
  const passesSetFilter = (value, selectedSet, allowEmpty) => {
    if (!selectedSet) return true;
    if (selectedSet.size === 0) return false;
    if (!value) return !!allowEmpty;
    return selectedSet.has(value);
  };

  if (COL.Q !== undefined && includedQ) {
    const qval = normalizeText(row[COL.Q]);
    if (!passesSetFilter(qval, includedQ, true)) return false;
  }

  if (COL.U !== undefined && includedU) {
    const uval = normalizeText(row[COL.U]);
    if (!passesSetFilter(uval, includedU, true)) return false;
  }

  if (COL.REMOTE !== undefined && includedRemote) {
    const remoteVal = normalizeText(row[COL.REMOTE]);
    if (!passesSetFilter(remoteVal, includedRemote, true)) return false;
  }

  return true;
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
        const v = normalizeText(row[COL.Q]);
        if (v) tempQ.add(v);
      }
      if (COL.U !== undefined && row.length > COL.U) {
        const v = normalizeText(row[COL.U]);
        if (v) tempU.add(v);
      }
      if (COL.AA !== undefined && row.length > COL.AA) {
        const v = normalizeText(row[COL.AA]);
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
      .map(row => COL.AA !== undefined && row.length > COL.AA ? normalizeText(row[COL.AA]) : null)
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
      const v = normalizeText(row[remoteCol]);
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
function calcRenewal(
  rows,
  p1s,
  p1e,
  p2s,
  p2e,
  includedQ,
  includedU,
  gracePeriodDays,
  includedRemote
) {
  if (COL.END === undefined) return null;
  const GRACE_MS = gracePeriodDays * 24 * 60 * 60 * 1000;
  const fmt = d => d ? d.toLocaleDateString('ru-RU') : '—';

  // --- ШАГ 1: собираем ИНН из выбранного периода (по дате ОКОНЧАНИЯ серта) ---
  const setINN_K1 = new Set();
  const innK1LastEnd = {};   // последняя дата окончания в периоде для каждого ИНН
  const rowsK1 = [];         // для экспорта «отвалились»
  let rowsScanned = 0, rowsPassedFilters = 0, rowsNoINN = 0, rowsNoEnd = 0, rowsOutOfK1 = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    rowsScanned++;
    if (!rowPassesSelectedFilters(row, includedQ, includedU, includedRemote)) continue;
    rowsPassedFilters++;

    const inn = normalizeInn(row[COL.F]);
    if (!inn) { rowsNoINN++; continue; }

    const endDate = parseDate(row[COL.END]);
    if (!endDate) { rowsNoEnd++; continue; }

    if (endDate < p1s || endDate > p1e) { rowsOutOfK1++; continue; }

    setINN_K1.add(inn);
    rowsK1.push({ inn, rowIndex: i });
    if (!innK1LastEnd[inn] || endDate > innK1LastEnd[inn]) innK1LastEnd[inn] = endDate;
  }

  // --- ШАГ 2: ищем новый сертификат в выбранном периоде (по дате НАЧАЛА, строго внутри p2s..p2e) ---
  // Важно: окончание и новый старт считаются в одном и том же выбранном периоде
  // Ищем среди ВСЕХ строк (без фильтра по дате окончания) — нужен старт внутри периода
  const innK2Start = {};     // ближайшая дата начала нового серта в периоде
  const innK2RowIndex = {};
  let rowsFoundInK2 = 0;

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (!rowPassesSelectedFilters(row, includedQ, includedU, includedRemote)) continue;

    const inn = normalizeInn(row[COL.F]);
    if (!inn || !setINN_K1.has(inn)) continue;

    const startDate = parseDate(row[COL.V]);
    if (!startDate) continue;
    // Новый сертификат должен начаться внутри выбранного периода
    if (startDate < p2s || startDate > p2e) continue;

    // Берём самый ранний старт в периоде
    if (!innK2Start[inn] || startDate < innK2Start[inn]) {
      innK2Start[inn] = startDate;
      innK2RowIndex[inn] = i;
      rowsFoundInK2++;
    }
  }

  // --- ШАГ 3: классификация ---
  const renewedINN = new Set();
  const retainedINN = new Set();
  const lapsedINN = new Set();
  // Собираем по 10 примеров для каждой категории отдельно
  const debugRenewed = [];
  const debugRetained = [];
  const debugLapsed = [];

  for (const inn of setINN_K1) {
    const prevEnd  = innK1LastEnd[inn];
    const nextStart = innK2Start[inn];

    let category, reason;

    if (!nextStart) {
      lapsedINN.add(inn);
      category = "lapsed";
      reason = `нет нового серта в выбранном периоде (${fmt(p2s)}–${fmt(p2e)})`;
    } else if (!prevEnd) {
      retainedINN.add(inn);
      category = "retained";
      reason = `нет даты окончания в выбранном периоде (нестандартно)`;
    } else if (nextStart <= prevEnd) {
      renewedINN.add(inn);
      category = "renewed";
      reason = `новый старт ${fmt(nextStart)} ≤ конец старого ${fmt(prevEnd)} (купил заранее)`;
    } else if (nextStart.getTime() <= prevEnd.getTime() + GRACE_MS) {
      retainedINN.add(inn);
      category = "retained";
      reason = `новый старт ${fmt(nextStart)}, конец старого ${fmt(prevEnd)}, разрыв ${Math.round((nextStart-prevEnd)/86400000)} дн ≤ ${gracePeriodDays} дн`;
    } else {
      lapsedINN.add(inn);
      category = "lapsed";
      reason = `новый старт ${fmt(nextStart)}, конец старого ${fmt(prevEnd)}, разрыв ${Math.round((nextStart-prevEnd)/86400000)} дн > ${gracePeriodDays} дн`;
    }

    const entry = { inn, reason, prevEnd, nextStart };
    if (category === "renewed"  && debugRenewed.length  < 10) debugRenewed.push(entry);
    if (category === "retained" && debugRetained.length < 10) debugRetained.push(entry);
    if (category === "lapsed"   && debugLapsed.length   < 10) debugLapsed.push(entry);
  }

  // --- ШАГ 4: индексы строк для экспорта ---
  const renewedRowIndexes = new Set();
  const retainedRowIndexes = new Set();
  const lapsedRowIndexes = new Set();

  for (const inn of renewedINN)  if (innK2RowIndex[inn]  !== undefined) renewedRowIndexes.add(innK2RowIndex[inn]);
  for (const inn of retainedINN) if (innK2RowIndex[inn]  !== undefined) retainedRowIndexes.add(innK2RowIndex[inn]);

  const seenLapsed = new Set();
  for (let k = rowsK1.length - 1; k >= 0; k--) {
    const { inn, rowIndex } = rowsK1[k];
    if (lapsedINN.has(inn) && !seenLapsed.has(inn)) {
      lapsedRowIndexes.add(rowIndex);
      seenLapsed.add(inn);
    }
  }

  const denominator   = setINN_K1.size;
  const renewalCount  = renewedINN.size;
  const retainedCount = retainedINN.size;
  const lapsedCount   = lapsedINN.size;

  const renewalPct  = denominator > 0 ? (renewalCount  / denominator * 100).toFixed(2) : "0";
  const retainedPct = denominator > 0 ? (retainedCount / denominator * 100).toFixed(2) : "0";
  const lapsedPct   = denominator > 0 ? (lapsedCount   / denominator * 100).toFixed(2) : "0";

  // --- Диагностический лог ---
  const debugLog = {
    rowsScanned,
    rowsPassedFilters,
    rowsNoINN,
    rowsNoEnd,
    rowsOutOfK1,
    totalInK1: setINN_K1.size,
    rowsFoundInK2,
    uniqueInnWithK2: Object.keys(innK2Start).length,
    debugRenewed,
    debugRetained,
    debugLapsed
  };

  return {
    renewalCount, retainedCount, lapsedCount, denominator,
    renewalPct, retainedPct, lapsedPct,
    renewedRowIndexes, retainedRowIndexes, lapsedRowIndexes,
    debugLog
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
  const p2s = new Date(p1s);
  const p2e = new Date(p1e);
  p1e.setHours(23, 59, 59, 999);
  p2e.setHours(23, 59, 59, 999);

  if (isNaN(p1s) || isNaN(p1e)) {
    return alert("Укажите корректные даты периода!");
  }

  const gracePeriodDays = getGracePeriodDays();

  const includedQ = new Set();
  document.querySelectorAll("#q-filters input:checked").forEach(cb => includedQ.add(normalizeText(cb.value)));

  const includedU = new Set();
  document.querySelectorAll("#u-main input:checked, #u-kcr input:checked").forEach(cb => includedU.add(normalizeText(cb.value)));

  const includedRemote = new Set();
  document.querySelectorAll("#remote-filters input:checked").forEach(cb => includedRemote.add(normalizeText(cb.value)));

  const allRemoteValues = new Set();
  if (COL.REMOTE !== undefined) {
    for (let i = 1; i < allData.length; i++) {
      const v = normalizeText(allData[i][COL.REMOTE]);
      if (v) allRemoteValues.add(v);
    }
  }
  const showRemoteBreakdown = includedRemote.size >= 2 && allRemoteValues.size >= 2;

  // --- Расчёт по СНИЛС ---
  // Серт с окончанием: дата ОКОНЧАНИЯ попадает в выбранный период
  // Новый серт: дата НАЧАЛА попадает в тот же выбранный период
  let setJ1 = new Set(), setJ2 = new Set();
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (!rowPassesSelectedFilters(row, includedQ, includedU, includedRemote)) continue;

    const jval = COL.J !== undefined ? normalizeText(row[COL.J]).toUpperCase() : "";
    if (!jval) continue;

    if (COL.END !== undefined) {
      const endDate = parseDate(row[COL.END]);
      if (endDate && endDate >= p1s && endDate <= p1e) setJ1.add(jval);
    }
    const startDate = parseDate(row[COL.V]);
    if (startDate && startDate >= p2s && startDate <= p2e) setJ2.add(jval);
  }

  const matchJ = [...setJ1].filter(snils => setJ2.has(snils)).length;
  const convJ = setJ1.size ? (matchJ / setJ1.size * 100).toFixed(2) : "0";

  // --- Расчёт по ИНН ---
  const renewal = calcRenewal(allData, p1s, p1e, p2s, p2e, includedQ, includedU, gracePeriodDays, includedRemote);
  lastExportData = { renewal, header: allData[0] };

  // --- Разбивка по удалённой схеме ---
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
                  <tr><td>ИНН с окончанием в периоде</td><td style="text-align:right">${r.denominator.toLocaleString('ru-RU')}</td><td>—</td></tr>
                  <tr style="background:#faf5ff;color:#6d28d9"><td>🔄 Продлились</td><td style="text-align:right">${r.renewalCount}</td><td>${r.renewalPct}%</td></tr>
                  <tr style="background:#f0fdf4;color:#166534"><td>✅ Удержались</td><td style="text-align:right">${r.retainedCount}</td><td>${r.retainedPct}%</td></tr>
                  <tr style="background:#fff1f2;color:#b91c1c"><td>❌ Отвалились</td><td style="text-align:right">${r.lapsedCount}</td><td>${r.lapsedPct}%</td></tr>
                </table>
              </div>`).join("")}
          </div>
        </div>`;
    }
  }

  const renewalBlock = renewal ? `
    <div style="min-width:460px;">
      <h3 style="text-align:center; color:#7c3aed;">По ИНН — Продление / Удержание / Отвал</h3>
      <table class="result-table">
        <tr><th>Метрика</th><th style="text-align:right">Кол-во</th><th style="text-align:right">%</th></tr>
        <tr><td>ИНН с окончанием серта в периоде</td><td style="text-align:right">${renewal.denominator.toLocaleString('ru-RU')}</td><td>—</td></tr>
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

  // --- Блок диагностического лога ---
  let debugHTML = "";
  if (renewal && renewal.debugLog) {
    const dl = renewal.debugLog;

    const makeTable = (rows) => {
      if (!rows.length) return `<p style="color:#94a3b8; font-style:italic;">Нет записей</p>`;
      return `<table class="result-table" style="font-size:0.85em;">
        <tr><th>ИНН</th><th>Оконч. в периоде</th><th>Новый старт в периоде</th><th>Причина</th></tr>
        ${rows.map(r => `
          <tr>
            <td style="font-family:monospace">${r.inn}</td>
            <td>${r.prevEnd ? r.prevEnd.toLocaleDateString('ru-RU') : '—'}</td>
            <td>${r.nextStart ? r.nextStart.toLocaleDateString('ru-RU') : '—'}</td>
            <td style="color:#64748b; font-size:0.9em;">${r.reason}</td>
          </tr>`).join("")}
      </table>`;
    };

    debugHTML = `
      <div style="margin-top:40px; border-top:2px dashed #e2e8f0; padding-top:30px;">
        <details>
          <summary style="cursor:pointer; font-size:1.1em; font-weight:700; color:#475569; padding:12px 16px; background:#f8fafc; border-radius:12px; border:1px solid #e2e8f0; list-style:none; display:flex; align-items:center; gap:8px;">
            🔍 Диагностический лог (что именно нашла программа)
          </summary>
          <div style="margin-top:16px; display:flex; flex-direction:column; gap:20px;">

            <div style="background:#f1f5f9; border-radius:12px; padding:16px 20px; font-size:0.9em; line-height:1.9;">
              <strong>📋 Общая статистика обработки строк:</strong><br>
              Всего строк в файле: <b>${dl.rowsScanned.toLocaleString('ru-RU')}</b><br>
              Прошли фильтры (тариф / спецпред / удал.схема): <b>${dl.rowsPassedFilters.toLocaleString('ru-RU')}</b><br>
              Пропущено — нет ИНН: <b>${dl.rowsNoINN}</b><br>
              Пропущено — нет даты окончания: <b>${dl.rowsNoEnd}</b><br>
              Пропущено — дата окончания вне периода: <b>${dl.rowsOutOfK1.toLocaleString('ru-RU')}</b><br>
              <b>→ Уникальных ИНН попало в период: ${dl.totalInK1.toLocaleString('ru-RU')}</b><br>
              <hr style="border:none; border-top:1px solid #cbd5e1; margin:8px 0;">
              Уникальных ИНН у которых найден новый серт в периоде: <b>${dl.uniqueInnWithK2.toLocaleString('ru-RU')}</b><br>
              Уникальных ИНН без нового серта в периоде: <b>${(dl.totalInK1 - dl.uniqueInnWithK2).toLocaleString('ru-RU')}</b>
            </div>

            <div>
              <h4 style="color:#6d28d9; margin:0 0 8px 0;">🔄 Продлились — до 10 примеров (найдено: ${dl.debugRenewed.length})</h4>
              ${makeTable(dl.debugRenewed)}
            </div>

            <div>
              <h4 style="color:#166534; margin:0 0 8px 0;">✅ Удержались — до 10 примеров (найдено: ${dl.debugRetained.length})</h4>
              ${makeTable(dl.debugRetained)}
            </div>

            <div>
              <h4 style="color:#b91c1c; margin:0 0 8px 0;">❌ Отвалились — до 10 примеров (найдено: ${dl.debugLapsed.length})</h4>
              ${makeTable(dl.debugLapsed)}
            </div>

          </div>
        </details>
      </div>`;
  }

  const resultDiv = document.getElementById("result");
  resultDiv.innerHTML = `
    <div class="center">
      <h2>✅ Результат анализа</h2>
      <p style="font-size:1.1em; margin:20px 0; background:#f0f9ff; padding:14px 20px; border-radius:12px; display:inline-block; text-align:left; line-height:1.8;">
        <strong>Выбранный период</strong>: ${p1s.toLocaleDateString('ru-RU')} — ${p1e.toLocaleDateString('ru-RU')}
      </p>
      <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:40px; margin-top:30px;">
        <div style="min-width:420px;">
          <h3 style="text-align:center; color:#1e40af;">По СНИЛС — Удержание</h3>
          <table class="result-table">
            <tr><th>Метрика</th><th style="text-align:right">Значение</th></tr>
            <tr><td>СНИЛС с окончанием серта в периоде</td><td>${setJ1.size.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Из них купили серт в периоде</td><td>${matchJ.toLocaleString('ru-RU')}</td></tr>
            <tr style="background:#f0fdf4;color:#166534"><td><strong>% Удержания</strong></td><td><strong>${convJ}%</strong></td></tr>
          </table>
        </div>
        ${renewalBlock}
      </div>
      ${remoteBreakdownHTML}
      ${exportButtons}
      ${debugHTML}
    </div>`;

  resultDiv.style.display = "block";
  resultDiv.scrollIntoView({ behavior: "smooth" });
}
