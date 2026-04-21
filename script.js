// ====================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ======================
let allData = [];
let uniqueTariffs = [];
let uniqueUValues = [];
let newPerYearGlobal = {};
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
  };
  const result = {};
  headerRow.forEach((cell, idx) => {
    const key = String(cell || "").trim().toLowerCase();
    if (MAP[key]) result[MAP[key]] = idx;
  });
  const required = ["V", "END", "AA", "U", "Q", "J", "F"];
  const missing = required.filter(k => result[k] === undefined);
  if (missing.length) console.warn("Не найдены колонки:", missing, "в заголовке:", headerRow);
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
    const innByYear = {};
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
      if (COL.F !== undefined && COL.V !== undefined && row.length > Math.max(COL.F, COL.V)) {
        const inn = String(row[COL.F] || "").trim().toUpperCase();
        const date = parseDate(row[COL.V]);
        if (inn && date) {
          const year = date.getFullYear();
          if (!innByYear[year]) innByYear[year] = new Set();
          innByYear[year].add(inn);
        }
      }
    }
    uniqueTariffs = Array.from(tempQ).sort((a, b) => a.localeCompare(b, 'ru'));
    uniqueUValues = Array.from(tempU).sort((a, b) => a.localeCompare(b, 'ru'));
    const uniqueAA = Array.from(tempAA).sort((a, b) => a.localeCompare(b, 'ru'));
    newPerYearGlobal = {};
    const cumulative = new Set();
    Object.keys(innByYear).sort((a,b)=>Number(a)-Number(b)).forEach(year => {
      let count = 0;
      innByYear[year].forEach(inn => {
        if (!cumulative.has(inn)) { count++; cumulative.add(inn); }
      });
      newPerYearGlobal[year] = count;
    });
    populateFilters(uniqueAA);
    populateGeneralStats();
    document.getElementById("loaded-config").style.display = "block";
    document.getElementById("analyzeBtn").disabled = false;
    statusEl.innerHTML = `✅ <strong>${file.name}</strong><br><small>Строк данных: ${(allData.length - 1).toLocaleString('ru-RU')}</small>`;
  } catch (e) {
    statusEl.innerHTML = `❌ Ошибка: ${e.message}`;
    console.error(e);
  }
}

function populateFilters(uniqueAA) {
  document.getElementById("q-filters").innerHTML = uniqueTariffs.map(val =>
    `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
  ).join("");
  const main = [], kcr = [];
  uniqueUValues.forEach(val => {
    val.toUpperCase().includes("KCR") ? kcr.push(val) : main.push(val);
  });
  document.getElementById("u-main").innerHTML = main.map(val =>
    `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
  ).join("");
  document.getElementById("u-kcr").innerHTML = kcr.map(val =>
    `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
  ).join("");
  const aaContainer = document.getElementById("aa-filter-container");
  if (aaContainer && uniqueAA && uniqueAA.length) {
    aaContainer.innerHTML = `
      <div style="display:flex; justify-content:space-between; align-items:center; margin-bottom:10px;">
        <h4 style="margin:0;">Удостоверяющий центр <small style="font-weight:400; color:#64748b;">(% считается только по отмеченным)</small></h4>
        <div style="display:flex; gap:8px;">
          <button type="button" onclick="selectAllAA(true)" style="font-size:0.8em; padding:4px 10px; background:#3b82f6; color:white; border:none; border-radius:6px; cursor:pointer; box-shadow:none; transform:none;">Все</button>
          <button type="button" onclick="selectAllAA(false)" style="font-size:0.8em; padding:4px 10px; background:#94a3b8; color:white; border:none; border-radius:6px; cursor:pointer; box-shadow:none; transform:none;">Ни одного</button>
        </div>
      </div>
      <div id="aa-filters" class="checkbox-group">
        ${uniqueAA.map(val =>
          `<label><input type="checkbox" value="${val.replace(/"/g, "&quot;")}" checked> ${val}</label>`
        ).join("")}
      </div>`;
  }
}

function selectAllAA(checked) {
  document.querySelectorAll("#aa-filters input[type='checkbox']").forEach(cb => cb.checked = checked);
}

function selectAllQ(checked) {
  document.querySelectorAll("#q-filters input[type='checkbox']").forEach(cb => cb.checked = checked);
  if (checked) {
    kcrEnabled = true;
    const btn = document.getElementById("kcr-button");
    btn.textContent = "Убрать КЦР";
    btn.style.background = "#ef4444";
  }
}

function selectAllU(checked) {
  document.querySelectorAll("#u-main input[type='checkbox'], #u-kcr input[type='checkbox']").forEach(cb => cb.checked = checked);
}

function toggleKCR() {
  kcrEnabled = !kcrEnabled;
  const btn = document.getElementById("kcr-button");
  btn.textContent = kcrEnabled ? "Убрать КЦР" : "Вернуть КЦР";
  btn.style.background = kcrEnabled ? "#ef4444" : "#22c55e";
  document.querySelectorAll("#q-filters input[type='checkbox']").forEach(cb => {
    if (cb.value.toUpperCase().includes("КЦР") || cb.value.toUpperCase().includes("KCR")) cb.checked = kcrEnabled;
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

// ====================== СТАТИСТИКА AA ======================
function populateGeneralStats() {
  const html = `
    <h3 style="margin-bottom:10px;">Статистика по удостоверяющим центрам</h3>
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
      if (d) { years.add(d.getFullYear()); months.add(d.getMonth() + 1); }
    }
  }
  Array.from(years).sort((a,b)=>b-a).forEach(y => yearSelect.appendChild(new Option(y, y)));
  const monthNames = ["","Январь","Февраль","Март","Апрель","Май","Июнь","Июль","Август","Сентябрь","Октябрь","Ноябрь","Декабрь"];
  Array.from(months).sort((a,b)=>a-b).forEach(m => monthSelect.appendChild(new Option(monthNames[m], m)));
}

function filterAAStats() {
  const yearF = document.getElementById("aa-year-filter").value;
  const monthF = document.getElementById("aa-month-filter").value;
  const tbody = document.getElementById("aa-tbody");
  const countMap = {};
  if (COL.AA === undefined || COL.V === undefined) {
    tbody.innerHTML = `<tr><td colspan="3" style="text-align:center;color:#ef4444;">Колонки не определены</td></tr>`;
    return;
  }
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row.length <= COL.AA) continue;
    const center = String(row[COL.AA] || "").trim();
    if (!center || center === EXCLUDED_UC) continue;
    const date = parseDate(row[COL.V]);
    if (!date) continue;
    if (yearF && date.getFullYear() != yearF) continue;
    if (monthF && (date.getMonth()+1) != monthF) continue;
    countMap[center] = (countMap[center] || 0) + 1;
  }
  const sorted = Object.entries(countMap).sort((a,b) => b[1] - a[1]);
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
//
// ПРАВИЛЬНАЯ ЛОГИКА:
//
// Шаг 1. Находим все ИНН, у которых ЕСТЬ сертификат в Периоде 1.
//         Для каждого такого ИНН берём максимальную дату окончания из П1 = prevEnd.
//
// Шаг 2. По всей базе (не ограничиваясь П2!) ищем следующий сертификат этого ИНН,
//         выпущенный СТРОГО ПОСЛЕ конца П1 (date > p1e). Берём самый ранний = nextStart.
//         (Фильтры Q/U не применяем — интересует сам факт любого нового сертификата)
//
// Шаг 3. Классифицируем:
//   - Нет nextStart                              → Отвалился
//   - nextStart <= prevEnd                       → Продлился  (купил до окончания)
//   - nextStart <= prevEnd + gracePeriodDays     → Удержался  (купил в грейс-период)
//   - nextStart >  prevEnd + gracePeriodDays     → Отвалился  (вернулся, но слишком поздно)
//
function calcRenewal(rows, p1s, p1e, p2s, p2e, includedQ, includedU, gracePeriodDays) {
  if (COL.END === undefined) return null;
  const GRACE_MS = gracePeriodDays * 24 * 60 * 60 * 1000;

  // --- Шаг 1: собираем ИНН из П1 (с фильтрами) + их максимальную дату окончания ---
  const setINN_P1 = new Set();
  const rowsP1 = [];
  const innP1LastEnd = {}; // ИНН → максимальная дата окончания сертификата из П1

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const qval = String(row[COL.Q] || "").trim();
    if (qval && !includedQ.has(qval)) continue;
    const uval = String(row[COL.U] || "").trim();
    if (uval && !includedU.has(uval)) continue;
    const inn = String(row[COL.F] || "").trim().toUpperCase();
    if (!inn) continue;
    const date = parseDate(row[COL.V]);
    if (!date) continue;

    if (date >= p1s && date <= p1e) {
      setINN_P1.add(inn);
      rowsP1.push({ inn, rowIndex: i });
      const endDate = parseDate(row[COL.END]);
      if (endDate) {
        if (!innP1LastEnd[inn] || endDate > innP1LastEnd[inn]) {
          innP1LastEnd[inn] = endDate;
        }
      }
    }
  }

  // --- Шаг 2: по ВСЕЙ базе ищем следующий сертификат после П1 (БЕЗ фильтров Q/U) ---
  // Берём самую раннюю дату начала ПОСЛЕ p1e для каждого ИНН из П1
  const innNextStart = {}; // ИНН → ближайшая дата начала после П1
  const innNextRowIndex = {}; // для скачивания

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const inn = String(row[COL.F] || "").trim().toUpperCase();
    if (!inn || !setINN_P1.has(inn)) continue; // интересуют только ИНН из П1
    const date = parseDate(row[COL.V]);
    if (!date) continue;
    if (date <= p1e) continue; // только сертификаты ПОСЛЕ конца П1

    if (!innNextStart[inn] || date < innNextStart[inn]) {
      innNextStart[inn] = date;
      innNextRowIndex[inn] = i;
    }
  }

  // --- Шаг 3: классификация ---
  const renewedINN = new Set();
  const retainedINN = new Set();
  const lapsedINN = new Set();

  for (const inn of setINN_P1) {
    const prevEnd = innP1LastEnd[inn];
    const nextStart = innNextStart[inn];

    if (!nextStart) {
      // Вообще не вернулся
      lapsedINN.add(inn);
      continue;
    }

    if (!prevEnd) {
      // Нет даты окончания в П1 — по факту возврата считаем удержанием
      retainedINN.add(inn);
      continue;
    }

    if (nextStart <= prevEnd) {
      renewedINN.add(inn);   // Продлился: новый серт до окончания предыдущего
    } else if (nextStart.getTime() <= prevEnd.getTime() + GRACE_MS) {
      retainedINN.add(inn);  // Удержался: вернулся в грейс-период
    } else {
      lapsedINN.add(inn);    // Отвалился: вернулся, но слишком поздно
    }
  }

  // --- Строки для скачивания ---
  // Продлились / удержались → строка с nextStart по этому ИНН
  const renewedRowIndexes = new Set();
  const retainedRowIndexes = new Set();
  for (const inn of renewedINN) {
    if (innNextRowIndex[inn] !== undefined) renewedRowIndexes.add(innNextRowIndex[inn]);
  }
  for (const inn of retainedINN) {
    if (innNextRowIndex[inn] !== undefined) retainedRowIndexes.add(innNextRowIndex[inn]);
  }

  // Отвалились → последняя строка этого ИНН в П1
  const lapsedRowIndexes = new Set();
  const seenLapsed = new Set();
  for (let k = rowsP1.length - 1; k >= 0; k--) {
    const { inn, rowIndex } = rowsP1[k];
    if (lapsedINN.has(inn) && !seenLapsed.has(inn)) {
      lapsedRowIndexes.add(rowIndex);
      seenLapsed.add(inn);
    }
  }

  const denominator = setINN_P1.size;
  const renewalCount  = renewedINN.size;
  const retainedCount = retainedINN.size;
  const lapsedCount   = lapsedINN.size;
  const renewalPct  = denominator > 0 ? (renewalCount  / denominator * 100).toFixed(2) : "0";
  const retainedPct = denominator > 0 ? (retainedCount / denominator * 100).toFixed(2) : "0";
  const lapsedPct   = denominator > 0 ? (lapsedCount   / denominator * 100).toFixed(2) : "0";

  return {
    renewalCount, retainedCount, lapsedCount,
    denominator, renewalPct, retainedPct, lapsedPct,
    renewedINN, renewedRowIndexes,
    retainedINN, retainedRowIndexes,
    lapsedINN, lapsedRowIndexes,
  };
}

// ====================== СКАЧИВАНИЕ XLSX ======================
function downloadExport(type) {
  if (!lastExportData) return;
  const { renewal, header } = lastExportData;
  let wsData = [header];
  let filename = "";
  if (type === "renewed") {
    if (!renewal) return;
    for (let i = 1; i < allData.length; i++) {
      if (renewal.renewedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "продления.xlsx";
  } else if (type === "retained") {
    if (!renewal) return;
    for (let i = 1; i < allData.length; i++) {
      if (renewal.retainedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "удержанные.xlsx";
  } else if (type === "lapsed") {
    if (!renewal) return;
    for (let i = 1; i < allData.length; i++) {
      if (renewal.lapsedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "отвалившиеся.xlsx";
  }
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), "Лист1");
  XLSX.writeFile(wb, filename);
}

// ====================== ТАБЛИЦА УЦ ДЛЯ ПЕРИОДА ======================
function renderAATable(countMap, title, selectedAA) {
  const sorted = Object.entries(countMap).sort((a,b) => b[1] - a[1]);
  const totalAll = sorted.reduce((s,[,c]) => s + c, 0);
  const totalForPct = (selectedAA && selectedAA.size > 0)
    ? sorted.filter(([c]) => selectedAA.has(c)).reduce((s,[,c]) => s + c, 0)
    : totalAll;
  if (!sorted.length) return `<div style="min-width:340px;"><h4 style="text-align:center;color:#64748b;">${title}</h4><p style="text-align:center;color:#94a3b8;">Нет данных</p></div>`;
  const rows = sorted.map(([c, n]) => {
    const isSelected = !selectedAA || selectedAA.size === 0 || selectedAA.has(c);
    const pct = (isSelected && totalForPct > 0) ? (n / totalForPct * 100).toFixed(1) : "—";
    const dim = isSelected ? "" : "opacity:0.4;";
    return `<tr style="${dim}">
      <td>${c}</td>
      <td style="text-align:right">${n.toLocaleString('ru-RU')}</td>
      <td style="text-align:right;color:#1e40af;font-weight:600">${pct !== "—" ? pct + "%" : "—"}</td>
    </tr>`;
  }).join("");
  return `
    <div style="min-width:340px;">
      <h4 style="text-align:center; margin-bottom:10px;">${title}</h4>
      <table class="result-table">
        <thead><tr><th>УЦ</th><th style="text-align:right">Кол-во</th><th style="text-align:right">%</th></tr></thead>
        <tbody>
          ${rows}
          <tr style="background:#f1f5f9;font-weight:700;">
            <td>Итого</td>
            <td style="text-align:right">${totalAll.toLocaleString('ru-RU')}</td>
            <td style="text-align:right">${totalForPct === totalAll ? "100%" : `(база: ${totalForPct.toLocaleString('ru-RU')})`}</td>
          </tr>
        </tbody>
      </table>
    </div>`;
}

// ====================== ОСНОВНОЙ АНАЛИЗ ======================
async function analyzeFiles() {
  if (!allData.length) return alert("Сначала загрузите файл!");
  if (COL.V === undefined || COL.Q === undefined || COL.U === undefined || COL.J === undefined) {
    return alert("Не удалось определить все необходимые колонки из заголовка файла. Проверьте строку 1.");
  }
  const p1s = new Date(document.getElementById("p1-start").value);
  const p1e = new Date(document.getElementById("p1-end").value);
  const p2s = new Date(document.getElementById("p2-start").value);
  const p2e = new Date(document.getElementById("p2-end").value);
  p1e.setHours(23,59,59,999);
  p2e.setHours(23,59,59,999);
  if (isNaN(p1s) || isNaN(p1e) || isNaN(p2s) || isNaN(p2e)) {
    return alert("Укажите корректные даты для обоих периодов!");
  }

  const graceSelect = document.getElementById("grace-period");
  const gracePeriodDays = graceSelect ? parseInt(graceSelect.value) : 90;

  const includedQ = new Set();
  document.querySelectorAll("#q-filters input:checked").forEach(cb => includedQ.add(cb.value));
  const includedU = new Set();
  document.querySelectorAll("#u-main input:checked, #u-kcr input:checked").forEach(cb => includedU.add(cb.value));
  const selectedAA = new Set();
  document.querySelectorAll("#aa-filters input:checked").forEach(cb => selectedAA.add(cb.value));

  // Статистика УЦ по периодам (для отображения таблиц)
  const aaP1 = {}, aaP2 = {};
  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const qval = String(row[COL.Q] || "").trim();
    if (qval && !includedQ.has(qval)) continue;
    const uval = String(row[COL.U] || "").trim();
    if (uval && !includedU.has(uval)) continue;
    const date = parseDate(row[COL.V]);
    if (!date) continue;
    const center = COL.AA !== undefined ? String(row[COL.AA] || "").trim() : "";
    const centerClean = center === EXCLUDED_UC ? "" : center;
    if (date >= p1s && date <= p1e) {
      if (centerClean) aaP1[centerClean] = (aaP1[centerClean] || 0) + 1;
    } else if (date >= p2s && date <= p2e) {
      if (centerClean) aaP2[centerClean] = (aaP2[centerClean] || 0) + 1;
    }
  }

  // --- Расчёт продления (новая логика) ---
  const renewal = calcRenewal(allData, p1s, p1e, p2s, p2e, includedQ, includedU, gracePeriodDays);
  lastExportData = { renewal, header: allData[0] };

  const graceLabel = gracePeriodDays === 90 ? "квартал (90 дней)" : "год (365 дней)";

  const renewalBlock = renewal ? `
    <div style="min-width:460px;">
      <h3 style="text-align:center; color:#7c3aed;">По ИНН — Продление / Удержание / Отвал</h3>
      <p style="text-align:center; font-size:0.88em; color:#64748b; margin-bottom:14px;">
        Поиск следующего сертификата ведётся по <strong>всей базе</strong> после конца Периода 1.<br>
        Грейс-период для удержания: <strong>${graceLabel}</strong> после окончания предыдущего сертификата.
      </p>
      <table class="result-table">
        <tr><th>Метрика</th><th style="text-align:right">Кол-во</th><th style="text-align:right">%</th></tr>
        <tr>
          <td>Уникальных ИНН в Периоде 1 (знаменатель)</td>
          <td style="text-align:right">${renewal.denominator.toLocaleString('ru-RU')}</td>
          <td>—</td>
        </tr>
        <tr style="background:#faf5ff;color:#6d28d9">
          <td>🔄 <strong>Продлились</strong> <small style="font-weight:400">(следующий серт. куплен до окончания предыдущего)</small></td>
          <td style="text-align:right"><strong>${renewal.renewalCount.toLocaleString('ru-RU')}</strong></td>
          <td style="text-align:right"><strong>${renewal.renewalPct}%</strong></td>
        </tr>
        <tr style="background:#f0fdf4;color:#166534">
          <td>✅ <strong>Удержались</strong> <small style="font-weight:400">(вернулись в грейс-период после окончания)</small></td>
          <td style="text-align:right"><strong>${renewal.retainedCount.toLocaleString('ru-RU')}</strong></td>
          <td style="text-align:right"><strong>${renewal.retainedPct}%</strong></td>
        </tr>
        <tr style="background:#fff1f2;color:#b91c1c">
          <td>❌ <strong>Отвалились</strong> <small style="font-weight:400">(не вернулись или вернулись позже грейс-периода)</small></td>
          <td style="text-align:right"><strong>${renewal.lapsedCount.toLocaleString('ru-RU')}</strong></td>
          <td style="text-align:right"><strong>${renewal.lapsedPct}%</strong></td>
        </tr>
      </table>
    </div>`
    : `<p style="color:#ef4444; text-align:center;">⚠️ Колонка «Дата окончания» не найдена — продление не рассчитано</p>`;

  const exportButtons = `
    <div style="margin-top:40px; border-top:2px solid #e2e8f0; padding-top:30px;">
      <h3 style="margin-bottom:16px; color:#374151;">📥 Скачать выгрузки (xlsx)</h3>
      <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:14px;">
        ${renewal ? `<button type="button" onclick="downloadExport('renewed')" style="background:linear-gradient(90deg,#7c3aed,#a78bfa); padding:12px 24px; font-size:1em; box-shadow:0 4px 14px rgba(124,58,237,0.35);">🔄 Продлились (${renewal.renewalCount.toLocaleString('ru-RU')})</button>` : ""}
        ${renewal ? `<button type="button" onclick="downloadExport('retained')" style="background:linear-gradient(90deg,#166534,#22c55e); padding:12px 24px; font-size:1em; box-shadow:0 4px 14px rgba(22,101,52,0.35);">✅ Удержались (${renewal.retainedCount.toLocaleString('ru-RU')})</button>` : ""}
        ${renewal ? `<button type="button" onclick="downloadExport('lapsed')" style="background:linear-gradient(90deg,#b91c1c,#ef4444); padding:12px 24px; font-size:1em; box-shadow:0 4px 14px rgba(185,28,28,0.35);">❌ Отвалились (${renewal.lapsedCount.toLocaleString('ru-RU')})</button>` : ""}
      </div>
      <p style="font-size:0.82em; color:#94a3b8; margin-top:12px; text-align:center;">
        Продлились / Удержались → первый новый сертификат после П1 &nbsp;|&nbsp; Отвалились → последняя строка Периода 1 по ИНН
      </p>
    </div>`;

  const resultDiv = document.getElementById("result");
  resultDiv.innerHTML = `
    <div class="center">
      <h2>✅ Результат сравнения</h2>
      <p style="font-size:1.25em; margin:25px 0;">
        <strong>Период 1:</strong> ${p1s.toLocaleDateString('ru-RU')} — ${p1e.toLocaleDateString('ru-RU')}<br>
        <strong>Период 2:</strong> ${p2s.toLocaleDateString('ru-RU')} — ${p2e.toLocaleDateString('ru-RU')}
      </p>
      <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:40px; margin-top:30px;">
        ${renewalBlock}
      </div>
      <div style="margin-top:40px; border-top:2px solid #e2e8f0; padding-top:30px;">
        <h3 style="margin-bottom:8px;">📊 Удостоверяющие центры по периодам</h3>
        ${selectedAA.size > 0 ? `<p style="text-align:center;font-size:0.88em;color:#3b82f6;margin-bottom:20px;">% считается по выбранным УЦ (${selectedAA.size} шт.)</p>` : ""}
        <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:40px;">
          ${renderAATable(aaP1, `Период 1 (${p1s.toLocaleDateString('ru-RU')} — ${p1e.toLocaleDateString('ru-RU')})`, selectedAA)}
          ${renderAATable(aaP2, `Период 2 (${p2s.toLocaleDateString('ru-RU')} — ${p2e.toLocaleDateString('ru-RU')})`, selectedAA)}
        </div>
      </div>
      ${exportButtons}
    </div>`;
  resultDiv.style.display = "block";
  resultDiv.scrollIntoView({ behavior: "smooth", block: "start" });
}