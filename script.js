// ====================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ======================
let allData = [];
let uniqueTariffs = [];
let uniqueUValues = [];
let newPerYearGlobal = {};
let kcrEnabled = true;

// Индексы колонок — определяются из заголовка
let COL = {};

// ====================== ОПРЕДЕЛЕНИЕ КОЛОНОК ======================
function detectColumns(headerRow) {
  const MAP = {
    "дата начала":             "V",
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

  const required = ["V", "AA", "U", "Q", "J", "F"];
  const missing = required.filter(k => result[k] === undefined);
  if (missing.length) {
    console.warn("Не найдены колонки:", missing, "в заголовке:", headerRow);
  }
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

    const tempQ = new Set(), tempU = new Set();
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

    newPerYearGlobal = {};
    const cumulative = new Set();
    Object.keys(innByYear).sort((a,b)=>Number(a)-Number(b)).forEach(year => {
      let count = 0;
      innByYear[year].forEach(inn => {
        if (!cumulative.has(inn)) { count++; cumulative.add(inn); }
      });
      newPerYearGlobal[year] = count;
    });

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

function populateFilters() {
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

// ====================== СТАТИСТИКА AA ======================
function populateGeneralStats() {
  const html = `
    <h3 style="margin-bottom:20px;">Статистика по удостоверяющим центрам</h3>
    <div style="margin-bottom:20px; text-align:center;">
      <select id="aa-year-filter" onchange="filterAAStats()" style="padding:10px 16px; border-radius:12px; border:2px solid #e2e8f0; margin-right:10px;">
        <option value="">Все годы</option>
      </select>
      <select id="aa-month-filter" onchange="filterAAStats()" style="padding:10px 16px; border-radius:12px; border:2px solid #e2e8f0;">
        <option value="">Все месяцы</option>
      </select>
    </div>
    <table class="result-table">
      <thead><tr><th>Удостоверяющий центр</th><th>Количество записей</th></tr></thead>
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
  const years = new Set();
  const months = new Set();

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
    tbody.innerHTML = `<tr><td colspan="2" style="text-align:center;color:#ef4444;">Колонки не определены</td></tr>`;
    return;
  }

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    if (row.length <= COL.AA) continue;
    const center = String(row[COL.AA] || "").trim();
    if (!center) continue;

    const date = parseDate(row[COL.V]);
    if (!date) continue;

    if (yearF && date.getFullYear() != yearF) continue;
    if (monthF && (date.getMonth()+1) != monthF) continue;

    countMap[center] = (countMap[center] || 0) + 1;
  }

  const sorted = Object.entries(countMap).sort((a,b) => b[1] - a[1]);
  tbody.innerHTML = sorted.length
    ? sorted.map(([center, cnt]) => `<tr><td>${center}</td><td style="text-align:right">${cnt.toLocaleString('ru-RU')}</td></tr>`).join("")
    : `<tr><td colspan="2" style="text-align:center;color:#64748b;">Нет данных</td></tr>`;
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

  const includedQ = new Set();
  document.querySelectorAll("#q-filters input:checked").forEach(cb => includedQ.add(cb.value));
  const includedU = new Set();
  document.querySelectorAll("#u-main input:checked, #u-kcr input:checked").forEach(cb => includedU.add(cb.value));

  let totalJ1 = 0, setJ1 = new Set();
  let totalJ2 = 0, setJ2 = new Set();

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];

    const qval = String(row[COL.Q] || "").trim();
    if (qval && !includedQ.has(qval)) continue;

    const uval = String(row[COL.U] || "").trim();
    if (uval && !includedU.has(uval)) continue;

    const date = parseDate(row[COL.V]);
    if (!date) continue;

    const jval = String(row[COL.J] || "").trim().toUpperCase();

    if (date >= p1s && date <= p1e) {
      if (jval) { totalJ1++; setJ1.add(jval); }
    } else if (date >= p2s && date <= p2e) {
      if (jval) { totalJ2++; setJ2.add(jval); }
    }
  }

  const matchJ = [...setJ1].filter(x => setJ2.has(x)).length;
  const convJ = setJ1.size ? (matchJ / setJ1.size * 100).toFixed(2) : "0";

  const resultDiv = document.getElementById("result");
  resultDiv.innerHTML = `
    <div class="center">
      <h2>✅ Результат сравнения</h2>
      <p style="font-size:1.25em; margin:25px 0;">
        <strong>Период 1:</strong> ${p1s.toLocaleDateString('ru-RU')} — ${p1e.toLocaleDateString('ru-RU')}<br>
        <strong>Период 2:</strong> ${p2s.toLocaleDateString('ru-RU')} — ${p2e.toLocaleDateString('ru-RU')}
      </p>
      <div style="display:flex; justify-content:center; margin-top:30px;">
        <div style="min-width:420px;">
          <h3 style="text-align:center; color:#1e40af;">По СНИЛС</h3>
          <table class="result-table">
            <tr><th>Метрика</th><th>Значение</th></tr>
            <tr><td>Всего строк (П1)</td><td>${totalJ1.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Уникальных (П1)</td><td>${setJ1.size.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Всего строк (П2)</td><td>${totalJ2.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Уникальных (П2)</td><td>${setJ2.size.toLocaleString('ru-RU')}</td></tr>
            <tr style="background:#f8fafc"><td><strong>Уникальных совпадений</strong></td><td><strong>${matchJ.toLocaleString('ru-RU')}</strong></td></tr>
            <tr style="background:#f0fdf4;color:#166534"><td><strong>Конверсия / Удержание</strong></td><td><strong>${convJ}%</strong></td></tr>
          </table>
        </div>
      </div>
    </div>`;

  resultDiv.style.display = "block";
  resultDiv.scrollIntoView({ behavior: "smooth", block: "start" });
}