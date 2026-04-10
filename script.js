// ====================== ГЛОБАЛЬНЫЕ ПЕРЕМЕННЫЕ ======================
let allData = [];
let uniqueTariffs = [];
let uniqueUValues = [];
let newPerYearGlobal = {};
let kcrEnabled = true;
let COL = {};

// Хранилище последних результатов для скачивания
let lastExportData = null;

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
    if (!center) continue;
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
function calcRenewal(rows, p1s, p1e, p2s, p2e, includedQ, includedU) {
  if (COL.END === undefined) return null;

  // Все серты по ИНН без фильтров — для поиска предшественника
  const allCerts = {};
  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    const inn = String(row[COL.F] || "").trim().toUpperCase();
    if (!inn) continue;
    const start = parseDate(row[COL.V]);
    const end   = parseDate(row[COL.END]);
    if (!start) continue;
    if (!allCerts[inn]) allCerts[inn] = [];
    allCerts[inn].push({ start, end });
  }
  for (const inn in allCerts) allCerts[inn].sort((a, b) => a.start - b.start);

  const setINN_P1 = new Set();
  const rowsP1 = [];
  const certsP2 = [];

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
    } else if (date >= p2s && date <= p2e) {
      certsP2.push({ inn, start: date, rowIndex: i });
    }
  }

  // Продлились: ИНН из П2 с перекрывающимся предшественником
  const renewedINN = new Set();
  const renewedRowIndexes = new Set();
  for (const { inn, start: newStart, rowIndex } of certsP2) {
    const certs = allCerts[inn];
    if (!certs) continue;
    const hasOverlap = certs.some(c =>
      c.start < newStart && c.end !== null && c.end >= newStart
    );
    if (hasOverlap) { renewedINN.add(inn); renewedRowIndexes.add(rowIndex); }
  }

  // Просрочились: были в П1, не появились в П2 вообще
  const setINN_P2 = new Set(certsP2.map(c => c.inn));
  const lapsedINN = new Set([...setINN_P1].filter(inn => !setINN_P2.has(inn)));

  // Берём последнюю строку каждого просроченного ИНН из П1
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
  const renewalCount = renewedINN.size;
  const renewalPct = denominator > 0 ? (renewalCount / denominator * 100).toFixed(2) : "0";

  return {
    renewalCount, denominator, renewalPct,
    renewedINN, renewedRowIndexes,
    lapsedINN, lapsedRowIndexes,
    setINN_P2
  };
}

// ====================== СКАЧИВАНИЕ XLSX ======================
function downloadExport(type) {
  if (!lastExportData) return;
  const { renewal, setJ1, setJ2, p2Rows, header } = lastExportData;

  let wsData = [header];
  let filename = "";

  if (type === "renewed") {
    // Строки П2, где ИНН продлился
    if (!renewal) return;
    p2Rows.forEach(r => { if (renewal.renewedINN.has(r.inn)) wsData.push(r.rowData); });
    filename = "продления.xlsx";
  } else if (type === "retained") {
    // Строки П2, где СНИЛС совпал с П1
    p2Rows.forEach(r => { if (r.snils && setJ1.has(r.snils)) wsData.push(r.rowData); });
    filename = "удержанные.xlsx";
  } else if (type === "lapsed") {
    // Строки П1 для просроченных ИНН (последняя запись)
    if (!renewal) return;
    for (let i = 1; i < allData.length; i++) {
      if (renewal.lapsedRowIndexes.has(i)) wsData.push(allData[i]);
    }
    filename = "просроченные.xlsx";
  }

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(wsData), "Лист1");
  XLSX.writeFile(wb, filename);
}

// ====================== ТАБЛИЦА УЦ ДЛЯ ПЕРИОДА ======================
function renderAATable(countMap, title) {
  const sorted = Object.entries(countMap).sort((a,b) => b[1] - a[1]);
  const total = sorted.reduce((s,[,c]) => s + c, 0);
  if (!sorted.length) return `<div style="min-width:340px;"><h4 style="text-align:center;color:#64748b;">${title}</h4><p style="text-align:center;color:#94a3b8;">Нет данных</p></div>`;
  const rows = sorted.map(([c, n]) => {
    const pct = (n / total * 100).toFixed(1);
    return `<tr><td>${c}</td><td style="text-align:right">${n.toLocaleString('ru-RU')}</td><td style="text-align:right;color:#1e40af;font-weight:600">${pct}%</td></tr>`;
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
            <td style="text-align:right">${total.toLocaleString('ru-RU')}</td>
            <td style="text-align:right">100%</td>
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

  const includedQ = new Set();
  document.querySelectorAll("#q-filters input:checked").forEach(cb => includedQ.add(cb.value));
  const includedU = new Set();
  document.querySelectorAll("#u-main input:checked, #u-kcr input:checked").forEach(cb => includedU.add(cb.value));

  // --- Удержание (по СНИЛС) + статистика УЦ по периодам ---
  let totalJ1 = 0, setJ1 = new Set();
  let totalJ2 = 0, setJ2 = new Set();
  const p2Rows = [];
  const aaP1 = {}, aaP2 = {};

  for (let i = 1; i < allData.length; i++) {
    const row = allData[i];
    const qval = String(row[COL.Q] || "").trim();
    if (qval && !includedQ.has(qval)) continue;
    const uval = String(row[COL.U] || "").trim();
    if (uval && !includedU.has(uval)) continue;
    const date = parseDate(row[COL.V]);
    if (!date) continue;

    const jval = String(row[COL.J] || "").trim().toUpperCase();
    const inn  = COL.F !== undefined ? String(row[COL.F] || "").trim().toUpperCase() : "";
    const center = COL.AA !== undefined ? String(row[COL.AA] || "").trim() : "";

    if (date >= p1s && date <= p1e) {
      if (jval) { totalJ1++; setJ1.add(jval); }
      if (center) aaP1[center] = (aaP1[center] || 0) + 1;
    } else if (date >= p2s && date <= p2e) {
      if (jval) { totalJ2++; setJ2.add(jval); }
      if (center) aaP2[center] = (aaP2[center] || 0) + 1;
      p2Rows.push({ inn, snils: jval, rowData: row, rowIndex: i });
    }
  }

  const matchJ = [...setJ1].filter(x => setJ2.has(x)).length;
  const convJ = setJ1.size ? (matchJ / setJ1.size * 100).toFixed(2) : "0";

  // --- Продление ---
  const renewal = calcRenewal(allData, p1s, p1e, p2s, p2e, includedQ, includedU);

  // Сохраняем для кнопок скачивания
  lastExportData = { renewal, setJ1, setJ2, p2Rows, header: allData[0] };

  // --- HTML блоков ---
  const renewalBlock = renewal
    ? `
      <div style="min-width:420px;">
        <h3 style="text-align:center; color:#7c3aed;">По ИНН — Продление</h3>
        <p style="text-align:center; font-size:0.9em; color:#64748b; margin-bottom:12px;">
          Новый сертификат выпущен <em>до истечения</em> предыдущего по тому же ИНН
        </p>
        <table class="result-table">
          <tr><th>Метрика</th><th>Значение</th></tr>
          <tr><td>Уникальных ИНН (П1, знаменатель)</td><td>${renewal.denominator.toLocaleString('ru-RU')}</td></tr>
          <tr><td>Продлений (уник. ИНН)</td><td>${renewal.renewalCount.toLocaleString('ru-RU')}</td></tr>
          <tr><td>Просрочено (ИНН П1, не вернулись)</td><td>${renewal.lapsedINN.size.toLocaleString('ru-RU')}</td></tr>
          <tr style="background:#faf5ff;color:#6d28d9"><td><strong>% Продления</strong></td><td><strong>${renewal.renewalPct}%</strong></td></tr>
        </table>
      </div>`
    : `<p style="color:#ef4444; text-align:center;">⚠️ Колонка «Дата окончания» не найдена — продление не рассчитано</p>`;

  const exportButtons = `
    <div style="margin-top:40px; border-top:2px solid #e2e8f0; padding-top:30px;">
      <h3 style="margin-bottom:16px; color:#374151;">📥 Скачать выгрузки (xlsx)</h3>
      <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:14px;">
        ${renewal ? `<button type="button" onclick="downloadExport('renewed')"
          style="background:linear-gradient(90deg,#7c3aed,#a78bfa); padding:12px 24px; font-size:1em; box-shadow:0 4px 14px rgba(124,58,237,0.35);">
          🔄 Продлились (${renewal.renewalCount.toLocaleString('ru-RU')})
        </button>` : ""}
        <button type="button" onclick="downloadExport('retained')"
          style="background:linear-gradient(90deg,#166534,#22c55e); padding:12px 24px; font-size:1em; box-shadow:0 4px 14px rgba(22,101,52,0.35);">
          ✅ Удержались (${matchJ.toLocaleString('ru-RU')})
        </button>
        ${renewal ? `<button type="button" onclick="downloadExport('lapsed')"
          style="background:linear-gradient(90deg,#b91c1c,#ef4444); padding:12px 24px; font-size:1em; box-shadow:0 4px 14px rgba(185,28,28,0.35);">
          ❌ Просрочились (${renewal.lapsedINN.size.toLocaleString('ru-RU')})
        </button>` : ""}
      </div>
      <p style="font-size:0.82em; color:#94a3b8; margin-top:12px; text-align:center;">
        Продлились / Удержались → строки Периода 2 &nbsp;|&nbsp; Просрочились → последняя строка Периода 1 по ИНН
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
        <div style="min-width:420px;">
          <h3 style="text-align:center; color:#1e40af;">По СНИЛС — Удержание</h3>
          <p style="text-align:center; font-size:0.9em; color:#64748b; margin-bottom:12px;">
            Один и тот же физлицо (СНИЛС) купил сертификат в обоих периодах
          </p>
          <table class="result-table">
            <tr><th>Метрика</th><th>Значение</th></tr>
            <tr><td>Всего строк (П1)</td><td>${totalJ1.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Уникальных СНИЛС (П1)</td><td>${setJ1.size.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Всего строк (П2)</td><td>${totalJ2.toLocaleString('ru-RU')}</td></tr>
            <tr><td>Уникальных СНИЛС (П2)</td><td>${setJ2.size.toLocaleString('ru-RU')}</td></tr>
            <tr style="background:#f8fafc"><td><strong>Совпадений СНИЛС</strong></td><td><strong>${matchJ.toLocaleString('ru-RU')}</strong></td></tr>
            <tr style="background:#f0fdf4;color:#166534"><td><strong>% Удержания</strong></td><td><strong>${convJ}%</strong></td></tr>
          </table>
        </div>
        ${renewalBlock}
      </div>

      <div style="margin-top:40px; border-top:2px solid #e2e8f0; padding-top:30px;">
        <h3 style="margin-bottom:20px;">📊 Удостоверяющие центры по периодам</h3>
        <div style="display:flex; flex-wrap:wrap; justify-content:center; gap:40px;">
          ${renderAATable(aaP1, `Период 1 (${p1s.toLocaleDateString('ru-RU')} — ${p1e.toLocaleDateString('ru-RU')})`)}
          ${renderAATable(aaP2, `Период 2 (${p2s.toLocaleDateString('ru-RU')} — ${p2e.toLocaleDateString('ru-RU')})`)}
        </div>
      </div>

      ${exportButtons}
    </div>`;

  resultDiv.style.display = "block";
  resultDiv.scrollIntoView({ behavior: "smooth", block: "start" });
}