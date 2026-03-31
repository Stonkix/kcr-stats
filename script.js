async function analyzeFiles() {
  const file1 = document.getElementById('file1').files[0];
  const file2 = document.getElementById('file2').files[0];

  if (!file1 || !file2) {
    alert('Пожалуйста, загрузите оба файла!');
    return;
  }

  const data1 = await readExcel(file1);
  const data2 = await readExcel(file2);

  const colIndex = 5; // Колонка F (A=0, B=1, C=2, D=3, E=4, F=5)

  const set2025 = new Set();
  const set2026 = new Set();
  let totalRows2025 = 0;
  let totalRows2026 = 0;

  // Обработка файла 2025
  data1.forEach(row => {
    if (row.length > colIndex) {
      const val = String(row[colIndex]).trim();
      if (val !== "") {
        totalRows2025++;
        set2025.add(val.toUpperCase());
      }
    }
  });

  // Обработка файла 2026
  const intersections = new Set();
  data2.forEach(row => {
    if (row.length > colIndex) {
      const val = String(row[colIndex]).trim();
      if (val !== "") {
        totalRows2026++;
        const upper = val.toUpperCase();
        set2026.add(upper);
        if (set2025.has(upper)) {
          intersections.add(upper);
        }
      }
    }
  });

  const unique2025 = set2025.size;
  const unique2026 = set2026.size;
  const matches = intersections.size;
  const conversion = unique2025 > 0 ? (matches / unique2025 * 100).toFixed(2) : 0;

  // Вывод результата
  const resultHTML = `
    <h2 class="center">Результат анализа</h2>
    <table>
      <tr><th>Метрика</th><th>Значение</th></tr>
      <tr><td>Всего строк в колонке F (2025)</td><td>${totalRows2025}</td></tr>
      <tr><td>Уникальных значений в 2025</td><td>${unique2025}</td></tr>
      <tr><td>Всего строк в колонке F (2026)</td><td>${totalRows2026}</td></tr>
      <tr><td>Уникальных значений в 2026</td><td>${unique2026}</td></tr>
      <tr class="highlight"><td><strong>Уникальных совпадений</strong></td><td><strong>${matches}</strong></td></tr>
      <tr class="highlight green"><td><strong>Конверсия 2025 → 2026</strong></td><td><strong>${conversion}%</strong></td></tr>
    </table>
    <p class="center" style="margin-top:20px; color:#64748b;">
      Анализ выполнен в браузере • Данные не отправляются на сервер
    </p>
  `;

  const resultDiv = document.getElementById('result');
  resultDiv.innerHTML = resultHTML;
  resultDiv.style.display = 'block';
}

function readExcel(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = function(e) {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      resolve(json);
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
}