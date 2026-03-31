function parsePeriodFromFilename(filename) {
    const yearMatch = filename.match(/\d{4}/);
    if (yearMatch) {
        return `Январь—Март ${yearMatch[0]}`;
    }
    return filename.replace(/\.(xlsx|xls)$/i, '');
}

function handleFileSelect(num) {
    const fileInput = document.getElementById(`file${num}`);
    const file = fileInput.files[0];
    if (!file) return;

    const period = parsePeriodFromFilename(file.name);
    const statusEl = document.getElementById(`status${num}`);
    const periodEl = document.getElementById(`period${num}`);

    periodEl.textContent = period;
    statusEl.innerHTML = `✅ <strong>${file.name}</strong><br><small>Загружен успешно</small>`;

    const btn = document.getElementById('analyzeBtn');
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];
    btn.disabled = !(file1 && file2);
}

async function analyzeFiles() {
    const file1 = document.getElementById('file1').files[0];
    const file2 = document.getElementById('file2').files[0];

    const analyzeBtn = document.getElementById('analyzeBtn');
    analyzeBtn.textContent = "⏳ Анализ...";
    analyzeBtn.disabled = true;

    try {
        const data1 = await readExcel(file1);
        const data2 = await readExcel(file2);

        const colIndex = 5; // Колонка F
        const set1 = new Set();
        const set2 = new Set();
        let totalRows1 = 0;
        let totalRows2 = 0;

        data1.forEach(row => {
            if (row.length > colIndex) {
                const val = String(row[colIndex]).trim();
                if (val !== "") {
                    totalRows1++;
                    set1.add(val.toUpperCase());
                }
            }
        });

        const intersections = new Set();
        data2.forEach(row => {
            if (row.length > colIndex) {
                const val = String(row[colIndex]).trim();
                if (val !== "") {
                    totalRows2++;
                    const upperVal = val.toUpperCase();
                    set2.add(upperVal);
                    if (set1.has(upperVal)) {
                        intersections.add(upperVal);
                    }
                }
            }
        });

        const unique1 = set1.size;
        const matches = intersections.size;
        const conversion = unique1 > 0 ? (matches / unique1 * 100).toFixed(2) : 0;

        const period1 = parsePeriodFromFilename(file1.name);
        const period2 = parsePeriodFromFilename(file2.name);

        const resultDiv = document.getElementById('result');
        
        // Формируем контент
        resultDiv.innerHTML = `
            <div class="center">
                <h2 style="margin-top: 0;">✅ Результат сравнения</h2>
                <p style="font-size: 1.25em; margin-bottom: 25px;">
                    <strong>${period1}</strong> → <strong>${period2}</strong>
                </p>
                <table class="result-table" style="margin: 0 auto; width: 100%; max-width: 600px; border-collapse: collapse;">
                    <tr><th style="text-align: left; padding: 12px; border: 1px solid #e2e8f0; background: #f1f5f9;">Метрика</th><th style="text-align: left; padding: 12px; border: 1px solid #e2e8f0; background: #f1f5f9;">Значение</th></tr>
                    <tr><td style="padding: 12px; border: 1px solid #e2e8f0;">Всего строк (${period1})</td><td style="padding: 12px; border: 1px solid #e2e8f0;">${totalRows1}</td></tr>
                    <tr><td style="padding: 12px; border: 1px solid #e2e8f0;">Уникальных (${period1})</td><td style="padding: 12px; border: 1px solid #e2e8f0;">${unique1}</td></tr>
                    <tr><td style="padding: 12px; border: 1px solid #e2e8f0;">Всего строк (${period2})</td><td style="padding: 12px; border: 1px solid #e2e8f0;">${totalRows2}</td></tr>
                    <tr><td style="padding: 12px; border: 1px solid #e2e8f0;">Уникальных (${period2})</td><td style="padding: 12px; border: 1px solid #e2e8f0;">${set2.size}</td></tr>
                    <tr style="background: #f8fafc;">
                        <td style="padding: 12px; border: 1px solid #e2e8f0;"><strong>Уникальных совпадений</strong></td>
                        <td style="padding: 12px; border: 1px solid #e2e8f0;"><strong>${matches}</strong></td>
                    </tr>
                    <tr style="background: #f0fdf4; color: #166534;">
                        <td style="padding: 12px; border: 1px solid #e2e8f0;"><strong>Конверсия / Удержание</strong></td>
                        <td style="padding: 12px; border: 1px solid #e2e8f0;"><strong>${conversion}%</strong></td>
                    </tr>
                </table>
                <p style="margin-top: 25px; color: #64748b; font-size: 0.9em;">Анализ завершен успешно!</p>
            </div>
        `;

        // 1. Показываем блок
        resultDiv.style.display = 'block';

        // 2. ПЛАВНЫЙ ПЕРЕХОД ВНИЗ К РЕЗУЛЬТАТУ
        resultDiv.scrollIntoView({ behavior: 'smooth', block: 'start' });

    } catch (error) {
        alert("Ошибка при чтении файлов: " + error.message);
    } finally {
        analyzeBtn.textContent = "🚀 Сравнить файлы";
        analyzeBtn.disabled = false;
    }
}

function readExcel(file) {
    return new Promise((resolve, reject) => {
        const reader = new FileReader();
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const worksheet = workbook.Sheets[workbook.SheetNames[0]];
                resolve(XLSX.utils.sheet_to_json(worksheet, { header: 1 }));
            } catch (err) { reject(err); }
        };
        reader.readAsArrayBuffer(file);
    });
}