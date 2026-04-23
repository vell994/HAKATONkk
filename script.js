// ======================= ГЛОБАЛЬНЫЕ ДАННЫЕ =======================
let tableData = [];
let certConfig = {
    align: 'center',
    width: 297,
    height: 210,
    bgImage: null,
    degree: 'none',        // 'none', '1', '2', '3'
    footerText: 'Директор АНО ДПО "Форсайт" В. В. Гартунг',
    colors: {
        title: '#8e44ad',
        degree: '#e67e22',
        name: '#2c3e50',
        body: '#333333',
        footer: '#7f8c8d'
    }
};
let currentPage = 0;
const rowsPerPage = 10;

// ======================= ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ =======================
function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/[&<>]/g, function(m) {
        if (m === '&') return '&amp;';
        if (m === '<') return '&lt;';
        if (m === '>') return '&gt;';
        return m;
    }).replace(/[\n]/g, '<br>');
}

function downloadFile(blob, name) {
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = name;
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    setTimeout(() => URL.revokeObjectURL(url), 100);
}

function syncColorPickersFromConfig() {
    const ids = ['title', 'degree', 'name', 'body', 'footer'];
    ids.forEach(id => {
        const picker = document.getElementById(`${id}Color`);
        const span = document.getElementById(`${id}ColorVal`);
        if (picker) picker.value = certConfig.colors[id];
        if (span) span.textContent = certConfig.colors[id];
    });
}

function resetColor(type) {
    const defaults = { title: '#8e44ad', degree: '#e67e22', name: '#2c3e50', body: '#333333', footer: '#7f8c8d' };
    certConfig.colors[type] = defaults[type];
    localStorage.setItem('certColors', JSON.stringify(certConfig.colors));
    syncColorPickersFromConfig();
    renderPreview();
}

function saveSettings() {
    localStorage.setItem('certColors', JSON.stringify(certConfig.colors));
    localStorage.setItem('certBg', certConfig.bgImage || '');
    localStorage.setItem('certFooterText', certConfig.footerText);
    localStorage.setItem('certDegree', certConfig.degree);
}

// Получить текст степени
function getDegreeText(degreeValue) {
    if (degreeValue === '1') return '🥇 Первое место';
    if (degreeValue === '2') return '🥈 Второе место';
    if (degreeValue === '3') return '🥉 Третье место';
    return '';
}

// ======================= ИНИЦИАЛИЗАЦИЯ =======================
document.addEventListener('DOMContentLoaded', () => {
    const savedColors = localStorage.getItem('certColors');
    if (savedColors) Object.assign(certConfig.colors, JSON.parse(savedColors));
    
    const savedBg = localStorage.getItem('certBg');
    if (savedBg && savedBg !== 'null') certConfig.bgImage = savedBg;
    
    const savedFooter = localStorage.getItem('certFooterText');
    if (savedFooter) certConfig.footerText = savedFooter;
    
    const savedDegree = localStorage.getItem('certDegree');
    if (savedDegree) certConfig.degree = savedDegree;
    
    syncColorPickersFromConfig();
    
    // Выбор степени
    const degreeSelect = document.getElementById('degreeSelect');
    if (degreeSelect) {
        degreeSelect.value = certConfig.degree;
        degreeSelect.addEventListener('change', (e) => {
            certConfig.degree = e.target.value;
            saveSettings();
            renderPreview();
        });
    }
    
    // Нижний текст
    const footerInput = document.getElementById('footerText');
    if (footerInput) {
        footerInput.value = certConfig.footerText;
        footerInput.addEventListener('input', (e) => {
            certConfig.footerText = e.target.value;
            saveSettings();
            renderPreview();
        });
    }
    
    // Загрузка фона
    const bgFileInput = document.getElementById('bgImageFile');
    if (bgFileInput) {
        bgFileInput.addEventListener('change', (e) => {
            const file = e.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = (ev) => {
                    certConfig.bgImage = ev.target.result;
                    saveSettings();
                    renderPreview();
                };
                reader.readAsDataURL(file);
            }
        });
    }
    
    // Сброс фона
    const clearBgBtn = document.getElementById('clearBgBtn');
    if (clearBgBtn) {
        clearBgBtn.addEventListener('click', () => {
            certConfig.bgImage = null;
            saveSettings();
            renderPreview();
            if (bgFileInput) bgFileInput.value = '';
        });
    }
    
    // Обработчики цвета
    ['title', 'degree', 'name', 'body', 'footer'].forEach(key => {
        const el = document.getElementById(`${key}Color`);
        if (el) {
            el.addEventListener('input', (e) => {
                certConfig.colors[key] = e.target.value;
                document.getElementById(`${key}ColorVal`).textContent = e.target.value;
                saveSettings();
                renderPreview();
            });
        }
    });
    
    // Кнопка применения пользовательского размера
    const applyCustomBtn = document.getElementById('applyCustomSize');
    if (applyCustomBtn) {
        applyCustomBtn.addEventListener('click', () => {
            const w = parseFloat(document.getElementById('customW').value);
            const h = parseFloat(document.getElementById('customH').value);
            if (!isNaN(w) && w > 0) certConfig.width = w;
            if (!isNaN(h) && h > 0) certConfig.height = h;
            renderPreview();
        });
    }
    
    // При изменении полей custom вручную можно тоже обновлять, но добавим для удобства
    document.getElementById('customW')?.addEventListener('input', () => {});
    document.getElementById('customH')?.addEventListener('input', () => {});
    
    // Загрузка таблицы
    const stored = localStorage.getItem('excelDataForCertificate');
    if (stored) {
        try { tableData = JSON.parse(stored); } catch(e) { console.error(e); }
    }
    if (!tableData || tableData.length === 0) {
        tableData = [
            ['ФИО Участника', 'Описание'],
            ['Иванов Иван Иванович', 'освоил программу по модулям: основы живописи, графический дизайн, архитектурный дизайн'],
            ['Петрова Анна Сергеевна', 'проявила отличные результаты в вокале и хореографии']
        ];
        localStorage.setItem('excelDataForCertificate', JSON.stringify(tableData));
    }
    
    renderTable();
    renderPreview();
});

// ======================= ЗАГРУЗКА EXCEL =======================
document.getElementById('excelFile').addEventListener('change', function(e) {
    const file = e.target.files[0];
    if (!file) return;
    const reader = new FileReader();
    reader.onload = (ev) => {
        const data = new Uint8Array(ev.target.result);
        const wb = XLSX.read(data, { type: 'array' });
        const sheet = wb.Sheets[wb.SheetNames[0]];
        const json = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: '' });
        if (json.length) {
            tableData = json;
            localStorage.setItem('excelDataForCertificate', JSON.stringify(tableData));
            currentPage = 0;
            renderTable();
            renderPreview();
            document.getElementById('generateCertBtn').disabled = false;
            document.getElementById('hintText').textContent = `Строк: ${tableData.length - 1}`;
        }
    };
    reader.readAsArrayBuffer(file);
    this.value = '';
});

// ======================= НАСТРОЙКИ =======================
function setAlign(align) {
    certConfig.align = align;
    document.querySelectorAll('.btn-group button[id^="btn-align-"]').forEach(b => b.classList.remove('active'));
    document.getElementById(`btn-align-${align}`).classList.add('active');
    renderPreview();
}

function updateSizeSettings() {
    const val = document.getElementById('sizeSelect').value;
    const customDiv = document.getElementById('customSizeDiv');
    if (val === 'custom') {
        customDiv.style.display = 'flex';
        // Не применяем автоматически, ждём кнопку "Применить"
    } else {
        customDiv.style.display = 'none';
        if (val === 'a4-l') { certConfig.width = 297; certConfig.height = 210; }
        if (val === 'a4-p') { certConfig.width = 210; certConfig.height = 297; }
        renderPreview();
    }
}

// ======================= УПРАВЛЕНИЕ ТАБЛИЦЕЙ =======================
function addColumn() {
    const colName = prompt("Название колонки:", `Колонка ${(tableData[0]?.length || 0) + 1}`);
    if (!colName) return;
    tableData.forEach((row, idx) => idx === 0 ? row.push(colName) : row.push(''));
    saveAndRenderTable();
}

function deleteColumn(colIdx) {
    if (tableData[0].length <= 1) return alert("Нельзя удалить последнюю колонку!");
    if (!confirm("Удалить колонку?")) return;
    tableData.forEach(row => row.splice(colIdx, 1));
    saveAndRenderTable();
}

function deleteRow(rowIdx) {
    if (tableData.length <= 1) return alert("Нет строк для удаления!");
    if (!confirm("Удалить строку?")) return;
    tableData.splice(rowIdx, 1);
    saveAndRenderTable();
}

function resetTable() {
    if (!confirm("Сбросить таблицу?")) return;
    tableData = [tableData[0] || ['ФИО', 'Описание']];
    currentPage = 0;
    saveAndRenderTable();
}

function saveAndRenderTable() {
    localStorage.setItem('excelDataForCertificate', JSON.stringify(tableData));
    renderTable();
    renderPreview();
}

function renderTable() {
    const thead = document.getElementById('tableHead');
    const tbody = document.getElementById('tableBody');
    if (!tableData.length) {
        thead.innerHTML = '<th>Нет данных</th>';
        tbody.innerHTML = '';
        return;
    }
    const headers = tableData[0];
    thead.innerHTML = headers.map((h, i) => `
        <th><div><span contenteditable="true" onblur="updateHeader(${i}, this.textContent)">${escapeHtml(h)}</span>
        <button class="del-col" onclick="deleteColumn(${i})">×</button></div></th>
    `).join('');

    const dataRows = tableData.slice(1);
    const start = currentPage * rowsPerPage;
    const pageData = dataRows.slice(start, start + rowsPerPage);
    tbody.innerHTML = pageData.map((row, idx) => {
        const realIdx = start + idx + 1;
        return `<tr>${row.map((cell, cIdx) => `<td contenteditable="true" data-row="${realIdx}" data-col="${cIdx}">${escapeHtml(cell ?? '')}</td>`).join('')}
        <td style="width:40px; text-align:center"><button class="del-row-btn" onclick="deleteRow(${realIdx})">❌</button></td></tr>`;
    }).join('');

    document.querySelectorAll('td[contenteditable="true"]').forEach(td => {
        td.removeEventListener('blur', handleCellBlur);
        td.addEventListener('blur', handleCellBlur);
    });
    updatePaginationUI();
}

function handleCellBlur() {
    const r = parseInt(this.dataset.row);
    const c = parseInt(this.dataset.col);
    if (tableData[r] && !isNaN(r) && !isNaN(c)) {
        tableData[r][c] = this.textContent;
        localStorage.setItem('excelDataForCertificate', JSON.stringify(tableData));
        if (r === 1) renderPreview();
    }
}

function updateHeader(colIdx, newText) {
    if (tableData[0] && colIdx >= 0) tableData[0][colIdx] = newText.trim();
    localStorage.setItem('excelDataForCertificate', JSON.stringify(tableData));
}

function updatePaginationUI() {
    const totalRows = Math.max(0, tableData.length - 1);
    const totalPages = Math.ceil(totalRows / rowsPerPage) || 1;
    document.getElementById('pageInfo').textContent = `Страница ${currentPage + 1} из ${totalPages}`;
    document.getElementById('prevBtn').disabled = currentPage === 0;
    document.getElementById('nextBtn').disabled = currentPage >= totalPages - 1;
}

function changePage(delta) {
    const totalRows = Math.max(0, tableData.length - 1);
    const totalPages = Math.ceil(totalRows / rowsPerPage) || 1;
    const newPage = currentPage + delta;
    if (newPage >= 0 && newPage < totalPages) {
        currentPage = newPage;
        renderTable();
    }
}

// ======================= ПРЕДПРОСМОТР =======================
function renderPreview() {
    const container = document.getElementById('certificatePreview');
    container.innerHTML = '';
    const maxWidth = 800;
    const scale = maxWidth / certConfig.width;
    const height = certConfig.height * scale;
    container.style.width = `${maxWidth}px`;
    container.style.height = `${height}px`;
    container.style.position = 'relative';
    container.style.overflow = 'hidden';
    container.style.background = '#fff';
    container.style.borderRadius = '16px';

    // Фон (если загружен)
    if (certConfig.bgImage) {
        const bgImg = document.createElement('img');
        bgImg.src = certConfig.bgImage;
        bgImg.style.cssText = 'position:absolute; top:0; left:0; width:100%; height:100%; object-fit:cover; z-index:0;';
        container.appendChild(bgImg);
    } // иначе просто белый фон (никакого градиента)

    const rowData = (tableData.length > 1 && tableData[1]) ? tableData[1] : tableData[0];
    const fullName = rowData[0] || 'Фамилия Имя';
    const mainText = rowData.slice(1).filter(m => m && m.trim() !== '').join(' ') || 'достиг(ла) выдающихся успехов';
    const degreeText = getDegreeText(certConfig.degree);

    const content = document.createElement('div');
    content.style.cssText = `
        position: relative; z-index: 1; width:100%; height:100%;
        padding: 8% 10%; box-sizing: border-box;
        display: flex; flex-direction: column;
        font-family: 'Times New Roman', Times, serif;
        text-align: center;
        background: transparent;
    `;

    let degreeHtml = '';
    if (degreeText) {
        degreeHtml = `<div style="font-size: 28px; font-weight: bold; color: ${certConfig.colors.degree}; margin-bottom: 15px;">${degreeText}</div>`;
    }

    content.innerHTML = `
        <div style="flex:1; display:flex; flex-direction:column; justify-content:center;">
            <h1 style="font-size: 52px; color: ${certConfig.colors.title}; margin-bottom: 30px; font-weight: bold;">Сертификат участника</h1>
            ${degreeHtml}
            <div style="font-size: 38px; font-weight: bold; color: ${certConfig.colors.name}; margin: 15px 0 20px; letter-spacing: 1px;">${escapeHtml(fullName)}</div>
            <div style="font-size: 22px; color: ${certConfig.colors.body}; line-height: 1.5; max-width: 85%; margin: 0 auto;">
                ${escapeHtml(mainText)}
            </div>
        </div>
        <div style="margin-top: 50px; font-size: 18px; color: ${certConfig.colors.footer}; border-top: 1px solid rgba(0,0,0,0.1); padding-top: 20px;">
            ${escapeHtml(certConfig.footerText)}
        </div>
    `;
    container.appendChild(content);
}

// ======================= ГЕНЕРАЦИЯ СЕРТИФИКАТОВ =======================
async function generateCertificates() {
    const dataRows = tableData.slice(1).filter(r => r && r.some(c => c && String(c).trim()));
    if (!dataRows.length) return alert("Нет данных для генерации!");

    const btn = document.getElementById('generateCertBtn');
    const origText = btn.textContent;
    btn.disabled = true;
    btn.textContent = '⏳ Генерация...';

    const format = document.getElementById('exportFormat').value;
    const tempDiv = document.createElement('div');
    tempDiv.style.cssText = 'position:absolute; left:-9999px; top:-9999px;';
    document.body.appendChild(tempDiv);

    try {
        const files = [];
        const { jsPDF } = window.jspdf;
        const scale = 2;
        const wPx = certConfig.width * 3.78 * scale;
        const hPx = certConfig.height * 3.78 * scale;

        for (let i = 0; i < dataRows.length; i++) {
            btn.textContent = `⏳ ${i+1}/${dataRows.length}...`;
            const row = dataRows[i];
            const wrapper = document.createElement('div');
            wrapper.style.cssText = `width:${wPx}px; height:${hPx}px; position:relative; background:#fff; overflow:hidden; font-family:'Times New Roman', Times, serif; border-radius: 16px;`;

            // Фон
            if (certConfig.bgImage) {
                const bg = document.createElement('img');
                bg.src = certConfig.bgImage;
                bg.style.cssText = 'width:100%; height:100%; object-fit:cover; position:absolute; top:0; left:0;';
                wrapper.appendChild(bg);
                await new Promise(resolve => { if (bg.complete) resolve(); else bg.onload = resolve; });
            }

            const contentDiv = document.createElement('div');
            const padding = 50 * scale;
            contentDiv.style.cssText = `position:relative; z-index:2; padding:${padding}px; display:flex; flex-direction:column; justify-content:center; height:100%; box-sizing:border-box; text-align:center; background:transparent;`;

            const fullName = row[0] || 'ФИО';
            const mainText = row.slice(1).filter(m => m && m.trim() !== '').join(' ') || 'достиг(ла) выдающихся успехов';
            const degreeText = getDegreeText(certConfig.degree);
            const degreeHtml = degreeText ? `<div style="font-size: ${28*scale}px; font-weight: bold; color: ${certConfig.colors.degree}; margin-bottom: ${15*scale}px;">${degreeText}</div>` : '';

            contentDiv.innerHTML = `
                <div style="flex:1; display:flex; flex-direction:column; justify-content:center;">
                    <h1 style="font-size: ${52*scale}px; color: ${certConfig.colors.title}; margin-bottom: ${30*scale}px; font-weight: bold;">Сертификат участника</h1>
                    ${degreeHtml}
                    <div style="font-size: ${38*scale}px; font-weight: bold; color: ${certConfig.colors.name}; margin: ${15*scale}px 0 ${20*scale}px;">${escapeHtml(fullName)}</div>
                    <div style="font-size: ${22*scale}px; color: ${certConfig.colors.body}; line-height: 1.5; max-width: 85%; margin: 0 auto;">
                        ${escapeHtml(mainText)}
                    </div>
                </div>
                <div style="margin-top: ${50*scale}px; font-size: ${18*scale}px; color: ${certConfig.colors.footer}; border-top: 1px solid rgba(0,0,0,0.1); padding-top: ${20*scale}px;">
                    ${escapeHtml(certConfig.footerText)}
                </div>
            `;
            wrapper.appendChild(contentDiv);
            tempDiv.appendChild(wrapper);
            await new Promise(r => setTimeout(r, 50));

            const canvas = await html2canvas(wrapper, { scale: 1, useCORS: true, backgroundColor: '#ffffff' });
            if (format === 'pdf') {
                const orientation = certConfig.width > certConfig.height ? 'l' : 'p';
                const pdf = new jsPDF({ orientation, unit: 'mm', format: [certConfig.width, certConfig.height] });
                pdf.addImage(canvas.toDataURL('image/jpeg', 0.95), 'JPEG', 0, 0, certConfig.width, certConfig.height);
                files.push({ data: pdf.output('blob'), name: `Сертификат_${i+1}.pdf` });
            } else {
                const mime = format === 'png' ? 'image/png' : 'image/jpeg';
                const ext = format === 'png' ? 'png' : 'jpg';
                const blob = await new Promise(resolve => canvas.toBlob(resolve, mime, 0.95));
                files.push({ data: blob, name: `Сертификат_${i+1}.${ext}` });
            }
            tempDiv.innerHTML = '';
        }

        if (files.length > 1) {
            const zip = new JSZip();
            files.forEach(f => zip.file(f.name, f.data));
            const zipBlob = await zip.generateAsync({ type: 'blob' });
            downloadFile(zipBlob, 'Сертификаты.zip');
        } else if (files.length === 1) {
            downloadFile(files[0].data, files[0].name);
        }
    } catch (err) {
        alert('Ошибка: ' + err.message);
    } finally {
        document.body.removeChild(tempDiv);
        btn.disabled = false;
        btn.textContent = origText;
    }
}