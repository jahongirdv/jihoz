// ============================================================
// 1. ДАННЫЕ
// ============================================================
let equipmentDB = {};
let modelsDB = [];

function initData() {
    if (typeof equipmentDatabase !== 'undefined' && Object.keys(equipmentDatabase).length > 0) {
        equipmentDB = { ...equipmentDatabase };
        console.log(`✅ Использую данные из type.js: ${Object.keys(equipmentDB).length} записей`);
    } else {
        equipmentDB = {
            'гастрофиброскопы': { section: 'Диагностическая аппаратура и системы мониторинга', point: 'Приборы для эндоскопии' },
            'электрокардиографы': { section: 'Диагностическая аппаратура и системы мониторинга', point: 'Электрокардиографы (ЭКГ)' },
            'кт': { section: 'Диагностическая аппаратура и системы мониторинга', point: 'Рентген-диагностическое оборудование' },
            'ивл': { section: 'Аппараты для анестезиологии и реанимации', point: 'Аппараты для искусственной вентиляции лёгких (ИВЛ)' },
        };
        console.warn('⚠️ equipmentDatabase не найден, использую резервные данные');
    }
    if (typeof modelsDatabase !== 'undefined' && modelsDatabase.length > 0) {
        modelsDB = [...modelsDatabase];
        console.log(`✅ Использую модели из type.js: ${modelsDB.length}`);
    } else {
        modelsDB = [
            { model: 'С856', manufacturer: 'United Products & Instruments, Inc', country: 'США', equipment: 'Лабораторная центрифуга' },
        ];
        console.warn('⚠️ modelsDatabase не найден, использую резервные модели');
    }
}

// ============================================================
// 2. ПОДСКАЗКИ ДЛЯ НАИМЕНОВАНИЯ
// ============================================================
function findEquipment(query) {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    const results = [];
    for (const key in equipmentDB) {
        if (key.includes(q) || q.includes(key)) {
            results.push({ name: key, section: equipmentDB[key].section, point: equipmentDB[key].point });
        }
    }
    return results;
}

function showEquipmentSuggestions(matches, input, sectionDisplay, pointDisplay, suggestionsList) {
    suggestionsList.innerHTML = '';
    if (!matches.length) { suggestionsList.style.display = 'none'; return; }
    suggestionsList.style.display = 'block';
    const header = document.createElement('div');
    header.className = 'suggestions-header';
    header.textContent = `🔍 Найдено ${matches.length} совпадений. Выберите:`;
    suggestionsList.appendChild(header);
    matches.forEach(match => {
        const item = document.createElement('div');
        item.className = 'suggestion-item';
        item.innerHTML = `
            <div class="suggestion-name">${match.name}</div>
            <div class="suggestion-class">
                <span class="suggestion-section">${match.section}</span>
                <span class="suggestion-point">${match.point}</span>
            </div>
        `;
        item.addEventListener('click', function() {
            input.value = match.name;
            sectionDisplay.textContent = match.section;
            pointDisplay.textContent = match.point;
            suggestionsList.style.display = 'none';
            sectionDisplay.classList.add('found', 'pulse');
            pointDisplay.classList.add('found', 'pulse');
            setTimeout(() => {
                sectionDisplay.classList.remove('found', 'pulse');
                pointDisplay.classList.remove('found', 'pulse');
            }, 1500);
        });
        suggestionsList.appendChild(item);
    });
}

// ============================================================
// 3. ПОДСКАЗКИ ДЛЯ МОДЕЛЕЙ
// ============================================================
function findModels(query) {
    const q = query.trim().toLowerCase();
    if (!q) return [];
    return modelsDB.filter(item => 
        item.model.toLowerCase().includes(q) || q.includes(item.model.toLowerCase())
    );
}

function showModelSuggestions(matches, input, firmInput, countryInput, countryFirmInput, equipmentInput, suggestionsList) {
    suggestionsList.innerHTML = '';
    if (!matches.length) { suggestionsList.style.display = 'none'; return; }
    suggestionsList.style.display = 'block';
    const header = document.createElement('div');
    header.className = 'suggestions-header';
    header.textContent = `🔍 Найдено ${matches.length} моделей. Выберите:`;
    suggestionsList.appendChild(header);
    matches.forEach(match => {
        const item = document.createElement('div');
        item.className = 'suggestion-item';
        item.innerHTML = `
            <div class="suggestion-name">${match.model}</div>
            <div class="suggestion-class">
                <span class="suggestion-section">${match.manufacturer}</span>
                <span class="suggestion-point">${match.country}</span>
            </div>
            ${match.equipment ? `<div style="font-size:12px;color:#8a9eb5;">${match.equipment}</div>` : ''}
        `;
        item.addEventListener('click', function() {
            input.value = match.model;
            firmInput.value = match.manufacturer;
            countryInput.value = match.country;
            countryFirmInput.value = match.country;
            suggestionsList.style.display = 'none';
            if (!equipmentInput.value.trim() && match.equipment) {
                equipmentInput.value = match.equipment;
                equipmentInput.dispatchEvent(new Event('input', { bubbles: true }));
            }
            firmInput.style.borderColor = '#66bb6a';
            countryInput.style.borderColor = '#66bb6a';
            countryFirmInput.style.borderColor = '#66bb6a';
            setTimeout(() => {
                firmInput.style.borderColor = '';
                countryInput.style.borderColor = '';
                countryFirmInput.style.borderColor = '';
            }, 2000);
        });
        suggestionsList.appendChild(item);
    });
}




// ============================================================
// ЭКСПОРТ В XLSX С ФОРМАТИРОВАНИЕМ (ExcelJS) – только жирный шрифт
// ============================================================
async function exportSelectedFields() {
    // Проверяем загрузку библиотек
    if (typeof ExcelJS === 'undefined') {
        alert('❌ Библиотека ExcelJS не загружена!');
        return;
    }
    if (typeof saveAs === 'undefined') {
        alert('❌ Библиотека FileSaver не загружена!');
        return;
    }

    // Собираем отмеченные чекбоксы
    const checkboxes = document.querySelectorAll('.export-field-checkbox:checked');
    if (checkboxes.length === 0) {
        alert('❌ Выберите хотя бы одно поле для экспорта.');
        return;
    }

    // Получаем данные из формы
    const data = {
        name: document.getElementById('equipmentName')?.value || '',
        section: document.getElementById('sectionDisplay')?.textContent || '—',
        point: document.getElementById('pointDisplay')?.textContent || '—',
        manufacturer: document.getElementById('firmInput')?.value || '',
        country: document.getElementById('countryInput')?.value || '',
        model: document.getElementById('modelInput')?.value || '',
        serial: document.getElementById('serialInput')?.value || ''
    };

    // Проверка на пустое наименование (если выбрано)
    if (data.name === '' && [...checkboxes].some(cb => cb.dataset.field === 'name')) {
        alert('❌ Поле "Наименование" пустое. Заполните его.');
        return;
    }

    // Соответствие полей и заголовков
    const fieldMap = {
        'section': 'Раздел',
        'point': 'Пункт',
        'name': 'Наименование',
        'model': 'Модель',
        'manufacturer': 'Фирма производитель',
        'country': 'Страна производства',
        'serial': 'Серийный номер'
    };

    // Формируем массивы заголовков и значений в том порядке, в котором отмечены чекбоксы
    const headers = [];
    const values = [];
    checkboxes.forEach(cb => {
        const field = cb.dataset.field;
        headers.push(fieldMap[field] || field);
        values.push(data[field] || '');
    });

    // Создаём книгу и лист
    const workbook = new ExcelJS.Workbook();
    const worksheet = workbook.addWorksheet('Данные');

    // --- ШАПКА (только жирный шрифт, без фона) ---
    const headerRow = worksheet.addRow(headers);
    headerRow.font = {
        name: 'Times New Roman',
        size: 12,
        bold: true,
        color: { argb: 'FF000000' } // чёрный цвет
    };
    headerRow.alignment = {
        horizontal: 'center',
        vertical: 'middle'
    };
    headerRow.height = 25;

    // --- ДАННЫЕ (обычный шрифт, выравнивание по левому краю) ---
    const dataRow = worksheet.addRow(values);
    dataRow.font = {
        name: 'Times New Roman',
        size: 12
    };
    dataRow.alignment = {
        horizontal: 'left',
        vertical: 'middle'
    };
    dataRow.height = 20;

    // --- Автоматическая ширина колонок (чтобы текст помещался) ---
    worksheet.columns.forEach((col, index) => {
        let maxLength = 0;
        const column = worksheet.getColumn(index + 1);
        column.eachCell({ includeEmpty: true }, (cell) => {
            const value = cell.value ? cell.value.toString() : '';
            maxLength = Math.max(maxLength, value.length);
        });
        // Устанавливаем ширину с запасом, но не более 60
        column.width = Math.min(Math.max(maxLength + 4, 20), 60);
    });

    // --- Сохраняем и скачиваем ---
    try {
        const buffer = await workbook.xlsx.writeBuffer();
        const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
        const now = new Date();
        const dateStr = now.toISOString().slice(0,10);
        const fileName = `Экспорт_полей_${dateStr}.xlsx`;
        saveAs(blob, fileName);
        // Файл скачивается без лишних уведомлений
    } catch (e) {
        alert('❌ Ошибка при создании Excel: ' + e.message);
        console.error(e);
    }
}





// ============================================================
// 5. ИНИЦИАЛИЗАЦИЯ (ОДИН РАЗ)
// ============================================================
document.addEventListener('DOMContentLoaded', function() {
    initData();

    // === ПОДСКАЗКИ ===
    const eqInput = document.getElementById('equipmentName');
    const sectionDisplay = document.getElementById('sectionDisplay');
    const pointDisplay = document.getElementById('pointDisplay');
    const eqSuggestions = document.getElementById('suggestionsList');
    const findBtn = document.getElementById('findBtn');
    const modelInput = document.getElementById('modelInput');
    const firmInput = document.getElementById('firmInput');
    const countryInput = document.getElementById('countryInput');
    const countryFirmInput = document.getElementById('countryFirmInput');
    const modelSuggestions = document.getElementById('modelSuggestionsList');

    if (eqInput) {
        function updateEquipment() {
            const query = eqInput.value;
            if (!query) {
                sectionDisplay.textContent = '—';
                pointDisplay.textContent = '—';
                eqSuggestions.style.display = 'none';
                return;
            }
            const matches = findEquipment(query);
            if (!matches.length) {
                sectionDisplay.textContent = 'Не найдено';
                pointDisplay.textContent = 'Не найдено';
                showEquipmentSuggestions([], eqInput, sectionDisplay, pointDisplay, eqSuggestions);
                return;
            }
            if (matches.length === 1) {
                sectionDisplay.textContent = matches[0].section;
                pointDisplay.textContent = matches[0].point;
                showEquipmentSuggestions([], eqInput, sectionDisplay, pointDisplay, eqSuggestions);
                return;
            }
            showEquipmentSuggestions(matches, eqInput, sectionDisplay, pointDisplay, eqSuggestions);
        }
        eqInput.addEventListener('input', updateEquipment);
        eqInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                const first = eqSuggestions.querySelector('.suggestion-item');
                first ? first.click() : updateEquipment();
            }
        });
        if (findBtn) findBtn.addEventListener('click', updateEquipment);
    }

    if (modelInput) {
        function updateModels() {
            const query = modelInput.value;
            if (!query) {
                modelSuggestions.style.display = 'none';
                modelSuggestions.innerHTML = '';
                firmInput.value = countryInput.value = countryFirmInput.value = '';
                return;
            }
            const matches = findModels(query);
            if (!matches.length) {
                modelSuggestions.style.display = 'none';
                modelSuggestions.innerHTML = '';
                return;
            }
            const exact = matches.find(m => m.model.toLowerCase() === query.toLowerCase());
            if (exact) {
                firmInput.value = exact.manufacturer;
                countryInput.value = exact.country;
                countryFirmInput.value = exact.country;
                if (!eqInput.value.trim() && exact.equipment) {
                    eqInput.value = exact.equipment;
                    setTimeout(() => eqInput.dispatchEvent(new Event('input', { bubbles: true })), 50);
                }
                modelSuggestions.style.display = 'none';
                modelSuggestions.innerHTML = '';
                firmInput.style.borderColor = '#66bb6a';
                countryInput.style.borderColor = '#66bb6a';
                countryFirmInput.style.borderColor = '#66bb6a';
                setTimeout(() => {
                    firmInput.style.borderColor = '';
                    countryInput.style.borderColor = '';
                    countryFirmInput.style.borderColor = '';
                }, 2000);
                return;
            }
            showModelSuggestions(matches, modelInput, firmInput, countryInput, countryFirmInput, eqInput, modelSuggestions);
        }
        let modelTimeout;
        modelInput.addEventListener('input', function() {
            clearTimeout(modelTimeout);
            const query = this.value.trim();
            if (!query) {
                modelSuggestions.style.display = 'none';
                modelSuggestions.innerHTML = '';
                firmInput.value = countryInput.value = countryFirmInput.value = '';
                return;
            }
            modelTimeout = setTimeout(updateModels, 200);
        });
        modelInput.addEventListener('keydown', function(e) {
            if (e.key === 'Enter') {
                const first = modelSuggestions.querySelector('.suggestion-item');
                first ? first.click() : updateModels();
            }
        });
    }

    // Закрытие списков
    document.addEventListener('click', function(e) {
        if (!e.target.closest('.input-with-button')) {
            const eqList = document.getElementById('suggestionsList');
            if (eqList) { eqList.style.display = 'none'; eqList.innerHTML = ''; }
        }
        if (!e.target.closest('.model-input-wrapper')) {
            const modelList = document.getElementById('modelSuggestionsList');
            if (modelList) { modelList.style.display = 'none'; modelList.innerHTML = ''; }
        }
    });

    // === КНОПКИ ЭКСПОРТА ===
    const exportHeaderBtn = document.getElementById('exportHeaderBtn');
    if (exportHeaderBtn) exportHeaderBtn.addEventListener('click', exportSelectedFields);
    const exportFormBtn = document.getElementById('exportFormBtn');
    if (exportFormBtn) exportFormBtn.addEventListener('click', exportSelectedFields);

    // === ИЗОБРАЖЕНИЯ ===
    const fileInput = document.getElementById('fileInput');
    const uploadBox = document.getElementById('uploadBox');
    const previewContainer = document.getElementById('previewContainer');
    const lightboxOverlay = document.getElementById('lightboxOverlay');
    const lightboxImage = document.getElementById('lightboxImage');
    const lightboxPrev = document.getElementById('lightboxPrev');
    const lightboxNext = document.getElementById('lightboxNext');
    const lightboxCounter = document.getElementById('lightboxCounter');
    const lightboxClose = document.getElementById('lightboxClose');

    if (fileInput && uploadBox) {
        uploadBox.addEventListener('click', () => fileInput.click());
        fileInput.addEventListener('change', function() {
            const files = Array.from(this.files);
            if (!files.length) return;
            const empty = previewContainer.querySelector('.empty-text');
            if (empty) empty.remove();
            files.forEach(file => {
                const reader = new FileReader();
                reader.onload = function(ev) {
                    const src = ev.target.result;
                    imageList.push(src);
                    const item = document.createElement('div');
                    item.className = 'preview-item';
                    const img = document.createElement('img');
                    img.src = src;
                    img.alt = file.name;
                    img.dataset.index = imageList.length - 1;
                    const removeBtn = document.createElement('button');
                    removeBtn.className = 'remove-btn';
                    removeBtn.textContent = '×';
                    removeBtn.addEventListener('click', function(e) {
                        e.stopPropagation();
                        const idx = parseInt(img.dataset.index);
                        imageList.splice(idx, 1);
                        document.querySelectorAll('.preview-item img').forEach((el, i) => {
                            el.dataset.index = i;
                        });
                        item.remove();
                        if (!previewContainer.children.length) {
                            const emptyDiv = document.createElement('div');
                            emptyDiv.className = 'empty-text';
                            emptyDiv.textContent = '📸 Здесь будут миниатюры загруженных фото';
                            previewContainer.appendChild(emptyDiv);
                        }
                        if (lightboxOverlay.classList.contains('active')) {
                            if (currentIndex >= imageList.length) currentIndex = Math.max(0, imageList.length - 1);
                            if (imageList.length === 0) {
                                lightboxOverlay.classList.remove('active');
                                document.body.style.overflow = '';
                            } else {
                                updateLightbox(currentIndex);
                            }
                        }
                    });
                    img.addEventListener('click', function() {
                        const idx = parseInt(this.dataset.index);
                        openLightbox(idx);
                    });
                    item.appendChild(img);
                    item.appendChild(removeBtn);
                    previewContainer.appendChild(item);
                };
                reader.readAsDataURL(file);
            });
            this.value = '';
        });
    }

    let imageList = [];
    let currentIndex = 0;

    function openLightbox(index) {
        if (!imageList.length) return;
        currentIndex = index;
        updateLightbox(index);
        lightboxOverlay.classList.add('active');
        document.body.style.overflow = 'hidden';
    }

    function updateLightbox(index) {
        if (index < 0) index = 0;
        if (index >= imageList.length) index = imageList.length - 1;
        currentIndex = index;
        lightboxImage.src = imageList[currentIndex];
        lightboxCounter.textContent = `${currentIndex + 1} / ${imageList.length}`;
        lightboxPrev.style.display = imageList.length > 1 ? 'block' : 'none';
        lightboxNext.style.display = imageList.length > 1 ? 'block' : 'none';
    }

    function closeLightbox() {
        lightboxOverlay.classList.remove('active');
        document.body.style.overflow = '';
    }

    if (lightboxPrev) {
        lightboxPrev.addEventListener('click', function(e) {
            e.stopPropagation();
            if (currentIndex > 0) {
                currentIndex--;
                updateLightbox(currentIndex);
            }
        });
    }
    if (lightboxNext) {
        lightboxNext.addEventListener('click', function(e) {
            e.stopPropagation();
            if (currentIndex < imageList.length - 1) {
                currentIndex++;
                updateLightbox(currentIndex);
            }
        });
    }
    if (lightboxClose) {
        lightboxClose.addEventListener('click', function(e) {
            e.stopPropagation();
            closeLightbox();
        });
    }
    if (lightboxOverlay) {
        lightboxOverlay.addEventListener('click', function(e) {
            if (e.target === this) closeLightbox();
        });
    }
    document.addEventListener('keydown', function(e) {
        if (e.key === 'Escape' && lightboxOverlay.classList.contains('active')) {
            closeLightbox();
        }
        if (lightboxOverlay.classList.contains('active')) {
            if (e.key === 'ArrowLeft') lightboxPrev?.click();
            else if (e.key === 'ArrowRight') lightboxNext?.click();
        }
    });

    console.log('✅ Бот с подсказками, предпросмотром и экспортом готов');
});