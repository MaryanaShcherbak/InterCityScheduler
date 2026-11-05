import { ExcelReader } from './services/excelReader.js';
import { ExcelWriter } from './services/excelWriter.js';
import { DictionaryService } from './services/dictionaryService.js';

const GOOGLE_APP_URL = "https://script.google.com/macros/s/AKfycbx1hpIw1RKBT0HauOKUVIGvbgENB9ryG_79MtxGKARK6KYlFf9dtpLGYHLFwUbsOOtR/exec";

let schedules = [];
let appDictionaries = null;
let currentRenderSession = 0;

// UI элементы
const fileInput = document.getElementById('fileInput');
const savePrologBtn = document.getElementById('savePrologBtn');
const saveNewsletterBtn = document.getElementById('saveNewsletterBtn');
const saveTrafficLightBtn = document.getElementById('saveTrafficLightBtn');
const pointsCountInput = document.getElementById('pointsCountInput');
const sourceLabel = document.getElementById('sourceLabel');
const statusInfo = document.getElementById('statusInfo');
const schedulesTable = document.getElementById('schedulesTable');

// Элементы прогресс-бара
const progressContainer = document.getElementById('progressContainer');
const progressBarFill = document.getElementById('progressBarFill');
const progressText = document.getElementById('progressText');

// Элементы модального окна
const nodesModal = document.getElementById('nodesModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const saveNodesBtn = document.getElementById('saveNodesBtn');
const unknownNodesTableBody = document.querySelector('#unknownNodesTable tbody');

// Управление прогресс-баром
function showIndeterminateProgress(label) {
  progressContainer.style.display = 'flex';
  progressBarFill.classList.add('indeterminate');
  progressBarFill.style.width = '100%';
  progressText.textContent = "...";
  if (label) statusInfo.textContent = label;
}

function setProgress(percent, label = null) {
  progressContainer.style.display = 'flex';
  progressBarFill.classList.remove('indeterminate');
  const clamped = Math.max(0, Math.min(100, Math.round(percent)));
  progressBarFill.style.width = `${clamped}%`;
  progressText.textContent = `${clamped}%`;
  if (label) statusInfo.textContent = label;
}

function hideProgress() {
  setTimeout(() => {
    progressContainer.style.display = 'none';
    progressBarFill.classList.remove('indeterminate');
    progressBarFill.style.width = '0%';
  }, 300);
}

// Быстрая загрузка без тормозов
fileInput.addEventListener('change', async (e) => {
  const files = Array.from(e.target.files);
  if (files.length === 0) return;

  try {
    showLoader(`Зчитування ${files.length > 1 ? files.length + ' файлів' : files[0].name}...`);
    statusInfo.textContent = "Зчитування файлів...";

    // 50 мс дают браузеру время плавно отрисовать крутилку на экране
    await new Promise(r => setTimeout(r, 50));
    await initPromise;

    let combinedSchedules = [];
    let detectedSources = new Set();
    let lastColor = "#0000cd";

    for (let i = 0; i < files.length; i++) {
      const file = files[i];
      if (files.length > 1) {
        showLoader(`Зчитування файлу ${i + 1} з ${files.length}: ${file.name}...`);
        await new Promise(r => setTimeout(r, 20));
      }

      const result = await ExcelReader.readWorkbook(file, appDictionaries);
      combinedSchedules.push(...result.schedules);
      detectedSources.add(result.sourceLabel);
      lastColor = result.sourceColor;
    }

    showLoader("Сортування та підготовка таблиці...");
    await new Promise(r => setTimeout(r, 20));

    schedules = ExcelReader.sortSchedules(combinedSchedules);

    sourceLabel.textContent = Array.from(detectedSources).join(" + ") || "Завантажено";
    sourceLabel.style.color = detectedSources.size === 1 ? lastColor : "#2a52be";
    statusInfo.textContent = `Успішно завантажено графіків: ${schedules.length} (файлів: ${files.length})`;

    renderTableChunked(schedules, 100);
    checkUnknownNodes();

  } catch (err) {
    console.error("Помилка читання файлів:", err);
    alert("Помилка при читанні файлів: " + err.message);
    statusInfo.textContent = "Помилка завантаження";
  } finally {
    hideLoader();
    fileInput.value = "";
  }
});


// Загрузка справочников
const initPromise = (async function initApp() {
  statusInfo.textContent = "Завантаження довідників...";
  try {
    appDictionaries = await DictionaryService.loadFromGoogleAPI(GOOGLE_APP_URL);
    statusInfo.textContent = "Довідники успішно завантажені";
  } catch (e) {
    console.error("Помилка завантаження довідників:", e);
    statusInfo.textContent = "Увага: не вдалося завантажити довідники з Google";
  }
  return appDictionaries;
})();

// Обработка загрузки файлов с точным прогрессом чтения
fileInput.addEventListener('change', async (e) => {
  const files = Array.from(e.target.files);
  if (files.length === 0) return;

  try {
    statusInfo.textContent = "Очікування довідників...";
    await initPromise;

    setProgress(0, `Початок завантаження (${files.length} файл.)...`);

    let combinedSchedules = [];
    let detectedSources = new Set();
    let lastColor = "#0000cd";

    for (let i = 0; i < files.length; i++) {
      const file = files[i];

      const result = await ExcelReader.readWorkbook(file, appDictionaries, (filePercent, stepText) => {
        // Вычисляем общий процент с учетом индекса текущего файла
        const globalPercent = ((i + filePercent / 100) / files.length) * 100;
        const label = files.length > 1 
          ? `[${i + 1}/${files.length}] ${file.name}: ${stepText}` 
          : `${file.name}: ${stepText}`;
        setProgress(globalPercent, label);
      });

      combinedSchedules.push(...result.schedules);
      detectedSources.add(result.sourceLabel);
      lastColor = result.sourceColor;
    }

    // Завершаем и скрываем прогресс-бар: чтение полностью окончено
    setProgress(100, "Файли успішно зчитано");
    hideProgress();

    schedules = ExcelReader.sortSchedules(combinedSchedules);

    sourceLabel.textContent = Array.from(detectedSources).join(" + ") || "Завантажено";
    sourceLabel.style.color = detectedSources.size === 1 ? lastColor : "#2a52be";
    statusInfo.textContent = `Успішно завантажено графіків: ${schedules.length} (файлів: ${files.length})`;

    // Отрисовка таблицы происходит независимо, без индикатора
    renderTableChunked(schedules, 100);

    checkUnknownNodes();

  } catch (err) {
    console.error("Помилка читання файлів:", err);
    alert("Помилка при читанні файлів: " + err.message);
    statusInfo.textContent = "Помилка завантаження";
    hideProgress();
  } finally {
    fileInput.value = "";
  }
});

// Аппаратное скрытие колонок количества точек через CSS
if (pointsCountInput) {
  pointsCountInput.addEventListener('input', () => {
    updatePointsVisibility(parseInt(pointsCountInput.value, 10) || 12);
  });
}

function updatePointsVisibility(pointsCount) {
  const maxVisibleT = pointsCount * 2;
  let styleTag = document.getElementById('dynamicPointsStyle');
  if (!styleTag) {
    styleTag = document.createElement('style');
    styleTag.id = 'dynamicPointsStyle';
    document.head.appendChild(styleTag);
  }

  let cssRules = [];
  for (let i = 3; i <= 22; i++) {
    if (i > maxVisibleT) {
      cssRules.push(`.th-T${i}, .col-T${i}`);
    }
  }

  styleTag.textContent = cssRules.length > 0 ? `${cssRules.join(', ')} { display: none !important; }` : "";
}

function ensureTableHeader() {
  const thead = schedulesTable.querySelector('thead');
  if (thead.innerHTML.trim() !== "") return;

  let theadHtml = `
    <tr>
      <th style="width: 110px;">Дата і час початку</th>
      <th style="width: 110px;">Дата і час завершення</th>
      <th style="width: 90px;">Код</th>
      <th style="width: 50px;">Номер</th>
      <th style="width: 110px;">Назва</th>
      <th style="width: 110px;">Група маршрутів</th>
      <th style="width: 110px;">Наряд</th>
      <th style="width: 130px;">Тип переміщення</th>
      <th style="width: 90px;">Тип ТЗ</th>
      <th style="width: 110px;">Формат завантаження</th>
      <th style="width: 220px;">Маршрут</th>
      <th style="width: 25px;">ПН</th>
      <th style="width: 25px;">ВТ</th>
      <th style="width: 25px;">СР</th>
      <th style="width: 25px;">ЧТ</th>
      <th style="width: 25px;">ПТ</th>
      <th style="width: 25px;">СБ</th>
      <th style="width: 25px;">НД</th>
      <th style="width: 50px;">Дедлайн</th>
      <th class="th-T1" style="width: 50px;">Подача</th>
      <th class="th-T2" style="width: 50px;">Виїзд</th>
  `;

  for (let i = 3; i <= 22; i++) {
    const isArr = (i % 2 !== 0);
    theadHtml += `<th class="th-T${i}" style="width: 50px;">${isArr ? 'Приїзд' : 'Виїзд'}</th>`;
  }

  theadHtml += `
      <th class="th-T23" style="width: 50px;">Приїзд</th>
      <th class="th-T24" style="width: 50px;">Виїзд</th>
      <th style="width: 180px;">Коментар</th>
      <th style="width: 60px;">Посегментні дедлайни</th>
      <th style="width: 60px;">ДВ</th>
      <th style="width: 60px;">П</th>
      <th style="width: 60px;">В</th>
    </tr>
  `;
  thead.innerHTML = theadHtml;
}

// Чанковая отрисовка строк без вмешательства в прогресс-бар
function renderTableChunked(data, chunkSize = 100) {
  ensureTableHeader();
  currentRenderSession++;
  const session = currentRenderSession;

  const tbody = schedulesTable.querySelector('tbody');
  tbody.innerHTML = "";

  const total = data.length;
  if (total === 0) return;

  let index = 0;
  updatePointsVisibility(parseInt(pointsCountInput ? pointsCountInput.value : 12, 10));

  function renderChunk() {
    if (session !== currentRenderSession) return;

    const end = Math.min(index + chunkSize, total);
    let chunkHtml = "";

    for (let i = index; i < end; i++) {
      chunkHtml += getRowHtml(data[i]);
    }

    tbody.insertAdjacentHTML('beforeend', chunkHtml);
    index = end;

    if (index < total) {
      requestAnimationFrame(renderChunk);
    }
  }

  renderChunk();
}

function getRowHtml(s) {
  let rowStyle = "";
  if (s.endDate && !isNaN(s.endDate.getTime())) {
    rowStyle = 'style="background-color: #fa8072;"';
  } else if (s.isNew) {
    rowStyle = 'style="background-color: #90ee90;"';
  } else if (s.changed) {
    rowStyle = 'style="background-color: #ffff00;"';
  }

  let html = `
    <tr ${rowStyle}>
      <td>${formatDateTime(s.startDate)}</td>
      <td>${formatDateTime(s.endDate)}</td>
      <td>${s.grf || ""}</td>
      <td>${s.number || ""}</td>
      <td>${s.name || ""}</td>
      <td>${s.routeGroup || ""}</td>
      <td>${s.order || ""}</td>
      <td>${s.movingType || ""}</td>
      <td>${s.vehicleType || ""}</td>
      <td>${s.loadFormat || ""}</td>
      <td>${s.route || ""}</td>
      <td><input type="checkbox" class="grid-chk" ${s.logMon ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.logTue ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.logWed ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.logThu ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.logFri ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.logSat ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.logSun ? "checked" : ""}></td>
      <td><input type="checkbox" class="grid-chk" ${s.deadline ? "checked" : ""}></td>
  `;

  for (let i = 1; i <= 24; i++) {
    html += `<td class="col-T${i}" style="text-align: center;">${formatTime(s[`t${i}`])}</td>`;
  }

  html += `
      <td>${s.comment || ""}</td>
      <td>${s.deadlineSegments || ""}</td>
      <td>${s.deadlineSegmentsSP || ""}</td>
      <td>${s.deadlineSegmentsP || ""}</td>
      <td>${s.deadlineSegmentsC || ""}</td>
    </tr>
  `;
  return html;
}

function formatDateTime(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const day = String(d.getDate()).padStart(2, '0');
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const year = d.getFullYear();
  const h = String(d.getHours()).padStart(2, '0');
  const m = String(d.getMinutes()).padStart(2, '0');
  return `${day}.${month}.${year} ${h}:${m}`;
}

function formatTime(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const h = String(d.getHours()).padStart(2, '0');
  const m = String(d.getMinutes()).padStart(2, '0');
  return `${h}:${m}`;
}

// Проверка неизвестных узлов
function checkUnknownNodes() {
  const missingAX = Array.from(DictionaryService.unknownNodesProlog).filter(n => n && n.trim());
  const missingPravosl = Array.from(DictionaryService.unknownNodesNews).filter(n => n && n.trim());

  if (missingAX.length === 0 && missingPravosl.length === 0) return;

  unknownNodesTableBody.innerHTML = "";
  
  missingAX.forEach(node => {
    unknownNodesTableBody.appendChild(createNodeRow(node, ""));
  });

  missingPravosl.forEach(node => {
    unknownNodesTableBody.appendChild(createNodeRow("", node));
  });

  nodesModal.style.display = 'flex';
}

function createNodeRow(ax, pravosl) {
  const tr = document.createElement('tr');
  tr.innerHTML = `
    <td><input type="text" class="node-ax" value="${ax || ''}" placeholder="Код АХ (напр: Kv-385)"></td>
    <td><input type="text" class="node-pravosl" value="${pravosl || ''}" placeholder="Назва православна"></td>
    <td><input type="text" class="node-city" placeholder="Місто"></td>
    <td><input type="text" class="node-cityring" placeholder="Місто (Кільце)"></td>
    <td><input type="text" class="node-cat" placeholder="В/П"></td>
  `;
  return tr;
}

saveNodesBtn.addEventListener('click', async () => {
  saveNodesBtn.textContent = "Збереження...";
  saveNodesBtn.disabled = true;

  const newNodes = [];
  const rows = unknownNodesTableBody.querySelectorAll('tr');
  
  rows.forEach(row => {
    const ax = row.querySelector('.node-ax').value.trim();
    const pravosl = row.querySelector('.node-pravosl').value.trim();
    const city = row.querySelector('.node-city').value.trim();
    const cityRing = row.querySelector('.node-cityring').value.trim();
    const nodeCat = row.querySelector('.node-cat').value.trim();

    if (ax && pravosl) {
      newNodes.push({ ax, pravosl, city, cityRing, nodeCat });
    }
  });

  if (newNodes.length === 0) {
    alert("Заповніть код АХ та Православну назву для вузлів.");
    saveNodesBtn.textContent = "💾 Зберегти у довідник";
    saveNodesBtn.disabled = false;
    return;
  }

  try {
    await DictionaryService.saveNewNodes(newNodes);
    statusInfo.textContent = "Оновлення довідників...";
    appDictionaries = await DictionaryService.loadFromGoogleAPI(GOOGLE_APP_URL);
    
    DictionaryService.setCities(schedules, appDictionaries);
    DictionaryService.setAXCities(schedules, appDictionaries);
    schedules.forEach(s => s.setRoute());
    renderTableChunked(schedules, 100);

    nodesModal.style.display = 'none';
    alert("Нові вузли успішно додані до довідника!");
  } catch (err) {
    console.error(err);
    alert("Помилка збереження: " + err.message);
  } finally {
    saveNodesBtn.textContent = "💾 Зберегти у довідник";
    saveNodesBtn.disabled = false;
  }
});

closeModalBtn.addEventListener('click', () => {
  nodesModal.style.display = 'none';
});

// Экспорт
savePrologBtn.addEventListener('click', () => {
  if (schedules.length === 0) return alert("Немає даних для збереження");
  ExcelWriter.exportToProlog(schedules, "Графіки_ProCTOlog.xlsx", false);
});

saveNewsletterBtn.addEventListener('click', () => {
  if (schedules.length === 0) return alert("Немає даних для збереження");
  ExcelWriter.exportToNewsLetter(schedules, "Графіки_Розсилка.xlsx");
});

if (saveTrafficLightBtn) {
  saveTrafficLightBtn.addEventListener('click', () => {
    if (schedules.length === 0) return alert("Немає даних для збереження");
    ExcelWriter.exportToProlog(schedules, "Графіки_Світлофор.xlsx", true);
  });
}

// Элементы оверлея-крутилки
const globalLoader = document.getElementById('globalLoader');
const loaderText = document.getElementById('loaderText');

function showLoader(text = "Обробка даних...") {
  if (loaderText) loaderText.textContent = text;
  if (globalLoader) globalLoader.style.display = 'flex';
}

function hideLoader() {
  if (globalLoader) globalLoader.style.display = 'none';
}