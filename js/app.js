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

    initVirtualTable(schedules, 100);
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
    populateFilterDropdowns(schedules);
    schedulesFiltered = schedules;
    sourceLabel.textContent = Array.from(detectedSources).join(" + ") || "Завантажено";
    sourceLabel.style.color = detectedSources.size === 1 ? lastColor : "#2a52be";
    statusInfo.textContent = `Успішно завантажено графіків: ${schedules.length} (файлів: ${files.length})`;

    // Отрисовка таблицы происходит независимо, без индикатора
    initVirtualTable(schedules, 100);

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

// --- ВИРТУАЛИЗАЦИЯ СКРОЛЛА ---
const ROW_HEIGHT = 22; // Фиксированная высота строки из стилей
const BUFFER_ROWS = 10; // Дополнительные строки сверху и снизу для плавности скролла

let currentDataToRender = [];
let isScrollThrottled = false;

const gridContainer = document.querySelector('.grid-container');

// Слушатель скролла с частотой обновления экрана (60 FPS)
gridContainer.addEventListener('scroll', () => {
  if (!isScrollThrottled) {
    window.requestAnimationFrame(() => {
      renderVisibleRows();
      isScrollThrottled = false;
    });
    isScrollThrottled = true;
  }
}, { passive: true });

// Пересчет видимого окна при изменении размера окна браузера
window.addEventListener('resize', () => {
  if (currentDataToRender.length > 0) renderVisibleRows();
});

/**
 * Инициализация отображения набора данных с виртуальным скроллом
 */
function initVirtualTable(data, resetScroll = true) {
  ensureTableHeader();
  currentDataToRender = data || [];

  if (resetScroll) {
    gridContainer.scrollTop = 0;
  }

  updatePointsVisibility(parseInt(pointsCountInput ? pointsCountInput.value : 12, 10));
  renderVisibleRows();
}

/**
 * Отрисовка ТОЛЬКО видимых 30-40 строк в окне просмотра
 */
function renderVisibleRows() {
  const tbody = schedulesTable.querySelector('tbody');
  const total = currentDataToRender.length;

  if (total === 0) {
    tbody.innerHTML = '<tr><td colspan="50" style="text-align: center; padding: 20px;">Немає даних для відображення</td></tr>';
    return;
  }

  const containerHeight = gridContainer.clientHeight || 800;
  const scrollTop = gridContainer.scrollTop;

  // Вычисляем видимый диапазон индексов строк
  let startIndex = Math.floor(scrollTop / ROW_HEIGHT) - BUFFER_ROWS;
  if (startIndex < 0) startIndex = 0;

  let endIndex = Math.ceil((scrollTop + containerHeight) / ROW_HEIGHT) + BUFFER_ROWS;
  if (endIndex > total) endIndex = total;

  // Вычисляем высоту верхнего и нижнего отступов (распорок)
  const topSpacerHeight = startIndex * ROW_HEIGHT;
  const bottomSpacerHeight = (total - endIndex) * ROW_HEIGHT;

  let html = '';

  // Верхняя распорка
  if (topSpacerHeight > 0) {
    html += `<tr class="virtual-spacer" style="height: ${topSpacerHeight}px;"><td colspan="50"></td></tr>`;
  }

  // Отрисовываем строго видимые строки (всего ~30-50 шт)
  for (let i = startIndex; i < endIndex; i++) {
    html += getRowHtml(currentDataToRender[i]);
  }

  // Нижняя распорка
  if (bottomSpacerHeight > 0) {
    html += `<tr class="virtual-spacer" style="height: ${bottomSpacerHeight}px;"><td colspan="50"></td></tr>`;
  }

  tbody.innerHTML = html;
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
    initVirtualTable(schedules, 100);

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
// Элементы фильтров
const fRoute = document.getElementById('fRoute');
const fCityA = document.getElementById('fCityA');
const fCityB = document.getElementById('fCityB');
const fVehType = document.getElementById('fVehType');
const fDeliveryType = document.getElementById('fDeliveryType');
const fRegion = document.getElementById('fRegion');
const fRouteGroup = document.getElementById('fRouteGroup');
const fLoadFormat = document.getElementById('fLoadFormat');
const fMovingType = document.getElementById('fMovingType');
const fLogDay = document.getElementById('fLogDay');
const fRelays = document.getElementById('fRelays');
const fDeadlines = document.getElementById('fDeadlines');
const fChanged = document.getElementById('fChanged');

const applyFiltersBtn = document.getElementById('applyFiltersBtn');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');

let schedulesFiltered = [];

// 1. Наполнение выпадающих списков уникальными значениями из загруженных данных
function populateFilterDropdowns(data) {
  const citiesA = new Set();
  const citiesB = new Set();
  const deliveryTypes = new Set();
  const movingTypes = new Set();
  const regions = new Set();
  const vehTypes = new Set();
  const loadFormats = new Set();

  data.forEach(s => {
    if (s.points && s.points.length > 0) {
      if (s.points[0].city) citiesA.add(s.points[0].city);
      const last = s.points[s.points.length - 1];
      if (last.city) citiesB.add(last.city);
    }
    if (s.deliveryType) deliveryTypes.add(s.deliveryType);
    if (s.movingType) movingTypes.add(s.movingType);
    if (s.region) regions.add(s.region);
    if (s.vehicleType) vehTypes.add(s.vehicleType);
    if (s.loadFormat) loadFormats.add(s.loadFormat);
  });

  fillSelect(fCityA, Array.from(citiesA).sort());
  fillSelect(fCityB, Array.from(citiesB).sort());
  fillSelect(fDeliveryType, Array.from(deliveryTypes).sort());
  fillSelect(fMovingType, Array.from(movingTypes).sort());
  fillSelect(fRegion, Array.from(regions).sort());
  fillSelect(fVehType, Array.from(vehTypes).sort());
  fillSelect(fLoadFormat, Array.from(loadFormats).sort());
}

function fillSelect(selectElem, items) {
  const currentVal = selectElem.value;
  selectElem.innerHTML = '<option value=""></option>';
  items.forEach(val => {
    const opt = document.createElement('option');
    opt.value = val;
    opt.textContent = val;
    selectElem.appendChild(opt);
  });
  selectElem.value = currentVal;
}

// 2. Алгоритм фильтрации (точное повторение FilterScheds из Form1.vb)
function filterSchedules() {
  let routeQuery = (fRoute.value || "").trim();
  let regex = null;
  if (routeQuery !== "") {
    if (!routeQuery.includes("*")) routeQuery = "*" + routeQuery + "*";
    // Преобразование wildcard * в регулярное выражение
    const pattern = "^" + routeQuery.replace(/[.+^${}()|[\]\\]/g, '\\$&').replace(/\*/g, '.*') + "$";
    regex = new RegExp(pattern, "i");
  }

  const cityA = fCityA.value;
  const cityB = fCityB.value;
  const vehType = fVehType.value;
  const delType = fDeliveryType.value;
  const region = fRegion.value;
  const movingType = fMovingType.value;
  const loadFormat = fLoadFormat.value;
  const logDay = fLogDay.value ? parseInt(fLogDay.value, 10) : null;
  const relays = fRelays.value;
  const deadlines = fDeadlines.value;
  const routeGroupQuery = (fRouteGroup.value || "").trim().toUpperCase();
  const onlyChanged = fChanged.checked;

  schedulesFiltered = schedules.filter(s => {
    // Маршрут (Route или RoutePravosl)
    if (regex) {
      const matchRoute = s.route && regex.test(s.route);
      const matchPravosl = s.routePravosl && regex.test(s.routePravosl);
      if (!matchRoute && !matchPravosl) return false;
    }

    // Города первой и последней точек
    const firstCity = s.points?.[0]?.city || "";
    const lastCity = s.points?.[s.points.length - 1]?.city || "";
    if (cityA !== "" && firstCity !== cityA) return false;
    if (cityB !== "" && lastCity !== cityB) return false;

    // Свойства графика
    if (vehType !== "" && s.vehicleType !== vehType) return false;
    if (delType !== "" && s.deliveryType !== delType) return false;
    if (region !== "" && s.region !== region) return false;
    if (movingType !== "" && s.movingType !== movingType) return false;
    if (loadFormat !== "" && s.loadFormat !== loadFormat) return false;

    // Дедлайны
    if (deadlines === "Так" && !s.deadline) return false;
    if (deadlines === "Ні" && s.deadline) return false;

    // Эстафеты (наличие буквы 'S' в имени)
    if (relays === "Так" && !(s.name && s.name.toUpperCase().includes("S"))) return false;
    if (relays === "Ні" && (s.name && s.name.toUpperCase().includes("S"))) return false;

    // Группа маршрута
    if (routeGroupQuery !== "" && !(s.routeGroup && s.routeGroup.toUpperCase().includes(routeGroupQuery))) return false;

    // Только с изменениями
    if (onlyChanged && !s.changed && !s.isNew) return false;

    // Логистический день
    if (logDay !== null && s.logisticDay !== logDay) return false;

    return true;
  });

  statusInfo.textContent = `Відображено: ${schedulesFiltered.length} з ${schedules.length}`;
  initVirtualTable(schedulesFiltered, 100);
}

// 3. Очистка фильтров (ClearFilter из Form1.vb)
function clearFilters() {
  fRoute.value = "";
  fCityA.value = "";
  fCityB.value = "";
  fVehType.value = "";
  fDeliveryType.value = "";
  fRegion.value = "";
  fRouteGroup.value = "";
  fLoadFormat.value = "";
  fMovingType.value = "";
  fLogDay.value = "";
  fRelays.value = "";
  fDeadlines.value = "";
  fChanged.checked = false;

  schedulesFiltered = schedules;
  statusInfo.textContent = `Успішно завантажено графіків: ${schedules.length}`;
  initVirtualTable(schedules, 100);
}

applyFiltersBtn.addEventListener('click', filterSchedules);
clearFiltersBtn.addEventListener('click', clearFilters);