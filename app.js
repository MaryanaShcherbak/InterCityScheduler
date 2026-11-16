import { ExcelReader } from './services/excelReader.js';
import { ExcelWriter } from './services/excelWriter.js';
import { DictionaryService } from './services/dictionaryService.js';

const GOOGLE_APP_URL = "https://script.google.com/macros/s/AKfycbx1hpIw1RKBT0HauOKUVIGvbgENB9ryG_79MtxGKARK6KYlFf9dtpLGYHLFwUbsOOtR/exec";

let schedules = [];
let schedulesFiltered = [];
let appDictionaries = null;
let currentDataToRender = [];

// Состояние выделения строк
let selectedIndices = new Set();
let lastClickedIndex = -1;

// Состояние редактирования в SchedForm
let currentEditingSchedule = null;
let currentEditingIndex = -1;
let localPoints = [];

// --- ЭЛЕМЕНТЫ ОСНОВНОГО ИНТЕРФЕЙСА ---
const fileInput = document.getElementById('fileInput');
const savePrologBtn = document.getElementById('savePrologBtn');
const saveNewsletterBtn = document.getElementById('saveNewsletterBtn');
const saveTrafficLightBtn = document.getElementById('saveTrafficLightBtn');
const pointsCountInput = document.getElementById('pointsCountInput');
const sourceLabel = document.getElementById('sourceLabel');
const statusInfo = document.getElementById('statusInfo');
const schedulesTable = document.getElementById('schedulesTable');
const gridContainer = document.querySelector('.grid-container');

// Оверлей-крутилка
const globalLoader = document.getElementById('globalLoader');
const loaderText = document.getElementById('loaderText');

// Кнопки действий
const btnCreateNew = document.getElementById('btnCreateNew');
const btnCreateCopy = document.getElementById('btnCreateCopy');
const btnDeleteSched = document.getElementById('btnDeleteSched');
const btnMassClose = document.getElementById('btnMassClose');

// Модальное окно массового закрытия
const massCloseModal = document.getElementById('massCloseModal');
const massCloseDateInput = document.getElementById('massCloseDateInput');
const confirmMassCloseBtn = document.getElementById('confirmMassCloseBtn');
const cancelMassCloseBtn = document.getElementById('cancelMassCloseBtn');
const closeMassCloseModalBtn = document.getElementById('closeMassCloseModalBtn');

// Модальное окно неизвестных узлов
const nodesModal = document.getElementById('nodesModal');
const closeModalBtn = document.getElementById('closeModalBtn');
const saveNodesBtn = document.getElementById('saveNodesBtn');
const unknownNodesTableBody = document.querySelector('#unknownNodesTable tbody');

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

// Элементы модального окна SchedForm
const schedModal = document.getElementById('schedModal');
const closeSchedModalBtn = document.getElementById('closeSchedModalBtn');
const saveSchedBtn = document.getElementById('saveSchedBtn');
const schedModalTitle = document.getElementById('schedModalTitle');

const sfIsActive = document.getElementById('sfIsActive');
const sfStartDate = document.getElementById('sfStartDate');
const sfEndDate = document.getElementById('sfEndDate');
const sfClearEndDateBtn = document.getElementById('sfClearEndDateBtn');
const sfVehType = document.getElementById('sfVehType');
const sfDeliveryType = document.getElementById('sfDeliveryType');
const sfMovingType = document.getElementById('sfMovingType');
const sfLoadFormat = document.getElementById('sfLoadFormat');

const sfRouteGroup = document.getElementById('sfRouteGroup');
const sfNumber = document.getElementById('sfNumber');
const sfOrder = document.getElementById('sfOrder');
const sfName = document.getElementById('sfName');
const sfDeadline = document.getElementById('sfDeadline');
const sfDeadlines = document.getElementById('sfDeadlines');
const sfDeadlinesSP = document.getElementById('sfDeadlinesSP');
const sfDeadlinesP = document.getElementById('sfDeadlinesP');
const sfDeadlinesC = document.getElementById('sfDeadlinesC');
const sfComment = document.getElementById('sfComment');

const sfLogMon = document.getElementById('sfLogMon');
const sfLogTue = document.getElementById('sfLogTue');
const sfLogWed = document.getElementById('sfLogWed');
const sfLogThu = document.getElementById('sfLogThu');
const sfLogFri = document.getElementById('sfLogFri');
const sfLogSat = document.getElementById('sfLogSat');
const sfLogSun = document.getElementById('sfLogSun');

const sfPointsTableBody = document.querySelector('#sfPointsTable tbody');
const sfAddPointBtn = document.getElementById('sfAddPointBtn');
const sfRemovePointBtn = document.getElementById('sfRemovePointBtn');
// --- 10. ВКЛАДКА «ЗВІТИ» (4 ОТЧЁТА) ---
const reportsDropdown = document.getElementById('reportsDropdown');
const reportsMenuBtn = document.getElementById('reportsMenuBtn');
const repDeadlinerBtn = document.getElementById('repDeadlinerBtn');
const repBadDeadlinesBtn = document.getElementById('repBadDeadlinesBtn');
const repCheckTimesBtn = document.getElementById('repCheckTimesBtn');
const repCheckUnloadBtn = document.getElementById('repCheckUnloadBtn');

const reportModal = document.getElementById('reportModal');
const reportModalTitle = document.getElementById('reportModalTitle');
const reportModalHint = document.getElementById('reportModalHint');
const closeReportModalBtn = document.getElementById('closeReportModalBtn');
const exportReportExcelBtn = document.getElementById('exportReportExcelBtn');
const reportTable = document.getElementById('reportTable');

let currentReportType = "";
let currentReportRows = [];
let currentReportHeaders = [];
// --- УПРАВЛЕНИЕ СПИННЕРОМ ---
function showLoader(text = "Обробка даних...") {
  if (loaderText) loaderText.textContent = text;
  if (globalLoader) globalLoader.style.display = 'flex';
}

function hideLoader() {
  if (globalLoader) globalLoader.style.display = 'none';
}

// --- 1. ЗАГРУЗКА СПРАВОЧНИКОВ ---
const initPromise = (async function initApp() {
  if (statusInfo) statusInfo.textContent = "Завантаження довідників...";
  try {
    appDictionaries = await DictionaryService.loadFromGoogleAPI(GOOGLE_APP_URL);
    if (statusInfo) statusInfo.textContent = "Довідники успішно завантажені";
  } catch (e) {
    console.error("Помилка завантаження довідників:", e);
    if (statusInfo) statusInfo.textContent = "Увага: не вдалося завантажити довідники з Google";
  }
  return appDictionaries;
})();

// --- 2. ЗАГРУЗКА EXCEL ФАЙЛОВ ---
fileInput?.addEventListener('change', async (e) => {
  const files = Array.from(e.target.files);
  if (files.length === 0) return;

  try {
    showLoader(`Зчитування ${files.length > 1 ? files.length + ' файлів' : files[0].name}...`);
    if (statusInfo) statusInfo.textContent = "Зчитування файлів...";

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
    schedulesFiltered = schedules;
    selectedIndices.clear();
    lastClickedIndex = -1;

    populateFilterDropdowns(schedules);

    if (sourceLabel) {
      sourceLabel.textContent = Array.from(detectedSources).join(" + ") || "Завантажено";
      sourceLabel.style.color = detectedSources.size === 1 ? lastColor : "#2a52be";
    }
    if (statusInfo) {
      statusInfo.textContent = `Успішно завантажено графіків: ${schedules.length} (файлів: ${files.length})`;
    }

    initVirtualTable(schedulesFiltered, true);
    checkUnknownNodes();

  } catch (err) {
    console.error("Помилка читання файлів:", err);
    alert("Помилка при читанні файлів: " + err.message);
    if (statusInfo) statusInfo.textContent = "Помилка завантаження";
  } finally {
    hideLoader();
    fileInput.value = "";
  }
});

// --- 3. ВИРТУАЛЬНАЯ ТАБЛИЦА И КОЛОНКИ ---
const ROW_HEIGHT = 22;
const BUFFER_ROWS = 10;
let isScrollThrottled = false;

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

  const fixWidth = (w) => `style="width: ${w}px; min-width: ${w}px; max-width: ${w}px;"`;

  let theadHtml = `
    <tr>
      <th ${fixWidth(110)}>Дата і час початку</th>
      <th ${fixWidth(110)}>Дата і час завершення</th>
      <th ${fixWidth(90)}>Код</th>
      <th ${fixWidth(50)}>Номер</th>
      <th ${fixWidth(110)}>Назва</th>
      <th ${fixWidth(110)}>Група маршрутів</th>
      <th ${fixWidth(110)}>Наряд</th>
      <th ${fixWidth(130)}>Тип переміщення</th>
      <th ${fixWidth(90)}>Тип ТЗ</th>
      <th ${fixWidth(110)}>Формат завантаження</th>
      <th ${fixWidth(220)}>Маршрут</th>
      <th ${fixWidth(25)}>ПН</th>
      <th ${fixWidth(25)}>ВТ</th>
      <th ${fixWidth(25)}>СР</th>
      <th ${fixWidth(25)}>ЧТ</th>
      <th ${fixWidth(25)}>ПТ</th>
      <th ${fixWidth(25)}>СБ</th>
      <th ${fixWidth(25)}>НД</th>
      <th ${fixWidth(50)}>Дедлайн</th>
      <th class="th-T1" ${fixWidth(50)}>Подача</th>
      <th class="th-T2" ${fixWidth(50)}>Виїзд</th>
  `;

  for (let i = 3; i <= 22; i++) {
    const isArr = (i % 2 !== 0);
    theadHtml += `<th class="th-T${i}" ${fixWidth(50)}>${isArr ? 'Приїзд' : 'Виїзд'}</th>`;
  }

  theadHtml += `
      <th class="th-T23" ${fixWidth(50)}>Приїзд</th>
      <th class="th-T24" ${fixWidth(50)}>Виїзд</th>
      <th ${fixWidth(180)}>Коментар</th>
      <th ${fixWidth(60)}>Посегментні дедлайни</th>
      <th ${fixWidth(40)}>ДВ</th>
      <th ${fixWidth(40)}>П</th>
      <th ${fixWidth(40)}>В</th>
    </tr>
  `;
  thead.innerHTML = theadHtml;
}

gridContainer?.addEventListener('scroll', () => {
  if (!isScrollThrottled) {
    window.requestAnimationFrame(() => {
      renderVisibleRows();
      isScrollThrottled = false;
    });
    isScrollThrottled = true;
  }
}, { passive: true });

window.addEventListener('resize', () => {
  if (currentDataToRender.length > 0) renderVisibleRows();
});

function initVirtualTable(data, resetScroll = true) {
  ensureTableHeader();
  currentDataToRender = data || [];
  if (resetScroll && gridContainer) {
    gridContainer.scrollTop = 0;
  }
  updatePointsVisibility(parseInt(pointsCountInput ? pointsCountInput.value : 12, 10));
  renderVisibleRows();
}

function renderVisibleRows() {
  const tbody = schedulesTable.querySelector('tbody');
  const total = currentDataToRender.length;

  if (total === 0) {
    tbody.innerHTML = '<tr><td colspan="50" style="text-align: center; padding: 20px;">Немає даних для відображення</td></tr>';
    return;
  }

  const containerHeight = gridContainer ? gridContainer.clientHeight : 800;
  const scrollTop = gridContainer ? gridContainer.scrollTop : 0;

  let startIndex = Math.floor(scrollTop / ROW_HEIGHT) - BUFFER_ROWS;
  if (startIndex < 0) startIndex = 0;

  let endIndex = Math.ceil((scrollTop + containerHeight) / ROW_HEIGHT) + BUFFER_ROWS;
  if (endIndex > total) endIndex = total;

  const topSpacerHeight = startIndex * ROW_HEIGHT;
  const bottomSpacerHeight = (total - endIndex) * ROW_HEIGHT;

  let html = '';
  if (topSpacerHeight > 0) {
    html += `<tr class="virtual-spacer" style="height: ${topSpacerHeight}px;"><td colspan="50"></td></tr>`;
  }

  for (let i = startIndex; i < endIndex; i++) {
    html += getRowHtml(currentDataToRender[i], i);
  }

  if (bottomSpacerHeight > 0) {
    html += `<tr class="virtual-spacer" style="height: ${bottomSpacerHeight}px;"><td colspan="50"></td></tr>`;
  }

  tbody.innerHTML = html;
}

function getRowHtml(s, index) {
  let rowStyle = "";
  if (s.endDate && !isNaN(s.endDate.getTime())) {
    rowStyle = 'style="background-color: #fa8072;"';
  } else if (s.isNew) {
    rowStyle = 'style="background-color: #90ee90;"';
  } else if (s.changed) {
    rowStyle = 'style="background-color: #ffff00;"';
  }

  const isSelected = selectedIndices.has(index) ? "selected-row" : "";

  let html = `
    <tr data-index="${index}" class="${isSelected}" ${rowStyle}>
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

// --- 4. ВЫДЕЛЕНИЕ СТРОК И ДВОЙНОЙ КЛИК ---
function updateSelectionUI() {
  const rows = schedulesTable.querySelectorAll('tbody tr[data-index]');
  rows.forEach((tr) => {
    const idx = parseInt(tr.dataset.index, 10);
    if (selectedIndices.has(idx)) {
      tr.classList.add('selected-row');
    } else {
      tr.classList.remove('selected-row');
    }
  });
}

const tableBody = schedulesTable.querySelector('tbody');

// Одиночный клик (без перерисовки DOM, чтобы работал dblclick)
tableBody.addEventListener('click', (e) => {
  const tr = e.target.closest('tr');
  if (!tr || tr.classList.contains('virtual-spacer')) return;

  const clickedIndex = parseInt(tr.dataset.index, 10);
  if (isNaN(clickedIndex)) return;

  if (e.ctrlKey || e.metaKey) {
    if (selectedIndices.has(clickedIndex)) {
      selectedIndices.delete(clickedIndex);
    } else {
      selectedIndices.add(clickedIndex);
    }
    lastClickedIndex = clickedIndex;
  } else if (e.shiftKey && lastClickedIndex !== -1) {
    selectedIndices.clear();
    const start = Math.min(lastClickedIndex, clickedIndex);
    const end = Math.max(lastClickedIndex, clickedIndex);
    for (let i = start; i <= end; i++) {
      selectedIndices.add(i);
    }
  } else {
    selectedIndices.clear();
    selectedIndices.add(clickedIndex);
    lastClickedIndex = clickedIndex;
  }

  updateSelectionUI();
});

// Двойной клик по строке -> открытие окна редактирования SchedForm
tableBody.addEventListener('dblclick', (e) => {
  const tr = e.target.closest('tr');
  if (!tr || tr.classList.contains('virtual-spacer')) return;

  const rowIndex = parseInt(tr.dataset.index, 10);
  const sched = currentDataToRender[rowIndex];
  if (sched) {
    openSchedModal(sched, rowIndex);
  }
});

function getActiveSelectedIndex() {
  if (lastClickedIndex !== -1 && selectedIndices.has(lastClickedIndex)) {
    return lastClickedIndex;
  }
  const first = selectedIndices.values().next().value;
  return first !== undefined ? first : -1;
}

// --- 5. КНОПКИ СОЗДАНИЯ, КОПИРОВАНИЯ, УДАЛЕНИЯ И МАССОВОГО ЗАКРЫТИЯ ---
btnCreateNew?.addEventListener('click', () => {
  const newSched = {
    isActive: true,
    startDate: new Date(),
    endDate: null,
    grf: "",
    number: "",
    name: "",
    routeGroup: "",
    order: "",
    movingType: "",
    vehicleType: "",
    loadFormat: "",
    route: "",
    routePravosl: "",
    deadline: false,
    deadlineSegments: "",
    deadlineSegmentsSP: "",
    deadlineSegmentsP: "",
    deadlineSegmentsC: "",
    comment: "",
    logMon: true,
    logTue: false,
    logWed: false,
    logThu: false,
    logFri: false,
    logSat: false,
    logSun: false,
    isNew: true,
    changed: true,
    changedFields: [],
    points: [
      { name: "", arrival: new Date(1975, 0, 1, 8, 0), departure: new Date(1975, 0, 1, 8, 30), movingType: "", distance: 0, loadRamp1: 0, loadRamp2: 0 },
      { name: "", arrival: new Date(1975, 0, 1, 14, 0), departure: new Date(1975, 0, 1, 14, 30), movingType: "", distance: 0, loadRamp1: 0, loadRamp2: 0 }
    ],
    setRoute() {
      if (!this.points || this.points.length === 0) return;
      this.route = this.points.map(p => p.name).join('_');
      this.routePravosl = this.points.map(p => p.namePravosl || p.name).join(' - ');
    },
    calculateTimes() {
      if (this.points && this.points.length > 0) {
        this.t1 = this.points[0].arrival;
        this.t2 = this.points[0].departure;
        this.t23 = this.points[this.points.length - 1].arrival;
        this.t24 = this.points[this.points.length - 1].departure;
      }
    },
    normalizeTimes() {}
  };

  openSchedModal(newSched, -1);
});

btnCreateCopy?.addEventListener('click', () => {
  const targetIndex = getActiveSelectedIndex();
  if (targetIndex < 0 || !currentDataToRender[targetIndex]) {
    alert("Виділіть графік у таблиці для копіювання!");
    return;
  }

  const baseSched = currentDataToRender[targetIndex];
  const copiedSched = cloneSchedule(baseSched);
  openSchedModal(copiedSched, targetIndex);
});

btnDeleteSched?.addEventListener('click', () => {
  const targetIndex = getActiveSelectedIndex();
  if (targetIndex < 0 || !currentDataToRender[targetIndex]) {
    alert("Виділіть рядок у таблиці для видалення!");
    return;
  }

  const sched = currentDataToRender[targetIndex];
  if (!sched.grf || sched.grf.trim() === "") {
    if (confirm(`Видалити створений графік "${sched.route || sched.name || 'новий'}"?`)) {
      const idxFiltered = schedulesFiltered.indexOf(sched);
      if (idxFiltered !== -1) schedulesFiltered.splice(idxFiltered, 1);

      const idxAll = schedules.indexOf(sched);
      if (idxAll !== -1) schedules.splice(idxAll, 1);

      selectedIndices.clear();
      lastClickedIndex = -1;
      renderVisibleRows();
      if (statusInfo) statusInfo.textContent = `Графік успішно видалено. Всього: ${schedules.length}`;
    }
  } else {
    alert("Для видалення діючого графіка необхідно лише проставити йому дату закриття");
  }
});

// Массовое закрытие
btnMassClose?.addEventListener('click', () => {
  if (selectedIndices.size === 0) {
    alert("Виділіть хоча б один рядок (графік) для закриття!");
    return;
  }

  const today = new Date();
  const yyyy = today.getFullYear();
  const mm = String(today.getMonth() + 1).padStart(2, '0');
  const dd = String(today.getDate()).padStart(2, '0');
  massCloseDateInput.value = `${yyyy}-${mm}-${dd}`;

  massCloseModal.style.display = 'flex';
  massCloseDateInput.focus();
});

confirmMassCloseBtn?.addEventListener('click', () => {
  const dateVal = massCloseDateInput.value;
  if (!dateVal) {
    alert("Оберіть дату закриття!");
    return;
  }

  const parts = dateVal.split('-');
  const closeDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 0, 0, 0);

  let changedCount = 0;
  selectedIndices.forEach((idx) => {
    const sched = currentDataToRender[idx];
    if (sched) {
      sched.endDate = closeDate;
      if (!sched.changedFields.includes("EndDate")) {
        sched.changedFields.push("EndDate");
      }
      sched.changed = true;
      changedCount++;
    }
  });

  massCloseModal.style.display = 'none';
  renderVisibleRows();
  alert(`Успішно закрито графіків: ${changedCount}`);
});

cancelMassCloseBtn?.addEventListener('click', () => { massCloseModal.style.display = 'none'; });
closeMassCloseModalBtn?.addEventListener('click', () => { massCloseModal.style.display = 'none'; });

// --- 6. ФИЛЬТРЫ ---
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
  if (!selectElem) return;
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

function filterSchedules() {
  let routeQuery = (fRoute.value || "").trim();
  let regex = null;
  if (routeQuery !== "") {
    if (!routeQuery.includes("*")) routeQuery = "*" + routeQuery + "*";
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
    if (regex) {
      const matchRoute = s.route && regex.test(s.route);
      const matchPravosl = s.routePravosl && regex.test(s.routePravosl);
      if (!matchRoute && !matchPravosl) return false;
    }

    const firstCity = s.points?.[0]?.city || "";
    const lastCity = s.points?.[s.points.length - 1]?.city || "";
    if (cityA !== "" && firstCity !== cityA) return false;
    if (cityB !== "" && lastCity !== cityB) return false;

    if (vehType !== "" && s.vehicleType !== vehType) return false;
    if (delType !== "" && s.deliveryType !== delType) return false;
    if (region !== "" && s.region !== region) return false;
    if (movingType !== "" && s.movingType !== movingType) return false;
    if (loadFormat !== "" && s.loadFormat !== loadFormat) return false;

    if (deadlines === "Так" && !s.deadline) return false;
    if (deadlines === "Ні" && s.deadline) return false;

    if (relays === "Так" && !(s.name && s.name.toUpperCase().includes("S"))) return false;
    if (relays === "Ні" && (s.name && s.name.toUpperCase().includes("S"))) return false;

    if (routeGroupQuery !== "" && !(s.routeGroup && s.routeGroup.toUpperCase().includes(routeGroupQuery))) return false;
    if (onlyChanged && !s.changed && !s.isNew) return false;
    if (logDay !== null && s.logisticDay !== logDay) return false;

    return true;
  });

  selectedIndices.clear();
  lastClickedIndex = -1;
  if (statusInfo) statusInfo.textContent = `Відображено: ${schedulesFiltered.length} з ${schedules.length}`;
  initVirtualTable(schedulesFiltered, true);
}

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
  selectedIndices.clear();
  lastClickedIndex = -1;
  if (statusInfo) statusInfo.textContent = `Успішно завантажено графіків: ${schedules.length}`;
  initVirtualTable(schedules, true);
}

applyFiltersBtn?.addEventListener('click', filterSchedules);
clearFiltersBtn?.addEventListener('click', clearFilters);

// --- 7. МОДАЛЬНОЕ ОКНО РЕДАКТИРОВАНИЯ SCHEDFORM ---
function openSchedModal(sched, index) {
  currentEditingSchedule = sched;
  currentEditingIndex = index;
  schedModalTitle.textContent = `Графік: ${sched.grf || sched.name || "Новий"}`;

  copySelectOptions(fVehType, sfVehType);
  copySelectOptions(fDeliveryType, sfDeliveryType);
  copySelectOptions(fMovingType, sfMovingType);
  copySelectOptions(fLoadFormat, sfLoadFormat);

  sfIsActive.checked = !!sched.isActive;
  sfStartDate.value = toInputDateTime(sched.startDate);
  sfEndDate.value = toInputDateTime(sched.endDate);

  sfVehType.value = sched.vehicleType || "";
  sfDeliveryType.value = sched.deliveryType || "";
  sfMovingType.value = sched.movingType || "";
  sfLoadFormat.value = sched.loadFormat || "";

  sfRouteGroup.value = sched.routeGroup || "";
  sfNumber.value = sched.number || "";
  sfOrder.value = sched.order || "";
  sfName.value = sched.name || "";
  sfDeadline.checked = !!sched.deadline;
  sfDeadlines.value = sched.deadlineSegments || "";
  sfDeadlinesSP.value = sched.deadlineSegmentsSP || "";
  sfDeadlinesP.value = sched.deadlineSegmentsP || "";
  sfDeadlinesC.value = sched.deadlineSegmentsC || "";
  sfComment.value = sched.comment || "";

  sfLogMon.checked = !!sched.logMon;
  sfLogTue.checked = !!sched.logTue;
  sfLogWed.checked = !!sched.logWed;
  sfLogThu.checked = !!sched.logThu;
  sfLogFri.checked = !!sched.logFri;
  sfLogSat.checked = !!sched.logSat;
  sfLogSun.checked = !!sched.logSun;

  localPoints = (sched.points || []).map(p => ({
    name: p.name || "",
    namePravosl: p.namePravosl || "",
    arrival: p.arrival ? new Date(p.arrival.getTime()) : null,
    departure: p.departure ? new Date(p.departure.getTime()) : null,
    movingType: p.movingType || "",
    distance: p.distance || 0,
    loadRamp1: p.loadRamp1 || 0,
    loadRamp2: p.loadRamp2 || 0,
    city: p.city || "",
    nodeCat: p.nodeCat || ""
  }));

  renderLocalPointsTable();
  schedModal.style.display = 'flex';
}

function copySelectOptions(sourceSelect, targetSelect) {
  if (sourceSelect && targetSelect) {
    targetSelect.innerHTML = sourceSelect.innerHTML;
  }
}

function toInputDateTime(d) {
  if (!d || !(d instanceof Date) || isNaN(d.getTime())) return "";
  const year = d.getFullYear();
  const month = String(d.getMonth() + 1).padStart(2, '0');
  const day = String(d.getDate()).padStart(2, '0');
  const hours = String(d.getHours()).padStart(2, '0');
  const minutes = String(d.getMinutes()).padStart(2, '0');
  return `${year}-${month}-${day}T${hours}:${minutes}`;
}

function renderLocalPointsTable() {
  sfPointsTableBody.innerHTML = "";
  localPoints.forEach((pt) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="text" class="pt-name" value="${pt.name}"></td>
      <td><input type="text" class="pt-arr" value="${formatTime(pt.arrival)}" placeholder="00:00" maxlength="5"></td>
      <td><input type="text" class="pt-dep" value="${formatTime(pt.departure)}" placeholder="00:00" maxlength="5"></td>
      <td><input type="text" class="pt-mov" value="${pt.movingType}"></td>
      <td><input type="number" class="pt-dist" value="${Math.max(0, pt.distance || 0)}" min="0" step="any" style="width: 55px;"></td>
      <td><input type="number" class="pt-r1" value="${Math.max(0, pt.loadRamp1 || 0)}" min="0" style="width: 45px;"></td>
      <td><input type="number" class="pt-r2" value="${Math.max(0, pt.loadRamp2 || 0)}" min="0" style="width: 45px;"></td>
    `;
    sfPointsTableBody.appendChild(tr);

    attachTimeInputAutoFormat(tr.querySelector('.pt-arr'));
    attachTimeInputAutoFormat(tr.querySelector('.pt-dep'));

    // Сброс отрицательного расстояния при ручном вводе или вставке
    const distInput = tr.querySelector('.pt-dist');
    distInput.addEventListener('input', () => {
      if (parseFloat(distInput.value) < 0) {
        distInput.value = 0;
      }
    });
  });
}

function attachTimeInputAutoFormat(input) {
  input.addEventListener('blur', () => {
    let val = input.value.trim().replace(/[^0-9]/g, '');
    if (val.length === 3 || val.length === 4) {
      val = val.padStart(4, '0');
      let h = Math.min(23, parseInt(val.slice(0, 2), 10));
      let m = Math.min(59, parseInt(val.slice(2), 10));
      input.value = `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
    }
  });
}

function parseTimeInput(timeStr) {
  if (!timeStr) return new Date(1975, 0, 1, 0, 0);
  let clean = timeStr.trim().replace(/[^0-9:]/g, '');
  if (!clean.includes(':') && (clean.length === 3 || clean.length === 4)) {
    clean = clean.padStart(4, '0');
    clean = clean.slice(0, 2) + ':' + clean.slice(2);
  }
  const parts = clean.split(':');
  const h = Math.min(23, Math.max(0, parseInt(parts[0], 10) || 0));
  const m = Math.min(59, Math.max(0, parseInt(parts[1], 10) || 0));
  return new Date(1975, 0, 1, h, m);
}

sfAddPointBtn?.addEventListener('click', () => {
  localPoints.push({
    name: "Новий",
    arrival: new Date(1975, 0, 1, 12, 0),
    departure: new Date(1975, 0, 1, 12, 10),
    movingType: sfMovingType.value || "",
    distance: 0,
    loadRamp1: 0,
    loadRamp2: 0
  });
  renderLocalPointsTable();
});

sfRemovePointBtn?.addEventListener('click', () => {
  if (localPoints.length > 2) {
    localPoints.pop();
    renderLocalPointsTable();
  } else {
    alert("У маршруті мусить бути мінімум дві точки!");
  }
});

sfClearEndDateBtn?.addEventListener('click', () => {
  sfEndDate.value = "";
});

function fillNumber() {
  const mov = sfMovingType.value;
  if (mov === "Порожній") { sfNumber.value = "88"; return; }
  if (mov === "Фрахт") { sfNumber.value = "99"; return; }
  if (mov === "Тара") { sfNumber.value = "999"; return; }

  const currentNum = sfNumber.value.trim();
  if (!currentNum || ["1", "2", "3", "4", "5", "6", "7"].includes(currentNum)) {
    if (sfLogMon.checked) sfNumber.value = "1";
    else if (sfLogTue.checked) sfNumber.value = "2";
    else if (sfLogWed.checked) sfNumber.value = "3";
    else if (sfLogThu.checked) sfNumber.value = "4";
    else if (sfLogFri.checked) sfNumber.value = "5";
    else if (sfLogSat.checked) sfNumber.value = "6";
    else if (sfLogSun.checked) sfNumber.value = "7";
  }
}

[sfLogMon, sfLogTue, sfLogWed, sfLogThu, sfLogFri, sfLogSat, sfLogSun].forEach(cb => {
  cb?.addEventListener('change', fillNumber);
});
sfMovingType?.addEventListener('change', fillNumber);

saveSchedBtn?.addEventListener('click', () => {
  if (!currentEditingSchedule) return;

const rows = sfPointsTableBody.querySelectorAll('tr');
  for (let i = 0; i < rows.length; i++) {
    const tr = rows[i];
    const rawDist = parseFloat(tr.querySelector('.pt-dist').value);

    if (!isNaN(rawDist) && rawDist < 0) {
      alert(`Відстань на вузлі "${tr.querySelector('.pt-name').value.trim() || i + 1}" не може бути від'ємною!`);
      return;
    }

    localPoints[i].name = tr.querySelector('.pt-name').value.trim();
    localPoints[i].arrival = parseTimeInput(tr.querySelector('.pt-arr').value);
    localPoints[i].departure = parseTimeInput(tr.querySelector('.pt-dep').value);
    localPoints[i].movingType = tr.querySelector('.pt-mov').value.trim();
    localPoints[i].distance = Math.max(0, rawDist || 0);
    localPoints[i].loadRamp1 = Math.max(0, parseInt(tr.querySelector('.pt-r1').value, 10) || 0);
    localPoints[i].loadRamp2 = Math.max(0, parseInt(tr.querySelector('.pt-r2').value, 10) || 0);
  }

  if (localPoints.length < 2) {
    alert("У маршруті мусить бути мінімум дві точки!");
    return;
  }
  if (sfLoadFormat.value.includes("Короб") && sfVehType.value === "20т") {
    alert("Невідповідний тип ТЗ до формату завантаження");
    return;
  }
  if (!sfLoadFormat.value && sfMovingType.value !== "Порожній") {
    alert("Потрібно заповнити формат завантаження");
    return;
  }
  if (!sfVehType.value) {
    alert("Потрібно вказати тип ТЗ");
    return;
  }
  if (!sfMovingType.value) {
    alert("Потрібно вказати тип завантаження");
    return;
  }
  if (!sfLogMon.checked && !sfLogTue.checked && !sfLogWed.checked &&
      !sfLogThu.checked && !sfLogFri.checked && !sfLogSat.checked && !sfLogSun.checked) {
    alert("Вкажи хоча б один робочий день");
    return;
  }

const sched = currentEditingSchedule;
  const p0 = localPoints[0];
  const pLast = localPoints[localPoints.length - 1];

  // Проверка времени погрузки на начальном узле (остается блокирующей)
  const tDiff0 = getTimeDiffMinutes(p0.arrival, p0.departure);
  if (tDiff0 > (23 * 60 + 50) || tDiff0 < 1) {
    alert(`На вузлі ${p0.name} стрьомний час на завантаження`);
    return;
  }

  // Проверка времени разгрузки на конечном узле (показывает предупреждение, но не блокирует сохранение)
  const tDiffLast = getTimeDiffMinutes(pLast.arrival, pLast.departure);
  if (tDiffLast > 120 || tDiffLast < 1) {
    alert(`На вузлі ${pLast.name} стрьомний час на вивантаження`);
  }

  // Проверка времени разгрузки на промежуточных узлах (показывает предупреждение, но не блокирует сохранение)
  for (let i = 1; i < localPoints.length - 1; i++) {
    const diff = getTimeDiffMinutes(localPoints[i].arrival, localPoints[i].departure);
    if (diff > 360 || diff < 1) {
      alert(`На вузлі ${localPoints[i].name} стрьомний час на вивантаження`);
    }
  }

  function trackChange(field, oldVal, newVal) {
    if (oldVal !== newVal && !sched.changedFields.includes(field)) {
      sched.changedFields.push(field);
    }
  }

  trackChange("VehicleType", sched.vehicleType, sfVehType.value);
  trackChange("DeliveryType", sched.deliveryType, sfDeliveryType.value);
  trackChange("MovingType", sched.movingType, sfMovingType.value);
  trackChange("LoadFormat", sched.loadFormat, sfLoadFormat.value);
  trackChange("RouteGroup", sched.routeGroup, sfRouteGroup.value);
  trackChange("Number", sched.number, sfNumber.value);
  trackChange("Order", sched.order, sfOrder.value);
  trackChange("Name", sched.name, sfName.value);
  trackChange("Deadline", sched.deadline, sfDeadline.checked);
  trackChange("DeadlineSegments", sched.deadlineSegments, sfDeadlines.value);
  trackChange("DeadlineSegmentsSP", sched.deadlineSegmentsSP, sfDeadlinesSP.value);
  trackChange("DeadlineSegmentsP", sched.deadlineSegmentsP, sfDeadlinesP.value);
  trackChange("DeadlineSegmentsC", sched.deadlineSegmentsC, sfDeadlinesC.value);
  trackChange("Comment", sched.comment, sfComment.value);

  trackChange("LogMon", sched.logMon, sfLogMon.checked);
  trackChange("LogTue", sched.logTue, sfLogTue.checked);
  trackChange("LogWed", sched.logWed, sfLogWed.checked);
  trackChange("LogThu", sched.logThu, sfLogThu.checked);
  trackChange("LogFri", sched.logFri, sfLogFri.checked);
  trackChange("LogSat", sched.logSat, sfLogSat.checked);
  trackChange("LogSun", sched.logSun, sfLogSun.checked);

  const selDays = [];
  if (sfLogMon.checked) selDays.push(1);
  if (sfLogTue.checked) selDays.push(2);
  if (sfLogWed.checked) selDays.push(3);
  if (sfLogThu.checked) selDays.push(4);
  if (sfLogFri.checked) selDays.push(5);
  if (sfLogSat.checked) selDays.push(6);
  if (sfLogSun.checked) selDays.push(7);

  if (sched.isNew) {
    selDays.forEach((d, idx) => {
      const s = (idx === 0) ? sched : cloneSchedule(sched);

      s.points = localPoints.map(p => ({
        ...p,
        arrival: p.arrival ? new Date(p.arrival.getTime()) : null,
        departure: p.departure ? new Date(p.departure.getTime()) : null
      }));
      s.setRoute();

      s.startDate = sfStartDate.value ? new Date(sfStartDate.value) : new Date();
      s.endDate = sfEndDate.value ? new Date(sfEndDate.value) : null;
      s.vehicleType = sfVehType.value;
      s.deliveryType = sfDeliveryType.value;
      s.movingType = sfMovingType.value;
      s.loadFormat = sfLoadFormat.value;
      s.routeGroup = sfRouteGroup.value;

      let calcNumber = sfNumber.value.trim();
      if (s.movingType === "Порожній") calcNumber = "88";
      else if (s.movingType === "Фрахт") calcNumber = "99";
      else if (s.movingType === "Тара") calcNumber = "999";
      else if (!calcNumber || ["1", "2", "3", "4", "5", "6", "7"].includes(calcNumber)) {
        calcNumber = String(d);
      }
      s.number = calcNumber;

      s.order = sfOrder.value;
      s.name = sfName.value;
      s.deadline = sfDeadline.checked;
      s.deadlineSegments = sfDeadlines.value;
      s.deadlineSegmentsSP = sfDeadlinesSP.value;
      s.deadlineSegmentsP = sfDeadlinesP.value;
      s.deadlineSegmentsC = sfDeadlinesC.value;
      s.comment = sfComment.value;

      s.logMon = (d === 1);
      s.logTue = (d === 2);
      s.logWed = (d === 3);
      s.logThu = (d === 4);
      s.logFri = (d === 5);
      s.logSat = (d === 6);
      s.logSun = (d === 7);
      s.logisticDay = d;

      s.astroDay = s.logisticDay;
      if (s.points[0]?.departure && s.points[0].departure.getHours() < 12) {
        s.astroDay += 1;
        if (s.astroDay === 8) s.astroDay = 1;
      }

      s.mon = (s.astroDay === 1);
      s.tue = (s.astroDay === 2);
      s.wed = (s.astroDay === 3);
      s.thu = (s.astroDay === 4);
      s.fri = (s.astroDay === 5);
      s.sat = (s.astroDay === 6);
      s.sun = (s.astroDay === 7);

      s.isNew = true;
      s.changed = true;
      s.calculateTimes();
      s.normalizeTimes();

      if (currentEditingIndex >= 0) {
        schedulesFiltered.splice(currentEditingIndex + 1, 0, s);
      } else {
        schedulesFiltered.unshift(s);
      }
      if (schedules !== schedulesFiltered) {
        schedules.unshift(s);
      }
    });
  } else {
    sched.points = localPoints;
    sched.setRoute();
    sched.startDate = sfStartDate.value ? new Date(sfStartDate.value) : sched.startDate;
    sched.endDate = sfEndDate.value ? new Date(sfEndDate.value) : null;
    sched.vehicleType = sfVehType.value;
    sched.deliveryType = sfDeliveryType.value;
    sched.movingType = sfMovingType.value;
    sched.loadFormat = sfLoadFormat.value;
    sched.routeGroup = sfRouteGroup.value;
    sched.number = sfNumber.value;
    sched.order = sfOrder.value;
    sched.name = sfName.value;
    sched.deadline = sfDeadline.checked;
    sched.deadlineSegments = sfDeadlines.value;
    sched.deadlineSegmentsSP = sfDeadlinesSP.value;
    sched.deadlineSegmentsP = sfDeadlinesP.value;
    sched.deadlineSegmentsC = sfDeadlinesC.value;
    sched.comment = sfComment.value;

    sched.logMon = sfLogMon.checked;
    sched.logTue = sfLogTue.checked;
    sched.logWed = sfLogWed.checked;
    sched.logThu = sfLogThu.checked;
    sched.logFri = sfLogFri.checked;
    sched.logSat = sfLogSat.checked;
    sched.logSun = sfLogSun.checked;

    sched.logisticDay = 1;
    if (sched.logTue) sched.logisticDay = 2;
    if (sched.logWed) sched.logisticDay = 3;
    if (sched.logThu) sched.logisticDay = 4;
    if (sched.logFri) sched.logisticDay = 5;
    if (sched.logSat) sched.logisticDay = 6;
    if (sched.logSun) sched.logisticDay = 7;

    sched.astroDay = sched.logisticDay;
    if (sched.points[0]?.departure && sched.points[0].departure.getHours() < 12) {
      sched.astroDay += 1;
      if (sched.astroDay === 8) sched.astroDay = 1;
    }

    sched.mon = (sched.astroDay === 1);
    sched.tue = (sched.astroDay === 2);
    sched.wed = (sched.astroDay === 3);
    sched.thu = (sched.astroDay === 4);
    sched.fri = (sched.astroDay === 5);
    sched.sat = (sched.astroDay === 6);
    sched.sun = (sched.astroDay === 7);

    sched.changed = true;
    sched.calculateTimes();
    sched.normalizeTimes();
  }

  if (statusInfo) statusInfo.textContent = `Відображено: ${schedulesFiltered.length} з ${schedules.length}`;
  renderVisibleRows();
  schedModal.style.display = 'none';
});

closeSchedModalBtn?.addEventListener('click', () => {
  schedModal.style.display = 'none';
});

function getTimeDiffMinutes(d1, d2) {
  if (!d1 || !d2) return 0;
  const m1 = d1.getHours() * 60 + d1.getMinutes();
  const m2 = d2.getHours() * 60 + d2.getMinutes();
  let diff = m2 - m1;
  if (diff < 0) diff += 1440;
  return diff;
}

function cloneSchedule(src) {
  const s = Object.assign(Object.create(Object.getPrototypeOf(src)), src);
  s.points = (src.points || []).map(p => ({
    ...p,
    arrival: p.arrival ? new Date(p.arrival.getTime()) : null,
    departure: p.departure ? new Date(p.departure.getTime()) : null
  }));
  s.changedFields = [];
  s.grf = "";
  s.isNew = true;
  s.changed = true;
  return s;
}

// --- 8. МОДАЛЬНОЕ ОКНО НЕИЗВЕСТНЫХ УЗЛОВ ---
function checkUnknownNodes() {
  const missingAX = Array.from(DictionaryService.unknownNodesProlog).filter(n => n && n.trim());
  const missingPravosl = Array.from(DictionaryService.unknownNodesNews).filter(n => n && n.trim());

  if (missingAX.length === 0 && missingPravosl.length === 0) return;

  unknownNodesTableBody.innerHTML = "";
  missingAX.forEach(node => unknownNodesTableBody.appendChild(createNodeRow(node, "")));
  missingPravosl.forEach(node => unknownNodesTableBody.appendChild(createNodeRow("", node)));

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

saveNodesBtn?.addEventListener('click', async () => {
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
    if (statusInfo) statusInfo.textContent = "Оновлення довідників...";
    appDictionaries = await DictionaryService.loadFromGoogleAPI(GOOGLE_APP_URL);

    DictionaryService.setCities(schedules, appDictionaries);
    DictionaryService.setAXCities(schedules, appDictionaries);
    schedules.forEach(s => s.setRoute());
    initVirtualTable(schedulesFiltered, false);

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

closeModalBtn?.addEventListener('click', () => {
  nodesModal.style.display = 'none';
});

// --- 9. ЭКСПОРТ ---
savePrologBtn?.addEventListener('click', async () => {
  if (schedulesFiltered.length === 0) return alert("Немає даних для збереження");
  showLoader("Формування файлу ProCTOlog...");
  await new Promise(r => setTimeout(r, 50));
  try {
    await ExcelWriter.exportToProlog(schedulesFiltered, "Графіки_ProCTOlog.xlsx", false);
  } finally {
    hideLoader();
  }
});

saveNewsletterBtn?.addEventListener('click', async () => {
  if (schedulesFiltered.length === 0) return alert("Немає даних для збереження");
  showLoader("Генерація файлу для розсилки...");
  await new Promise(r => setTimeout(r, 50));
  try {
    await ExcelWriter.exportToNewsLetter(schedulesFiltered, "Графіки_Розсилка.xlsx");
  } finally {
    hideLoader();
  }
});

saveTrafficLightBtn?.addEventListener('click', async () => {
  if (schedulesFiltered.length === 0) return alert("Немає даних для збереження");
  showLoader("Формування файлу Світлофор...");
  await new Promise(r => setTimeout(r, 50));
  try {
    await ExcelWriter.exportToProlog(schedulesFiltered, "Графіки_Світлофор.xlsx", true);
  } finally {
    hideLoader();
  }
});
// Открытие / закрытие выпадающего меню Звіти
reportsMenuBtn?.addEventListener('click', (e) => {
  e.stopPropagation();
  reportsDropdown?.classList.toggle('open');
});

document.addEventListener('click', () => {
  reportsDropdown?.classList.remove('open');
});

closeReportModalBtn?.addEventListener('click', () => {
  reportModal.style.display = 'none';
});

// Вспомогательная функция преобразования сегментов дедлайнов в названия узлов (ConvertDeadlines)
function getScheduleConvertedDeadlines(sched) {
  if (!sched.deadline || !sched.deadlineSegments || !sched.deadlineSegments.trim()) return [];
  const rawParts = sched.deadlineSegments.split(',').map(x => x.trim()).filter(Boolean);
  const result = [];

  for (const dd of rawParts) {
    const ab = dd.split('-');
    if (ab.length !== 2) continue;
    const a = parseInt(ab[0], 10) - 1;
    const b = parseInt(ab[1], 10) - 1;
    if (!isNaN(a) && !isNaN(b) && sched.points?.[a] && sched.points?.[b]) {
      result.push(`${sched.points[a].name}_${sched.points[b].name}`);
    }
  }
  return result;
}

// ==========================================
// ОТЧЁТ 1: ДЕЛДАЙНЕР (CheckDeadlineSegments)
// ==========================================
function buildDeadlinerReport(list) {
  const map = new Map();

  for (const sched of list) {
    const deadlines = getScheduleConvertedDeadlines(sched);
    for (const dir of deadlines) {
      let seg = map.get(dir);
      if (!seg) {
        seg = {
          direction: dir,
          mon: 0, tue: 0, wed: 0, thu: 0, fri: 0, sat: 0, sun: 0,
          diagnosis: "",
          schedules: []
        };
        map.set(dir, seg);
      }
      switch (sched.logisticDay) {
        case 1: seg.mon++; break;
        case 2: seg.tue++; break;
        case 3: seg.wed++; break;
        case 4: seg.thu++; break;
        case 5: seg.fri++; break;
        case 6: seg.sat++; break;
        case 7: seg.sun++; break;
      }
      seg.schedules.push(sched.grf || "");
    }
  }

  const rows = Array.from(map.values()).sort((a, b) => a.direction.localeCompare(b.direction));
  for (const seg of rows) {
    if (seg.mon !== seg.tue || seg.tue !== seg.wed || seg.wed !== seg.thu || seg.thu !== seg.fri) {
      seg.diagnosis = "Бідэ";
    }
  }
  return rows;
}

repDeadlinerBtn?.addEventListener('click', () => {
  reportsDropdown?.classList.remove('open');
  if (schedulesFiltered.length === 0) return alert("Немає завантажених графіків");

  const rows = buildDeadlinerReport(schedulesFiltered);
  currentReportType = "deadliner";
  currentReportRows = rows;
  currentReportHeaders = ["Напрямок", "ПН", "ВТ", "СР", "ЧТ", "ПТ", "СБ", "НД", "Діагноз"];

  renderReportModal(
    "Делдайнер",
    `Знайдено напрямків: ${rows.length}. Подвійний клік по напрямку відфільтрує графіки в головній таблиці.`,
    currentReportHeaders,
    rows.map(r => ({
      cells: [r.direction, r.mon, r.tue, r.wed, r.thu, r.fri, r.sat, r.sun, r.diagnosis],
      rowClass: r.diagnosis === "Бідэ" ? "diag-bad" : "",
      payload: r
    }))
  );
});

// ==========================================
// ОТЧЁТ 2: ПОГАНІ ДЕДЛАЙНИ (DeadlinesAreBad)
// ==========================================
function isDeadlineBad(sched) {
  const segStr = (sched.deadlineSegments || "").trim();
  if ((sched.deadline && segStr === "") || (!sched.deadline && segStr !== "")) {
    return true;
  }

  if (sched.deadline) {
    const ddls = segStr.split(',').map(x => x.trim()).filter(Boolean);
    if (ddls.length === 0) return true;

    for (const ddl of ddls) {
      const parts = ddl.split('-');
      if (parts.length !== 2) return true;
      const a = parseInt(parts[0], 10) - 1;
      const b = parseInt(parts[1], 10) - 1;

      if (isNaN(a) || isNaN(b) || !sched.points?.[a] || !sched.points?.[b]) {
        return true;
      }

      const ptA = sched.points[a];
      const ptB = sched.points[b];
      const cityA = (ptA.city || "").trim();
      const cityB = (ptB.city || "").trim();
      const catA = (ptA.nodeCat || "").trim();
      const catB = (ptB.nodeCat || "").trim();

      if ((cityA !== "" && cityA === cityB) ||
          (catA === "В" && catB === "П") ||
          (catA === "П" && catB === "В")) {
        return true;
      }
    }
  }
  return false;
}

repBadDeadlinesBtn?.addEventListener('click', () => {
  reportsDropdown?.classList.remove('open');
  if (schedulesFiltered.length === 0) return alert("Немає завантажених графіків");

  const badList = schedulesFiltered.filter(s => isDeadlineBad(s));
  currentReportType = "badDeadlines";
  currentReportRows = badList;
  currentReportHeaders = [
    "Дата початку", "Код", "Номер", "Назва", "Маршрут",
    "ПН", "ВТ", "СР", "ЧТ", "ПТ", "СБ", "НД",
    "Дедлайн", "Посегментні дедлайни", "Коментар"
  ];

  renderReportModal(
    "Погані дедлайни",
    `Знайдено графіків з помилками дедлайнів: ${badList.length}. Подвійний клік відкриває картку графіка.`,
    currentReportHeaders,
    badList.map(s => ({
      cells: [
        formatDateTime(s.startDate),
        s.grf || "",
        s.number || "",
        s.name || "",
        s.route || "",
        s.logMon ? "+" : "",
        s.logTue ? "+" : "",
        s.logWed ? "+" : "",
        s.logThu ? "+" : "",
        s.logFri ? "+" : "",
        s.logSat ? "+" : "",
        s.logSun ? "+" : "",
        s.deadline ? "Так" : "Ні",
        s.deadlineSegments || "",
        s.comment || ""
      ],
      rowClass: "",
      payload: s
    }))
  );
});

// ==========================================
// ОТЧЁТ 3: ПЕРЕВІРКА ТРАНЗИТНОГО ЧАСУ (CheckTimes)
// ==========================================
function getMatrixList() {
  if (!appDictionaries) return [];
  return appDictionaries.matrixes || appDictionaries.matrix || appDictionaries.Matrixes || [];
}

function findMatrixEntry(matrixList, nodeA, nodeB) {
  if (!nodeA || !nodeB) return null;
  const a = nodeA.trim().toLowerCase();
  const b = nodeB.trim().toLowerCase();
  return matrixList.find(m => {
    const mA = (m.nodeA ?? m.NodeA ?? m.from ?? "").trim().toLowerCase();
    const mB = (m.nodeB ?? m.NodeB ?? m.to ?? "").trim().toLowerCase();
    return mA === a && mB === b;
  }) || null;
}

function parseMatrixDurationMinutes(val) {
  if (val === null || val === undefined || val === "") return 0;
  if (val instanceof Date && !isNaN(val.getTime())) {
    return val.getHours() * 60 + val.getMinutes();
  }
  if (typeof val === "number") {
    // Если дробь суток Excel (< 1)
    if (val > 0 && val < 1) return Math.round(val * 24 * 60);
    return Math.round(val);
  }
  const str = String(val).trim();
  if (str.includes("T")) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) return d.getHours() * 60 + d.getMinutes();
  }
  if (str.includes(":")) {
    const parts = str.split(":");
    return (parseInt(parts[0], 10) || 0) * 60 + (parseInt(parts[1], 10) || 0);
  }
  return 0;
}

function formatMinutesHHMM(totalMin) {
  const h = Math.floor(totalMin / 60);
  const m = totalMin % 60;
  return `${String(h).padStart(2, '0')}:${String(m).padStart(2, '0')}`;
}

function buildCheckTimesReport(list, matrixList) {
  const res = [];
  for (const sched of list) {
    if (!sched.points || sched.points.length < 2) continue;
    for (let i = 1; i < sched.points.length; i++) {
      const prevPt = sched.points[i - 1];
      const curPt = sched.points[i];
      const tdMin = getTimeDiffMinutes(prevPt.departure, curPt.arrival);

      const mtr = findMatrixEntry(matrixList, prevPt.name, curPt.name);
      if (mtr) {
        const mtrMin = parseMatrixDurationMinutes(mtr.duration ?? mtr.Duration ?? mtr.time);
        if (tdMin !== mtrMin) {
          const dz = tdMin < mtrMin
            ? `Сегмент ${i} - графік менше сітки`
            : `Сегмент ${i} - графік більше сітки`;
          const deltaMin = Math.abs(tdMin - mtrMin);

          res.push({
            grf: sched.grf || "",
            route: sched.route || "",
            movingType: sched.movingType || "",
            dz,
            schedTime: formatMinutesHHMM(tdMin),
            matrixTime: formatMinutesHHMM(mtrMin),
            delta: formatMinutesHHMM(deltaMin),
            _sched: sched
          });
        }
      }
    }
  }
  return res;
}

repCheckTimesBtn?.addEventListener('click', () => {
  reportsDropdown?.classList.remove('open');
  if (schedulesFiltered.length === 0) return alert("Немає завантажених графіків");

  const matrixList = getMatrixList();
  if (matrixList.length === 0) {
    alert("Увага: довідник матриці відстаней порожній або ще не завантажився.");
  }

  const rows = buildCheckTimesReport(schedulesFiltered, matrixList);
  currentReportType = "checkTimes";
  currentReportRows = rows;
  currentReportHeaders = ["Графік", "Маршрут", "Тип переміщення", "Невідповідність", "Графік (час)", "Сітка", "Дельта"];

  renderReportModal(
    "Порівняння транзитного часу графіків із сіткою",
    `Виявлено розбіжностей: ${rows.length}. Подвійний клік по рядку відкриває картку графіка.`,
    currentReportHeaders,
    rows.map(r => ({
      cells: [r.grf, r.route, r.movingType, r.dz, r.schedTime, r.matrixTime, r.delta],
      rowClass: "",
      payload: r._sched
    }))
  );
});

// ==========================================
// ОТЧЁТ 4: СТРЬОМНИЙ ЧАС НА ВИВАНТАЖЕННЯ (CheckUnloadTimes)
// ==========================================
function buildCheckUnloadTimesReport(list) {
  const res = [];
  for (const sched of list) {
    if (!sched.points || sched.points.length < 3) continue;
    for (let i = 1; i <= sched.points.length - 2; i++) {
      const pt = sched.points[i];
      const diffMin = getTimeDiffMinutes(pt.arrival, pt.departure);
      if (diffMin > 360 || diffMin < 1) {
        res.push({
          grf: sched.grf || "",
          route: sched.route || "",
          node: pt.name || "",
          arrival: formatTime(pt.arrival),
          departure: formatTime(pt.departure),
          _sched: sched
        });
      }
    }
  }
  return res;
}

repCheckUnloadBtn?.addEventListener('click', () => {
  reportsDropdown?.classList.remove('open');
  if (schedulesFiltered.length === 0) return alert("Немає завантажених графіків");

  const rows = buildCheckUnloadTimesReport(schedulesFiltered);
  currentReportType = "checkUnload";
  currentReportRows = rows;
  currentReportHeaders = ["Графік", "Маршрут", "Вузол", "Приїзд", "Виїзд"];

  renderReportModal(
    "Стрьомний час на вивантаження",
    `Знайдено транзитних вузлів з часом > 6 год або < 1 хв: ${rows.length}. Подвійний клік відкриває картку графіка.`,
    currentReportHeaders,
    rows.map(r => ({
      cells: [r.grf, r.route, r.node, r.arrival, r.departure],
      rowClass: "",
      payload: r._sched
    }))
  );
});

// ==========================================
// ОТРИСОВКА ОКНА ОТЧЁТА И ОБРАБОТКА ДВОЙНОГО КЛИКА
// ==========================================
function renderReportModal(title, hint, headers, rowItems) {
  reportModalTitle.textContent = title;
  reportModalHint.textContent = hint;

  const thead = reportTable.querySelector('thead');
  const tbody = reportTable.querySelector('tbody');

  thead.innerHTML = `<tr>${headers.map(h => `<th>${h}</th>`).join('')}</tr>`;

  if (rowItems.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${headers.length}" style="text-align: center; padding: 18px;">Записів не знайдено</td></tr>`;
  } else {
    tbody.innerHTML = rowItems.map((item, idx) => `
      <tr data-rep-idx="${idx}" class="${item.rowClass}">
        ${item.cells.map(c => `<td>${c ?? ""}</td>`).join('')}
      </tr>
    `).join('');
  }

  // Сохраняем payload для двойного клика
  reportTable._rowItems = rowItems;
  reportModal.style.display = 'flex';
}

reportTable?.querySelector('tbody')?.addEventListener('dblclick', (e) => {
  const tr = e.target.closest('tr[data-rep-idx]');
  if (!tr) return;

  const idx = parseInt(tr.dataset.repIdx, 10);
  const item = reportTable._rowItems?.[idx];
  if (!item || !item.payload) return;

  if (currentReportType === "deadliner") {
    // Поведение DeadlinerForm_CellDoubleClick: сброс фильтров и поиск по направлению
    const dir = item.payload.direction || "";
    reportModal.style.display = 'none';
    clearFilters();
    fRoute.value = `*${dir.replace(/_/g, '*')}*`;
    filterSchedules();
  } else {
    // Открываем карточку графика для редактирования
    const sched = item.payload;
    const mainIdx = schedulesFiltered.indexOf(sched);
    openSchedModal(sched, mainIdx);
  }
});

// Выгрузка текущего отчёта в Excel
exportReportExcelBtn?.addEventListener('click', () => {
  const items = reportTable._rowItems || [];
  if (items.length === 0) return alert("Немає даних для експорту");

  const wsData = [currentReportHeaders, ...items.map(i => i.cells)];
  const ws = XLSX.utils.aoa_to_sheet(wsData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, "Звіт");
  XLSX.writeFile(wb, `${reportModalTitle.textContent || "Звіт"}.xlsx`);
});