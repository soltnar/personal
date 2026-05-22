const APP_VERSION = "1.7.1";
const DAY_CUTOFF_SECONDS = 4 * 3600;

const universalInput = document.getElementById("universalInput");
const attendanceInput = document.getElementById("attendanceInput");
const staffInput = document.getElementById("staffInput");
const statusEl = document.getElementById("status");
const dateSelect = document.getElementById("dateSelect");
const restaurantSelect = document.getElementById("restaurantSelect");
const calcBtn = document.getElementById("calcBtn");
const csvBtn = document.getElementById("csvBtn");
const xlsxBtn = document.getElementById("xlsxBtn");
const gsSendBtn = document.getElementById("gsSendBtn");
const gsWebhookUrlInput = document.getElementById("gsWebhookUrl");
const gsSpreadsheetIdInput = document.getElementById("gsSpreadsheetId");
const gsSheetNameInput = document.getElementById("gsSheetName");
const gsStatusEl = document.getElementById("gsStatus");
const summaryEl = document.getElementById("summary");
const tableBody = document.querySelector("#resultTable tbody");
const appVersionEl = document.getElementById("appVersion");

let baseRecords = [];
let mappedRecords = [];
let staffRestaurantMap = new Map();
let staffConflicts = 0;
let staffConflictKeys = new Set();
let mappingStats = { matched: 0, total: 0 };
let lastResultRows = [];

appVersionEl.textContent = APP_VERSION;

function excelDateToSerialDay(value) {
  const days = Number(value);
  if (!Number.isFinite(days)) return NaN;
  return Math.floor(days);
}

function serialDayToISO(serialDay) {
  const utcValue = (serialDay - 25569) * 86400;
  const date = new Date(utcValue * 1000);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

function parseExcelTimeToSeconds(value) {
  if (value === null || value === undefined || value === "") return NaN;
  const numeric = Number(value);
  if (Number.isFinite(numeric)) return Math.round(numeric * 24 * 3600);

  const text = String(value).trim();
  const m = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return NaN;
  return Number(m[1]) * 3600 + Number(m[2]) * 60 + Number(m[3] || 0);
}

function prettyDate(iso) {
  const [y, m, d] = iso.split("-");
  return `${d}.${m}.${y}`;
}

function normalize(text) {
  return String(text || "")
    .toLowerCase()
    .replace(/ё/g, "е")
    .replace(/\s+/g, " ")
    .trim();
}

function normalizeFio(text) {
  return normalize(text).replace(/[^a-zа-я0-9 ]/gi, "");
}

function classifyRole(roleText) {
  const role = normalize(roleText);
  if (/повар|шеф/.test(role)) return "Кухня";
  if (/официант|менеджер зала|мойщ|мойк/.test(role)) return "Зал";
  if (/логист|курьер|водител/.test(role)) return "Доставка";
  if (/барменедж|барбэк|барбек|бармен/.test(role)) return "Бар";
  return null;
}

function formatShift(value) {
  return Number.isInteger(value) ? String(value) : value.toFixed(1);
}

function escapeHtml(text) {
  return String(text)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function getSelectedValues(selectEl) {
  return Array.from(selectEl.selectedOptions).map((o) => o.value);
}

function getCheckedGroups() {
  return Array.from(document.querySelectorAll(".groupCheck:checked")).map((el) => el.value);
}

function fillMultiSelect(selectEl, values, selectedValues = []) {
  const selectedSet = new Set(selectedValues.length ? selectedValues : values);
  selectEl.innerHTML = "";
  values.forEach((v) => {
    const opt = document.createElement("option");
    opt.value = v;
    opt.textContent = v;
    opt.selected = selectedSet.has(v);
    selectEl.appendChild(opt);
  });
}

function findHeaderIndex(header, candidates) {
  for (const name of candidates) {
    const idx = header.indexOf(name);
    if (idx !== -1) return idx;
  }
  return -1;
}

function readWorkbookRows(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
}

function detectFileType(rows) {
  if (!rows.length) return "unknown";
  const header = rows[0].map((h) => String(h).trim());
  const has = (name) => header.includes(name);

  if (has("Источник") && has("Направление") && has("Дата") && has("Время") && (has("ФИО") || (has("Фамилия") && has("Имя")))) {
    return "attendance";
  }

  if (has("ФИО") && has("Название подразделения")) {
    return "staff";
  }

  return "unknown";
}

function parseStaffRows(rows) {
  if (!rows.length) throw new Error("Файл сотрудников пустой.");

  const header = rows[0].map((h) => String(h).trim());
  const fioIdx = header.indexOf("ФИО");
  const restaurantIdx = header.indexOf("Название подразделения");

  if (fioIdx === -1 || restaurantIdx === -1) {
    throw new Error("В файле сотрудников нужны колонки: ФИО и Название подразделения.");
  }

  const map = new Map();
  let conflicts = 0;
  const conflictKeys = new Set();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const fio = String(row[fioIdx] || "").trim();
    const restaurant = String(row[restaurantIdx] || "").trim();
    if (!fio || !restaurant) continue;

    const key = normalizeFio(fio);
    if (!key) continue;

    if (!map.has(key)) {
      map.set(key, restaurant);
    } else if (map.get(key) !== restaurant) {
      conflicts += 1;
      conflictKeys.add(key);
    }
  }

  return { map, conflicts, conflictKeys };
}

function parseAttendanceRows(rows) {
  if (!rows.length) return [];

  const header = rows[0].map((h) => String(h).trim());
  const idx = {
    date: findHeaderIndex(header, ["Дата"]),
    time: findHeaderIndex(header, ["Время"]),
    source: findHeaderIndex(header, ["Источник"]),
    direction: findHeaderIndex(header, ["Направление"]),
    surname: findHeaderIndex(header, ["Фамилия"]),
    name: findHeaderIndex(header, ["Имя"]),
    middle: findHeaderIndex(header, ["Отчество"]),
    fio: findHeaderIndex(header, ["ФИО"]),
    role: findHeaderIndex(header, ["Должность"]),
    address: findHeaderIndex(header, ["Адрес"])
  };

  const required = ["date", "time", "source", "direction", "role", "address"];
  const missing = required.filter((k) => idx[k] === -1);
  if (missing.length) {
    throw new Error(`Не найдены нужные колонки в файле проходной: ${missing.join(", ")}`);
  }
  if (idx.fio === -1 && (idx.surname === -1 || idx.name === -1)) {
    throw new Error("Не найдены колонки ФИО или Фамилия+Имя в файле проходной.");
  }

  const parsed = [];

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const source = String(row[idx.source] || "").trim();
    if (source !== "Проходная") continue;

    const dateSerialDay = excelDateToSerialDay(row[idx.date]);
    if (!Number.isFinite(dateSerialDay)) continue;

    const timeSec = parseExcelTimeToSeconds(row[idx.time]);
    if (!Number.isFinite(timeSec)) continue;

    const operationalSerialDay = timeSec < DAY_CUTOFF_SECONDS ? dateSerialDay - 1 : dateSerialDay;
    const dateIso = serialDayToISO(operationalSerialDay);
    if (!dateIso) continue;
    const absSec = dateSerialDay * 86400 + timeSec;

    const roleRaw = String(row[idx.role] || "").trim();
    const group = classifyRole(roleRaw);
    if (!group) continue;

    const person = idx.fio !== -1
      ? String(row[idx.fio] || "").trim()
      : [row[idx.surname], row[idx.name], row[idx.middle]].filter(Boolean).join(" ").trim();
    if (!person) continue;

    const direction = String(row[idx.direction] || "").trim();
    if (direction !== "Вход" && direction !== "Выход") continue;

    parsed.push({
      dateIso,
      absSec,
      person,
      personKey: normalizeFio(person),
      group,
      direction,
      restaurantFromGate: String(row[idx.address] || "").trim() || "Не указан"
    });
  }

  return parsed;
}

function applyStaffData(staffData) {
  staffRestaurantMap = staffData.map;
  staffConflicts = staffData.conflicts;
  staffConflictKeys = staffData.conflictKeys;
  if (baseRecords.length) rebuildMappedRecords();
  refreshStatus();
}

function applyAttendanceData(records) {
  baseRecords = records;
  rebuildMappedRecords();
  lastResultRows = [];
  tableBody.innerHTML = "";
  summaryEl.textContent = "Выберите фильтры и нажмите «Рассчитать».";
  csvBtn.disabled = true;
  xlsxBtn.disabled = true;
  gsSendBtn.disabled = true;
  refreshStatus();
}

async function processWorkbookFile(file) {
  const buf = await file.arrayBuffer();
  const rows = readWorkbookRows(buf);
  const type = detectFileType(rows);

  if (type === "attendance") {
    applyAttendanceData(parseAttendanceRows(rows));
    return { file: file.name, type: "проходная" };
  }

  if (type === "staff") {
    applyStaffData(parseStaffRows(rows));
    return { file: file.name, type: "сотрудники" };
  }

  return { file: file.name, type: "не распознан" };
}

function rebuildMappedRecords() {
  mappingStats = { matched: 0, total: baseRecords.length };

  mappedRecords = baseRecords.map((r) => {
    const mappedRestaurant = staffRestaurantMap.get(r.personKey);
    if (mappedRestaurant) mappingStats.matched += 1;

    return {
      ...r,
      restaurant: mappedRestaurant || "Не определен в списке сотрудников",
      hasConflict: staffConflictKeys.has(r.personKey)
    };
  });

  const prevDates = getSelectedValues(dateSelect);
  const prevRestaurants = getSelectedValues(restaurantSelect);

  const dates = [...new Set(mappedRecords.map((r) => r.dateIso))].sort();
  const restaurants = [...new Set(mappedRecords.map((r) => r.restaurant))].sort((a, b) => a.localeCompare(b, "ru"));

  fillMultiSelect(dateSelect, dates, prevDates);
  fillMultiSelect(restaurantSelect, restaurants, prevRestaurants);
}

function calcWorkedSeconds(events) {
  const sorted = [...events].sort((a, b) => a.absSec - b.absSec);
  let total = 0;
  let inWork = false;
  let startSec = 0;

  sorted.forEach((e) => {
    if (e.direction === "Вход") {
      inWork = true;
      startSec = e.absSec;
      return;
    }
    if (e.direction === "Выход" && inWork && e.absSec >= startSec) {
      total += e.absSec - startSec;
      inWork = false;
    }
  });

  if (total === 0 && sorted.length >= 2) {
    const fallback = sorted[sorted.length - 1].absSec - sorted[0].absSec;
    if (fallback > 0) total = fallback;
  }

  return total;
}

function workedSecondsToShift(workedSeconds) {
  if (workedSeconds <= 0) return 0;
  return workedSeconds > 7 * 3600 ? 1 : 0.5;
}

function calculate(records) {
  const selectedDates = getSelectedValues(dateSelect);
  const selectedRestaurants = getSelectedValues(restaurantSelect);
  const selectedGroups = new Set(getCheckedGroups());

  const filtered = records.filter(
    (r) => selectedDates.includes(r.dateIso) && selectedRestaurants.includes(r.restaurant) && selectedGroups.has(r.group)
  );

  const personDay = new Map();

  filtered.forEach((r) => {
    const key = `${r.dateIso}||${r.restaurant}||${r.group}||${r.person}`;
    if (!personDay.has(key)) {
      personDay.set(key, {
        dateIso: r.dateIso,
        restaurant: r.restaurant,
        group: r.group,
        person: r.person,
        hasConflict: false,
        events: []
      });
    }
    if (r.hasConflict) personDay.get(key).hasConflict = true;
    personDay.get(key).events.push({ direction: r.direction, absSec: r.absSec });
  });

  const restaurantDay = new Map();

  Array.from(personDay.values()).forEach((item) => {
    const shiftValue = workedSecondsToShift(calcWorkedSeconds(item.events));
    if (shiftValue === 0) return;

    const key = `${item.dateIso}||${item.restaurant}`;
    if (!restaurantDay.has(key)) {
      restaurantDay.set(key, {
        dateIso: item.dateIso,
        restaurant: item.restaurant,
        kitchen: 0,
        hall: 0,
        delivery: 0,
        bar: 0,
        total: 0,
        hasConflict: false,
        details: { kitchen: [], hall: [], delivery: [], bar: [] }
      });
    }

    const row = restaurantDay.get(key);
    if (item.group === "Кухня") row.details.kitchen.push({ person: item.person, shift: shiftValue, hasConflict: item.hasConflict });
    if (item.group === "Зал") row.details.hall.push({ person: item.person, shift: shiftValue, hasConflict: item.hasConflict });
    if (item.group === "Доставка") row.details.delivery.push({ person: item.person, shift: shiftValue, hasConflict: item.hasConflict });
    if (item.group === "Бар") row.details.bar.push({ person: item.person, shift: shiftValue, hasConflict: item.hasConflict });

    if (item.group === "Кухня") row.kitchen += shiftValue;
    if (item.group === "Зал") row.hall += shiftValue;
    if (item.group === "Доставка") row.delivery += shiftValue;
    if (item.group === "Бар") row.bar += shiftValue;
    if (item.hasConflict) row.hasConflict = true;
    row.total += shiftValue;
  });

  return Array.from(restaurantDay.values())
    .map((row) => {
      row.details.kitchen.sort((a, b) => a.person.localeCompare(b.person, "ru"));
      row.details.hall.sort((a, b) => a.person.localeCompare(b.person, "ru"));
      row.details.delivery.sort((a, b) => a.person.localeCompare(b.person, "ru"));
      row.details.bar.sort((a, b) => a.person.localeCompare(b.person, "ru"));
      return row;
    })
    .sort((a, b) => (a.dateIso !== b.dateIso ? a.dateIso.localeCompare(b.dateIso) : a.restaurant.localeCompare(b.restaurant, "ru")));
}

function renderPeopleList(items) {
  if (!items.length) return `<div class="emptyList">Нет сотрудников</div>`;
  return `<ul>${items.map((p) => `<li>${escapeHtml(p.person)} — ${formatShift(p.shift)}${p.hasConflict ? ' <span class="conflictBadge">конфликт ФИО</span>' : ''}</li>`).join("")}</ul>`;
}

function buildDetailsHtml(row) {
  return `
    <div class="detailsWrap">
      <div class="detailsCol"><h4>Кухня (${formatShift(row.kitchen)})</h4>${renderPeopleList(row.details.kitchen)}</div>
      <div class="detailsCol"><h4>Зал (${formatShift(row.hall)})</h4>${renderPeopleList(row.details.hall)}</div>
      <div class="detailsCol"><h4>Доставка (${formatShift(row.delivery)})</h4>${renderPeopleList(row.details.delivery)}</div>
      <div class="detailsCol"><h4>Бар (${formatShift(row.bar)})</h4>${renderPeopleList(row.details.bar)}</div>
    </div>
  `;
}

function renderTable(rows) {
  tableBody.innerHTML = "";

  if (!rows.length) {
    summaryEl.textContent = "По выбранным фильтрам данных нет.";
    csvBtn.disabled = true;
    xlsxBtn.disabled = true;
    gsSendBtn.disabled = true;
    return;
  }

  let totalKitchen = 0;
  let totalHall = 0;
  let totalDelivery = 0;
  let totalBar = 0;

  rows.forEach((r) => {
    totalKitchen += r.kitchen;
    totalHall += r.hall;
    totalDelivery += r.delivery;
    totalBar += r.bar;

    const tr = document.createElement("tr");
    const detailsTr = document.createElement("tr");
    detailsTr.className = "detailsRow";
    detailsTr.style.display = "none";

    const detailsCell = document.createElement("td");
    detailsCell.colSpan = 8;
    detailsCell.innerHTML = buildDetailsHtml(r);
    detailsTr.appendChild(detailsCell);

    const toggleId = `toggle-${r.dateIso}-${Math.random().toString(36).slice(2, 8)}`;
    tr.innerHTML = `
      <td><button class="detailBtn" id="${toggleId}" type="button">Показать</button></td>
      <td>${prettyDate(r.dateIso)}</td>
      <td>${escapeHtml(r.restaurant)}${r.hasConflict ? ' <span class="conflictBadge">есть конфликт</span>' : ''}</td>
      <td>${formatShift(r.kitchen)}</td>
      <td>${formatShift(r.hall)}</td>
      <td>${formatShift(r.delivery)}</td>
      <td>${formatShift(r.bar)}</td>
      <td>${formatShift(r.total)}</td>
    `;

    tableBody.appendChild(tr);
    tableBody.appendChild(detailsTr);

    tr.querySelector(`#${toggleId}`).addEventListener("click", (e) => {
      const open = detailsTr.style.display !== "none";
      detailsTr.style.display = open ? "none" : "";
      e.currentTarget.textContent = open ? "Показать" : "Скрыть";
    });
  });

  summaryEl.textContent = `Строк: ${rows.length}. Кухня: ${formatShift(totalKitchen)}, Зал: ${formatShift(totalHall)}, Доставка: ${formatShift(totalDelivery)}, Бар: ${formatShift(totalBar)}, Всего смен: ${formatShift(totalKitchen + totalHall + totalDelivery + totalBar)}.`;
  csvBtn.disabled = false;
  xlsxBtn.disabled = false;
  gsSendBtn.disabled = false;
}

function toCSV(rows) {
  const head = ["Дата", "Ресторан", "Кухня", "Зал", "Доставка", "Бар", "Итого"];
  const lines = [head.join(";")];
  rows.forEach((r) => {
    lines.push([
      prettyDate(r.dateIso),
      `"${String(r.restaurant).replaceAll('"', '""')}"`,
      formatShift(r.kitchen),
      formatShift(r.hall),
      formatShift(r.delivery),
      formatShift(r.bar),
      formatShift(r.total)
    ].join(";"));
  });
  return lines.join("\n");
}

function buildMatrix(rows, fieldName) {
  const restaurants = [...new Set(rows.map((r) => r.restaurant))].sort((a, b) => a.localeCompare(b, "ru"));
  const dates = [...new Set(rows.map((r) => r.dateIso))].sort();
  const map = new Map(rows.map((r) => [`${r.restaurant}||${r.dateIso}`, r[fieldName]]));

  const aoa = [["Ресторан", ...dates.map(prettyDate)]];
  restaurants.forEach((restaurant) => {
    const line = [restaurant];
    dates.forEach((dateIso) => line.push(map.get(`${restaurant}||${dateIso}`) || 0));
    aoa.push(line);
  });

  return aoa;
}

function exportExcelPivot(rows) {
  const groups = getCheckedGroups();
  const wb = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "total")), "Итого");
  if (groups.includes("Кухня")) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "kitchen")), "Кухня");
  if (groups.includes("Зал")) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "hall")), "Зал");
  if (groups.includes("Доставка")) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "delivery")), "Доставка");
  if (groups.includes("Бар")) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "bar")), "Бар");

  XLSX.writeFile(wb, `итог_персонал_${new Date().toISOString().slice(0, 10)}.xlsx`);
}

async function sendToGoogleSheets() {
  if (!lastResultRows.length) {
    gsStatusEl.textContent = "Сначала выполните расчет.";
    return;
  }

  const webhookUrl = gsWebhookUrlInput.value.trim();
  const spreadsheetId = gsSpreadsheetIdInput.value.trim();
  const sheetName = gsSheetNameInput.value.trim();

  if (!webhookUrl || !spreadsheetId) {
    gsStatusEl.textContent = "Заполните Webhook URL и Spreadsheet ID.";
    return;
  }

  const payload = {
    spreadsheetId,
    sheetName,
    generatedAt: new Date().toISOString(),
    groups: getCheckedGroups(),
    rows: lastResultRows,
    matrix: {
      total: buildMatrix(lastResultRows, "total"),
      kitchen: buildMatrix(lastResultRows, "kitchen"),
      hall: buildMatrix(lastResultRows, "hall"),
      delivery: buildMatrix(lastResultRows, "delivery"),
      bar: buildMatrix(lastResultRows, "bar")
    }
  };

  gsSendBtn.disabled = true;
  gsStatusEl.textContent = "Отправка данных в Google Sheets...";

  try {
    const resp = await fetch(webhookUrl, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload)
    });

    if (!resp.ok) {
      const txt = await resp.text();
      throw new Error(`${resp.status}: ${txt.slice(0, 200)}`);
    }

    gsStatusEl.textContent = "Данные успешно отправлены в Google Sheets.";
  } catch (err) {
    gsStatusEl.textContent = `Ошибка отправки: ${err.message}`;
  } finally {
    gsSendBtn.disabled = false;
  }
}

function refreshStatus() {
  if (!baseRecords.length) {
    statusEl.textContent = "Загрузите файл проходной.";
    return;
  }

  const staffLoaded = staffRestaurantMap.size > 0;
  const staffPart = staffLoaded
    ? ` Список сотрудников: сопоставлено ${mappingStats.matched} из ${mappingStats.total} записей.${staffConflicts ? ` Конфликтов ФИО: ${staffConflicts}.` : ""}`
    : " Список сотрудников не загружен, рестораны не будут определены.";

  statusEl.textContent = `Записей проходной: ${baseRecords.length}.${staffPart}`;
}

attendanceInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = readWorkbookRows(await file.arrayBuffer());
    applyAttendanceData(parseAttendanceRows(rows));
  } catch (err) {
    statusEl.textContent = `Ошибка файла проходной: ${err.message}`;
  }
});

staffInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = readWorkbookRows(await file.arrayBuffer());
    applyStaffData(parseStaffRows(rows));
    if (baseRecords.length) summaryEl.textContent = "Список сотрудников загружен. Пересчитайте данные.";
  } catch (err) {
    statusEl.textContent = `Ошибка файла сотрудников: ${err.message}`;
    staffRestaurantMap = new Map();
    staffConflicts = 0;
    staffConflictKeys = new Set();
    if (baseRecords.length) {
      rebuildMappedRecords();
      refreshStatus();
    }
  }
});

universalInput.addEventListener("change", async (e) => {
  const files = Array.from(e.target.files || []);
  if (!files.length) return;

  const results = [];
  for (const file of files) {
    try {
      const res = await processWorkbookFile(file);
      results.push(`${res.file}: ${res.type}`);
    } catch (err) {
      results.push(`${file.name}: ошибка (${err.message})`);
    }
  }

  summaryEl.textContent = `Общая загрузка: ${results.join("; ")}`;
});

calcBtn.addEventListener("click", () => {
  if (!mappedRecords.length) {
    summaryEl.textContent = "Сначала загрузите файл проходной.";
    return;
  }
  const checkedGroups = getCheckedGroups();
  if (!checkedGroups.length) {
    summaryEl.textContent = "Выберите хотя бы одну группу должностей.";
    tableBody.innerHTML = "";
    csvBtn.disabled = true;
    xlsxBtn.disabled = true;
    gsSendBtn.disabled = true;
    return;
  }
  lastResultRows = calculate(mappedRecords);
  renderTable(lastResultRows);
});

csvBtn.addEventListener("click", () => {
  if (!lastResultRows.length) return;
  const csv = toCSV(lastResultRows);
  const blob = new Blob(["\uFEFF" + csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const a = document.createElement("a");
  a.href = url;
  a.download = "count_staff_by_day.csv";
  document.body.appendChild(a);
  a.click();
  a.remove();
  URL.revokeObjectURL(url);
});

xlsxBtn.addEventListener("click", () => {
  if (!lastResultRows.length) return;
  exportExcelPivot(lastResultRows);
});

gsSendBtn.addEventListener("click", sendToGoogleSheets);
