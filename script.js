const APP_VERSION = "1.5.0";
const DAY_CUTOFF_SECONDS = 4 * 3600;

const attendanceInput = document.getElementById("attendanceInput");
const staffInput = document.getElementById("staffInput");
const statusEl = document.getElementById("status");
const dateSelect = document.getElementById("dateSelect");
const restaurantSelect = document.getElementById("restaurantSelect");
const calcBtn = document.getElementById("calcBtn");
const csvBtn = document.getElementById("csvBtn");
const xlsxBtn = document.getElementById("xlsxBtn");
const summaryEl = document.getElementById("summary");
const tableBody = document.querySelector("#resultTable tbody");
const appVersionEl = document.getElementById("appVersion");

let baseRecords = [];
let mappedRecords = [];
let staffRestaurantMap = new Map();
let staffConflicts = 0;
let mappingStats = { matched: 0, total: 0 };
let lastResultRows = [];

appVersionEl.textContent = APP_VERSION;

function excelDateToISO(value) {
  const days = Number(value);
  if (!Number.isFinite(days)) return "";
  const utcDays = Math.floor(days - 25569);
  const utcValue = utcDays * 86400;
  const date = new Date(utcValue * 1000);
  if (Number.isNaN(date.getTime())) return "";
  return date.toISOString().slice(0, 10);
}

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
  if (Number.isFinite(numeric)) {
    return Math.round(numeric * 24 * 3600);
  }

  const text = String(value).trim();
  const m = text.match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (!m) return NaN;
  const h = Number(m[1]);
  const min = Number(m[2]);
  const sec = Number(m[3] || 0);
  return h * 3600 + min * 60 + sec;
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
    .replaceAll("\"", "&quot;")
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

function parseStaffWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });
  if (!rows.length) throw new Error("Файл сотрудников пустой.");

  const header = rows[0].map((h) => String(h).trim());
  const fioIdx = header.indexOf("ФИО");
  const restaurantIdx = header.indexOf("Название подразделения");

  if (fioIdx === -1 || restaurantIdx === -1) {
    throw new Error("В файле сотрудников нужны колонки: ФИО и Название подразделения.");
  }

  const map = new Map();
  let conflicts = 0;

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
    }
  }

  return { map, conflicts };
}

function parseAttendanceWorkbook(arrayBuffer) {
  const wb = XLSX.read(arrayBuffer, { type: "array" });
  const sheet = wb.Sheets[wb.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, raw: true, defval: "" });

  if (!rows.length) return [];

  const header = rows[0].map((h) => String(h).trim());
  const idx = {
    date: header.indexOf("Дата"),
    time: header.indexOf("Время"),
    source: header.indexOf("Источник"),
    direction: header.indexOf("Направление"),
    surname: header.indexOf("Фамилия"),
    name: header.indexOf("Имя"),
    middle: header.indexOf("Отчество"),
    role: header.indexOf("Должность"),
    address: header.indexOf("Адрес")
  };

  const required = ["date", "time", "source", "direction", "surname", "name", "role", "address"];
  const missing = required.filter((k) => idx[k] === -1);
  if (missing.length) {
    throw new Error(`Не найдены нужные колонки в файле проходной: ${missing.join(", ")}`);
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

    const person = [row[idx.surname], row[idx.name], row[idx.middle]].filter(Boolean).join(" ").trim();
    if (!person) continue;

    const direction = String(row[idx.direction] || "").trim();
    if (direction !== "Вход" && direction !== "Выход") continue;

    parsed.push({
      dateIso,
      timeSec,
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

function rebuildMappedRecords() {
  mappingStats = { matched: 0, total: baseRecords.length };

  mappedRecords = baseRecords.map((r) => {
    const mappedRestaurant = staffRestaurantMap.get(r.personKey);
    if (mappedRestaurant) mappingStats.matched += 1;

    return {
      ...r,
      restaurant: mappedRestaurant || "Не определен в списке сотрудников"
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
    (r) =>
      selectedDates.includes(r.dateIso) &&
      selectedRestaurants.includes(r.restaurant) &&
      selectedGroups.has(r.group)
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
        events: []
      });
    }
    personDay.get(key).events.push({ direction: r.direction, absSec: r.absSec });
  });

  const restaurantDay = new Map();

  Array.from(personDay.values()).forEach((item) => {
    const workedSeconds = calcWorkedSeconds(item.events);
    const shiftValue = workedSecondsToShift(workedSeconds);
    if (shiftValue === 0) return;

    const key = `${item.dateIso}||${item.restaurant}`;
    if (!restaurantDay.has(key)) {
      restaurantDay.set(key, {
        dateIso: item.dateIso,
        restaurant: item.restaurant,
        kitchen: 0,
        hall: 0,
        delivery: 0,
        total: 0,
        details: {
          kitchen: [],
          hall: [],
          delivery: []
        }
      });
    }

    const row = restaurantDay.get(key);
    if (item.group === "Кухня") {
      row.kitchen += shiftValue;
      row.details.kitchen.push({ person: item.person, shift: shiftValue });
    }
    if (item.group === "Зал") {
      row.hall += shiftValue;
      row.details.hall.push({ person: item.person, shift: shiftValue });
    }
    if (item.group === "Доставка") {
      row.delivery += shiftValue;
      row.details.delivery.push({ person: item.person, shift: shiftValue });
    }
    row.total += shiftValue;
  });

  return Array.from(restaurantDay.values()).map((row) => {
    row.details.kitchen.sort((a, b) => a.person.localeCompare(b.person, "ru"));
    row.details.hall.sort((a, b) => a.person.localeCompare(b.person, "ru"));
    row.details.delivery.sort((a, b) => a.person.localeCompare(b.person, "ru"));
    return row;
  }).sort((a, b) => {
    if (a.dateIso !== b.dateIso) return a.dateIso.localeCompare(b.dateIso);
    return a.restaurant.localeCompare(b.restaurant, "ru");
  });
}

function renderPeopleList(items) {
  if (!items.length) return `<div class="emptyList">Нет сотрудников</div>`;
  return `<ul>${items
    .map((p) => `<li>${escapeHtml(p.person)} — ${formatShift(p.shift)}</li>`)
    .join("")}</ul>`;
}

function buildDetailsHtml(row) {
  return `
    <div class="detailsWrap">
      <div class="detailsCol">
        <h4>Кухня (${formatShift(row.kitchen)})</h4>
        ${renderPeopleList(row.details.kitchen)}
      </div>
      <div class="detailsCol">
        <h4>Зал (${formatShift(row.hall)})</h4>
        ${renderPeopleList(row.details.hall)}
      </div>
      <div class="detailsCol">
        <h4>Доставка (${formatShift(row.delivery)})</h4>
        ${renderPeopleList(row.details.delivery)}
      </div>
    </div>
  `;
}

function renderTable(rows) {
  tableBody.innerHTML = "";

  if (!rows.length) {
    summaryEl.textContent = "По выбранным фильтрам данных нет.";
    csvBtn.disabled = true;
    xlsxBtn.disabled = true;
    return;
  }

  let totalKitchen = 0;
  let totalHall = 0;
  let totalDelivery = 0;

  rows.forEach((r) => {
    totalKitchen += r.kitchen;
    totalHall += r.hall;
    totalDelivery += r.delivery;

    const tr = document.createElement("tr");
    tr.className = "mainRow";
    const detailsTr = document.createElement("tr");
    detailsTr.className = "detailsRow";
    detailsTr.style.display = "none";

    const detailsCell = document.createElement("td");
    detailsCell.colSpan = 7;
    detailsCell.innerHTML = buildDetailsHtml(r);
    detailsTr.appendChild(detailsCell);

    const toggleId = `toggle-${r.dateIso}-${Math.random().toString(36).slice(2, 8)}`;
    tr.innerHTML = `
      <td><button class="detailBtn" id="${toggleId}" type="button">Показать</button></td>
      <td>${prettyDate(r.dateIso)}</td>
      <td>${escapeHtml(r.restaurant)}</td>
      <td>${formatShift(r.kitchen)}</td>
      <td>${formatShift(r.hall)}</td>
      <td>${formatShift(r.delivery)}</td>
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

  summaryEl.textContent = `Строк: ${rows.length}. Кухня: ${formatShift(totalKitchen)}, Зал: ${formatShift(totalHall)}, Доставка: ${formatShift(totalDelivery)}, Всего смен: ${formatShift(totalKitchen + totalHall + totalDelivery)}.`;
  csvBtn.disabled = false;
  xlsxBtn.disabled = false;
}

function toCSV(rows) {
  const head = ["Дата", "Ресторан", "Кухня", "Зал", "Доставка", "Итого"];
  const lines = [head.join(";")];
  rows.forEach((r) => {
    lines.push([
      prettyDate(r.dateIso),
      `"${String(r.restaurant).replaceAll('"', '""')}"`,
      formatShift(r.kitchen),
      formatShift(r.hall),
      formatShift(r.delivery),
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
    dates.forEach((dateIso) => {
      line.push(map.get(`${restaurant}||${dateIso}`) || 0);
    });
    aoa.push(line);
  });

  return aoa;
}

function exportExcelPivot(rows) {
  const groups = getCheckedGroups();
  const wb = XLSX.utils.book_new();

  const totalSheet = XLSX.utils.aoa_to_sheet(buildMatrix(rows, "total"));
  XLSX.utils.book_append_sheet(wb, totalSheet, "Итого");

  if (groups.includes("Кухня")) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "kitchen")), "Кухня");
  }
  if (groups.includes("Зал")) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "hall")), "Зал");
  }
  if (groups.includes("Доставка")) {
    XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(buildMatrix(rows, "delivery")), "Доставка");
  }

  const dateLabel = new Date().toISOString().slice(0, 10);
  XLSX.writeFile(wb, `итог_персонал_${dateLabel}.xlsx`);
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
    const buf = await file.arrayBuffer();
    baseRecords = parseAttendanceWorkbook(buf);
    rebuildMappedRecords();

    lastResultRows = [];
    tableBody.innerHTML = "";
    summaryEl.textContent = "Выберите фильтры и нажмите «Рассчитать».";
    csvBtn.disabled = true;
    xlsxBtn.disabled = true;

    refreshStatus();
  } catch (err) {
    statusEl.textContent = `Ошибка файла проходной: ${err.message}`;
    baseRecords = [];
    mappedRecords = [];
    lastResultRows = [];
    tableBody.innerHTML = "";
    summaryEl.textContent = "Нет данных для отображения.";
    csvBtn.disabled = true;
    xlsxBtn.disabled = true;
  }
});

staffInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;

  try {
    const buf = await file.arrayBuffer();
    const staff = parseStaffWorkbook(buf);
    staffRestaurantMap = staff.map;
    staffConflicts = staff.conflicts;

    if (baseRecords.length) {
      rebuildMappedRecords();
      summaryEl.textContent = "Список сотрудников загружен. Пересчитайте данные.";
    }

    refreshStatus();
  } catch (err) {
    statusEl.textContent = `Ошибка файла сотрудников: ${err.message}`;
    staffRestaurantMap = new Map();
    staffConflicts = 0;
    if (baseRecords.length) {
      rebuildMappedRecords();
      refreshStatus();
    }
  }
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
