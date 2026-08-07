const APP_VERSION = "2.2.3";
const DAY_CUTOFF_SECONDS = 4 * 3600;

const universalInput = document.getElementById("universalInput");
const attendanceInput = document.getElementById("attendanceInput");
const staffInput = document.getElementById("staffInput");
const revenueInput = document.getElementById("revenueInput");
const statusEl = document.getElementById("status");
const dateSelect = document.getElementById("dateSelect");
const restaurantSelect = document.getElementById("restaurantSelect");
const warehouseTypeControl = document.getElementById("warehouseTypeControl");
const calcBtn = document.getElementById("calcBtn");
const csvBtn = document.getElementById("csvBtn");
const xlsxBtn = document.getElementById("xlsxBtn");
const summaryEl = document.getElementById("summary");
const tableBody = document.querySelector("#resultTable tbody");
const appVersionEl = document.getElementById("appVersion");
const revenueDbStatusEl = document.getElementById("revenueDbStatus");
const revenueDbLoginBtn = document.getElementById("revenueDbLogin");
const revenueDbLogoutBtn = document.getElementById("revenueDbLogout");
const loadRevenueDbBtn = document.getElementById("loadRevenueDbBtn");
const sabyFromInput = document.getElementById("sabyFrom");
const sabyToInput = document.getElementById("sabyTo");
const loadSabyBtn = document.getElementById("loadSabyBtn");
const sabyApiStatusEl = document.getElementById("sabyApiStatus");

const SUPABASE_URL = "https://wqxbnwcdkobgeyhdmqup.supabase.co";
const SUPABASE_PUBLISHABLE_KEY = "sb_publishable_WzfB8mJAOBXpeNWa34hBEQ_11QhCyqa";
const REVENUE_API_URL = `${SUPABASE_URL}/functions/v1/revenue-api`;
const PERSONNEL_API_URL = `${SUPABASE_URL}/functions/v1/personnel-api`;

let baseRecords = [];
let mappedRecords = [];
let staffRestaurantMap = new Map();
let staffConflicts = 0;
let staffConflictKeys = new Set();
let mappingStats = { matched: 0, total: 0 };
let lastResultRows = [];
let revenueRows = [];
let revenueStats = { rows: 0, matched: 0 };
let supabaseClient = null;
let revenueDbSession = null;
let revenueDbLoading = false;
let sabyApiLoading = false;

appVersionEl.textContent = APP_VERSION;

function isoLocalDate(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");
  return `${year}-${month}-${day}`;
}

function setDefaultSabyPeriod() {
  if (!sabyFromInput || !sabyToInput) return;
  const today = new Date();
  const weekAgo = new Date(today);
  weekAgo.setDate(today.getDate() - 7);
  sabyFromInput.value ||= isoLocalDate(weekAgo);
  sabyToInput.value ||= isoLocalDate(today);
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

function hasRibsMarker(value) {
  const text = String(value || "").toLowerCase().replace(/ё/g, "е");
  return text.includes("ribs") || text.includes("рибс") || /[rр][iи][bб][sс]/i.test(text);
}

function canonicalRestaurantName(value) {
  const original = String(value || "").trim();
  const key = revenueNameKey(original);
  if (!key) return "";
  if (key.includes("белинского") && key.includes("61") && hasRibsMarker(original)) {
    return "Белинского, 61 Рибс";
  }
  if (key.includes("белинского") && key.includes("61") && key.includes("достав")) {
    return "Белинского, 61 Самурай";
  }
  if ((key.includes("б покровская") || key.includes("большая покровская") || key.includes("бп")) && key.includes("63")) {
    return "Б. Покровская, 63 Самурай";
  }
  if ((key.includes("б покровская") || key.includes("большая покровская") || key.includes("бп")) && key.includes("59")) {
    return "Б. Покровская, 59 Самурай";
  }
  if (key.includes("винедо") || key.includes("vinedo")) {
    return "Октябрьская, 1 Винедо";
  }
  if (key.includes("детский центр жюль верн") || key.includes("ударник") || (key.includes("ленина") && key.includes("64"))) {
    return "Ленина, 64 Ударник";
  }
  return original;
}

function isNonRestaurantDepartment(value) {
  const key = revenueNameKey(value);
  return !key
    || key === "не определен в списке сотрудников"
    || key === "технический персонал"
    || key.includes("сотрудники сторонних организаций")
    || key.includes("сдача отчетности");
}

function classifyRole(roleText) {
  const role = normalize(roleText);
  if (/повар|шеф/.test(role)) return "Кухня";
  if (/официант|менеджер зала|мойщ|мойк|администратор|кассир|оператор/.test(role)) return "Зал";
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

function setRevenueDbStatus(message, kind = "") {
  if (!revenueDbStatusEl) return;
  revenueDbStatusEl.textContent = message;
  revenueDbStatusEl.className = `status${kind ? ` is-${kind}` : ""}`;
}

function setSabyApiStatus(message, kind = "") {
  if (!sabyApiStatusEl) return;
  sabyApiStatusEl.textContent = message;
  sabyApiStatusEl.className = `status${kind ? ` is-${kind}` : ""}`;
}

function updateRevenueDbButtons() {
  if (!revenueDbLoginBtn || !revenueDbLogoutBtn || !loadRevenueDbBtn) return;
  const signedIn = Boolean(revenueDbSession?.access_token);
  revenueDbLoginBtn.hidden = signedIn;
  revenueDbLogoutBtn.hidden = !signedIn;
  loadRevenueDbBtn.disabled = revenueDbLoading || !signedIn || !dateSelect.options.length;
  if (loadSabyBtn) {
    loadSabyBtn.disabled = sabyApiLoading || !signedIn || !sabyFromInput?.value || !sabyToInput?.value;
  }
}

async function initRevenueDatabase() {
  if (!window.supabase?.createClient) {
    setRevenueDbStatus("Библиотека Supabase не загрузилась. Можно использовать Excel-файл выручек как резерв.", "error");
    setSabyApiStatus("Автоматическая загрузка недоступна. Откройте резервную загрузку Excel.", "error");
    updateRevenueDbButtons();
    return;
  }

  supabaseClient = window.supabase.createClient(SUPABASE_URL, SUPABASE_PUBLISHABLE_KEY);
  const { data } = await supabaseClient.auth.getSession();
  revenueDbSession = data.session;
  updateRevenueDbButtons();
  setRevenueDbStatus(
    revenueDbSession
      ? "Доступ к базе выручки есть. После загрузки проходной выручка подтянется автоматически."
      : "Для автозагрузки выручки войдите через Google. Excel-файл выручек остается резервным вариантом.",
    revenueDbSession ? "success" : ""
  );
  setSabyApiStatus(
    revenueDbSession
      ? "Доступ подтверждён. Выберите период и загрузите данные."
      : "Для автоматической загрузки войдите через Google.",
    revenueDbSession ? "success" : ""
  );

  supabaseClient.auth.onAuthStateChange((_event, session) => {
    revenueDbSession = session;
    updateRevenueDbButtons();
    if (session) {
      setRevenueDbStatus("Вход выполнен. Можно подтянуть выручку из базы.", "success");
      setSabyApiStatus("Вход выполнен. Выберите период и нажмите «Загрузить данные из Saby».", "success");
    } else {
      setRevenueDbStatus("Для автозагрузки выручки войдите через Google.", "");
      setSabyApiStatus("Для автоматической загрузки войдите через Google.", "");
    }
  });
}

async function signInRevenueDatabase() {
  if (!supabaseClient) return;
  revenueDbLoginBtn.disabled = true;
  setRevenueDbStatus("Открываем вход через Google…", "loading");
  const { error } = await supabaseClient.auth.signInWithOAuth({
    provider: "google",
    options: { redirectTo: "https://soltnar.github.io/personal/" }
  });
  if (error) {
    setRevenueDbStatus(`Не удалось открыть вход: ${error.message}`, "error");
    revenueDbLoginBtn.disabled = false;
  }
}

async function signOutRevenueDatabase() {
  if (!supabaseClient) return;
  await supabaseClient.auth.signOut();
  revenueDbSession = null;
  updateRevenueDbButtons();
  setRevenueDbStatus("Вы вышли из базы выручки. Excel-файл выручек остается доступен вручную.", "");
  setSabyApiStatus("Вы вышли. Автоматическая загрузка приостановлена.", "");
}

function getRevenueDbDateRange() {
  const selected = getSelectedValues(dateSelect);
  const dates = selected.length ? selected : Array.from(dateSelect.options).map((option) => option.value);
  const validDates = dates.filter((date) => /^\d{4}-\d{2}-\d{2}$/.test(date)).sort();
  if (!validDates.length) return null;
  return { from: validDates[0], to: validDates[validDates.length - 1] };
}

function parseSabyDateTime(value) {
  const match = String(value || "").match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?/);
  if (!match) return null;
  const [, year, month, day, hour, minute, second = "0"] = match;
  const dayNumber = Math.floor(Date.UTC(Number(year), Number(month) - 1, Number(day)) / 86400000);
  const timeSec = Number(hour) * 3600 + Number(minute) * 60 + Number(second);
  const operationalDay = timeSec < DAY_CUTOFF_SECONDS ? dayNumber - 1 : dayNumber;
  const operationalDate = new Date(operationalDay * 86400000).toISOString().slice(0, 10);
  return { dateIso: operationalDate, absSec: dayNumber * 86400 + timeSec };
}

function parseSabyEmployeeDate(value) {
  const match = String(value || "").match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  return match ? `${match[3]}-${match[2]}-${match[1]}` : "";
}

function prepareSabyData(rows, from, to, employees = []) {
  const departmentSets = new Map();
  const departmentLabels = new Map();
  const officialByPerson = new Map();
  const attendance = [];

  employees.forEach((employee) => {
    const personKey = normalizeFio(employee.person);
    const hired = parseSabyEmployeeDate(employee.hired);
    const fired = parseSabyEmployeeDate(employee.fired);
    if (!personKey || (hired && hired > to) || (fired && fired < from)) return;
    if (!officialByPerson.has(personKey)) officialByPerson.set(personKey, []);
    officialByPerson.get(personKey).push(employee);

    const department = canonicalRestaurantName(employee.department);
    if (department && !isNonRestaurantDepartment(department)) {
      if (!departmentSets.has(personKey)) departmentSets.set(personKey, new Set());
      departmentSets.get(personKey).add(department);
      departmentLabels.set(personKey, department);
    }
  });

  rows.forEach((row) => {
    const person = String(row.person || "").trim();
    const personKey = normalizeFio(person);
    const timing = parseSabyDateTime(row.dateTime);
    const official = officialByPerson.get(personKey) || [];
    const officialRole = official.map((employee) => employee.position).find(classifyRole);
    const group = classifyRole(officialRole || row.position);
    const department = canonicalRestaurantName(official[0]?.department || row.department);
    if (!personKey || !timing || !group || timing.dateIso < from || timing.dateIso > to) return;

    // Internal report fields remain a fallback for employees that the official
    // directory did not return for the selected employment period.
    if (!official.length && department && !isNonRestaurantDepartment(department)) {
      if (!departmentSets.has(personKey)) departmentSets.set(personKey, new Set());
      departmentSets.get(personKey).add(department);
      departmentLabels.set(personKey, department);
    }

    attendance.push({
      dateIso: timing.dateIso,
      absSec: timing.absSec,
      person,
      personKey,
      group,
      direction: Number(row.actionType) === 1 ? "Вход" : "Выход",
      restaurantFromGate: String(row.address || "").trim() || "Не указан"
    });
  });

  const map = new Map();
  const conflictKeys = new Set();
  let conflicts = 0;
  departmentSets.forEach((departments, personKey) => {
    const values = [...departments];
    if (values.length === 1) map.set(personKey, values[0]);
    if (values.length > 1) {
      conflicts += 1;
      conflictKeys.add(personKey);
      map.set(personKey, departmentLabels.get(personKey));
    }
  });

  return { attendance, staff: { map, conflicts, conflictKeys } };
}

const SABY_CACHE_POLL_MS = 4000;
const SABY_CACHE_WAIT_MS = 15 * 60 * 1000;
const SABY_API_TIMEOUT_MS = 20000;
const SABY_EMPLOYEES_TIMEOUT_MS = 90000;

function addDaysIso(dateIso, days) {
  const [year, month, day] = dateIso.split("-").map(Number);
  const date = new Date(Date.UTC(year, month - 1, day));
  date.setUTCDate(date.getUTCDate() + days);
  return date.toISOString().slice(0, 10);
}

async function fetchSabyChunk(from, to, accessToken) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), SABY_API_TIMEOUT_MS);
  try {
    const response = await fetch(`${PERSONNEL_API_URL}?from=${encodeURIComponent(from)}&to=${encodeURIComponent(to)}`, {
      method: "GET",
      cache: "no-store",
      signal: controller.signal,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        apikey: SUPABASE_PUBLISHABLE_KEY,
        "Content-Type": "application/json"
      }
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) {
      const message = response.status === 546
        ? "Saby не успел ответить (546)"
        : (payload.error || `HTTP ${response.status}`);
      throw new Error(message);
    }
    if (!Array.isArray(payload.rows)) throw new Error("Сервер вернул неизвестный формат данных");
    return payload;
  } catch (error) {
    if (error?.name === "AbortError") throw new Error(`сервер кэша не ответил за ${Math.round(SABY_API_TIMEOUT_MS / 1000)} секунд`);
    throw error;
  } finally {
    clearTimeout(timeoutId);
  }
}

async function fetchSabyEmployees(accessToken) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), SABY_EMPLOYEES_TIMEOUT_MS);
  try {
    const response = await fetch(`${PERSONNEL_API_URL}?resource=employees`, {
      method: "GET",
      cache: "no-store",
      signal: controller.signal,
      headers: {
        Authorization: `Bearer ${accessToken}`,
        apikey: SUPABASE_PUBLISHABLE_KEY,
        "Content-Type": "application/json"
      }
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || `HTTP ${response.status}`);
    return Array.isArray(payload.employees) ? payload.employees : [];
  } finally {
    clearTimeout(timeoutId);
  }
}

async function loadSabyFromApi() {
  if (sabyApiLoading || !supabaseClient) return;
  const from = sabyFromInput?.value || "";
  const to = sabyToInput?.value || "";
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) {
    setSabyApiStatus("Проверьте даты периода.", "error");
    return;
  }

  const { data } = await supabaseClient.auth.getSession();
  revenueDbSession = data.session;
  if (!revenueDbSession?.access_token) {
    setSabyApiStatus("Сначала войдите через Google.", "error");
    updateRevenueDbButtons();
    return;
  }

  sabyApiLoading = true;
  updateRevenueDbButtons();
  loadSabyBtn.classList.add("is-busy");
  loadSabyBtn.setAttribute("aria-busy", "true");
  let allRows = [];
  let employees = [];
  const startedAt = Date.now();

  try {
    // The official employee directory is independent from attendance. Load it
    // in parallel so a slow directory page cannot delay or break cache polling.
    const employeesPromise = fetchSabyEmployees(revenueDbSession.access_token).catch(() => []);
    while (Date.now() - startedAt < SABY_CACHE_WAIT_MS) {
      const payload = await fetchSabyChunk(from, to, revenueDbSession.access_token);
      allRows = payload.rows;
      if (Array.isArray(payload.employees) && payload.employees.length) employees = payload.employees;
      const ready = payload.readyDates?.length || 0;
      const total = ready + (payload.pendingDates?.length || 0);
      if (payload.complete) break;
      const currentDate = payload.diagnostics?.syncingDate || payload.pendingDates?.[0] || "следующая дата";
      setSabyApiStatus(`Кэшируем проходную: готово ${ready} из ${total} дней. Сейчас загружается ${currentDate}…`, "loading");
      await new Promise((resolve) => setTimeout(resolve, SABY_CACHE_POLL_MS));
    }

    const finalPayload = await fetchSabyChunk(from, to, revenueDbSession.access_token);
    allRows = finalPayload.rows;
    if (!finalPayload.complete) throw new Error(`не успели заполнить ${finalPayload.pendingDates?.length || 0} дней; готовые даты сохранены`);
    employees = await Promise.race([
      employeesPromise,
      new Promise((resolve) => setTimeout(() => resolve([]), 8000))
    ]);

    const prepared = prepareSabyData(allRows, from, to, employees);
    if (!prepared.attendance.length) throw new Error("За выбранный период не найдено событий проходной по учитываемым должностям");
    applyStaffData(prepared.staff);
    applyAttendanceData(prepared.attendance, { selectAllDates: true, skipAutoRevenue: true });
    lastResultRows = calculate(mappedRecords);
    renderTable(lastResultRows);
    setSabyApiStatus(
      `Готово за ${Math.max(1, Math.round((Date.now() - startedAt) / 1000))} сек.: ${allRows.length} событий проходной, ${employees.length} сотрудников. В расчёт вошло ${prepared.attendance.length}.`,
      "success"
    );
    await loadRevenueFromDatabase({ silent: true, range: { from, to } });
  } catch (error) {
    setSabyApiStatus(`Не удалось загрузить данные Saby: ${error?.message || "неизвестная ошибка"}. Доступна резервная загрузка Excel.`, "error");
  } finally {
    sabyApiLoading = false;
    loadSabyBtn.classList.remove("is-busy");
    loadSabyBtn.removeAttribute("aria-busy");
    updateRevenueDbButtons();
  }
}

async function loadRevenueFromDatabase(options = {}) {
  if (revenueDbLoading) return;
  if (!supabaseClient) {
    setRevenueDbStatus("База выручки пока недоступна. Используйте Excel-файл выручек.", "error");
    return;
  }
  const range = options.range || getRevenueDbDateRange();
  if (!range) {
    setRevenueDbStatus("Сначала загрузите проходную, чтобы появились даты для запроса выручки.", "");
    updateRevenueDbButtons();
    return;
  }

  const { data } = await supabaseClient.auth.getSession();
  revenueDbSession = data.session;
  if (!revenueDbSession?.access_token) {
    setRevenueDbStatus("Для автозагрузки выручки войдите через Google.", "");
    updateRevenueDbButtons();
    return;
  }

  revenueDbLoading = true;
  updateRevenueDbButtons();
  if (!options.silent) setRevenueDbStatus(`Читаем выручку из базы за ${range.from} - ${range.to}…`, "loading");

  try {
    const response = await fetch(`${REVENUE_API_URL}?from=${encodeURIComponent(range.from)}&to=${encodeURIComponent(range.to)}`, {
      method: "GET",
      cache: "no-store",
      headers: {
        Authorization: `Bearer ${revenueDbSession.access_token}`,
        apikey: SUPABASE_PUBLISHABLE_KEY,
        "Content-Type": "application/json"
      }
    });
    const payload = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(payload.error || `Ошибка ${response.status}`);
    if (!Array.isArray(payload.rows)) throw new Error("База вернула неизвестный формат выручки");

    const rows = payload.rows
      .map((row) => ({
        dateIso: String(row.date || ""),
        warehouse: String(row.restaurant || "").trim(),
        revenue: Number(row.revenue),
        source: "База выручки"
      }))
      .filter((row) => /^\d{4}-\d{2}-\d{2}$/.test(row.dateIso) && row.warehouse && Number.isFinite(row.revenue) && !isBlockedRevenueName(row.warehouse));

    applyRevenueData(aggregateRevenueRows(rows));
    if (lastResultRows.length) {
      lastResultRows = calculate(mappedRecords);
      renderTable(lastResultRows);
    }

    const loadedAt = payload.generatedAt ? new Date(payload.generatedAt).toLocaleString("ru-RU") : new Date().toLocaleString("ru-RU");
    setRevenueDbStatus(`Выручка из базы загружена: ${rows.length} строк за ${range.from} - ${range.to}. ${loadedAt}.`, "success");
    if (!options.silent) summaryEl.textContent = "Выручка из базы загружена. Нажмите «Рассчитать» или используйте текущий пересчет.";
  } catch (error) {
    setRevenueDbStatus(`Не удалось загрузить выручку из базы: ${error.message || "неизвестная ошибка"}. Можно загрузить Excel-файл выручек вручную.`, "error");
  } finally {
    revenueDbLoading = false;
    updateRevenueDbButtons();
    refreshStatus();
  }
}

function maybeAutoLoadRevenueFromDatabase() {
  updateRevenueDbButtons();
  if (revenueDbSession?.access_token && dateSelect.options.length) {
    loadRevenueFromDatabase({ silent: true });
  }
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


const DEFAULT_REVENUE_EXCLUSIONS = [
  "Онлайн оплата, СБП",
  "Основной склад",
  "БУРГЕР БИК ООО Чайка",
  "Бургер Бик",
  "Фабрика разделка",
  "Шале №15",
  "Совнаркомовская 13",
  "НТО ООО Приспех пр-кт Гагарина, д. 35",
  "Юность ул. Зеленский Съезд, д. 8/10",
  "ИП Амельченко Евгений Андреевич",
  "Фудтрак Амельченко пл. Маркина, д. 12А",
  "Фабрика кондитерка",
  "ул. Большая Покровская, д. 13",
  "Швейцария БИК \"ПРИСПЕХ\"",
  "Фабрика пекарня",
  "Рождественская",
  "ВЕНУСТО"
].map(revenueNameKey);

function readWorkbook(arrayBuffer) {
  return XLSX.read(arrayBuffer, { type: "array" });
}

function cellToText(cell) {
  if (!cell) return "";
  if (cell.w != null) return String(cell.w).trim();
  if (cell.v != null) return String(cell.v).trim();
  return "";
}

function repairWorksheetRef(ws) {
  if (!ws) return;
  const cells = Object.keys(ws).filter((k) => !k.startsWith("!"));
  if (!cells.length) return;
  let minR = Infinity;
  let minC = Infinity;
  let maxR = -1;
  let maxC = -1;
  cells.forEach((addr) => {
    const decoded = XLSX.utils.decode_cell(addr);
    minR = Math.min(minR, decoded.r);
    minC = Math.min(minC, decoded.c);
    maxR = Math.max(maxR, decoded.r);
    maxC = Math.max(maxC, decoded.c);
  });
  ws["!ref"] = XLSX.utils.encode_range({ s: { r: minR, c: minC }, e: { r: maxR, c: maxC } });
}

function parseRevenueWorkbook(arrayBuffer, fileName) {
  const wb = readWorkbook(arrayBuffer);
  const parsed = [];
  wb.SheetNames.forEach((sheetName) => {
    const ws = wb.Sheets[sheetName];
    repairWorksheetRef(ws);
    parsed.push(...parseRevenuePeriodicByDays(ws, fileName, sheetName));
  });
  return aggregateRevenueRows(parsed);
}

function parseRevenuePeriodicByDays(ws, fileName, sheetName) {
  if (!ws) return [];
  const rows = XLSX.utils.sheet_to_json(ws, { header: "A", raw: false, defval: "", blankrows: false });
  const result = [];
  let currentWarehouse = null;

  rows.forEach((row) => {
    const a = String(row.A || "").trim();
    const dateIso = normalizeRevenueDate(a);
    const revenue = revenueToNumber(row.E);

    if (dateIso) {
      if (currentWarehouse && revenue != null && revenue > 0) {
        result.push({ dateIso, warehouse: currentWarehouse, revenue, source: `${fileName} / ${sheetName}` });
      }
      return;
    }

    if (a && isBlockedRevenueName(a)) {
      currentWarehouse = null;
      return;
    }

    const warehouse = cleanRevenueWarehouseName(a);
    if (warehouse) currentWarehouse = warehouse;
  });

  return result;
}

function aggregateRevenueRows(rows) {
  const map = new Map();
  rows.forEach((row) => {
    const parts = splitRevenueWarehouseName(row.warehouse);
    const key = `${row.dateIso}||${parts.restaurant}||${parts.warehouse}`;
    const existing = map.get(key);
    if (existing) {
      existing.revenue += row.revenue;
    } else {
      map.set(key, {
        dateIso: row.dateIso,
        restaurant: parts.restaurant,
        restaurantKey: restaurantMatchKey(parts.restaurant),
        warehouse: parts.warehouse,
        warehouseKind: getRevenueWarehouseKind(parts.warehouse),
        revenue: row.revenue,
        source: row.source
      });
    }
  });
  return Array.from(map.values());
}

function normalizeRevenueDate(value) {
  const text = String(value || "").trim();
  const shortMatch = text.match(/^(\d{2})\.(\d{2})\.(\d{2})$/);
  if (shortMatch) {
    const yy = Number(shortMatch[3]);
    const year = yy >= 70 ? 1900 + yy : 2000 + yy;
    return `${year}-${shortMatch[2]}-${shortMatch[1]}`;
  }
  const fullMatch = text.match(/^(\d{2})\.(\d{2})\.(\d{4})$/);
  if (fullMatch) return `${fullMatch[3]}-${fullMatch[2]}-${fullMatch[1]}`;
  return null;
}

function cleanRevenueWarehouseName(value) {
  const text = String(value || "").trim();
  if (!text) return null;
  if (/^\d+[.,]?\d*$/.test(text)) return null;
  if (/^отчет по продажам$/i.test(text)) return null;
  if (/^построен:/i.test(text)) return null;
  if (/^детализация:/i.test(text)) return null;
  if (/^наша компания:/i.test(text)) return null;
  if (/^выручки?$/i.test(text)) return null;
  if (/^сумма$/i.test(text)) return null;
  if (/^кол-во$/i.test(text)) return null;
  if (/^ед\.\s*изм\.?$/i.test(text)) return null;
  if (/^продажа$/i.test(text)) return null;
  if (/^лист\d*$/i.test(text)) return null;
  if (/^\d{2}\.\d{2}\.\d{4}\s*-\s*\d{2}\.\d{2}\.\d{4}$/.test(text)) return null;
  if (isBlockedRevenueName(text)) return null;
  return text;
}

function revenueToNumber(value) {
  if (value == null || value === "") return null;
  let s = String(value).trim();
  if (!s) return null;
  s = s.replace(/\u00a0/g, " ").replace(/\s/g, "");
  s = s.replace(/[^\d,.\-]/g, "");
  if (!s || s === "-" || s === "," || s === ".") return null;
  const lastComma = s.lastIndexOf(",");
  const lastDot = s.lastIndexOf(".");
  if (lastComma >= 0 && lastDot >= 0) {
    if (lastComma > lastDot) {
      s = s.replace(/\./g, "").replace(",", ".");
    } else {
      s = s.replace(/,/g, "");
    }
  } else if (lastComma >= 0) {
    s = s.replace(",", ".");
  }
  if (!/^-?\d+(\.\d+)?$/.test(s)) return null;
  const n = Number(s);
  return Number.isFinite(n) ? n : null;
}

function splitRevenueWarehouseName(name) {
  const original = String(name || "").trim();
  const key = revenueNameKey(original);
  let restaurant = original;

  if ((key.includes("бп") || key.includes("б покровская") || key.includes("большая покровская")) && key.includes("59")) {
    restaurant = "Б. Покровская, 59 Самурай";
  } else if ((key.includes("бп") || key.includes("б покровская") || key.includes("большая покровская")) && key.includes("63")) {
    restaurant = "Б. Покровская, 63 Самурай";
  } else if (key.includes("вп 14")) {
    restaurant = "Верхне-Печерская, 14Б Самурай";
  } else if (key.includes("каспарус") || (key.includes("циолковского") && key.includes("19"))) {
    restaurant = "Циолковского, 19А Самурай";
  } else if (key.includes("геологов")) {
    restaurant = "Геологов 7А Самурай";
  } else if ((key.includes("ленина") && key.includes("64")) || key.includes("ударник")) {
    restaurant = "Ленина, 64 Ударник";
  } else if (key.includes("ошар") || key.includes("ресторан xix")) {
    restaurant = "Ошарская, 8А 19";
  } else if (key.includes("винедо") || key.includes("vinedo")) {
    restaurant = "Октябрьская, 1 Винедо";
  } else if (key.includes("белинского") && key.includes("61") && key.includes("достав")) {
    restaurant = "Белинского, 61 доставка Самурай";
  } else if ((key.includes("белинского") && key.includes("61") && hasRibsMarker(original)) || /^RIBS\b/i.test(original) || hasRibsMarker(original)) {
    restaurant = "Белинского, 61 Рибс";
  } else if (key.includes("белинского") && key.includes("61")) {
    restaurant = "Белинского, 61 Самурай";
  } else if ((key.includes("гагарина") && key.includes("35")) || key.includes("парк швейцария")) {
    restaurant = "Парк Швейцария Самурай";
  } else if (key.includes("моторн") && (key.includes("пер") || key.includes("2к1") || key.includes("2 1"))) {
    restaurant = "Моторный, 2/1 доставка Самурай";
  } else if (key.includes("коминтерн") && key.includes("166")) {
    restaurant = "Коминтерна 166 CALL CENTRE";
  } else if (key.includes("коминтерн") && key.includes("115")) {
    restaurant = "Коминтерна, 115 Самурай";
  } else if (key.includes("волжская") && key.includes("13")) {
    restaurant = "Волжская, 13 Самурай";
  } else if (key.includes("веденяпина") && key.includes("1а")) {
    restaurant = "Веденяпина, 1А Самурай";
  } else if (key.includes("октября") && key.includes("2")) {
    restaurant = "Октября, 2 Самурай";
  } else if (key.includes("ленина") && key.includes("36")) {
    restaurant = "Ленина, 36 Самурай";
  } else if (key.includes("детский центр жюль верн")) {
    restaurant = "Ленина, 64 Ударник";
  } else if (/^Самурай,\s*/i.test(original)) {
    restaurant = original.replace(/\s*\([^)]*\)\s*$/g, "").replace(/^Самурай,\s*/i, "").trim();
    restaurant = restaurant + " Самурай";
  } else {
    restaurant = original.replace(/\s*\([^)]*\)\s*$/g, "").trim();
  }

  return { restaurant: canonicalRestaurantName(restaurant || original), warehouse: original };
}

function revenueNameKey(value) {
  return String(value || "")
    .toLowerCase()
    .replaceAll("ё", "е")
    .replace(/\u00a0/g, " ")
    .replace(/[.,;:()\-_/]+/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function restaurantMatchKey(value) {
  let key = revenueNameKey(value);
  key = key
    .replace(/\bсамурай\b/g, "")
    .replace(/\bcall\b|\bcentre\b|\bcenter\b/g, "")
    .replace(/\bд\b/g, "")
    .replace(/\s+/g, " ")
    .trim();
  if (key.includes("коминтерна") && key.includes("166")) key = key.replace(/\bдоставка\b/g, "").trim();
  return key;
}

function isBlockedRevenueName(value) {
  const key = revenueNameKey(value);
  return DEFAULT_REVENUE_EXCLUSIONS.some((rule) => rule && key.includes(rule));
}

function getRevenueWarehouseKind(name) {
  const key = revenueNameKey(name);
  if (key.includes("кухня")) return "kitchen";
  if (key.includes("бар")) return "bar";
  return "single";
}

function getSelectedWarehouseTypes() {
  return [...document.querySelectorAll('input[name="warehouseType"]:checked')]
    .map((input) => input.value);
}

function revenueMatchesWarehouseType(row) {
  const selectedTypes = getSelectedWarehouseTypes();
  if (selectedTypes.includes("all")) return true;
  if (!selectedTypes.length) return false;
  return selectedTypes.includes(row.warehouseKind);
}

function getRevenueFor(restaurant, dateIso) {
  if (!revenueRows.length) return 0;
  const key = restaurantMatchKey(restaurant);
  return revenueRows
    .filter((row) => row.dateIso === dateIso && revenueMatchesWarehouseType(row) && revenueRestaurantMatches(row.restaurantKey, key))
    .reduce((sum, row) => sum + row.revenue, 0);
}

function revenueRestaurantMatches(revenueKey, staffKey) {
  if (!revenueKey || !staffKey) return false;
  return revenueKey === staffKey || revenueKey.includes(staffKey) || staffKey.includes(revenueKey);
}

function formatMoney(value) {
  return Math.round(value || 0).toLocaleString("ru-RU");
}

function round2(n) {
  return Math.round((n + Number.EPSILON) * 100) / 100;
}

function detectFileType(rows) {
  if (!rows.length) return "unknown";
  const header = rows[0].map((h) => String(h).trim());
  const has = (name) => header.includes(name);

  if (has("Источник") && has("Направление") && has("Дата") && has("Время") && (has("ФИО") || (has("Фамилия") && has("Имя")))) {
    return "attendance";
  }

  if (has("ФИО") && (has("Название подразделения") || has("Подразделение"))) {
    return "staff";
  }

  const joined = rows.slice(0, 20).flat().map((v) => String(v || "")).join(" ");
  if (/Отчет по продажам|Детализация:\s*Склад|Продажа|Сумма/i.test(joined)) {
    return "revenue";
  }

  return "unknown";
}

function parseStaffRows(rows) {
  if (!rows.length) throw new Error("Файл сотрудников пустой.");

  const header = rows[0].map((h) => String(h).trim());
  const fioIdx = header.indexOf("ФИО");
  const restaurantIdx = findHeaderIndex(header, ["Название подразделения", "Подразделение"]);

  if (fioIdx === -1 || restaurantIdx === -1) {
    throw new Error("В файле сотрудников нужны колонки: ФИО и Название подразделения/Подразделение.");
  }

  const map = new Map();
  let conflicts = 0;
  const conflictKeys = new Set();

  for (let i = 1; i < rows.length; i += 1) {
    const row = rows[i];
    const fio = String(row[fioIdx] || "").trim();
    const restaurant = canonicalRestaurantName(row[restaurantIdx]);
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

function applyAttendanceData(records, options = {}) {
  baseRecords = records;
  rebuildMappedRecords();
  if (options.selectAllDates) {
    Array.from(dateSelect.options).forEach((option) => { option.selected = true; });
  }
  lastResultRows = [];
  tableBody.innerHTML = "";
  summaryEl.textContent = "Выберите фильтры и нажмите «Рассчитать».";
  csvBtn.disabled = true;
  xlsxBtn.disabled = true;
  refreshStatus();
  if (!options.skipAutoRevenue) maybeAutoLoadRevenueFromDatabase();
}

function applyRevenueData(rows) {
  const isFirstRevenueLoad = revenueRows.length === 0;
  revenueRows = rows;
  revenueStats = { rows: rows.length, matched: 0 };
  if (baseRecords.length) rebuildMappedRecords(isFirstRevenueLoad);
  if (lastResultRows.length) {
    lastResultRows = calculate(mappedRecords);
    renderTable(lastResultRows);
  }
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

  if (type === "revenue") {
    const parsedRevenueRows = parseRevenueWorkbook(buf, file.name);
    applyRevenueData(parsedRevenueRows);
    return { file: file.name, type: `выручки (${parsedRevenueRows.length})` };
  }

  return { file: file.name, type: "не распознан" };
}

function rebuildMappedRecords(selectAllRestaurants = false) {
  mappingStats = { matched: 0, total: baseRecords.length };

  mappedRecords = baseRecords.map((r) => {
    const mappedRestaurant = canonicalRestaurantName(staffRestaurantMap.get(r.personKey));
    if (mappedRestaurant) mappingStats.matched += 1;

    return {
      ...r,
      restaurant: mappedRestaurant || "Подразделение не определено",
      hasConflict: staffConflictKeys.has(r.personKey)
    };
  });

  const prevDates = getSelectedValues(dateSelect);
  const prevRestaurants = selectAllRestaurants ? [] : getSelectedValues(restaurantSelect);

  const dates = [...new Set(mappedRecords.map((r) => r.dateIso))].sort();
  // Attendance can cover only part of the restaurants. Revenue rows keep the
  // full restaurant list available even when a location has no gate events.
  const restaurants = [
    ...new Set([
      ...mappedRecords.map((r) => r.restaurant).filter((name) => !isNonRestaurantDepartment(name)),
      ...revenueRows.map((r) => r.restaurant)
    ])
  ].sort((a, b) => a.localeCompare(b, "ru"));

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

  selectedDates.forEach((dateIso) => {
    selectedRestaurants.forEach((restaurant) => {
      restaurantDay.set(`${dateIso}||${restaurant}`, {
        dateIso,
        restaurant,
        kitchen: 0,
        hall: 0,
        delivery: 0,
        bar: 0,
        total: 0,
        revenue: 0,
        hasConflict: false,
        details: { kitchen: [], hall: [], delivery: [], bar: [] }
      });
    });
  });

  Array.from(personDay.values()).forEach((item) => {
    const shiftValue = workedSecondsToShift(calcWorkedSeconds(item.events));
    if (shiftValue === 0) return;

    const key = `${item.dateIso}||${item.restaurant}`;
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
      row.revenue = getRevenueFor(row.restaurant, row.dateIso);
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
    return;
  }

  let totalKitchen = 0;
  let totalHall = 0;
  let totalDelivery = 0;
  let totalBar = 0;
  let totalRevenue = 0;

  rows.forEach((r) => {
    totalKitchen += r.kitchen;
    totalHall += r.hall;
    totalDelivery += r.delivery;
    totalBar += r.bar;
    totalRevenue += r.revenue || 0;

    const tr = document.createElement("tr");
    const detailsTr = document.createElement("tr");
    detailsTr.className = "detailsRow";
    detailsTr.style.display = "none";

    const detailsCell = document.createElement("td");
    detailsCell.colSpan = 9;
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
      <td>${formatMoney(r.revenue)}</td>
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

  summaryEl.textContent = `Строк: ${rows.length}. Кухня: ${formatShift(totalKitchen)}, Зал: ${formatShift(totalHall)}, Доставка: ${formatShift(totalDelivery)}, Бар: ${formatShift(totalBar)}, Выручка: ${formatMoney(totalRevenue)}, Всего смен: ${formatShift(totalKitchen + totalHall + totalDelivery + totalBar)}.`;
  csvBtn.disabled = false;
  xlsxBtn.disabled = false;
}

function toCSV(rows) {
  const head = ["Дата", "Ресторан", "Кухня", "Зал", "Доставка", "Бар", "Выручка", "Итого"];
  const lines = [head.join(";")];
  rows.forEach((r) => {
    lines.push([
      prettyDate(r.dateIso),
      `"${String(r.restaurant).replaceAll('"', '""')}"`,
      formatShift(r.kitchen),
      formatShift(r.hall),
      formatShift(r.delivery),
      formatShift(r.bar),
      r.revenue || 0,
      formatShift(r.total)
    ].join(";"));
  });
  return lines.join("\n");
}

function buildMatrix(rows, fieldName) {
  const restaurants = [...new Set(rows.map((r) => r.restaurant))].sort((a, b) => a.localeCompare(b, "ru"));
  const dates = [...new Set(rows.map((r) => r.dateIso))].sort();
  const staffMap = new Map(rows.map((r) => [`${r.restaurant}||${r.dateIso}`, r[fieldName]]));
  const revenueMap = new Map(rows.map((r) => [`${r.restaurant}||${r.dateIso}`, r.revenue || 0]));

  const aoa = [["Ресторан", "Показатель", ...dates.map(prettyDate)]];
  restaurants.forEach((restaurant) => {
    const revenueLine = [restaurant, "Выручка"];
    const staffLine = ["", "Кол-во персонала"];
    dates.forEach((dateIso) => {
      revenueLine.push(round2(revenueMap.get(`${restaurant}||${dateIso}`) || 0));
      staffLine.push(staffMap.get(`${restaurant}||${dateIso}`) || 0);
    });
    aoa.push(revenueLine, staffLine);
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

function refreshStatus() {
  if (!baseRecords.length) {
    statusEl.textContent = "Данные проходной пока не загружены.";
    return;
  }

  const staffLoaded = staffRestaurantMap.size > 0;
  const staffPart = staffLoaded
    ? ` Рестораны по подразделениям: сопоставлено ${mappingStats.matched} из ${mappingStats.total} записей.${staffConflicts ? ` Конфликтов ФИО: ${staffConflicts}.` : ""}`
    : " Подразделения ресторанов в проходной не найдены.";
  const revenuePart = revenueRows.length ? ` Выручек: ${revenueRows.length} строк.` : " Выручка не загружена.";

  statusEl.textContent = `Записей проходной: ${baseRecords.length}.${staffPart}${revenuePart}`;
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

revenueInput.addEventListener("change", async (e) => {
  const file = e.target.files?.[0];
  if (!file) return;
  try {
    const rows = parseRevenueWorkbook(await file.arrayBuffer(), file.name);
    applyRevenueData(rows);
    summaryEl.textContent = `Файл выручек загружен: ${rows.length} строк. Пересчитайте данные.`;
  } catch (err) {
    statusEl.textContent = `Ошибка файла выручек: ${err.message}`;
    revenueRows = [];
    if (baseRecords.length) rebuildMappedRecords();
    refreshStatus();
  }
});

if (revenueDbLoginBtn) revenueDbLoginBtn.addEventListener("click", signInRevenueDatabase);
if (revenueDbLogoutBtn) revenueDbLogoutBtn.addEventListener("click", signOutRevenueDatabase);
if (loadRevenueDbBtn) loadRevenueDbBtn.addEventListener("click", () => loadRevenueFromDatabase());
if (loadSabyBtn) loadSabyBtn.addEventListener("click", loadSabyFromApi);
if (sabyFromInput) sabyFromInput.addEventListener("change", updateRevenueDbButtons);
if (sabyToInput) sabyToInput.addEventListener("change", updateRevenueDbButtons);

warehouseTypeControl.addEventListener("change", (event) => {
  const changed = event.target.closest('input[name="warehouseType"]');
  if (!changed) return;
  const inputs = [...warehouseTypeControl.querySelectorAll('input[name="warehouseType"]')];
  if (changed.value === "all" && changed.checked) {
    inputs.forEach((input) => {
      if (input !== changed) input.checked = false;
    });
  } else if (changed.value !== "all" && changed.checked) {
    const allInput = inputs.find((input) => input.value === "all");
    if (allInput) allInput.checked = false;
  }
  if (!mappedRecords.length) return;
  summaryEl.textContent = "Состав складов выручки изменен. Нажмите «Рассчитать».";
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
    summaryEl.textContent = "Сначала загрузите данные из Saby или откройте резервную загрузку Excel.";
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

setDefaultSabyPeriod();
initRevenueDatabase();
