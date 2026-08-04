import "jsr:@supabase/functions-js/edge-runtime.d.ts";

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const SABY_APP_CLIENT_ID = Deno.env.get("SABY_APP_CLIENT_ID")!;
const SABY_APP_SECRET = Deno.env.get("SABY_APP_SECRET")!;
const SABY_SERVICE_KEY = Deno.env.get("SABY_SERVICE_KEY")!;
const ALLOWED_EMAIL = "soltnar@gmail.com";
const PAGE_SIZE = 1000;
const MAX_PAGES = 40;
const SABY_REQUEST_TIMEOUT_MS = 25_000;
const SABY_REQUEST_ATTEMPTS = 2;

const cors = {
  "Access-Control-Allow-Origin": "https://soltnar.github.io",
  "Access-Control-Allow-Headers": "authorization, apikey, content-type",
  "Access-Control-Allow-Methods": "GET, OPTIONS",
};

const json = (body: unknown, status = 200) => new Response(JSON.stringify(body), {
  status,
  headers: { ...cors, "Content-Type": "application/json" },
});

function validatePeriod(from: string, to: string) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(from) || !/^\d{4}-\d{2}-\d{2}$/.test(to) || from > to) {
    throw new Error("Некорректный период");
  }
  const fromDate = new Date(`${from}T12:00:00Z`);
  const toDate = new Date(`${to}T12:00:00Z`);
  const days = Math.round((toDate.getTime() - fromDate.getTime()) / 86400000) + 1;
  if (days > 62) throw new Error("За один запрос можно получить не более 62 дней");
}

function nextDate(date: string) {
  const value = new Date(`${date}T12:00:00Z`);
  value.setUTCDate(value.getUTCDate() + 1);
  return value.toISOString().slice(0, 10);
}

async function authorize(req: Request) {
  const authorization = req.headers.get("authorization") || "";
  if (!authorization.startsWith("Bearer ")) return false;
  const response = await fetch(`${SUPABASE_URL}/auth/v1/user`, {
    headers: { authorization, apikey: SERVICE_KEY },
  });
  if (!response.ok) return false;
  const user = await response.json();
  return String(user.email || "").toLowerCase() === ALLOWED_EMAIL;
}

async function sabyToken() {
  const response = await fetch("https://online.sbis.ru/oauth/service/", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      app_client_id: SABY_APP_CLIENT_ID,
      app_secret: SABY_APP_SECRET,
      secret_key: SABY_SERVICE_KEY,
    }),
  });
  const auth = await response.json();
  if (!response.ok || !auth.token) throw new Error("Saby не выдал сервисный токен");
  return auth.token as string;
}

function record(fields: Array<[unknown, string, unknown]>) {
  return {
    d: fields.map((field) => field[2]),
    s: fields.map((field) => ({ t: field[0], n: field[1] })),
    _type: "record",
    f: 0,
  };
}

function attendanceRequest(from: string, to: string, page: number) {
  const filter = record([
    ["Число целое", "CheckpointType", 0],
    ["Число целое", "ActionType", null],
    ["Логическое", "OnlyMain", false],
    [{ n: "Массив", t: "Строка" }, "IDs", null],
    [{ n: "Массив", t: "Число целое" }, "Departments", null],
    [{ n: "Массив", t: "Число целое" }, "Employees", null],
    [{ n: "Массив", t: "Число целое" }, "Rooms", []],
    [{ n: "Массив", t: "Дата" }, "Dates", [from, nextDate(to)]],
    [{ n: "Массив", t: "Число целое" }, "Minutes", []],
    ["Строка", "SearchString", ""],
    ["Число целое", "Organization", -2],
    ["Число целое", "TimezoneOffset", 180],
  ]);
  return {
    jsonrpc: "2.0",
    protocol: 7,
    method: "ActivityFixation.GetPersonsMainEvents",
    params: {
      "Фильтр": filter,
      "Сортировка": null,
      "Навигация": record([
        ["Логическое", "ЕстьЕще", true],
        ["Число целое", "РазмерСтраницы", PAGE_SIZE],
        ["Число целое", "Страница", page],
      ]),
      "ДопПоля": [],
    },
    id: page + 1,
  };
}

function rowsFromResult(result: { s?: Array<{ n: string }>; d?: unknown[][] }) {
  const names = (result.s || []).map((field) => field.n);
  return (result.d || []).map((values) => Object.fromEntries(
    names.map((name, index) => [name, values[index]]),
  ));
}

function cleanText(value: unknown) {
  return typeof value === "string" ? value.trim() : "";
}

async function fetchAttendance(from: string, to: string) {
  const startedAt = Date.now();
  const token = await sabyToken();
  const rows: Record<string, unknown>[] = [];
  let pagesLoaded = 0;

  const fetchPage = async (page: number) => {
    for (let attempt = 1; attempt <= SABY_REQUEST_ATTEMPTS; attempt += 1) {
      try {
        const response = await fetch("https://online.saby.ru/service/", {
          method: "POST",
          signal: AbortSignal.timeout(SABY_REQUEST_TIMEOUT_MS),
          headers: {
            "Content-Type": "application/json; charset=utf-8",
            "X-SBISAccessToken": token,
            "X-CalledMethod": "ActivityFixation.GetPersonsMainEvents",
            "X-Requested-With": "XMLHttpRequest",
          },
          body: JSON.stringify(attendanceRequest(from, to, page)),
        });
        const data = await response.json();
        if (!response.ok || data.error) {
          throw new Error(`Saby не вернул проходную: ${data.error?.message || response.status}`);
        }
        return {
          rows: rowsFromResult(data.result),
          hasMore: data.result?.n !== false,
        };
      } catch (error) {
        if (attempt === SABY_REQUEST_ATTEMPTS) throw error;
        console.log(JSON.stringify({ event: "attendance_page_retry", page, attempt }));
      }
    }
    throw new Error("Saby не вернул страницу проходной");
  };

  // Saby throttles parallel report requests, so larger pages are faster and
  // more reliable than concurrent pagination.
  for (let page = 0; page < MAX_PAGES; page += 1) {
    const result = await fetchPage(page);
    pagesLoaded += 1;
    rows.push(...result.rows);
    if (result.rows.length < PAGE_SIZE || !result.hasMore) break;
    if (page === MAX_PAGES - 1) throw new Error("Превышен лимит страниц проходной");
  }

  const unique = new Map<string, Record<string, unknown>>();
  rows.forEach((row) => {
    if (Number(row.Source) !== 3) return;
    const location = row.LocationInfo && typeof row.LocationInfo === "object"
      ? row.LocationInfo as Record<string, unknown>
      : {};
    const person = [row.LastName, row.FirstName, row.MiddleName].map(cleanText).filter(Boolean).join(" ")
      || cleanText(row.Person);
    const normalized = {
      dateTime: cleanText(row.DateTime),
      actionType: Number(row.ActionType),
      person,
      position: cleanText(row.Position),
      department: cleanText(row.DepartmentName),
      address: cleanText(location.Address) || cleanText(row.Address),
      accessPoint: cleanText(location.AccessPointName),
    };
    if (!normalized.dateTime || !normalized.person || !Number.isFinite(normalized.actionType)) return;
    const key = [normalized.dateTime, normalized.actionType, normalized.person, normalized.accessPoint].join("|");
    unique.set(key, normalized);
  });
  console.log(JSON.stringify({
    event: "attendance_loaded",
    from,
    to,
    pagesLoaded,
    sourceRows: rows.length,
    uniqueRows: unique.size,
    elapsedMs: Date.now() - startedAt,
  }));
  return { rows: [...unique.values()], pagesLoaded, elapsedMs: Date.now() - startedAt };
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: cors });
  if (req.method !== "GET") return json({ error: "Метод не поддерживается" }, 405);
  try {
    if (!(await authorize(req))) return json({ error: "Доступ запрещён" }, 403);
    const url = new URL(req.url);
    const from = url.searchParams.get("from") || "";
    const to = url.searchParams.get("to") || "";
    validatePeriod(from, to);
    const attendance = await fetchAttendance(from, to);
    return json({
      from,
      to,
      generatedAt: new Date().toISOString(),
      rows: attendance.rows,
      diagnostics: { pages: attendance.pagesLoaded, elapsedMs: attendance.elapsedMs },
    });
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : "Ошибка сервера" }, 400);
  }
});
