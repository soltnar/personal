import "jsr:@supabase/functions-js/edge-runtime.d.ts";
import { nextAutomaticDate, failureValues, syncActive, retryDue } from './schedule.mjs';

const SUPABASE_URL = Deno.env.get("SUPABASE_URL")!;
const SERVICE_KEY = Deno.env.get("SUPABASE_SERVICE_ROLE_KEY")!;
const SABY_APP_CLIENT_ID = Deno.env.get("SABY_APP_CLIENT_ID")!;
const SABY_APP_SECRET = Deno.env.get("SABY_APP_SECRET")!;
const SABY_SERVICE_KEY = Deno.env.get("SABY_SERVICE_KEY")!;
const ALLOWED_EMAIL = "soltnar@gmail.com";
// Match Saby's actual UI pagination. Larger sizes intermittently hang even on
// page zero (verified against the same date and token in server probes).
const PAGE_SIZE = 30;
const MAX_PAGES = 200;
const SABY_REQUEST_TIMEOUT_MS = 25_000;
const SABY_REQUEST_ATTEMPTS = 1;
const EMPLOYEE_PAGE_SIZE = 100;
const EMPLOYEE_CACHE_MS = 60 * 60 * 1000;
const CACHE_FORMAT_VERSION = 4;

let employeeCache: { expiresAt: number; rows: Record<string, unknown>[] } | null = null;
const CACHE_TABLE = "attendance_daily_cache";
const CACHE_STALE_MS = 5 * 60_000;

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

function addDays(date: string, days: number) {
  const value = new Date(`${date}T12:00:00Z`);
  value.setUTCDate(value.getUTCDate() + days);
  return value.toISOString().slice(0, 10);
}

function datesInRange(from: string, to: string) {
  const dates: string[] = [];
  for (let date = from; date <= to; date = nextDate(date)) dates.push(date);
  return dates;
}

async function cacheRequest(path: string, init: RequestInit = {}) {
  const response = await fetch(`${SUPABASE_URL}/rest/v1/${CACHE_TABLE}${path}`, {
    ...init,
    signal: AbortSignal.timeout(20000),
    headers: {
      apikey: SERVICE_KEY,
      Authorization: `Bearer ${SERVICE_KEY}`,
      "Content-Type": "application/json",
      Prefer: "return=representation,resolution=merge-duplicates",
      ...(init.headers || {}),
    },
  });
  if (!response.ok) throw new Error(`Ошибка кэша проходной: ${response.status}`);
  const text = await response.text();
  return text ? JSON.parse(text) : [];
}

function progressFilter(progress: Record<string,unknown> | null | undefined) {
  if(!progress)return '&sync_progress=is.null';
  // Compare only small queue/lease identifiers. Never put accumulated page rows
  // in a URL: even a few pages exceed the HTTP client's URL length limit.
  return ['requested_at','lease_id','lease_at'].map(key=>`&sync_progress->>${key}=${progress[key]?'eq.'+encodeURIComponent(String(progress[key])):'is.null'}`).join('');
}

async function readCache(from: string, to: string) {
  const query = `?select=work_date,status,rows,error,updated_at,sync_progress&work_date=gte.${from}&work_date=lte.${to}&order=work_date.asc`;
  return await cacheRequest(query) as Array<{
    work_date: string;
    status: string;
    rows: Record<string, unknown>[] | null;
    error: string | null;
    updated_at: string;
    sync_progress?: {signature:string;next_page:number;rows:Record<string,unknown>[];complete:boolean;attempt_day:string;stable:boolean;failed_at?:string;paused?:boolean;failures?:number;retry_after?:string;requested_at?:string;lease_at?:string;lease_id?:string};
  }>;
}

async function writeCache(workDate: string, values: Record<string, unknown>) {
  return await cacheRequest("?on_conflict=work_date", {
    method: "POST",
    body: JSON.stringify({ work_date: workDate, ...values, updated_at: new Date().toISOString() }),
  });
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

async function authorizeCron(req: Request) {
  const key = req.headers.get("x-cron-key") || "";
  if (!key) return false;
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(key));
  const hash = [...new Uint8Array(digest)].map((value) => value.toString(16).padStart(2, "0")).join("");
  const response = await fetch(`${SUPABASE_URL}/rest/v1/attendance_sync_config?select=secret_hash&id=eq.1&secret_hash=eq.${hash}`, {
    headers: { apikey: SERVICE_KEY, Authorization: `Bearer ${SERVICE_KEY}` },
  });
  const rows = response.ok ? await response.json() : [];
  return Array.isArray(rows) && rows.length === 1;
}

async function sabyToken() {
  const response = await fetch("https://online.sbis.ru/oauth/service/", {
    signal: AbortSignal.timeout(15000),
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

function operationalDate(value: unknown) {
  const match = cleanText(value).match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
  if (!match) return "";
  const [, year, month, day, hour] = match;
  const date = new Date(Date.UTC(Number(year), Number(month) - 1, Number(day)));
  if (Number(hour) < 4) date.setUTCDate(date.getUTCDate() - 1);
  return date.toISOString().slice(0, 10);
}

async function fetchEmployees(token: string) {
  if (employeeCache && employeeCache.expiresAt > Date.now()) return employeeCache.rows;

  const employees: Record<string, unknown>[] = [];
  for (let page = 0; page < MAX_PAGES; page += 1) {
    let data: Record<string, any> | null = null;
    for (let attempt = 1; attempt <= SABY_REQUEST_ATTEMPTS; attempt += 1) {
      try {
        const response = await fetch("https://online.saby.ru/service/?srv=1", {
          method: "POST",
          signal: AbortSignal.timeout(SABY_REQUEST_TIMEOUT_MS),
          headers: {
            "Content-Type": "application/json; charset=utf-8",
            "X-SBISAccessToken": token,
          },
          body: JSON.stringify({
            jsonrpc: "2.0",
            method: "СБИС.СписокСотрудников",
            params: {
              "Параметр": {
                "Фильтр": { "ВернутьУволенных": "Да" },
                "Навигация": { "РазмерСтраницы": String(EMPLOYEE_PAGE_SIZE), "Страница": String(page) },
              },
            },
            id: page + 1,
          }),
        });
        data = await response.json();
        if (!response.ok || data?.error) {
          throw new Error(`Saby не вернул список сотрудников: ${data?.error?.message || response.status}`);
        }
        break;
      } catch (error) {
        if (attempt === SABY_REQUEST_ATTEMPTS) throw error;
        console.log(JSON.stringify({ event: "official_employees_page_retry", page, attempt }));
      }
    }
    const employeeData = data || {};
    const pageRows = Array.isArray(employeeData.result?.["Сотрудник"])
      ? employeeData.result["Сотрудник"] as Record<string, unknown>[]
      : [];
    employees.push(...pageRows);
    if (pageRows.length < EMPLOYEE_PAGE_SIZE || employeeData.result?.["Навигация"]?.["ЕстьЕще"] !== "Да") break;
  }

  const rows = employees.map((employee) => {
    const position = employee["Должность"] && typeof employee["Должность"] === "object"
      ? employee["Должность"] as Record<string, unknown> : {};
    const department = employee["Подразделение"] && typeof employee["Подразделение"] === "object"
      ? employee["Подразделение"] as Record<string, unknown> : {};
    return {
      person: [employee["Фамилия"], employee["Имя"], employee["Отчество"]].map(cleanText).filter(Boolean).join(" "),
      position: cleanText(position["Название"]),
      department: cleanText(department["Название"]),
      hired: cleanText(employee["Принят"]),
      fired: cleanText(employee["Уволен"]),
    };
  }).filter((employee) => employee.person);
  employeeCache = { expiresAt: Date.now() + EMPLOYEE_CACHE_MS, rows };
  console.log(JSON.stringify({ event: "official_employees_loaded", rows: rows.length }));
  return rows;
}

async function fetchAttendance(from: string, to: string, token: string, progress: Awaited<ReturnType<typeof readCache>>[number]['sync_progress'] | null = null, saveProgress: (value:Record<string,unknown>)=>Promise<unknown> = async()=>{}) {
  const startedAt = Date.now();
  const rows: Record<string, unknown>[] = progress?.rows ? [...progress.rows] : [];
  let pagesLoaded = 0;
  if(progress && !progress.complete && progress.next_page>=MAX_PAGES)throw new Error('Превышен лимит страниц проходной');

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
        if(response.status===429){
          const error=new Error(`Saby ограничил частоту запросов (HTTP 429). Retry-After: ${response.headers.get('Retry-After')||'не указан'}. Повторите позже.`);
          error.name='SabyRateLimitError';throw error;
        }
        const data = await response.json();
        if (!response.ok || data.error) {
          throw new Error(`Saby не вернул проходную: ${data.error?.message || response.status}`);
        }
        return {
          rows: rowsFromResult(data.result),
          hasMore: typeof data.result?.n === "boolean" ? data.result.n : null,
        };
      } catch (error) {
        if(error instanceof Error && error.name==='SabyRateLimitError')throw error;
        if (attempt === SABY_REQUEST_ATTEMPTS) throw error;
        console.log(JSON.stringify({ event: "attendance_page_retry", page, attempt, pageSize:PAGE_SIZE, error:error instanceof Error?error.message:'unknown' }));
        await new Promise(resolve=>setTimeout(resolve,2000));
      }
    }
    throw new Error("Saby не вернул страницу проходной");
  };

  // Serial requests with a >=1.6s gap: below 38 page starts/minute per loader.
  for (let page = progress?.complete ? MAX_PAGES : progress?.next_page || 0; page < MAX_PAGES; page += 1) {
    if(page>0)await new Promise(resolve=>setTimeout(resolve,1600));
    const result = await fetchPage(page);
    pagesLoaded += 1;
    if(result.hasMore===true&&!result.rows.length)throw new Error('Saby вернул пустую страницу с признаком продолжения');
    const signature=JSON.stringify(result.rows);
    for(let offset=0;result.rows.length&&offset<rows.length;offset+=PAGE_SIZE){
      if(JSON.stringify(rows.slice(offset,offset+result.rows.length))===signature)throw new Error('Saby повторяет страницу проходной; загрузка остановлена без потери сохранённых страниц');
    }
    rows.push(...result.rows);
    const complete=result.hasMore===false || (result.hasMore===null&&result.rows.length<PAGE_SIZE);
    await saveProgress({next_page:page+1,rows,complete});
    if (complete) break;
    if(pagesLoaded>=4||Date.now()-startedAt>=10000){const error=new Error('Продолжение сохранённой загрузки');error.name='AttendanceBatchPaused';throw error;}
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
      personId: String(row.Person || ""),
      position: cleanText(row.Position),
      department: cleanText(row.DepartmentName) || [location.LocationName, location.Address, location.AccessPointName]
        .map(cleanText).filter(Boolean).join(" "),
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

async function syncCacheDates(workDates: string[], previous: Awaited<ReturnType<typeof readCache>> = []) {
  const startedAt=Date.now(), firstDate=workDates[0],lastDate=workDates.at(-1)!;
  const old=previous.find(entry=>entry.work_date===firstDate)!;
  const signature=JSON.stringify([firstDate,lastDate,PAGE_SIZE,CACHE_FORMAT_VERSION]);
  const today=new Intl.DateTimeFormat('en-CA',{timeZone:'Europe/Moscow',year:'numeric',month:'2-digit',day:'2-digit'}).format(new Date());
  const saved=old.sync_progress;
  const resumable=saved?.signature===signature&&(saved.stable||saved.attempt_day===today);
  let lastProgress:Record<string,any>={...(resumable?saved:{}),signature,requested_at:saved?.requested_at||new Date().toISOString(),
    next_page:resumable?saved!.next_page:0,rows:resumable?saved!.rows:[],complete:resumable?saved!.complete:false,
    attempt_day:today,stable:nextDate(lastDate)<today,lease_id:saved?.lease_id,lease_at:new Date().toISOString(),paused:false};
  const patch=async(values:Record<string,unknown>)=>{
    const result=await cacheRequest(`?work_date=eq.${firstDate}&sync_progress->>lease_id=eq.${encodeURIComponent(saved?.lease_id||'')}`,{method:'PATCH',body:JSON.stringify(values)});
    if(!result.length){const error=new Error('Загрузку продолжает другой обработчик');error.name='AttendanceLeaseLost';throw error;}
  };
  try{
    await patch({sync_progress:lastProgress});
    const token=await sabyToken();
    const attendance=await fetchAttendance(firstDate,lastDate,token,lastProgress as any,async values=>{
      lastProgress={...lastProgress,...values,lease_at:new Date().toISOString(),paused:false,failures:0,retry_after:null};
      await patch({sync_progress:lastProgress});
    });
    const rows=attendance.rows.filter(row=>operationalDate(row.dateTime)===firstDate).map(row=>({...row,cacheVersion:CACHE_FORMAT_VERSION}));
    await patch({status:'ready',rows,error:null,sync_progress:null,elapsed_ms:attendance.elapsedMs,updated_at:new Date().toISOString()});
    console.log(JSON.stringify({event:'attendance_cache_batch_ready',from:firstDate,to:lastDate,rows:rows.length,elapsedMs:Date.now()-startedAt}));
  }catch(error){
    if(error instanceof Error&&error.name==='AttendanceLeaseLost')return;
    const paused=error instanceof Error&&error.name==='AttendanceBatchPaused';
    const failures=paused?0:Number(lastProgress.failures||0)+1;
    const message=error instanceof Error?error.message:'Ошибка Saby';
    try{await patch({...(!paused?failureValues(old,message,Date.now()-startedAt):{}),sync_progress:{...lastProgress,paused:true,failures,
      retry_after:paused?null:new Date(Date.now()+Math.min(900000,30000*2**Math.min(failures-1,5))).toISOString(),
      failed_at:paused?null:new Date().toISOString()}});}catch(saveError){if(!(saveError instanceof Error&&saveError.name==='AttendanceLeaseLost'))throw saveError;}
    if(!paused)console.error(JSON.stringify({event:'attendance_cache_batch_error',from:firstDate,to:lastDate,error:message,elapsedMs:Date.now()-startedAt}));
  }
}

Deno.serve(async (req) => {
  if (req.method === "OPTIONS") return new Response(null, { headers: cors });
  if (req.method !== "GET") return json({ error: "Метод не поддерживается" }, 405);
  try {
    const url = new URL(req.url);
    const refreshMode = url.searchParams.get("refresh") || "";
    const cronRequest = refreshMode === "recent" || refreshMode === "backfill";
    if (!(cronRequest ? await authorizeCron(req) : await authorize(req))) return json({ error: "Доступ запрещён" }, 403);
    if (!cronRequest && url.searchParams.get("resource") === "employees") {
      const employees = await fetchEmployees(await sabyToken());
      return json({ employees, generatedAt: new Date().toISOString() });
    }
    const todayMoscow = new Intl.DateTimeFormat("en-CA", { timeZone: "Europe/Moscow", year: "numeric", month: "2-digit", day: "2-digit" }).format(new Date());
    let automaticDate: string | null = null;
    if(refreshMode==='recent') {
      const entries=[];
      for(let offset=0;;offset+=1000){
        const page=await cacheRequest(`?select=work_date,status,updated_at,sync_progress&work_date=lte.${todayMoscow}&order=work_date.asc&limit=1000&offset=${offset}`);
        entries.push(...page);if(page.length<1000)break;
      }
      automaticDate=nextAutomaticDate(entries,todayMoscow);
      if(!automaticDate)return json({complete:true,skipped:true,reason:'Нет дней для автоматической загрузки сегодня'});
    }
    const from = automaticDate || url.searchParams.get("from") || "";
    const to = automaticDate || url.searchParams.get("to") || "";
    validatePeriod(from, to);
    const requestedDates = datesInRange(from, to);
    const cached = await readCache(from, to);
    const byDate = new Map(cached.map((entry) => [entry.work_date, entry]));
    const now = Date.now();
    const forceBefore=refreshMode==='force'?url.searchParams.get('before')||'':'';
    if(refreshMode==='force'&&(!Number.isFinite(Date.parse(forceBefore))||Date.parse(forceBefore)>now+60000))throw new Error('Некорректный момент обновления');
    const refreshCutoff=forceBefore||`${todayMoscow}T00:00:00Z`;
    const maySync=cronRequest||refreshMode==='force';
    const isCurrentCacheEntry=(entry:typeof cached[number]|undefined)=>Boolean(entry?.status==='ready'&&Array.isArray(entry.rows)&&entry.rows.every(row=>Number(row.cacheVersion)===CACHE_FORMAT_VERSION));
    // Persist the entire requested date range before starting any source reads.
    // Polling and cron both consume this same durable queue.
    if(maySync){
      const queued=requestedDates.filter(date=>{
        const entry=byDate.get(date);
        if(entry?.sync_progress?.requested_at||syncActive(entry,now))return false;
        return !isCurrentCacheEntry(entry)||(refreshMode==='force'?entry!.updated_at<refreshCutoff:!!automaticDate);
      }).map(date=>({date,entry:byDate.get(date)}));
      for(let offset=0;offset<queued.length;offset+=8){
        await Promise.all(queued.slice(offset,offset+8).map(async({date,entry})=>{
          const progress={requested_at:forceBefore||new Date().toISOString(),next_page:0,rows:[],complete:false,paused:true,failures:0};
          const result=entry?await cacheRequest(`?work_date=eq.${date}&updated_at=eq.${encodeURIComponent(entry.updated_at)}${progressFilter(entry.sync_progress)}`,{method:'PATCH',body:JSON.stringify({sync_progress:progress,error:null})})
            :await cacheRequest('?on_conflict=work_date',{method:'POST',headers:{Prefer:'return=representation,resolution=ignore-duplicates'},body:JSON.stringify({work_date:date,status:'loading',rows:[],error:null,sync_progress:progress,updated_at:new Date().toISOString()})});
          if(result.length)byDate.set(date,result[0]);
        }));
      }
    }
    const current=[...byDate.values()];
    const activeLoading=current.some(entry=>syncActive(entry,now));
    const nextMissing=current.find(entry=>entry.sync_progress?.requested_at&&retryDue(entry,now)&&!syncActive(entry,now));
    let syncingDates:string[]=[];
    if(maySync&&!activeLoading&&nextMissing){
      const lease={...nextMissing.sync_progress,lease_id:crypto.randomUUID(),lease_at:new Date().toISOString(),paused:false};
      const claim=await cacheRequest(`?work_date=eq.${nextMissing.work_date}${progressFilter(nextMissing.sync_progress)}`,{method:'PATCH',body:JSON.stringify({sync_progress:lease})});
      if(claim.length){syncingDates=[nextMissing.work_date];EdgeRuntime.waitUntil(syncCacheDates(syncingDates,claim));}
    }

    const rows = current.filter(isCurrentCacheEntry).flatMap((entry) => entry.rows || []);
    const readyDates = current.filter(isCurrentCacheEntry)
      .filter((entry) => refreshMode !== "force" || (entry.updated_at >= refreshCutoff && !entry.sync_progress?.requested_at))
      .map((entry) => entry.work_date);
    const pendingDates = requestedDates.filter((date) => !readyDates.includes(date));
    const errors:unknown[]=[]; // Source failures are retried by the persistent queue.
    const progressEntry=current.find(entry=>entry.sync_progress?.requested_at);

    return json({
      from,
      to,
      generatedAt: new Date().toISOString(),
      rows,
      employees: [],
      complete: pendingDates.length === 0,
      readyDates,
      pendingDates,
      errors,
      diagnostics: {
        cache: true,
        persistentQueue: true,
        cacheHits: readyDates.length,
        cacheTotal: requestedDates.length,
        syncingDate: syncingDates[0] || null,
        syncingDates,
        progressPage: progressEntry?.sync_progress?.next_page || 0,
        progressRows: progressEntry?.sync_progress?.rows?.length || 0,
        retryAfter: progressEntry?.sync_progress?.retry_after || null,
        lastError: progressEntry?.error || null,
        queuedDays: current.filter(entry=>entry.sync_progress?.requested_at).length,
        employees: 0,
        employeeSource: "separate-request",
      },
    });
  } catch (error) {
    return json({ error: error instanceof Error ? error.message : "Ошибка сервера" }, 400);
  }
});
