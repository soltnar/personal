import fs from "node:fs";
import path from "node:path";

const [projectPath, dateFrom = "01.08.2026", dateTo = "31.08.2026"] = process.argv.slice(2);
if (!projectPath) throw new Error("Usage: node tools/probe-saby-timesheets.mjs <project-with-saby-env> [from] [to]");

function loadEnv(filePath) {
  return Object.fromEntries(fs.readFileSync(filePath, "utf8").split(/\r?\n/)
    .map((line) => line.match(/^([^#=\s]+)=(.*)$/)).filter(Boolean)
    .map((match) => [match[1], match[2].trim().replace(/^['"]|['"]$/g, "")]));
}

const env = { ...loadEnv(path.join(projectPath, ".env")), ...process.env };
const keyFile = fs.readdirSync(projectPath).find((name) => name.endsWith(".key"));
const serviceKey = env.SABY_SERVICE_KEY || fs.readFileSync(path.join(projectPath, keyFile), "utf8").trim();
const authResponse = await fetch("https://online.sbis.ru/oauth/service/", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({ app_client_id: env.SABY_APP_CLIENT_ID, app_secret: env.SABY_APP_SECRET, secret_key: serviceKey }),
});
const auth = await authResponse.json();

const response = await fetch("https://online.sbis.ru/service/?srv=1", {
  method: "POST",
  headers: { "Content-Type": "application/json; charset=utf-8", "X-SBISAccessToken": auth.token },
  body: JSON.stringify({
    jsonrpc: "2.0",
    method: "СБИС.СписокДокументов",
    params: {
      Фильтр: {
        ДатаС: dateFrom,
        ДатаПо: dateTo,
        Тип: "ТабельДокумент",
        Навигация: { РазмерСтраницы: "10", Страница: "0" },
      },
    },
    id: 1,
  }),
});
const data = await response.json();
const documents = data.result?.Документ || [];
const listedSheets = Array.isArray(documents[0]?.Табели) ? documents[0].Табели : [];
const listedDays = listedSheets.flatMap((sheet) => Array.isArray(sheet.ДанныеТабеля) ? sheet.ДанныеТабеля : []);
let readSummary = null;
if (documents[0]?.Идентификатор) {
  const readResponse = await fetch("https://online.sbis.ru/service/?srv=1", {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8", "X-SBISAccessToken": auth.token },
    body: JSON.stringify({
      jsonrpc: "2.0",
      method: "СБИС.ПрочитатьДокумент",
      params: { Документ: { Идентификатор: documents[0].Идентификатор } },
      id: 2,
    }),
  });
  const readData = await readResponse.json();
  const document = readData.result || {};
  const sheets = Array.isArray(document.Табели) ? document.Табели : [];
  const days = sheets.flatMap((sheet) => Array.isArray(sheet.ДанныеТабеля) ? sheet.ДанныеТабеля : []);
  readSummary = {
    ok: readResponse.ok && !readData.error,
    error: readData.error || null,
    resultKeys: Object.keys(document),
    sheetCount: sheets.length,
    sheetKeys: sheets[0] ? Object.keys(sheets[0]) : [],
    dayCount: days.length,
    dayKeys: days[0] ? Object.keys(days[0]) : [],
  };
}
console.log(JSON.stringify({
  httpStatus: response.status,
  ok: response.ok && !data.error,
  error: data.error || null,
  count: Array.isArray(documents) ? documents.length : null,
  documentKeys: documents[0] ? Object.keys(documents[0]) : [],
  listedSheetCount: listedSheets.length,
  listedSheetKeys: listedSheets[0] ? Object.keys(listedSheets[0]) : [],
  listedDayCount: listedDays.length,
  listedDayKeys: listedDays[0] ? Object.keys(listedDays[0]) : [],
  navigation: data.result?.Навигация || null,
  readSummary,
}, null, 2));
