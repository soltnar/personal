import fs from "node:fs";
import path from "node:path";

const [projectPath] = process.argv.slice(2);
if (!projectPath) throw new Error("Usage: node tools/probe-saby-employees.mjs <project-with-saby-env>");

function loadEnv(filePath) {
  if (!fs.existsSync(filePath)) return {};
  return Object.fromEntries(
    fs.readFileSync(filePath, "utf8")
      .split(/\r?\n/)
      .map((line) => line.match(/^([^#=\s]+)=(.*)$/))
      .filter(Boolean)
      .map((match) => [match[1], match[2].trim().replace(/^['"]|['"]$/g, "")])
  );
}

const env = { ...loadEnv(path.join(projectPath, ".env")), ...process.env };
const keyFile = fs.readdirSync(projectPath).find((name) => name.endsWith(".key"));
const serviceKey = env.SABY_SERVICE_KEY || (keyFile
  ? fs.readFileSync(path.join(projectPath, keyFile), "utf8").trim()
  : "");

const authResponse = await fetch("https://online.sbis.ru/oauth/service/", {
  method: "POST",
  headers: { "Content-Type": "application/json" },
  body: JSON.stringify({
    app_client_id: env.SABY_APP_CLIENT_ID,
    app_secret: env.SABY_APP_SECRET,
    secret_key: serviceKey,
  }),
});
const auth = await authResponse.json();
if (!authResponse.ok || !auth.token) throw new Error(`Saby OAuth failed: ${authResponse.status}`);

const rpcResponse = await fetch("https://online.saby.ru/service/?srv=1", {
  method: "POST",
  headers: {
    "Content-Type": "application/json; charset=utf-8",
    "X-SBISAccessToken": auth.token,
  },
  body: JSON.stringify({
    jsonrpc: "2.0",
    method: "СБИС.СписокСотрудников",
    params: {
      Параметр: {
        Фильтр: { ВернутьУволенных: "Да" },
        Навигация: { РазмерСтраницы: "3", Страница: "0" },
      },
    },
    id: 1,
  }),
});
const data = await rpcResponse.json();
const result = data.result || {};
const employees = result.Сотрудник || result.employees || [];
console.log(JSON.stringify({
  httpStatus: rpcResponse.status,
  ok: rpcResponse.ok && !data.error,
  error: data.error || null,
  resultKeys: Object.keys(result),
  employeeCount: Array.isArray(employees) ? employees.length : null,
  employeeKeys: Array.isArray(employees) && employees[0] ? Object.keys(employees[0]) : [],
  navigation: result.Навигация || null,
}, null, 2));
