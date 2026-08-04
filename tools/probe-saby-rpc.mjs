import fs from "node:fs";
import path from "node:path";

const [harPath, revenueProjectPath, methodName, requestedPageSize] = process.argv.slice(2);
if (!harPath || !revenueProjectPath || !methodName) {
  throw new Error("Usage: node tools/probe-saby-rpc.mjs <har> <revenue-project> <method>");
}

function loadEnv(filePath) {
  if (!fs.existsSync(filePath)) return {};
  return Object.fromEntries(
    fs
      .readFileSync(filePath, "utf8")
      .split(/\r?\n/)
      .map((line) => line.match(/^([^#=\s]+)=(.*)$/))
      .filter(Boolean)
      .map((match) => [match[1], match[2].trim().replace(/^['"]|['"]$/g, "")])
  );
}

function findRequest(value, method) {
  if (Array.isArray(value)) {
    for (const item of value) {
      const found = findRequest(item, method);
      if (found) return found;
    }
    return null;
  }
  if (!value || typeof value !== "object") return null;
  if (value.method === method) return value;
  for (const item of Object.values(value)) {
    const found = findRequest(item, method);
    if (found) return found;
  }
  return null;
}

function collectRecordFields(value, fields = new Set()) {
  if (Array.isArray(value)) {
    value.forEach((item) => collectRecordFields(item, fields));
  } else if (value && typeof value === "object") {
    if (Array.isArray(value.s) && Array.isArray(value.d)) {
      value.s.forEach((field) => field?.n && fields.add(field.n));
    }
    Object.values(value).forEach((item) => collectRecordFields(item, fields));
  }
  return fields;
}

const har = JSON.parse(fs.readFileSync(harPath, "utf8"));
let requestBody = null;
let calledMethodHeader = methodName;
for (const entry of har.log.entries) {
  const text = entry.request?.postData?.text;
  if (!text) continue;
  try {
    requestBody = findRequest(JSON.parse(text), methodName);
  } catch {
    // Ignore non-JSON telemetry payloads.
  }
  if (requestBody) {
    calledMethodHeader = entry.request.headers?.find(
      (header) => header.name.toLowerCase() === "x-calledmethod"
    )?.value || methodName;
    break;
  }
}
if (!requestBody) throw new Error(`Method not found in HAR: ${methodName}`);

if (requestedPageSize && requestBody.params?.["Навигация"]?.s && requestBody.params?.["Навигация"]?.d) {
  const navigation = requestBody.params["Навигация"];
  const sizeIndex = navigation.s.findIndex((field) => field.n === "РазмерСтраницы");
  const pageIndex = navigation.s.findIndex((field) => field.n === "Страница");
  if (sizeIndex >= 0) navigation.d[sizeIndex] = Number(requestedPageSize);
  if (pageIndex >= 0) navigation.d[pageIndex] = 0;
}

const env = { ...loadEnv(path.join(revenueProjectPath, ".env")), ...process.env };
const keyFile = fs.readdirSync(revenueProjectPath).find((name) => name.endsWith(".key"));
const serviceKey = env.SABY_SERVICE_KEY || (keyFile
  ? fs.readFileSync(path.join(revenueProjectPath, keyFile), "utf8").trim()
  : "");
if (!env.SABY_APP_CLIENT_ID || !env.SABY_APP_SECRET || !serviceKey) {
  throw new Error("Saby OAuth credentials are incomplete");
}

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

const rpcResponse = await fetch("https://online.saby.ru/service/", {
  method: "POST",
  headers: {
    "Content-Type": "application/json; charset=utf-8",
    "X-SBISAccessToken": auth.token,
    "X-CalledMethod": calledMethodHeader,
    "X-Requested-With": "XMLHttpRequest",
  },
  body: JSON.stringify(requestBody),
});
const responseText = await rpcResponse.text();
let data;
try {
  data = JSON.parse(responseText);
} catch {
  data = null;
}

const rowCount = Array.isArray(data?.result?.d) ? data.result.d.length : null;
const summary = {
  method: methodName,
  httpStatus: rpcResponse.status,
  ok: rpcResponse.ok && !data?.error,
  rowCount,
  fields: [...collectRecordFields(data)].sort(),
  errorCode: data?.error?.code ?? null,
  errorMessage: data?.error?.message ?? (data ? null : "Non-JSON response"),
};
console.log(JSON.stringify(summary, null, 2));
