import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(fileURLToPath(import.meta.url));
const codePath = path.join(root, "Code.gs");
const htmlPath = path.join(root, "Index.html");
const manifestPath = path.join(root, "appsscript.json");

function readText(filePath) {
  return fs.readFileSync(filePath, "utf8");
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

const manifest = JSON.parse(readText(manifestPath));
assert(manifest.timeZone === "Asia/Tokyo", "appsscript.json timeZone must be Asia/Tokyo.");
assert(manifest.runtimeVersion === "V8", "appsscript.json runtimeVersion must be V8.");
assert(manifest.webapp?.executeAs === "USER_DEPLOYING", "Web app must execute as the deploying user.");
assert(manifest.webapp?.access === "MYSELF", "Web app access must be MYSELF.");
assert(!String(manifest.webapp?.access || "").includes("ANYONE"), "Web app must not be deployed for anyone access.");

const scopes = new Set(manifest.oauthScopes || []);
for (const requiredScope of [
  "https://www.googleapis.com/auth/spreadsheets",
  "https://www.googleapis.com/auth/drive",
  "https://www.googleapis.com/auth/gmail.compose",
  "https://www.googleapis.com/auth/gmail.send",
]) {
  assert(scopes.has(requiredScope), `Missing OAuth scope: ${requiredScope}`);
}

assert(
  (manifest.dependencies?.enabledAdvancedServices || []).some(
    (service) => service.userSymbol === "Drive" && service.serviceId === "drive",
  ),
  "Advanced Drive service must be enabled to convert Excel templates.",
);

const code = readText(codePath);
new Function(code);
const cloudFunctions = new Function(`${code}; rows_ = () => []; return { daysInMonth_, rowsForReport_, previewRows_ };`)();
// 個人情報はコードに直書きせず Script Properties(APP_CONFIG) から cfg_() 経由で読み込む。
assert(code.includes("function cfg_()"), "Code.gs must load config from Script Properties via cfg_().");
assert(code.includes("getProperty('APP_CONFIG')"), "Code.gs must read APP_CONFIG from Script Properties.");
// ブロックリストは gitignore 済みの実設定(app-config.json)から動的に取得する。
// これにより validate.mjs 自体には個人情報を持たせない。設定が無い環境では検査をスキップ。
const appConfigFile = path.join(root, "..", "data", "app-config.json");
if (fs.existsSync(appConfigFile)) {
  const c = JSON.parse(readText(appConfigFile));
  const sr = c.salesReport || {};
  const piiValues = [
    c.allowedEmail, c.userId, c.userName, c.recipient, c.recipientName,
    c.commuteRoute, c.commuteCost, c.salesFileName, c.clientFilePerson,
    ...(c.ccRecipients || []),
    sr.managementNumber, sr.employeeName,
    ...Object.values(sr.contract || {}),
  ].filter((value) => typeof value === "string" && value.length >= 2);
  for (const pii of piiValues) {
    assert(!code.includes(pii), `Code.gs must not contain hard-coded personal information: ${pii}`);
  }
}
assert(code.includes("function assertAccess_()"), "Code.gs must include assertAccess_().");
assert(
  code.includes("XFrameOptionsMode.DEFAULT") && !code.includes("XFrameOptionsMode.ALLOWALL"),
  "Code.gs must use DEFAULT X-Frame options (SAMEORIGIN behavior, no ALLOWALL) to prevent clickjacking.",
);
assert(code.includes("LockService.getScriptLock()"), "Code.gs must use LockService for writes.");
assert(code.includes("for (let index = 0; index < 31; index += 1)"), "Report generation must process all 31 template rows.");
assert(code.includes("clearContent()"), "Report generation must clear unused template rows for short months.");
assert(!code.includes("rows.slice(0, 31).forEach"), "Report generation must not skip unused template rows in short months.");
assert(cloudFunctions.daysInMonth_("2026-02").length === 28, "February 2026 must generate 28 days.");
assert(cloudFunctions.daysInMonth_("2028-02").length === 29, "Leap-year February must generate 29 days.");
assert(cloudFunctions.daysInMonth_("2026-04").length === 30, "April must generate 30 days.");
assert(cloudFunctions.daysInMonth_("2026-05").length === 31, "May must generate 31 days.");
assert(cloudFunctions.rowsForReport_("2026-02", false).length === 28, "February report source rows must have 28 days.");
assert(cloudFunctions.previewRows_("2026-02", "sales").length === 28, "February preview must have 28 rows.");
assert(code.includes("function assertDriveFileId_(value)"), "Code.gs must validate template file IDs.");
assert(code.includes("function assertTemplateFileId_(value)"), "Code.gs must validate Excel/Sheets template file IDs.");
assert(code.includes("function copyTemplateAsSpreadsheet_(templateId, name, folder)"), "Code.gs must convert Excel templates before editing them.");
assert(code.includes("case 'listTemplateFiles'"), "Code.gs must expose a template file picker API.");
assert(code.includes("function listTemplateFiles_()"), "Code.gs must list Excel template candidates.");

const html = readText(htmlPath);
assert(html.includes("function formatMonthLabel"), "Index.html must format the selected month for display.");
assert(html.includes("対象年月"), "Index.html must show the current target month.");
assert(html.includes('<div class="target-month"><span>対象年月</span><input type="month" data-field="month"'), "Index.html must let users select the target month in the header.");
assert(!html.includes("<label>対象年月<input type=\"month\" data-field=\"month\""), "Index.html must not show the target month picker inside the monthly report card.");
assert(html.includes("画面表示でエラーが発生しました"), "Index.html must show render errors instead of a blank screen.");
assert(html.includes("対象: ${escapeHtml(targetMonthLabel)}分"), "Index.html must show the month-end confirmation target.");
assert(html.includes("対象: ${escapeHtml(targetMonthLabel)}分の作業報告書"), "Index.html must show the Gmail target month and file type.");
assert(!html.includes("月末確認（${escapeHtml(targetMonthLabel)}分）"), "Index.html must not repeat the target month in the month-end heading.");
assert(!html.includes("${escapeHtml(targetMonthLabel)}分のGmail下書きを作成"), "Index.html must not repeat the target month in the Gmail draft button.");
assert(html.includes("openTemplatePicker"), "Index.html must open a Drive-backed template picker from the template ID fields.");
assert(html.includes("chooseTemplateFile"), "Index.html must copy the chosen Drive file ID into the template ID field.");
assert(html.includes("picker-panel"), "Index.html must show Drive candidates in an on-demand picker panel.");
assert(html.includes("Excel形式の候補"), "Index.html must describe Excel template candidates.");
const scriptMatch = html.match(/<script>([\s\S]*)<\/script>/);
assert(scriptMatch, "Index.html must contain a script block.");
const bootstrap = {
  user: { email: "sender@example.com", displayName: "氏名", recipient: "recipient@example.com", ccRecipients: [] },
  settings: { rootFolderName: "勤怠管理", salesTemplateId: "", clientTemplateId: "" },
  attendance: { days: [], reports: [], gmailDrafts: [], monthlyConfirmation: {} },
  month: "2026-06",
  today: "2026-06-29",
};
const browserGlobals = "const google={script:{run:{withSuccessHandler(){return this},withFailureHandler(){return this},api(){}}}}; const document={getElementById(){return {innerHTML:'',addEventListener(){}}},querySelectorAll(){return []}}; const confirm=()=>false;";
new Function(`${browserGlobals}\n${scriptMatch[1].replace("<?!= bootstrapJson ?>", JSON.stringify(bootstrap))}`);
new Function(`${browserGlobals}\n${scriptMatch[1].replace("<?!= bootstrapJson ?>", "{}")}`);

console.log("cloud-apps-script validation passed");
