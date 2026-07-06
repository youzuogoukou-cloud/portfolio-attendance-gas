// 非個人情報の定数のみを保持する。氏名・メールアドレス・宛先・派遣契約情報などの
// 個人情報はソースに直書きせず、Script Properties の `APP_CONFIG`(JSON文字列) から
// cfg_() 経由で読み込む。セットアップ手順とキー構造は config.example.gs を参照。
const APP = {
  rootFolderName: '勤怠管理',
  storeSpreadsheetName: '勤怠管理データ',
  categories: ['①', '②', '③', '休日', '休暇'],
  defaultTimes: { startTime: '09:00', endTime: '17:40', breakMinutes: 60 },
  mimeXlsx: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  mimeGoogleSheets: 'application/vnd.google-apps.spreadsheet',
};

// 個人情報を含む設定は Script Properties から遅延ロード（1実行内でメモ化）。
let __appConfig = null;
function cfg_() {
  if (__appConfig) return __appConfig;
  const raw = PropertiesService.getScriptProperties().getProperty('APP_CONFIG');
  if (!raw) {
    throw new Error('APP_CONFIG が未設定です。config.example.gs を config.local.gs にコピーして実値を入力し、installAppConfig() を一度だけ実行してください。');
  }
  __appConfig = JSON.parse(raw);
  return __appConfig;
}

const SHEETS = {
  attendance_days: ['id', 'user_id', 'work_date', 'work_type', 'category', 'start_time', 'end_time', 'break_minutes', 'note', 'source', 'created_at', 'updated_at', 'holiday_name'],
  monthly_confirmations: ['id', 'user_id', 'month', 'attendance_reviewed', 'report_preview_reviewed', 'sales_attachment_confirmed', 'confirmed_at', 'confirmed_by'],
  reports: ['id', 'user_id', 'month', 'target_type', 'status', 'drive_file_id', 'xlsx_file_id', 'file_name', 'content_type', 'generated_at', 'updated_at'],
  gmail_drafts: ['id', 'user_id', 'month', 'gmail_draft_id', 'sales_report_id', 'recipient', 'subject', 'cc_recipients', 'sender_email', 'status', 'gmail_message_id', 'created_at', 'sent_at'],
  settings: ['key', 'value', 'updated_at'],
  audit_events: ['id', 'event_type', 'resource_type', 'resource_id', 'payload_json', 'created_at'],
};

const HOLIDAYS = {
  '2026-01-01': '元日',
  '2026-01-12': '成人の日',
  '2026-02-11': '建国記念の日',
  '2026-02-23': '天皇誕生日',
  '2026-03-20': '春分の日',
  '2026-04-29': '昭和の日',
  '2026-05-03': '憲法記念日',
  '2026-05-04': 'みどりの日',
  '2026-05-05': 'こどもの日',
  '2026-05-06': '振替休日',
  '2026-07-20': '海の日',
  '2026-08-11': '山の日',
  '2026-09-21': '敬老の日',
  '2026-09-22': '国民の休日',
  '2026-09-23': '秋分の日',
  '2026-10-12': 'スポーツの日',
  '2026-11-03': '文化の日',
  '2026-11-23': '勤労感謝の日',
};

function doGet() {
  assertAccess_();
  const template = HtmlService.createTemplateFromFile('Index');
  template.bootstrapJson = JSON.stringify(getBootstrap_());
  return template.evaluate()
    .setTitle('勤怠管理クラウド')
    // クリックジャッキング防止。DEFAULT は X-Frame-Options: SAMEORIGIN を出し、
    // 同一オリジン以外からのiframe埋め込みを禁止する（ALLOWALL の逆）。
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function api(action, payload) {
  assertAccess_();
  const lock = LockService.getScriptLock();
  lock.waitLock(20 * 1000);
  try {
    ensureStore_();
    switch (action) {
      case 'bootstrap':
        return getBootstrap_();
      case 'monthState':
        return { attendance: monthState_(assertMonth_(payload.month || currentMonth_())) };
      case 'dayState':
        return { day: dayState_((payload && payload.date) || '') };
      case 'saveSettings':
        return saveSettings_(payload || {});
      case 'listTemplateFiles':
        return { templateFiles: listTemplateFiles_() };
      case 'saveAttendance':
        return saveAttendance_(payload || {});
      case 'reflectToday':
        return reflectToday_(payload || {});
      case 'generateReports':
        return generateReports_(payload || {});
      case 'loadPreview': {
        const previewMonth = String(payload.month || currentMonth_());
        const previewTarget = String(payload.targetType || 'sales');
        // 実ファイルが存在する対象だけプレビュー内容を返す。無い場合は空＋未作成フラグを返し、
        // UI 側で「まだ作成されていません」を表示する(データ由来の“作られる予定”は出さない)。
        const reportExists = liveReports_(previewMonth).some((r) => r.targetType === previewTarget);
        return {
          rows: reportExists ? previewRows_(previewMonth, previewTarget) : [],
          reportExists,
          previewTarget,
          previewMonth,
        };
      }
      case 'confirmMonth':
        return confirmMonth_(payload || {});
      case 'clearMonthConfirmation':
        return clearMonthConfirmation_(payload || {});
      case 'createGmailDraft':
        return createGmailDraft_(payload || {});
      case 'sendGmailDraft':
        return sendGmailDraft_(payload || {});
      case 'syncDraftStatus':
        return syncDraftStatus_(payload || {});
      default:
        throw new Error('Unknown action: ' + action);
    }
  } finally {
    lock.releaseLock();
  }
}

function assertAccess_() {
  const allowed = String(cfg_().allowedEmail || '').toLowerCase();
  const email = String(Session.getActiveUser().getEmail() || '').toLowerCase();
  // メールが取得できない場合(空文字)も拒否する。空文字素通りは認可バイパスになるため。
  if (email !== allowed) {
    throw new Error('このアプリは ' + cfg_().allowedEmail + ' 専用です。現在のアカウント: ' + (email || '(不明)'));
  }
}

function getBootstrap_() {
  ensureStore_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const month = today.slice(0, 7);
  return {
    user: {
      email: cfg_().allowedEmail,
      displayName: cfg_().userName,
      recipient: cfg_().recipient,
      ccRecipients: cfg_().ccRecipients,
      recipientName: cfg_().recipientName,
      commuteRoute: cfg_().commuteRoute,
      commuteCost: cfg_().commuteCost,
    },
    today,
    todayDay: dayState_(today),
    month,
    settings: settings_(),
    attendance: monthState_(month),
  };
}

function ensureStore_(skipSeed) {
  const props = PropertiesService.getScriptProperties();
  let spreadsheetId = props.getProperty('STORE_SPREADSHEET_ID');
  let ss;
  let created = false;
  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create(APP.storeSpreadsheetName);
    spreadsheetId = ss.getId();
    props.setProperty('STORE_SPREADSHEET_ID', spreadsheetId);
    created = true;
  }
  // ストアのTZをスクリプトTZ(Asia/Tokyo)へ固定する。未設定だと '09:00' 等の時刻セルを
  // getValues で Date として読む際に作成アカウント既定TZで解釈され、formatDate(スクリプトTZ)で
  // 時差ぶんズレる(例: 9時→6時)。アプリ表示・Excel生成の双方をこの値で統一する。
  try {
    if (ss.getSpreadsheetTimeZone() !== Session.getScriptTimeZone()) {
      ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
    }
  } catch (err) {
    // TZ設定に失敗しても処理は続行する
  }
  Object.keys(SHEETS).forEach((name) => ensureSheet_(ss, name, SHEETS[name]));
  if (!skipSeed) seedSettings_();
  if (created) moveFileToFolder_(DriveApp.getFileById(spreadsheetId), rootFolder_());
  return ss;
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  const current = sheet.getRange(1, 1, 1, headers.length).getValues()[0];
  if (headers.some((header, index) => current[index] !== header)) {
    sheet.clear();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function seedSettings_() {
  const defaults = {
    rootFolderName: APP.rootFolderName,
    salesTemplateId: '',
    clientTemplateId: '',
  };
  const existing = settings_();
  Object.keys(defaults).forEach((key) => {
    if (!(key in existing)) upsertRow_('settings', { key, value: defaults[key], updated_at: now_() }, ['key']);
  });
}

function settings_() {
  const rows = rows_('settings');
  return rows.reduce((acc, row) => {
    acc[row.key] = row.value;
    return acc;
  }, {});
}

function saveSettings_(payload) {
  const allowed = ['rootFolderName', 'salesTemplateId', 'clientTemplateId'];
  allowed.forEach((key) => {
    if (key in payload) upsertRow_('settings', { key, value: sanitizeSettingValue_(key, payload[key]), updated_at: now_() }, ['key']);
  });
  addAudit_('settings_saved', 'settings', 'settings', {});
  return { settings: settings_(), attendance: monthState_(String(payload.month || currentMonth_())) };
}

function sanitizeSettingValue_(key, value) {
  const text = String(value || '').trim();
  if (key === 'salesTemplateId' || key === 'clientTemplateId') return text ? assertTemplateFileId_(text) : '';
  return text;
}

function listTemplateFiles_() {
  const files = DriveApp.searchFiles("mimeType = '" + APP.mimeXlsx + "' and trashed = false");
  const results = [];
  while (files.hasNext() && results.length < 100) {
    const file = files.next();
    results.push({
      id: file.getId(),
      name: file.getName(),
      mimeType: file.getMimeType(),
      updatedAt: file.getLastUpdated().toISOString(),
    });
  }
  results.sort((a, b) => String(b.updatedAt).localeCompare(String(a.updatedAt)));
  return results.slice(0, 50);
}

function monthState_(month) {
  return {
    month,
    days: rows_('attendance_days').filter((row) => String(row.work_date || '').startsWith(month)).map(publicDay_),
    monthlyConfirmation: confirmation_(month),
    reports: liveReports_(month),
    gmailDrafts: rows_('gmail_drafts').filter((row) => row.user_id === cfg_().userId && row.month === month).map(publicDraft_),
  };
}

// 対象月の帳票を「作成済み」として返す。reports シートの行だけでは、実ファイルを削除/ゴミ箱に
// してもレコードが残り誤って作成済み判定になる。対象種別ごとに最新1件へ集約し、実際に Drive に
// ファイルが存在するものだけを返す(=月次作成ボタンの可否が実ファイルと一致する)。
function liveReports_(month) {
  const rows = rows_('reports').filter((row) => row.user_id === cfg_().userId && row.month === month);
  const latest = {};
  rows.forEach((row) => {
    const prev = latest[row.target_type];
    if (!prev || String(row.generated_at || '') > String(prev.generated_at || '')) latest[row.target_type] = row;
  });
  return Object.keys(latest)
    .map((t) => latest[t])
    .filter((row) => driveFileExists_(row.xlsx_file_id))
    .map(publicReport_);
}

function driveFileExists_(fileId) {
  if (!fileId) return false;
  try {
    return !DriveApp.getFileById(fileId).isTrashed();
  } catch (err) {
    return false; // 見つからない/権限無し等は「存在しない」扱い
  }
}

function saveAttendance_(payload) {
  const workDate = assertDate_(payload.date || payload.workDate);
  const category = assertCategory_(payload.category || '①');
  const isOff = category === '休日' || category === '休暇';
  const row = {
    id: existingAttendanceId_(workDate) || id_('att'),
    user_id: cfg_().userId,
    work_date: workDate,
    work_type: isOff ? 'holiday' : 'workday',
    category,
    start_time: isOff ? '' : assertTime_(payload.startTime || APP.defaultTimes.startTime),
    end_time: isOff ? '' : assertTime_(payload.endTime || APP.defaultTimes.endTime),
    break_minutes: isOff ? 0 : assertBreakMinutes_(payload.breakMinutes),
    note: String(payload.note || ''),
    source: 'manual',
    created_at: now_(),
    updated_at: now_(),
    holiday_name: HOLIDAYS[workDate] || '',
  };
  upsertRow_('attendance_days', row, ['user_id', 'work_date']);
  addAudit_('attendance_saved', 'attendance_day', workDate, { workDate, category });
  return { day: publicDay_(row), attendance: monthState_(workDate.slice(0, 7)) };
}

// 指定日の勤怠を「営業帳票に反映される値」でそのまま返す。フォーム表示をこの1ソースに統一し、
// クライアント側の古いキャッシュに依存せずプレビュー/Excelと必ず一致させる。
// 保存済みならその値、未保存なら既定値(平日9:00-17:40 / 土日祝は休日)を返す。
function dayState_(date) {
  const d = assertDate_(date);
  const rows = rowsForReport_(d.slice(0, 7), true);
  const found = rows.filter((row) => row.workDate === d)[0];
  return found || {
    workDate: d,
    workType: 'workday',
    category: '①',
    startTime: APP.defaultTimes.startTime,
    endTime: APP.defaultTimes.endTime,
    breakMinutes: APP.defaultTimes.breakMinutes,
    note: '',
  };
}

function reflectToday_(payload) {
  const saved = saveAttendance_(payload);
  const month = saved.day.workDate.slice(0, 7);
  // 月次作成と同じくモード連動(既定=営業用のみ)。営業用は xlsx を直接書き換える方式で
  // スプレッドシート変換を挟まないため速く確実。両方を固定生成していた旧実装は、客先用の
  // 変換失敗やテンプレ未設定で反映ごと失敗し得たため、モードで対象を絞る。
  const reports = generateReports_({ month, mode: payload.mode || 'salesOnly', includeSavedAttendance: true }).reports;
  return { day: saved.day, reports, attendance: monthState_(month) };
}

function generateReports_(payload) {
  const month = assertMonth_(payload.month || currentMonth_());
  const targetTypes = Array.isArray(payload.targetTypes) && payload.targetTypes.length
    ? payload.targetTypes
    : targetTypesForMode_(payload.mode || 'salesOnly');
  const reports = targetTypes.map((targetType) => createReport_(month, assertTargetType_(targetType), Boolean(payload.includeSavedAttendance)));
  return { reports, attendance: monthState_(month) };
}

function createReport_(month, targetType, includeSavedAttendance) {
  const config = settings_();
  const templateId = targetType === 'sales' ? config.salesTemplateId : config.clientTemplateId;
  if (!templateId) throw new Error((targetType === 'sales' ? '営業用' : '客先用') + 'テンプレートIDを設定してください。');
  const rows = rowsForReport_(month, includeSavedAttendance);
  const parts = month.split('-');
  const monthFolderName = parts[0] + '年' + parts[1] + '月';
  const folder = childFolder_(childFolder_(rootFolder_(), reportKindLabel_(targetType)), monthFolderName);
  const fileName = reportFileName_(month, targetType);
  const oldFiles = folder.getFilesByName(fileName.replace(/\.xlsx$/, ''));
  while (oldFiles.hasNext()) oldFiles.next().setTrashed(true);
  const existingXlsx = folder.getFilesByName(fileName);
  while (existingXlsx.hasNext()) existingXlsx.next().setTrashed(true);
  let xlsx;
  if (targetType === 'sales') {
    // 営業用: Sheets変換を行わず xlsx を直接編集する。元テンプレの数式・書式・外部参照の
    // キャッシュ値を完全に保持し、入力セル(就業期間D8/H8 と 各日の 区分D/始業E/終業F/休憩G/備考N)
    // だけを書き換える。これで変換由来の #ERROR や数式の作り替えが起きず、元テンプレに忠実。
    xlsx = folder.createFile(buildSalesXlsx_(templateId, rows, month, fileName));
  } else {
    // 客先用(出勤簿): 定義名参照などがあり、従来どおり Sheets 変換方式で記入する。
    let copy;
    try {
      copy = copyTemplateAsSpreadsheet_(templateId, fileName.replace(/\.xlsx$/, ''), folder);
    } catch (err) {
      throw new Error('客先用テンプレートを準備できませんでした。' + (err && err.message ? err.message : err));
    }
    const ss = SpreadsheetApp.openById(copy.getId());
    // xlsxから変換した直後のシートは既定TZになり日付が-1日ずれるためスクリプトのTZ(日本時間)へ揃える。
    try {
      ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
      ss.setSpreadsheetLocale('ja_JP'); // TEXT(日付,"aaa")=曜日 が効くよう日本語ロケールにする
    } catch (err) {
      // TZ/ロケール設定に失敗しても生成は続行する
    }
    fillClientReport_(ss, rows, month);
    SpreadsheetApp.flush();
    xlsx = folder.createFile(exportSpreadsheetAsXlsx_(copy.getId(), fileName));
    try {
      copy.setTrashed(true); // 中間スプレッドシートは残さない
    } catch (err) {
      // 中間スプレッドシートの削除に失敗しても生成は成功扱いとする
    }
  }
  const record = {
    id: id_('report'),
    user_id: cfg_().userId,
    month,
    target_type: targetType,
    status: 'ready',
    drive_file_id: '',
    xlsx_file_id: xlsx.getId(),
    file_name: fileName,
    content_type: APP.mimeXlsx,
    generated_at: now_(),
    updated_at: now_(),
  };
  upsertRow_('reports', record, ['user_id', 'month', 'target_type']);
  addAudit_('report_generated', 'report', record.id, { month, targetType, includeSavedAttendance });
  return publicReport_(record);
}

// ===== 営業用: xlsx を直接編集（Sheets変換なし）=====
// テンプレの xlsx(zip) を展開し、報告書ワークシートの入力セルだけ書き換えて再圧縮する。
// 数式・書式・外部参照のキャッシュ値はそのまま残るため、元テンプレに忠実な出力になる。
function buildSalesXlsx_(templateId, rows, month, fileName) {
  const blob = DriveApp.getFileById(templateId).getBlob().setContentType('application/zip');
  const parts = Utilities.unzip(blob);
  const worksheetPath = salesWorksheetPath_(parts);
  const out = parts.map((part) => {
    const name = part.getName();
    if (name === worksheetPath) {
      const xml = updateSalesWorksheetXml_(part.getDataAsString('UTF-8'), rows, month);
      return Utilities.newBlob(xml, 'application/xml', name);
    }
    if (name === 'xl/workbook.xml') {
      // 開いた時に全再計算させる。日付B列(=IF(D8..))等が就業期間D8/H8から再計算され、
      // テンプレに残る旧月のキャッシュ値が当月へ更新される。契約欄など外部参照は
      // externalLink のキャッシュ値から再計算されるため #ERROR にならない。
      return Utilities.newBlob(forceFullCalcOnLoad_(part.getDataAsString('UTF-8')), 'application/xml', name);
    }
    return part;
  });
  return Utilities.zip(out, fileName).setContentType(APP.mimeXlsx);
}

// workbook.xml の calcPr に fullCalcOnLoad="1" を付け、Excel が開いた時に全再計算するようにする。
function forceFullCalcOnLoad_(workbookXml) {
  if (/<calcPr\b[^>]*\bfullCalcOnLoad=/.test(workbookXml)) return workbookXml;
  if (/<calcPr\b[^>]*\/>/.test(workbookXml)) {
    return workbookXml.replace(/<calcPr\b([^>]*)\/>/, '<calcPr$1 fullCalcOnLoad="1"/>');
  }
  return workbookXml.replace('</workbook>', '<calcPr calcId="191029" fullCalcOnLoad="1"/></workbook>');
}

// workbook.xml と rels から報告書シート(様式J)のワークシートXMLのパスを解決する。
function salesWorksheetPath_(parts) {
  const byName = {};
  parts.forEach((part) => { byName[part.getName()] = part; });
  const workbook = byName['xl/workbook.xml'] ? byName['xl/workbook.xml'].getDataAsString('UTF-8') : '';
  const rels = byName['xl/_rels/workbook.xml.rels'] ? byName['xl/_rels/workbook.xml.rels'].getDataAsString('UTF-8') : '';
  const sheetTags = (workbook.match(/<sheet\b[^>]*\/>/g) || []).map(xmlAttrs_);
  let target = null;
  for (let i = 0; i < sheetTags.length; i += 1) {
    if (sheetTags[i].name === '派遣就業状況報告書(様式J)') { target = sheetTags[i]; break; }
  }
  if (!target) target = sheetTags[0];
  if (!target || !target['r:id']) throw new Error('テンプレートのワークシートが見つかりません。');
  const relTags = (rels.match(/<Relationship\b[^>]*\/>/g) || []).map(xmlAttrs_);
  let rel = null;
  for (let j = 0; j < relTags.length; j += 1) {
    if (relTags[j].Id === target['r:id']) { rel = relTags[j]; break; }
  }
  if (!rel || !rel.Target) throw new Error('ワークシート参照が見つかりません。');
  return 'xl/' + String(rel.Target).replace(/^\/?xl\//, '');
}

// 報告書ワークシートXMLの入力セルだけ書き換える（数式・他セル・外部参照は温存）。
// 日付B列・曜日C列・実労働などはテンプレ数式が就業期間(D8/H8)から自動計算する。
// 数式セル(<c>..<f>..</f><v>..</v>..</c>)の <f>(共有数式含む)・s・t 属性を温存したまま、
// キャッシュ値 <v> だけを差し替える。プレビュー/保護ビューは <v> を表示するのでここが当月になる。
// value が null/空 のときは <v> を除去する(プレビュー空欄・Excelは開いた時に再計算)。
function setFormulaCachedValue_(sheetXml, cellRef, value) {
  const cellPattern = new RegExp('<c\\b(?=[^>]*\\br="' + cellRef + '")[^>]*>[\\s\\S]*?<\\/c>');
  const match = sheetXml.match(cellPattern);
  if (!match) return sheetXml; // セルが無ければ何もしない
  const cell = match[0];
  let open = cell.match(/^<c\b[^>]*>/)[0];
  let inner = cell.slice(open.length, cell.length - 4); // 末尾の </c> を除く
  const blank = value === null || value === undefined || value === '';
  const isNumber = typeof value === 'number';
  // セル型を値に合わせる: 数値なら t 属性を除去(=数値型)、文字列なら t="str" を付与。
  // これをしないと、テンプレで空欄(数式結果="")だった日のセルが t="str" のまま数値を持ち、
  // プレビューが数値を文字列扱いして書式(h:mm や小数)が効かず生の小数で表示される。
  if (isNumber) {
    open = open.replace(/\s+t="[^"]*"/, '');
  } else if (!blank) {
    open = /\st="[^"]*"/.test(open) ? open.replace(/\st="[^"]*"/, ' t="str"') : open.replace(/^<c\b/, '<c t="str"');
  }
  const newV = blank ? '' : '<v>' + xlsxEscape_(isNumber ? String(value) : value) + '</v>';
  const vPattern = /<v\s*\/>|<v>[\s\S]*?<\/v>/;
  if (vPattern.test(inner)) inner = inner.replace(vPattern, newV);
  else if (!blank) inner += newV; // <v> が無ければ末尾(=</f>の後)に追加
  return sheetXml.replace(cellPattern, open + inner + '</c>');
}

// 指定セルのキャッシュ数値(<v>)を読む。丸め単位 S6/T6(契約の外部参照由来)の取得に使う。
function sheetCachedNumber_(sheetXml, cellRef) {
  const m = sheetXml.match(new RegExp('<c\\b(?=[^>]*\\br="' + cellRef + '")[^>]*>[\\s\\S]*?<\\/c>'));
  if (!m) return null;
  const vm = m[0].match(/<v>([\s\S]*?)<\/v>/);
  if (!vm) return null;
  const n = Number(vm[1]);
  return isFinite(n) ? n : null;
}

function updateSalesWorksheetXml_(sheetXml, rows, month) {
  let xml = sheetXml;
  const firstDate = month + '-01';
  const parts = month.split('-').map(Number);
  const lastRow = rows[rows.length - 1];
  const lastDate = (lastRow && lastRow.workDate) || firstDate;
  // 入力セル(数式なし): 就業期間 D8/H8
  xml = xlsxSetCell_(xml, 'D8', excelDateSerial_(firstDate)); // 就業期間 開始
  xml = xlsxSetCell_(xml, 'H8', excelDateSerial_(lastDate));  // 就業期間 終了
  // 実労働の丸め単位(=S6/T6 のキャッシュ値, 契約の外部参照 [2]入力用 由来)。取れなければ10分。
  const unit = sheetCachedNumber_(xml, 'S6') || (10 / 1440);
  const totalUnitHours = (sheetCachedNumber_(xml, 'T6') || (10 / 1440)) * 24;
  // 数式は残し、キャッシュ値 <v> だけを当月へ更新(プレビュー表示用)。年A1・月E1/A11。
  xml = setFormulaCachedValue_(xml, 'A1', parts[0]);
  xml = setFormulaCachedValue_(xml, 'E1', parts[1]);
  xml = setFormulaCachedValue_(xml, 'A11', parts[1]); // =MONTH(D8) 日別表の月見出し
  let sumI = 0;
  let sumJ = 0;
  for (let index = 0; index < 31; index += 1) {
    const row = rows[index];
    const r = 11 + index;
    if (!row) {
      // 該当日なし: 日付/曜日/実労働のキャッシュを空に、入力欄はクリア
      xml = setFormulaCachedValue_(xml, 'B' + r, null);
      xml = setFormulaCachedValue_(xml, 'C' + r, null);
      xml = setFormulaCachedValue_(xml, 'I' + r, null);
      xml = setFormulaCachedValue_(xml, 'J' + r, null);
      xml = xlsxSetCell_(xml, 'D' + r, null);
      xml = xlsxSetCell_(xml, 'E' + r, null);
      xml = xlsxSetCell_(xml, 'F' + r, null);
      xml = xlsxSetCell_(xml, 'G' + r, null);
      xml = xlsxSetCell_(xml, 'N' + r, null);
      continue;
    }
    const holiday = row.workType === 'holiday';
    const start = holiday ? null : timeToExcelSerial_(row.startTime);
    const end = holiday ? null : timeToExcelSerial_(row.endTime);
    const breakFraction = (holiday || !row.breakMinutes) ? 0 : row.breakMinutes / 1440;
    // 入力セル(数式なし)を値で記入: D区分/E始業/F終業/G休憩/N備考
    xml = xlsxSetCell_(xml, 'D' + r, holiday ? '休日' : (row.category || '①'));
    xml = xlsxSetCell_(xml, 'E' + r, start);
    xml = xlsxSetCell_(xml, 'F' + r, end);
    xml = xlsxSetCell_(xml, 'G' + r, (holiday || !row.breakMinutes) ? null : breakFraction);
    xml = xlsxSetCell_(xml, 'N' + r, holiday ? '' : (row.note || ''));
    // 数式セルのキャッシュ値 <v> を更新: B日付/C曜日
    xml = setFormulaCachedValue_(xml, 'B' + r, excelDateSerial_(row.workDate));
    xml = setFormulaCachedValue_(xml, 'C' + r, salesWeekday_(row.workDate));
    // 実労働 I=(終業-始業-休憩), J=FLOOR(I,unit)*24。E/F未入力(休日等)は空。
    if (start !== null && end !== null) {
      const workFraction = end - start - breakFraction;
      const jHours = Math.floor(workFraction / unit + 1e-9) * unit * 24;
      sumI += workFraction;
      sumJ += jHours;
      xml = setFormulaCachedValue_(xml, 'I' + r, workFraction);
      xml = setFormulaCachedValue_(xml, 'J' + r, jHours);
    } else {
      xml = setFormulaCachedValue_(xml, 'I' + r, null);
      xml = setFormulaCachedValue_(xml, 'J' + r, null);
    }
  }
  // 合計セルのキャッシュ値を更新: I6=ΣI, J6=FLOOR(ΣJ, T6*24)。K6/L6/M6("-")は不変。
  xml = setFormulaCachedValue_(xml, 'I6', sumI);
  xml = setFormulaCachedValue_(xml, 'J6', sumJ === 0 ? 0 : Math.floor(sumJ / totalUnitHours + 1e-9) * totalUnitHours);
  return xml;
}

// 曜日ラベル(TEXT(日付,"aaa") 相当)。UTC基準で計算しタイムゾーンずれを避ける。
function salesWeekday_(workDate) {
  const labels = ['日', '月', '火', '水', '木', '金', '土'];
  const p = String(workDate).split('-').map(Number);
  return labels[new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay()];
}

// --- xlsx セル操作ヘルパ（server.mjs の実装を移植）---
function excelDateSerial_(date) {
  const p = String(date).split('-').map(Number);
  return Math.round((Date.UTC(p[0], p[1] - 1, p[2]) - Date.UTC(1899, 11, 30)) / 86400000);
}

function timeToExcelSerial_(time) {
  if (!time) return null;
  const p = String(time).split(':').map(Number);
  if (!isFinite(p[0]) || !isFinite(p[1])) return null; // パース不能な時刻は NaN を書かず空欄にする
  return (p[0] * 60 + p[1]) / 1440;
}

function xlsxEscape_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function decodeXml_(value) {
  return String(value == null ? '' : value)
    .replace(/&quot;/g, '"').replace(/&gt;/g, '>').replace(/&lt;/g, '<').replace(/&amp;/g, '&');
}

function xmlAttrs_(tag) {
  const attrs = {};
  const re = /\s([\w:]+)="([^"]*)"/g;
  let m = re.exec(String(tag));
  while (m !== null) {
    attrs[m[1]] = decodeXml_(m[2]);
    m = re.exec(String(tag));
  }
  return attrs;
}

function xlsxColumnIndex_(cellRef) {
  const letters = (String(cellRef).match(/^[A-Z]+/) || [''])[0];
  let total = 0;
  for (let i = 0; i < letters.length; i += 1) total = total * 26 + letters.charCodeAt(i) - 64;
  return total;
}

function xlsxCellXml_(cellRef, value, style) {
  const stylePart = style ? ' s="' + style + '"' : '';
  if (value === null || value === undefined || value === '') return '<c r="' + cellRef + '"' + stylePart + '/>';
  if (typeof value === 'number') return '<c r="' + cellRef + '"' + stylePart + '><v>' + value + '</v></c>';
  return '<c r="' + cellRef + '"' + stylePart + ' t="inlineStr"><is><t>' + xlsxEscape_(value) + '</t></is></c>';
}

// 指定セルを value で置換する。既存の書式(s属性)は保持。無ければ行内/新規行に挿入する。
function xlsxSetCell_(sheetXml, cellRef, value) {
  const cellPattern = new RegExp('<c\\b(?=[^>]*\\br="' + cellRef + '")[^>]*\\/>|<c\\b(?=[^>]*\\br="' + cellRef + '")[^>]*>[\\s\\S]*?<\\/c>');
  const existingMatch = sheetXml.match(cellPattern);
  const existing = existingMatch ? existingMatch[0] : null;
  const styleMatch = existing ? existing.match(/\bs="([^"]+)"/) : null;
  const style = styleMatch ? styleMatch[1] : '';
  const nextCell = xlsxCellXml_(cellRef, value, style);
  if (existing) return sheetXml.replace(cellPattern, nextCell);
  const rowNumber = Number((String(cellRef).match(/\d+$/) || [0])[0]);
  const rowPattern = new RegExp('(<row[^>]*\\br="' + rowNumber + '"[^>]*>)([\\s\\S]*?)(</row>)');
  const rowReplaced = sheetXml.replace(rowPattern, (match, open, content, close) => {
    const cellRe = /<c\b(?=[^>]*\br="([A-Z]+\d+)")[^>]*\/>|<c\b(?=[^>]*\br="([A-Z]+\d+)")[^>]*>[\s\S]*?<\/c>/g;
    let insertBefore = null;
    let cm = cellRe.exec(content);
    while (cm !== null) {
      const ref = cm[1] || cm[2];
      if (xlsxColumnIndex_(ref) > xlsxColumnIndex_(cellRef)) { insertBefore = cm[0]; break; }
      cm = cellRe.exec(content);
    }
    if (!insertBefore) return open + content + nextCell + close;
    return open + content.replace(insertBefore, nextCell + insertBefore) + close;
  });
  if (rowReplaced !== sheetXml) return rowReplaced;
  const rowXml = '<row r="' + rowNumber + '">' + nextCell + '</row>';
  const rowRe = /<row\b(?=[^>]*\br="(\d+)")[^>]*>[\s\S]*?<\/row>/g;
  let insertBeforeRow = null;
  let rm = rowRe.exec(sheetXml);
  while (rm !== null) {
    if (Number(rm[1]) > rowNumber) { insertBeforeRow = rm[0]; break; }
    rm = rowRe.exec(sheetXml);
  }
  if (insertBeforeRow) return sheetXml.replace(insertBeforeRow, rowXml + insertBeforeRow);
  return sheetXml.replace('</sheetData>', rowXml + '</sheetData>');
}

function fillClientReport_(ss, rows, month) {
  const settingSheet = ss.getSheetByName('基本設定と使い方');
  const [year, monthNumber] = month.split('-').map(Number);
  if (settingSheet) {
    settingSheet.getRange('B3').setValue(year);
    settingSheet.getRange('B4').setValue(monthNumber);
    const holidays = rows.filter((row) => row.holidayName).slice(0, 5);
    for (let i = 0; i < 5; i += 1) settingSheet.getRange(5 + i, 3).setValue(holidays[i] ? Number(holidays[i].workDate.slice(8, 10)) : '');
  }
  const sheet = ss.getSheetByName('出勤簿') || ss.getSheets()[0];
  for (let index = 0; index < 31; index += 1) {
    const row = rows[index];
    const r = 6 + index;
    if (!row) {
      sheet.getRange(r, 3, 1, 7).clearContent();
      continue;
    }
    const isHoliday = row.workType === 'holiday';
    sheet.getRange(r, 3).setValue(isHoliday ? '' : timeDate_(row.startTime) || '');
    sheet.getRange(r, 4).setValue(isHoliday ? '' : timeDate_(row.endTime) || '');
    sheet.getRange(r, 5).setValue(isHoliday || !row.breakMinutes ? '' : row.breakMinutes / 1440);
    sheet.getRange(r, 9).setValue(isHoliday ? '' : row.note || '');
  }
}

function previewRows_(month, targetType) {
  assertTargetType_(targetType);
  return rowsForReport_(month, true).map((row) => ({
    workDate: row.workDate,
    workType: row.workType,
    category: row.category,
    startTime: row.startTime,
    endTime: row.endTime,
    breakMinutes: row.breakMinutes,
    note: row.note,
    holidayName: row.holidayName,
  }));
}

function confirmMonth_(payload) {
  const month = assertMonth_(payload.month || currentMonth_());
  // 営業用の実ファイルが無ければ月末確認できない(送付対象が存在しないため)。
  if (!liveReports_(month).some((r) => r.targetType === 'sales')) {
    throw new Error('ファイルがないと月末確認ができません。');
  }
  const record = {
    id: id_('conf'),
    user_id: cfg_().userId,
    month,
    attendance_reviewed: true,
    report_preview_reviewed: true,
    sales_attachment_confirmed: true,
    confirmed_at: now_(),
    confirmed_by: cfg_().allowedEmail,
  };
  upsertRow_('monthly_confirmations', record, ['user_id', 'month']);
  addAudit_('monthly_confirmation_saved', 'monthly_confirmation', month, { month });
  return { monthlyConfirmation: confirmation_(month), attendance: monthState_(month) };
}

// 月末確認を解除(クリア)する。3フラグを false に戻し、Gmail下書き作成のゲートを再び有効にする。
function clearMonthConfirmation_(payload) {
  const month = assertMonth_(payload.month || currentMonth_());
  const existing = rows_('monthly_confirmations').find((item) => item.user_id === cfg_().userId && item.month === month);
  const record = Object.assign({ id: id_('conf'), user_id: cfg_().userId, month }, existing || {}, {
    attendance_reviewed: false,
    report_preview_reviewed: false,
    sales_attachment_confirmed: false,
    confirmed_at: '',
    confirmed_by: '',
  });
  upsertRow_('monthly_confirmations', record, ['user_id', 'month']);
  addAudit_('monthly_confirmation_cleared', 'monthly_confirmation', month, { month });
  return { monthlyConfirmation: confirmation_(month), attendance: monthState_(month) };
}

function createGmailDraft_(payload) {
  const month = assertMonth_(payload.month || currentMonth_());
  const confirmation = confirmation_(month);
  if (!confirmation.attendanceReviewed || !confirmation.reportPreviewReviewed || !confirmation.salesAttachmentConfirmed) {
    throw new Error('月末確認を完了してください。');
  }
  // 既存の営業用レポート(実ファイルが存在する最新のもの)を使う。無い場合でも勝手に生成しない。
  // (以前は createReport_ で自動生成しており、送信時に意図しない報告書が作られてしまっていた。)
  const salesReports = rows_('reports').filter((row) => row.user_id === cfg_().userId && row.month === month && row.target_type === 'sales');
  let report = null;
  for (let i = salesReports.length - 1; i >= 0; i -= 1) {
    if (driveFileExists_(salesReports[i].xlsx_file_id)) { report = salesReports[i]; break; }
  }
  if (!report) {
    throw new Error('営業用の作業報告書がまだありません。先に「月次作成」または「今日の内容をファイルに反映」で作成してください。');
  }
  const file = DriveApp.getFileById(report.xlsx_file_id);
  // テスト送信: 宛先を自分(ログインアカウント)にし CC 無し・件名に[テスト]を付ける。本番宛先には送らない。
  const testMode = Boolean(payload.testMode);
  const recipient = testMode ? cfg_().allowedEmail : cfg_().recipient;
  const ccRecipients = testMode ? [] : cfg_().ccRecipients;
  const subject = (testMode ? '[テスト] ' : '') + '【作業報告書】' + month + ' ' + cfg_().userName;
  const body = String(payload.bodyText || defaultGmailBody_(month));
  const draft = GmailApp.createDraft(recipient, subject, body, {
    cc: ccRecipients.join(','),
    attachments: [file.getBlob().setName(report.file_name || report.fileName)],
  });
  const record = {
    id: id_('draft'),
    user_id: cfg_().userId,
    month,
    gmail_draft_id: draft.getId(),
    sales_report_id: report.id || report.reportId,
    recipient: recipient,
    subject,
    cc_recipients: ccRecipients.join(','),
    sender_email: cfg_().allowedEmail,
    status: 'draft_ready',
    gmail_message_id: '',
    created_at: now_(),
    sent_at: '',
  };
  appendRow_('gmail_drafts', record);
  addAudit_('gmail_draft_created', 'gmail_draft', record.id, { month, salesReportId: record.sales_report_id });
  return { draft: publicDraft_(record), attendance: monthState_(month) };
}

function sendGmailDraft_(payload) {
  const draftId = String(payload.draftId || '');
  const draft = rows_('gmail_drafts').find((row) => row.id === draftId && row.user_id === cfg_().userId);
  if (!draft) throw new Error('Gmail下書きが見つかりません。');
  if (draft.status === 'sent') throw new Error('このメールは送信済みです。');
  if (draft.status !== 'draft_ready') throw new Error('送信できる下書きではありません。');
  try {
    const gmailDraft = GmailApp.getDraft(draft.gmail_draft_id);
    gmailDraft.send();
  } catch (error) {
    // Gmail上で下書きが手動削除された等で送信できない場合。例外で返すと画面が更新されず送信ボタンが
    // 残るため、記録を missing に更新した最新状態と通知を返し、クライアント側で無効化・案内させる。
    const missing = Object.assign({}, draft, { status: 'missing', updated_at: now_() });
    upsertRow_('gmail_drafts', missing, ['id']);
    addAudit_('gmail_draft_missing', 'gmail_draft', draft.id, { month: draft.month });
    return {
      draft: publicDraft_(missing),
      attendance: monthState_(draft.month),
      notice: 'Gmail上の下書きが削除されています。新しい下書きを作成してください。',
    };
  }
  const next = Object.assign({}, draft, { status: 'sent', gmail_message_id: draft.gmail_draft_id, sent_at: now_() });
  upsertRow_('gmail_drafts', next, ['id']);
  addAudit_('gmail_message_sent', 'gmail_draft', draft.id, { month: draft.month });
  return { draft: publicDraft_(next), attendance: monthState_(draft.month) };
}

function syncDraftStatus_(payload) {
  const month = assertMonth_(payload.month || currentMonth_());
  rows_('gmail_drafts').filter((row) => row.month === month && row.status === 'draft_ready').forEach((row) => {
    try {
      GmailApp.getDraft(row.gmail_draft_id);
    } catch (error) {
      upsertRow_('gmail_drafts', Object.assign({}, row, { status: 'missing' }), ['id']);
      addAudit_('gmail_draft_missing', 'gmail_draft', row.id, { month });
    }
  });
  return { attendance: monthState_(month) };
}

function rowsForReport_(month, includeSavedAttendance) {
  const days = daysInMonth_(month);
  const saved = {};
  if (includeSavedAttendance) {
    rows_('attendance_days').filter((row) => row.user_id === cfg_().userId && String(row.work_date).startsWith(month)).forEach((row) => {
      saved[row.work_date] = publicDay_(row);
    });
  }
  return days.map((date) => {
    if (saved[date]) return saved[date];
    const holidayName = HOLIDAYS[date] || '';
    const dow = dateObject_(date).getDay();
    const isHoliday = holidayName || dow === 0 || dow === 6;
    return {
      workDate: date,
      workType: isHoliday ? 'holiday' : 'workday',
      category: isHoliday ? '休日' : '①',
      startTime: isHoliday ? '' : APP.defaultTimes.startTime,
      endTime: isHoliday ? '' : APP.defaultTimes.endTime,
      breakMinutes: isHoliday ? 0 : APP.defaultTimes.breakMinutes,
      note: '',
      source: 'generated',
      holidayName,
    };
  });
}

function rows_(sheetName) {
  const ss = ensureStore_(true);
  const sheet = ss.getSheetByName(sheetName);
  const values = sheet.getDataRange().getValues();
  const headers = SHEETS[sheetName];
  return values.slice(1).filter((row) => row.some((value) => value !== '')).map((row) => {
    const obj = {};
    headers.forEach((header, index) => obj[header] = normalizeStoreValue_(header, row[index]));
    return obj;
  });
}

// Sheet は日付文字列(例 '2026-07-03')を書くとセルを日付型へ自動変換し、読み戻すと Date を返す。
// そのままだと work_date/month の文字列照合(startsWith や ===、saved[date] キー)が崩れ、保存済み
// 勤怠が生成にマージされない(=常に既定値)。ここで日付/月の列だけ文字列キーへ正規化して吸収する。
function normalizeStoreValue_(header, value) {
  if (value instanceof Date) {
    if (header === 'work_date') return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
    if (header === 'month') return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM');
    // 始業/終業は 'HH:mm' 文字列へ。Sheet が '09:00' を時刻型に変換して Date で読み戻すため、
    // 文字列に戻さないと timeToExcelSerial_ の String(Date).split(':') が NaN になり、
    // 保存済みの日の時刻セルが壊れる(<v>NaN</v>)。
    if (header === 'start_time' || header === 'end_time') return Utilities.formatDate(value, Session.getScriptTimeZone(), 'HH:mm');
    // created_at 等の日時列も ISO 文字列へ。Date を含む戻り値は google.script.run の
    // シリアライズで成功コールバックが発火せず「作成中」のまま固まる原因になるため、
    // 返す値は必ず素の文字列にしておく。
    return value.toISOString();
  }
  return value;
}

function appendRow_(sheetName, object) {
  const ss = ensureStore_(true);
  const sheet = ss.getSheetByName(sheetName);
  sheet.appendRow(SHEETS[sheetName].map((header) => object[header] ?? ''));
}

function upsertRow_(sheetName, object, keys) {
  const ss = ensureStore_(true);
  const sheet = ss.getSheetByName(sheetName);
  const headers = SHEETS[sheetName];
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i += 1) {
    // 日付/月の列は Sheet が Date に変換して保持することがあるため、比較時に正規化して突き合わせる
    // (正規化しないと既存行を見つけられず重複行が増え、同日勤怠が二重になる)。
    const matches = keys.every((key) => normalizeStoreValue_(key, values[i][headers.indexOf(key)]) === normalizeStoreValue_(key, object[key]));
    if (matches) {
      const current = {};
      headers.forEach((header, index) => current[header] = values[i][index]);
      const next = Object.assign({}, current, object);
      sheet.getRange(i + 1, 1, 1, headers.length).setValues([headers.map((header) => next[header] ?? '')]);
      return;
    }
  }
  appendRow_(sheetName, object);
}

function addAudit_(eventType, resourceType, resourceId, payload) {
  appendRow_('audit_events', {
    id: id_('audit'),
    event_type: eventType,
    resource_type: resourceType,
    resource_id: resourceId,
    payload_json: JSON.stringify(payload || {}),
    created_at: now_(),
  });
}

function rootFolder_() {
  const name = settings_().rootFolderName || APP.rootFolderName;
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

// 診断用: エディタから直接実行してフォルダ作成が動くかを確認する。
// 実行後、実行ログ（Ctrl+Enter/表示>ログ）に作成先フォルダのURLが出る。
function debugFolderCreation() {
  const root = rootFolder_();
  const kind = childFolder_(root, '営業用');
  const monthFolder = childFolder_(kind, '2026年07月');
  const file = monthFolder.createFile('フォルダ作成テスト.txt', 'ok', 'text/plain');
  const info = {
    ルート: root.getName() + ' -> ' + root.getUrl(),
    種別: kind.getName() + ' -> ' + kind.getUrl(),
    年月: monthFolder.getName() + ' -> ' + monthFolder.getUrl(),
    テストファイル: file.getName() + ' -> ' + file.getUrl(),
  };
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}

function childFolder_(parent, name) {
  const folders = parent.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : parent.createFolder(name);
}

function moveFileToFolder_(file, folder) {
  folder.addFile(file);
  const parents = file.getParents();
  while (parents.hasNext()) {
    const parent = parents.next();
    if (parent.getId() !== folder.getId()) parent.removeFile(file);
  }
}

function exportSpreadsheetAsXlsx_(spreadsheetId, fileName) {
  const url = 'https://docs.google.com/spreadsheets/d/' + spreadsheetId + '/export?format=xlsx';
  const response = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  if (response.getResponseCode() >= 300) throw new Error('Excelエクスポートに失敗しました: HTTP ' + response.getResponseCode());
  return response.getBlob().setName(fileName);
}

function existingAttendanceId_(workDate) {
  const row = rows_('attendance_days').find((item) => item.user_id === cfg_().userId && item.work_date === workDate);
  return row ? row.id : '';
}

function confirmation_(month) {
  const row = rows_('monthly_confirmations').find((item) => item.user_id === cfg_().userId && item.month === month);
  return {
    attendanceReviewed: Boolean(row && row.attendance_reviewed),
    reportPreviewReviewed: Boolean(row && row.report_preview_reviewed),
    salesAttachmentConfirmed: Boolean(row && row.sales_attachment_confirmed),
  };
}

function publicDay_(row) {
  return {
    workDate: row.work_date || row.workDate,
    workType: row.work_type || row.workType || 'workday',
    category: row.category || '①',
    startTime: row.start_time || row.startTime || '',
    endTime: row.end_time || row.endTime || '',
    breakMinutes: Number(row.break_minutes ?? row.breakMinutes ?? 0),
    note: row.note || '',
    source: row.source || '',
    holidayName: row.holiday_name || row.holidayName || '',
    status: row.source === 'manual' ? 'reflected' : 'draft',
  };
}

function publicReport_(row) {
  return {
    reportId: row.id || row.reportId,
    month: row.month,
    targetType: row.target_type || row.targetType,
    status: row.status,
    driveFileId: row.drive_file_id || row.driveFileId,
    xlsxFileId: row.xlsx_file_id || row.xlsxFileId,
    fileName: row.file_name || row.fileName,
    generatedAt: row.generated_at || row.generatedAt,
  };
}

function publicDraft_(row) {
  return {
    draftId: row.id,
    gmailDraftId: row.gmail_draft_id,
    salesReportId: row.sales_report_id,
    recipient: row.recipient,
    subject: row.subject,
    ccRecipients: String(row.cc_recipients || '').split(',').filter(Boolean),
    senderEmail: row.sender_email,
    status: row.status,
    gmailMessageId: row.gmail_message_id || '',
    createdAt: row.created_at,
    sentAt: row.sent_at || '',
  };
}

function targetTypesForMode_(mode) {
  if (mode === 'clientOnly') return ['client'];
  if (mode === 'both') return ['sales', 'client'];
  return ['sales'];
}

function reportKindLabel_(targetType) {
  return targetType === 'client' ? '客先用' : '営業用';
}

function reportFileName_(month, targetType) {
  const parts = month.split('-');
  if (targetType === 'client') return '出勤簿_' + parts[0] + '年' + parts[1] + '月_' + cfg_().clientFilePerson + '.xlsx';
  return cfg_().salesFileName;
}

function defaultGmailBody_(month) {
  const parts = month.split('-').map(Number);
  const c = cfg_();
  return c.recipientName + '\n\nお疲れ様です。' + c.userName + 'です。\n\n' + parts[0] + '年' + parts[1] + '月の作業報告書を添付しました。交通費は、下記の通りです。\nご確認のほど、よろしくお願いいたします。\n\n経路：' + c.commuteRoute + '\n交通費： ' + c.commuteCost;
}

function assertMonth_(value) {
  const month = String(value || '');
  if (!/^\d{4}-\d{2}$/.test(month)) throw new Error('対象年月が不正です。');
  return month;
}

function assertDate_(value) {
  const date = String(value || '');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(date)) throw new Error('日付が不正です。');
  return date;
}

function assertCategory_(value) {
  const category = String(value || '①');
  if (APP.categories.indexOf(category) < 0) throw new Error('区分が不正です。');
  return category;
}

function assertTime_(value) {
  const time = String(value || '');
  if (!/^\d{2}:\d{2}$/.test(time)) throw new Error('時刻が不正です。');
  const minute = Number(time.slice(3, 5));
  if (minute % 10 !== 0) throw new Error('時刻は10分単位で入力してください。');
  return time;
}

function assertBreakMinutes_(value) {
  const minutes = Number(value ?? APP.defaultTimes.breakMinutes);
  if (!Number.isFinite(minutes) || minutes < 0 || minutes > 600 || minutes % 10 !== 0) throw new Error('休憩時間が不正です。');
  return minutes;
}

function assertTargetType_(value) {
  const targetType = String(value || '');
  if (targetType !== 'sales' && targetType !== 'client') throw new Error('帳票種別が不正です。');
  return targetType;
}

function assertDriveFileId_(value) {
  const id = String(value || '').trim();
  if (!/^[A-Za-z0-9_-]{10,200}$/.test(id)) throw new Error('テンプレートIDが不正です。');
  return id;
}

function assertTemplateFileId_(value) {
  const id = assertDriveFileId_(value);
  const file = DriveApp.getFileById(id);
  const mimeType = file.getMimeType();
  if (mimeType !== APP.mimeXlsx && mimeType !== APP.mimeGoogleSheets) {
    throw new Error('テンプレートIDはExcelファイルまたはGoogle Sheets形式のファイルを指定してください。');
  }
  return id;
}

function copyTemplateAsSpreadsheet_(templateId, name, folder) {
  let file;
  try {
    file = DriveApp.getFileById(assertDriveFileId_(templateId));
  } catch (err) {
    throw new Error('テンプレートファイルが見つかりません。テンプレートIDを確認してください。');
  }
  const mimeType = file.getMimeType();
  if (mimeType === APP.mimeGoogleSheets) return file.makeCopy(name, folder);
  if (mimeType !== APP.mimeXlsx) {
    throw new Error('テンプレートはExcel(.xlsx)またはGoogleスプレッドシート形式のファイルにしてください。');
  }
  if (typeof Drive === 'undefined' || !Drive.Files) {
    throw new Error('ExcelをGoogleスプレッドシートに変換できません。Apps Scriptの「サービス」でDrive API(詳細サービス)を有効化してください。');
  }
  const blob = file.getBlob();
  let converted;
  try {
    if (Drive.Files.insert) {
      // Drive 詳細サービス v2
      converted = Drive.Files.insert(
        { title: name, mimeType: APP.mimeGoogleSheets, parents: [{ id: folder.getId() }] },
        blob,
        { convert: true, supportsAllDrives: true }
      );
    } else if (Drive.Files.create) {
      // Drive 詳細サービス v3
      converted = Drive.Files.create(
        { name: name, mimeType: APP.mimeGoogleSheets, parents: [folder.getId()] },
        blob,
        { supportsAllDrives: true }
      );
    } else {
      throw new Error('Drive API(詳細サービス)のFiles操作が見つかりません。');
    }
  } catch (err) {
    throw new Error('Excelテンプレートの変換に失敗しました（' + (err && err.message ? err.message : err) + '）。テンプレートが有効な.xlsxか確認してください。');
  }
  if (!converted || !converted.id) {
    throw new Error('Excelテンプレートを変換できませんでした。Drive API(詳細サービス)の有効化と権限をご確認ください。');
  }
  const result = DriveApp.getFileById(converted.id);
  // 変換時に parents が無視され My Drive 直下に作られる場合があるため、確実にフォルダへ移動する。
  try {
    result.moveTo(folder);
  } catch (err) {
    try {
      folder.addFile(result);
      DriveApp.getRootFolder().removeFile(result);
    } catch (err2) {
      // 移動に失敗しても、この中間ファイルは呼び出し元で削除されるため続行する
    }
  }
  return result;
}

function currentMonth_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM');
}

function now_() {
  return new Date().toISOString();
}

function id_(prefix) {
  return prefix + '_' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
}

function daysInMonth_(month) {
  const parts = month.split('-').map(Number);
  const last = new Date(parts[0], parts[1], 0).getDate();
  return Array.from({ length: last }, (_, index) => month + '-' + String(index + 1).padStart(2, '0'));
}

function dateObject_(date) {
  const parts = String(date).split('-').map(Number);
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function timeDate_(time) {
  if (!time) return null;
  const parts = String(time).split(':').map(Number);
  return new Date(1899, 11, 30, parts[0], parts[1], 0);
}

function weekdayLabel_(date) {
  return ['日', '月', '火', '水', '木', '金', '土'][dateObject_(date).getDay()];
}
