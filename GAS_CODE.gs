var SETTINGS_DOC_NAME = '給与明細設定値';
var SETTINGS_DOC_NAME_CANDIDATES = ['給与明細設定値', '給与明細設定'];
var DEFAULT_MASTER_FOLDER_ID = '1CD96PrUWhdAIXVtfrQCInTuBuceeukuZ'; // 初期値（プロパティ未設定時のフォールバック）
var DEFAULT_EMPLOYEE_PARENT_ID = '1mrG28x7dH9yiZ1uKRhmKzmzgdenKddt6'; // 従業員フォルダを新規作成する場所（初期値）
function getEmployeeParentId() {
  var v = PropertiesService.getScriptProperties().getProperty('EMPLOYEE_PARENT_ID');
  return v || DEFAULT_EMPLOYEE_PARENT_ID;
}

var DEFAULT_RETENTION_DAYS = 1826; // 初期値：5年（法令に応じて管理者が変更してください）
function getRetentionDays() {
  var v = PropertiesService.getScriptProperties().getProperty('RETENTION_DAYS');
  return v ? parseInt(v) : DEFAULT_RETENTION_DAYS;
}
function addDaysToDateString(dateStr, days) {
  var d = new Date(dateStr + 'T00:00:00');
  d.setDate(d.getDate() + days);
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return y + '-' + m + '-' + day;
}
function todayDateString() {
  var d = new Date();
  var y = d.getFullYear();
  var m = ('0' + (d.getMonth() + 1)).slice(-2);
  var day = ('0' + d.getDate()).slice(-2);
  return y + '-' + m + '-' + day;
}
function logSystemEvent(eventType, detail) {
  var logDoc = getOrCreateMasterDoc('操作ログ');
  var logLine = new Date().toISOString() + ':system:system:' + eventType + ':' + detail.replace(/\n/g, ' ');
  appendLine(logDoc, logLine);
}
var BOOTSTRAP_DOC_ID = '15GaNuDtOCUW311AS-dKXSbu7x2-mVGHblkIe25KlY4I'; // GASのURLを書いた、公開閲覧可能なドキュメントのID
var APP_VERSION_GAS = '2026.07.13.1'; // このGASコードの版数。デプロイのたびに手動で書き換えてください（自動更新お知らせの検知に使われます）

function getMasterFolderId() {
  var v = PropertiesService.getScriptProperties().getProperty('MASTER_FOLDER_ID');
  return v || DEFAULT_MASTER_FOLDER_ID;
}
var MASTER_FOLDER_ID = getMasterFolderId();

var DEFAULT_INQUIRY_EMAIL = 'a_aoyama@dsbz.jp';
var DEFAULT_INQUIRY_NAME = '青山';
function getInquiryEmail() {
  var v = PropertiesService.getScriptProperties().getProperty('INQUIRY_EMAIL');
  return v || DEFAULT_INQUIRY_EMAIL;
}
function getInquiryName() {
  var v = PropertiesService.getScriptProperties().getProperty('INQUIRY_NAME');
  return v || DEFAULT_INQUIRY_NAME;
}

var DEFAULT_TRANSPORT_EMAIL = 'a_aoyama@dsbz.jp';
function getTransportEmail() {
  var v = PropertiesService.getScriptProperties().getProperty('TRANSPORT_EMAIL');
  return v || DEFAULT_TRANSPORT_EMAIL;
}

var DEFAULT_LEAVE_TYPE_EMAIL = 'a_aoyama@dsbz.jp';
function getLeaveTypeEmail() {
  var v = PropertiesService.getScriptProperties().getProperty('LEAVE_TYPE_EMAIL');
  return v || DEFAULT_LEAVE_TYPE_EMAIL;
}

// ============================================================
// 交通費申請
// ============================================================
var TRANSPORT_FOLDER_NAME = '交通費申請';
var TRANSPORT_TEMPLATE_NAME = '交通費申請書(原紙)';
var TRANSPORT_LEDGER_NAME = '交通費申請一覧';
var TRANSPORT_REQUEST_MARKER = '##TRANSPORT_REQUEST##';
var TRANSPORT_TRIP_HEADER_ROW = 4;
var TRANSPORT_TRIP_START_ROW = 5;
var TRANSPORT_TRIP_MAX_ROWS = 31; // 原紙の行5〜35（合計行は36）

function findTransportFolder() {
  var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  var it = masterFolder.getFoldersByName(TRANSPORT_FOLDER_NAME);
  return it.hasNext() ? it.next() : null;
}
function findTransportTemplateFile() {
  // 原紙はリクエストごとのコピー置き場（交通費申請フォルダ）ではなく、管理情報などと同じマスターフォルダ直下に置く運用
  var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  var it = masterFolder.getFilesByName(TRANSPORT_TEMPLATE_NAME);
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) return f;
  }
  return null;
}
function getOrCreateTransportLedgerDoc(transportFolder) {
  var it = transportFolder.getFilesByName(TRANSPORT_LEDGER_NAME);
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType() === MimeType.GOOGLE_DOCS) return DocumentApp.openById(f.getId());
  }
  var newDoc = DocumentApp.create(TRANSPORT_LEDGER_NAME);
  var newFile = DriveApp.getFileById(newDoc.getId());
  transportFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);
  return newDoc;
}
// 交通費申請一覧の1行は「requestId:empId:empName:appliedAtISO:status:decidedAtISO:reasonB64:payloadB64」。
// reasonB64は却下理由（無ければ空）、payloadB64は{trips,total,sheetFileId,sheetUrl}のJSONをUTF-8セーフなbase64で埋め込む
// （コロンを含む自由記述の理由や、明細行を安全にテキスト行へ収めるため）。
function encodeTransportPayload(obj) {
  return Utilities.base64Encode(JSON.stringify(obj), Utilities.Charset.UTF_8);
}
function decodeTransportPayload(b64) {
  if (!b64) return null;
  try { return JSON.parse(Utilities.newBlob(Utilities.base64Decode(b64, Utilities.Charset.UTF_8)).getDataAsString('UTF-8')); } catch (e) { return null; }
}
function parseTransportLedgerLine(line) {
  var parts = line.split(':');
  if (parts.length < 8) return null;
  var reasonB64 = parts[6];
  var payload = decodeTransportPayload(parts[7]) || { trips: [], total: 0 };
  return {
    requestId: parts[0], empId: parts[1], empName: parts[2],
    appliedAt: parts[3] ? Utilities.newBlob(Utilities.base64Decode(parts[3], Utilities.Charset.UTF_8)).getDataAsString('UTF-8') : '',
    status: parts[4],
    decidedAt: parts[5] ? Utilities.newBlob(Utilities.base64Decode(parts[5], Utilities.Charset.UTF_8)).getDataAsString('UTF-8') : '',
    rejectReason: reasonB64 ? Utilities.newBlob(Utilities.base64Decode(reasonB64, Utilities.Charset.UTF_8)).getDataAsString('UTF-8') : '',
    trips: payload.trips || [], total: payload.total || 0,
    sheetFileId: payload.sheetFileId || '', sheetUrl: payload.sheetUrl || ''
  };
}
function serializeTransportLedgerLine(r) {
  // appliedAt/decidedAtはISO日時（コロンを含む）なので、コロン区切りの行フォーマットと衝突しないようbase64化する
  var appliedAtB64 = Utilities.base64Encode(r.appliedAt || '', Utilities.Charset.UTF_8);
  var decidedAtB64 = r.decidedAt ? Utilities.base64Encode(r.decidedAt, Utilities.Charset.UTF_8) : '';
  var reasonB64 = r.rejectReason ? Utilities.base64Encode(r.rejectReason, Utilities.Charset.UTF_8) : '';
  var payloadB64 = encodeTransportPayload({ trips: r.trips, total: r.total, sheetFileId: r.sheetFileId, sheetUrl: r.sheetUrl });
  return [r.requestId, r.empId, r.empName, appliedAtB64, r.status, decidedAtB64, reasonB64, payloadB64].join(':');
}

function getSettingsDocFile(folder) {
  for (var i = 0; i < SETTINGS_DOC_NAME_CANDIDATES.length; i++) {
    var it = folder.getFilesByName(SETTINGS_DOC_NAME_CANDIDATES[i]);
    while (it.hasNext()) {
      var f = it.next();
      if (f.getMimeType() === MimeType.GOOGLE_DOCS) return f;
    }
  }
  return null;
}

function findMasterListFile() {
  var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  var candidates = ['管理情報', '社員一覧'];
  for (var ci = 0; ci < candidates.length; ci++) {
    var lf = masterFolder.getFilesByName(candidates[ci]);
    while (lf.hasNext()) {
      var cand = lf.next();
      if (cand.getMimeType() === MimeType.GOOGLE_DOCS) return cand;
    }
  }
  return null;
}

function getTabBodyByTitle(doc, title) {
  var tabs = doc.getTabs();
  for (var i = 0; i < tabs.length; i++) {
    if (tabs[i].getTitle() === title) return tabs[i].asDocumentTab().getBody();
  }
  return null;
}

function digitsOnly(s) {
  return (s || '').replace(/[^0-9]/g, '');
}

function parseUserRecords(text) {
  var lines = text.split('\n');
  var records = [];
  var cur = null;
  for (var i = 0; i < lines.length; i++) {
    var ln = lines[i].trim();
    if (!ln) continue;
    var norm = ln.replace(/：/g, ':');
    var idx = norm.indexOf(':');
    if (idx < 0) continue;
    var key = norm.substring(0, idx).trim();
    var val = norm.substring(idx + 1).trim();
    if (key === 'ユーザーID') {
      if (cur) records.push(cur);
      cur = {};
    }
    if (cur) cur[key] = val;
  }
  if (cur) records.push(cur);
  return records;
}

function serializeUserRecords(records) {
  var lines = [];
  records.forEach(function(r, idx) {
    if (idx > 0) lines.push('');
    lines.push('ユーザーID:' + (r['ユーザーID'] || ''));
    if (r['苗字'] !== undefined) lines.push('苗字：' + r['苗字']);
    if (r['名前'] !== undefined) lines.push('名前：' + r['名前']);
    if (r['みょうじ'] !== undefined) lines.push('みょうじ：' + r['みょうじ']);
    if (r['なまえ'] !== undefined) lines.push('なまえ：' + r['なまえ']);
    if (r['生年月日'] !== undefined) lines.push('生年月日:' + r['生年月日']);
    if (r['管理者権限'] !== undefined) lines.push('管理者権限:' + r['管理者権限']);
    if (r['役員フラグ'] !== undefined) lines.push('役員フラグ:' + r['役員フラグ']);
    if (r['削除フラグ'] !== undefined) lines.push('削除フラグ:' + r['削除フラグ']);
    if (r['削除日'] !== undefined) lines.push('削除日:' + r['削除日']);
    if (r['削除予定日'] !== undefined) lines.push('削除予定日:' + r['削除予定日']);
    if (r['削除理由'] !== undefined) lines.push('削除理由:' + r['削除理由']);
    if (r['告知済み'] !== undefined) lines.push('告知済み:' + r['告知済み']);
  });
  return lines.join('\n');
}

function findMainListRow(mainText, targetId) {
  var rows = mainText.split('\n');
  for (var i = 0; i < rows.length; i++) {
    var row = rows[i].trim();
    if (!row) continue;
    var cols = row.replace(/：/g, ':').split(':');
    if (cols.length < 4) continue;
    if (cols[1].trim() === targetId) {
      return { index: i, cols: cols, name: cols[4] ? cols[4].trim() : '' };
    }
  }
  return null;
}

function getOrCreateMasterDoc(name) {
  var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  var it = masterFolder.getFilesByName(name);
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType() === MimeType.GOOGLE_DOCS) return DocumentApp.openById(f.getId());
  }
  var newDoc = DocumentApp.create(name);
  var newFile = DriveApp.getFileById(newDoc.getId());
  masterFolder.addFile(newFile);
  DriveApp.getRootFolder().removeFile(newFile);
  return newDoc;
}

function appendLine(doc, line) {
  var body = doc.getBody();
  body.appendParagraph(line);
  doc.saveAndClose();
}

var COMPANY_SETTING_KEYS = ['有給次回付与', '付与数', '給料日', '締日', '賞与月1', '賞与日1', '賞与月2', '賞与日2', '残業時間警告基準'];
// 「会社共通設定」ヘッダーが無い旧形式のドキュメントでも設定値を読み取れるよう、
// ヘッダーの有無に関わらずキー名で判定する（ヘッダーがあれば単に読み飛ばす）。
function parseCompanySettingsFromText(text) {
  var result = {};
  var inPaydaySection = false;
  var lines = text.split('\n');
  for (var i = 0; i < lines.length; i++) {
    var row = lines[i].trim();
    if (row === '給料日特例') { inPaydaySection = true; continue; }
    if (row === '会社共通設定') { inPaydaySection = false; continue; }
    if (!row) { inPaydaySection = false; continue; }
    if (inPaydaySection) continue;
    var idx = row.indexOf(':');
    if (idx <= 0) continue;
    var key = row.substring(0, idx).trim();
    if (COMPANY_SETTING_KEYS.indexOf(key) !== -1) {
      result[key] = row.substring(idx + 1).trim();
    }
  }
  return result;
}

function parsePaydayOverrides(text) {
  var lines = text.split('\n');
  var overrides = {};
  var inSection = false;
  for (var i = 0; i < lines.length; i++) {
    var row = lines[i].trim();
    if (row === '給料日特例') { inSection = true; continue; }
    if (row === '会社共通設定') { inSection = false; continue; }
    if (!row) { inSection = false; continue; }
    if (inSection) {
      var idx = row.indexOf(':');
      if (idx > 0) overrides[row.substring(0, idx).trim()] = row.substring(idx + 1).trim();
    }
  }
  return overrides;
}

function buildUserDetails(docS, mainTextS, targetIdS) {
  var rowS = findMainListRow(mainTextS, targetIdS);
  if (!rowS) return null;
  var userInfo = { sei:'', mei:'', dob:'', role: rowS.cols[0].trim(), isExecutive:false, deleted:false, deletedDate:'', scheduledDate:'', deleteReason:'' };
  var userBodyS = getTabBodyByTitle(docS, 'ユーザー情報');
  if (userBodyS) {
    var recordsS = parseUserRecords(userBodyS.getText());
    for (var si = 0; si < recordsS.length; si++) {
      if (recordsS[si]['ユーザーID'] === targetIdS) {
        userInfo.sei = recordsS[si]['苗字'] || '';
        userInfo.mei = recordsS[si]['名前'] || '';
        userInfo.dob = recordsS[si]['生年月日'] || '';
        userInfo.isExecutive = recordsS[si]['役員フラグ'] === '1';
        userInfo.deleted = recordsS[si]['削除フラグ'] === '1';
        userInfo.deletedDate = recordsS[si]['削除日'] || '';
        userInfo.scheduledDate = recordsS[si]['削除予定日'] || '';
        userInfo.deleteReason = recordsS[si]['削除理由'] || '';
        break;
      }
    }
  }
  var byYearS = {};
  try {
    var targetFolderS = DriveApp.getFolderById(rowS.cols[3].trim());
    var settingsDocS = getSettingsDocFile(targetFolderS);
    if (settingsDocS) {
      var settingsTextS = DocumentApp.openById(settingsDocS.getId()).getBody().getText();
      var sLinesS = settingsTextS.split('\n');
      for (var yi = 0; yi < sLinesS.length; yi++) {
        var yMatch = sLinesS[yi].trim().match(/^(\d{4})有給残:([\d.]+)$/);
        if (yMatch) byYearS[yMatch[1]] = yMatch[2];
      }
    }
  } catch (errS) { /* フォルダが見つからない等は無視してbyYearは空のまま返す */ }
  return { id: targetIdS, folderId: rowS.cols[3].trim(), userInfo: userInfo, byYear: byYearS };
}

// ============================================================
// 管理情報スプレッドシート（社員一覧・ユーザー情報・会社共通設定・給料日特例）
// 旧形式（管理情報ドキュメント）からの移行後は、こちらが正となる。
// ============================================================
var USER_INFO_HEADER = ['ユーザーID', '苗字', '名前', 'みょうじ', 'なまえ', '生年月日', '管理者権限', '役員フラグ', '削除フラグ', '削除日', '削除予定日', '削除理由', '告知済み'];

// 「休暇制度」シートのヘッダーと、国が定める法定休暇のデフォルト行
// URL・必要書類は暫定値（会社の実情に合わせて管理画面から修正してください）
var LEAVE_TYPES_HEADER = ['ID', '休暇名', '給与', '条件', '申請可', 'URL', '必要書類'];
var DEFAULT_LEAVE_TYPES = [
  ['産前産後休業', '×', '出産予定日の6週間前（多胎妊娠は14週間前）〜出産後8週間', '1', 'https://www.mhlw.go.jp/bosei/', '母子健康手帳（写し）|診断書'],
  ['育児休業', '×', '原則として子が1歳になるまで（要件を満たせば延長可）', '1', 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kodomo/shokuba_kosodate/ikuji/index.html', '母子健康手帳（写し）'],
  ['介護休業', '×', '対象家族1人につき通算93日まで（3回を上限に分割可）', '1', 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/kaigo/index.html', '介護保険被保険者証（写し）|診断書'],
  ['子の看護休暇', '×', '小学校就学前の子の世話・通院等、年5日（対象の子が2人以上は年10日）', '1', 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/kodomo/shokuba_kosodate/ikuji/index.html', '母子健康手帳（写し）'],
  ['介護休暇', '×', '要介護状態の家族の世話等、年5日（対象家族が2人以上は年10日）', '1', 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/kaigo/index.html', '介護保険被保険者証（写し）'],
  ['生理休暇', '×', '生理日の就業が著しく困難なとき', '1', 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/index.html', ''],
  ['公民権行使の休暇', '×', '選挙権の行使・裁判員など公の職務を執行するとき', '1', 'https://www.mhlw.go.jp/stf/seisakunitsuite/bunya/koyou_roudou/roudoukijun/index.html', '']
];

// 「休暇制度」シートを取得し、なければ作成して法定休暇のデフォルト行を登録する
// 既存シートがURL・必要書類列を追加する前のバージョンで作られていた場合は、列を補完し
// 法定休暇のデフォルト行にはURL・必要書類のデフォルト値を後から補完する（自己修復）
function getOrCreateLeaveTypesSheet(ss) {
  var sheet = ss.getSheetByName('休暇制度');
  if (!sheet) {
    sheet = ss.insertSheet('休暇制度');
    appendRowAsText(sheet, LEAVE_TYPES_HEADER);
    var baseTime = new Date().getTime();
    DEFAULT_LEAVE_TYPES.forEach(function (row, i) {
      appendRowAsText(sheet, [String(baseTime + i), row[0], row[1], row[2], row[3], row[4] || '', row[5] || '']);
    });
    return sheet;
  }
  var lastCol = sheet.getLastColumn();
  if (lastCol < LEAVE_TYPES_HEADER.length) {
    setRangeAsText(sheet, 1, lastCol + 1, 1, LEAVE_TYPES_HEADER.length - lastCol, [LEAVE_TYPES_HEADER.slice(lastCol)]);
  }
  var data = getSheetData(ss, '休暇制度');
  var defaultsByName = {};
  DEFAULT_LEAVE_TYPES.forEach(function (row) { defaultsByName[row[0]] = row; });
  data.rows.forEach(function (row, i) {
    var def = defaultsByName[row[1]];
    if (!def) return;
    var currentUrl = row[5] || '';
    var currentDocs = row[6] || '';
    if (!currentUrl && !currentDocs && (def[4] || def[5])) {
      setRangeAsText(sheet, i + 2, 6, 1, 2, [[def[4] || '', def[5] || '']]);
    }
  });
  return sheet;
}

// 「必要書類リスト」シート：休暇申請時の必要書類ドロップダウンの選択肢を管理
var DEFAULT_REQUIRED_DOCUMENTS = ['住民票（写し）', 'マイナンバーカード（両面コピー）', '母子健康手帳（写し）', '診断書', '医師の意見書', '介護保険被保険者証（写し）', '戸籍謄本', '在職証明書'];
function getOrCreateRequiredDocumentsSheet(ss) {
  var sheet = ss.getSheetByName('必要書類リスト');
  if (sheet) return sheet;
  sheet = ss.insertSheet('必要書類リスト');
  appendRowAsText(sheet, ['書類名']);
  DEFAULT_REQUIRED_DOCUMENTS.forEach(function (name) {
    appendRowAsText(sheet, [name]);
  });
  return sheet;
}

function findMasterSpreadsheetFile() {
  var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
  // Driveへ.xlsxをアップロードしてGoogleスプレッドシートへ自動変換した場合、
  // 見た目上のファイル名に拡張子が残ることがあるため、両方の名前で探す。
  var candidateNames = ['管理情報', '管理情報.xlsx'];
  for (var ni = 0; ni < candidateNames.length; ni++) {
    var it = masterFolder.getFilesByName(candidateNames[ni]);
    while (it.hasNext()) {
      var f = it.next();
      if (f.getMimeType() === MimeType.GOOGLE_SHEETS) return f;
    }
  }
  return null;
}

function openMasterSpreadsheet() {
  var f = findMasterSpreadsheetFile();
  return f ? SpreadsheetApp.openById(f.getId()) : null;
}

// 日付らしい文字列（生年月日・削除日・対象年月など）をAPI経由で新規セルに書き込むと、
// Googleスプレッドシートが自動的に日付型に変換してしまい、読み戻すとDateオブジェクトの
// 文字列表現に化けてしまう。書き込み前に必ずテキスト書式を明示することで防ぐ。
function appendRowAsText(sheet, values) {
  var row = sheet.getLastRow() + 1;
  var range = sheet.getRange(row, 1, 1, values.length);
  range.setNumberFormat('@');
  SpreadsheetApp.flush();
  range.setValues([values]);
  range.setHorizontalAlignment('left');
}
function setCellAsText(sheet, row, col, value) {
  var range = sheet.getRange(row, col);
  range.setNumberFormat('@');
  SpreadsheetApp.flush();
  range.setValue(value);
  range.setHorizontalAlignment('left');
}
function setRangeAsText(sheet, row, col, numRows, numCols, values) {
  var range = sheet.getRange(row, col, numRows, numCols);
  range.setNumberFormat('@');
  SpreadsheetApp.flush();
  range.setValues(values);
  range.setHorizontalAlignment('left');
}

function getSheetData(ss, sheetName) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { sheet: null, header: [], rows: [] };
  var lastRow = sheet.getLastRow();
  if (lastRow < 1) return { sheet: sheet, header: [], rows: [] };
  var values = sheet.getDataRange().getValues();
  return { sheet: sheet, header: values[0], rows: values.slice(1) };
}

function rowToObject(header, row) {
  var obj = {};
  for (var i = 0; i < header.length; i++) obj[String(header[i]).trim()] = row[i];
  return obj;
}

// パスワードをSHA-256でハッシュ化し、"sha256:"プレフィックス付きの16進文字列にする
function hashPasswordSha256(pw) {
  var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, pw, Utilities.Charset.UTF_8);
  var hex = digest.map(function(b) {
    var v = (b < 0) ? b + 256 : b;
    var h = v.toString(16);
    return h.length === 1 ? '0' + h : h;
  }).join('');
  return 'sha256:' + hex;
}

// 指定シートを、指定列（0始まり）のIDの昇順（数値として比較、数値化できない場合は文字列比較）に並べ替える
function sortSheetByIdColumn(ss, sheetName, idColIndex) {
  var sheet = ss.getSheetByName(sheetName);
  var data = getSheetData(ss, sheetName);
  if (data.rows.length < 2) return;
  data.rows.sort(function(a, b) {
    var idA = parseInt(String(a[idColIndex]).trim(), 10);
    var idB = parseInt(String(b[idColIndex]).trim(), 10);
    if (isNaN(idA) && isNaN(idB)) return String(a[idColIndex]).trim().localeCompare(String(b[idColIndex]).trim());
    if (isNaN(idA)) return 1;
    if (isNaN(idB)) return -1;
    return idA - idB;
  });
  setRangeAsText(sheet, 2, 1, data.rows.length, data.header.length, data.rows);
}
// 社員一覧・ユーザー情報の両シートを社員ID順に並べ替える
function sortEmployeeSheetById(ss) {
  sortSheetByIdColumn(ss, '社員一覧', 1);
  sortSheetByIdColumn(ss, 'ユーザー情報', 0);
}

// 社員一覧シートに既存登録されていない社員ID候補を、指定の番号より大きい範囲から昇順で最大5件返す
function findAvailableEmployeeIds(ss, afterId) {
  var data = getSheetData(ss, '社員一覧');
  var used = {};
  for (var i = 0; i < data.rows.length; i++) {
    used[String(data.rows[i][1]).trim()] = true;
  }
  var candidates = [];
  var n = afterId + 1;
  while (candidates.length < 5 && n <= 999) {
    var candId = ('000' + n).slice(-3);
    if (!used[candId]) candidates.push(candId);
    n++;
  }
  return candidates;
}

// 社員一覧シートから該当行を検索（1行目はヘッダー）
function findEmployeeRowSS(ss, targetId) {
  var data = getSheetData(ss, '社員一覧');
  for (var i = 0; i < data.rows.length; i++) {
    var row = data.rows[i];
    if (String(row[1]).trim() === targetId) {
      return {
        rowIndex: i + 2, role: String(row[0]).trim(), id: String(row[1]).trim(),
        password: String(row[2]).trim(), folderId: String(row[3]).trim(),
        name: row[4] ? String(row[4]).trim() : ''
      };
    }
  }
  return null;
}

function findUserInfoRowSS(ss, targetId) {
  var data = getSheetData(ss, 'ユーザー情報');
  for (var i = 0; i < data.rows.length; i++) {
    if (String(data.rows[i][0]).trim() === targetId) {
      return { rowIndex: i + 2, record: rowToObject(data.header, data.rows[i]) };
    }
  }
  return null;
}

function readCompanySettingsSS(ss) {
  var data = getSheetData(ss, '会社共通設定');
  var result = {};
  data.rows.forEach(function (row) {
    var key = String(row[0] || '').trim();
    if (key) result[key] = String(row[1] === undefined || row[1] === null ? '' : row[1]);
  });
  return result;
}

function readPaydayOverridesSS(ss) {
  var data = getSheetData(ss, '給料日特例');
  var result = {};
  data.rows.forEach(function (row) {
    var key = String(row[0] || '').trim();
    if (key) result[key] = String(row[1] === undefined || row[1] === null ? '' : row[1]);
  });
  return result;
}

function buildUserDetailsSS(ss, targetId) {
  var empRow = findEmployeeRowSS(ss, targetId);
  if (!empRow) return null;
  var userInfo = { sei: '', mei: '', seiKana: '', meiKana: '', dob: '', role: empRow.role, isExecutive: false, deleted: false, deletedDate: '', scheduledDate: '', deleteReason: '' };
  var userRow = findUserInfoRowSS(ss, targetId);
  if (userRow) {
    var rec = userRow.record;
    userInfo.sei = rec['苗字'] || '';
    userInfo.mei = rec['名前'] || '';
    userInfo.seiKana = rec['みょうじ'] || '';
    userInfo.meiKana = rec['なまえ'] || '';
    userInfo.dob = rec['生年月日'] ? String(rec['生年月日']) : '';
    userInfo.isExecutive = String(rec['役員フラグ']) === '1';
    userInfo.deleted = String(rec['削除フラグ']) === '1';
    userInfo.deletedDate = rec['削除日'] ? String(rec['削除日']) : '';
    userInfo.scheduledDate = rec['削除予定日'] ? String(rec['削除予定日']) : '';
    userInfo.deleteReason = rec['削除理由'] || '';
  }
  var byYear = {};
  var individualGrantDate = '';
  var individualGrantDays = '';
  try {
    var targetFolder = DriveApp.getFolderById(empRow.folderId);
    var settingsDocFile = getSettingsDocFile(targetFolder);
    if (settingsDocFile) {
      var settingsText = DocumentApp.openById(settingsDocFile.getId()).getBody().getText();
      var sLines = settingsText.split('\n');
      for (var yi = 0; yi < sLines.length; yi++) {
        var lineTrim = sLines[yi].trim();
        var yMatch = lineTrim.match(/^(\d{4})有給残:([\d.]+)$/);
        if (yMatch) byYear[yMatch[1]] = yMatch[2];
        if (lineTrim.indexOf('有給次回付与:') === 0) {
          var gd = lineTrim.slice('有給次回付与:'.length).trim();
          if (/^\d{8}$/.test(gd)) individualGrantDate = gd.slice(0,4) + '-' + gd.slice(4,6) + '-' + gd.slice(6,8);
        }
        if (lineTrim.indexOf('付与数:') === 0) {
          individualGrantDays = lineTrim.slice('付与数:'.length).trim();
        }
      }
    }
  } catch (errS) { /* フォルダが見つからない等は無視してbyYearは空のまま返す */ }
  return { id: targetId, folderId: empRow.folderId, userInfo: userInfo, byYear: byYear, individualGrantDate: individualGrantDate, individualGrantDays: individualGrantDays };
}

function doGet(e) {
  var action = e.parameter.action;
  var callback = e.parameter.callback;
  var out;
  try {
    if (action === 'list') {
      var folder = DriveApp.getFolderById(e.parameter.folder);
      var it = folder.getFiles();
      var arr = [];
      while (it.hasNext()) {
        var f = it.next();
        arr.push({id: f.getId(), name: f.getName(), mimeType: f.getMimeType()});
      }
      out = {files: arr};
    } else if (action === 'file') {
      var file = DriveApp.getFileById(e.parameter.id);
      var blob = file.getBlob();
      out = {name: file.getName(), mimeType: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes())};
    } else if (action === 'deleteFile') {
      // 明細PDFの差し替え時、既存ファイルをゴミ箱に移動するために使用
      var delFileId = (e.parameter.id || '').trim();
      if (!delFileId) {
        out = { error: 'ファイルIDを指定してください' };
      } else {
        try {
          DriveApp.getFileById(delFileId).setTrashed(true);
          out = { success: true };
        } catch (delFileErr) {
          out = { error: 'ファイルの削除に失敗しました：' + delFileErr.message };
        }
      }
    } else if (action === 'readSettings') {
      var folder2 = DriveApp.getFolderById(e.parameter.folder);
      var f2 = getSettingsDocFile(folder2);
      if (!f2) { out = {exists: false}; }
      else {
        var text = DocumentApp.openById(f2.getId()).getBody().getText();
        out = {exists: true, content: text};
      }
    } else if (action === 'writeSettings') {
      var empIdCheck = (e.parameter.empId || '').trim();
      var folderParam = (e.parameter.folder || '').trim();
      if (empIdCheck) {
        var verifySS = openMasterSpreadsheet();
        var verifyFolderId = null;
        var verifySourceFound = false;
        if (verifySS) {
          verifySourceFound = true;
          var verifyEmpRow = findEmployeeRowSS(verifySS, empIdCheck);
          verifyFolderId = verifyEmpRow ? verifyEmpRow.folderId : null;
        } else {
          var verifyListFile = findMasterListFile();
          if (verifyListFile) {
            verifySourceFound = true;
            var verifyDoc = DocumentApp.openById(verifyListFile.getId());
            var verifyMainText = verifyDoc.getTabs()[0].asDocumentTab().getBody().getText();
            var verifyRow = findMainListRow(verifyMainText, empIdCheck);
            verifyFolderId = verifyRow ? verifyRow.cols[3].trim() : null;
          }
        }
        if (verifySourceFound && verifyFolderId !== folderParam) {
          logSystemEvent('write_settings_blocked', '社員ID「' + empIdCheck + '」に紐づかないフォルダ（' + folderParam + '）への書き込みをブロックしました');
          out = { error: 'フォルダとユーザーの整合性が取れないため、書き込みを中止しました（安全のためのチェックです）。ページを再読み込みして、再度ログインしてから試してください。' };
        }
      }
      if (!out) {
        var folder3 = DriveApp.getFolderById(folderParam);
        var content = e.parameter.content || '';
        var f3 = getSettingsDocFile(folder3);
        var doc;
        if (f3) {
          doc = DocumentApp.openById(f3.getId());
        } else {
          var newDoc = DocumentApp.create(SETTINGS_DOC_NAME);
          var newFile = DriveApp.getFileById(newDoc.getId());
          folder3.addFile(newFile);
          DriveApp.getRootFolder().removeFile(newFile);
          doc = newDoc;
        }
        doc.getBody().editAsText().setText(content);
        doc.saveAndClose();
        out = {success: true};
      }
    } else if (action === 'sendInquiry') {
      var subject = e.parameter.subject || 'お問い合わせ';
      var body = e.parameter.body || '';
      var attachmentIdsRaw = (e.parameter.attachmentIds || '').trim();
      if (attachmentIdsRaw) {
        var attachmentIdList = attachmentIdsRaw.split(',').filter(function(s){ return s; });
        var attachmentBlobs = [];
        for (var ai = 0; ai < attachmentIdList.length && ai < 5; ai++) {
          try {
            var attachFile = DriveApp.getFileById(attachmentIdList[ai].trim());
            attachmentBlobs.push(attachFile.getBlob());
          } catch (attachErr) {
            // 取得できない添付は無視して続行
          }
        }
        MailApp.sendEmail(getInquiryEmail(), subject, body, { attachments: attachmentBlobs });
        for (var aj = 0; aj < attachmentIdList.length && aj < 5; aj++) {
          try {
            DriveApp.getFileById(attachmentIdList[aj].trim()).setTrashed(true);
          } catch (trashErr) {
            // 削除に失敗しても送信自体は成功しているため無視
          }
        }
      } else {
        MailApp.sendEmail(getInquiryEmail(), subject, body);
      }
      out = {success: true};
    } else if (action === 'sendLeaveRequestEmail') {
      var lrEmpName = e.parameter.empName || '';
      var lrDate = e.parameter.date || '';
      var lrReason = e.parameter.reason || '（未入力）';
      var lrSubject = '給与台帳：有給申請';
      var lrBody = lrEmpName + 'さんから有給申請が届きました。' + lrDate + '。理由：' + lrReason + 'のため。内容を確認し、アプリより、有給許可を行ってください。\n※本メールは給与台帳より自動送信されています。';
      MailApp.sendEmail(getInquiryEmail(), lrSubject, lrBody);
      out = {success: true};
    } else if (action === 'uploadChunkStart') {
      var upFolder = DriveApp.getFolderById(e.parameter.folder);
      var tempDoc = DocumentApp.create('_upload_tmp_' + new Date().getTime());
      var tempFile = DriveApp.getFileById(tempDoc.getId());
      upFolder.addFile(tempFile);
      DriveApp.getRootFolder().removeFile(tempFile);
      out = { success: true, tempDocId: tempDoc.getId() };
    } else if (action === 'uploadChunkAppend') {
      var appendDoc = DocumentApp.openById(e.parameter.tempDocId);
      appendDoc.getBody().appendParagraph(e.parameter.data || '');
      appendDoc.saveAndClose();
      out = { success: true };
    } else if (action === 'uploadChunkFinish') {
      var finishDoc = DocumentApp.openById(e.parameter.tempDocId);
      var fullText = finishDoc.getBody().getText().replace(/\n/g, '').replace(/\s/g, '');
      var targetFolder = DriveApp.getFolderById(e.parameter.folder);
      var fileName = e.parameter.filename || 'アップロードファイル';
      var mimeType = e.parameter.mimeType || 'application/pdf';
      var bytes = Utilities.base64Decode(fullText);
      var blob = Utilities.newBlob(bytes, mimeType, fileName);
      var newFile = targetFolder.createFile(blob);
      DriveApp.getFileById(finishDoc.getId()).setTrashed(true);
      out = { success: true, fileId: newFile.getId(), fileName: newFile.getName() };
    } else if (action === 'getAppIcon') {
      var masterFolderIcon = DriveApp.getFolderById(MASTER_FOLDER_ID);
      var iconFiles = masterFolderIcon.getFilesByName('icon.ico');
      var iconFile = null;
      if (iconFiles.hasNext()) {
        iconFile = iconFiles.next();
      } else {
        var iconFilesPng = masterFolderIcon.getFilesByName('icon.png');
        if (iconFilesPng.hasNext()) iconFile = iconFilesPng.next();
      }
      if (!iconFile) {
        out = { exists: false };
      } else {
        var iconBlob = iconFile.getBlob();
        out = { exists: true, mimeType: iconBlob.getContentType(), base64: Utilities.base64Encode(iconBlob.getBytes()) };
      }
    } else if (action === 'getInquiryContact') {
      out = { success: true, email: getInquiryEmail(), name: getInquiryName() };
    } else if (action === 'updateInquiryContact') {
      var newEmail = (e.parameter.newEmail || '').trim();
      var newName = (e.parameter.newName || '').trim();
      if (!newEmail) {
        out = { error: 'メールアドレスを入力してください' };
      } else {
        PropertiesService.getScriptProperties().setProperty('INQUIRY_EMAIL', newEmail);
        PropertiesService.getScriptProperties().setProperty('INQUIRY_NAME', newName || DEFAULT_INQUIRY_NAME);
        out = { success: true };
      }
    } else if (action === 'getTransportContact') {
      out = { success: true, email: getTransportEmail() };
    } else if (action === 'updateTransportContact') {
      var newTransportEmail = (e.parameter.newEmail || '').trim();
      if (!newTransportEmail) {
        out = { error: 'メールアドレスを入力してください' };
      } else {
        PropertiesService.getScriptProperties().setProperty('TRANSPORT_EMAIL', newTransportEmail);
        out = { success: true };
      }
    } else if (action === 'getLeaveTypeContact') {
      out = { success: true, email: getLeaveTypeEmail() };
    } else if (action === 'updateLeaveTypeContact') {
      var newLeaveTypeEmail = (e.parameter.newEmail || '').trim();
      if (!newLeaveTypeEmail) {
        out = { error: 'メールアドレスを入力してください' };
      } else {
        PropertiesService.getScriptProperties().setProperty('LEAVE_TYPE_EMAIL', newLeaveTypeEmail);
        out = { success: true };
      }
    } else if (action === 'companySettings') {
      var ss0 = openMasterSpreadsheet();
      if (!ss0) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        out = { success: true, companySettings: readCompanySettingsSS(ss0), paydayOverrides: readPaydayOverridesSS(ss0) };
      }
    } else if (action === 'updatePaydayOverride') {
      var ssP = openMasterSpreadsheet();
      if (!ssP) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var ymP = (e.parameter.ym || '').trim();
        var dateP = (e.parameter.date || '').trim(); // 空なら特例を削除
        var sheetDataP = getSheetData(ssP, '給料日特例');
        var sheetP = sheetDataP.sheet;
        if (!sheetP) {
          sheetP = ssP.insertSheet('給料日特例');
          sheetP.appendRow(['対象年月', '給料日']);
          sheetDataP = getSheetData(ssP, '給料日特例');
        }
        var foundRowIdxP = -1;
        for (var pi = 0; pi < sheetDataP.rows.length; pi++) {
          if (String(sheetDataP.rows[pi][0]).trim() === ymP) { foundRowIdxP = pi + 2; break; }
        }
        if (dateP) {
          if (foundRowIdxP > 0) {
            setCellAsText(sheetP, foundRowIdxP, 2, dateP);
          } else {
            appendRowAsText(sheetP, [ymP, dateP]);
          }
        } else if (foundRowIdxP > 0) {
          sheetP.deleteRow(foundRowIdxP);
        }
        out = { success: true };
      }
    } else if (action === 'findUserByInfo') {
      var ssU = openMasterSpreadsheet();
      if (!ssU) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var dataU = getSheetData(ssU, 'ユーザー情報');
        var inSei = (e.parameter.sei || '').trim();
        var inMei = (e.parameter.mei || '').trim();
        var inDob = digitsOnly(e.parameter.dob || '');
        var matchedId = null;
        for (var ri = 0; ri < dataU.rows.length; ri++) {
          var r = rowToObject(dataU.header, dataU.rows[ri]);
          if (String(r['苗字'] || '') === inSei && String(r['名前'] || '') === inMei && digitsOnly(String(r['生年月日'] || '')) === inDob) {
            matchedId = String(r['ユーザーID']);
            break;
          }
        }
        if (matchedId) {
          out = { success: true, userId: matchedId };
        } else {
          out = { error: '一致する情報が見つかりませんでした' };
        }
      }
    } else if (action === 'resetPassword') {
      var ssR = openMasterSpreadsheet();
      if (!ssR) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetId = (e.parameter.id || '').trim();
        var newPw = (e.parameter.newPassword || '').trim().replace(/:/g, '');
        var empRowR = findEmployeeRowSS(ssR, targetId);
        if (empRowR) {
          ssR.getSheetByName('社員一覧').getRange(empRowR.rowIndex, 3).setValue(hashPasswordSha256(newPw));
          out = { success: true };
        } else {
          out = { error: '該当するユーザーIDが見つかりませんでした' };
        }
      }
    } else if (action === 'postNews') {
      var newsDocP = getOrCreateMasterDoc('ニュース');
      var newsId = new Date().getTime();
      var scopeParamP = e.parameter.scope || 'all';
      var scope = (scopeParamP === 'admin' || scopeParamP.indexOf('emp:') === 0) ? scopeParamP : 'all';
      var priorityP = e.parameter.priority || 'normal';
      var titleP = (e.parameter.title || '').replace(/\n/g, ' ').replace(/:/g, '：');
      var msg = (e.parameter.message || '').replace(/\n/g, ' ').replace(/:/g, '：');
      appendLine(newsDocP, [newsId, scope, priorityP, titleP, '', '', msg].join(':'));
      out = { success: true };
    } else if (action === 'checkAppVersion') {
      var htmlVersionCV = (e.parameter.htmlVersion || '').trim();
      var lockCV = LockService.getScriptLock();
      lockCV.waitLock(10000);
      try {
        var propsCV = PropertiesService.getScriptProperties();
        var lastGasCV = propsCV.getProperty('LAST_NOTIFIED_GAS_VERSION');
        var lastHtmlCV = propsCV.getProperty('LAST_NOTIFIED_HTML_VERSION');
        var changedPartsCV = [];
        if (lastGasCV === null) {
          propsCV.setProperty('LAST_NOTIFIED_GAS_VERSION', APP_VERSION_GAS);
        } else if (lastGasCV !== APP_VERSION_GAS) {
          changedPartsCV.push('サーバー側（GAS）');
          propsCV.setProperty('LAST_NOTIFIED_GAS_VERSION', APP_VERSION_GAS);
        }
        if (htmlVersionCV) {
          if (lastHtmlCV === null) {
            propsCV.setProperty('LAST_NOTIFIED_HTML_VERSION', htmlVersionCV);
          } else if (lastHtmlCV !== htmlVersionCV) {
            changedPartsCV.push('画面（アプリ本体）');
            propsCV.setProperty('LAST_NOTIFIED_HTML_VERSION', htmlVersionCV);
          }
        }
        if (changedPartsCV.length > 0) {
          var newsDocCV = getOrCreateMasterDoc('ニュース');
          var newsIdCV = new Date().getTime();
          var titleCV = 'システムが更新されました';
          var msgCV = changedPartsCV.join('・') + 'が更新されました。表示がおかしい場合は再読み込みをお試しください。';
          appendLine(newsDocCV, [newsIdCV, 'all', 'maintenance', titleCV, '', '', msgCV].join(':'));
          out = { success: true, notified: true };
        } else {
          out = { success: true, notified: false };
        }
      } finally {
        lockCV.releaseLock();
      }
    } else if (action === 'getNews') {
      var newsDocG = getOrCreateMasterDoc('ニュース');
      var newsLines = newsDocG.getBody().getText().split('\n');
      var newsList = [];
      var includeDeletedG = e.parameter.includeDeleted === '1';
      for (var ni = 0; ni < newsLines.length; ni++) {
        var nrow = newsLines[ni].trim();
        if (!nrow) continue;
        var nParts = nrow.split(':');
        if (nParts.length < 3) continue;
        var itemG;
        if (nParts.length >= 7) {
          // 新形式（編集日時・削除日時つき）
          itemG = { id: nParts[0], scope: nParts[1], priority: nParts[2], title: nParts[3], editedAt: nParts[4], deletedAt: nParts[5], message: nParts.slice(6).join(':') };
        } else if (nParts.length >= 5) {
          itemG = { id: nParts[0], scope: nParts[1], priority: nParts[2], title: nParts[3], editedAt: '', deletedAt: '', message: nParts.slice(4).join(':') };
        } else if (nParts.length === 4) {
          // 旧形式（重要度なし・タイトルあり）との互換
          itemG = { id: nParts[0], scope: nParts[1], priority: 'normal', title: nParts[2], editedAt: '', deletedAt: '', message: nParts.slice(3).join(':') };
        } else {
          // 最旧形式（タイトルなし）との互換
          itemG = { id: nParts[0], scope: nParts[1], priority: 'normal', title: nParts.slice(2).join(':'), editedAt: '', deletedAt: '', message: '' };
        }
        if (itemG.deletedAt && !includeDeletedG) continue;
        newsList.push(itemG);
      }
      out = { success: true, news: newsList };
    } else if (action === 'editNews' || action === 'deleteNews' || action === 'restoreNews') {
      var newsDocX = getOrCreateMasterDoc('ニュース');
      var bodyX = newsDocX.getBody();
      var linesX = bodyX.getText().split('\n');
      var idX = (e.parameter.id || '').trim();
      var foundX = false;
      for (var xi = 0; xi < linesX.length; xi++) {
        var rowX = linesX[xi].trim();
        if (!rowX) continue;
        var partsX = rowX.split(':');
        if (partsX[0] !== idX) continue;
        var isNewFormatX = partsX.length >= 7;
        var scopeX = partsX[1] || 'all';
        var priorityX = partsX[2] || 'normal';
        var titleX = isNewFormatX ? partsX[3] : (partsX.length >= 5 ? partsX[3] : (partsX.length === 4 ? partsX[2] : ''));
        var editedAtX = isNewFormatX ? partsX[4] : '';
        var deletedAtX = isNewFormatX ? partsX[5] : '';
        var msgX = isNewFormatX ? partsX.slice(6).join(':') : (partsX.length >= 5 ? partsX.slice(4).join(':') : '');
        if (action === 'editNews') {
          priorityX = e.parameter.priority || priorityX;
          titleX = (e.parameter.title || '').replace(/\n/g, ' ').replace(/:/g, '：');
          msgX = (e.parameter.message || '').replace(/\n/g, ' ').replace(/:/g, '：');
          editedAtX = String(new Date().getTime());
        } else if (action === 'deleteNews') {
          if (priorityX !== 'critical') {
            // 重要ニュース以外は復元不要のため完全削除する
            linesX.splice(xi, 1);
            foundX = true;
            break;
          }
          deletedAtX = String(new Date().getTime());
        } else if (action === 'restoreNews') {
          deletedAtX = '';
        }
        linesX[xi] = [idX, scopeX, priorityX, titleX, editedAtX, deletedAtX, msgX].join(':');
        foundX = true;
        break;
      }
      if (!foundX) {
        out = { error: '対象のニュースが見つかりませんでした' };
      } else {
        bodyX.editAsText().setText(linesX.join('\n'));
        newsDocX.saveAndClose();
        out = { success: true };
      }
    } else if (action === 'bulkDeleteNews') {
      var newsDocBD = getOrCreateMasterDoc('ニュース');
      var bodyBD = newsDocBD.getBody();
      var linesBD = bodyBD.getText().split('\n');
      var idsBD = (e.parameter.ids || '').split(',').map(function (s) { return s.trim(); }).filter(function (s) { return s; });
      var idsSetBD = {};
      idsBD.forEach(function (id) { idsSetBD[id] = true; });
      var newLinesBD = [];
      var deletedCountBD = 0;
      for (var bi = 0; bi < linesBD.length; bi++) {
        var rowBD = linesBD[bi].trim();
        if (!rowBD) continue;
        var idBD = rowBD.split(':')[0];
        if (idsSetBD[idBD]) {
          // 一括削除は復元不可の完全削除
          deletedCountBD++;
          continue;
        }
        newLinesBD.push(rowBD);
      }
      bodyBD.editAsText().setText(newLinesBD.join('\n'));
      newsDocBD.saveAndClose();
      out = { success: true, deletedCount: deletedCountBD };
    } else if (action === 'logEvent') {
      var logDoc = getOrCreateMasterDoc('操作ログ');
      var logLine = new Date().toISOString() + ':' + (e.parameter.empId || '') + ':' + (e.parameter.empName || '') + ':' + (e.parameter.eventType || '') + ':' + (e.parameter.detail || '').replace(/\n/g, ' ');
      appendLine(logDoc, logLine);
      out = { success: true };
    } else if (action === 'getLogs') {
      var logDocG = getOrCreateMasterDoc('操作ログ');
      var logLines = logDocG.getBody().getText().split('\n').filter(function(l){ return l.trim(); });
      var recent = logLines.slice(Math.max(0, logLines.length - 200));
      out = { success: true, logs: recent };
    } else if (action === 'adminFindUsersByQuery') {
      var ssQ = openMasterSpreadsheet();
      if (!ssQ) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var queryQ = (e.parameter.query || '').trim();
        var byIdQ = findEmployeeRowSS(ssQ, queryQ);
        if (byIdQ) {
          var detailsQ = buildUserDetailsSS(ssQ, queryQ);
          out = { success: true, mode: 'single', detail: detailsQ };
        } else {
          var dataQ = getSheetData(ssQ, 'ユーザー情報');
          var matchesQ = [];
          for (var qi = 0; qi < dataQ.rows.length; qi++) {
            var recQ = rowToObject(dataQ.header, dataQ.rows[qi]);
            if (String(recQ['みょうじ'] || '') === queryQ) {
              matchesQ.push({
                id: String(recQ['ユーザーID']),
                sei: recQ['苗字'] || '',
                mei: recQ['名前'] || '',
                seiKana: recQ['みょうじ'] || '',
                meiKana: recQ['なまえ'] || '',
                deleted: String(recQ['削除フラグ']) === '1'
              });
            }
          }
          if (matchesQ.length === 0) {
            out = { error: '該当するユーザーが見つかりませんでした' };
          } else if (matchesQ.length === 1) {
            var detailsQ1 = buildUserDetailsSS(ssQ, matchesQ[0].id);
            out = { success: true, mode: 'single', detail: detailsQ1 };
          } else {
            out = { success: true, mode: 'multiple', matches: matchesQ };
          }
        }
      }
    } else if (action === 'adminDeleteUser') {
      var ssDel = openMasterSpreadsheet();
      if (!ssDel) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdDel = (e.parameter.targetId || '').trim();
        var reasonDel = (e.parameter.reason || '').trim();
        var userRowDel = findUserInfoRowSS(ssDel, targetIdDel);
        var todayDel = todayDateString();
        var scheduledDel = addDaysToDateString(todayDel, getRetentionDays());
        var sheetUserDel = ssDel.getSheetByName('ユーザー情報');
        if (userRowDel) {
          var hIdx = {};
          getSheetData(ssDel, 'ユーザー情報').header.forEach(function (h, i) { hIdx[String(h).trim()] = i + 1; });
          sheetUserDel.getRange(userRowDel.rowIndex, hIdx['削除フラグ']).setValue('1');
          setCellAsText(sheetUserDel, userRowDel.rowIndex, hIdx['削除日'], todayDel);
          setCellAsText(sheetUserDel, userRowDel.rowIndex, hIdx['削除予定日'], scheduledDel);
          sheetUserDel.getRange(userRowDel.rowIndex, hIdx['削除理由']).setValue(reasonDel || '（理由未入力）');
          sheetUserDel.getRange(userRowDel.rowIndex, hIdx['告知済み']).setValue('0');
        } else if (findEmployeeRowSS(ssDel, targetIdDel)) {
          // ユーザー情報行が無い社員一覧のみの社員も削除予定にできるよう、新規に行を追加する
          appendRowAsText(sheetUserDel, [targetIdDel, '', '', '', '', '', '', '', '1', todayDel, scheduledDel, reasonDel || '（理由未入力）', '0']);
        } else {
          out = { error: '該当するユーザーが見つかりませんでした' };
        }
        if (!out) {
          logSystemEvent('user_delete_scheduled', targetIdDel + ' を削除予定に設定（完全削除予定日：' + scheduledDel + '、理由：' + reasonDel + '）');
          out = { success: true, scheduledDate: scheduledDel };
        }
      }
    } else if (action === 'adminRestoreUser') {
      var ssRes = openMasterSpreadsheet();
      if (!ssRes) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdRes = (e.parameter.targetId || '').trim();
        var reasonRes = (e.parameter.reason || '').trim();
        var userRowRes = findUserInfoRowSS(ssRes, targetIdRes);
        if (!userRowRes) {
          out = { error: '該当するユーザーが見つかりませんでした' };
        } else {
          var hIdxRes = {};
          getSheetData(ssRes, 'ユーザー情報').header.forEach(function (h, i) { hIdxRes[String(h).trim()] = i + 1; });
          var sheetUserRes = ssRes.getSheetByName('ユーザー情報');
          ['削除フラグ', '削除日', '削除予定日', '削除理由', '告知済み'].forEach(function (k) {
            if (hIdxRes[k]) sheetUserRes.getRange(userRowRes.rowIndex, hIdxRes[k]).setValue('');
          });
          logSystemEvent('user_restore', targetIdRes + ' の削除を取り消しました（理由：' + reasonRes + '）');
          out = { success: true };
        }
      }
    } else if (action === 'adminCreateUser') {
      var ssC = openMasterSpreadsheet();
      if (!ssC) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var newIdC = (e.parameter.newId || '').trim();
        var newPwC = (e.parameter.newPassword || '').trim().replace(/:/g, '');
        var newRoleC = (e.parameter.role || '0').trim();
        var newSeiC = (e.parameter.sei || '').trim();
        var newMeiC = (e.parameter.mei || '').trim();
        var newSeiKanaC = (e.parameter.seiKana || '').trim();
        var newMeiKanaC = (e.parameter.meiKana || '').trim();
        var newDobC = (e.parameter.dob || '').trim();
        var newBaseSalaryC = (e.parameter.baseSalary || '0').trim();
        var newIsExecutiveC = (e.parameter.isExecutive || '0').trim();
        if (!newIdC || !newPwC) {
          out = { error: '社員IDとパスワードを入力してください' };
        } else if (findEmployeeRowSS(ssC, newIdC)) {
          out = { error: 'その社員IDはすでに使われています' };
        } else {
          var parentFolderC = DriveApp.getFolderById(getEmployeeParentId());
          var newEmpFolderC = parentFolderC.createFolder(newSeiC + newMeiC || newIdC);
          var newSettingsDocC = DocumentApp.create(SETTINGS_DOC_NAME);
          var newSettingsFileC = DriveApp.getFileById(newSettingsDocC.getId());
          newEmpFolderC.addFile(newSettingsFileC);
          DriveApp.getRootFolder().removeFile(newSettingsFileC);
          var curYearC = new Date().getFullYear();
          var initialTextC = [
            '基本給:' + newBaseSalaryC,
            '',
            '残業時間履歴',
            '',
            '給与実績履歴',
            '',
            curYearC + '有給残:0',
            '',
            '有給使用履歴'
          ].join('\n');
          newSettingsDocC.getBody().editAsText().setText(initialTextC);
          newSettingsDocC.saveAndClose();

          appendRowAsText(ssC.getSheetByName('社員一覧'), [newRoleC, newIdC, hashPasswordSha256(newPwC), newEmpFolderC.getId(), newSeiC + newMeiC]);
          appendRowAsText(ssC.getSheetByName('ユーザー情報'), [newIdC, newSeiC, newMeiC, newSeiKanaC, newMeiKanaC, newDobC, newRoleC, newIsExecutiveC, '', '', '', '', '']);
          sortEmployeeSheetById(ssC);

          out = { success: true, id: newIdC, folderId: newEmpFolderC.getId() };
        }
      }
    } else if (action === 'adminAvailableIds') {
      var ssAvail = openMasterSpreadsheet();
      if (!ssAvail) {
        out = { error: '「管理情報」スプレッドシートが見つかりません' };
      } else {
        var afterIdAvail = parseInt(e.parameter.afterId || '0', 10);
        if (isNaN(afterIdAvail) || afterIdAvail < 0) afterIdAvail = 0;
        out = { success: true, ids: findAvailableEmployeeIds(ssAvail, afterIdAvail) };
      }
    } else if (action === 'adminSearchUser') {
      var ssS = openMasterSpreadsheet();
      if (!ssS) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdS = (e.parameter.targetId || '').trim();
        var detailsS = buildUserDetailsSS(ssS, targetIdS);
        if (!detailsS) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          out = { success: true, id: detailsS.id, folderId: detailsS.folderId, userInfo: detailsS.userInfo, byYear: detailsS.byYear, individualGrantDate: detailsS.individualGrantDate, individualGrantDays: detailsS.individualGrantDays };
        }
      }
    } else if (action === 'adminUpdateUserInfo') {
      var ssI = openMasterSpreadsheet();
      if (!ssI) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdI = (e.parameter.targetId || '').trim();
        var newRole = (e.parameter.role || '0').trim();
        var empRowI = findEmployeeRowSS(ssI, targetIdI);
        if (!empRowI) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          ssI.getSheetByName('社員一覧').getRange(empRowI.rowIndex, 1).setValue(newRole);
          var userRowI = findUserInfoRowSS(ssI, targetIdI);
          var valsI = [targetIdI, e.parameter.sei || '', e.parameter.mei || '', e.parameter.seiKana || '', e.parameter.meiKana || '', e.parameter.dob || '', newRole, e.parameter.isExecutive || '0', '', '', '', '', ''];
          if (userRowI) {
            var hIdxI = {};
            getSheetData(ssI, 'ユーザー情報').header.forEach(function (h, i) { hIdxI[String(h).trim()] = i + 1; });
            var sheetUserI = ssI.getSheetByName('ユーザー情報');
            sheetUserI.getRange(userRowI.rowIndex, hIdxI['苗字']).setValue(e.parameter.sei || '');
            sheetUserI.getRange(userRowI.rowIndex, hIdxI['名前']).setValue(e.parameter.mei || '');
            sheetUserI.getRange(userRowI.rowIndex, hIdxI['みょうじ']).setValue(e.parameter.seiKana || '');
            sheetUserI.getRange(userRowI.rowIndex, hIdxI['なまえ']).setValue(e.parameter.meiKana || '');
            setCellAsText(sheetUserI, userRowI.rowIndex, hIdxI['生年月日'], e.parameter.dob || '');
            sheetUserI.getRange(userRowI.rowIndex, hIdxI['管理者権限']).setValue(newRole);
            sheetUserI.getRange(userRowI.rowIndex, hIdxI['役員フラグ']).setValue(e.parameter.isExecutive || '0');
          } else {
            appendRowAsText(ssI.getSheetByName('ユーザー情報'), valsI);
          }
          out = { success: true };
        }
      }
    } else if (action === 'adminUpdateLeaveBalance') {
      var ssL = openMasterSpreadsheet();
      if (!ssL) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdL = (e.parameter.targetId || '').trim();
        var rowL = findEmployeeRowSS(ssL, targetIdL);
        if (!rowL) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          var yearL = (e.parameter.year || '').trim();
          var newValueL = (e.parameter.newValue || '0').trim();
          var targetFolderL = DriveApp.getFolderById(rowL.folderId);
          var settingsDocFileL = getSettingsDocFile(targetFolderL);
          var settingsDocL;
          if (settingsDocFileL) {
            settingsDocL = DocumentApp.openById(settingsDocFileL.getId());
          } else {
            settingsDocL = DocumentApp.create(SETTINGS_DOC_NAME);
            var newFileL = DriveApp.getFileById(settingsDocL.getId());
            targetFolderL.addFile(newFileL);
            DriveApp.getRootFolder().removeFile(newFileL);
          }
          var bodyL = settingsDocL.getBody();
          var linesL = bodyL.getText().split('\n');
          var foundL = false;
          for (var li = 0; li < linesL.length; li++) {
            if (linesL[li].trim() === '' ) continue;
            var lm = linesL[li].trim().match(/^(\d{4})有給残:([\d.]+)$/);
            if (lm && lm[1] === yearL) {
              linesL[li] = yearL + '有給残:' + newValueL;
              foundL = true;
              break;
            }
          }
          if (!foundL) {
            linesL.push(yearL + '有給残:' + newValueL);
          }
          bodyL.editAsText().setText(linesL.join('\n'));
          settingsDocL.saveAndClose();
          out = { success: true };
        }
      }
    } else if (action === 'adminUpdateGrantDate') {
      var ssGD = openMasterSpreadsheet();
      if (!ssGD) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdGD = (e.parameter.targetId || '').trim();
        var rowGD = findEmployeeRowSS(ssGD, targetIdGD);
        if (!rowGD) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          var newDateGD = (e.parameter.newDate || '').trim();
          var newDaysGD = (e.parameter.newDays || '').trim();
          if (!/^\d{4}-\d{2}-\d{2}$/.test(newDateGD)) {
            out = { error: '日付の形式が正しくありません' };
          } else {
            var targetFolderGD = DriveApp.getFolderById(rowGD.folderId);
            var settingsDocFileGD = getSettingsDocFile(targetFolderGD);
            var settingsDocGD;
            if (settingsDocFileGD) {
              settingsDocGD = DocumentApp.openById(settingsDocFileGD.getId());
            } else {
              settingsDocGD = DocumentApp.create(SETTINGS_DOC_NAME);
              var newFileGD = DriveApp.getFileById(settingsDocGD.getId());
              targetFolderGD.addFile(newFileGD);
              DriveApp.getRootFolder().removeFile(newFileGD);
            }
            var bodyGD = settingsDocGD.getBody();
            var linesGD = bodyGD.getText().split('\n');
            var dateLineGD = '有給次回付与:' + newDateGD.replace(/-/g, '');
            var foundDateGD = false;
            for (var gi = 0; gi < linesGD.length; gi++) {
              if (linesGD[gi].trim().indexOf('有給次回付与:') === 0) {
                linesGD[gi] = dateLineGD;
                foundDateGD = true;
                break;
              }
            }
            if (!foundDateGD) linesGD.push(dateLineGD);
            if (newDaysGD !== '' && !isNaN(parseFloat(newDaysGD))) {
              var daysLineGD = '付与数:' + newDaysGD;
              var foundDaysGD = false;
              for (var gj = 0; gj < linesGD.length; gj++) {
                if (linesGD[gj].trim().indexOf('付与数:') === 0) {
                  linesGD[gj] = daysLineGD;
                  foundDaysGD = true;
                  break;
                }
              }
              if (!foundDaysGD) linesGD.push(daysLineGD);
            }
            bodyGD.editAsText().setText(linesGD.join('\n'));
            settingsDocGD.saveAndClose();
            out = { success: true };
          }
        }
      }
    } else if (action === 'adminResolveLeaveRequest') {
      var ssR = openMasterSpreadsheet();
      if (!ssR) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var targetIdR = (e.parameter.targetId || '').trim();
        var rowR = findEmployeeRowSS(ssR, targetIdR);
        if (!rowR) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          var requestIdR = (e.parameter.requestId || '').trim();
          var typeR = e.parameter.type === 'delete' ? 'delete' : 'use';
          var decisionR = e.parameter.decision === 'approve' ? 'approve' : 'reject';
          var reasonR = (e.parameter.reason || '').replace(/\n/g, ' ');
          var targetFolderR = DriveApp.getFolderById(rowR.folderId);
          var settingsDocFileR = getSettingsDocFile(targetFolderR);
          if (!settingsDocFileR) {
            out = { error: '対象ユーザーの設定ドキュメントが見つかりません' };
          } else {
            var settingsDocR = DocumentApp.openById(settingsDocFileR.getId());
            var bodyR = settingsDocR.getBody();
            var linesR = bodyR.getText().split('\n');
            var byYearR = {};
            for (var yi = 0; yi < linesR.length; yi++) {
              var ym = linesR[yi].trim().match(/^(\d{4})有給残:([\d.]+)$/);
              if (ym) byYearR[ym[1]] = parseFloat(ym[2]);
            }
            var usageStartR = -1, usageEndR = linesR.length;
            for (var si = 0; si < linesR.length; si++) {
              if (linesR[si].trim() === '有給使用履歴') { usageStartR = si + 1; break; }
            }
            if (usageStartR === -1) {
              out = { error: '対象の申請が見つかりませんでした' };
            } else {
              for (var ei = usageStartR; ei < linesR.length; ei++) {
                if (linesR[ei].trim() === '') { usageEndR = ei; break; }
              }
              var markerR = typeR === 'delete' ? 'PENDING_DELETE' : 'PENDING_USE';
              var suffixR = '｜' + markerR + '｜' + requestIdR;
              var targetLineIdx = -1;
              for (var ui = usageStartR; ui < usageEndR; ui++) {
                if (linesR[ui].indexOf(suffixR) !== -1) { targetLineIdx = ui; break; }
              }
              if (targetLineIdx === -1) {
                out = { error: '対象の申請が見つかりませんでした（既に処理済みの可能性があります）' };
              } else {
                var lineM = linesR[targetLineIdx].trim().match(/^(\d{8}):(-?\d+(?:\.\d+)?):([^@]*)(?:@(.*))?$/);
                if (!lineM) {
                  out = { error: '申請データの形式が正しくありません' };
                } else {
                  var dateStrR = lineM[1];
                  var amountR = parseFloat(lineM[2]);
                  var noteWithMarkerR = lineM[3];
                  var consumedStrR = lineM[4] || '';
                  var cleanNoteR = noteWithMarkerR.split(suffixR).join('');
                  var consumedListR = [];
                  if (consumedStrR) {
                    var consumedPartsR = consumedStrR.split(',');
                    for (var pi = 0; pi < consumedPartsR.length; pi++) {
                      var kv = consumedPartsR[pi].split('=');
                      if (kv.length === 2) consumedListR.push({year: kv[0], amount: parseFloat(kv[1])});
                    }
                  }
                  var restoreBalanceR = function(){
                    if (consumedListR.length > 0) {
                      for (var ci = 0; ci < consumedListR.length; ci++) {
                        var cy = consumedListR[ci].year;
                        byYearR[cy] = Math.round(((byYearR[cy] || 0) + consumedListR[ci].amount) * 10) / 10;
                      }
                    } else {
                      var fy = String(parseInt(dateStrR.substring(0, 4)));
                      byYearR[fy] = Math.round(((byYearR[fy] || 0) + amountR) * 10) / 10;
                    }
                  };
                  if (typeR === 'use') {
                    if (decisionR === 'approve') {
                      linesR[targetLineIdx] = dateStrR + ':' + amountR + ':' + cleanNoteR + (consumedStrR ? '@' + consumedStrR : '');
                    } else {
                      restoreBalanceR();
                      linesR[targetLineIdx] = dateStrR + ':' + amountR + ':不許可：' + reasonR;
                    }
                  } else {
                    if (decisionR === 'approve') {
                      restoreBalanceR();
                      linesR.splice(targetLineIdx, 1);
                    } else {
                      linesR[targetLineIdx] = dateStrR + ':' + amountR + ':' + cleanNoteR + (consumedStrR ? '@' + consumedStrR : '');
                    }
                  }
                  for (var wi = 0; wi < linesR.length; wi++) {
                    var wm = linesR[wi].trim().match(/^(\d{4})有給残:([\d.]+)$/);
                    if (wm && byYearR.hasOwnProperty(wm[1])) {
                      linesR[wi] = wm[1] + '有給残:' + byYearR[wm[1]];
                    }
                  }
                  bodyR.editAsText().setText(linesR.join('\n'));
                  settingsDocR.saveAndClose();
                  out = { success: true };
                }
              }
            }
          }
        }
      }
    } else if (action === 'updateCompanySettings') {
      var ssCS = openMasterSpreadsheet();
      if (!ssCS) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var updatesCS = {};
        try {
          updatesCS = JSON.parse(e.parameter.updates || '{}');
        } catch (parseErrCS) {
          updatesCS = {};
        }
        var dataCS = getSheetData(ssCS, '会社共通設定');
        var sheetCS = dataCS.sheet;
        if (!sheetCS) {
          sheetCS = ssCS.insertSheet('会社共通設定');
          appendRowAsText(sheetCS, ['設定項目', '値']);
          dataCS = getSheetData(ssCS, '会社共通設定');
        }
        var rowIdxByKeyCS = {};
        dataCS.rows.forEach(function (row, i) { rowIdxByKeyCS[String(row[0]).trim()] = i + 2; });
        for (var ukCS in updatesCS) {
          if (rowIdxByKeyCS[ukCS]) {
            setCellAsText(sheetCS, rowIdxByKeyCS[ukCS], 2, updatesCS[ukCS]);
          } else {
            appendRowAsText(sheetCS, [ukCS, updatesCS[ukCS]]);
            rowIdxByKeyCS[ukCS] = sheetCS.getLastRow();
          }
        }
        out = { success: true };
      }
    } else if (action === 'getRetentionDays') {
      out = { success: true, retentionDays: getRetentionDays() };
    } else if (action === 'updateRetentionDays') {
      var newDaysR = parseInt(e.parameter.days || '0');
      if (!newDaysR || newDaysR <= 0) {
        out = { error: '正しい日数を入力してください' };
      } else {
        PropertiesService.getScriptProperties().setProperty('RETENTION_DAYS', String(newDaysR));
        out = { success: true };
      }
    } else if (action === 'getLeaveTypes') {
      var ssLT = openMasterSpreadsheet();
      if (!ssLT) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var sheetLT = getOrCreateLeaveTypesSheet(ssLT);
        var dataLT = getSheetData(ssLT, '休暇制度');
        var typesLT = dataLT.rows.map(function (row) {
          return {
            id: String(row[0]), name: row[1] || '', pay: row[2] || '', condition: row[3] || '',
            applyEnabled: String(row[4]) === '1', url: row[5] || '',
            documents: row[6] ? String(row[6]).split('|').filter(function (d) { return d; }) : []
          };
        });
        out = { success: true, types: typesLT };
      }
    } else if (action === 'adminAddLeaveType' || action === 'adminUpdateLeaveType') {
      var ssLA = openMasterSpreadsheet();
      if (!ssLA) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var nameLA = (e.parameter.name || '').trim();
        var payLA = (e.parameter.pay || '').trim();
        var conditionLA = (e.parameter.condition || '').trim();
        var applyEnabledLA = e.parameter.applyEnabled === '1' ? '1' : '0';
        var urlLA = (e.parameter.url || '').trim();
        var documentsLA = (e.parameter.documents || '').trim();
        if (!nameLA) {
          out = { error: '休暇名を入力してください' };
        } else {
          var sheetLA = getOrCreateLeaveTypesSheet(ssLA);
          if (action === 'adminAddLeaveType') {
            var idLA = String(new Date().getTime());
            appendRowAsText(sheetLA, [idLA, nameLA, payLA, conditionLA, applyEnabledLA, urlLA, documentsLA]);
            out = { success: true, id: idLA };
          } else {
            var idLU = (e.parameter.id || '').trim();
            var dataLU = getSheetData(ssLA, '休暇制度');
            var rowIdxLU = -1;
            for (var lui = 0; lui < dataLU.rows.length; lui++) {
              if (String(dataLU.rows[lui][0]) === idLU) { rowIdxLU = lui + 2; break; }
            }
            if (rowIdxLU === -1) {
              out = { error: '対象の休暇制度が見つかりませんでした' };
            } else {
              setRangeAsText(sheetLA, rowIdxLU, 2, 1, 6, [[nameLA, payLA, conditionLA, applyEnabledLA, urlLA, documentsLA]]);
              out = { success: true, id: idLU };
            }
          }
        }
      }
    } else if (action === 'adminDeleteLeaveType') {
      var ssLD = openMasterSpreadsheet();
      if (!ssLD) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        var idLD = (e.parameter.id || '').trim();
        var sheetLD = getOrCreateLeaveTypesSheet(ssLD);
        var dataLD = getSheetData(ssLD, '休暇制度');
        var rowIdxLD = -1;
        for (var ldi = 0; ldi < dataLD.rows.length; ldi++) {
          if (String(dataLD.rows[ldi][0]) === idLD) { rowIdxLD = ldi + 2; break; }
        }
        if (rowIdxLD === -1) {
          out = { error: '対象の休暇制度が見つかりませんでした' };
        } else {
          sheetLD.deleteRow(rowIdxLD);
          out = { success: true };
        }
      }
    } else if (action === 'getRequiredDocumentOptions') {
      var ssRD = openMasterSpreadsheet();
      if (!ssRD) {
        out = {error: '「管理情報」スプレッドシートが見つかりません（管理者は先に移行処理を実行してください）'};
      } else {
        getOrCreateRequiredDocumentsSheet(ssRD);
        var dataRD = getSheetData(ssRD, '必要書類リスト');
        var optionsRD = dataRD.rows.map(function (row) { return row[0]; }).filter(function (v) { return v; });
        out = { success: true, options: optionsRD };
      }
    } else if (action === 'applyLeaveType') {
      var empIdAL = (e.parameter.empId || '').trim();
      var empNameAL = (e.parameter.empName || '').trim();
      var nameAL = (e.parameter.name || '').trim();
      var fromAL = (e.parameter.from || '').trim();
      var toAL = (e.parameter.to || '').trim();
      var reasonAL = (e.parameter.reason || '').trim();
      var docsSummaryAL = (e.parameter.docsSummary || '').trim();
      if (!nameAL || !fromAL) {
        out = { error: '休暇名・開始日は必須です' };
      } else {
        var periodAL = toAL ? (fromAL + '〜' + toAL) : (fromAL + '〜（終了日未定）');
        var bodyAL = '休暇制度の申請がありました。\n\n'
          + '休暇名：' + nameAL + '\n'
          + '申請者：' + empIdAL + '：' + empNameAL + '\n'
          + '期間：' + periodAL + '\n'
          + (reasonAL ? ('理由・備考：' + reasonAL + '\n') : '')
          + '\n必要書類：' + (docsSummaryAL || 'なし');
        MailApp.sendEmail(getLeaveTypeEmail(), '【給与台帳】休暇制度の申請：' + nameAL, bodyAL);
        out = { success: true };
      }
    } else if (action === 'updateMasterFolderId') {
      var newFolderId = (e.parameter.newFolderId || '').trim();
      if (!newFolderId) {
        out = { error: 'フォルダIDを入力してください' };
      } else {
        DriveApp.getFolderById(newFolderId); // 存在確認（無ければ例外→catchでエラー返却）
        PropertiesService.getScriptProperties().setProperty('MASTER_FOLDER_ID', newFolderId);
        out = { success: true };
      }
    } else if (action === 'updateScriptUrl') {
      var newUrl = (e.parameter.newUrl || '').trim();
      if (!newUrl) {
        out = { error: 'URLを入力してください' };
      } else if (BOOTSTRAP_DOC_ID === 'PUT_BOOTSTRAP_DOC_ID_HERE') {
        out = { error: 'BOOTSTRAP_DOC_IDが未設定です（コード内で設定してください）' };
      } else {
        var bDoc = DocumentApp.openById(BOOTSTRAP_DOC_ID);
        bDoc.getBody().editAsText().setText(newUrl);
        bDoc.saveAndClose();
        out = { success: true };
      }
    } else if (action === 'migrateToSpreadsheet') {
      var existingSSFile = findMasterSpreadsheetFile();
      if (existingSSFile) {
        out = { error: 'すでに「管理情報」スプレッドシートが存在します（移行済みの可能性があります）。' };
      } else {
        var oldListFile = findMasterListFile();
        if (!oldListFile) {
          out = { error: '移行元の「管理情報」ドキュメントが見つかりません。' };
        } else {
          var oldDoc = DocumentApp.openById(oldListFile.getId());
          var mainTextOld = oldDoc.getTabs()[0].asDocumentTab().getBody().getText();
          var companySettingsOld = parseCompanySettingsFromText(mainTextOld);
          var paydayOverridesOld = parsePaydayOverrides(mainTextOld);

          // 社員一覧（区分:社員ID:パスワード:フォルダID:氏名）を抽出。
          // 区分は必ず0か1のため、これを手がかりに設定値や特例行と区別する。
          var empRowsOld = [];
          var rowsOld = mainTextOld.split('\n');
          for (var roi = 0; roi < rowsOld.length; roi++) {
            var rowOld = rowsOld[roi].trim();
            if (!rowOld) continue;
            var colsOld = rowOld.replace(/：/g, ':').split(':');
            if (colsOld.length < 4) continue;
            if (colsOld[0].trim() !== '0' && colsOld[0].trim() !== '1') continue;
            empRowsOld.push([colsOld[0].trim(), colsOld[1].trim(), colsOld[2].trim(), colsOld[3].trim(), colsOld[4] ? colsOld[4].trim() : '']);
          }
          var userBodyOld = getTabBodyByTitle(oldDoc, 'ユーザー情報');
          var userRecordsOld = userBodyOld ? parseUserRecords(userBodyOld.getText()) : [];

          var masterFolderMig = DriveApp.getFolderById(MASTER_FOLDER_ID);
          var newSS = SpreadsheetApp.create('管理情報');
          var newSSFile = DriveApp.getFileById(newSS.getId());
          masterFolderMig.addFile(newSSFile);
          DriveApp.getRootFolder().removeFile(newSSFile);

          var sheetEmp = newSS.getSheets()[0];
          sheetEmp.setName('社員一覧');
          sheetEmp.getRange('A:E').setNumberFormat('@'); // ID・パスワード等は常にテキスト扱いにする（列全体への保険）
          appendRowAsText(sheetEmp, ['区分', '社員ID', 'パスワード', 'フォルダID', '氏名']);
          if (empRowsOld.length > 0) setRangeAsText(sheetEmp, 2, 1, empRowsOld.length, 5, empRowsOld);

          var sheetUser = newSS.insertSheet('ユーザー情報');
          sheetUser.getRange('A:M').setNumberFormat('@'); // 生年月日・削除日などが日付として自動変換されるのを防ぐ（列全体への保険）
          appendRowAsText(sheetUser, USER_INFO_HEADER);
          if (userRecordsOld.length > 0) {
            var userRowsOld = userRecordsOld.map(function (r) {
              return USER_INFO_HEADER.map(function (k) { return r[k] !== undefined ? r[k] : ''; });
            });
            setRangeAsText(sheetUser, 2, 1, userRowsOld.length, USER_INFO_HEADER.length, userRowsOld);
          }

          var sheetCompany = newSS.insertSheet('会社共通設定');
          sheetCompany.getRange('A:B').setNumberFormat('@');
          appendRowAsText(sheetCompany, ['設定項目', '値']);
          COMPANY_SETTING_KEYS.forEach(function (k) {
            if (companySettingsOld[k] !== undefined) appendRowAsText(sheetCompany, [k, companySettingsOld[k]]);
          });

          var sheetPayday = newSS.insertSheet('給料日特例');
          sheetPayday.getRange('A:B').setNumberFormat('@');
          appendRowAsText(sheetPayday, ['対象年月', '給料日']);
          Object.keys(paydayOverridesOld).forEach(function (ym) {
            appendRowAsText(sheetPayday, [ym, paydayOverridesOld[ym]]);
          });

          // 旧ドキュメントは削除せず、バックアップとして名前を変えて残す（非破壊的な移行）
          try { oldListFile.setName('管理情報（旧・ドキュメント版バックアップ）'); } catch (renameErr) { /* リネーム失敗は無視 */ }

          logSystemEvent('migrate_to_spreadsheet', '管理情報をスプレッドシートに移行しました（社員' + empRowsOld.length + '件、ユーザー情報' + userRecordsOld.length + '件）');
          out = { success: true, employeeCount: empRowsOld.length, userInfoCount: userRecordsOld.length, spreadsheetId: newSS.getId() };
        }
      }
    } else if (action === 'submitTransportRequest') {
      var tEmpId = (e.parameter.empId || '').trim();
      var tEmpName = (e.parameter.empName || '').trim();
      var trips;
      try { trips = JSON.parse(e.parameter.tripsJson || '[]'); } catch (parseErrT) { trips = null; }
      if (!tEmpId || !trips || !Array.isArray(trips) || trips.length === 0) {
        out = { error: '申請内容が正しくありません' };
      } else if (trips.length > TRANSPORT_TRIP_MAX_ROWS) {
        out = { error: '一度に申請できる件数は' + TRANSPORT_TRIP_MAX_ROWS + '件までです' };
      } else {
        var transportFolder = findTransportFolder();
        if (!transportFolder) {
          out = { error: '「交通費申請」フォルダが見つかりません（管理者にご確認ください）' };
        } else {
          var templateFile = findTransportTemplateFile();
          if (!templateFile) {
            out = { error: '「交通費申請書(原紙)」が見つかりません（管理者にご確認ください）' };
          } else {
            var nowT = new Date();
            var tsStr = Utilities.formatDate(nowT, Session.getScriptTimeZone() || 'Asia/Tokyo', 'yyyyMMdd_HHmmss');
            var copiedFile = templateFile.makeCopy(tEmpId + '_' + tsStr, transportFolder);
            var copiedSS = SpreadsheetApp.openById(copiedFile.getId());
            var sheetT = copiedSS.getSheets()[0];
            sheetT.getRange(2, 4).setValue(todayDateString());
            sheetT.getRange(3, 4).setValue(tEmpName || tEmpId);
            var totalT = 0;
            var tripRowsT = [];
            for (var ti = 0; ti < trips.length; ti++) {
              var tr = trips[ti];
              var amt = parseFloat(tr.amount) || 0;
              totalT += amt;
              tripRowsT.push([tr.y || '', tr.m || '', tr.d || '', tr.from || '', tr.to || '', tr.method || '', tr.roundTrip || '', amt, tr.note || '']);
            }
            sheetT.getRange(TRANSPORT_TRIP_START_ROW, 1, tripRowsT.length, 9).setValues(tripRowsT);

            var requestId = 'tr_' + Utilities.getUuid().replace(/-/g, '').slice(0, 16);
            var nowIsoT = nowT.toISOString();
            var recordT = {
              requestId: requestId, empId: tEmpId, empName: tEmpName, appliedAt: nowIsoT,
              status: 'pending', decidedAt: '', rejectReason: '',
              trips: trips, total: totalT, sheetFileId: copiedFile.getId(), sheetUrl: copiedFile.getUrl()
            };
            var ledgerDocC = getOrCreateTransportLedgerDoc(transportFolder);
            appendLine(ledgerDocC, serializeTransportLedgerLine(recordT));

            var newsDocT = getOrCreateMasterDoc('ニュース');
            var newsIdT = new Date().getTime();
            var newsPayloadT = encodeTransportPayload({ requestId: requestId, empId: tEmpId, empName: tEmpName, total: totalT, trips: trips, appliedAt: nowIsoT });
            appendLine(newsDocT, [newsIdT, 'admin', 'critical', '交通費申請', '', '', TRANSPORT_REQUEST_MARKER + newsPayloadT].join(':'));

            try {
              var mailBodyT = (tEmpName || tEmpId) + '（' + tEmpId + '）さんから交通費申請が届きました。\n件数：' + trips.length + '件\n合計金額：' + totalT + '円\n\nアプリより内容を確認し、承認・却下を行ってください。\n※本メールは給与台帳より自動送信されています。';
              MailApp.sendEmail(getTransportEmail(), '給与台帳：交通費申請', mailBodyT);
            } catch (mailErrT) { /* メール送信失敗は申請自体の成功を妨げない */ }

            logSystemEvent('transport_request', tEmpId + ':' + tEmpName + ' が交通費申請（' + trips.length + '件・' + totalT + '円）');
            out = { success: true, requestId: requestId, sheetUrl: copiedFile.getUrl() };
          }
        }
      }
    } else if (action === 'getTransportRequests') {
      var transportFolderG = findTransportFolder();
      if (!transportFolderG) {
        out = { success: true, requests: [] };
      } else {
        var ledgerDocG = getOrCreateTransportLedgerDoc(transportFolderG);
        var linesG = ledgerDocG.getBody().getText().split('\n').filter(function (l) { return l.trim(); });
        var empFilterG = (e.parameter.empId || '').trim();
        var requestsG = [];
        for (var gi = 0; gi < linesG.length; gi++) {
          var recG = parseTransportLedgerLine(linesG[gi].trim());
          if (!recG) continue;
          if (empFilterG && recG.empId !== empFilterG) continue;
          requestsG.push(recG);
        }
        requestsG.sort(function (a, b) { return (b.appliedAt || '').localeCompare(a.appliedAt || ''); });
        out = { success: true, requests: requestsG };
      }
    } else if (action === 'resolveTransportRequest') {
      var transportFolderR = findTransportFolder();
      if (!transportFolderR) {
        out = { error: '「交通費申請」フォルダが見つかりません' };
      } else {
        var requestIdR2 = (e.parameter.requestId || '').trim();
        var decisionR2 = e.parameter.decision === 'reject' ? 'reject' : 'approve';
        var reasonR2 = (e.parameter.reason || '').trim();
        if (decisionR2 === 'reject' && !reasonR2) {
          out = { error: '却下の理由を入力してください' };
        } else {
          var ledgerDocR2 = getOrCreateTransportLedgerDoc(transportFolderR);
          var bodyR2 = ledgerDocR2.getBody();
          var linesR2 = bodyR2.getText().split('\n');
          var targetIdxR2 = -1, targetRecR2 = null;
          for (var ri2 = 0; ri2 < linesR2.length; ri2++) {
            var rowR2 = linesR2[ri2].trim();
            if (!rowR2) continue;
            var recR2 = parseTransportLedgerLine(rowR2);
            if (recR2 && recR2.requestId === requestIdR2) { targetIdxR2 = ri2; targetRecR2 = recR2; break; }
          }
          if (targetIdxR2 === -1) {
            out = { error: '対象の申請が見つかりませんでした' };
          } else if (targetRecR2.status !== 'pending') {
            out = { error: 'この申請は既に処理済みです' };
          } else {
            targetRecR2.status = decisionR2 === 'approve' ? 'approved' : 'rejected';
            targetRecR2.decidedAt = new Date().toISOString();
            targetRecR2.rejectReason = decisionR2 === 'reject' ? reasonR2 : '';
            linesR2[targetIdxR2] = serializeTransportLedgerLine(targetRecR2);
            bodyR2.editAsText().setText(linesR2.join('\n'));
            ledgerDocR2.saveAndClose();
            out = { success: true, empId: targetRecR2.empId, empName: targetRecR2.empName, total: targetRecR2.total };
          }
        }
      }
    } else if (action === 'reviseTransportRequest') {
      // 保守画面から、判定済み（あるいは未処理）の交通費申請を後から修正するための操作。理由は常に必須。
      var transportFolderV = findTransportFolder();
      if (!transportFolderV) {
        out = { error: '「交通費申請」フォルダが見つかりません' };
      } else {
        var requestIdV = (e.parameter.requestId || '').trim();
        var decisionV = e.parameter.decision === 'reject' ? 'reject' : 'approve';
        var reasonV = (e.parameter.reason || '').trim();
        if (!reasonV) {
          out = { error: '修正の理由を入力してください' };
        } else {
          var ledgerDocV = getOrCreateTransportLedgerDoc(transportFolderV);
          var bodyV = ledgerDocV.getBody();
          var linesV = bodyV.getText().split('\n');
          var targetIdxV = -1, targetRecV = null;
          for (var vi = 0; vi < linesV.length; vi++) {
            var rowV = linesV[vi].trim();
            if (!rowV) continue;
            var recV = parseTransportLedgerLine(rowV);
            if (recV && recV.requestId === requestIdV) { targetIdxV = vi; targetRecV = recV; break; }
          }
          if (targetIdxV === -1) {
            out = { error: '対象の申請が見つかりませんでした' };
          } else {
            targetRecV.status = decisionV === 'approve' ? 'approved' : 'rejected';
            targetRecV.decidedAt = new Date().toISOString();
            targetRecV.rejectReason = reasonV;
            linesV[targetIdxV] = serializeTransportLedgerLine(targetRecV);
            bodyV.editAsText().setText(linesV.join('\n'));
            ledgerDocV.saveAndClose();
            logSystemEvent('transport_request_revise', targetRecV.empId + ':' + targetRecV.empName + ' の交通費申請（' + requestIdV + '）の判定を' + (decisionV === 'approve' ? '許諾' : '不許可') + 'に修正（理由：' + reasonV + '）');
            out = { success: true, empId: targetRecV.empId, empName: targetRecV.empName, total: targetRecV.total, status: targetRecV.status };
          }
        }
      }
    } else if (action === 'debugMatch') {
      var listFileM = findMasterListFile();
      if (!listFileM) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docM = DocumentApp.openById(listFileM.getId());
        var textM = docM.getTabs()[0].asDocumentTab().getBody().getText();
        var inputIdM = (e.parameter.id || '').trim();
        var inputPwM = (e.parameter.password || '').trim();
        var rowsM = textM.split('\n');
        var debugRows = [];
        var matchedRow = null;
        for (var mi3 = 0; mi3 < rowsM.length; mi3++) {
          var rowM = rowsM[mi3].trim();
          if (!rowM) continue;
          var colsM = rowM.replace(/：/g, ':').split(':');
          var idMatch = (colsM.length >= 2) ? (colsM[1].trim() === inputIdM) : false;
          var pwMatch = (colsM.length >= 3) ? (colsM[2].trim() === inputPwM) : false;
          debugRows.push({ raw: rowM, cols: colsM, colsLength: colsM.length, idMatch: idMatch, pwMatch: pwMatch });
          if (colsM.length >= 4 && idMatch && pwMatch) matchedRow = rowM;
        }
        out = { success: true, inputId: inputIdM, inputIdLength: inputIdM.length, inputPassword: inputPwM, rows: debugRows, matchedRow: matchedRow };
      }
    } else if (action === 'debugParams') {
      out = {
        success: true,
        rawId: e.parameter.id,
        rawIdType: typeof e.parameter.id,
        rawIdLength: (e.parameter.id === undefined || e.parameter.id === null) ? -1 : String(e.parameter.id).length,
        rawPassword: e.parameter.password
      };
    } else if (action === 'debugMainText') {
      var listFileD = findMasterListFile();
      if (!listFileD) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docD = DocumentApp.openById(listFileD.getId());
        var tabsD = docD.getTabs();
        var tabTitlesD = tabsD.map(function(t){ return t.getTitle(); });
        var textD = tabsD[0].asDocumentTab().getBody().getText();
        out = { success: true, tabCount: tabsD.length, tabTitles: tabTitlesD, text: textD, textLength: textD.length };
      }
    } else if (action === 'employeeLogin') {
      var loginSS = openMasterSpreadsheet();
      if (loginSS) {
        // ---- 新形式：管理情報スプレッドシート ----
        var inputIdSS = (e.parameter.id || '').trim();
        var inputPwSS = (e.parameter.password || '').trim();
        var empRowSS = findEmployeeRowSS(loginSS, inputIdSS);
        var pwOkSS = false;
        if (empRowSS) {
          if (empRowSS.password.indexOf('sha256:') === 0) {
            pwOkSS = (empRowSS.password === hashPasswordSha256(inputPwSS));
          } else {
            // 旧形式（平文）：一致すればログインを許可し、この場でハッシュ形式に自動移行する
            pwOkSS = (empRowSS.password === inputPwSS);
            if (pwOkSS) {
              loginSS.getSheetByName('社員一覧').getRange(empRowSS.rowIndex, 3).setValue(hashPasswordSha256(inputPwSS));
            }
          }
        }
        if (!empRowSS || !pwOkSS) {
          out = { error: 'IDまたはパスワードが違います' };
        } else {
          var nameSS = empRowSS.name || inputIdSS;
          var isDeletedSS = false, isExecutiveSS = false;
          var userRowSS = findUserInfoRowSS(loginSS, inputIdSS);
          if (userRowSS) {
            var recSS = userRowSS.record;
            var seiSS = recSS['苗字'] || '', meiSS = recSS['名前'] || '';
            if (seiSS || meiSS) nameSS = seiSS + meiSS;
            isDeletedSS = String(recSS['削除フラグ']) === '1';
            isExecutiveSS = String(recSS['役員フラグ']) === '1';
          }
          if (isDeletedSS) {
            out = { error: 'このアカウントは削除処理中のため、ログインできません。管理者にお問い合わせください。' };
          } else {
            out = { success: true, folderId: empRowSS.folderId, name: nameSS, isAdmin: (empRowSS.role === '1'), isExecutive: isExecutiveSS, companySettings: readCompanySettingsSS(loginSS), paydayOverrides: readPaydayOverridesSS(loginSS) };
          }
        }
      } else {
      // ---- 旧形式：管理情報ドキュメント（スプレッドシートへの移行が済むまでのフォールバック） ----
      var masterFolder = DriveApp.getFolderById(MASTER_FOLDER_ID);
      var listFile = null;
      var candidates = ['管理情報', '社員一覧'];
      for (var ci = 0; ci < candidates.length && !listFile; ci++) {
        var lf = masterFolder.getFilesByName(candidates[ci]);
        while (lf.hasNext()) {
          var cand = lf.next();
          if (cand.getMimeType() === MimeType.GOOGLE_DOCS) { listFile = cand; break; }
        }
      }
      if (!listFile) {
        out = {error: '「管理情報」または「社員一覧」ドキュメントが見つかりません'};
      } else {
        var docL = DocumentApp.openById(listFile.getId());
        var listText = docL.getTabs()[0].asDocumentTab().getBody().getText();
        var inputId = (e.parameter.id || '').trim();
        var inputPw = (e.parameter.password || '').trim();
        var found = null;
        var companySettings = parseCompanySettingsFromText(listText);
        var rows = listText.split('\n');
        for (var i = 0; i < rows.length; i++) {
          var row = rows[i].trim();
          if (!row) continue;
          var cols = row.replace(/：/g, ':').split(':');
          if (cols.length < 4) continue;
          if (cols[1].trim() === inputId && cols[2].trim() === inputPw) {
            found = { role: cols[0].trim(), folderId: cols[3].trim(), name: cols[4] ? cols[4].trim() : inputId };
          }
        }
        if (found) {
          var userBodyL = getTabBodyByTitle(docL, 'ユーザー情報');
          var isDeletedL = false;
          var isExecutiveL = false;
          if (userBodyL) {
            var recordsL = parseUserRecords(userBodyL.getText());
            for (var rj = 0; rj < recordsL.length; rj++) {
              if (recordsL[rj]['ユーザーID'] === inputId) {
                var seiL = recordsL[rj]['苗字'] || '';
                var meiL = recordsL[rj]['名前'] || '';
                if (seiL || meiL) found.name = (seiL + meiL);
                if (recordsL[rj]['削除フラグ'] === '1') isDeletedL = true;
                if (recordsL[rj]['役員フラグ'] === '1') isExecutiveL = true;
                break;
              }
            }
          }
          if (isDeletedL) {
            out = { error: 'このアカウントは削除処理中のため、ログインできません。管理者にお問い合わせください。' };
          } else {
            out = { success: true, folderId: found.folderId, name: found.name, isAdmin: (found.role === '1'), isExecutive: isExecutiveL, companySettings: companySettings, paydayOverrides: parsePaydayOverrides(listText) };
          }
        } else {
          out = { error: 'IDまたはパスワードが違います' };
        }
      }
      }
    } else {
      out = {error: 'unknown action'};
    }
  } catch (err) {
    out = {error: String(err)};
  }
  var json = JSON.stringify(out);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// この関数は「トリガー」（時間主導型・毎日実行）として管理者が1回だけ設定してください。
// Apps Scriptエディタ左側の時計アイコン→「トリガーを追加」→実行する関数：checkDeletionSchedule→
// イベントの種類：時間主導型→日付ベースのタイマー→午前2時〜3時など、好きな時間帯を選んでください。
function checkDeletionSchedule() {
  var ss = openMasterSpreadsheet();
  if (!ss) return;
  var userSheet = ss.getSheetByName('ユーザー情報');
  var empSheet = ss.getSheetByName('社員一覧');
  if (!userSheet || !empSheet) return;
  var userData = getSheetData(ss, 'ユーザー情報');
  var empData = getSheetData(ss, '社員一覧');
  var today = todayDateString();
  var changed = false;
  var empChanged = false;
  var remainingUserRows = [];
  var remainingEmpRows = empData.rows.slice();
  var hIdx = {};
  userData.header.forEach(function (h, i) { hIdx[String(h).trim()] = i; });

  userData.rows.forEach(function (rowIn) {
    var row = rowIn.slice();
    var r = rowToObject(userData.header, row);
    if (String(r['削除フラグ']) !== '1') { remainingUserRows.push(row); return; }

    var scheduledDate = r['削除予定日'] ? String(r['削除予定日']) : '';
    if (!scheduledDate) { remainingUserRows.push(row); return; }

    // 1か月前告知（まだ告知していなければ）
    if (String(r['告知済み']) !== '1') {
      var oneMonthBefore = addDaysToDateString(scheduledDate, -30);
      if (today >= oneMonthBefore) {
        var msg = r['削除日'] + 'に削除された' + r['ユーザーID'] + ':' + (r['苗字'] || '') + '　' + (r['名前'] || '') + 'のアカウント情報が、' + scheduledDate + 'に完全削除される予定です';
        var newsDoc = getOrCreateMasterDoc('ニュース');
        appendLine(newsDoc, new Date().getTime() + ':admin:critical:完全削除予告:' + msg.replace(/:/g, '：'));
        row[hIdx['告知済み']] = '1';
        changed = true;
        logSystemEvent('user_delete_notice', r['ユーザーID'] + ' の完全削除予告を管理者に通知（予定日：' + scheduledDate + '）');
      }
    }

    // 完全削除の実行
    if (today >= scheduledDate) {
      var empIdxFound = -1, folderIdToTrash = null;
      for (var mi = 0; mi < remainingEmpRows.length; mi++) {
        if (String(remainingEmpRows[mi][1]).trim() === String(r['ユーザーID'])) { empIdxFound = mi; folderIdToTrash = String(remainingEmpRows[mi][3]).trim(); break; }
      }
      if (empIdxFound >= 0) {
        try {
          DriveApp.getFolderById(folderIdToTrash).setTrashed(true);
        } catch (errTrash) { /* フォルダが既に無い場合等は無視 */ }
        remainingEmpRows.splice(empIdxFound, 1);
        empChanged = true;
      }
      logSystemEvent('user_delete_completed', r['ユーザーID'] + ':' + (r['苗字'] || '') + (r['名前'] || '') + ' のアカウント情報を完全削除しました（削除理由：' + (r['削除理由'] || '') + '）');
      changed = true;
      return; // このレコードはremainingUserRowsに含めない（完全削除）
    }

    remainingUserRows.push(row);
  });

  if (changed && userData.rows.length > 0) {
    userSheet.getRange(2, 1, userData.rows.length, userData.header.length).clearContent();
    if (remainingUserRows.length > 0) userSheet.getRange(2, 1, remainingUserRows.length, userData.header.length).setValues(remainingUserRows);
  }
  if (empChanged && empData.rows.length > 0) {
    empSheet.getRange(2, 1, empData.rows.length, empData.header.length).clearContent();
    if (remainingEmpRows.length > 0) empSheet.getRange(2, 1, remainingEmpRows.length, empData.header.length).setValues(remainingEmpRows);
  }
}