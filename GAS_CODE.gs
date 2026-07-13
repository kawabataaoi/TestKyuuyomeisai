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

function parsePaydayOverrides(text) {
  var lines = text.split('\n');
  var overrides = {};
  var inSection = false;
  for (var i = 0; i < lines.length; i++) {
    var row = lines[i].trim();
    if (!row) continue;
    if (row === '給料日特例') { inSection = true; continue; }
    if (row === '会社共通設定') { inSection = false; continue; }
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
        var verifyListFile = findMasterListFile();
        if (verifyListFile) {
          var verifyDoc = DocumentApp.openById(verifyListFile.getId());
          var verifyMainText = verifyDoc.getTabs()[0].asDocumentTab().getBody().getText();
          var verifyRow = findMainListRow(verifyMainText, empIdCheck);
          if (!verifyRow || verifyRow.cols[3].trim() !== folderParam) {
            logSystemEvent('write_settings_blocked', '社員ID「' + empIdCheck + '」に紐づかないフォルダ（' + folderParam + '）への書き込みをブロックしました');
            out = { error: 'フォルダとユーザーの整合性が取れないため、書き込みを中止しました（安全のためのチェックです）。ページを再読み込みして、再度ログインしてから試してください。' };
          }
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
      MailApp.sendEmail(getInquiryEmail(), subject, body);
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
    } else if (action === 'companySettings') {
      var masterFolder0 = DriveApp.getFolderById(MASTER_FOLDER_ID);
      var listFile0 = null;
      var candidates0 = ['管理情報', '社員一覧'];
      for (var cj = 0; cj < candidates0.length && !listFile0; cj++) {
        var lf0 = masterFolder0.getFilesByName(candidates0[cj]);
        while (lf0.hasNext()) {
          var cand0 = lf0.next();
          if (cand0.getMimeType() === MimeType.GOOGLE_DOCS) { listFile0 = cand0; break; }
        }
      }
      if (!listFile0) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var listText0 = DocumentApp.openById(listFile0.getId()).getTabs()[0].asDocumentTab().getBody().getText();
        var cs0 = {};
        var inSection0 = false;
        var rows0 = listText0.split('\n');
        for (var k = 0; k < rows0.length; k++) {
          var row0 = rows0[k].trim();
          if (!row0) continue;
          if (row0 === '会社共通設定') { inSection0 = true; continue; }
          if (inSection0) {
            var kv0 = row0.split(':');
            if (kv0.length >= 2) cs0[kv0[0].trim()] = kv0.slice(1).join(':').trim();
          }
        }
        out = { success: true, companySettings: cs0, paydayOverrides: parsePaydayOverrides(listText0) };
      }
    } else if (action === 'updatePaydayOverride') {
      var listFileP = findMasterListFile();
      if (!listFileP) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docP = DocumentApp.openById(listFileP.getId());
        var bodyP = docP.getBody();
        var linesP = bodyP.getText().split('\n');
        var ymP = (e.parameter.ym || '').trim();
        var dateP = (e.parameter.date || '').trim(); // 空なら特例を削除
        var sectionIdxP = -1;
        var endIdxP = linesP.length;
        for (var pi = 0; pi < linesP.length; pi++) {
          if (linesP[pi].trim() === '給料日特例') { sectionIdxP = pi; continue; }
          if (sectionIdxP >= 0 && linesP[pi].trim() === '会社共通設定') { endIdxP = pi; break; }
        }
        if (sectionIdxP < 0) {
          // セクション自体が無ければ末尾に新設
          linesP.push('');
          linesP.push('給料日特例');
          if (dateP) linesP.push(ymP + ':' + dateP);
          linesP.push('');
        } else {
          var replacedP = false;
          for (var pj = sectionIdxP + 1; pj < endIdxP; pj++) {
            var rowP = linesP[pj].trim();
            if (!rowP) continue;
            var idxP2 = rowP.indexOf(':');
            if (idxP2 > 0 && rowP.substring(0, idxP2).trim() === ymP) {
              if (dateP) { linesP[pj] = ymP + ':' + dateP; }
              else { linesP.splice(pj, 1); }
              replacedP = true;
              break;
            }
          }
          if (!replacedP && dateP) {
            linesP.splice(sectionIdxP + 1, 0, ymP + ':' + dateP);
          }
        }
        bodyP.editAsText().setText(linesP.join('\n'));
        docP.saveAndClose();
        out = { success: true };
      }
    } else if (action === 'findUserByInfo') {
      var listFileU = findMasterListFile();
      if (!listFileU) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docU = DocumentApp.openById(listFileU.getId());
        var userBody = getTabBodyByTitle(docU, 'ユーザー情報');
        if (!userBody) {
          out = {error: '「ユーザー情報」タブが見つかりません'};
        } else {
          var uLines = userBody.getText().split('\n');
          var records = [];
          var cur = null;
          for (var ui = 0; ui < uLines.length; ui++) {
            var uln = uLines[ui].trim();
            if (!uln) continue;
            var norm = uln.replace(/：/g, ':');
            var uidx = norm.indexOf(':');
            if (uidx < 0) continue;
            var ukey = norm.substring(0, uidx).trim();
            var uval = norm.substring(uidx + 1).trim();
            if (ukey === 'ユーザーID') {
              if (cur) records.push(cur);
              cur = {};
            }
            if (cur) cur[ukey] = uval;
          }
          if (cur) records.push(cur);
          var inSei = (e.parameter.sei || '').trim();
          var inMei = (e.parameter.mei || '').trim();
          var inDob = digitsOnly(e.parameter.dob || '');
          var matchedId = null;
          for (var ri = 0; ri < records.length; ri++) {
            var r = records[ri];
            if (r['苗字'] === inSei && r['名前'] === inMei && digitsOnly(r['生年月日']) === inDob) {
              matchedId = r['ユーザーID'];
              break;
            }
          }
          if (matchedId) {
            out = { success: true, userId: matchedId };
          } else {
            out = { error: '一致する情報が見つかりませんでした' };
          }
        }
      }
    } else if (action === 'resetPassword') {
      var listFileR = findMasterListFile();
      if (!listFileR) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docR = DocumentApp.openById(listFileR.getId());
        var mainBody = docR.getTabs()[0].asDocumentTab().getBody();
        var mainLines = mainBody.getText().split('\n');
        var targetId = (e.parameter.id || '').trim();
        var newPw = (e.parameter.newPassword || '').trim().replace(/:/g, '');
        var changed = false;
        for (var mi = 0; mi < mainLines.length; mi++) {
          var mrow = mainLines[mi].trim();
          if (!mrow) continue;
          var mcols = mrow.replace(/：/g, ':').split(':');
          if (mcols.length < 4) continue;
          if (mcols[1].trim() === targetId) {
            mcols[2] = newPw;
            mainLines[mi] = mcols.join(':');
            changed = true;
            break;
          }
        }
        if (changed) {
          mainBody.editAsText().setText(mainLines.join('\n'));
          docR.saveAndClose();
          out = { success: true };
        } else {
          out = { error: '該当するユーザーIDが見つかりませんでした' };
        }
      }
    } else if (action === 'postNews') {
      var newsDocP = getOrCreateMasterDoc('ニュース');
      var newsId = new Date().getTime();
      var scope = e.parameter.scope === 'admin' ? 'admin' : 'all';
      var priorityP = e.parameter.priority || 'normal';
      var titleP = (e.parameter.title || '').replace(/\n/g, ' ').replace(/:/g, '：');
      var msg = (e.parameter.message || '').replace(/\n/g, ' ').replace(/:/g, '：');
      appendLine(newsDocP, [newsId, scope, priorityP, titleP, '', '', msg].join(':'));
      out = { success: true };
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
      var listFileQ = findMasterListFile();
      if (!listFileQ) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docQ = DocumentApp.openById(listFileQ.getId());
        var mainTextQ = docQ.getTabs()[0].asDocumentTab().getBody().getText();
        var queryQ = (e.parameter.query || '').trim();
        var byIdQ = findMainListRow(mainTextQ, queryQ);
        if (byIdQ) {
          var detailsQ = buildUserDetails(docQ, mainTextQ, queryQ);
          out = { success: true, mode: 'single', detail: detailsQ };
        } else {
          var userBodyQ = getTabBodyByTitle(docQ, 'ユーザー情報');
          var matchesQ = [];
          if (userBodyQ) {
            var recordsQ = parseUserRecords(userBodyQ.getText());
            for (var qi = 0; qi < recordsQ.length; qi++) {
              if (recordsQ[qi]['みょうじ'] === queryQ) {
                matchesQ.push({
                  id: recordsQ[qi]['ユーザーID'],
                  sei: recordsQ[qi]['苗字'] || '',
                  mei: recordsQ[qi]['名前'] || '',
                  seiKana: recordsQ[qi]['みょうじ'] || '',
                  meiKana: recordsQ[qi]['なまえ'] || '',
                  deleted: recordsQ[qi]['削除フラグ'] === '1'
                });
              }
            }
          }
          if (matchesQ.length === 0) {
            out = { error: '該当するユーザーが見つかりませんでした' };
          } else if (matchesQ.length === 1) {
            var detailsQ1 = buildUserDetails(docQ, mainTextQ, matchesQ[0].id);
            out = { success: true, mode: 'single', detail: detailsQ1 };
          } else {
            out = { success: true, mode: 'multiple', matches: matchesQ };
          }
        }
      }
    } else if (action === 'adminDeleteUser') {
      var listFileDel = findMasterListFile();
      if (!listFileDel) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docDel = DocumentApp.openById(listFileDel.getId());
        var targetIdDel = (e.parameter.targetId || '').trim();
        var reasonDel = (e.parameter.reason || '').trim();
        var userBodyDel = getTabBodyByTitle(docDel, 'ユーザー情報');
        if (!userBodyDel) {
          out = { error: '「ユーザー情報」タブが見つかりません' };
        } else {
          var recordsDel = parseUserRecords(userBodyDel.getText());
          var foundDel = false;
          var todayDel = todayDateString();
          var scheduledDel = addDaysToDateString(todayDel, getRetentionDays());
          for (var di = 0; di < recordsDel.length; di++) {
            if (recordsDel[di]['ユーザーID'] === targetIdDel) {
              recordsDel[di]['削除フラグ'] = '1';
              recordsDel[di]['削除日'] = todayDel;
              recordsDel[di]['削除予定日'] = scheduledDel;
              recordsDel[di]['削除理由'] = reasonDel || '（理由未入力）';
              recordsDel[di]['告知済み'] = '0';
              foundDel = true;
              break;
            }
          }
          if (!foundDel) {
            out = { error: '該当するユーザーが見つかりませんでした' };
          } else {
            userBodyDel.editAsText().setText(serializeUserRecords(recordsDel));
            docDel.saveAndClose();
            logSystemEvent('user_delete_scheduled', targetIdDel + ' を削除予定に設定（完全削除予定日：' + scheduledDel + '、理由：' + reasonDel + '）');
            out = { success: true, scheduledDate: scheduledDel };
          }
        }
      }
    } else if (action === 'adminRestoreUser') {
      var listFileRes = findMasterListFile();
      if (!listFileRes) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docRes = DocumentApp.openById(listFileRes.getId());
        var targetIdRes = (e.parameter.targetId || '').trim();
        var reasonRes = (e.parameter.reason || '').trim();
        var userBodyRes = getTabBodyByTitle(docRes, 'ユーザー情報');
        if (!userBodyRes) {
          out = { error: '「ユーザー情報」タブが見つかりません' };
        } else {
          var recordsRes = parseUserRecords(userBodyRes.getText());
          var foundRes = false;
          for (var ri3 = 0; ri3 < recordsRes.length; ri3++) {
            if (recordsRes[ri3]['ユーザーID'] === targetIdRes) {
              delete recordsRes[ri3]['削除フラグ'];
              delete recordsRes[ri3]['削除日'];
              delete recordsRes[ri3]['削除予定日'];
              delete recordsRes[ri3]['削除理由'];
              delete recordsRes[ri3]['告知済み'];
              foundRes = true;
              break;
            }
          }
          if (!foundRes) {
            out = { error: '該当するユーザーが見つかりませんでした' };
          } else {
            userBodyRes.editAsText().setText(serializeUserRecords(recordsRes));
            docRes.saveAndClose();
            logSystemEvent('user_restore', targetIdRes + ' の削除を取り消しました（理由：' + reasonRes + '）');
            out = { success: true };
          }
        }
      }
    } else if (action === 'adminCreateUser') {
      var listFileC = findMasterListFile();
      if (!listFileC) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docC = DocumentApp.openById(listFileC.getId());
        var mainTextC = docC.getTabs()[0].asDocumentTab().getBody().getText();
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
        } else if (findMainListRow(mainTextC, newIdC)) {
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

          var mainBodyC = docC.getTabs()[0].asDocumentTab().getBody();
          mainBodyC.appendParagraph(newRoleC + ':' + newIdC + ':' + newPwC + ':' + newEmpFolderC.getId());

          var userBodyC = getTabBodyByTitle(docC, 'ユーザー情報');
          if (userBodyC) {
            var recordsC = parseUserRecords(userBodyC.getText());
            recordsC.push({
              'ユーザーID': newIdC,
              '苗字': newSeiC,
              '名前': newMeiC,
              'みょうじ': newSeiKanaC,
              'なまえ': newMeiKanaC,
              '生年月日': newDobC,
              '管理者権限': newRoleC,
              '役員フラグ': newIsExecutiveC
            });
            userBodyC.editAsText().setText(serializeUserRecords(recordsC));
          }
          docC.saveAndClose();
          out = { success: true, id: newIdC, folderId: newEmpFolderC.getId() };
        }
      }
    } else if (action === 'adminSearchUser') {
      var listFileS = findMasterListFile();
      if (!listFileS) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docS = DocumentApp.openById(listFileS.getId());
        var mainTextS = docS.getTabs()[0].asDocumentTab().getBody().getText();
        var targetIdS = (e.parameter.targetId || '').trim();
        var detailsS = buildUserDetails(docS, mainTextS, targetIdS);
        if (!detailsS) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          out = { success: true, id: detailsS.id, folderId: detailsS.folderId, userInfo: detailsS.userInfo, byYear: detailsS.byYear };
        }
      }
    } else if (action === 'adminUpdateUserInfo') {
      var listFileI = findMasterListFile();
      if (!listFileI) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docI = DocumentApp.openById(listFileI.getId());
        var targetIdI = (e.parameter.targetId || '').trim();
        var newRole = (e.parameter.role || '0').trim();
        var mainBodyI = docI.getTabs()[0].asDocumentTab().getBody();
        var mainLinesI = mainBodyI.getText().split('\n');
        var changedI = false;
        for (var mi2 = 0; mi2 < mainLinesI.length; mi2++) {
          var mrow2 = mainLinesI[mi2].trim();
          if (!mrow2) continue;
          var mcols2 = mrow2.replace(/：/g, ':').split(':');
          if (mcols2.length < 4) continue;
          if (mcols2[1].trim() === targetIdI) {
            mcols2[0] = newRole;
            mainLinesI[mi2] = mcols2.join(':');
            changedI = true;
            break;
          }
        }
        if (!changedI) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          mainBodyI.editAsText().setText(mainLinesI.join('\n'));
          var userBodyI = getTabBodyByTitle(docI, 'ユーザー情報');
          if (userBodyI) {
            var recordsI = parseUserRecords(userBodyI.getText());
            var foundI = false;
            for (var ri2 = 0; ri2 < recordsI.length; ri2++) {
              if (recordsI[ri2]['ユーザーID'] === targetIdI) {
                recordsI[ri2]['苗字'] = e.parameter.sei || '';
                recordsI[ri2]['名前'] = e.parameter.mei || '';
                recordsI[ri2]['生年月日'] = e.parameter.dob || '';
                recordsI[ri2]['管理者権限'] = newRole;
                recordsI[ri2]['役員フラグ'] = e.parameter.isExecutive || '0';
                foundI = true;
                break;
              }
            }
            if (!foundI) {
              recordsI.push({ 'ユーザーID': targetIdI, '苗字': e.parameter.sei || '', '名前': e.parameter.mei || '', '生年月日': e.parameter.dob || '', '管理者権限': newRole, '役員フラグ': e.parameter.isExecutive || '0' });
            }
            userBodyI.editAsText().setText(serializeUserRecords(recordsI));
          }
          docI.saveAndClose();
          out = { success: true };
        }
      }
    } else if (action === 'adminUpdateLeaveBalance') {
      var listFileL = findMasterListFile();
      if (!listFileL) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var mainTextL = DocumentApp.openById(listFileL.getId()).getTabs()[0].asDocumentTab().getBody().getText();
        var targetIdL = (e.parameter.targetId || '').trim();
        var rowL = findMainListRow(mainTextL, targetIdL);
        if (!rowL) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          var yearL = (e.parameter.year || '').trim();
          var newValueL = (e.parameter.newValue || '0').trim();
          var targetFolderL = DriveApp.getFolderById(rowL.cols[3].trim());
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
      var listFileGD = findMasterListFile();
      if (!listFileGD) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var mainTextGD = DocumentApp.openById(listFileGD.getId()).getTabs()[0].asDocumentTab().getBody().getText();
        var targetIdGD = (e.parameter.targetId || '').trim();
        var rowGD = findMainListRow(mainTextGD, targetIdGD);
        if (!rowGD) {
          out = { error: '該当する社員IDが見つかりませんでした' };
        } else {
          var newDateGD = (e.parameter.newDate || '').trim();
          var newDaysGD = (e.parameter.newDays || '').trim();
          if (!/^\d{4}-\d{2}-\d{2}$/.test(newDateGD)) {
            out = { error: '日付の形式が正しくありません' };
          } else {
            var targetFolderGD = DriveApp.getFolderById(rowGD.cols[3].trim());
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
    } else if (action === 'updateCompanySettings') {
      var listFileCS = findMasterListFile();
      if (!listFileCS) {
        out = {error: '「管理情報」ドキュメントが見つかりません'};
      } else {
        var docCS = DocumentApp.openById(listFileCS.getId());
        var bodyCS = docCS.getTabs()[0].asDocumentTab().getBody();
        var linesCS = bodyCS.getText().split('\n');
        var updatesCS = {};
        try {
          updatesCS = JSON.parse(e.parameter.updates || '{}');
        } catch (parseErrCS) {
          updatesCS = {};
        }
        var sectionStartCS = -1;
        var sectionEndCS = linesCS.length;
        for (var csi = 0; csi < linesCS.length; csi++) {
          if (linesCS[csi].trim() === '会社共通設定') { sectionStartCS = csi; continue; }
          if (sectionStartCS >= 0 && linesCS[csi].trim() === '給料日特例') { sectionEndCS = csi; break; }
        }
        var existingCS = {};
        var orderCS = [];
        if (sectionStartCS >= 0) {
          for (var csj = sectionStartCS + 1; csj < sectionEndCS; csj++) {
            var rowCS = linesCS[csj].trim();
            if (!rowCS) continue;
            var idxCS = rowCS.indexOf(':');
            if (idxCS > 0) {
              var keyCS = rowCS.substring(0, idxCS).trim();
              existingCS[keyCS] = rowCS.substring(idxCS + 1).trim();
              orderCS.push(keyCS);
            }
          }
        }
        for (var ukCS in updatesCS) {
          if (existingCS[ukCS] === undefined) orderCS.push(ukCS);
          existingCS[ukCS] = updatesCS[ukCS];
        }
        var newSectionLinesCS = ['会社共通設定'];
        orderCS.forEach(function(k){ newSectionLinesCS.push(k + ':' + existingCS[k]); });
        newSectionLinesCS.push('');
        var beforeCS = sectionStartCS >= 0 ? linesCS.slice(0, sectionStartCS) : linesCS.slice(0, linesCS.length);
        var afterCS = sectionStartCS >= 0 ? linesCS.slice(sectionEndCS) : [];
        var finalLinesCS = beforeCS.concat(newSectionLinesCS).concat(afterCS);
        bodyCS.editAsText().setText(finalLinesCS.join('\n'));
        docCS.saveAndClose();
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
        var companySettings = {};
        var inCompanySection = false;
        var rows = listText.split('\n');
        for (var i = 0; i < rows.length; i++) {
          var row = rows[i].trim();
          if (!row) continue;
          if (row === '会社共通設定') { inCompanySection = true; continue; }
          if (inCompanySection) {
            var kv = row.split(':');
            if (kv.length >= 2) companySettings[kv[0].trim()] = kv.slice(1).join(':').trim();
            continue;
          }
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
  var listFile = findMasterListFile();
  if (!listFile) return;
  var doc = DocumentApp.openById(listFile.getId());
  var userBody = getTabBodyByTitle(doc, 'ユーザー情報');
  if (!userBody) return;
  var records = parseUserRecords(userBody.getText());
  var mainBody = doc.getTabs()[0].asDocumentTab().getBody();
  var mainLines = mainBody.getText().split('\n');
  var today = todayDateString();
  var changed = false;
  var mainChanged = false;
  var remainingRecords = [];

  for (var i = 0; i < records.length; i++) {
    var r = records[i];
    if (r['削除フラグ'] !== '1') { remainingRecords.push(r); continue; }

    var scheduledDate = r['削除予定日'];
    if (!scheduledDate) { remainingRecords.push(r); continue; }

    // 1か月前告知（まだ告知していなければ）
    if (r['告知済み'] !== '1') {
      var oneMonthBefore = addDaysToDateString(scheduledDate, -30);
      if (today >= oneMonthBefore) {
        var msg = r['削除日'] + 'に削除された' + r['ユーザーID'] + ':' + (r['苗字'] || '') + '　' + (r['名前'] || '') + 'のアカウント情報が、' + scheduledDate + 'に完全削除される予定です';
        var newsDoc = getOrCreateMasterDoc('ニュース');
        appendLine(newsDoc, new Date().getTime() + ':admin:critical:完全削除予告:' + msg.replace(/:/g, '：'));
        r['告知済み'] = '1';
        changed = true;
        logSystemEvent('user_delete_notice', r['ユーザーID'] + ' の完全削除予告を管理者に通知（予定日：' + scheduledDate + '）');
      }
    }

    // 完全削除の実行
    if (today >= scheduledDate) {
      var mainIdx = -1;
      var folderIdToTrash = null;
      for (var mi = 0; mi < mainLines.length; mi++) {
        var mrow = mainLines[mi].trim();
        if (!mrow) continue;
        var mcols = mrow.replace(/：/g, ':').split(':');
        if (mcols.length < 4) continue;
        if (mcols[1].trim() === r['ユーザーID']) { mainIdx = mi; folderIdToTrash = mcols[3].trim(); break; }
      }
      if (mainIdx >= 0) {
        try {
          var folder = DriveApp.getFolderById(folderIdToTrash);
          folder.setTrashed(true);
        } catch (errTrash) { /* フォルダが既に無い場合等は無視 */ }
        mainLines.splice(mainIdx, 1);
        mainChanged = true;
      }
      logSystemEvent('user_delete_completed', r['ユーザーID'] + ':' + (r['苗字'] || '') + (r['名前'] || '') + ' のアカウント情報を完全削除しました（削除理由：' + (r['削除理由'] || '') + '）');
      changed = true;
      // このレコードはremainingRecordsに含めない（完全削除）
      continue;
    }

    remainingRecords.push(r);
  }

  if (changed) {
    userBody.editAsText().setText(serializeUserRecords(remainingRecords));
  }
  if (mainChanged) {
    mainBody.editAsText().setText(mainLines.join('\n'));
  }
  if (changed || mainChanged) {
    doc.saveAndClose();
  }
}