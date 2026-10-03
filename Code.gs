// ===========================================================
// เวอร์ชันสมบูรณ์ รวม AIS + TRUE
// แก้ไข: TRUE ใช้ header row 3 (index 2) และ data เริ่มที่ row 4 (index 3)
//         AOR=col X (index 23), SmartQC=col Y (index 24) จากภาพจริง
// ===========================================================

// ── LINE Messaging API: Channel Access Token ของบอท SPE_SLA ──
var DEFAULT_LINE_TOKEN = 'YKtVKOIprzQoLKqB7foUkyxIwvzGaWxY/lnBmm4GaoJVNVDgbEUOTs8MOZRWBtEfzX8X6k0pX+pJSyave60Ka//baM6waKsQE/Ho43TkMod6YcyLcreDpjVC85MCXv7NxSj47Bh6bI2a2Xuls5hnkAdB04t89/1O/w1cDnyilFU=';
var DEFAULT_LINE_GROUP_ID = 'C9d136fee255c27308ede4164cad0e27d\nC42aae0c059a87a75d1b8166953108d70\nC3ade9979ac5d2b606210a02797b861b3'; // 3 กลุ่มเริ่มต้น: SPE-SLA, Super Star, TLN_AIS PATDOC

// ─────────────────────────────────────────────────────────────
// doGet: รวมเป็นอันเดียว — serve HTML + action=groupid
// (ถูกรวมกับ doGet ที่บรรทัด 609 ให้เหลืออันเดียว)
// ─────────────────────────────────────────────────────────────

// ─────────────────────────────────────────────────────────────
// doPost: LINE Webhook Receiver
//   - รับ events จาก LINE เมื่อมีข้อความในกลุ่ม
//   - บันทึก groupId อัตโนมัติ
//   - ถ้า source.type == 'group' และข้อความมี /groupid หรือ /idgroup → reply ด้วย Group ID
//   - รองรับ saveImportData และ sendLineAlert จาก Web App frontend
// ─────────────────────────────────────────────────────────────
function doPost(e) {
  if (!e || !e.postData || !e.postData.contents) {
    Logger.log('ℹ️ doPost: ไม่พบ postData (กรณีนี้เกิดจากการกดปุ่ม Run ใน Apps Script Editor โดยตรง ซึ่งเป็นปกติครับ เนื่องจาก Webhook จริงจะส่ง postData มาจาก LINE โดยอัตโนมัติ)');
    return ContentService.createTextOutput('No postData').setMimeType(ContentService.MimeType.TEXT);
  }
  try {
    var body = JSON.parse(e.postData.contents);

    // ── กรณี LINE Webhook Event ──
    if (body.events) {
      var events = body.events;
      for (var i = 0; i < events.length; i++) {
        var evt = events[i];
        var source = evt.source || {};

        // บันทึก groupId อัตโนมัติทันทีที่ Bot ได้รับ event จากกลุ่ม
        if (source.type === 'group' && source.groupId) {
          PropertiesService.getScriptProperties().setProperty('SAVED_LINE_GROUP_ID', source.groupId);
          Logger.log('✅ บันทึก Group ID: ' + source.groupId);

          // ถ้ามีข้อความ → reply Group ID กลับในกลุ่ม
          if (evt.type === 'message' && evt.message && evt.message.type === 'text') {
            var replyToken = evt.replyToken;
            var text = (evt.message.text || '').toLowerCase().trim();
            if (text.indexOf('/groupid') >= 0 || text.indexOf('groupid') >= 0 || text.indexOf('/idgroup') >= 0 || text.indexOf('idgroup') >= 0) {
              replyLineMessage_(replyToken, '🤖 SPE_SLA Bot\n\n✅ LINE Group ID:\n' + source.groupId + '\n\nกรุณาคัดลอก ID นี้ไปกรอกในหน้า Dashboard ที่ช่อง "LINE Group ID" ครับ');
            }
          }
        }
      }
      return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
    }

    // ── กรณี API call จาก Frontend (saveImportData / sendLineAlert) ──
    if (body.action === 'saveImportData' && body.operator && body.records) {
      return ContentService.createTextOutput(
        JSON.stringify(saveImportData(body.operator, body.records))
      ).setMimeType(ContentService.MimeType.JSON);
    }

    if (body.action === 'sendLineAlert' && body.token && body.message) {
      return ContentService.createTextOutput(
        JSON.stringify(sendLineAlert(body.token, body.message, body.imageBase64 || null, body.groupId || ''))
      ).setMimeType(ContentService.MimeType.JSON);
    }

  } catch(err) {
    Logger.log('doPost error: ' + err.message);
  }
  return ContentService.createTextOutput('OK').setMimeType(ContentService.MimeType.TEXT);
}

// ── ส่ง Reply Message กลับใน LINE Group ──
function replyLineMessage_(replyToken, text) {
  if (!replyToken) return;
  var payload = {
    replyToken: replyToken,
    messages: [{ type: 'text', text: text }]
  };
  UrlFetchApp.fetch('https://api.line.me/v2/bot/message/reply', {
    method: 'post',
    contentType: 'application/json',
    headers: { 'Authorization': 'Bearer ' + DEFAULT_LINE_TOKEN },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}


function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Executive Dashboard')
      .addItem('🚀 เปิด SLA Web App / Dashboard & Import', 'showDashboardModal')
      .addSeparator()
      .addItem('🔄 อัปเดตข้อมูลล่าสุด AIS (Refresh)', 'createAisDashboard')
      .addItem('🔄 อัปเดตข้อมูลล่าสุด TRUE (Refresh)', 'createTrueDashboard')
      .addSeparator()
      .addItem('📲 ส่งแจ้งเตือน LINE Group (AIS)', 'sendLineAlertAIS')
      .addItem('📲 ส่งแจ้งเตือน LINE Group (TRUE)', 'sendLineAlertTRUE')
      .addToUi();
}

function onEdit(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var cell = e.range.getA1Notation();
  if (sheetName === "Dashboard สรุปงาน" && cell === "E2") createAisDashboard();
  else if (sheetName === "Dashboard สรุปงาน TRUE" && cell === "E2") createTrueDashboard();
}

function toDateOnly_(value, tz) {
  if (!value) return null;
  if (value instanceof Date) {
    if (isNaN(value.getTime())) return null;
    var formatted = Utilities.formatDate(value, tz, "yyyy-MM-dd");
    var parts = formatted.split("-");
    return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
  }

  // กรณีเป็นตัวเลข Excel serial date เช่น 45558
  if (typeof value === "number" && value > 30000 && value < 60000) {
    var dt = new Date((value - 25569) * 86400000);
    if (!isNaN(dt.getTime())) {
      var formatted = Utilities.formatDate(dt, tz, "yyyy-MM-dd");
      var parts = formatted.split("-");
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    }
  }

  var s = String(value).trim();
  if (!s || s === "-" || s.toLowerCase() === "n/a") return null;

  // ตรวจสอบกรณีเป็นตัวเลขเดี่ยวๆ ในรูป string
  var num = Number(s);
  if (!isNaN(num)) {
    if (num > 30000 && num < 60000) {
      var dt = new Date((num - 25569) * 86400000);
      if (!isNaN(dt.getTime())) {
        var formatted = Utilities.formatDate(dt, tz, "yyyy-MM-dd");
        var parts = formatted.split("-");
        return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
      }
    }
    return null; // ตัวเลขที่ไม่ใช่ serial date (เช่น 1938, 0, 1) ให้ข้าม
  }

  // รองรับรูปแบบ DD/MM/YYYY หรือ DD-MM-YYYY
  var m1 = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (m1) {
    var day   = parseInt(m1[1], 10);
    var month = parseInt(m1[2], 10) - 1;
    var year  = parseInt(m1[3], 10);
    var dt = new Date(year, month, day);
    if (!isNaN(dt.getTime())) return dt;
  }

  // รองรับรูปแบบ YYYY/MM/DD หรือ YYYY-MM-DD
  var m2 = s.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})/);
  if (m2) {
    var year  = parseInt(m2[1], 10);
    var month = parseInt(m2[2], 10) - 1;
    var day   = parseInt(m2[3], 10);
    var dt = new Date(year, month, day);
    if (!isNaN(dt.getTime())) return dt;
  }

  var d = new Date(s);
  if (isNaN(d.getTime())) return null;
  var formatted = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  var parts = formatted.split("-");
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

function colIdxToLetter_(idx) {
  var s="", i=idx+1;
  while(i>0){var r=(i-1)%26; s=String.fromCharCode(65+r)+s; i=Math.floor((i-1)/26);}
  return s;
}

// ---------------------------------------------------------
// Helper: หาแถวที่เป็น header จริง (แถวที่มีคำว่า "DU ID" หรือ "DUID")
// ---------------------------------------------------------
function findHeaderRow_(data) {
  for (var r = 0; r < Math.min(10, data.length); r++) {
    for (var c = 0; c < data[r].length; c++) {
      var val = data[r][c] ? String(data[r][c]).trim().toLowerCase() : "";
      if (val === "du id" || val.replace(/\s/g, "").includes("duid")) {
        return r;
      }
    }
  }
  return 0; // หาไม่เจอ ใช้แถวแรกเป็น fallback
}

// ---------------------------------------------------------
// Helper: หา index คอลัมน์จาก header โดยใช้เงื่อนไขฟังก์ชัน ถ้าไม่เจอใช้ fallback
// ---------------------------------------------------------
function findColIndex_(headerRow, matchFn, fallbackIndex) {
  if (!headerRow) return fallbackIndex;
  for (var c = 0; c < headerRow.length; c++) {
    var val = headerRow[c] ? String(headerRow[c]).trim().toLowerCase() : "";
    if (matchFn(val)) return c;
  }
  return fallbackIndex;
}

// ---------------------------------------------------------
// ✅ AIS column indices: อิงตาม Old Code.gs เป็นหลัก (AE=30, AF=31, AH=33, AJ=35, AL=37)
// ---------------------------------------------------------
function getAisColumnIndices_(data) {
  var headerRowIdx = findHeaderRow_(data);
  var headers = data[headerRowIdx] || [];

  // โครงสร้างชีต AIS ตาม Old Code.gs (คอลัมน์ >= 38)
  if (headers.length >= 38) {
    var idxDuid = findColIndex_(headers, function(v) {
      return v === "du id" || v.replace(/\s/g, "").includes("duid");
    }, 0);

    return {
      headerRowIdx : headerRowIdx,
      idxDuid      : idxDuid,
      idxInstall   : 30, // Col AE - วันที่ติดตั้งเสร็จ
      idxSmartQc   : 31, // Col AF - Smart QC
      idxPat       : 33, // Col AH - PAT Subcon
      idxPatRemark : 35, // Col AJ - PAT Remark
      idxOwnerDoc  : 37  // Col AL - Owner Doc
    };
  } else {
    // กรณีชีตรูปแบบย่อ (เช่น 17 คอลัมน์ ISDP)
    var idxDuid = findColIndex_(headers, function(v) {
      return v === "du id" || v.replace(/\s/g, "").includes("duid");
    }, 0);
    var idxInstall = findColIndex_(headers, function(v) {
      return v.includes("actual start") || v.includes("actual end");
    }, 7);
    var idxSmartQc = findColIndex_(headers, function(v) {
      return v.includes("close time") || v.includes("smartqc");
    }, 12);
    var idxPat = findColIndex_(headers, function(v) {
      return v.includes("delivery attachment upload") || v.includes("attachment upload") || v.includes("pat date");
    }, 14);
    var idxOwnerDoc = findColIndex_(headers, function(v) {
      return v === "owner" || v.includes("owner");
    }, 9);
    var idxPatRemark = findColIndex_(headers, function(v) {
      return v === "remarks" || v.includes("remark");
    }, 16);

    return {
      headerRowIdx : headerRowIdx,
      idxDuid      : idxDuid,
      idxInstall   : idxInstall,
      idxSmartQc   : idxSmartQc,
      idxPat       : idxPat,
      idxPatRemark : idxPatRemark,
      idxOwnerDoc  : idxOwnerDoc
    };
  }
}

// ---------------------------------------------------------
// ✅ TRUE column indices: อิงตาม Old Code.gs (B=1, W=22, X=23, Y=24, AA=26, AC=28, AG=32)
// ---------------------------------------------------------
function getTrueColumnIndices_(data) {
  var headerRowIdx = findHeaderRow_(data);
  var headers = data[headerRowIdx] || [];

  var idxDuid = findColIndex_(headers, function(v) {
    return v === "du id" || v.replace(/\s/g, "").includes("duid");
  }, 1);

  var idxVerifyPhoto = findColIndex_(headers, function(v) {
    return v === "verify photo" || v.includes("verify photo");
  }, 22);

  var idxAor = findColIndex_(headers, function(v) {
    return v === "aor" || v.includes("aor");
  }, 23);

  var idxSmartQc = findColIndex_(headers, function(v) {
    return v.replace(/\s/g, "").includes("smartqc");
  }, 24);

  var idxPatRemark = findColIndex_(headers, function(v) {
    return v.includes("alarm category");
  }, 26);

  var idxPat = findColIndex_(headers, function(v) {
    return v.includes("pat site folder") && v.indexOf("power") === -1;
  }, 28);

  var idxOwnerDoc = findColIndex_(headers, function(v) {
    return v.includes("owner") || v.includes("pat owner");
  }, 32);

  return {
    headerRowIdx  : headerRowIdx,
    idxDuid       : idxDuid,
    idxVerifyPhoto: idxVerifyPhoto,
    idxAor        : idxAor,
    idxSmartQc    : idxSmartQc,
    idxPatRemark  : idxPatRemark,
    idxPat        : idxPat,
    idxOwnerDoc   : idxOwnerDoc
  };
}

// ---------------------------------------------------------
// 3. AIS Dashboard
// ✅ ดึงจากชีต AIS โดยตรง อิงตามการทำงานของ Old Code.gs
// ---------------------------------------------------------
function createAisDashboard() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(lockErr) { SpreadsheetApp.getUi().alert("⏳ กำลังอัปเดตอยู่ กรุณารอสักครู่"); return; }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = "Asia/Bangkok";
    var sourceSheet = ss.getSheetByName("AIS");
    if (!sourceSheet) { SpreadsheetApp.getUi().alert("⚠️ ไม่พบชีต 'AIS'"); return; }

    var dashSheet = _getOrCreateSheet_(ss, "Dashboard สรุปงาน");
    var selectedYear = _readAndClearDash_(dashSheet);

    var data = sourceSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return;

    var idx          = getAisColumnIndices_(data);
    var headerRowIdx = idx.headerRowIdx;
    var idxDuid      = idx.idxDuid;
    var idxInstall   = idx.idxInstall;
    var idxSmartQc   = idx.idxSmartQc;
    var idxPat       = idx.idxPat;
    var idxOwnerDoc  = idx.idxOwnerDoc;
    var idxPatRemark = idx.idxPatRemark;
    var startRow     = headerRowIdx + 1;

    var today = toDateOnly_(new Date(), ssTz);
    var availableYears = {};
    for (var i = startRow; i < data.length; i++) {
      var dd = toDateOnly_(data[i][idxInstall], ssTz);
      if (dd && !isNaN(dd.getTime())) availableYears[dd.getFullYear()] = true;
    }
    var yearList = Object.keys(availableYears).sort(function(a,b){return b-a;});
    yearList.unshift("ทั้งหมด");

    var acc = _newAccumulators_();
    var C   = _colors_();

    for (var i = startRow; i < data.length; i++) {
      var row = data[i];
      if (row.length <= idxPatRemark) continue;

      var installDate = toDateOnly_(row[idxInstall], ssTz);
      if (!installDate) continue;
      if (selectedYear !== "ทั้งหมด" && installDate.getFullYear().toString() !== selectedYear) continue;
      acc.totalInstall++;

      var duidVal    = row[idxDuid]     ? String(row[idxDuid]).trim()     : "ไม่ระบุ DUID";
      var ownerDoc   = row[idxOwnerDoc] ? String(row[idxOwnerDoc]).trim() : "-";
      var smartQcDate = toDateOnly_(row[idxSmartQc], ssTz);
      var patDate     = toDateOnly_(row[idxPat],     ssTz);
      var patRemark   = row[idxPatRemark] ? String(row[idxPatRemark]).trim() : "";

      _calcSQC_(acc, today, installDate, smartQcDate, duidVal, ownerDoc, 3, "");
      _calcPAT_(acc, today, installDate, smartQcDate, patDate, duidVal, ownerDoc);

      // ตรวจสอบงาน Rework / PAT Not Pass: ข้อความ remark ที่ไม่ใช่ Pass / First Time Pass
      var remLower = patRemark.toLowerCase().trim();
      var isNotPass = false;
      if (remLower && remLower !== "pass" && remLower !== "first time pass" && remLower !== "n/a") {
        isNotPass = true;
      } else if (remLower.includes("not pass") || remLower.includes("reject") || remLower.includes("revise") || remLower.includes("fail") || remLower.includes("แก้ไข")) {
        isNotPass = true;
      }

      if (isNotPass) {
        var od = Math.max(0, patDate ? Math.floor((today - patDate) / 86400000) : Math.floor((today - installDate) / 86400000));
        acc.listPatNotPass.push([duidVal, od + " วัน", C.ORANGE_BG, C.ORANGE_FG, ownerDoc]);
      }
    }

    function getAvg(t,c){return c>0?(t/c).toFixed(1):"0.0";}
    var debugLine2 = " | Install=" + colIdxToLetter_(idxInstall) +
                     " SQC=" + colIdxToLetter_(idxSmartQc) + " PAT=" + colIdxToLetter_(idxPat) +
                     " Owner=" + colIdxToLetter_(idxOwnerDoc) + " (hdr row " + (headerRowIdx + 1) + ")";

    _buildDashboard(dashSheet, {
      title: "AIS INSTALLATION & SLA DASHBOARD", selectedYear: selectedYear, yearList: yearList, totalInstall: acc.totalInstall,
      ...acc, sqcCardLabel: "EXECUTIVE SUMMARY: SMART QC", patCardLabel: "EXECUTIVE SUMMARY: PAT SUBCON SUBMIT",
      pendingSqcTitle: "ACTION REQUIRED: SMART QC", pendingPatTitle: "ACTION REQUIRED: PAT SUBCON",
      notPassTitle: "REWORK REQUIRED: PAT NOT PASS", debugLine2: debugLine2, getAvg: getAvg
    });
  } catch(err) {
    Logger.log(err.message + "\n" + err.stack);
    SpreadsheetApp.getUi().alert("⚠️ Error:\n" + err.message);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------
// 4. TRUE Dashboard
// ✅ ดึงจากชีต True โดยตรง อิงตามการทำงานของ Old Code.gs
// ---------------------------------------------------------
function createTrueDashboard() {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); }
  catch(lockErr) { SpreadsheetApp.getUi().alert("⏳ กำลังอัปเดตอยู่ กรุณารอสักครู่"); return; }
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = "Asia/Bangkok";
    var sourceSheet = ss.getSheetByName("True");
    if (!sourceSheet) { SpreadsheetApp.getUi().alert("⚠️ ไม่พบชีต 'True'"); return; }

    var dashSheet = _getOrCreateSheet_(ss, "Dashboard สรุปงาน TRUE");
    var selectedYear = _readAndClearDash_(dashSheet);

    var data = sourceSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return;

    var idx            = getTrueColumnIndices_(data);
    var headerRowIdx   = idx.headerRowIdx;
    var idxDuid        = idx.idxDuid;
    var idxVerifyPhoto = idx.idxVerifyPhoto;
    var idxAor         = idx.idxAor;
    var idxSmartQc     = idx.idxSmartQc;
    var idxPat         = idx.idxPat;
    var idxPatRemark   = idx.idxPatRemark;
    var idxOwnerDoc    = idx.idxOwnerDoc;
    var startRow       = headerRowIdx + 1;

    var today = toDateOnly_(new Date(), ssTz);
    var availableYears = {};
    for (var i = startRow; i < data.length; i++) {
      var dd = toDateOnly_(data[i][idxVerifyPhoto], ssTz);
      if (dd && !isNaN(dd.getTime())) availableYears[dd.getFullYear()] = true;
    }
    var yearList = Object.keys(availableYears).sort(function(a,b){return b-a;});
    yearList.unshift("ทั้งหมด");

    var acc = _newAccumulators_();
    var C   = _colors_();

    for (var i = startRow; i < data.length; i++) {
      var row = data[i];
      if (row.length <= idxPatRemark) continue;

      var installDate = toDateOnly_(row[idxVerifyPhoto], ssTz);
      if (!installDate) continue;
      if (selectedYear !== "ทั้งหมด" && installDate.getFullYear().toString() !== selectedYear) continue;
      acc.totalInstall++;

      var duidVal   = row[idxDuid]      ? String(row[idxDuid]).trim()     : "ไม่ระบุ DUID";
      var ownerDoc  = row[idxOwnerDoc]  ? String(row[idxOwnerDoc]).trim() : "-";
      var patRemark = row[idxPatRemark] ? String(row[idxPatRemark]).trim(): "";

      var aorDate     = toDateOnly_(row[idxAor],     ssTz);
      var smartQcDate = toDateOnly_(row[idxSmartQc], ssTz);
      var patDate     = toDateOnly_(row[idxPat],     ssTz);

      // Smart QC: ถ้ามี AOR ให้เช็คทั้งคู่ หรือถ้าไม่มี AOR ให้เช็ค SmartQC
      var sqcDone = (idxAor !== undefined && row[idxAor] !== undefined && row[idxAor] !== "") ? (aorDate && smartQcDate) : !!smartQcDate;
      var missingLabel = (!aorDate ? " ⚠️AOR" : "") + (!smartQcDate ? " ⚠️SQC" : "");
      _calcSQC_(acc, today, installDate, sqcDone ? (smartQcDate || aorDate) : null, duidVal, ownerDoc, 3, missingLabel);

      // PAT
      _calcPAT_(acc, today, installDate, smartQcDate, patDate, duidVal, ownerDoc);

      // Alarm Rework
      var hasAlarm = patRemark !== "" && patRemark.toLowerCase().includes("alarm") && !patRemark.toLowerCase().includes("no alarm");
      if (hasAlarm) {
        var od = Math.max(0, patDate ? Math.floor((today - patDate) / 86400000) : Math.floor((today - installDate) / 86400000));
        acc.listPatNotPass.push([duidVal, od + " วัน", C.ORANGE_BG, C.ORANGE_FG, ownerDoc]);
      }
    }

    function getAvg(t,c){return c>0?(t/c).toFixed(1):"0.0";}
    var debugLine2 = " | VP=" + colIdxToLetter_(idxVerifyPhoto) +
                     (idxAor !== undefined ? " AOR=" + colIdxToLetter_(idxAor) : "") +
                     " SQC=" + colIdxToLetter_(idxSmartQc) + " PAT=" + colIdxToLetter_(idxPat) +
                     " Alarm=" + colIdxToLetter_(idxPatRemark) + " Owner=" + colIdxToLetter_(idxOwnerDoc) +
                     " (hdr row " + (headerRowIdx + 1) + ")";

    _buildDashboard(dashSheet, {
      title: "TRUE INSTALLATION & SLA DASHBOARD", selectedYear: selectedYear, yearList: yearList, totalInstall: acc.totalInstall,
      ...acc, sqcCardLabel: "EXECUTIVE SUMMARY: 08.1 SmartQC+AOR", patCardLabel: "EXECUTIVE SUMMARY: 14.1 A129 PAT Site Folder",
      pendingSqcTitle: "ACTION REQUIRED: 08.1 SmartQC+AOR", pendingPatTitle: "ACTION REQUIRED: 14.1 A129 PAT Site Folder",
      notPassTitle: "REWORK REQUIRED: ALARM FOUND", debugLine2: debugLine2, getAvg: getAvg
    });
  } catch(err) {
    Logger.log(err.message + "\n" + err.stack);
    SpreadsheetApp.getUi().alert("⚠️ Error:\n" + err.message);
  } finally {
    lock.releaseLock();
  }
}

// ---------------------------------------------------------
// helpers ใช้ร่วมกัน
// ---------------------------------------------------------
function _colors_() {
  return { RED_BG:"#fee2e2",RED_FG:"#dc2626", BLUE_BG:"#e0f2fe",BLUE_FG:"#0284c7", ORANGE_BG:"#ffedd5",ORANGE_FG:"#ea580c" };
}

function _newAccumulators_() {
  return {
    totalInstall:0,
    doneSQC_InSLA:0,totalDays_doneSQC_InSLA:0,doneSQC_OverSLA:0,totalDays_doneSQC_OverSLA:0,
    pendingSQC_InSLA:0,totalDays_pendingSQC_InSLA:0,pendingSQC_OverSLA:0,totalDays_pendingSQC_OverSLA:0,
    donePat_InSLA:0,totalDays_donePat_InSLA:0,donePat_OverSLA:0,totalDays_donePat_OverSLA:0,
    pendingPat_InSLA:0,totalDays_pendingPat_InSLA:0,pendingPat_OverSLA:0,totalDays_pendingPat_OverSLA:0,
    listPendingSmartQc:[],listPendingPat:[],listPatNotPass:[]
  };
}

// ✅ คำนวณ Smart QC
//    sqcDateDone = null หรือ วันที่ในอนาคต → Pending, มีค่าในอดีต/วันนี้ → Done
function _calcSQC_(acc, today, installDate, sqcDateDone, duidVal, ownerDoc, slaDays, missingLabel) {
  var C = _colors_();
  var isDone = sqcDateDone && (sqcDateDone <= today);
  if (!isDone) {
    var diff = Math.max(0, Math.floor((today - installDate) / 86400000));
    if (diff > slaDays) {
      acc.pendingSQC_OverSLA++; acc.totalDays_pendingSQC_OverSLA += diff;
      acc.listPendingSmartQc.push([duidVal, diff + " วัน" + (missingLabel || ""), C.RED_BG, C.RED_FG, ownerDoc]);
    } else {
      acc.pendingSQC_InSLA++; acc.totalDays_pendingSQC_InSLA += diff;
      acc.listPendingSmartQc.push([duidVal, diff + " วัน" + (missingLabel || ""), C.BLUE_BG, C.BLUE_FG, ownerDoc]);
    }
  } else {
    var diff = Math.max(0, Math.floor((sqcDateDone - installDate) / 86400000));
    if (diff > slaDays) { acc.doneSQC_OverSLA++; acc.totalDays_doneSQC_OverSLA += diff; }
    else                { acc.doneSQC_InSLA++;   acc.totalDays_doneSQC_InSLA += diff; }
  }
}

// ✅ คำนวณ PAT
//    patDate = null หรือ วันที่ในอนาคต (เช่น แพลนไว้ยังไม่ถึง) → Pending, มีค่าในอดีต/วันนี้ → Done
function _calcPAT_(acc, today, installDate, smartQcDate, patDate, duidVal, ownerDoc) {
  var C = _colors_();
  var isDone = patDate && (patDate <= today);
  if (!isDone) {
    var baseDate = smartQcDate || installDate;
    if (baseDate) {
      var diff = Math.max(0, Math.floor((today - baseDate) / 86400000));
      if (diff > 5) {
        acc.pendingPat_OverSLA++; acc.totalDays_pendingPat_OverSLA += diff;
        acc.listPendingPat.push([duidVal, diff + " วัน", C.RED_BG, C.RED_FG, ownerDoc]);
      } else {
        acc.pendingPat_InSLA++; acc.totalDays_pendingPat_InSLA += diff;
        acc.listPendingPat.push([duidVal, diff + " วัน", C.BLUE_BG, C.BLUE_FG, ownerDoc]);
      }
    }
  } else {
    var base = smartQcDate || installDate;
    var diff = Math.max(0, Math.floor((patDate - base) / 86400000));
    if (diff > 5) { acc.donePat_OverSLA++; acc.totalDays_donePat_OverSLA += diff; }
    else          { acc.donePat_InSLA++;   acc.totalDays_donePat_InSLA += diff; }
  }
}

function _getOrCreateSheet_(ss, name) {
  var sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  return sh;
}

function _readAndClearDash_(dashSheet) {
  var selectedYear = "ทั้งหมด";
  try { var v=dashSheet.getRange("E2").getValue(); if(v&&v!=="") selectedYear=v.toString().trim(); } catch(e){}
  dashSheet.clear();
  var charts = dashSheet.getCharts();
  for(var i=0;i<charts.length;i++) dashSheet.removeChart(charts[i]);
  dashSheet.setHiddenGridlines(true);
  return selectedYear;
}

// ---------------------------------------------------------
// 5. _buildDashboard: วาด UI
// ---------------------------------------------------------
function _buildDashboard(dashSheet, o) {
  var getAvg=o.getAvg;
  var C_BG="#f8fafc",C_CARD="#ffffff",C_HEAD_BG="#0f172a",C_HEAD_TXT="#ffffff",C_BORDER="#cbd5e1";
  var maxCol=11,maxRow=250;

  dashSheet.getRange(1,1,maxRow,maxCol).setBackground(C_BG).setFontFamily("Arial").setVerticalAlignment("middle").setWrap(true);
  dashSheet.setRowHeights(1,maxRow,28); dashSheet.setRowHeight(1,35); dashSheet.setRowHeight(2,35);

  dashSheet.getRange("A1:K2").setBackground(C_HEAD_BG).setFontColor(C_HEAD_TXT);
  dashSheet.getRange("A1").setValue(o.title).setFontSize(18).setFontWeight("bold").setVerticalAlignment("bottom");
  var line2="ข้อมูลอัปเดต ณ วันที่: "+Utilities.formatDate(new Date(),"GMT+7","dd/MM/yyyy HH:mm")+(o.debugLine2||"");
  dashSheet.getRange("A2").setValue(line2).setFontSize(9).setFontWeight("bold").setFontColor("#94a3b8").setVerticalAlignment("top");

  dashSheet.getRange("C2").setValue("📅 เลือกปี (Year):").setFontSize(11).setFontWeight("bold").setFontColor("#94a3b8").setHorizontalAlignment("right");
  var rule=SpreadsheetApp.newDataValidation().requireValueInList(o.yearList).build();
  dashSheet.getRange("E2").setDataValidation(rule).setValue(o.selectedYear).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true,true,true,true,false,false,"#334155",SpreadsheetApp.BorderStyle.SOLID);
  dashSheet.getRange("E1").setValue("🔄 Auto-Refresh เมื่อเปลี่ยนปี").setFontSize(9).setFontWeight("bold").setFontColor("#475569").setVerticalAlignment("bottom").setHorizontalAlignment("center");
  dashSheet.getRange("I1:K2").merge().setValue(o.totalInstall+" ไซต์").setFontSize(26).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
  dashSheet.getRange("G1:H2").merge().setValue("จำนวนงานติดตั้งเสร็จ\n(ตามปีที่เลือก)").setFontSize(11).setFontWeight("bold").setFontColor("#cbd5e1").setHorizontalAlignment("right").setVerticalAlignment("middle");

  function createCard(rangeStr,title,dataArr){
    var rng=dashSheet.getRange(rangeStr);
    rng.setBackground(C_CARD).setBorder(true,true,true,true,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
    dashSheet.getRange(rng.getRow(),rng.getColumn(),1,3).merge().setValue(title).setBackground("#f8fafc").setFontColor("#0f172a").setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center").setBorder(null,null,true,null,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
    var dr=dashSheet.getRange(rng.getRow()+1,rng.getColumn(),5,3);
    dr.setValues(dataArr).setFontWeight("bold").setFontColor("#1e293b").setFontSize(11);
    dashSheet.getRange(rng.getRow()+1,rng.getColumn()+1,5,2).setHorizontalAlignment("center");
    dr.setBorder(null,null,null,null,false,true,"#f1f5f9",SpreadsheetApp.BorderStyle.SOLID);
  }
  createCard("A4:C9",o.sqcCardLabel,[
    ["สถานะงาน","จำนวนไซต์","Aging เฉลี่ย (วัน)"],
    ["🟢 ปิดตามกำหนด (Done in SLA)",           o.doneSQC_InSLA,    Number(getAvg(o.totalDays_doneSQC_InSLA,   o.doneSQC_InSLA))],
    ["🟡 ปิดงานล่าช้า (Done Late)",             o.doneSQC_OverSLA,  Number(getAvg(o.totalDays_doneSQC_OverSLA, o.doneSQC_OverSLA))],
    ["🔵 รอตรวจสอบ (Pending in SLA)",          o.pendingSQC_InSLA, Number(getAvg(o.totalDays_pendingSQC_InSLA,o.pendingSQC_InSLA))],
    ["🔴 ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)",o.pendingSQC_OverSLA,Number(getAvg(o.totalDays_pendingSQC_OverSLA,o.pendingSQC_OverSLA))]
  ]);
  createCard("E4:G9",o.patCardLabel,[
    ["สถานะงาน","จำนวนไซต์","Aging เฉลี่ย (วัน)"],
    ["🟢 ส่งตามกำหนด (Done in SLA)",            o.donePat_InSLA,    Number(getAvg(o.totalDays_donePat_InSLA,   o.donePat_InSLA))],
    ["🟡 ส่งงานล่าช้า (Done Late)",              o.donePat_OverSLA,  Number(getAvg(o.totalDays_donePat_OverSLA, o.donePat_OverSLA))],
    ["🔵 รอส่งงาน (Pending in SLA)",            o.pendingPat_InSLA, Number(getAvg(o.totalDays_pendingPat_InSLA,o.pendingPat_InSLA))],
    ["🔴 ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", o.pendingPat_OverSLA,Number(getAvg(o.totalDays_pendingPat_OverSLA,o.pendingPat_OverSLA))]
  ]);

  dashSheet.getRange("Y1:Z30").clearContent();
  var sqcData=[["Status","Count"],
    ["ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)",o.pendingSQC_OverSLA>0?o.pendingSQC_OverSLA:0.0001],
    ["ปิดงานล่าช้า (Done Late)",            o.doneSQC_OverSLA>0?o.doneSQC_OverSLA:0.0001],
    ["ปิดตามกำหนด (Done in SLA)",           o.doneSQC_InSLA>0?o.doneSQC_InSLA:0.0001],
    ["รอตรวจสอบ (Pending in SLA)",          o.pendingSQC_InSLA>0?o.pendingSQC_InSLA:0.0001]];
  var patData=[["Status","Count"],
    ["ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)",o.pendingPat_OverSLA>0?o.pendingPat_OverSLA:0.0001],
    ["ส่งงานล่าช้า (Done Late)",            o.donePat_OverSLA>0?o.donePat_OverSLA:0.0001],
    ["ส่งตามกำหนด (Done in SLA)",           o.donePat_InSLA>0?o.donePat_InSLA:0.0001],
    ["รอส่งงาน (Pending in SLA)",           o.pendingPat_InSLA>0?o.pendingPat_InSLA:0.0001]];
  var slices={0:{color:'#ef4444'},1:{color:'#eab308'},2:{color:'#22c55e'},3:{color:'#3b82f6'}};
  dashSheet.getRange(1,25,sqcData.length,2).setValues(sqcData);
  dashSheet.getRange(15,25,patData.length,2).setValues(patData);
  dashSheet.getRange("A11:C11").merge().setValue("สัดส่วนผลงานภาพรวม 100% (Smart QC)").setFontSize(13).setFontWeight("bold").setFontColor("#1e293b").setHorizontalAlignment("center");
  dashSheet.getRange("E11:G11").merge().setValue("สัดส่วนผลงานภาพรวม 100% (PAT Subcon)").setFontSize(13).setFontWeight("bold").setFontColor("#1e293b").setHorizontalAlignment("center");

  function insertDonut(dr,ar,ac){
    dashSheet.insertChart(dashSheet.newChart().setChartType(Charts.ChartType.PIE).addRange(dr).setPosition(ar,ac,0,0)
      .setOption('pieHole',0.45).setOption('pieSliceText','percentage')
      .setOption('pieSliceTextStyle',{color:'#ffffff',fontSize:14,bold:true,fontName:'Arial'})
      .setOption('slices',slices)
      .setOption('legend',{position:'bottom',textStyle:{color:'#0f172a',fontSize:13,bold:true,fontName:'Arial'}})
      .setOption('pieSliceBorderColor','#ffffff').setOption('backgroundColor',{fill:'#f8fafc'})
      .setOption('width',520).setOption('height',450).setOption('chartArea',{left:'5%',top:'2%',width:'90%',height:'70%'}).build());
  }
  insertDonut(dashSheet.getRange(1,25,sqcData.length,2),12,1);
  insertDonut(dashSheet.getRange(15,25,patData.length,2),12,5);

  var lsr=30;
  function wList(sc,title,list,empty){
    dashSheet.getRange(lsr,sc,1,2).merge().setValue(title).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true,true,true,true,null,null,"#1e293b",SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange(lsr+1,sc).setValue("Site DUID"); dashSheet.getRange(lsr+1,sc+1).setValue("Aging");
    dashSheet.getRange(lsr+1,sc,1,2).setFontColor("#1e293b").setFontWeight("bold").setBackground("#f8fafc").setHorizontalAlignment("center").setBorder(false,true,true,true,true,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
    if(list.length>0){
      for(var k=0;k<list.length;k++) dashSheet.getRange(lsr+2+k,sc,1,2).setValues([[list[k][0],list[k][1]]]).setBackground(list[k][2]).setFontColor(list[k][3]).setFontWeight("bold").setBorder(false,true,true,true,true,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(lsr+2,sc+1,list.length,1).setHorizontalAlignment("center").setFontWeight("bold");
    } else { dashSheet.getRange(lsr+2,sc,1,2).merge().setValue(empty).setBackground(C_CARD).setFontColor("#94a3b8").setFontWeight("bold").setHorizontalAlignment("center").setBorder(false,true,true,true,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID); }
  }
  function wListOwner(sc,title,list,empty){
    dashSheet.getRange(lsr,sc,1,2).merge().setValue(title).setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true,true,true,true,null,null,"#1e293b",SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange(lsr,sc+2).setValue("OWNER DOC").setBackground("#ca8a04").setFontColor("#ffffff").setFontWeight("bold").setHorizontalAlignment("center").setBorder(true,true,true,true,null,null,"#ca8a04",SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange(lsr+1,sc).setValue("Site DUID"); dashSheet.getRange(lsr+1,sc+1).setValue("Aging");
    dashSheet.getRange(lsr+1,sc,1,2).setFontColor("#1e293b").setFontWeight("bold").setBackground("#f8fafc").setHorizontalAlignment("center").setBorder(false,true,true,true,true,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange(lsr+1,sc+2).setValue("Assign").setFontColor("#1e293b").setFontWeight("bold").setBackground("#fef9c3").setHorizontalAlignment("center").setBorder(false,true,true,true,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
    if(list.length>0){
      for(var k=0;k<list.length;k++){
        dashSheet.getRange(lsr+2+k,sc,1,2).setValues([[list[k][0],list[k][1]]]).setBackground(list[k][2]).setFontColor(list[k][3]).setFontWeight("bold").setBorder(false,true,true,true,true,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
        dashSheet.getRange(lsr+2+k,sc+2).setValue(list[k][4]||"-").setBackground("#dcfce7").setFontColor("#15803d").setFontWeight("bold").setHorizontalAlignment("center").setBorder(false,true,true,true,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
      }
      dashSheet.getRange(lsr+2,sc+1,list.length,1).setHorizontalAlignment("center").setFontWeight("bold");
    } else {
      dashSheet.getRange(lsr+2,sc,1,2).merge().setValue(empty).setBackground(C_CARD).setFontColor("#94a3b8").setFontWeight("bold").setHorizontalAlignment("center").setBorder(false,true,true,true,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(lsr+2,sc+2).setValue("-").setBackground("#dcfce7").setFontColor("#15803d").setHorizontalAlignment("center").setBorder(false,true,true,true,false,false,C_BORDER,SpreadsheetApp.BorderStyle.SOLID);
    }
  }

  wList(1,     o.pendingSqcTitle, o.listPendingSmartQc, "All sites are up to date 🎉");
  wListOwner(5,o.pendingPatTitle, o.listPendingPat,     "All sites are up to date 🎉");
  wListOwner(9,o.notPassTitle,    o.listPatNotPass,      "No rework required 🎉");

  dashSheet.setColumnWidth(1,320);dashSheet.setColumnWidth(2,90); dashSheet.setColumnWidth(3,150);
  dashSheet.setColumnWidth(4,20); dashSheet.setColumnWidth(5,320);dashSheet.setColumnWidth(6,90);
  dashSheet.setColumnWidth(7,150);dashSheet.setColumnWidth(8,20); dashSheet.setColumnWidth(9,300);
  dashSheet.setColumnWidth(10,90);dashSheet.setColumnWidth(11,150);

  var lastRow=lsr+2+Math.max(o.listPendingSmartQc.length,o.listPendingPat.length,o.listPatNotPass.length);
  if(lastRow<maxRow) dashSheet.getRange(lastRow+1,1,maxRow-lastRow,maxCol).setBackground(C_BG).setBorder(false,false,false,false,false,false);
}

// ===========================================================
// 6. Web App, Import Handler & LINE Alert Integration
// ===========================================================

/**
 * doGet: serve Web App + รองรับ ?action=groupid
 */
function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  if (action === 'groupid') {
    var gid = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || DEFAULT_LINE_GROUP_ID;
    return ContentService.createTextOutput('LINE_GROUP_ID=' + gid)
      .setMimeType(ContentService.MimeType.TEXT);
  }
  // Serve Dashboard HTML
  try {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('AIS & TRUE INSTALLATION & SLA DASHBOARD')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  } catch (err) {
    return ContentService.createTextOutput('Error: ' + err.message)
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

/**
 * เปิด Modal Dialog ขนาดใหญ่ใน Google Sheets
 */
function showDashboardModal() {
  var html = HtmlService.createHtmlOutputFromFile('Index')
      .setWidth(1200)
      .setHeight(850);
  SpreadsheetApp.getUi().showModalDialog(html, '📊 AIS & TRUE SLA Dashboard & Automation');
}

/**
 * เปิด Sidebar ใน Google Sheets
 */
function showDashboardSidebar() {
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('📊 AIS & TRUE Dashboard');
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * บันทึกข้อมูลที่ Import จาก Excel ลงใน Google Sheet (AIS หรือ True)
 * แล้วสั่งรีเฟรช Dashboard อัตโนมัติ
 */
function saveImportData(operator, records) {
  if (!records || records.length === 0) return { success: false, message: 'ไม่มีข้อมูลนำเข้า' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetName = (operator === 'AIS') ? 'AIS' : 'True';
  var sheet = ss.getSheetByName(sheetName) || ss.getSheetByName(operator === 'AIS' ? '56A0S0Q' : '56A0UPS');
  if (!sheet) return { success: false, message: 'ไม่พบชีต: ' + sheetName };
  var activeSheetName = sheet.getName();
  var isNewFormat = (activeSheetName === '56A0S0Q' || activeSheetName === '56A0UPS');

  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), isNewFormat ? (operator === 'AIS' ? 45 : 70) : 40);
  if (lastRow < 2) return { success: false, message: 'ชีตว่างเปล่า ไม่มีข้อมูลเดิม' };

  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var data = range.getValues();
  var duidRowMap = {};
  
  // ── AIS ──
  if (operator === 'AIS') {
    var duidCol = 0; // Col A
    var startRow = isNewFormat ? 2 : 2; // ข้าม header และ subtotal
    for (var r = startRow; r < data.length; r++) {
      var d = data[r][duidCol] ? String(data[r][duidCol]).trim() : '';
      if (d) duidRowMap[d] = r; // 0-based array index
    }

    var appendRows = [];
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var duid = rec.duid;
      if (!duid) continue;

      var targetIdx = duidRowMap[duid];
      if (isNewFormat) {
        if (targetIdx !== undefined) {
          // อัปเดตในหน่วยความจำ (In-Memory) รวดเร็วระดับมิลลิวินาที
          if (rec.installDate) data[targetIdx][8] = rec.installDate;   // col I (index 8)
          if (rec.smartQcDate) data[targetIdx][20] = rec.smartQcDate; // col U (index 20)
          if (rec.patDate) data[targetIdx][30] = rec.patDate;         // col AE (index 30)
          if (rec.owner && rec.owner !== '-') data[targetIdx][31] = rec.owner; // col AF (index 31)
          if (rec.patRemark) data[targetIdx][38] = rec.patRemark;     // col AM (index 38)
          if (rec.patStatus) data[targetIdx][39] = rec.patStatus;     // col AN (index 39)
        } else {
          var newRow = new Array(lastCol).fill('');
          newRow[0] = duid;
          newRow[8] = rec.installDate || '';
          newRow[20] = rec.smartQcDate || '';
          newRow[30] = rec.patDate || '';
          newRow[31] = rec.owner || '-';
          newRow[38] = rec.patRemark || '';
          newRow[39] = rec.patStatus || '';
          appendRows.push(newRow);
        }
      } else {
        if (targetIdx !== undefined) {
          if (rec.installDate) data[targetIdx][30] = rec.installDate;
          if (rec.smartQcDate) data[targetIdx][31] = rec.smartQcDate;
          if (rec.patDate) data[targetIdx][33] = rec.patDate;
          if (rec.patRemark) data[targetIdx][35] = rec.patRemark;
          if (rec.patStatus) data[targetIdx][36] = rec.patStatus;
          if (rec.owner && rec.owner !== '-') data[targetIdx][37] = rec.owner;
        } else {
          var newRow = new Array(lastCol).fill('');
          newRow[0] = duid;
          newRow[30] = rec.installDate || '';
          newRow[31] = rec.smartQcDate || '';
          newRow[33] = rec.patDate || '';
          newRow[35] = rec.patRemark || '';
          newRow[36] = rec.patStatus || '';
          newRow[37] = rec.owner || '-';
          appendRows.push(newRow);
        }
      }
    }

    // เขียนข้อมูลอัปเดตลงชีตครั้งเดียวรวดเดียว (1 Single Write)
    range.setValues(data);

    // หากมีแถวใหม่ที่เพิ่มเข้ามา เขียนต่อท้ายใน 1 คำสั่ง
    if (appendRows.length > 0) {
      sheet.getRange(lastRow + 1, 1, appendRows.length, lastCol).setValues(appendRows);
    }

    createAisDashboard();

  } else {
    // ── TRUE ──
    var duidCol = isNewFormat ? 0 : 1;
    var startRow = isNewFormat ? 2 : 3;
    for (var r = startRow; r < data.length; r++) {
      var d = data[r][duidCol] ? String(data[r][duidCol]).trim() : '';
      if (d) duidRowMap[d] = r;
    }

    var appendRows = [];
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var duid = rec.duid;
      if (!duid) continue;

      var targetIdx = duidRowMap[duid];
      if (isNewFormat) {
        if (targetIdx !== undefined) {
          if (rec.smartQcDate) data[targetIdx][36] = rec.smartQcDate; // col AK (index 36)
          if (rec.aorDate) data[targetIdx][40] = rec.aorDate;         // col AO (index 40)
          if (rec.alarmRemark) data[targetIdx][44] = rec.alarmRemark; // col AS (index 44)
          if (rec.patDate) data[targetIdx][46] = rec.patDate;         // col AU (index 46)
          if (rec.owner && rec.owner !== '-') data[targetIdx][47] = rec.owner; // col AV (index 47)
          if (rec.installDate) data[targetIdx][66] = rec.installDate; // col BO (index 66)
        } else {
          var newRow = new Array(lastCol).fill('');
          newRow[0] = duid;
          newRow[36] = rec.smartQcDate || '';
          newRow[40] = rec.aorDate || '';
          newRow[44] = rec.alarmRemark || '';
          newRow[46] = rec.patDate || '';
          newRow[47] = rec.owner || '-';
          newRow[66] = rec.installDate || '';
          appendRows.push(newRow);
        }
      } else {
        if (targetIdx !== undefined) {
          if (rec.installDate) data[targetIdx][22] = rec.installDate;
          if (rec.aorDate) data[targetIdx][23] = rec.aorDate;
          if (rec.smartQcDate) data[targetIdx][24] = rec.smartQcDate;
          if (rec.alarmRemark) data[targetIdx][26] = rec.alarmRemark;
          if (rec.patDate) data[targetIdx][28] = rec.patDate;
          if (rec.owner && rec.owner !== '-') data[targetIdx][32] = rec.owner;
        } else {
          var newRow = new Array(lastCol).fill('');
          newRow[1] = duid;
          newRow[22] = rec.installDate || '';
          newRow[23] = rec.aorDate || '';
          newRow[24] = rec.smartQcDate || '';
          newRow[26] = rec.alarmRemark || '';
          newRow[28] = rec.patDate || '';
          newRow[32] = rec.owner || '-';
          appendRows.push(newRow);
        }
      }
    }

    range.setValues(data);

    if (appendRows.length > 0) {
      sheet.getRange(lastRow + 1, 1, appendRows.length, lastCol).setValues(appendRows);
    }

    createTrueDashboard();
  }

  return { success: true, count: records.length, message: 'นำเข้าข้อมูล ' + operator + ' สำเร็จแล้ว' };
}

/**
 * นำเข้าทั้งตาราง 2D Array จากไฟล์ Excel วางลงในชีตเป้าหมายโดยตรงใน 1 วินาที
 */
function importFullSheet(sheetTarget, full2dTable) {
  if (!full2dTable || full2dTable.length === 0) return { success: false, message: 'ไม่มีข้อมูล' };
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetTarget);
  if (!sheet) {
    sheet = ss.insertSheet(sheetTarget);
  }

  var maxCols = 0;
  for (var r = 0; r < full2dTable.length; r++) {
    if (full2dTable[r] && full2dTable[r].length > maxCols) maxCols = full2dTable[r].length;
  }
  if (maxCols === 0) return { success: false, message: 'ข้อมูลไม่มีคอลัมน์' };

  var padded = full2dTable.map(function(row) {
    var r = row ? row.slice() : [];
    while (r.length < maxCols) r.push('');
    return r;
  });

  // เคลียร์และเขียนทับทั้งหมดใน 1 คำสั่ง
  sheet.clearContents();
  sheet.getRange(1, 1, padded.length, maxCols).setValues(padded);

  // คำนวณแดชบอร์ดตามค่าย
  if (sheetTarget === '56A0S0Q' || sheetTarget === 'AIS') {
    createAisDashboard();
  } else if (sheetTarget === '56A0UPS' || sheetTarget === 'True') {
    createTrueDashboard();
  }

  return { success: true, count: padded.length, message: 'นำเข้าตาราง ' + sheetTarget + ' ทั้งหมดสำเร็จเรียบร้อยแล้ว' };
}

/**
 * อัปโหลด base64 image ไปยัง Google Drive แล้วคืน URL สาธารณะ
 * LINE Messaging API ต้องการ URL (ไม่รองรับ base64 โดยตรง)
 */
function uploadImageToDrive_(base64Data) {
  var cleanBase64 = base64Data.indexOf(',') > -1 ? base64Data.split(',')[1] : base64Data;
  var blob = Utilities.newBlob(
    Utilities.base64Decode(cleanBase64),
    'image/png',
    'SLA_Dashboard_' + new Date().getTime() + '.png'
  );

  // หา/สร้างโฟลเดอร์ SLA Dashboard Images ใน Drive
  var folderName = 'SLA Dashboard Images';
  var folders = DriveApp.getFoldersByName(folderName);
  var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

  // อัปโหลดไฟล์ + เปิดสิทธิ์ให้ทุกคนดูได้ (จำเป็นสำหรับ LINE)
  var file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

  // URL ที่ LINE สามารถ fetch ได้โดยตรง
  var fileId = file.getId();
  return 'https://drive.google.com/uc?export=view&id=' + fileId;
}

/**
 * ส่งแจ้งเตือนเข้า LINE ผ่าน LINE Notify Token หรือ LINE Messaging API (บอท SPE_SLA)
 */
function sendLineAlert(token, message, imageBase64, groupId) {
  if (!token) throw new Error('กรุณาระบุ LINE Token หรือ Webhook URL');
  
  // บันทึก Token ล่าสุดไว้ใน User Properties
  try {
    PropertiesService.getUserProperties().setProperty('LINE_NOTIFY_TOKEN', token);
    if (groupId) PropertiesService.getScriptProperties().setProperty('SAVED_LINE_GROUP_ID', groupId);
  } catch(e) {}

  if (!groupId) {
    groupId = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || '';
  }

  if (token.indexOf('http://') === 0 || token.indexOf('https://') === 0) {
    // 1. กรณีเป็น Webhook URL
    var payload = {
      message: message,
      imageBase64: imageBase64 || null,
      groupId: groupId
    };
    var options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(token, options);
    return { success: true, response: res.getContentText() };
  } else if (token.length > 80) {
    // 2. กรณีเป็น LINE Messaging API Channel Access Token (บอท SPE_SLA)
    // รองรับหลาย groupId คั่นด้วย comma หรือ newline
    var pushUrl = 'https://api.line.me/v2/bot/message/push';
    var broadcastUrl = 'https://api.line.me/v2/bot/message/broadcast';
    var headers = { 'Authorization': 'Bearer ' + token };

    // ── อัปโหลดรูปไปยัง Google Drive แล้วรับ URL สาธารณะ ──
    var imageUrl = null;
    if (imageBase64) {
      try {
        imageUrl = uploadImageToDrive_(imageBase64);
      } catch(imgUpErr) {
        Logger.log('Image upload skipped: ' + imgUpErr.message);
      }
    }

    // สร้าง messages array (text + image ถ้ามี)
    var msgs = [{ type: 'text', text: message }];
    if (imageUrl) {
      msgs.push({
        type: 'image',
        originalContentUrl: imageUrl,
        previewImageUrl: imageUrl
      });
    }

    // Parse groupIds: แยกด้วย newline หรือ comma แล้ว trim ค่าว่าง
    var groupIds = [];
    if (groupId) {
      groupIds = String(groupId).split(/[\n,]+/)
        .map(function(id) { return id.trim(); })
        .filter(function(id) { return id.length > 0; });
    }

    if (groupIds.length === 0) {
      // Broadcast ถ้าไม่มี Group ID
      var payload = JSON.stringify({ messages: msgs });
      var options = {
        method: 'post',
        contentType: 'application/json',
        headers: headers,
        payload: payload,
        muteHttpExceptions: true
      };
      var res = UrlFetchApp.fetch(broadcastUrl, options);
      var resCode = res.getResponseCode();
      if (resCode !== 200) {
        throw new Error('LINE Broadcast Error (' + resCode + '): ' + res.getContentText());
      }
      return { success: true, response: res.getContentText(), groups: ['broadcast'] };
    }

    // Push ไปทุก Group ID
    var results = [];
    var errors = [];
    for (var g = 0; g < groupIds.length; g++) {
      var gid = groupIds[g];
      try {
        var pushPayload = JSON.stringify({ to: gid, messages: msgs });
        var pushOptions = {
          method: 'post',
          contentType: 'application/json',
          headers: headers,
          payload: pushPayload,
          muteHttpExceptions: true
        };
        var pushRes = UrlFetchApp.fetch(pushUrl, pushOptions);
        var pushCode = pushRes.getResponseCode();
        if (pushCode !== 200) {
          errors.push(gid + ': ' + pushCode + ' ' + pushRes.getContentText());
        } else {
          results.push(gid);
          // บันทึก Group ID แรกที่สำเร็จ
          if (g === 0) PropertiesService.getScriptProperties().setProperty('SAVED_LINE_GROUP_ID', gid);
        }
      } catch(pushErr) {
        errors.push(gid + ': ' + pushErr.message);
      }
    }

    if (errors.length > 0 && results.length === 0) {
      throw new Error('ส่งทุกกลุ่มล้มเหลว: ' + errors.join(' | '));
    }
    return {
      success: true,
      sent: results.length,
      failed: errors.length,
      groups: results,
      errors: errors,
      imageUrl: imageUrl || null
    };

  } else {
    // 3. กรณีเป็น LINE Notify Token (ปกติความยาว 43 ตัวอักษร)
    var url = 'https://notify-api.line.me/api/notify';
    var payload = {
      'message': '\n' + message
    };

    if (imageBase64) {
      try {
        var cleanBase64 = imageBase64.indexOf(',') > -1 ? imageBase64.split(',')[1] : imageBase64;
        var imageBlob = Utilities.newBlob(Utilities.base64Decode(cleanBase64), 'image/png', 'dashboard.png');
        payload['imageFile'] = imageBlob;
      } catch(imgErr) {
        Logger.log('Attach image error: ' + imgErr.message);
      }
    }

    var options = {
      method: 'post',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      payload: payload,
      muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    var resCode = res.getResponseCode();
    if (resCode !== 200) {
      throw new Error('LINE API Error (' + resCode + '): ' + res.getContentText());
    }
    return { success: true, response: res.getContentText() };
  }
}

/**
 * เมนูลัด: ส่งแจ้งเตือน LINE สำหรับ AIS จากเมนูชีต
 */
function sendLineAlertAIS() {
  var ui = SpreadsheetApp.getUi();
  var token = PropertiesService.getUserProperties().getProperty('LINE_NOTIFY_TOKEN') || DEFAULT_LINE_TOKEN;
  if (!token) {
    var prompt = ui.prompt('📲 ส่งแจ้งเตือน LINE Group (AIS)', 'กรุณากรอก LINE Channel Access Token:', ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    token = prompt.getResponseText().trim();
  }
  if (!token) { ui.alert('⚠️ ไม่พบ Token'); return; }

  var groupId = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || DEFAULT_LINE_GROUP_ID;

  // สร้างข้อความสรุป AIS
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('Dashboard สรุปงาน');
  if (!dash) { ui.alert('⚠️ ไม่พบชีต Dashboard สรุปงาน กรุณากด Refresh ก่อน'); return; }

  var totalVal = dash.getRange('I1').getValue();
  var dateVal = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  
  var msg = '📊 [AIS INSTALLATION & SLA REPORT]\n' +
            '📅 ณ วันที่: ' + dateVal + '\n' +
            '🎯 งานติดตั้งเสร็จ: ' + totalVal + '\n' +
            '------------------------------------\n' +
            '🔹 SMART QC:\n' +
            '🟢 ปิดตามกำหนด: ' + dash.getRange('B5').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('C5').getValue() + ' วัน)\n' +
            '🟡 ปิดงานล่าช้า: ' + dash.getRange('B6').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('C6').getValue() + ' วัน)\n' +
            '🔵 รอตรวจสอบ: ' + dash.getRange('B7').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('C7').getValue() + ' วัน)\n' +
            '🔴 ค้างวิกฤต: ' + dash.getRange('B8').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('C8').getValue() + ' วัน)\n\n' +
            '🔹 PAT SUBCON:\n' +
            '🟢 ส่งตามกำหนด: ' + dash.getRange('F5').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('G5').getValue() + ' วัน)\n' +
            '🟡 ส่งงานล่าช้า: ' + dash.getRange('F6').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('G6').getValue() + ' วัน)\n' +
            '🔵 รอส่งงาน: ' + dash.getRange('F7').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('G7').getValue() + ' วัน)\n' +
            '🔴 ค้างวิกฤต: ' + dash.getRange('F8').getValue() + ' ไซต์\n' +
            '------------------------------------\n' +
            '📱 รายละเอียดเพิ่มเติมดูได้ในแดชบอร์ด';

  try {
    sendLineAlert(token, msg, null, groupId);
    ui.alert('✅ ส่งสรุปแจ้งเตือนเข้า LINE Group สำเร็จแล้ว!');
  } catch(e) {
    ui.alert('⚠️ เกิดข้อผิดพลาดในการส่ง LINE:\n' + e.message);
  }
}

/**
 * เมนูลัด: ส่งแจ้งเตือน LINE สำหรับ TRUE จากเมนูชีต
 */
function sendLineAlertTRUE() {
  var ui = SpreadsheetApp.getUi();
  var token = PropertiesService.getUserProperties().getProperty('LINE_NOTIFY_TOKEN') || DEFAULT_LINE_TOKEN;
  if (!token) {
    var prompt = ui.prompt('📲 ส่งแจ้งเตือน LINE Group (TRUE)', 'กรุณากรอก LINE Channel Access Token:', ui.ButtonSet.OK_CANCEL);
    if (prompt.getSelectedButton() !== ui.Button.OK) return;
    token = prompt.getResponseText().trim();
  }
  if (!token) { ui.alert('⚠️ ไม่พบ Token'); return; }

  var groupId = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || DEFAULT_LINE_GROUP_ID;

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var dash = ss.getSheetByName('Dashboard สรุปงาน TRUE');
  if (!dash) { ui.alert('⚠️ ไม่พบชีต Dashboard สรุปงาน TRUE กรุณากด Refresh ก่อน'); return; }

  var totalVal = dash.getRange('I1').getValue();
  var dateVal = Utilities.formatDate(new Date(), 'GMT+7', 'dd/MM/yyyy HH:mm');
  
  var msg = '📊 [TRUE INSTALLATION & SLA REPORT]\n' +
            '📅 ณ วันที่: ' + dateVal + '\n' +
            '🎯 งานติดตั้งเสร็จ: ' + totalVal + '\n' +
            '------------------------------------\n' +
            '🔹 08.1 SmartQC+AOR:\n' +
            '🟢 ปิดตามกำหนด: ' + dash.getRange('B5').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('C5').getValue() + ' วัน)\n' +
            '🟡 ปิดงานล่าช้า: ' + dash.getRange('B6').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('C6').getValue() + ' วัน)\n' +
            '🔵 รอตรวจสอบ: ' + dash.getRange('B7').getValue() + ' ไซต์\n' +
            '🔴 ค้างวิกฤต: ' + dash.getRange('B8').getValue() + ' ไซต์\n\n' +
            '🔹 14.1 A129 PAT Site Folder:\n' +
            '🟢 ส่งตามกำหนด: ' + dash.getRange('F5').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('G5').getValue() + ' วัน)\n' +
            '🟡 ส่งงานล่าช้า: ' + dash.getRange('F6').getValue() + ' ไซต์ (เฉลี่ย ' + dash.getRange('G6').getValue() + ' วัน)\n' +
            '🔵 รอส่งงาน: ' + dash.getRange('F7').getValue() + ' ไซต์\n' +
            '🔴 ค้างวิกฤต: ' + dash.getRange('F8').getValue() + ' ไซต์\n' +
            '------------------------------------\n' +
            '📱 รายละเอียดเพิ่มเติมดูได้ในแดชบอร์ด';

  try {
    sendLineAlert(token, msg, null, groupId);
    ui.alert('✅ ส่งสรุปแจ้งเตือนเข้า LINE Group สำเร็จแล้ว!');
  } catch(e) {
    ui.alert('⚠️ เกิดข้อผิดพลาดในการส่ง LINE:\n' + e.message);
  }
}