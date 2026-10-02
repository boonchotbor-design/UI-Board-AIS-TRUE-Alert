// ===========================================================
// เวอร์ชันสมบูรณ์ รวม AIS + TRUE
// แก้ไข: TRUE ใช้ header row 3 (index 2) และ data เริ่มที่ row 4 (index 3)
//         AOR=col X (index 23), SmartQC=col Y (index 24) จากภาพจริง
// ===========================================================

// ── LINE Messaging API: Channel Access Token ของบอท SPE_SLA ──
var DEFAULT_LINE_TOKEN = 'YKtVKOIprzQoLKqB7foUkyxIwvzGaWxY/lnBmm4GaoJVNVDgbEUOTs8MOZRWBtEfzX8X6k0pX+pJSyave60Ka//baM6waKsQE/Ho43TkMod6YcyLcreDpjVC85MCXv7NxSj47Bh6bI2a2Xuls5hnkAdB04t89/1O/w1cDnyilFU=';

// ─────────────────────────────────────────────────────────────
// doGet: ดึง Group ID ที่บันทึกไว้ (เรียกผ่าน Web App URL ?action=groupid)
// ─────────────────────────────────────────────────────────────
function doGet(e) {
  var action = e && e.parameter && e.parameter.action ? e.parameter.action : '';
  if (action === 'groupid') {
    var gid = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || 'ยังไม่พบ Group ID';
    return ContentService.createTextOutput('LINE_GROUP_ID=' + gid)
      .setMimeType(ContentService.MimeType.TEXT);
  }
  // Default: open dashboard
  try {
    var html = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('AIS & TRUE SLA Dashboard')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    return html;
  } catch(e2) {
    return ContentService.createTextOutput('Dashboard OK - ' + new Date().toISOString())
      .setMimeType(ContentService.MimeType.TEXT);
  }
}

// ─────────────────────────────────────────────────────────────
// doPost: LINE Webhook Receiver
//   - รับ events จาก LINE เมื่อมีข้อความในกลุ่ม
//   - บันทึก groupId อัตโนมัติ
//   - ถ้า source.type == 'group' และข้อความมี /groupid → reply ด้วย Group ID
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
            var text = evt.message.text || '';
            if (text.toLowerCase().indexOf('/groupid') >= 0 || text.toLowerCase().indexOf('groupid') >= 0) {
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
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return null;
  var formatted = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  var parts = formatted.split("-");
  return new Date(parseInt(parts[0],10), parseInt(parts[1],10)-1, parseInt(parts[2],10));
}

function colIdxToLetter_(idx) {
  var s="", i=idx+1;
  while(i>0){var r=(i-1)%26; s=String.fromCharCode(65+r)+s; i=Math.floor((i-1)/26);}
  return s;
}

// ---------------------------------------------------------
// ✅ หา TRUE column indices โดย scan ทุก cell ใน 10 แถวแรก
//    คืน object ที่มี headerRowIdx และ index แต่ละ column
//    พร้อม fallback จากภาพจริง (ยืนยันจาก screenshot)
// ---------------------------------------------------------
function getTrueColumnIndices_(data) {
  var SCAN_ROWS = Math.min(10, data.length);

  // สร้าง map: normalized_key -> col_index
  // scan ทุก row เพื่อรองรับ multi-row header
  var normMap = {}; // key = lower+no-space-dot, value = col index
  var rawMap  = {}; // key = lower trimmed,       value = col index
  var headerRowIdx = 2; // default: row 3 (0-based=2) จากภาพจริง

  for (var r = 0; r < SCAN_ROWS; r++) {
    for (var c = 0; c < data[r].length; c++) {
      var raw = data[r][c] ? String(data[r][c]).trim() : "";
      if (raw === "") continue;
      var lo   = raw.toLowerCase();
      var norm = lo.replace(/[\s.\-_\/]/g, "");

      if (!rawMap[lo])   rawMap[lo]   = c;
      if (!normMap[norm]) normMap[norm] = c;

      // หา header row จาก "DU ID"
      if (lo === "du id" || norm === "duid") headerRowIdx = r;
    }
  }

  // helper: ลอง keys หลายรูปแบบ คืน index แรกที่เจอ หรือ fallback
  function pick(candidates, fallback) {
    for (var k = 0; k < candidates.length; k++) {
      var lo   = candidates[k].toLowerCase().trim();
      var norm = lo.replace(/[\s.\-_\/]/g, "");
      if (rawMap[lo]   !== undefined) return rawMap[lo];
      if (normMap[norm] !== undefined) return normMap[norm];
    }
    return fallback;
  }

  // ✅ indices ยืนยันจากภาพ sheet จริง (fallback = col จากภาพ)
  var result = {
    headerRowIdx : headerRowIdx,
    idxDuid      : pick(["du id","duid"],                             1),   // col B
    idxVerifyPhoto: pick(["verify photo","verifyphoto"],              22),  // col W
    idxAor       : pick(["aor"],                                      23),  // col X ✅
    idxSmartQc   : pick(["08.1 smartqc","081smartqc","smartqc",
                          "08.1smartqc","081 smartqc"],               24),  // col Y ✅
    idxExtAlarm  : pick(["ext alarm test","extalarmtest"],            25),  // col Z
    idxPatRemark : pick(["integration alarm category",
                          "alarm category","integrationalarmc"],      26),  // col AA
    idxCrossSector: pick(["cross sector verification"],               27),  // col AB
    idxPat       : pick(["14.1 a129 pat site folder",
                          "pat site folder","patsitef"],              28),  // col AC
    idxOwnerDoc  : pick(["pat owner","patowner","owner doc",
                          "ownerdoc","return x-doc","returnxdoc"],    32),  // col AG
  };

  Logger.log("=== getTrueColumnIndices_ ===");
  for (var key in result) {
    var idx = result[key];
    if (key === "headerRowIdx") { Logger.log("  headerRowIdx: " + idx + " (sheet row " + (idx+1) + ")"); continue; }
    Logger.log("  "+key+": "+idx+" ("+colIdxToLetter_(idx)+") header=["+_getHeaderLabel(data,SCAN_ROWS,idx)+"]");
  }
  return result;
}

function _getHeaderLabel(data, scanRows, colIdx) {
  var parts = [];
  for (var r = 0; r < scanRows; r++) {
    var v = data[r][colIdx] ? String(data[r][colIdx]).trim() : "";
    if (v !== "") parts.push(v);
  }
  return parts.join(" / ");
}

// ---------------------------------------------------------
// 3. AIS Dashboard
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

    var headers      = data[0];
    var idxInstall   = 30, idxSmartQc=31, idxPat=33, idxPatRemark=35, idxOwnerDoc=37;
    var idxDuid = 0;
    for (var c=0;c<headers.length;c++) {
      if ((headers[c]?String(headers[c]).trim().toLowerCase():"").includes("duid")) {idxDuid=c;break;}
    }

    var today = toDateOnly_(new Date(), ssTz);
    var availableYears={};
    for (var i=1;i<data.length;i++) {
      var d=data[i][idxInstall];
      if(d&&d!==""){var dd=d instanceof Date?d:new Date(d);if(!isNaN(dd.getTime()))availableYears[dd.getFullYear()]=true;}
    }
    var yearList=Object.keys(availableYears).sort(function(a,b){return b-a;});yearList.unshift("ทั้งหมด");

    var acc = _newAccumulators_();
    var C = _colors_();

    for (var i=1;i<data.length;i++) {
      var row=data[i];
      if(row.length<=idxPatRemark) continue;
      var installDate=toDateOnly_(row[idxInstall],ssTz); if(!installDate) continue;
      if(selectedYear!=="ทั้งหมด"&&installDate.getFullYear().toString()!==selectedYear) continue;
      acc.totalInstall++;
      var duidVal   =row[idxDuid]     ?String(row[idxDuid]).trim()    :"ไม่ระบุ DUID";
      var ownerDoc  =row[idxOwnerDoc] ?String(row[idxOwnerDoc]).trim():"-";
      var smartQcDate=toDateOnly_(row[idxSmartQc],ssTz);
      var patDate    =toDateOnly_(row[idxPat],ssTz);
      var patRemark  =row[idxPatRemark]?String(row[idxPatRemark]).trim():"";

      _calcSQC_(acc, today, installDate, smartQcDate, duidVal, ownerDoc, 3, "");
      _calcPAT_(acc, today, installDate, smartQcDate, patDate, duidVal, ownerDoc);
      if (patRemark.toLowerCase().includes("not pass")) {
        var od=Math.max(0,patDate?Math.floor((today-patDate)/86400000):Math.floor((today-installDate)/86400000));
        acc.listPatNotPass.push([duidVal,od+" วัน",C.ORANGE_BG,C.ORANGE_FG,ownerDoc]);
      }
    }

    function getAvg(t,c){return c>0?(t/c).toFixed(1):"0.0";}
    _buildDashboard(dashSheet,{
      title:"AIS INSTALLATION & SLA DASHBOARD", selectedYear, yearList, totalInstall:acc.totalInstall,
      ...acc, sqcCardLabel:"EXECUTIVE SUMMARY: SMART QC", patCardLabel:"EXECUTIVE SUMMARY: PAT SUBCON SUBMIT",
      pendingSqcTitle:"ACTION REQUIRED: SMART QC", pendingPatTitle:"ACTION REQUIRED: PAT SUBCON",
      notPassTitle:"REWORK REQUIRED: PAT NOT PASS", debugLine2:null, getAvg
    });
  } catch(err){Logger.log(err.message+"\n"+err.stack);SpreadsheetApp.getUi().alert("⚠️ Error:\n"+err.message);}
  finally{lock.releaseLock();}
}

// ---------------------------------------------------------
// 4. TRUE Dashboard
// ✅ ใช้ getTrueColumnIndices_ scan header จริง
//    AOR + SmartQC ต้องมีทั้งคู่ถึงจะ Done
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

    // ✅ หา column indices จาก header จริง
    var idx = getTrueColumnIndices_(data);
    var headerRowIdx  = idx.headerRowIdx;
    var idxDuid       = idx.idxDuid;
    var idxVerifyPhoto= idx.idxVerifyPhoto;
    var idxAor        = idx.idxAor;
    var idxSmartQc    = idx.idxSmartQc;
    var idxPat        = idx.idxPat;
    var idxPatRemark  = idx.idxPatRemark;
    var idxOwnerDoc   = idx.idxOwnerDoc;

    var today = toDateOnly_(new Date(), ssTz);
    var availableYears={};
    // ✅ data เริ่มหลัง headerRowIdx
    for (var i=headerRowIdx+1;i<data.length;i++) {
      var d=data[i][idxVerifyPhoto];
      if(d&&d!==""){var dd=d instanceof Date?d:new Date(d);if(!isNaN(dd.getTime()))availableYears[dd.getFullYear()]=true;}
    }
    var yearList=Object.keys(availableYears).sort(function(a,b){return b-a;});yearList.unshift("ทั้งหมด");

    var acc = _newAccumulators_();
    var C = _colors_();

    for (var i=headerRowIdx+1;i<data.length;i++) {
      var row=data[i];
      var installDate=toDateOnly_(row[idxVerifyPhoto],ssTz); if(!installDate) continue;
      if(selectedYear!=="ทั้งหมด"&&installDate.getFullYear().toString()!==selectedYear) continue;
      acc.totalInstall++;

      var duidVal   =row[idxDuid]     ?String(row[idxDuid]).trim()    :"ไม่ระบุ DUID";
      var ownerDoc  =row[idxOwnerDoc] ?String(row[idxOwnerDoc]).trim():"-";
      var patRemark =row[idxPatRemark]?String(row[idxPatRemark]).trim():"";

      // ✅ แปลงวันที่ทั้ง 2 milestone
      var aorDate     = toDateOnly_(row[idxAor],      ssTz);
      var smartQcDate = toDateOnly_(row[idxSmartQc],  ssTz);
      var patDate     = toDateOnly_(row[idxPat],      ssTz);

      // ✅ Debug log เฉพาะ RYG0757_Flash_battery
      if (duidVal.toLowerCase().includes("ryg0757_flash_battery")) {
        Logger.log("=== DEBUG RYG0757_Flash_battery (sheet row "+(i+1)+") ===");
        Logger.log("  VerifyPhoto["+idxVerifyPhoto+"]("+colIdxToLetter_(idxVerifyPhoto)+"): "+installDate+" raw=["+row[idxVerifyPhoto]+"]");
        Logger.log("  AOR["+idxAor+"]("+colIdxToLetter_(idxAor)+"): "+aorDate+" raw=["+row[idxAor]+"] type="+typeof row[idxAor]);
        Logger.log("  SmartQC["+idxSmartQc+"]("+colIdxToLetter_(idxSmartQc)+"): "+smartQcDate+" raw=["+row[idxSmartQc]+"] type="+typeof row[idxSmartQc]);
        Logger.log("  sqcDone: "+(aorDate&&smartQcDate?"YES ✅":"NO ❌ missing="+((!aorDate?"AOR ":"")+(!smartQcDate?"SmartQC":""))));
        // dump raw values col 20-30
        for(var cc=20;cc<=30;cc++) Logger.log("    raw col["+cc+"]("+colIdxToLetter_(cc)+"): ["+row[cc]+"] type="+typeof row[cc]);
      }

      // ✅ Smart QC: Done = ต้องมีทั้ง AOR และ SmartQC
      var sqcDone = aorDate && smartQcDate;
      var missingLabel = (!aorDate?" ⚠️AOR":"")+(!smartQcDate?" ⚠️SQC":"");
      _calcSQC_(acc, today, installDate, sqcDone?smartQcDate:null, duidVal, ownerDoc, 3, missingLabel);

      // PAT
      _calcPAT_(acc, today, installDate, smartQcDate, patDate, duidVal, ownerDoc);

      // Alarm Rework: มีค่าและไม่ใช่ "no alarm"
      var hasAlarm = patRemark!=="" && !patRemark.toLowerCase().includes("no alarm");
      if (hasAlarm) {
        var od=Math.max(0,patDate?Math.floor((today-patDate)/86400000):Math.floor((today-installDate)/86400000));
        acc.listPatNotPass.push([duidVal,od+" วัน",C.ORANGE_BG,C.ORANGE_FG,ownerDoc]);
      }
    }

    function getAvg(t,c){return c>0?(t/c).toFixed(1):"0.0";}
    var debugLine2=" | VP="+colIdxToLetter_(idxVerifyPhoto)+" AOR="+colIdxToLetter_(idxAor)+
                   " SQC="+colIdxToLetter_(idxSmartQc)+" PAT="+colIdxToLetter_(idxPat)+
                   " Alarm="+colIdxToLetter_(idxPatRemark)+" Owner="+colIdxToLetter_(idxOwnerDoc)+
                   " (hdr row "+(headerRowIdx+1)+")";

    _buildDashboard(dashSheet,{
      title:"TRUE INSTALLATION & SLA DASHBOARD", selectedYear, yearList, totalInstall:acc.totalInstall,
      ...acc, sqcCardLabel:"EXECUTIVE SUMMARY: 08.1 SmartQC+AOR", patCardLabel:"EXECUTIVE SUMMARY: 14.1 A129 PAT Site Folder",
      pendingSqcTitle:"ACTION REQUIRED: 08.1 SmartQC+AOR", pendingPatTitle:"ACTION REQUIRED: 14.1 A129 PAT Site Folder",
      notPassTitle:"REWORK REQUIRED: ALARM FOUND", debugLine2, getAvg
    });
  } catch(err){Logger.log(err.message+"\n"+err.stack);SpreadsheetApp.getUi().alert("⚠️ Error:\n"+err.message);}
  finally{lock.releaseLock();}
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
//    sqcDateDone = null → Pending, มีค่า → Done
//    missingLabel = ข้อความเพิ่มเติม เช่น " ⚠️AOR"
function _calcSQC_(acc, today, installDate, sqcDateDone, duidVal, ownerDoc, slaDays, missingLabel) {
  var C = _colors_();
  if (!sqcDateDone) {
    var diff = Math.max(0, Math.floor((today - installDate) / 86400000));
    if (diff > slaDays) {
      acc.pendingSQC_OverSLA++; acc.totalDays_pendingSQC_OverSLA+=diff;
      acc.listPendingSmartQc.push([duidVal, diff+" วัน"+(missingLabel||""), C.RED_BG, C.RED_FG, ownerDoc]);
    } else {
      acc.pendingSQC_InSLA++; acc.totalDays_pendingSQC_InSLA+=diff;
      acc.listPendingSmartQc.push([duidVal, diff+" วัน"+(missingLabel||""), C.BLUE_BG, C.BLUE_FG, ownerDoc]);
    }
  } else {
    var diff = Math.max(0, Math.floor((sqcDateDone - installDate) / 86400000));
    if (diff > slaDays) { acc.doneSQC_OverSLA++; acc.totalDays_doneSQC_OverSLA+=diff; }
    else                { acc.doneSQC_InSLA++;   acc.totalDays_doneSQC_InSLA+=diff; }
  }
}

// ✅ คำนวณ PAT
function _calcPAT_(acc, today, installDate, smartQcDate, patDate, duidVal, ownerDoc) {
  var C = _colors_();
  if (!patDate) {
    if (smartQcDate) {
      var diff = Math.max(0, Math.floor((today - smartQcDate) / 86400000));
      if (diff > 5) { acc.pendingPat_OverSLA++; acc.totalDays_pendingPat_OverSLA+=diff; acc.listPendingPat.push([duidVal,diff+" วัน",C.RED_BG,C.RED_FG,ownerDoc]); }
      else          { acc.pendingPat_InSLA++;   acc.totalDays_pendingPat_InSLA+=diff;   acc.listPendingPat.push([duidVal,diff+" วัน",C.BLUE_BG,C.BLUE_FG,ownerDoc]); }
    }
  } else {
    var base = smartQcDate || installDate;
    var diff = Math.max(0, Math.floor((patDate - base) / 86400000));
    if (diff > 5) { acc.donePat_OverSLA++; acc.totalDays_donePat_OverSLA+=diff; }
    else          { acc.donePat_InSLA++;   acc.totalDays_donePat_InSLA+=diff; }
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
 * เสิร์ฟ Web App สำหรับเปิดบน Browser อิสระ
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
      .evaluate()
      .setTitle('AIS & TRUE INSTALLATION & SLA DASHBOARD')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * เปิด Modal Dialog ขนาดใหญ่ใน Google Sheets
 */
function showDashboardModal() {
  var html = HtmlService.createTemplateFromFile('Index')
      .evaluate()
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
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) return { success: false, message: 'ไม่พบชีต: ' + sheetName };

  var data = sheet.getDataRange().getValues();
  var duidRowMap = {};
  
  // สร้าง map DUID ที่มีอยู่เดิมในชีต
  if (operator === 'AIS') {
    // Row 2 คือ Header (index 1), Data เริ่ม row index 2
    for (var r = 2; r < data.length; r++) {
      var d = data[r][0] ? String(data[r][0]).trim() : '';
      if (d) duidRowMap[d] = r + 1; // 1-based row index
    }

    // ทำการ update หรือ append
    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var duid = rec.duid;
      if (!duid) continue;

      var targetRow = duidRowMap[duid];
      if (targetRow) {
        // อัปเดตข้อมูลวันที่ และสถานะ
        if (rec.installDate) sheet.getRange(targetRow, 31).setValue(rec.installDate); // col AE (31)
        if (rec.smartQcDate) sheet.getRange(targetRow, 32).setValue(rec.smartQcDate); // col AF (32)
        if (rec.patDate) sheet.getRange(targetRow, 34).setValue(rec.patDate); // col AH (34)
        if (rec.patRemark) sheet.getRange(targetRow, 36).setValue(rec.patRemark); // col AJ (36)
        if (rec.patStatus) sheet.getRange(targetRow, 37).setValue(rec.patStatus); // col AK (37)
        if (rec.owner && rec.owner !== '-') sheet.getRange(targetRow, 38).setValue(rec.owner); // col AL (38)
      } else {
        // Append แถวใหม่
        var newRow = new Array(45).fill('');
        newRow[0] = duid;
        newRow[30] = rec.installDate || '';
        newRow[31] = rec.smartQcDate || '';
        newRow[33] = rec.patDate || '';
        newRow[35] = rec.patRemark || '';
        newRow[36] = rec.patStatus || '';
        newRow[37] = rec.owner || '-';
        sheet.appendRow(newRow);
        duidRowMap[duid] = sheet.getLastRow();
      }
    }

    // สั่งคำนวณและวาด Dashboard AIS ใหม่ทันที
    createAisDashboard();

  } else {
    // TRUE
    // Data เริ่มต้นที่ row 4 (index 3), col B คือ DUID
    for (var r = 3; r < data.length; r++) {
      var d = data[r][1] ? String(data[r][1]).trim() : '';
      if (d) duidRowMap[d] = r + 1;
    }

    for (var i = 0; i < records.length; i++) {
      var rec = records[i];
      var duid = rec.duid;
      if (!duid) continue;

      var targetRow = duidRowMap[duid];
      if (targetRow) {
        if (rec.installDate) sheet.getRange(targetRow, 23).setValue(rec.installDate); // col W (23)
        if (rec.aorDate) sheet.getRange(targetRow, 24).setValue(rec.aorDate); // col X (24)
        if (rec.smartQcDate) sheet.getRange(targetRow, 25).setValue(rec.smartQcDate); // col Y (25)
        if (rec.alarmRemark) sheet.getRange(targetRow, 27).setValue(rec.alarmRemark); // col AA (27)
        if (rec.patDate) sheet.getRange(targetRow, 29).setValue(rec.patDate); // col AC (29)
        if (rec.owner && rec.owner !== '-') sheet.getRange(targetRow, 33).setValue(rec.owner); // col AG (33)
      } else {
        var newRow = new Array(35).fill('');
        newRow[1] = duid;
        newRow[22] = rec.installDate || '';
        newRow[23] = rec.aorDate || '';
        newRow[24] = rec.smartQcDate || '';
        newRow[26] = rec.alarmRemark || '';
        newRow[28] = rec.patDate || '';
        newRow[32] = rec.owner || '-';
        sheet.appendRow(newRow);
        duidRowMap[duid] = sheet.getLastRow();
      }
    }

    // สั่งคำนวณและวาด Dashboard TRUE ใหม่ทันที
    createTrueDashboard();
  }

  return { success: true, count: records.length, message: 'นำเข้าข้อมูล ' + operator + ' สำเร็จแล้ว' };
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
    // ถ้ารู้ groupId ให้ส่งแบบ push ตรงเข้ากลุ่ม, ถ้าไม่รู้ให้ broadcast
    var url = groupId ? 'https://api.line.me/v2/bot/message/push' : 'https://api.line.me/v2/bot/message/broadcast';
    var payload = {
      messages: [
        {
          type: 'text',
          text: message
        }
      ]
    };
    if (groupId) payload.to = groupId;

    var options = {
      method: 'post',
      contentType: 'application/json',
      headers: {
        'Authorization': 'Bearer ' + token
      },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    var res = UrlFetchApp.fetch(url, options);
    var resCode = res.getResponseCode();
    if (resCode !== 200) {
      throw new Error('LINE Messaging API Error (' + resCode + '): ' + res.getContentText());
    }
    return { success: true, response: res.getContentText() };
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

  var groupId = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || '';

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

  var groupId = PropertiesService.getScriptProperties().getProperty('SAVED_LINE_GROUP_ID') || '';

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