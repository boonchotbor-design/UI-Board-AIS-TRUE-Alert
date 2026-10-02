// ===========================================================
// ไฟล์นี้คือเวอร์ชันสมบูรณ์ รวม AIS เดิม + TRUE ใหม่ ไว้ในไฟล์เดียว
// วิธีใช้: เปิด Apps Script -> เลือกทั้งหมดในไฟล์ Code.gs (Ctrl+A) -> ลบทิ้ง
// -> วางไฟล์นี้ทับทั้งหมด -> Save -> รีเฟรชหน้าสเปรดชีต
// ===========================================================

// ---------------------------------------------------------
// 1. เมนูบนหน้า Google Sheets
// ---------------------------------------------------------
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Executive Dashboard')
      .addItem('🔄 อัปเดตข้อมูลล่าสุด AIS (Refresh)', 'createAisDashboard')
      .addItem('🔄 อัปเดตข้อมูลล่าสุด TRUE (Refresh)', 'createTrueDashboard')
      .addToUi();
}

// ---------------------------------------------------------
// 2. Auto-Refresh เมื่อเลือกปี (Magic Trigger)
// ---------------------------------------------------------
function onEdit(e) {
  if (!e) return;
  var sheet = e.source.getActiveSheet();
  var sheetName = sheet.getName();
  var cell = e.range.getA1Notation();

  if (sheetName === "Dashboard สรุปงาน" && cell === "E2") {
    createAisDashboard();
  } else if (sheetName === "Dashboard สรุปงาน TRUE" && cell === "E2") {
    createTrueDashboard();
  }
}

// ---------------------------------------------------------
// helper: แปลงค่าวันที่ให้เหลือแค่ปี/เดือน/วัน (ตัดเวลาออก) โดยอิงตาม
// timezone ของสเปรดชีตเอง แทนที่จะพึ่ง timezone ของ Apps Script project
// วิธีนี้ป้องกันปัญหา Aging ติดลบ (-1 วัน) ที่เกิดจาก timezone ไม่ตรงกัน
// ---------------------------------------------------------
function toDateOnly_(value, tz) {
  if (!value) return null;
  var d = value instanceof Date ? value : new Date(value);
  if (isNaN(d.getTime())) return null;
  var formatted = Utilities.formatDate(d, tz, "yyyy-MM-dd");
  var parts = formatted.split("-");
  return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
}

// ---------------------------------------------------------
// 3. ฟังก์ชันหลักสำหรับ AIS
// ---------------------------------------------------------
function createAisDashboard() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    SpreadsheetApp.getUi().alert("⏳ กำลังอัปเดตข้อมูลอยู่ (มีคนอื่น/รอบอื่นกำลังรันอยู่) กรุณารอสักครู่แล้วลองใหม่อีกครั้ง");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = "Asia/Bangkok"; // ล็อค timezone ตรงๆ กันปัญหา getSpreadsheetTimeZone() ไม่ทำงานใน simple trigger
    var sourceSheetName = "AIS";
    var sourceSheet = ss.getSheetByName(sourceSheetName);

    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert("⚠️ ไม่พบชีตที่ชื่อ '" + sourceSheetName + "' กรุณาตรวจสอบอีกครั้ง");
      return;
    }

    var dashSheetName = "Dashboard สรุปงาน";
    var dashSheet = ss.getSheetByName(dashSheetName);

    var selectedYear = "ทั้งหมด";
    if (dashSheet) {
      var oldFilter = dashSheet.getRange("E2").getValue();
      if (oldFilter && oldFilter !== "") {
        selectedYear = oldFilter.toString().trim();
      }
      dashSheet.clear();
    } else {
      dashSheet = ss.insertSheet(dashSheetName);
    }

    var existingCharts = dashSheet.getCharts();
    for (var i = 0; i < existingCharts.length; i++) {
      dashSheet.removeChart(existingCharts[i]);
    }

    dashSheet.setHiddenGridlines(true);

    var data = sourceSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return;

    var headers = data[0];

    var idxInstall   = 30; // AE
    var idxSmartQc   = 31; // AF
    var idxPat       = 33; // AH
    var idxPatRemark = 35; // AJ
    var idxOwnerDoc  = 37; // AL

    var idxDuid = 0;
    for (var c = 0; c < headers.length; c++) {
      var headerName = headers[c] ? String(headers[c]).trim().toLowerCase() : "";
      if (headerName.includes("duid")) { idxDuid = c; break; }
    }

    var today = toDateOnly_(new Date(), ssTz);
    

    var availableYears = {};
    for (var i = 1; i < data.length; i++) {
      var instDate = data[i][idxInstall];
      if (instDate && instDate !== "") {
        var d = instDate instanceof Date ? instDate : new Date(instDate);
        if (!isNaN(d.getTime())) {
          availableYears[d.getFullYear()] = true;
        }
      }
    }
    var yearList = Object.keys(availableYears).sort(function (a, b) { return b - a; });
    yearList.unshift("ทั้งหมด");

    var totalInstall = 0;

    var doneSQC_InSLA = 0;      var totalDays_doneSQC_InSLA = 0;
    var doneSQC_OverSLA = 0;    var totalDays_doneSQC_OverSLA = 0;
    var pendingSQC_InSLA = 0;   var totalDays_pendingSQC_InSLA = 0;
    var pendingSQC_OverSLA = 0; var totalDays_pendingSQC_OverSLA = 0;

    var donePat_InSLA = 0;      var totalDays_donePat_InSLA = 0;
    var donePat_OverSLA = 0;    var totalDays_donePat_OverSLA = 0;
    var pendingPat_InSLA = 0;   var totalDays_pendingPat_InSLA = 0;
    var pendingPat_OverSLA = 0; var totalDays_pendingPat_OverSLA = 0;

    var listPendingSmartQc = [];
    var listPendingPat     = [];
    var listPatNotPass     = [];

    var C_BG       = "#f8fafc";
    var C_CARD     = "#ffffff";
    var C_HEAD_BG  = "#0f172a";
    var C_HEAD_TXT = "#ffffff";
    var C_BORDER   = "#cbd5e1";

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (row.length <= idxPatRemark) continue;

      var installDateVal = row[idxInstall];
      var smartQcDateVal = row[idxSmartQc];
      var patDateVal     = row[idxPat];
      var patRemarkVal   = row[idxPatRemark] ? String(row[idxPatRemark]).trim() : "";
      var duidVal        = row[idxDuid]      ? String(row[idxDuid]).trim()      : "ไม่ระบุ DUID";
      var ownerDocVal    = row[idxOwnerDoc]  ? String(row[idxOwnerDoc]).trim()  : "-";

      if (!installDateVal || installDateVal === "") continue;
      var installDate = toDateOnly_(installDateVal, ssTz);
      if (!installDate) continue;

      var dataYear = installDate.getFullYear().toString();
      if (selectedYear !== "ทั้งหมด" && dataYear !== selectedYear) continue;

      totalInstall++;

      var smartQcDate = null;
      if (smartQcDateVal && smartQcDateVal !== "") {
        smartQcDate = toDateOnly_(smartQcDateVal, ssTz);
      }
      var patDate = null;
      if (patDateVal && patDateVal !== "") {
        patDate = toDateOnly_(patDateVal, ssTz);
      }

      if (!smartQcDate) {
        var diffDays = Math.floor((today - installDate) / (1000 * 60 * 60 * 24));
        if (diffDays > 3) {
          pendingSQC_OverSLA++; totalDays_pendingSQC_OverSLA += diffDays;
          listPendingSmartQc.push([duidVal, diffDays + " วัน", "#fee2e2", "#dc2626", ownerDocVal]);
        } else {
          pendingSQC_InSLA++; totalDays_pendingSQC_InSLA += diffDays;
          listPendingSmartQc.push([duidVal, diffDays + " วัน", "#e0f2fe", "#0284c7", ownerDocVal]);
        }
      } else {
        var diffDays = Math.floor((smartQcDate - installDate) / (1000 * 60 * 60 * 24));
        if (diffDays > 3) { doneSQC_OverSLA++; totalDays_doneSQC_OverSLA += diffDays; }
        else              { doneSQC_InSLA++;   totalDays_doneSQC_InSLA += diffDays; }
      }

      if (!patDate) {
        if (smartQcDate) {
          var diffDaysPat = Math.floor((today - smartQcDate) / (1000 * 60 * 60 * 24));
          if (diffDaysPat > 5) {
            pendingPat_OverSLA++; totalDays_pendingPat_OverSLA += diffDaysPat;
            listPendingPat.push([duidVal, diffDaysPat + " วัน", "#fee2e2", "#dc2626", ownerDocVal]);
          } else {
            pendingPat_InSLA++; totalDays_pendingPat_InSLA += diffDaysPat;
            listPendingPat.push([duidVal, diffDaysPat + " วัน", "#e0f2fe", "#0284c7", ownerDocVal]);
          }
        }
      } else {
        if (smartQcDate) {
          var diffDays = Math.floor((patDate - smartQcDate) / (1000 * 60 * 60 * 24));
          if (diffDays > 5) { donePat_OverSLA++; totalDays_donePat_OverSLA += diffDays; }
          else              { donePat_InSLA++;   totalDays_donePat_InSLA += diffDays; }
        } else {
          var diffDays = Math.floor((patDate - installDate) / (1000 * 60 * 60 * 24));
          if (diffDays > 5) { donePat_OverSLA++; totalDays_donePat_OverSLA += diffDays; }
          else              { donePat_InSLA++;   totalDays_donePat_InSLA += diffDays; }
        }
      }

      if (patRemarkVal.toLowerCase().includes("not pass")) {
        var overDays = patDate
          ? Math.floor((today - patDate)     / (1000 * 60 * 60 * 24))
          : Math.floor((today - installDate) / (1000 * 60 * 60 * 24));
        listPatNotPass.push([duidVal, overDays + " วัน", "#ffedd5", "#ea580c", ownerDocVal]);
      }
    }

    function getAvg(totalDays, count) { return count > 0 ? (totalDays / count).toFixed(1) : "0.0"; }

    var maxCol = 11;
    var maxRow = 250;
    dashSheet.getRange(1, 1, maxRow, maxCol)
             .setBackground(C_BG).setFontFamily("Arial").setVerticalAlignment("middle").setWrap(true);

    dashSheet.setRowHeights(1, maxRow, 28);
    dashSheet.setRowHeight(1, 35);
    dashSheet.setRowHeight(2, 35);

    dashSheet.getRange("A1:K2").setBackground(C_HEAD_BG).setFontColor(C_HEAD_TXT);
    dashSheet.getRange("A1").setValue("AIS INSTALLATION & SLA DASHBOARD")
             .setFontSize(18).setFontWeight("bold").setVerticalAlignment("bottom");
    dashSheet.getRange("A2")
             .setValue("ข้อมูลอัปเดต ณ วันที่: " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm"))
             .setFontSize(10).setFontWeight("bold").setFontColor("#94a3b8").setVerticalAlignment("top");

    dashSheet.getRange("C2").setValue("📅 เลือกปี (Year):")
             .setFontSize(11).setFontWeight("bold").setFontColor("#94a3b8").setHorizontalAlignment("right");
    var filterCell = dashSheet.getRange("E2");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(yearList).build();
    filterCell.setDataValidation(rule).setValue(selectedYear)
              .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold")
              .setHorizontalAlignment("center")
              .setBorder(true, true, true, true, false, false, "#334155", SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange("E1").setValue("🔄 Auto-Refresh เมื่อเปลี่ยนปี")
             .setFontSize(9).setFontWeight("bold").setFontColor("#475569")
             .setVerticalAlignment("bottom").setHorizontalAlignment("center");

    dashSheet.getRange("I1:K2").merge().setValue(totalInstall + " ไซต์")
             .setFontSize(26).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    dashSheet.getRange("G1:H2").merge()
             .setValue("จำนวนงานติดตั้งเสร็จ\n(ตามปีที่เลือก)")
             .setFontSize(11).setFontWeight("bold").setFontColor("#cbd5e1")
             .setHorizontalAlignment("right").setVerticalAlignment("middle");

    function createBoldSummaryCard(rangeStr, title, dataArr) {
      var range = dashSheet.getRange(rangeStr);
      range.setBackground(C_CARD)
           .setBorder(true, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      dashSheet.getRange(range.getRow(), range.getColumn(), 1, 3).merge()
               .setValue(title).setBackground("#f8fafc").setFontColor("#0f172a")
               .setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(null, null, true, null, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      var dataRange = dashSheet.getRange(range.getRow() + 1, range.getColumn(), 5, 3);
      dataRange.setValues(dataArr).setFontWeight("bold").setFontColor("#1e293b").setFontSize(11);
      dashSheet.getRange(range.getRow() + 1, range.getColumn() + 1, 5, 2).setHorizontalAlignment("center");
      dataRange.setBorder(null, null, null, null, false, true, "#f1f5f9", SpreadsheetApp.BorderStyle.SOLID);
    }

    createBoldSummaryCard("A4:C9", "EXECUTIVE SUMMARY: SMART QC", [
      ["สถานะงาน", "จำนวนไซต์", "Aging เฉลี่ย (วัน)"],
      ["🟢 ปิดตามกำหนด (Done in SLA)",            doneSQC_InSLA,     Number(getAvg(totalDays_doneSQC_InSLA,   doneSQC_InSLA))],
      ["🟡 ปิดงานล่าช้า (Done Late)",              doneSQC_OverSLA,   Number(getAvg(totalDays_doneSQC_OverSLA,   doneSQC_OverSLA))],
      ["🔵 รอตรวจสอบ (Pending in SLA)",           pendingSQC_InSLA,  Number(getAvg(totalDays_pendingSQC_InSLA,  pendingSQC_InSLA))],
      ["🔴 ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingSQC_OverSLA, Number(getAvg(totalDays_pendingSQC_OverSLA, pendingSQC_OverSLA))]
    ]);

    createBoldSummaryCard("E4:G9", "EXECUTIVE SUMMARY: PAT SUBCON SUBMIT", [
      ["สถานะงาน", "จำนวนไซต์", "Aging เฉลี่ย (วัน)"],
      ["🟢 ส่งตามกำหนด (Done in SLA)",            donePat_InSLA,     Number(getAvg(totalDays_donePat_InSLA,   donePat_InSLA))],
      ["🟡 ส่งงานล่าช้า (Done Late)",              donePat_OverSLA,   Number(getAvg(totalDays_donePat_OverSLA,   donePat_OverSLA))],
      ["🔵 รอส่งงาน (Pending in SLA)",            pendingPat_InSLA,  Number(getAvg(totalDays_pendingPat_InSLA,  pendingPat_InSLA))],
      ["🔴 ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingPat_OverSLA, Number(getAvg(totalDays_pendingPat_OverSLA, pendingPat_OverSLA))]
    ]);

    dashSheet.getRange("Y1:Z30").clearContent();

    var sqcChartData = [
      ["Status", "Count"],
      ["ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingSQC_OverSLA > 0 ? pendingSQC_OverSLA : 0.0001],
      ["ปิดงานล่าช้า (Done Late)",             doneSQC_OverSLA   > 0 ? doneSQC_OverSLA   : 0.0001],
      ["ปิดตามกำหนด (Done in SLA)",            doneSQC_InSLA     > 0 ? doneSQC_InSLA     : 0.0001],
      ["รอตรวจสอบ (Pending in SLA)",           pendingSQC_InSLA  > 0 ? pendingSQC_InSLA  : 0.0001]
    ];
    var sqcSlices = { 0: { color: '#ef4444' }, 1: { color: '#eab308' }, 2: { color: '#22c55e' }, 3: { color: '#3b82f6' } };

    var patChartData = [
      ["Status", "Count"],
      ["ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingPat_OverSLA > 0 ? pendingPat_OverSLA : 0.0001],
      ["ส่งงานล่าช้า (Done Late)",             donePat_OverSLA   > 0 ? donePat_OverSLA   : 0.0001],
      ["ส่งตามกำหนด (Done in SLA)",            donePat_InSLA     > 0 ? donePat_InSLA     : 0.0001],
      ["รอส่งงาน (Pending in SLA)",            pendingPat_InSLA  > 0 ? pendingPat_InSLA  : 0.0001]
    ];
    var patSlices = { 0: { color: '#ef4444' }, 1: { color: '#eab308' }, 2: { color: '#22c55e' }, 3: { color: '#3b82f6' } };

    dashSheet.getRange(1, 25, sqcChartData.length, 2).setValues(sqcChartData);
    dashSheet.getRange(15, 25, patChartData.length, 2).setValues(patChartData);

    dashSheet.getRange("A11:C11").merge()
             .setValue("สัดส่วนผลงานภาพรวม 100% (Smart QC)")
             .setFontSize(13).setFontWeight("bold").setFontColor("#1e293b").setHorizontalAlignment("center");
    dashSheet.getRange("E11:G11").merge()
             .setValue("สัดส่วนผลงานภาพรวม 100% (PAT Subcon)")
             .setFontSize(13).setFontWeight("bold").setFontColor("#1e293b").setHorizontalAlignment("center");

    dashSheet.insertChart(
      dashSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(dashSheet.getRange(1, 25, sqcChartData.length, 2))
        .setPosition(12, 1, 0, 0)
        .setOption('pieHole', 0.45)
        .setOption('pieSliceText', 'percentage')
        .setOption('pieSliceTextStyle', { color: '#ffffff', fontSize: 14, bold: true, fontName: 'Arial' })
        .setOption('slices', sqcSlices)
        .setOption('legend', { position: 'bottom', textStyle: { color: '#0f172a', fontSize: 13, bold: true, fontName: 'Arial' } })
        .setOption('pieSliceBorderColor', '#ffffff')
        .setOption('backgroundColor', { fill: '#f8fafc' })
        .setOption('width',  520)
        .setOption('height', 450)
        .setOption('chartArea', { left: '5%', top: '2%', width: '90%', height: '70%' })
        .build()
    );

    dashSheet.insertChart(
      dashSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(dashSheet.getRange(15, 25, patChartData.length, 2))
        .setPosition(12, 5, 0, 0)
        .setOption('pieHole', 0.45)
        .setOption('pieSliceText', 'percentage')
        .setOption('pieSliceTextStyle', { color: '#ffffff', fontSize: 14, bold: true, fontName: 'Arial' })
        .setOption('slices', patSlices)
        .setOption('legend', { position: 'bottom', textStyle: { color: '#0f172a', fontSize: 13, bold: true, fontName: 'Arial' } })
        .setOption('pieSliceBorderColor', '#ffffff')
        .setOption('backgroundColor', { fill: '#f8fafc' })
        .setOption('width',  520)
        .setOption('height', 450)
        .setOption('chartArea', { left: '5%', top: '2%', width: '90%', height: '70%' })
        .build()
    );

    var listStartRow = 30;

    function writeProListCard(startCol, title, dataList, emptyMsg) {
      dashSheet.getRange(listStartRow, startCol, 1, 2).merge()
               .setValue(title).setBackground("#1e293b").setFontColor("#ffffff")
               .setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(true, true, true, true, null, null, "#1e293b", SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow + 1, startCol).setValue("Site DUID");
      dashSheet.getRange(listStartRow + 1, startCol + 1).setValue("Aging");
      dashSheet.getRange(listStartRow + 1, startCol, 1, 2)
               .setFontColor("#1e293b").setFontWeight("bold").setBackground("#f8fafc")
               .setHorizontalAlignment("center")
               .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      if (dataList.length > 0) {
        for (var k = 0; k < dataList.length; k++) {
          dashSheet.getRange(listStartRow + 2 + k, startCol, 1, 2)
                   .setValues([[dataList[k][0], dataList[k][1]]])
                   .setBackground(dataList[k][2]).setFontColor(dataList[k][3])
                   .setFontWeight("bold")
                   .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        }
        dashSheet.getRange(listStartRow + 2, startCol + 1, dataList.length, 1)
                 .setHorizontalAlignment("center").setFontWeight("bold");
      } else {
        dashSheet.getRange(listStartRow + 2, startCol, 1, 2).merge()
                 .setValue(emptyMsg).setBackground(C_CARD).setFontColor("#94a3b8")
                 .setFontWeight("bold").setHorizontalAlignment("center")
                 .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      }
    }

    function writeProListCardWithOwner(startCol, title, dataList, emptyMsg) {
      dashSheet.getRange(listStartRow, startCol, 1, 2).merge()
               .setValue(title).setBackground("#1e293b").setFontColor("#ffffff")
               .setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(true, true, true, true, null, null, "#1e293b", SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow, startCol + 2, 1, 1)
               .setValue("OWNER DOC").setBackground("#ca8a04").setFontColor("#ffffff")
               .setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(true, true, true, true, null, null, "#ca8a04", SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow + 1, startCol).setValue("Site DUID");
      dashSheet.getRange(listStartRow + 1, startCol + 1).setValue("Aging");
      dashSheet.getRange(listStartRow + 1, startCol, 1, 2)
               .setFontColor("#1e293b").setFontWeight("bold").setBackground("#f8fafc")
               .setHorizontalAlignment("center")
               .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow + 1, startCol + 2)
               .setValue("Assign").setFontColor("#1e293b").setFontWeight("bold")
               .setBackground("#fef9c3").setHorizontalAlignment("center")
               .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      if (dataList.length > 0) {
        for (var k = 0; k < dataList.length; k++) {
          dashSheet.getRange(listStartRow + 2 + k, startCol, 1, 2)
                   .setValues([[dataList[k][0], dataList[k][1]]])
                   .setBackground(dataList[k][2]).setFontColor(dataList[k][3])
                   .setFontWeight("bold")
                   .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
          dashSheet.getRange(listStartRow + 2 + k, startCol + 2)
                   .setValue(dataList[k][4] || "-")
                   .setBackground("#dcfce7").setFontColor("#15803d")
                   .setFontWeight("bold").setHorizontalAlignment("center")
                   .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        }
        dashSheet.getRange(listStartRow + 2, startCol + 1, dataList.length, 1)
                 .setHorizontalAlignment("center").setFontWeight("bold");
      } else {
        dashSheet.getRange(listStartRow + 2, startCol, 1, 2).merge()
                 .setValue(emptyMsg).setBackground(C_CARD).setFontColor("#94a3b8")
                 .setFontWeight("bold").setHorizontalAlignment("center")
                 .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        dashSheet.getRange(listStartRow + 2, startCol + 2)
                 .setValue("-").setBackground("#dcfce7").setFontColor("#15803d")
                 .setHorizontalAlignment("center")
                 .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      }
    }

    writeProListCard(1, "ACTION REQUIRED: SMART QC",     listPendingSmartQc, "All sites are up to date 🎉");
    writeProListCardWithOwner(5, "ACTION REQUIRED: PAT SUBCON",   listPendingPat,     "All sites are up to date 🎉");
    writeProListCardWithOwner(9, "REWORK REQUIRED: PAT NOT PASS", listPatNotPass,     "No rework required 🎉");

    dashSheet.setColumnWidth(1,  320);
    dashSheet.setColumnWidth(2,  90);
    dashSheet.setColumnWidth(3,  150);
    dashSheet.setColumnWidth(4,  20);
    dashSheet.setColumnWidth(5,  320);
    dashSheet.setColumnWidth(6,  90);
    dashSheet.setColumnWidth(7,  150);
    dashSheet.setColumnWidth(8,  20);
    dashSheet.setColumnWidth(9,  300);
    dashSheet.setColumnWidth(10, 90);
    dashSheet.setColumnWidth(11, 150);

    var lastUsedRow = listStartRow + 2 + Math.max(listPendingSmartQc.length, listPendingPat.length, listPatNotPass.length);
    if (lastUsedRow < maxRow) {
      dashSheet.getRange(lastUsedRow + 1, 1, maxRow - lastUsedRow, maxCol)
               .setBackground(C_BG).setBorder(false, false, false, false, false, false);
    }

  } catch (err) {
    Logger.log("createAisDashboard error: " + err.message + "\n" + err.stack);
    SpreadsheetApp.getUi().alert("⚠️ เกิดข้อผิดพลาดขณะสร้าง Dashboard:\n\n" + err.message);
  } finally {
    lock.releaseLock();
  }
}
// ---------------------------------------------------------
// 4. ฟังก์ชันหลักสำหรับ TRUE (เวอร์ชันแก้ไข: หา header/คอลัมน์แบบ dynamic)
// ---------------------------------------------------------

// helper: แปลง index (0-based) เป็นตัวอักษรคอลัมน์ เช่น 21 -> "V"
function colIdxToLetter_(idx) {
  var letter = "";
  idx = idx + 1; // เปลี่ยนเป็น 1-based
  while (idx > 0) {
    var rem = (idx - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    idx = Math.floor((idx - 1) / 26);
  }
  return letter;
}

// helper: หาแถวที่เป็น header จริง (แถวที่มีคำว่า "DU ID" อยู่)
function findTrueHeaderRow_(data) {
  for (var r = 0; r < Math.min(10, data.length); r++) {
    for (var c = 0; c < data[r].length; c++) {
      var val = data[r][c] ? String(data[r][c]).trim().toLowerCase() : "";
      if (val === "du id" || val.replace(/\s/g, "").includes("duid")) {
        return r;
      }
    }
  }
  return 0; // หาไม่เจอ ใช้แถว 1 เป็น fallback
}

// helper: หา index คอลัมน์จาก header โดยเช็คเงื่อนไขที่ส่งเข้ามา ถ้าไม่เจอใช้ fallback
function findTrueColIndex_(headerRow, matchFn, fallbackIndex) {
  for (var c = 0; c < headerRow.length; c++) {
    var val = headerRow[c] ? String(headerRow[c]).trim().toLowerCase() : "";
    if (matchFn(val)) return c;
  }
  return fallbackIndex;
}

function createTrueDashboard() {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
  } catch (lockErr) {
    SpreadsheetApp.getUi().alert("⏳ กำลังอัปเดตข้อมูลอยู่ (มีคนอื่น/รอบอื่นกำลังรันอยู่) กรุณารอสักครู่แล้วลองใหม่อีกครั้ง");
    return;
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ssTz = "Asia/Bangkok"; // ล็อค timezone ตรงๆ กันปัญหา getSpreadsheetTimeZone() ไม่ทำงานใน simple trigger
    var sourceSheetName = "True";
    var sourceSheet = ss.getSheetByName(sourceSheetName);

    if (!sourceSheet) {
      SpreadsheetApp.getUi().alert("⚠️ ไม่พบชีตที่ชื่อ '" + sourceSheetName + "' กรุณาตรวจสอบอีกครั้ง");
      return;
    }

    var dashSheetName = "Dashboard สรุปงาน TRUE";
    var dashSheet = ss.getSheetByName(dashSheetName);

    var selectedYear = "ทั้งหมด";
    if (dashSheet) {
      var oldFilter = dashSheet.getRange("E2").getValue();
      if (oldFilter && oldFilter !== "") {
        selectedYear = oldFilter.toString().trim();
      }
      dashSheet.clear();
    } else {
      dashSheet = ss.insertSheet(dashSheetName);
    }

    var existingCharts = dashSheet.getCharts();
    for (var i = 0; i < existingCharts.length; i++) {
      dashSheet.removeChart(existingCharts[i]);
    }

    dashSheet.setHiddenGridlines(true);

    var data = sourceSheet.getDataRange().getValues();
    if (!data || data.length <= 1) return;

    // --- หาแถว header จริง (ไม่ใช้ data[0] ตายตัวอีกต่อไป) ---
    var headerRowIdx = findTrueHeaderRow_(data);
    var headers = data[headerRowIdx];

    // --- หา index คอลัมน์จากชื่อ header จริง (fallback เป็นค่าตายตัวเดิมถ้าหาไม่เจอ) ---
    var idxDuid = findTrueColIndex_(headers, function (v) {
      return v === "du id" || v.replace(/\s/g, "").includes("duid");
    }, 1);

    var idxVerifyPhoto = findTrueColIndex_(headers, function (v) {
      return v === "verify photo"; // ใช้ Verify Photo (W) เป็นวันที่ตั้งต้นแทน Integration
    }, 22);

    var idxSmartQc = findTrueColIndex_(headers, function (v) {
      return v.replace(/\s/g, "").includes("smartqc");
    }, 24);

    var idxPatRemark = findTrueColIndex_(headers, function (v) {
      return v.includes("alarm category");
    }, 26);

    var idxPat = findTrueColIndex_(headers, function (v) {
      return v.includes("pat site folder") && v.indexOf("power") === -1; // กัน match "A129.1 Power PAT Site Folder"
    }, 28);

    var idxOwnerDoc = 32; // AG = Owner Doc

    var today = toDateOnly_(new Date(), ssTz);
    

    var availableYears = {};
    for (var i = headerRowIdx + 1; i < data.length; i++) {
      var instDate = data[i][idxVerifyPhoto];
      if (instDate && instDate !== "") {
        var d = instDate instanceof Date ? instDate : new Date(instDate);
        if (!isNaN(d.getTime())) {
          availableYears[d.getFullYear()] = true;
        }
      }
    }
    var yearList = Object.keys(availableYears).sort(function (a, b) { return b - a; });
    yearList.unshift("ทั้งหมด");

    var totalInstall = 0;

    var doneSQC_InSLA = 0;      var totalDays_doneSQC_InSLA = 0;
    var doneSQC_OverSLA = 0;    var totalDays_doneSQC_OverSLA = 0;
    var pendingSQC_InSLA = 0;   var totalDays_pendingSQC_InSLA = 0;
    var pendingSQC_OverSLA = 0; var totalDays_pendingSQC_OverSLA = 0;

    var donePat_InSLA = 0;      var totalDays_donePat_InSLA = 0;
    var donePat_OverSLA = 0;    var totalDays_donePat_OverSLA = 0;
    var pendingPat_InSLA = 0;   var totalDays_pendingPat_InSLA = 0;
    var pendingPat_OverSLA = 0; var totalDays_pendingPat_OverSLA = 0;

    var listPendingSmartQc = [];
    var listPendingPat     = [];
    var listPatNotPass     = [];

    var C_BG       = "#f8fafc";
    var C_CARD     = "#ffffff";
    var C_HEAD_BG  = "#0f172a";
    var C_HEAD_TXT = "#ffffff";
    var C_BORDER   = "#cbd5e1";

    // --- เริ่มลูปข้อมูลจริงหลัง header row แทนที่จะ hardcode i=1 ---
    for (var i = headerRowIdx + 1; i < data.length; i++) {
      var row = data[i];
      if (row.length <= idxPatRemark) continue;

      var installDateVal = row[idxVerifyPhoto];
      var smartQcDateVal = row[idxSmartQc];
      var patDateVal     = row[idxPat];
      var patRemarkVal   = row[idxPatRemark] ? String(row[idxPatRemark]).trim() : "";
      var duidVal        = row[idxDuid]      ? String(row[idxDuid]).trim()      : "ไม่ระบุ DUID";
      var ownerDocVal    = row[idxOwnerDoc]  ? String(row[idxOwnerDoc]).trim()  : "-";

      if (!installDateVal || installDateVal === "") continue;
      var installDate = toDateOnly_(installDateVal, ssTz);
      if (!installDate) continue;

      var dataYear = installDate.getFullYear().toString();
      if (selectedYear !== "ทั้งหมด" && dataYear !== selectedYear) continue;

      totalInstall++;

      var smartQcDate = null;
      if (smartQcDateVal && smartQcDateVal !== "") {
        smartQcDate = toDateOnly_(smartQcDateVal, ssTz);
      }
      var patDate = null;
      if (patDateVal && patDateVal !== "") {
        patDate = toDateOnly_(patDateVal, ssTz);
      }

      if (!smartQcDate) {
        var diffDays = Math.floor((today - installDate) / (1000 * 60 * 60 * 24));
        if (diffDays > 3) {
          pendingSQC_OverSLA++; totalDays_pendingSQC_OverSLA += diffDays;
          listPendingSmartQc.push([duidVal, diffDays + " วัน", "#fee2e2", "#dc2626", ownerDocVal]);
        } else {
          pendingSQC_InSLA++; totalDays_pendingSQC_InSLA += diffDays;
          listPendingSmartQc.push([duidVal, diffDays + " วัน", "#e0f2fe", "#0284c7", ownerDocVal]);
        }
      } else {
        var diffDays = Math.floor((smartQcDate - installDate) / (1000 * 60 * 60 * 24));
        if (diffDays > 3) { doneSQC_OverSLA++; totalDays_doneSQC_OverSLA += diffDays; }
        else              { doneSQC_InSLA++;   totalDays_doneSQC_InSLA += diffDays; }
      }

      if (!patDate) {
        if (smartQcDate) {
          var diffDaysPat = Math.floor((today - smartQcDate) / (1000 * 60 * 60 * 24));
          if (diffDaysPat > 5) {
            pendingPat_OverSLA++; totalDays_pendingPat_OverSLA += diffDaysPat;
            listPendingPat.push([duidVal, diffDaysPat + " วัน", "#fee2e2", "#dc2626", ownerDocVal]);
          } else {
            pendingPat_InSLA++; totalDays_pendingPat_InSLA += diffDaysPat;
            listPendingPat.push([duidVal, diffDaysPat + " วัน", "#e0f2fe", "#0284c7", ownerDocVal]);
          }
        }
      } else {
        if (smartQcDate) {
          var diffDays = Math.floor((patDate - smartQcDate) / (1000 * 60 * 60 * 24));
          if (diffDays > 5) { donePat_OverSLA++; totalDays_donePat_OverSLA += diffDays; }
          else              { donePat_InSLA++;   totalDays_donePat_InSLA += diffDays; }
        } else {
          var diffDays = Math.floor((patDate - installDate) / (1000 * 60 * 60 * 24));
          if (diffDays > 5) { donePat_OverSLA++; totalDays_donePat_OverSLA += diffDays; }
          else              { donePat_InSLA++;   totalDays_donePat_InSLA += diffDays; }
        }
      }

      // TRUE ใช้ "Integration alarm category": มีค่าและไม่ใช่ "No alarm" = ต้อง rework
      var hasAlarmIssue = patRemarkVal !== "" &&
                           !patRemarkVal.toLowerCase().includes("no alarm");
      if (hasAlarmIssue) {
        var overDays = patDate
          ? Math.floor((today - patDate)     / (1000 * 60 * 60 * 24))
          : Math.floor((today - installDate) / (1000 * 60 * 60 * 24));
        listPatNotPass.push([duidVal, overDays + " วัน", "#ffedd5", "#ea580c", ownerDocVal]);
      }
    }

    function getAvg(totalDays, count) { return count > 0 ? (totalDays / count).toFixed(1) : "0.0"; }

    var maxCol = 11;
    var maxRow = 250;
    dashSheet.getRange(1, 1, maxRow, maxCol)
             .setBackground(C_BG).setFontFamily("Arial").setVerticalAlignment("middle").setWrap(true);

    dashSheet.setRowHeights(1, maxRow, 28);
    dashSheet.setRowHeight(1, 35);
    dashSheet.setRowHeight(2, 35);

    dashSheet.getRange("A1:K2").setBackground(C_HEAD_BG).setFontColor(C_HEAD_TXT);
    dashSheet.getRange("A1").setValue("TRUE INSTALLATION & SLA DASHBOARD")
             .setFontSize(18).setFontWeight("bold").setVerticalAlignment("bottom");
    dashSheet.getRange("A2")
             .setValue("ข้อมูลอัปเดต ณ วันที่: " + Utilities.formatDate(new Date(), "GMT+7", "dd/MM/yyyy HH:mm") +
                        "   |   ตรวจพบ " + totalInstall + " แถว (header แถวที่ " + (headerRowIdx + 1) +
                        ", VerifyPhoto=" + colIdxToLetter_(idxVerifyPhoto) +
                        ", SmartQC=" + colIdxToLetter_(idxSmartQc) +
                        ", PAT=" + colIdxToLetter_(idxPat) +
                        ", AlarmCat=" + colIdxToLetter_(idxPatRemark) +
                        ", Owner=" + colIdxToLetter_(idxOwnerDoc) + ")")
             .setFontSize(9).setFontWeight("bold").setFontColor("#94a3b8").setVerticalAlignment("top");

    dashSheet.getRange("C2").setValue("📅 เลือกปี (Year):")
             .setFontSize(11).setFontWeight("bold").setFontColor("#94a3b8").setHorizontalAlignment("right");
    var filterCell = dashSheet.getRange("E2");
    var rule = SpreadsheetApp.newDataValidation().requireValueInList(yearList).build();
    filterCell.setDataValidation(rule).setValue(selectedYear)
              .setBackground("#1e293b").setFontColor("#ffffff").setFontWeight("bold")
              .setHorizontalAlignment("center")
              .setBorder(true, true, true, true, false, false, "#334155", SpreadsheetApp.BorderStyle.SOLID);
    dashSheet.getRange("E1").setValue("🔄 Auto-Refresh เมื่อเปลี่ยนปี")
             .setFontSize(9).setFontWeight("bold").setFontColor("#475569")
             .setVerticalAlignment("bottom").setHorizontalAlignment("center");

    dashSheet.getRange("I1:K2").merge().setValue(totalInstall + " ไซต์")
             .setFontSize(26).setFontWeight("bold").setHorizontalAlignment("right").setVerticalAlignment("middle");
    dashSheet.getRange("G1:H2").merge()
             .setValue("จำนวนงานติดตั้งเสร็จ\n(ตามปีที่เลือก)")
             .setFontSize(11).setFontWeight("bold").setFontColor("#cbd5e1")
             .setHorizontalAlignment("right").setVerticalAlignment("middle");

    function createBoldSummaryCard(rangeStr, title, dataArr) {
      var range = dashSheet.getRange(rangeStr);
      range.setBackground(C_CARD)
           .setBorder(true, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
      dashSheet.getRange(range.getRow(), range.getColumn(), 1, 3).merge()
               .setValue(title).setBackground("#f8fafc").setFontColor("#0f172a")
               .setFontSize(12).setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(null, null, true, null, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      var dataRange = dashSheet.getRange(range.getRow() + 1, range.getColumn(), 5, 3);
      dataRange.setValues(dataArr).setFontWeight("bold").setFontColor("#1e293b").setFontSize(11);
      dashSheet.getRange(range.getRow() + 1, range.getColumn() + 1, 5, 2).setHorizontalAlignment("center");
      dataRange.setBorder(null, null, null, null, false, true, "#f1f5f9", SpreadsheetApp.BorderStyle.SOLID);
    }

    createBoldSummaryCard("A4:C9", "EXECUTIVE SUMMARY: SMART QC", [
      ["สถานะงาน", "จำนวนไซต์", "Aging เฉลี่ย (วัน)"],
      ["🟢 ปิดตามกำหนด (Done in SLA)",            doneSQC_InSLA,     Number(getAvg(totalDays_doneSQC_InSLA,   doneSQC_InSLA))],
      ["🟡 ปิดงานล่าช้า (Done Late)",              doneSQC_OverSLA,   Number(getAvg(totalDays_doneSQC_OverSLA,   doneSQC_OverSLA))],
      ["🔵 รอตรวจสอบ (Pending in SLA)",           pendingSQC_InSLA,  Number(getAvg(totalDays_pendingSQC_InSLA,  pendingSQC_InSLA))],
      ["🔴 ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingSQC_OverSLA, Number(getAvg(totalDays_pendingSQC_OverSLA, pendingSQC_OverSLA))]
    ]);

    createBoldSummaryCard("E4:G9", "EXECUTIVE SUMMARY: PAT SUBCON SUBMIT", [
      ["สถานะงาน", "จำนวนไซต์", "Aging เฉลี่ย (วัน)"],
      ["🟢 ส่งตามกำหนด (Done in SLA)",            donePat_InSLA,     Number(getAvg(totalDays_donePat_InSLA,   donePat_InSLA))],
      ["🟡 ส่งงานล่าช้า (Done Late)",              donePat_OverSLA,   Number(getAvg(totalDays_donePat_OverSLA,   donePat_OverSLA))],
      ["🔵 รอส่งงาน (Pending in SLA)",            pendingPat_InSLA,  Number(getAvg(totalDays_pendingPat_InSLA,  pendingPat_InSLA))],
      ["🔴 ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingPat_OverSLA, Number(getAvg(totalDays_pendingPat_OverSLA, pendingPat_OverSLA))]
    ]);

    dashSheet.getRange("Y1:Z30").clearContent();

    var sqcChartData = [
      ["Status", "Count"],
      ["ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingSQC_OverSLA > 0 ? pendingSQC_OverSLA : 0.0001],
      ["ปิดงานล่าช้า (Done Late)",             doneSQC_OverSLA   > 0 ? doneSQC_OverSLA   : 0.0001],
      ["ปิดตามกำหนด (Done in SLA)",            doneSQC_InSLA     > 0 ? doneSQC_InSLA     : 0.0001],
      ["รอตรวจสอบ (Pending in SLA)",           pendingSQC_InSLA  > 0 ? pendingSQC_InSLA  : 0.0001]
    ];
    var sqcSlices = { 0: { color: '#ef4444' }, 1: { color: '#eab308' }, 2: { color: '#22c55e' }, 3: { color: '#3b82f6' } };

    var patChartData = [
      ["Status", "Count"],
      ["ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)", pendingPat_OverSLA > 0 ? pendingPat_OverSLA : 0.0001],
      ["ส่งงานล่าช้า (Done Late)",             donePat_OverSLA   > 0 ? donePat_OverSLA   : 0.0001],
      ["ส่งตามกำหนด (Done in SLA)",            donePat_InSLA     > 0 ? donePat_InSLA     : 0.0001],
      ["รอส่งงาน (Pending in SLA)",            pendingPat_InSLA  > 0 ? pendingPat_InSLA  : 0.0001]
    ];
    var patSlices = { 0: { color: '#ef4444' }, 1: { color: '#eab308' }, 2: { color: '#22c55e' }, 3: { color: '#3b82f6' } };

    dashSheet.getRange(1, 25, sqcChartData.length, 2).setValues(sqcChartData);
    dashSheet.getRange(15, 25, patChartData.length, 2).setValues(patChartData);

    dashSheet.getRange("A11:C11").merge()
             .setValue("สัดส่วนผลงานภาพรวม 100% (Smart QC)")
             .setFontSize(13).setFontWeight("bold").setFontColor("#1e293b").setHorizontalAlignment("center");
    dashSheet.getRange("E11:G11").merge()
             .setValue("สัดส่วนผลงานภาพรวม 100% (PAT Subcon)")
             .setFontSize(13).setFontWeight("bold").setFontColor("#1e293b").setHorizontalAlignment("center");

    dashSheet.insertChart(
      dashSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(dashSheet.getRange(1, 25, sqcChartData.length, 2))
        .setPosition(12, 1, 0, 0)
        .setOption('pieHole', 0.45)
        .setOption('pieSliceText', 'percentage')
        .setOption('pieSliceTextStyle', { color: '#ffffff', fontSize: 14, bold: true, fontName: 'Arial' })
        .setOption('slices', sqcSlices)
        .setOption('legend', { position: 'bottom', textStyle: { color: '#0f172a', fontSize: 13, bold: true, fontName: 'Arial' } })
        .setOption('pieSliceBorderColor', '#ffffff')
        .setOption('backgroundColor', { fill: '#f8fafc' })
        .setOption('width',  520)
        .setOption('height', 450)
        .setOption('chartArea', { left: '5%', top: '2%', width: '90%', height: '70%' })
        .build()
    );

    dashSheet.insertChart(
      dashSheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(dashSheet.getRange(15, 25, patChartData.length, 2))
        .setPosition(12, 5, 0, 0)
        .setOption('pieHole', 0.45)
        .setOption('pieSliceText', 'percentage')
        .setOption('pieSliceTextStyle', { color: '#ffffff', fontSize: 14, bold: true, fontName: 'Arial' })
        .setOption('slices', patSlices)
        .setOption('legend', { position: 'bottom', textStyle: { color: '#0f172a', fontSize: 13, bold: true, fontName: 'Arial' } })
        .setOption('pieSliceBorderColor', '#ffffff')
        .setOption('backgroundColor', { fill: '#f8fafc' })
        .setOption('width',  520)
        .setOption('height', 450)
        .setOption('chartArea', { left: '5%', top: '2%', width: '90%', height: '70%' })
        .build()
    );

    var listStartRow = 30;

    function writeProListCard(startCol, title, dataList, emptyMsg) {
      dashSheet.getRange(listStartRow, startCol, 1, 2).merge()
               .setValue(title).setBackground("#1e293b").setFontColor("#ffffff")
               .setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(true, true, true, true, null, null, "#1e293b", SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow + 1, startCol).setValue("Site DUID");
      dashSheet.getRange(listStartRow + 1, startCol + 1).setValue("Aging");
      dashSheet.getRange(listStartRow + 1, startCol, 1, 2)
               .setFontColor("#1e293b").setFontWeight("bold").setBackground("#f8fafc")
               .setHorizontalAlignment("center")
               .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      if (dataList.length > 0) {
        for (var k = 0; k < dataList.length; k++) {
          dashSheet.getRange(listStartRow + 2 + k, startCol, 1, 2)
                   .setValues([[dataList[k][0], dataList[k][1]]])
                   .setBackground(dataList[k][2]).setFontColor(dataList[k][3])
                   .setFontWeight("bold")
                   .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        }
        dashSheet.getRange(listStartRow + 2, startCol + 1, dataList.length, 1)
                 .setHorizontalAlignment("center").setFontWeight("bold");
      } else {
        dashSheet.getRange(listStartRow + 2, startCol, 1, 2).merge()
                 .setValue(emptyMsg).setBackground(C_CARD).setFontColor("#94a3b8")
                 .setFontWeight("bold").setHorizontalAlignment("center")
                 .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      }
    }

    function writeProListCardWithOwner(startCol, title, dataList, emptyMsg) {
      dashSheet.getRange(listStartRow, startCol, 1, 2).merge()
               .setValue(title).setBackground("#1e293b").setFontColor("#ffffff")
               .setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(true, true, true, true, null, null, "#1e293b", SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow, startCol + 2, 1, 1)
               .setValue("OWNER DOC").setBackground("#ca8a04").setFontColor("#ffffff")
               .setFontWeight("bold").setHorizontalAlignment("center")
               .setBorder(true, true, true, true, null, null, "#ca8a04", SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow + 1, startCol).setValue("Site DUID");
      dashSheet.getRange(listStartRow + 1, startCol + 1).setValue("Aging");
      dashSheet.getRange(listStartRow + 1, startCol, 1, 2)
               .setFontColor("#1e293b").setFontWeight("bold").setBackground("#f8fafc")
               .setHorizontalAlignment("center")
               .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      dashSheet.getRange(listStartRow + 1, startCol + 2)
               .setValue("Assign").setFontColor("#1e293b").setFontWeight("bold")
               .setBackground("#fef9c3").setHorizontalAlignment("center")
               .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      if (dataList.length > 0) {
        for (var k = 0; k < dataList.length; k++) {
          dashSheet.getRange(listStartRow + 2 + k, startCol, 1, 2)
                   .setValues([[dataList[k][0], dataList[k][1]]])
                   .setBackground(dataList[k][2]).setFontColor(dataList[k][3])
                   .setFontWeight("bold")
                   .setBorder(false, true, true, true, true, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
          dashSheet.getRange(listStartRow + 2 + k, startCol + 2)
                   .setValue(dataList[k][4] || "-")
                   .setBackground("#dcfce7").setFontColor("#15803d")
                   .setFontWeight("bold").setHorizontalAlignment("center")
                   .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        }
        dashSheet.getRange(listStartRow + 2, startCol + 1, dataList.length, 1)
                 .setHorizontalAlignment("center").setFontWeight("bold");
      } else {
        dashSheet.getRange(listStartRow + 2, startCol, 1, 2).merge()
                 .setValue(emptyMsg).setBackground(C_CARD).setFontColor("#94a3b8")
                 .setFontWeight("bold").setHorizontalAlignment("center")
                 .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
        dashSheet.getRange(listStartRow + 2, startCol + 2)
                 .setValue("-").setBackground("#dcfce7").setFontColor("#15803d")
                 .setHorizontalAlignment("center")
                 .setBorder(false, true, true, true, false, false, C_BORDER, SpreadsheetApp.BorderStyle.SOLID);
      }
    }

    writeProListCard(1, "ACTION REQUIRED: SMART QC",     listPendingSmartQc, "All sites are up to date 🎉");
    writeProListCardWithOwner(5, "ACTION REQUIRED: PAT SUBCON",   listPendingPat,     "All sites are up to date 🎉");
    writeProListCardWithOwner(9, "REWORK REQUIRED: ALARM FOUND", listPatNotPass,     "No rework required 🎉");

    dashSheet.setColumnWidth(1,  320);
    dashSheet.setColumnWidth(2,  90);
    dashSheet.setColumnWidth(3,  150);
    dashSheet.setColumnWidth(4,  20);
    dashSheet.setColumnWidth(5,  320);
    dashSheet.setColumnWidth(6,  90);
    dashSheet.setColumnWidth(7,  150);
    dashSheet.setColumnWidth(8,  20);
    dashSheet.setColumnWidth(9,  300);
    dashSheet.setColumnWidth(10, 90);
    dashSheet.setColumnWidth(11, 150);

    var lastUsedRow = listStartRow + 2 + Math.max(listPendingSmartQc.length, listPendingPat.length, listPatNotPass.length);
    if (lastUsedRow < maxRow) {
      dashSheet.getRange(lastUsedRow + 1, 1, maxRow - lastUsedRow, maxCol)
               .setBackground(C_BG).setBorder(false, false, false, false, false, false);
    }

  } catch (err) {
    Logger.log("createTrueDashboard error: " + err.message + "\n" + err.stack);
    SpreadsheetApp.getUi().alert("⚠️ เกิดข้อผิดพลาดขณะสร้าง Dashboard:\n\n" + err.message);
  } finally {
    lock.releaseLock();
  }
}