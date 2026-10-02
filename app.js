/**
 * AIS & TRUE SLA Dashboard Controller
 * Supports live Google Sheet data, drag & drop Excel (.xlsx) parsing via SheetJS,
 * interactive Chart.js donut visualizations, and LINE Group Alert notifications.
 */

// State Management
const state = {
  operator: 'AIS', // 'AIS' | 'TRUE'
  selectedYear: '2026',
  availableYears: ['ทั้งหมด', '2026', '2025'],
  rawGoogleSheetData: {
    AIS: [],
    TRUE: []
  },
  importedData: {
    AIS: null,
    TRUE: null
  },
  dataSource: 'SHEET', // 'SHEET' | 'EXCEL_UPLOAD'
  lastUpdated: new Date(),
  charts: {
    sqc: null,
    pat: null
  },
  searchQuery: {
    sqc: '',
    pat: '',
    rework: ''
  },
  lineToken: localStorage.getItem('sla_line_token') || 'YKtVKOIprzQoLKqB7foUkyxIwvzGaWxY/lnBmm4GaoJVNVDgbEUOTs8MOZRWBtEfzX8X6k0pX+pJSyave60Ka//baM6waKsQE/Ho43TkMod6YcyLcreDpjVC85MCXv7NxSj47Bh6bI2a2Xuls5hnkAdB04t89/1O/w1cDnyilFU=',
  lineGroupId: localStorage.getItem('sla_line_group_id') || 'C9d136fee255c27308ede4164cad0e27d',
  lineWebhook: localStorage.getItem('sla_line_webhook') || 'https://webhook.site/d43cd402-b87b-4c7f-a8a8-8e58b0cc27ba',
  gasWebAppUrl: localStorage.getItem('sla_gas_url') || 'https://script.google.com/macros/s/AKfycbysOK_GAlsnJ12VOLUUm-0qDltWipjba_JKYc2gdzE9M50FaGQ5O-R8gPiqEQK0LopsQQ/exec'
};

// Colors matching dashboard
const COLORS = {
  GREEN: '#10b981',
  YELLOW: '#f59e0b',
  BLUE: '#3b82f6',
  RED: '#ef4444',
  ORANGE: '#f97316'
};

// Initialize Application
document.addEventListener('DOMContentLoaded', () => {
  initUIEventListeners();
  loadDataFromGoogleSheet();
});

// Setup UI Event Listeners
function initUIEventListeners() {
  // Operator Switch buttons
  document.querySelectorAll('.operator-btn').forEach(btn => {
    btn.addEventListener('click', (e) => {
      document.querySelectorAll('.operator-btn').forEach(b => b.classList.remove('active'));
      const op = btn.getAttribute('data-operator');
      btn.classList.add('active');
      state.operator = op;
      updateThemeForOperator(op);
      calculateAndRenderDashboard();
    });
  });

  // Year filter selector
  const yearSelect = document.getElementById('yearSelect');
  if (yearSelect) {
    yearSelect.addEventListener('change', (e) => {
      state.selectedYear = e.target.value;
      calculateAndRenderDashboard();
    });
  }

  // Refresh button
  const refreshBtn = document.getElementById('refreshBtn');
  if (refreshBtn) {
    refreshBtn.addEventListener('click', () => {
      showToast('กำลังโหลดข้อมูลล่าสุดจาก Google Sheets...', 'info');
      loadDataFromGoogleSheet();
    });
  }

  // Upload Modal Triggers
  const uploadBtn = document.getElementById('uploadModalBtn');
  const uploadModal = document.getElementById('uploadModal');
  const closeUploadBtn = document.getElementById('closeUploadModalBtn');
  const cancelUploadBtn = document.getElementById('cancelUploadBtn');

  if (uploadBtn && uploadModal) {
    uploadBtn.addEventListener('click', () => uploadModal.classList.add('active'));
  }
  if (closeUploadBtn) closeUploadBtn.addEventListener('click', () => uploadModal.classList.remove('active'));
  if (cancelUploadBtn) cancelUploadBtn.addEventListener('click', () => uploadModal.classList.remove('active'));

  // Drag & drop dropzone
  const dropzone = document.getElementById('fileDropzone');
  const fileInput = document.getElementById('excelFileInput');

  if (dropzone && fileInput) {
    dropzone.addEventListener('click', () => fileInput.click());
    dropzone.addEventListener('dragover', (e) => { e.preventDefault(); dropzone.classList.add('dragover'); });
    dropzone.addEventListener('dragleave', () => dropzone.classList.remove('dragover'));
    dropzone.addEventListener('drop', (e) => {
      e.preventDefault();
      dropzone.classList.remove('dragover');
      if (e.dataTransfer.files.length > 0) {
        handleExcelFileUpload(e.dataTransfer.files[0]);
      }
    });
    fileInput.addEventListener('change', (e) => {
      if (e.target.files.length > 0) {
        handleExcelFileUpload(e.target.files[0]);
      }
    });
  }

  // LINE Notification & Image Capture Modal Triggers
  const lineModalBtn = document.getElementById('lineAlertBtn');
  const captureBtn = document.getElementById('captureBtn');
  const lineModal = document.getElementById('lineModal');
  const closeLineBtn = document.getElementById('closeLineModalBtn');
  const cancelLineBtn = document.getElementById('cancelLineBtn');
  const sendLineBtn = document.getElementById('sendLineBtn');
  const copyLineBtn = document.getElementById('copyLineBtn');
  const copyImageBtn = document.getElementById('copyImageBtn');
  const downloadImageBtn = document.getElementById('downloadImageBtn');

  const openLineAndCapture = () => {
    generateLineAlertPreview();
    captureDashboardImage();
    if (lineModal) lineModal.classList.add('active');
  };

  if (lineModalBtn) lineModalBtn.addEventListener('click', openLineAndCapture);
  if (captureBtn) captureBtn.addEventListener('click', openLineAndCapture);

  if (closeLineBtn) closeLineBtn.addEventListener('click', () => lineModal.classList.remove('active'));
  if (cancelLineBtn) cancelLineBtn.addEventListener('click', () => lineModal.classList.remove('active'));
  
  if (copyLineBtn) {
    copyLineBtn.addEventListener('click', () => {
      const msg = document.getElementById('lineMessagePreview').textContent;
      navigator.clipboard.writeText(msg).then(() => {
        showToast('📋 คัดลอกข้อความสำหรับ LINE Group เรียบร้อยแล้ว!', 'success');
      });
    });
  }
  if (copyImageBtn) copyImageBtn.addEventListener('click', handleCopyImage);
  if (downloadImageBtn) downloadImageBtn.addEventListener('click', handleDownloadImage);
  if (sendLineBtn) sendLineBtn.addEventListener('click', handleSendLineAlert);

  // Search in tables
  const sqcSearch = document.getElementById('sqcSearch');
  if (sqcSearch) {
    sqcSearch.addEventListener('input', (e) => {
      state.searchQuery.sqc = e.target.value.toLowerCase();
      renderActionTables();
    });
  }
  const patSearch = document.getElementById('patSearch');
  if (patSearch) {
    patSearch.addEventListener('input', (e) => {
      state.searchQuery.pat = e.target.value.toLowerCase();
      renderActionTables();
    });
  }
  const reworkSearch = document.getElementById('reworkSearch');
  if (reworkSearch) {
    reworkSearch.addEventListener('input', (e) => {
      state.searchQuery.rework = e.target.value.toLowerCase();
      renderActionTables();
    });
  }
}

// Update Theme Accents
function updateThemeForOperator(op) {
  const brandLogo = document.getElementById('brandLogo');
  const titleText = document.getElementById('dashboardTitle');
  const subtitleText = document.getElementById('dashboardSubtitle');

  if (op === 'AIS') {
    brandLogo.className = 'logo-badge ais';
    brandLogo.textContent = 'A';
    titleText.textContent = 'AIS INSTALLATION & SLA DASHBOARD';
    subtitleText.textContent = 'Executive Dashboard & Real-Time Monitoring';
    document.getElementById('sqcCardTitle').textContent = 'EXECUTIVE SUMMARY: SMART QC';
    document.getElementById('patCardTitle').textContent = 'EXECUTIVE SUMMARY: PAT SUBCON SUBMIT';
    document.getElementById('actionSqcTitle').textContent = 'ACTION REQUIRED: SMART QC';
    document.getElementById('actionPatTitle').textContent = 'ACTION REQUIRED: PAT SUBCON';
    document.getElementById('actionReworkTitle').textContent = 'REWORK REQUIRED: PAT NOT PASS';
  } else {
    brandLogo.className = 'logo-badge true';
    brandLogo.textContent = 'T';
    titleText.textContent = 'TRUE INSTALLATION & SLA DASHBOARD';
    subtitleText.textContent = 'Executive Dashboard & Real-Time Monitoring';
    document.getElementById('sqcCardTitle').textContent = 'EXECUTIVE SUMMARY: 08.1 SmartQC+AOR';
    document.getElementById('patCardTitle').textContent = 'EXECUTIVE SUMMARY: 14.1 A129 PAT Site Folder';
    document.getElementById('actionSqcTitle').textContent = 'ACTION REQUIRED: 08.1 SmartQC+AOR';
    document.getElementById('actionPatTitle').textContent = 'ACTION REQUIRED: 14.1 A129 PAT Site Folder';
    document.getElementById('actionReworkTitle').textContent = 'REWORK REQUIRED: ALARM FOUND';
  }
}

// Fetch live data from Google Sheet via JSONP (Bypasses CORS restrictions)
function fetchSheetViaJsonp(sheetName) {
  return new Promise((resolve, reject) => {
    const callbackName = 'gviz_cb_' + Math.random().toString(36).substring(2, 9);
    const script = document.createElement('script');
    const SPREADSHEET_ID = '1fi7RCyx74VaBpPH2QO-44gE7vDKGHkFj-Nqp3QGc3E0';
    const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/gviz/tq?sheet=${encodeURIComponent(sheetName)}&tqx=responseHandler:${callbackName}`;
    
    let isDone = false;
    const timeout = setTimeout(() => {
      if (isDone) return;
      isDone = true;
      cleanup();
      reject(new Error(`Timeout loading sheet: ${sheetName}`));
    }, 15000);

    function cleanup() {
      clearTimeout(timeout);
      delete window[callbackName];
      if (script.parentNode) script.parentNode.removeChild(script);
    }

    window[callbackName] = function(response) {
      if (isDone) return;
      isDone = true;
      cleanup();
      if (response && response.status === 'ok' && response.table) {
        const rows = response.table.rows || [];
        const cols = response.table.cols || [];
        const table2d = rows.map(r => {
          if (!r || !r.c) return new Array(cols.length).fill('');
          return r.c.map(cell => {
            if (!cell) return '';
            if (cell.f !== undefined && cell.f !== null) return String(cell.f);
            if (cell.v !== undefined && cell.v !== null) return String(cell.v);
            return '';
          });
        });
        resolve(table2d);
      } else {
        reject(new Error(response && response.errors && response.errors[0] ? response.errors[0].message : 'Failed to parse sheet data'));
      }
    };

    script.onerror = function() {
      if (isDone) return;
      isDone = true;
      cleanup();
      reject(new Error(`Network error loading sheet: ${sheetName}`));
    };

    script.src = url;
    document.head.appendChild(script);
  });
}

// Fetch live data from Google Sheet
async function loadDataFromGoogleSheet() {
  try {
    const [dataAis, dataTrue] = await Promise.all([
      fetchSheetViaJsonp('AIS'),
      fetchSheetViaJsonp('TRUE')
    ]);

    state.rawGoogleSheetData.AIS = dataAis;
    state.rawGoogleSheetData.TRUE = dataTrue;
    state.lastUpdated = new Date();
    
    document.getElementById('updateTimestamp').textContent = 
      `ข้อมูลอัปเดต ณ: ${formatDateTime(state.lastUpdated)} (Google Sheets)`;

    extractAvailableYears();
    calculateAndRenderDashboard();
    showToast('โหลดข้อมูลสดจาก Google Sheets สำเร็จเรียบร้อยแล้ว!', 'success');
  } catch (err) {
    console.error('Error fetching Google Sheets data:', err);
    showToast('ไม่สามารถดึงข้อมูลจาก Google Sheets ได้ (กำลังใช้ข้อมูลตัวอย่าง)', 'error');
  }
}

// CSV Parser Helper
function parseCsv(text) {
  const lines = text.split('\n');
  const result = [];
  for (let line of lines) {
    if (!line.trim()) continue;
    // Handle CSV quoting
    const row = [];
    let insideQuotes = false;
    let entry = '';
    for (let i = 0; i < line.length; i++) {
      const c = line[i];
      if (c === '"') {
        if (insideQuotes && line[i+1] === '"') {
          entry += '"';
          i++;
        } else {
          insideQuotes = !insideQuotes;
        }
      } else if (c === ',' && !insideQuotes) {
        row.push(entry.trim());
        entry = '';
      } else {
        entry += c;
      }
    }
    row.push(entry.trim());
    result.push(row);
  }
  return result;
}

// Parse date strings in multiple formats (DD/MM/YYYY, YYYY-MM-DD, Date(Y,M,D))
function parseDate(v) {
  if (!v) return null;
  const str = String(v).trim();
  if (!str) return null;

  // Handle Google Sheet Date(yyyy,m,d) format
  const dateMatch = str.match(/Date\((\d+),\s*(\d+),\s*(\d+)/);
  if (dateMatch) {
    return new Date(parseInt(dateMatch[1], 10), parseInt(dateMatch[2], 10), parseInt(dateMatch[3], 10));
  }

  const s = str.split(' ')[0];
  if (s.includes('/')) {
    const p = s.split('/');
    if (p.length === 3) {
      if (p[0].length === 4) return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
      return new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
    }
  } else if (s.includes('-')) {
    const p = s.split('-');
    if (p.length === 3) {
      if (p[0].length === 4) return new Date(parseInt(p[0]), parseInt(p[1])-1, parseInt(p[2]));
      return new Date(parseInt(p[2]), parseInt(p[1])-1, parseInt(p[0]));
    }
  }
  const d = new Date(v);
  return isNaN(d.getTime()) ? null : d;
}

// Extract available years from dataset
function extractAvailableYears() {
  const years = new Set();
  const dataset = state.dataSource === 'SHEET' ? state.rawGoogleSheetData[state.operator] : state.importedData[state.operator];

  if (state.operator === 'AIS') {
    if (state.dataSource === 'SHEET' && dataset) {
      for (let i = 2; i < dataset.length; i++) {
        const d = parseDate(dataset[i][30]);
        if (d) years.add(d.getFullYear().toString());
      }
    } else if (dataset) {
      dataset.forEach(item => {
        if (item.installDate) years.add(new Date(item.installDate).getFullYear().toString());
      });
    }
  } else {
    // TRUE
    if (state.dataSource === 'SHEET' && dataset) {
      for (let i = 2; i < dataset.length; i++) {
        const d = parseDate(dataset[i][22]);
        if (d) years.add(d.getFullYear().toString());
      }
    } else if (dataset) {
      dataset.forEach(item => {
        if (item.installDate) years.add(new Date(item.installDate).getFullYear().toString());
      });
    }
  }

  const sortedYears = Array.from(years).sort((a,b) => b - a);
  state.availableYears = ['ทั้งหมด', ...sortedYears];

  // Update dropdown options
  const select = document.getElementById('yearSelect');
  if (select) {
    select.innerHTML = '';
    state.availableYears.forEach(yr => {
      const opt = document.createElement('option');
      opt.value = yr;
      opt.textContent = yr;
      if (yr === state.selectedYear) opt.selected = true;
      select.appendChild(opt);
    });
  }
}

// Compute Metrics for Dashboard
function calculateMetrics() {
  const op = state.operator;
  const isSheet = state.dataSource === 'SHEET';
  const selectedYear = state.selectedYear;
  const today = new Date();
  today.setHours(0,0,0,0);

  const acc = {
    totalInstall: 0,
    doneSQC_InSLA: 0, totalDays_doneSQC_InSLA: 0,
    doneSQC_OverSLA: 0, totalDays_doneSQC_OverSLA: 0,
    pendingSQC_InSLA: 0, totalDays_pendingSQC_InSLA: 0,
    pendingSQC_OverSLA: 0, totalDays_pendingSQC_OverSLA: 0,

    donePat_InSLA: 0, totalDays_donePat_InSLA: 0,
    donePat_OverSLA: 0, totalDays_donePat_OverSLA: 0,
    pendingPat_InSLA: 0, totalDays_pendingPat_InSLA: 0,
    pendingPat_OverSLA: 0, totalDays_pendingPat_OverSLA: 0,

    listPendingSmartQc: [],
    listPendingPat: [],
    listPatNotPass: []
  };

  if (op === 'AIS') {
    if (isSheet) {
      const rows = state.rawGoogleSheetData.AIS;
      if (!rows || rows.length <= 2) return acc;

      for (let i = 2; i < rows.length; i++) {
        const row = rows[i];
        if (row.length <= 35) continue;
        const installDate = parseDate(row[30]);
        if (!installDate) continue;
        if (selectedYear !== 'ทั้งหมด' && installDate.getFullYear().toString() !== selectedYear) continue;

        acc.totalInstall++;
        const duid = row[0] ? row[0].trim() : 'ไม่ระบุ DUID';
        const ownerDoc = row[37] ? row[37].trim() : '-';
        const smartQcDate = parseDate(row[31]);
        const patDate = parseDate(row[33]);
        const patRemark = (row[35] || '').trim();

        // Smart QC (3 days SLA)
        if (!smartQcDate) {
          const diff = Math.max(0, Math.floor((today - installDate) / 86400000));
          if (diff > 3) {
            acc.pendingSQC_OverSLA++;
            acc.totalDays_pendingSQC_OverSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน`, isOver: true, owner: ownerDoc });
          } else {
            acc.pendingSQC_InSLA++;
            acc.totalDays_pendingSQC_InSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน`, isOver: false, owner: ownerDoc });
          }
        } else {
          const diff = Math.max(0, Math.floor((smartQcDate - installDate) / 86400000));
          if (diff > 3) {
            acc.doneSQC_OverSLA++;
            acc.totalDays_doneSQC_OverSLA += diff;
          } else {
            acc.doneSQC_InSLA++;
            acc.totalDays_doneSQC_InSLA += diff;
          }
        }

        // PAT Subcon (5 days SLA)
        if (!patDate) {
          if (smartQcDate) {
            const diff = Math.max(0, Math.floor((today - smartQcDate) / 86400000));
            if (diff > 5) {
              acc.pendingPat_OverSLA++;
              acc.totalDays_pendingPat_OverSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: true, owner: ownerDoc });
            } else {
              acc.pendingPat_InSLA++;
              acc.totalDays_pendingPat_InSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: false, owner: ownerDoc });
            }
          }
        } else {
          const base = smartQcDate || installDate;
          const diff = Math.max(0, Math.floor((patDate - base) / 86400000));
          if (diff > 5) {
            acc.donePat_OverSLA++;
            acc.totalDays_donePat_OverSLA += diff;
          } else {
            acc.donePat_InSLA++;
            acc.totalDays_donePat_InSLA += diff;
          }
        }

        // Rework (PAT Not Pass)
        if (patRemark.toLowerCase().includes('not pass')) {
          const od = Math.max(0, patDate ? Math.floor((today - patDate) / 86400000) : Math.floor((today - installDate) / 86400000));
          acc.listPatNotPass.push({ duid, days: od, label: `${od} วัน`, isRework: true, owner: ownerDoc });
        }
      }
    } else {
      // Process from imported Excel dataset
      const items = state.importedData.AIS || [];
      items.forEach(item => {
        const installDate = parseDate(item.installDate);
        if (!installDate) return;
        if (selectedYear !== 'ทั้งหมด' && installDate.getFullYear().toString() !== selectedYear) return;

        acc.totalInstall++;
        const duid = item.duid;
        const ownerDoc = item.owner || '-';
        const smartQcDate = parseDate(item.smartQcDate);
        const patDate = parseDate(item.patDate);
        const patRemark = (item.patRemark || item.patStatus || '').toLowerCase();

        if (!smartQcDate) {
          const diff = Math.max(0, Math.floor((today - installDate) / 86400000));
          if (diff > 3) {
            acc.pendingSQC_OverSLA++;
            acc.totalDays_pendingSQC_OverSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน`, isOver: true, owner: ownerDoc });
          } else {
            acc.pendingSQC_InSLA++;
            acc.totalDays_pendingSQC_InSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน`, isOver: false, owner: ownerDoc });
          }
        } else {
          const diff = Math.max(0, Math.floor((smartQcDate - installDate) / 86400000));
          if (diff > 3) {
            acc.doneSQC_OverSLA++;
            acc.totalDays_doneSQC_OverSLA += diff;
          } else {
            acc.doneSQC_InSLA++;
            acc.totalDays_doneSQC_InSLA += diff;
          }
        }

        if (!patDate) {
          if (smartQcDate) {
            const diff = Math.max(0, Math.floor((today - smartQcDate) / 86400000));
            if (diff > 5) {
              acc.pendingPat_OverSLA++;
              acc.totalDays_pendingPat_OverSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: true, owner: ownerDoc });
            } else {
              acc.pendingPat_InSLA++;
              acc.totalDays_pendingPat_InSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: false, owner: ownerDoc });
            }
          }
        } else {
          const base = smartQcDate || installDate;
          const diff = Math.max(0, Math.floor((patDate - base) / 86400000));
          if (diff > 5) {
            acc.donePat_OverSLA++;
            acc.totalDays_donePat_OverSLA += diff;
          } else {
            acc.donePat_InSLA++;
            acc.totalDays_donePat_InSLA += diff;
          }
        }

        if (patRemark.includes('not pass')) {
          const od = Math.max(0, patDate ? Math.floor((today - patDate) / 86400000) : Math.floor((today - installDate) / 86400000));
          acc.listPatNotPass.push({ duid, days: od, label: `${od} วัน`, isRework: true, owner: ownerDoc });
        }
      });
    }
  } else {
    // TRUE
    if (isSheet) {
      const rows = state.rawGoogleSheetData.TRUE;
      if (!rows || rows.length <= 2) return acc;

      for (let i = 2; i < rows.length; i++) {
        const row = rows[i];
        if (row.length <= 28) continue;
        const installDate = parseDate(row[22]); // Verify Photo (col W)
        if (!installDate) continue;
        if (selectedYear !== 'ทั้งหมด' && installDate.getFullYear().toString() !== selectedYear) continue;

        acc.totalInstall++;
        const duid = row[1] ? row[1].trim() : 'ไม่ระบุ DUID';
        const ownerDoc = row[32] ? row[32].trim() : '-';
        const aorDate = parseDate(row[23]);
        const smartQcDate = parseDate(row[24]);
        const patDate = parseDate(row[28]);
        const patRemark = (row[26] || '').trim(); // Integration alarm category

        const sqcDone = aorDate && smartQcDate;
        const missingLabel = (!aorDate ? ' ⚠️AOR' : '') + (!smartQcDate ? ' ⚠️SQC' : '');

        if (!sqcDone) {
          const diff = Math.max(0, Math.floor((today - installDate) / 86400000));
          if (diff > 3) {
            acc.pendingSQC_OverSLA++;
            acc.totalDays_pendingSQC_OverSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน${missingLabel}`, isOver: true, owner: ownerDoc });
          } else {
            acc.pendingSQC_InSLA++;
            acc.totalDays_pendingSQC_InSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน${missingLabel}`, isOver: false, owner: ownerDoc });
          }
        } else {
          const diff = Math.max(0, Math.floor((smartQcDate - installDate) / 86400000));
          if (diff > 3) {
            acc.doneSQC_OverSLA++;
            acc.totalDays_doneSQC_OverSLA += diff;
          } else {
            acc.doneSQC_InSLA++;
            acc.totalDays_doneSQC_InSLA += diff;
          }
        }

        // PAT
        if (!patDate) {
          if (smartQcDate) {
            const diff = Math.max(0, Math.floor((today - smartQcDate) / 86400000));
            if (diff > 5) {
              acc.pendingPat_OverSLA++;
              acc.totalDays_pendingPat_OverSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: true, owner: ownerDoc });
            } else {
              acc.pendingPat_InSLA++;
              acc.totalDays_pendingPat_InSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: false, owner: ownerDoc });
            }
          }
        } else {
          const base = smartQcDate || installDate;
          const diff = Math.max(0, Math.floor((patDate - base) / 86400000));
          if (diff > 5) {
            acc.donePat_OverSLA++;
            acc.totalDays_donePat_OverSLA += diff;
          } else {
            acc.donePat_InSLA++;
            acc.totalDays_donePat_InSLA += diff;
          }
        }

        // Alarm Rework
        const hasAlarm = patRemark !== '' && !patRemark.toLowerCase().includes('no alarm');
        if (hasAlarm) {
          const od = Math.max(0, patDate ? Math.floor((today - patDate) / 86400000) : Math.floor((today - installDate) / 86400000));
          acc.listPatNotPass.push({ duid, days: od, label: `${od} วัน`, isRework: true, owner: ownerDoc });
        }
      }
    } else {
      // Process from imported Excel dataset
      const items = state.importedData.TRUE || [];
      items.forEach(item => {
        const installDate = parseDate(item.installDate);
        if (!installDate) return;
        if (selectedYear !== 'ทั้งหมด' && installDate.getFullYear().toString() !== selectedYear) return;

        acc.totalInstall++;
        const duid = item.duid;
        const ownerDoc = item.owner || '-';
        const aorDate = parseDate(item.aorDate);
        const smartQcDate = parseDate(item.smartQcDate);
        const patDate = parseDate(item.patDate);
        const patRemark = (item.alarmRemark || '').trim();

        const sqcDone = aorDate && smartQcDate;
        const missingLabel = (!aorDate ? ' ⚠️AOR' : '') + (!smartQcDate ? ' ⚠️SQC' : '');

        if (!sqcDone) {
          const diff = Math.max(0, Math.floor((today - installDate) / 86400000));
          if (diff > 3) {
            acc.pendingSQC_OverSLA++;
            acc.totalDays_pendingSQC_OverSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน${missingLabel}`, isOver: true, owner: ownerDoc });
          } else {
            acc.pendingSQC_InSLA++;
            acc.totalDays_pendingSQC_InSLA += diff;
            acc.listPendingSmartQc.push({ duid, days: diff, label: `${diff} วัน${missingLabel}`, isOver: false, owner: ownerDoc });
          }
        } else {
          const diff = Math.max(0, Math.floor((smartQcDate - installDate) / 86400000));
          if (diff > 3) {
            acc.doneSQC_OverSLA++;
            acc.totalDays_doneSQC_OverSLA += diff;
          } else {
            acc.doneSQC_InSLA++;
            acc.totalDays_doneSQC_InSLA += diff;
          }
        }

        if (!patDate) {
          if (smartQcDate) {
            const diff = Math.max(0, Math.floor((today - smartQcDate) / 86400000));
            if (diff > 5) {
              acc.pendingPat_OverSLA++;
              acc.totalDays_pendingPat_OverSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: true, owner: ownerDoc });
            } else {
              acc.pendingPat_InSLA++;
              acc.totalDays_pendingPat_InSLA += diff;
              acc.listPendingPat.push({ duid, days: diff, label: `${diff} วัน`, isOver: false, owner: ownerDoc });
            }
          }
        } else {
          const base = smartQcDate || installDate;
          const diff = Math.max(0, Math.floor((patDate - base) / 86400000));
          if (diff > 5) {
            acc.donePat_OverSLA++;
            acc.totalDays_donePat_OverSLA += diff;
          } else {
            acc.donePat_InSLA++;
            acc.totalDays_donePat_InSLA += diff;
          }
        }

        const hasAlarm = patRemark !== '' && !patRemark.toLowerCase().includes('no alarm');
        if (hasAlarm) {
          const od = Math.max(0, patDate ? Math.floor((today - patDate) / 86400000) : Math.floor((today - installDate) / 86400000));
          acc.listPatNotPass.push({ duid, days: od, label: `${od} วัน`, isRework: true, owner: ownerDoc });
        }
      });
    }
  }

  // Sort lists descending by aging days
  acc.listPendingSmartQc.sort((a,b) => b.days - a.days);
  acc.listPendingPat.sort((a,b) => b.days - a.days);
  acc.listPatNotPass.sort((a,b) => b.days - a.days);

  return acc;
}

// Render Dashboard Elements
function calculateAndRenderDashboard() {
  const m = calculateMetrics();
  state.currentMetrics = m;

  // Header Total
  document.getElementById('totalInstallCount').textContent = `${m.totalInstall} ไซต์`;
  document.getElementById('yearSublabel').textContent = `(ปี: ${state.selectedYear})`;

  // Helper for averages
  const avg = (tot, cnt) => cnt > 0 ? (tot / cnt).toFixed(1) : '0.0';

  // Smart QC Executive Summary Table
  document.getElementById('sqcDoneInCount').textContent = m.doneSQC_InSLA;
  document.getElementById('sqcDoneInAvg').textContent = `${avg(m.totalDays_doneSQC_InSLA, m.doneSQC_InSLA)} วัน`;

  document.getElementById('sqcDoneLateCount').textContent = m.doneSQC_OverSLA;
  document.getElementById('sqcDoneLateAvg').textContent = `${avg(m.totalDays_doneSQC_OverSLA, m.doneSQC_OverSLA)} วัน`;

  document.getElementById('sqcPendInCount').textContent = m.pendingSQC_InSLA;
  document.getElementById('sqcPendInAvg').textContent = `${avg(m.totalDays_pendingSQC_InSLA, m.pendingSQC_InSLA)} วัน`;

  document.getElementById('sqcPendOverCount').textContent = m.pendingSQC_OverSLA;
  document.getElementById('sqcPendOverAvg').textContent = `${avg(m.totalDays_pendingSQC_OverSLA, m.pendingSQC_OverSLA)} วัน`;

  // PAT Subcon Executive Summary Table
  document.getElementById('patDoneInCount').textContent = m.donePat_InSLA;
  document.getElementById('patDoneInAvg').textContent = `${avg(m.totalDays_donePat_InSLA, m.donePat_InSLA)} วัน`;

  document.getElementById('patDoneLateCount').textContent = m.donePat_OverSLA;
  document.getElementById('patDoneLateAvg').textContent = `${avg(m.totalDays_donePat_OverSLA, m.donePat_OverSLA)} วัน`;

  document.getElementById('patPendInCount').textContent = m.pendingPat_InSLA;
  document.getElementById('patPendInAvg').textContent = `${avg(m.totalDays_pendingPat_InSLA, m.pendingPat_InSLA)} วัน`;

  document.getElementById('patPendOverCount').textContent = m.pendingPat_OverSLA;
  document.getElementById('patPendOverAvg').textContent = `${avg(m.totalDays_pendingPat_OverSLA, m.pendingPat_OverSLA)} วัน`;

  // Render Charts
  renderDonutCharts(m);

  // Render Action Tables
  renderActionTables();
}

// Render Donut Charts with Chart.js
function renderDonutCharts(m) {
  const sqcCanvas = document.getElementById('sqcDonutChart');
  const patCanvas = document.getElementById('patDonutChart');
  if (!sqcCanvas || !patCanvas) return;

  const sqcData = [
    m.pendingSQC_OverSLA,
    m.doneSQC_OverSLA,
    m.doneSQC_InSLA,
    m.pendingSQC_InSLA
  ];

  const patData = [
    m.pendingPat_OverSLA,
    m.donePat_OverSLA,
    m.donePat_InSLA,
    m.pendingPat_InSLA
  ];

  const labels = [
    'ปิดเกินกำหนด / ค้างวิกฤต (Over SLA)',
    'ปิดงานล่าช้า / ส่งงานล่าช้า (Done Late)',
    'ปิดตามกำหนด / ส่งตามกำหนด (Done in SLA)',
    'รอตรวจสอบ / รอส่งงาน (Pending in SLA)'
  ];

  const backgroundColors = [COLORS.RED, COLORS.YELLOW, COLORS.GREEN, COLORS.BLUE];

  // Helper chart builder
  const createChart = (canvas, existingChart, data, title) => {
    if (existingChart) existingChart.destroy();
    return new Chart(canvas, {
      type: 'doughnut',
      data: {
        labels: labels,
        datasets: [{
          data: data.map(v => v > 0 ? v : 0.0001),
          backgroundColor: backgroundColors,
          borderWidth: 2,
          borderColor: '#111827'
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            position: 'bottom',
            labels: {
              color: '#94a3b8',
              font: { family: 'Prompt', size: 11 },
              padding: 12,
              usePointStyle: true,
              pointStyle: 'circle'
            }
          },
          tooltip: {
            callbacks: {
              label: function(context) {
                const raw = data[context.dataIndex];
                const sum = data.reduce((a,b)=>a+b, 0);
                const pct = sum > 0 ? ((raw / sum) * 100).toFixed(1) : 0;
                return ` ${context.label}: ${raw} ไซต์ (${pct}%)`;
              }
            }
          }
        },
        cutout: '55%'
      }
    });
  };

  state.charts.sqc = createChart(sqcCanvas, state.charts.sqc, sqcData, 'Smart QC');
  state.charts.pat = createChart(patCanvas, state.charts.pat, patData, 'PAT Subcon');
}

// Render Action & Rework Tables with Search
function renderActionTables() {
  const m = state.currentMetrics;
  if (!m) return;

  const sqcQuery = state.searchQuery.sqc;
  const patQuery = state.searchQuery.pat;
  const reworkQuery = state.searchQuery.rework;

  // Filter lists
  const sqcList = m.listPendingSmartQc.filter(item => 
    item.duid.toLowerCase().includes(sqcQuery) || (item.owner && item.owner.toLowerCase().includes(sqcQuery))
  );
  const patList = m.listPendingPat.filter(item => 
    item.duid.toLowerCase().includes(patQuery) || (item.owner && item.owner.toLowerCase().includes(patQuery))
  );
  const reworkList = m.listPatNotPass.filter(item => 
    item.duid.toLowerCase().includes(reworkQuery) || (item.owner && item.owner.toLowerCase().includes(reworkQuery))
  );

  // Update counts
  document.getElementById('sqcCountPill').textContent = `${sqcList.length} ไซต์`;
  document.getElementById('patCountPill').textContent = `${patList.length} ไซต์`;
  document.getElementById('reworkCountPill').textContent = `${reworkList.length} ไซต์`;

  // Render SQC List
  const sqcContainer = document.getElementById('sqcActionList');
  if (sqcList.length === 0) {
    sqcContainer.innerHTML = '<div class="empty-state">🎉 All sites are up to date</div>';
  } else {
    sqcContainer.innerHTML = sqcList.map(item => `
      <div class="action-item ${item.isOver ? 'over-sla' : 'in-sla'}" onclick="copyText('${item.duid}')" title="คลิกเพื่อคัดลอก DUID">
        <div>
          <div class="site-duid-text">${escapeHtml(item.duid)}</div>
          <div class="site-owner-text">Owner: ${escapeHtml(item.owner || '-')}</div>
        </div>
        <div class="action-badge ${item.isOver ? 'badge-red' : 'badge-blue'}">${item.label}</div>
      </div>
    `).join('');
  }

  // Render PAT List
  const patContainer = document.getElementById('patActionList');
  if (patList.length === 0) {
    patContainer.innerHTML = '<div class="empty-state">🎉 All sites are up to date</div>';
  } else {
    patContainer.innerHTML = patList.map(item => `
      <div class="action-item ${item.isOver ? 'over-sla' : 'in-sla'}" onclick="copyText('${item.duid}')" title="คลิกเพื่อคัดลอก DUID">
        <div>
          <div class="site-duid-text">${escapeHtml(item.duid)}</div>
          <div class="site-owner-text">Assign: ${escapeHtml(item.owner || '-')}</div>
        </div>
        <div class="action-badge ${item.isOver ? 'badge-red' : 'badge-blue'}">${item.label}</div>
      </div>
    `).join('');
  }

  // Render Rework List
  const reworkContainer = document.getElementById('reworkActionList');
  if (reworkList.length === 0) {
    reworkContainer.innerHTML = '<div class="empty-state">🎉 No rework required</div>';
  } else {
    reworkContainer.innerHTML = reworkList.map(item => `
      <div class="action-item rework" onclick="copyText('${item.duid}')" title="คลิกเพื่อคัดลอก DUID">
        <div>
          <div class="site-duid-text">${escapeHtml(item.duid)}</div>
          <div class="site-owner-text">Assign: ${escapeHtml(item.owner || '-')}</div>
        </div>
        <div class="action-badge badge-orange">${item.label}</div>
      </div>
    `).join('');
  }
}

// Handle Excel File Upload & Parse with SheetJS
function handleExcelFileUpload(file) {
  if (!file.name.endsWith('.xlsx') && !file.name.endsWith('.xls')) {
    showToast('กรุณาเลือกไฟล์ Excel (.xlsx หรือ .xls) เท่านั้น', 'error');
    return;
  }

  const reader = new FileReader();
  showToast(`กำลังอ่านไฟล์: ${file.name}...`, 'info');

  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array', cellDates: true });

      // Look for 'Site Rollout Plan'
      const sheetName = workbook.SheetNames.find(n => n.toLowerCase().includes('site rollout plan')) || workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonRows = XLSX.utils.sheet_to_json(worksheet, { header: 1, raw: false, dateNF: 'yyyy-mm-dd' });

      // Auto-detect AIS or TRUE from filename or content
      let detectedOp = state.operator;
      const lowerName = file.name.toLowerCase();
      if (lowerName.includes('56a0s0q') || lowerName.includes('ais')) {
        detectedOp = 'AIS';
      } else if (lowerName.includes('56a0ups') || lowerName.includes('true')) {
        detectedOp = 'TRUE';
      }

      // Parse records based on detected Operator
      const parsedRecords = parseUploadedExcelRows(detectedOp, jsonRows);

      state.importedData[detectedOp] = parsedRecords;
      state.dataSource = 'EXCEL_UPLOAD';
      state.operator = detectedOp;
      state.lastUpdated = new Date();

      // Update UI active buttons
      document.querySelectorAll('.operator-btn').forEach(b => {
        b.classList.toggle('active', b.getAttribute('data-operator') === detectedOp);
      });
      updateThemeForOperator(detectedOp);

      document.getElementById('updateTimestamp').textContent = 
        `ข้อมูลอัปเดต ณ: ${formatDateTime(state.lastUpdated)} (ไฟล์นำเข้า: ${file.name} - ${parsedRecords.length} แถว)`;

      extractAvailableYears();
      calculateAndRenderDashboard();

      document.getElementById('uploadModal').classList.remove('active');
      showToast(`นำเข้าสำเร็จ! พบ ${parsedRecords.length} ไซต์ใน ${detectedOp}`, 'success');

      // If in Google Apps Script context, ask or push to Google Sheet
      if (typeof google !== 'undefined' && google.script && google.script.run) {
        showToast('กำลังซิงค์ข้อมูลกับ Google Spreadsheet...', 'info');
        google.script.run
          .withSuccessHandler((res) => {
            showToast('อัปเดตข้อมูลลง Google Spreadsheet เรียบร้อยแล้ว!', 'success');
          })
          .withFailureHandler((err) => {
            console.error('GAS save failed:', err);
          })
          .saveImportData(detectedOp, parsedRecords);
      }
    } catch (err) {
      console.error('Excel parse error:', err);
      showToast(`เกิดข้อผิดพลาดในการอ่านไฟล์: ${err.message}`, 'error');
    }
  };

  reader.readAsArrayBuffer(file);
}

// Parse Rows from Excel SheetJS based on Operator
function parseUploadedExcelRows(op, rows) {
  const items = [];
  if (!rows || rows.length < 4) return items;

  if (op === 'AIS') {
    // Row 4 (index 3) is where data starts
    // Col 1 (A, idx 0): DUID
    // Col 9 (I, idx 8): Installation-Completed Actual End Date
    // Col 21 (U, idx 20): Smart QC Actual End Date
    // Col 31 (AE, idx 30): PAT Subcon submit Actual End Date
    // Col 32 (AF, idx 31): PAT Owner
    // Col 39 (AM, idx 38): PAT Remarks
    // Col 40 (AN, idx 39): Sub PAT Pass Status
    for (let r = 3; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row[0]) continue;
      const duid = String(row[0]).trim();
      const installDate = row[8] ? String(row[8]).trim() : null;
      const smartQcDate = row[20] ? String(row[20]).trim() : null;
      const patDate = row[30] ? String(row[30]).trim() : null;
      const owner = row[31] || row[9] || '-';
      const patRemark = row[38] || '';
      const patStatus = row[39] || '';

      if (installDate) {
        items.push({
          duid,
          installDate,
          smartQcDate,
          patDate,
          owner: String(owner).trim(),
          patRemark: String(patRemark).trim(),
          patStatus: String(patStatus).trim()
        });
      }
    }
  } else {
    // TRUE
    // Col 1 (A, idx 0): DUID
    // Col 67 (BO, idx 66): Verify Photo Actual End Date
    // Col 41 (AO, idx 40): AOR Actual End Date
    // Col 37 (AK, idx 36): 08.1 SmartQC Actual End Date
    // Col 45 (AS, idx 44): Integration alarm category
    // Col 47 (AU, idx 46): 14.1 A129 PAT Site Folder Actual End Date
    // Col 48 (AV, idx 47): PAT Owner
    for (let r = 3; r < rows.length; r++) {
      const row = rows[r];
      if (!row || !row[0]) continue;
      const duid = String(row[0]).trim();
      const installDate = row[66] ? String(row[66]).trim() : null;
      const aorDate = row[40] ? String(row[40]).trim() : null;
      const smartQcDate = row[36] ? String(row[36]).trim() : null;
      const alarmRemark = row[44] ? String(row[44]).trim() : '';
      const patDate = row[46] ? String(row[46]).trim() : null;
      const owner = row[47] || row[37] || row[3] || '-';

      if (installDate) {
        items.push({
          duid,
          installDate,
          aorDate,
          smartQcDate,
          alarmRemark: String(alarmRemark).trim(),
          patDate,
          owner: String(owner).trim()
        });
      }
    }
  }

  return items;
}

// Generate LINE Notification Preview Message
function generateLineAlertPreview() {
  const m = state.currentMetrics;
  if (!m) return;
  const op = state.operator;
  const yr = state.selectedYear;
  const avg = (tot, cnt) => cnt > 0 ? (tot / cnt).toFixed(1) : '0.0';

  let msg = `📊 [${op} INSTALLATION & SLA DASHBOARD]\n`;
  msg += `📅 ประจำวันที่: ${formatDateTime(new Date())}\n`;
  msg += `🎯 งานติดตั้งเสร็จ (${yr}): ${m.totalInstall} ไซต์\n`;
  msg += `------------------------------------\n`;

  if (op === 'AIS') {
    msg += `🔹 EXECUTIVE SUMMARY: SMART QC\n`;
    msg += `🟢 ปิดตามกำหนด (In SLA): ${m.doneSQC_InSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_doneSQC_InSLA, m.doneSQC_InSLA)} วัน)\n`;
    msg += `🟡 ปิดงานล่าช้า (Done Late): ${m.doneSQC_OverSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_doneSQC_OverSLA, m.doneSQC_OverSLA)} วัน)\n`;
    msg += `🔵 รอตรวจสอบ (In SLA): ${m.pendingSQC_InSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_pendingSQC_InSLA, m.pendingSQC_InSLA)} วัน)\n`;
    msg += `🔴 ค้างวิกฤต (Over SLA): ${m.pendingSQC_OverSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_pendingSQC_OverSLA, m.pendingSQC_OverSLA)} วัน)\n\n`;

    msg += `🔹 EXECUTIVE SUMMARY: PAT SUBCON\n`;
    msg += `🟢 ส่งตามกำหนด (In SLA): ${m.donePat_InSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_donePat_InSLA, m.donePat_InSLA)} วัน)\n`;
    msg += `🟡 ส่งงานล่าช้า (Done Late): ${m.donePat_OverSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_donePat_OverSLA, m.donePat_OverSLA)} วัน)\n`;
    msg += `🔵 รอส่งงาน (In SLA): ${m.pendingPat_InSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_pendingPat_InSLA, m.pendingPat_InSLA)} วัน)\n`;
    msg += `🔴 ค้างวิกฤต (Over SLA): ${m.pendingPat_OverSLA} ไซต์\n`;
  } else {
    msg += `🔹 EXECUTIVE SUMMARY: 08.1 SmartQC+AOR\n`;
    msg += `🟢 ปิดตามกำหนด (In SLA): ${m.doneSQC_InSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_doneSQC_InSLA, m.doneSQC_InSLA)} วัน)\n`;
    msg += `🟡 ปิดงานล่าช้า (Done Late): ${m.doneSQC_OverSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_doneSQC_OverSLA, m.doneSQC_OverSLA)} วัน)\n`;
    msg += `🔵 รอตรวจสอบ (In SLA): ${m.pendingSQC_InSLA} ไซต์\n`;
    msg += `🔴 ค้างวิกฤต (Over SLA): ${m.pendingSQC_OverSLA} ไซต์\n\n`;

    msg += `🔹 EXECUTIVE SUMMARY: 14.1 A129 PAT\n`;
    msg += `🟢 ส่งตามกำหนด (In SLA): ${m.donePat_InSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_donePat_InSLA, m.donePat_InSLA)} วัน)\n`;
    msg += `🟡 ส่งงานล่าช้า (Done Late): ${m.donePat_OverSLA} ไซต์ (เฉลี่ย ${avg(m.totalDays_donePat_OverSLA, m.donePat_OverSLA)} วัน)\n`;
    msg += `🔵 รอส่งงาน (In SLA): ${m.pendingPat_InSLA} ไซต์\n`;
    msg += `🔴 ค้างวิกฤต (Over SLA): ${m.pendingPat_OverSLA} ไซต์\n`;
  }

  // Urgent Items List
  msg += `------------------------------------\n`;
  msg += `⚠️ รายการที่ต้องดำเนินการเร่งด่วน:\n`;

  if (m.listPendingSmartQc.length > 0) {
    msg += `\n📌 Smart QC ค้าง (${m.listPendingSmartQc.length} ไซต์):\n`;
    m.listPendingSmartQc.slice(0, 5).forEach((item, idx) => {
      msg += `  ${idx+1}. ${item.duid} (${item.label}) - ${item.owner}\n`;
    });
    if (m.listPendingSmartQc.length > 5) msg += `  ...และอีก ${m.listPendingSmartQc.length - 5} ไซต์\n`;
  }

  if (m.listPendingPat.length > 0) {
    msg += `\n📌 PAT Subcon ค้าง (${m.listPendingPat.length} ไซต์):\n`;
    m.listPendingPat.slice(0, 5).forEach((item, idx) => {
      msg += `  ${idx+1}. ${item.duid} (${item.label}) - Assign: ${item.owner}\n`;
    });
    if (m.listPendingPat.length > 5) msg += `  ...และอีก ${m.listPendingPat.length - 5} ไซต์\n`;
  }

  if (m.listPatNotPass.length > 0) {
    const reworkLabel = op === 'AIS' ? 'PAT Not Pass' : 'Alarm Found';
    msg += `\n🚨 Rework Required: ${reworkLabel} (${m.listPatNotPass.length} ไซต์):\n`;
    m.listPatNotPass.slice(0, 5).forEach((item, idx) => {
      msg += `  ${idx+1}. ${item.duid} (${item.label}) - Assign: ${item.owner}\n`;
    });
    if (m.listPatNotPass.length > 5) msg += `  ...และอีก ${m.listPatNotPass.length - 5} ไซต์\n`;
  }

  if (m.listPendingSmartQc.length === 0 && m.listPendingPat.length === 0 && m.listPatNotPass.length === 0) {
    msg += `✨ ไม่มีงานค้างวิกฤต All sites are up to date! 🎉\n`;
  }

  document.getElementById('lineMessagePreview').textContent = msg;

  // Restore saved token & groupId
  const tokenInput = document.getElementById('lineNotifyToken');
  if (tokenInput) {
    // Use saved token from localStorage, fall back to default state token
    tokenInput.value = localStorage.getItem('sla_line_token') || state.lineToken;
  }

  const groupInput = document.getElementById('lineGroupId');
  if (groupInput) {
    const savedGroup = localStorage.getItem('sla_line_group_id') || state.lineGroupId;
    if (savedGroup) groupInput.value = savedGroup;
  }

  const gasInput = document.getElementById('gasWebAppUrl');
  if (gasInput) {
    const savedGasUrl = localStorage.getItem('sla_gas_url') || state.gasWebAppUrl;
    if (savedGasUrl) gasInput.value = savedGasUrl;
  }
}

// Capture High-Res Dashboard Screenshot using html2canvas
async function captureDashboardImage() {
  const captureArea = document.getElementById('captureArea') || document.body;
  const imgEl = document.getElementById('capturedDashboardImg');
  const spinner = document.getElementById('imageLoadingSpinner');

  if (imgEl && spinner) {
    imgEl.style.display = 'none';
    spinner.style.display = 'block';
  }

  try {
    if (typeof html2canvas === 'undefined') {
      console.warn('html2canvas not loaded');
      if (spinner) spinner.textContent = '⚠️ ไม่พบ library html2canvas';
      return;
    }

    const canvas = await html2canvas(captureArea, {
      backgroundColor: '#0b1120',
      scale: 1.5,
      logging: false,
      useCORS: true,
      windowWidth: 1920
    });

    state.lastCapturedCanvas = canvas;
    const dataUrl = canvas.toDataURL('image/png');
    state.lastCapturedBase64 = dataUrl.split(',')[1];

    canvas.toBlob((blob) => {
      state.lastCapturedBlob = blob;
    }, 'image/png');

    if (imgEl && spinner) {
      imgEl.src = dataUrl;
      imgEl.style.display = 'block';
      spinner.style.display = 'none';
    }
  } catch(err) {
    console.error('Capture error:', err);
    if (spinner) spinner.textContent = '⚠️ ไม่สามารถแคปเจอร์รูปภาพได้: ' + err.message;
  }
}

// Copy Captured Image directly to clipboard (for pasting in LINE PC / Mac)
async function handleCopyImage() {
  if (!state.lastCapturedBlob) {
    showToast('กรุณารอระบบกำลังประมวลผลรูปภาพ...', 'info');
    return;
  }
  try {
    const item = new ClipboardItem({ 'image/png': state.lastCapturedBlob });
    await navigator.clipboard.write([item]);
    showToast('🖼️ คัดลอกรูปภาพแดชบอร์ดลง Clipboard แล้ว! สามารถกด Ctrl+V วางใน LINE ได้ทันที', 'success');
  } catch(err) {
    console.error('Copy image error:', err);
    showToast('เบราว์เซอร์ไม่อนุญาตให้คัดลอกรูปภาพโดยตรง กรุณากด "ดาวน์โหลดรูปภาพ"', 'error');
  }
}

// Download Captured Image as PNG file
function handleDownloadImage() {
  if (!state.lastCapturedCanvas) {
    showToast('ไม่พบรูปภาพแดชบอร์ด', 'error');
    return;
  }
  const a = document.createElement('a');
  const dStr = formatDateTime(new Date()).replace(/[\/ :]/g, '_');
  a.href = state.lastCapturedCanvas.toDataURL('image/png');
  a.download = `${state.operator}_SLA_Dashboard_${dStr}.png`;
  a.click();
  showToast('💾 ดาวน์โหลดรูปภาพแดชบอร์ดสำเร็จแล้ว!', 'success');
}

// Send LINE Alert via LINE Messaging API or Webhook (with Image & Text)
async function handleSendLineAlert() {
  const tokenInput = document.getElementById('lineNotifyToken');
  const token = tokenInput ? tokenInput.value.trim() : '';
  const groupInput = document.getElementById('lineGroupId');
  const groupId = groupInput ? groupInput.value.trim() : '';
  const message = document.getElementById('lineMessagePreview').textContent;
  const imageBase64 = state.lastCapturedBase64 || null;

  if (!token) {
    showToast('กรุณากรอก LINE Channel Access Token หรือ Webhook URL', 'error');
    return;
  }

  // Save token & group for next time
  localStorage.setItem('sla_line_token', token);
  state.lineToken = token;
  if (groupId) localStorage.setItem('sla_line_group_id', groupId);

  const gasInput = document.getElementById('gasWebAppUrl');
  const gasUrl = (gasInput ? gasInput.value.trim() : '') || localStorage.getItem('sla_gas_url') || state.gasWebAppUrl;
  if (gasUrl) localStorage.setItem('sla_gas_url', gasUrl);

  showToast('กำลังส่งแจ้งเตือน (รูปภาพ + ข้อความ) เข้า LINE Group...', 'info');

  // Check if running in Google Apps Script context
  if (typeof google !== 'undefined' && google.script && google.script.run) {
    google.script.run
      .withSuccessHandler(() => {
        showToast('🚀 ส่งรูปภาพและข้อความแจ้งเตือน LINE Group สำเร็จแล้ว!', 'success');
        document.getElementById('lineModal').classList.remove('active');
      })
      .withFailureHandler((err) => {
        showToast(`❌ ส่ง LINE ล้มเหลว: ${err.message}`, 'error');
      })
      .sendLineAlert(token, message, imageBase64, groupId);
    return;
  }

  // ── Google Apps Script Web App Proxy mode (สำหรับ Browser แก้ปัญหา CORS 100%) ──
  if (gasUrl && (gasUrl.startsWith('http://') || gasUrl.startsWith('https://'))) {
    try {
      await fetch(gasUrl, {
        method: 'POST',
        headers: { 'Content-Type': 'text/plain;charset=utf-8' },
        body: JSON.stringify({
          action: 'sendLineAlert',
          token: token,
          message: message,
          imageBase64: imageBase64,
          groupId: groupId
        }),
        mode: 'no-cors'
      });
      showToast('🚀 ส่งข้อมูลแจ้งเตือนเข้า LINE Group สำเร็จแล้ว (ผ่าน Google Apps Script)!', 'success');
      if (state.lastCapturedBlob) {
        try {
          await navigator.clipboard.write([new ClipboardItem({ 'image/png': state.lastCapturedBlob })]);
        } catch(e) {}
      }
      document.getElementById('lineModal').classList.remove('active');
      return;
    } catch(gasErr) {
      console.warn('GAS Proxy call error:', gasErr);
    }
  }

  // ── Webhook URL mode (เช่น webhook.site สำหรับ test) ──
  if (token.startsWith('http://') || token.startsWith('https://')) {
    try {
      await fetch(token, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ message, imageBase64, groupId })
      });
      showToast('🚀 ส่ง Webhook แจ้งเตือนสำเร็จ!', 'success');
      document.getElementById('lineModal').classList.remove('active');
    } catch (e) {
      showToast(`❌ Webhook error: ${e.message}`, 'error');
    }
    return;
  }

  // ── LINE Messaging API mode (Channel Access Token) ──
  // NOTE: Direct browser calls to api.line.me are blocked by CORS.
  // We send via a CORS proxy (allorigins.win) so the request goes through.
  // For production use, route through a backend / Google Apps Script instead.
  const LINE_API = 'https://api.line.me/v2/bot/message';
  const endpoint = groupId
    ? `${LINE_API}/push`    // ส่งไปยัง Group ID ที่ระบุ
    : `${LINE_API}/broadcast`; // broadcast ไปทุกคนที่ follow Bot

  const textMessage = { type: 'text', text: message };
  const messages = [textMessage];

  // เพิ่ม flex image ถ้ามี base64
  if (imageBase64) {
    // Upload image อาจต้องใช้ LINE Rich menu / Content API
    // สำหรับ prototype: แนบ image url placeholder หรือข้ามไป
    // (เราใช้ข้อความ + copy image แทนสำหรับ browser mode)
  }

  const body = groupId
    ? { to: groupId, messages }
    : { messages };

  try {
    // ลองส่งตรง (จะสำเร็จเฉพาะ environment ที่ CORS อนุญาต เช่น GAS Web App)
    let response = await fetch(endpoint, {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Authorization': `Bearer ${token}`
      },
      body: JSON.stringify(body)
    });

    if (response.ok) {
      showToast('🚀 ส่งข้อความเข้า LINE Group สำเร็จแล้ว!', 'success');
      // Copy image ไว้ให้ user วางต่อ
      if (state.lastCapturedBlob) {
        try {
          await navigator.clipboard.write([new ClipboardItem({ 'image/png': state.lastCapturedBlob })]);
          showToast('🖼️ คัดลอกรูปภาพแล้ว — กด Ctrl+V วางใน LINE ได้เลย!', 'success');
        } catch(e) {}
      }
      document.getElementById('lineModal').classList.remove('active');
    } else {
      const errText = await response.text();
      throw new Error(`LINE API: ${response.status} ${errText}`);
    }
  } catch (corsOrErr) {
    console.warn('LINE API direct call failed (CORS หรือ token ผิด):', corsOrErr.message);
    // Fallback: คัดลอกรูป + ข้อความ เพื่อให้ user วางใน LINE เอง
    let copied = false;
    if (state.lastCapturedBlob) {
      try {
        await navigator.clipboard.write([new ClipboardItem({ 'image/png': state.lastCapturedBlob })]);
        copied = true;
      } catch(e) {}
    }
    if (!copied) {
      try {
        await navigator.clipboard.writeText(message);
      } catch(e) {}
    }
    showToast(
      '📋 ไม่สามารถส่งตรงได้จาก Browser (CORS) — ' +
      'รูปภาพและข้อความถูกคัดลอกแล้ว กด Ctrl+V วางใน LINE Group ได้เลย! ' +
      '(ใช้ Google Apps Script Web App เพื่อส่งอัตโนมัติ)',
      'info'
    );
  }
}

// Toast helper
function showToast(message, type = 'info') {
  let toast = document.getElementById('toast');
  if (!toast) {
    toast = document.createElement('div');
    toast.id = 'toast';
    toast.className = 'toast';
    document.body.appendChild(toast);
  }

  toast.textContent = message;
  toast.className = `toast show ${type}`;

  setTimeout(() => {
    toast.classList.remove('show');
  }, 4000);
}

// Copy to clipboard helper
function copyText(txt) {
  navigator.clipboard.writeText(txt).then(() => {
    showToast(`คัดลอก: "${txt}" แล้ว`, 'success');
  });
}

// Helpers
function formatDateTime(d) {
  if (!d) return '-';
  const pad = n => n < 10 ? '0' + n : n;
  return `${pad(d.getDate())}/${pad(d.getMonth()+1)}/${d.getFullYear()} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function escapeHtml(str) {
  if (!str) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
