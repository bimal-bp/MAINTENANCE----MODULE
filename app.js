/* ===================================================================
   SGX MINERALS - MAINTENANCE MODULE ERP
   Core Application Logic
   =================================================================== */

// ===================== FLEET MASTER =====================
const FLEET = [
  {id:'AP39UQ0095', category:'Tipper'},
  {id:'AP39UQ0097', category:'Tipper'},
  {id:'AP39UW9880', category:'Tipper'},
  {id:'AP39UW9881', category:'Tipper'},
  {id:'AP39UY4651', category:'Tipper'},
  {id:'AP39UY4652', category:'Tipper'},
  {id:'AP39WC0926', category:'Tipper'},
  {id:'AP39WC0927', category:'Tipper'},
  {id:'AP39WC0928', category:'Tipper'},
  {id:'AP39WF2052', category:'Tipper'},
  {id:'AP39WF2057', category:'Tipper'},
  {id:'AP39WF2058', category:'Tipper'},
  {id:'EX-01', category:'Excavator'},
  {id:'EX-02', category:'Excavator'},
  {id:'LD-01', category:'Loader'},
  {id:'LD-02', category:'Loader'}
];

const SERVICE_INTERVALS = {'500':500, '1000':1000, '2000':2000, '4000':4000};

// ===================== STORAGE LAYER =====================
const DB_KEYS = {
  dailyLogs: 'sgx_mnt_dailylogs',
  hardware: 'sgx_mnt_hardware',
  lubricants: 'sgx_mnt_lubricants',
  workshop: 'sgx_mnt_workshop',
  materialIssues: 'sgx_mnt_materialissues',
  serviceRecords: 'sgx_mnt_servicerecords',
  breakdowns: 'sgx_mnt_breakdowns'
};

function dbGet(key) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : [];
  } catch(e) { console.error('DB read error', key, e); return []; }
}
function dbSet(key, data) {
  try { localStorage.setItem(key, JSON.stringify(data)); }
  catch(e) { console.error('DB write error', key, e); alert('Storage error: ' + e.message); }
}
function uid() { return 'id_' + Date.now() + '_' + Math.floor(Math.random()*100000); }

// ===================== UTILITIES =====================
function fmtCurrency(v) {
  v = Number(v) || 0;
  return '₹' + v.toLocaleString('en-IN', {maximumFractionDigits: 0});
}
function fmtCurrency2(v) {
  v = Number(v) || 0;
  return '₹' + v.toLocaleString('en-IN', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}
function fmtNum(v, dec) {
  v = Number(v) || 0;
  return v.toLocaleString('en-IN', {minimumFractionDigits: dec||0, maximumFractionDigits: dec||2});
}
function todayStr() {
  const d = new Date();
  return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
}
function escapeHtml(str) {
  if (str === null || str === undefined) return '';
  return String(str).replace(/[&<>"']/g, m => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]));
}
function showToast(msg, isError) {
  const toastEl = document.getElementById('appToast');
  const body = document.getElementById('appToastBody');
  body.textContent = msg;
  toastEl.classList.remove('text-bg-success','text-bg-danger');
  toastEl.classList.add(isError ? 'text-bg-danger' : 'text-bg-success');
  const toast = bootstrap.Toast.getOrCreateInstance(toastEl, {delay: 2500});
  toast.show();
}
function timeToMinutes(t) {
  if (!t) return null;
  const parts = t.split(':');
  return (parseInt(parts[0],10) * 60) + parseInt(parts[1],10);
}
function diffHoursFromTimes(fromT, toT) {
  const fm = timeToMinutes(fromT), tm = timeToMinutes(toT);
  if (fm === null || tm === null) return 0;
  let diff = tm - fm;
  if (diff < 0) diff += 24*60; // crosses midnight
  return diff / 60;
}
function diffHoursFromDatetime(start, end) {
  if (!start || !end) return 0;
  const s = new Date(start), e = new Date(end);
  const diffMs = e - s;
  if (diffMs <= 0) return 0;
  return diffMs / (1000*60*60);
}

// ===================== NAVIGATION =====================
const PAGE_TITLES = {
  dashboard: ['Dashboard', 'Fleet maintenance overview & KPIs'],
  dailylog: ['Daily Log Entry', 'Crusher / mining fleet daily operations log'],
  inventory: ['Stores & Inventory', 'Hardware, lubricants & workshop consumables'],
  materialissue: ['Material Issue', 'Issue materials to equipment & auto-deduct stock'],
  service: ['Service Schedule', 'Preventive maintenance schedule & due alerts'],
  breakdown: ['Breakdown History', 'Equipment breakdown & repair log'],
  costanalysis: ['Cost Analysis', 'Maintenance cost breakdown & equipment ranking'],
  fleetkpi: ['Fleet KPIs', 'Availability, utilization & reliability metrics'],
  settings: ['Settings / Data', 'Data backup, restore & fleet master']
};

function navigateTo(page) {
  document.querySelectorAll('.page-section').forEach(s => s.classList.remove('active'));
  const target = document.getElementById('page-' + page);
  if (target) target.classList.add('active');
  document.querySelectorAll('.sidebar .nav-link').forEach(l => l.classList.remove('active'));
  const link = document.querySelector('.sidebar .nav-link[data-page="' + page + '"]');
  if (link) link.classList.add('active');
  const titleInfo = PAGE_TITLES[page] || ['Dashboard',''];
  document.getElementById('pageTitle').textContent = titleInfo[0];
  document.getElementById('pageSubtitle').textContent = titleInfo[1];

  // close mobile sidebar
  document.getElementById('sidebar').classList.remove('show');
  document.getElementById('sidebarOverlay').classList.remove('show');

  // refresh relevant page data
  if (page === 'dashboard') renderDashboard();
  if (page === 'dailylog') renderDailyLogTable();
  if (page === 'inventory') renderAllInventoryTables();
  if (page === 'materialissue') renderMaterialIssueTable();
  if (page === 'service') renderServiceTable();
  if (page === 'breakdown') renderBreakdownTable();
  if (page === 'costanalysis') renderCostAnalysis();
  if (page === 'fleetkpi') renderFleetKpi();
  if (page === 'settings') renderFleetMasterTable();
}

// ===================== EXPORT / PRINT UTILITIES =====================
function exportTableToExcel(tableId, filename) {
  const table = document.getElementById(tableId);
  if (!table) return;
  // Clone and strip the Actions column before export
  const clone = table.cloneNode(true);
  const actionCols = clone.querySelectorAll('.no-print');
  actionCols.forEach(el => el.remove());
  const wb = XLSX.utils.table_to_book(clone, {sheet: "Sheet1"});
  XLSX.writeFile(wb, filename + '_' + todayStr() + '.xlsx');
  showToast('Exported to Excel successfully.');
}

function printTable(tableId, title) {
  const table = document.getElementById(tableId);
  if (!table) return;
  const clone = table.cloneNode(true);
  clone.querySelectorAll('.no-print').forEach(el => el.remove());
  const win = window.open('', '_blank');
  win.document.write(`
    <html><head><title>${escapeHtml(title)}</title>
    <style>
      body{font-family:Arial,sans-serif;padding:20px;}
      h2{color:#0f2942;border-bottom:2px solid #c9952c;padding-bottom:8px;}
      table{width:100%;border-collapse:collapse;font-size:12px;margin-top:14px;}
      th,td{border:1px solid #ccc;padding:6px 8px;text-align:left;}
      th{background:#0f2942;color:#fff;}
      .meta{font-size:11px;color:#666;margin-bottom:10px;}
    </style>
    </head><body>
    <h2>SGX Minerals Pvt. Ltd. — ${escapeHtml(title)}</h2>
    <div class="meta">Generated on ${new Date().toLocaleString('en-IN')}</div>
    ${clone.outerHTML}
    </body></html>
  `);
  win.document.close();
  win.focus();
  setTimeout(() => win.print(), 400);
}

// ===================== BACKUP / RESTORE / RESET =====================
function backupAllData() {
  const backup = {};
  Object.keys(DB_KEYS).forEach(k => backup[k] = dbGet(DB_KEYS[k]));
  backup._exportedAt = new Date().toISOString();
  const blob = new Blob([JSON.stringify(backup, null, 2)], {type: 'application/json'});
  const a = document.createElement('a');
  a.href = URL.createObjectURL(blob);
  a.download = 'SGX_Maintenance_Backup_' + todayStr() + '.json';
  a.click();
  showToast('Backup downloaded.');
}

function restoreAllData(evt) {
  const file = evt.target.files[0];
  if (!file) return;
  const reader = new FileReader();
  reader.onload = function(e) {
    try {
      const data = JSON.parse(e.target.result);
      Object.keys(DB_KEYS).forEach(k => {
        if (data[k]) dbSet(DB_KEYS[k], data[k]);
      });
      showToast('Data restored successfully. Reloading...');
      setTimeout(() => location.reload(), 1200);
    } catch(err) {
      showToast('Invalid backup file.', true);
    }
  };
  reader.readAsText(file);
  evt.target.value = '';
}

function resetAllData() {
  if (!confirm('This will permanently delete ALL data (daily logs, inventory, service records, breakdowns). This cannot be undone. Continue?')) return;
  Object.values(DB_KEYS).forEach(k => localStorage.removeItem(k));
  showToast('All data reset.');
  setTimeout(() => location.reload(), 800);
}

// ===================== INIT =====================
document.addEventListener('DOMContentLoaded', function() {
  document.getElementById('footerYear').textContent = new Date().getFullYear();
  document.getElementById('todayDateDisplay').textContent = new Date().toLocaleDateString('en-IN', {weekday:'long', year:'numeric', month:'long', day:'numeric'});

  populateVehicleDropdowns();

  // Sidebar nav
  document.querySelectorAll('.sidebar .nav-link').forEach(link => {
    link.addEventListener('click', function() {
      navigateTo(this.getAttribute('data-page'));
    });
  });

  // Mobile sidebar toggle
  document.getElementById('sidebarToggle').addEventListener('click', function() {
    document.getElementById('sidebar').classList.toggle('show');
    document.getElementById('sidebarOverlay').classList.toggle('show');
  });
  document.getElementById('sidebarOverlay').addEventListener('click', function() {
    document.getElementById('sidebar').classList.remove('show');
    this.classList.remove('show');
  });

  // Inventory tabs
  document.querySelectorAll('#invTabs .nav-link').forEach(tab => {
    tab.addEventListener('click', function() {
      document.querySelectorAll('#invTabs .nav-link').forEach(t => t.classList.remove('active'));
      this.classList.add('active');
      const target = this.getAttribute('data-inv');
      document.querySelectorAll('.inv-panel').forEach(p => p.style.display = 'none');
      document.getElementById('inv-' + target).style.display = 'block';
    });
  });

  // Daily log status dropdown listeners
  document.getElementById('dlDayStatus').addEventListener('change', toggleDailyLogBreakdownFields);
  document.getElementById('dlNightStatus').addEventListener('change', toggleDailyLogBreakdownFields);
  ['dlDayTrips','dlNightTrips'].forEach(id => document.getElementById(id).addEventListener('input', calcDailyLogDerived));
  ['dlOpenKm','dlCloseKm'].forEach(id => document.getElementById(id).addEventListener('input', calcDailyLogDerived));
  ['dlDayBdFrom','dlDayBdTo','dlNightBdFrom','dlNightBdTo'].forEach(id => document.getElementById(id).addEventListener('input', calcDailyLogDerived));

  // Inventory form listeners
  ['invOpening','invReceived','invIssued'].forEach(id => document.getElementById(id).addEventListener('input', calcInventoryClosing));

  // Search/filter listeners
  document.getElementById('dlFilterDate').addEventListener('input', renderDailyLogTable);
  document.getElementById('dlFilterVehicle').addEventListener('change', renderDailyLogTable);
  document.getElementById('dlSearchBox').addEventListener('input', renderDailyLogTable);
  document.getElementById('hwSearchBox').addEventListener('input', () => renderInventoryTable('hardware'));
  document.getElementById('lubSearchBox').addEventListener('input', () => renderInventoryTable('lubricants'));
  document.getElementById('wsSearchBox').addEventListener('input', () => renderInventoryTable('workshop'));
  document.getElementById('miSearchBox').addEventListener('input', renderMaterialIssueTable);
  document.getElementById('svFilterCategory').addEventListener('change', renderServiceTable);
  document.getElementById('svFilterStatus').addEventListener('change', renderServiceTable);
  document.getElementById('bdFilterCategory').addEventListener('change', renderBreakdownTable);
  document.getElementById('bdSearchBox').addEventListener('input', renderBreakdownTable);

  renderDashboard();
  updateServiceDueBadge();
});

function populateVehicleDropdowns() {
  const selects = ['dlVehicle','miEquipment','svEquipment','bdEquipment'];
  selects.forEach(selId => {
    const sel = document.getElementById(selId);
    if (!sel) return;
    // keep first placeholder option
    const placeholder = sel.options[0];
    sel.innerHTML = '';
    sel.appendChild(placeholder);
    FLEET.forEach(eq => {
      const opt = document.createElement('option');
      opt.value = eq.id;
      opt.textContent = eq.id + ' (' + eq.category + ')';
      sel.appendChild(opt);
    });
  });

  // daily log filter vehicle dropdown
  const filterSel = document.getElementById('dlFilterVehicle');
  FLEET.forEach(eq => {
    const opt = document.createElement('option');
    opt.value = eq.id;
    opt.textContent = eq.id;
    filterSel.appendChild(opt);
  });
}

function getFleetCategory(eqId) {
  const eq = FLEET.find(f => f.id === eqId);
  return eq ? eq.category : '';
}

/* ===================================================================
   DAILY LOG MODULE
   =================================================================== */

function toggleDailyLogBreakdownFields() {
  const dayStatus = document.getElementById('dlDayStatus').value;
  const nightStatus = document.getElementById('dlNightStatus').value;
  document.getElementById('dlDayBreakdownSection').style.display = (dayStatus === 'B') ? 'block' : 'none';
  document.getElementById('dlNightBreakdownSection').style.display = (nightStatus === 'B') ? 'block' : 'none';
  document.getElementById('dlMaintSection').style.display = (dayStatus === 'B' || nightStatus === 'B') ? 'block' : 'none';
  calcDailyLogDerived();
}

function calcDailyLogDerived() {
  const dayTrips = parseFloat(document.getElementById('dlDayTrips').value) || 0;
  const nightTrips = parseFloat(document.getElementById('dlNightTrips').value) || 0;
  document.getElementById('dlFtdTrips').value = (dayTrips + nightTrips);

  const openKm = parseFloat(document.getElementById('dlOpenKm').value) || 0;
  const closeKm = parseFloat(document.getElementById('dlCloseKm').value) || 0;
  const ftdRun = closeKm - openKm;
  document.getElementById('dlFtdRun').value = ftdRun.toFixed(1);

  const dayStatus = document.getElementById('dlDayStatus').value;
  const nightStatus = document.getElementById('dlNightStatus').value;
  let totalBdHours = 0;
  if (dayStatus === 'B') {
    totalBdHours += diffHoursFromTimes(document.getElementById('dlDayBdFrom').value, document.getElementById('dlDayBdTo').value);
  }
  if (nightStatus === 'B') {
    totalBdHours += diffHoursFromTimes(document.getElementById('dlNightBdFrom').value, document.getElementById('dlNightBdTo').value);
  }
  document.getElementById('dlTotalBdHours').textContent = totalBdHours.toFixed(2) + ' hrs';
}

function openDailyLogModal(editId) {
  document.getElementById('dailyLogForm').reset();
  document.getElementById('dlEditId').value = '';
  document.getElementById('dlDayBreakdownSection').style.display = 'none';
  document.getElementById('dlNightBreakdownSection').style.display = 'none';
  document.getElementById('dlMaintSection').style.display = 'none';
  document.getElementById('dlFtdTrips').value = '0';
  document.getElementById('dlFtdRun').value = '0';
  document.getElementById('dlTotalBdHours').textContent = '0.00 hrs';

  if (editId) {
    const logs = dbGet(DB_KEYS.dailyLogs);
    const log = logs.find(l => l.id === editId);
    if (log) {
      document.getElementById('dlEditId').value = log.id;
      document.getElementById('dlDate').value = log.date;
      document.getElementById('dlVehicle').value = log.vehicle;
      document.getElementById('dlDayStatus').value = log.dayStatus;
      document.getElementById('dlNightStatus').value = log.nightStatus;
      document.getElementById('dlDayTrips').value = log.dayTrips;
      document.getElementById('dlNightTrips').value = log.nightTrips;
      document.getElementById('dlFtdTons').value = log.ftdTons;
      document.getElementById('dlSthTrips').value = log.sthTrips;
      document.getElementById('dlObTrips').value = log.obTrips;
      document.getElementById('dlDiesel').value = log.diesel;
      document.getElementById('dlOpenKm').value = log.openKm;
      document.getElementById('dlCloseKm').value = log.closeKm;
      document.getElementById('dlDayBdFrom').value = log.dayBdFrom || '';
      document.getElementById('dlDayBdTo').value = log.dayBdTo || '';
      document.getElementById('dlNightBdFrom').value = log.nightBdFrom || '';
      document.getElementById('dlNightBdTo').value = log.nightBdTo || '';
      document.getElementById('dlMaintDesc').value = log.maintDesc || '';
      document.getElementById('dlRemarks').value = log.remarks || '';
      toggleDailyLogBreakdownFields();
    }
  } else {
    document.getElementById('dlDate').value = todayStr();
  }
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('dailyLogModal'));
  modal.show();
}

function saveDailyLog() {
  const date = document.getElementById('dlDate').value;
  const vehicle = document.getElementById('dlVehicle').value;
  const dayStatus = document.getElementById('dlDayStatus').value;
  const nightStatus = document.getElementById('dlNightStatus').value;

  if (!date || !vehicle || !dayStatus || !nightStatus) {
    showToast('Please fill all required fields (Date, Vehicle, Day Status, Night Status).', true);
    return;
  }

  const dayTrips = parseFloat(document.getElementById('dlDayTrips').value) || 0;
  const nightTrips = parseFloat(document.getElementById('dlNightTrips').value) || 0;
  const ftdTrips = dayTrips + nightTrips;
  const openKm = parseFloat(document.getElementById('dlOpenKm').value) || 0;
  const closeKm = parseFloat(document.getElementById('dlCloseKm').value) || 0;
  const ftdRun = closeKm - openKm;

  let totalBdHours = 0;
  const dayBdFrom = document.getElementById('dlDayBdFrom').value;
  const dayBdTo = document.getElementById('dlDayBdTo').value;
  const nightBdFrom = document.getElementById('dlNightBdFrom').value;
  const nightBdTo = document.getElementById('dlNightBdTo').value;
  if (dayStatus === 'B') totalBdHours += diffHoursFromTimes(dayBdFrom, dayBdTo);
  if (nightStatus === 'B') totalBdHours += diffHoursFromTimes(nightBdFrom, nightBdTo);

  const editId = document.getElementById('dlEditId').value;
  const logs = dbGet(DB_KEYS.dailyLogs);

  const record = {
    id: editId || uid(),
    date, vehicle, dayStatus, nightStatus,
    dayTrips, nightTrips, ftdTrips,
    ftdTons: parseFloat(document.getElementById('dlFtdTons').value) || 0,
    sthTrips: parseFloat(document.getElementById('dlSthTrips').value) || 0,
    obTrips: parseFloat(document.getElementById('dlObTrips').value) || 0,
    diesel: parseFloat(document.getElementById('dlDiesel').value) || 0,
    openKm, closeKm, ftdRun,
    dayBdFrom: dayStatus === 'B' ? dayBdFrom : '',
    dayBdTo: dayStatus === 'B' ? dayBdTo : '',
    nightBdFrom: nightStatus === 'B' ? nightBdFrom : '',
    nightBdTo: nightStatus === 'B' ? nightBdTo : '',
    maintDesc: document.getElementById('dlMaintDesc').value.trim(),
    remarks: document.getElementById('dlRemarks').value.trim(),
    totalBdHours,
    createdAt: editId ? (logs.find(l=>l.id===editId)?.createdAt || new Date().toISOString()) : new Date().toISOString()
  };

  if (editId) {
    const idx = logs.findIndex(l => l.id === editId);
    if (idx > -1) logs[idx] = record;
  } else {
    logs.push(record);
  }
  dbSet(DB_KEYS.dailyLogs, logs);

  // Also update service record current KM/HMR if exists
  updateServiceCurrentKmFromLog(vehicle, closeKm);

  bootstrap.Modal.getInstance(document.getElementById('dailyLogModal')).hide();
  showToast('Daily log entry saved successfully.');
  renderDailyLogTable();
  renderDashboard();
}

function updateServiceCurrentKmFromLog(vehicle, closeKm) {
  if (!closeKm) return;
  const records = dbGet(DB_KEYS.serviceRecords);
  let changed = false;
  records.forEach(r => {
    if (r.equipment === vehicle && closeKm > r.currentKm) {
      r.currentKm = closeKm;
      r.nextDue = r.lastKm + SERVICE_INTERVALS[r.serviceType];
      r.balance = r.nextDue - r.currentKm;
      r.status = computeServiceStatus(r.balance);
      changed = true;
    }
  });
  if (changed) dbSet(DB_KEYS.serviceRecords, records);
}

function deleteDailyLog(id) {
  if (!confirm('Delete this daily log entry?')) return;
  let logs = dbGet(DB_KEYS.dailyLogs);
  logs = logs.filter(l => l.id !== id);
  dbSet(DB_KEYS.dailyLogs, logs);
  showToast('Entry deleted.');
  renderDailyLogTable();
  renderDashboard();
}

function renderDailyLogTable() {
  const logs = dbGet(DB_KEYS.dailyLogs).slice().sort((a,b) => (b.date+b.createdAt).localeCompare(a.date+a.createdAt));
  const filterDate = document.getElementById('dlFilterDate').value;
  const filterVehicle = document.getElementById('dlFilterVehicle').value;
  const search = document.getElementById('dlSearchBox').value.toLowerCase();

  const filtered = logs.filter(l => {
    if (filterDate && l.date !== filterDate) return false;
    if (filterVehicle && l.vehicle !== filterVehicle) return false;
    if (search && !(l.vehicle.toLowerCase().includes(search) || (l.remarks||'').toLowerCase().includes(search) || (l.maintDesc||'').toLowerCase().includes(search))) return false;
    return true;
  });

  const tbody = document.getElementById('dailyLogTableBody');
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="16"><div class="empty-state"><i class="bi bi-inbox"></i>No daily log entries found</div></td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(l => `
    <tr>
      <td>${formatDateDisplay(l.date)}</td>
      <td><strong>${escapeHtml(l.vehicle)}</strong></td>
      <td><span class="badge-status ${l.dayStatus==='R'?'badge-running':'badge-breakdown'}">${l.dayStatus}</span></td>
      <td><span class="badge-status ${l.nightStatus==='R'?'badge-running':'badge-breakdown'}">${l.nightStatus}</span></td>
      <td>${fmtNum(l.dayTrips)}</td>
      <td>${fmtNum(l.nightTrips)}</td>
      <td><strong>${fmtNum(l.ftdTrips)}</strong></td>
      <td>${fmtNum(l.ftdTons,2)}</td>
      <td>${fmtNum(l.sthTrips)}</td>
      <td>${fmtNum(l.obTrips)}</td>
      <td>${fmtNum(l.diesel,2)}</td>
      <td>${fmtNum(l.openKm,1)}</td>
      <td>${fmtNum(l.closeKm,1)}</td>
      <td><strong>${fmtNum(l.ftdRun,1)}</strong></td>
      <td>${l.totalBdHours ? fmtNum(l.totalBdHours,2) : '-'}</td>
      <td class="no-print">
        <button class="btn btn-sm btn-outline-navy btn-icon-action" onclick="openDailyLogModal('${l.id}')"><i class="bi bi-pencil"></i></button>
        <button class="btn btn-sm btn-outline-danger btn-icon-action" onclick="deleteDailyLog('${l.id}')"><i class="bi bi-trash"></i></button>
      </td>
    </tr>
  `).join('');
}

function formatDateDisplay(dateStr) {
  if (!dateStr) return '-';
  const d = new Date(dateStr + 'T00:00:00');
  if (isNaN(d)) return dateStr;
  return d.toLocaleDateString('en-IN', {day:'2-digit', month:'short', year:'numeric'});
}

/* ===================================================================
   INVENTORY MODULE (Hardware / Lubricants / Workshop Consumables)
   =================================================================== */

function getInventoryKey(storeType) {
  return DB_KEYS[storeType]; // 'hardware' | 'lubricants' | 'workshop'
}

function calcInventoryClosing() {
  const opening = parseFloat(document.getElementById('invOpening').value) || 0;
  const received = parseFloat(document.getElementById('invReceived').value) || 0;
  const issued = parseFloat(document.getElementById('invIssued').value) || 0;
  const closing = opening + received - issued;
  document.getElementById('invClosing').value = closing.toFixed(2);
}

function openInventoryModal(storeType, editId) {
  document.getElementById('inventoryForm').reset();
  document.getElementById('invEditId').value = '';
  document.getElementById('invStoreType').value = storeType;
  document.getElementById('invMaterialCodeWrap').style.display = (storeType === 'hardware') ? 'block' : 'none';
  document.getElementById('invClosing').value = '0';

  const titles = {hardware: 'Hardware Store Material', lubricants: 'Lubricants Store Material', workshop: 'Workshop Consumable'};
  document.getElementById('invModalTitle').innerHTML = '<i class="bi bi-boxes"></i> ' + (editId ? 'Edit ' : 'Add ') + titles[storeType];

  if (storeType === 'lubricants') {
    document.getElementById('invUnit').value = 'Ltr';
  }

  if (editId) {
    const items = dbGet(getInventoryKey(storeType));
    const item = items.find(i => i.id === editId);
    if (item) {
      document.getElementById('invEditId').value = item.id;
      document.getElementById('invMaterialCode').value = item.code || '';
      document.getElementById('invMaterialName').value = item.name;
      document.getElementById('invUnit').value = item.unit;
      document.getElementById('invOpening').value = item.opening;
      document.getElementById('invReceived').value = item.received;
      document.getElementById('invIssued').value = item.issued;
      calcInventoryClosing();
    }
  }
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('inventoryModal'));
  modal.show();
}

function saveInventoryItem() {
  const storeType = document.getElementById('invStoreType').value;
  const name = document.getElementById('invMaterialName').value.trim();
  const unit = document.getElementById('invUnit').value;
  if (!name || !unit) {
    showToast('Please fill all required fields.', true);
    return;
  }
  const opening = parseFloat(document.getElementById('invOpening').value) || 0;
  const received = parseFloat(document.getElementById('invReceived').value) || 0;
  const issued = parseFloat(document.getElementById('invIssued').value) || 0;
  const closing = opening + received - issued;
  const editId = document.getElementById('invEditId').value;

  const items = dbGet(getInventoryKey(storeType));
  const record = {
    id: editId || uid(),
    code: storeType === 'hardware' ? document.getElementById('invMaterialCode').value.trim() : '',
    name, unit, opening, received, issued, closing
  };

  if (editId) {
    const idx = items.findIndex(i => i.id === editId);
    if (idx > -1) items[idx] = record;
  } else {
    items.push(record);
  }
  dbSet(getInventoryKey(storeType), items);
  bootstrap.Modal.getInstance(document.getElementById('inventoryModal')).hide();
  showToast('Material saved successfully.');
  renderInventoryTable(storeType);
}

function deleteInventoryItem(storeType, id) {
  if (!confirm('Delete this material record?')) return;
  let items = dbGet(getInventoryKey(storeType));
  items = items.filter(i => i.id !== id);
  dbSet(getInventoryKey(storeType), items);
  showToast('Material deleted.');
  renderInventoryTable(storeType);
}

function renderAllInventoryTables() {
  renderInventoryTable('hardware');
  renderInventoryTable('lubricants');
  renderInventoryTable('workshop');
}

function renderInventoryTable(storeType) {
  const items = dbGet(getInventoryKey(storeType)).slice().sort((a,b) => a.name.localeCompare(b.name));
  const searchIds = {hardware:'hwSearchBox', lubricants:'lubSearchBox', workshop:'wsSearchBox'};
  const bodyIds = {hardware:'hwTableBody', lubricants:'lubTableBody', workshop:'wsTableBody'};
  const search = document.getElementById(searchIds[storeType]).value.toLowerCase();

  const filtered = items.filter(i => !search || i.name.toLowerCase().includes(search) || (i.code||'').toLowerCase().includes(search));
  const tbody = document.getElementById(bodyIds[storeType]);
  const colspan = storeType === 'hardware' ? 8 : 7;

  if (filtered.length === 0) {
    tbody.innerHTML = `<tr><td colspan="${colspan}"><div class="empty-state"><i class="bi bi-inbox"></i>No materials found</div></td></tr>`;
    return;
  }

  tbody.innerHTML = filtered.map(i => {
    const lowStock = i.closing <= 0;
    const codeCol = storeType === 'hardware' ? `<td>${escapeHtml(i.code || '-')}</td>` : '';
    return `
    <tr class="${lowStock ? 'row-red' : ''}">
      ${codeCol}
      <td><strong>${escapeHtml(i.name)}</strong></td>
      <td>${escapeHtml(i.unit)}</td>
      <td>${fmtNum(i.opening,2)}</td>
      <td>${fmtNum(i.received,2)}</td>
      <td>${fmtNum(i.issued,2)}</td>
      <td><strong>${fmtNum(i.closing,2)}</strong> ${lowStock ? '<span class="badge-status badge-due">OUT OF STOCK</span>' : ''}</td>
      <td class="no-print">
        <button class="btn btn-sm btn-outline-navy btn-icon-action" onclick="openInventoryModal('${storeType}','${i.id}')"><i class="bi bi-pencil"></i></button>
        <button class="btn btn-sm btn-outline-danger btn-icon-action" onclick="deleteInventoryItem('${storeType}','${i.id}')"><i class="bi bi-trash"></i></button>
      </td>
    </tr>`;
  }).join('');
}

/* ===================================================================
   MATERIAL ISSUE MODULE (auto-deducts inventory stock)
   =================================================================== */

function populateMaterialOptions() {
  const storeType = document.getElementById('miStore').value;
  const sel = document.getElementById('miMaterial');
  sel.innerHTML = '<option value="">Select Material</option>';
  document.getElementById('miAvailStock').textContent = '-';
  if (!storeType) return;
  const items = dbGet(getInventoryKey(storeType)).slice().sort((a,b) => a.name.localeCompare(b.name));
  items.forEach(item => {
    const opt = document.createElement('option');
    opt.value = item.id;
    opt.textContent = `${item.name} (${item.unit}) — Stock: ${fmtNum(item.closing,2)}`;
    opt.dataset.unit = item.unit;
    opt.dataset.closing = item.closing;
    sel.appendChild(opt);
  });
}

function updateMiAvailableStock() {
  const sel = document.getElementById('miMaterial');
  const opt = sel.options[sel.selectedIndex];
  if (opt && opt.value) {
    document.getElementById('miAvailStock').textContent = fmtNum(parseFloat(opt.dataset.closing),2) + ' ' + opt.dataset.unit;
  } else {
    document.getElementById('miAvailStock').textContent = '-';
  }
}

function openMaterialIssueModal() {
  document.getElementById('materialIssueForm').reset();
  document.getElementById('miDate').value = todayStr();
  document.getElementById('miMaterial').innerHTML = '<option value="">Select Material</option>';
  document.getElementById('miAvailStock').textContent = '-';
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('materialIssueModal'));
  modal.show();
}

function saveMaterialIssue() {
  const date = document.getElementById('miDate').value;
  const equipment = document.getElementById('miEquipment').value;
  const storeType = document.getElementById('miStore').value;
  const materialId = document.getElementById('miMaterial').value;
  const qty = parseFloat(document.getElementById('miQty').value);
  const purpose = document.getElementById('miPurpose').value.trim();
  const issuedBy = document.getElementById('miIssuedBy').value.trim();

  if (!date || !equipment || !storeType || !materialId || !qty || qty <= 0 || !purpose || !issuedBy) {
    showToast('Please fill all required fields with valid values.', true);
    return;
  }

  const items = dbGet(getInventoryKey(storeType));
  const item = items.find(i => i.id === materialId);
  if (!item) {
    showToast('Material not found.', true);
    return;
  }
  if (qty > item.closing) {
    if (!confirm(`Warning: Issuing ${qty} ${item.unit} exceeds available stock of ${item.closing} ${item.unit}. Stock will go negative. Continue?`)) {
      return;
    }
  }

  // Deduct stock
  item.issued = (item.issued || 0) + qty;
  item.closing = item.opening + item.received - item.issued;
  dbSet(getInventoryKey(storeType), items);

  // Record issue
  const issues = dbGet(DB_KEYS.materialIssues);
  issues.push({
    id: uid(), date, equipment, storeType, materialId,
    materialName: item.name, unit: item.unit, qty, purpose, issuedBy,
    createdAt: new Date().toISOString()
  });
  dbSet(DB_KEYS.materialIssues, issues);

  bootstrap.Modal.getInstance(document.getElementById('materialIssueModal')).hide();
  showToast('Material issued and stock updated.');
  renderMaterialIssueTable();
}

function deleteMaterialIssue(id) {
  if (!confirm('Delete this material issue record? This will NOT automatically restore the deducted stock.')) return;
  let issues = dbGet(DB_KEYS.materialIssues);
  issues = issues.filter(i => i.id !== id);
  dbSet(DB_KEYS.materialIssues, issues);
  showToast('Issue record deleted.');
  renderMaterialIssueTable();
}

const STORE_LABELS = {hardware: 'Hardware', lubricants: 'Lubricants', workshop: 'Workshop'};

function renderMaterialIssueTable() {
  const issues = dbGet(DB_KEYS.materialIssues).slice().sort((a,b) => (b.date+b.createdAt).localeCompare(a.date+a.createdAt));
  const search = document.getElementById('miSearchBox').value.toLowerCase();
  const filtered = issues.filter(i => !search ||
    i.equipment.toLowerCase().includes(search) ||
    i.materialName.toLowerCase().includes(search) ||
    i.purpose.toLowerCase().includes(search) ||
    i.issuedBy.toLowerCase().includes(search));

  const tbody = document.getElementById('miTableBody');
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="9"><div class="empty-state"><i class="bi bi-inbox"></i>No material issue records found</div></td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(i => `
    <tr>
      <td>${formatDateDisplay(i.date)}</td>
      <td><strong>${escapeHtml(i.equipment)}</strong></td>
      <td>${escapeHtml(i.materialName)}</td>
      <td>${STORE_LABELS[i.storeType] || i.storeType}</td>
      <td>${fmtNum(i.qty,2)}</td>
      <td>${escapeHtml(i.unit)}</td>
      <td>${escapeHtml(i.purpose)}</td>
      <td>${escapeHtml(i.issuedBy)}</td>
      <td class="no-print">
        <button class="btn btn-sm btn-outline-danger btn-icon-action" onclick="deleteMaterialIssue('${i.id}')"><i class="bi bi-trash"></i></button>
      </td>
    </tr>
  `).join('');
}

/* ===================================================================
   SERVICE SCHEDULE MODULE
   =================================================================== */

function computeServiceStatus(balance) {
  if (balance <= 0) return 'Due';
  if (balance < 100) return 'Due Soon';
  return 'OK';
}

function autoFillServiceCategory() {
  const eq = document.getElementById('svEquipment').value;
  document.getElementById('svCategory').value = getFleetCategory(eq);
}

function calcServiceFields() {
  const lastKm = parseFloat(document.getElementById('svLastKm').value) || 0;
  const currentKm = parseFloat(document.getElementById('svCurrentKm').value) || 0;
  const serviceType = document.getElementById('svServiceType').value;
  if (!serviceType) {
    document.getElementById('svNextDue').value = '';
    document.getElementById('svBalance').value = '';
    document.getElementById('svStatusDisplay').textContent = '-';
    return;
  }
  const interval = SERVICE_INTERVALS[serviceType];
  const nextDue = lastKm + interval;
  const balance = nextDue - currentKm;
  document.getElementById('svNextDue').value = nextDue.toFixed(1);
  document.getElementById('svBalance').value = balance.toFixed(1);
  const status = computeServiceStatus(balance);
  const statusEl = document.getElementById('svStatusDisplay');
  statusEl.textContent = status + ' (Balance: ' + balance.toFixed(1) + ')';
  statusEl.className = 'readonly-pill ' + (status === 'Due' ? 'text-danger' : status === 'Due Soon' ? 'text-warning' : 'text-success');
}

function openServiceModal(editId) {
  document.getElementById('serviceForm').reset();
  document.getElementById('svEditId').value = '';
  document.getElementById('svCategory').value = '';
  document.getElementById('svNextDue').value = '';
  document.getElementById('svBalance').value = '';
  document.getElementById('svStatusDisplay').textContent = '-';
  document.getElementById('svStatusDisplay').className = 'readonly-pill';

  if (editId) {
    const records = dbGet(DB_KEYS.serviceRecords);
    const r = records.find(x => x.id === editId);
    if (r) {
      document.getElementById('svEditId').value = r.id;
      document.getElementById('svEquipment').value = r.equipment;
      document.getElementById('svCategory').value = r.category;
      document.getElementById('svServiceType').value = r.serviceType;
      document.getElementById('svLastDate').value = r.lastDate;
      document.getElementById('svLastKm').value = r.lastKm;
      document.getElementById('svCurrentKm').value = r.currentKm;
      calcServiceFields();
    }
  } else {
    document.getElementById('svLastDate').value = todayStr();
  }
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('serviceModal'));
  modal.show();
}

function saveServiceRecord() {
  const equipment = document.getElementById('svEquipment').value;
  const serviceType = document.getElementById('svServiceType').value;
  const lastDate = document.getElementById('svLastDate').value;
  const lastKm = parseFloat(document.getElementById('svLastKm').value);
  const currentKm = parseFloat(document.getElementById('svCurrentKm').value);

  if (!equipment || !serviceType || !lastDate || isNaN(lastKm) || isNaN(currentKm)) {
    showToast('Please fill all required fields.', true);
    return;
  }

  const interval = SERVICE_INTERVALS[serviceType];
  const nextDue = lastKm + interval;
  const balance = nextDue - currentKm;
  const status = computeServiceStatus(balance);
  const editId = document.getElementById('svEditId').value;

  const records = dbGet(DB_KEYS.serviceRecords);
  const record = {
    id: editId || uid(),
    equipment, category: getFleetCategory(equipment), serviceType,
    lastDate, lastKm, currentKm, nextDue, balance, status
  };

  if (editId) {
    const idx = records.findIndex(r => r.id === editId);
    if (idx > -1) records[idx] = record;
  } else {
    // Remove any existing record for same equipment+serviceType combo to avoid duplicates
    const existingIdx = records.findIndex(r => r.equipment === equipment && r.serviceType === serviceType);
    if (existingIdx > -1) records.splice(existingIdx, 1);
    records.push(record);
  }
  dbSet(DB_KEYS.serviceRecords, records);
  bootstrap.Modal.getInstance(document.getElementById('serviceModal')).hide();
  showToast('Service record saved.');
  renderServiceTable();
  updateServiceDueBadge();
}

function deleteServiceRecord(id) {
  if (!confirm('Delete this service record?')) return;
  let records = dbGet(DB_KEYS.serviceRecords);
  records = records.filter(r => r.id !== id);
  dbSet(DB_KEYS.serviceRecords, records);
  showToast('Service record deleted.');
  renderServiceTable();
  updateServiceDueBadge();
}

function renderServiceTable() {
  const records = dbGet(DB_KEYS.serviceRecords).slice().sort((a,b) => a.balance - b.balance);
  const catFilter = document.getElementById('svFilterCategory').value;
  const statusFilter = document.getElementById('svFilterStatus').value;

  const filtered = records.filter(r => {
    if (catFilter && r.category !== catFilter) return false;
    if (statusFilter && r.status !== statusFilter) return false;
    return true;
  });

  const tbody = document.getElementById('serviceTableBody');
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="10"><div class="empty-state"><i class="bi bi-inbox"></i>No service records found</div></td></tr>';
  } else {
    tbody.innerHTML = filtered.map(r => {
      const rowClass = r.status === 'Due' ? 'row-red' : r.status === 'Due Soon' ? 'row-orange' : '';
      const badgeClass = r.status === 'Due' ? 'badge-due' : r.status === 'Due Soon' ? 'badge-warn' : 'badge-ok';
      return `
      <tr class="${rowClass}">
        <td><strong>${escapeHtml(r.equipment)}</strong></td>
        <td>${escapeHtml(r.category)}</td>
        <td>${r.serviceType} Service</td>
        <td>${formatDateDisplay(r.lastDate)}</td>
        <td>${fmtNum(r.lastKm,1)}</td>
        <td>${fmtNum(r.currentKm,1)}</td>
        <td>${fmtNum(r.nextDue,1)}</td>
        <td><strong>${fmtNum(r.balance,1)}</strong></td>
        <td><span class="badge-status ${badgeClass}">${r.status}</span></td>
        <td class="no-print">
          <button class="btn btn-sm btn-outline-navy btn-icon-action" onclick="openServiceModal('${r.id}')"><i class="bi bi-pencil"></i></button>
          <button class="btn btn-sm btn-outline-danger btn-icon-action" onclick="deleteServiceRecord('${r.id}')"><i class="bi bi-trash"></i></button>
        </td>
      </tr>`;
    }).join('');
  }

  renderServiceAlerts();
}

function renderServiceAlerts() {
  const records = dbGet(DB_KEYS.serviceRecords);
  const dueRecords = records.filter(r => r.status === 'Due' || r.status === 'Due Soon').sort((a,b) => a.balance - b.balance);
  const container = document.getElementById('serviceAlertsContainer');

  if (dueRecords.length === 0) {
    container.innerHTML = '';
    return;
  }
  container.innerHTML = dueRecords.map(r => {
    const cls = r.status === 'Due' ? 'red' : 'orange';
    const icon = r.status === 'Due' ? 'bi-exclamation-octagon-fill' : 'bi-exclamation-triangle-fill';
    return `<div class="alert-due-strip ${cls}"><i class="bi ${icon}"></i> ${escapeHtml(r.equipment)} — ${r.serviceType} Service ${r.status === 'Due' ? 'OVERDUE' : 'due soon'} (Balance: ${fmtNum(r.balance,1)} KM/HMR)</div>`;
  }).join('');
}

function updateServiceDueBadge() {
  const records = dbGet(DB_KEYS.serviceRecords);
  const dueCount = records.filter(r => r.status === 'Due' || r.status === 'Due Soon').length;
  const badge = document.getElementById('serviceDueBadge');
  if (dueCount > 0) {
    badge.style.display = 'inline-block';
    badge.textContent = dueCount;
  } else {
    badge.style.display = 'none';
  }
}

/* ===================================================================
   BREAKDOWN HISTORY MODULE
   =================================================================== */

function calcBreakdownDowntime() {
  const start = document.getElementById('bdStart').value;
  const end = document.getElementById('bdEnd').value;
  const hours = diffHoursFromDatetime(start, end);
  document.getElementById('bdDowntime').value = hours.toFixed(2);
}

function calcBreakdownTotalCost() {
  const spare = parseFloat(document.getElementById('bdSpareCost').value) || 0;
  const lube = parseFloat(document.getElementById('bdLubeCost').value) || 0;
  const labour = parseFloat(document.getElementById('bdLabourCost').value) || 0;
  const vendor = parseFloat(document.getElementById('bdVendorCost').value) || 0;
  const total = spare + lube + labour + vendor;
  document.getElementById('bdTotalCostDisplay').textContent = fmtCurrency2(total);
}

function openBreakdownModal(editId) {
  document.getElementById('breakdownForm').reset();
  document.getElementById('bdEditId').value = '';
  document.getElementById('bdDowntime').value = '0.00';
  document.getElementById('bdTotalCostDisplay').textContent = '₹0.00';
  ['bdSpareCost','bdLubeCost','bdLabourCost','bdVendorCost'].forEach(id => document.getElementById(id).value = '0');

  if (editId) {
    const records = dbGet(DB_KEYS.breakdowns);
    const r = records.find(x => x.id === editId);
    if (r) {
      document.getElementById('bdEditId').value = r.id;
      document.getElementById('bdDate').value = r.date;
      document.getElementById('bdEquipment').value = r.equipment;
      document.getElementById('bdCategory').value = r.category;
      document.getElementById('bdProblem').value = r.problem;
      document.getElementById('bdRootCause').value = r.rootCause || '';
      document.getElementById('bdAction').value = r.action || '';
      document.getElementById('bdStart').value = r.start;
      document.getElementById('bdEnd').value = r.end;
      document.getElementById('bdParts').value = r.parts || '';
      document.getElementById('bdMechanic').value = r.mechanic || '';
      document.getElementById('bdVendor').value = r.vendor || '';
      document.getElementById('bdSpareCost').value = r.spareCost;
      document.getElementById('bdLubeCost').value = r.lubeCost;
      document.getElementById('bdLabourCost').value = r.labourCost;
      document.getElementById('bdVendorCost').value = r.vendorCost;
      calcBreakdownDowntime();
      calcBreakdownTotalCost();
    }
  } else {
    document.getElementById('bdDate').value = todayStr();
  }
  const modal = bootstrap.Modal.getOrCreateInstance(document.getElementById('breakdownModal'));
  modal.show();
}

function saveBreakdownRecord() {
  const date = document.getElementById('bdDate').value;
  const equipment = document.getElementById('bdEquipment').value;
  const category = document.getElementById('bdCategory').value;
  const problem = document.getElementById('bdProblem').value.trim();
  const start = document.getElementById('bdStart').value;
  const end = document.getElementById('bdEnd').value;

  if (!date || !equipment || !category || !problem || !start || !end) {
    showToast('Please fill all required fields (Date, Equipment, Category, Problem, Start/End time).', true);
    return;
  }
  if (new Date(end) < new Date(start)) {
    showToast('End time cannot be before start time.', true);
    return;
  }

  const downtime = diffHoursFromDatetime(start, end);
  const spareCost = parseFloat(document.getElementById('bdSpareCost').value) || 0;
  const lubeCost = parseFloat(document.getElementById('bdLubeCost').value) || 0;
  const labourCost = parseFloat(document.getElementById('bdLabourCost').value) || 0;
  const vendorCost = parseFloat(document.getElementById('bdVendorCost').value) || 0;
  const totalCost = spareCost + lubeCost + labourCost + vendorCost;
  const editId = document.getElementById('bdEditId').value;

  const records = dbGet(DB_KEYS.breakdowns);
  const record = {
    id: editId || uid(),
    date, equipment, category, problem,
    rootCause: document.getElementById('bdRootCause').value.trim(),
    action: document.getElementById('bdAction').value.trim(),
    start, end, downtime,
    parts: document.getElementById('bdParts').value.trim(),
    mechanic: document.getElementById('bdMechanic').value.trim(),
    vendor: document.getElementById('bdVendor').value.trim(),
    spareCost, lubeCost, labourCost, vendorCost, totalCost,
    createdAt: editId ? (records.find(r=>r.id===editId)?.createdAt || new Date().toISOString()) : new Date().toISOString()
  };

  if (editId) {
    const idx = records.findIndex(r => r.id === editId);
    if (idx > -1) records[idx] = record;
  } else {
    records.push(record);
  }
  dbSet(DB_KEYS.breakdowns, records);
  bootstrap.Modal.getInstance(document.getElementById('breakdownModal')).hide();
  showToast('Breakdown record saved.');
  renderBreakdownTable();
  renderDashboard();
}

function deleteBreakdownRecord(id) {
  if (!confirm('Delete this breakdown record?')) return;
  let records = dbGet(DB_KEYS.breakdowns);
  records = records.filter(r => r.id !== id);
  dbSet(DB_KEYS.breakdowns, records);
  showToast('Breakdown record deleted.');
  renderBreakdownTable();
  renderDashboard();
}

function renderBreakdownTable() {
  const records = dbGet(DB_KEYS.breakdowns).slice().sort((a,b) => (b.date+b.createdAt).localeCompare(a.date+a.createdAt));
  const catFilter = document.getElementById('bdFilterCategory').value;
  const search = document.getElementById('bdSearchBox').value.toLowerCase();

  const filtered = records.filter(r => {
    if (catFilter && r.category !== catFilter) return false;
    if (search && !(r.equipment.toLowerCase().includes(search) || r.problem.toLowerCase().includes(search) || (r.mechanic||'').toLowerCase().includes(search) || (r.vendor||'').toLowerCase().includes(search))) return false;
    return true;
  });

  const tbody = document.getElementById('breakdownTableBody');
  if (filtered.length === 0) {
    tbody.innerHTML = '<tr><td colspan="15"><div class="empty-state"><i class="bi bi-inbox"></i>No breakdown records found</div></td></tr>';
    return;
  }

  tbody.innerHTML = filtered.map(r => `
    <tr>
      <td>${formatDateDisplay(r.date)}</td>
      <td><strong>${escapeHtml(r.equipment)}</strong></td>
      <td><span class="badge-status badge-warn">${escapeHtml(r.category)}</span></td>
      <td style="max-width:220px;white-space:normal;">${escapeHtml(r.problem)}</td>
      <td>${formatDateTimeDisplay(r.start)}</td>
      <td>${formatDateTimeDisplay(r.end)}</td>
      <td><strong>${fmtNum(r.downtime,2)}</strong></td>
      <td>${escapeHtml(r.mechanic || '-')}</td>
      <td>${escapeHtml(r.vendor || '-')}</td>
      <td>${fmtCurrency(r.spareCost)}</td>
      <td>${fmtCurrency(r.lubeCost)}</td>
      <td>${fmtCurrency(r.labourCost)}</td>
      <td>${fmtCurrency(r.vendorCost)}</td>
      <td><strong>${fmtCurrency(r.totalCost)}</strong></td>
      <td class="no-print">
        <button class="btn btn-sm btn-outline-navy btn-icon-action" onclick="openBreakdownModal('${r.id}')"><i class="bi bi-pencil"></i></button>
        <button class="btn btn-sm btn-outline-danger btn-icon-action" onclick="deleteBreakdownRecord('${r.id}')"><i class="bi bi-trash"></i></button>
      </td>
    </tr>
  `).join('');
}

function formatDateTimeDisplay(dt) {
  if (!dt) return '-';
  const d = new Date(dt);
  if (isNaN(d)) return dt;
  return d.toLocaleString('en-IN', {day:'2-digit', month:'short', hour:'2-digit', minute:'2-digit'});
}

/* ===================================================================
   ANALYTICS ENGINE — shared cost & KPI computation
   =================================================================== */

// Diesel price assumption for cost calc (editable constant)
const DIESEL_PRICE_PER_LTR = 95;

function computeEquipmentCostMap() {
  const map = {};
  FLEET.forEach(eq => {
    map[eq.id] = {equipment: eq.id, category: eq.category, fuelCost: 0, spareCost: 0, lubeCost: 0, labourCost: 0, breakdownCost: 0, totalCost: 0, totalDiesel: 0, totalTrips: 0, totalTons: 0, totalKmRun: 0};
  });

  const logs = dbGet(DB_KEYS.dailyLogs);
  logs.forEach(l => {
    if (!map[l.vehicle]) return;
    map[l.vehicle].totalDiesel += (l.diesel || 0);
    map[l.vehicle].fuelCost += (l.diesel || 0) * DIESEL_PRICE_PER_LTR;
    map[l.vehicle].totalTrips += (l.ftdTrips || 0);
    map[l.vehicle].totalTons += (l.ftdTons || 0);
    map[l.vehicle].totalKmRun += Math.max(0, l.ftdRun || 0);
  });

  const breakdowns = dbGet(DB_KEYS.breakdowns);
  breakdowns.forEach(b => {
    if (!map[b.equipment]) return;
    map[b.equipment].spareCost += (b.spareCost || 0);
    map[b.equipment].lubeCost += (b.lubeCost || 0);
    map[b.equipment].labourCost += (b.labourCost || 0);
    map[b.equipment].breakdownCost += (b.totalCost || 0); // total repair cost incl. vendor
  });

  Object.values(map).forEach(m => {
    m.totalCost = m.fuelCost + m.spareCost + m.lubeCost + m.labourCost + (m.breakdownCost - m.spareCost - m.lubeCost - m.labourCost);
    // breakdownCost already includes spare+lube+labour+vendor, so totalCost = fuelCost + breakdownCost
    m.totalCost = m.fuelCost + m.breakdownCost;
  });

  return map;
}

function computeOverallTotals() {
  const map = computeEquipmentCostMap();
  const totals = {fuelCost:0, spareCost:0, lubeCost:0, labourCost:0, breakdownCost:0, totalCost:0, totalDiesel:0, totalTrips:0, totalTons:0, totalKmRun:0};
  Object.values(map).forEach(m => {
    totals.fuelCost += m.fuelCost;
    totals.spareCost += m.spareCost;
    totals.lubeCost += m.lubeCost;
    totals.labourCost += m.labourCost;
    totals.breakdownCost += m.breakdownCost;
    totals.totalCost += m.totalCost;
    totals.totalDiesel += m.totalDiesel;
    totals.totalTrips += m.totalTrips;
    totals.totalTons += m.totalTons;
    totals.totalKmRun += m.totalKmRun;
  });
  return totals;
}

function computeFleetAvailability() {
  const logs = dbGet(DB_KEYS.dailyLogs);
  if (logs.length === 0) return {availability: 0, utilization: 0};
  let totalShifts = 0, runningShifts = 0, totalTrips = 0;
  logs.forEach(l => {
    totalShifts += 2; // day + night
    if (l.dayStatus === 'R') runningShifts++;
    if (l.nightStatus === 'R') runningShifts++;
    totalTrips += (l.ftdTrips || 0);
  });
  const availability = totalShifts > 0 ? (runningShifts / totalShifts) * 100 : 0;
  // Utilization: trips achieved vs theoretical max (assume 2 trips/shift baseline reference of 4 trips/day per vehicle as "full utilization")
  const distinctDays = new Set(logs.map(l => l.date + '_' + l.vehicle)).size;
  const theoreticalMaxTrips = distinctDays * 4;
  const utilization = theoreticalMaxTrips > 0 ? Math.min(100, (totalTrips / theoreticalMaxTrips) * 100) : 0;
  return {availability, utilization};
}

function renderDashboard() {
  const totals = computeOverallTotals();
  document.getElementById('kpiDieselCost').textContent = fmtCurrency(totals.fuelCost);
  document.getElementById('kpiSpareCost').textContent = fmtCurrency(totals.spareCost);
  document.getElementById('kpiLubeCost').textContent = fmtCurrency(totals.lubeCost);
  document.getElementById('kpiLabourCost').textContent = fmtCurrency(totals.labourCost);
  document.getElementById('kpiBreakdownCost').textContent = fmtCurrency(totals.breakdownCost);
  document.getElementById('kpiTotalCost').textContent = fmtCurrency(totals.totalCost);

  const {availability, utilization} = computeFleetAvailability();
  document.getElementById('kpiAvailability').textContent = availability.toFixed(1) + '%';
  document.getElementById('kpiUtilization').textContent = utilization.toFixed(1) + '%';
  document.getElementById('kpiCostPerTrip').textContent = totals.totalTrips > 0 ? fmtCurrency(totals.totalCost / totals.totalTrips) : '₹0';
  document.getElementById('kpiCostPerTon').textContent = totals.totalTons > 0 ? fmtCurrency(totals.totalCost / totals.totalTons) : '₹0';

  renderEquipmentCostTable('eqCostTableBody');
  renderDashboardCharts();
  renderDashboardServiceAlerts();
}

function renderEquipmentCostTable(bodyId) {
  const map = computeEquipmentCostMap();
  const rows = Object.values(map).sort((a,b) => b.totalCost - a.totalCost);
  const tbody = document.getElementById(bodyId);
  tbody.innerHTML = rows.map(r => `
    <tr>
      <td><strong>${escapeHtml(r.equipment)}</strong> <span class="text-muted" style="font-size:11px;">(${escapeHtml(r.category)})</span></td>
      <td>${fmtCurrency(r.fuelCost)}</td>
      <td>${fmtCurrency(r.spareCost)}</td>
      <td>${fmtCurrency(r.lubeCost)}</td>
      <td>${fmtCurrency(r.labourCost)}</td>
      <td>${fmtCurrency(r.breakdownCost)}</td>
      <td><strong>${fmtCurrency(r.totalCost)}</strong></td>
    </tr>
  `).join('');
}

function renderDashboardServiceAlerts() {
  const records = dbGet(DB_KEYS.serviceRecords);
  const dueRecords = records.filter(r => r.status === 'Due' || r.status === 'Due Soon').sort((a,b) => a.balance - b.balance).slice(0, 8);
  const container = document.getElementById('dashboardServiceAlerts');
  if (dueRecords.length === 0) {
    container.innerHTML = '<div class="empty-state"><i class="bi bi-check-circle"></i>No alerts to show</div>';
    return;
  }
  container.innerHTML = dueRecords.map(r => {
    const cls = r.status === 'Due' ? 'red' : 'orange';
    const icon = r.status === 'Due' ? 'bi-exclamation-octagon-fill' : 'bi-exclamation-triangle-fill';
    return `<div class="alert-due-strip ${cls}"><i class="bi ${icon}"></i> ${escapeHtml(r.equipment)} — ${r.serviceType} Service (Bal: ${fmtNum(r.balance,1)})</div>`;
  }).join('');
}

/* ===================== CHARTS ===================== */
let chartInstances = {};
function destroyChart(key) {
  if (chartInstances[key]) { chartInstances[key].destroy(); delete chartInstances[key]; }
}

function getMonthlyAggregates() {
  const logs = dbGet(DB_KEYS.dailyLogs);
  const breakdowns = dbGet(DB_KEYS.breakdowns);
  const monthMap = {}; // 'YYYY-MM' -> {diesel, fuelCost, spareCost, lubeCost, labourCost, breakdownCost}

  function ensureMonth(key) {
    if (!monthMap[key]) monthMap[key] = {diesel:0, fuelCost:0, breakdownCost:0};
    return monthMap[key];
  }

  logs.forEach(l => {
    const key = l.date ? l.date.slice(0,7) : 'unknown';
    const m = ensureMonth(key);
    m.diesel += (l.diesel || 0);
    m.fuelCost += (l.diesel || 0) * DIESEL_PRICE_PER_LTR;
  });
  breakdowns.forEach(b => {
    const key = b.date ? b.date.slice(0,7) : 'unknown';
    const m = ensureMonth(key);
    m.breakdownCost += (b.totalCost || 0);
  });

  const sortedKeys = Object.keys(monthMap).filter(k => k !== 'unknown').sort();
  return {sortedKeys, monthMap};
}

function monthLabel(ym) {
  const [y,m] = ym.split('-');
  const d = new Date(parseInt(y), parseInt(m)-1, 1);
  return d.toLocaleDateString('en-IN', {month:'short', year:'2-digit'});
}

function renderDashboardCharts() {
  const {sortedKeys, monthMap} = getMonthlyAggregates();
  const labels = sortedKeys.map(monthLabel);

  destroyChart('costTrend');
  const ctx1 = document.getElementById('chartCostTrend');
  chartInstances.costTrend = new Chart(ctx1, {
    type: 'line',
    data: {
      labels: labels.length ? labels : ['No Data'],
      datasets: [{
        label: 'Total Maintenance Cost (₹)',
        data: sortedKeys.length ? sortedKeys.map(k => (monthMap[k].fuelCost||0) + (monthMap[k].breakdownCost||0)) : [0],
        borderColor: '#c9952c', backgroundColor: 'rgba(201,149,44,0.15)', fill: true, tension: 0.3
      }]
    },
    options: {responsive:true, maintainAspectRatio:true, plugins:{legend:{display:false}}}
  });

  destroyChart('diesel');
  const ctx2 = document.getElementById('chartDiesel');
  chartInstances.diesel = new Chart(ctx2, {
    type: 'bar',
    data: {
      labels: labels.length ? labels : ['No Data'],
      datasets: [{
        label: 'Diesel (Ltr)',
        data: sortedKeys.length ? sortedKeys.map(k => monthMap[k].diesel || 0) : [0],
        backgroundColor: '#0f2942'
      }]
    },
    options: {responsive:true, maintainAspectRatio:true, plugins:{legend:{display:false}}}
  });

  destroyChart('bdCat');
  const breakdowns = dbGet(DB_KEYS.breakdowns);
  const catTotals = {};
  breakdowns.forEach(b => { catTotals[b.category] = (catTotals[b.category]||0) + 1; });
  const catLabels = Object.keys(catTotals);
  const ctx3 = document.getElementById('chartBreakdownCat');
  chartInstances.bdCat = new Chart(ctx3, {
    type: 'doughnut',
    data: {
      labels: catLabels.length ? catLabels : ['No Data'],
      datasets: [{
        data: catLabels.length ? catLabels.map(c => catTotals[c]) : [1],
        backgroundColor: ['#0f2942','#c9952c','#1e9e5a','#d6422f','#2178b4','#e08b1a','#647184','#1c3f60']
      }]
    },
    options: {responsive:true, maintainAspectRatio:true}
  });
}

/* ===================== COST ANALYSIS PAGE ===================== */
function renderCostAnalysis() {
  const totals = computeOverallTotals();
  document.getElementById('caCostPerKM').textContent = totals.totalKmRun > 0 ? fmtCurrency(totals.totalCost / totals.totalKmRun) : '₹0';
  document.getElementById('caCostPerHMR').textContent = totals.totalKmRun > 0 ? fmtCurrency(totals.totalCost / totals.totalKmRun) : '₹0';
  document.getElementById('caCostPerTrip2').textContent = totals.totalTrips > 0 ? fmtCurrency(totals.totalCost / totals.totalTrips) : '₹0';
  document.getElementById('caCostPerTon2').textContent = totals.totalTons > 0 ? fmtCurrency(totals.totalCost / totals.totalTons) : '₹0';

  renderEquipmentCostTable('eqCostTable2Body');

  destroyChart('costBreakdown');
  const ctx1 = document.getElementById('chartCostBreakdown');
  chartInstances.costBreakdown = new Chart(ctx1, {
    type: 'pie',
    data: {
      labels: ['Fuel Cost','Spare Cost','Lubricant Cost','Labour Cost','Vendor Cost'],
      datasets: [{
        data: [totals.fuelCost, totals.spareCost, totals.lubeCost, totals.labourCost, Math.max(0,totals.breakdownCost - totals.spareCost - totals.lubeCost - totals.labourCost)],
        backgroundColor: ['#2178b4','#0f2942','#c9952c','#1e9e5a','#d6422f']
      }]
    },
    options: {responsive:true, maintainAspectRatio:true}
  });

  destroyChart('topEquipment');
  const map = computeEquipmentCostMap();
  const top5 = Object.values(map).sort((a,b) => b.totalCost - a.totalCost).slice(0,5);
  const ctx2 = document.getElementById('chartTopEquipment');
  chartInstances.topEquipment = new Chart(ctx2, {
    type: 'bar',
    data: {
      labels: top5.map(e => e.equipment),
      datasets: [{label:'Total Cost (₹)', data: top5.map(e => e.totalCost), backgroundColor: '#c9952c'}]
    },
    options: {responsive:true, maintainAspectRatio:true, indexAxis:'y', plugins:{legend:{display:false}}}
  });
}

/* ===================== FLEET KPI PAGE ===================== */
function renderFleetKpi() {
  const {availability, utilization} = computeFleetAvailability();
  document.getElementById('fkAvailability').textContent = availability.toFixed(1) + '%';
  document.getElementById('fkUtilization').textContent = utilization.toFixed(1) + '%';

  const breakdowns = dbGet(DB_KEYS.breakdowns);
  document.getElementById('fkTotalBD').textContent = breakdowns.length;

  const logs = dbGet(DB_KEYS.dailyLogs);
  // avg running hours between breakdowns (approx using total km/hmr run / number of breakdowns as proxy is unreliable;
  // instead compute average days between breakdowns per equipment, then average across fleet)
  let mtbfHours = 0;
  if (breakdowns.length > 0) {
    const totalDowntime = breakdowns.reduce((s,b) => s + (b.downtime||0), 0);
    const totalRunHours = logs.length * 12; // assume 12 hrs/shift logged as running baseline proxy
    mtbfHours = breakdowns.length > 0 ? (totalRunHours - totalDowntime) / breakdowns.length : 0;
  }
  document.getElementById('fkMTBF').textContent = fmtNum(Math.max(0,mtbfHours),1) + ' hrs';

  const tbody = document.getElementById('fleetKpiTableBody');
  const rows = FLEET.map(eq => {
    const eqLogs = logs.filter(l => l.vehicle === eq.id);
    const totalDays = eqLogs.length;
    let runningShifts = 0, totalShifts = 0, totalTrips = 0;
    eqLogs.forEach(l => {
      totalShifts += 2;
      if (l.dayStatus === 'R') runningShifts++;
      if (l.nightStatus === 'R') runningShifts++;
      totalTrips += (l.ftdTrips || 0);
    });
    const avail = totalShifts > 0 ? (runningShifts/totalShifts)*100 : 0;
    const theoreticalMax = totalDays * 4;
    const util = theoreticalMax > 0 ? Math.min(100, (totalTrips/theoreticalMax)*100) : 0;
    return {equipment: eq.id, totalDays, runningShifts, breakdownShifts: totalShifts-runningShifts, avail, totalTrips, util};
  });

  if (logs.length === 0) {
    tbody.innerHTML = '<tr><td colspan="7"><div class="empty-state"><i class="bi bi-inbox"></i>No data available yet — add daily log entries</div></td></tr>';
  } else {
    tbody.innerHTML = rows.map(r => `
      <tr>
        <td><strong>${escapeHtml(r.equipment)}</strong></td>
        <td>${r.totalDays}</td>
        <td>${r.runningShifts}</td>
        <td>${r.breakdownShifts}</td>
        <td>${r.avail.toFixed(1)}%</td>
        <td>${fmtNum(r.totalTrips)}</td>
        <td>${r.util.toFixed(1)}%</td>
      </tr>
    `).join('');
  }
}

/* ===================== SETTINGS PAGE ===================== */
function renderFleetMasterTable() {
  const tbody = document.getElementById('fleetMasterTableBody');
  tbody.innerHTML = FLEET.map(eq => `<tr><td><strong>${eq.id}</strong></td><td>${eq.category}</td></tr>`).join('');
}

/* ===================================================================
   DEMO DATA LOADER
   =================================================================== */
function loadDemoData() {
  if (!confirm('This will add sample demo data on top of existing data. Continue?')) return;

  const today = new Date();
  function dateOffset(daysAgo) {
    const d = new Date(today);
    d.setDate(d.getDate() - daysAgo);
    return d.getFullYear() + '-' + String(d.getMonth()+1).padStart(2,'0') + '-' + String(d.getDate()).padStart(2,'0');
  }

  // Daily Logs - last 10 days for first 6 tippers + excavators/loaders
  const logs = dbGet(DB_KEYS.dailyLogs);
  const demoVehicles = ['AP39UQ0095','AP39UQ0097','AP39UW9880','AP39UY4651','EX-01','LD-01'];
  let km = {};
  demoVehicles.forEach(v => km[v] = 12000 + Math.random()*5000);

  for (let day = 9; day >= 0; day--) {
    demoVehicles.forEach(v => {
      const isBreakdown = Math.random() < 0.12;
      const dayStatus = isBreakdown && Math.random() < 0.5 ? 'B' : 'R';
      const nightStatus = isBreakdown && dayStatus === 'R' ? 'B' : 'R';
      const dayTrips = dayStatus === 'R' ? Math.floor(8 + Math.random()*6) : Math.floor(Math.random()*3);
      const nightTrips = nightStatus === 'R' ? Math.floor(6 + Math.random()*5) : Math.floor(Math.random()*2);
      const openKm = km[v];
      const run = (dayStatus==='R'?40:5) + (nightStatus==='R'?35:5) + Math.random()*20;
      const closeKm = openKm + run;
      km[v] = closeKm;

      logs.push({
        id: uid(), date: dateOffset(day), vehicle: v, dayStatus, nightStatus,
        dayTrips, nightTrips, ftdTrips: dayTrips+nightTrips,
        ftdTons: (dayTrips+nightTrips) * (18 + Math.random()*4),
        sthTrips: Math.floor(Math.random()*4), obTrips: Math.floor(Math.random()*3),
        diesel: 40 + Math.random()*60,
        openKm: Math.round(openKm*10)/10, closeKm: Math.round(closeKm*10)/10, ftdRun: Math.round(run*10)/10,
        dayBdFrom: dayStatus==='B' ? '09:00':'', dayBdTo: dayStatus==='B' ? '13:30':'',
        nightBdFrom: nightStatus==='B' ? '21:00':'', nightBdTo: nightStatus==='B' ? '23:45':'',
        maintDesc: isBreakdown ? 'Hydraulic hose leak repaired' : '',
        remarks: '', totalBdHours: (dayStatus==='B'?4.5:0)+(nightStatus==='B'?2.75:0),
        createdAt: new Date().toISOString()
      });
    });
  }
  dbSet(DB_KEYS.dailyLogs, logs);

  // Hardware Store
  const hw = dbGet(DB_KEYS.hardware);
  [
    {code:'HW-001', name:'Hex Bolt M16x50', unit:'Pcs', opening:500, received:200, issued:150},
    {code:'HW-002', name:'Wheel Nut (Tipper)', unit:'Pcs', opening:300, received:100, issued:80},
    {code:'HW-003', name:'Bearing 6205', unit:'Pcs', opening:40, received:20, issued:15},
    {code:'HW-004', name:'V-Belt A-Section', unit:'Pcs', opening:25, received:10, issued:8},
    {code:'HW-005', name:'Air Filter (Tipper)', unit:'Pcs', opening:18, received:12, issued:14}
  ].forEach(m => hw.push({id: uid(), ...m, closing: m.opening+m.received-m.issued}));
  dbSet(DB_KEYS.hardware, hw);

  // Lubricants Store
  const lub = dbGet(DB_KEYS.lubricants);
  [
    {name:'Engine Oil 15W40', unit:'Ltr', opening:400, received:200, issued:180},
    {name:'Hydraulic Oil 68', unit:'Ltr', opening:300, received:100, issued:90},
    {name:'Grease EP2', unit:'Kg', opening:60, received:30, issued:25},
    {name:'Gear Oil 85W140', unit:'Ltr', opening:150, received:60, issued:55},
    {name:'Coolant', unit:'Ltr', opening:100, received:40, issued:35}
  ].forEach(m => lub.push({id: uid(), code:'', ...m, closing: m.opening+m.received-m.issued}));
  dbSet(DB_KEYS.lubricants, lub);

  // Workshop Consumables
  const ws = dbGet(DB_KEYS.workshop);
  [
    {name:'Welding Rod 3.15mm', unit:'Kg', opening:50, received:20, issued:18},
    {name:'Cotton Waste', unit:'Kg', opening:30, received:15, issued:12},
    {name:'Brake Fluid', unit:'Ltr', opening:20, received:10, issued:6},
    {name:'Cutting Disc', unit:'Pcs', opening:40, received:20, issued:22}
  ].forEach(m => ws.push({id: uid(), code:'', ...m, closing: m.opening+m.received-m.issued}));
  dbSet(DB_KEYS.workshop, ws);

  // Material Issues
  const mi = dbGet(DB_KEYS.materialIssues);
  const hwItems = dbGet(DB_KEYS.hardware);
  if (hwItems.length) {
    mi.push({id: uid(), date: dateOffset(2), equipment:'AP39UQ0095', storeType:'hardware', materialId: hwItems[0].id, materialName: hwItems[0].name, unit: hwItems[0].unit, qty: 8, purpose:'Wheel bolt replacement', issuedBy:'Krishna', createdAt: new Date().toISOString()});
  }
  dbSet(DB_KEYS.materialIssues, mi);

  // Service Records
  const sv = dbGet(DB_KEYS.serviceRecords);
  [
    {equipment:'AP39UQ0095', serviceType:'500', lastKm: 12500, currentKm: 12950},
    {equipment:'AP39UQ0097', serviceType:'1000', lastKm: 11000, currentKm: 11920},
    {equipment:'AP39UW9880', serviceType:'2000', lastKm: 9000, currentKm: 9300},
    {equipment:'EX-01', serviceType:'500', lastKm: 4200, currentKm: 4550},
    {equipment:'LD-01', serviceType:'1000', lastKm: 3000, currentKm: 3990}
  ].forEach(r => {
    const interval = SERVICE_INTERVALS[r.serviceType];
    const nextDue = r.lastKm + interval;
    const balance = nextDue - r.currentKm;
    sv.push({id: uid(), equipment:r.equipment, category:getFleetCategory(r.equipment), serviceType:r.serviceType,
      lastDate: dateOffset(20), lastKm:r.lastKm, currentKm:r.currentKm, nextDue, balance, status: computeServiceStatus(balance)});
  });
  dbSet(DB_KEYS.serviceRecords, sv);

  // Breakdowns
  const bd = dbGet(DB_KEYS.breakdowns);
  const bdDate1 = dateOffset(3), bdDate2 = dateOffset(6);
  bd.push({id: uid(), date: bdDate1, equipment:'AP39UW9880', category:'Hydraulic', problem:'Hydraulic hose burst causing oil leak',
    rootCause:'Hose wear and tear', action:'Replaced hydraulic hose and refilled oil',
    start: bdDate1+'T09:00', end: bdDate1+'T13:30', downtime: 4.5,
    parts:'Hydraulic hose 1/2" x 2m', mechanic:'Siva Prasad', vendor:'',
    spareCost: 2400, lubeCost: 800, labourCost: 600, vendorCost: 0, totalCost: 3800, createdAt: new Date().toISOString()});
  bd.push({id: uid(), date: bdDate2, equipment:'EX-02', category:'Engine', problem:'Engine overheating',
    rootCause:'Radiator clogged', action:'Cleaned radiator, replaced coolant',
    start: bdDate2+'T21:00', end: bdDate2+'T23:45', downtime: 2.75,
    parts:'Coolant 20L', mechanic:'Vijay', vendor:'AP Diesel Services',
    spareCost: 0, lubeCost: 1200, labourCost: 500, vendorCost: 1800, totalCost: 3500, createdAt: new Date().toISOString()});
  dbSet(DB_KEYS.breakdowns, bd);

  showToast('Demo data loaded successfully.');
  setTimeout(() => location.reload(), 1000);
}
