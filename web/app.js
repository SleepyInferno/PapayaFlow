/* =========================================================================
   PapayaFlow — app.js
   Complete dashboard script: upload flow, theme toggle, skeleton,
   stat cards, sort, expandable rows, threshold flags.
   ES5-compatible syntax throughout (var, function declarations).
   ========================================================================= */

/* ---- 1. XSS prevention utility ---------------------------------------- */
function escapeHtml(str) {
  if (str == null) return '';
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

/* ---- 2. Number formatting utilities (UI-SPEC Section 7.5) -------------- */
function formatPages(n)  { return (n || 0).toLocaleString(); }
function formatCost(n)   { return '$' + (n || 0).toFixed(2).replace(/\B(?=(\d{3})+(?!\d))/g, ','); }
function formatPct(n)    { return Math.round(n || 0) + '%'; }
function formatInt(n)    { return (n || 0).toString(); }

/* ---- 3. Module-level state --------------------------------------------- */
var deptData = [];
var sortState = { col: 'TotalCost', dir: 'desc' };
var currentThreshold = 30;

/* ---- 4. State machine show/hide functions ------------------------------ */

function showUpload() {
  document.getElementById('upload-section').style.display = '';
  document.getElementById('dashboard-section').style.display = 'none';
  document.getElementById('loading-skeleton').style.display = 'none';
  document.getElementById('new-report-btn').style.display = 'none';
  var banner = document.getElementById('status-banner');
  banner.textContent = '';
  banner.className = 'status-banner';
  banner.style.display = 'none';
}

function showSkeleton() {
  document.getElementById('upload-section').style.display = 'none';
  document.getElementById('loading-skeleton').style.display = 'block';
  document.getElementById('dashboard-section').style.display = 'none';
  document.getElementById('new-report-btn').style.display = 'none';
}

function showDashboard(data) {
  document.getElementById('upload-section').style.display = 'none';
  document.getElementById('loading-skeleton').style.display = 'none';
  document.getElementById('dashboard-section').style.display = 'block';
  document.getElementById('new-report-btn').style.display = '';
  renderDashboard(data);
}

function showError(message) {
  document.getElementById('upload-section').style.display = '';
  document.getElementById('loading-skeleton').style.display = 'none';
  document.getElementById('dashboard-section').style.display = 'none';
  document.getElementById('new-report-btn').style.display = 'none';
  var banner = document.getElementById('status-banner');
  banner.className = 'status-banner error';
  banner.innerHTML = (message ? escapeHtml(message) + '<br>' : '') + 'Check your files and try again.';
  banner.style.display = '';
}

/* ---- 5-10. Theme toggle init ------------------------------------------- */

function initThemeToggle() {
  var btn = document.getElementById('theme-toggle');
  function updateButton() {
    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    btn.textContent = isDark ? '\u2600' : '\uD83C\uDF19';
    btn.setAttribute('aria-label', isDark ? 'Switch to light mode' : 'Switch to dark mode');
  }
  updateButton();
  btn.addEventListener('click', function() {
    var isDark = document.documentElement.getAttribute('data-theme') === 'dark';
    var next = isDark ? 'light' : 'dark';
    document.documentElement.setAttribute('data-theme', next);
    localStorage.setItem('papayaflow-theme', next);
    updateButton();
  });
}

/* ---- Dashboard rendering functions (Task 2) ---------------------------- */

function isHighColor(dept, threshold) {
  var t = parseInt(threshold, 10);
  if (isNaN(t) || t < 0 || t >= 100) return false;
  return dept.PctColor >= t;
}

function createSortComparator(col, dir) {
  var d = dir === 'asc' ? 1 : -1;
  return function(a, b) {
    var va = a[col], vb = b[col];
    if (va == null) va = (typeof vb === 'string') ? '' : 0;
    if (vb == null) vb = (typeof va === 'string') ? '' : 0;
    if (typeof va === 'string') return va.localeCompare(vb) * d;
    return (va < vb ? -1 : va > vb ? 1 : 0) * d;
  };
}

function updateSortArrows() {
  var ths = document.querySelectorAll('#dept-table thead th[data-col]');
  for (var i = 0; i < ths.length; i++) {
    var th = ths[i];
    var existing = th.querySelector('.sort-arrow');
    if (existing) th.removeChild(existing);
    th.removeAttribute('aria-sort');

    if (th.getAttribute('data-col') === sortState.col) {
      var arrow = document.createElement('span');
      arrow.className = 'sort-arrow';
      arrow.textContent = sortState.dir === 'asc' ? ' \u25b2' : ' \u25bc';
      th.appendChild(arrow);
      th.setAttribute('aria-sort', sortState.dir === 'asc' ? 'ascending' : 'descending');
    } else {
      th.setAttribute('aria-sort', 'none');
    }
  }
}

function renderUserTable(users) {
  var html = '<table class="user-table">' +
    '<thead><tr>' +
    '<th>User</th><th>Pages</th><th>Print</th><th>Copy</th>' +
    '<th>B&amp;W</th><th>Color</th><th>Color %</th>' +
    '<th>1-Sided</th><th>2-Sided</th><th>Scans</th><th>Fax</th>' +
    '<th>Jobs</th><th>Cost</th>' +
    '</tr></thead><tbody>';

  for (var i = 0; i < users.length; i++) {
    var u = users[i];
    var userPrint = (u.BW || 0) + (u.Color || 0);
    var userPctColor = userPrint > 0 ? Math.round((u.Color / userPrint) * 100) : 0;

    html += '<tr>' +
      '<td>' + escapeHtml(u.UPN) + '</td>' +
      '<td>' + formatPages(u.Pages) + '</td>' +
      '<td>' + formatPages(u.Print) + '</td>' +
      '<td>' + formatPages(u.Copy) + '</td>' +
      '<td>' + formatPages(u.BW) + '</td>' +
      '<td>' + formatPages(u.Color) + '</td>' +
      '<td>' + userPctColor + '%</td>' +
      '<td>' + formatInt(u.OneSided) + '</td>' +
      '<td>' + formatInt(u.TwoSided) + '</td>' +
      '<td>' + formatInt(u.Scans) + '</td>' +
      '<td>' + formatInt(u.Fax) + '</td>' +
      '<td>' + formatInt(u.Jobs) + '</td>' +
      '<td>' + formatCost(u.Cost) + '</td>' +
    '</tr>';
  }

  html += '</tbody></table>';
  return html;
}

function renderDeptRow(dept, index, threshold) {
  var flagged = isHighColor(dept, threshold);
  var rowClass = 'dept-row' + (flagged ? ' high-color' : '');
  var pctCell = formatPct(dept.PctColor);
  var pctCellClass = flagged ? ' class="cost-cell"' : '';
  var badge = flagged ? ' <span class="badge-high-color">HIGH COLOR</span>' : '';

  var deptRow = '<tr class="' + rowClass + '" data-dept-index="' + index + '" aria-expanded="false">' +
    '<td class="chevron-cell"><span class="chevron">&#9654;</span></td>' +
    '<td>' + escapeHtml(dept.Name) + '</td>' +
    '<td>' + formatInt(dept.ActiveUsers) + '</td>' +
    '<td>' + formatPages(dept.TotalPages) + '</td>' +
    '<td>' + formatPages(dept.TotalPrint) + '</td>' +
    '<td>' + formatPages(dept.TotalCopy) + '</td>' +
    '<td>' + formatPages(dept.TotalBW) + '</td>' +
    '<td>' + formatPages(dept.TotalColor) + '</td>' +
    '<td' + pctCellClass + '>' + pctCell + badge + '</td>' +
    '<td>' + formatInt(dept.TotalOneSided) + '</td>' +
    '<td>' + formatInt(dept.TotalTwoSided) + '</td>' +
    '<td>' + formatInt(dept.TotalScans) + '</td>' +
    '<td>' + formatInt(dept.TotalFax) + '</td>' +
    '<td>' + formatInt(dept.TotalJobs) + '</td>' +
    '<td>' + formatCost(dept.TotalCost) + '</td>' +
  '</tr>';

  var detailRow = '<tr class="detail-row" data-dept-index="' + index + '">' +
    '<td colspan="15">' + renderUserTable(dept.Users || []) + '</td>' +
  '</tr>';

  return deptRow + detailRow;
}

function renderTable() {
  var threshold = currentThreshold;
  var sorted = deptData.slice().sort(createSortComparator(sortState.col, sortState.dir));
  var html = '';
  for (var i = 0; i < sorted.length; i++) {
    html += renderDeptRow(sorted[i], i, threshold);
  }
  document.getElementById('dept-tbody').innerHTML = html;
  updateSortArrows();
}

function renderDashboard(data) {
  deptData = data.Departments || [];

  var dr = data.DateRange;
  var dateStr = (dr && dr.From && dr.To) ? (dr.From + ' \u2013 ' + dr.To) : '';
  document.getElementById('date-range').textContent = dateStr;

  var org = data.OrgTotals || {};
  document.getElementById('stat-total-pages').textContent  = formatPages(org.TotalPages);
  document.getElementById('stat-total-cost').textContent   = formatCost(org.TotalCost);
  document.getElementById('stat-active-users').textContent = formatInt(org.ActiveUsers);
  document.getElementById('stat-color-pct').textContent    = formatPct(org.PctColor);
  document.getElementById('stat-departments').textContent  = formatInt(deptData.length);

  sortState = { col: 'TotalCost', dir: 'desc' };
  currentThreshold = parseInt(document.getElementById('color-threshold').value, 10);
  if (isNaN(currentThreshold)) currentThreshold = 30;

  document.getElementById('empty-state').style.display = deptData.length === 0 ? '' : 'none';

  renderTable();
}

/* ---- CSV export --------------------------------------------------------- */

function csvEscape(val) {
  if (val == null) return '';
  var s = String(val);
  // If value contains comma, double-quote, or newline, wrap in double-quotes
  // and escape internal double-quotes by doubling them
  if (s.indexOf(',') !== -1 || s.indexOf('"') !== -1 || s.indexOf('\n') !== -1) {
    return '"' + s.replace(/"/g, '""') + '"';
  }
  return s;
}

function buildCsv(departments) {
  var rows = [];

  // Header row
  rows.push([
    'Type', 'Department', 'User',
    'ActiveUsers', 'TotalPages', 'TotalPrint', 'TotalCopy',
    'TotalBW', 'TotalColor', 'PctColor',
    'TotalOneSided', 'TotalTwoSided', 'TotalScans', 'TotalFax',
    'TotalJobs', 'TotalCost'
  ].map(csvEscape).join(','));

  for (var i = 0; i < departments.length; i++) {
    var d = departments[i];

    // Department summary row — Type=Department, User column left blank
    rows.push([
      'Department',
      d.Name,
      '',
      d.ActiveUsers,
      d.TotalPages,
      d.TotalPrint,
      d.TotalCopy,
      d.TotalBW,
      d.TotalColor,
      d.PctColor,
      d.TotalOneSided,
      d.TotalTwoSided,
      d.TotalScans,
      d.TotalFax,
      d.TotalJobs,
      d.TotalCost
    ].map(csvEscape).join(','));

    // Per-user detail rows — Type=User, Department column repeated
    var users = d.Users || [];
    for (var j = 0; j < users.length; j++) {
      var u = users[j];
      var userPrint = (u.BW || 0) + (u.Color || 0);
      var userPctColor = userPrint > 0 ? (Math.round((u.Color / userPrint) * 100 * 10) / 10) : 0;
      rows.push([
        'User',
        d.Name,
        u.UPN,
        '',
        u.Pages,
        u.Print,
        u.Copy,
        u.BW,
        u.Color,
        userPctColor,
        u.OneSided,
        u.TwoSided,
        u.Scans,
        u.Fax,
        u.Jobs,
        u.Cost
      ].map(csvEscape).join(','));
    }
  }

  return rows.join('\r\n');
}

/* ---- DOMContentLoaded bootstrap --------------------------------------- */

document.addEventListener('DOMContentLoaded', function() {

  initThemeToggle();
  showUpload();

  /* File input change handlers — enable/disable process button */
  function updateProcessBtn() {
    var hasPdf = document.getElementById('pdf-input').files.length > 0;
    var hasCsv = document.getElementById('csv-input').files.length > 0;
    document.getElementById('process-btn').disabled = !(hasPdf && hasCsv);
  }
  document.getElementById('pdf-input').addEventListener('change', updateProcessBtn);
  document.getElementById('csv-input').addEventListener('change', updateProcessBtn);

  /* Upload form submit handler */
  document.getElementById('upload-form').addEventListener('submit', async function(e) {
    e.preventDefault();
    var pdfFile = document.getElementById('pdf-input').files[0];
    var csvFile = document.getElementById('csv-input').files[0];
    if (!pdfFile || !csvFile) { alert('Select both files'); return; }

    var fd = new FormData();
    fd.append('pdf', pdfFile);
    fd.append('csv', csvFile);

    var banner = document.getElementById('status-banner');
    banner.className = 'status-banner processing';
    banner.textContent = 'Processing your files\u2026';
    banner.style.display = '';

    showSkeleton();

    try {
      var res = await fetch('/process', { method: 'POST', body: fd });
      var data = await res.json();
      if (!res.ok) {
        showError(data.Error || data.error || 'Server returned an error.');
        return;
      }
      showDashboard(data);
    } catch (err) {
      showError(err.message || 'Network error.');
    }
  });

  /* New Report button — immediate reset, no confirmation dialog */
  document.getElementById('new-report-btn').addEventListener('click', function() {
    deptData = [];
    sortState = { col: 'TotalCost', dir: 'desc' };
    currentThreshold = 30;
    document.getElementById('color-threshold').value = '30';
    document.getElementById('pdf-input').value = '';
    document.getElementById('csv-input').value = '';
    document.getElementById('process-btn').disabled = true;
    showUpload();
  });

  /* Drag-and-drop on upload area */
  var uploadArea = document.getElementById('upload-area');
  uploadArea.addEventListener('dragover', function(e) {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
  });
  uploadArea.addEventListener('dragenter', function(e) {
    e.preventDefault();
    uploadArea.classList.add('drag-over');
  });
  uploadArea.addEventListener('dragleave', function() {
    uploadArea.classList.remove('drag-over');
  });
  uploadArea.addEventListener('drop', function(e) {
    e.preventDefault();
    uploadArea.classList.remove('drag-over');
    var files = e.dataTransfer.files;
    for (var i = 0; i < files.length; i++) {
      var f = files[i];
      var name = f.name.toLowerCase();
      if (name.endsWith('.pdf')) {
        var dt = new DataTransfer();
        dt.items.add(f);
        document.getElementById('pdf-input').files = dt.files;
      } else if (name.endsWith('.csv')) {
        var dt2 = new DataTransfer();
        dt2.items.add(f);
        document.getElementById('csv-input').files = dt2.files;
      }
    }
    updateProcessBtn();
  });

  /* Sort click handler — event delegation on thead */
  document.querySelector('#dept-table thead').addEventListener('click', function(e) {
    var th = e.target.closest('th[data-col]');
    if (!th) return;
    var col = th.getAttribute('data-col');
    if (sortState.col === col) {
      sortState.dir = sortState.dir === 'asc' ? 'desc' : 'asc';
    } else {
      sortState.col = col;
      sortState.dir = 'asc';
    }
    renderTable();
  });

  /* Expand/collapse click handler — event delegation on tbody */
  document.getElementById('dept-tbody').addEventListener('click', function(e) {
    var row = e.target.closest('tr.dept-row');
    if (!row) return;
    var idx = row.getAttribute('data-dept-index');
    var detail = document.querySelector('tr.detail-row[data-dept-index="' + idx + '"]');
    if (!detail) return;
    detail.classList.toggle('visible');
    row.classList.toggle('expanded');
    row.setAttribute('aria-expanded', row.classList.contains('expanded') ? 'true' : 'false');
  });

  /* Color threshold input handler — recalculates on every keystroke */
  document.getElementById('color-threshold').addEventListener('input', function() {
    var val = parseInt(this.value, 10);
    currentThreshold = (isNaN(val) || val < 0 || val >= 100) ? 100 : val;
    renderTable();
  });

  /* Export CSV button — builds CSV from in-memory deptData and triggers download */
  document.getElementById('export-btn').addEventListener('click', function() {
    if (!deptData || deptData.length === 0) return;
    var csv = buildCsv(deptData);
    var blob = new Blob([csv], { type: 'text/csv;charset=utf-8;' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = 'papayaflow-export.csv';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
  });

});
