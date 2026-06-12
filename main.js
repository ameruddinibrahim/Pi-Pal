// ─── State ───────────────────────────────────────────────────────────────────

let state = {
  clusters: [],
  reports: [],
  activeReportId: null,
  activeTab: 'clusters'
};

const STATUS = ['todo', 'in-progress', 'done'];
const STATUS_LABELS = { 'todo': 'To Do', 'in-progress': 'In Progress', 'done': 'Done' };

function uid() {
  return Math.random().toString(36).slice(2, 9);
}

// ─── Persistence ─────────────────────────────────────────────────────────────

function save() {
  localStorage.setItem('seo-toolkit', JSON.stringify(state));
}

function load() {
  const raw = localStorage.getItem('seo-toolkit');
  if (raw) {
    state = { ...state, ...JSON.parse(raw) };
    return;
  }
  // Migrate data from old cluster-only version
  const legacy = localStorage.getItem('kco-state');
  if (legacy) {
    const old = JSON.parse(legacy);
    state.clusters = old.clusters || [];
    save();
  }
}

// ─── Tab Management ───────────────────────────────────────────────────────────

function switchTab(tab) {
  state.activeTab = tab;
  document.getElementById('view-clusters').style.display = tab === 'clusters' ? 'block' : 'none';
  document.getElementById('view-reports').style.display = tab === 'reports' ? 'block' : 'none';
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tab);
  });
  renderHeaderActions();
}

function renderHeaderActions() {
  const el = document.getElementById('header-actions');
  if (state.activeTab === 'clusters') {
    el.innerHTML = `
      <button id="export-btn" class="btn-secondary">Export CSV</button>
      <button id="add-cluster-btn" class="btn-primary">+ New Cluster</button>
    `;
    document.getElementById('export-btn').addEventListener('click', exportCSV);
    document.getElementById('add-cluster-btn').addEventListener('click', addCluster);
  } else {
    el.innerHTML = `<label class="btn-primary" for="file-input" style="cursor:pointer">Import .xlsx</label>`;
  }
}

// ─── Cluster Actions ──────────────────────────────────────────────────────────

function addCluster() {
  state.clusters.push({ id: uid(), name: 'New Cluster', targetUrl: '', keywords: [] });
  save();
  renderClusters();
}

function removeCluster(id) {
  state.clusters = state.clusters.filter(c => c.id !== id);
  save();
  renderClusters();
}

function updateCluster(id, field, value) {
  const cluster = state.clusters.find(c => c.id === id);
  if (cluster) cluster[field] = value;
  save();
}

function addKeyword(clusterId, text) {
  if (!text.trim()) return false;
  const cluster = state.clusters.find(c => c.id === clusterId);
  if (cluster) cluster.keywords.push({ id: uid(), text: text.trim(), status: 'todo' });
  save();
  renderClusters();
  return true;
}

function removeKeyword(clusterId, keywordId) {
  const cluster = state.clusters.find(c => c.id === clusterId);
  if (cluster) cluster.keywords = cluster.keywords.filter(k => k.id !== keywordId);
  save();
  renderClusters();
}

function cycleStatus(clusterId, keywordId) {
  const cluster = state.clusters.find(c => c.id === clusterId);
  const kw = cluster && cluster.keywords.find(k => k.id === keywordId);
  if (kw) kw.status = STATUS[(STATUS.indexOf(kw.status) + 1) % STATUS.length];
  save();
  renderClusters();
}

function exportCSV() {
  const rows = [['Keyword', 'Cluster', 'Target URL', 'Status']];
  state.clusters.forEach(c => {
    c.keywords.forEach(kw => {
      rows.push([`"${kw.text}"`, `"${c.name}"`, `"${c.targetUrl}"`, STATUS_LABELS[kw.status]]);
    });
  });
  if (rows.length === 1) { alert('No keywords to export.'); return; }
  const csv = rows.map(r => r.join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = 'keyword-clusters.csv';
  a.click();
}

// ─── Report Actions ───────────────────────────────────────────────────────────

function parseXLSXFile(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    const workbook = XLSX.read(new Uint8Array(e.target.result), { type: 'array' });
    let lastId = null;

    workbook.SheetNames.forEach(sheetName => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1 });
      if (rows.length < 2) return;

      const results = rows.slice(1)
        .filter(row => row[0] != null)
        .map(row => ({
          position:   row[0] || null,
          type:       row[1] || '',
          domain:     row[2] || '',
          url:        row[3] || '',
          pageAS:     row[4] || 0,
          refDomains: row[5] || 0,
          backlinks:  row[6] || 0,
          traffic:    row[7] || 0,
          urlKeywords: row[8] || 0,
          serpFeature: row[9] || null
        }));

      const report = {
        id: uid(),
        keyword: sheetName,
        importedAt: new Date().toISOString().split('T')[0],
        results
      };

      const existing = state.reports.findIndex(r => r.keyword === sheetName);
      if (existing >= 0) {
        state.reports[existing] = report;
        lastId = report.id;
      } else {
        state.reports.unshift(report);
        lastId = report.id;
      }
    });

    if (lastId) state.activeReportId = lastId;
    save();
    switchTab('reports');
    renderReports();
    showToast('Report imported successfully');
  };
  reader.readAsArrayBuffer(file);
}

function deleteReport(id) {
  state.reports = state.reports.filter(r => r.id !== id);
  if (state.activeReportId === id) {
    state.activeReportId = state.reports[0]?.id || null;
  }
  save();
  renderReports();
}

function assignKeywordToCluster(keyword, clusterId) {
  const cluster = state.clusters.find(c => c.id === clusterId);
  if (!cluster) return false;
  const exists = cluster.keywords.some(k => k.text.toLowerCase() === keyword.toLowerCase());
  if (!exists) {
    cluster.keywords.push({ id: uid(), text: keyword, status: 'todo' });
    save();
    return true;
  }
  return false;
}

// ─── Report Analysis ──────────────────────────────────────────────────────────

function analyzeReport(report) {
  const organic = report.results.filter(r => r.type === 'Organic');
  const top10 = organic.filter(r => r.position <= 10);

  const avg = (arr, key) => arr.length
    ? Math.round(arr.reduce((s, r) => s + (r[key] || 0), 0) / arr.length)
    : 0;

  const serpFeatures = [...new Set([
    ...report.results.filter(r => r.type !== 'Organic').map(r => r.type),
    ...report.results.filter(r => r.serpFeature).map(r => r.serpFeature)
  ])].filter(Boolean);

  return {
    organic,
    top10,
    serpFeatures,
    avgPageAS:    avg(top10, 'pageAS'),
    avgBacklinks: avg(top10, 'backlinks'),
    avgTraffic:   avg(top10, 'traffic'),
    avgRefDomains: avg(top10, 'refDomains')
  };
}

// ─── Render Clusters ──────────────────────────────────────────────────────────

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function renderCluster(c) {
  const total = c.keywords.length;
  const done = c.keywords.filter(k => k.status === 'done').length;
  return `
    <div class="cluster-card" data-id="${c.id}">
      <div class="cluster-header">
        <input class="cluster-name" value="${esc(c.name)}"
          data-action="update-name" data-cluster-id="${c.id}">
        <button class="btn-icon" data-action="remove-cluster"
          data-cluster-id="${c.id}" title="Delete cluster">✕</button>
      </div>
      <input class="cluster-url" value="${esc(c.targetUrl)}"
        placeholder="Target page URL..."
        data-action="update-url" data-cluster-id="${c.id}">
      ${total > 0 ? `<div class="cluster-stats">${total} keyword${total !== 1 ? 's' : ''} · ${done} done</div>` : ''}
      <div class="keywords-list">
        ${c.keywords.map(kw => `
          <div class="keyword-row">
            <span class="kw-text">${esc(kw.text)}</span>
            <button class="status-badge status-${kw.status}"
              data-action="cycle-status"
              data-cluster-id="${c.id}"
              data-keyword-id="${kw.id}">${STATUS_LABELS[kw.status]}</button>
            <button class="btn-icon"
              data-action="remove-keyword"
              data-cluster-id="${c.id}"
              data-keyword-id="${kw.id}" title="Remove">✕</button>
          </div>
        `).join('')}
      </div>
      <div class="add-kw-form">
        <input class="kw-input" placeholder="Add a keyword…" data-cluster-id="${c.id}">
        <button class="btn-add-kw" data-action="add-keyword" data-cluster-id="${c.id}">Add</button>
      </div>
    </div>
  `;
}

function renderClusters() {
  const grid = document.getElementById('clusters-grid');
  const empty = document.getElementById('empty-state');
  if (state.clusters.length === 0) {
    grid.innerHTML = '';
    empty.style.display = 'flex';
  } else {
    empty.style.display = 'none';
    grid.innerHTML = state.clusters.map(renderCluster).join('');
  }
}

// ─── Render Reports ───────────────────────────────────────────────────────────

function renderReportCard(report) {
  const { avgPageAS, organic } = analyzeReport(report);
  const topDomain = organic[0]?.domain || '—';
  const isActive = state.activeReportId === report.id;
  return `
    <div class="report-card ${isActive ? 'active' : ''}" data-report-id="${report.id}">
      <div class="report-card-header">
        <span class="report-keyword">${esc(report.keyword)}</span>
        <button class="btn-icon" data-action="delete-report"
          data-report-id="${report.id}" title="Delete">✕</button>
      </div>
      <div class="report-card-meta">
        <span>#1: ${esc(topDomain)}</span>
        <span>AS: ${avgPageAS}</span>
      </div>
      <div class="report-card-date">${report.importedAt}</div>
    </div>
  `;
}

function renderReportDetail(report) {
  const { organic, top10, serpFeatures, avgPageAS, avgBacklinks, avgTraffic, avgRefDomains } = analyzeReport(report);

  const clusterOptions = state.clusters
    .map(c => `<option value="${c.id}">${esc(c.name)}</option>`)
    .join('');

  return `
    <div class="report-detail-inner">
      <div class="report-detail-header">
        <div>
          <h2>${esc(report.keyword)}</h2>
          <span class="report-date">Imported ${report.importedAt} · ${organic.length} organic results</span>
        </div>
        ${state.clusters.length > 0 ? `
          <div class="assign-cluster">
            <select id="assign-select">
              <option value="">+ Add keyword to cluster…</option>
              ${clusterOptions}
            </select>
          </div>
        ` : ''}
      </div>

      ${serpFeatures.length > 0 ? `
        <div class="serp-features">
          ${serpFeatures.map(f => `<span class="feature-badge">${esc(f)}</span>`).join('')}
        </div>
      ` : ''}

      <div class="benchmark-grid">
        <div class="benchmark-card">
          <div class="benchmark-value">${avgPageAS}</div>
          <div class="benchmark-label">Avg Page AS<br>Top 10</div>
        </div>
        <div class="benchmark-card">
          <div class="benchmark-value">${avgBacklinks.toLocaleString()}</div>
          <div class="benchmark-label">Avg Backlinks<br>Top 10</div>
        </div>
        <div class="benchmark-card">
          <div class="benchmark-value">${avgRefDomains.toLocaleString()}</div>
          <div class="benchmark-label">Avg Ref. Domains<br>Top 10</div>
        </div>
        <div class="benchmark-card">
          <div class="benchmark-value">${avgTraffic.toLocaleString()}</div>
          <div class="benchmark-label">Avg Traffic<br>Top 10</div>
        </div>
      </div>

      <div class="competitors-table-wrap">
        <table class="competitors-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Domain</th>
              <th>Page AS</th>
              <th>Traffic</th>
              <th>Backlinks</th>
              <th>Ref. Domains</th>
              <th>Feature</th>
            </tr>
          </thead>
          <tbody>
            ${organic.slice(0, 20).map(r => `
              <tr class="${r.position <= 10 ? 'top-10' : ''}">
                <td>${r.position}</td>
                <td class="domain-cell">
                  <a href="${esc(r.url)}" target="_blank" rel="noopener">${esc(r.domain)}</a>
                </td>
                <td>${r.pageAS || '—'}</td>
                <td>${r.traffic ? r.traffic.toLocaleString() : '—'}</td>
                <td>${r.backlinks ? r.backlinks.toLocaleString() : '—'}</td>
                <td>${r.refDomains ? r.refDomains.toLocaleString() : '—'}</td>
                <td>${r.serpFeature ? `<span class="feature-badge small">${esc(r.serpFeature)}</span>` : ''}</td>
              </tr>
            `).join('')}
          </tbody>
        </table>
      </div>
    </div>
  `;
}

function renderReports() {
  const listEl  = document.getElementById('reports-list');
  const emptyEl = document.getElementById('reports-empty');
  const detailEl = document.getElementById('report-detail');

  if (state.reports.length === 0) {
    listEl.innerHTML = '';
    emptyEl.style.display = 'flex';
    detailEl.innerHTML = '<div id="report-placeholder"><p>Import a .xlsx file to see your SERP analysis</p></div>';
    return;
  }

  emptyEl.style.display = 'none';
  listEl.innerHTML = state.reports.map(renderReportCard).join('');

  const active = state.reports.find(r => r.id === state.activeReportId) || state.reports[0];
  state.activeReportId = active.id;
  detailEl.innerHTML = renderReportDetail(active);

  const assignSelect = document.getElementById('assign-select');
  if (assignSelect) {
    assignSelect.addEventListener('change', () => {
      const clusterId = assignSelect.value;
      if (!clusterId) return;
      const cluster = state.clusters.find(c => c.id === clusterId);
      const added = assignKeywordToCluster(active.keyword, clusterId);
      assignSelect.value = '';
      showToast(added
        ? `"${active.keyword}" added to "${cluster.name}"`
        : `Already in "${cluster.name}"`
      );
    });
  }
}

// ─── Toast ────────────────────────────────────────────────────────────────────

function showToast(msg) {
  const toast = document.getElementById('toast');
  toast.textContent = msg;
  toast.classList.add('visible');
  clearTimeout(toast._timer);
  toast._timer = setTimeout(() => toast.classList.remove('visible'), 2500);
}

// ─── Events ───────────────────────────────────────────────────────────────────

document.addEventListener('click', e => {
  const action = e.target.dataset.action;
  const cid = e.target.dataset.clusterId;
  const kid = e.target.dataset.keywordId;
  const rid = e.target.dataset.reportId;

  if (action === 'remove-cluster')  { removeCluster(cid); return; }
  if (action === 'remove-keyword')  { removeKeyword(cid, kid); return; }
  if (action === 'cycle-status')    { cycleStatus(cid, kid); return; }
  if (action === 'delete-report')   { deleteReport(rid); return; }
  if (action === 'add-keyword') {
    const input = document.querySelector(`.kw-input[data-cluster-id="${cid}"]`);
    if (addKeyword(cid, input.value)) {
      document.querySelector(`.kw-input[data-cluster-id="${cid}"]`)?.focus();
    }
    return;
  }

  // Click on a report card (not on an action button)
  const card = e.target.closest('.report-card');
  if (card && !action) {
    state.activeReportId = card.dataset.reportId;
    renderReports();
  }
});

document.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.classList.contains('kw-input')) {
    const cid = e.target.dataset.clusterId;
    if (addKeyword(cid, e.target.value)) {
      document.querySelector(`.kw-input[data-cluster-id="${cid}"]`)?.focus();
    }
  }
});

document.addEventListener('change', e => {
  const action = e.target.dataset.action;
  const cid = e.target.dataset.clusterId;
  if (action === 'update-name') updateCluster(cid, 'name', e.target.value);
  if (action === 'update-url')  updateCluster(cid, 'targetUrl', e.target.value);
});

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => switchTab(btn.dataset.tab));
});

const fileInput = document.getElementById('file-input');
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) {
    parseXLSXFile(e.target.files[0]);
    e.target.value = '';
  }
});

// Drag and drop
const reportsView = document.getElementById('view-reports');
reportsView.addEventListener('dragover', e => {
  e.preventDefault();
  reportsView.classList.add('drag-over');
});
reportsView.addEventListener('dragleave', () => reportsView.classList.remove('drag-over'));
reportsView.addEventListener('drop', e => {
  e.preventDefault();
  reportsView.classList.remove('drag-over');
  const file = e.dataTransfer.files[0];
  if (file?.name.endsWith('.xlsx')) parseXLSXFile(file);
});

document.getElementById('empty-add-btn').addEventListener('click', addCluster);

// ─── Init ─────────────────────────────────────────────────────────────────────

load();
renderClusters();
renderReports();
renderHeaderActions();
switchTab(state.activeTab || 'clusters');
