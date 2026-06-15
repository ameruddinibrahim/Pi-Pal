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

// ─── Config ───────────────────────────────────────────────────────────────────
// All thresholds in one place for easy tuning.

const CONFIG = {
  weakASThreshold: 22,   // Page AS strictly below this = "weak page" (tuned from 25 to match real-world validation)
  lowRDThreshold:  20,   // Ref.Domains below this = "low-link page"
  winnableMin:     70,   // Score >= this → Winnable
  competitiveMin:  40,   // Score >= this → Competitive (else Hard)
  skewedDiff:      15,   // avg − median > this → flag avg as skewed
  ugcDomains:      ['reddit.com', 'quora.com', 'medium.com', 'stackexchange.com', 'stackoverflow.com'],
  myDomain:        'vantagemarkets',

  // Industry-average organic CTR by position (desktop). Tune as your own GSC data comes in.
  ctrCurve: { 1: 0.28, 2: 0.15, 3: 0.11, 4: 0.08, 5: 0.07, 6: 0.05, 7: 0.04, 8: 0.035, 9: 0.03, 10: 0.025 },
  // When an AI Overview is present it pushes organic results down and absorbs clicks.
  // Multiplier applied to every organic CTR (0.65 = ~35% fewer clicks).
  aiOverviewCTRMultiplier: 0.65
};

const BUCKET_META = {
  winnable:    { label: '🟢 Winnable',    cls: 'winnable',    verdict: 'Quality content can rank here without a big backlink profile. Prioritise.' },
  competitive: { label: '🟡 Competitive', cls: 'competitive', verdict: 'Page 1 + AI Overview realistic; top 3 will take time and links.' },
  hard:        { label: '🔴 Hard',        cls: 'hard',        verdict: 'Authority-dominated. Needs backlinks and patience, or target long-tail variants instead.' }
};

const FORMAT_TIPS = {
  'AI Overview':      'Add a crisp 2–3 sentence definition at the top of the page to compete for the AI Overview snippet.',
  'People also ask':  'Add an FAQ section with FAQ schema markup to capture People Also Ask slots.',
  'Video':            'Embed a short explainer video — YouTube videos can rank independently in the Video carousel.',
  'Video carousel':   'Embed a short explainer video — YouTube videos can rank independently in the Video carousel.'
};

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
  const legacy = localStorage.getItem('kco-state');
  if (legacy) {
    state.clusters = JSON.parse(legacy).clusters || [];
    save();
  }
}

// ─── Tab Management ───────────────────────────────────────────────────────────

function switchTab(tab) {
  state.activeTab = tab;
  document.getElementById('view-clusters').style.display = tab === 'clusters' ? 'block' : 'none';
  document.getElementById('view-reports').style.display  = tab === 'reports'  ? 'block' : 'none';
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
          position:    row[0] || null,
          type:        row[1] || '',
          domain:      row[2] || '',
          url:         row[3] || '',
          pageAS:      row[4] || 0,
          refDomains:  row[5] || 0,
          backlinks:   row[6] || 0,
          traffic:     row[7] || 0,
          urlKeywords: row[8] || 0,
          serpFeature: row[9] || null
        }));

      const report = { id: uid(), keyword: sheetName, importedAt: new Date().toISOString().split('T')[0], results };
      const existing = state.reports.findIndex(r => r.keyword === sheetName);
      if (existing >= 0) { state.reports[existing] = report; lastId = report.id; }
      else { state.reports.unshift(report); lastId = report.id; }
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
  if (state.activeReportId === id) state.activeReportId = state.reports[0]?.id || null;
  save();
  renderReports();
}

function setSearchVolume(id, value) {
  const report = state.reports.find(r => r.id === id);
  if (report) report.searchVolume = value;
  save();
}

function assignKeywordToCluster(keyword, clusterId) {
  const cluster = state.clusters.find(c => c.id === clusterId);
  if (!cluster) return false;
  if (cluster.keywords.some(k => k.text.toLowerCase() === keyword.toLowerCase())) return false;
  cluster.keywords.push({ id: uid(), text: keyword, status: 'todo' });
  save();
  return true;
}

// ─── Report Analysis ──────────────────────────────────────────────────────────

function median(arr) {
  if (!arr.length) return 0;
  const s = [...arr].sort((a, b) => a - b);
  const mid = Math.floor(s.length / 2);
  return s.length % 2 === 0 ? (s[mid - 1] + s[mid]) / 2 : s[mid];
}

function analyzeReport(report) {
  const organic = report.results.filter(r => r.type === 'Organic');

  // De-duplicate top-10 by domain
  const seen = new Set();
  const top10 = organic
    .filter(r => r.position <= 10)
    .filter(r => { if (seen.has(r.domain)) return false; seen.add(r.domain); return true; });

  const serpFeatures = [...new Set([
    ...report.results.filter(r => r.type !== 'Organic').map(r => r.type),
    ...report.results.filter(r => r.serpFeature).map(r => r.serpFeature)
  ])].filter(Boolean);

  // Only use rows with real data for metric calculations
  const withAS = top10.filter(r => r.pageAS > 0);
  const withRD = top10.filter(r => r.refDomains > 0);

  const avg = (arr, key) => arr.length ? Math.round(arr.reduce((s, r) => s + r[key], 0) / arr.length) : 0;

  const avgPageAS     = avg(withAS, 'pageAS');
  const avgBacklinks  = avg(top10,  'backlinks');
  const avgRefDomains = avg(withRD, 'refDomains');
  const avgTraffic    = avg(top10,  'traffic');
  const medianAS      = Math.round(median(withAS.map(r => r.pageAS)));
  const isSkewed      = withAS.length >= 3 && (avgPageAS - medianAS) > CONFIG.skewedDiff;

  // ── Winnability ──────────────────────────────────────────────────────────

  // Component 1: Weak-page density (max 40)
  const weakCount  = withAS.filter(r => r.pageAS < CONFIG.weakASThreshold).length;
  const weakPct    = withAS.length ? (weakCount / withAS.length) * 100 : 0;
  const comp1      = Math.round(Math.min(40, weakPct * 0.8));

  // Component 2: Low-link density (max 25)
  const lowRDCount = withRD.filter(r => r.refDomains < CONFIG.lowRDThreshold).length;
  const lowRDPct   = withRD.length ? (lowRDCount / withRD.length) * 100 : 0;
  const comp2      = Math.round(Math.min(25, lowRDPct * 0.5));

  // Component 3: Median authority floor (max 20)
  const comp3 = medianAS < 15 ? 20 : medianAS < 30 ? 13 : medianAS < 50 ? 6 : 0;

  // Component 4: UGC / forum in top 10 (max 10)
  const hasUGC = top10.some(r => CONFIG.ugcDomains.some(d => r.domain.includes(d)));
  const comp4  = hasUGC ? 10 : 0;

  // Component 5: Capturable SERP features (max 5)
  const hasAIOverview = serpFeatures.includes('AI Overview');
  const hasPAA        = serpFeatures.includes('People also ask');
  const comp5         = (hasAIOverview ? 3 : 0) + (hasPAA ? 2 : 0);

  const score  = Math.min(100, comp1 + comp2 + comp3 + comp4 + comp5);
  const bucket = score >= CONFIG.winnableMin ? 'winnable' : score >= CONFIG.competitiveMin ? 'competitive' : 'hard';

  // Soft spots: 3 lowest-AS pages in top 10
  const softSpots = [...withAS].sort((a, b) => a.pageAS - b.pageAS).slice(0, 3);

  // My domain check
  const myDomainResult = top10.find(r => r.domain.includes(CONFIG.myDomain));

  return {
    organic, top10, serpFeatures,
    avgPageAS, avgBacklinks, avgRefDomains, avgTraffic,
    medianAS, isSkewed,
    missingASCount: top10.length - withAS.length,
    winnability: { score, bucket, comp1, comp2, comp3, comp4, comp5,
      details: { weakPct, weakCount, lowRDPct, lowRDCount, hasUGC, hasAIOverview, hasPAA,
        missingCount: top10.length - withAS.length } },
    softSpots,
    myDomainResult
  };
}

// Estimate monthly clicks per position from a search volume.
// Applies the AI Overview penalty when that feature is present in the SERP.
function estimateTraffic(searchVolume, hasAIOverview) {
  const vol = Number(searchVolume) || 0;
  const mult = hasAIOverview ? CONFIG.aiOverviewCTRMultiplier : 1;
  return Object.entries(CONFIG.ctrCurve).map(([pos, baseCtr]) => {
    const ctr = baseCtr * mult;
    return {
      position: Number(pos),
      ctr,
      clicks: Math.round(vol * ctr)
    };
  });
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
        <input class="cluster-name" value="${esc(c.name)}" data-action="update-name" data-cluster-id="${c.id}">
        <button class="btn-icon" data-action="remove-cluster" data-cluster-id="${c.id}" title="Delete cluster">✕</button>
      </div>
      <input class="cluster-url" value="${esc(c.targetUrl)}" placeholder="Target page URL..."
        data-action="update-url" data-cluster-id="${c.id}">
      ${total > 0 ? `<div class="cluster-stats">${total} keyword${total !== 1 ? 's' : ''} · ${done} done</div>` : ''}
      <div class="keywords-list">
        ${c.keywords.map(kw => `
          <div class="keyword-row">
            <span class="kw-text">${esc(kw.text)}</span>
            <button class="status-badge status-${kw.status}" data-action="cycle-status"
              data-cluster-id="${c.id}" data-keyword-id="${kw.id}">${STATUS_LABELS[kw.status]}</button>
            <button class="btn-icon" data-action="remove-keyword"
              data-cluster-id="${c.id}" data-keyword-id="${kw.id}" title="Remove">✕</button>
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
  const grid  = document.getElementById('clusters-grid');
  const empty = document.getElementById('empty-state');
  if (state.clusters.length === 0) { grid.innerHTML = ''; empty.style.display = 'flex'; }
  else { empty.style.display = 'none'; grid.innerHTML = state.clusters.map(renderCluster).join(''); }
}

// ─── Render Reports ───────────────────────────────────────────────────────────

function renderReportCard(report) {
  const { organic, winnability } = analyzeReport(report);
  const { score, bucket } = winnability;
  const topDomain = organic[0]?.domain || '—';
  const isActive  = state.activeReportId === report.id;
  return `
    <div class="report-card ${isActive ? 'active' : ''}" data-report-id="${report.id}">
      <div class="report-card-header">
        <span class="report-keyword">${esc(report.keyword)}</span>
        <button class="btn-icon" data-action="delete-report" data-report-id="${report.id}" title="Delete">✕</button>
      </div>
      <div class="report-card-meta">
        <span>#1: ${esc(topDomain)}</span>
        <span class="card-bucket-badge bucket-${bucket}">${score}</span>
      </div>
      <div class="report-card-date">${report.importedAt}</div>
    </div>
  `;
}

function renderReportDetail(report) {
  const {
    organic, top10, serpFeatures,
    avgPageAS, avgBacklinks, avgRefDomains,
    medianAS, isSkewed, missingASCount,
    winnability, softSpots, myDomainResult
  } = analyzeReport(report);

  const { score, bucket, comp1, comp2, comp3, comp4, comp5, details } = winnability;
  const meta = BUCKET_META[bucket];

  const clusterOptions = state.clusters.map(c => `<option value="${c.id}">${esc(c.name)}</option>`).join('');

  const formatOpps = Object.entries(FORMAT_TIPS)
    .filter(([feature]) => serpFeatures.includes(feature))
    .filter(([feature], i, arr) => arr.findIndex(([f]) => FORMAT_TIPS[f] === FORMAT_TIPS[feature]) === i);

  const searchVolume = report.searchVolume || '';
  const trafficRows = estimateTraffic(searchVolume, details.hasAIOverview);
  const top3Clicks = trafficRows.slice(0, 3).reduce((s, r) => s + r.clicks, 0);

  return `
    <div class="report-detail-inner">

      <!-- Header -->
      <div class="report-detail-header">
        <div class="report-title-block">
          <div class="report-title-row">
            <h2>${esc(report.keyword)}</h2>
            <span class="winnability-badge bucket-${bucket}">${meta.label} · ${score}/100</span>
          </div>
          <span class="report-date">Imported ${report.importedAt} · ${organic.length} organic results</span>
          ${myDomainResult ? `<span class="my-domain-flag">Vantage found at position #${myDomainResult.position} — optimise existing page</span>` : ''}
        </div>
        ${state.clusters.length > 0 ? `
          <div class="assign-cluster">
            <select id="assign-select">
              <option value="">+ Add to cluster…</option>
              ${clusterOptions}
            </select>
          </div>
        ` : ''}
      </div>

      <!-- Winnability verdict + breakdown -->
      <div class="winnability-panel bucket-bg-${bucket}">
        <p class="verdict-text">${meta.verdict}</p>
        <details class="score-breakdown">
          <summary>Score breakdown</summary>
          <div class="breakdown-grid">
            <div class="breakdown-row">
              <span>Weak-page density <span class="breakdown-sub">(${Math.round(details.weakPct)}% of top-10 have AS &lt; ${CONFIG.weakASThreshold})</span></span>
              <span class="breakdown-pts">${comp1}<span class="breakdown-max">/40</span></span>
            </div>
            <div class="breakdown-row">
              <span>Low-link density <span class="breakdown-sub">(${Math.round(details.lowRDPct)}% of top-10 have &lt;${CONFIG.lowRDThreshold} ref. domains)</span></span>
              <span class="breakdown-pts">${comp2}<span class="breakdown-max">/25</span></span>
            </div>
            <div class="breakdown-row">
              <span>Median authority floor <span class="breakdown-sub">(median Page AS: ${medianAS})</span></span>
              <span class="breakdown-pts">${comp3}<span class="breakdown-max">/20</span></span>
            </div>
            <div class="breakdown-row">
              <span>UGC / forum in top 10 <span class="breakdown-sub">${details.hasUGC ? 'Reddit / Quora present — Google is starved for quality content' : 'None detected'}</span></span>
              <span class="breakdown-pts">${comp4}<span class="breakdown-max">/10</span></span>
            </div>
            <div class="breakdown-row">
              <span>Capturable SERP features <span class="breakdown-sub">${[details.hasAIOverview && 'AI Overview', details.hasPAA && 'People also ask'].filter(Boolean).join(', ') || 'None'}</span></span>
              <span class="breakdown-pts">${comp5}<span class="breakdown-max">/5</span></span>
            </div>
            ${details.missingCount > 0 ? `<p class="breakdown-note">${details.missingCount} top-10 row${details.missingCount > 1 ? 's' : ''} had no Page AS data and were excluded from density calculations.</p>` : ''}
          </div>
        </details>
      </div>

      <!-- Format opportunities -->
      ${formatOpps.length > 0 ? `
        <div class="format-opps">
          <h3>Format Opportunities</h3>
          <div class="opp-list">
            ${formatOpps.map(([feature, tip]) => `
              <div class="opp-card">
                <span class="feature-badge">${esc(feature)}</span>
                <p>${tip}</p>
              </div>
            `).join('')}
          </div>
        </div>
      ` : ''}

      <!-- Benchmarks -->
      <div class="benchmark-grid">
        <div class="benchmark-card">
          <div class="benchmark-value">${avgPageAS}</div>
          <div class="benchmark-label">Avg Page AS${isSkewed ? '<br><span class="skewed-note">skewed by outliers</span>' : ''}</div>
        </div>
        <div class="benchmark-card highlight">
          <div class="benchmark-value">${medianAS}</div>
          <div class="benchmark-label">Median Page AS<br><span class="honest-note">the honest number</span></div>
        </div>
        <div class="benchmark-card">
          <div class="benchmark-value">${avgBacklinks.toLocaleString()}</div>
          <div class="benchmark-label">Avg Backlinks<br>Top 10</div>
        </div>
        <div class="benchmark-card">
          <div class="benchmark-value">${avgRefDomains.toLocaleString()}</div>
          <div class="benchmark-label">Avg Ref. Domains<br>Top 10</div>
        </div>
      </div>

      <!-- Traffic potential -->
      <div class="traffic-panel">
        <div class="traffic-header">
          <h3>Traffic Potential</h3>
          <div class="volume-input">
            <label for="volume-${report.id}">Monthly search volume</label>
            <input type="number" id="volume-${report.id}" min="0" placeholder="e.g. 5400"
              value="${searchVolume}" data-action="set-volume" data-report-id="${report.id}">
          </div>
        </div>
        ${searchVolume ? `
          <div class="traffic-headline">
            Ranking top 3 ≈ <strong>${top3Clicks.toLocaleString()}</strong> clicks/mo
            ${details.hasAIOverview ? `<span class="ai-penalty-flag">AI Overview present — CTR reduced ${Math.round((1 - CONFIG.aiOverviewCTRMultiplier) * 100)}%</span>` : ''}
          </div>
          <table class="traffic-table">
            <thead><tr><th>Position</th><th>CTR</th><th>Est. clicks / mo</th></tr></thead>
            <tbody>
              ${trafficRows.map(r => `
                <tr class="${r.position <= 3 ? 'top-3' : ''}">
                  <td>#${r.position}</td>
                  <td>${(r.ctr * 100).toFixed(1)}%</td>
                  <td>${r.clicks.toLocaleString()}</td>
                </tr>
              `).join('')}
            </tbody>
          </table>
          <p class="traffic-note">Based on industry-average organic CTR by position${details.hasAIOverview ? ', adjusted for the AI Overview' : ''}. A rough planning estimate, not a guarantee.</p>
        ` : `
          <p class="traffic-empty">Enter the keyword's monthly search volume (from the SEMrush Keyword Overview) to estimate the clicks you'd win at each position.</p>
        `}
      </div>

      <!-- Soft spots -->
      ${softSpots.length > 0 ? `
        <div class="soft-spots">
          <h3>Soft Spots</h3>
          <p class="soft-spots-caption">These low-authority pages are ranking — they're who you can realistically displace.</p>
          <div class="soft-spots-list">
            ${softSpots.map(r => `
              <div class="soft-spot-row">
                <span class="soft-pos">#${r.position}</span>
                <a class="soft-domain" href="${esc(r.url)}" target="_blank" rel="noopener">${esc(r.domain)}</a>
                <span class="soft-stat">AS <strong>${r.pageAS}</strong></span>
                <span class="soft-stat">RD <strong>${r.refDomains || '—'}</strong></span>
              </div>
            `).join('')}
          </div>
        </div>
      ` : ''}

      <!-- Competitor table -->
      <div class="competitors-table-wrap">
        <table class="competitors-table">
          <thead>
            <tr><th>#</th><th>Domain</th><th>Page AS</th><th>Traffic</th><th>Backlinks</th><th>Ref. Domains</th><th>Feature</th></tr>
          </thead>
          <tbody>
            ${organic.slice(0, 20).map(r => `
              <tr class="${r.position <= 10 ? 'top-10' : ''}${r.domain.includes(CONFIG.myDomain) ? ' my-domain-row' : ''}">
                <td>${r.position}</td>
                <td class="domain-cell"><a href="${esc(r.url)}" target="_blank" rel="noopener">${esc(r.domain)}</a></td>
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
  const listEl   = document.getElementById('reports-list');
  const emptyEl  = document.getElementById('reports-empty');
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
      showToast(added ? `"${active.keyword}" added to "${cluster.name}"` : `Already in "${cluster.name}"`);
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

  if (action === 'remove-cluster') { removeCluster(cid); return; }
  if (action === 'remove-keyword') { removeKeyword(cid, kid); return; }
  if (action === 'cycle-status')   { cycleStatus(cid, kid); return; }
  if (action === 'delete-report')  { deleteReport(rid); return; }
  if (action === 'add-keyword') {
    const input = document.querySelector(`.kw-input[data-cluster-id="${cid}"]`);
    if (addKeyword(cid, input.value)) document.querySelector(`.kw-input[data-cluster-id="${cid}"]`)?.focus();
    return;
  }

  const card = e.target.closest('.report-card');
  if (card && !action) { state.activeReportId = card.dataset.reportId; renderReports(); }
});

document.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.classList.contains('kw-input')) {
    const cid = e.target.dataset.clusterId;
    if (addKeyword(cid, e.target.value)) document.querySelector(`.kw-input[data-cluster-id="${cid}"]`)?.focus();
  }
});

document.addEventListener('change', e => {
  const action = e.target.dataset.action;
  const cid    = e.target.dataset.clusterId;
  if (action === 'update-name') updateCluster(cid, 'name',      e.target.value);
  if (action === 'update-url')  updateCluster(cid, 'targetUrl', e.target.value);
  if (action === 'set-volume') {
    setSearchVolume(e.target.dataset.reportId, e.target.value);
    renderReports();
  }
});

document.querySelectorAll('.tab-btn').forEach(btn => {
  btn.addEventListener('click', () => switchTab(btn.dataset.tab));
});

const fileInput = document.getElementById('file-input');
fileInput.addEventListener('change', e => {
  if (e.target.files[0]) { parseXLSXFile(e.target.files[0]); e.target.value = ''; }
});

const reportsView = document.getElementById('view-reports');
reportsView.addEventListener('dragover',  e => { e.preventDefault(); reportsView.classList.add('drag-over'); });
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
