// ─── State ───────────────────────────────────────────────────────────────────

let state = { clusters: [] };

const STATUS = ['todo', 'in-progress', 'done'];
const STATUS_LABELS = { 'todo': 'To Do', 'in-progress': 'In Progress', 'done': 'Done' };

function uid() {
  return Math.random().toString(36).slice(2, 9);
}

// ─── Persistence ─────────────────────────────────────────────────────────────

function save() {
  localStorage.setItem('kco-state', JSON.stringify(state));
}

function load() {
  const raw = localStorage.getItem('kco-state');
  if (raw) state = JSON.parse(raw);
}

// ─── Actions ─────────────────────────────────────────────────────────────────

function addCluster() {
  state.clusters.push({ id: uid(), name: 'New Cluster', targetUrl: '', keywords: [] });
  save();
  render();
}

function removeCluster(id) {
  state.clusters = state.clusters.filter(c => c.id !== id);
  save();
  render();
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
  render();
  return true;
}

function removeKeyword(clusterId, keywordId) {
  const cluster = state.clusters.find(c => c.id === clusterId);
  if (cluster) cluster.keywords = cluster.keywords.filter(k => k.id !== keywordId);
  save();
  render();
}

function cycleStatus(clusterId, keywordId) {
  const cluster = state.clusters.find(c => c.id === clusterId);
  const kw = cluster && cluster.keywords.find(k => k.id === keywordId);
  if (kw) kw.status = STATUS[(STATUS.indexOf(kw.status) + 1) % STATUS.length];
  save();
  render();
}

// ─── Export ───────────────────────────────────────────────────────────────────

function exportCSV() {
  const rows = [['Keyword', 'Cluster', 'Target URL', 'Status']];
  state.clusters.forEach(c => {
    c.keywords.forEach(kw => {
      rows.push([`"${kw.text}"`, `"${c.name}"`, `"${c.targetUrl}"`, STATUS_LABELS[kw.status]]);
    });
  });
  if (rows.length === 1) {
    alert('No keywords to export.');
    return;
  }
  const csv = rows.map(r => r.join(',')).join('\n');
  const a = document.createElement('a');
  a.href = URL.createObjectURL(new Blob([csv], { type: 'text/csv' }));
  a.download = 'keyword-clusters.csv';
  a.click();
}

// ─── Render ───────────────────────────────────────────────────────────────────

function esc(str) {
  return String(str)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
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
              data-keyword-id="${kw.id}">
              ${STATUS_LABELS[kw.status]}
            </button>
            <button class="btn-icon"
              data-action="remove-keyword"
              data-cluster-id="${c.id}"
              data-keyword-id="${kw.id}"
              title="Remove">✕</button>
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

function render() {
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

// ─── Events ───────────────────────────────────────────────────────────────────

document.addEventListener('click', e => {
  const action = e.target.dataset.action;
  if (!action) return;
  const cid = e.target.dataset.clusterId;
  const kid = e.target.dataset.keywordId;

  if (action === 'remove-cluster') removeCluster(cid);
  if (action === 'remove-keyword') removeKeyword(cid, kid);
  if (action === 'cycle-status') cycleStatus(cid, kid);
  if (action === 'add-keyword') {
    const input = document.querySelector(`.kw-input[data-cluster-id="${cid}"]`);
    if (addKeyword(cid, input.value)) {
      document.querySelector(`.kw-input[data-cluster-id="${cid}"]`).focus();
    }
  }
});

document.addEventListener('keydown', e => {
  if (e.key === 'Enter' && e.target.classList.contains('kw-input')) {
    const cid = e.target.dataset.clusterId;
    if (addKeyword(cid, e.target.value)) {
      document.querySelector(`.kw-input[data-cluster-id="${cid}"]`).focus();
    }
  }
});

document.addEventListener('change', e => {
  const action = e.target.dataset.action;
  const cid = e.target.dataset.clusterId;
  if (action === 'update-name') updateCluster(cid, 'name', e.target.value);
  if (action === 'update-url') updateCluster(cid, 'targetUrl', e.target.value);
});

document.getElementById('add-cluster-btn').addEventListener('click', addCluster);
document.getElementById('export-btn').addEventListener('click', exportCSV);
document.getElementById('empty-add-btn').addEventListener('click', addCluster);

// ─── Init ─────────────────────────────────────────────────────────────────────

load();
render();
