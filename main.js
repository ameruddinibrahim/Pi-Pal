const titleInput = document.getElementById('page-title');
const descInput = document.getElementById('meta-desc');
const urlInput = document.getElementById('page-url');

const previewTitle = document.getElementById('preview-title');
const previewDesc = document.getElementById('preview-desc');
const previewUrl = document.getElementById('preview-url');
const previewSite = document.getElementById('preview-site');

const titleCount = document.getElementById('title-count');
const descCount = document.getElementById('desc-count');

const metaOutput = document.getElementById('meta-output');
const copyBtn = document.getElementById('copy-btn');

const TITLE_LIMIT = 60;
const DESC_LIMIT = 160;

function updateCounter(el, count, limit) {
  el.textContent = count;
  el.className = 'count';
  if (count === 0) return;
  if (count <= Math.floor(limit * 0.85)) el.classList.add('good');
  else if (count <= limit) el.classList.add('warn');
  else el.classList.add('over');
}

function truncate(text, limit) {
  return text.length > limit ? text.slice(0, limit) + '...' : text;
}

function formatUrl(raw) {
  try {
    const url = new URL(raw);
    const parts = url.pathname.split('/').filter(Boolean);
    const breadcrumb = [url.hostname, ...parts].join(' › ');
    return { display: breadcrumb, site: url.hostname };
  } catch {
    return { display: raw || 'https://example.com', site: raw ? raw.replace(/https?:\/\//, '').split('/')[0] : 'example.com' };
  }
}

function generateMetaTags(title, desc, url) {
  const lines = [];

  if (title) {
    lines.push(`<title>${title}</title>`);
    lines.push(`<meta name="title" content="${title}">`);
  }
  if (desc) {
    lines.push(`<meta name="description" content="${desc}">`);
  }

  if (title || desc || url) {
    lines.push('');
    lines.push('<!-- Open Graph -->');
    if (title) lines.push(`<meta property="og:title" content="${title}">`);
    if (desc) lines.push(`<meta property="og:description" content="${desc}">`);
    if (url) lines.push(`<meta property="og:url" content="${url}">`);

    lines.push('');
    lines.push('<!-- Twitter Card -->');
    lines.push(`<meta name="twitter:card" content="summary_large_image">`);
    if (title) lines.push(`<meta name="twitter:title" content="${title}">`);
    if (desc) lines.push(`<meta name="twitter:description" content="${desc}">`);
  }

  return lines.join('\n');
}

function update() {
  const title = titleInput.value;
  const desc = descInput.value;
  const url = urlInput.value;

  updateCounter(titleCount, title.length, TITLE_LIMIT);
  updateCounter(descCount, desc.length, DESC_LIMIT);

  previewTitle.textContent = title
    ? truncate(title, TITLE_LIMIT)
    : 'Your page title will appear here';

  previewDesc.textContent = desc
    ? truncate(desc, DESC_LIMIT)
    : 'Your meta description will appear here. Make it compelling and informative to improve click-through rates from search results.';

  const { display, site } = formatUrl(url);
  previewUrl.textContent = display;
  previewSite.textContent = site;

  metaOutput.textContent = generateMetaTags(title, desc, url);
}

copyBtn.addEventListener('click', () => {
  const text = metaOutput.textContent;
  if (!text.trim()) return;
  navigator.clipboard.writeText(text).then(() => {
    copyBtn.textContent = 'Copied!';
    copyBtn.classList.add('copied');
    setTimeout(() => {
      copyBtn.textContent = 'Copy';
      copyBtn.classList.remove('copied');
    }, 2000);
  });
});

titleInput.addEventListener('input', update);
descInput.addEventListener('input', update);
urlInput.addEventListener('input', update);

update();
