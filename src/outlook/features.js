// Helpers
const status = msg => (document.getElementById('status').textContent = msg);
const getHtml = () => new Promise(r => {
  Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, x => r(x.value || ""));
});
const setHtml = h => new Promise(r => {
  Office.context.mailbox.item.body.setAsync(h, { coercionType: Office.CoercionType.Html }, () => r());
});
const wrapDiv = html => { const d = document.createElement('div'); d.innerHTML = html; return d; };

// Features
async function formatTables() {
  status("Formatting tables…");
  const html = await getHtml(), div = wrapDiv(html);
  const border = '1pt solid #d9d9d9', pad = '0.1in';
  div.querySelectorAll('table').forEach(t => {
    t.style.borderCollapse = 'collapse';
    t.style.border = border;
    t.querySelectorAll('th,td').forEach(c => { c.style.border = border; c.style.padding = pad; });
  });
  await setHtml(div.innerHTML); status("Tables formatted ✓");
}

async function resizeImages60() {
  status("Resizing images to 60% of original…");
  const html = await getHtml(), div = document.createElement('div'); div.innerHTML = html;

  const pxFrom = (style, prop) => {
    const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(\\d+(?:\\.\\d+)?)px\\s*(?:!important)?`, 'i').exec(style || '');
    return m ? parseFloat(m[1]) : NaN;
  };
  const nearly = (a,b) => !isNaN(a) && !isNaN(b) && Math.abs(a - b) <= 1;

  const loadNatural = (src) => new Promise(resolve => {
    // Skip non-loadable schemes (cid:, about:, etc.)
    if (!src || /^(cid|about):/i.test(src)) return resolve({ w: NaN, h: NaN });
    const im = new Image();
    let done = false;
    const finish = (w, h) => { if (!done) { done = true; resolve({ w, h }); } };
    im.onload = () => finish(im.naturalWidth || NaN, im.naturalHeight || NaN);
    im.onerror = () => finish(NaN, NaN);
    // Timeout so we don't hang if blocked by CSP/network
    setTimeout(() => finish(NaN, NaN), 1500);
    im.src = src;
  });

  const imgs = Array.from(div.querySelectorAll('img'));
  // Preload naturals for only those that lack stored/attribute/style px
  const preloadPromises = imgs.map(async img => {
    if (img.getAttribute('data-oj-ow') || img.getAttribute('data-oj-oh')) return null;
    const attrW = parseInt(img.getAttribute('width'), 10);
    const attrH = parseInt(img.getAttribute('height'), 10);
    const st = img.getAttribute('style') || '';
    const stW = pxFrom(st, 'width');
    const stH = pxFrom(st, 'height');
    if (attrW || attrH || !isNaN(stW) || !isNaN(stH)) return null;
    // Try to get natural size once
    const { w, h } = await loadNatural(img.getAttribute('src'));
    if (w) img.setAttribute('data-oj-ow', String(w));
    if (h) img.setAttribute('data-oj-oh', String(h));
    return null;
  });

  await Promise.all(preloadPromises);

  let updated = 0, skipped = 0, unknown = 0;

  imgs.forEach(img => {
    // 0) If we’ve already finalized once, skip.
    if (img.getAttribute('data-oj-done') === '60') { skipped++; return; }

    // 1) Establish ORIGINAL pixels once (priority: stored → attrs → inline px → naturals we just fetched)
    let ow = parseInt(img.getAttribute('data-oj-ow'), 10);
    let oh = parseInt(img.getAttribute('data-oj-oh'), 10);

    if (!ow || !oh) {
      const attrW = parseInt(img.getAttribute('width'), 10);
      const attrH = parseInt(img.getAttribute('height'), 10);
      const st = img.getAttribute('style') || '';
      const stW = pxFrom(st, 'width');
      const stH = pxFrom(st, 'height');

      if (!ow) ow = attrW || (!isNaN(stW) ? Math.round(stW) : parseInt(img.getAttribute('data-oj-ow'), 10) || 0);
      if (!oh) oh = attrH || (!isNaN(stH) ? Math.round(stH) : parseInt(img.getAttribute('data-oj-oh'), 10) || 0);

      // Persist originals if found
      if (ow) img.setAttribute('data-oj-ow', String(ow));
      if (oh) img.setAttribute('data-oj-oh', String(oh));
    }

    // If we still don't know any pixel originals, we cannot safely compute 60% without risk of enlarging — skip.
    if (!ow && !oh) { unknown++; return; }

    // 2) Compute 60% targets
    const tw = ow ? Math.max(1, Math.round(ow * 0.6)) : NaN;
    const th = oh ? Math.max(1, Math.round(oh * 0.6)) : NaN;

    // 3) If already at target, mark done and skip
    const stNow = img.getAttribute('style') || '';
    const curW = pxFrom(stNow, 'width');
    const curH = pxFrom(stNow, 'height');
    const alreadyAtTarget = (ow ? nearly(curW, tw) : true) && (oh ? nearly(curH, th) : true);
    if (alreadyAtTarget) {
      img.setAttribute('data-oj-done', '60'); // idempotent in future
      skipped++;
      return;
    }

    // 4) Apply pixel sizes (never use % → never enlarge)
    if (!isNaN(tw)) img.style.setProperty('width',  tw + 'px', 'important');
    if (!isNaN(th)) img.style.setProperty('height', th + 'px', 'important');
    img.style.setProperty('max-width', '100%', 'important');

    // 5) Finalize
    img.setAttribute('data-oj-done', '60');
    updated++;
  });

  await setHtml(div.innerHTML);
  let msg = updated ? `Images set to 60% (${updated}) ✓` : 'No images to resize.';
  if (skipped) msg += ` — ${skipped} already at 60%`;
  if (unknown) msg += ` — ${unknown} without measurable original (skipped)`;
  status(msg);
}

async function setWholeBodyFont(family, sizePt) {
  status(`Setting ${family}${sizePt ? (' ' + sizePt) : ''}…`);
  const html = await getHtml(), div = wrapDiv(html);
  div.querySelectorAll('p,div,span,td,th,li,blockquote,pre,h1,h2,h3,h4,h5,h6')
    .forEach(el => { el.style.fontFamily = family; if (sizePt) el.style.fontSize = sizePt + 'pt'; });
  await setHtml(div.innerHTML); status("Font applied ✓");
}

const setArial11 = () => setWholeBodyFont('Arial, Helvetica, sans-serif', 11);
const setTimes = () => setWholeBodyFont('"Times New Roman", Times, serif');

function setSelectionBlue() {
  status("Coloring selection blue…");
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Html, res => {
    let sel = res?.value?.data || "";
    if (!sel.trim()) { status("Select some text first."); return; }
    const openBlueSpan = /<span\b([^>]*\bstyle\s*=\s*"(?:[^"]*;)?[^"]*color\s*:\s*(?:#0078d4|blue)\b[^"]*"(?:[^">]*)?)>/gi;
    const closeSpan = /<\/span>/gi;
    sel = sel.replace(openBlueSpan, "").replace(closeSpan, "");
    const wrapped = `<span style="color:#0078d4;">${sel}</span>`;
    Office.context.mailbox.item.setSelectedDataAsync(wrapped, { coercionType: Office.CoercionType.Html }, () => status("Selection colored ✓"));
  });
}

// expose for ui.js
window.OfficeJsFeatures = {
  formatTables, resizeImages60, setArial11, setTimes, setSelectionBlue
};
