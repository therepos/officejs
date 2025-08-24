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
  const nearly = (a,b) => !isNaN(a) && !isNaN(b) && Math.abs(a-b) <= 1;

  let updated = 0, skipped = 0, unable = 0;

  div.querySelectorAll('img').forEach(img => {
    // 0) If already processed, skip
    const cls = (img.getAttribute('class') || '');
    if (/\boj-sized-60\b/.test(cls)) { skipped++; return; }

    // 1) Establish ORIGINAL pixel size once (attributes -> inline px -> natural*)
    let ow = parseInt(img.getAttribute('width'), 10);
    let oh = parseInt(img.getAttribute('height'), 10);

    const style = img.getAttribute('style') || '';
    const styleW = pxFrom(style, 'width');
    const styleH = pxFrom(style, 'height');

    if (!ow && !isNaN(styleW)) ow = Math.round(styleW);
    if (!oh && !isNaN(styleH)) oh = Math.round(styleH);

    // natural* often 0 in detached DOM; harmless if present
    if ((!ow || !oh) && img.naturalWidth && img.naturalHeight) {
      if (!ow) ow = img.naturalWidth;
      if (!oh) oh = img.naturalHeight;
    }

    // If we still can't determine a pixel original, do nothing (avoid % which can enlarge)
    if (!ow && !oh) { unable++; return; }

    // 2) Compute 60% targets
    const tw = ow ? Math.max(1, Math.round(ow * 0.6)) : NaN;
    const th = oh ? Math.max(1, Math.round(oh * 0.6)) : NaN;

    const curW = pxFrom(style, 'width');
    const curH = pxFrom(style, 'height');

    // 3) If already at target, just mark & skip
    if ((ow ? nearly(curW, tw) : true) && (oh ? nearly(curH, th) : true)) {
      img.setAttribute('class', (cls + ' oj-sized-60').trim());
      skipped++; return;
    }

    // 4) Apply pixels (no % fallback → no accidental upsizing)
    if (!isNaN(tw)) img.style.setProperty('width',  tw + 'px', 'important'); else img.style.setProperty('width',  'auto', 'important');
    if (!isNaN(th)) img.style.setProperty('height', th + 'px', 'important'); else img.style.setProperty('height', 'auto', 'important');
    img.style.setProperty('max-width', '100%', 'important');

    // Freeze "originals" for future reference (optional but helpful)
    if (!img.hasAttribute('width')  && ow) img.setAttribute('width',  String(ow));
    if (!img.hasAttribute('height') && oh) img.setAttribute('height', String(oh));

    // 5) Mark as processed so re-clicks do nothing
    img.setAttribute('class', (cls + ' oj-sized-60').trim());
    updated++;
  });

  await setHtml(div.innerHTML);
  status(updated
    ? `Images set to 60% (${updated}) ✓${skipped ? ` — ${skipped} already at 60%` : ''}${unable ? ` — ${unable} without measurable pixels` : ''}`
    : 'No images to resize.');
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
