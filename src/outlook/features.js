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

  let updated = 0, skipped = 0, pending = 0;

  div.querySelectorAll('img').forEach(img => {
    // Already done? bail early.
    if (img.getAttribute('data-oj-scaled') === '60') { skipped++; return; }

    // 1) Get or infer pixel originals (attrs -> style px -> natural*)
    let ow = parseInt(img.getAttribute('data-orig-width'), 10);
    let oh = parseInt(img.getAttribute('data-orig-height'), 10);

    if (!ow || !oh) {
      const attrW = parseInt(img.getAttribute('width'), 10);
      const attrH = parseInt(img.getAttribute('height'), 10);
      const style  = img.getAttribute('style') || '';
      const styleW = pxFrom(style, 'width');
      const styleH = pxFrom(style, 'height');

      if (!ow) ow = attrW || (!isNaN(styleW) ? Math.round(styleW) : 0);
      if (!oh) oh = attrH || (!isNaN(styleH) ? Math.round(styleH) : 0);

      // As a last resort, try natural sizes (often 0 on first click, nonzero later)
      if ((!ow || !oh) && img.naturalWidth && img.naturalHeight) {
        ow = ow || img.naturalWidth;
        oh = oh || img.naturalHeight;
      }

      if (ow) img.setAttribute('data-orig-width', String(ow));
      if (oh) img.setAttribute('data-orig-height', String(oh));
    }

    // If we still can't determine a pixel original, skip this time (don’t enlarge via %)
    if (!ow && !oh) { pending++; return; }

    const targetW = ow ? Math.max(1, Math.round(ow * 0.6)) : NaN;
    const targetH = oh ? Math.max(1, Math.round(oh * 0.6)) : NaN;

    const style = img.getAttribute('style') || '';
    const curW  = pxFrom(style, 'width');
    const curH  = pxFrom(style, 'height');

    // Already at target? mark + skip.
    if ((ow ? nearly(curW, targetW) : true) && (oh ? nearly(curH, targetH) : true)) {
      img.setAttribute('data-oj-scaled', '60');
      skipped++; 
      return;
    }

    // 2) Apply target pixels (no % fallback to avoid upsizing)
    if (!isNaN(targetW)) img.style.setProperty('width',  targetW + 'px', 'important');
    if (!isNaN(targetH)) img.style.setProperty('height', targetH + 'px', 'important');
    img.style.setProperty('max-width', '100%', 'important');

    // Keep width/height attributes as the “originals”; don’t remove them.
    img.setAttribute('data-oj-scaled', '60');
    updated++;
  });

  await setHtml(div.innerHTML);
  let msg = updated ? `Images set to 60% (${updated}) ✓` : 'No images to resize.';
  if (skipped) msg += ` — ${skipped} already at 60%`;
  if (pending) msg += ` — ${pending} waiting for size info`;
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
