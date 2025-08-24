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
  const is60pct = (style) => /(?:^|;)\s*width\s*:\s*60%\s*(?:!important)?\s*(?:;|$)/i.test(style || '');
  const nearly = (a,b) => !isNaN(a) && !isNaN(b) && Math.abs(a-b) <= 1;

  let updated = 0, skipped = 0;

  div.querySelectorAll('img').forEach(img => {
    const style = img.getAttribute('style') || '';
    let origW = parseInt(img.getAttribute('width'), 10);
    let origH = parseInt(img.getAttribute('height'), 10);

    const styleWpx = pxFrom(style, 'width');
    const styleHpx = pxFrom(style, 'height');

    // If no presentational attrs, freeze current px styles as "originals" once
    if (!origW && !isNaN(styleWpx)) { origW = Math.round(styleWpx); img.setAttribute('width', String(origW)); }
    if (!origH && !isNaN(styleHpx)) { origH = Math.round(styleHpx); img.setAttribute('height', String(origH)); }

    // If we still don’t know original pixels, use % and make it idempotent
    if (!origW && !origH) {
      if (is60pct(style)) { skipped++; return; }            // already at 60%
      img.style.setProperty('width',  '60%', 'important');  // one-time visible change
      img.style.setProperty('height', 'auto', 'important');
      img.style.setProperty('max-width', '100%', 'important');
      updated++;
      return;
    }

    // Compute target(s) from originals
    const targetW = origW ? Math.round(origW * 0.6) : NaN;
    const targetH = origH ? Math.round(origH * 0.6) : NaN;

    const curW = pxFrom(style, 'width');
    const curH = pxFrom(style, 'height');

    // If already at target, skip
    if ((origW ? nearly(curW, targetW) : true) && (origH ? nearly(curH, targetH) : true)) {
      skipped++; return;
    }

    // Apply target(s)
    if (!isNaN(targetW)) img.style.setProperty('width',  targetW + 'px', 'important');
    else                 img.style.setProperty('width',  '60%', 'important');

    if (!isNaN(targetH)) img.style.setProperty('height', targetH + 'px', 'important');
    else                 img.style.setProperty('height', 'auto', 'important');

    img.style.setProperty('max-width', '100%', 'important');

    // IMPORTANT: do NOT remove width/height attributes. They are our "original" reference.
    updated++;
  });

  await setHtml(div.innerHTML);
  status(updated ? `Images set to 60% (${updated}) ✓${skipped?` — ${skipped} already at 60%`:''}` : 'No images to resize.');
}


  await setHtml(div.innerHTML);
  status(updated ? `Images set to 60% (${updated}) ✓${skipped?` — ${skipped} already at 60%`:''}` : 'No images to resize.');
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
