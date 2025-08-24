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
    const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(\\d+(?:\\.\\d+)?)px`, 'i').exec(style || '');
    return m ? parseFloat(m[1]) : NaN;
  };
  const pctFrom = (style, prop) => {
    const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(\\d+(?:\\.\\d+)?)%`, 'i').exec(style || '');
    return m ? parseFloat(m[1]) : NaN;
  };

  let changed = 0;

  div.querySelectorAll('img').forEach(img => {
    // 1) Capture originals once (don’t rely on naturalWidth in detached DOM)
    let ow = parseInt(img.getAttribute('data-orig-width'), 10);
    let oh = parseInt(img.getAttribute('data-orig-height'), 10);

    if (!ow || !oh) {
      const attrW = parseInt(img.getAttribute('width'), 10);
      const attrH = parseInt(img.getAttribute('height'), 10);
      const style = img.getAttribute('style') || '';
      const styleWpx = pxFrom(style, 'width');
      const styleHpx = pxFrom(style, 'height');

      if (!ow) ow = !isNaN(styleWpx) ? styleWpx : (attrW || 0);
      if (!oh) oh = !isNaN(styleHpx) ? styleHpx : (attrH || 0);

      if (ow) img.setAttribute('data-orig-width', String(ow));
      if (oh) img.setAttribute('data-orig-height', String(oh));
    }

    // 2) Apply 60% scaling (with solid fallbacks)
    if (ow) {
      img.style.setProperty('width', Math.round(ow * 0.6) + 'px', 'important');
    } else {
      // scale % if present, otherwise force 60% so first click is visible
      const style = img.getAttribute('style') || '';
      const wpct = pctFrom(style, 'width');
      img.style.setProperty('width', !isNaN(wpct) ? (wpct * 0.6).toFixed(2) + '%' : '60%', 'important');
    }

    if (oh) {
      img.style.setProperty('height', Math.round(oh * 0.6) + 'px', 'important');
    } else {
      img.style.setProperty('height', 'auto', 'important');
    }

    // play nice with layouts and Outlook’s sanitizer
    img.style.setProperty('max-width', '100%', 'important');
    img.removeAttribute('width');
    img.removeAttribute('height');

    changed++;
  });

  await setHtml(div.innerHTML);
  status(changed ? `Images resized (${changed}) ✓` : 'No images to resize.');
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
