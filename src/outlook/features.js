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

  const cssNum = (style, name) => {
    const m = new RegExp(`--${name}\\s*:\\s*(\\d+(?:\\.\\d+)?)\\b`).exec(style || '');
    return m ? parseFloat(m[1]) : NaN;
  };
  const pxFrom = (style, prop) => {
    const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(\\d+(?:\\.\\d+)?)px`, 'i').exec(style || '');
    return m ? parseFloat(m[1]) : NaN;
  };
  const setCssVar = (img, name, val) => {
    if (isNaN(val)) return;
    const re = new RegExp(`--${name}\\s*:\\s*[^;]+;?`);
    let s = img.getAttribute('style') || '';
    s = re.test(s) ? s.replace(re, `--${name}:${val};`) : (s + `;--${name}:${val};`);
    img.setAttribute('style', s);
  };
  const hasClass = (el, c) => new RegExp(`\\b${c}\\b`).test(el.getAttribute('class') || '');
  const addClass = (el, c) => el.setAttribute('class', ((el.getAttribute('class') || '') + ' ' + c).trim());
  const eq = (a,b)=>!isNaN(a)&&!isNaN(b)&&Math.abs(a-b)<=1;

  let updated = 0, skipped = 0;

  div.querySelectorAll('img').forEach(img => {
    const style = img.getAttribute('style') || '';

    // 1) Read stored originals if present (persisted in inline style)
    let ow = cssNum(style, 'oj-orig-w');
    let oh = cssNum(style, 'oj-orig-h');

    // 2) If not stored yet, infer once from existing HTML (attrs or inline px)
    if (isNaN(ow) || isNaN(oh)) {
      const attrW = parseInt(img.getAttribute('width'), 10);
      const attrH = parseInt(img.getAttribute('height'), 10);
      const styleW = pxFrom(style, 'width');
      const styleH = pxFrom(style, 'height');
      if (isNaN(ow)) ow = !isNaN(styleW) ? styleW : (attrW || NaN);
      if (isNaN(oh)) oh = !isNaN(styleH) ? styleH : (attrH || NaN);
    }

    // 3) If we already processed (class marker) and we have no numeric originals, do nothing
    if (hasClass(img, 'oj-sized-60') && isNaN(ow) && isNaN(oh)) { skipped++; return; }

    // Targets from originals, if known
    const tw = !isNaN(ow) ? Math.round(ow * 0.6) : NaN;
    const th = !isNaN(oh) ? Math.round(oh * 0.6) : NaN;

    const curW = pxFrom(style, 'width');
    const curH = pxFrom(style, 'height');

    // 4) If already at 60% (within 1px), skip
    if ((!isNaN(tw) && eq(curW, tw)) && (isNaN(th) || eq(curH, th))) { skipped++; return; }

    // 5) Apply scaling once
    if (!isNaN(tw)) img.style.setProperty('width',  tw + 'px', 'important');
    else             img.style.setProperty('width',  '60%',   'important'); // fallback, one-time

    if (!isNaN(th)) img.style.setProperty('height', th + 'px', 'important');
    else             img.style.setProperty('height', 'auto',   'important');

    img.style.setProperty('max-width', '100%', 'important');

    // Remove presentational attrs that can override
    img.removeAttribute('width'); img.removeAttribute('height');

    // Persist originals (so next click knows not to shrink again)
    if (!isNaN(ow)) setCssVar(img, 'oj-orig-w', ow);
    if (!isNaN(oh)) setCssVar(img, 'oj-orig-h', oh);
    addClass(img, 'oj-sized-60');

    updated++;
  });

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
