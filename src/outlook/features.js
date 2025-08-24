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
  const nearly = (a, b) => !isNaN(a) && !isNaN(b) && Math.abs(a - b) <= 1;

  div.querySelectorAll('img').forEach(img => {
    // 1) Record originals ONCE (even if one dimension is unknown → store 0)
    if (!img.hasAttribute('data-orig-width') || !img.hasAttribute('data-orig-height')) {
      const style = img.getAttribute('style') || '';
      const styleW = pxFrom(style, 'width')  || 0;
      const styleH = pxFrom(style, 'height') || 0;
      const attrW  = parseInt(img.getAttribute('width'), 10)  || 0;
      const attrH  = parseInt(img.getAttribute('height'), 10) || 0;
      const natW   = img.naturalWidth  || 0;
      const natH   = img.naturalHeight || 0;

      // Priority: inline px style → attribute → natural
      const ow = Math.round(styleW) || attrW || natW || 0;
      const oh = Math.round(styleH) || attrH || natH || 0;

      // Freeze BOTH attributes so this branch never runs again
      img.setAttribute('data-orig-width',  String(ow));
      img.setAttribute('data-orig-height', String(oh));
    }

    // 2) Use stored originals to set scaled size
    const ow = parseInt(img.getAttribute('data-orig-width'), 10)  || 0;
    const oh = parseInt(img.getAttribute('data-orig-height'), 10) || 0;

    // If we never discovered any pixels, skip (avoid % fallback to prevent enlarging)
    if (!ow && !oh) return;

    const targetW = ow ? Math.max(1, Math.round(ow * 0.6)) : NaN;
    const targetH = oh ? Math.max(1, Math.round(oh * 0.6)) : NaN;

    // If already at target, do nothing
    const styleNow = img.getAttribute('style') || '';
    const curW = pxFrom(styleNow, 'width');
    const curH = pxFrom(styleNow, 'height');
    if ((ow ? nearly(curW, targetW) : true) && (oh ? nearly(curH, targetH) : true)) return;

    if (!isNaN(targetW)) img.style.width  = targetW + 'px';
    if (!isNaN(targetH)) img.style.height = targetH + 'px';
    img.style.maxWidth = '100%';
  });

  await setHtml(div.innerHTML);
  status("Images resized to 60% ✓");
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
