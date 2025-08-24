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
  const html = await getHtml(), div = wrapDiv(html);

  const getPxFromStyle = (styleStr, prop) => {
    if (!styleStr) return null;
    const m = new RegExp(`(?:^|;)\\s*${prop}\\s*:\\s*(\\d+(?:\\.\\d+)?)px`, 'i').exec(styleStr);
    return m ? parseFloat(m[1]) : null;
  };

  div.querySelectorAll('img').forEach(img => {
    // 1) Establish originals: prefer pre-existing data-*; else read from attributes/style.
    let ow = parseInt(img.getAttribute('data-orig-width'), 10);
    let oh = parseInt(img.getAttribute('data-orig-height'), 10);

    if (!ow || !oh) {
      const attrW = parseInt(img.getAttribute('width'), 10);
      const attrH = parseInt(img.getAttribute('height'), 10);
      const style = img.getAttribute('style') || "";
      const styleW = getPxFromStyle(style, 'width');
      const styleH = getPxFromStyle(style, 'height');

      // Try to infer a pair that preserves aspect ratio
      // Priority: (styleW, styleH) → (attrW, attrH) → (styleW, attrH) / (attrW, styleH)
      // If only one dimension is known, we’ll scale that and let the other be auto.
      ow = ow || styleW || attrW || 0;
      oh = oh || styleH || attrH || 0;

      // As a last resort, if neither px is known and images aren't loaded, leave originals unset.
      if (ow) img.setAttribute('data-orig-width', String(ow));
      if (oh) img.setAttribute('data-orig-height', String(oh));
    }

    // 2) Apply 60% scaling
    if (ow && oh) {
      img.style.width = Math.round(ow * 0.6) + 'px';
      img.style.height = Math.round(oh * 0.6) + 'px';
    } else if (ow && !oh) {
      // Only width known: scale width, keep aspect
      img.style.width = Math.round(ow * 0.6) + 'px';
      img.style.height = 'auto';
    } else if (!ow && oh) {
      // Only height known: scale height, let width auto
      img.style.height = Math.round(oh * 0.6) + 'px';
      img.style.width = 'auto';
    } else {
      // Nothing known yet: use relative width; first click will still work
      img.style.width = '60%';
      img.style.height = 'auto';
    }

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
