(function () {
  // ----- UI (Fluent-like buttons) -----
  const css = `
  body{font:14px "Segoe UI",Arial,sans-serif;padding:12px;line-height:1.5}
  h3{margin:0 0 10px}
  :root{--btn-bg:#f3f2f1;--btn-bg-h:#edebe9;--btn-bg-a:#e1dfdd;--btn-b:#8a8886;--btn-text:#323130;--btn-focus:#0078d4;--btn-p-bg:#0078d4;--btn-p-bg-h:#106ebe;--btn-p-text:#fff;--radius:4px}
  .m-btn{display:inline-flex;align-items:center;gap:8px;font:600 13px "Segoe UI",Arial,sans-serif;color:var(--btn-text);background:var(--btn-bg);border:1px solid var(--btn-b);border-radius:var(--radius);padding:6px 12px;cursor:pointer;user-select:none;transition:background .12s,border-color .12s,box-shadow .12s;width:100%;text-align:left}
  .m-btn:hover{background:var(--btn-bg-h)} .m-btn:active{background:var(--btn-bg-a)}
  .m-btn:focus{outline:none;box-shadow:0 0 0 2px var(--btn-focus) inset}
  .m-btn.primary{background:var(--btn-p-bg);color:var(--btn-p-text);border-color:var(--btn-p-bg)}
  .m-btn.primary:hover{background:var(--btn-p-bg-h);border-color:var(--btn-p-bg-h)}
  .m-icon{width:16px;height:16px;display:inline-block} .stack{display:flex;flex-direction:column;gap:8px}
  #status{margin-top:12px;font-size:12px;color:#666}`;
  const html = `
  <h3>OfficeJS Tools</h3>
  <div class="stack">
    <button id="btnFormat"  class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M3 2h10c.55 0 1 .45 1 1v10c0 .55-.45 1-1 1H3c-.55 0-1-.45-1-1V3c0-.55.45-1 1-1Zm0 1v3h4V3H3Zm5 0v3h5V3H8ZM3 7v3h4V7H3Zm5 0v3h5V7H8ZM3 11v2h4v-2H3Zm5 0v2h5v-2H8Z" fill="#9f9f9fff"/></svg>Format Tables</button>
    <button id="btnImg60"   class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M3 2h10c.55 0 1 .45 1 1v10c0 .55-.45 1-1 1H3c-.55 0-1-.45-1-1V3c0-.55.45-1 1-1Zm0 1v7.5l2.65-2.65c.2-.2.51-.2.7 0l2.15 2.15 3.65-3.64c.2-.2.51-.2.7 0L13 8.3V3H3Zm0 10h10V9.7l-1.15-1.15-3.65 3.64c-.2.2-.51.2-.7 0L5.35 9.9 3 12.24V13Z"/>
<circle cx="11" cy="5" r="1"/></svg>Images 60%</button>
    <button id="btnArial11" class="m-btn primary"><svg class="m-icon" viewBox="0 0 16 16"><path d="M7.8 4h.5l3.7 8h-1.6l-.8-2H6.4l-.8 2H4.1L7.8 4zm.3 2.2L6.9 8.9h2.1L8.1 6.2z" fill="currentColor"/></svg>Arial 11</button>
    <button id="btnTimes"   class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M4 4h8v1H9v7H7V5H4z" fill="currentColor"/></svg>Times New Roman</button>
    <button id="btnSelBlue" class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M6.66 9.25h2.68l.77 2.25h1.48L8.9 2h-1.8l-2.7 9.5h1.43l.83-2.25Zm1.34-3.73c.17.52.33 1.05.49 1.59l.34 1.04H7.49l.35-1.04c.15-.54.31-1.07.48-1.59Zm-6 8.23c0-.41.34-.75.75-.75h10.5c.42 0 .75.34.75.75s-.33.75-.75.75H2.75a.75.75 0 0 1-.75-.75Z" fill="#0078d4"/></svg>Blue Text</button>
  </div>
  <div id="status"></div>`;

  function el(html) { const d = document.createElement('div'); d.innerHTML = html; return d.children[0] || d; }

  Office.onReady(() => {
    // inject CSS + UI
    const style = document.createElement('style'); style.textContent = css; document.head.appendChild(style);
    const app = document.getElementById('app'); app.innerHTML = html;

    const status = msg => document.getElementById('status').textContent = msg;

    const getHtml = () => new Promise(r => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, x => r(x.value || ""));
    });
    const setHtml = h => new Promise(r => {
      Office.context.mailbox.item.body.setAsync(h, { coercionType: Office.CoercionType.Html }, () => r());
    });

    async function formatTables() {
      status("Formatting tables…");
      const html = await getHtml(), div = el(`<div>${html}</div>`);
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

      div.querySelectorAll('img').forEach(img => {
        // 1) If we haven’t stored original dimensions, record them once
        if (!img.hasAttribute('data-orig-width') || !img.hasAttribute('data-orig-height')) {
          if (img.naturalWidth && img.naturalHeight) {
            img.setAttribute('data-orig-width', img.naturalWidth);
            img.setAttribute('data-orig-height', img.naturalHeight);
          } else if (img.width && img.height) {
            img.setAttribute('data-orig-width', img.width);
            img.setAttribute('data-orig-height', img.height);
          }
        }

        // 2) Use stored originals to set scaled size
        const ow = parseInt(img.getAttribute('data-orig-width'), 10);
        const oh = parseInt(img.getAttribute('data-orig-height'), 10);

        if (ow && oh) {
          img.style.width = Math.round(ow * 0.6) + "px";
          img.style.height = Math.round(oh * 0.6) + "px";
        }
        img.style.maxWidth = "100%";
      });

      await setHtml(div.innerHTML);
      status("Images resized to 60% ✓");
    }

    async function setWholeBodyFont(family, sizePt) {
      status(`Setting ${family}${sizePt ? (' ' + sizePt) : ''}…`);
      const html = await getHtml(), div = el(`<div>${html}</div>`);
      div.querySelectorAll('p,div,span,td,th,li,blockquote,pre,h1,h2,h3,h4,h5,h6')
        .forEach(elm => { elm.style.fontFamily = family; if (sizePt) elm.style.fontSize = sizePt + 'pt'; });
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

    // wire buttons
    document.getElementById('btnFormat').onclick = formatTables;
    document.getElementById('btnImg60').onclick = resizeImages60;
    document.getElementById('btnArial11').onclick = setArial11;
    document.getElementById('btnTimes').onclick = setTimes;
    document.getElementById('btnSelBlue').onclick = setSelectionBlue;
  });
})();
