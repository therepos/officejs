(function(){
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
    <button id="btnFormat"  class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M1 3h14v2H1zM1 7h14v2H1zM1 11h14v2H1z" fill="currentColor"/></svg>Format tables (1pt grey, 0.1in)</button>
    <button id="btnImg60"   class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M2 3h12v10H2zM3 11l3-3 2 2 3-4 2 3v3H3z" fill="currentColor"/></svg>Resize all images to 60%</button>
    <button id="btnArial11" class="m-btn primary"><svg class="m-icon" viewBox="0 0 16 16"><path d="M7.8 4h.5l3.7 8h-1.6l-.8-2H6.4l-.8 2H4.1L7.8 4zm.3 2.2L6.9 8.9h2.1L8.1 6.2z" fill="currentColor"/></svg>Set entire message to Arial 11</button>
    <button id="btnTimes"   class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M4 4h8v1H9v7H7V5H4z" fill="currentColor"/></svg>Set entire message to Times New Roman</button>
    <button id="btnSelBlue" class="m-btn"><svg class="m-icon" viewBox="0 0 16 16"><path d="M2 12h12v2H2zM3 3h10v2H3z" fill="currentColor"/></svg>Make selected text Blue</button>
  </div>
  <div id="status"></div>`;

  function el(html){ const d=document.createElement('div'); d.innerHTML=html; return d.children[0]||d; }

  Office.onReady(() => {
    // inject CSS + UI
    const style = document.createElement('style'); style.textContent = css; document.head.appendChild(style);
    const app = document.getElementById('app'); app.innerHTML = html;

    const status = msg => document.getElementById('status').textContent = msg;

    const getHtml = () => new Promise(r=>{
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Html, x=>r(x.value||""));
    });
    const setHtml = h => new Promise(r=>{
      Office.context.mailbox.item.body.setAsync(h,{coercionType:Office.CoercionType.Html},()=>r());
    });

    async function formatTables(){
      status("Formatting tables…");
      const html = await getHtml(), div=el(`<div>${html}</div>`);
      const border='1pt solid #d9d9d9', pad='0.1in';
      div.querySelectorAll('table').forEach(t=>{
        t.style.borderCollapse='collapse';
        t.style.border=border;
        t.querySelectorAll('th,td').forEach(c=>{ c.style.border=border; c.style.padding=pad; });
      });
      await setHtml(div.innerHTML); status("Tables formatted ✓");
    }

    async function resizeImages60(){
      status("Resizing images to 60%…");
      const html = await getHtml(), div=el(`<div>${html}</div>`), scale=0.6;
      div.querySelectorAll('img').forEach(img=>{
        const w = img.width || img.naturalWidth, h = img.height || img.naturalHeight;
        if (w && h){ img.style.width=Math.round(w*scale)+'px'; img.style.height=Math.round(h*scale)+'px'; }
        else { img.style.width='60%'; img.style.height='auto'; }
        img.style.maxWidth='100%';
      });
      await setHtml(div.innerHTML); status("Images resized ✓");
    }

    async function setWholeBodyFont(family, sizePt){
      status(`Setting ${family}${sizePt?(' '+sizePt):''}…`);
      const html = await getHtml(), div=el(`<div>${html}</div>`);
      div.querySelectorAll('p,div,span,td,th,li,blockquote,pre,h1,h2,h3,h4,h5,h6')
        .forEach(elm=>{ elm.style.fontFamily=family; if(sizePt) elm.style.fontSize=sizePt+'pt'; });
      await setHtml(div.innerHTML); status("Font applied ✓");
    }
    const setArial11 = () => setWholeBodyFont('Arial, Helvetica, sans-serif', 11);
    const setTimes   = () => setWholeBodyFont('"Times New Roman", Times, serif');

    function setSelectionBlue(){
      status("Coloring selection blue…");
      Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Html, res=>{
        let sel = res?.value?.data || "";
        if(!sel.trim()){ status("Select some text first."); return; }
        const openBlueSpan = /<span\b([^>]*\bstyle\s*=\s*"(?:[^"]*;)?[^"]*color\s*:\s*(?:#0078d4|blue)\b[^"]*"(?:[^">]*)?)>/gi;
        const closeSpan=/<\/span>/gi;
        sel = sel.replace(openBlueSpan,"").replace(closeSpan,"");
        const wrapped = `<span style="color:#0078d4;">${sel}</span>`;
        Office.context.mailbox.item.setSelectedDataAsync(wrapped,{coercionType:Office.CoercionType.Html},()=>status("Selection colored ✓"));
      });
    }

    // wire buttons
    document.getElementById('btnFormat').onclick = formatTables;
    document.getElementById('btnImg60').onclick  = resizeImages60;
    document.getElementById('btnArial11').onclick= setArial11;
    document.getElementById('btnTimes').onclick  = setTimes;
    document.getElementById('btnSelBlue').onclick= setSelectionBlue;
  });
})();
