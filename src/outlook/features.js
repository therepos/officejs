// ----------------- helpers -----------------
const status = (msg) =>
  (document.getElementById("status").textContent = msg);

const getHtml = () =>
  new Promise((r) => {
    Office.context.mailbox.item.body.getAsync(
      Office.CoercionType.Html,
      (x) => r(x.value || "")
    );
  });

const setHtml = (h) =>
  new Promise((r) => {
    Office.context.mailbox.item.body.setAsync(
      h,
      { coercionType: Office.CoercionType.Html },
      () => r()
    );
  });

const wrapDiv = (html) => {
  const d = document.createElement("div");
  d.innerHTML = html;
  return d;
};

// ----------------- features -----------------
async function formatTables() {
  status("Formatting tables…");
  const html = await getHtml(),
    div = wrapDiv(html);
  const border = "1pt solid #d9d9d9",
    pad = "0.1in";
  div.querySelectorAll("table").forEach((t) => {
    t.style.borderCollapse = "collapse";
    t.style.border = border;
    t.querySelectorAll("th,td").forEach((c) => {
      c.style.border = border;
      c.style.padding = pad;
    });
  });
  await setHtml(div.innerHTML);
  status("Tables formatted ✓");
}

// Compose only
async function resizeImages60() {
  Office.context.mailbox.item.body.getSelectedDataAsync(
    Office.CoercionType.Html,
    async (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return;

      const html = (res.value || "").trim();
      if (!html) return;

      // Parse robustly (avoid regex-only parsing)
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, "text/html");

      // Handle cases like <a><img/></a> or plain <img/>
      const imgs = Array.from(doc.querySelectorAll("img"));
      if (!imgs.length) return;

      // Probe all images (parallel)
      const probed = await Promise.all(
        imgs.map(async (img) => {
          const src = img.getAttribute("src") || "";
          // Can't probe cid: sources from add-in domain
          if (!src || src.startsWith("cid:")) return { img, w: 0, h: 0 };

          const probe = new Image();
          try { probe.crossOrigin = "anonymous"; } catch {}
          const size = await new Promise<{ w: number; h: number }>((resolve) => {
            probe.onload = () => resolve({ w: probe.naturalWidth, h: probe.naturalHeight });
            probe.onerror = () => resolve({ w: 0, h: 0 });
            probe.src = src;
          });
          return { img, ...size };
        })
      );

      // Mutate DOM: reset & set to 60% where we know natural size
      probed.forEach(({ img, w, h }) => {
        // Strip explicit sizing (attrs + inline style)
        img.removeAttribute("width");
        img.removeAttribute("height");
        const style = (img.getAttribute("style") || "")
          .split(";")
          .map(s => s.trim())
          .filter(s => s && !/^(width|height|transform|zoom)\s*:/i.test(s))
          .join("; ");
        style ? img.setAttribute("style", style) : img.removeAttribute("style");

        if (w > 0 && h > 0) {
          img.setAttribute("width", String(Math.round(w * 0.6)));
          img.setAttribute("height", String(Math.round(h * 0.6)));
        }
        // else: we leave it intrinsic (best effort for cid:/blocked sources)
      });

      // Replace selection with updated HTML
      Office.context.mailbox.item.body.setSelectedDataAsync(
        doc.body.innerHTML,
        { coercionType: Office.CoercionType.Html },
        () => {}
      );
    }
  );
}

async function setWholeBodyFont(family, sizePt) {
  status(`Setting ${family}${sizePt ? " " + sizePt : ""}…`);
  const html = await getHtml(),
    div = wrapDiv(html);
  div
    .querySelectorAll(
      "p,div,span,td,th,li,blockquote,pre,h1,h2,h3,h4,h5,h6"
    )
    .forEach((el) => {
      el.style.fontFamily = family;
      if (sizePt) el.style.fontSize = sizePt + "pt";
    });
  await setHtml(div.innerHTML);
  status("Font applied ✓");
}

const setArial11 = () =>
  setWholeBodyFont("Arial, Helvetica, sans-serif", 11);
const setTimes = () =>
  setWholeBodyFont('"Times New Roman", Times, serif');

function setSelectionBlue() {
  status("Coloring selection blue…");
  Office.context.mailbox.item.getSelectedDataAsync(
    Office.CoercionType.Html,
    (res) => {
      let sel = res?.value?.data || "";
      if (!sel.trim()) {
        status("Select some text first.");
        return;
      }
      const openBlueSpan =
        /<span\b([^>]*\bstyle\s*=\s*"(?:[^"]*;)?[^"]*color\s*:\s*(?:#0078d4|blue)\b[^"]*"(?:[^">]*)?)>/gi;
      const closeSpan = /<\/span>/gi;
      sel = sel.replace(openBlueSpan, "").replace(closeSpan, "");
      const wrapped = `<span style="color:#0078d4;">${sel}</span>`;
      Office.context.mailbox.item.setSelectedDataAsync(
        wrapped,
        { coercionType: Office.CoercionType.Html },
        () => status("Selection colored ✓")
      );
    }
  );
}

// ----------------- UI -----------------
Office.onReady(() => {
  const $ = (id) => document.getElementById(id);

  $("btnFormat").onclick = formatTables;
  $("btnImg60").onclick = resizeImages60;
  $("btnArial11").onclick = setArial11;
  $("btnTimes").onclick = setTimes;
  $("btnSelBlue").onclick = setSelectionBlue;

});
