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

async function resizeImages60() {
// Requires: Office.js; Compose mode
// 1) Get the selected HTML (should be an <img>)
  Office.context.mailbox.item.body.getSelectedDataAsync(
    Office.CoercionType.Html,
    async (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) return;
      const html = (res.value || "").trim();
      const imgMatch = html.match(/<img\b[^>]*>/i);
      if (!imgMatch) return;

      const tag = imgMatch[0];

      // Extract src
      const srcMatch = tag.match(/\bsrc=(["'])(.*?)\1/i);
      if (!srcMatch) return;
      const src = srcMatch[2];

      // 2) Load the image in the add-in to read its natural size
      const probe = new Image();
      // (crossOrigin helps some remote images; safe to try)
      try { probe.crossOrigin = "anonymous"; } catch {}
      const natural = await new Promise<{w:number;h:number}>((resolve, reject) => {
        probe.onload = () => resolve({ w: probe.naturalWidth, h: probe.naturalHeight });
        probe.onerror = reject;
        probe.src = src;
      }).catch(() => null as any);
      if (!natural || !natural.w || !natural.h) return;

      // 3) Build a clean <img> (reset to original) then scale to 60%
      const w = Math.round(natural.w * 0.6);
      const h = Math.round(natural.h * 0.6);

      // Keep all original non-size attributes (alt, class, etc.), but strip width/height/style sizing
      const cleaned = tag
        // remove width/height attrs
        .replace(/\swidth\s*=\s*["'][^"']*["']/ig, "")
        .replace(/\sheight\s*=\s*["'][^"']*["']/ig, "")
        // remove inline size styles (width/height/transform/zoom)
        .replace(/\sstyle\s*=\s*["'][^"']*["']/ig, (m) => {
          const val = m.slice(m.indexOf("=") + 1).replace(/^["']|["']$/g, "");
          const filtered = val
            .split(";")
            .map(s => s.trim())
            .filter(s => s && !/^width\s*:|^height\s*:|^transform\s*:|^zoom\s*:/.test(s.toLowerCase()))
            .join("; ");
          return filtered ? ` style="${filtered}"` : "";
        });

      // New tag: intrinsic reset (no size attrs), then explicit 60% px size
      const newImg = cleaned.replace(/<img\b/i, `<img width="${w}" height="${h}"`);

      // 4) Replace selection with resized image
      Office.context.mailbox.item.body.setSelectedDataAsync(
        newImg,
        { coercionType: Office.CoercionType.Html },
        () => {/* no-op */}
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
