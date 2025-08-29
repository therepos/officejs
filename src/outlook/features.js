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
  status("Resizing selected image…");

  Office.context.mailbox.item.getSelectedDataAsync(
    Office.CoercionType.Html,
    async (res) => {
      if (res.status !== Office.AsyncResultStatus.Succeeded) {
        status("Failed to get selection.");
        return;
      }

      const html = res.value?.data || "";
      if (!html.trim()) {
        status("Select an image first.");
        return;
      }

      const div = wrapDiv(html);
      const img = div.querySelector("img");

      if (!img) {
        status("Selected content is not an image.");
        return;
      }

      // Reset and scale
      img.removeAttribute("width");
      img.removeAttribute("height");
      img.style.width = "";
      img.style.height = "";
      img.style.transform = "scale(0.6)";
      img.style.transformOrigin = "top left";

      Office.context.mailbox.item.setSelectedDataAsync(
        div.innerHTML,
        { coercionType: Office.CoercionType.Html },
        (result) => {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            status("Selected image resized ✓");
          } else {
            status("Failed to update image.");
          }
        }
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
      const html = res?.value?.data || "";
      if (!html.trim()) {
        status("Select some text first.");
        return;
      }

      const div = wrapDiv(html);
      const span = document.createElement("span");
      span.style.color = "#0078d4";
      span.innerHTML = div.innerHTML;

      Office.context.mailbox.item.setSelectedDataAsync(
        span.outerHTML,
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
