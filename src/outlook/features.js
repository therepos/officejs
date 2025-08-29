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

  // Step 1: Get selected HTML
  const selectedHtml = await new Promise((resolve, reject) => {
    Office.context.mailbox.item.getSelectedDataAsync(
      Office.CoercionType.Html,
      (res) => {
        if (res.status === Office.AsyncResultStatus.Succeeded) {
          resolve(res.value?.data || "");
        } else {
          reject("Failed to get selection.");
        }
      }
    );
  });

  if (!selectedHtml.trim()) {
    status("Select an image first.");
    return;
  }

  const selectedDiv = wrapDiv(selectedHtml);
  const selectedImg = selectedDiv.querySelector("img");

  if (!selectedImg) {
    status("Selected content is not an image.");
    return;
  }

  const selectedSrc = selectedImg.getAttribute("src");
  if (!selectedSrc) {
    status("Could not identify selected image.");
    return;
  }

  // Step 2: Get full body HTML
  const fullHtml = await getHtml();
  const fullDiv = wrapDiv(fullHtml);

  // Step 3: Find matching image in full body
  const matchingImg = Array.from(fullDiv.querySelectorAll("img")).find(
    (img) => img.getAttribute("src") === selectedSrc
  );

  if (!matchingImg) {
    status("Could not find selected image in body.");
    return;
  }

  // Step 4: Reset and scale
  matchingImg.removeAttribute("width");
  matchingImg.removeAttribute("height");
  matchingImg.style.width = "";
  matchingImg.style.height = "";
  matchingImg.style.transform = "scale(0.6)";
  matchingImg.style.transformOrigin = "top left";

  // Step 5: Write back full HTML
  await setHtml(fullDiv.innerHTML);
  status("Selected image resized ✓");
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
