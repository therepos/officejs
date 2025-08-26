const status = (msg) => {
  const el = document.getElementById("status");
  if (el) el.textContent = msg || "";
};

Office.onReady(() => {
  const $ = (id) => document.getElementById(id);
  $("btnArial8").onclick = setWorkbookArial8;
  $("btnGridOff").onclick = hideGridlinesAll;
});

// Set all used cells in every sheet to Arial 8
async function setWorkbookArial8() {
  status("Applying Arial 8 to entire workbook…");
  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items/name");
    await ctx.sync();

    sheets.items.forEach((ws) => {
      const used = ws.getUsedRange(true); // include formats if sheet looks empty
      used.format.font.name = "Arial";
      used.format.font.size = 8;
    });

    await ctx.sync();
  });
  status("Done ✓");
}

// Hide gridlines on all worksheets (closest to 'page break off' in JS API)
async function hideGridlinesAll() {
  status("Hiding gridlines on all sheets…");
  await Excel.run(async (ctx) => {
    const sheets = ctx.workbook.worksheets;
    sheets.load("items");
    await ctx.sync();
    sheets.items.forEach((ws) => (ws.showGridlines = false));
    await ctx.sync();
  });
  status("Done ✓");
}
