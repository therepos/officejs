const status = (m)=>{ const el=document.getElementById("status"); if(el) el.textContent=m||""; };

Office.onReady(()=> {
  document.getElementById("btnArial8").onclick = setWorkbookArial8;
  document.getElementById("btnGridOff").onclick = hideGridlinesAll;
});

async function setWorkbookArial8(){
  status("Applying Arial 8…");
  await Excel.run(async (ctx)=>{
    const sheets = ctx.workbook.worksheets; sheets.load("items");
    await ctx.sync();
    sheets.items.forEach(ws=>{
      const used = ws.getUsedRange(true);
      used.format.font.name = "Arial";
      used.format.font.size = 8;
    });
    await ctx.sync();
  });
  status("Done ✓");
}

async function hideGridlinesAll(){
  status("Hiding gridlines…");
  await Excel.run(async (ctx)=>{
    const sheets = ctx.workbook.worksheets; sheets.load("items");
    await ctx.sync();
    sheets.items.forEach(ws => ws.showGridlines = false);
    await ctx.sync();
  });
  status("Done ✓");
}
