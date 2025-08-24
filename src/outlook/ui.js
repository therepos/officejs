Office.onReady(() => {
  const F = window.OfficeJsFeatures;
  document.getElementById('btnFormat').onclick   = F.formatTables;
  document.getElementById('btnImg60').onclick    = F.resizeImages60;
  document.getElementById('btnArial11').onclick  = F.setArial11;
  document.getElementById('btnTimes').onclick    = F.setTimes;
  document.getElementById('btnSelBlue').onclick  = F.setSelectionBlue;
});
