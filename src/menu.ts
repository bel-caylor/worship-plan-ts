// src/menu.ts
export function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Worship')
    .addItem('Link song media', 'linkSongMedia')
    .addItem('Sync songs from Drive', 'syncSongsFromDrive')
    .addItem('Build leader list (from Usage Log)', 'buildLeadersFromPlanner')
    .addItem('Rebuild song Usage (from Weekly Planner)', 'rebuildSongUsageFromPlanner')
    .addToUi();
}




export function showMenuNow() { try { onOpen(); } catch (e) { try { Logger.log(e); SpreadsheetApp.getActive().toast(String(e)); } catch(_) {} } }

export function setupMenuTrigger() {
  try {
    const ss = SpreadsheetApp.getActive();
    const triggers = ScriptApp.getProjectTriggers() || [];
    const has = triggers.some(t => t.getHandlerFunction && t.getHandlerFunction() === 'onOpen');
    if (!has) {
      ScriptApp.newTrigger('onOpen').forSpreadsheet(ss).onOpen().create();
    }
    SpreadsheetApp.getActive().toast('Worship menu trigger installed');
  } catch (e) {
    try { Logger.log(e); SpreadsheetApp.getActive().toast('Failed to install trigger: '+e); } catch(_) {}
  }
}
