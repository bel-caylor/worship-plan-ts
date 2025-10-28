// src/menu.ts
export function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Worship')
    .addItem('Link song media', 'linkSongMedia')
    .addItem('Build leader list (from Usage Log)', 'buildLeadersFromPlanner')
    .addToUi();
}
