import { doGet, doPost, doOptions } from './http';
import { onOpen as menuOnOpen, showMenuNow as menuShow, setupMenuTrigger as menuSetup } from './menu';
import { linkSongMedia, rebuildSongUsageFromPlanner, syncSongsFromDrive } from './features/songs';
import { buildLeadersFromPlanner } from './features/leaders';
import { repairServiceDateTimeColumns, syncYouTubeStreamsCatalog, matchServicesFromYouTubeStreams, resetYouTubeStreamsSyncState } from './features/services';
import { getFilesForFolderUrl } from './util/drive';
import { rpc } from './rpc';

// Top-level wrappers so Apps Script Run menu can see them
export function onOpen() {
try { menuOnOpen(); } catch (e) { try { Logger.log(e); } catch (_) {} }
}

export function showMenuNow() {
authorizeOrderDocExport();
try { menuShow(); } catch (e) { try { Logger.log(e); } catch (_) {} }
}

export function setupMenuTrigger() {
try { menuSetup(); } catch (e) { try { Logger.log(e); } catch (_) {} }
}

// Run once from the Apps Script editor to prompt for the Drive scope needed by
// the Order of Worship DOCX export flow.
export function authorizeOrderDocExport() {
  const file = DriveApp.createFile(
    `Order Doc Export Auth ${new Date().toISOString()}.txt`,
    'Authorization check for Order of Worship DOCX export.'
  );
  const id = file.getId();
  file.setTrashed(true);
  return { ok: true, fileId: id };
}

// Expose to GAS global so web app + Run menu can call them
declare const global: any;
global.doGet = doGet;
global.doPost = doPost;
global.doOptions = doOptions;
global.rpc = rpc;
global.getFilesForFolderUrl = getFilesForFolderUrl;

global.linkSongMedia = linkSongMedia;
global.buildLeadersFromPlanner = buildLeadersFromPlanner;
global.rebuildSongUsageFromPlanner = rebuildSongUsageFromPlanner;
global.syncSongsFromDrive = syncSongsFromDrive;
global.repairServiceDateTimeColumns = repairServiceDateTimeColumns;
global.syncYouTubeStreamsCatalog = syncYouTubeStreamsCatalog;
global.resetYouTubeStreamsSyncState = resetYouTubeStreamsSyncState;
global.matchServicesFromYouTubeStreams = matchServicesFromYouTubeStreams;

global.onOpen = onOpen;
global.showMenuNow = showMenuNow;
global.setupMenuTrigger = setupMenuTrigger;
global.authorizeOrderDocExport = authorizeOrderDocExport;
