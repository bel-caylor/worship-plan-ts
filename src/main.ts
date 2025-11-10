import { doGet, doPost } from './http';
import { onOpen as menuOnOpen, showMenuNow as menuShow, setupMenuTrigger as menuSetup } from './menu';
import { linkSongMedia, rebuildSongUsageFromPlanner, syncSongsFromDrive } from './features/songs';
import { buildLeadersFromPlanner } from './features/leaders';
import { getFilesForFolderUrl } from './util/drive';
import { rpc } from './rpc';

// Top-level wrappers so Apps Script Run menu can see them
export function onOpen() {
try { menuOnOpen(); } catch (e) { try { Logger.log(e); } catch (_) {} }
}

export function showMenuNow() {
try { menuShow(); } catch (e) { try { Logger.log(e); } catch (_) {} }
}

export function setupMenuTrigger() {
try { menuSetup(); } catch (e) { try { Logger.log(e); } catch (_) {} }
}

// Expose to GAS global so web app + Run menu can call them
declare const global: any;
global.doGet = doGet;
global.doPost = doPost;
global.rpc = rpc;
global.getFilesForFolderUrl = getFilesForFolderUrl;

global.linkSongMedia = linkSongMedia;
global.buildLeadersFromPlanner = buildLeadersFromPlanner;
global.rebuildSongUsageFromPlanner = rebuildSongUsageFromPlanner;
global.syncSongsFromDrive = syncSongsFromDrive;

global.onOpen = onOpen;
global.showMenuNow = showMenuNow;
global.setupMenuTrigger = setupMenuTrigger;
