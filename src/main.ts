import { doGet } from './http';
import { onOpen } from './menu';
import { linkSongMedia } from './features/songs';
import { buildLeadersFromPlanner } from './features/leaders';
import { getFilesForFolderUrl } from './util/drive';
import { rpc } from './rpc';

declare const global: any;
global.doGet = doGet;
global.onOpen = onOpen;
global.linkSongMedia = linkSongMedia;
global.buildLeadersFromPlanner = buildLeadersFromPlanner;
global.getFilesForFolderUrl = getFilesForFolderUrl;
global.rpc = rpc;
