// src/util/drive.ts
import { AUDIO_EXT, AUDIO_MIME_PREFIX } from '../constants';
import { normalize, similarity } from './text';

let _folderCacheByRoot: Record<string, { id:string; name:string; url:string }[] | null> = {};
export function ensureFolderCache(rootId: string) {
  if (!_folderCacheByRoot[rootId]) _folderCacheByRoot[rootId] = listAllSubfolders(rootId);
  return _folderCacheByRoot[rootId]!;
}

export function listAllSubfolders(rootId: string) {
    const root = DriveApp.getFolderById(rootId);
    const stack = [root];
    const all: { id: string; name: string; url: string }[] = [];

    while (stack.length) {
        const folder = stack.pop()!;
        const subfolders = folder.getFolders();
        while (subfolders.hasNext()) {
            const f = subfolders.next();
            all.push({ id: f.getId(), name: f.getName(), url: f.getUrl() });
            stack.push(f);
        }
    }
    return all;
}

export function listAudioInFolder(folderId: string, limit: number) {
    const folder = DriveApp.getFolderById(folderId);
    const out: { name: string; url: string }[] = [];
    const files = folder.getFiles();
    while (files.hasNext() && out.length < limit) {
        const file = files.next();
        const mime = file.getMimeType() || '';
        const name = file.getName();
        const ext = (name.split('.').pop() || '').toLowerCase();
        if (mime.startsWith(AUDIO_MIME_PREFIX) || AUDIO_EXT.has(ext)) {
            out.push({ name, url: file.getUrl() });
        }
    }
    return out;
}

export function getFilesForFolderUrl(folderUrl: string, limit: number = 200) {
    if (!folderUrl) return [];
    const m = /\/folders\/([A-Za-z0-9_-]+)/.exec(folderUrl);
    if (!m) return [];
    const folder = DriveApp.getFolderById(m[1]);
    const out: Array<{ name: string; url: string; mimeType: string }> = [];
    const it = folder.getFiles();
    while (it.hasNext() && out.length < limit) {
        const f = it.next();
        out.push({ name: f.getName(), url: f.getUrl(), mimeType: f.getMimeType() });
    }
    out.sort((a, b) => a.name.localeCompare(b.name));
    return out;
}

export function findBestFolderForSong(songName: string, rootId: string) {
  const folders = ensureFolderCache(rootId);
  const normSong = normalize(songName);
  const exact = folders.find(f => normalize(f.name) === normSong);
  if (exact) return { ...exact, score: 1 };
  const contains = folders.find(f => normalize(f.name).includes(normSong));
  if (contains) return { ...contains, score: 0.9 };
  let best: any=null;
  for (const f of folders) {
    const s = similarity(normalize(f.name), normSong);
    if (!best || s > best.score) best = { ...f, score: s };
  }
  return best && best.score >= 0.35 ? best : null;
}
