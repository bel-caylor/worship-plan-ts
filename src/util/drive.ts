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
    const extractId = (url: string) => {
        const direct = /\/folders\/([A-Za-z0-9_-]+)/.exec(url);
        if (direct) return direct[1];
        const query = /[?&]id=([A-Za-z0-9_-]+)/.exec(url);
        if (query) return query[1];
        return null;
    };
    const folderId = extractId(folderUrl);
    if (!folderId) return [];
    let folder: GoogleAppsScript.Drive.Folder;
    try {
        folder = DriveApp.getFolderById(folderId);
    } catch (_) {
        return [];
    }
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
  const songTokens = new Set(normSong.split(' ').filter(Boolean));
  const exact = folders.find(f => normalize(f.name) === normSong);
  if (exact) return { ...exact, score: 1 };
  const contains = folders.find(f => normalize(f.name).includes(normSong));
  if (contains) return { ...contains, score: 0.9 };
  let best: { id: string; name: string; url: string; score: number; overlap: number } | null = null;
  for (const f of folders) {
    const normalizedName = normalize(f.name);
    const s = similarity(normalizedName, normSong);
    const folderTokens = new Set(normalizedName.split(' ').filter(Boolean));
    let overlap = 0;
    for (const tok of songTokens) if (folderTokens.has(tok)) overlap++;
    if (!best || s > best.score) best = { ...f, score: s, overlap };
  }
  if (!best) return null;
  if (best.score >= 0.65) return best;
  if (best.overlap > 0) return best;
  return null;
}
