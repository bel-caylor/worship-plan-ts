// src/features/scripture.ts
import { recommendSongsForPassage, summarizeScriptureThemes, SongRecommendation } from '../util/ai';
import { loadSongLyricSamples } from './song-ai';

type SummaryInput = {
  text: string;
  reference?: string;
  songLimit?: number;
  leader?: string;
};

export function summarizePassageWithSongs(input: SummaryInput) {
  const text = String(input?.text || '').trim();
  if (!text) throw new Error('Passage text is required');
  const reference = String(input?.reference || '').trim();
  const songLimit = Math.max(1, Math.min(12, Number(input?.songLimit ?? 10)));
  const leader = String(input?.leader || '').trim();

  const summaryResult = summarizeScriptureThemes({ text, reference });

  let recommendedSongs: SongRecommendation[] = [];
  let recommendedSongsError = '';
  try {
    const catalog = loadSongLyricSamples(songLimit * 3, 420, leader);
    if (catalog.length) {
      const aiResult = recommendSongsForPassage({
        passageText: text,
        reference,
        songs: catalog,
        limit: songLimit
      });
      recommendedSongs = Array.isArray(aiResult?.songs) ? aiResult.songs : [];
      if (aiResult?.error) recommendedSongsError = aiResult.error;
    } else {
      recommendedSongsError = 'No songs with lyrics are available';
    }
  } catch (err) {
    recommendedSongsError = (err && (err as any).message) ? (err as any).message : 'Failed to load songs for recommendations';
  }

  return {
    summary: summaryResult.summary,
    summaryError: summaryResult.error,
    recommendedSongs,
    recommendedSongsError
  };
}
