// src/util/ai.ts
import { Row } from '../constants';

type SongLite = { name: string; lyrics?: string };

export function aiScripturesForLyrics(input: { songs: SongLite[]; theme?: string; k?: number }) {
  const songs = Array.isArray(input?.songs) ? input.songs : [];
  const theme = String(input?.theme || '').trim();
  const k = Number(input?.k ?? 8);

  const key = String(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '').trim();
  if (!key) return { refs: [], error: 'OPENAI_API_KEY not set in Script Properties' };

  const summarize = (s: SongLite) => {
    const name = String(s?.name || '').trim();
    let lyr = String(s?.lyrics || '').trim();
    if (lyr.length > 1200) lyr = lyr.slice(0, 1200) + '...';
    return `- ${name}: ${lyr || '(no lyrics available)'}\n`;
  };
  const songsBlock = songs.map(summarize).join('');
  const sys = `You are a helpful worship-planning assistant. Suggest concise Scripture references (book chapter:verses) that complement the provided song lyrics and theme. Prioritize clarity and relevance, avoid duplicates.`;
  const user = `Songs and lyrics:\n${songsBlock}\nTheme: ${theme || '(none)'}\n\nReturn a JSON array of up to ${k} references, example: ["John 1:1-5", "Psalm 27:1"].`;

  // Local heuristic fallback using keywords in lyrics/theme
  const deriveRefs = (): string[] => {
    const map: { keys: RegExp[]; refs: string[] }[] = [
      { keys: [/\blight\b/, /bright/, /shine/], refs: ['John 1:1-5','1 John 1:5-7','Psalm 27:1','Ephesians 5:8-14','2 Corinthians 4:6'] },
      { keys: [/\bkingdom\b/, /\breign\b/, /throne/], refs: ['Daniel 7:13-14','Revelation 11:15','Psalm 145:11-13'] },
      { keys: [/nation/, /every\s+knee/, /bow/, /tongue/], refs: ['Philippians 2:9-11','Revelation 7:9-10'] },
      { keys: [/\blove\b/], refs: ['Romans 5:8','1 John 4:9-10','John 3:16'] },
      { keys: [/\bgrace\b/], refs: ['Ephesians 2:8-9','Titus 2:11'] },
      { keys: [/\bmercy\b/], refs: ['Psalm 103:8-12','Lamentations 3:22-23'] },
      { keys: [/righteous/, /clothed/], refs: ['2 Corinthians 5:21','Isaiah 61:10'] },
      { keys: [/free(dom)?/, /chains?\s+fell|set\s+free/], refs: ['Galatians 5:1','John 8:36'] },
      { keys: [/ancient\s+of\s+days/], refs: ['Daniel 7:9-14'] },
    ];
    const text = [songs.map(s => (s.lyrics || '')).join('\n\n'), theme].join('\n').toLowerCase();
    const out: string[] = [];
    const seen = new Set<string>();
    const push = (r: string) => { const v = r.trim(); const key = v.toLowerCase(); if (v && !seen.has(key)) { seen.add(key); out.push(v); } };
    for (const m of map) {
      if (m.keys.some(rx => rx.test(text))) m.refs.forEach(push);
      if (out.length >= k) break;
    }
    return out.slice(0, k);
  };

  try {
    const url = 'https://api.openai.com/v1/chat/completions';
    const payload = {
      model: 'gpt-4o-mini',
      temperature: 0.2,
      messages: [
        { role: 'system', content: sys },
        { role: 'user', content: user }
      ]
    } as any;
    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
    const data = JSON.parse(res.getContentText());
    const text = (data && data.choices && data.choices[0] && data.choices[0].message && data.choices[0].message.content) ? String(data.choices[0].message.content) : '';
    let refs: string[] = [];
    try {
      refs = JSON.parse(text);
      if (!Array.isArray(refs)) refs = [];
    } catch (_) {
      // try to extract JSON array from within fences
      const m = /\[\s*"[\s\S]*?\]/.exec(text);
      if (m) {
        try { refs = JSON.parse(m[0]); } catch (_) {}
      }
      if (!Array.isArray(refs)) {
        refs = text.split(/\n|,|;|\|/).map(s => String(s).trim()).filter(v => /[A-Za-z]/.test(v));
      }
    }
    // de-dupe and limit
    const out: string[] = [];
    const seen = new Set<string>();
    for (const r of refs) { const v = String(r || '').trim(); if (v && !seen.has(v.toLowerCase())) { seen.add(v.toLowerCase()); out.push(v); if (out.length >= k) break; } }
    if (out.length > 0) return { refs: out };
    // Fallback heuristics if AI didn't return anything
    const heur = deriveRefs();
    return { refs: heur };
  } catch (e) {
    try { Logger.log('AI error: ' + (e as any)?.message); } catch(_) {}
    // Even if AI fails, offer heuristic suggestions
    const heur = deriveRefs();
    return { refs: heur, error: 'AI request failed' };
  }
}
