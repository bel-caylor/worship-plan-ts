// src/util/ai.ts
type SongLite = { name: string; lyrics?: string };

export type SongMetadataInput = {
  name: string;
  lyrics?: string;
  hints?: string[];
  forcedSeason?: string;
  kScriptures?: number;
  allowKeywords?: boolean;
};

export type SongMetadataResult = {
  themes: string[];
  keywords: string[];
  season: string;
  scriptures: string[];
  error?: string;
};

const STOP_WORDS = new Set([
  'the','and','for','with','that','this','from','into','your','you','have','will','are','our','but','all','any','when','than','then','unto','yours','him','her','them','they','ever','cada','como','para','pero','porque','con','por','sin','unos','unas','los','las','del','una','uno','sobre','donde','dios','sera','sera','esto','esta','estas','estos'
]);

const KEYWORD_BLACKLIST = new Set([
  'folder','folders','path','paths','file','files','doc','docs','wdoc','source','archive','language','spanish','advent','future','lyrics','document','documents'
]);

const SCRIPTURE_HINTS: { keys: RegExp[]; refs: string[] }[] = [
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

export function aiSongMetadata(input: SongMetadataInput): SongMetadataResult {
  const name = String(input?.name || '').trim();
  const lyrics = String(input?.lyrics || '').trim();
  const hints = Array.isArray(input?.hints) ? input!.hints.map(h => String(h || '').trim()).filter(Boolean) : [];
  const forcedSeason = String(input?.forcedSeason || '').trim();
  const kScriptures = Math.max(1, Math.min(10, Number(input?.kScriptures ?? 5)));
  const allowKeywords = !!input?.allowKeywords;

  const fallback = heuristicSongMetadata({ name, lyrics, hints, forcedSeason, kScriptures, allowKeywords });
  const key = getOpenAiKey();
  if (!key) return { ...fallback, error: 'OPENAI_API_KEY not set in Script Properties' };

  const snippet = lyrics.length > 1600 ? lyrics.slice(0, 1600) + '...' : lyrics;
  const hintsBlock = hints.length ? hints.join('\n') : '(no additional hints)';
  const seasonDirective = forcedSeason
    ? `Always set the season to exactly "${forcedSeason}".`
    : 'Infer the best-fitting liturgical season (Christmas, Advent, Easter, Lent, Pentecost, General, etc).';

  const sys = 'You are a worship-planning assistant. Given a song title and lyrics, suggest concise metadata.';
  const user = `Song title: ${name || '(unknown)'}
Lyrics snippet:
${snippet || '(lyrics unavailable)'}

Hints:
${hintsBlock}

${seasonDirective}
Return strict JSON with the following shape:
{
  "themes": ["Theme 1","Theme 2"],
  "keywords": ["Keyword 1","Keyword 2"],
  "season": "Season Name",
  "scriptures": ["Book 1:1-5","Book 2:3"]
}
Use concise phrases; limit lists to at most ${kScriptures} scripture references and 6 keywords.`;

  try {
    const res = callOpenAi(sys, user, key);
    const parsed = safeJsonParse(res);
    const themes = normalizeList(parsed?.themes).slice(0, 5);
    const keywords = allowKeywords ? sanitizeKeywords(normalizeList(parsed?.keywords).slice(0, 6), name) : [];
    const scriptures = normalizeList(parsed?.scriptures).slice(0, kScriptures);
    const seasonRaw = forcedSeason || normalizeSeason(parsed?.season || '');
    return {
      themes: themes.length ? themes : fallback.themes,
      keywords: keywords.length ? keywords : fallback.keywords,
      scriptures: scriptures.length ? scriptures : fallback.scriptures,
      season: seasonRaw || fallback.season,
    };
  } catch (err) {
    try { Logger.log('AI metadata error: ' + (err as any)?.message); } catch (_) {}
    return { ...fallback, error: 'AI metadata request failed' };
  }
}

export function aiScripturesForLyrics(input: {
  songs?: SongLite[];
  theme?: string;
  keywords?: string;
  primaryScripture?: string;
  k?: number;
  reason?: string;
}) {
  const songs = Array.isArray(input?.songs) ? input.songs : [];
  const theme = String(input?.theme || '').trim();
  const keywords = String(input?.keywords || '').trim();
  const primary = String(input?.primaryScripture || '').trim();
  const reason = String(input?.reason || '').trim();
  const k = Math.max(1, Math.min(12, Number(input?.k ?? 8)));

  const textBlock = [songs.map(s => s.lyrics || '').join('\n\n'), theme, keywords, primary].join('\n');
  const heuristics = heuristicScriptureRefs(textBlock, k, primary);
  const key = getOpenAiKey();
  if (!key) return { refs: heuristics, error: 'OPENAI_API_KEY not set in Script Properties' };

  const summarize = (s: SongLite) => {
    const title = String(s?.name || '').trim();
    let lyr = String(s?.lyrics || '').trim();
    if (lyr.length > 1200) lyr = lyr.slice(0, 1200) + '...';
    return `- ${title}: ${lyr || '(no lyrics available)'}\n`;
  };
  const songsBlock = songs.length ? songs.map(summarize).join('') : '(no song lyrics provided)';
  const contextLines: string[] = [];
  if (theme) contextLines.push(`Theme: ${theme}`);
  if (keywords) contextLines.push(`Keywords: ${keywords}`);
  if (primary) contextLines.push(`Main Scripture: ${primary}`);
  const contextBlock = contextLines.length ? contextLines.join('\n') : '(no additional context)';
  const reasonLine = reason ? `Reason: ${reason}` : '';
  const sys = `You are a helpful worship-planning assistant. Suggest concise Scripture references (book chapter:verses) that complement the provided songs and planning context. Prioritize clarity and relevance, avoid duplicates.`;
  const user = `Songs and lyrics:
${songsBlock}

Planning context:
${contextBlock}
${reasonLine ? `\n${reasonLine}\n` : '\n'}Return a JSON array of up to ${k} references, example: ["John 1:1-5", "Psalm 27:1"].`;

  try {
    const res = callOpenAi(sys, user, key);
    const parsed = safeJsonParse(res);
    const refs = normalizeList(parsed);
    if (refs.length) return { refs: refs.slice(0, k) };
    return { refs: heuristics };
  } catch (err) {
    try { Logger.log('AI error: ' + (err as any)?.message); } catch (_) {}
    return { refs: heuristics, error: 'AI request failed' };
  }
}

let _lastAiCallTs = 0;
const AI_MIN_DELAY_MS = 1200;

function throttleAiRequests() {
  const now = Date.now();
  const wait = Math.max(0, _lastAiCallTs + AI_MIN_DELAY_MS - now);
  if (wait > 0) Utilities.sleep(wait);
  _lastAiCallTs = Date.now();
}

function callOpenAi(systemPrompt: string, userPrompt: string, key: string) {
  const payload = {
    model: 'gpt-4o-mini',
    temperature: 0.2,
    messages: [
      { role: 'system', content: systemPrompt },
      { role: 'user', content: userPrompt }
    ]
  } as any;
  throttleAiRequests();
  let res: GoogleAppsScript.URL_Fetch.HTTPResponse;
  try {
    res = UrlFetchApp.fetch('https://api.openai.com/v1/chat/completions', {
      method: 'post',
      headers: { 'Authorization': 'Bearer ' + key, 'Content-Type': 'application/json' },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (err) {
    throw new Error(`AI HTTP error: ${(err as any)?.message || err}`);
  }
  const status = res.getResponseCode();
  const body = res.getContentText();
  let data: any = null;
  try { data = JSON.parse(body); } catch (_) {}
  if (status >= 400) {
    const detail = data?.error?.message || `HTTP ${status}`;
    throw new Error(detail);
  }
  const text = data?.choices?.[0]?.message?.content;
  if (text) return text;
  if (data?.error?.message) throw new Error(data.error.message);
  throw new Error('Empty AI response');
}

function safeJsonParse(text: string | string[]) {
  if (Array.isArray(text)) return text;
  const raw = String(text || '').trim();
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (_) {
    const match = /\{[\s\S]*\}|\[[\s\S]*\]/.exec(raw);
    if (match) {
      try { return JSON.parse(match[0]); } catch (_) { /* ignore */ }
    }
  }
  return raw.split(/\n|,|;|\|/).map(s => String(s).trim()).filter(Boolean);
}

function getOpenAiKey() {
  return String(PropertiesService.getScriptProperties().getProperty('OPENAI_API_KEY') || '').trim();
}

function normalizeList(value: unknown): string[] {
  if (Array.isArray(value)) {
    return value
      .map(v => String(v || '').trim())
      .filter(Boolean)
      .map(normalizeCase);
  }
  if (typeof value === 'string') {
    return value
      .split(/[,;\n|]+/)
      .map(v => String(v || '').trim())
      .filter(Boolean)
      .map(normalizeCase);
  }
  return [];
}

function normalizeCase(value: string) {
  const trimmed = value.replace(/\s+/g, ' ').trim();
  if (!trimmed) return '';
  return trimmed.replace(/\b\w/g, c => c.toUpperCase());
}

function normalizeSeason(value: string) {
  const s = value.toLowerCase();
  if (!s) return '';
  if (s.includes('christmas') || s.includes('navidad')) return 'Christmas';
  if (s.includes('advent')) return 'Advent';
  if (s.includes('lent') || s.includes('cuaresma')) return 'Lent';
  if (s.includes('easter') || s.includes('resurrection') || s.includes('pascua')) return 'Easter';
  if (s.includes('pentecost')) return 'Pentecost';
  return s.replace(/\b\w/g, c => c.toUpperCase());
}

function heuristicSongMetadata(input: {
  name: string;
  lyrics: string;
  hints: string[];
  forcedSeason?: string;
  kScriptures: number;
  allowKeywords?: boolean;
}): SongMetadataResult {
  const baseText = [input.name, input.lyrics, input.hints.join(' ')].join('\n').toLowerCase();
  const season = input.forcedSeason || detectSeasonFromText(baseText);
  const themes = detectThemesFromText(baseText, season);
  const keywords = input.allowKeywords ? extractKeywords(input.lyrics, 6, input.name) : [];
  const scriptures = heuristicScriptureRefs(baseText, input.kScriptures);
  return {
    themes,
    keywords: sanitizeKeywords(keywords, input.name),
    season,
    scriptures
  };
}

function detectSeasonFromText(text: string) {
  if (/advent|emmanuel|manger|bethlehem|noel|shepherds|angel|incarnat/.test(text)) return 'Christmas';
  if (/christmas|navidad/.test(text)) return 'Christmas';
  if (/lent|ashes|fast|forty\s+days/.test(text)) return 'Lent';
  if (/easter|resurrect|empty\s+tomb|grave|cross|crucif/.test(text)) return 'Easter';
  if (/pentecost|spirit\s+fire|tongues/.test(text)) return 'Pentecost';
  return '';
}

function detectThemesFromText(text: string, season: string) {
  const cues: { rx: RegExp; label: string }[] = [
    { rx: /(worship|adore|glorif|praise)/, label: 'Adoration' },
    { rx: /(hope|wait|longing|expect)/, label: 'Hope' },
    { rx: /(joy|rejoice|glad)/, label: 'Joy' },
    { rx: /(peace|shalom|calm|rest)/, label: 'Peace' },
    { rx: /(light|shine|star|radiant)/, label: 'Light' },
    { rx: /(grace|mercy|forgive|forgiven)/, label: 'Grace' },
    { rx: /(cross|blood|lamb|sacrifice|redeem)/, label: 'Redemption' },
    { rx: /(spirit|breath|wind|fire)/, label: 'Holy Spirit' },
    { rx: /(nations|every\s+tongue|mission|send|go)/, label: 'Mission' },
    { rx: /(king|kingdom|reign|throne)/, label: 'Kingship' },
    { rx: /(victory|conquer|triumph)/, label: 'Victory' },
  ];
  const out: string[] = [];
  for (const cue of cues) {
    if (cue.rx.test(text)) out.push(cue.label);
  }
  if (season === 'Christmas') out.unshift('Incarnation');
  if (!out.length) out.push('Worship');
  return Array.from(new Set(out)).slice(0, 4);
}

function extractKeywords(lyrics: string, limit: number, title?: string) {
  const freq: Record<string, number> = {};
  const corpus = String(lyrics || '').toLowerCase();
  if (!corpus.trim()) return [];
  const words = corpus.match(/[a-záéíóúñü]+/gi) || [];
  const titleTokens = new Set(normalizeTitleTokens(title));
  for (const word of words) {
    const w = word.trim();
    if (w.length < 4) continue;
    if (STOP_WORDS.has(w)) continue;
    if (KEYWORD_BLACKLIST.has(w)) continue;
    if (titleTokens.has(w)) continue;
    freq[w] = (freq[w] || 0) + 1;
  }
  const sorted = Object.keys(freq).sort((a, b) => freq[b] - freq[a]);
  return sorted.slice(0, limit).map(w => w.replace(/\b\w/, c => c.toUpperCase()));
}

function heuristicScriptureRefs(text: string, k: number, primary?: string) {
  const out: string[] = [];
  const seen = new Set<string>();
  const push = (val: string) => {
    const v = String(val || '').trim();
    const key = v.toLowerCase();
    if (v && !seen.has(key)) {
      seen.add(key);
      out.push(v);
    }
  };
  if (primary) push(primary);
  const lower = String(text || '').toLowerCase();
  for (const hint of SCRIPTURE_HINTS) {
    if (hint.keys.some(rx => rx.test(lower))) {
      hint.refs.forEach(push);
      if (out.length >= k) break;
    }
  }
  return out.slice(0, k);
}

function sanitizeKeywords(keywords: string[], title: string) {
  const titleTokens = new Set(normalizeTitleTokens(title));
  const out: string[] = [];
  const seen = new Set<string>();
  for (const phrase of keywords) {
    const raw = String(phrase || '').trim();
    if (!raw) continue;
    const lowerTokens = raw.toLowerCase().split(/\s+/).filter(Boolean);
    if (!lowerTokens.length) continue;
    if (lowerTokens.every(tok => KEYWORD_BLACKLIST.has(tok) || titleTokens.has(tok))) continue;
    if (lowerTokens.some(tok => KEYWORD_BLACKLIST.has(tok) || titleTokens.has(tok))) continue;
    const normalized = raw.replace(/\s+/g, ' ').trim();
    const key = normalized.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(normalized);
  }
  return out;
}

function normalizeTitleTokens(title?: string) {
  return String(title || '')
    .toLowerCase()
    .replace(/\([^)]*\)|\[[^\]]*\]/g, ' ')
    .replace(/[^a-z0-9\s]/g, ' ')
    .split(/\s+/)
    .filter(Boolean);
}
