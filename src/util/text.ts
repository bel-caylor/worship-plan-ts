// src/util/text.ts
const norm = (s: string) => String(s).toLowerCase().replace(/\s+/g, '');

export function splitTokens(s: string) {
    return String(s || '')
        .split(/[\/,;|&]|,\s*/g)
        .map(t => t.trim())
        .filter(Boolean);
}

export function normalize(s: string) {
    return s
        .toLowerCase()
        .replace(/[\p{P}\p{S}]+/gu, ' ')
        .replace(/\s+/g, ' ')
        .trim();
}

export function levenshtein(a: string, b: string) {
    const m = a.length, n = b.length;
    if (m === 0) return n;
    if (n === 0) return m;

    const dp = new Array(n + 1);
    for (let j = 0; j <= n; j++) dp[j] = j;

    for (let i = 1; i <= m; i++) {
        let prev = dp[0];
        dp[0] = i;
        for (let j = 1; j <= n; j++) {
            const tmp = dp[j];
            dp[j] = Math.min(
                dp[j] + 1,              // deletion
                dp[j - 1] + 1,          // insertion
                prev + (a[i - 1] === b[j - 1] ? 0 : 1) // substitution
            );
            prev = tmp;
        }
    }
    return dp[n];
}

export function similarity(a: string, b: string) {
    const d = levenshtein(a, b);
    const maxLen = Math.max(a.length, b.length) || 1;
    return 1 - d / maxLen;
}
