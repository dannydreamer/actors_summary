/**
 * SimCenter Actor Tool — Core Parser
 * Works in both browser (global export) and Node.js (module.exports).
 */

(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory();
  } else {
    root.SimCenterParser = factory();
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function () {
  'use strict';

  // ─── Skip logic ───────────────────────────────────────────────────────────

  function isSheetSkipped(sheetName) {
    const n = sheetName.toLowerCase();
    return n.includes('נדחה') || n.includes('מבוטל');
  }

  // ─── Name parsing ─────────────────────────────────────────────────────────

  /**
   * Parse a raw actor cell value into an array of name strings.
   * Handles:
   *   - Comma separation
   *   - Hebrew ו-conjunction ( ו as separator, e.g. "קובי ירחי ושאזו שאז")
   *   - Leading ו glued to first name (e.g. "ויעלה לוי" → "יעלה לוי")
   *   - Even-word-count chunks with no separator → split into 2-word pairs
   *     (handles "שאזו שאז אליאב מוסרי" → ["שאזו שאז", "אליאב מוסרי"])
   *   - * stripping and whitespace trimming
   */
  function parseActorNames(cellValue) {
    if (!cellValue && cellValue !== 0) return [];
    const raw = String(cellValue).trim();
    if (!raw) return [];

    // Strip leading ו if glued (no space) to first character
    const stripped = raw.startsWith('ו') && raw.length > 1 && raw[1] !== ' '
      ? raw.slice(1)
      : raw;

    const names = [];
    const commaParts = stripped.split(',');

    for (const part of commaParts) {
      const wawParts = part.split(' ו');
      for (const chunk of wawParts) {
        const trimmed = chunk.replace(/\*/g, '').trim();
        if (!trimmed) continue;

        const words = trimmed.split(/\s+/).filter(Boolean);

        // Even-word-count chunks with 4+ words and no separator → 2-word pairs
        if (words.length >= 4 && words.length % 2 === 0) {
          for (let i = 0; i < words.length; i += 2) {
            const name = words[i] + ' ' + words[i + 1];
            names.push(name);
          }
        } else {
          if (trimmed.length >= 2) names.push(trimmed);
        }
      }
    }

    return names;
  }

  // ─── Deduplication ────────────────────────────────────────────────────────

  /**
   * Build a map of single-word name → canonical full name.
   * Only maps when exactly ONE full name starts with that first word.
   * If two full names share a first word, the single-word form is left alone
   * (it will fail validation, prompting the user to fix the source sheet).
   */
  function buildDeduplicationMap(allNames) {
    const firstWordToFull = new Map();
    for (const name of allNames) {
      const words = name.trim().split(/\s+/);
      if (words.length >= 2) {
        const fw = words[0];
        if (!firstWordToFull.has(fw)) firstWordToFull.set(fw, new Set());
        firstWordToFull.get(fw).add(name.trim());
      }
    }
    const dedup = new Map();
    for (const name of allNames) {
      const words = name.trim().split(/\s+/);
      if (words.length === 1) {
        const candidates = firstWordToFull.get(name.trim());
        if (candidates && candidates.size === 1) {
          dedup.set(name.trim(), [...candidates][0]);
        }
      }
    }
    return dedup;
  }

  function canonicalize(name, dedupMap) {
    return dedupMap.get(name.trim()) || name.trim();
  }

  // ─── Date helpers ─────────────────────────────────────────────────────────

  function formatDate(raw) {
    if (raw === '' || raw == null) return '';
    return String(raw).trim();
  }

  function compareDates(a, b) {
    const parse = d => {
      const [day, month] = d.split('.').map(Number);
      return month * 100 + day;
    };
    return parse(a) - parse(b);
  }

  function inferMonthYear(workshops) {
    for (const w of workshops) {
      const parts = w.date.split('.');
      if (parts.length === 2) {
        return {
          month: parts[1].padStart(2, '0'),
          year: new Date().getFullYear(),
        };
      }
    }
    return { month: '00', year: new Date().getFullYear() };
  }

  // ─── Main workbook parser ─────────────────────────────────────────────────

  /**
   * Parse a SheetJS workbook object.
   *
   * Returns one of:
   *   { ok: false, errors: [{ name, date, client }] }
   *   { ok: true,  workshops: [...], actors: [...] }
   *
   * workshops: [{ date, client, actors: string[] }]  — sorted by date
   * actors:    [{ name, count, clients: string[] }]   — sorted alphabetically
   */
  function parseWorkbook(workbook) {
    const rawWorkshopData = [];
    const allRawNames = [];

    for (const sheetName of workbook.SheetNames) {
      if (isSheetSkipped(sheetName)) continue;

      const ws = workbook.Sheets[sheetName];
      // XLSX must be available as a global (browser) or required externally (Node)
      const XLSX = typeof module !== 'undefined' ? require('./node_modules/xlsx') : window.XLSX;
      const rows = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });

      const row2 = rows[1] || [];
      const row3 = rows[2] || [];

      const rawDate   = row2[0];
      const rawClient = String(row2[1] || '').split('\n')[0].trim();
      const actorNames = [
        ...parseActorNames(row3[4]),
        ...parseActorNames(row3[5]),
      ];

      allRawNames.push(...actorNames);
      rawWorkshopData.push({ sheetName, rawDate, rawClient, actorNames });
    }

    const dedupMap = buildDeduplicationMap(allRawNames);

    // Validation pass
    const errors = [];
    for (const { rawDate, rawClient, actorNames } of rawWorkshopData) {
      for (const name of actorNames) {
        const canonical = canonicalize(name, dedupMap);
        if (canonical.trim().split(/\s+/).filter(Boolean).length < 2) {
          errors.push({ name: canonical, date: formatDate(rawDate), client: rawClient });
        }
      }
    }

    if (errors.length > 0) return { ok: false, errors };

    // Build output
    const workshops = [];
    const actorMap  = new Map();

    for (const { rawDate, rawClient, actorNames } of rawWorkshopData) {
      const canonicalActors = [...new Set(actorNames.map(n => canonicalize(n, dedupMap)))];

      workshops.push({ date: formatDate(rawDate), client: rawClient, actors: canonicalActors });

      for (const actor of canonicalActors) {
        if (!actorMap.has(actor)) actorMap.set(actor, { count: 0, clients: [] });
        actorMap.get(actor).count += 1;
        actorMap.get(actor).clients.push(rawClient);
      }
    }

    workshops.sort((a, b) => compareDates(a.date, b.date));

    const actors = [...actorMap.entries()]
      .sort((a, b) => a[0].localeCompare(b[0], 'he'))
      .map(([name, { count, clients }]) => ({ name, count, clients }));

    return { ok: true, workshops, actors };
  }

  // ─── Public API ───────────────────────────────────────────────────────────

  return {
    isSheetSkipped,
    parseActorNames,
    buildDeduplicationMap,
    parseWorkbook,
    inferMonthYear,
  };
}));
