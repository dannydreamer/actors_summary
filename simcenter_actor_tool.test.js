/**
 * SimCenter Actor Tool — Test Suite
 * Run: node simcenter_actor_tool.test.js
 */

'use strict';

const XLSX   = require('./node_modules/xlsx');
const parser = require('./parse.js');

const {
  isSheetSkipped,
  parseActorNames,
  buildDeduplicationMap,
  parseWorkbook,
} = parser;

// ─────────────────────────────────────────────
//  TEST HARNESS
// ─────────────────────────────────────────────

let passed = 0, failed = 0;

function assert(description, actual, expected) {
  const ok = JSON.stringify(actual) === JSON.stringify(expected);
  if (ok) {
    console.log(`  ✓  ${description}`);
    passed++;
  } else {
    console.log(`  ✗  ${description}`);
    console.log(`       expected: ${JSON.stringify(expected)}`);
    console.log(`       actual:   ${JSON.stringify(actual)}`);
    failed++;
  }
}

// ─────────────────────────────────────────────
//  UNIT TESTS: isSheetSkipped
// ─────────────────────────────────────────────

console.log('\n── isSheetSkipped ──');
assert('skips sheet containing נדחה',     isSheetSkipped('נדחה 11.2 שיעור'),    true);
assert('skips sheet containing מבוטל',    isSheetSkipped('מבוטל   15.2 כיתה'), true);
assert('skips sheet with נדחה mid-name',  isSheetSkipped('שיעור נדחה 3.3'),    true);
assert('does not skip regular sheet',     isSheetSkipped('9.2 מטה ניצנים'),    false);
assert('does not skip empty name',        isSheetSkipped(''),                   false);
assert('skips all 5 February skip sheets',
  [
    'נדחה         11.2 מדריכות לשוני',
    'נדחה 12.2 מגזר כללי רכזות ניצני',
    'נדחה       15.2 רכזי קודש חמ"ד',
    'מבוטל   15.2 תכנית קדימה',
    'מבוטל   26.2 קידום נוער - חני',
  ].every(isSheetSkipped),
  true
);

// ─────────────────────────────────────────────
//  UNIT TESTS: parseActorNames
// ─────────────────────────────────────────────

console.log('\n── parseActorNames ──');

// Spec examples
assert('comma separation',
  parseActorNames('יהודית אבדל, דניאל עובדיה, עידית עיד'),
  ['יהודית אבדל', 'דניאל עובדיה', 'עידית עיד']);

assert('ו-conjunction',
  parseActorNames('דודי גורדון ושאזו שאז'),
  ['דודי גורדון', 'שאזו שאז']);

assert('leading ו stripped',
  parseActorNames('ויעלה לוי'),
  ['יעלה לוי']);

assert('4-word even split → two 2-word names',
  parseActorNames('שאזו שאז אליאב מוסרי'),
  ['שאזו שאז', 'אליאב מוסרי']);

assert('ו-conjunction variant',
  parseActorNames('איתן החמוד וקובי ירחי'),
  ['איתן החמוד', 'קובי ירחי']);

// Edge cases
assert('trailing spaces and * stripped',
  parseActorNames('קובי ירחי*, ישי מאיר '),
  ['קובי ירחי', 'ישי מאיר']);

assert('3-word name (odd) stays intact',
  parseActorNames('קרן ברל כצנלסון'),
  ['קרן ברל כצנלסון']);

assert('3-word + comma-separated 3-word',
  parseActorNames('קרן ברל כצנלסון, תגל פישר פרייס'),
  ['קרן ברל כצנלסון', 'תגל פישר פרייס']);

assert('empty string → []',  parseActorNames(''),        []);
assert('null → []',           parseActorNames(null),      []);
assert('undefined → []',      parseActorNames(undefined), []);

// All real February cell values that exercised special paths
assert('Feb: יאיר להמן ואיתן החמוד',
  parseActorNames('יאיר להמן ואיתן החמוד'), ['יאיר להמן', 'איתן החמוד']);
assert('Feb: דניאל עובדיה ותמר תמרוני',
  parseActorNames('דניאל עובדיה ותמר תמרוני'), ['דניאל עובדיה', 'תמר תמרוני']);
assert('Feb: קובי ירחי וישי כהן',
  parseActorNames('קובי ירחי וישי כהן'), ['קובי ירחי', 'ישי כהן']);
assert('Feb: איתן החמוד ועוזי בוס',
  parseActorNames('איתן החמוד ועוזי בוס'), ['איתן החמוד', 'עוזי בוס']);
assert('Feb: יעלה, עדי זינגר (single-word pre-dedup)',
  parseActorNames('יעלה, עדי זינגר '), ['יעלה', 'עדי זינגר']);

// ─────────────────────────────────────────────
//  UNIT TESTS: buildDeduplicationMap
// ─────────────────────────────────────────────

console.log('\n── buildDeduplicationMap ──');

{
  const m = buildDeduplicationMap(['קובי ירחי', 'קובי', 'ישי מאיר']);
  assert('single-word maps to unique full name', m.get('קובי'), 'קובי ירחי');
  assert('full name not in dedup map',           m.has('קובי ירחי'), false);
}
{
  const m = buildDeduplicationMap(['יעלה לוי', 'יעלה כהן', 'יעלה']);
  assert('ambiguous first name → no dedup entry', m.has('יעלה'), false);
}
{
  const m = buildDeduplicationMap(['דניאל עובדיה', 'דניאל']);
  assert('דניאל → דניאל עובדיה when unambiguous', m.get('דניאל'), 'דניאל עובדיה');
}

// ─────────────────────────────────────────────
//  INTEGRATION: raw February data
// ─────────────────────────────────────────────

console.log('\n── Integration: raw February 2026 ──');

const rawWb = XLSX.readFile('פברואר  2026.xlsx');
const rawResult = parseWorkbook(rawWb);

assert('raw data blocked by validation (ambiguous יעלה)', rawResult.ok, false);
if (!rawResult.ok) {
  assert('exactly 1 validation error', rawResult.errors.length, 1);
  assert('error is for יעלה', rawResult.errors[0].name, 'יעלה');
  console.log(`  ℹ  error: "${rawResult.errors[0].name}" @ ${rawResult.errors[0].date} ${rawResult.errors[0].client}`);
}

// ─────────────────────────────────────────────
//  INTEGRATION: patched February data
//  (יעלה → יעלה לוי in sheet 11.2 ביהס המסורתי)
// ─────────────────────────────────────────────

console.log('\n── Integration: patched February 2026 ──');

const patchedWb = XLSX.readFile('פברואר  2026.xlsx');
{
  const ws   = patchedWb.Sheets['11.2 ביהס המסורתי'];
  const addr = XLSX.utils.encode_cell({ r: 2, c: 5 });
  if (ws[addr]) ws[addr].v = 'יעלה לוי, עדי זינגר';
  else ws[addr] = { v: 'יעלה לוי, עדי זינגר', t: 's' };
}

const r = parseWorkbook(patchedWb);

assert('patched data passes validation', r.ok, true);

if (r.ok) {
  const { workshops, actors } = r;

  assert('26 workshops (31 sheets − 5 skipped)', workshops.length, 26);
  assert('22 unique actors',                      actors.length,    22);
  assert('sorted: first date 1.2',   workshops[0].date,                  '1.2');
  assert('sorted: last date 24.2',   workshops[workshops.length-1].date, '24.2');

  const byName = Object.fromEntries(actors.map(a => [a.name, a]));

  // All expected actors present
  const expectedActors = [
    'איתן החמוד','אליאב מוסרי','דודי גורדון','דורין עטר','דניאל עובדיה',
    'טל טל','יאיר להמן','יהודית אבדל','יובל יוב','יעלה כהן','יעלה לוי',
    'ישי כהן','עדי זינגר','עוזי בוס','עידית עיד','ציפי ציפורה','קובי ירחי',
    'קרן ברל כצנלסון','רן כהן','שאזו שאז','תגל פישר פרייס','תמר תמרוני',
  ];
  assert('all 22 expected actors present',
    expectedActors.every(n => byName[n]), true);

  // Workshop counts
  const counts = [
    ['איתן החמוד', 6], ['קובי ירחי', 6], ['שאזו שאז', 4], ['דודי גורדון', 4],
    ['יהודית אבדל', 3], ['ישי כהן', 3], ['דניאל עובדיה', 3], ['עידית עיד', 3],
    ['יעלה לוי', 2], ['אליאב מוסרי', 2], ['ציפי ציפורה', 2], ['תגל פישר פרייס', 2],
    ['טל טל', 1], ['יובל יוב', 1], ['יאיר להמן', 1], ['עדי זינגר', 1],
    ['עוזי בוס', 1], ['דורין עטר', 1], ['יעלה כהן', 1],
    ['קרן ברל כצנלסון', 1], ['רן כהן', 1], ['תמר תמרוני', 1],
  ];
  assert('all actor workshop counts correct',
    counts.every(([name, n]) => byName[name]?.count === n), true);

  // Spot-check actor sets per workshop (against expected output xlsx)
  const expWb  = XLSX.readFile('actor_report_feb_2026.xlsx');
  const expRows = XLSX.utils.sheet_to_json(
    expWb.Sheets['סיכום סדנאות'], { header: 1, defval: '' }
  ).slice(1);

  assert('workshop count matches expected output', workshops.length, expRows.length);

  const actorSetMismatches = workshops.reduce((acc, w, i) => {
    const mine = w.actors.slice().sort().join(', ');
    const exp  = (expRows[i][2] === '—' ? [] : (expRows[i][2] || '').split(', ').map(s => s.trim())).sort().join(', ');
    return mine === exp ? acc : acc + 1;
  }, 0);
  assert('actor sets match expected output in all 26 workshops', actorSetMismatches, 0);

  // Actor counts match expected output xlsx
  const expActorRows = XLSX.utils.sheet_to_json(
    expWb.Sheets['סיכום שחקנים'], { header: 1, defval: '' }
  ).slice(1, -1); // strip header and סה"כ row
  const expCounts = Object.fromEntries(expActorRows.map(r => [r[0], r[1]]));
  const countMismatches = actors.filter(a => expCounts[a.name] !== undefined && a.count !== expCounts[a.name]);
  assert('all actor counts match expected output xlsx', countMismatches.length, 0);
}

// ─────────────────────────────────────────────
//  SUMMARY
// ─────────────────────────────────────────────

console.log(`\n${'─'.repeat(50)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed === 0) console.log('All tests passed! ✓');
else process.exitCode = 1;
