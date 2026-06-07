/**
 * SimCenter Actor Tool — Test Suite (Sessions 1, 2, 3)
 * Run: node simcenter_actor_tool.test.js
 */

'use strict';

const XLSX      = require('./node_modules/xlsx');
const parser    = require('./parse.js');
const generator = require('./generate.js');

const { isSheetSkipped, parseActorNames, buildDeduplicationMap, parseWorkbook } = parser;
const { buildWorkbook, inferMonthYear } = generator;

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
//  SESSION 1 — PARSER / VALIDATOR
// ─────────────────────────────────────────────

console.log('\n── isSheetSkipped ──');
assert('skips sheet containing נדחה',    isSheetSkipped('נדחה 11.2 שיעור'),    true);
assert('skips sheet containing מבוטל',   isSheetSkipped('מבוטל   15.2 כיתה'), true);
assert('skips sheet with נדחה mid-name', isSheetSkipped('שיעור נדחה 3.3'),    true);
assert('does not skip regular sheet',    isSheetSkipped('9.2 מטה ניצנים'),    false);
assert('does not skip empty name',       isSheetSkipped(''),                   false);
assert('skips all 5 February skip sheets',
  ['נדחה         11.2 מדריכות לשוני', 'נדחה 12.2 מגזר כללי רכזות ניצני',
   'נדחה       15.2 רכזי קודש חמ"ד',  'מבוטל   15.2 תכנית קדימה',
   'מבוטל   26.2 קידום נוער - חני'].every(isSheetSkipped), true);

console.log('\n── parseActorNames ──');
assert('comma separation',
  parseActorNames('יהודית אבדל, דניאל עובדיה, עידית עיד'),
  ['יהודית אבדל', 'דניאל עובדיה', 'עידית עיד']);
assert('ו-conjunction',
  parseActorNames('דודי גורדון ושאזו שאז'), ['דודי גורדון', 'שאזו שאז']);
assert('leading ו stripped',
  parseActorNames('ויעלה לוי'), ['יעלה לוי']);
assert('4-word even split → two 2-word names',
  parseActorNames('שאזו שאז אליאב מוסרי'), ['שאזו שאז', 'אליאב מוסרי']);
assert('ו-conjunction variant',
  parseActorNames('איתן החמוד וקובי ירחי'), ['איתן החמוד', 'קובי ירחי']);
assert('trailing spaces and * stripped',
  parseActorNames('קובי ירחי*, ישי מאיר '), ['קובי ירחי', 'ישי מאיר']);
assert('3-word name stays intact',
  parseActorNames('קרן ברל כצנלסון'), ['קרן ברל כצנלסון']);
assert('3-word + comma-separated 3-word',
  parseActorNames('קרן ברל כצנלסון, תגל פישר פרייס'),
  ['קרן ברל כצנלסון', 'תגל פישר פרייס']);
assert('empty string → []',  parseActorNames(''),        []);
assert('null → []',           parseActorNames(null),      []);
assert('undefined → []',      parseActorNames(undefined), []);

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
//  SESSION 2 — INTEGRATION (parser + validator)
// ─────────────────────────────────────────────

console.log('\n── Integration: raw February 2026 ──');
const rawWb     = XLSX.readFile('פברואר  2026.xlsx');
const rawResult = parseWorkbook(rawWb);
assert('raw data blocked (ambiguous יעלה)', rawResult.ok, false);
if (!rawResult.ok) {
  assert('exactly 1 validation error', rawResult.errors.length, 1);
  assert('error name is יעלה',         rawResult.errors[0].name, 'יעלה');
  console.log(`  ℹ  error: "${rawResult.errors[0].name}" @ ${rawResult.errors[0].date} ${rawResult.errors[0].client}`);
}

function buildPatchedWorkbook() {
  const wb   = XLSX.readFile('פברואר  2026.xlsx');
  const ws   = wb.Sheets['11.2 ביהס המסורתי'];
  const addr = XLSX.utils.encode_cell({ r: 2, c: 5 });
  if (ws[addr]) ws[addr].v = 'יעלה לוי, עדי זינגר';
  else ws[addr] = { v: 'יעלה לוי, עדי זינגר', t: 's' };
  return wb;
}

console.log('\n── Integration: patched February 2026 ──');
const patchedResult = parseWorkbook(buildPatchedWorkbook());
assert('patched data passes validation', patchedResult.ok, true);

if (patchedResult.ok) {
  const { workshops, actors } = patchedResult;
  assert('26 workshops', workshops.length, 26);
  assert('22 actors',    actors.length,    22);
  assert('sorted: first 1.2',  workshops[0].date,                  '1.2');
  assert('sorted: last 24.2',  workshops[workshops.length-1].date, '24.2');

  const byName = Object.fromEntries(actors.map(a => [a.name, a]));
  const counts = [
    ['איתן החמוד',6],['קובי ירחי',6],['שאזו שאז',4],['דודי גורדון',4],
    ['יהודית אבדל',3],['ישי כהן',3],['דניאל עובדיה',3],['עידית עיד',3],
    ['יעלה לוי',2],['אליאב מוסרי',2],['ציפי ציפורה',2],['תגל פישר פרייס',2],
    ['טל טל',1],['יובל יוב',1],['יאיר להמן',1],['עדי זינגר',1],
    ['עוזי בוס',1],['דורין עטר',1],['יעלה כהן',1],
    ['קרן ברל כצנלסון',1],['רן כהן',1],['תמר תמרוני',1],
  ];
  assert('all actor counts correct',
    counts.every(([n, c]) => byName[n]?.count === c), true);
}

// ─────────────────────────────────────────────
//  SESSION 3 — EXCEL OUTPUT
// ─────────────────────────────────────────────

console.log('\n── Session 3: Excel output ──');

const { workshops: febWorkshops, actors: febActors } =
  patchedResult.ok ? patchedResult : { workshops: [], actors: [] };

const outWb = buildWorkbook(febWorkshops, febActors);
const refWb = XLSX.readFile('actor_report_feb_2026.xlsx');

// ── Sheet names ──────────────────────────────
assert('output has 2 sheets',             outWb.SheetNames.length, 2);
assert('sheet 1 name: סיכום סדנאות',     outWb.SheetNames[0], 'סיכום סדנאות');
assert('sheet 2 name: סיכום שחקנים',    outWb.SheetNames[1], 'סיכום שחקנים');

// ── Sheet 1: סיכום סדנאות ───────────────────
const outWs1 = XLSX.utils.sheet_to_json(outWb.Sheets['סיכום סדנאות'],
  { header: 1, defval: '' });
const refWs1 = XLSX.utils.sheet_to_json(refWb.Sheets['סיכום סדנאות'],
  { header: 1, defval: '' });

assert('sheet 1 header',
  outWs1[0], ['תאריך', 'שם לקוח', 'שחקנים']);
assert('sheet 1 row count matches reference',
  outWs1.length, refWs1.length);

// Date column sorted and matches reference
assert('date column matches reference',
  outWs1.slice(1).map(r => r[0]),
  refWs1.slice(1).map(r => r[0]));

// Actor column: sorted sets match reference for all 26 rows
const actorColMismatches = outWs1.slice(1).filter((row, i) => {
  const refRow    = refWs1[i + 1] || [];
  const outActors = (row[2] || '').split(', ').map(s => s.trim()).filter(Boolean).sort();
  const refActors = (refRow[2] === '—' ? '' : (refRow[2] || ''))
    .split(', ').map(s => s.trim()).filter(Boolean).sort();
  return JSON.stringify(outActors) !== JSON.stringify(refActors);
});
assert('all 26 workshop actor lists match reference', actorColMismatches.length, 0);

// Workshops with no actors show '—'
const dashRows = outWs1.slice(1).filter(r => r[2] === '—');
assert('empty-actor workshops show —',
  febWorkshops.filter(w => w.actors.length === 0).length, dashRows.length);

// ── Sheet 2: סיכום שחקנים ───────────────────
const outWs2 = XLSX.utils.sheet_to_json(outWb.Sheets['סיכום שחקנים'],
  { header: 1, defval: '' });
const refWs2 = XLSX.utils.sheet_to_json(refWb.Sheets['סיכום שחקנים'],
  { header: 1, defval: '' });

assert('sheet 2 header',
  outWs2[0], ['שם שחקן/ית', 'מספר סדנאות', 'שמות סדנאות', 'קיבלנו הזמנת תשלום']);
assert('sheet 2 row count matches reference',
  outWs2.length, refWs2.length);

// Actor rows (exclude header and total)
const outActorRows = outWs2.slice(1, -1);
const refActorRows = refWs2.slice(1, -1);
const outActorMap  = Object.fromEntries(outActorRows.map(r => [r[0], r[1]]));
const refActorMap  = Object.fromEntries(refActorRows.map(r => [r[0], r[1]]));

assert('all 22 actor names present', refActorRows.every(r => outActorMap[r[0]] !== undefined), true);
assert('all actor counts match reference', refActorRows.every(r => outActorMap[r[0]] === r[1]), true);
assert('payment column empty in all rows', outActorRows.every(r => r[3] === ''), true);

// Total row
const outTotal = outWs2[outWs2.length - 1];
assert('total row label',            outTotal[0], 'סה"כ שחקנים');
assert('total row count = 22',       outTotal[1], 22);
assert('total row matches reference', outTotal[1], refWs2[refWs2.length - 1][1]);

// ── Structure ────────────────────────────────
assert('sheet 1 col widths: 3 cols', outWb.Sheets['סיכום סדנאות']['!cols'].length, 3);
assert('sheet 2 col widths: 4 cols', outWb.Sheets['סיכום שחקנים']['!cols'].length, 4);
assert('sheet 1 actors col width 52',
  outWb.Sheets['סיכום סדנאות']['!cols'][2].wch, 52);
assert('sheet 2 payment col width 22',
  outWb.Sheets['סיכום שחקנים']['!cols'][3].wch, 22);
assert('sheet 1 is RTL', outWb.Sheets['סיכום סדנאות']['!dir'], 'rtl');
assert('sheet 2 is RTL', outWb.Sheets['סיכום שחקנים']['!dir'], 'rtl');

// ── Filename inference ───────────────────────
const { month, year } = inferMonthYear(febWorkshops);
assert('month inferred as 02', month, '02');
assert('filename correct',
  `actor_report_${month}_${year}.xlsx`,
  `actor_report_02_${new Date().getFullYear()}.xlsx`);

// ─────────────────────────────────────────────
//  SUMMARY
// ─────────────────────────────────────────────

console.log(`\n${'─'.repeat(50)}`);
console.log(`Results: ${passed} passed, ${failed} failed`);
if (failed === 0) console.log('All tests passed! ✓');
else process.exitCode = 1;
