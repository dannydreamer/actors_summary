/**
 * SimCenter Actor Tool — Excel Report Generator
 * Works in both browser (global export) and Node.js (module.exports).
 * Depends on SheetJS (XLSX) being available in scope.
 */

(function (root, factory) {
  if (typeof module !== 'undefined' && module.exports) {
    module.exports = factory(require('./node_modules/xlsx'));
  } else {
    root.SimCenterGenerator = factory(root.XLSX);
  }
}(typeof globalThis !== 'undefined' ? globalThis : this, function (XLSX) {
  'use strict';

  // ─── Helpers ──────────────────────────────────────────────────────────────

  function cellAt(c, r) {
    return XLSX.utils.encode_cell({ c, r });
  }

  function ensureCell(ws, addr) {
    if (!ws[addr]) ws[addr] = { v: '', t: 's' };
    if (!ws[addr].s) ws[addr].s = {};
  }

  function applyAlternatingRows(ws, rowCount, colCount) {
    for (let r = 1; r < rowCount; r += 2) {
      for (let c = 0; c < colCount; c++) {
        const addr = cellAt(c, r);
        ensureCell(ws, addr);
        ws[addr].s.fill = { fgColor: { rgb: 'F3F4F6' }, patternType: 'solid' };
      }
    }
  }

  function applyPaymentColumnFill(ws, rowCount) {
    for (let r = 1; r < rowCount; r++) {
      const addr = cellAt(3, r);
      ensureCell(ws, addr);
      ws[addr].s = { fill: { fgColor: { rgb: 'D1FAE5' }, patternType: 'solid' } };
    }
  }

  // ─── Main builder ─────────────────────────────────────────────────────────

  /**
   * Build a SheetJS workbook from parsed output.
   * @param {Array} workshops  — [{ date, client, actors }]
   * @param {Array} actors     — [{ name, count, clients }]
   * @returns XLSX workbook object
   */
  function buildWorkbook(workshops, actors) {
    const wb = XLSX.utils.book_new();

    // ── Sheet 1: סיכום סדנאות ──────────────────────────────────────────────
    const ws1Data = [['תאריך', 'שם לקוח', 'שחקנים']];
    workshops.forEach(w => ws1Data.push([
      w.date,
      w.client,
      w.actors.length ? w.actors.join(', ') : '—',
    ]));

    const ws1 = XLSX.utils.aoa_to_sheet(ws1Data);
    ws1['!cols'] = [{ wch: 8 }, { wch: 36 }, { wch: 52 }];
    ws1['!dir']  = 'rtl';
    applyAlternatingRows(ws1, ws1Data.length, 3);
    XLSX.utils.book_append_sheet(wb, ws1, 'סיכום סדנאות');

    // ── Sheet 2: סיכום שחקנים ──────────────────────────────────────────────
    const ws2Data = [['שם שחקן/ית', 'מספר סדנאות', 'שמות סדנאות', 'קיבלנו הזמנת תשלום']];
    actors.forEach(a => ws2Data.push([a.name, a.count, a.clients.join(', '), '']));
    ws2Data.push(['סה"כ שחקנים', actors.length, '', '']);

    const ws2 = XLSX.utils.aoa_to_sheet(ws2Data);
    ws2['!cols'] = [{ wch: 22 }, { wch: 14 }, { wch: 52 }, { wch: 22 }];
    ws2['!dir']  = 'rtl';
    applyAlternatingRows(ws2, ws2Data.length, 4);
    applyPaymentColumnFill(ws2, ws2Data.length);
    XLSX.utils.book_append_sheet(wb, ws2, 'סיכום שחקנים');

    return wb;
  }

  /**
   * Infer month/year from the first workshop date (d.m format).
   */
  function inferMonthYear(workshops) {
    for (const w of workshops) {
      const parts = w.date.split('.');
      if (parts.length === 2) {
        return {
          month: parts[1].padStart(2, '0'),
          year:  new Date().getFullYear(),
        };
      }
    }
    return { month: '00', year: new Date().getFullYear() };
  }

  /**
   * Generate and trigger download of the report xlsx.
   * Browser-only — uses XLSX.writeFile which invokes a save dialog.
   */
  function downloadReport(workshops, actors) {
    const wb = buildWorkbook(workshops, actors);
    const { month, year } = inferMonthYear(workshops);
    XLSX.writeFile(wb, `actor_report_${month}_${year}.xlsx`);
  }

  // ─── Public API ───────────────────────────────────────────────────────────

  return { buildWorkbook, inferMonthYear, downloadReport };
}));
