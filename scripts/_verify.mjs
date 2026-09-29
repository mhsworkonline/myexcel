import { chromium } from '@playwright/test';
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 860 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e.stack || e)));
await page.goto('http://localhost:3100/');
await page.waitForFunction(() => window.__myexcel && window.__myexcel.edit);
const results = await page.evaluate(() => {
  const X = window.__myexcel;
  const out = [];
  const check = (name, cond, extra) => out.push((cond ? 'PASS ' : 'FAIL ') + name + (extra !== undefined ? ' :: ' + JSON.stringify(extra) : ''));
  const st = () => X.S();
  const sh = () => st().wb.activeSheet;
  const val = (r, c) => X.values.getScalar(sh(), r, c);
  const sel = (r1, c1, r2 = r1, c2 = c1) => X.store.setState({ sel: { ranges: [{ r1, c1, r2, c2 }], active: { r: r1, c: c1 }, anchor: { r: r1, c: c1 } } });
  const put = (r, c, t) => X.store.transact('t', (tx) => X.edit.writeInput(tx, sh(), r, c, t));
  try {
    // data
    ['A', 'B', 'C'].forEach((h, i) => put(0, i, h));
    for (let r = 1; r <= 5; r++) { put(r, 0, 'n' + (6 - r)); put(r, 1, String(r * 10)); put(r, 2, '=B' + (r + 1) + '*2'); }
    check('formula', val(1, 2) === 20, val(1, 2));
    // insert row above row 3 → formulas shift
    sel(2, 0); X.structure.insertRows();
    check('insert row shifts formula', sh().getCell(3, 2)?.f === '=B4*2', sh().getCell(3, 2)?.f);
    X.store.undo();
    check('undo insert', sh().getCell(2, 2)?.f === '=B3*2', sh().getCell(2, 2)?.f);
    // delete column B → #REF!
    sel(0, 1, 1048575, 1); X.structure.deleteCols();
    check('delete col -> #REF!', String(sh().getCell(1, 1)?.f).includes('#REF!'), sh().getCell(1, 1)?.f);
    X.store.undo();
    check('undo delete col', val(1, 2) === 20, val(1, 2));
    // sort by A ascending with header
    X.sortFilter.sortRange({ r1: 0, c1: 0, r2: 5, c2: 2 }, [{ col: 0, desc: false, by: 'value' }], true);
    check('sort', val(1, 0) === 'n1' && val(1, 1) === 50 && val(1, 2) === 100, [val(1, 0), val(1, 1), val(1, 2)]);
    X.store.undo();
    // autofilter
    sel(0, 0); X.sortFilter.toggleAutoFilter();
    X.sortFilter.setColumnFilter(1, { type: 'custom', and: true, c1: { op: 'gt', val: '25' } });
    check('filter hides rows', sh().filteredRows.has(1) && sh().filteredRows.has(2) && !sh().filteredRows.has(3), [...sh().filteredRows]);
    X.sortFilter.clearAllFilters();
    check('clear filter', sh().filteredRows.size === 0);
    X.sortFilter.toggleAutoFilter();
    // merge
    sel(7, 0, 8, 2); put(7, 0, 'Merged'); X.format.mergeCells('center', true);
    check('merge', sh().merges.length === 1 && st().wb.styles.get(sh().styleIdAt(7, 0)).hAlign === 'center');
    // fill series
    put(10, 0, 'Jan'); X.fill.autoFill({ r1: 10, c1: 0, r2: 10, c2: 0 }, { r1: 10, c1: 0, r2: 13, c2: 0 });
    check('fill months', val(13, 0) === 'Apr', val(13, 0));
    put(10, 1, '1'); put(11, 1, '3'); X.fill.autoFill({ r1: 10, c1: 1, r2: 11, c2: 1 }, { r1: 10, c1: 1, r2: 14, c2: 1 });
    check('fill linear', val(14, 1) === 9, val(14, 1));
    put(10, 2, '1/31/2024'); X.fill.autoFill({ r1: 10, c1: 2, r2: 10, c2: 2 }, { r1: 10, c1: 2, r2: 11, c2: 2 });
    check('fill date', val(11, 2) === val(10, 2) + 1);
    // find & replace
    const hits = X.find.findAll({ ...X.find.defaultFind, what: 'n*' });
    check('find wildcard', hits.length >= 5, hits.length);
    X.find.replaceAll({ ...X.find.defaultFind, what: 'n1', entireCell: true }, 'one');
    check('replace', X.find.findAll({ ...X.find.defaultFind, what: 'one' }).length === 1);
    X.store.closeDialog();
    // number formats
    sel(1, 1); X.format.setNumFmt('[$₹-4009] #,##0.00'); put(1, 1, '12345678');
    const d = X.values.displayOf(sh(), 1, 1, st().wb.styles.get(sh().styleIdAt(1, 1)));
    check('INR display', d.text === '₹ 1,23,45,678.00', d.text);
    // input parsing
    put(20, 0, '10%'); check('percent input', val(20, 0) === 0.1 && st().wb.styles.get(sh().styleIdAt(20, 0)).numFmt === '0%');
    put(20, 1, "'007"); check('apostrophe text', val(20, 1) === '007');
    // names
    X.formulas.defineName('Rate', "Sheet1!$B$2");
    put(21, 0, '=Rate*2'); check('named range', val(21, 0) === 24691356, val(21, 0));
    // conditional formatting + validation
    sel(1, 1, 5, 1); X.data.addCFRule({ type: 'cellIs', operator: 'greaterThan', formulas: ['=25'], style: { fillColor: '#FFC7CE' } });
    check('cf rule added', sh().conditionalFormats.length === 1);
    X.data.setValidation({ type: 'list', formula1: 'Yes,No', allowBlank: true, showDropdown: true, showInput: false, showError: true, errorStyle: 'stop' }, [{ r1: 30, c1: 0, r2: 35, c2: 0 }]);
    check('dv', sh().validations.length === 1);
    // table
    X.data.createTable({ r1: 40, c1: 0, r2: 42, c2: 1 }, false);
    check('table', sh().tables.length === 1 && val(40, 0) === 'Column1', val(40, 0));
    // sheets
    X.structure.addSheet(); check('add sheet', st().wb.sheets.length === 2);
    X.structure.renameSheet(st().wb.activeSheetId, 'Data 2'); check('rename', sh().name === 'Data 2');
    X.store.undo(); X.store.undo();
    check('undo sheet ops', st().wb.sheets.length === 1, st().wb.sheets.map((s) => s.name));
    // hide / unhide
    sel(3, 0, 4, 16383); X.format.hideRowsCols('row', true);
    check('hide rows', sh().hiddenRows.has(3) && sh().hiddenRows.has(4));
    X.format.hideRowsCols('row', false); check('unhide rows', !sh().hiddenRows.has(3));
    // clipboard copy/paste internal with formula shift
    sel(1, 2); X.store.setState({});
    // circular
    put(50, 0, '=A51+1'); check('circular', !!st().engine.findCircular(sh()));
  } catch (e) { out.push('EXCEPTION ' + (e.stack || e)); }
  return out;
});
console.log(results.join('\n'));
if (errors.length) console.log('PAGE ERRORS\n' + errors.join('\n'));
await browser.close();
