import { chromium } from '@playwright/test';
const browser = await chromium.launch();
const page = await browser.newPage({ viewport: { width: 1440, height: 860 } });
const errors = [];
page.on('pageerror', (e) => errors.push(String(e.stack || e)));
await page.goto('http://localhost:3100/');
await page.waitForFunction(() => window.__myexcel && window.__myexcel.pivot);
const out = await page.evaluate(() => {
  const X = window.__myexcel;
  const res = [];
  const check = (n, c, x) => res.push((c ? 'PASS ' : 'FAIL ') + n + (x !== undefined ? ' :: ' + JSON.stringify(x) : ''));
  const sh = () => X.S().wb.activeSheet;
  const put = (r, c, t) => X.store.transact('t', (tx) => X.edit.writeInput(tx, sh(), r, c, t));
  try {
    const rows = [['Region', 'Product', 'Month', 'Sales', 'Units'], ['East', 'Apples', '1/15/2024', '100', '10'], ['West', 'Apples', '2/10/2024', '150', '12'], ['East', 'Pears', '2/20/2024', '80', '8'], ['West', 'Pears', '3/5/2024', '120', '11'], ['East', 'Apples', '3/9/2024', '60', '5'], ['North', 'Plums', '3/19/2024', '90', '9']];
    rows.forEach((r, i) => r.forEach((v, j) => put(i, j, v)));
    const srcId = sh().id;
    X.pivot.createPivot(srcId, { r1: 0, c1: 0, r2: 6, c2: 4 }, { sheetId: null, r: 2, c: 0 });
    const ps = sh();
    const id = ps.pivots[0].id;
    X.pivot.updatePivot(id, { rows: ['Region'], cols: ['Product'], values: [{ field: 'Sales', agg: 'sum' }] });
    const cell = (r, c) => X.values.getScalar(sh(), r, c);
    check('pivot header', cell(2, 0) === 'Sum of Sales' && cell(3, 0) === 'Row Labels', [cell(2, 0), cell(3, 0), cell(3, 1)]);
    check('pivot east apples', cell(4, 0) === 'East' && cell(4, 1) === 160, [cell(4, 0), cell(4, 1)]);
    check('pivot grand total', cell(7, 0) === 'Grand Total' && cell(7, 4) === 600, [cell(7, 0), cell(7, 4)]);
    X.pivot.updatePivot(id, { rows: ['Month'], cols: [], values: [{ field: 'Sales', agg: 'sum' }, { field: 'Units', agg: 'average' }], grouping: [{ field: 'Month', type: 'date', by: 'months' }] });
    check('pivot date grouping', cell(3, 0) === 'Jan' && cell(5, 0) === 'Mar' && cell(5, 1) === 270, [cell(3, 0), cell(5, 0), cell(5, 1), cell(5, 2)]);
    X.pivot.updatePivot(id, { calculatedFields: [{ name: 'PerUnit', formula: '=Sales/Units' }], values: [{ field: 'Sales', agg: 'sum' }, { field: 'PerUnit', agg: 'sum' }] });
    check('pivot calc field', Math.abs(cell(3, 2) - 10) < 1e-9, cell(3, 2));
    X.pivot.updatePivot(id, { filters: [{ field: 'Region', selected: ['East'] }] });
    check('pivot filter', cell(2, 0) === 'Region' && cell(2, 1) === 'East', [cell(2, 0), cell(2, 1)]);
    // back to source for charts etc
    X.structure.activateSheet(srcId);
    X.store.setState({ sel: { ranges: [{ r1: 0, c1: 1, r2: 6, c2: 1 }, { r1: 0, c1: 3, r2: 6, c2: 3 }], active: { r: 0, c: 1 }, anchor: { r: 0, c: 1 } } });
    X.store.setState({ sel: { ranges: [{ r1: 0, c1: 3, r2: 6, c2: 4 }], active: { r: 0, c: 3 }, anchor: { r: 0, c: 3 } } });
    X.charts.insertChart('column');
    check('chart inserted', sh().charts.length === 1 && sh().charts[0].series.length === 2, sh().charts[0]?.series);
    const d = X.charts.chartData(sh().charts[0], sh());
    check('chart data', d.series[0].name === 'Sales' && d.series[0].values[0] === 100, d.series[0]);
    // goal seek: F2 = D2*2, find D2 so F2 = 500
    put(1, 5, '=D2*2');
    const gs = X.whatif.goalSeek('F2', 500, 'D2');
    check('goal seek', gs.found && Math.abs(gs.value - 250) < 1e-6, gs);
    // solver: maximize H1 = G1*(10-G1) → G1 = 5
    put(0, 6, '1'); put(0, 7, '=G1*(10-G1)');
    const sv = X.whatif.solve({ objective: 'H1', goal: 'max', targetValue: 0, variables: 'G1', constraints: [], nonNegative: true });
    check('solver', sv.ok && Math.abs(sv.values[0] - 5) < 1e-3, sv);
    // scenario
    X.whatif.addScenario('Best', 'G1', [7]);
    X.whatif.showScenario(sh().scenarios[0].id);
    check('scenario show', X.values.getScalar(sh(), 0, 6) === 7);
    // data table: column input G1, formula in J10 = H1
    put(9, 9, '=H1'); put(10, 8, '1'); put(11, 8, '2'); put(12, 8, '3');
    X.whatif.dataTable({ r1: 9, c1: 8, r2: 12, c2: 9 }, '', 'G1');
    check('data table', X.values.getScalar(sh(), 10, 9) === 9 && X.values.getScalar(sh(), 12, 9) === 21, [X.values.getScalar(sh(), 10, 9), X.values.getScalar(sh(), 12, 9)]);
    // print pages
    const pages = X.print.paginate(sh());
    check('paginate', pages.length >= 1, pages.length);
  } catch (e) { res.push('EXCEPTION ' + (e.stack || e)); }
  return res;
});
console.log(out.join('\n'));
await page.waitForTimeout(400);
await page.screenshot({ path: 'screenshots/phase3-chart.png' });
await page.evaluate(() => { const X = window.__myexcel; X.structure.activateSheet(X.S().wb.sheets[0].id); X.store.setState({ sel: { ranges: [{ r1: 4, c1: 0, r2: 4, c2: 0 }], active: { r: 4, c: 0 }, anchor: { r: 4, c: 0 } } }); });
await page.waitForTimeout(400);
await page.screenshot({ path: 'screenshots/phase3-pivot.png' });
if (errors.length) console.log('ERRORS\n' + errors.join('\n'));
await browser.close();
