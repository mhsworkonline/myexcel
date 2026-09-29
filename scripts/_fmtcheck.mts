import { formatValue as f, FMT, formatGeneral, adjustDecimals } from '../src/model/numfmt.ts';
const cases: [any, string][] = [
  [1234.567, '0.00'], [1234.567, '#,##0.00'], [-1234.5, '#,##0.00_);[Red](#,##0.00)'], [0.1234, '0%'], [0.1234, '0.00%'],
  [12345678.9, FMT.inr], [123456789, '[$₹-4009] #,##0'], [1234.5, FMT.accounting], [0, FMT.accounting], [-5, FMT.accounting],
  [45293, 'm/d/yyyy'], [45293.5, 'dddd, mmmm d, yyyy h:mm AM/PM'], [0.75, 'h:mm:ss'], [1.5, '[h]:mm'], [12345, '0.00E+00'],
  [0.5, '# ?/?'], [1.75, '# ?/?'], [5551234567, '(000) 000-0000'], [123, '000-00-0000'], ['abc', '@'], [5.5, '.00'],
  [1/3, 'General'], [123456789012, 'General'], [1234567, '#,##0,"K"'], [-3, '0;(0)'], [45293, 'mmm-yy'], [0.000012, 'General'],
  [123.456, '"$"#,##0.00'], [150, '[>100][Red]0;0'],
];
for (const [v, fm] of cases) { const r = f(v, fm); console.log(JSON.stringify(v).padEnd(14), fm.padEnd(40), '→', JSON.stringify(r.text), r.left !== undefined ? `L=${JSON.stringify(r.left)} R=${JSON.stringify(r.right)}` : '', r.color ?? ''); }
console.log(adjustDecimals('0.00', 1), adjustDecimals('0.00', -1), adjustDecimals(undefined, 1, 3.5), adjustDecimals('#,##0', 1));
