// Converts typed cell input into a value + an auto-applied number format, like Excel.
import { partsToSerial } from './numfmt';

export interface ParsedInput {
  kind: 'empty' | 'formula' | 'number' | 'string' | 'boolean' | 'error';
  value: number | string | boolean | null;
  formula?: string;
  /** Number format Excel would auto-apply (only when the cell is currently General). */
  autoFormat?: string;
}

const MONTHS = ['jan', 'feb', 'mar', 'apr', 'may', 'jun', 'jul', 'aug', 'sep', 'oct', 'nov', 'dec'];
const ERRORS = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A', '#GETTING_DATA', '#SPILL!', '#CALC!'];

function monthIndex(s: string): number {
  const i = MONTHS.indexOf(s.slice(0, 3).toLowerCase());
  return i;
}

function validDate(y: number, m: number, d: number): boolean {
  if (m < 1 || m > 12 || d < 1 || y < 1900 || y > 9999) return false;
  const dim = new Date(Date.UTC(y, m, 0)).getUTCDate();
  return d <= dim;
}

function fixYear(y: number, raw: string): number {
  if (raw.length <= 2) return y < 30 ? 2000 + y : 1900 + y;
  return y;
}

function parseTime(s: string): { frac: number; hasAmPm: boolean; hasSeconds: boolean } | null {
  const m = /^(\d{1,2}):(\d{1,2})(?::(\d{1,2}(?:\.\d+)?))?\s*(am|pm|a|p)?$/i.exec(s.trim());
  if (!m) return null;
  let h = parseInt(m[1], 10);
  const mi = parseInt(m[2], 10);
  const sec = m[3] ? parseFloat(m[3]) : 0;
  if (mi > 59 || sec >= 60) return null;
  const ap = m[4]?.toLowerCase();
  if (ap) {
    if (h > 12 || h === 0) return null;
    if (ap.startsWith('p') && h < 12) h += 12;
    if (ap.startsWith('a') && h === 12) h = 0;
  }
  if (h > 9999) return null;
  return { frac: (h * 3600 + mi * 60 + sec) / 86400, hasAmPm: !!ap, hasSeconds: !!m[3] };
}

/** Returns serial + format for date strings Excel (en-US) recognises. */
export function parseDateString(s: string): { serial: number; fmt: string } | null {
  const t = s.trim();
  let m: RegExpExecArray | null;
  // m/d/yyyy or m-d-yyyy (US)
  if ((m = /^(\d{1,2})[/-](\d{1,2})[/-](\d{2,4})(?:\s+(.+))?$/.exec(t))) {
    const mo = +m[1];
    const d = +m[2];
    const y = fixYear(+m[3], m[3]);
    if (!validDate(y, mo, d)) return null;
    let serial = partsToSerial(y, mo, d);
    if (m[4]) {
      const tm = parseTime(m[4]);
      if (!tm) return null;
      serial += tm.frac;
      return { serial, fmt: 'm/d/yyyy h:mm' };
    }
    return { serial, fmt: 'm/d/yyyy' };
  }
  // yyyy-mm-dd / yyyy/mm/dd
  if ((m = /^(\d{4})[/-](\d{1,2})[/-](\d{1,2})(?:[\sT]+(.+))?$/.exec(t))) {
    const y = +m[1];
    const mo = +m[2];
    const d = +m[3];
    if (!validDate(y, mo, d)) return null;
    let serial = partsToSerial(y, mo, d);
    if (m[4]) {
      const tm = parseTime(m[4].replace(/Z$|[+-]\d\d:?\d\d$/, ''));
      if (!tm) return null;
      serial += tm.frac;
      return { serial, fmt: 'm/d/yyyy h:mm' };
    }
    return { serial, fmt: 'm/d/yyyy' };
  }
  // m/d (current year)
  if ((m = /^(\d{1,2})[/-](\d{1,2})$/.exec(t))) {
    const y = new Date().getFullYear();
    const mo = +m[1];
    const d = +m[2];
    if (!validDate(y, mo, d)) return null;
    return { serial: partsToSerial(y, mo, d), fmt: 'd-mmm' };
  }
  // d-mmm-yy, d mmm yyyy, d-mmm
  if ((m = /^(\d{1,2})[\s-]([A-Za-z]{3,9})(?:[\s-,]+(\d{2,4}))?$/.exec(t))) {
    const mi = monthIndex(m[2]);
    if (mi < 0) return null;
    const y = m[3] ? fixYear(+m[3], m[3]) : new Date().getFullYear();
    if (!validDate(y, mi + 1, +m[1])) return null;
    return { serial: partsToSerial(y, mi + 1, +m[1]), fmt: m[3] ? 'd-mmm-yy' : 'd-mmm' };
  }
  // mmm d, yyyy / mmmm d yyyy
  if ((m = /^([A-Za-z]{3,9})\.?\s+(\d{1,2}),?\s+(\d{2,4})$/.exec(t))) {
    const mi = monthIndex(m[1]);
    if (mi < 0) return null;
    const y = fixYear(+m[3], m[3]);
    if (!validDate(y, mi + 1, +m[2])) return null;
    return { serial: partsToSerial(y, mi + 1, +m[2]), fmt: 'd-mmm-yy' };
  }
  // mmm-yy / mmm yyyy
  if ((m = /^([A-Za-z]{3,9})[\s-](\d{2,4})$/.exec(t))) {
    const mi = monthIndex(m[1]);
    if (mi < 0) return null;
    const y = fixYear(+m[2], m[2]);
    if (!validDate(y, mi + 1, 1)) return null;
    return { serial: partsToSerial(y, mi + 1, 1), fmt: 'mmm-yy' };
  }
  return null;
}

export function parseNumberString(raw: string): { value: number; fmt?: string } | null {
  let s = raw.trim();
  if (!s) return null;
  let neg = false;
  if (/^\(.*\)$/.test(s)) {
    neg = true;
    s = s.slice(1, -1).trim();
  }
  if (s.startsWith('-')) {
    neg = !neg;
    s = s.slice(1).trim();
  } else if (s.startsWith('+')) s = s.slice(1).trim();
  let currency: string | undefined;
  const cm = /^([$€£¥₹])\s*/.exec(s);
  if (cm) {
    currency = cm[1];
    s = s.slice(cm[0].length);
  } else {
    const tail = /\s*([$€£¥₹])$/.exec(s);
    if (tail) {
      currency = tail[1];
      s = s.slice(0, -tail[0].length);
    }
  }
  if (s.startsWith('-') && !neg) {
    neg = true;
    s = s.slice(1);
  }
  let pct = false;
  if (s.endsWith('%')) {
    pct = true;
    s = s.slice(0, -1).trim();
  }
  let thousands = false;
  if (/^\d{1,3}(,\d{2,3})+(\.\d*)?$/.test(s)) {
    // accept both western (1,234,567) and Indian (12,34,567) grouping
    const intPart = s.split('.')[0];
    const groups = intPart.split(',');
    const last = groups[groups.length - 1];
    if (last.length !== 3) return null;
    thousands = true;
    s = s.replace(/,/g, '');
  }
  if (!/^(\d+\.?\d*|\.\d+)([eE][+-]?\d+)?$/.test(s)) return null;
  let v = parseFloat(s);
  if (!isFinite(v)) return null;
  if (neg) v = -v;
  const decimals = (/\.(\d+)/.exec(s)?.[1].length ?? 0);
  if (pct) return { value: v / 100, fmt: decimals ? '0.' + '0'.repeat(Math.min(decimals, 10)) + '%' : '0%' };
  if (currency) {
    const d = decimals ? '.00' : '';
    if (currency === '₹') return { value: v, fmt: `[$₹-4009] #,##0${d}` };
    const sym = currency === '$' ? '"$"' : `[$${currency}]`;
    return { value: v, fmt: `${sym}#,##0${d}` };
  }
  if (/[eE]/.test(s)) return { value: v, fmt: '0.00E+00' };
  if (thousands) return { value: v, fmt: decimals ? '#,##0.00' : '#,##0' };
  return { value: v };
}

export function parseInput(input: string): ParsedInput {
  if (input === '') return { kind: 'empty', value: null };
  if (input.startsWith("'")) return { kind: 'string', value: input.slice(1) };
  if ((input.startsWith('=') || input.startsWith('+') || (input.startsWith('-') && /[A-Za-z(]/.test(input))) && input.length > 1) {
    // "+A1" and "-A1" become formulas like in Excel
    if (input.startsWith('=')) return { kind: 'formula', value: null, formula: input };
    if (/^[+-]\s*\d+(\.\d+)?%?$/.test(input) === false && /[A-Za-z(]/.test(input)) {
      return { kind: 'formula', value: null, formula: '=' + input };
    }
  }
  const up = input.trim().toUpperCase();
  if (up === 'TRUE') return { kind: 'boolean', value: true };
  if (up === 'FALSE') return { kind: 'boolean', value: false };
  if (ERRORS.includes(up)) return { kind: 'error', value: up };
  const num = parseNumberString(input);
  if (num) return { kind: 'number', value: num.value, autoFormat: num.fmt };
  const tm = parseTime(input);
  if (tm && /^\s*\d{1,2}:\d/.test(input)) {
    const fmt = tm.hasAmPm ? (tm.hasSeconds ? 'h:mm:ss AM/PM' : 'h:mm AM/PM') : tm.hasSeconds ? 'h:mm:ss' : 'h:mm';
    return { kind: 'number', value: tm.frac, autoFormat: fmt };
  }
  const dt = parseDateString(input);
  if (dt) return { kind: 'number', value: dt.serial, autoFormat: dt.fmt };
  return { kind: 'string', value: input };
}
