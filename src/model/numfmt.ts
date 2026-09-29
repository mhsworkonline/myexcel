// Excel number-format code interpreter.

export type ScalarValue = number | string | boolean | null | undefined;

export interface FormatResult {
  text: string;
  color?: string;
  /** When the format contains a `*x` fill, text is split into a left and right part. */
  left?: string;
  right?: string;
  /** True when a numeric format was applied (numbers right-align by default). */
  isNumber?: boolean;
}

const COLOR_NAMES: Record<string, string> = {
  black: '#000000',
  blue: '#0000FF',
  cyan: '#00FFFF',
  green: '#00FF00',
  magenta: '#FF00FF',
  red: '#FF0000',
  white: '#FFFFFF',
  yellow: '#FFFF00',
};

const INDEXED_COLORS = [
  '#000000', '#FFFFFF', '#FF0000', '#00FF00', '#0000FF', '#FFFF00', '#FF00FF', '#00FFFF',
  '#800000', '#008000', '#000080', '#808000', '#800080', '#008080', '#C0C0C0', '#808080',
  '#9999FF', '#993366', '#FFFFCC', '#CCFFFF', '#660066', '#FF8080', '#0066CC', '#CCCCFF',
  '#000080', '#FF00FF', '#FFFF00', '#00FFFF', '#800080', '#800000', '#008080', '#0000FF',
  '#00CCFF', '#CCFFFF', '#CCFFCC', '#FFFF99', '#99CCFF', '#FF99CC', '#CC99FF', '#FFCC99',
  '#3366FF', '#33CCCC', '#99CC00', '#FFCC00', '#FF9900', '#FF6600', '#666699', '#969696',
  '#003366', '#339966', '#003300', '#333300', '#993300', '#993366', '#333399', '#333333',
];

const MONTHS = ['January', 'February', 'March', 'April', 'May', 'June', 'July', 'August', 'September', 'October', 'November', 'December'];
const DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];

type Tok =
  | { t: 'lit'; v: string }
  | { t: 'digit'; v: '0' | '#' | '?' }
  | { t: 'dot' }
  | { t: 'comma' }
  | { t: 'pct' }
  | { t: 'exp'; sign: '+' | '-' }
  | { t: 'slash' }
  | { t: 'text' }
  | { t: 'fill'; v: string }
  | { t: 'skip'; v: string }
  | { t: 'date'; v: string }
  | { t: 'ampm'; v: string }
  | { t: 'elapsed'; v: string };

interface Section {
  toks: Tok[];
  color?: string;
  cond?: { op: string; val: number };
  isDate: boolean;
  isText: boolean;
  hasNumber: boolean;
  indian: boolean;
  isGeneral: boolean;
}

interface ParsedFormat {
  sections: Section[];
}

const cache = new Map<string, ParsedFormat>();

function splitSections(fmt: string): string[] {
  const out: string[] = [];
  let cur = '';
  let q = false;
  let br = false;
  for (let i = 0; i < fmt.length; i++) {
    const ch = fmt[i];
    if (ch === '\\' && !q) {
      cur += ch + (fmt[i + 1] ?? '');
      i++;
      continue;
    }
    if (ch === '"') q = !q;
    else if (ch === '[' && !q) br = true;
    else if (ch === ']' && !q) br = false;
    if (ch === ';' && !q && !br) {
      out.push(cur);
      cur = '';
    } else cur += ch;
  }
  out.push(cur);
  return out;
}

function parseSection(src: string): Section {
  const toks: Tok[] = [];
  const sec: Section = { toks, isDate: false, isText: false, hasNumber: false, indian: false, isGeneral: false };
  let i = 0;
  const lower = src.toLowerCase();
  while (i < src.length) {
    const ch = src[i];
    const lc = lower[i];
    if (ch === '"') {
      const j = src.indexOf('"', i + 1);
      const end = j < 0 ? src.length : j;
      toks.push({ t: 'lit', v: src.slice(i + 1, end) });
      i = end + 1;
      continue;
    }
    if (ch === '\\') {
      toks.push({ t: 'lit', v: src[i + 1] ?? '' });
      i += 2;
      continue;
    }
    if (ch === '_') {
      toks.push({ t: 'skip', v: src[i + 1] ?? ' ' });
      i += 2;
      continue;
    }
    if (ch === '*') {
      toks.push({ t: 'fill', v: src[i + 1] ?? ' ' });
      i += 2;
      continue;
    }
    if (ch === '[') {
      const j = src.indexOf(']', i);
      const inner = src.slice(i + 1, j < 0 ? src.length : j);
      i = j < 0 ? src.length : j + 1;
      const il = inner.toLowerCase();
      if (COLOR_NAMES[il]) sec.color = COLOR_NAMES[il];
      else if (/^color\s*\d+$/.test(il)) {
        const n = parseInt(il.replace(/\D/g, ''), 10);
        sec.color = INDEXED_COLORS[(n - 1) % INDEXED_COLORS.length];
      } else if (/^(<=|>=|<>|<|>|=)/.test(inner)) {
        const m = /^(<=|>=|<>|<|>|=)\s*(-?[\d.]+)/.exec(inner);
        if (m) sec.cond = { op: m[1], val: parseFloat(m[2]) };
      } else if (inner.startsWith('$')) {
        const dash = inner.indexOf('-');
        const sym = dash >= 0 ? inner.slice(1, dash) : inner.slice(1);
        const loc = dash >= 0 ? inner.slice(dash + 1).toLowerCase() : '';
        if (sym) toks.push({ t: 'lit', v: sym });
        if (loc === '4009' || loc === '439' || loc === '449' || loc.endsWith('-in') || loc === 'hi-in') sec.indian = true;
      } else if (/^(h+|m+|s+)$/i.test(inner)) {
        toks.push({ t: 'elapsed', v: il });
        sec.isDate = true;
      }
      continue;
    }
    if (lower.startsWith('general', i)) {
      sec.isGeneral = true;
      toks.push({ t: 'lit', v: '\u0000GENERAL' });
      i += 7;
      continue;
    }
    if (lower.startsWith('am/pm', i)) {
      toks.push({ t: 'ampm', v: 'AM/PM' });
      sec.isDate = true;
      i += 5;
      continue;
    }
    if (lower.startsWith('a/p', i)) {
      toks.push({ t: 'ampm', v: 'A/P' });
      sec.isDate = true;
      i += 3;
      continue;
    }
    if ('ymdhs'.includes(lc) || (lc === 'e' && sec.isDate) || (lc === 'b' && /^b[12]/.test(lower.slice(i)))) {
      if (lc === 'e' && !sec.isDate) {
        // handled below as exponent
      } else {
        let j = i;
        while (j < src.length && lower[j] === lc) j++;
        toks.push({ t: 'date', v: lower.slice(i, j) });
        sec.isDate = true;
        i = j;
        continue;
      }
    }
    if (ch === '0' || ch === '#' || ch === '?') {
      // Fractional seconds after a date token e.g. ss.00
      toks.push({ t: 'digit', v: ch });
      sec.hasNumber = true;
      i++;
      continue;
    }
    if (ch === '.') {
      toks.push({ t: 'dot' });
      i++;
      continue;
    }
    if (ch === ',') {
      toks.push({ t: 'comma' });
      i++;
      continue;
    }
    if (ch === '%') {
      toks.push({ t: 'pct' });
      i++;
      continue;
    }
    if ((ch === 'E' || ch === 'e') && (src[i + 1] === '+' || src[i + 1] === '-')) {
      toks.push({ t: 'exp', sign: src[i + 1] as '+' | '-' });
      i += 2;
      continue;
    }
    if (ch === '/') {
      toks.push({ t: 'slash' });
      i++;
      continue;
    }
    if (ch === '@') {
      toks.push({ t: 'text' });
      sec.isText = true;
      i++;
      continue;
    }
    toks.push({ t: 'lit', v: ch });
    i++;
  }
  return sec;
}

function parseFormat(fmt: string): ParsedFormat {
  let p = cache.get(fmt);
  if (p) return p;
  const secs = splitSections(fmt).map(parseSection);
  p = { sections: secs };
  if (cache.size > 2000) cache.clear();
  cache.set(fmt, p);
  return p;
}

function testCond(c: { op: string; val: number }, v: number): boolean {
  switch (c.op) {
    case '<': return v < c.val;
    case '<=': return v <= c.val;
    case '>': return v > c.val;
    case '>=': return v >= c.val;
    case '=': return v === c.val;
    case '<>': return v !== c.val;
  }
  return false;
}

// ---------- General ----------

export function formatGeneral(n: number, maxChars = 11): string {
  if (!isFinite(n)) return '#NUM!';
  if (n === 0) return '0';
  const abs = Math.abs(n);
  if (abs >= 1e11 || abs < 1e-9) {
    for (let d = 5; d >= 0; d--) {
      const s = n.toExponential(d).replace(/\.?0+e/, 'e').replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2');
      if (s.length <= maxChars || d === 0) return s;
    }
  }
  const intDigits = Math.floor(Math.log10(abs)) + 1;
  const sign = n < 0 ? 1 : 0;
  let decimals = Math.max(0, maxChars - sign - Math.max(intDigits, 1) - 1);
  decimals = Math.min(decimals, 15 - Math.max(intDigits, 0));
  if (decimals < 0) decimals = 0;
  let s = n.toFixed(Math.min(decimals, 20));
  if (s.includes('.')) s = s.replace(/0+$/, '').replace(/\.$/, '');
  if (s === '-0') s = '0';
  if (s.length > maxChars && Math.abs(n) >= 1) {
    const e = n.toExponential(Math.max(0, maxChars - 6)).replace(/\.?0+e/, 'e').replace('e', 'E').replace(/E([+-])(\d)$/, 'E$10$2');
    return e;
  }
  return s;
}

// ---------- Dates ----------

export interface DateParts {
  y: number;
  m: number; // 1-12
  d: number;
  H: number;
  M: number;
  S: number;
  ms: number;
  dow: number;
}

export function serialToParts(serial: number): DateParts {
  let days = Math.floor(serial);
  let frac = serial - days;
  // round to milliseconds to avoid 59.9999 seconds
  let totalMs = Math.round(frac * 86400000);
  if (totalMs >= 86400000) {
    days += 1;
    totalMs -= 86400000;
  }
  const base = Date.UTC(1899, 11, 30);
  const dt = new Date(base + days * 86400000);
  const H = Math.floor(totalMs / 3600000);
  const M = Math.floor((totalMs % 3600000) / 60000);
  const S = Math.floor((totalMs % 60000) / 1000);
  const ms = totalMs % 1000;
  return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate(), H, M, S, ms, dow: dt.getUTCDay() };
}

export function partsToSerial(y: number, m: number, d: number, H = 0, M = 0, S = 0): number {
  const base = Date.UTC(1899, 11, 30);
  const t = Date.UTC(y, m - 1, d);
  return Math.round((t - base) / 86400000) + (H * 3600 + M * 60 + S) / 86400;
}

function pad(n: number, w: number): string {
  let s = String(Math.abs(Math.trunc(n)));
  while (s.length < w) s = '0' + s;
  return (n < 0 ? '-' : '') + s;
}

function formatDateSection(sec: Section, value: number): string {
  const p = serialToParts(value);
  const hasAmPm = sec.toks.some((t) => t.t === 'ampm');
  const toks = sec.toks;
  // Decide which 'm' tokens are minutes: m following h or preceding s.
  const kinds: string[] = toks.map((t) => (t.t === 'date' ? t.v : ''));
  const isMinute = (idx: number): boolean => {
    for (let j = idx - 1; j >= 0; j--) {
      const k = toks[j];
      if (k.t === 'date') return k.v[0] === 'h';
      if (k.t === 'elapsed') return k.v[0] === 'h';
      if (k.t === 'digit' || k.t === 'dot') continue;
    }
    for (let j = idx + 1; j < toks.length; j++) {
      const k = toks[j];
      if (k.t === 'date') return k.v[0] === 's';
      if (k.t === 'elapsed') return k.v[0] === 's';
    }
    return false;
  };
  void kinds;
  let out = '';
  let fracDigitsAfterSec = 0;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === 'date' && t.v[0] === 's') {
      // check for .000 afterwards
      let j = i + 1;
      if (toks[j]?.t === 'dot') {
        j++;
        while (toks[j]?.t === 'digit') {
          fracDigitsAfterSec++;
          j++;
        }
      }
    }
  }
  // Rounding for seconds when no fractional part displayed
  let { S, M, H } = p;
  let ms = p.ms;
  if (fracDigitsAfterSec === 0 && ms >= 500) {
    S += 1;
    ms = 0;
    if (S === 60) {
      S = 0;
      M += 1;
      if (M === 60) {
        M = 0;
        H += 1;
      }
    }
  }
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.t) {
      case 'lit':
        out += t.v;
        break;
      case 'skip':
        out += ' ';
        break;
      case 'fill':
        break;
      case 'ampm': {
        const pm = H >= 12;
        out += t.v === 'AM/PM' ? (pm ? 'PM' : 'AM') : pm ? 'P' : 'A';
        break;
      }
      case 'elapsed': {
        const totalSec = Math.round(value * 86400);
        if (t.v[0] === 'h') out += pad(Math.floor(totalSec / 3600), t.v.length);
        else if (t.v[0] === 'm') out += pad(Math.floor(totalSec / 60), t.v.length);
        else out += pad(totalSec, t.v.length);
        break;
      }
      case 'date': {
        const v = t.v;
        const c = v[0];
        if (c === 'y' || c === 'e') out += v.length <= 2 && c === 'y' ? pad(p.y % 100, 2) : String(p.y);
        else if (c === 'm') {
          if (v.length <= 2 && isMinute(i)) {
            out += pad(M, v.length);
            break;
          }
          if (v.length === 1) out += String(p.m);
          else if (v.length === 2) out += pad(p.m, 2);
          else if (v.length === 3) out += MONTHS[p.m - 1].slice(0, 3);
          else if (v.length === 5) out += MONTHS[p.m - 1][0];
          else out += MONTHS[p.m - 1];
        } else if (c === 'd') {
          if (v.length === 1) out += String(p.d);
          else if (v.length === 2) out += pad(p.d, 2);
          else if (v.length === 3) out += DAYS[p.dow].slice(0, 3);
          else out += DAYS[p.dow];
        } else if (c === 'h') {
          let h = H;
          if (hasAmPm) {
            h = h % 12;
            if (h === 0) h = 12;
          }
          out += v.length >= 2 ? pad(h, 2) : String(h);
        } else if (c === 's') {
          out += v.length >= 2 ? pad(S, 2) : String(S);
          if (toks[i + 1]?.t === 'dot' && toks[i + 2]?.t === 'digit') {
            let n = 0;
            let j = i + 2;
            while (toks[j]?.t === 'digit') {
              n++;
              j++;
            }
            const frac = Math.floor((ms / 1000) * Math.pow(10, n));
            out += '.' + pad(frac, n);
            i = j - 1;
          }
        } else if (c === 'b') {
          out += String(p.y + 543);
        }
        break;
      }
      case 'dot':
        out += '.';
        break;
      case 'comma':
        out += ',';
        break;
      case 'slash':
        out += '/';
        break;
      case 'pct':
        out += '%';
        break;
      case 'digit':
        out += t.v === '0' ? '0' : '';
        break;
      default:
        break;
    }
  }
  return out;
}

// ---------- Numbers ----------

function groupDigits(s: string, indian: boolean): string {
  if (s.length <= 3) return s;
  if (!indian) return s.replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  const last3 = s.slice(-3);
  const rest = s.slice(0, -3);
  return rest.replace(/\B(?=(\d{2})+(?!\d))/g, ',') + ',' + last3;
}

function gcd(a: number, b: number): number {
  while (b) [a, b] = [b, a % b];
  return a;
}

function approxFraction(x: number, maxDen: number): [number, number] {
  let bestN = 0;
  let bestD = 1;
  let bestErr = Infinity;
  for (let d = 1; d <= maxDen; d++) {
    const n = Math.round(x * d);
    const err = Math.abs(x - n / d);
    if (err < bestErr - 1e-12) {
      bestErr = err;
      bestN = n;
      bestD = d;
    }
  }
  const g = gcd(bestN, bestD) || 1;
  return [bestN / g, bestD / g];
}

function renderPlaceholders(phs: ('0' | '#' | '?')[], digits: string, leftFill: boolean): string {
  // leftFill=true → integer part (right-aligned digits); false → decimal part
  let out = '';
  if (leftFill) {
    let di = digits.length - 1;
    for (let pi = phs.length - 1; pi >= 0; pi--) {
      const ph = phs[pi];
      if (di >= 0) {
        if (pi === 0) {
          out = digits.slice(0, di + 1) + out;
          di = -1;
        } else {
          out = digits[di] + out;
          di--;
        }
      } else if (ph === '0') out = '0' + out;
      else if (ph === '?') out = ' ' + out;
    }
    if (phs.length === 0 && digits.length) out = digits;
    return out;
  }
  for (let pi = 0; pi < phs.length; pi++) {
    const ch = digits[pi];
    out += ch ?? (phs[pi] === '0' ? '0' : phs[pi] === '?' ? ' ' : '');
  }
  return out;
}

/** Distribute integer digits over placeholders right-to-left; the first placeholder takes any overflow. */
function intPlaceholderChars(phs: ('0' | '#' | '?')[], digits: string): string[] {
  const res: string[] = new Array(phs.length).fill('');
  let di = digits.length - 1;
  for (let pi = phs.length - 1; pi >= 0; pi--) {
    if (di >= 0) {
      if (pi === 0) {
        res[0] = digits.slice(0, di + 1);
        di = -1;
      } else {
        res[pi] = digits[di];
        di--;
      }
    } else res[pi] = phs[pi] === '0' ? '0' : phs[pi] === '?' ? ' ' : '';
  }
  return res;
}

function formatNumberSection(sec: Section, value: number, forceMinus: boolean): { text: string; fillAt: number; fillChar?: string } {
  const toks = sec.toks;
  let pctCount = 0;
  let hasExp = false;
  let hasSlash = false;
  let dotIdx = -1;
  let lastDigit = -1;
  let firstDigit = -1;
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    if (t.t === 'pct') pctCount++;
    else if (t.t === 'exp') hasExp = true;
    else if (t.t === 'slash') hasSlash = true;
    else if (t.t === 'dot' && dotIdx < 0) dotIdx = i;
    else if (t.t === 'digit') {
      if (firstDigit < 0) firstDigit = i;
      lastDigit = i;
    }
  }
  let v = value * Math.pow(100, pctCount);
  // scaling commas: commas directly after the last digit placeholder (before dot or end)
  let grouping = false;
  const commaScale: Set<number> = new Set();
  for (let i = 0; i < toks.length; i++) {
    if (toks[i].t !== 'comma') continue;
    // comma between digit placeholders → grouping
    let prevDigit = false;
    let nextDigit = false;
    for (let j = i - 1; j >= 0; j--) {
      if (toks[j].t === 'digit') { prevDigit = true; break; }
      if (toks[j].t !== 'comma') break;
    }
    for (let j = i + 1; j < toks.length; j++) {
      if (toks[j].t === 'digit') { nextDigit = true; break; }
      if (toks[j].t !== 'comma') break;
    }
    const intEnd = dotIdx >= 0 ? dotIdx : lastDigit + 1;
    if (prevDigit && nextDigit && i < intEnd) grouping = true;
    else if (prevDigit && !nextDigit) commaScale.add(i);
  }
  v = v / Math.pow(1000, commaScale.size);
  const neg = v < 0;
  let abs = Math.abs(v);

  let fillAt = -1;
  let fillChar: string | undefined;

  if (hasSlash && !hasExp) {
    return formatFraction(sec, abs, neg && forceMinus);
  }

  // Collect placeholders by region
  const intPh: ('0' | '#' | '?')[] = [];
  const decPh: ('0' | '#' | '?')[] = [];
  const expPh: ('0' | '#' | '?')[] = [];
  let region: 'int' | 'dec' | 'exp' = 'int';
  for (const t of toks) {
    if (t.t === 'dot' && region === 'int') region = 'dec';
    else if (t.t === 'exp') region = 'exp';
    else if (t.t === 'digit') {
      if (region === 'int') intPh.push(t.v);
      else if (region === 'dec') decPh.push(t.v);
      else expPh.push(t.v);
    }
  }
  let intStr = '';
  let decStr = '';
  let expStr = '';
  let expNeg = false;
  if (hasExp) {
    let exp = abs === 0 ? 0 : Math.floor(Math.log10(abs));
    const intCount = Math.max(1, intPh.length);
    // engineering-style when int placeholders > 1: exponent multiple of intCount
    if (intPh.length > 1) exp = Math.floor(exp / intCount) * intCount;
    let mant = abs / Math.pow(10, exp);
    mant = +mant.toFixed(decPh.length);
    if (mant >= Math.pow(10, intPh.length > 1 ? intCount : 1)) {
      exp += intPh.length > 1 ? intCount : 1;
      mant = +(abs / Math.pow(10, exp)).toFixed(decPh.length);
    }
    const parts = mant.toFixed(decPh.length).split('.');
    intStr = parts[0] === '0' && !intPh.includes('0') ? '' : parts[0];
    decStr = parts[1] ?? '';
    expNeg = exp < 0;
    expStr = String(Math.abs(exp));
  } else {
    const fixed = abs.toFixed(Math.min(decPh.length, 20));
    const parts = fixed.split('.');
    intStr = parts[0];
    decStr = parts[1] ?? '';
    if (intStr === '0') intStr = '';
    if (grouping) intStr = groupDigits(intStr, sec.indian);
  }
  // trim decimal '#' trailing zeros
  let decOut = renderPlaceholders(decPh, decStr, false);
  if (decPh.length) {
    // remove trailing zeros for '#'
    let k = decPh.length - 1;
    const arr = decOut.split('');
    while (k >= 0 && decPh[k] === '#' && arr[k] === '0') {
      arr[k] = '';
      k--;
    }
    decOut = arr.join('');
  }
  let intOut: string;
  if (grouping && intStr.includes(',')) {
    // pad with zeros for extra '0' placeholders
    const zeroCount = intPh.filter((p) => p === '0').length;
    const digitsOnly = intStr.replace(/,/g, '');
    let padded = digitsOnly;
    while (padded.length < zeroCount) padded = '0' + padded;
    intOut = groupDigits(padded, sec.indian);
  } else if (grouping) {
    const zeroCount = intPh.filter((p) => p === '0').length;
    let padded = intStr;
    while (padded.length < zeroCount) padded = '0' + padded;
    intOut = padded.length > 3 ? groupDigits(padded, sec.indian) : padded;
    const qCount = intPh.filter((p) => p === '?').length;
    while (intOut.length < qCount) intOut = ' ' + intOut;
  } else intOut = '';

  // Build output by walking tokens
  let out = neg && forceMinus && (abs !== 0 || /[1-9]/.test(intStr + decStr)) ? '-' : '';
  if (neg && forceMinus && !/[1-9]/.test(intStr + decOut)) out = '';
  region = 'int';
  let intEmitted = false;
  let intIdx = 0;
  const posChars = grouping ? [] : intPlaceholderChars(intPh, intStr);
  let decIdx = 0;
  let expIdx = 0;
  const expOut = renderPlaceholders(expPh, expStr, true);
  for (let i = 0; i < toks.length; i++) {
    const t = toks[i];
    switch (t.t) {
      case 'digit':
        if (region === 'int') {
          if (grouping) {
            if (!intEmitted) {
              out += intOut;
              intEmitted = true;
            }
          } else out += posChars[intIdx] ?? '';
          intIdx++;
        } else if (region === 'dec') {
          out += decOut[decIdx] ?? '';
          decIdx++;
        } else {
          if (expIdx === 0) out += expOut;
          expIdx++;
        }
        break;
      case 'dot':
        if (region === 'int') {
          region = 'dec';
          // integer digits with no integer placeholders (e.g. ".00") still show
          if (intPh.length === 0) out += intStr.replace(/,/g, '');
        }
        out += '.';
        break;
      case 'comma':
        if (!grouping && !commaScale.has(i) && region === 'int' && intIdx === 0) out += ',';
        break;
      case 'pct':
        out += '%';
        break;
      case 'exp':
        region = 'exp';
        out += 'E' + (expNeg ? '-' : t.sign === '+' ? '+' : '');
        break;
      case 'lit':
        out += t.v;
        break;
      case 'skip':
        out += ' ';
        break;
      case 'fill':
        fillAt = out.length;
        fillChar = t.v;
        break;
      case 'text':
        break;
      default:
        break;
    }
  }
  // Remove trailing '.' when no decimals rendered and no placeholders exist? Excel keeps it (e.g. "0." shows "5.")
  return { text: out, fillAt, fillChar };
}

function formatFraction(sec: Section, abs: number, minus: boolean): { text: string; fillAt: number; fillChar?: string } {
  const toks = sec.toks;
  const slashIdx = toks.findIndex((t) => t.t === 'slash');
  // denominator placeholders or fixed denominator
  const denToks = toks.slice(slashIdx + 1);
  let denDigits = 0;
  let fixedDen = '';
  for (const t of denToks) {
    if (t.t === 'digit') {
      denDigits++;
    } else if (t.t === 'lit' && /^\d+$/.test(t.v)) fixedDen += t.v;
    else break;
  }
  // numerator part: find whether there's an integer part (digits, then lit space, then digits before slash)
  const before = toks.slice(0, slashIdx);
  let groups = 0;
  let inDigits = false;
  for (const t of before) {
    if (t.t === 'digit') {
      if (!inDigits) groups++;
      inDigits = true;
    } else inDigits = false;
  }
  const hasInt = groups >= 2;
  let whole = hasInt ? Math.floor(abs) : 0;
  let frac = hasInt ? abs - whole : abs;
  let n: number;
  let d: number;
  if (fixedDen) {
    d = parseInt(fixedDen, 10);
    n = Math.round(frac * d);
  } else {
    [n, d] = approxFraction(frac, Math.pow(10, Math.max(1, denDigits)) - 1);
  }
  if (hasInt && n === d) {
    whole += 1;
    n = 0;
  }
  let out = minus ? '-' : '';
  if (hasInt) {
    if (n === 0) return { text: out + String(whole), fillAt: -1 };
    out += (whole ? String(whole) + ' ' : '') + `${n}/${d}`;
  } else out += `${n}/${d}`;
  return { text: out, fillAt: -1 };
}

export function isDateFormat(fmt: string | undefined): boolean {
  if (!fmt || fmt === 'General') return false;
  const p = parseFormat(fmt);
  return p.sections[0]?.isDate ?? false;
}

export function isTextFormat(fmt: string | undefined): boolean {
  return fmt === '@';
}

export function formatValue(value: ScalarValue, fmt: string | undefined, generalWidth = 11): FormatResult {
  if (value === null || value === undefined || value === '') return { text: '' };
  if (typeof value === 'boolean') return { text: value ? 'TRUE' : 'FALSE' };
  const f = fmt && fmt !== 'General' ? fmt : '';
  if (typeof value === 'string') {
    if (!f) return { text: value };
    const p = parseFormat(f);
    const textSec = p.sections.length >= 4 ? p.sections[3] : p.sections.find((s) => s.isText);
    if (!textSec) return { text: value };
    let out = '';
    let fillAt = -1;
    let fillChar: string | undefined;
    for (const t of textSec.toks) {
      if (t.t === 'text') out += value;
      else if (t.t === 'lit') out += t.v;
      else if (t.t === 'skip') out += ' ';
      else if (t.t === 'fill') {
        fillAt = out.length;
        fillChar = t.v;
      }
    }
    return withFill({ text: out, color: textSec.color }, fillAt, fillChar);
  }
  // number
  if (!isFinite(value)) return { text: '#NUM!' };
  if (!f) return { text: formatGeneral(value, generalWidth), isNumber: true };
  const p = parseFormat(f);
  const secs = p.sections;
  let sec: Section;
  let forceMinus = true;
  let v = value;
  const hasConds = secs.some((s) => s.cond);
  if (hasConds) {
    const idx = secs.findIndex((s, i) => i < 3 && s.cond && testCond(s.cond, value));
    if (idx >= 0) {
      sec = secs[idx];
      forceMinus = value < 0 && !(sec.cond && (sec.cond.op === '<' || sec.cond.op === '<=') && sec.cond.val <= 0);
    } else {
      // fallthrough: first section without condition
      const plain = secs.filter((s, i) => i < 3 && !s.cond && !s.isText);
      sec = plain[0] ?? secs[secs.length - 1];
    }
  } else if (v > 0 || secs.length === 1 || (v === 0 && secs.length < 3)) {
    sec = secs[0];
    if (secs.length > 1 && v === 0) sec = secs[0];
  } else if (v < 0) {
    sec = secs[1] ?? secs[0];
    forceMinus = !secs[1];
    if (secs[1]) v = Math.abs(v);
  } else {
    sec = secs[2] ?? secs[0];
  }
  if (sec.isText && !sec.hasNumber) {
    // Text-only section applied to a number shows the number as General.
    return { text: formatGeneral(value, generalWidth), isNumber: true };
  }
  if (sec.isGeneral) {
    let out = '';
    for (const t of sec.toks) {
      if (t.t === 'lit') out += t.v === '\u0000GENERAL' ? formatGeneral(forceMinus ? v : Math.abs(v), generalWidth) : t.v;
      else if (t.t === 'skip') out += ' ';
    }
    return { text: out, color: sec.color, isNumber: true };
  }
  if (sec.isDate) {
    if (v < 0) return { text: '#'.repeat(10), isNumber: true };
    return { text: formatDateSection(sec, v), color: sec.color, isNumber: true };
  }
  const r = formatNumberSection(sec, v, forceMinus);
  return withFill({ text: r.text, color: sec.color, isNumber: true }, r.fillAt, r.fillChar);
}

function withFill(res: FormatResult, fillAt: number, fillChar?: string): FormatResult {
  if (fillAt < 0) return res;
  res.left = res.text.slice(0, fillAt);
  res.right = res.text.slice(fillAt);
  void fillChar;
  return res;
}

// ---------- Format builders used by the ribbon and Format Cells dialog ----------

export const FMT = {
  general: 'General',
  number: '0.00',
  currency: '"$"#,##0.00',
  accounting: '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)',
  shortDate: 'm/d/yyyy',
  longDate: 'dddd, mmmm d, yyyy',
  time: 'h:mm:ss AM/PM',
  percent: '0%',
  percent2: '0.00%',
  fraction: '# ?/?',
  scientific: '0.00E+00',
  text: '@',
  comma: '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)',
  inr: '[$₹-4009] #,##0.00',
  inrAccounting: '_ [$₹-4009] * #,##0.00_ ;_ [$₹-4009] * \\-#,##0.00_ ;_ [$₹-4009] * "-"??_ ;_ @_ ',
};

export function buildNumberFormat(decimals: number, thousands: boolean, negStyle = 0): string {
  const base = (thousands ? '#,##0' : '0') + (decimals > 0 ? '.' + '0'.repeat(decimals) : '');
  switch (negStyle) {
    case 1: return `${base};[Red]${base}`;
    case 2: return `${base}_);\\(${base}\\)`;
    case 3: return `${base}_);[Red]\\(${base}\\)`;
    default: return base;
  }
}

export function currencyPrefix(symbol: string): string {
  if (symbol === '₹') return '[$₹-4009] ';
  if (symbol === '$') return '"$"';
  if (symbol === '') return '';
  return `[$${symbol}]`;
}

export function buildCurrencyFormat(decimals: number, symbol: string, negStyle = 0): string {
  const num = '#,##0' + (decimals > 0 ? '.' + '0'.repeat(decimals) : '');
  const pre = currencyPrefix(symbol);
  const base = pre + num;
  switch (negStyle) {
    case 1: return `${base};[Red]${base}`;
    case 2: return `${base}_);\\(${base}\\)`;
    case 3: return `${base}_);[Red]\\(${base}\\)`;
    default: return `${base};-${base}`;
  }
}

export function buildAccountingFormat(decimals: number, symbol: string): string {
  const num = '#,##0' + (decimals > 0 ? '.' + '0'.repeat(decimals) : '');
  const q = decimals > 0 ? '.' + '?'.repeat(decimals) : '';
  const sym = symbol === '₹' ? '[$₹-4009] ' : symbol ? `"${symbol}"` : '';
  return `_(${sym}* ${num}_);_(${sym}* \\(${num}\\);_(${sym}* "-"${q}_);_(@_)`;
}

export function buildPercentFormat(decimals: number): string {
  return '0' + (decimals > 0 ? '.' + '0'.repeat(decimals) : '') + '%';
}

export function buildScientificFormat(decimals: number): string {
  return '0' + (decimals > 0 ? '.' + '0'.repeat(decimals) : '') + 'E+00';
}

/** Adjust decimals in a format code (Increase/Decrease Decimal buttons). */
export function adjustDecimals(fmt: string | undefined, delta: number, sampleValue?: number): string {
  let f = fmt && fmt !== 'General' ? fmt : '';
  if (!f) {
    // derive from current General display
    let dec = 0;
    if (typeof sampleValue === 'number') {
      const s = formatGeneral(sampleValue);
      const m = /\.(\d+)/.exec(s);
      dec = m ? m[1].length : 0;
    }
    const nd = Math.max(0, dec + delta);
    return nd === 0 ? '0' : '0.' + '0'.repeat(nd);
  }
  const secs = splitSections(f).map((s) => {
    if (delta > 0) {
      // add a 0 after the last '0' digit placeholder in decimal part (or add ".0")
      const re = /(0|#)(?![^"]*"(?:[^"]*"[^"]*")*[^"]*$)/g;
      let last = -1;
      let m: RegExpExecArray | null;
      while ((m = re.exec(s))) last = m.index;
      if (last < 0) return s;
      const hasDot = /\.(?=[0#?])/.test(s);
      return hasDot ? s.slice(0, last + 1) + '0' + s.slice(last + 1) : s.slice(0, last + 1) + '.0' + s.slice(last + 1);
    } else {
      const m = /\.([0#?]+)/.exec(s);
      if (!m) return s;
      const digits = m[1];
      const nd = digits.slice(0, -1);
      const rep = nd.length ? '.' + nd : '';
      return s.slice(0, m.index) + rep + s.slice(m.index + m[0].length);
    }
  });
  return secs.join(';');
}

export const DATE_FORMATS = [
  'm/d/yyyy',
  'dddd, mmmm d, yyyy',
  'm/d',
  'm/d/yy',
  'mm/dd/yy',
  'd-mmm',
  'd-mmm-yy',
  'dd-mmm-yy',
  'mmm-yy',
  'mmmm-yy',
  'mmmm d, yyyy',
  'm/d/yy h:mm AM/PM',
  'm/d/yy h:mm',
  'mmmmm',
  'mmmmm-yy',
  'd-mmm-yyyy',
  'dd/mm/yyyy',
  'yyyy-mm-dd',
];

export const TIME_FORMATS = ['h:mm:ss AM/PM', 'h:mm', 'h:mm AM/PM', 'h:mm:ss', 'mm:ss', 'mm:ss.0', '[h]:mm:ss', 'm/d/yy h:mm AM/PM', 'm/d/yy h:mm'];

export const CUSTOM_FORMATS = [
  'General',
  '0',
  '0.00',
  '#,##0',
  '#,##0.00',
  '#,##0_);(#,##0)',
  '#,##0_);[Red](#,##0)',
  '#,##0.00_);(#,##0.00)',
  '#,##0.00_);[Red](#,##0.00)',
  '"$"#,##0_);("$"#,##0)',
  '"$"#,##0.00_);[Red]("$"#,##0.00)',
  '0%',
  '0.00%',
  '0.00E+00',
  '##0.0E+0',
  '# ?/?',
  '# ??/??',
  'm/d/yyyy',
  'd-mmm-yy',
  'd-mmm',
  'mmm-yy',
  'h:mm AM/PM',
  'h:mm:ss AM/PM',
  'h:mm',
  'h:mm:ss',
  'm/d/yyyy h:mm',
  'mm:ss',
  'mm:ss.0',
  '@',
  '[h]:mm:ss',
  FMT.accounting,
  FMT.comma,
  FMT.inr,
  '[$₹-4009] #,##0',
  '[>=10000000]##\\,##\\,##\\,##0;[>=100000]##\\,##\\,##0;##,##0',
];

export const SPECIAL_FORMATS = [
  { name: 'Zip Code', code: '00000' },
  { name: 'Zip Code + 4', code: '00000-0000' },
  { name: 'Phone Number', code: '[<=9999999]###-####;(###) ###-####' },
  { name: 'Social Security Number', code: '000-00-0000' },
  { name: 'Indian Rupee (lakh/crore)', code: '[$₹-4009] #,##0.00' },
];
