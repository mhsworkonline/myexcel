// OpenDocument Spreadsheet (.ods) read/write.
import JSZip from 'jszip';
import { colToName, nameToCol } from '../model/address';
import { tokenize } from '../model/formula';
import { formatValue, isDateFormat, partsToSerial } from '../model/numfmt';
import { Sheet } from '../model/sheet';
import type { CellStyle } from '../model/styles';
import type { Cell, CellValue } from '../model/types';
import { Workbook } from '../model/workbook';
import { attr, children, esc, findAll, local, parseXml, textOf, XNode } from './xml';

// ---------- formula conversion ----------

function quoteOdfSheet(n: string): string {
  return /^[A-Za-z_][\w]*$/.test(n) ? n : `'${n.replace(/'/g, "''")}'`;
}

export function toOdfFormula(f: string): string {
  let out = '';
  for (const t of tokenize(f.replace(/^=/, '='))) {
    if (t.type === 'ref' && t.ref) {
      const text = t.text.includes('!') ? t.text.slice(t.text.lastIndexOf('!') + 1) : t.text;
      const pre = t.ref.sheet ? `$${quoteOdfSheet(t.ref.sheet)}` : '';
      const [a, b] = text.split(':');
      out += b ? `[${pre}.${a}:${pre}.${b}]` : `[${pre}.${a}]`;
    } else if (t.type === 'sep' && t.text === ',') out += ';';
    else out += t.text;
  }
  return 'of:' + out;
}

export function fromOdfFormula(f: string): string {
  let s = f.replace(/^(of:|oooc:)/, '');
  s = s.replace(/\[\$?('(?:[^']|'')*'|[^.\]:]*)?\.(\$?[A-Z]+\$?\d+)(?::\$?('(?:[^']|'')*'|[^.\]:]*)?\.(\$?[A-Z]+\$?\d+))?\]/g, (_m, sh: string | undefined, a: string, _sh2: string | undefined, b: string | undefined) => {
    const prefix = sh ? `${sh}!` : '';
    return prefix + a + (b ? ':' + b : '');
  });
  // ';' → ',' outside strings
  let out = '';
  let q = false;
  for (const ch of s) {
    if (ch === '"') q = !q;
    out += !q && ch === ';' ? ',' : ch;
  }
  return out.startsWith('=') ? out : '=' + out;
}

// ---------- writing ----------

function styleKey(s: CellStyle): string {
  return JSON.stringify([s.bold, s.italic, s.underline, s.strike, s.fontColor, s.fillColor, s.hAlign, s.vAlign, s.wrap, s.fontName, s.fontSize]);
}

function cellStyleXml(name: string, s: CellStyle): string {
  const text: string[] = [];
  if (s.bold) text.push('fo:font-weight="bold"');
  if (s.italic) text.push('fo:font-style="italic"');
  if (s.underline) text.push('style:text-underline-style="solid" style:text-underline-width="auto" style:text-underline-color="font-color"');
  if (s.strike) text.push('style:text-line-through-style="solid"');
  if (s.fontColor) text.push(`fo:color="${s.fontColor}"`);
  if (s.fontName) text.push(`style:font-name="${esc(s.fontName)}"`);
  if (s.fontSize) text.push(`fo:font-size="${s.fontSize}pt"`);
  const cell: string[] = [];
  if (s.fillColor) cell.push(`fo:background-color="${s.fillColor}"`);
  if (s.wrap) cell.push('fo:wrap-option="wrap"');
  if (s.vAlign) cell.push(`style:vertical-align="${s.vAlign === 'middle' ? 'middle' : s.vAlign === 'top' ? 'top' : 'bottom'}"`);
  const para = s.hAlign && ['left', 'center', 'right', 'justify'].includes(s.hAlign) ? `<style:paragraph-properties fo:text-align="${s.hAlign === 'left' ? 'start' : s.hAlign === 'right' ? 'end' : s.hAlign}"/>` : '';
  return `<style:style style:name="${name}" style:family="table-cell">${cell.length ? `<style:table-cell-properties ${cell.join(' ')}/>` : ''}${para}${text.length ? `<style:text-properties ${text.join(' ')}/>` : ''}</style:style>`;
}

export async function writeOds(wb: Workbook, getValue: (s: Sheet, r: number, c: number) => unknown): Promise<Uint8Array> {
  const styleNames = new Map<string, string>();
  const autoStyles: string[] = [];
  const colStyles = new Map<number, string>();
  const tables: string[] = [];
  for (const sheet of wb.sheets) {
    const used = sheet.usedRange();
    const r2 = used?.r2 ?? 0;
    const c2 = used?.c2 ?? 0;
    let cols = '';
    for (let c = 0; c <= c2; c++) {
      const w = sheet.colWidth(c);
      let cs = colStyles.get(w);
      if (!cs) {
        cs = `co${colStyles.size + 1}`;
        colStyles.set(w, cs);
        autoStyles.push(`<style:style style:name="${cs}" style:family="table-column"><style:table-column-properties style:column-width="${(w / 96).toFixed(4)}in"/></style:style>`);
      }
      cols += `<table:table-column table:style-name="${cs}"${sheet.hiddenCols.has(c) ? ' table:visibility="collapse"' : ''}/>`;
    }
    const covered = new Set<string>();
    for (const m of sheet.merges) for (let r = m.r1; r <= m.r2; r++) for (let c = m.c1; c <= m.c2; c++) if (r !== m.r1 || c !== m.c1) covered.add(`${r},${c}`);
    let rows = '';
    for (let r = 0; r <= r2; r++) {
      let cells = '';
      for (let c = 0; c <= c2; c++) {
        if (covered.has(`${r},${c}`)) {
          cells += '<table:covered-table-cell/>';
          continue;
        }
        const cell = sheet.getCell(r, c);
        const style = wb.styles.get(sheet.styleIdAt(r, c));
        let sAttr = '';
        const sk = styleKey(style);
        if (sk !== styleKey({})) {
          let sn = styleNames.get(sk);
          if (!sn) {
            sn = `ce${styleNames.size + 1}`;
            styleNames.set(sk, sn);
            autoStyles.push(cellStyleXml(sn, style));
          }
          sAttr = ` table:style-name="${sn}"`;
        }
        const m = sheet.merges.find((x) => x.r1 === r && x.c1 === c);
        const span = m ? ` table:number-columns-spanned="${m.c2 - m.c1 + 1}" table:number-rows-spanned="${m.r2 - m.r1 + 1}"` : '';
        if (!cell || (cell.v === undefined && cell.f === undefined)) {
          cells += `<table:table-cell${sAttr}${span}/>`;
          continue;
        }
        const v = cell.f ? getValue(sheet, r, c) : cell.v;
        const fAttr = cell.f ? ` table:formula="${esc(toOdfFormula(cell.f))}"` : '';
        const shown = typeof v === 'number' ? formatValue(v, style.numFmt).text : v === null || v === undefined ? '' : typeof v === 'object' ? String((v as { error?: string }).error ?? '') : String(v);
        let typeAttr: string;
        if (typeof v === 'number') {
          if (isDateFormat(style.numFmt)) {
            const d = new Date(Date.UTC(1899, 11, 30) + v * 86400000);
            typeAttr = ` office:value-type="date" office:date-value="${d.toISOString().slice(0, 19)}"`;
          } else if (style.numFmt?.includes('%')) typeAttr = ` office:value-type="percentage" office:value="${v}"`;
          else typeAttr = ` office:value-type="float" office:value="${v}"`;
        } else if (typeof v === 'boolean') typeAttr = ` office:value-type="boolean" office:boolean-value="${v}"`;
        else typeAttr = ' office:value-type="string"';
        const note = cell.note ? `<office:annotation><text:p>${esc(cell.note.text)}</text:p></office:annotation>` : '';
        const p = cell.link ? `<text:p><text:a xlink:href="${esc(cell.link)}">${esc(shown)}</text:a></text:p>` : `<text:p>${esc(shown)}</text:p>`;
        cells += `<table:table-cell${sAttr}${span}${typeAttr}${fAttr}>${note}${p}</table:table-cell>`;
      }
      const h = sheet.rowHeights.get(r);
      let rs = '';
      if (h) {
        rs = ` table:style-name="ro${h}"`;
        if (!autoStyles.some((x) => x.includes(`"ro${h}"`))) autoStyles.push(`<style:style style:name="ro${h}" style:family="table-row"><style:table-row-properties style:row-height="${(h / 96).toFixed(4)}in" style:use-optimal-row-height="false"/></style:style>`);
      }
      rows += `<table:table-row${rs}${sheet.hiddenRows.has(r) ? ' table:visibility="collapse"' : ''}>${cells}</table:table-row>`;
    }
    tables.push(`<table:table table:name="${esc(sheet.name)}">${cols}${rows}</table:table>`);
  }
  const names = wb.names.length
    ? `<table:named-expressions>${wb.names.map((n) => `<table:named-range table:name="${esc(n.name)}" table:base-cell-address="$${esc(quoteOdfSheet(wb.sheets[0].name))}.$A$1" table:cell-range-address="${esc(n.ref.replace(/!/g, '.').replace(/^/, '$'))}"/>`).join('')}</table:named-expressions>`
    : '';
  const ns =
    'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" xmlns:xlink="http://www.w3.org/1999/xlink" xmlns:of="urn:oasis:names:tc:opendocument:xmlns:of:1.2"';
  const content = `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-content ${ns} office:version="1.2"><office:automatic-styles>${autoStyles.join('')}</office:automatic-styles><office:body><office:spreadsheet>${tables.join('')}${names}</office:spreadsheet></office:body></office:document-content>`;
  const zip = new JSZip();
  zip.file('mimetype', 'application/vnd.oasis.opendocument.spreadsheet', { compression: 'STORE' });
  zip.file('content.xml', content);
  zip.file('styles.xml', `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-styles ${ns} office:version="1.2"/>`);
  zip.file('meta.xml', `<?xml version="1.0" encoding="UTF-8"?>\n<office:document-meta xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" xmlns:meta="urn:oasis:names:tc:opendocument:xmlns:meta:1.0" office:version="1.2"><office:meta><meta:generator>MyExcel</meta:generator></office:meta></office:document-meta>`);
  zip.file(
    'META-INF/manifest.xml',
    `<?xml version="1.0" encoding="UTF-8"?>\n<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.spreadsheet"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="styles.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="meta.xml" manifest:media-type="text/xml"/></manifest:manifest>`,
  );
  return zip.generateAsync({ type: 'uint8array', mimeType: 'application/vnd.oasis.opendocument.spreadsheet' });
}

// ---------- reading ----------

function parseStyles(root: XNode): Map<string, CellStyle> {
  const out = new Map<string, CellStyle>();
  for (const st of findAll(root, 'style')) {
    if (attr(st, 'family') !== 'table-cell') continue;
    const s: CellStyle = {};
    for (const p of st.children) {
      const ln = local(p);
      if (ln === 'text-properties') {
        if (attr(p, 'font-weight') === 'bold') s.bold = true;
        if (attr(p, 'font-style') === 'italic') s.italic = true;
        if (attr(p, 'text-underline-style') && attr(p, 'text-underline-style') !== 'none') s.underline = 'single';
        if (attr(p, 'text-line-through-style') && attr(p, 'text-line-through-style') !== 'none') s.strike = true;
        const col = attr(p, 'color');
        if (col && col !== '#000000') s.fontColor = col.toUpperCase();
        const fs = attr(p, 'font-size');
        if (fs && fs.endsWith('pt') && parseFloat(fs) !== 11) s.fontSize = parseFloat(fs);
        const fn = attr(p, 'font-name');
        if (fn) s.fontName = fn;
      } else if (ln === 'table-cell-properties') {
        const bg = attr(p, 'background-color');
        if (bg && bg !== 'transparent') s.fillColor = bg.toUpperCase();
        if (attr(p, 'wrap-option') === 'wrap') s.wrap = true;
        const va = attr(p, 'vertical-align');
        if (va === 'top' || va === 'middle') s.vAlign = va;
      } else if (ln === 'paragraph-properties') {
        const ta = attr(p, 'text-align');
        if (ta === 'center') s.hAlign = 'center';
        else if (ta === 'end' || ta === 'right') s.hAlign = 'right';
        else if (ta === 'start' || ta === 'left') s.hAlign = 'left';
      }
    }
    out.set(attr(st, 'name') ?? '', s);
  }
  return out;
}

function parseLen(s: string | undefined): number | undefined {
  if (!s) return undefined;
  const n = parseFloat(s);
  if (s.endsWith('in')) return n * 96;
  if (s.endsWith('cm')) return (n / 2.54) * 96;
  if (s.endsWith('mm')) return (n / 25.4) * 96;
  if (s.endsWith('pt')) return (n * 4) / 3;
  return undefined;
}

export async function readOds(data: ArrayBuffer): Promise<Workbook> {
  const zip = await JSZip.loadAsync(data);
  const content = parseXml((await zip.file('content.xml')?.async('string')) ?? '');
  const styles = parseStyles(content);
  const colW = new Map<string, number>();
  const rowH = new Map<string, number>();
  for (const st of findAll(content, 'style')) {
    const fam = attr(st, 'family');
    if (fam === 'table-column') {
      const w = parseLen(attr(children(st, 'table-column-properties')[0], 'column-width'));
      if (w) colW.set(attr(st, 'name') ?? '', Math.round(w));
    } else if (fam === 'table-row') {
      const h = parseLen(attr(children(st, 'table-row-properties')[0], 'row-height'));
      if (h) rowH.set(attr(st, 'name') ?? '', Math.round(h));
    }
  }
  const wb = new Workbook();
  const dateFmt = wb.styles.intern({ numFmt: 'm/d/yyyy' });
  const pctFmt = wb.styles.intern({ numFmt: '0%' });
  for (const t of findAll(content, 'table')) {
    if (t.name !== 'table:table') continue;
    const sheet = new Sheet((attr(t, 'name') ?? `Sheet${wb.sheets.length + 1}`).slice(0, 31));
    let c = 0;
    for (const col of findAll(t, 'table-column')) {
      const rep = Math.min(1024, +(attr(col, 'number-columns-repeated') ?? 1));
      const w = colW.get(attr(col, 'style-name') ?? '');
      for (let k = 0; k < rep; k++) {
        if (w && w !== 64) sheet.colWidths.set(c, w);
        if (attr(col, 'visibility') === 'collapse') sheet.hiddenCols.add(c);
        c++;
      }
    }
    let r = 0;
    for (const row of findAll(t, 'table-row')) {
      const rrep = Math.min(+(attr(row, 'number-rows-repeated') ?? 1), 100000);
      const hasContent = row.children.some((x) => x.children.length || Object.keys(x.attrs).some((k) => k.includes('value') || k.includes('formula') || k.includes('spanned')));
      if (!hasContent) {
        r += rrep;
        continue;
      }
      for (let k = 0; k < Math.min(rrep, 1000); k++) {
        const h = rowH.get(attr(row, 'style-name') ?? '');
        if (h && h !== 20) sheet.rowHeights.set(r, h);
        if (attr(row, 'visibility') === 'collapse') sheet.hiddenRows.add(r);
        let cc = 0;
        for (const cellEl of row.children) {
          const ln = local(cellEl);
          if (ln !== 'table-cell' && ln !== 'covered-table-cell') continue;
          const crep = Math.min(+(attr(cellEl, 'number-columns-repeated') ?? 1), 16384);
          if (ln === 'covered-table-cell' || (!cellEl.children.length && !attr(cellEl, 'value-type') && !attr(cellEl, 'style-name') && !attr(cellEl, 'number-columns-spanned'))) {
            cc += crep;
            continue;
          }
          for (let j = 0; j < Math.min(crep, 256); j++) {
            const cell: Cell = {};
            const type = attr(cellEl, 'value-type');
            const f = attr(cellEl, 'formula');
            let v: CellValue = null;
            if (type === 'float' || type === 'currency' || type === 'percentage') v = +(attr(cellEl, 'value') ?? 0);
            else if (type === 'boolean') v = attr(cellEl, 'boolean-value') === 'true';
            else if (type === 'date') {
              const d = attr(cellEl, 'date-value') ?? '';
              const m = /^(\d{4})-(\d\d)-(\d\d)(?:T(\d\d):(\d\d):(\d\d))?/.exec(d);
              if (m) v = partsToSerial(+m[1], +m[2], +m[3], +(m[4] ?? 0), +(m[5] ?? 0), +(m[6] ?? 0));
              cell.s = dateFmt;
            } else if (type === 'time') {
              const m = /PT(\d+)H(\d+)M([\d.]+)S/.exec(attr(cellEl, 'time-value') ?? '');
              if (m) v = (+m[1] * 3600 + +m[2] * 60 + +m[3]) / 86400;
            } else if (type === 'string' || cellEl.children.length) v = children(cellEl, 'p').map(textOf).join('\n') || null;
            if (type === 'percentage') cell.s = pctFmt;
            if (v !== null) cell.v = v;
            if (f) cell.f = fromOdfFormula(f);
            const st = styles.get(attr(cellEl, 'style-name') ?? '');
            if (st && Object.keys(st).length) cell.s = wb.styles.merge(cell.s, st);
            const ann = children(cellEl, 'annotation')[0];
            if (ann) cell.note = { text: children(ann, 'p').map(textOf).join('\n') };
            const a = findAll(cellEl, 'a')[0];
            if (a) cell.link = attr(a, 'href');
            const cs = +(attr(cellEl, 'number-columns-spanned') ?? 1);
            const rs = +(attr(cellEl, 'number-rows-spanned') ?? 1);
            if (cs > 1 || rs > 1) sheet.merges.push({ r1: r, c1: cc, r2: r + rs - 1, c2: cc + cs - 1 });
            if (Object.keys(cell).length) sheet.setCellRaw(r, cc, cell);
            cc++;
          }
          cc += Math.max(0, crep - 256);
        }
        r++;
      }
      r += Math.max(0, rrep - 1000);
    }
    sheet.touch();
    wb.sheets.push(sheet);
  }
  if (!wb.sheets.length) wb.sheets.push(new Sheet('Sheet1'));
  wb.activeSheetId = wb.sheets[0].id;
  for (const nr of findAll(content, 'named-range')) {
    const addr = (attr(nr, 'cell-range-address') ?? '').replace(/^\$/, '').replace(/\.\$?/g, (m) => '!' + (m.length > 1 ? '$' : ''));
    const name = attr(nr, 'name');
    if (name) wb.names.push({ name, ref: addr.replace(/:\$?[^!]*!/, ':') });
  }
  void colToName;
  void nameToCol;
  return wb;
}
