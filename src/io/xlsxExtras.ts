// Post-processing of ExcelJS output for parts ExcelJS doesn't support:
// native charts (DrawingML), constant / sheet-scoped defined names, and MyExcel metadata.
import JSZip from 'jszip';
import { colToName, parseRange, Range } from '../model/address';
import { sheetLayout } from '../model/layout';
import { newId, Sheet } from '../model/sheet';
import type { ChartSeries, ChartSpec, ChartType, PivotSpec, Scenario, SparklineGroup } from '../model/types';
import type { Workbook } from '../model/workbook';
import { attr, child, children, esc, findAll, local, parseXml, textOf, XNode } from './xml';
import { readXlsx, ValueReader, writeXlsx } from './xlsx';

const EMU = 9525;
const NS_R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const REL_DRAWING = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing';
const REL_CHART = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart';
const REL_CUSTOMXML = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXml';
const META_NS = 'urn:myexcel:meta';

interface SheetPart {
  name: string;
  path: string; // e.g. xl/worksheets/sheet1.xml
  relsPath: string;
}

function relsPathFor(path: string): string {
  const i = path.lastIndexOf('/');
  return `${path.slice(0, i)}/_rels/${path.slice(i + 1)}.rels`;
}

function resolveTarget(base: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = base.split('/').slice(0, -1);
  for (const seg of target.split('/')) {
    if (seg === '..') parts.pop();
    else if (seg !== '.') parts.push(seg);
  }
  return parts.join('/');
}

async function sheetParts(zip: JSZip): Promise<SheetPart[]> {
  const wbXml = parseXml((await zip.file('xl/workbook.xml')?.async('string')) ?? '');
  const relsXml = parseXml((await zip.file('xl/_rels/workbook.xml.rels')?.async('string')) ?? '');
  const rels = new Map(findAll(relsXml, 'Relationship').map((r) => [r.attrs.Id, r.attrs.Target]));
  return findAll(wbXml, 'sheet').map((s) => {
    const rid = attr(s, 'id') ?? '';
    const path = resolveTarget('xl/workbook.xml', rels.get(rid) ?? '');
    return { name: s.attrs.name, path, relsPath: relsPathFor(path) };
  });
}

// ---------- value caches ----------

function resolveRefIn(wb: Workbook, host: Sheet, ref: string): { sheet: Sheet; range: Range } | null {
  let t = ref.trim().replace(/^=/, '');
  let sheet = host;
  const bang = t.lastIndexOf('!');
  if (bang > 0) {
    const s = wb.sheetByName(t.slice(0, bang).replace(/^'|'$/g, '').replace(/''/g, "'"));
    if (!s) return null;
    sheet = s;
    t = t.slice(bang + 1);
  }
  const rg = parseRange(t.replace(/\$/g, ''));
  return rg ? { sheet, range: rg } : null;
}

function cacheValues(wb: Workbook, host: Sheet, ref: string | undefined, getValue: ValueReader): unknown[] {
  if (!ref) return [];
  const res = resolveRefIn(wb, host, ref);
  if (!res) return [];
  const out: unknown[] = [];
  for (let r = res.range.r1; r <= Math.min(res.range.r2, res.range.r1 + 10000); r++)
    for (let c = res.range.c1; c <= Math.min(res.range.c2, res.range.c1 + 500); c++) out.push(getValue(res.sheet, r, c));
  return out;
}

function absRef(ref: string): string {
  // make sure references are sheet-qualified & absolute for Excel
  return ref.replace(/(^|[^$A-Z])([A-Z]{1,3})(\d+)/g, (m, p, c, r) => `${p}$${c}$${r}`);
}

function strRef(tag: string, ref: string | undefined, vals: unknown[]): string {
  if (!ref) return '';
  const pts = vals.map((v, i) => `<c:pt idx="${i}"><c:v>${esc(String(v ?? ''))}</c:v></c:pt>`).join('');
  return `<c:${tag}><c:strRef><c:f>${esc(absRef(ref))}</c:f><c:strCache><c:ptCount val="${vals.length}"/>${pts}</c:strCache></c:strRef></c:${tag}>`;
}

function numRef(tag: string, ref: string, vals: unknown[]): string {
  const pts = vals.map((v, i) => (typeof v === 'number' ? `<c:pt idx="${i}"><c:v>${v}</c:v></c:pt>` : '')).join('');
  return `<c:${tag}><c:numRef><c:f>${esc(absRef(ref))}</c:f><c:numCache><c:formatCode>General</c:formatCode><c:ptCount val="${vals.length}"/>${pts}</c:numCache></c:numRef></c:${tag}>`;
}

const PALETTE = ['4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47', '264478', '9E480E', '636363', '997300'];

function seriesXml(wb: Workbook, host: Sheet, ch: ChartSpec, s: ChartSeries, i: number, kind: 'bar' | 'line' | 'area' | 'pie' | 'scatter', getValue: ValueReader): string {
  const color = (s.color ?? '#' + PALETTE[i % PALETTE.length]).replace('#', '');
  let tx = '';
  if (s.nameRef) tx = strRef('tx', s.nameRef, cacheValues(wb, host, s.nameRef, getValue));
  else if (s.name) tx = `<c:tx><c:v>${esc(s.name)}</c:v></c:tx>`;
  const sp =
    kind === 'line' || kind === 'scatter'
      ? `<c:spPr><a:ln w="28575" cap="rnd"><a:solidFill><a:srgbClr val="${color}"/></a:solidFill><a:round/></a:ln></c:spPr>`
      : kind === 'pie'
        ? ''
        : `<c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr>`;
  const marker =
    kind === 'line'
      ? ch.type === 'lineMarkers' || ch.type === 'combo'
        ? `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:marker>`
        : '<c:marker><c:symbol val="none"/></c:marker>'
      : kind === 'scatter'
        ? `<c:marker><c:symbol val="circle"/><c:size val="5"/><c:spPr><a:solidFill><a:srgbClr val="${color}"/></a:solidFill></c:spPr></c:marker>`
        : '';
  const inv = kind === 'bar' ? '<c:invertIfNegative val="0"/>' : '';
  const dl = ch.dataLabels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
  const vals = cacheValues(wb, host, s.values, getValue);
  if (kind === 'scatter') {
    const xv = s.xValues ? numRef('xVal', s.xValues, cacheValues(wb, host, s.xValues, getValue)) : '';
    return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}<c:spPr><a:ln w="19050"><a:noFill/></a:ln></c:spPr>${marker}${dl}${xv}${numRef('yVal', s.values, vals)}<c:smooth val="0"/></c:ser>`;
  }
  const cat = s.categories ? strRef('cat', s.categories, cacheValues(wb, host, s.categories, getValue)) : '';
  const smooth = kind === 'line' ? '<c:smooth val="0"/>' : '';
  return `<c:ser><c:idx val="${i}"/><c:order val="${i}"/>${tx}${sp}${inv}${marker}${dl}${cat}${numRef('val', s.values, vals)}${smooth}</c:ser>`;
}

function axes(catId: number, valId: number, horizontal: boolean, gridlines: boolean, scatter: boolean, secondary = false, xTitle?: string, yTitle?: string): string {
  const title = (t?: string) => (t ? `<c:title><c:tx><c:rich><a:bodyPr/><a:p><a:r><a:t>${esc(t)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title>` : '');
  const catPos = secondary ? 'b' : horizontal ? 'l' : 'b';
  const valPos = secondary ? 'r' : horizontal ? 'b' : 'l';
  const catAx = scatter
    ? `<c:valAx><c:axId val="${catId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="b"/>${title(xTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valId}"/><c:crosses val="autoZero"/><c:crossBetween val="midCat"/></c:valAx>`
    : `<c:catAx><c:axId val="${catId}"/><c:scaling><c:orientation val="${horizontal ? 'minMax' : 'minMax'}"/></c:scaling><c:delete val="${secondary ? 1 : 0}"/><c:axPos val="${catPos}"/>${secondary ? '' : title(xTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="${valId}"/><c:crosses val="autoZero"/><c:auto val="1"/><c:lblAlgn val="ctr"/><c:lblOffset val="100"/></c:catAx>`;
  const valAx = `<c:valAx><c:axId val="${valId}"/><c:scaling><c:orientation val="minMax"/></c:scaling><c:delete val="0"/><c:axPos val="${valPos}"/>${gridlines && !secondary ? '<c:majorGridlines><c:spPr><a:ln w="9525"><a:solidFill><a:srgbClr val="D9D9D9"/></a:solidFill></a:ln></c:spPr></c:majorGridlines>' : ''}${secondary ? '' : title(yTitle)}<c:numFmt formatCode="General" sourceLinked="1"/><c:tickLblPos val="nextTo"/><c:crossAx val="${catId}"/><c:crosses val="${secondary ? 'max' : 'autoZero'}"/><c:crossBetween val="${scatter ? 'midCat' : 'between'}"/></c:valAx>`;
  return catAx + valAx;
}

export function chartXml(wb: Workbook, host: Sheet, ch: ChartSpec, getValue: ValueReader): string {
  const t = ch.type;
  const plot: string[] = [];
  const ser = (list: ChartSeries[], kind: Parameters<typeof seriesXml>[5], offset = 0) => list.map((s, i) => seriesXml(wb, host, ch, s, i + offset, kind, getValue)).join('');
  const dLbls = ch.dataLabels ? '<c:dLbls><c:showLegendKey val="0"/><c:showVal val="1"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/></c:dLbls>' : '';
  if (t === 'pie' || t === 'doughnut') {
    const tag = t === 'pie' ? 'pieChart' : 'doughnutChart';
    plot.push(`<c:${tag}><c:varyColors val="1"/>${ser(ch.series, 'pie')}${dLbls}<c:firstSliceAng val="0"/>${t === 'doughnut' ? '<c:holeSize val="50"/>' : ''}</c:${tag}>`);
  } else if (t === 'scatter') {
    plot.push(`<c:scatterChart><c:scatterStyle val="lineMarker"/><c:varyColors val="0"/>${ser(ch.series, 'scatter')}<c:axId val="1001"/><c:axId val="1002"/></c:scatterChart>`);
    plot.push(axes(1001, 1002, false, ch.gridlines, true, false, ch.xTitle, ch.yTitle));
  } else {
    const typeOf = (s: ChartSeries): 'bar' | 'line' | 'area' =>
      t === 'combo' ? (s.seriesType === 'line' ? 'line' : s.seriesType === 'area' ? 'area' : 'bar') : t.includes('olumn') || t === 'column' || t.includes('ar') && !t.includes('rea') ? 'bar' : t.includes('rea') ? 'area' : 'line';
    const horizontal = t === 'bar' || t === 'stackedBar';
    const stacked = t.startsWith('stacked');
    const groups: Record<string, { list: ChartSeries[]; idx: number[] }> = {};
    ch.series.forEach((s, i) => {
      const k = typeOf(s) + (s.secondaryAxis ? '2' : '');
      (groups[k] ??= { list: [], idx: [] }).list.push(s);
      groups[k].idx.push(i);
    });
    let needSecondary = false;
    for (const [k, g] of Object.entries(groups)) {
      const secondary = k.endsWith('2');
      if (secondary) needSecondary = true;
      const ids = secondary ? '<c:axId val="2001"/><c:axId val="2002"/>' : '<c:axId val="1001"/><c:axId val="1002"/>';
      const sers = g.list.map((s, j) => seriesXml(wb, host, ch, s, g.idx[j], k.startsWith('bar') ? 'bar' : k.startsWith('line') ? 'line' : 'area', getValue)).join('');
      if (k.startsWith('bar'))
        plot.push(`<c:barChart><c:barDir val="${horizontal ? 'bar' : 'col'}"/><c:grouping val="${stacked ? 'stacked' : 'clustered'}"/><c:varyColors val="0"/>${sers}${dLbls}<c:gapWidth val="219"/><c:overlap val="${stacked ? 100 : -27}"/>${ids}</c:barChart>`);
      else if (k.startsWith('line')) plot.push(`<c:lineChart><c:grouping val="standard"/><c:varyColors val="0"/>${sers}${dLbls}<c:marker val="1"/>${ids}</c:lineChart>`);
      else plot.push(`<c:areaChart><c:grouping val="${stacked ? 'stacked' : 'standard'}"/><c:varyColors val="0"/>${sers}${dLbls}${ids}</c:areaChart>`);
    }
    plot.push(axes(1001, 1002, horizontal, ch.gridlines, false, false, ch.xTitle, ch.yTitle));
    if (needSecondary) plot.push(axes(2001, 2002, false, false, false, true));
  }
  const title =
    ch.showTitle && ch.title
      ? `<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="1400" b="0"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:r><a:rPr lang="en-US"/><a:t>${esc(ch.title)}</a:t></a:r></a:p></c:rich></c:tx><c:overlay val="0"/></c:title><c:autoTitleDeleted val="0"/>`
      : '<c:autoTitleDeleted val="1"/>';
  const legendPos: Record<string, string> = { right: 'r', left: 'l', top: 't', bottom: 'b' };
  const legend = ch.legend !== 'none' ? `<c:legend><c:legendPos val="${legendPos[ch.legend]}"/><c:overlay val="0"/></c:legend>` : '';
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n` +
    `<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="${NS_R}">` +
    `<c:roundedCorners val="0"/><c:chart>${title}<c:plotArea><c:layout/>${plot.join('')}</c:plotArea>${legend}<c:plotVisOnly val="1"/><c:dispBlanksAs val="gap"/></c:chart>` +
    `<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz="900"><a:solidFill><a:srgbClr val="595959"/></a:solidFill></a:defRPr></a:pPr><a:endParaRPr lang="en-US"/></a:p></c:txPr></c:chartSpace>`
  );
}

function anchorXml(sheet: Sheet, ch: ChartSpec, idx: number, rid: string): string {
  const { rows, cols } = sheetLayout(sheet);
  const cell = (x: number, y: number) => {
    const c = cols.indexAt(x);
    const r = rows.indexAt(y);
    return { c, r, dx: Math.max(0, Math.round((x - cols.offset(c)) * EMU)), dy: Math.max(0, Math.round((y - rows.offset(r)) * EMU)) };
  };
  const a = cell(ch.anchor.x, ch.anchor.y);
  const b = cell(ch.anchor.x + ch.anchor.w, ch.anchor.y + ch.anchor.h);
  return (
    `<xdr:twoCellAnchor editAs="oneCell"><xdr:from><xdr:col>${a.c}</xdr:col><xdr:colOff>${a.dx}</xdr:colOff><xdr:row>${a.r}</xdr:row><xdr:rowOff>${a.dy}</xdr:rowOff></xdr:from>` +
    `<xdr:to><xdr:col>${b.c}</xdr:col><xdr:colOff>${b.dx}</xdr:colOff><xdr:row>${b.r}</xdr:row><xdr:rowOff>${b.dy}</xdr:rowOff></xdr:to>` +
    `<xdr:graphicFrame macro=""><xdr:nvGraphicFramePr><xdr:cNvPr id="${idx + 2}" name="Chart ${idx + 1}"/><xdr:cNvGraphicFramePr/></xdr:nvGraphicFramePr>` +
    `<xdr:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/></xdr:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart">` +
    `<c:chart xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:r="${NS_R}" r:id="${rid}"/></a:graphicData></a:graphic></xdr:graphicFrame><xdr:clientData/></xdr:twoCellAnchor>`
  );
}

function addOverride(ct: string, part: string, type: string): string {
  if (ct.includes(`PartName="/${part}"`)) return ct;
  return ct.replace('</Types>', `<Override PartName="/${part}" ContentType="${type}"/></Types>`);
}

function addRel(rels: string | null, id: string, type: string, target: string): string {
  const base = rels ?? `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"></Relationships>`;
  return base.replace('</Relationships>', `<Relationship Id="${id}" Type="${type}" Target="${target}"/></Relationships>`);
}

/** Insert <drawing r:id> at the schema-correct position inside a worksheet part. */
function insertDrawingEl(xml: string, rid: string): string {
  const el = `<drawing r:id="${rid}"/>`;
  let out = xml;
  if (!/xmlns:r=/.test(out.slice(0, out.indexOf('>', out.indexOf('<worksheet')) + 1))) out = out.replace('<worksheet', `<worksheet xmlns:r="${NS_R}"`);
  for (const before of ['<legacyDrawing', '<legacyDrawingHF', '<picture', '<oleObjects', '<controls', '<webPublishItems', '<tableParts', '<extLst']) {
    const i = out.indexOf(before);
    if (i >= 0) return out.slice(0, i) + el + out.slice(i);
  }
  return out.replace('</worksheet>', el + '</worksheet>');
}

interface MetaSheet {
  name: string;
  pivots?: (Omit<PivotSpec, 'sourceSheetId'> & { sourceSheet: string })[];
  sparklines?: SparklineGroup[];
  scenarios?: Scenario[];
  charts?: ChartSpec[];
}

function buildMeta(wb: Workbook): string {
  const sheets: MetaSheet[] = wb.sheets.map((s) => ({
    name: s.name,
    pivots: s.pivots.length ? s.pivots.map(({ sourceSheetId, ...p }) => ({ ...p, sourceSheet: wb.sheetById(sourceSheetId)?.name ?? '' })) : undefined,
    sparklines: s.sparklines.length ? s.sparklines : undefined,
    scenarios: s.scenarios.length ? s.scenarios : undefined,
    charts: s.charts.length ? s.charts : undefined,
  }));
  return JSON.stringify({ version: 1, sheets });
}

function b64(s: string): string {
  const bytes = new TextEncoder().encode(s);
  let bin = '';
  for (let i = 0; i < bytes.length; i += 0x8000) bin += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return btoa(bin);
}

function unb64(s: string): string {
  const bin = atob(s.trim());
  const bytes = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i++) bytes[i] = bin.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

function isRangeRef(ref: string): boolean {
  return ref.split(',').every((r) => /^(('[^']+'|[A-Za-z_][\w.]*)!)?\$?[A-Z]{1,3}\$?\d+(:\$?[A-Z]{1,3}\$?\d+)?$/.test(r.trim()));
}

export async function writeXlsxFull(wb: Workbook, getValue: ValueReader): Promise<ArrayBuffer> {
  const base = await writeXlsx(wb, getValue);
  const zip = await JSZip.loadAsync(base);
  let ct = (await zip.file('[Content_Types].xml')!.async('string'))!;
  const parts = await sheetParts(zip);
  let chartNo = 1;
  let drawingNo = 1;
  for (const s of wb.sheets) {
    if (!s.charts.length) continue;
    const part = parts.find((p) => p.name === s.name);
    if (!part) continue;
    while (zip.file(`xl/drawings/drawing${drawingNo}.xml`)) drawingNo++;
    const dPath = `xl/drawings/drawing${drawingNo}.xml`;
    let dRels: string | null = null;
    const anchors: string[] = [];
    s.charts.forEach((ch, i) => {
      while (zip.file(`xl/charts/chart${chartNo}.xml`)) chartNo++;
      const cPath = `xl/charts/chart${chartNo}.xml`;
      zip.file(cPath, chartXml(wb, s, ch, getValue));
      ct = addOverride(ct, cPath, 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml');
      const rid = `rId${i + 1}`;
      dRels = addRel(dRels, rid, REL_CHART, `../charts/chart${chartNo}.xml`);
      anchors.push(anchorXml(s, ch, i, rid));
      chartNo++;
    });
    zip.file(
      dPath,
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">${anchors.join('')}</xdr:wsDr>`,
    );
    zip.file(relsPathFor(dPath), dRels!);
    ct = addOverride(ct, dPath, 'application/vnd.openxmlformats-officedocument.drawing+xml');
    const sheetRels = (await zip.file(part.relsPath)?.async('string')) ?? null;
    const rid = 'rIdMxDrawing1';
    zip.file(part.relsPath, addRel(sheetRels, rid, REL_DRAWING, `../drawings/drawing${drawingNo}.xml`));
    const sheetXml = (await zip.file(part.path)!.async('string'))!;
    zip.file(part.path, insertDrawingEl(sheetXml, rid));
    drawingNo++;
  }
  // defined names (constants, formulas and sheet scope that ExcelJS can't express)
  let wbXml = (await zip.file('xl/workbook.xml')!.async('string'))!;
  const keep = [...wbXml.matchAll(/<definedName [^>]*name="(_xlnm[^"]*)"[^>]*>[\s\S]*?<\/definedName>/g)].map((m) => m[0]);
  const ours = wb.names.map((n) => {
    const local = n.scope ? ` localSheetId="${wb.sheets.findIndex((s) => s.id === n.scope)}"` : '';
    const cmt = n.comment ? ` comment="${esc(n.comment)}"` : '';
    return `<definedName name="${esc(n.name)}"${local}${cmt}>${esc(isRangeRef(n.ref) ? n.ref : n.ref.replace(/^=/, ''))}</definedName>`;
  });
  wbXml = wbXml.replace(/<definedNames>[\s\S]*?<\/definedNames>/, '');
  const all = [...keep, ...ours];
  if (all.length) {
    const dn = `<definedNames>${all.join('')}</definedNames>`;
    const at = wbXml.indexOf('<calcPr');
    wbXml = at >= 0 ? wbXml.slice(0, at) + dn + wbXml.slice(at) : wbXml.replace('</sheets>', '</sheets>' + dn);
  }
  // MyExcel metadata as a custom XML part
  const meta = buildMeta(wb);
  if (meta.includes('"pivots"') || meta.includes('"sparklines"') || meta.includes('"scenarios"') || meta.includes('"charts"')) {
    let n = 1;
    while (zip.file(`customXml/item${n}.xml`)) n++;
    zip.file(`customXml/item${n}.xml`, `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<myexcel xmlns="${META_NS}" encoding="base64">${b64(meta)}</myexcel>`);
    zip.file(
      `customXml/itemProps${n}.xml`,
      `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\n<ds:datastoreItem ds:itemID="{6F1C5B2A-${String(n).padStart(4, '0')}-4C1D-9E2B-4D59A11E0001}" xmlns:ds="http://schemas.openxmlformats.org/officeDocument/2006/customXml"><ds:schemaRefs><ds:schemaRef ds:uri="${META_NS}"/></ds:schemaRefs></ds:datastoreItem>`,
    );
    zip.file(`customXml/_rels/item${n}.xml.rels`, addRel(null, 'rId1', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/customXmlProps', `itemProps${n}.xml`));
    ct = addOverride(ct, `customXml/itemProps${n}.xml`, 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml');
    if (!/Extension="xml"/.test(ct)) ct = ct.replace('<Override', '<Default Extension="xml" ContentType="application/xml"/><Override');
    const wbRels = (await zip.file('xl/_rels/workbook.xml.rels')!.async('string'))!;
    zip.file('xl/_rels/workbook.xml.rels', addRel(wbRels, 'rIdMxMeta1', REL_CUSTOMXML, `../customXml/item${n}.xml`));
  }
  zip.file('xl/workbook.xml', wbXml);
  zip.file('[Content_Types].xml', ct);
  return zip.generateAsync({ type: 'arraybuffer', compression: 'DEFLATE' });
}

// ---------- reading ----------

function chartTypeFrom(plot: XNode): { type: ChartType; kinds: string[] } {
  const kinds = plot.children.map(local).filter((n) => n.endsWith('Chart'));
  const first = child(plot, kinds[0] ?? '');
  const k0 = kinds[0] ?? 'barChart';
  if (kinds.length > 1) return { type: 'combo', kinds };
  if (k0 === 'pieChart' || k0 === 'pie3DChart' || k0 === 'ofPieChart') return { type: 'pie', kinds };
  if (k0 === 'doughnutChart') return { type: 'doughnut', kinds };
  if (k0 === 'scatterChart') return { type: 'scatter', kinds };
  if (k0.startsWith('area')) return { type: attr(child(first, 'grouping'), 'val')?.startsWith('stacked') ? 'stackedArea' : 'area', kinds };
  if (k0.startsWith('line')) {
    const noMarkers = findAll(first!, 'symbol').every((s) => attr(s, 'val') === 'none');
    return { type: noMarkers ? 'line' : 'lineMarkers', kinds };
  }
  const dir = attr(child(first, 'barDir'), 'val');
  const stacked = attr(child(first, 'grouping'), 'val')?.startsWith('stacked') || attr(child(first, 'grouping'), 'val') === 'percentStacked';
  return { type: dir === 'bar' ? (stacked ? 'stackedBar' : 'bar') : stacked ? 'stackedColumn' : 'column', kinds };
}

function refFormula(n: XNode | undefined): string | undefined {
  const f = n ? findAll(n, 'f')[0] : undefined;
  return f ? textOf(f).trim() : undefined;
}

function parseChart(xml: string): Omit<ChartSpec, 'id' | 'anchor'> | null {
  const doc = parseXml(xml);
  const chart = findAll(doc, 'chart').find((n) => child(n, 'plotArea'));
  if (!chart) return null;
  const plot = child(chart, 'plotArea')!;
  const { type } = chartTypeFrom(plot);
  const series: ChartSeries[] = [];
  for (const group of plot.children.filter((c) => local(c).endsWith('Chart'))) {
    const gk = local(group);
    for (const s of children(group, 'ser')) {
      const tx = child(s, 'tx');
      const nameRef = refFormula(tx);
      const nameV = tx && !nameRef ? textOf(findAll(tx, 'v')[0]) : undefined;
      const color = attr(findAll(child(s, 'spPr') ?? { name: '', attrs: {}, children: [], text: '' }, 'srgbClr')[0], 'val');
      const se: ChartSeries = {
        values: refFormula(child(s, 'val') ?? child(s, 'yVal')) ?? '',
        categories: refFormula(child(s, 'cat')),
        xValues: refFormula(child(s, 'xVal')),
        nameRef,
        name: nameV || undefined,
        color: color ? '#' + color : undefined,
      };
      if (type === 'combo') se.seriesType = gk.startsWith('line') ? 'line' : gk.startsWith('area') ? 'area' : 'column';
      const axIds = children(group, 'axId').map((a) => attr(a, 'val'));
      if (type === 'combo' && axIds.length && group !== plot.children.find((c) => local(c).endsWith('Chart'))) {
        const firstAx = children(plot.children.find((c) => local(c).endsWith('Chart')), 'axId').map((a) => attr(a, 'val'));
        if (axIds[0] !== firstAx[0]) se.secondaryAxis = true;
      }
      if (se.values) series.push(se);
    }
  }
  const titleNode = child(chart, 'title');
  const title = titleNode ? findAll(titleNode, 't').map(textOf).join('') || undefined : undefined;
  const legendPos = attr(child(child(chart, 'legend'), 'legendPos'), 'val');
  const legendMap: Record<string, ChartSpec['legend']> = { r: 'right', l: 'left', t: 'top', b: 'bottom', tr: 'right' };
  const valAx = child(plot, 'valAx');
  const axTitle = (ax: XNode | undefined) => {
    const t = child(ax, 'title');
    return t ? findAll(t, 't').map(textOf).join('') || undefined : undefined;
  };
  const catAx = child(plot, 'catAx') ?? children(plot, 'valAx')[1];
  return {
    type,
    title: title ?? (series.length === 1 && attr(child(chart, 'autoTitleDeleted'), 'val') !== '1' ? undefined : undefined),
    showTitle: !!titleNode || (attr(child(chart, 'autoTitleDeleted'), 'val') === '0' && series.length === 1),
    legend: child(chart, 'legend') ? legendMap[legendPos ?? 'r'] ?? 'right' : 'none',
    dataLabels: findAll(plot, 'showVal').some((n) => attr(n, 'val') === '1'),
    gridlines: !!findAll(plot, 'majorGridlines').length,
    xTitle: axTitle(catAx),
    yTitle: axTitle(valAx),
    style: 1,
    series,
  };
}

export async function readXlsxFull(data: ArrayBuffer): Promise<Workbook> {
  const wb = await readXlsx(data);
  let zip: JSZip;
  try {
    zip = await JSZip.loadAsync(data);
  } catch {
    return wb;
  }
  const parts = await sheetParts(zip);
  // metadata first (lossless for our own files)
  let meta: { sheets: MetaSheet[] } | null = null;
  for (const f of Object.keys(zip.files).filter((p) => /^customXml\/item\d+\.xml$/.test(p))) {
    const x = await zip.file(f)!.async('string');
    if (!x.includes(META_NS)) continue;
    const body = parseXml(x);
    const root = findAll(body, 'myexcel')[0];
    try {
      meta = JSON.parse(unb64(textOf(root)));
    } catch {
      meta = null;
    }
  }
  for (const s of wb.sheets) {
    const m = meta?.sheets.find((x) => x.name === s.name);
    if (m?.charts) {
      s.charts = m.charts.map((c) => ({ ...c, id: newId('ch') }));
      continue;
    }
    // native charts
    const part = parts.find((p) => p.name === s.name);
    if (!part) continue;
    const relsXml = await zip.file(part.relsPath)?.async('string');
    if (!relsXml) continue;
    const drawRel = findAll(parseXml(relsXml), 'Relationship').find((r) => r.attrs.Type === REL_DRAWING);
    if (!drawRel) continue;
    const dPath = resolveTarget(part.path, drawRel.attrs.Target);
    const dXml = await zip.file(dPath)?.async('string');
    if (!dXml) continue;
    const dRels = findAll(parseXml((await zip.file(relsPathFor(dPath))?.async('string')) ?? ''), 'Relationship');
    const { rows, cols } = sheetLayout(s);
    for (const anchor of [...findAll(parseXml(dXml), 'twoCellAnchor'), ...findAll(parseXml(dXml), 'oneCellAnchor')]) {
      const chartEl = findAll(anchor, 'chart')[0];
      if (!chartEl) continue;
      const rid = attr(chartEl, 'id');
      const rel = dRels.find((r) => r.attrs.Id === rid);
      if (!rel) continue;
      const cXml = await zip.file(resolveTarget(dPath, rel.attrs.Target))?.async('string');
      if (!cXml) continue;
      const spec = parseChart(cXml);
      if (!spec || !spec.series.length) continue;
      const pos = (n: XNode | undefined) => {
        const c = +(textOf(child(n, 'col')) || 0);
        const r = +(textOf(child(n, 'row')) || 0);
        return { x: cols.offset(c) + +(textOf(child(n, 'colOff')) || 0) / EMU, y: rows.offset(r) + +(textOf(child(n, 'rowOff')) || 0) / EMU };
      };
      const from = pos(child(anchor, 'from'));
      let to = child(anchor, 'to') ? pos(child(anchor, 'to')) : null;
      if (!to) {
        const ext = child(anchor, 'ext');
        to = { x: from.x + +(attr(ext, 'cx') ?? 4572000) / EMU, y: from.y + +(attr(ext, 'cy') ?? 2743200) / EMU };
      }
      s.charts.push({ ...spec, id: newId('ch'), anchor: { x: from.x, y: from.y, w: Math.max(80, to.x - from.x), h: Math.max(60, to.y - from.y) } });
    }
  }
  if (meta) {
    for (const s of wb.sheets) {
      const m = meta.sheets.find((x) => x.name === s.name);
      if (!m) continue;
      if (m.sparklines) s.sparklines = m.sparklines;
      if (m.scenarios) s.scenarios = m.scenarios;
      if (m.pivots)
        s.pivots = m.pivots.map(({ sourceSheet, ...p }) => ({ ...p, sourceSheetId: wb.sheetByName(sourceSheet)?.id ?? s.id }));
    }
  }
  // names ExcelJS dropped (constants / formulas / scoped)
  const wbXml = parseXml((await zip.file('xl/workbook.xml')?.async('string')) ?? '');
  for (const dn of findAll(wbXml, 'definedName')) {
    const name = dn.attrs.name;
    if (!name || name.startsWith('_xlnm')) continue;
    const ls = dn.attrs.localSheetId;
    const scope = ls !== undefined ? wb.sheets[+ls]?.id : undefined;
    const ref = textOf(dn).trim();
    const existing = wb.names.find((n) => n.name === name);
    if (existing) {
      existing.scope = scope;
      existing.ref = ref;
      if (dn.attrs.comment) existing.comment = dn.attrs.comment;
    } else wb.names.push({ name, ref, scope, comment: dn.attrs.comment });
  }
  void colToName;
  return wb;
}
