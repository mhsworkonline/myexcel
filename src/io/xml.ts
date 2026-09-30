// Minimal XML parser/serializer helpers (enough for OOXML/ODF parts; no DTD support).

export interface XNode {
  name: string; // qualified name, e.g. "c:ser"
  attrs: Record<string, string>;
  children: XNode[];
  text: string;
}

const ENT: Record<string, string> = { lt: '<', gt: '>', amp: '&', quot: '"', apos: "'" };

export function decodeEntities(s: string): string {
  return s.replace(/&(#x[0-9a-f]+|#\d+|lt|gt|amp|quot|apos);/gi, (_, e: string) => {
    if (e[0] === '#') return String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10));
    return ENT[e] ?? _;
  });
}

export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

export function parseXml(src: string): XNode {
  const root: XNode = { name: '#root', attrs: {}, children: [], text: '' };
  const stack: XNode[] = [root];
  const re = /<!\[CDATA\[([\s\S]*?)\]\]>|<!--[\s\S]*?-->|<\?[\s\S]*?\?>|<!DOCTYPE[^>]*>|<\/([^\s>]+)\s*>|<([^\s/>]+)((?:\s+[^\s=]+\s*=\s*(?:"[^"]*"|'[^']*'))*)\s*(\/?)>|([^<]+)/g;
  let m: RegExpExecArray | null;
  while ((m = re.exec(src))) {
    const top = stack[stack.length - 1];
    if (m[1] !== undefined) top.text += m[1];
    else if (m[2]) {
      if (stack.length > 1) stack.pop();
    } else if (m[3]) {
      const attrs: Record<string, string> = {};
      const ar = /([^\s=]+)\s*=\s*("([^"]*)"|'([^']*)')/g;
      let a: RegExpExecArray | null;
      while ((a = ar.exec(m[4] ?? ''))) attrs[a[1]] = decodeEntities(a[3] ?? a[4] ?? '');
      const node: XNode = { name: m[3], attrs, children: [], text: '' };
      top.children.push(node);
      if (!m[5]) stack.push(node);
    } else if (m[6] !== undefined) top.text += decodeEntities(m[6]);
  }
  return root;
}

/** Local name without namespace prefix. */
export function local(n: XNode): string {
  const i = n.name.indexOf(':');
  return i >= 0 ? n.name.slice(i + 1) : n.name;
}

export function child(n: XNode | undefined, name: string): XNode | undefined {
  return n?.children.find((c) => local(c) === name);
}

export function children(n: XNode | undefined, name: string): XNode[] {
  return n ? n.children.filter((c) => local(c) === name) : [];
}

/** Depth-first search for all descendants with a local name. */
export function findAll(n: XNode, name: string, out: XNode[] = []): XNode[] {
  for (const c of n.children) {
    if (local(c) === name) out.push(c);
    findAll(c, name, out);
  }
  return out;
}

export function textOf(n: XNode | undefined): string {
  if (!n) return '';
  return n.text + n.children.map(textOf).join('');
}

export function attr(n: XNode | undefined, name: string): string | undefined {
  if (!n) return undefined;
  if (name in n.attrs) return n.attrs[name];
  for (const k of Object.keys(n.attrs)) if (k.endsWith(':' + name)) return n.attrs[k];
  return undefined;
}
