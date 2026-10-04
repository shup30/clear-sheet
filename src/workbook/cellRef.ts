/** Cell reference utilities: A1 notation, parsing, ranges. */

export function colToLetter(n: number): string {
  let s = '';
  for (n++; n; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
}

export function letterToCol(letters: string): number {
  let n = 0;
  for (const ch of letters.toUpperCase()) {
    const v = ch.charCodeAt(0) - 64;
    if (v < 1 || v > 26) return -1;
    n = n * 26 + v;
  }
  return n - 1;
}

export function toAddress(r: number, c: number): string {
  return `${colToLetter(c)}${r + 1}`;
}

export interface ParsedRef { r: number; c: number }

export function parseAddress(a1: string): ParsedRef | null {
  const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(a1.trim());
  if (!m) return null;
  const c = letterToCol(m[1]);
  const r = parseInt(m[2], 10) - 1;
  if (c < 0 || c > 16383 || r < 0 || r > 1048575) return null;
  return { r, c };
}

export interface SheetRef extends ParsedRef { sheet?: string }

const SHEET_REF_RE = /^(?:'((?:[^']|'')+)'|([A-Za-z0-9_.]+))?!(\$?[A-Za-z]{1,3}\$?[0-9]{1,7})$/;

export function parseSheetRef(token: string): SheetRef | null {
  const m = SHEET_REF_RE.exec(token.trim());
  if (!m) {
    const p = parseAddress(token.trim());
    return p ? { ...p } : null;
  }
  const sheet = (m[1] ?? m[2]).replace(/''/g, "'");
  const p = parseAddress(m[3]);
  return p ? { ...p, sheet } : null;
}

export function quoteSheet(name: string): string {
  return /[^A-Za-z0-9_.]/.test(name) ? `'${name.replace(/'/g, "''")}'` : name;
}

export interface Range { r1: number; c1: number; r2: number; c2: number; sheet?: string }

export function parseRange(token: string): Range | null {
  const excl = token.indexOf('!');
  let sheet: string | undefined;
  let body = token;
  if (excl >= 0) {
    let s = token.slice(0, excl).trim();
    if (s.startsWith("'") && s.endsWith("'")) s = s.slice(1, -1).replace(/''/g, "'");
    sheet = s;
    body = token.slice(excl + 1);
  }
  const parts = body.split(':');
  if (parts.length === 1) {
    const p = parseAddress(parts[0]);
    return p ? { r1: p.r, c1: p.c, r2: p.r, c2: p.c, sheet } : null;
  }
  if (parts.length === 2) {
    const a = parseAddress(parts[0]);
    const b = parseAddress(parts[1]);
    if (!a || !b) return null;
    return {
      r1: Math.min(a.r, b.r), c1: Math.min(a.c, b.c),
      r2: Math.max(a.r, b.r), c2: Math.max(a.c, b.c), sheet,
    };
  }
  return null;
}

export function cycleRefAbsolute(expr: string, pos: number): { nextExpr: string; nextPos: number } {
  const regex = /"(?:[^"]|"")*"|(\$?[A-Za-z]{1,3})(\$?[0-9]{1,7})/g;
  let match: RegExpExecArray | null;
  while ((match = regex.exec(expr)) !== null) {
    if (match[0].startsWith('"')) continue;
    const start = match.index;
    const end = start + match[0].length;
    if (pos >= start && pos <= end) {
      const colPart = match[1];
      const rowPart = match[2];
      const hasColDollar = colPart.startsWith('$');
      const hasRowDollar = rowPart.startsWith('$');
      const pureCol = colPart.replace(/\$/g, '');
      const pureRow = rowPart.replace(/\$/g, '');

      let nextToken: string;
      if (!hasColDollar && !hasRowDollar) {
        nextToken = `$${pureCol}$${pureRow}`;
      } else if (hasColDollar && hasRowDollar) {
        nextToken = `${pureCol}$${pureRow}`;
      } else if (!hasColDollar && hasRowDollar) {
        nextToken = `$${pureCol}${pureRow}`;
      } else {
        nextToken = `${pureCol}${pureRow}`;
      }

      const nextExpr = expr.slice(0, start) + nextToken + expr.slice(end);
      const delta = nextToken.length - match[0].length;
      const nextPos = Math.max(start, Math.min(nextExpr.length, pos + delta));
      return { nextExpr, nextPos };
    }
  }
  return { nextExpr: expr, nextPos: pos };
}
