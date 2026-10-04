import { parseRange, type Range } from './cellRef';

/** Minimal Excel-compatible formula engine: parse, evaluate, dependency-tracked recalc. */

export type EvalVal = number | string | boolean | null | { err: string };

export const ERR_DIV = '#DIV/0!';
export const ERR_VALUE = '#VALUE!';
export const ERR_REF = '#REF!';
export const ERR_NAME = '#NAME?';
export const ERR_NA = '#N/A';
export const ERR_CYCLE = '#CYCLE!';
export const ERR_NUM = '#NUM!';

export function isErr(v: EvalVal): v is { err: string } {
  return typeof v === 'object' && v !== null && 'err' in (v as object);
}

export function valToString(v: EvalVal): string {
  if (v === null) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'number') return String(v);
  if (typeof v === 'string') return v;
  return v.err;
}

function toNumber(v: EvalVal): number | null {
  if (typeof v === 'number') return v;
  if (typeof v === 'boolean') return v ? 1 : 0;
  if (v === null) return 0;
  if (typeof v === 'string') {
    const t = v.trim();
    if (t === '') return 0;
    const n = Number(t);
    return isFinite(n) ? n : null;
  }
  return null;
}

function toBool(v: EvalVal): boolean | null {
  if (typeof v === 'boolean') return v;
  if (typeof v === 'number') return v !== 0;
  if (v === null) return false;
  if (typeof v === 'string') {
    const t = v.trim().toLowerCase();
    if (t === 'true') return true;
    if (t === 'false') return false;
    const n = Number(t);
    if (t !== '' && isFinite(n)) return n !== 0;
    return null;
  }
  return null;
}

// ---------------- Tokenizer ----------------

type Tok =
  | { t: 'num'; v: number }
  | { t: 'str'; v: string }
  | { t: 'id'; v: string }
  | { t: 'ref'; v: string; sheet?: string }
  | { t: 'op'; v: string }
  | { t: 'lp' } | { t: 'rp' } | { t: 'comma' } | { t: 'colon' } | { t: 'bang' } | { t: 'pct' };

const REF_RE = /^\$?[A-Za-z]{1,3}\$?[0-9]{1,7}/;
const ID_RE = /^[A-Za-z_][A-Za-z0-9_.]*/;

function tokenize(src: string): Tok[] | { err: string } {
  const toks: Tok[] = [];
  let i = 0;
  const n = src.length;
  while (i < n) {
    const ch = src[i];
    if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { i++; continue; }
    if (ch === '#') {
      const m = /^#[A-Za-z0-9/?!]+/.exec(src.slice(i));
      const lit = m ? m[0] : '#';
      toks.push({ t: 'id', v: lit.toUpperCase() });
      i += lit.length; continue;
    }    if (ch === '"') {
      let j = i + 1, s = '';
      while (j < n) {
        if (src[j] === '"') {
          if (src[j + 1] === '"') { s += '"'; j += 2; continue; }
          break;
        }
        s += src[j]; j++;
      }
      if (j >= n) return { err: ERR_VALUE };
      toks.push({ t: 'str', v: s }); i = j + 1; continue;
    }
    if (ch === "'") {
      // quoted sheet name: 'My Sheet'!A1
      let j = i + 1, s = '';
      while (j < n) {
        if (src[j] === "'") {
          if (src[j + 1] === "'") { s += "'"; j += 2; continue; }
          break;
        }
        s += src[j]; j++;
      }
      if (j >= n || src[j + 1] !== '!') return { err: ERR_NAME };
      // look ahead for ref
      const rest = src.slice(j + 2);
      const m = REF_RE.exec(rest);
      if (!m) return { err: ERR_REF };
      toks.push({ t: 'ref', v: m[0], sheet: s.replace(/''/g, "'") });
      i = j + 2 + m[0].length; continue;
    }
    if (/[0-9.]/.test(ch)) {
      const m = /^[0-9]*\.?[0-9]+([eE][+-]?[0-9]+)?/.exec(src.slice(i));
      if (!m || (m[0] === '.' )) { return { err: ERR_VALUE }; }
      toks.push({ t: 'num', v: parseFloat(m[0]) }); i += m[0].length; continue;
    }
    const point=REF_RE.exec(src.slice(i));
    if(point && !/[A-Za-z0-9_.!(]/.test(src[i+point[0].length] ?? ' ')) { toks.push({t:'ref',v:point[0]});i+=point[0].length;continue; }
    if (/[A-Za-z_]/.test(ch)) {
      const m = ID_RE.exec(src.slice(i));
      if (!m) return { err: ERR_NAME };
      const word = m[0];
      let j = i + word.length;
      if (src[j] === '!') {
        const rest = src.slice(j + 1);
        const rm = REF_RE.exec(rest);
        if (!rm) return { err: ERR_REF };
        toks.push({ t: 'ref', v: rm[0], sheet: word });
        i = j + 1 + rm[0].length; continue;
      }
      // bare ref?
      if (REF_RE.test(word) && /^[A-Za-z]{1,3}[0-9]{1,7}$/.test(word.replace(/\$/g, ''))) {
        toks.push({ t: 'ref', v: word });
      } else {
        toks.push({ t: 'id', v: word.toUpperCase() });
      }
      i += word.length; continue;
    }
    if (ch === '$') {
      const m = REF_RE.exec(src.slice(i));
      if (m) { toks.push({ t: 'ref', v: m[0] }); i += m[0].length; continue; }
      return { err: ERR_REF };
    }
    if (ch === '(') { toks.push({ t: 'lp' }); i++; continue; }
    if (ch === ')') { toks.push({ t: 'rp' }); i++; continue; }
    if (ch === ',') { toks.push({ t: 'comma' }); i++; continue; }
    if (ch === ';') { toks.push({ t: 'comma' }); i++; continue; }
    if (ch === ':') { toks.push({ t: 'colon' }); i++; continue; }
    if (ch === '!') { toks.push({ t: 'bang' }); i++; continue; }
    if (ch === '%') { toks.push({ t: 'pct' }); i++; continue; }
    if (ch === '<' && src[i + 1] === '>') { toks.push({ t: 'op', v: '<>' }); i += 2; continue; }
    if (ch === '<' && src[i + 1] === '=') { toks.push({ t: 'op', v: '<=' }); i += 2; continue; }
    if (ch === '>' && src[i + 1] === '=') { toks.push({ t: 'op', v: '>=' }); i += 2; continue; }
    if ('+-*/^&=<>'.includes(ch)) { toks.push({ t: 'op', v: ch }); i++; continue; }
    return { err: ERR_NAME };
  }
  return toks;
}

// ---------------- AST ----------------

export type Node =
  | { k: 'num'; v: number } | { k: 'str'; v: string } | { k: 'bool'; v: boolean }
  | { k: 'err'; v: string }
  | { k: 'ref'; sheet?: string; r: number; c: number }
  | { k: 'range'; sheet?: string; r1: number; c1: number; r2: number; c2: number }
  | { k: 'call'; name: string; args: Node[] }
  | { k: 'un'; op: string; a: Node }
  | { k: 'bin'; op: string; a: Node; b: Node };

export function parseFormula(src: string): Node | { err: string } {
  const t = tokenize(src.length > 8192 ? src.slice(0, 8192) : src);
  if (!Array.isArray(t)) return t;
  const toks = t;
  let pos = 0;
  const peek = (): Tok | undefined => toks[pos];
  const next = (): Tok | undefined => toks[pos++];

  function parseExpr(): Node | { err: string } { return parseCmp(); }
  function parseCmp(): Node | { err: string } {
    let a = parseConcat();
    if (typeof a === 'object' && 'err' in a) return a;
    for (;;) {
      const p = peek();
      if (p?.t === 'op' && ['=', '<', '>', '<=', '>=', '<>'].includes(p.v)) {
        next();
        const b = parseConcat();
        if (typeof b === 'object' && 'err' in b) return b;
        a = { k: 'bin', op: p.v, a: a as Node, b: b as Node };
      } else return a;
    }
  }
  function parseConcat(): Node | { err: string } {
    let a = parseAdd();
    if (typeof a === 'object' && 'err' in a) return a;
    for (;;) {
      const p = peek();
      if (p?.t === 'op' && p.v === '&') {
        next();
        const b = parseAdd();
        if (typeof b === 'object' && 'err' in b) return b;
        a = { k: 'bin', op: '&', a: a as Node, b: b as Node };
      } else return a;
    }
  }
  function parseAdd(): Node | { err: string } {
    let a = parseMul();
    if (typeof a === 'object' && 'err' in a) return a;
    for (;;) {
      const p = peek();
      if (p?.t === 'op' && (p.v === '+' || p.v === '-')) {
        next();
        const b = parseMul();
        if (typeof b === 'object' && 'err' in b) return b;
        a = { k: 'bin', op: p.v, a: a as Node, b: b as Node };
      } else return a;
    }
  }
  function parseMul(): Node | { err: string } {
    let a = parsePow();
    if (typeof a === 'object' && 'err' in a) return a;
    for (;;) {
      const p = peek();
      if (p?.t === 'op' && (p.v === '*' || p.v === '/')) {
        next();
        const b = parsePow();
        if (typeof b === 'object' && 'err' in b) return b;
        a = { k: 'bin', op: p.v, a: a as Node, b: b as Node };
      } else return a;
    }
  }
  function parsePow(): Node | { err: string } {
    const a = parseUnary();
    if (typeof a === 'object' && 'err' in a) return a;
    const p = peek();
    if (p?.t === 'op' && p.v === '^') {
      next();
      const b = parseUnary();
      if (typeof b === 'object' && 'err' in b) return b;
      return { k: 'bin', op: '^', a: a as Node, b: b as Node };
    }
    return a;
  }
  function parseUnary(): Node | { err: string } {
    const p = peek();
    if (p?.t === 'op' && (p.v === '-' || p.v === '+')) {
      next();
      const a = parseUnary();
      if (typeof a === 'object' && 'err' in a) return a;
      return p.v === '-' ? { k: 'un', op: '-', a: a as Node } : (a as Node);
    }
    return parsePostfix();
  }
  function parsePostfix(): Node | { err: string } {
    const a = parsePrimary();
    if (typeof a === 'object' && 'err' in a) return a;
    if (peek()?.t === 'pct') { next(); return { k: 'bin', op: '/', a: a as Node, b: { k: 'num', v: 100 } }; }
    return a;
  }
  function refToRC(v: string): { r: number; c: number } | { err: string } {
    const m = /^\$?([A-Za-z]{1,3})\$?([0-9]{1,7})$/.exec(v);
    if (!m) return { err: ERR_REF };
    let c = 0;
    for (const ch of m[1].toUpperCase()) c = c * 26 + (ch.charCodeAt(0) - 64);
    c -= 1;
    const r = parseInt(m[2], 10) - 1;
    if (c < 0 || c > 16383 || r < 0 || r > 1048575) return { err: ERR_REF };
    return { r, c };
  }
  function parsePrimary(): Node | { err: string } {
    const p = next();
    if (!p) return { err: ERR_VALUE };
    if (p.t === 'num') return { k: 'num', v: p.v };
    if (p.t === 'str') return { k: 'str', v: p.v };
    if (p.t === 'lp') {
      const e = parseExpr();
      if (typeof e === 'object' && 'err' in e) return e;
      if (peek()?.t !== 'rp') return { err: ERR_VALUE };
      next();
      return e as Node;
    }
    if (p.t === 'id') {
      const up = p.v;
      if (up === 'TRUE') return { k: 'bool', v: true };
      if (up === 'FALSE') return { k: 'bool', v: false };
      if (up.startsWith('#')) return { k: 'err', v: up };
      if (peek()?.t === 'lp') {
        next();
        const args: Node[] = [];
        if (peek()?.t !== 'rp') {
          for (;;) {
            const a = parseExpr();
            if (typeof a === 'object' && 'err' in a) return a;
            args.push(a as Node);
            const s = peek();
            if (s?.t === 'comma') { next(); continue; }
            break;
          }
        }
        if (peek()?.t !== 'rp') return { err: ERR_VALUE };
        next();
        if (args.length > 255) return { err: ERR_VALUE };
        return { k: 'call', name: up, args };
      }
      return { err: ERR_NAME };
    }
    if (p.t === 'ref') {
      const rc = refToRC(p.v);
      if ('err' in rc) return rc;
      if (peek()?.t === 'colon') {
        const q = toks[pos + 1];
        if (q?.t === 'ref' && (q.sheet ?? p.sheet) === (p.sheet ?? q.sheet)) {
          next();
          const second = next() as { t: string; v?: string };
          const rc2 = refToRC((second as { v: string }).v);
          if ('err' in rc2) return rc2;
          return {
            k: 'range', sheet: p.sheet,
            r1: Math.min(rc.r, rc2.r), c1: Math.min(rc.c, rc2.c),
            r2: Math.max(rc.r, rc2.r), c2: Math.max(rc.c, rc2.c),
          };
        }
        return { err: ERR_REF };
      }
      return { k: 'ref', sheet: p.sheet, r: rc.r, c: rc.c };
    }
    return { err: ERR_VALUE };
  }

  const e = parseExpr();
  if (typeof e === 'object' && 'err' in e) return e;
  if (pos !== toks.length) return { err: ERR_VALUE };
  return e as Node;
}

/** Collect cell refs + range rects (for dependency tracking). */
export function collectDeps(node: Node, refs: Set<string>, ranges: Range[], selfSheet: string): void {
  switch (node.k) {
    case 'ref': {
      const sh = node.sheet ?? selfSheet;
      refs.add(`${sh}\0${node.r}:${node.c}`);
      break;
    }
    case 'range':
      ranges.push({ r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2, sheet: node.sheet ?? selfSheet });
      break;
    case 'call': for (const a of node.args) collectDeps(a, refs, ranges, selfSheet); break;
    case 'un': collectDeps(node.a, refs, ranges, selfSheet); break;
    case 'bin': collectDeps(node.a, refs, ranges, selfSheet); collectDeps(node.b, refs, ranges, selfSheet); break;
  }
}

// ---------------- Engine ----------------

export interface FormulaSource {
  sheetNames(): string[];
  /** Resolve sheet name case-insensitively; `from` is the referencing sheet. */
  resolveSheet(name: string | undefined, from: string): string | null;
  rawValue(sheet: string, r: number, c: number): EvalVal;
  nonEmptyCount(sheet: string, r1: number, c1: number, r2: number, c2: number): number;
  sparseCells(sheet: string, r1: number, c1: number, r2: number, c2: number): { r: number; c: number }[];
}

interface FormulaRec {
  sheet: string; r: number; c: number;
  ast: Node;
  refs: Set<string>;
  ranges: Range[];
  value: EvalVal;
}

const key = (sheet: string, r: number, c: number) => `${sheet}\0${r}:${c}`;

export class FormulaEngine {
  private formulas = new Map<string, FormulaRec>();
  /** referenced-sheet -> dependents */
  private bySheet = new Map<string, Set<string>>();
  /** referenced cell key -> dependents (for point refs) */
  private byCell = new Map<string, Set<string>>();

  constructor(private src: FormulaSource) {}

  hasFormula(sheet: string, r: number, c: number): boolean {
    return this.formulas.has(key(sheet, r, c));
  }

  cachedValue(sheet: string, r: number, c: number): EvalVal | undefined {
    return this.formulas.get(key(sheet, r, c))?.value;
  }

  formulaCount(): number { return this.formulas.size; }

  setFormula(sheet: string, r: number, c: number, expr: string): EvalVal {
    this.removeFormula(sheet, r, c);
    const ast = parseFormula(expr);
    if (typeof ast === 'object' && 'err' in ast) {
      const rec: FormulaRec = {
        sheet, r, c,
        ast: { k: 'err', v: ast.err }, refs: new Set(), ranges: [], value: { err: (ast as { err: string }).err },
      };
      this.formulas.set(key(sheet, r, c), rec);
      this.index(key(sheet, r, c), rec);
      return rec.value;
    }
    const refs = new Set<string>();
    const ranges: Range[] = [];
    collectDeps(ast as Node, refs, ranges, sheet);
    // Excel sheet names are case-insensitive; index dependencies canonically.
    for (const ref of [...refs]) {
      const split = ref.indexOf('\0');
      const canonical = this.src.resolveSheet(ref.slice(0,split),sheet);
      if (canonical) { refs.delete(ref); refs.add(canonical + ref.slice(split)); }
    }
    for (const range of ranges) range.sheet = this.src.resolveSheet(range.sheet,sheet) ?? range.sheet;
    const rec: FormulaRec = { sheet, r, c, ast: ast as Node, refs, ranges, value: null };
    this.formulas.set(key(sheet, r, c), rec);
    this.index(key(sheet, r, c), rec);
    // Evaluation is batched by recalc/recalcAll after all edits are applied.
    return rec.value;
  }

  private index(k: string, rec: FormulaRec): void {
    const addSheet = (sh: string) => {
      let s = this.bySheet.get(sh);
      if (!s) { s = new Set(); this.bySheet.set(sh, s); }
      s.add(k);
    };
    addSheet(rec.sheet);
    for (const rk of rec.refs) {
      const sh = rk.split('\0')[0];
      addSheet(sh);
      let s = this.byCell.get(rk);
      if (!s) { s = new Set(); this.byCell.set(rk, s); }
      s.add(k);
    }
    for (const rg of rec.ranges) addSheet(rg.sheet ?? rec.sheet);
  }

  removeFormula(sheet: string, r: number, c: number): void {
    const k = key(sheet, r, c);
    const rec = this.formulas.get(k);
    if (!rec) return;
    this.formulas.delete(k);
    for (const set of this.bySheet.values()) set.delete(k);
    for (const set of this.byCell.values()) set.delete(k);
  }

  renameSheet(oldName: string, newName: string): void {
    // Re-key formulas on renamed sheet; cross-sheet refs keep old spelling (best effort).
    const move: [string, FormulaRec][] = [];
    for (const [k, rec] of this.formulas) {
      if (rec.sheet === oldName) { move.push([k, rec]); }
    }
    for (const [k] of move) this.formulas.delete(k);
    for (const [, rec] of move) {
      rec.sheet = newName;
      this.formulas.set(key(newName, rec.r, rec.c), rec);
    }
    this.rebuildIndex();
  }

  removeSheet(sheet: string): void {
    for (const [k, rec] of [...this.formulas]) {
      if (rec.sheet === sheet) this.formulas.delete(k);
    }
    this.rebuildIndex();
  }

  private rebuildIndex(): void {
    this.bySheet.clear(); this.byCell.clear();
    for (const [k, rec] of this.formulas) this.index(k, rec);
  }

  /** Find dependents affected by a change to (sheet,r,c). */
  private affectedBy(sheet: string, r: number, c: number): Set<string> {
    const out = new Set<string>();
    const ck = key(sheet, r, c);
    const direct = this.byCell.get(ck);
    if (direct) for (const d of direct) out.add(d);
    const cands = this.bySheet.get(sheet);
    if (cands) {
      for (const d of cands) {
        if (out.has(d)) continue;
        const rec = this.formulas.get(d);
        if (!rec) continue;
        for (const rg of rec.ranges) {
          const rs = rg.sheet ?? rec.sheet;
          if (rs === sheet && r >= rg.r1 && r <= rg.r2 && c >= rg.c1 && c <= rg.c2) { out.add(d); break; }
        }
      }
    }
    return out;
  }

  /**
   * Recalculate cells affected by the given changes.
   * Returns map of formula key -> new value (only formulas whose value changed OR all affected).
   */
  recalc(changed: { sheet: string; r: number; c: number }[]): Map<string, EvalVal> {
    const affected = new Set<string>();
    for (const ch of changed) if (this.hasFormula(ch.sheet,ch.r,ch.c)) affected.add(key(ch.sheet,ch.r,ch.c));
    const queue = [...changed];
    const seen = new Set<string>();
    while (queue.length) {
      const ch = queue.pop()!;
      for (const d of this.affectedBy(ch.sheet, ch.r, ch.c)) {
        if (seen.has(d)) continue;
        seen.add(d); affected.add(d);
        const rec = this.formulas.get(d);
        if (rec) queue.push({ sheet: rec.sheet, r: rec.r, c: rec.c });
      }
    }
    const updated = new Map<string, EvalVal>();
    this.evaluationCache = new Map();
    for (const k of affected) {
      const rec = this.formulas.get(k);
      if (!rec) continue;
      const v = this.evaluate(k, rec, new Set());
      rec.value = v;
      updated.set(k,v);
    }
    this.evaluationCache = undefined;
    return updated;
  }

  /** Re-evaluate every formula (e.g. after bulk load). Returns changed keys. */
  recalcAll(): Map<string, EvalVal> {
    const updated = new Map<string, EvalVal>();
    this.evaluationCache = new Map();
    for (const [k,rec] of this.formulas) {
      rec.value = this.evaluate(k,rec,new Set());
      updated.set(k,rec.value);
    }
    this.evaluationCache = undefined;
    return updated;
  }

  // ---------- evaluation ----------

  private evaluationCache?: Map<string,EvalVal>;
  private evaluate(k: string, rec: FormulaRec, stack: Set<string>): EvalVal {
    if (stack.has(k)) return {err:ERR_CYCLE};
    if (stack.size > 512) return {err:ERR_VALUE};
    if (this.evaluationCache?.has(k)) return this.evaluationCache.get(k)!;
    const next = new Set(stack); next.add(k);
    const value = this.evalNode(rec.ast,rec.sheet,next);
    this.evaluationCache?.set(k,value);
    return value;
  }

  private lookup(sheetRef: string | undefined, from: string, r: number, c: number, stack: Set<string>): EvalVal {
    const sheet = this.src.resolveSheet(sheetRef, from);
    if (!sheet) return { err: ERR_REF };
    if (r < 0 || r > 1048575 || c < 0 || c > 16383) return { err: ERR_REF };
    const k = key(sheet, r, c);
    if (stack.has(k)) return { err: ERR_CYCLE };
    const f = this.formulas.get(k);
    if (f) return this.evaluate(k,f,stack);
    return this.src.rawValue(sheet, r, c);
  }

  private rangeValues(rg: Range, from: string, stack: Set<string>): { vals: EvalVal[]; rows: number; cols: number; nonEmpty: number } {
    const sheet = this.src.resolveSheet(rg.sheet, from);
    if (!sheet) return { vals: [{ err: ERR_REF }], rows: 0, cols: 0, nonEmpty: 0 };
    const r1 = Math.max(0, rg.r1), c1 = Math.max(0, rg.c1);
    const r2 = Math.min(1048575, rg.r2), c2 = Math.min(16383, rg.c2);
    const rows = r2 - r1 + 1, cols = c2 - c1 + 1;
    const total = rows * cols;
    if (total > 2000000) return { vals: [{ err: ERR_VALUE }], rows, cols, nonEmpty: 0 };
    const vals: EvalVal[] = [];
    // Keep blanks in their positions: INDEX/MATCH/VLOOKUP require a rectangle.
    for (let r=r1;r<=r2;r++) for(let c=c1;c<=c2;c++) vals.push(this.lookup(sheet,from,r,c,stack));
    const nonEmpty = this.src.nonEmptyCount(sheet, r1, c1, r2, c2);
    return { vals, rows, cols, nonEmpty };
  }

  private evalNode(node: Node, from: string, stack: Set<string>): EvalVal {
    switch (node.k) {
      case 'num': return node.v;
      case 'str': return node.v;
      case 'bool': return node.v;
      case 'err': return { err: node.v };
      case 'ref': return this.lookup(node.sheet, from, node.r, node.c, stack);
      case 'range': {
        // standalone range: Excel returns top-left in scalar context
        const rg: Range = { r1: node.r1, c1: node.c1, r2: node.r2, c2: node.c2, sheet: node.sheet };
        const { vals } = this.rangeValues(rg, from, stack);
        return vals[0] ?? null;
      }
      case 'un': {
        const a = this.evalNode(node.a, from, stack);
        if (isErr(a)) return a;
        const n = toNumber(a);
        if (n === null) return { err: ERR_VALUE };
        return -n;
      }
      case 'bin': return this.evalBin(node.op, node.a, node.b, from, stack);
      case 'call': return this.evalCall(node.name, node.args, from, stack);
    }
  }

  private evalBin(op: string, aN: Node, bN: Node, from: string, stack: Set<string>): EvalVal {
    // Range operands in scalar context: take top-left
    const a0 = this.evalNode(aN, from, stack);
    if (isErr(a0)) return a0;
    const b0 = this.evalNode(bN, from, stack);
    if (isErr(b0)) return b0;
    if (op === '&') return valToString(a0) + valToString(b0);
    if (op === '=' || op === '<>' || op === '<' || op === '>' || op === '<=' || op === '>=') {
      return compare(op, a0, b0);
    }
    const a = toNumber(a0);
    const b = toNumber(b0);
    if (a === null || b === null) return { err: ERR_VALUE };
    switch (op) {
      case '+': return a + b;
      case '-': return a - b;
      case '*': return a * b;
      case '/': return b === 0 ? { err: ERR_DIV } : a / b;
      case '^': return Math.pow(a, b);
    }
    return { err: ERR_VALUE };
  }

  private evalCall(name: string, args: Node[], from: string, stack: Set<string>): EvalVal {
    const fn = FUNCTIONS[name];
    if (!fn) return { err: ERR_NAME };
    // Lazily evaluate args: pass thunks so IF/IFERROR short-circuit
    const lazy = args.map((a) => ({
      node: a,
      val: (): EvalVal => this.evalNode(a, from, stack),
      range: (): { vals: EvalVal[]; rows: number; cols: number; nonEmpty: number } | { err: string } => {
        if (a.k === 'range') {
          return this.rangeValues({ r1: a.r1, c1: a.c1, r2: a.r2, c2: a.c2, sheet: a.sheet }, from, stack);
        }
        if (a.k === 'ref') {
          const sh = this.src.resolveSheet(a.sheet, from);
          if (!sh) return { err: ERR_REF };
          return { vals: [this.lookup(a.sheet, from, a.r, a.c, stack)], rows: 1, cols: 1, nonEmpty: 1 };
        }
        return { vals: [this.evalNode(a, from, stack)], rows: 1, cols: 1, nonEmpty: 1 };
      },
    }));
    try {
      return fn(lazy, from, this);
    } catch {
      return { err: ERR_VALUE };
    }
  }

  /** Public helper for function impls: resolve a range arg to a 2-D grid. */
  gridOf(arg: LazyArg, from: string, stack: Set<string>): { grid: EvalVal[][]; rows: number; cols: number } | { err: string } {
    const n = arg.node;
    if (n.k === 'range') {
      const sheet = this.src.resolveSheet(n.sheet, from);
      if (!sheet) return { err: ERR_REF };
      const grid: EvalVal[][] = [];
      for (let r = n.r1; r <= n.r2; r++) {
        const row: EvalVal[] = [];
        for (let c = n.c1; c <= n.c2; c++) row.push(this.lookup(n.sheet, from, r, c, stack));
        grid.push(row);
      }
      return { grid, rows: n.r2 - n.r1 + 1, cols: n.c2 - n.c1 + 1 };
    }
    if (n.k === 'ref') {
      return { grid: [[this.lookup(n.sheet, from, n.r, n.c, stack)]], rows: 1, cols: 1 };
    }
    return { grid: [[arg.val()]], rows: 1, cols: 1 };
  }
}

export interface LazyArg {
  node: Node;
  val(): EvalVal;
  range(): { vals: EvalVal[]; rows: number; cols: number; nonEmpty: number } | { err: string };
}



function compare(op: string, a: EvalVal, b: EvalVal): boolean {
  // Excel type order: number < text < logical; null/blank special-cased
  if (a === null && b === null) return op === '=' || op === '<=' || op === '>=';
  if (a === null) return compareBlank(op, b, true);
  if (b === null) return compareBlank(op, a, false);
  const rank = (v: EvalVal) => (typeof v === 'number' ? 0 : typeof v === 'string' ? 1 : 2);
  if (rank(a) !== rank(b)) {
    const d = rank(a) - rank(b);
    switch (op) {
      case '=': return false;
      case '<>': return true;
      case '<': return d < 0;
      case '>': return d > 0;
      case '<=': return d <= 0;
      case '>=': return d >= 0;
    }
  }
  let eq: boolean; let lt: boolean;
  if (typeof a === 'string' && typeof b === 'string') {
    const al = a.toLowerCase(), bl = b.toLowerCase();
    eq = al === bl; lt = al < bl;
  } else if (typeof a === 'boolean' && typeof b === 'boolean') {
    eq = a === b; lt = !a && b;
  } else {
    eq = (a as number) === (b as number); lt = (a as number) < (b as number);
  }
  switch (op) {
    case '=': return eq;
    case '<>': return !eq;
    case '<': return lt;
    case '>': return !eq && !lt;
    case '<=': return eq || lt;
    case '>=': return !lt;
    default: return false;
  }
}

function compareBlank(op: string, other: EvalVal, blankFirst: boolean): boolean {
  if (typeof other === 'string') {
    if (other === '') return op === '=' || op === '<=' || op === '>=';
    // blank < non-empty text
    return blankFirst ? ['<', '<=', '<>'].includes(op) : ['>', '>=', '<>'].includes(op);
  }
  if (typeof other === 'number') {
    if (other === 0) return op === '=' || op === '<=' || op === '>=';
    return blankFirst ? other > 0 ? ['<', '<=', '<>'].includes(op) : ['>', '>=', '<>'].includes(op)
      : other > 0 ? ['>', '>=', '<>'].includes(op) : ['<', '<=', '<>'].includes(op);
  }
  return op === '<>';
}

function flattenivals(lazy: LazyArg[]): { vals: EvalVal[]; err?: string } {
  const out: EvalVal[] = [];
  for (const a of lazy) {
    const r = a.range();
    if ('err' in r) return { vals: [], err: r.err };
    for (const v of r.vals) { if (isErr(v)) return { vals: [], err: (v as { err: string }).err }; out.push(v); }
  }
  return { vals: out };
}

function numsOf(vals: EvalVal[]): number[] {
  const out: number[] = [];
  for (const v of vals) {
    if (typeof v === 'number') out.push(v);
    else if (typeof v === 'boolean') out.push(v ? 1 : 0);
    // text ignored in refs (but direct literals counted by caller convention)
  }
  return out;
}

type Fn = (args: LazyArg[], from: string, eng: FormulaEngine) => EvalVal;

const FUNCTIONS: Record<string, Fn> = {
  SUM(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    let s = 0;
    for (const v of f.vals) {
      if (typeof v === 'number') s += v;
      else if (typeof v === 'boolean') s += v ? 1 : 0;
      else if (v === null) continue;
      else return { err: ERR_VALUE };
    }
    return s;
  },
  AVERAGE(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    const ns = numsOf(f.vals.filter((v) => v !== null));
    if (!ns.length) return { err: ERR_DIV };
    return ns.reduce((a, b) => a + b, 0) / ns.length;
  },
  MIN(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    const ns = numsOf(f.vals);
    return ns.length ? Math.min(...ns) : 0;
  },
  MAX(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    const ns = numsOf(f.vals);
    return ns.length ? Math.max(...ns) : 0;
  },
  COUNT(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    return f.vals.filter((v) => typeof v === 'number').length;
  },
  COUNTA(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    return f.vals.filter((v) => v !== null && v !== '').length;
  },
  COUNTBLANK(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const r = args[0].range();
    if ('err' in r) return { err: r.err };
    return r.rows * r.cols - r.nonEmpty;
  },
  COUNTIF(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const r = args[0].range();
    if ('err' in r) return { err: r.err };
    const crit = args[1].val();
    if (isErr(crit)) return crit;
    return r.vals.filter((v) => matchCriteria(v, crit)).length;
  },
  SUMIF(args) {
    if (args.length < 2 || args.length > 3) return { err: ERR_VALUE };
    const r = args[0].range();
    if ('err' in r) return { err: r.err };
    const crit = args[1].val();
    if (isErr(crit)) return crit;
    let sumVals: EvalVal[];
    if (args[2]) {
      const s = args[2].range();
      if ('err' in s) return { err: s.err };
      sumVals = s.vals;
    } else sumVals = r.vals;
    let s = 0;
    for (let i = 0; i < r.vals.length; i++) {
      if (matchCriteria(r.vals[i], crit)) {
        const v = sumVals[Math.min(i, sumVals.length - 1)];
        if (isErr(v)) return v;
        const n = toNumber(v);
        if (n === null) return { err: ERR_VALUE };
        s += n;
      }
    }
    return s;
  },
  AVERAGEIF(args) {
    if (args.length < 2 || args.length > 3) return { err: ERR_VALUE };
    const r = args[0].range();
    if ('err' in r) return { err: r.err };
    const crit = args[1].val();
    if (isErr(crit)) return crit;
    let avgVals: EvalVal[];
    if (args[2]) {
      const s = args[2].range();
      if ('err' in s) return { err: s.err };
      avgVals = s.vals;
    } else avgVals = r.vals;
    const ns: number[] = [];
    for (let i = 0; i < r.vals.length; i++) {
      if (matchCriteria(r.vals[i], crit)) {
        const v = avgVals[Math.min(i, avgVals.length - 1)];
        if (isErr(v)) return v;
        if (typeof v === 'number') ns.push(v);
      }
    }
    if (!ns.length) return { err: ERR_DIV };
    return ns.reduce((a, b) => a + b, 0) / ns.length;
  },
  IF(args) {
    if (args.length < 2 || args.length > 3) return { err: ERR_VALUE };
    const c = args[0].val();
    if (isErr(c)) return c;
    const b = toBool(c);
    if (b === null) return { err: ERR_VALUE };
    if (b) return args[1].val();
    return args[2] ? args[2].val() : false;
  },
  IFERROR(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const v = args[0].val();
    return isErr(v) ? args[1].val() : v;
  },
  AND(args) {
    if (!args.length) return { err: ERR_VALUE };
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    for (const v of f.vals) {
      const b = toBool(v);
      if (b === null) return { err: ERR_VALUE };
      if (!b) return false;
    }
    return true;
  },
  OR(args) {
    if (!args.length) return { err: ERR_VALUE };
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    for (const v of f.vals) {
      const b = toBool(v);
      if (b === null) return { err: ERR_VALUE };
      if (b) return true;
    }
    return false;
  },
  NOT(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    const b = toBool(v);
    return b === null ? { err: ERR_VALUE } : !b;
  },
  ROUND(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const a = args[0].val(); const b = args[1].val();
    if (isErr(a)) return a; if (isErr(b)) return b;
    const n = toNumber(a); const d = toNumber(b);
    if (n === null || d === null) return { err: ERR_VALUE };
    const f = Math.pow(10, Math.trunc(d));
    return Math.round(n * f) / f;
  },
  ROUNDUP(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const a = args[0].val(); const b = args[1].val();
    if (isErr(a)) return a; if (isErr(b)) return b;
    const n = toNumber(a); const d = toNumber(b);
    if (n === null || d === null) return { err: ERR_VALUE };
    const f = Math.pow(10, Math.trunc(d));
    return (n >= 0 ? Math.ceil(n * f) : Math.floor(n * f)) / f;
  },
  ROUNDDOWN(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const a = args[0].val(); const b = args[1].val();
    if (isErr(a)) return a; if (isErr(b)) return b;
    const n = toNumber(a); const d = toNumber(b);
    if (n === null || d === null) return { err: ERR_VALUE };
    const f = Math.pow(10, Math.trunc(d));
    return (n >= 0 ? Math.floor(n * f) : Math.ceil(n * f)) / f;
  },
  INT(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    const n = toNumber(v);
    return n === null ? { err: ERR_VALUE } : Math.floor(n);
  },
  ABS(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    const n = toNumber(v);
    return n === null ? { err: ERR_VALUE } : Math.abs(n);
  },
  MOD(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const a = args[0].val(); const b = args[1].val();
    if (isErr(a)) return a; if (isErr(b)) return b;
    const n = toNumber(a); const d = toNumber(b);
    if (n === null || d === null) return { err: ERR_VALUE };
    if (d === 0) return { err: ERR_DIV };
    return n % d;
  },
  POWER(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const a = args[0].val(); const b = args[1].val();
    if (isErr(a)) return a; if (isErr(b)) return b;
    const n = toNumber(a); const d = toNumber(b);
    if (n === null || d === null) return { err: ERR_VALUE };
    return Math.pow(n, d);
  },
  SQRT(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    const n = toNumber(v);
    if (n === null) return { err: ERR_VALUE };
    if (n < 0) return { err: ERR_VALUE };
    return Math.sqrt(n);
  },
  CONCAT(args) {
    const parts: string[] = [];
    for (const a of args) {
      const r = a.range();
      if ('err' in r) return { err: r.err };
      for (const v of r.vals) {
        if (isErr(v)) return v;
        parts.push(valToString(v));
      }
    }
    return parts.join('');
  },
  CONCATENATE(args) { return FUNCTIONS.CONCAT(args, '', args as never as FormulaEngine); },
  TEXTJOIN(args) {
    if (args.length < 2) return { err: ERR_VALUE };
    const delim = args[0].val();
    if (isErr(delim)) return delim;
    const ignoreEmpty = toBool(args[1].val());
    const parts: string[] = [];
    for (const a of args.slice(2)) {
      const r = a.range();
      if ('err' in r) return { err: r.err };
      for (const v of r.vals) {
        if (isErr(v)) return v;
        const s = valToString(v);
        if (s === '' && ignoreEmpty) continue;
        parts.push(s);
      }
    }
    return parts.join(valToString(delim));
  },
  LEFT(args) {
    if (args.length < 1 || args.length > 2) return { err: ERR_VALUE };
    const s = args[0].val();
    if (isErr(s)) return s;
    let n = 1;
    if (args[1]) {
      const nv = args[1].val();
      if (isErr(nv)) return nv;
      const nn = toNumber(nv);
      if (nn === null) return { err: ERR_VALUE };
      n = Math.trunc(nn);
    }
    return valToString(s).slice(0, Math.max(0, n));
  },
  RIGHT(args) {
    if (args.length < 1 || args.length > 2) return { err: ERR_VALUE };
    const s = args[0].val();
    if (isErr(s)) return s;
    let n = 1;
    if (args[1]) {
      const nv = args[1].val();
      if (isErr(nv)) return nv;
      const nn = toNumber(nv);
      if (nn === null) return { err: ERR_VALUE };
      n = Math.trunc(nn);
    }
    const str = valToString(s);
    return str.slice(Math.max(0, str.length - n));
  },
  MID(args) {
    if (args.length !== 3) return { err: ERR_VALUE };
    const s = args[0].val(); const st = args[1].val(); const ln = args[2].val();
    if (isErr(s)) return s; if (isErr(st)) return st; if (isErr(ln)) return ln;
    const start = toNumber(st); const len = toNumber(ln);
    if (start === null || len === null) return { err: ERR_VALUE };
    return valToString(s).slice(Math.max(0, Math.trunc(start) - 1), Math.max(0, Math.trunc(start) - 1 + Math.trunc(len)));
  },
  LEN(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    return valToString(v).length;
  },
  UPPER(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    return valToString(v).toUpperCase();
  },
  LOWER(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    return valToString(v).toLowerCase();
  },
  TRIM(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    return valToString(v).trim().replace(/\s+/g, ' ');
  },
  VALUE(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    const n = toNumber(typeof v === 'string' ? v.trim() : v);
    return n === null ? { err: ERR_VALUE } : n;
  },
  TODAY() {
    const now = new Date();
    const utc = Date.UTC(now.getFullYear(), now.getMonth(), now.getDate());
    return utc / 86400000 + 25569;
  },
  NOW() {
    return Date.now() / 86400000 + 25569;
  },
  DATE(args) {
    if (args.length !== 3) return { err: ERR_VALUE };
    const y = toNumber(args[0].val()); const m = toNumber(args[1].val()); const d = toNumber(args[2].val());
    if (y === null || m === null || d === null) return { err: ERR_VALUE };
    return Date.UTC(Math.trunc(y), Math.trunc(m) - 1, Math.trunc(d)) / 86400000 + 25569;
  },
  YEAR(args) {
    const v = args[0]?.val();
    if (v === undefined) return { err: ERR_VALUE };
    if (isErr(v)) return v;
    const n = toNumber(v);
    if (n === null) return { err: ERR_VALUE };
    return new Date(Math.round((n - 25569) * 86400000)).getUTCFullYear();
  },
  MONTH(args) {
    const v = args[0]?.val();
    if (v === undefined) return { err: ERR_VALUE };
    if (isErr(v)) return v;
    const n = toNumber(v);
    if (n === null) return { err: ERR_VALUE };
    return new Date(Math.round((n - 25569) * 86400000)).getUTCMonth() + 1;
  },
  DAY(args) {
    const v = args[0]?.val();
    if (v === undefined) return { err: ERR_VALUE };
    if (isErr(v)) return v;
    const n = toNumber(v);
    if (n === null) return { err: ERR_VALUE };
    return new Date(Math.round((n - 25569) * 86400000)).getUTCDate();
  },
  DAYS(args) {
    if (args.length !== 2) return { err: ERR_VALUE };
    const a = toNumber(args[0].val()); const b = toNumber(args[1].val());
    if (a === null || b === null) return { err: ERR_VALUE };
    return Math.trunc(a - b);
  },
  VLOOKUP(args, from, eng) {
    if (args.length < 3 || args.length > 4) return { err: ERR_VALUE };
    const look = args[0].val();
    if (isErr(look)) return look;
    const g = eng.gridOf(args[1], from, new Set());
    if ('err' in g) return { err: g.err };
    const col = toNumber(args[2].val());
    if (col === null) return { err: ERR_VALUE };
    const ci = Math.trunc(col);
    if (ci < 1 || ci > g.cols) return { err: ERR_REF };
    const exact = args[3] ? toBool(args[3].val()) : true;
    let best = -1;
    for (let r = 0; r < g.rows; r++) {
      const cell = g.grid[r][0];
      if (isErr(cell)) continue;
      if (compare('=', cell, look)) { best = r; break; }
      if (!exact && typeof cell === 'number' && typeof look === 'number' && cell <= look) best = r;
    }
    if (best < 0) return { err: ERR_NA };
    return g.grid[best][ci - 1];
  },
  XLOOKUP(args, from, eng) {
    if (args.length < 3 || args.length > 6) return { err: ERR_VALUE };
    const look = args[0].val();
    if (isErr(look)) return look;
    const la = eng.gridOf(args[1], from, new Set());
    if ('err' in la) return { err: la.err };
    const ra = eng.gridOf(args[2], from, new Set());
    if ('err' in ra) return { err: ra.err };
    const notFound = args[3] ? args[3].val() : { err: ERR_NA };
    const searchMode = args[5] ? toNumber(args[5].val()) : 1;
    const flatL: EvalVal[] = la.grid.flat();
    const order = flatL.map((_, i) => i);
    if (searchMode === -1) order.reverse();
    for (const i of order) {
      const cell = flatL[i];
      if (isErr(cell)) continue;
      if (compare('=', cell, look)) {
        const flatR: EvalVal[] = ra.grid.flat();
        return flatR[Math.min(i, flatR.length - 1)] ?? notFound;
      }
    }
    return notFound;
  },
  INDEX(args, from, eng) {
    if (args.length < 2 || args.length > 3) return { err: ERR_VALUE };
    const g = eng.gridOf(args[0], from, new Set());
    if ('err' in g) return { err: g.err };
    const r = toNumber(args[1].val());
    if (r === null) return { err: ERR_VALUE };
    const c = args[2] ? toNumber(args[2].val()) : 1;
    if (c === null) return { err: ERR_VALUE };
    const ri = Math.trunc(r) - 1, ci = Math.trunc(c) - 1;
    if (ri < 0 || ri >= g.rows || ci < 0 || ci >= g.cols) return { err: ERR_REF };
    return g.grid[ri][ci];
  },
  MATCH(args, from, eng) {
    if (args.length < 2 || args.length > 3) return { err: ERR_VALUE };
    const look = args[0].val();
    if (isErr(look)) return look;
    const g = eng.gridOf(args[1], from, new Set());
    if ('err' in g) return { err: g.err };
    const flat = g.grid.flat();
    for (let i = 0; i < flat.length; i++) {
      if (isErr(flat[i])) continue;
      if (compare('=', flat[i], look)) return i + 1;
    }
    return { err: ERR_NA };
  },
  ROW(args) {
    if (!args.length) return { err: ERR_VALUE };
    const n = args[0].node;
    if (n.k === 'ref') return n.r + 1;
    if (n.k === 'range') return n.r1 + 1;
    return { err: ERR_VALUE };
  },
  COLUMN(args) {
    if (!args.length) return { err: ERR_VALUE };
    const n = args[0].node;
    if (n.k === 'ref') return n.c + 1;
    if (n.k === 'range') return n.c1 + 1;
    return { err: ERR_VALUE };
  },
  ISBLANK(args) {
    if (args.length !== 1) return { err: ERR_VALUE };
    const v = args[0].val();
    if (isErr(v)) return v;
    return v === null || v === '';
  },
  ISNUMBER(args) {
    const v = args[0]?.val();
    if (v === undefined) return { err: ERR_VALUE };
    if (isErr(v)) return v;
    return typeof v === 'number';
  },
  ISTEXT(args) {
    const v = args[0]?.val();
    if (v === undefined) return { err: ERR_VALUE };
    if (isErr(v)) return v;
    return typeof v === 'string';
  },
  COUNTIFS(args) {
    if (args.length < 2 || args.length % 2 !== 0) return { err: ERR_VALUE };
    const r0 = args[0].range();
    if ('err' in r0) return { err: r0.err };
    const n = r0.vals.length;
    const critRanges: EvalVal[][] = [r0.vals];
    const critVals: EvalVal[] = [];

    const c0 = args[1].val();
    if (isErr(c0)) return c0;
    critVals.push(c0);

    for (let k = 1; k < args.length / 2; k++) {
      const r = args[2 * k].range();
      if ('err' in r) return { err: r.err };
      if (r.vals.length !== n) return { err: ERR_VALUE };
      critRanges.push(r.vals);
      const c = args[2 * k + 1].val();
      if (isErr(c)) return c;
      critVals.push(c);
    }

    let count = 0;
    for (let i = 0; i < n; i++) {
      let match = true;
      for (let k = 0; k < critRanges.length; k++) {
        if (!matchCriteria(critRanges[k][i], critVals[k])) {
          match = false;
          break;
        }
      }
      if (match) count++;
    }
    return count;
  },
  SUMIFS(args) {
    if (args.length < 3 || args.length % 2 === 0) return { err: ERR_VALUE };
    const sumR = args[0].range();
    if ('err' in sumR) return { err: sumR.err };
    const n = sumR.vals.length;
    const critRanges: EvalVal[][] = [];
    const critVals: EvalVal[] = [];

    for (let k = 0; k < (args.length - 1) / 2; k++) {
      const r = args[1 + 2 * k].range();
      if ('err' in r) return { err: r.err };
      if (r.vals.length !== n) return { err: ERR_VALUE };
      critRanges.push(r.vals);
      const c = args[2 + 2 * k].val();
      if (isErr(c)) return c;
      critVals.push(c);
    }

    let sum = 0;
    for (let i = 0; i < n; i++) {
      let match = true;
      for (let k = 0; k < critRanges.length; k++) {
        if (!matchCriteria(critRanges[k][i], critVals[k])) {
          match = false;
          break;
        }
      }
      if (match) {
        const v = sumR.vals[i];
        if (isErr(v)) return v;
        const num = toNumber(v);
        if (num === null) return { err: ERR_VALUE };
        sum += num;
      }
    }
    return sum;
  },
  AVERAGEIFS(args) {
    if (args.length < 3 || args.length % 2 === 0) return { err: ERR_VALUE };
    const avgR = args[0].range();
    if ('err' in avgR) return { err: avgR.err };
    const n = avgR.vals.length;
    const critRanges: EvalVal[][] = [];
    const critVals: EvalVal[] = [];

    for (let k = 0; k < (args.length - 1) / 2; k++) {
      const r = args[1 + 2 * k].range();
      if ('err' in r) return { err: r.err };
      if (r.vals.length !== n) return { err: ERR_VALUE };
      critRanges.push(r.vals);
      const c = args[2 + 2 * k].val();
      if (isErr(c)) return c;
      critVals.push(c);
    }

    const matchedNums: number[] = [];
    for (let i = 0; i < n; i++) {
      let match = true;
      for (let k = 0; k < critRanges.length; k++) {
        if (!matchCriteria(critRanges[k][i], critVals[k])) {
          match = false;
          break;
        }
      }
      if (match) {
        const v = avgR.vals[i];
        if (isErr(v)) return v;
        if (typeof v === 'number') matchedNums.push(v);
      }
    }
    if (!matchedNums.length) return { err: ERR_DIV };
    return matchedNums.reduce((a, b) => a + b, 0) / matchedNums.length;
  },
  IFS(args) {
    if (args.length < 2 || args.length % 2 !== 0) return { err: ERR_VALUE };
    for (let i = 0; i < args.length; i += 2) {
      const cond = args[i].val();
      if (isErr(cond)) return cond;
      const b = toBool(cond);
      if (b === null) return { err: ERR_VALUE };
      if (b) return args[i + 1].val();
    }
    return { err: ERR_NA };
  },
  SWITCH(args) {
    if (args.length < 3) return { err: ERR_VALUE };
    const target = args[0].val();
    if (isErr(target)) return target;
    const hasDefault = (args.length - 1) % 2 === 1;
    const numPairs = hasDefault ? args.length - 2 : args.length - 1;
    for (let i = 1; i <= numPairs; i += 2) {
      const caseVal = args[i].val();
      if (isErr(caseVal)) return caseVal;
      if (compare('=', target, caseVal)) {
        return args[i + 1].val();
      }
    }
    if (hasDefault) {
      return args[args.length - 1].val();
    }
    return { err: ERR_NA };
  },
  MEDIAN(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    const ns = numsOf(f.vals.filter(v => v !== null));
    if (!ns.length) return { err: ERR_NUM };
    ns.sort((a, b) => a - b);
    const mid = Math.floor(ns.length / 2);
    return ns.length % 2 !== 0 ? ns[mid] : (ns[mid - 1] + ns[mid]) / 2;
  },
  STDEV(args) {
    return FUNCTIONS['STDEV.S'](args, '', args as never as FormulaEngine);
  },
  'STDEV.S'(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    const ns = numsOf(f.vals.filter(v => v !== null));
    if (ns.length < 2) return { err: ERR_DIV };
    const mean = ns.reduce((a, b) => a + b, 0) / ns.length;
    const ss = ns.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0);
    return Math.sqrt(ss / (ns.length - 1));
  },
  'STDEV.P'(args) {
    const f = flattenivals(args);
    if (f.err) return { err: f.err };
    const ns = numsOf(f.vals.filter(v => v !== null));
    if (!ns.length) return { err: ERR_DIV };
    const mean = ns.reduce((a, b) => a + b, 0) / ns.length;
    const ss = ns.reduce((acc, v) => acc + Math.pow(v - mean, 2), 0);
    return Math.sqrt(ss / ns.length);
  },
  PMT(args) {
    if (args.length < 3 || args.length > 5) return { err: ERR_VALUE };
    const r = toNumber(args[0].val());
    const n = toNumber(args[1].val());
    const p = toNumber(args[2].val());
    if (r === null || n === null || p === null) return { err: ERR_VALUE };
    let fv = 0;
    if (args[3]) {
      const fvNum = toNumber(args[3].val());
      if (fvNum === null) return { err: ERR_VALUE };
      fv = fvNum;
    }
    let type = 0;
    if (args[4]) {
      const typeNum = toNumber(args[4].val());
      if (typeNum === null) return { err: ERR_VALUE };
      type = typeNum !== 0 ? 1 : 0;
    }
    if (r === 0) return -(p + fv) / n;
    const pvif = Math.pow(1 + r, n);
    let pmt = (r / (pvif - 1)) * -(p * pvif + fv);
    if (type === 1) pmt /= (1 + r);
    return pmt;
  },
  SEQUENCE(args) {
    if (!args.length) return { err: ERR_VALUE };
    const r = toNumber(args[0].val());
    if (r === null || r < 1) return { err: ERR_VALUE };
    const start = args[2] ? toNumber(args[2].val()) : 1;
    if (start === null) return { err: ERR_VALUE };
    return start;
  },
  SORT(args, from, eng) {
    if (!args.length) return { err: ERR_VALUE };
    const g = eng.gridOf(args[0], from, new Set());
    if ('err' in g) return { err: g.err };
    const colIdx = args[1] ? toNumber(args[1].val()) : 1;
    if (colIdx === null) return { err: ERR_VALUE };
    const ci = Math.trunc(colIdx) - 1;
    if (ci < 0 || ci >= g.cols) return { err: ERR_REF };
    const order = args[2] ? toNumber(args[2].val()) : 1;
    const desc = order === -1;
    const rows = [...g.grid];
    rows.sort((a, b) => {
      const cmp = compare('<', a[ci], b[ci]);
      const eq = compare('=', a[ci], b[ci]);
      if (eq) return 0;
      return (cmp ? -1 : 1) * (desc ? -1 : 1);
    });
    return rows[0]?.[0] ?? null;
  },
  UNIQUE(args, from, eng) {
    if (!args.length) return { err: ERR_VALUE };
    const g = eng.gridOf(args[0], from, new Set());
    if ('err' in g) return { err: g.err };
    return g.grid[0]?.[0] ?? null;
  },
};

function matchCriteria(v: EvalVal, crit: EvalVal): boolean {
  if (isErr(v)) return false;
  if (typeof crit === 'string') {
    const t = crit.trim();
    const m = /^(<=|>=|<>|=|<|>)(.*)$/.exec(t);
    if (m) {
      const rhsRaw = m[2];
      const num = Number(rhsRaw);
      const rhs: EvalVal = rhsRaw !== '' && isFinite(num) ? num : rhsRaw;
      return compare(m[1], v, rhs);
    }
    // wildcard support: * and ?
    if (/[*?]/.test(t)) {
      if (typeof v !== 'string') return false;
      const re = new RegExp('^' + t.split('').map((ch) =>
        ch === '*' ? '.*' : ch === '?' ? '.' : ch.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')).join('') + '$', 'i');
      return re.test(v);
    }
    if (typeof v === 'string') return v.toLowerCase() === t.toLowerCase();
    const n = Number(t);
    return typeof v === 'number' && isFinite(n) && v === n;
  }
  return compare('=', v, crit);
}

/** Rewrite only reference tokens, never Excel string literals or function names. */
function rewriteRefs(expr: string, fn: (token: string) => string): string {
  return expr.replace(/"(?:[^"]|"")*"|(?<![A-Za-z0-9_.\[])(?:(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?(?![A-Za-z0-9_.\](])/g, token => token.startsWith('"') ? token : fn(token));
}

export function renameFormulaSheet(expr: string, oldName: string, newName: string | null): string {
  return rewriteRefs(expr, token => {
    const range = parseRange(token);
    if (range?.sheet?.toLowerCase() !== oldName.toLowerCase()) return token;
    if (newName === null) return ERR_REF;
    return "'" + newName.replaceAll("'", "''") + "'!" + token.slice(token.lastIndexOf('!')+1);
  });
}

export function shiftFormula(expr: string, sheet: string, axis: 'row'|'col', at: number, delta: number, delFrom = -1, owner = sheet): string {
  return rewriteRefs(expr, token => {
    const range = parseRange(token);
    if (!range || (range.sheet ?? owner).toLowerCase() !== sheet.toLowerCase()) return token;
    const bang = token.lastIndexOf('!');
    const prefix = bang >= 0 ? token.slice(0,bang+1) : '';
    const points = token.slice(bang+1).split(':');
    const parsed = points.map(p => /^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(p)!);
    const indices = parsed.map(m => axis === 'row' ? Number(m[4])-1 : [...m[2].toUpperCase()].reduce((a,c)=>a*26+c.charCodeAt(0)-64,0)-1);
    if (delta < 0 && delFrom >= 0) {
      if (indices.length === 1 && indices[0] >= delFrom && indices[0] < at) return ERR_REF;
      if (indices.length === 2) {
        const lo=Math.min(...indices),hi=Math.max(...indices);
        if(lo>=delFrom && hi<at) return ERR_REF;
        for(let i=0;i<2;i++) if(indices[i]>=delFrom && indices[i]<at) indices[i] = indices[i]===lo ? at : delFrom-1;
      }
    }
    return prefix + parsed.map((m,i) => {
      let n = indices[i]; if (n>=at) n+=delta;
      if (n<0 || n >= (axis==='row'?1048576:16384)) return ERR_REF;
      let letters=''; for(let v=n+1;v;v=Math.floor((v-1)/26)) letters=String.fromCharCode(65+(v-1)%26)+letters;
      return axis==='row' ? `${m[1]}${m[2]}${m[3]}${n+1}` : `${m[1]}${letters}${m[3]}${m[4]}`;
    }).join(':');
  });
}

/** Relative references move on copy; absolute axes and string literals do not. */
export function translateFormula(expr: string, dr: number, dc: number): string {
  return rewriteRefs(expr, token => {
    const bang=token.lastIndexOf('!');
    const prefix=bang<0?'':token.slice(0,bang+1);
    return prefix+token.slice(bang+1).split(':').map(point=>{
      const m=/^(\$?)([A-Za-z]+)(\$?)(\d+)$/.exec(point)!;
      let c=[...m[2].toUpperCase()].reduce((n,ch)=>n*26+ch.charCodeAt(0)-64,0)-1;
      let r=Number(m[4])-1;
      if(!m[1])c+=dc;if(!m[3])r+=dr;
      if(r<0||c<0||r>=1048576||c>=16384)return ERR_REF;
      let letters='';for(let n=c+1;n;n=Math.floor((n-1)/26))letters=String.fromCharCode(65+(n-1)%26)+letters;
      return `${m[1]}${letters}${m[3]}${r+1}`;
    }).join(':');
  });
}
