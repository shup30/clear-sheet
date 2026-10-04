import type { CellStyle } from '../../electron/types';

export type { CellStyle };

export const FONT_CHOICES = ['Calibri', 'Arial', 'Segoe UI', 'Times New Roman', 'Courier New', 'Verdana'];
export const SIZE_CHOICES = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24];

/** Common number formats (subset of Excel ids + custom). */
export const NUM_FORMATS: { label: string; fmt: string }[] = [
  { label: 'General', fmt: 'General' },
  { label: 'Number', fmt: '0.00' },
  { label: 'Comma', fmt: '#,##0.00' },
  { label: 'Currency ($)', fmt: '$#,##0.00' },
  { label: 'Currency (€)', fmt: '€#,##0.00' },
  { label: 'Percent', fmt: '0%' },
  { label: 'Percent (2dp)', fmt: '0.00%' },
  { label: 'Date (M/D/YYYY)', fmt: 'm/d/yyyy' },
  { label: 'Date (D-MMM-YY)', fmt: 'd-mmm-yy' },
  { label: 'Time', fmt: 'h:mm:ss' },
  { label: 'Date+Time', fmt: 'm/d/yyyy h:mm' },
  { label: 'Text', fmt: '@' },
];

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];

/** Excel serial date -> JS Date (with 1900 system; ignores 1904). */
export function excelSerialToDate(n: number): Date {
  const ms = Math.round((n - 25569) * 86400 * 1000);
  return new Date(ms);
}

export function dateToExcelSerial(d: Date): number {
  return d.getTime() / 86400000 + 25569;
}

function pad(n: number, w = 2): string {
  return String(n).padStart(w, '0');
}

function formatDatePart(d: Date, token: string): string {
  switch (token) {
    case 'yyyy': return String(d.getUTCFullYear());
    case 'yy': return String(d.getUTCFullYear()).slice(2);
    case 'mmmm': return d.toLocaleString('en-US', { month: 'long', timeZone: 'UTC' });
    case 'mmm': return MONTHS[d.getUTCMonth()];
    case 'mm': return pad(d.getUTCMonth() + 1);
    case 'm': return String(d.getUTCMonth() + 1);
    case 'dddd': return d.toLocaleString('en-US', { weekday: 'long', timeZone: 'UTC' });
    case 'ddd': return d.toLocaleString('en-US', { weekday: 'short', timeZone: 'UTC' });
    case 'dd': return pad(d.getUTCDate());
    case 'd': return String(d.getUTCDate());
    case 'hh': return pad(d.getUTCHours());
    case 'h': return String(d.getUTCHours());
    case 'ss': return pad(d.getUTCSeconds());
    case 's': return String(d.getUTCSeconds());
    default: return token;
  }
}

/** Minimal date-format renderer for common Excel tokens. */
export function formatExcelDate(serial: number, fmt: string): string {
  const d = excelSerialToDate(serial);
  if (isNaN(d.getTime())) return String(serial);
  const lower = fmt.toLowerCase();
  // minutes vs months: treat m after h as minutes
  let out = '';
  const re = /(yyyy|yy|mmmm|mmm|mm|m|dddd|ddd|dd|d|hh|h|ss|s|am\/pm|:|\/|-|,| )/gi;
  let m: RegExpExecArray | null;
  let afterH = false;
  while ((m = re.exec(fmt))) {
    const t = m[0];
    if (/^h/i.test(t)) afterH = true;
    if (/^m+$/i.test(t) && afterH && lower.includes('h')) {
      // minutes
      out += t.length === 1 ? String(d.getUTCMinutes()) : pad(d.getUTCMinutes());
    } else if (/^m+$/i.test(t) && !afterH) {
      out += formatDatePart(d, t.toLowerCase() === t ? t : t.toLowerCase());
      // normalize month tokens
      if (t === 'M') out = out; // noop
    } else if (/^(yyyy|yy|mmm|mm|dd|hh|ss)$/i.test(t)) {
      out += formatDatePart(d, t.toLowerCase());
    } else if (/^(d|m|h|s)$/i.test(t)) {
      out += formatDatePart(d, t.toLowerCase());
    } else {
      out += t === 'am/pm' || t === 'AM/PM' ? (d.getUTCHours() < 12 ? 'AM' : 'PM') : t;
    }
  }
  return out || d.toISOString().slice(0, 10);
}

export function isDateFormat(fmt?: string): boolean {
  if (!fmt || fmt === 'General' || fmt === '@') return false;
  return /[ymdhhs]/i.test(fmt) && !/^[#,0?.%$-]+$/.test(fmt);
}

/** Format a raw value with an Excel-like number format. */
export function formatValue(v: string | number | boolean | null, fmt?: string): string {
  if (v === null || v === undefined) return '';
  if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
  if (typeof v === 'string') return v;
  if (!fmt || fmt === 'General' || fmt === '@') {
    return String(v);
  }
  if (fmt.endsWith('%')) {
    const dp = fmt.includes('.') ? (fmt.split('.')[1].split('%')[0].length) : 0;
    return (v * 100).toFixed(dp) + '%';
  }
  if (isDateFormat(fmt)) return formatExcelDate(v, fmt);
  if (fmt.startsWith('$') || fmt.startsWith('€') || fmt.startsWith('£')) {
    const sym = fmt[0];
    const dp = fmt.includes('.') ? fmt.split('.')[1].replace(/[^0]/g, '').length || 2 : 0;
    const useComma = fmt.includes(',');
    const body = useComma
      ? v.toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp })
      : v.toFixed(dp);
    return `${sym}${body}`;
  }
  if (fmt.includes(',')) {
    const dp = fmt.includes('.') ? fmt.split('.')[1].length : 0;
    return v.toLocaleString('en-US', { minimumFractionDigits: dp, maximumFractionDigits: dp });
  }
  if (/^0(\.0+)?$/.test(fmt)) {
    const dp = fmt.includes('.') ? fmt.split('.')[1].length : 0;
    return v.toFixed(dp);
  }
  return String(v);
}

/** Parse user-typed input into stored value + optional formula + display format hint. */
export function parseInput(text: string): { v: string | number | boolean | null; f?: string; numFmt?: string } {
  const t = text.trim();
  if (t === '') return { v: null };
  if (text.startsWith("'")) return {v:text.slice(1)};
  if (t.startsWith('=')) return { v: null, f: t.slice(1) };
  if (/^(true|false)$/i.test(t)) return { v: t.toLowerCase() === 'true' };
  // percent literal "12%"
  const pct = /^(-?[\d,]*\.?\d+)%$/.exec(t.replace(/\s/g, ''));
  if (pct) {
    const n = Number(pct[1].replace(/,/g, ''));
    if (isFinite(n)) return { v: n / 100, numFmt: '0%' };
  }
  // plain number
  if (/^-?[\d,]*\.?\d+([eE][+-]?\d+)?$/.test(t)) {
    const n = Number(t.replace(/,/g, ''));
    if (isFinite(n)) return { v: n };
  }
  // date literal (store ISO string; engine coerces for date fns)
  const dt = Date.parse(t);
  if (!isNaN(dt) && /[/\-.]/.test(t) && /\d{2,}/.test(t)) {
    // keep as serial + date format so Excel round-trips as a date
    return { v: dateToExcelSerial(new Date(dt)), numFmt: 'm/d/yyyy' };
  }
  return { v: text };
}

/** Convert our CellStyle to XLSX cell.s (SheetJS style object). */
export function styleToXlsx(s?: CellStyle): any | undefined {
  if (!s) return undefined;
  const out: any = {};
  if (s.fontName || s.fontSize || s.bold || s.italic || s.underline || s.color) {
    out.font = {
      ...(s.fontName ? { name: s.fontName } : {}),
      ...(s.fontSize ? { sz: s.fontSize } : {}),
      ...(s.bold ? { bold: true } : {}),
      ...(s.italic ? { italic: true } : {}),
      ...(s.underline ? { underline: true } : {}),
      ...(s.color ? { color: { rgb: cssToArgb(s.color) } } : {}),
    };
  }
  if (s.bg) out.fill = { patternType: 'solid', fgColor: { rgb: cssToArgb(s.bg) } };
  if (s.hAlign || s.vAlign || s.wrap) {
    out.alignment = {
      ...(s.hAlign ? { horizontal: s.hAlign } : {}),
      ...(s.vAlign ? { vertical: s.vAlign } : {}),
      ...(s.wrap ? { wrapText: true } : {}),
    };
  }
  if (s.numFmt && s.numFmt !== 'General') out.numFmt = s.numFmt;
  if (s.border && (s.border.top || s.border.bottom || s.border.left || s.border.right)) {
    const side = { style: 'thin', color: { rgb: 'FF000000' } };
    out.border = {
      ...(s.border.top ? { top: side } : {}),
      ...(s.border.bottom ? { bottom: side } : {}),
      ...(s.border.left ? { left: side } : {}),
      ...(s.border.right ? { right: side } : {}),
    };
  }
  return Object.keys(out).length ? out : undefined;
}

/** Best-effort mapping from XLSX cell (s/z) to our CellStyle. */
export function xlsxToStyle(cell: any): CellStyle | undefined {
  const s = cell?.s;
  const z = cell?.z;
  if (!s && !z) return undefined;
  const st: CellStyle = {};
  const font = s?.font;
  if (font) {
    if (font.bold) st.bold = true;
    if (font.italic) st.italic = true;
    if (font.underline) st.underline = true;
    if (font.name) st.fontName = String(font.name);
    if (font.sz) st.fontSize = Number(font.sz);
    if (font.color?.rgb) st.color = argbToCss(String(font.color.rgb));
  }
  const fill = s?.fill;
  const fg = fill?.fgColor?.rgb;
  if (fg && fill?.patternType && fill.patternType !== 'none') st.bg = argbToCss(String(fg));
  const al = s?.alignment;
  if (al) {
    if (['left', 'center', 'right'].includes(al.horizontal)) st.hAlign = al.horizontal;
    if (['top', 'middle', 'bottom', 'center'].includes(al.vertical)) st.vAlign = al.vertical === 'center' ? 'middle' : al.vertical;
    if (al.wrapText) st.wrap = true;
  }
  const border = s?.border;
  if (border) {
    st.border = {
      top: !!border.top, bottom: !!border.bottom, left: !!border.left, right: !!border.right,
    };
    if (!st.border.top && !st.border.bottom && !st.border.left && !st.border.right) delete st.border;
  }
  if (typeof z === 'string' && z && z !== 'General') st.numFmt = z;
  else if (typeof s?.numFmt === 'string' && s.numFmt !== 'General') st.numFmt = s.numFmt;
  return Object.keys(st).length ? st : undefined;
}

function cssToArgb(css: string): string {
  let h = css.trim().replace('#', '');
  if (h.length === 3) h = h.split('').map((c) => c + c).join('');
  if (h.length === 6) h = 'FF' + h;
  return h.toUpperCase().slice(0, 8).padStart(8, 'F');
}

function argbToCss(argb: string): string {
  const h = argb.replace('#', '');
  if (h.length === 8) return '#' + h.slice(2);
  if (h.length === 6) return '#' + h;
  return '#000000';
}
