import { WorkbookModel } from './model';
import { inputToChange } from './commands';

/** TSV clipboard helpers (Excel-compatible). */

export function escapeTSVCell(s: string): string {
  return /[\t\r\n"]/.test(s) ? '"' + s.replaceAll('"', '""') + '"' : s;
}

export function rangeToTSV(model: WorkbookModel, sheet: string, r1: number, c1: number, r2: number, c2: number, withFormulas: boolean): string {
  const lines: string[] = [];
  for (let r = r1; r <= r2; r++) {
    const row: string[] = [];
    for (let c = c1; c <= c2; c++) row.push(escapeTSVCell(model.cellTextForCopy(sheet, r, c, withFormulas)));
    lines.push(row.join('\t'));
  }
  return lines.join('\r\n');
}

/** Parse clipboard text (TSV with quoted cells, or plain lines) into a grid. */
export function parseTSV(text: string): string[][] {
  const rows: string[][] = [];
  let row: string[] = [];
  let cur = '';
  let inQuotes = false;
  for (let i = 0; i < text.length; i++) {
    const ch = text[i];
    if (inQuotes) {
      if (ch === '"') {
        if (text[i + 1] === '"') { cur += '"'; i++; }
        else inQuotes = false;
      } else cur += ch;
    } else {
      if (ch === '"' && cur === '') inQuotes = true;
      else if (ch === '\t') { row.push(cur); cur = ''; }
      else if (ch === '\r') { row.push(cur); rows.push(row); row=[]; cur=''; if(text[i+1]==='\n')i++; }
      else if (ch === '\n') { row.push(cur); rows.push(row); row = []; cur = ''; }
      else cur += ch;
    }
  }
  row.push(cur);
  // drop single trailing empty row from final newline
  if (row.length === 1 && row[0] === '' && rows.length) return rows;
  rows.push(row);
  return rows.filter((r, i) => !(i === rows.length - 1 && r.length === 1 && r[0] === ''));
}

export function tsvToChanges(grid: string[][]): { v: string | number | boolean | null; f?: string }[] {
  return grid.flat().map((t) => inputToChange(t));
}
