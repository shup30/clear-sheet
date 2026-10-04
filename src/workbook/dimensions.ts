import type { Command } from './commands';
import type { WorkbookModel } from './model';

/** Sizes are stored at 100% zoom; the grid adapter scales only their display. */
export function resizeCommand(model: WorkbookModel, sheet: string, axis: 'row' | 'col', sizes: Map<number, number>): Command {
  const sh = model.getSheet(sheet);
  if (!sh) throw new Error('Worksheet no longer exists.');
  const map = () => {
    const current = model.getSheet(sheet);
    if (!current) throw new Error('Worksheet no longer exists.');
    return axis === 'row' ? current.rowHeights : current.colWidths;
  };
  const before = new Map([...sizes.keys()].map(index => [index, map().get(index)]));
  for (const [index, size] of sizes) {
    if (!Number.isInteger(index) || index < 0 || !Number.isFinite(size) || size <= 0) throw new Error('Invalid row or column size.');
  }
  return {
    label: axis === 'row' ? 'Resize rows' : 'Resize columns', sheets: [sheet], viewChange: 'dimensions',
    do() { for (const [index, size] of sizes) map().set(index, size); model.dirty = true; },
    undo() { for (const [index, size] of before) { if (size === undefined) map().delete(index); else map().set(index, size); } model.dirty = true; },
  };
}
