import { insertDeleteCommand, type Command } from './commands';
import type { WorkbookModel } from './model';

export interface Span { start: number; count: number }

export function mergeSpans(spans: Span[]): Span[] {
  const result: Span[] = [];
  for (const span of [...spans].sort((a, b) => a.start - b.start)) {
    const last = result.at(-1);
    if (last && span.start <= last.start + last.count) last.count = Math.max(last.start + last.count, span.start + span.count) - last.start;
    else result.push({ ...span });
  }
  return result;
}

/** Operate bottom/right first so disjoint selected rows retain their addresses. */
export function axisCommand(model: WorkbookModel, sheet: string, axis: 'row' | 'col', mode: 'insert' | 'delete', spans: Span[]): Command {
  const ordered = mergeSpans(spans).reverse();
  let applied: Command[] = [];
  return {
    label: `${mode} ${axis}s`, sheets: model.sheetNames(),
    do() {
      applied = [];
      try {
        for (const span of ordered) {
          const command = insertDeleteCommand(model, sheet, axis, mode, span.start, span.count);
          command.do(); applied.push(command);
        }
      } catch (error) { for (const command of [...applied].reverse()) command.undo(); throw error; }
    },
    undo() { for (const command of [...applied].reverse()) command.undo(); },
  };
}
