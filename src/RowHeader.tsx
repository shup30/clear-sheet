import { useRef } from 'react';
import type { ICellRendererParams } from 'ag-grid-community';
import type { Row } from '../electron/types';
import { useViewer } from './store';
import { resizeCommand } from './workbook/dimensions';

export function RowHeader({ data, node, api }: ICellRendererParams<Row>) {
  const drag = useRef<{ startY: number; height: number; current: number; zoom: number; state: ReturnType<typeof useViewer.getState> } | null>(null);
  if (!data) return null;
  const row = data.id - 1;
  return <div className="row-heading"><span>{data.id}</span><div className="row-resize-handle"
    role="separator" aria-orientation="horizontal" aria-label={`Resize row ${data.id}`} title="Drag to resize row"
    onClick={e => e.stopPropagation()}
    onPointerDown={e => {
      if (e.button !== 0) return;
      const state = useViewer.getState();
      if (state.busy || !state.model || !state.sheetName) return;
      e.preventDefault(); e.stopPropagation(); e.currentTarget.setPointerCapture(e.pointerId);
      const height = state.model.getSheet(state.sheetName)?.rowHeights.get(row) ?? 30;
      drag.current = { startY: e.clientY, height, current: height, zoom: state.zoom / 100, state };
    }}
    onPointerMove={e => {
      const active = drag.current; if (!active) return;
      e.stopPropagation();
      active.current = Math.round(Math.max(20, Math.min(600, active.height + (e.clientY - active.startY) / active.zoom)));
      node.setRowHeight(active.current * active.zoom); api.onRowHeightChanged();
    }}
    onPointerUp={e => {
      const active = drag.current; if (!active) return;
      drag.current = null; e.stopPropagation(); e.currentTarget.releasePointerCapture(e.pointerId);
      const state = useViewer.getState();
      if (!state.busy && state.model === active.state.model && state.sheetName === active.state.sheetName && active.current !== active.height) {
        state.exec(resizeCommand(state.model!, state.sheetName!, 'row', new Map([[row, active.current]])));
      } else { node.setRowHeight(active.height * active.zoom); api.onRowHeightChanged(); }
    }}
    onPointerCancel={() => {
      const active = drag.current; drag.current = null;
      if (active) { node.setRowHeight(active.height * active.zoom); api.onRowHeightChanged(); }
    }}
  /></div>;
}
