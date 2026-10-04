import React, { useMemo } from 'react';
import { Box, Typography } from '@mui/material';
import type { WorkbookModel } from '../workbook/model';
import type { Selection } from '../store';

interface StatusBarProps {
  sheetName: string | null;
  model: WorkbookModel | null;
  dirty: boolean;
  selLabel: string;
  selection: Selection | null;
}

export const StatusBar: React.FC<StatusBarProps> = ({
  sheetName,
  model,
  dirty,
  selLabel,
  selection,
}) => {
  const sheetStats = useMemo(() => {
    if (!sheetName || !model) return 'Choose a sheet';
    const sh = model.getSheet(sheetName);
    if (!sh) return 'Choose a sheet';
    return `${sh.rowCount.toLocaleString()} rows · ${sh.colCount} columns`;
  }, [sheetName, model]);

  // Compute live selection statistics (Sum, Average, Count)
  const selectionMetrics = useMemo(() => {
    if (!sheetName || !model || !selection || !selection.ranges.length) return null;
    let sum = 0;
    let numCount = 0;
    let totalCount = 0;
    let cellLimit = 20000;

    for (const rg of selection.ranges) {
      const area = (rg.r2 - rg.r1 + 1) * (rg.c2 - rg.c1 + 1);
      totalCount += area;
      if (totalCount > cellLimit) break;

      for (let r = rg.r1; r <= rg.r2; r++) {
        for (let c = rg.c1; c <= rg.c2; c++) {
          const rec = model.getRec(sheetName, r, c);
          if (rec && typeof rec.v === 'number') {
            sum += rec.v;
            numCount++;
          }
        }
      }
    }

    if (totalCount < 2) return null;

    const avg = numCount > 0 ? (sum / numCount).toLocaleString(undefined, { maximumFractionDigits: 2 }) : '0';
    return {
      sum: sum.toLocaleString(undefined, { maximumFractionDigits: 2 }),
      avg,
      count: totalCount,
      numCount,
    };
  }, [sheetName, model, selection]);

  return (
    <footer className="status-bar">
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 2 }}>
        <Typography variant="caption" sx={{ fontSize: '0.75rem', color: 'text.secondary', fontWeight: 500 }}>
          {sheetStats}
        </Typography>

        {selectionMetrics && (
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5, pl: 1.5, borderLeft: 1, borderColor: 'divider' }}>
            <Typography variant="caption" sx={{ fontSize: '0.75rem', fontFamily: 'monospace', color: 'text.primary' }}>
              Count: <b>{selectionMetrics.count}</b>
            </Typography>
            {selectionMetrics.numCount > 0 && (
              <>
                <Typography variant="caption" sx={{ fontSize: '0.75rem', fontFamily: 'monospace', color: 'text.primary' }}>
                  Sum: <b>{selectionMetrics.sum}</b>
                </Typography>
                <Typography variant="caption" sx={{ fontSize: '0.75rem', fontFamily: 'monospace', color: 'text.primary' }}>
                  Avg: <b>{selectionMetrics.avg}</b>
                </Typography>
              </>
            )}
          </Box>
        )}
      </Box>

      {/* Must contain "All changes saved" / "Unsaved changes" for tests */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1.5 }}>
        <Typography
          variant="caption"
          sx={{
            fontSize: '0.75rem',
            color: dirty ? 'warning.main' : 'text.secondary',
            fontWeight: dirty ? 600 : 400,
          }}
        >
          {dirty ? '● Unsaved changes (Ctrl+S to save)' : 'All changes saved'}
        </Typography>
        <Typography variant="caption" sx={{ fontSize: '0.75rem', color: 'text.secondary' }}>
          ·
        </Typography>
        <Typography variant="caption" sx={{ fontSize: '0.75rem', fontFamily: 'monospace', fontWeight: 600, color: 'text.primary' }}>
          {selLabel}
        </Typography>
        <Typography variant="caption" sx={{ fontSize: '0.75rem', color: 'text.secondary' }}>
          · Double-click / F2 / type to edit
        </Typography>
      </Box>
    </footer>
  );
};
