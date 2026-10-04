import React, { useState } from 'react';
import {
  Toolbar as MuiToolbar, Box, Tooltip, IconButton, Button, Divider,
  FormControl, InputLabel, Select, MenuItem, Menu, TextField
} from '@mui/material';
import {
  Undo2, Redo2, Scissors, Copy, Clipboard, Bold, Italic, Underline,
  ChevronDown, Search, Grid, Eraser, WrapText,
  BarChart2, Sigma, Sliders
} from 'lucide-react';
import { FONT_CHOICES, SIZE_CHOICES, NUM_FORMATS } from '../workbook/format';
import type { CellStyle } from '../../electron/types';

interface CommandBarProps {
  canUndo: boolean;
  canRedo: boolean;
  hasSelection: boolean;
  hasSelectedCell: boolean;
  busy: boolean;
  zoom: number;
  query: string;
  searchRef: React.RefObject<HTMLInputElement | null>;
  onUndo: () => void;
  onRedo: () => void;
  onCopy: (cut: boolean) => void;
  onPaste: () => void;
  onCopyCell: () => void;
  onCopyRows: () => void;
  onFormat: (patch: Partial<CellStyle> | null, numFmt?: string, label?: string) => void;
  onInsertDel: (axis: 'row' | 'col', mode: 'insert' | 'delete') => void;
  onOpenDim: (kind: 'colWidth' | 'rowHeight') => void;
  onHideUnhide: (axis: 'row' | 'col', hide: boolean) => void;
  onUnhideAll: () => void;
  onResetFilters: () => void;
  onFreeze: (rows: number, cols: number) => void;
  onFreezeAtSelection: () => void;
  onOpenFind: () => void;
  onSearchChange: (value: string) => void;
  onZoomChange: (value: number) => void;
  isBoldActive: boolean;
  isItalicActive: boolean;
  isUnderlineActive: boolean;
  isWrapActive: boolean;
  onOpenFormatCells?: () => void;
  onOpenChart?: () => void;
  onAutoSum?: () => void;
}

export const CommandBar: React.FC<CommandBarProps> = ({
  canUndo,
  canRedo,
  hasSelection,
  hasSelectedCell,
  busy,
  zoom,
  query,
  searchRef,
  onUndo,
  onRedo,
  onCopy,
  onPaste,
  onCopyCell,
  onCopyRows,
  onFormat,
  onInsertDel,
  onOpenDim,
  onHideUnhide,
  onUnhideAll,
  onResetFilters,
  onFreeze,
  onFreezeAtSelection,
  onOpenFind,
  onSearchChange,
  onZoomChange,
  isBoldActive,
  isItalicActive,
  isUnderlineActive,
  isWrapActive,
  onOpenFormatCells,
  onOpenChart,
  onAutoSum,
}) => {
  const [insertMenu, setInsertMenu] = useState<null | HTMLElement>(null);
  const [freezeMenu, setFreezeMenu] = useState<null | HTMLElement>(null);

  return (
    <MuiToolbar className="ribbon command-bar" variant="dense" disableGutters sx={{ px: 1.5, py: 0.5, minHeight: 38 }}>
      {/* Group: History */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
        <Tooltip title="Undo (Ctrl+Z)">
          <span>
            <IconButton size="small" disabled={!canUndo || busy} onClick={onUndo} aria-label="Undo">
              <Undo2 size={15} />
            </IconButton>
          </span>
        </Tooltip>
        <Tooltip title="Redo (Ctrl+Y)">
          <span>
            <IconButton size="small" disabled={!canRedo || busy} onClick={onRedo} aria-label="Redo">
              <Redo2 size={15} />
            </IconButton>
          </span>
        </Tooltip>
      </Box>

      <Divider orientation="vertical" flexItem sx={{ mx: 0.75, my: 0.5 }} />

      {/* Group: Clipboard */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
        <Tooltip title="Cut (Ctrl+X)">
          <span>
            <IconButton size="small" disabled={!hasSelection || busy} onClick={() => onCopy(true)} aria-label="Cut">
              <Scissors size={15} />
            </IconButton>
          </span>
        </Tooltip>
        <Tooltip title="Copy (Ctrl+C)">
          <span>
            <IconButton size="small" disabled={!hasSelection || busy} onClick={() => onCopy(false)} aria-label="Copy">
              <Copy size={15} />
            </IconButton>
          </span>
        </Tooltip>
        <Tooltip title="Paste (Ctrl+V)">
          <span>
            <IconButton size="small" disabled={!hasSelection || busy} onClick={onPaste} aria-label="Paste">
              <Clipboard size={15} />
            </IconButton>
          </span>
        </Tooltip>
        <Button size="small" variant="text" disabled={!hasSelectedCell || busy} onClick={onCopyCell} sx={{ fontSize: '0.75rem', px: 0.75 }}>
          Copy cell
        </Button>
        <Button size="small" variant="text" disabled={busy} onClick={onCopyRows} sx={{ fontSize: '0.75rem', px: 0.75 }}>
          Copy rows
        </Button>
      </Box>

      <Divider orientation="vertical" flexItem sx={{ mx: 0.75, my: 0.5 }} />

      {/* Group: Typography & Styles */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
        <FormControl size="small" sx={{ minWidth: 105 }}>
          <InputLabel id="font-label" sx={{ fontSize: '0.75rem' }}>Font</InputLabel>
          <Select
            labelId="font-label"
            label="Font"
            value=""
            displayEmpty
            onChange={(e) => onFormat({ fontName: String(e.target.value) }, undefined, 'Font')}
            aria-label="Font family"
            sx={{ height: 28, fontSize: '0.75rem' }}
          >
            {FONT_CHOICES.map((f) => (
              <MenuItem key={f} value={f} sx={{ fontFamily: f, fontSize: '0.8125rem' }}>
                {f}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        <FormControl size="small" sx={{ minWidth: 65 }}>
          <InputLabel id="size-label" sx={{ fontSize: '0.75rem' }}>Size</InputLabel>
          <Select
            labelId="size-label"
            label="Size"
            value=""
            displayEmpty
            onChange={(e) => onFormat({ fontSize: Number(e.target.value) }, undefined, 'Font size')}
            aria-label="Font size"
            sx={{ height: 28, fontSize: '0.75rem' }}
          >
            {SIZE_CHOICES.map((n) => (
              <MenuItem key={n} value={n} sx={{ fontSize: '0.8125rem' }}>
                {n}
              </MenuItem>
            ))}
          </Select>
        </FormControl>

        <Tooltip title="Bold (Ctrl+B)">
          <span>
            <IconButton
              size="small"
              aria-label="Bold"
              color={isBoldActive ? 'primary' : 'default'}
              onClick={() => onFormat({ bold: !isBoldActive }, undefined, 'Bold')}
              sx={{ bgcolor: isBoldActive ? 'action.selected' : 'transparent' }}
            >
              <Bold size={15} />
            </IconButton>
          </span>
        </Tooltip>

        <Tooltip title="Italic (Ctrl+I)">
          <span>
            <IconButton
              size="small"
              aria-label="Italic"
              color={isItalicActive ? 'primary' : 'default'}
              onClick={() => onFormat({ italic: !isItalicActive }, undefined, 'Italic')}
              sx={{ bgcolor: isItalicActive ? 'action.selected' : 'transparent' }}
            >
              <Italic size={15} />
            </IconButton>
          </span>
        </Tooltip>

        <Tooltip title="Underline (Ctrl+U)">
          <span>
            <IconButton
              size="small"
              aria-label="Underline"
              color={isUnderlineActive ? 'primary' : 'default'}
              onClick={() => onFormat({ underline: !isUnderlineActive }, undefined, 'Underline')}
              sx={{ bgcolor: isUnderlineActive ? 'action.selected' : 'transparent' }}
            >
              <Underline size={15} />
            </IconButton>
          </span>
        </Tooltip>

        <Tooltip title="Text color">
          <Box sx={{ position: 'relative', display: 'inline-flex', alignItems: 'center' }}>
            <input
              type="color"
              aria-label="Text color"
              className="color-picker-input"
              defaultValue="#000000"
              onChange={(e) => onFormat({ color: e.target.value }, undefined, 'Text color')}
            />
          </Box>
        </Tooltip>

        <Tooltip title="Fill color">
          <Box sx={{ position: 'relative', display: 'inline-flex', alignItems: 'center' }}>
            <input
              type="color"
              aria-label="Fill color"
              className="color-picker-input fill"
              defaultValue="#ffffff"
              onChange={(e) => onFormat({ bg: e.target.value }, undefined, 'Fill color')}
            />
          </Box>
        </Tooltip>

        <Select
          size="small"
          value=""
          displayEmpty
          aria-label="Alignment"
          onChange={(e) => {
            const v = String(e.target.value);
            if (v === 'left' || v === 'center' || v === 'right') onFormat({ hAlign: v }, undefined, 'Alignment');
          }}
          sx={{ height: 28, fontSize: '0.75rem', minWidth: 70 }}
        >
          <MenuItem value="" sx={{ fontSize: '0.8125rem' }}>Align</MenuItem>
          <MenuItem value="left" sx={{ fontSize: '0.8125rem' }}>Left</MenuItem>
          <MenuItem value="center" sx={{ fontSize: '0.8125rem' }}>Center</MenuItem>
          <MenuItem value="right" sx={{ fontSize: '0.8125rem' }}>Right</MenuItem>
        </Select>

        <Tooltip title="Wrap text">
          <span>
            <Button
              size="small"
              aria-label="Wrap text"
              variant={isWrapActive ? 'contained' : 'text'}
              onClick={() => onFormat({ wrap: !isWrapActive }, undefined, 'Wrap text')}
              startIcon={<WrapText size={14} />}
              sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
            >
              Wrap
            </Button>
          </span>
        </Tooltip>

        <Tooltip title="All borders">
          <span>
            <Button
              size="small"
              aria-label="Borders"
              onClick={() => onFormat({ border: { top: true, bottom: true, left: true, right: true } }, undefined, 'Borders')}
              sx={{ height: 28, minWidth: 32, px: 0.75 }}
            >
              <Grid size={15} />
            </Button>
          </span>
        </Tooltip>

        <Tooltip title="Clear formatting">
          <span>
            <Button
              size="small"
              aria-label="Clear formatting"
              onClick={() => onFormat(null, undefined, 'Clear formatting')}
              startIcon={<Eraser size={13} />}
              sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
            >
              ✕Fmt
            </Button>
          </span>
        </Tooltip>
      </Box>

      <Divider orientation="vertical" flexItem sx={{ mx: 0.75, my: 0.5 }} />

      {/* Group: Number format */}
      <Box sx={{ display: 'flex', alignItems: 'center' }}>
        <FormControl size="small" sx={{ minWidth: 115 }}>
          <InputLabel id="num-label" sx={{ fontSize: '0.75rem' }}>Number</InputLabel>
          <Select
            labelId="num-label"
            label="Number"
            value=""
            displayEmpty
            aria-label="Number format"
            onChange={(e) => {
              const v = String(e.target.value);
              if (v) onFormat({}, v, 'Number format');
            }}
            sx={{ height: 28, fontSize: '0.75rem' }}
          >
            {NUM_FORMATS.map((f) => (
              <MenuItem key={f.label} value={f.fmt} sx={{ fontSize: '0.8125rem' }}>
                {f.label}
              </MenuItem>
            ))}
          </Select>
        </FormControl>
      </Box>

      <Divider orientation="vertical" flexItem sx={{ mx: 0.75, my: 0.5 }} />

      {/* Group: Structure & Tools */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5 }}>
        <Button
          size="small"
          aria-label="Insert menu"
          onClick={(e) => setInsertMenu(e.currentTarget)}
          endIcon={<ChevronDown size={14} />}
          sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
        >
          Insert/Delete
        </Button>
        <Menu anchorEl={insertMenu} open={!!insertMenu} onClose={() => setInsertMenu(null)}>
          <MenuItem onClick={() => { setInsertMenu(null); onInsertDel('row', 'insert'); }}>Insert row(s)</MenuItem>
          <MenuItem onClick={() => { setInsertMenu(null); onInsertDel('row', 'delete'); }}>Delete row(s)</MenuItem>
          <MenuItem onClick={() => { setInsertMenu(null); onInsertDel('col', 'insert'); }}>Insert column(s)</MenuItem>
          <MenuItem onClick={() => { setInsertMenu(null); onInsertDel('col', 'delete'); }}>Delete column(s)</MenuItem>
          <Divider sx={{ my: 0.5 }} />
          <MenuItem onClick={() => { setInsertMenu(null); onOpenDim('colWidth'); }}>Column width…</MenuItem>
          <MenuItem onClick={() => { setInsertMenu(null); onOpenDim('rowHeight'); }}>Row height…</MenuItem>
          <Divider sx={{ my: 0.5 }} />
          <MenuItem onClick={() => { setInsertMenu(null); onHideUnhide('row', true); }}>Hide rows</MenuItem>
          <MenuItem onClick={() => { setInsertMenu(null); onHideUnhide('col', true); }}>Hide columns</MenuItem>
          <MenuItem onClick={() => { setInsertMenu(null); onUnhideAll(); }}>Unhide all</MenuItem>
        </Menu>

        <Button
          size="small"
          aria-label="Freeze menu"
          onClick={(e) => setFreezeMenu(e.currentTarget)}
          endIcon={<ChevronDown size={14} />}
          sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
        >
          Freeze
        </Button>
        <Menu anchorEl={freezeMenu} open={!!freezeMenu} onClose={() => setFreezeMenu(null)}>
          <MenuItem onClick={() => { setFreezeMenu(null); onFreeze(1, 0); }}>Freeze first row</MenuItem>
          <MenuItem onClick={() => { setFreezeMenu(null); onFreeze(0, 1); }}>Freeze first column</MenuItem>
          <MenuItem onClick={() => { setFreezeMenu(null); onFreezeAtSelection(); }}>Freeze at selection</MenuItem>
          <MenuItem onClick={() => { setFreezeMenu(null); onFreeze(0, 0); }}>Unfreeze</MenuItem>
        </Menu>

        <Button size="small" onClick={onResetFilters} sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}>
          Reset filters
        </Button>

        <Button size="small" onClick={onOpenFind} sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}>
          Find/Replace
        </Button>

        <Divider orientation="vertical" flexItem sx={{ my: 0.5 }} />

        {onAutoSum && (
          <Tooltip title="AutoSum (Alt+=)">
            <Button
              size="small"
              onClick={onAutoSum}
              startIcon={<Sigma size={13} />}
              sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
            >
              AutoSum
            </Button>
          </Tooltip>
        )}

        {onOpenChart && (
          <Tooltip title="Insert Chart from Selection">
            <Button
              size="small"
              onClick={onOpenChart}
              startIcon={<BarChart2 size={13} />}
              sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
            >
              Chart
            </Button>
          </Tooltip>
        )}

        {onOpenFormatCells && (
          <Tooltip title="Format Cells (Ctrl+1)">
            <Button
              size="small"
              onClick={onOpenFormatCells}
              startIcon={<Sliders size={13} />}
              sx={{ height: 28, fontSize: '0.75rem', px: 0.75 }}
            >
              Format Cells…
            </Button>
          </Tooltip>
        )}
      </Box>

      <Box sx={{ flex: 1 }} />

      {/* Right: Quick filter & zoom */}
      <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
        <TextField
          size="small"
          placeholder="Search sheet · Ctrl+F"
          value={query}
          inputRef={searchRef}
          onChange={(e) => onSearchChange(e.target.value)}
          slotProps={{
            htmlInput: { 'aria-label': 'Search sheet' },
            input: {
              startAdornment: <Search size={14} style={{ marginRight: 6, opacity: 0.5 }} />,
            }
          }}
          sx={{
            width: 175,
            '& .MuiInputBase-root': { height: 28, fontSize: '0.75rem' },
          }}
        />

        <Select
          size="small"
          value={zoom}
          inputProps={{ 'aria-label': 'Zoom' }}
          onChange={(e) => onZoomChange(Number(e.target.value))}
          sx={{ height: 28, fontSize: '0.75rem', minWidth: 80 }}
        >
          {[75, 90, 100, 110, 125, 150, 175, 200].map((n) => (
            <MenuItem key={n} value={n} sx={{ fontSize: '0.8125rem' }}>
              {n}%
            </MenuItem>
          ))}
        </Select>
      </Box>
    </MuiToolbar>
  );
};
