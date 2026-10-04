import React, { useState, useMemo } from 'react';
import {
  Dialog, DialogContent, Box, InputBase, List, ListItemButton,
  ListItemIcon, ListItemText, Typography, Chip
} from '@mui/material';
import {
  Search, Save, FolderOpen, RotateCcw, RotateCw, Plus,
  Sun, Moon, ZoomIn, Eraser, FilterX, HelpCircle, Columns, Rows,
  BarChart2, Sigma, Sliders
} from 'lucide-react';
import type { ColorMode } from '../theme/tokens';

export interface CommandItem {
  id: string;
  title: string;
  category: string;
  shortcut?: string;
  icon: React.ReactNode;
  action: () => void;
}

interface CommandPaletteProps {
  open: boolean;
  onClose: () => void;
  mode: ColorMode;
  onToggleTheme: () => void;
  onSave: () => void;
  onSaveAs: () => void;
  onOpen: () => void;
  onOpenFind: () => void;
  onUndo: () => void;
  onRedo: () => void;
  onAddSheet: () => void;
  onFreezeFirstRow: () => void;
  onFreezeFirstCol: () => void;
  onUnfreeze: () => void;
  onResetFilters: () => void;
  onClearFormatting: () => void;
  onSetZoom: (zoom: number) => void;
  onFormatCells?: () => void;
  onInsertChart?: () => void;
  onAutoSum?: () => void;
  onAddCf?: (type: 'gt' | 'colorScale' | 'dataBar') => void;
  onClearCf?: () => void;
}

export const CommandPalette: React.FC<CommandPaletteProps> = ({
  open,
  onClose,
  mode,
  onToggleTheme,
  onSave,
  onSaveAs,
  onOpen,
  onOpenFind,
  onUndo,
  onRedo,
  onAddSheet,
  onFreezeFirstRow,
  onFreezeFirstCol,
  onUnfreeze,
  onResetFilters,
  onClearFormatting,
  onSetZoom,
  onFormatCells,
  onInsertChart,
  onAutoSum,
  onAddCf,
  onClearCf,
}) => {
  const [query, setQuery] = useState('');
  const [selectedIndex, setSelectedIndex] = useState(0);

  const commands = useMemo<CommandItem[]>(() => [
    { id: 'save', title: 'Save Workbook', category: 'File', shortcut: 'Ctrl+S', icon: <Save size={16} />, action: onSave },
    { id: 'save-as', title: 'Save As…', category: 'File', shortcut: 'Ctrl+Shift+S', icon: <Save size={16} />, action: onSaveAs },
    { id: 'open', title: 'Open File…', category: 'File', shortcut: 'Ctrl+O', icon: <FolderOpen size={16} />, action: onOpen },
    { id: 'find', title: 'Find and Replace', category: 'Edit', shortcut: 'Ctrl+H', icon: <Search size={16} />, action: onOpenFind },
    { id: 'undo', title: 'Undo', category: 'Edit', shortcut: 'Ctrl+Z', icon: <RotateCcw size={16} />, action: onUndo },
    { id: 'redo', title: 'Redo', category: 'Edit', shortcut: 'Ctrl+Y', icon: <RotateCw size={16} />, action: onRedo },
    { id: 'add-sheet', title: 'Add New Worksheet', category: 'Sheets', icon: <Plus size={16} />, action: onAddSheet },
    { id: 'freeze-row', title: 'Freeze First Row', category: 'View', icon: <Rows size={16} />, action: onFreezeFirstRow },
    { id: 'freeze-col', title: 'Freeze First Column', category: 'View', icon: <Columns size={16} />, action: onFreezeFirstCol },
    { id: 'unfreeze', title: 'Unfreeze Panes', category: 'View', icon: <Columns size={16} />, action: onUnfreeze },
    { id: 'reset-filters', title: 'Reset Filters & Sorting', category: 'Data', icon: <FilterX size={16} />, action: onResetFilters },
    { id: 'format-cells', title: 'Format Cells…', category: 'Format', shortcut: 'Ctrl+1', icon: <Sliders size={16} />, action: () => onFormatCells?.() },
    { id: 'cf-color-scale', title: 'Conditional Formatting: Color Scale (Heatmap)', category: 'Format', icon: <Sliders size={16} />, action: () => onAddCf?.('colorScale') },
    { id: 'cf-data-bar', title: 'Conditional Formatting: Data Bars', category: 'Format', icon: <Sliders size={16} />, action: () => onAddCf?.('dataBar') },
    { id: 'cf-highlight-gt', title: 'Conditional Formatting: Highlight > 0', category: 'Format', icon: <Sliders size={16} />, action: () => onAddCf?.('gt') },
    { id: 'cf-clear', title: 'Clear Conditional Formatting', category: 'Format', icon: <Eraser size={16} />, action: () => onClearCf?.() },
    { id: 'insert-chart', title: 'Insert Chart…', category: 'Insert', shortcut: '', icon: <BarChart2 size={16} />, action: () => onInsertChart?.() },
    { id: 'auto-sum', title: 'AutoSum', category: 'Formulas', shortcut: 'Alt+=', icon: <Sigma size={16} />, action: () => onAutoSum?.() },
    { id: 'clear-fmt', title: 'Clear Cell Formatting', category: 'Format', icon: <Eraser size={16} />, action: onClearFormatting },
    { id: 'theme', title: `Switch to ${mode === 'light' ? 'Dark' : 'Light'} Mode`, category: 'Appearance', icon: mode === 'light' ? <Moon size={16} /> : <Sun size={16} />, action: onToggleTheme },
    { id: 'zoom-100', title: 'Zoom: 100% Normal', category: 'View', icon: <ZoomIn size={16} />, action: () => onSetZoom(100) },
    { id: 'zoom-125', title: 'Zoom: 125% Larger', category: 'View', icon: <ZoomIn size={16} />, action: () => onSetZoom(125) },
    { id: 'zoom-150', title: 'Zoom: 150% Extra Large', category: 'View', icon: <ZoomIn size={16} />, action: () => onSetZoom(150) },
  ], [mode, onToggleTheme, onSave, onSaveAs, onOpen, onOpenFind, onUndo, onRedo, onAddSheet, onFreezeFirstRow, onFreezeFirstCol, onUnfreeze, onResetFilters, onClearFormatting, onSetZoom, onFormatCells, onInsertChart, onAutoSum, onAddCf, onClearCf]);

  const filtered = useMemo(() => {
    if (!query.trim()) return commands;
    const q = query.toLowerCase();
    return commands.filter(c => c.title.toLowerCase().includes(q) || c.category.toLowerCase().includes(q));
  }, [commands, query]);

  const executeCommand = (cmd: CommandItem) => {
    onClose();
    setQuery('');
    setSelectedIndex(0);
    cmd.action();
  };

  return (
    <Dialog
      open={open}
      onClose={() => {
        onClose();
        setQuery('');
      }}
      maxWidth="sm"
      fullWidth
      slotProps={{
        backdrop: {
          sx: { backdropFilter: 'blur(4px)', backgroundColor: 'rgba(0, 0, 0, 0.4)' },
        },
      }}
      sx={{
        '& .MuiDialog-paper': {
          mt: 8,
          verticalAlign: 'top',
          borderRadius: 3,
          overflow: 'hidden',
          boxShadow: '0 20px 40px rgba(0, 0, 0, 0.3)',
        }
      }}
    >
      <DialogContent sx={{ p: 0 }}>
        {/* Search header */}
        <Box sx={{ display: 'flex', alignItems: 'center', px: 2, py: 1.5, borderBottom: 1, borderColor: 'divider', gap: 1.5 }}>
          <Search size={18} style={{ opacity: 0.5 }} />
          <InputBase
            autoFocus
            fullWidth
            placeholder="Type a command or search action…"
            value={query}
            onChange={(e) => {
              setQuery(e.target.value);
              setSelectedIndex(0);
            }}
            onKeyDown={(e) => {
              if (e.key === 'ArrowDown') {
                e.preventDefault();
                setSelectedIndex(i => Math.min(filtered.length - 1, i + 1));
              } else if (e.key === 'ArrowUp') {
                e.preventDefault();
                setSelectedIndex(i => Math.max(0, i - 1));
              } else if (e.key === 'Enter' && filtered[selectedIndex]) {
                e.preventDefault();
                executeCommand(filtered[selectedIndex]);
              }
            }}
            sx={{ fontSize: '0.9375rem' }}
          />
          <Chip label="ESC" size="small" sx={{ fontSize: '0.6875rem', height: 20 }} />
        </Box>

        {/* Command list */}
        <List sx={{ maxHeight: 340, overflowY: 'auto', p: 1 }}>
          {filtered.length === 0 ? (
            <Box sx={{ py: 4, textAlign: 'center', color: 'text.secondary' }}>
              <HelpCircle size={24} style={{ opacity: 0.4, marginBottom: 8 }} />
              <Typography variant="body2">No matching commands</Typography>
            </Box>
          ) : (
            filtered.map((cmd, i) => (
              <ListItemButton
                key={cmd.id}
                selected={i === selectedIndex}
                onClick={() => executeCommand(cmd)}
                sx={{
                  borderRadius: 1.5,
                  py: 1,
                  px: 1.5,
                  mb: 0.5,
                  '&.Mui-selected': {
                    bgcolor: 'primary.light',
                    color: 'primary.main',
                  },
                }}
              >
                <ListItemIcon sx={{ minWidth: 32, color: 'inherit' }}>
                  {cmd.icon}
                </ListItemIcon>
                <ListItemText
                  primary={cmd.title}
                  secondary={cmd.category}
                  primaryTypographyProps={{ fontSize: '0.8125rem', fontWeight: 500 }}
                  secondaryTypographyProps={{ fontSize: '0.6875rem' }}
                />
                {cmd.shortcut && (
                  <Chip
                    label={cmd.shortcut}
                    size="small"
                    sx={{
                      fontSize: '0.6875rem',
                      height: 20,
                      fontFamily: 'monospace',
                      bgcolor: 'action.hover',
                    }}
                  />
                )}
              </ListItemButton>
            ))
          )}
        </List>
      </DialogContent>
    </Dialog>
  );
};
