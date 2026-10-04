import React, { useState, useMemo } from 'react';
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Tabs,
  Tab,
  Box,
  Typography,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  Checkbox,
  FormControlLabel,
  Paper,
} from '@mui/material';
import type { CellStyle } from '../../electron/types';
import { formatValue } from '../workbook/format';

interface FormatCellsDialogProps {
  open: boolean;
  activeValue: unknown;
  currentStyle?: CellStyle;
  currentNumFmt?: string;
  onClose: () => void;
  onApply: (patch: Partial<CellStyle>, numFmt?: string) => void;
}

const COLOR_SWATCHES = [
  '#000000', '#ffffff', '#334155', '#64748b', '#94a3b8', '#e2e8f0',
  '#2563eb', '#38bdf8', '#0284c7', '#0d9488', '#10b981', '#4ade80',
  '#d97706', '#f59e0b', '#fef08a', '#dc2626', '#f87171', '#9333ea',
];

const BG_SWATCHES = [
  '', '#f8fafc', '#f1f5f9', '#e2e8f0', '#dbeafe', '#e0f2fe',
  '#ccfbf1', '#dcfce7', '#fef9c3', '#fef3c7', '#fee2e2', '#f3e8ff',
  '#1e293b', '#0f172a', '#1e3a8a', '#064e3b', '#78350f', '#4c0519',
];

export const FormatCellsDialog: React.FC<FormatCellsDialogProps> = ({
  open,
  activeValue,
  currentStyle,
  currentNumFmt,
  onClose,
  onApply,
}) => {
  const [tab, setTab] = useState(0);

  // Number Tab State
  const [numCategory, setNumCategory] = useState<'general' | 'number' | 'currency' | 'percent' | 'date' | 'text'>(() => {
    if (!currentNumFmt || currentNumFmt === 'General') return 'general';
    if (currentNumFmt.includes('$')) return 'currency';
    if (currentNumFmt.includes('%')) return 'percent';
    if (currentNumFmt.includes('yy') || currentNumFmt.includes('mm') || currentNumFmt.includes('dd')) return 'date';
    if (currentNumFmt === '@') return 'text';
    return 'number';
  });
  const [decimals, setDecimals] = useState(2);
  const [thousandsSep, setThousandsSep] = useState(true);

  // Alignment Tab State
  const [hAlign, setHAlign] = useState<'left' | 'center' | 'right' | undefined>(currentStyle?.hAlign);
  const [vAlign, setVAlign] = useState<'top' | 'middle' | 'bottom' | undefined>(currentStyle?.vAlign);
  const [wrap, setWrap] = useState<boolean>(!!currentStyle?.wrap);

  // Font Tab State
  const [fontName, setFontName] = useState<string>(currentStyle?.fontName || 'Inter Variable');
  const [fontSize, setFontSize] = useState<number>(currentStyle?.fontSize || 13);
  const [bold, setBold] = useState<boolean>(!!currentStyle?.bold);
  const [italic, setItalic] = useState<boolean>(!!currentStyle?.italic);
  const [underline, setUnderline] = useState<boolean>(!!currentStyle?.underline);
  const [color, setColor] = useState<string>(currentStyle?.color || '');

  // Fill Tab State
  const [bg, setBg] = useState<string>(currentStyle?.bg || '');

  // Computed Number Format String
  const computedNumFmt = useMemo(() => {
    switch (numCategory) {
      case 'general':
        return 'General';
      case 'number': {
        const dec = decimals > 0 ? '.' + '0'.repeat(decimals) : '';
        return thousandsSep ? `#,##0${dec}` : `0${dec}`;
      }
      case 'currency': {
        const dec = decimals > 0 ? '.' + '0'.repeat(decimals) : '';
        return `$#,##0${dec}`;
      }
      case 'percent': {
        const dec = decimals > 0 ? '.' + '0'.repeat(decimals) : '';
        return `0${dec}%`;
      }
      case 'date':
        return 'yyyy-mm-dd';
      case 'text':
        return '@';
      default:
        return 'General';
    }
  }, [numCategory, decimals, thousandsSep]);

  // Preview formatted text
  const previewText = useMemo(() => {
    const raw = typeof activeValue === 'number' ? activeValue : 1234.56;
    try {
      return formatValue(raw, computedNumFmt);
    } catch {
      return String(raw);
    }
  }, [activeValue, computedNumFmt]);

  const handleApply = () => {
    const patch: Partial<CellStyle> = {
      ...(bold !== undefined ? { bold } : {}),
      ...(italic !== undefined ? { italic } : {}),
      ...(underline !== undefined ? { underline } : {}),
      ...(fontName ? { fontName } : {}),
      ...(fontSize ? { fontSize } : {}),
      ...(color ? { color } : {}),
      ...(bg ? { bg } : {}),
      ...(hAlign ? { hAlign } : {}),
      ...(vAlign ? { vAlign } : {}),
      ...(wrap !== undefined ? { wrap } : {}),
    };
    onApply(patch, computedNumFmt);
    onClose();
  };

  return (
    <Dialog open={open} onClose={onClose} maxWidth="sm" fullWidth>
      <DialogTitle sx={{ pb: 1, fontSize: '1.05rem', fontWeight: 600 }}>Format Cells (Ctrl+1)</DialogTitle>
      <Tabs value={tab} onChange={(_, val) => setTab(val)} sx={{ px: 3, borderBottom: 1, borderColor: 'divider' }}>
        <Tab label="Number" sx={{ minWidth: 80, textTransform: 'none', fontWeight: 600 }} />
        <Tab label="Alignment" sx={{ minWidth: 80, textTransform: 'none', fontWeight: 600 }} />
        <Tab label="Font" sx={{ minWidth: 80, textTransform: 'none', fontWeight: 600 }} />
        <Tab label="Fill" sx={{ minWidth: 80, textTransform: 'none', fontWeight: 600 }} />
      </Tabs>

      <DialogContent sx={{ minHeight: 280, pt: 2.5 }}>
        {/* Tab 0: Number */}
        {tab === 0 && (
          <Box sx={{ display: 'flex', gap: 3 }}>
            <Box sx={{ width: 140, borderRight: 1, borderColor: 'divider', pr: 2 }}>
              <Typography variant="caption" sx={{ fontWeight: 600, color: 'text.secondary', display: 'block', mb: 1 }}>
                Category:
              </Typography>
              {(['general', 'number', 'currency', 'percent', 'date', 'text'] as const).map((cat) => (
                <Box
                  key={cat}
                  onClick={() => setNumCategory(cat)}
                  sx={{
                    px: 1.5,
                    py: 0.6,
                    borderRadius: 1,
                    cursor: 'pointer',
                    textTransform: 'capitalize',
                    fontSize: '0.85rem',
                    fontWeight: numCategory === cat ? 600 : 400,
                    bgcolor: numCategory === cat ? 'action.selected' : 'transparent',
                    '&:hover': { bgcolor: 'action.hover' },
                  }}
                >
                  {cat}
                </Box>
              ))}
            </Box>

            <Box sx={{ flex: 1, display: 'flex', flexDirection: 'column', gap: 2 }}>
              {numCategory !== 'general' && numCategory !== 'text' && numCategory !== 'date' && (
                <>
                  <FormControl size="small" sx={{ width: 160 }}>
                    <InputLabel>Decimal places</InputLabel>
                    <Select value={decimals} label="Decimal places" onChange={(e) => setDecimals(Number(e.target.value))}>
                      {[0, 1, 2, 3, 4, 5, 6].map((d) => (
                        <MenuItem key={d} value={d}>{d}</MenuItem>
                      ))}
                    </Select>
                  </FormControl>

                  {numCategory === 'number' && (
                    <FormControlLabel
                      control={<Checkbox checked={thousandsSep} onChange={(e) => setThousandsSep(e.target.checked)} size="small" />}
                      label="Use 1000 Separator (,)"
                    />
                  )}
                </>
              )}

              {/* Sample Box */}
              <Box sx={{ mt: 'auto' }}>
                <Typography variant="caption" sx={{ fontWeight: 600, color: 'text.secondary' }}>
                  Sample
                </Typography>
                <Paper
                  variant="outlined"
                  sx={{
                    p: 1.5,
                    mt: 0.5,
                    fontFamily: '"JetBrains Mono", monospace',
                    bgcolor: 'background.default',
                    textAlign: 'right',
                    fontSize: '0.95rem',
                  }}
                >
                  {previewText}
                </Paper>
              </Box>
            </Box>
          </Box>
        )}

        {/* Tab 1: Alignment */}
        {tab === 1 && (
          <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2.5 }}>
            <Box sx={{ display: 'flex', gap: 2 }}>
              <FormControl size="small" sx={{ flex: 1 }}>
                <InputLabel>Horizontal</InputLabel>
                <Select
                  value={hAlign || 'left'}
                  label="Horizontal"
                  onChange={(e) => setHAlign(e.target.value as 'left' | 'center' | 'right')}
                >
                  <MenuItem value="left">Left</MenuItem>
                  <MenuItem value="center">Center</MenuItem>
                  <MenuItem value="right">Right</MenuItem>
                </Select>
              </FormControl>

              <FormControl size="small" sx={{ flex: 1 }}>
                <InputLabel>Vertical</InputLabel>
                <Select
                  value={vAlign || 'middle'}
                  label="Vertical"
                  onChange={(e) => setVAlign(e.target.value as 'top' | 'middle' | 'bottom')}
                >
                  <MenuItem value="top">Top</MenuItem>
                  <MenuItem value="middle">Middle</MenuItem>
                  <MenuItem value="bottom">Bottom</MenuItem>
                </Select>
              </FormControl>
            </Box>

            <FormControlLabel
              control={<Checkbox checked={wrap} onChange={(e) => setWrap(e.target.checked)} size="small" />}
              label="Wrap text"
            />
          </Box>
        )}

        {/* Tab 2: Font */}
        {tab === 2 && (
          <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            <Box sx={{ display: 'flex', gap: 2 }}>
              <FormControl size="small" sx={{ flex: 2 }}>
                <InputLabel>Font</InputLabel>
                <Select value={fontName} label="Font" onChange={(e) => setFontName(e.target.value)}>
                  <MenuItem value="Inter Variable">Inter Variable</MenuItem>
                  <MenuItem value="JetBrains Mono">JetBrains Mono</MenuItem>
                  <MenuItem value="Arial">Arial</MenuItem>
                  <MenuItem value="Segoe UI">Segoe UI</MenuItem>
                  <MenuItem value="Calibri">Calibri</MenuItem>
                  <MenuItem value="Georgia">Georgia</MenuItem>
                </Select>
              </FormControl>

              <FormControl size="small" sx={{ flex: 1 }}>
                <InputLabel>Size</InputLabel>
                <Select value={fontSize} label="Size" onChange={(e) => setFontSize(Number(e.target.value))}>
                  {[9, 10, 11, 12, 13, 14, 16, 18, 20, 24].map((s) => (
                    <MenuItem key={s} value={s}>{s} px</MenuItem>
                  ))}
                </Select>
              </FormControl>
            </Box>

            <Box sx={{ display: 'flex', gap: 2, alignItems: 'center' }}>
              <FormControlLabel
                control={<Checkbox checked={bold} onChange={(e) => setBold(e.target.checked)} size="small" />}
                label="Bold"
              />
              <FormControlLabel
                control={<Checkbox checked={italic} onChange={(e) => setItalic(e.target.checked)} size="small" />}
                label="Italic"
              />
              <FormControlLabel
                control={<Checkbox checked={underline} onChange={(e) => setUnderline(e.target.checked)} size="small" />}
                label="Underline"
              />
            </Box>

            <Box>
              <Typography variant="caption" sx={{ fontWeight: 600, color: 'text.secondary', display: 'block', mb: 0.75 }}>
                Text Color
              </Typography>
              <Box sx={{ display: 'flex', gap: 0.75, flexWrap: 'wrap' }}>
                {COLOR_SWATCHES.map((swatch) => (
                  <Box
                    key={swatch}
                    onClick={() => setColor(swatch)}
                    sx={{
                      width: 24,
                      height: 24,
                      borderRadius: '50%',
                      bgcolor: swatch,
                      cursor: 'pointer',
                      border: color === swatch ? '2px solid #2563eb' : '1px solid rgba(0,0,0,0.2)',
                      boxShadow: color === swatch ? '0 0 0 2px rgba(37,99,235,0.4)' : 'none',
                    }}
                  />
                ))}
              </Box>
            </Box>
          </Box>
        )}

        {/* Tab 3: Fill */}
        {tab === 3 && (
          <Box sx={{ display: 'flex', flexDirection: 'column', gap: 2 }}>
            <Typography variant="caption" sx={{ fontWeight: 600, color: 'text.secondary' }}>
              Background Color
            </Typography>
            <Box sx={{ display: 'flex', gap: 1, flexWrap: 'wrap' }}>
              {BG_SWATCHES.map((swatch, idx) => (
                <Box
                  key={idx}
                  onClick={() => setBg(swatch)}
                  sx={{
                    width: 32,
                    height: 32,
                    borderRadius: 1,
                    bgcolor: swatch || 'transparent',
                    cursor: 'pointer',
                    border: bg === swatch ? '2px solid #2563eb' : '1px solid rgba(128,128,128,0.3)',
                    display: 'flex',
                    alignItems: 'center',
                    justifyContent: 'center',
                    fontSize: '0.65rem',
                    color: 'text.secondary',
                  }}
                >
                  {!swatch && 'None'}
                </Box>
              ))}
            </Box>
          </Box>
        )}

        {/* Live Preview Box */}
        <Box sx={{ mt: 3, pt: 2, borderTop: 1, borderColor: 'divider' }}>
          <Typography variant="caption" sx={{ fontWeight: 600, color: 'text.secondary', display: 'block', mb: 0.5 }}>
            Live Preview
          </Typography>
          <Box
            sx={{
              p: 1.5,
              borderRadius: 1,
              border: '1px solid',
              borderColor: 'divider',
              fontFamily: fontName,
              fontSize: `${fontSize}px`,
              fontWeight: bold ? 700 : 400,
              fontStyle: italic ? 'italic' : 'normal',
              textDecoration: underline ? 'underline' : 'none',
              color: color || 'inherit',
              bgcolor: bg || 'transparent',
              textAlign: hAlign || 'left',
              whiteSpace: wrap ? 'normal' : 'nowrap',
              height: 48,
              display: 'flex',
              alignItems: vAlign === 'top' ? 'flex-start' : vAlign === 'bottom' ? 'flex-end' : 'center',
              justifyContent: hAlign === 'center' ? 'center' : hAlign === 'right' ? 'flex-end' : 'flex-start',
            }}
          >
            {previewText}
          </Box>
        </Box>
      </DialogContent>

      <DialogActions sx={{ px: 3, py: 2 }}>
        <Button onClick={onClose} color="inherit">Cancel</Button>
        <Button onClick={handleApply} variant="contained" color="primary">Apply</Button>
      </DialogActions>
    </Dialog>
  );
};
