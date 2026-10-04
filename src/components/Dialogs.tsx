import React from 'react';
import {
  Dialog, DialogTitle, DialogContent, DialogActions, TextField,
  Box, FormControl, InputLabel, Select, MenuItem, FormControlLabel,
  Checkbox, Typography, List, ListItemButton, Button
} from '@mui/material';
import { colToLetter } from '../workbook/cellRef';

export interface FoundItem {
  sheet: string;
  r: number;
  c: number;
  text: string;
}

interface FindDialogProps {
  open: boolean;
  onClose: () => void;
  findText: string;
  onFindTextChange: (val: string) => void;
  replaceText: string;
  onReplaceTextChange: (val: string) => void;
  findScope: 'sheet' | 'book';
  onFindScopeChange: (scope: 'sheet' | 'book') => void;
  matchCase: boolean;
  onMatchCaseChange: (match: boolean) => void;
  found: FoundItem[];
  foundIdx: number;
  onFindNav: (dir: 1 | -1) => void;
  onFindAll: () => void;
  onReplace: (all: boolean) => void;
  onGotoFound: (item: FoundItem, index: number) => void;
}

export const FindDialog: React.FC<FindDialogProps> = ({
  open,
  onClose,
  findText,
  onFindTextChange,
  replaceText,
  onReplaceTextChange,
  findScope,
  onFindScopeChange,
  matchCase,
  onMatchCaseChange,
  found,
  foundIdx,
  onFindNav,
  onFindAll,
  onReplace,
  onGotoFound,
}) => {
  return (
    <Dialog open={open} onClose={onClose} maxWidth="sm" fullWidth>
      <DialogTitle sx={{ pb: 1, fontSize: '1rem', fontWeight: 600 }}>Find & Replace</DialogTitle>
      <DialogContent sx={{ display: 'flex', flexDirection: 'column', gap: 2, pt: 1 }}>
        <TextField
          label="Find"
          value={findText}
          onChange={(e) => onFindTextChange(e.target.value)}
          autoFocus
          fullWidth
          size="small"
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              onFindNav(1);
            }
          }}
        />
        <TextField
          label="Replace with"
          value={replaceText}
          onChange={(e) => onReplaceTextChange(e.target.value)}
          fullWidth
          size="small"
        />
        <Box display="flex" gap={2} alignItems="center">
          <FormControl size="small" sx={{ minWidth: 140 }}>
            <InputLabel>Scope</InputLabel>
            <Select
              value={findScope}
              label="Scope"
              onChange={(e) => onFindScopeChange(e.target.value === 'book' ? 'book' : 'sheet')}
            >
              <MenuItem value="sheet">Current sheet</MenuItem>
              <MenuItem value="book">Workbook</MenuItem>
            </Select>
          </FormControl>
          <FormControlLabel
            control={<Checkbox checked={matchCase} onChange={(e) => onMatchCaseChange(e.target.checked)} size="small" />}
            label={<Typography variant="body2">Match case</Typography>}
          />
        </Box>
        {found.length > 0 && (
          <Typography variant="caption" sx={{ color: 'text.secondary', fontWeight: 500 }}>
            {foundIdx + 1} / {found.length} matches
          </Typography>
        )}
        {found.length > 0 && (
          <List dense sx={{ maxHeight: 180, overflow: 'auto', border: 1, borderColor: 'divider', borderRadius: 1.5, p: 0.5 }}>
            {found.slice(0, 500).map((f, i) => (
              <ListItemButton
                key={`${f.sheet}:${f.r}:${f.c}`}
                selected={i === foundIdx}
                onClick={() => onGotoFound(f, i)}
                sx={{ borderRadius: 1, py: 0.5 }}
              >
                <Typography variant="body2" sx={{ fontFamily: 'monospace', fontSize: '0.75rem' }}>
                  <b>{f.sheet}!{colToLetter(f.c)}{f.r + 1}</b>: {f.text.slice(0, 75)}
                </Typography>
              </ListItemButton>
            ))}
          </List>
        )}
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 2 }}>
        <Button onClick={() => onFindNav(-1)} size="small">Find prev</Button>
        <Button onClick={() => onFindNav(1)} size="small">Find next</Button>
        <Button onClick={onFindAll} size="small">Find all</Button>
        <Button onClick={() => onReplace(false)} size="small">Replace</Button>
        <Button onClick={() => onReplace(true)} size="small">Replace all</Button>
        <Button onClick={onClose} size="small">Close</Button>
      </DialogActions>
    </Dialog>
  );
};

interface RenameDialogProps {
  open: boolean;
  value: string;
  onChange: (val: string) => void;
  onConfirm: () => void;
  onCancel: () => void;
}

export const RenameDialog: React.FC<RenameDialogProps> = ({
  open,
  value,
  onChange,
  onConfirm,
  onCancel,
}) => {
  return (
    <Dialog open={open} onClose={onCancel} maxWidth="xs" fullWidth>
      <DialogTitle sx={{ pb: 1, fontSize: '1rem', fontWeight: 600 }}>Rename worksheet</DialogTitle>
      <DialogContent>
        <TextField
          label="Sheet name"
          value={value}
          onChange={(e) => onChange(e.target.value)}
          fullWidth
          autoFocus
          size="small"
          sx={{ mt: 1 }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              onConfirm();
            }
          }}
        />
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 2 }}>
        <Button onClick={onCancel} size="small">Cancel</Button>
        <Button variant="contained" onClick={onConfirm} size="small">Rename</Button>
      </DialogActions>
    </Dialog>
  );
};

interface DeleteSheetDialogProps {
  open: boolean;
  sheetName: string | null;
  onConfirm: () => void;
  onCancel: () => void;
}

export const DeleteSheetDialog: React.FC<DeleteSheetDialogProps> = ({
  open,
  sheetName,
  onConfirm,
  onCancel,
}) => {
  return (
    <Dialog open={open} onClose={onCancel} maxWidth="xs" fullWidth>
      <DialogTitle sx={{ pb: 1, fontSize: '1rem', fontWeight: 600 }}>Delete worksheet?</DialogTitle>
      <DialogContent>
        <Typography variant="body2">
          Delete worksheet “{sheetName}”? You can undo this while the workbook remains open.
        </Typography>
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 2 }}>
        <Button onClick={onCancel} size="small">Cancel</Button>
        <Button variant="contained" color="error" onClick={onConfirm} size="small">Delete</Button>
      </DialogActions>
    </Dialog>
  );
};

interface DimensionDialogProps {
  kind: 'colWidth' | 'rowHeight' | null;
  value: string;
  onChange: (val: string) => void;
  onConfirm: () => void;
  onCancel: () => void;
}

export const DimensionDialog: React.FC<DimensionDialogProps> = ({
  kind,
  value,
  onChange,
  onConfirm,
  onCancel,
}) => {
  return (
    <Dialog open={kind !== null} onClose={onCancel} maxWidth="xs" fullWidth>
      <DialogTitle sx={{ pb: 1, fontSize: '1rem', fontWeight: 600 }}>
        {kind === 'colWidth' ? 'Column width (characters)' : 'Row height (pixels)'}
      </DialogTitle>
      <DialogContent>
        <TextField
          label={kind === 'colWidth' ? 'Width' : 'Height'}
          value={value}
          onChange={(e) => onChange(e.target.value)}
          fullWidth
          autoFocus
          size="small"
          inputProps={{ inputMode: 'numeric' }}
          sx={{ mt: 1 }}
          onKeyDown={(e) => {
            if (e.key === 'Enter') {
              e.preventDefault();
              onConfirm();
            }
          }}
        />
      </DialogContent>
      <DialogActions sx={{ px: 3, pb: 2 }}>
        <Button onClick={onCancel} size="small">Cancel</Button>
        <Button variant="contained" onClick={onConfirm} size="small">Apply</Button>
      </DialogActions>
    </Dialog>
  );
};
