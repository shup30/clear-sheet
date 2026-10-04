import React, { useState } from 'react';
import { Box, Tabs, Tab, Button, Menu, MenuItem, Divider } from '@mui/material';
import { Plus, ChevronDown, Copy, Edit3, ArrowLeft, ArrowRight, Trash2 } from 'lucide-react';

interface SheetTabsProps {
  sheetNames: string[];
  activeSheet: string | null;
  busy: boolean;
  onSelectSheet: (name: string) => void;
  onAddSheet: () => void;
  onRenameSheet: (name: string) => void;
  onDuplicateSheet: (name: string) => void;
  onMoveSheet: (name: string, direction: 'left' | 'right') => void;
  onDeleteSheet: (name: string) => void;
}

export const SheetTabs: React.FC<SheetTabsProps> = ({
  sheetNames,
  activeSheet,
  busy,
  onSelectSheet,
  onAddSheet,
  onRenameSheet,
  onDuplicateSheet,
  onMoveSheet,
  onDeleteSheet,
}) => {
  const [sheetMenu, setSheetMenu] = useState<null | HTMLElement>(null);

  const handleMenuClose = () => setSheetMenu(null);

  return (
    <Box className="sheet-tabs-bar" sx={{ display: 'flex', alignItems: 'center', px: 1, borderTop: 1, borderColor: 'divider', bgcolor: 'background.paper', minHeight: 35 }}>
      <Tabs
        value={activeSheet ?? false}
        onChange={(_, name) => onSelectSheet(name)}
        variant="scrollable"
        scrollButtons="auto"
        sx={{
          flex: 1,
          minHeight: 34,
          '& .MuiTabs-scrollButtons': {
            width: 24,
            height: 34,
            '&.Mui-disabled': { opacity: 0.3 },
          },
        }}
      >
        {sheetNames.map((name) => (
          <Tab
            key={name}
            label={
              <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.75 }}>
                <Box
                  sx={{
                    width: 6,
                    height: 6,
                    borderRadius: '50%',
                    bgcolor: activeSheet === name ? 'primary.main' : 'transparent',
                    transition: 'all 0.15s ease',
                  }}
                />
                <span>{name}</span>
              </Box>
            }
            value={name}
            disabled={busy}
            sx={{
              minHeight: 34,
              py: 0.5,
              px: 1.5,
              fontSize: '0.8125rem',
              fontWeight: activeSheet === name ? 600 : 500,
            }}
          />
        ))}
      </Tabs>

      <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.5, pl: 1 }}>
        <Button
          size="small"
          aria-label="Add sheet"
          onClick={onAddSheet}
          disabled={busy}
          startIcon={<Plus size={14} />}
          sx={{ height: 26, fontSize: '0.75rem', px: 1 }}
        >
          + Sheet
        </Button>

        <Button
          size="small"
          aria-label="Sheet menu"
          disabled={!activeSheet || busy}
          onClick={(e) => setSheetMenu(e.currentTarget)}
          endIcon={<ChevronDown size={14} />}
          sx={{ height: 26, fontSize: '0.75rem', px: 1 }}
        >
          Sheet ▾
        </Button>
      </Box>

      <Menu
        anchorEl={sheetMenu}
        open={!!sheetMenu}
        onClose={handleMenuClose}
        anchorOrigin={{ vertical: 'top', horizontal: 'right' }}
        transformOrigin={{ vertical: 'bottom', horizontal: 'right' }}
      >
        <MenuItem
          onClick={() => {
            handleMenuClose();
            if (activeSheet) onRenameSheet(activeSheet);
          }}
        >
          <Edit3 size={14} style={{ marginRight: 8, opacity: 0.7 }} />
          Rename…
        </MenuItem>
        <MenuItem
          onClick={() => {
            handleMenuClose();
            if (activeSheet) onDuplicateSheet(activeSheet);
          }}
        >
          <Copy size={14} style={{ marginRight: 8, opacity: 0.7 }} />
          Duplicate
        </MenuItem>
        <MenuItem
          onClick={() => {
            handleMenuClose();
            if (activeSheet) onMoveSheet(activeSheet, 'left');
          }}
        >
          <ArrowLeft size={14} style={{ marginRight: 8, opacity: 0.7 }} />
          Move left
        </MenuItem>
        <MenuItem
          onClick={() => {
            handleMenuClose();
            if (activeSheet) onMoveSheet(activeSheet, 'right');
          }}
        >
          <ArrowRight size={14} style={{ marginRight: 8, opacity: 0.7 }} />
          Move right
        </MenuItem>
        <Divider sx={{ my: 0.5 }} />
        <MenuItem
          onClick={() => {
            handleMenuClose();
            if (activeSheet) onDeleteSheet(activeSheet);
          }}
          sx={{ color: 'error.main' }}
        >
          <Trash2 size={14} style={{ marginRight: 8, opacity: 0.7 }} />
          Delete…
        </MenuItem>
      </Menu>
    </Box>
  );
};
