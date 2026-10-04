import React from 'react';
import { Box, Typography, Button, Chip, Tooltip, IconButton } from '@mui/material';
import { FolderOpen, Save, FileSpreadsheet, Search, Sun, Moon } from 'lucide-react';
import type { ColorMode } from '../theme/tokens';

interface TitleBarProps {
  bookName: string | null;
  filePath: string | null;
  dirty: boolean;
  busy: boolean;
  mode: ColorMode;
  onToggleTheme: () => void;
  onOpen: () => void;
  onSave: () => void;
  onSaveAs: () => void;
  onOpenCommandPalette: () => void;
}

export const TitleBar: React.FC<TitleBarProps> = ({
  bookName,
  filePath,
  dirty,
  busy,
  mode,
  onToggleTheme,
  onOpen,
  onSave,
  onSaveAs,
  onOpenCommandPalette,
}) => {
  return (
    <header className="title-bar">
      <Box className="title-bar-left" sx={{ display: 'flex', alignItems: 'center', gap: 1.5 }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <FileSpreadsheet size={18} className="app-brand-icon" />
          <Typography variant="subtitle2" fontWeight={700} sx={{ letterSpacing: '-0.01em', whiteSpace: 'nowrap' }}>
            Clear Sheet
          </Typography>
        </Box>

        {bookName && (
          <>
            <Box sx={{ width: '1px', height: 16, bgcolor: 'divider' }} />
            <Tooltip title={filePath ?? bookName} arrow>
              <Typography
                variant="body2"
                fontWeight={600}
                noWrap
                sx={{
                  maxWidth: 240,
                  color: 'text.primary',
                  cursor: 'default',
                  fontSize: '0.8125rem',
                }}
              >
                {bookName}
              </Typography>
            </Tooltip>
          </>
        )}

        {dirty ? (
          <Chip
            size="small"
            label="● Unsaved changes"
            color="warning"
            variant="outlined"
            sx={{
              height: 22,
              fontSize: '0.6875rem',
              fontWeight: 600,
              borderWidth: '1px',
              animation: 'pulseGlow 2s infinite',
            }}
          />
        ) : (
          <Chip
            size="small"
            label="Saved"
            variant="outlined"
            sx={{
              height: 22,
              fontSize: '0.6875rem',
              opacity: 0.75,
            }}
          />
        )}
      </Box>

      {/* Center: Command Palette Pill */}
      <Box className="title-bar-center" sx={{ display: 'flex', justifyContent: 'center', flex: 1 }}>
        <Button
          variant="outlined"
          onClick={onOpenCommandPalette}
          className="cmd-palette-pill"
          sx={{
            py: 0.25,
            px: 1.5,
            height: 26,
            minWidth: 180,
            borderRadius: '13px',
            borderColor: 'divider',
            color: 'text.secondary',
            bgcolor: mode === 'dark' ? 'rgba(255, 255, 255, 0.04)' : 'rgba(0, 0, 0, 0.03)',
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'space-between',
            textTransform: 'none',
            fontSize: '0.75rem',
            '&:hover': {
              borderColor: 'primary.main',
              bgcolor: mode === 'dark' ? 'rgba(255, 255, 255, 0.07)' : 'rgba(0, 0, 0, 0.05)',
              color: 'text.primary',
            },
          }}
        >
          <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.75 }}>
            <Search size={13} />
            <span>Search or command…</span>
          </Box>
          <Box
            component="span"
            sx={{
              fontSize: '0.6875rem',
              px: 0.6,
              py: 0.1,
              borderRadius: '3px',
              bgcolor: mode === 'dark' ? 'rgba(255, 255, 255, 0.1)' : 'rgba(0, 0, 0, 0.07)',
              fontFamily: 'monospace',
              letterSpacing: 0.5,
            }}
          >
            Ctrl+K
          </Box>
        </Button>
      </Box>

      {/* Right: Actions */}
      <Box className="title-bar-right" sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
        <Tooltip title={`Switch to ${mode === 'light' ? 'Dark' : 'Light'} mode`}>
          <IconButton size="small" onClick={onToggleTheme} aria-label="Toggle theme">
            {mode === 'light' ? <Moon size={16} /> : <Sun size={16} />}
          </IconButton>
        </Tooltip>

        <Button
          variant="outlined"
          size="small"
          disabled={busy}
          onClick={onOpen}
          startIcon={<FolderOpen size={14} />}
          sx={{ height: 28 }}
        >
          Open file
        </Button>

        <Button
          variant="contained"
          size="small"
          disabled={!bookName || busy}
          onClick={onSave}
          startIcon={<Save size={14} />}
          sx={{ height: 28 }}
        >
          Save
        </Button>

        <Button
          variant="outlined"
          size="small"
          disabled={!bookName || busy}
          onClick={onSaveAs}
          sx={{ height: 28 }}
        >
          Save As
        </Button>
      </Box>
    </header>
  );
};
