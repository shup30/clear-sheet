import React from 'react';
import { Box, Typography, Button, Paper } from '@mui/material';
import { FileSpreadsheet, Upload, Clock, Sparkles } from 'lucide-react';

interface EmptyStateProps {
  busy: boolean;
  recents: string[];
  onOpen: (path?: string) => void;
  onClearRecent: () => void;
}

export const EmptyState: React.FC<EmptyStateProps> = ({
  busy,
  recents,
  onOpen,
  onClearRecent,
}) => {
  return (
    <main className="welcome">
      <Box sx={{ maxWidth: 540, width: '100%', textAlign: 'center', py: 6 }}>
        {/* Animated Brand Badge */}
        <Box
          sx={{
            display: 'inline-flex',
            alignItems: 'center',
            gap: 1,
            px: 1.5,
            py: 0.5,
            mb: 3,
            borderRadius: 5,
            bgcolor: 'action.hover',
            border: 1,
            borderColor: 'divider',
          }}
        >
          <Sparkles size={14} color="#0D9488" />
          <Typography variant="caption" sx={{ fontWeight: 600, color: 'text.secondary', letterSpacing: 0.5 }}>
            HIGH-PERFORMANCE DATA VIEWER
          </Typography>
        </Box>

        {/* Hero Drop Container */}
        <Paper
          elevation={0}
          className="empty-drop-zone"
          onClick={() => void onOpen()}
          sx={{
            p: 5,
            borderRadius: 3,
            border: '2px dashed',
            borderColor: 'divider',
            bgcolor: 'background.paper',
            cursor: 'pointer',
            transition: 'all 0.2s cubic-bezier(0.16, 1, 0.3, 1)',
            '&:hover': {
              borderColor: 'primary.main',
              transform: 'translateY(-2px)',
              boxShadow: '0 12px 30px rgba(13, 148, 136, 0.12)',
            },
          }}
        >
          <Box
            sx={{
              width: 56,
              height: 56,
              borderRadius: 3,
              bgcolor: 'primary.light',
              color: 'primary.main',
              display: 'grid',
              placeItems: 'center',
              margin: '0 auto 16px',
            }}
          >
            <FileSpreadsheet size={28} />
          </Box>

          <Typography variant="h5" fontWeight={700} sx={{ mb: 1, letterSpacing: '-0.02em' }}>
            Open or drop a spreadsheet
          </Typography>

          <Typography variant="body2" color="text.secondary" sx={{ mb: 3 }}>
            Drop any workbook here or browse files on your computer.
          </Typography>

          <Button
            variant="contained"
            size="medium"
            disabled={busy}
            onClick={(e) => {
              e.stopPropagation();
              void onOpen();
            }}
            startIcon={<Upload size={16} />}
            sx={{ px: 3, py: 1, borderRadius: 2 }}
          >
            Browse files
          </Button>

          <Typography
            variant="caption"
            display="block"
            color="text.secondary"
            sx={{ mt: 3, fontSize: '0.6875rem', letterSpacing: 0.5 }}
          >
            XLSX · XLS · XLSM · XLSB · CSV · UP TO 200 MB
          </Typography>
        </Paper>

        {/* Recent files section */}
        <section className="recent" style={{ marginTop: 32, textAlign: 'left' }}>
          <Box display="flex" justifyContent="space-between" alignItems="center" sx={{ mb: 1.5 }}>
            <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
              <Clock size={15} style={{ opacity: 0.6 }} />
              <Typography variant="subtitle2" fontWeight={600}>
                Recent files
              </Typography>
            </Box>
            {recents.length > 0 && (
              <Button size="small" variant="text" onClick={onClearRecent} sx={{ fontSize: '0.75rem' }}>
                Clear history
              </Button>
            )}
          </Box>

          {recents.length > 0 ? (
            <Box sx={{ display: 'flex', flexDirection: 'column', gap: 0.75 }}>
              {recents.map((path) => (
                <Button
                  className="recent-file"
                  key={path}
                  disabled={busy}
                  onClick={() => void onOpen(path)}
                  title={path}
                  variant="outlined"
                  sx={{
                    justifyContent: 'flex-start',
                    py: 1,
                    px: 1.5,
                    borderRadius: 2,
                    textTransform: 'none',
                    bgcolor: 'background.paper',
                    borderColor: 'divider',
                  }}
                >
                  <FileSpreadsheet size={15} style={{ marginRight: 10, opacity: 0.6, flexShrink: 0 }} />
                  <Typography variant="body2" noWrap sx={{ fontSize: '0.8125rem' }}>
                    {path}
                  </Typography>
                </Button>
              ))}
            </Box>
          ) : (
            <Paper variant="outlined" sx={{ p: 2.5, textAlign: 'center', borderRadius: 2, bgcolor: 'transparent' }}>
              <Typography color="text.secondary" variant="body2">
                Opened files will appear here.
              </Typography>
            </Paper>
          )}
        </section>
      </Box>
    </main>
  );
};
