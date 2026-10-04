import { createTheme, type Theme } from '@mui/material/styles';
import { palette, typography, motion, type ColorMode } from './tokens';

export function createAppTheme(mode: ColorMode): Theme {
  const p = palette[mode];

  return createTheme({
    palette: {
      mode,
      primary: {
        main: p.primary,
        light: p.primaryHover,
        dark: p.primary,
        contrastText: '#FFFFFF',
      },
      background: {
        default: p.canvas,
        paper: p.surface,
      },
      text: {
        primary: p.textPrimary,
        secondary: p.textSecondary,
      },
      divider: p.borderSubtle,
      warning: {
        main: p.warning,
      },
      error: {
        main: p.error,
      },
    },
    typography: {
      fontFamily: typography.fontUi,
      button: {
        textTransform: 'none',
        fontWeight: 500,
        fontSize: '0.8125rem', // 13px
      },
      body1: {
        fontSize: '0.875rem',
      },
      body2: {
        fontSize: '0.8125rem',
      },
      caption: {
        fontSize: '0.75rem',
      },
    },
    shape: {
      borderRadius: 6,
    },
    components: {
      MuiCssBaseline: {
        styleOverrides: {
          body: {
            backgroundColor: p.canvas,
            color: p.textPrimary,
            fontFamily: typography.fontUi,
            overflow: 'hidden',
            userSelect: 'none',
          },
        },
      },
      MuiButton: {
        defaultProps: {
          disableElevation: true,
        },
        styleOverrides: {
          root: {
            borderRadius: 6,
            padding: '4px 10px',
            lineHeight: 1.5,
            transition: `all ${motion.fast}`,
          },
          containedPrimary: {
            backgroundColor: p.primary,
            color: '#FFFFFF',
            '&:hover': {
              backgroundColor: p.primaryHover,
            },
          },
          outlined: {
            borderColor: p.borderStrong,
            color: p.textPrimary,
            '&:hover': {
              borderColor: p.primary,
              backgroundColor: p.primaryLight,
            },
          },
        },
      },
      MuiIconButton: {
        styleOverrides: {
          root: {
            borderRadius: 6,
            color: p.textSecondary,
            padding: 5,
            transition: `all ${motion.fast}`,
            '&:hover': {
              color: p.textPrimary,
              backgroundColor: p.surfaceHover,
            },
          },
        },
      },
      MuiChip: {
        styleOverrides: {
          root: {
            borderRadius: 5,
            fontWeight: 500,
            fontSize: '0.75rem',
            height: 24,
          },
        },
      },
      MuiTooltip: {
        styleOverrides: {
          tooltip: {
            backgroundColor: mode === 'dark' ? '#1F2937' : '#0F172A',
            color: '#F8FAFC',
            fontSize: '0.75rem',
            borderRadius: 5,
            boxShadow: '0 4px 12px rgba(0, 0, 0, 0.25)',
            border: `1px solid ${p.borderSubtle}`,
          },
        },
      },
      MuiMenu: {
        styleOverrides: {
          paper: {
            borderRadius: 8,
            backgroundColor: p.surface,
            border: `1px solid ${p.borderStrong}`,
            boxShadow: mode === 'dark'
              ? '0 10px 25px -5px rgba(0, 0, 0, 0.5), 0 8px 10px -6px rgba(0, 0, 0, 0.5)'
              : '0 10px 25px -5px rgba(0, 0, 0, 0.1), 0 8px 10px -6px rgba(0, 0, 0, 0.05)',
            backgroundImage: 'none',
          },
        },
      },
      MuiMenuItem: {
        styleOverrides: {
          root: {
            fontSize: '0.8125rem',
            padding: '6px 12px',
            borderRadius: 4,
            margin: '2px 4px',
            transition: `background-color ${motion.fast}`,
            '&:hover': {
              backgroundColor: p.surfaceHover,
            },
          },
        },
      },
      MuiDialog: {
        styleOverrides: {
          paper: {
            borderRadius: 10,
            backgroundColor: p.surface,
            border: `1px solid ${p.borderStrong}`,
            boxShadow: mode === 'dark'
              ? '0 20px 25px -5px rgba(0, 0, 0, 0.6), 0 10px 10px -5px rgba(0, 0, 0, 0.5)'
              : '0 20px 25px -5px rgba(0, 0, 0, 0.1), 0 10px 10px -5px rgba(0, 0, 0, 0.04)',
            backgroundImage: 'none',
          },
        },
      },
      MuiTabs: {
        styleOverrides: {
          root: {
            minHeight: 34,
          },
          indicator: {
            backgroundColor: p.primary,
            height: 2,
            borderRadius: '2px 2px 0 0',
          },
        },
      },
      MuiTab: {
        styleOverrides: {
          root: {
            minHeight: 34,
            padding: '6px 14px',
            textTransform: 'none',
            fontWeight: 500,
            fontSize: '0.8125rem',
            color: p.textSecondary,
            '&.Mui-selected': {
              color: p.primary,
              fontWeight: 600,
            },
          },
        },
      },
      MuiDivider: {
        styleOverrides: {
          root: {
            borderColor: p.borderSubtle,
          },
        },
      },
    },
  });
}
