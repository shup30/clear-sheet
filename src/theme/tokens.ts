import { themeQuartz, colorSchemeDark, colorSchemeLight } from 'ag-grid-community';

export type ColorMode = 'light' | 'dark';

export const palette = {
  light: {
    mode: 'light' as const,
    canvas: '#F8FAFC', // Slate 50
    surface: '#FFFFFF',
    surfaceSubtle: '#F1F5F9', // Slate 100
    surfaceHover: '#E2E8F0', // Slate 200
    borderSubtle: '#E2E8F0',
    borderStrong: '#CBD5E1', // Slate 300
    primary: '#0D9488', // Emerald Teal 600
    primaryHover: '#0F766E',
    primaryLight: 'rgba(13, 148, 136, 0.08)',
    accentSubtle: 'rgba(13, 148, 136, 0.12)',
    accentGlow: 'rgba(13, 148, 136, 0.22)',
    textPrimary: '#0F172A', // Slate 900
    textSecondary: '#64748B', // Slate 500
    textMuted: '#94A3B8', // Slate 400
    warning: '#D97706',
    warningBg: 'rgba(217, 119, 6, 0.10)',
    error: '#DC2626',
    errorBg: 'rgba(220, 38, 38, 0.10)',
    gridCellBg: '#FFFFFF',
    gridHeaderBg: '#F8FAFC',
    gridRowHeaderBg: '#F1F5F9',
    gridBorder: '#E2E8F0',
  },
  dark: {
    mode: 'dark' as const,
    canvas: '#0B0F17', // Deep Obsidian
    surface: '#111827', // Slate 900
    surfaceSubtle: '#1F2937', // Slate 800
    surfaceHover: '#374151', // Slate 700
    borderSubtle: 'rgba(255, 255, 255, 0.09)',
    borderStrong: 'rgba(255, 255, 255, 0.18)',
    primary: '#14B8A6', // Emerald Teal 500
    primaryHover: '#2DD4BF',
    primaryLight: 'rgba(20, 184, 166, 0.12)',
    accentSubtle: 'rgba(20, 184, 166, 0.18)',
    accentGlow: 'rgba(20, 184, 166, 0.35)',
    textPrimary: '#F8FAFC', // Slate 50
    textSecondary: '#94A3B8', // Slate 400
    textMuted: '#64748B', // Slate 500
    warning: '#F59E0B',
    warningBg: 'rgba(245, 158, 11, 0.15)',
    error: '#EF4444',
    errorBg: 'rgba(239, 68, 68, 0.15)',
    gridCellBg: '#111827',
    gridHeaderBg: '#0F172A',
    gridRowHeaderBg: '#162032',
    gridBorder: 'rgba(255, 255, 255, 0.08)',
  }
};

export const typography = {
  fontUi: '"Inter Variable", "Inter", -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif',
  fontMono: '"JetBrains Mono", "SF Mono", Consolas, "Cascadia Code", monospace',
};

export const motion = {
  fast: '120ms cubic-bezier(0.16, 1, 0.3, 1)',
  standard: '180ms cubic-bezier(0.16, 1, 0.3, 1)',
};

export function getGridTheme(mode: ColorMode, zoom: number) {
  const p = palette[mode];
  const scale = zoom / 100;
  const baseTheme = mode === 'dark' ? themeQuartz.withPart(colorSchemeDark) : themeQuartz.withPart(colorSchemeLight);

  return baseTheme.withParams({
    accentColor: p.primary,
    fontFamily: typography.fontUi,
    fontSize: Math.round(13 * scale),
    rowHeight: Math.round(30 * scale),
    headerHeight: Math.round(34 * scale),
    backgroundColor: p.gridCellBg,
    foregroundColor: p.textPrimary,
    headerBackgroundColor: p.gridHeaderBg,
    headerTextColor: p.textSecondary,
    borderColor: p.gridBorder,
    rowBorder: { style: 'solid', width: 1, color: p.gridBorder },
    columnBorder: { style: 'solid', width: 1, color: p.gridBorder },
    selectedRowBackgroundColor: p.accentSubtle,
    oddRowBackgroundColor: mode === 'dark' ? 'rgba(255, 255, 255, 0.015)' : 'rgba(0, 0, 0, 0.015)',
  });
}
