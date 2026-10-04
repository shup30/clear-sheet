import type { Range } from './cellRef';
import type { EvalVal } from './formulas';

export type CFOperator = 'gt' | 'lt' | 'between' | 'eq' | 'contains';

export interface ConditionalFormatRule {
  id: string;
  sheet: string;
  range: Range;
  ruleType: 'highlight' | 'colorScale' | 'dataBar';
  operator?: CFOperator;
  val1?: number | string;
  val2?: number;
  bg?: string;
  color?: string;
  minColor?: string;
  midColor?: string;
  maxColor?: string;
  barColor?: string;
}

function hexToRgb(hex: string): [number, number, number] {
  const clean = hex.replace('#', '');
  if (clean.length === 3) {
    return [
      parseInt(clean[0] + clean[0], 16),
      parseInt(clean[1] + clean[1], 16),
      parseInt(clean[2] + clean[2], 16),
    ];
  }
  return [
    parseInt(clean.slice(0, 2), 16) || 0,
    parseInt(clean.slice(2, 4), 16) || 0,
    parseInt(clean.slice(4, 6), 16) || 0,
  ];
}

function interpolateRgb(c1: [number, number, number], c2: [number, number, number], t: number): string {
  const r = Math.round(c1[0] + (c2[0] - c1[0]) * t);
  const g = Math.round(c1[1] + (c2[1] - c1[1]) * t);
  const b = Math.round(c1[2] + (c2[2] - c1[2]) * t);
  return `rgb(${r}, ${g}, ${b})`;
}

/**
 * Evaluates conditional formatting rules on a given cell.
 * Returns partial CSS style properties if a rule matches.
 */
export function evaluateConditionalFormatting(
  rules: ConditionalFormatRule[],
  sheet: string,
  r: number,
  c: number,
  val: EvalVal,
  getRangeNums: (rg: Range) => { min: number; max: number; mid: number }
): { bg?: string; color?: string; background?: string } | null {
  for (const rule of rules) {
    if (rule.sheet.toLowerCase() !== sheet.toLowerCase()) continue;
    const rg = rule.range;
    if (r < rg.r1 || r > rg.r2 || c < rg.c1 || c > rg.c2) continue;

    // 1. Highlight cell rules
    if (rule.ruleType === 'highlight' && rule.operator) {
      if (val === null || typeof val === 'object') continue;

      if (rule.operator === 'contains') {
        const text = String(val).toLowerCase();
        const search = String(rule.val1 ?? '').toLowerCase();
        if (search && text.includes(search)) {
          return { bg: rule.bg, color: rule.color };
        }
      } else if (typeof val === 'number') {
        const v1 = typeof rule.val1 === 'number' ? rule.val1 : Number(rule.val1);
        const v2 = rule.val2 !== undefined ? Number(rule.val2) : undefined;

        if (rule.operator === 'gt' && !isNaN(v1) && val > v1) {
          return { bg: rule.bg, color: rule.color };
        }
        if (rule.operator === 'lt' && !isNaN(v1) && val < v1) {
          return { bg: rule.bg, color: rule.color };
        }
        if (rule.operator === 'eq' && !isNaN(v1) && val === v1) {
          return { bg: rule.bg, color: rule.color };
        }
        if (rule.operator === 'between' && !isNaN(v1) && v2 !== undefined && !isNaN(v2)) {
          const lo = Math.min(v1, v2);
          const hi = Math.max(v1, v2);
          if (val >= lo && val <= hi) {
            return { bg: rule.bg, color: rule.color };
          }
        }
      }
    }

    // 2. Color Scales (Heatmaps)
    if (rule.ruleType === 'colorScale' && typeof val === 'number') {
      const stats = getRangeNums(rule.range);
      if (stats.min >= stats.max) continue;
      const minCol = hexToRgb(rule.minColor || '#f87171');
      const maxCol = hexToRgb(rule.maxColor || '#4ade80');

      if (rule.midColor) {
        const midCol = hexToRgb(rule.midColor);
        if (val <= stats.mid) {
          const t = stats.mid > stats.min ? (val - stats.min) / (stats.mid - stats.min) : 0;
          return { bg: interpolateRgb(minCol, midCol, Math.max(0, Math.min(1, t))) };
        } else {
          const t = stats.max > stats.mid ? (val - stats.mid) / (stats.max - stats.mid) : 0;
          return { bg: interpolateRgb(midCol, maxCol, Math.max(0, Math.min(1, t))) };
        }
      } else {
        const t = (val - stats.min) / (stats.max - stats.min);
        return { bg: interpolateRgb(minCol, maxCol, Math.max(0, Math.min(1, t))) };
      }
    }

    // 3. Data Bars
    if (rule.ruleType === 'dataBar' && typeof val === 'number') {
      const stats = getRangeNums(rule.range);
      const span = stats.max - Math.min(0, stats.min);
      if (span <= 0) continue;
      const pct = Math.max(2, Math.min(100, Math.round(((val - Math.min(0, stats.min)) / span) * 100)));
      const barCol = rule.barColor || '#38bdf8';
      return {
        background: `linear-gradient(to right, ${barCol}55 ${pct}%, transparent ${pct}%)`,
      };
    }
  }

  return null;
}
