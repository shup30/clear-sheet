import React, { useState, useMemo, useRef } from 'react';
import {
  Dialog,
  DialogTitle,
  DialogContent,
  DialogActions,
  Button,
  Box,
  Typography,
  Select,
  MenuItem,
  FormControl,
  InputLabel,
  TextField,
  FormControlLabel,
  Checkbox,
  Paper,
} from '@mui/material';
import {
  BarChart2,
  Download,
  Copy,
} from 'lucide-react';
import type { WorkbookModel } from '../workbook/model';
import type { Selection } from '../store';
import { colToLetter } from '../workbook/cellRef';

export type ChartType = 'column' | 'bar' | 'line' | 'area' | 'pie' | 'scatter';

interface ChartModalProps {
  open: boolean;
  onClose: () => void;
  sheetName: string | null;
  selection: Selection | null;
  model: WorkbookModel | null;
}

const PALETTES = {
  office: ['#2563eb', '#ea580c', '#16a34a', '#d97706', '#9333ea', '#0891b2'],
  ocean: ['#0284c7', '#06b6d4', '#3b82f6', '#6366f1', '#0ea5e9', '#38bdf8'],
  sunset: ['#f43f5e', '#fb923c', '#facc15', '#e11d48', '#ea580c', '#d97706'],
  forest: ['#059669', '#10b981', '#34d399', '#0d9488', '#14b8a6', '#2dd4bf'],
};

export const ChartModal: React.FC<ChartModalProps> = ({
  open,
  onClose,
  sheetName,
  selection,
  model,
}) => {
  const [chartType, setChartType] = useState<ChartType>('column');
  const [title, setTitle] = useState<string>('Chart Analysis');
  const [paletteKey, setPaletteKey] = useState<keyof typeof PALETTES>('office');
  const [showLegend, setShowLegend] = useState(true);
  const [showGridlines, setShowGridlines] = useState(true);
  const [showLabels, setShowLabels] = useState(false);
  const [hoveredPoint, setHoveredPoint] = useState<{ label: string; series: string; val: number; x: number; y: number } | null>(null);

  const svgRef = useRef<SVGSVGElement>(null);

  // Compute selected data rectangle
  const parsedData = useMemo(() => {
    if (!model || !sheetName) return null;
    let r1 = 0, c1 = 0, r2 = 3, c2 = 3;
    if (selection && selection.ranges.length > 0) {
      const rg = selection.ranges[0];
      r1 = rg.r1; c1 = rg.c1; r2 = rg.r2; c2 = rg.c2;
    }

    const rows = r2 - r1 + 1;
    const cols = c2 - c1 + 1;

    // Detect header row and header col
    const topRowFirstVal = model.displayOf(sheetName, r1, c1).v;
    const topRowIsHeader = typeof topRowFirstVal === 'string' && rows > 1;
    const leftColIsCategory = cols > 1;

    const dataRowStart = topRowIsHeader ? r1 + 1 : r1;
    const dataColStart = leftColIsCategory ? c1 + 1 : c1;

    const categories: string[] = [];
    for (let r = dataRowStart; r <= r2; r++) {
      if (leftColIsCategory) {
        const text = model.displayOf(sheetName, r, c1).text;
        categories.push(text || `Row ${r + 1}`);
      } else {
        categories.push(`Row ${r + 1}`);
      }
    }

    const seriesList: { name: string; values: number[] }[] = [];
    for (let c = dataColStart; c <= c2; c++) {
      const seriesName = topRowIsHeader
        ? model.displayOf(sheetName, r1, c).text || `${colToLetter(c)}`
        : `Series ${colToLetter(c)}`;

      const values: number[] = [];
      for (let r = dataRowStart; r <= r2; r++) {
        const val = model.displayOf(sheetName, r, c).v;
        values.push(typeof val === 'number' ? val : 0);
      }
      seriesList.push({ name: seriesName, values });
    }

    return { categories, seriesList, rangeStr: `${colToLetter(c1)}${r1 + 1}:${colToLetter(c2)}${r2 + 1}` };
  }, [model, sheetName, selection]);

  const colors = PALETTES[paletteKey];

  // SVG Chart Geometry
  const width = 640;
  const height = 340;
  const margin = { top: 40, right: 30, bottom: 50, left: 60 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;

  // Compute min / max values
  const allValues = parsedData?.seriesList.flatMap((s) => s.values) ?? [0];
  const maxVal = Math.max(...allValues, 10);
  const minVal = Math.min(0, ...allValues);
  const range = maxVal - minVal || 1;

  const getY = (val: number) => {
    return margin.top + plotHeight - ((val - minVal) / range) * plotHeight;
  };

  // Export handlers
  const handleExportSVG = () => {
    if (!svgRef.current) return;
    const serializer = new XMLSerializer();
    const source = serializer.serializeToString(svgRef.current);
    const blob = new Blob([source], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = `${title.toLowerCase().replace(/\s+/g, '_')}.svg`;
    a.click();
    URL.revokeObjectURL(url);
  };

  const handleExportPNG = () => {
    if (!svgRef.current) return;
    const serializer = new XMLSerializer();
    const source = serializer.serializeToString(svgRef.current);
    const img = new Image();
    const svgBlob = new Blob([source], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(svgBlob);

    img.onload = () => {
      const canvas = document.createElement('canvas');
      canvas.width = width * 2;
      canvas.height = height * 2;
      const ctx = canvas.getContext('2d');
      if (ctx) {
        ctx.scale(2, 2);
        ctx.fillStyle = '#ffffff';
        ctx.fillRect(0, 0, width, height);
        ctx.drawImage(img, 0, 0);
        const pngUrl = canvas.toDataURL('image/png');
        const a = document.createElement('a');
        a.href = pngUrl;
        a.download = `${title.toLowerCase().replace(/\s+/g, '_')}.png`;
        a.click();
      }
      URL.revokeObjectURL(url);
    };
    img.src = url;
  };

  return (
    <Dialog open={open} onClose={onClose} maxWidth="md" fullWidth>
      <DialogTitle sx={{ display: 'flex', alignItems: 'center', justifyContent: 'space-between', pb: 1 }}>
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 1 }}>
          <BarChart2 size={20} color="var(--mui-palette-primary-main, #2563eb)" />
          <Typography variant="h6" sx={{ fontSize: '1.05rem', fontWeight: 700 }}>
            Insert Chart ({parsedData?.rangeStr})
          </Typography>
        </Box>
        <Box sx={{ display: 'flex', gap: 1 }}>
          <Button size="small" variant="outlined" startIcon={<Copy size={14} />} onClick={handleExportSVG}>
            SVG
          </Button>
          <Button size="small" variant="outlined" startIcon={<Download size={14} />} onClick={handleExportPNG}>
            PNG
          </Button>
        </Box>
      </DialogTitle>

      <DialogContent sx={{ display: 'flex', flexDirection: 'column', gap: 2, pt: 1 }}>
        {/* Controls Toolbar */}
        <Box sx={{ display: 'flex', gap: 1.5, flexWrap: 'wrap', alignItems: 'center' }}>
          <FormControl size="small" sx={{ width: 140 }}>
            <InputLabel>Chart Type</InputLabel>
            <Select value={chartType} label="Chart Type" onChange={(e) => setChartType(e.target.value as ChartType)}>
              <MenuItem value="column">Column</MenuItem>
              <MenuItem value="bar">Bar</MenuItem>
              <MenuItem value="line">Line</MenuItem>
              <MenuItem value="area">Area</MenuItem>
              <MenuItem value="pie">Pie</MenuItem>
              <MenuItem value="scatter">Scatter</MenuItem>
            </Select>
          </FormControl>

          <TextField
            size="small"
            label="Chart Title"
            value={title}
            onChange={(e) => setTitle(e.target.value)}
            sx={{ width: 180 }}
          />

          <FormControl size="small" sx={{ width: 130 }}>
            <InputLabel>Palette</InputLabel>
            <Select value={paletteKey} label="Palette" onChange={(e) => setPaletteKey(e.target.value as keyof typeof PALETTES)}>
              <MenuItem value="office">Office</MenuItem>
              <MenuItem value="ocean">Ocean</MenuItem>
              <MenuItem value="sunset">Sunset</MenuItem>
              <MenuItem value="forest">Forest</MenuItem>
            </Select>
          </FormControl>

          <FormControlLabel
            control={<Checkbox checked={showLegend} onChange={(e) => setShowLegend(e.target.checked)} size="small" />}
            label="Legend"
          />
          <FormControlLabel
            control={<Checkbox checked={showGridlines} onChange={(e) => setShowGridlines(e.target.checked)} size="small" />}
            label="Grid"
          />
          <FormControlLabel
            control={<Checkbox checked={showLabels} onChange={(e) => setShowLabels(e.target.checked)} size="small" />}
            label="Labels"
          />
        </Box>

        {/* Chart Viewport */}
        <Paper
          variant="outlined"
          sx={{
            p: 1.5,
            bgcolor: 'background.default',
            display: 'flex',
            justifyContent: 'center',
            alignItems: 'center',
            position: 'relative',
            minHeight: 360,
            overflow: 'hidden',
          }}
        >
          {parsedData && parsedData.seriesList.length > 0 ? (
            <svg
              ref={svgRef}
              width={width}
              height={height}
              viewBox={`0 0 ${width} ${height}`}
              style={{ maxWidth: '100%', height: 'auto', fontFamily: 'Inter, system-ui, sans-serif' }}
              onMouseLeave={() => setHoveredPoint(null)}
            >
              {/* Title */}
              <text x={width / 2} y={24} textAnchor="middle" fontSize={16} fontWeight={700} fill="currentColor">
                {title}
              </text>

              {/* Gridlines */}
              {showGridlines && chartType !== 'pie' && (
                <g opacity={0.15}>
                  {[0, 0.25, 0.5, 0.75, 1].map((ratio) => {
                    const y = margin.top + plotHeight * (1 - ratio);
                    const val = minVal + range * ratio;
                    return (
                      <g key={ratio}>
                        <line x1={margin.left} y1={y} x2={width - margin.right} y2={y} stroke="currentColor" strokeDasharray="3 3" />
                        <text x={margin.left - 8} y={y + 4} textAnchor="end" fontSize={10} fill="currentColor">
                          {Math.round(val)}
                        </text>
                      </g>
                    );
                  })}
                </g>
              )}

              {/* COLUMN CHART */}
              {chartType === 'column' && (
                <g>
                  {parsedData.categories.map((cat, catIdx) => {
                    const groupWidth = plotWidth / parsedData.categories.length;
                    const groupX = margin.left + catIdx * groupWidth;
                    const numSeries = parsedData.seriesList.length;
                    const barWidth = Math.max(6, (groupWidth * 0.75) / numSeries);
                    const groupPadding = (groupWidth - barWidth * numSeries) / 2;

                    return (
                      <g key={cat}>
                        {parsedData.seriesList.map((series, sIdx) => {
                          const val = series.values[catIdx] ?? 0;
                          const barX = groupX + groupPadding + sIdx * barWidth;
                          const barY = getY(Math.max(0, val));
                          const zeroY = getY(0);
                          const barH = Math.max(2, Math.abs(zeroY - getY(val)));
                          const color = colors[sIdx % colors.length];

                          return (
                            <rect
                              key={series.name}
                              x={barX}
                              y={barY}
                              width={barWidth - 2}
                              height={barH}
                              fill={color}
                              rx={2}
                              opacity={0.9}
                              style={{ cursor: 'pointer', transition: 'all 0.15s ease' }}
                              onMouseEnter={() => setHoveredPoint({ label: cat, series: series.name, val, x: barX, y: barY })}
                            />
                          );
                        })}
                        {/* Category label */}
                        <text
                          x={groupX + groupWidth / 2}
                          y={height - margin.bottom + 18}
                          textAnchor="middle"
                          fontSize={11}
                          fill="currentColor"
                          opacity={0.8}
                        >
                          {cat}
                        </text>
                      </g>
                    );
                  })}
                </g>
              )}

              {/* LINE / AREA CHART */}
              {(chartType === 'line' || chartType === 'area') && (
                <g>
                  {parsedData.seriesList.map((series, sIdx) => {
                    const color = colors[sIdx % colors.length];
                    const pts = series.values.map((v, i) => {
                      const x = margin.left + (i / Math.max(1, parsedData.categories.length - 1)) * plotWidth;
                      const y = getY(v);
                      return { x, y, v, cat: parsedData.categories[i] };
                    });

                    const pathStr = pts.reduce((acc, p, idx) => `${acc} ${idx === 0 ? 'M' : 'L'} ${p.x} ${p.y}`, '');
                    const zeroY = getY(0);
                    const areaStr = `${pathStr} L ${pts[pts.length - 1].x} ${zeroY} L ${pts[0].x} ${zeroY} Z`;

                    return (
                      <g key={series.name}>
                        {chartType === 'area' && <path d={areaStr} fill={color} opacity={0.25} />}
                        <path d={pathStr} fill="none" stroke={color} strokeWidth={2.5} strokeLinecap="round" strokeLinejoin="round" />
                        {pts.map((p, pIdx) => (
                          <circle
                            key={pIdx}
                            cx={p.x}
                            cy={p.y}
                            r={4}
                            fill="#ffffff"
                            stroke={color}
                            strokeWidth={2}
                            style={{ cursor: 'pointer' }}
                            onMouseEnter={() => setHoveredPoint({ label: p.cat, series: series.name, val: p.v, x: p.x, y: p.y })}
                          />
                        ))}
                      </g>
                    );
                  })}
                  {/* Category labels */}
                  {parsedData.categories.map((cat, i) => {
                    const x = margin.left + (i / Math.max(1, parsedData.categories.length - 1)) * plotWidth;
                    return (
                      <text key={cat} x={x} y={height - margin.bottom + 18} textAnchor="middle" fontSize={11} fill="currentColor" opacity={0.8}>
                        {cat}
                      </text>
                    );
                  })}
                </g>
              )}

              {/* PIE CHART */}
              {chartType === 'pie' && (
                <g transform={`translate(${width / 2}, ${height / 2})`}>
                  {(() => {
                    const firstSeries = parsedData.seriesList[0];
                    if (!firstSeries) return null;
                    const total = firstSeries.values.reduce((a, b) => a + Math.max(0, b), 0) || 1;
                    let currentAngle = 0;
                    const radius = Math.min(plotWidth, plotHeight) / 2.3;

                    return firstSeries.values.map((val, idx) => {
                      const positiveVal = Math.max(0, val);
                      const sliceAngle = (positiveVal / total) * Math.PI * 2;
                      const x1 = Math.cos(currentAngle) * radius;
                      const y1 = Math.sin(currentAngle) * radius;
                      const x2 = Math.cos(currentAngle + sliceAngle) * radius;
                      const y2 = Math.sin(currentAngle + sliceAngle) * radius;
                      const largeArc = sliceAngle > Math.PI ? 1 : 0;
                      const d = `M 0 0 L ${x1} ${y1} A ${radius} ${radius} 0 ${largeArc} 1 ${x2} ${y2} Z`;
                      const color = colors[idx % colors.length];
                      const midAngle = currentAngle + sliceAngle / 2;
                      currentAngle += sliceAngle;

                      return (
                        <path
                          key={idx}
                          d={d}
                          fill={color}
                          stroke="#ffffff"
                          strokeWidth={1.5}
                          style={{ cursor: 'pointer' }}
                          onMouseEnter={() =>
                            setHoveredPoint({
                              label: parsedData.categories[idx] || `Item ${idx + 1}`,
                              series: firstSeries.name,
                              val,
                              x: width / 2 + Math.cos(midAngle) * (radius / 1.5),
                              y: height / 2 + Math.sin(midAngle) * (radius / 1.5),
                            })
                          }
                        />
                      );
                    });
                  })()}
                </g>
              )}

              {/* Legend */}
              {showLegend && (
                <g transform={`translate(${margin.left}, ${height - 18})`}>
                  {parsedData.seriesList.map((series, sIdx) => {
                    const color = colors[sIdx % colors.length];
                    const itemX = sIdx * 110;
                    return (
                      <g key={series.name} transform={`translate(${itemX}, 0)`}>
                        <rect width={12} height={12} fill={color} rx={2} />
                        <text x={18} y={10} fontSize={11} fill="currentColor">
                          {series.name}
                        </text>
                      </g>
                    );
                  })}
                </g>
              )}
            </svg>
          ) : (
            <Typography variant="body2" sx={{ color: 'text.secondary' }}>
              Select a range of cells to generate a chart.
            </Typography>
          )}

          {/* Tooltip Overlay */}
          {hoveredPoint && (
            <Box
              sx={{
                position: 'absolute',
                left: hoveredPoint.x,
                top: hoveredPoint.y - 36,
                transform: 'translate(-50%, -100%)',
                bgcolor: 'rgba(15, 23, 42, 0.9)',
                color: '#ffffff',
                px: 1.25,
                py: 0.5,
                borderRadius: 1,
                fontSize: '0.75rem',
                fontFamily: 'monospace',
                pointerEvents: 'none',
                boxShadow: '0 4px 12px rgba(0,0,0,0.3)',
                whiteSpace: 'nowrap',
                zIndex: 10,
              }}
            >
              <strong>{hoveredPoint.label}</strong> ({hoveredPoint.series}): {hoveredPoint.val}
            </Box>
          )}
        </Paper>
      </DialogContent>

      <DialogActions sx={{ px: 3, py: 1.5 }}>
        <Button onClick={onClose} variant="contained" color="primary">
          Done
        </Button>
      </DialogActions>
    </Dialog>
  );
};
