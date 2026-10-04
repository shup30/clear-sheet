import React, { useState, useEffect, useRef, useCallback } from 'react';
import { Box, Tooltip, IconButton, Paper, Typography } from '@mui/material';
import { Check, X, FunctionSquare } from 'lucide-react';
import { getFormulaContext, type FunctionMeta } from '../workbook/intellisense';
import { cycleRefAbsolute } from '../workbook/cellRef';

interface FormulaBarProps {
  address: string;
  value: string;
  isEditing: boolean;
  onFocus: () => void;
  onBlur?: () => void;
  onChange: (val: string) => void;
  onCommit: () => void;
  onCancel: () => void;
}

export const FormulaBar: React.FC<FormulaBarProps> = ({
  address,
  value,
  isEditing,
  onFocus,
  onBlur,
  onChange,
  onCommit,
  onCancel,
}) => {
  const inputRef = useRef<HTMLInputElement>(null);
  const [cursorPos, setCursorPos] = useState<number>(value.length);
  const [selectedSuggestIdx, setSelectedSuggestIdx] = useState<number>(0);
  const [showDropdown, setShowDropdown] = useState<boolean>(false);

  // Update cursor position on click / keyup / select
  const syncCursor = useCallback(() => {
    if (inputRef.current) {
      setCursorPos(inputRef.current.selectionStart ?? value.length);
    }
  }, [value.length]);

  const ctx = React.useMemo(() => {
    if (!isEditing || !value.startsWith('=')) {
      return { isFormula: false, activeParamIndex: 0, suggestions: [] as FunctionMeta[], prefix: '', prefixStart: 0 };
    }
    return getFormulaContext(value, cursorPos);
  }, [value, cursorPos, isEditing]);

  const suggestions = ctx.suggestions;
  const activeFunction = ctx.activeFunction;
  const activeParamIndex = ctx.activeParamIndex;

  useEffect(() => {
    setSelectedSuggestIdx(0);
    setShowDropdown(suggestions.length > 0);
  }, [suggestions]);

  const acceptSuggestion = (func: FunctionMeta) => {
    const before = value.slice(0, ctx.prefixStart);
    const after = value.slice(cursorPos);
    const completed = `${before}${func.name}(`;
    const nextVal = `${completed}${after}`;
    onChange(nextVal);
    setShowDropdown(false);
    setTimeout(() => {
      if (inputRef.current) {
        inputRef.current.focus();
        inputRef.current.setSelectionRange(completed.length, completed.length);
        setCursorPos(completed.length);
      }
    }, 0);
  };

  return (
    <section className="formula formula-bar-container" style={{ position: 'relative' }}>
      {/* Address Pill */}
      <span className="address" title="Active cell / range">
        {address || '—'}
      </span>

      {/* FX Icon Badge */}
      <span className="fx" title="Formula">
        fx
      </span>

      {/* Edit actions (shown when editing formula) */}
      {isEditing && (
        <Box sx={{ display: 'flex', alignItems: 'center', gap: 0.25, pr: 0.5 }}>
          <Tooltip title="Cancel (Esc)">
            <IconButton size="small" onClick={onCancel} sx={{ p: 0.3, color: 'text.secondary' }}>
              <X size={13} />
            </IconButton>
          </Tooltip>
          <Tooltip title="Commit (Enter)">
            <IconButton size="small" onClick={onCommit} sx={{ p: 0.3, color: 'primary.main' }}>
              <Check size={13} />
            </IconButton>
          </Tooltip>
        </Box>
      )}

      {/* Monospace Formula Input */}
      <input
        ref={inputRef}
        aria-label="Formula bar"
        value={value}
        placeholder="Enter a value or formula (e.g. =SUM(A1:A10))"
        onFocus={() => {
          onFocus();
          syncCursor();
        }}
        onBlur={() => {
          // Delay blur to allow suggestion clicking
          setTimeout(() => {
            setShowDropdown(false);
            if (onBlur) onBlur();
          }, 150);
        }}
        onClick={syncCursor}
        onKeyUp={syncCursor}
        onChange={(e) => {
          onChange(e.target.value);
          setCursorPos(e.target.selectionStart ?? e.target.value.length);
        }}
        onKeyDown={(e) => {
          // Autocomplete navigation
          if (showDropdown && suggestions.length > 0) {
            if (e.key === 'ArrowDown') {
              e.preventDefault();
              setSelectedSuggestIdx((prev) => (prev + 1) % suggestions.length);
              return;
            }
            if (e.key === 'ArrowUp') {
              e.preventDefault();
              setSelectedSuggestIdx((prev) => (prev - 1 + suggestions.length) % suggestions.length);
              return;
            }
            if (e.key === 'Tab' || (e.key === 'Enter' && selectedSuggestIdx >= 0)) {
              e.preventDefault();
              acceptSuggestion(suggestions[selectedSuggestIdx]);
              return;
            }
          }

          if (e.key === 'F4') {
            e.preventDefault();
            const cur = inputRef.current?.selectionStart ?? value.length;
            const cycled = cycleRefAbsolute(value, cur);
            onChange(cycled.nextExpr);
            setTimeout(() => {
              if (inputRef.current) {
                inputRef.current.setSelectionRange(cycled.nextPos, cycled.nextPos);
                setCursorPos(cycled.nextPos);
              }
            }, 0);
            return;
          }

          if (e.key === 'Enter') {
            e.preventDefault();
            setShowDropdown(false);
            onCommit();
            (e.target as HTMLInputElement).blur();
          } else if (e.key === 'Escape') {
            e.preventDefault();
            setShowDropdown(false);
            onCancel();
            (e.target as HTMLInputElement).blur();
          }
          e.stopPropagation();
        }}
      />

      {/* Autocomplete Suggestions Popover */}
      {showDropdown && suggestions.length > 0 && isEditing && (
        <Paper
          elevation={8}
          className="formula-autocomplete-popup"
          sx={{
            position: 'absolute',
            top: '100%',
            left: '95px',
            mt: 0.5,
            minWidth: 320,
            maxWidth: 480,
            zIndex: 1300,
            bgcolor: 'background.paper',
            border: '1px solid',
            borderColor: 'divider',
            borderRadius: 1.5,
            boxShadow: '0 8px 24px rgba(0,0,0,0.28)',
            overflow: 'hidden',
          }}
        >
          {suggestions.map((fn, idx) => {
            const isSelected = idx === selectedSuggestIdx;
            return (
              <Box
                key={fn.name}
                onMouseDown={(e) => {
                  e.preventDefault();
                  acceptSuggestion(fn);
                }}
                sx={{
                  px: 1.5,
                  py: 0.75,
                  cursor: 'pointer',
                  display: 'flex',
                  alignItems: 'center',
                  gap: 1.25,
                  bgcolor: isSelected ? 'action.selected' : 'transparent',
                  '&:hover': { bgcolor: 'action.hover' },
                  borderBottom: idx < suggestions.length - 1 ? '1px solid' : 'none',
                  borderColor: 'divider',
                }}
              >
                <FunctionSquare size={16} style={{ opacity: 0.75, flexShrink: 0 }} />
                <Box sx={{ minWidth: 0, flex: 1 }}>
                  <Typography variant="body2" sx={{ fontFamily: 'monospace', fontWeight: 700, fontSize: '0.8125rem' }}>
                    {fn.name}
                  </Typography>
                  <Typography variant="caption" sx={{ color: 'text.secondary', display: 'block', fontSize: '0.72rem' }} noWrap>
                    {fn.description}
                  </Typography>
                </Box>
              </Box>
            );
          })}
        </Paper>
      )}

      {/* Active Function Signature & Parameter Guidance Tooltip */}
      {activeFunction && !showDropdown && isEditing && (
        <Box
          className="formula-syntax-hint"
          sx={{
            display: 'flex',
            flexDirection: 'column',
            gap: 0.35,
            maxWidth: 520,
            pointerEvents: 'none',
          }}
        >
          <Box sx={{ display: 'flex', alignItems: 'center', flexWrap: 'wrap', gap: 0.5 }}>
            <span style={{ fontWeight: 700, color: 'var(--mui-palette-primary-main, #60a5fa)' }}>
              {activeFunction.name}(
            </span>
            {activeFunction.params.map((param, i) => {
              const isActive = i === activeParamIndex || (i === activeFunction.params.length - 1 && activeParamIndex >= i);
              return (
                <span
                  key={param.name}
                  style={{
                    fontWeight: isActive ? 800 : 400,
                    textDecoration: isActive ? 'underline' : 'none',
                    color: isActive ? '#fef08a' : 'inherit',
                  }}
                >
                  {param.name}
                  {i < activeFunction.params.length - 1 ? ', ' : ''}
                </span>
              );
            })}
            <span>)</span>
          </Box>
          {activeFunction.params[Math.min(activeParamIndex, activeFunction.params.length - 1)] && (
            <Box sx={{ fontSize: '0.6875rem', opacity: 0.85, fontStyle: 'italic' }}>
              <strong>{activeFunction.params[Math.min(activeParamIndex, activeFunction.params.length - 1)].name}:</strong>{' '}
              {activeFunction.params[Math.min(activeParamIndex, activeFunction.params.length - 1)].description}
            </Box>
          )}
        </Box>
      )}
    </section>
  );
};
