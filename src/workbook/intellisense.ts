import { parseRange, type Range } from './cellRef';

export interface ParamMeta {
  name: string;
  description: string;
  optional?: boolean;
}

export interface FunctionMeta {
  name: string;
  syntax: string;
  description: string;
  params: ParamMeta[];
}

export const FORMULA_FUNCTIONS: Record<string, FunctionMeta> = {
  SUM: {
    name: 'SUM',
    syntax: 'SUM(number1, [number2], ...)',
    description: 'Adds all the numbers in a range of cells.',
    params: [
      { name: 'number1', description: 'The first number, cell reference, or range to sum.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  AVERAGE: {
    name: 'AVERAGE',
    syntax: 'AVERAGE(number1, [number2], ...)',
    description: 'Returns the average (arithmetic mean) of its arguments.',
    params: [
      { name: 'number1', description: 'The first number, cell reference, or range to average.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  MIN: {
    name: 'MIN',
    syntax: 'MIN(number1, [number2], ...)',
    description: 'Returns the smallest number in a set of values.',
    params: [
      { name: 'number1', description: 'First number, cell reference, or range.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  MAX: {
    name: 'MAX',
    syntax: 'MAX(number1, [number2], ...)',
    description: 'Returns the largest number in a set of values.',
    params: [
      { name: 'number1', description: 'First number, cell reference, or range.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  COUNT: {
    name: 'COUNT',
    syntax: 'COUNT(value1, [value2], ...)',
    description: 'Counts how many cells contain numbers.',
    params: [
      { name: 'value1', description: 'First item, cell reference, or range to count numbers in.' },
      { name: 'value2', description: 'Optional additional items or ranges.', optional: true },
    ],
  },
  COUNTA: {
    name: 'COUNTA',
    syntax: 'COUNTA(value1, [value2], ...)',
    description: 'Counts how many cells are not empty.',
    params: [
      { name: 'value1', description: 'First item, cell reference, or range to count non-empty cells in.' },
      { name: 'value2', description: 'Optional additional items or ranges.', optional: true },
    ],
  },
  COUNTBLANK: {
    name: 'COUNTBLANK',
    syntax: 'COUNTBLANK(range)',
    description: 'Counts the number of empty cells in a specified range.',
    params: [
      { name: 'range', description: 'The range from which the blank cells are to be counted.' },
    ],
  },
  COUNTIF: {
    name: 'COUNTIF',
    syntax: 'COUNTIF(range, criteria)',
    description: 'Counts the number of cells within a range that meet the given condition.',
    params: [
      { name: 'range', description: 'The range of cells to evaluate.' },
      { name: 'criteria', description: 'The condition in the form of a number, expression, or text (e.g. ">5", "Apple").' },
    ],
  },
  COUNTIFS: {
    name: 'COUNTIFS',
    syntax: 'COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Counts the number of cells specified by a given set of conditions or criteria.',
    params: [
      { name: 'criteria_range1', description: 'The first range in which to evaluate the associated criteria.' },
      { name: 'criteria1', description: 'The condition that defines which cells will be counted.' },
      { name: 'criteria_range2', description: 'Additional range to evaluate.', optional: true },
      { name: 'criteria2', description: 'Additional criteria to evaluate.', optional: true },
    ],
  },
  SUMIF: {
    name: 'SUMIF',
    syntax: 'SUMIF(range, criteria, [sum_range])',
    description: 'Adds the cells specified by a given condition or criteria.',
    params: [
      { name: 'range', description: 'The range of cells you want evaluated by the criteria.' },
      { name: 'criteria', description: 'The condition that defines which cells are added.' },
      { name: 'sum_range', description: 'The actual cells to sum, if different from range.', optional: true },
    ],
  },
  SUMIFS: {
    name: 'SUMIFS',
    syntax: 'SUMIFS(sum_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Adds cells in a range that meet multiple criteria.',
    params: [
      { name: 'sum_range', description: 'The actual cells to sum.' },
      { name: 'criteria_range1', description: 'The first range to evaluate.' },
      { name: 'criteria1', description: 'The criteria that determines which cells in criteria_range1 to include.' },
      { name: 'criteria_range2', description: 'Additional range to evaluate.', optional: true },
      { name: 'criteria2', description: 'Additional criteria to evaluate.', optional: true },
    ],
  },
  AVERAGEIF: {
    name: 'AVERAGEIF',
    syntax: 'AVERAGEIF(range, criteria, [average_range])',
    description: 'Returns the average of all cells in a range that meet a given criteria.',
    params: [
      { name: 'range', description: 'The range of cells you want evaluated.' },
      { name: 'criteria', description: 'The condition that defines which cells to average.' },
      { name: 'average_range', description: 'The actual cells to average, if different from range.', optional: true },
    ],
  },
  AVERAGEIFS: {
    name: 'AVERAGEIFS',
    syntax: 'AVERAGEIFS(average_range, criteria_range1, criteria1, [criteria_range2, criteria2], ...)',
    description: 'Returns the average of all cells that meet multiple criteria.',
    params: [
      { name: 'average_range', description: 'The actual cells to average.' },
      { name: 'criteria_range1', description: 'The first range to evaluate.' },
      { name: 'criteria1', description: 'The criteria to evaluate against criteria_range1.' },
      { name: 'criteria_range2', description: 'Additional range to evaluate.', optional: true },
      { name: 'criteria2', description: 'Additional criteria to evaluate.', optional: true },
    ],
  },
  MEDIAN: {
    name: 'MEDIAN',
    syntax: 'MEDIAN(number1, [number2], ...)',
    description: 'Returns the median of the given numbers (the middle value).',
    params: [
      { name: 'number1', description: 'First number, cell reference, or range.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  STDEV: {
    name: 'STDEV',
    syntax: 'STDEV(number1, [number2], ...)',
    description: 'Estimates standard deviation based on a sample (ignores logical values and text).',
    params: [
      { name: 'number1', description: 'First number, cell reference, or range of sample.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  'STDEV.S': {
    name: 'STDEV.S',
    syntax: 'STDEV.S(number1, [number2], ...)',
    description: 'Calculates standard deviation based on a sample (replaces STDEV).',
    params: [
      { name: 'number1', description: 'First number or range in the sample.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  'STDEV.P': {
    name: 'STDEV.P',
    syntax: 'STDEV.P(number1, [number2], ...)',
    description: 'Calculates standard deviation based on the entire population.',
    params: [
      { name: 'number1', description: 'First number or range in the population.' },
      { name: 'number2', description: 'Optional additional numbers or ranges.', optional: true },
    ],
  },
  IF: {
    name: 'IF',
    syntax: 'IF(logical_test, value_if_true, [value_if_false])',
    description: 'Checks whether a condition is met, and returns one value if TRUE, and another if FALSE.',
    params: [
      { name: 'logical_test', description: 'Any value or expression that can be evaluated to TRUE or FALSE.' },
      { name: 'value_if_true', description: 'The value that is returned if logical_test is TRUE.' },
      { name: 'value_if_false', description: 'The value that is returned if logical_test is FALSE.', optional: true },
    ],
  },
  IFS: {
    name: 'IFS',
    syntax: 'IFS(logical_test1, value_if_true1, [logical_test2, value_if_true2], ...)',
    description: 'Checks whether one or more conditions are met and returns a value corresponding to the first TRUE condition.',
    params: [
      { name: 'logical_test1', description: 'Condition that evaluates to TRUE or FALSE.' },
      { name: 'value_if_true1', description: 'Result if logical_test1 is TRUE.' },
      { name: 'logical_test2', description: 'Additional condition.', optional: true },
      { name: 'value_if_true2', description: 'Additional result.', optional: true },
    ],
  },
  SWITCH: {
    name: 'SWITCH',
    syntax: 'SWITCH(expression, val1, result1, [default_or_val2], ...)',
    description: 'Evaluates an expression against a list of values and returns the result corresponding to the first matching value.',
    params: [
      { name: 'expression', description: 'The value or expression to compare against.' },
      { name: 'val1', description: 'The first value to be compared against expression.' },
      { name: 'result1', description: 'The corresponding result if val1 matches expression.' },
      { name: 'default_or_val2', description: 'Default value if no match, or next value.', optional: true },
    ],
  },
  IFERROR: {
    name: 'IFERROR',
    syntax: 'IFERROR(value, value_if_error)',
    description: 'Returns value_if_error if expression is an error and the value of the expression itself otherwise.',
    params: [
      { name: 'value', description: 'The argument that is checked for an error.' },
      { name: 'value_if_error', description: 'The value to return if the formula evaluates to an error.' },
    ],
  },
  AND: {
    name: 'AND',
    syntax: 'AND(logical1, [logical2], ...)',
    description: 'Returns TRUE if all of its arguments are TRUE.',
    params: [
      { name: 'logical1', description: 'The first condition that you want to test.' },
      { name: 'logical2', description: 'Optional additional conditions.', optional: true },
    ],
  },
  OR: {
    name: 'OR',
    syntax: 'OR(logical1, [logical2], ...)',
    description: 'Returns TRUE if any argument is TRUE.',
    params: [
      { name: 'logical1', description: 'The first condition that you want to test.' },
      { name: 'logical2', description: 'Optional additional conditions.', optional: true },
    ],
  },
  NOT: {
    name: 'NOT',
    syntax: 'NOT(logical)',
    description: 'Reverses the logic of its argument.',
    params: [
      { name: 'logical', description: 'A value or expression that can be evaluated to TRUE or FALSE.' },
    ],
  },
  VLOOKUP: {
    name: 'VLOOKUP',
    syntax: 'VLOOKUP(lookup_value, table_array, col_index, [range_lookup])',
    description: 'Looks for a value in the leftmost column of a table, and returns a value in the same row from a column you specify.',
    params: [
      { name: 'lookup_value', description: 'The value to find in the first column of the table.' },
      { name: 'table_array', description: 'The table of information in which data is looked up.' },
      { name: 'col_index', description: 'The column number in table_array from which the matching value is returned.' },
      { name: 'range_lookup', description: 'TRUE for approximate match, FALSE for exact match (default TRUE).', optional: true },
    ],
  },
  XLOOKUP: {
    name: 'XLOOKUP',
    syntax: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], [match_mode], [search_mode])',
    description: 'Searches a range or array, and returns an item corresponding to the first match it finds.',
    params: [
      { name: 'lookup_value', description: 'The value to search for.' },
      { name: 'lookup_array', description: 'The array or range to search.' },
      { name: 'return_array', description: 'The array or range to return results from.' },
      { name: 'if_not_found', description: 'Returned if no match is found.', optional: true },
      { name: 'match_mode', description: 'Specify match type: 0 Exact match (default), -1 Exact or next smaller, 1 Exact or next larger, 2 Wildcard.', optional: true },
      { name: 'search_mode', description: 'Specify search order: 1 First to last (default), -1 Last to first.', optional: true },
    ],
  },
  INDEX: {
    name: 'INDEX',
    syntax: 'INDEX(array, row_num, [column_num])',
    description: 'Returns a value or reference of the cell at the intersection of a particular row and column in a given range.',
    params: [
      { name: 'array', description: 'A range of cells or an array constant.' },
      { name: 'row_num', description: 'Selects the row in array from which to return a value.' },
      { name: 'column_num', description: 'Selects the column in array from which to return a value.', optional: true },
    ],
  },
  MATCH: {
    name: 'MATCH',
    syntax: 'MATCH(lookup_value, lookup_array, [match_type])',
    description: 'Returns the relative position of an item in an array that matches a specified value.',
    params: [
      { name: 'lookup_value', description: 'The value that you want to match in lookup_array.' },
      { name: 'lookup_array', description: 'The range of cells being searched.' },
      { name: 'match_type', description: '1 less than, 0 exact match, -1 greater than.', optional: true },
    ],
  },
  SORT: {
    name: 'SORT',
    syntax: 'SORT(array, [sort_index], [sort_order])',
    description: 'Sorts the contents of a range or array.',
    params: [
      { name: 'array', description: 'The range or array to sort.' },
      { name: 'sort_index', description: 'A number indicating the row or column to sort by.', optional: true },
      { name: 'sort_order', description: '1 for ascending (default), -1 for descending.', optional: true },
    ],
  },
  UNIQUE: {
    name: 'UNIQUE',
    syntax: 'UNIQUE(array)',
    description: 'Returns a list of unique values in a list or range.',
    params: [
      { name: 'array', description: 'The range or array from which to return unique rows or values.' },
    ],
  },
  PMT: {
    name: 'PMT',
    syntax: 'PMT(rate, nper, pv, [fv], [type])',
    description: 'Calculates the payment for a loan based on constant payments and a constant interest rate.',
    params: [
      { name: 'rate', description: 'The interest rate for the loan per period.' },
      { name: 'nper', description: 'The total number of payments for the loan.' },
      { name: 'pv', description: 'The present value: total amount that a series of future payments is worth now.' },
      { name: 'fv', description: 'The future value, or cash balance after the last payment (default 0).', optional: true },
      { name: 'type', description: '0 = payment at end of period, 1 = payment at start (default 0).', optional: true },
    ],
  },
  ROUND: {
    name: 'ROUND',
    syntax: 'ROUND(number, num_digits)',
    description: 'Rounds a number to a specified number of digits.',
    params: [
      { name: 'number', description: 'The number you want to round.' },
      { name: 'num_digits', description: 'The number of digits to which you want to round.' },
    ],
  },
  ROUNDUP: {
    name: 'ROUNDUP',
    syntax: 'ROUNDUP(number, num_digits)',
    description: 'Rounds a number up, away from 0.',
    params: [
      { name: 'number', description: 'The number you want to round.' },
      { name: 'num_digits', description: 'The number of digits to which you want to round.' },
    ],
  },
  ROUNDDOWN: {
    name: 'ROUNDDOWN',
    syntax: 'ROUNDDOWN(number, num_digits)',
    description: 'Rounds a number down, toward 0.',
    params: [
      { name: 'number', description: 'The number you want to round.' },
      { name: 'num_digits', description: 'The number of digits to which you want to round.' },
    ],
  },
  INT: {
    name: 'INT',
    syntax: 'INT(number)',
    description: 'Rounds a number down to the nearest integer.',
    params: [
      { name: 'number', description: 'The real number you want to round down.' },
    ],
  },
  ABS: {
    name: 'ABS',
    syntax: 'ABS(number)',
    description: 'Returns the absolute value of a number.',
    params: [
      { name: 'number', description: 'The real number of which you want the absolute value.' },
    ],
  },
  MOD: {
    name: 'MOD',
    syntax: 'MOD(number, divisor)',
    description: 'Returns the remainder after number is divided by divisor.',
    params: [
      { name: 'number', description: 'The number for which you want to find the remainder.' },
      { name: 'divisor', description: 'The number by which you want to divide number.' },
    ],
  },
  POWER: {
    name: 'POWER',
    syntax: 'POWER(number, power)',
    description: 'Returns the result of a number raised to a power.',
    params: [
      { name: 'number', description: 'The base number.' },
      { name: 'power', description: 'The exponent to which the base number is raised.' },
    ],
  },
  SQRT: {
    name: 'SQRT',
    syntax: 'SQRT(number)',
    description: 'Returns a positive square root.',
    params: [
      { name: 'number', description: 'The number for which you want the square root.' },
    ],
  },
  SEQUENCE: {
    name: 'SEQUENCE',
    syntax: 'SEQUENCE(rows, [columns], [start], [step])',
    description: 'Generates a list of sequential numbers in an array.',
    params: [
      { name: 'rows', description: 'The number of rows to return.' },
      { name: 'columns', description: 'The number of columns to return (default 1).', optional: true },
      { name: 'start', description: 'The first number in the sequence (default 1).', optional: true },
      { name: 'step', description: 'The amount to increment each value (default 1).', optional: true },
    ],
  },
  CONCAT: {
    name: 'CONCAT',
    syntax: 'CONCAT(text1, [text2], ...)',
    description: 'Combines the text from multiple ranges and/or strings.',
    params: [
      { name: 'text1', description: 'Text item, cell reference, or range.' },
      { name: 'text2', description: 'Optional additional text items or ranges.', optional: true },
    ],
  },
  CONCATENATE: {
    name: 'CONCATENATE',
    syntax: 'CONCATENATE(text1, [text2], ...)',
    description: 'Joins several text strings into one text string.',
    params: [
      { name: 'text1', description: 'Text item or cell reference.' },
      { name: 'text2', description: 'Optional additional text items.', optional: true },
    ],
  },
  TEXTJOIN: {
    name: 'TEXTJOIN',
    syntax: 'TEXTJOIN(delimiter, ignore_empty, text1, [text2], ...)',
    description: 'Combines the text from multiple ranges and/or strings with a specified delimiter.',
    params: [
      { name: 'delimiter', description: 'A string to insert between each text item.' },
      { name: 'ignore_empty', description: 'If TRUE, ignores empty cells.' },
      { name: 'text1', description: 'First text item or range to join.' },
      { name: 'text2', description: 'Additional text items or ranges.', optional: true },
    ],
  },
  LEFT: {
    name: 'LEFT',
    syntax: 'LEFT(text, [num_chars])',
    description: 'Returns the specified number of characters from the start of a text string.',
    params: [
      { name: 'text', description: 'The text string containing the characters you want to extract.' },
      { name: 'num_chars', description: 'Specifies how many characters to extract (default 1).', optional: true },
    ],
  },
  RIGHT: {
    name: 'RIGHT',
    syntax: 'RIGHT(text, [num_chars])',
    description: 'Returns the specified number of characters from the end of a text string.',
    params: [
      { name: 'text', description: 'The text string containing the characters you want to extract.' },
      { name: 'num_chars', description: 'Specifies how many characters to extract (default 1).', optional: true },
    ],
  },
  MID: {
    name: 'MID',
    syntax: 'MID(text, start_num, num_chars)',
    description: 'Returns a specific number of characters from a text string, starting at the position you specify.',
    params: [
      { name: 'text', description: 'The text string containing the characters you want to extract.' },
      { name: 'start_num', description: 'The position of the first character you want to extract (1-based).' },
      { name: 'num_chars', description: 'Specifies the number of characters to extract.' },
    ],
  },
  LEN: {
    name: 'LEN',
    syntax: 'LEN(text)',
    description: 'Returns the number of characters in a text string.',
    params: [
      { name: 'text', description: 'The text whose length you want to find.' },
    ],
  },
  UPPER: {
    name: 'UPPER',
    syntax: 'UPPER(text)',
    description: 'Converts text to uppercase.',
    params: [
      { name: 'text', description: 'The text you want converted to uppercase.' },
    ],
  },
  LOWER: {
    name: 'LOWER',
    syntax: 'LOWER(text)',
    description: 'Converts text to lowercase.',
    params: [
      { name: 'text', description: 'The text you want converted to lowercase.' },
    ],
  },
  TRIM: {
    name: 'TRIM',
    syntax: 'TRIM(text)',
    description: 'Removes all spaces from text except for single spaces between words.',
    params: [
      { name: 'text', description: 'The text from which you want spaces removed.' },
    ],
  },
  VALUE: {
    name: 'VALUE',
    syntax: 'VALUE(text)',
    description: 'Converts a text string that represents a number to a number.',
    params: [
      { name: 'text', description: 'The text enclosed in quotation marks or a reference to a cell.' },
    ],
  },
  TODAY: {
    name: 'TODAY',
    syntax: 'TODAY()',
    description: 'Returns the current date formatted as a date serial number.',
    params: [],
  },
  NOW: {
    name: 'NOW',
    syntax: 'NOW()',
    description: 'Returns the current date and time serial number.',
    params: [],
  },
  DATE: {
    name: 'DATE',
    syntax: 'DATE(year, month, day)',
    description: 'Returns the serial number that represents a particular date.',
    params: [
      { name: 'year', description: 'The year of the date.' },
      { name: 'month', description: 'The month of the date (1 to 12).' },
      { name: 'day', description: 'The day of the date (1 to 31).' },
    ],
  },
  YEAR: {
    name: 'YEAR',
    syntax: 'YEAR(serial_number)',
    description: 'Returns the year of a date, an integer between 1900 and 9999.',
    params: [
      { name: 'serial_number', description: 'The date serial number or cell reference.' },
    ],
  },
  MONTH: {
    name: 'MONTH',
    syntax: 'MONTH(serial_number)',
    description: 'Returns the month of a date, a number from 1 to 12.',
    params: [
      { name: 'serial_number', description: 'The date serial number or cell reference.' },
    ],
  },
  DAY: {
    name: 'DAY',
    syntax: 'DAY(serial_number)',
    description: 'Returns the day of the month, a number from 1 to 31.',
    params: [
      { name: 'serial_number', description: 'The date serial number or cell reference.' },
    ],
  },
  DAYS: {
    name: 'DAYS',
    syntax: 'DAYS(end_date, start_date)',
    description: 'Returns the number of days between two dates.',
    params: [
      { name: 'end_date', description: 'The end date serial number or reference.' },
      { name: 'start_date', description: 'The start date serial number or reference.' },
    ],
  },
  ROW: {
    name: 'ROW',
    syntax: 'ROW([reference])',
    description: 'Returns the row number of a reference.',
    params: [
      { name: 'reference', description: 'The cell or range of cells for which you want the row number.', optional: true },
    ],
  },
  COLUMN: {
    name: 'COLUMN',
    syntax: 'COLUMN([reference])',
    description: 'Returns the column number of a reference.',
    params: [
      { name: 'reference', description: 'The cell or range of cells for which you want the column number.', optional: true },
    ],
  },
  ISBLANK: {
    name: 'ISBLANK',
    syntax: 'ISBLANK(value)',
    description: 'Returns TRUE if the value refers to an empty cell.',
    params: [
      { name: 'value', description: 'The value or cell reference you want to test.' },
    ],
  },
  ISNUMBER: {
    name: 'ISNUMBER',
    syntax: 'ISNUMBER(value)',
    description: 'Returns TRUE if the value refers to a number.',
    params: [
      { name: 'value', description: 'The value or cell reference you want to test.' },
    ],
  },
  ISTEXT: {
    name: 'ISTEXT',
    syntax: 'ISTEXT(value)',
    description: 'Returns TRUE if the value refers to text.',
    params: [
      { name: 'value', description: 'The value or cell reference you want to test.' },
    ],
  },
};

export const FORMULA_COLOR_PALETTE = [
  '#2563eb', // Royal Blue
  '#dc2626', // Crimson Red
  '#9333ea', // Vivid Purple
  '#16a34a', // Emerald Green
  '#d97706', // Warm Amber
  '#0891b2', // Deep Cyan
  '#ea580c', // Bright Orange
  '#be185d', // Fuchsia
];

export interface FormulaReferenceHighlight {
  token: string;
  color: string;
  range: Range;
  startIdx: number;
  endIdx: number;
}

/**
 * Extracts cell/range references from formula text and assigns distinct colors from the palette.
 */
export function extractFormulaHighlights(formulaText: string): FormulaReferenceHighlight[] {
  if (!formulaText.startsWith('=')) return [];
  const highlights: FormulaReferenceHighlight[] = [];
  const regex = /"(?:[^"]|"")*"|(?<![A-Za-z0-9_.\[])((?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)!)?(\$?[A-Za-z]{1,3}\$?\d+(?::\$?[A-Za-z]{1,3}\$?\d+)?)(?![A-Za-z0-9_.\](])/g;
  let match: RegExpExecArray | null;
  let colorIdx = 0;
  while ((match = regex.exec(formulaText)) !== null) {
    if (match[0].startsWith('"')) continue;
    const token = match[0];
    const range = parseRange(token);
    if (range) {
      highlights.push({
        token,
        color: FORMULA_COLOR_PALETTE[colorIdx % FORMULA_COLOR_PALETTE.length],
        range,
        startIdx: match.index,
        endIdx: match.index + token.length,
      });
      colorIdx++;
    }
  }
  return highlights;
}

export interface FormulaContext {
  isFormula: boolean;
  activeFunction?: FunctionMeta;
  activeParamIndex: number;
  suggestions: FunctionMeta[];
  prefix: string;
  prefixStart: number;
}

/**
 * Analyzes the formula text up to cursor position to provide function signature help
 * and autocomplete suggestions.
 */
export function getFormulaContext(formulaText: string, cursorPos: number): FormulaContext {
  if (!formulaText.startsWith('=')) {
    return { isFormula: false, activeParamIndex: 0, suggestions: [], prefix: '', prefixStart: 0 };
  }

  const textBeforeCursor = formulaText.slice(0, cursorPos);
  const textAfterEq = textBeforeCursor.slice(1);

  // Check if user is typing an identifier (function name autocomplete)
  const identMatch = /[A-Za-z_][A-Za-z0-9_.]*$/.exec(textAfterEq);
  let suggestions: FunctionMeta[] = [];
  let prefix = '';
  let prefixStart = 0;

  if (identMatch) {
    prefix = identMatch[0].toUpperCase();
    prefixStart = 1 + identMatch.index;
    const allFuncs = Object.values(FORMULA_FUNCTIONS);
    suggestions = allFuncs
      .filter((f) => f.name.startsWith(prefix))
      .sort((a, b) => a.name.localeCompare(b.name))
      .slice(0, 8);
  }

  // Parse parenthesis stack to find active function and active parameter index
  const stack: { name: string; paramIndex: number }[] = [];
  let inString = false;
  let currentWord = '';

  for (let i = 1; i < textBeforeCursor.length; i++) {
    const ch = textBeforeCursor[i];

    if (ch === '"') {
      if (inString && textBeforeCursor[i + 1] === '"') {
        i++; // skip escaped quote
      } else {
        inString = !inString;
      }
      currentWord = '';
      continue;
    }

    if (inString) continue;

    if (/[A-Za-z0-9_.]/.test(ch)) {
      currentWord += ch;
    } else {
      if (ch === '(') {
        const fnName = currentWord.toUpperCase();
        stack.push({ name: fnName, paramIndex: 0 });
      } else if (ch === ')') {
        if (stack.length > 0) stack.pop();
      } else if (ch === ',' || ch === ';') {
        if (stack.length > 0) {
          stack[stack.length - 1].paramIndex++;
        }
      }
      currentWord = '';
    }
  }

  let activeFunction: FunctionMeta | undefined;
  let activeParamIndex = 0;

  if (stack.length > 0) {
    const top = stack[stack.length - 1];
    activeFunction = FORMULA_FUNCTIONS[top.name];
    activeParamIndex = top.paramIndex;
  }

  return {
    isFormula: true,
    activeFunction,
    activeParamIndex,
    suggestions,
    prefix,
    prefixStart,
  };
}
