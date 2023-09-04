import * as XLSX from 'xlsx-js-style';

type BooleanCell = {
  type: 'boolean';
  value: boolean;
};

type StringCell = {
  type: 'string';
  value: string;
};

type NumberCell = {
  type: 'number';
  value: number;
  formatted?: string;
  mask?: string;
};

type DateCell = {
  type: 'date';
  value: Date | string;
  formatted?: string;
  mask?: string;
};

type Cell = {
  formula?: string;
  formulaRange?: `${string}:${string}`;
  hyperlink?: {
    target: string;
    tooltip?: string;
  };
  comment?: {
    author: string;
    text: string;
  }[];
  style?: XLSX.CellStyle;
};

export type ExcelEntCellObject = Cell &
  (BooleanCell | StringCell | NumberCell | DateCell);

export type ExportMeExcelOptions = {
  headerStyle?: XLSX.CellStyle;
  bodyStyle?: XLSX.CellStyle;
  columnWidths?: number[];
  sheetProps?: XLSX.FullProperties;
};

export type ExportationType =
  | {
      type: 'buffer' | 'base64' | 'download';
    }
  | {
      type: 'filepath';
      path: string;
    };

export type ExportMeExcelAdvancedProps = {
  headers?: ExcelEntCellObject[];
  rows: ExcelEntCellObject[][];
  fileName: string;
  exportAs: ExportationType;
  options?: ExportMeExcelOptions;
};
