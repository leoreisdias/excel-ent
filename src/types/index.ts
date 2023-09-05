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

export type MergeProps = {
    start: { row: number; column: number };
    end: { row: number; column: number };
};

export type ExportMeContent = ExcelEntCellObject | number | string;

export type ExportMeExcelAdvancedProps = {
    data: {
        headerRow?: ExportMeContent[];
        content: ExportMeContent[][];
        contentStructure: 'rows' | 'columns';
    };
    fileName: string;
    exportAs: ExportationType;
    merges?: MergeProps[];
    options?: ExportMeExcelOptions;
};
