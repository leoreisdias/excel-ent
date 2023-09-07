import * as XLSX from "xlsx-js-style";

export type CellStyle = XLSX.CellStyle;

type BooleanCell = {
  type: "boolean";
  value: boolean;
};

type StringCell = {
  type: "string";
  value: string;
};

type NumberCell = {
  type: "number";
  value: number;
  formatted?: string;
  mask?: string;
};

type DateCell = {
  type: "date";
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
  style?: CellStyle;
};

export type ExcelEntCellObject = Cell &
  (BooleanCell | StringCell | NumberCell | DateCell);
