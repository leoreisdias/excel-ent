import * as XLSX from "xlsx-js-style";

import { CellStyle } from "./cell";
import { ExcelEntDataProps, PaginatedObjectContentProps } from "./contents";

export type ExportMeExcelOptions = {
  headerStyle?: CellStyle;
  bodyStyle?: CellStyle;
  columnWidths?: number[];
  rowHeights?: number[];
  globalRowHeight?: number;
  sheetProps?: XLSX.FullProperties;
};

export type ExportationType =
  | {
      type: "buffer" | "base64" | "download";
    }
  | {
      type: "filepath";
      path: string;
    };

export type MergeProps = {
  start: { row: number; column: number };
  end: { row: number; column: number };
};

export type ExportMeExcelAdvancedProps = {
  data: ExcelEntDataProps;
  fileName: string;
  exportAs: ExportationType;
  merges?: MergeProps[];
  options?: ExportMeExcelOptions;
  loggingMatrix?: boolean;
};

export type ExportMeExcelProps = {
  data: Record<string, any>[] | PaginatedObjectContentProps[];
  fileName: string;
  exportAs: ExportationType;
  options?: ExportMeExcelOptions & { stripedRows?: boolean };
};
