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

export type ExportationTypeBase64 = { type: "base64" };
export type ExportationTypeBuffer = { type: "buffer" };
export type ExportationTypeDownload = { type: "download" };
export type ExportationTypeFilePath = { type: "filepath"; path: string };

export type ExportationType =
  | ExportationTypeBase64
  | ExportationTypeBuffer
  | ExportationTypeDownload
  | ExportationTypeFilePath;

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

type ExcelAdvancedOptions = {
  data: ExcelEntDataProps;
  fileName: string;
  merges?: MergeProps[];
  options?: ExportMeExcelOptions;
  loggingMatrix?: boolean;
};

export type ExcelAdvancedOptionsBase64 = ExcelAdvancedOptions & {
  exportAs: ExportationTypeBase64;
};
export type ExcelAdvancedOptionsBuffer = ExcelAdvancedOptions & {
  exportAs: ExportationTypeBuffer;
};
export type ExcelAdvancedOptionsDownload = ExcelAdvancedOptions & {
  exportAs: ExportationTypeDownload;
};
export type ExcelAdvancedOptionsFilePath = ExcelAdvancedOptions & {
  exportAs: ExportationTypeFilePath;
};

type ExcelProps = {
  data: Record<string, any>[] | PaginatedObjectContentProps[];
  fileName: string;
  options?: ExportMeExcelOptions & { stripedRows?: boolean };
};

export type ExcelOptionsBase64 = ExcelProps & {
  exportAs: ExportationTypeBase64;
};
export type ExcelOptionsBuffer = ExcelProps & {
  exportAs: ExportationTypeBuffer;
};
export type ExcelOptionsDownload = ExcelProps & {
  exportAs: ExportationTypeDownload;
};
export type ExcelOptionsFilePath = ExcelProps & {
  exportAs: ExportationTypeFilePath;
};
