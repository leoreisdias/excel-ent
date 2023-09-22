import { CellStyle, ExcelEntCellObject } from "./cell";

export type ExcelEntContent =
  | ExcelEntCellObject
  | number
  | string
  | null
  | undefined;

export type MixedContent = {
  type: "row" | "column";
  value: ExcelEntContent[];
  style?: CellStyle;
};

type MixedStructure = {
  contentStructure: "mixed";
  content: MixedContent[];
};

type UniqueStructure = {
  contentStructure: "rows" | "columns";
  content: ExcelEntContent[][];
};

export type ExcelEntDataStructure = MixedStructure | UniqueStructure;

export type ExcelEntDataProps = {
  headerRow?: ExcelEntContent[];
} & ExcelEntDataStructure;

export type PaginatedObjectContentProps = {
  content: Record<string, any>[];
  sheetName: string;
};
