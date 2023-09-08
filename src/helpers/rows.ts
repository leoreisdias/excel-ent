import { ExcelEntContent, ExcelEntDataStructure, MixedContent } from "../types";
import { transposeMatrixWithPadding } from "./transform";

const getNextRow = (currentIndex: number, content: MixedContent[]) => {
  let nextRowIndex = currentIndex + 1;
  for (nextRowIndex; nextRowIndex < content.length; nextRowIndex++) {
    if (content[nextRowIndex].type === "row") {
      return nextRowIndex - 1;
    }
  }

  return -1;
};

const handleMixedStructure = (data: ExcelEntDataStructure) => {
  if (data.contentStructure !== "mixed") return [];

  const rows: ExcelEntContent[][] = [];

  let currentIndex = 0;
  for (currentIndex; currentIndex < data.content.length; currentIndex++) {
    const rowColumn = data.content[currentIndex];

    if (!rowColumn) break;

    if (rowColumn.type === "row") {
      rows.push(rowColumn.value);
    }

    if (rowColumn.type === "column") {
      const nextRowIndex = getNextRow(currentIndex, data.content);

      const columns = data.content.slice(
        currentIndex,
        nextRowIndex !== -1 ? nextRowIndex : undefined
      );

      const columnsContent = columns.map((item) => item.value);

      const transposed = transposeMatrixWithPadding(columnsContent);

      rows.push(...transposed);
      
      if (nextRowIndex === -1) break;

      currentIndex = nextRowIndex;
    }
  }

  return rows;
};

export const getRowsStructure = (
  data: ExcelEntDataStructure
): ExcelEntContent[][] => {
  if (data.contentStructure === "rows") {
    return data.content;
  }

  if (data.contentStructure === "columns")
    return transposeMatrixWithPadding(data.content);

  if (data.contentStructure === "mixed") {
    return handleMixedStructure(data);
  }

  return [];
};

export const getRowHeights = (
  rowHeights: number[],
  maxRows: number,
  globalRowHeight?: number
) => {
  if (globalRowHeight) {
    return Array.from({ length: maxRows }).map(() => globalRowHeight);
  }

  return rowHeights;
};
