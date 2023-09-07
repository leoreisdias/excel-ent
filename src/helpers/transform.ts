import * as XLSX from "xlsx-js-style";

import { ExcelEntContent } from "../types";
import { convertType } from "./convert";

export const transformIntoCellObject = (
  cell: ExcelEntContent,
  globalStyle: XLSX.CellStyle
): XLSX.CellObject => {
  if (!cell) {
    return {
      t: "z",
    };
  }

  if (typeof cell === "number") {
    return {
      t: "n",
      v: cell,
      s: globalStyle,
    };
  }

  if (typeof cell === "string") {
    return {
      t: "s",
      v: cell,
      s: globalStyle,
    };
  }

  const { style: cellStyle = {} } = cell;

  return {
    t: convertType(cell.type),
    v: cell.value,
    c: cell.comment?.map((comment) => ({
      t: comment.text,
      a: comment.author,
    })),
    F: cell.formulaRange,
    f: cell.formula,
    l: cell.hyperlink && {
      Target: cell.hyperlink.target,
      Tooltip: cell.hyperlink.tooltip,
    },
    s: {
      ...globalStyle,
      ...cellStyle,
    },
    z: cell.type === "number" || cell.type === "date" ? cell.mask : undefined,
    w:
      cell.type === "number" || cell.type === "date"
        ? cell.formatted
        : undefined,
  };
};

export const applyStrippedRowStyle = (
  style: XLSX.CellStyle | undefined
): XLSX.CellStyle => {
  if (!style)
    return {
      fill: {
        fgColor: { rgb: "F2F2F2" },
      },
    };

  return {
    ...style,
    fill: {
      ...style.fill,
      fgColor: { rgb: "F2F2F2" },
    },
  };
};

export const transposeMatrixWithPadding = (
  matrix: ExcelEntContent[][]
): ExcelEntContent[][] => {
  const numRows = matrix.length;
  const numCols = matrix.reduce((max, row) => Math.max(max, row.length), 0);

  const transposed = [];
  for (let col = 0; col < numCols; col++) {
    const newRow = [];
    for (let row = 0; row < numRows; row++) {
      const cellValue = matrix[row][col];
      if (cellValue !== undefined) {
        newRow.push(cellValue);
      } else {
        // Determine the maximum column length in this column
        const maxColumnLength = matrix.reduce(
          (max, currentRow) =>
            currentRow[col] !== undefined
              ? Math.max(max, currentRow[col]?.toString()?.length ?? 0)
              : max,
          0
        );

        // Pad the cell with spaces to match the maximum column length
        newRow.push(" ".repeat(maxColumnLength));
      }
    }
    transposed.push(newRow);
  }

  return transposed;
};
