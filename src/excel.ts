/* eslint-disable no-console */
import * as XLSX from "xlsx-js-style";

import { getRowHeights, getRowsStructure } from "./helpers/rows";
import {
  applyStrippedRowStyle,
  transformIntoCellObject,
} from "./helpers/transform";
import { MergeProps } from "./types";
import {
  ExportationType,
  ExportMeExcelAdvancedProps,
  ExportMeExcelOptions,
} from "./types/functions";

const executeXLSX = (
  data: XLSX.CellObject[][],
  columnWidths?: number[],
  rowsHeights?: number[],
  merges?: MergeProps[]
) => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws["!cols"] = columnWidths?.map((width) => ({ width }));
  ws["!rows"] = rowsHeights?.map((height) => ({ hpx: height }));
  ws["!merges"] = merges?.map((item) => ({
    s: { r: item.start.row, c: item.start.column },
    e: { r: item.end.row, c: item.end.column },
  }));

  XLSX.utils.book_append_sheet(wb, ws);

  return wb;
};

const exportFile = (
  exportAs: ExportationType,
  wb: XLSX.WorkBook,
  fileName: string
) => {
  if (exportAs.type === "base64") {
    return XLSX.write(wb, { type: "base64", bookType: "xlsx" });
  }

  if (exportAs.type === "buffer") {
    return XLSX.write(wb, { type: "buffer", bookType: "xlsx" });
  }

  if (exportAs.type === "download") {
    return XLSX.writeFile(wb, `${fileName}.xlsx`);
  }

  if (exportAs.type === "filepath") {
    return XLSX.writeFile(wb, exportAs.path);
  }
};

export const exportmeExcelAdvanced = ({
  fileName,
  data,
  options,
  exportAs,
  merges,
  loggingMatrix,
}: ExportMeExcelAdvancedProps) => {
  const {
    bodyStyle = {},
    columnWidths = [],
    rowHeights = [],
    headerStyle = {},
    sheetProps,
    globalRowHeight,
  } = options ?? {};

  const { headerRow, ...dataStructure } = data;

  const headerXLSX: XLSX.CellObject[] | undefined = headerRow?.map((cell) =>
    transformIntoCellObject(cell, headerStyle)
  );

  const rowsAdapter = getRowsStructure(dataStructure);

  const rowsXLSX: XLSX.CellObject[][] = rowsAdapter.map((item) =>
    item.map((cell) => transformIntoCellObject(cell, bodyStyle))
  );

  const finalMatrix = headerXLSX ? [headerXLSX, ...rowsXLSX] : rowsXLSX;

  const wb = executeXLSX(
    finalMatrix,
    columnWidths,
    getRowHeights(rowHeights, rowsXLSX.length, globalRowHeight),
    merges
  );

  wb.Props = sheetProps;

  if (loggingMatrix) {
    console.info(
      `💡 Excel-Ent - Logging Matrix: ${JSON.stringify(rowsAdapter)}`
    );
  }

  return exportFile(exportAs, wb, fileName);
};

export const exportmeExcel = (
  data: Record<string, any>[],
  fileName: string,
  exportAs: ExportationType,
  options?: ExportMeExcelOptions & { stripedRows?: boolean }
) => {
  const {
    bodyStyle = {},
    columnWidths = [],
    rowHeights = [],
    headerStyle = {},
    sheetProps,
    globalRowHeight,
    stripedRows,
  } = options ?? {};

  const headers: XLSX.CellObject[] = Object.keys(data[0]).map((item) => ({
    v: item,
    t: "s",
    s: headerStyle,
  }));

  const body: XLSX.CellObject[][] = data.map((item, index) =>
    Object.keys(item).map((key) => {
      const isRowPainted = stripedRows && index % 2 === 0;

      return {
        v: item[key],
        t: "s",
        s: isRowPainted ? applyStrippedRowStyle(bodyStyle) : bodyStyle,
      };
    })
  );

  const finalMatrix = [headers, ...body];

  const wb = executeXLSX(
    finalMatrix,
    columnWidths,
    getRowHeights(rowHeights, finalMatrix.length, globalRowHeight)
  );

  wb.Props = sheetProps;

  return exportFile(exportAs, wb, fileName);
};
