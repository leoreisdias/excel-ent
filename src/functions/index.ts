import * as XLSX from 'xlsx-js-style';

import {
  convertType,
  downloadFile,
  objectToSemicolons,
} from '../helpers/convert';
import {
  ExportationType,
  ExportMeExcelAdvancedProps,
  ExportMeExcelOptions,
} from '../types';

const validateData = (data: any[], fileName: string) => {
  if (
    !Array.isArray(data) ||
    typeof fileName !== 'string' ||
    Object.prototype.toString.call(fileName) !== '[object String]'
  ) {
    throw new Error(
      'Invalid input types: First Params should be an Array and the second one a String',
    );
  }
};

const executeXLSX = (data: XLSX.CellObject[][], columnWidths?: number[]) => {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(data);
  ws['!cols'] = columnWidths?.map(width => ({ width }));
  XLSX.utils.book_append_sheet(wb, ws);

  return wb;
};

const exportFile = (
  exportAs: ExportationType,
  wb: XLSX.WorkBook,
  fileName: string,
) => {
  if (exportAs.type === 'base64') {
    return XLSX.write(wb, { type: 'base64', bookType: 'xlsx' });
  }

  if (exportAs.type === 'buffer') {
    return XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
  }

  if (exportAs.type === 'download') {
    return XLSX.writeFile(wb, `${fileName}.xlsx`);
  }

  if (exportAs.type === 'filepath') {
    return XLSX.writeFile(wb, exportAs.path);
  }
};

export const exportmeExcelAdvanced = ({
  fileName,
  headers = [],
  rows,
  options,
  exportAs,
}: ExportMeExcelAdvancedProps) => {
  const {
    bodyStyle = {},
    columnWidths,
    headerStyle = {},
    sheetProps,
  } = options ?? {};

  const headerXLSX: XLSX.CellObject[] = headers.map(
    cell =>
      ({
        t: convertType(cell.type),
        v: cell.value,
        c: cell.comment?.map(comment => ({
          a: comment.author,
          t: comment.text,
        })),
        F: cell.formulaRange,
        f: cell.formula,
        l: cell.hyperlink && {
          Target: cell.hyperlink?.target,
          Tooltip: cell.hyperlink?.tooltip,
        },
        s: {
          ...(cell.style ?? {}),
          ...headerStyle,
        },
        z:
          cell.type === 'number' || cell.type === 'date'
            ? cell.mask
            : undefined,
        w:
          cell.type === 'number' || cell.type === 'date'
            ? cell.formatted
            : undefined,
      } as XLSX.CellObject),
  );

  const rowsXLSX: XLSX.CellObject[][] = rows.map(item =>
    item.map(
      cell =>
        ({
          t: convertType(cell.type),
          v: cell.value,
          c: cell.comment?.map(comment => ({
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
            ...(cell.style ?? {}),
            ...bodyStyle,
          },
          z:
            cell.type === 'number' || cell.type === 'date'
              ? cell.mask
              : undefined,
          w:
            cell.type === 'number' || cell.type === 'date'
              ? cell.formatted
              : undefined,
        } as XLSX.CellObject),
    ),
  );

  const wb = executeXLSX([headerXLSX, ...rowsXLSX], columnWidths);
  wb.Props = sheetProps;

  return exportFile(exportAs, wb, fileName);
};

export const exportmeExcel = (
  data: Record<string, any>[],
  fileName: string,
  exportAs: ExportationType,
  options?: ExportMeExcelOptions,
) => {
  validateData(data, fileName);

  const headers: XLSX.CellObject[] = Object.keys(data[0]).map(item => ({
    v: item,
    t: 's',
    s: options?.headerStyle,
  }));

  const body: XLSX.CellObject[][] = data.map(item =>
    Object.keys(item).map(key => ({
      v: item[key],
      t: 's',
      s: options?.bodyStyle,
    })),
  );

  const wb = executeXLSX([headers, ...body], options?.columnWidths);

  wb.Props = options?.sheetProps;

  return exportFile(exportAs, wb, fileName);
};

export const exportmeToCsv = (data: any[], fileName: string) => {
  if (
    typeof fileName !== 'string' ||
    Object.prototype.toString.call(fileName) !== '[object String]'
  ) {
    throw new Error(
      'Invalid input types: First Params should be an Array and the second one a String',
    );
  }

  if (window) {
    const computedCSV = new Blob([objectToSemicolons(data)], {
      type: 'text/csv;charset=utf-8',
    });

    const csvLink = window.URL.createObjectURL(computedCSV);
    downloadFile(csvLink, `${fileName}.csv`);
  } else {
    throw new Error(
      'Window is not definided: You must be using it in a browser',
    );
  }
};
