import * as XLSX from 'xlsx-js-style';

import { downloadFile, objectToSemicolons } from '../helpers/convert';
import {
    transformIntoCellObject,
    transposeMatrixWithPadding,
} from '../helpers/transform';
import {
    ExportationType,
    ExportMeContent,
    ExportMeExcelAdvancedProps,
    ExportMeExcelOptions,
    MergeProps,
} from '../types';

const executeXLSX = (
    data: XLSX.CellObject[][],
    columnWidths?: number[],
    merges?: MergeProps[],
) => {
    const wb = XLSX.utils.book_new();
    const ws = XLSX.utils.aoa_to_sheet(data);
    ws['!cols'] = columnWidths?.map((width) => ({ width }));
    ws['!merges'] = merges?.map((item) => ({
        s: { r: item.start.row, c: item.start.column },
        e: { r: item.end.row, c: item.end.column },
    }));

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

const getRowsStructure = (
    content: ExportMeContent[][],
    basedStructure: 'rows' | 'columns',
): ExportMeContent[][] => {
    if (basedStructure === 'rows') {
        return content;
    }

    const columnsToRows = transposeMatrixWithPadding(content);

    return columnsToRows;
};

export const exportmeExcelAdvanced = ({
    fileName,
    data,
    options,
    exportAs,
    merges,
}: ExportMeExcelAdvancedProps) => {
    const {
        bodyStyle = {},
        columnWidths,
        headerStyle = {},
        sheetProps,
    } = options ?? {};

    const { headerRow = [], content, contentStructure } = data;

    const headerXLSX: XLSX.CellObject[] = headerRow.map((cell) =>
        transformIntoCellObject(cell, headerStyle),
    );

    const rowsAdapter = getRowsStructure(content, contentStructure);

    const rowsXLSX: XLSX.CellObject[][] = rowsAdapter.map((item) =>
        item.map((cell) => transformIntoCellObject(cell, bodyStyle)),
    );

    const wb = executeXLSX([headerXLSX, ...rowsXLSX], columnWidths, merges);
    wb.Props = sheetProps;

    return exportFile(exportAs, wb, fileName);
};

export const exportmeExcel = (
    data: Record<string, any>[],
    fileName: string,
    exportAs: ExportationType,
    options?: ExportMeExcelOptions,
) => {
    const headers: XLSX.CellObject[] = Object.keys(data[0]).map((item) => ({
        v: item,
        t: 's',
        s: options?.headerStyle,
    }));

    const body: XLSX.CellObject[][] = data.map((item) =>
        Object.keys(item).map((key) => ({
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
