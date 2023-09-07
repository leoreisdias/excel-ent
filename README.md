# Excel-ent

[![NPM](https://img.shields.io/npm/v/excel-ent)](https://www.npmjs.com/package/excel-ent)
[![npm](https://img.shields.io/npm/l/excel-ent)](https://github.com/leoreisdias/excel-ent/blob/main/LICENSE)

# Sumário

- [Description](#description)
- [Installation](#installation)
- [Using excel-ent](#using-excel-ent)
  - [exportmeExcel](#exportmeexcel)
    - [Parameters](#parameters)
    - [Example](#example)
  - [exportmeToCsv](#exportmetocsv)
    - [Parameters](#parameters-1)
    - [Example](#example-1)
  - [exportmeExcelAdvanced](#exportmeexceladvanced)
    - [Parameters](#parameters-2)
    - [Understanding `exportmeExcelAdvanced` exclusive props](#exportmeexceladvanced-exclusive-props)
      - [`merges`](#merges)
      - [`loggingMatrix`](#loggingmatrix)
      - [the `data` property](#exportmeexceladvanced-the-data-property)
    - [About `ExcelEntContent`](#about-excelentcontent)
      - [Example](#example-2)
- [Types](#types)
- [License](#license)
- [Acknowledgments](#acknowledgments)

### Version 4

<img src="https://i.pinimg.com/originals/56/b7/64/56b7642388e52d62910a8806a18e10be.gif" alt="Above version 4 - Gear Fourth" width="160" height="100">

## Description

[excel-ent](https://github.com/leoreisdias/excel-ent.git) is a helper library that simplifies exporting data to XLSX and CSV using the SheetJS CE library.

## Installation

```bash
$ yarn add excel-ent

# or with npm

$ npm install excel-ent --save
```

## Using excel-ent

Excel-ent provides three main functions for exporting data: `exportmeExcel`, `exportmeToCsv` and `exportmeExcelAdvanced`.

### exportmeExcel

```ts
exportmeExcel(data: any[], fileName: string, 
  exportAs: {
    type: 'buffer' | 'base64' | 'download' | 'filepath';
    path?: string; // Required if exportAs type is 'filepath'
  }, 
  options?: {
    headerStyle?: XLSX.CellStyle;
    bodyStyle?: XLSX.CellStyle;
    columnWidths?: number[];
    rowHeights?: number[];
    globalRowHeight?: number;
    sheetProps?: XLSX.FullProperties;
    stripedRows?: boolean;
})
```

#### Parameters

- `data`: Required, must be an array of objects.
- `fileName`: Required, the name of the generated file.
- `options`: Optional, receives the following attributes:
  - `headerStyle` and `bodyStyle`: Both receive styles in the format of XLSX.CellStyle. You can check the available options [here in the xlsx-js-style](https://github.com/gitbrent/xlsx-js-style#cell-style-properties)
  - `columnWidths`: An array of numeric values indicating the minimum width for each column.
  - `rowHeights`: An array of numeric values indicating the minimum height for each row.
  - `globalRowHeight`: A numeric value that sets a minimum height for ALL rows in the matrix.
  - `sheetProps`: Additional properties for the worksheet, following XLSX.FullProperties. You can check the [official docs for more details...](https://docs.sheetjs.com/docs/csf/book#file-properties)
  - `stripedRows`: Optional, alternates row colors between white (customizable via cell styling) and light gray (F2F2F2) to improve data readability.
- `exportAs`: An object specifying how to export the file, with a type attribute that can be 'buffer', 'base64', 'download', or 'filepath'. If 'filepath' is chosen, the path attribute becomes required.

#### Example

```ts
import { exportmeExcel } from "excel-ent";

const data = [
  {
    id: 1,
    name: 'Some Name',
    age: 21,
  },
  {
    id: 2,
    name: 'Some New Name',
    age: 23,
  },
  {
    id: 3,
    name: 'Some Name Again',
    age: 22,
  },
];

 exportmeExcel(
  data,
  'Test File',
  {
    type: "download",
  },
  {
    columnWidths: [30, 30, 30], // Each of 3 rows width
    globalRowHeight: 20, // Height for ALL rows
    headerStyle: {
      fill: {
        fgColor: {
          rgb: "0a1c3e", // Color can't have the '#'
        },
      },
      font: {
        bold: true,
        color: {
          rgb: "ffffff", // Must have at least 6 letters (FFF wouldn't work) 
        },
      },
      alignment: {
        vertical: "center",
        horizontal: "center",
      },
    },
    bodyStyle: {
      font: {
        name: "sans-serif",
      },
      alignment: {
        vertical: "center",
        horizontal: "center",
      },
    },
    stripedRows: true,
  }
);
```

Output

![Example output](https://i.imgur.com/kqn1oZn.png)

---

### exportmeToCsv
`exportmeToCsv(data: any[], title: string)`

#### Parameters

- `data`: Required, must be an array of objects.
- `title`: Required, the name of the generated CSV file.

#### Example
```ts
import { exportmeToCsv } from "excel-ent";

const data = [
  {
    id: 1,
    name: "Some Name",
    age: 21,
  },
  {
    id: 2,
    name: "Some New Name",
    age: 23,
  },
  {
    id: 3,
    name: "Some Name Again",
    age: 22,
  },
];

exportmeToCsv(data, "MyReport");
```

### exportmeExcelAdvanced
```ts
exportmeExcelAdvanced(options: {
  data: ExcelMeDataProps;
  exportAs: ExportationType;
  merges?: MergeProps[];
  options?: ExportMeExcelOptions;
  loggingMatrix?: boolean;
  fileName: string;
  options?: {
    headerStyle?: XLSX.CellStyle;
    bodyStyle?: XLSX.CellStyle;
    columnWidths?: number[];
    rowHeights?: number[];
    globalRowHeight?: number;
    sheetProps?: XLSX.FullProperties;
  };
})
```

#### Parameters

- `fileName`: Required, the name of the generated file.
- `options`: Optional, receives the following attributes:
  - `headerStyle` and `bodyStyle`: Both receive styles in the format of XLSX.CellStyle. You can check the available options [here in the xlsx-js-style](https://github.com/gitbrent/xlsx-js-style#cell-style-properties)
  - `columnWidths`: An array of numeric values indicating the minimum width for each column.
  - `rowHeights`: An array of numeric values indicating the minimum height for each row.
  - `globalRowHeight`: A numeric value that sets a minimum height for ALL rows in the matrix.
  - `sheetProps`: Additional properties for the worksheet, following XLSX.FullProperties. You can check the [official docs for more details...](https://docs.sheetjs.com/docs/csf/book#file-properties)
- `exportAs`: An object specifying how to export the file, with a type attribute that can be 'buffer', 'base64', 'download', or 'filepath'. If 'filepath' is chosen, the path attribute becomes required.
- `merges`: Optional, merges cells within the worksheet based on specified start and end coordinates.
`loggingMatrix`: Optional, logs the resulting matrix before export for debugging purposes.
- `data`: Required, defines the data structure.

### `exportmeExcelAdvanced` exclusive props

#### `merges`

The merges property is an optional attribute that can be used with the `exportmeExcelAdvanced` function in the Excel-ent library. It accepts an array of MergeProps, where each MergeProps object defines the starting and ending coordinates (row and column) for merging cells within the worksheet.

    Default Value: None (No cell merging by default).
    Usage: By providing an array of MergeProps, you can merge specific cells in the worksheet, improving the visual organization of data.

Usage example

```ts
import { exportmeExcelAdvanced } from "excel-ent";

const merges = [
  { start: { row: 1, column: 1 }, end: { row: 1, column: 3 } }, // Merge cells in the first row from column 1 to 3
  { start: { row: 2, column: 2 }, end: { row: 3, column: 2 } }, // Merge cells in the second and third rows in column 2
];

return exportmeExcelAdvanced({
  fileName: `MergedCellsData`,
  options: {
    ...other,
    merges: merges, // Merge specified cells
  },
  // ... (other configurations)
});
```

#### `loggingMatrix`

When set to true, it enables the logging of the resulting matrix in the browser or server log just before exporting. This feature is designed to assist in debugging and understanding the data structure that will be used for export.

    Default Value: Disabled (false).
    Usage: By configuring loggingMatrix as true, you can view the matrix content in the log, which can be helpful for debugging and troubleshooting any issues related to data formatting or structure.

---
#### `exportmeExcelAdvanced`: the `data` property

The "data" property is a fundamental aspect of the Excel-ent library, and it plays a pivotal role in configuring the content structure of the worksheet to be exported. It is strongly typed using the `ExcelEntDataProps` interface.

`ExcelEntDataProps`:

- `headerRow` (Optional): An array of `ExcelEntContent`, which can be of type ExcelEntCellObject (as previously documented), or a number, string, null, or undefined.
- `contentStructure`: Determines the content structure within the worksheet and can be one of the following values: **"rows"**, **"columns"** or **"mixed."**

**Content Structure**:
The "content" property varies based on the selected "contentStructure."

When `contentStructure` is "rows":
- `content` is a matrix of `ExcelEntContent`. Each element can be an `ExcelEntCellObject`, number, string, null, or undefined. In this structure, each row in the content matrix corresponds to a row in the worksheet.

Example:
```ts
data: {
  contentStructure: "rows",
  headerRow: ['ID', 'Name', 'Age'],
  content: [
    [1, 'John'],
    [2,, 28],
    [3, 'Bob', 35],
  ],
}
```

Resulting Matrix and Excel structure:

```ts
[  
  ['ID', 'Name', 'Age'],
  [1, 'John',],
  [2,, 28], // null or undefined values result in empty cells
  [3, 'Bob', 35]
]
```

When `contentStructure` is "columns":
- `content` is a matrix of `ExcelEntContent`. Each element can be an `ExcelEntCellObject`, number, string, null, or undefined. In this structure, the Excel-ent library will transform the column matrix into a row matrix by taking the transpose of the original matrix. Each column in the content matrix corresponds to a row in the worksheet.

Example:
```ts
data: {
  contentStructure: "columns",
  headerRow: ['ID', 'Name', 'Age'],
  content: [
    [1, 2, 3],
    ['John', 'Alice', 'Bob'],
    [30, 28, 35],
  ],
}
```

Resulting Matrix (Transpose of the Original Matrix) and Excel structure:
```ts
[  
  ['ID', 'Name', 'Age'],
  [1, 'John', 30],
  [2, 'Alice', 28],
  [3, 'Bob', 35]
]
```

When contentStructure is "mixed":
- `content` is an array of `MixedContent` objects. Each `MixedContent` object represents either a row or a column in the matrix, specified by the `type` attribute.

Example:
```ts
data: {
  contentStructure: "mixed",
  content: [
    {
      type: 'row',
      value: ['ID', 'Name', 'Age'],
    },
    {
      type: 'column',
      value: [1, 2, 3],
    },
    {
      type: 'column',
      value: ['John', 'Alice', 'Bob'],
    },
    {
      type: 'column',
      value: [30, 28, 35],
    },
    {
      type: 'row',
      value: ['some cell'],
    },
  ],
}
```

Resulting Matrix and Excel structure:
```ts
[  
  ['ID', 'Name', 'Age'],
  [1, 'John', 30],
  [2, 'Alice', 28],
  [3, 'Bob', 35],
  ['some cell',,]
]
```

In the "mixed" content structure, you can specify whether each element represents a row or a column, allowing for flexible data organization. The Excel-ent library will handle the transformation accordingly.

The "data" property is central to configuring the content structure of your worksheet, offering flexibility in how you structure your data for export.

---

#### About `ExcelEntContent`

The "ExcelEntContent" type is used to define the content of individual cells in the Excel worksheet. It can take on various forms, including numbers, strings, null, undefined, or an object of type "ExcelEntCellObject."

The `ExcelEntCellObject` type is as follows:

```ts
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
```

#### Example

Here's an example of how to use `exportmeExcelAdvanced`:

```ts
import { exportmeExcelAdvanced } from "excel-ent";

const data = [
  {
    id: 1,
    name: "John",
    age: 30,
  },
  {
    id: 2,
    name: "Alice",
    age: 28,
  },
  {
    id: 3,
    name: "Bob",
    age: 35,
  },
];

const content: ExcelEntContent[][] = data.map((item, index) => [
  item.id,
  {
    type: 'string',
    value: item.name,
    style: {
        fill: {
          bgColor: {
            rgb: 'FFFF00',
          },
        },
    },
  },
  {
    type: 'number',
    value: item.value,
    formatted: `${item.value} years`,
    comment: [
      {
        author: 'User',
        text: 'This is a comment'
      }
    ]
  },
] as ExcelEntContent[]);

return exportmeExcelAdvanced({
  fileName: `Example`,
  options: {
    headerStyle: { 
      font: {
        sz: 40
      }
    },
    bodyStyle: { 
      font: {
        sz: 16
      }
    },
    sheetProps: {
      Title: `Additional Info`,
    },
  },
  data: {
      contentStructure: "column",
      headerRow: ['ID', 'Name', 'Age'],
      content: content
  },
  exportAs: {
    type: 'buffer',
  },
});
```

---
### Types

You can import the Excel-ent types to assist in usage and preparation.

[XLSX.CellStyles properties can be found here.](https://github.com/gitbrent/xlsx-js-style#cell-style-properties)

## License

excel-ent is [MIT licensed](LICENSE).

---

## Acknowledgments

Special thanks to the following libraries for their invaluable contributions:

- [SheetJS CE](https://github.com/SheetJS/sheetjs): A fundamental framework that serves as the backbone for Excel-ent, providing powerful export functionality and a wide range of features.

- [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style): This library played a pivotal role in enabling us to seamlessly incorporate styling and formatting options into Excel-ent.


### Thank you and be free to contribute