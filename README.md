# Excel-ent

[![NPM](https://img.shields.io/npm/v/excel-ent)](https://www.npmjs.com/package/excel-ent)
[![npm](https://img.shields.io/npm/l/excel-ent)](https://github.com/leoreisdias/excel-ent/blob/main/LICENSE)

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
exportmeExcel(data: any[], title: string, options?: {
  headerStyle?: XLSX.CellStyle;
  bodyStyle?: XLSX.CellStyle;
  columnWidths?: number[];
  sheetProps?: XLSX.FullProperties;
  exportAs: {
    type: 'buffer' | 'base64' | 'download' | 'filepath';
    path?: string; // Required if exportAs type is 'filepath'
  };
})
```

Parameters

- `data`: Required, must be an array of objects.
- `title`: Required, the name of the generated file.
- `options`: Optional, receives the following attributes:
  - `headerStyle` and `bodyStyle`: Both receive styles in the format of XLSX.CellStyle. You can check the available options [here in the xlsx-js-style](https://github.com/gitbrent/xlsx-js-style#cell-style-properties)
  - `columnWidths`: An array of numeric values indicating the minimum width for each column.
  - `sheetProps`: Additional properties for the worksheet, following XLSX.FullProperties. You can check the [official docs for more details...](https://docs.sheetjs.com/docs/csf/book#file-properties)
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

exportmeExcel(data, 'test', {
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
  columnWidths: [15, 20, 15], // Example column widths
  sheetProps: { Title: 'My Worksheet', Author: 'John Doe' }, // Example sheet properties
  exportAs: { type: 'download' }, // Example export type
});
```

### exportmeToCsv
`exportmeToCsv(data: any[], title: string)`

Parameters

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
  fileName: string;
  options?: {
    headerStyle?: XLSX.CellStyle;
    bodyStyle?: XLSX.CellStyle;
    columnWidths?: number[];
    sheetProps?: XLSX.FullProperties;
    exportAs: {
      type: 'buffer' | 'base64' | 'download' | 'filepath';
      path?: string; // Required if exportAs type is 'filepath'
    };
  };
  headers?: ExcelEntCellObject[];
  rows: ExcelEntCellObject[][];
})
```

Parameters

- `fileName`: Required, the name of the generated file.
- `options`: Optional, similar to the options in exportmeExcel.
- `headers`: Optional, an array of ExcelEntCellObject to specify header cells.
- `rows`: Required, a matrix of ExcelEntCellObject representing the data.

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

You can import the `ExcelEntCellObject` type to assist in usage and preparation.

[XLSX.CellStyles properties can be found here.](https://github.com/gitbrent/xlsx-js-style#cell-style-properties)

#### Example

Here's an example of how to use `exportmeExcelAdvanced`:

```ts
import { exportmeExcelAdvanced } from "excel-ent";

const rows: ExcelEntCellObject[][] = market.map(
  (item, index) =>
    [
      {
        type: 'string',
        value: item.name,
      },
      {
        type: 'number',
        value: item.value,
        formatted: maskCurrency(item.value),
        mask: 'R$ #.##0,00',
      },
    ] as ExcelEntCellObject[],
);

return exportmeExcelAdvanced({
  fileName: `Products`,
  options: {
    sheetProps: {
      Title: `Market Products`,
    },
    columnWidths: [50], // Setting width only for the first column
  },
  headers: [
    {
      type: 'string',
      value: 'Product',
      comment: [
        {
          author: 'ExcelEnt',
          text: 'Value',
        },
      ],
      style: {
        border: {
          bottom: {
            color: {
              rgb: '000000',
            },
            style: 'thin',
          },
        },
      },
    },
    {
      type: 'string',
      value: 'Valor',
      style: {
        border: {
          bottom: {
            color: {
              rgb: '000000',
            },
            style: 'thin',
          },
        },
      },
    },
  ],
  rows: rows,
  exportAs: {
    type: 'download',
  },
});
```

Example 2

```ts
import { exportmeExcelAdvanced } from "excel-ent";

const data = [
  {
    id: 4,
    name: 'Another Product',
    value: 45.99,
  },
  {
    id: 5,
    name: 'Yet Another Product',
    value: 59.99,
  },
];

const rows: ExcelEntCellObject[][] = data.map((item, index) => [
  {
    type: 'string',
    value: item.name,
    formula: '=A2&B2'
  },
  {
    type: 'number',
    value: item.value,
    formatted: maskCurrency(item.value),
    mask: 'R$ #.##0,00',
    hyperlink: {
      target: 'https://example.com',
      tooltip: 'Visit website'
    },
    comment: [
      {
        author: 'User',
        text: 'This is a comment'
      }
    ]
  },
] as ExcelEntCellObject[]);

return exportmeExcelAdvanced({
  fileName: `AdditionalProducts`,
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
      Title: `Additional Market Products`,
    },
  },
  headers: [
    {
      type: 'string',
      value: 'Product',
      comment: [
        {
          author: 'ExcelEnt',
          text: 'Value',
        },
      ],
      style: {
        fill: {
          bgColor: {
            rgb: 'FFFF00',
          },
        },
      },
    },
    {
      type: 'string',
      value: 'Valor'
    },
  ],
  rows: rows,
  exportAs: {
    type: 'buffer',
  },
});
```

## License

excel-ent is [MIT licensed](LICENSE).

---

## Acknowledgments

Special thanks to the following libraries for their invaluable contributions:

- [SheetJS CE](https://github.com/SheetJS/sheetjs): A fundamental framework that serves as the backbone for Excel-ent, providing powerful export functionality and a wide range of features.

- [xlsx-js-style](https://github.com/gitbrent/xlsx-js-style): This library played a pivotal role in enabling us to seamlessly incorporate styling and formatting options into Excel-ent.

