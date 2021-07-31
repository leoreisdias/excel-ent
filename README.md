# Excel-ent

## Description

[excel-ent](https://github.com/leoreisdias/excel-ent.git) is a helper lib to export data in XLS and CSV.

## Installation

```bash
$ npm install @numpod/excel-ent --save

# or with yarn

$ yarn add @numpod/excel-ent
```

## Using ezdate

Two main functions - exportmeToCsv & exportmeExcel

### exportmeExcel

exportmeExcel(data: any[], title: string)

#### Parameters

`data`
Required, must be an array of Object

`title`
Required, name of the archive generated

#### Example

```js
import { exportmeExcel } from "@numpod/excel-ent";

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

function handleExport() {
  exportmeExcel(data, "MyReport");
}
```

---

### exportmeToCsv

exportmeToCsv(data: any[], title: string)

#### Parameters

`data`
Required, must be an array of Object

`title`
Required, name of the archive generated

#### Example

```js
import { exportmeToCsv } from "@numpod/excel-ent";

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

function handleExport() {
  exportmeToCSV(data, "MyReport");
}
```

---

---
