# Excel-ent

<p align="left">
  <a href="#" target="blank"><img src="https://i.imgur.com/vkTQF5g.jpg" width="100" alt="NUMPOD" /></a>
</p>

[![NPM](https://img.shields.io/npm/v/@numpod/excel-ent)](https://www.npmjs.com/package/@numpod/excel-ent)

## Description

[excel-ent](https://github.com/leoreisdias/excel-ent.git) is a helper lib to export data in XLS and CSV.

## Installation

```bash
$ yarn add @numpod/excel-ent

# or with npm

$ npm install @numpod/excel-ent --save
```

## Using excel-ent

Two main functions <b>exportmeExcel</b> & <b>exportmeToCsv</b> 

### exportmeExcel

exportmeExcel(data: any[], title: string, options?: CSS.Properties)

#### Parameters

`data`
Required, must be an array of Object

`title`
Required, name of the archive generated

`options`
Optional, receives two attributes: <i>bodyStyle</i> and <i>headerStyle</i>

- `bodyStyle` and `headerStyle`: Both receives CSS Properties in <b>camelCase</b> and provide style for the header and body in the Excel File

### Example

```js
import { exportmeExcel } from "@numpod/excel-ent";

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
  bodyStyle: { fontSize: '20px' },
  headerStyle: { fontSize: '40px' },
});

```

### Example output
<img src="./excel-ent-example.png" alt="Example Output - Excel Screenshot" />

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

exportmeToCSV(data, "MyReport");

```

---

## License

@numpod/excel-ent is [MIT licensed](LICENSE).

---

## Thank you and be free to contribute.