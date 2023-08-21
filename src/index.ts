import * as CSS from "csstype";

export interface IExportmeExcelOptions {
  headerStyle?: CSS.Properties;
  bodyStyle?: CSS.Properties;
}

export const exportmeExcel = (
  data: any[],
  fileName: string,
  options?: IExportmeExcelOptions
) => {
  if (
    !Array.isArray(data) ||
    typeof fileName !== "string" ||
    Object.prototype.toString.call(fileName) !== "[object String]"
  ) {
    throw new Error(
      "Invalid input types: First Params should be an Array and the second one a String"
    );
  }

  const TEMPLATE_XLS = `
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="https://www.w3.org/TR/html40">
      <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"/>
      <head><!--[if gte mso 9]><xml>
      <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
      <![endif]--></head>
      <body>{table}</body></html>`;
      
  const MIME_XLS = "application/vnd.ms-excel;base64,";

  const parameters: any = {
    fileName,
    table: objectToTable(data, options?.headerStyle, options?.bodyStyle),
  };
  const buildedOutput = TEMPLATE_XLS.replace(
    /{(\w+)}/g,
    (x, y) => parameters[y]
  );

  if (window) {
    const excelBuild = new Blob([buildedOutput], {
      type: MIME_XLS,
    });

    const excelLink = window.URL.createObjectURL(excelBuild);
    downloadFile(excelLink, fileName + ".xls");
  } else {
    throw new Error(
      "Window is not definided: You must be using it in a browser"
    );
  }
};

export const exportmeToCsv = (data: any[], fileName: string) => {
  if (
    typeof fileName !== "string" ||
    Object.prototype.toString.call(fileName) !== "[object String]"
  ) {
    throw new Error(
      "Invalid input types: First Params should be an Array and the second one a String"
    );
  }

  if (window) {
    const computedCSV = new Blob([objectToSemicolons(data)], {
      type: "text/csv;charset=utf-8",
    });

    const csvLink = window.URL.createObjectURL(computedCSV);
    downloadFile(csvLink, fileName + ".csv");
  } else {
    throw new Error(
      "Window is not definided: You must be using it in a browser"
    );
  }
};

function objectToSemicolons(data: any[]) {
  const colsHead = Object.keys(data[0])
    .map((key) => [key])
    .join(";");
  const colsData = data
    .map((obj) => [
      Object.keys(obj)
        .map((col) => [obj[col]])
        .join(";"),
    ])
    .join("\n");

  return `${colsHead}\n${colsData}`;
}

function objectToTable(
  data: any[],
  headerStyle?: CSS.Properties,
  bodyStyle?: CSS.Properties
) {
  const colsHead = `<tr>${Object.keys(data[0])
    .map(
      (key) =>
        `<td align="center" style="${cssPropertyToStyleString(
          headerStyle
        )}"><b>${key}</b></td>`
    )
    .join("")}</tr>`;

  const colsData = data
    .map((obj) => [
      `<tr>
              ${Object.keys(obj)
                .map(
                  (col) =>
                    `<td style="${cssPropertyToStyleString(bodyStyle)}">${
                      obj[col] ? obj[col] : ""
                    }</td>`
                )
                .join("")}
          </tr>`,
    ])
    .join("");

  return `<table>${colsHead}${colsData}</table>`.trim();
}

function downloadFile(output: string, fileName: string) {
  const link = document.createElement("a");
  document.body.appendChild(link);
  link.download = fileName;
  link.href = output;
  link.click();
}

function cssPropertyToStyleString(cssObject?: any) {
  if (!cssObject) return;

  return Object.keys(cssObject)
    .map(
      (key) =>
        key.replace(/[A-Z]/g, (property) => `-${property.toLowerCase()}`) +
        `:${cssObject[key]}`
    )
    .join(";");
}
