export const exportmeExcel = (data: any[], fileName: string) => {
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
      <html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel" xmlns="http://www.w3.org/TR/REC-html40">
      <meta http-equiv="content-type" content="application/vnd.ms-excel; charset=UTF-8"/>
      <head><!--[if gte mso 9]><xml>
      <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>
      <![endif]--></head>
      <body>{table}</body></html>`;
  const MIME_XLS = "application/vnd.ms-excel;base64,";

  const parameters = {
    fileName,
    table: objectToTable(data),
  };
  const buildedOutput = TEMPLATE_XLS.replace(
    /{(\w+)}/g,
    (x, y) => parameters[y]
  );

  const excelBuild = new Blob([buildedOutput], {
    type: MIME_XLS,
  });

  const excelLink = window.URL.createObjectURL(excelBuild);
  downloadFile(excelLink, fileName);
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
  const computedCSV = new Blob([objectToSemicolons(data)], {
    type: "text/csv;charset=utf-8",
  });

  const csvLink = window.URL.createObjectURL(computedCSV);
  downloadFile(csvLink, fileName);
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

function objectToTable(data: any[]) {
  const colsHead = `<tr>${Object.keys(data[0])
    .map((key) => `<td>${key}</td>`)
    .join("")}</tr>`;

  const colsData = data
    .map((obj) => [
      `<tr>
              ${Object.keys(obj)
                .map((col) => `<td>${obj[col] ? obj[col] : ""}</td>`)
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

console.log(
  exportmeExcel(
    [
      {
        id: 1,
        nome: "Leonardo",
      },
      {
        id: 2,
        name: "jose",
      },
    ],
    "Teste de Excel"
  )
);
