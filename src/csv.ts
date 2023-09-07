import { objectToSemicolons } from "./helpers/convert";

const downloadFile = (output: string, fileName: string): void => {
  const link = document.createElement("a");
  document.body.appendChild(link);
  link.download = fileName;
  link.href = output;
  link.click();
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
    downloadFile(csvLink, `${fileName}.csv`);
  } else {
    throw new Error(
      "Window is not definided: You must be using it in a browser"
    );
  }
};
