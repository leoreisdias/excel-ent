"use strict";
exports.__esModule = true;
exports.exportmeToCsv = exports.exportmeExcel = void 0;
var buffer_1 = require("buffer");
var exportmeExcel = function (data, fileName) {
    if (!Array.isArray(data) ||
        typeof fileName !== "string" ||
        Object.prototype.toString.call(fileName) !== "[object String]") {
        throw new Error("Invalid input types: First Params should be an Array and the second one a String");
    }
    var TEMPLATE_XLS = "\n      <html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"http://www.w3.org/TR/REC-html40\">\n      <meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\"/>\n      <head><!--[if gte mso 9]><xml>\n      <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>\n      <![endif]--></head>\n      <body>{table}</body></html>";
    var MIME_XLS = "application/vnd.ms-excel;base64,";
    var parameters = {
        fileName: fileName,
        table: objectToTable(data)
    };
    var buildedOutput = TEMPLATE_XLS.replace(/{(\w+)}/g, function (x, y) { return parameters[y]; });
    var excelBuild = new buffer_1.Blob([buildedOutput], {
        type: MIME_XLS
    });
    var excelLink = window.URL.createObjectURL(excelBuild);
    downloadFile(excelLink, fileName);
};
exports.exportmeExcel = exportmeExcel;
var exportmeToCsv = function (data, fileName) {
    if (typeof fileName !== "string" ||
        Object.prototype.toString.call(fileName) !== "[object String]") {
        throw new Error("Invalid input types: First Params should be an Array and the second one a String");
    }
    var computedCSV = new buffer_1.Blob([objectToSemicolons(data)], {
        type: "text/csv;charset=utf-8"
    });
    var csvLink = window.URL.createObjectURL(computedCSV);
    downloadFile(csvLink, fileName);
};
exports.exportmeToCsv = exportmeToCsv;
function objectToSemicolons(data) {
    var colsHead = Object.keys(data[0])
        .map(function (key) { return [key]; })
        .join(";");
    var colsData = data
        .map(function (obj) { return [
        Object.keys(obj)
            .map(function (col) { return [obj[col]]; })
            .join(";"),
    ]; })
        .join("\n");
    return colsHead + "\n" + colsData;
}
function objectToTable(data) {
    var colsHead = "<tr>" + Object.keys(data[0])
        .map(function (key) { return "<td>" + key + "</td>"; })
        .join("") + "</tr>";
    var colsData = data
        .map(function (obj) { return [
        "<tr>\n              " + Object.keys(obj)
            .map(function (col) { return "<td>" + (obj[col] ? obj[col] : "") + "</td>"; })
            .join("") + "\n          </tr>",
    ]; })
        .join("");
    return ("<table>" + colsHead + colsData + "</table>").trim();
}
function downloadFile(output, fileName) {
    var link = document.createElement("a");
    document.body.appendChild(link);
    link.download = fileName;
    link.href = output;
    link.click();
}
console.log(exports.exportmeExcel([
    {
        id: 1,
        nome: "Leonardo"
    },
    {
        id: 2,
        name: "jose"
    },
], "Teste de Excel"));
