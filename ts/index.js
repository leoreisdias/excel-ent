"use strict";
exports.__esModule = true;
exports.exportmeToCsv = exports.exportmeExcel = void 0;
var exportmeExcel = function (data, fileName, options) {
    if (!Array.isArray(data) ||
        typeof fileName !== "string" ||
        Object.prototype.toString.call(fileName) !== "[object String]") {
        throw new Error("Invalid input types: First Params should be an Array and the second one a String");
    }
    var TEMPLATE_XLS = "\n      <html xmlns:o=\"urn:schemas-microsoft-com:office:office\" xmlns:x=\"urn:schemas-microsoft-com:office:excel\" xmlns=\"https://www.w3.org/TR/html40\">\n      <meta http-equiv=\"content-type\" content=\"application/vnd.ms-excel; charset=UTF-8\"/>\n      <head><!--[if gte mso 9]><xml>\n      <x:ExcelWorkbook><x:ExcelWorksheets><x:ExcelWorksheet><x:Name>{title}</x:Name><x:WorksheetOptions><x:DisplayGridlines/></x:WorksheetOptions></x:ExcelWorksheet></x:ExcelWorksheets></x:ExcelWorkbook></xml>\n      <![endif]--></head>\n      <body>{table}</body></html>";
    var MIME_XLS = "application/vnd.ms-excel;base64,";
    var parameters = {
        fileName: fileName,
        table: objectToTable(data, options === null || options === void 0 ? void 0 : options.headerStyle, options === null || options === void 0 ? void 0 : options.bodyStyle)
    };
    var buildedOutput = TEMPLATE_XLS.replace(/{(\w+)}/g, function (x, y) { return parameters[y]; });
    if (window) {
        var excelBuild = new Blob([buildedOutput], {
            type: MIME_XLS
        });
        var excelLink = window.URL.createObjectURL(excelBuild);
        downloadFile(excelLink, fileName);
    }
    else {
        throw new Error("Window is not definided: You must be using it in a browser");
    }
};
exports.exportmeExcel = exportmeExcel;
var exportmeToCsv = function (data, fileName) {
    if (typeof fileName !== "string" ||
        Object.prototype.toString.call(fileName) !== "[object String]") {
        throw new Error("Invalid input types: First Params should be an Array and the second one a String");
    }
    if (window) {
        var computedCSV = new Blob([objectToSemicolons(data)], {
            type: "text/csv;charset=utf-8"
        });
        var csvLink = window.URL.createObjectURL(computedCSV);
        downloadFile(csvLink, fileName);
    }
    else {
        throw new Error("Window is not definided: You must be using it in a browser");
    }
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
function objectToTable(data, headerStyle, bodyStyle) {
    var colsHead = "<tr>" + Object.keys(data[0])
        .map(function (key) {
        return "<td align=\"center\" style=\"" + cssPropertyToStyleString(headerStyle) + "\"><b>" + key.toUpperCase() + "</b></td>";
    })
        .join("") + "</tr>";
    var colsData = data
        .map(function (obj) { return [
        "<tr>\n              " + Object.keys(obj)
            .map(function (col) {
            return "<td style=\"" + cssPropertyToStyleString(bodyStyle) + "\">" + (obj[col] ? obj[col] : "") + "</td>";
        })
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
function cssPropertyToStyleString(cssObject) {
    if (!cssObject)
        return;
    return Object.keys(cssObject)
        .map(function (key) {
        return key.replace(/[A-Z]/g, function (property) { return "-" + property.toLowerCase(); }) +
            (":" + cssObject[key]);
    })
        .join(";");
}
