import * as CSS from "csstype";
export declare const exportmeExcel: (data: any[], fileName: string, options?: {
    headerStyle: CSS.Properties;
    bodyStyle: CSS.Properties;
}) => void;
export declare const exportmeToCsv: (data: any[], fileName: string) => void;
