declare module 'html-to-excel-converter' {
  export default class HTMLToExcelConverter {
    constructor();
    convertHtmlToExcel(html: string): Promise<Buffer>;
  }
} 