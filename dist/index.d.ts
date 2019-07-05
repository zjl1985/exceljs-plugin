export interface ExcelPluginConfig {}

export interface ExcelPlugin {
  hello(): void;
  exportByDom(dom: any, fileName: string): void;
}


declare module 'exceljs-plugin' {
  export const GaoxinExcelExport: ExcelPlugin;
}
