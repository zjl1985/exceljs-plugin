export interface ExcelPluginConfig {}

export interface Header {
  header: string;
  key: string;
  width?: number;
}

export interface ExcelPlugin {
  export(data: any[], header: Header[], fileName: string): void;
  exportByDom(dom: any, fileName: string): void;
}
