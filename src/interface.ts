export interface ExcelPluginOption {
  filName: string;
  header?: Header[];
  data: any[];
  mergeCells?: string[];
}

export interface Header {
  header: string;
  key: string;
  width?: number;
}

export interface ExcelPlugin {
  export(option: ExcelPluginOption): void;
  exportByDom(dom: any, fileName: string): void;
}
