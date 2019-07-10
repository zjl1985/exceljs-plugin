export interface ExcelPluginOption {
  fileName: string;
  header?: Header[];
  data: any[];
  mergeCells?: string[];
}

export interface Header {
  header: string;
  key: string;
  width?: number;
  style?: Style;
}

export interface Style {
  alignment?: {
    horizontal?:
      | 'left'
      | 'center'
      | 'right'
      | 'fill'
      | 'justify'
      | 'centerContinuous'
      | 'distributed';
    vertical?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
    wrapText?: boolean;
    indent?: number;
    readingOrder?: 'rtl' | 'ltr';
    textRotation?: number | 'vertical';
  };
}

export interface ExcelPlugin {
  export(option: ExcelPluginOption): void;
  exportByDom(dom: any, fileName: string): void;
}
