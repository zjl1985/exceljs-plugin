import { Font } from 'exceljs';

export interface ExcelPluginOption {
  fileName: string;
  header?: Header[];
  data: any[];
  //需要合并的单元格 ['A1:B1','A2:C12']
  mergeCells?: string[];
  //列宽 {A:12}
  columnWidth?: { [key: string]: number };
  //列样式
  columnStyle?: { [key: string]: Style };
  //是否自动换行
  enbaleWrapText?: boolean;
  headerFooter?: {
    firstHeader?: string;
    firstFooter?: string;
  };
  csv?: boolean;
}

export interface ExcelPluginOptionUserDefine {
  fileName: string;
  data: any[][];
  //需要合并的单元格 ['A1:B1','A2:C12']
  mergeCells?: string[];
  //列宽 {A:12}
  columnWidth?: { [key: string]: number };
  //列样式
  columnStyle?: { [key: string]: Style };
  //是否自动换行
  enbaleWrapText?: boolean;
}

export interface ExcelPluginByDomOption {
  //需要合并的单元格 ['A1:B1','A2:C12']
  mergeCells?: string[];
  //列宽 {A:12}
  columnWidth?: { [key: string]: number };
  //列样式
  columnStyle?: { [key: string]: Style };
  //是否自动换行
  enbaleWrapText?: boolean;
}

export interface Header {
  header: string;
  key: string;
  width?: number;
  style?: Style;
}

export interface hfRow {
  data: string[];
  height?: number;
}

export interface Style {
  alignment?: {
    horizontal?: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed';
    vertical?: 'top' | 'middle' | 'bottom' | 'distributed' | 'justify';
    wrapText?: boolean;
    indent?: number;
    readingOrder?: 'rtl' | 'ltr';
    textRotation?: number | 'vertical';
  };
  numFmt?: string;
  font?: Partial<Font>;
  enbaleWrapText?: boolean;
}

export interface ExcelPlugin {
  export(option: ExcelPluginOption): void;
  exportByDom(dom: any, fileName: string): void;
  exportByDomPlugin(dom: any, fileName: string, opt: any, headerAndFooter: any): void;
}
