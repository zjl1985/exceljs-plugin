import { Font, Borders } from 'exceljs';
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
  //批量设置单元格样式（优先级在列样式之后、单元格样式之前）
  styleToCellList?: {
    cellStyle: Style;
    cellList: string[];
  }[];
  //单元格样式{A2:{alignment:{horizontal:'left'}}}
  cellStyle?: { [key: string]: Style };
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
  //批量设置单元格样式（优先级在列样式之后、单元格样式之前）
  styleToCellList?: {
    cellStyle: Style;
    cellList: string[];
  }[];
  //单元格样式（优先级最高）
  cellStyle?: { [key: string]: Style };
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
  //批量设置单元格样式（优先级在列样式之后、单元格样式之前）
  styleToCellList?: {
    cellStyle: Style;
    cellList: string[];
  }[];
  //单元格样式{A2:{alignment:{horizontal:'left'}}}
  cellStyle?: { [key: string]: Style };
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
  //边框设置、默认为'thin',不需要边框直接border:null，设置某一方向border:{left:null}
  border?: Partial<Borders>;
  numFmt?: string;
  font?: Partial<Font>;
  enbaleWrapText?: boolean;
}

export interface ExcelPlugin {
  export(option: ExcelPluginOption): void;
  exportByDom(dom: any, fileName: string): void;
  exportByDomPlugin(dom: any, fileName: string, opt: ExcelPluginOption, headerAndFooter: any): void;
}
