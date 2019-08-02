import { HandleForDom } from './handle-for-dom';
import {
  ExcelPlugin,
  Header,
  ExcelPluginOption,
  ExcelPluginByDomOption,
  hfRow,
  ExcelPluginOptionUserDefine,
} from './interface';
import { HandleForData } from './handle-for-data';
export class ExcelPluginImpl implements ExcelPlugin {
  constructor() {}
  export(option: ExcelPluginOption): void {
    const handel = new HandleForData();
    handel.save(option);
  }

  exportByDom(dom: any, fileName: string): void {
    const handel = new HandleForDom();
    handel.save(dom, fileName);
  }

  exportByDomPlugin(
    dom: any,
    fileName: string,
    opt: ExcelPluginByDomOption,
    headerAndFooter?: { header?: hfRow[]; footer?: hfRow[] },
  ): void {
    const handel = new HandleForDom();
    handel.savePlugin(dom, fileName, opt, headerAndFooter);
  }

  exportUserDefine(opt: ExcelPluginOptionUserDefine) {
    const handel = new HandleForData();
    handel.saveUserDefine(opt);
  }
}
