import { HandleForDom } from './handle-for-dom';
import { ExcelPlugin, Header, ExcelPluginOption } from './interface';
import { HandleForData } from './handle-for-data';
export class ExcelPluginImpl implements ExcelPlugin {
  constructor() {}
  export(option: ExcelPluginOption): void {
    const handel = new HandleForData();
    handel.save(option);
  }

  exportByDom(dom: any, fileName: string): void {
    const handel = new HandleForDom();
    handel.save(dom, name);
  }
}
