import { HandleForDom } from './handle-for-dom';
import { ExcelPlugin, Header } from './interface';
export class ExcelPluginImpl implements ExcelPlugin {
  constructor() {}
  export(data: any[], header: Header[], fileName: string): void {
    console.table(data);
  }

  exportByDom(dom: any, fileName: string): void {
    const handel = new HandleForDom();
    handel.save(dom, name);
  }
}
