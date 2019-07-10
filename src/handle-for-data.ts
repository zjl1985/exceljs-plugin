import * as Excel from 'exceljs/dist/exceljs.min.js';
import { Workbook, Worksheet, Row, Cell } from 'exceljs';
import { FileProcess } from './file-process';
import { ExcelPluginOption, Header } from './interface';
export class HandleForData {
  public save(option: ExcelPluginOption) {
    if (option.data == undefined && option.data.length == 0) {
      console.warn('没有数据要导出,请不要盲目调用');
      return;
    }
    if (option.filName == undefined || option.filName.trim() === '') {
      console.warn('没有设置导出文件名,默认使用export');
      option.filName = 'export';
    }
    const workbook: Workbook = this.processWorkbook(option);
    const process = new FileProcess();
    process.saveFile(workbook, option.filName.trim());
  }

  public processWorkbook(
    option: ExcelPluginOption,
    mybook?: Workbook,
  ): Workbook {
    let workbook: Workbook = mybook ? mybook : new Excel.Workbook();
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
    if (option.header) {
      sheet.columns = option.header;
    } else {
      const columns: Header[] = [];
      for (const key in option.data[0]) {
        columns.push({ header: key, key: key, width: 15 });
      }
      sheet.columns = columns;
    }
    sheet.addRows(option.data);
    if (option.mergeCells && option.mergeCells.length > 0) {
      for (const cell of option.mergeCells) {
        // @ts-ignore
        sheet.mergeCells(cell);
      }
    }
    return workbook;
  }
}
