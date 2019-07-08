import * as Excel from 'exceljs/dist/exceljs.min.js';
import { Workbook, Worksheet, Row, Cell } from 'exceljs';
import { FileProcess } from './file-process';
import { Header } from './interface';
export class HandleForData {
  public save(data: any[], header: Header[], fileName: string) {
    const workbook: Workbook = this.processWorkbook(data, header);
    const process = new FileProcess();
    process.saveFile(workbook, fileName);
  }

  public processWorkbook(
    data: any[],
    header: Header[],
    mybook?: Workbook,
  ): Workbook {
    let workbook: Workbook;
    if (mybook) {
      workbook = mybook;
    } else {
      workbook = new Excel.Workbook();
    }
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
    sheet.columns = header;
    sheet.addRows(data);
    return workbook;
  }
}
