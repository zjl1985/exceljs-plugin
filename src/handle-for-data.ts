import { Workbook, Worksheet, Row, Cell } from 'exceljs';
import { FileProcess } from './file-process';
import { ExcelPluginOption, Header, ExcelPluginOptionUserDefine } from './interface';
declare const ExcelJS: any;

export class HandleForData {
  enbaleWrapText = true;

  public save(option: ExcelPluginOption) {
    if (option.data == undefined && option.data.length == 0) {
      console.warn('没有数据要导出,请不要盲目调用');
      return;
    }
    if (option.fileName == undefined || option.fileName.trim() === '') {
      console.warn('没有设置导出文件名,默认使用export');
      option.fileName = 'export';
    }
    const workbook: Workbook = this.processWorkbook(option);
    const process = new FileProcess();
    process.saveFile(workbook, option.fileName.trim(), option.csv);
  }

  public processWorkbook(option: ExcelPluginOption, mybook?: Workbook): Workbook {
    let workbook: Workbook = mybook ? mybook : new ExcelJS.Workbook();
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
    if (option.header) {
      for (const header of option.header) {
        if (!header.width) {
          header.width = 15;
        }
      }
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
    if (option.headerFooter) {
      sheet.headerFooter.differentFirst = true;
      if (option.headerFooter.firstHeader) {
        sheet.headerFooter.firstHeader = option.headerFooter.firstHeader;
      }
      if (option.headerFooter.firstFooter) {
        sheet.headerFooter.firstFooter = option.headerFooter.firstFooter;
      }
    }
    return workbook;
  }

  saveUserDefine(opt: ExcelPluginOptionUserDefine) {
    let workbook: Workbook = new ExcelJS.Workbook();
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
    sheet.addRows(opt.data);
    if (opt.enbaleWrapText !== undefined) {
      this.enbaleWrapText = opt.enbaleWrapText;
    }
    sheet.eachRow((row, rowNumber) => {
      row.eachCell((cell: Cell, colNumber: number) => {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'left',
          wrapText: this.enbaleWrapText,
        };
        cell.font = { size: 12, family: 1, bold: false };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    });
    for (let index = 0; index < sheet.columns.length; index++) {
      const col = sheet.columns[index];
      col.width = 16;
    }
    if (opt.columnWidth) {
      for (const key in opt.columnWidth) {
        sheet.getColumn(key).width = opt.columnWidth[key];
      }
    }
    if (opt.mergeCells && opt.mergeCells.length > 0) {
      for (const cell of opt.mergeCells) {
        // @ts-ignore
        sheet.mergeCells(cell);
      }
    }

    if (opt.columnStyle && Object.keys(opt.columnStyle).length > 0) {
      for (const key in opt.columnStyle) {
        sheet.getColumn(key).style = opt.columnStyle[key];
      }
    }
    const process = new FileProcess();
    process.saveFile(workbook, opt.fileName.trim());
  }
}
