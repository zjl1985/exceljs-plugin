import * as Excel from 'exceljs/dist/exceljs.min';
import { Workbook, Worksheet, Row, Cell } from 'exceljs';
import { INDEX_TO_LETTER } from './basic-data';
import { FileProcess } from './file-process';
import { Style, ExcelPluginByDomOption } from './interface';

export class HandleForDom {
  //占位符
  seat: number = 123456.654321;
  //高度系数
  heightRatio = 0.6;
  //宽度系数
  widthRatio = 0.12;

  enbaleWrapText: boolean = true;

  public save(dom: any, name: string) {
    const workbook: Workbook = this.processWorkbook(dom);
    const process = new FileProcess();
    process.saveFile(workbook, name);
  }

  public savePlugin(
    dom: any,
    name: string,
    opt?: ExcelPluginByDomOption,
    headerAndFooter?: { header?: string[][]; footer?: string[][] },
  ) {
    if (opt.enbaleWrapText !== undefined) {
      this.enbaleWrapText = opt.enbaleWrapText;
    }
    const workbook: Workbook = this.processWorkbook(dom, opt, headerAndFooter);
    const process = new FileProcess();
    process.saveFile(workbook, name);
  }

  public processWorkbook(
    dom: any,
    opt?: ExcelPluginByDomOption,
    headerAndFooter?: { header?: string[][]; footer?: string[][] },
  ): Workbook {
    const workbook: Workbook = new Excel.Workbook();
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
    if (
      headerAndFooter &&
      headerAndFooter.header &&
      headerAndFooter.header.length > 0
    ) {
      sheet.addRows(headerAndFooter.header);
    }
    this.buildHead(sheet, dom);
    this.buildBody(sheet, dom);
    if (
      headerAndFooter &&
      headerAndFooter.footer &&
      headerAndFooter.footer.length > 0
    ) {
      sheet.addRows(headerAndFooter.footer);
    }

    if (opt && opt.mergeCells && opt.mergeCells.length > 0) {
      for (const cell of opt.mergeCells) {
        sheet.mergeCells(cell);
      }
    }

    if (opt && opt.columnWidth) {
      for (const key in opt.columnWidth) {
        sheet.getColumn(key).width = opt.columnWidth[key];
      }
    }

    if (opt && opt.columnStyle) {
      for (const key in opt.columnStyle) {
        sheet.getColumn(key).style = opt.columnStyle[key];
      }
    }
    return workbook;
  }

  private buildHead(sheet: Worksheet, dom: any) {
    const headerRows: any[] = dom
      .getElementsByTagName('thead')[0]
      .getElementsByTagName('tr');
    const [rows, tableSize] = this.drawExcel(headerRows, sheet);
    const rowStyle = tableSize.rowStyle;
    this.setStyle(rows, rowStyle);
  }

  private buildBody(sheet: Worksheet, dom: any) {
    const bodyRows = dom
      .getElementsByTagName('tbody')[0]
      .getElementsByTagName('tr');
    const [rows, tableSize] = this.drawExcel(bodyRows, sheet);
    const rowStyle = tableSize.rowStyle;
    const colStyle = tableSize.colStyle;

    for (let index = 0; index < sheet.columns.length; index++) {
      const col = sheet.columns[index];
      col.width = colStyle[index].width;
    }
    this.setStyle(rows, rowStyle);
  }

  private setStyle(rows: Row[], rowStyle: any) {
    for (let index = 0; index < rows.length; index++) {
      const row = rows[index];
      row.height = rowStyle[index].height;
      row.eachCell((cell: Cell, colNumber: number) => {
        cell.alignment = {
          vertical: 'middle',
          horizontal: 'center',
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
    }
  }

  private drawExcel(
    headerRows: any[],
    sheet: Worksheet,
  ): [Row[], { rowStyle: any; colStyle: any }] {
    const start = sheet.rowCount;
    const rows = this.buildMatrix(headerRows);
    const rowCount = headerRows.length;
    const needMerge: String[] = [];
    const result: Row[] = [];
    const rowStyle = [];
    const colStyle = [];
    for (let i = 0; i < rowCount; i++) {
      const tr = headerRows[i];
      rowStyle.push({
        height: tr.offsetHeight * this.heightRatio,
      });
      const cellLength = tr.cells.length;
      let celIndex = 0;
      for (let j = 0; j < cellLength; j++) {
        const cell = tr.cells[j];
        const displayDom = cell.getElementsByClassName('display-excel');
        let displayText = '';
        if (displayDom.length > 0) {
          displayText = displayDom[0].innerText;
        } else {
          displayText = cell.innerText;
          if (displayText === '') {
            const input = cell.getElementsByTagName('input');
            if (input.length > 0) {
              displayText = input[0].value;
            }
          }
        }
        const colSpan = cell.colSpan;
        const rowSpan = cell.rowSpan;
        console.log(colSpan);
        if (colSpan === 1) {
          colStyle[j] = { width: cell.offsetWidth * this.widthRatio };
        }
        if (rows[i][celIndex] !== this.seat) {
          celIndex++;
          j--;
          continue;
        }
        rows[i][celIndex] = displayText;
        let letter = INDEX_TO_LETTER[celIndex + colSpan - 1];
        let toIndex = i + rowSpan + start;
        this.fillRows(colSpan, rows, i, celIndex, cell, rowSpan);
        if (rowSpan > 1 || colSpan > 1) {
          needMerge.push(
            `${INDEX_TO_LETTER[celIndex] + (i + 1 + start)}:${letter +
              toIndex}`,
          );
        }
        celIndex++;
      }
    }
    for (const item of rows) {
      result.push(sheet.addRow(item));
    }
    for (const cell of needMerge) {
      // @ts-ignore
      sheet.mergeCells(cell);
    }
    const tableSize = {
      rowStyle: rowStyle,
      colStyle: colStyle,
    };
    return [result, tableSize];
  }

  private fillRows(
    colSpan: any,
    rows: any[],
    i: number,
    celIndex: number,
    th: any,
    rowSpan: any,
  ) {
    if (colSpan > 1) {
      for (let index = 0; index < colSpan; index++) {
        if (rowSpan > 1) {
          for (let rowIndex = 0; rowIndex < rowSpan; rowIndex++) {
            rows[i + rowIndex][celIndex + index] = th.innerText;
          }
        } else {
          rows[i][celIndex + index] = th.innerText;
        }
      }
      return;
    }
    if (rowSpan > 1) {
      for (let index = 0; index < rowSpan; index++) {
        rows[i + index][celIndex] = th.innerText;
      }
    }
  }

  private buildMatrix(headerRows: any[]): any[] {
    const height: number = headerRows.length;
    let width: number = 0;
    for (const th of headerRows[0].cells) {
      const span = th.colSpan;
      width = width + span;
    }
    const matrix = new Array(height);
    for (let i = 0; i < height; i++) {
      const row = new Array(width);
      for (let j = 0; j < width; j++) {
        row[j] = this.seat;
      }
      matrix[i] = row;
    }
    return matrix;
  }
}
