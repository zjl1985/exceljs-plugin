import * as Excel from 'exceljs/dist/exceljs.min.js';
import { Workbook, Worksheet, Row, Cell } from 'exceljs';
import { INDEX_TO_LETTER } from './basic-data';
import { FileProcess } from './file-process';

export class HandleForDom {
  //占位符
  seat: number = 123456.654321;

  public save(dom: any, name: string) {
    const workbook: Workbook = this.processWorkbook(dom);
    const process = new FileProcess();
    process.saveFile(workbook, name);
  }

  public processWorkbook(dom: any): Workbook {
    const workbook: Workbook = new Excel.Workbook();
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
    this.buildHead(sheet, dom);
    this.buildBody(sheet, dom);
    for (const col of sheet.columns) {
      col.width = 10;
    }
    return workbook;
  }

  private buildHead(sheet: Worksheet, dom: any) {
    const headerRows: any[] = dom
      .getElementsByTagName('thead')[0]
      .getElementsByTagName('tr');
    const rows = this.drawExcel(headerRows, sheet);
    for (const row of rows) {
      row.eachCell((cell: Cell, colNumber: number) => {
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.font = { size: 14, family: 2, bold: true };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    }
  }

  private buildBody(sheet: Worksheet, dom: any) {
    const bodyRows = dom
      .getElementsByTagName('tbody')[0]
      .getElementsByTagName('tr');
    const rows = this.drawExcel(bodyRows, sheet);
    for (const row of rows) {
      row.eachCell((cell: Cell, colNumber: number) => {
        cell.alignment = { vertical: 'middle' };
        cell.font = { size: 14, family: 1, bold: false };
        cell.border = {
          top: { style: 'thin' },
          left: { style: 'thin' },
          bottom: { style: 'thin' },
          right: { style: 'thin' },
        };
      });
    }
  }

  private drawExcel(headerRows: any[], sheet: Worksheet): Row[] {
    const start = sheet.rowCount;
    const rows = this.buildMatrix(headerRows);
    const rowCount = headerRows.length;
    const needMerge: String[] = [];
    const result: Row[] = [];
    for (let i = 0; i < rowCount; i++) {
      const tr = headerRows[i];
      const thLength = tr.cells.length;
      let celIndex = 0;
      for (let j = 0; j < thLength; j++) {
        const th = tr.cells[j];
        const displayDom = th.getElementsByClassName('display-excel');
        let displayText = '';
        if (displayDom.length > 0) {
          displayText = displayDom[0].innerText;
        } else {
          displayText = th.innerText;
          if (displayText === '') {
            const input = th.getElementsByTagName('input');
            if (input.length > 0) {
              displayText = input[0].value;
            }
          }
        }
        const colSpan = th.colSpan ? th.colSpan : 0;
        const rowSpan = th.rowSpan ? th.rowSpan : 0;
        if (rows[i][celIndex] !== this.seat) {
          celIndex++;
          j--;
          continue;
        }
        rows[i][celIndex] = displayText;
        let letter = INDEX_TO_LETTER[celIndex + colSpan - 1];
        let toIndex = i + rowSpan + start;
        this.fillRows(colSpan, rows, i, celIndex, th, rowSpan);
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
    return result;
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
      const span = th.colSpan ? th.colSpan : 1;
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
