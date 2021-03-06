import { Workbook, Worksheet, Row, Cell } from 'exceljs';
import { INDEX_TO_LETTER } from './basic-data';
import { FileProcess } from './file-process';
import { ExcelPluginByDomOption, hfRow } from './interface';
declare const ExcelJS: any;
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
    headerAndFooter?: { header?: hfRow[]; footer?: hfRow[] },
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
    headerAndFooter?: { header?: hfRow[]; footer?: hfRow[] },
  ): Workbook {
    const workbook: Workbook = new ExcelJS.Workbook();
    const sheet: Worksheet = workbook.addWorksheet('sheet1');
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
    if (headerAndFooter && headerAndFooter.header && headerAndFooter.header.length > 0) {
      for (const item of headerAndFooter.header) {
        this.buildHF(sheet, item);
      }
    }
    this.buildHead(sheet, dom);
    this.buildBody(sheet, dom);
    if (headerAndFooter && headerAndFooter.footer && headerAndFooter.footer.length > 0) {
      for (const item of headerAndFooter.footer) {
        this.buildHF(sheet, item);
      }
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

    if (opt && opt.columnStyle && Object.keys(opt.columnStyle).length > 0) {
      for (const key in opt.columnStyle) {
        for (const styleKey in opt.columnStyle[key]) {
          sheet.getColumn(key)[styleKey] = opt.columnStyle[key][styleKey];
        }
      }
    }
    //批量更改某种样式下的单元格数组
    if (opt && opt.styleToCellList && opt.styleToCellList.length > 0) {
      opt.styleToCellList.forEach((item) => {
        const style = item.cellStyle;
        item.cellList.forEach((cellCode) => {
          for (const styleKey in style) {
            sheet.getCell(cellCode)[styleKey] = style[styleKey];
          }
        });
      });
    }

    if (opt && opt.cellStyle && Object.keys(opt.cellStyle).length > 0) {
      for (const cellCode in opt.cellStyle) {
        for (const styleKey in opt.cellStyle[cellCode]) {
          sheet.getCell(cellCode)[styleKey] = opt.cellStyle[cellCode][styleKey];
        }
      }
    }
    return workbook;
  }

  private buildHF(sheet: Worksheet, item: hfRow) {
    const row: Row = sheet.addRow(item.data);
    if (item.height) {
      row.height = item.height;
    }
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

  private buildHead(sheet: Worksheet, dom: any) {
    if (dom.getElementsByTagName('thead')) {
      const headerRows: any[] = dom.getElementsByTagName('thead')[0].getElementsByTagName('tr');
      const [rows, tableSize] = this.drawExcel(headerRows, sheet);
      const rowStyle = tableSize.rowStyle;
      this.setStyle(rows, rowStyle, headerRows);
    }
  }

  private buildBody(sheet: Worksheet, dom: any) {
    const bodyRows = dom.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
    const [rows, tableSize] = this.drawExcel(bodyRows, sheet);
    const rowStyle = tableSize.rowStyle;
    const colStyle = tableSize.colStyle;
    for (let index = 0; index < sheet.columns.length; index++) {
      const col = sheet.columns[index];
      if (colStyle.length > 0 && colStyle[index] && colStyle[index].width !== undefined) {
        col.width = colStyle[index].width;
      } else {
        col.width = 32;
      }
    }
    this.setStyle(rows, rowStyle, bodyRows);
  }

  private setStyle(rows: Row[], rowStyle: any, domRows) {
    for (let index = 0; index < rows.length; index++) {
      const row = rows[index];
      let tds;
      if (domRows) {
        tds = domRows[index].getElementsByTagName('td');
        if (tds === null || tds.length === 0) {
          tds = domRows[index].getElementsByTagName('th');
        }
      }
      let cellIndex = 0;
      row.eachCell((cell: Cell, colNumber: number) => {
        let textAlign: 'left' | 'center' | 'right' | 'fill' | 'justify' | 'centerContinuous' | 'distributed' = 'left';
        if (tds && tds[cellIndex]) {
          if (tds[cellIndex].style['textAlign'] !== '') {
            textAlign = tds[cellIndex].style['textAlign'];
          }
        }
        cellIndex++;
        cell.alignment = {
          vertical: 'middle',
          horizontal: textAlign,
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

  private drawExcel(headerRows: any[], sheet: Worksheet): [Row[], { rowStyle: any; colStyle: any }] {
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
        if (colSpan === 1) {
          colStyle[j] = { width: cell.offsetWidth * this.widthRatio };
        }
        if (rows[i][celIndex] !== this.seat) {
          celIndex++;
          j--;
          continue;
        }
        const cellType = cell.getAttribute('excel-cell-type');
        switch (cellType) {
          case 'number':
            rows[i][celIndex] = Number(displayText.trim());
            break;
          default:
            rows[i][celIndex] = displayText.trim();
            break;
        }

        let letter = INDEX_TO_LETTER[celIndex + colSpan - 1];
        let toIndex = i + rowSpan + start;
        this.fillRows(colSpan, rows, i, celIndex, displayText.trim(), rowSpan);
        if (rowSpan > 1 || colSpan > 1) {
          needMerge.push(`${INDEX_TO_LETTER[celIndex] + (i + 1 + start)}:${letter + toIndex}`);
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

  private fillRows(colSpan: any, rows: any[], i: number, celIndex: number, text: string, rowSpan: any) {
    if (colSpan > 1) {
      for (let index = 0; index < colSpan; index++) {
        if (rowSpan > 1) {
          for (let rowIndex = 0; rowIndex < rowSpan; rowIndex++) {
            rows[i + rowIndex][celIndex + index] = text;
          }
        } else {
          rows[i][celIndex + index] = text;
        }
      }
      return;
    }
    if (rowSpan > 1) {
      for (let index = 0; index < rowSpan; index++) {
        rows[i + index][celIndex] = text;
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
