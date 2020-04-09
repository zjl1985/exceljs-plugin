import { GaoxinExcelExport } from '../src';
import { ExcelPluginOption, Header, ExcelPluginOptionUserDefine } from '../src/interface';
import { HandleForData } from '../src/handle-for-data';
import { Workbook, Worksheet, Cell } from 'exceljs';
const ExcelJS = require('exceljs');
test('GaoxinExcelExport', () => {
  // const handel = new HandleForData();
  // const header: Header[] = [
  //   { header: 'a', key: 'a' },
  //   { header: 'bbbb', key: 'b', style: { alignment: { horizontal: 'center' } } },
  // ];
  // const data = [{ a: 'nihsdfffffffffffffffffffffffffffao', b: 1 }, { a: 'helsdafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafaflo', b: 'world' }];
  // const workbookMy = new Excel.Workbook();
  // const opt: ExcelPluginOption = {
  //   fileName: 'eee',
  //   header: header,
  //   data: data,
  // };
  // const workboot = handel.processWorkbook(opt, workbookMy);
  //  workboot.xlsx.writeFile('./test.xlsx').then(() => {
  //   expect('hello').toBe('hello');
  // });
  const title = [];
  const name = '上报数据项报表';
  title.push(name);

  const body = [
    [
      '全厂-二次能源-电力-购入已消费量（电力）',
      '全厂',
      '',
      '',
      '千瓦时',
      '',
      '吨标准煤/万千瓦时',
      '1.2290',
      '没有采集',
    ],
  ];
  body;

  const head = ['指标名称', '指标范围', '时间', '指标值', '计量单位', '折标值', '折标单位', '折标系数', '备注'];
  const result = [title, head, ...body];

  const opt: ExcelPluginOptionUserDefine = {
    fileName: name,
    data: result,
    mergeCells: ['A1:B1'],
    columnWidth: {
      A: 40,
    },
    columnStyle: {
      A: {
        alignment: {
          horizontal: 'center',
        },
      },
    },
  };

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
      for (const styleKey in opt.columnStyle[key]) {
        sheet.getColumn(key)[styleKey] = opt.columnStyle[key][styleKey];
      }
    }
  }
  workbook.xlsx.writeFile('./test.xlsx').then(() => {
    expect('hello').toBe('hello');
  });
});
