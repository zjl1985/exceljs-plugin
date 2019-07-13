import { GaoxinExcelExport } from '../src';
import { ExcelPluginOption, Header } from '../src/interface';
import { HandleForData } from '../src/handle-for-data';
const Excel = require('exceljs');
test('GaoxinExcelExport', () => {
  const handel = new HandleForData();
  const header: Header[] = [
    { header: 'a', key: 'a' },
    { header: 'bbbb', key: 'b', style: { alignment: { horizontal: 'center' } } },
  ];
  const data = [{ a: 'nihsdfffffffffffffffffffffffffffao', b: 1 }, { a: 'helsdafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafafaflo', b: 'world' }];
  const workbookMy = new Excel.Workbook();
  const opt: ExcelPluginOption = {
    fileName: 'eee',
    header: header,
    data: data,
  };
  const workboot = handel.processWorkbook(opt, workbookMy);
   workboot.xlsx.writeFile('./test.xlsx').then(() => {
    expect('hello').toBe('hello');
  });
});
