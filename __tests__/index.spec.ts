import { HandleForData } from '../src/handle-for-data';
import { Header } from '../dist/interface';
const Excel = require('exceljs')
test('adds 1 + 2 to equal 3', () => {
  const handel = new HandleForData();
  const header: Header[] = [
    { header: 'a', key: 'a' },
    { header: 'bbbb', key: 'b' },
  ];
  const data = [{ a: 'nihao', b: 1 }, { a: 'hello', b: 'world' }];
  const workbookMy = new Excel.Workbook()
  const workboot = handel.processWorkbook(data, header,workbookMy);
  console.log(workboot);
  workboot.xlsx.writeFile('./test.xlsx').then(() => {
    expect('hello').toBe('hello');
  });
});
