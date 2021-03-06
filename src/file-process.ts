import * as FileSaver from 'file-saver';
import { Workbook } from 'exceljs';

export class FileProcess {
  public saveFile(workbook: Workbook, name: string, csv?: boolean) {
    if (csv) {
      workbook.csv.writeBuffer().then((data: any) => {
        let blob = new Blob([data], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        FileSaver.saveAs(blob, `${name}.xlsx`);
      });
    } else {
      workbook.xlsx.writeBuffer().then((data: any) => {
        let blob = new Blob([data], {
          type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        });
        FileSaver.saveAs(blob, `${name}.xlsx`);
      });
    }
  }
}
